"""Linux desktop file chooser via the XDG Desktop Portal.

This module intentionally talks to the user's existing desktop session over
D-Bus. It does not bundle GTK or Qt; the desktop portal backend supplies the
native-looking dialog.
"""

from __future__ import annotations

import asyncio
import os
import uuid
from pathlib import Path
from typing import Sequence, Tuple
from urllib.parse import urlsplit


PORTAL_BUS = "org.freedesktop.portal.Desktop"
PORTAL_PATH = "/org/freedesktop/portal/desktop"
FILE_CHOOSER_IFACE = "org.freedesktop.portal.FileChooser"
REQUEST_IFACE = "org.freedesktop.portal.Request"
PROPERTIES_IFACE = "org.freedesktop.DBus.Properties"
PORTAL_VERSION_FOR_DIRECTORY_SELECTION = 3


class LinuxPortalUnavailable(RuntimeError):
    """The desktop session has no usable XDG file chooser portal."""


class LinuxPortalSelectionError(RuntimeError):
    """The desktop portal failed while opening or completing a chooser."""


def _dbus_error(reply: object) -> str:
    body = getattr(reply, "body", None) or []
    error_name = getattr(reply, "error_name", None) or "D-Bus error"
    detail = str(body[0]) if body else ""
    return f"{error_name}: {detail}" if detail else str(error_name)


def _message_failed(reply: object) -> bool:
    from dbus_next import MessageType

    return getattr(reply, "message_type", None) == MessageType.ERROR


def _variant_value(value: object) -> object:
    return getattr(value, "value", value)


def _file_uri_to_path(uri: str) -> str:
    """Convert a portal's file URI to the local path expected by core."""
    from urllib.request import url2pathname

    parsed = urlsplit(str(uri))
    if parsed.scheme.lower() != "file":
        raise LinuxPortalSelectionError("桌面选择器返回了非本地文件地址。")
    path = url2pathname(parsed.path)
    if parsed.netloc and parsed.netloc.lower() != "localhost":
        path = "//" + parsed.netloc + path
    return os.path.normpath(path)


def _directory_bytes(current: str) -> bytes:
    folder = Path(current).expanduser() if current else Path.home()
    if folder.is_file():
        folder = folder.parent
    if not folder.is_dir():
        folder = Path.home()
    return os.fsencode(str(folder)) + b"\0"


def _filters(file_types: Sequence[Tuple[str, str]]) -> list:
    """Translate the existing (label, glob string) filters to portal format."""
    output = []
    for label, globs in file_types:
        patterns = [[0, item] for item in str(globs).split() if item]
        if patterns:
            output.append([str(label), patterns])
    return output


async def _portal_version(timeout: float) -> int:
    from dbus_next import Message, MessageType
    from dbus_next.aio import MessageBus

    bus = None
    try:
        bus = await asyncio.wait_for(MessageBus().connect(), timeout=timeout)
        reply = await asyncio.wait_for(
            bus.call(
                Message(
                    destination=PORTAL_BUS,
                    path=PORTAL_PATH,
                    interface=PROPERTIES_IFACE,
                    member="Get",
                    signature="ss",
                    body=[FILE_CHOOSER_IFACE, "version"],
                )
            ),
            timeout=timeout,
        )
        if reply.message_type == MessageType.ERROR or not reply.body:
            raise LinuxPortalUnavailable(_dbus_error(reply))
        return int(_variant_value(reply.body[0]))
    except LinuxPortalUnavailable:
        raise
    except Exception as exc:
        raise LinuxPortalUnavailable(str(exc)) from exc
    finally:
        if bus is not None:
            bus.disconnect()


def portal_available(timeout: float = 1.5) -> bool:
    """Return whether this login session exposes directory-capable FileChooser."""
    if not os.environ.get("DBUS_SESSION_BUS_ADDRESS"):
        return False
    try:
        return asyncio.run(_portal_version(timeout)) >= PORTAL_VERSION_FOR_DIRECTORY_SELECTION
    except Exception:
        return False


async def _select_paths(
    *,
    title: str,
    folder: bool,
    current: str,
    multiple: bool,
    file_types: Sequence[Tuple[str, str]],
    timeout: float,
) -> Tuple[str, ...]:
    from dbus_next import Message, MessageType, Variant
    from dbus_next.aio import MessageBus

    bus = None
    match_rule = "type='signal',interface='org.freedesktop.portal.Request',member='Response'"
    handler = None
    match_installed = False
    try:
        bus = await asyncio.wait_for(MessageBus().connect(), timeout=timeout)
        if not bus.unique_name:
            raise LinuxPortalUnavailable("无法连接桌面 D-Bus 会话。")

        signals: asyncio.Queue = asyncio.Queue()

        def handle_message(message):
            if (
                message.message_type == MessageType.SIGNAL
                and message.interface == REQUEST_IFACE
                and message.member == "Response"
            ):
                signals.put_nowait(message)
            return None

        handler = handle_message
        bus.add_message_handler(handler)
        match_reply = await bus.call(
            Message(
                destination="org.freedesktop.DBus",
                path="/org/freedesktop/DBus",
                interface="org.freedesktop.DBus",
                member="AddMatch",
                signature="s",
                body=[match_rule],
            )
        )
        if _message_failed(match_reply):
            raise LinuxPortalUnavailable(_dbus_error(match_reply))
        match_installed = True

        token = "base_audit_" + uuid.uuid4().hex
        options = {
            "handle_token": Variant("s", token),
            "modal": Variant("b", True),
            "multiple": Variant("b", bool(multiple and not folder)),
            "directory": Variant("b", bool(folder)),
            "current_folder": Variant("ay", _directory_bytes(current)),
        }
        file_filters = _filters(file_types)
        if file_filters:
            options["filters"] = Variant("a(sa(us))", file_filters)

        reply = await bus.call(
            Message(
                destination=PORTAL_BUS,
                path=PORTAL_PATH,
                interface=FILE_CHOOSER_IFACE,
                member="OpenFile",
                signature="ssa{sv}",
                body=["", title, options],
            )
        )
        if _message_failed(reply) or not reply.body:
            raise LinuxPortalUnavailable(_dbus_error(reply))
        request_path = str(reply.body[0])

        deadline = asyncio.get_running_loop().time() + timeout
        response_message = None
        while response_message is None:
            remaining = deadline - asyncio.get_running_loop().time()
            if remaining <= 0:
                raise LinuxPortalSelectionError("桌面文件选择窗口等待超时。")
            signal = await asyncio.wait_for(signals.get(), timeout=remaining)
            if signal.path == request_path:
                response_message = signal

        response_code = int(response_message.body[0])
        results = response_message.body[1] if len(response_message.body) > 1 else {}
        if response_code == 1:
            return ()
        if response_code != 0:
            raise LinuxPortalSelectionError("桌面文件选择窗口未能完成操作。")

        uris = _variant_value(results.get("uris", []))
        return tuple(_file_uri_to_path(uri) for uri in (uris or []))
    except (LinuxPortalUnavailable, LinuxPortalSelectionError):
        raise
    except Exception as exc:
        raise LinuxPortalSelectionError(str(exc)) from exc
    finally:
        if bus is not None:
            if handler is not None:
                bus.remove_message_handler(handler)
            if match_installed:
                try:
                    await bus.call(
                        Message(
                            destination="org.freedesktop.DBus",
                            path="/org/freedesktop/DBus",
                            interface="org.freedesktop.DBus",
                            member="RemoveMatch",
                            signature="s",
                            body=[match_rule],
                        )
                    )
                except Exception:
                    pass
            bus.disconnect()


def select_paths(
    *,
    title: str,
    folder: bool,
    current: str,
    multiple: bool = False,
    file_types: Sequence[Tuple[str, str]] = (),
    timeout: float = 3600.0,
) -> Tuple[str, ...]:
    """Open the desktop chooser and return selected local filesystem paths."""
    if not portal_available(timeout=min(timeout, 2.0)):
        raise LinuxPortalUnavailable("当前桌面没有可用的 XDG 文件选择服务。")
    try:
        return asyncio.run(
            _select_paths(
                title=title,
                folder=folder,
                current=current,
                multiple=multiple,
                file_types=file_types,
                timeout=timeout,
            )
        )
    except (LinuxPortalUnavailable, LinuxPortalSelectionError):
        raise
    except Exception as exc:
        raise LinuxPortalSelectionError(str(exc)) from exc
