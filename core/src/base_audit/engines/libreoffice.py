"""LibreOffice Calc calculation support for UOS/Deepin and Linux.

Both direct DEB installations and the UOS application-store (Linglong)
edition are supported. Calculation always uses a private profile and staged
files, so it never touches a user's open LibreOffice session or a source file.
"""

from __future__ import annotations

import json
import os
import re
import shutil
import subprocess
import tempfile
import time
from dataclasses import dataclass
from pathlib import Path, PurePosixPath
from typing import Iterable
from urllib.parse import quote


LINGLONG_APP_DEFAULT = "org.libreoffice.libreoffice"
KARE_APP_ENV = "BASE_AUDIT_KARE_APP_ID"
KARE_UPPER_ENV = "BASE_AUDIT_KARE_UPPER_ROOT"
KARE_CONFIG_NAME = "libreoffice-kare.json"


def _system_command_env() -> dict[str, str]:
    """Frozen PyInstaller apps restore the host loader path for LibreOffice/ll-cli."""
    env = os.environ.copy()
    if "LD_LIBRARY_PATH_ORIG" in env:
        original = env.pop("LD_LIBRARY_PATH_ORIG")
        if original:
            env["LD_LIBRARY_PATH"] = original
        else:
            env.pop("LD_LIBRARY_PATH", None)
    return env


class LibreOfficeUnavailableError(RuntimeError):
    """Raised when a native LibreOffice Calc process cannot be used."""


@dataclass(frozen=True)
class CalcEngine:
    name: str
    source: str
    command_prefix: tuple[str, ...]
    display: str
    # KARE's overlay upperdir is its host-side backing store. Paths passed to
    # ``kare run`` must use the corresponding in-container namespace path.
    kare_upper_root: Path | None = None


@dataclass(frozen=True)
class LibreOfficeCalculationResult:
    workbook_path: Path
    engine_display: str
    elapsed_seconds: float


@dataclass(frozen=True)
class LibreOfficeBatchCalculationResult:
    workbook_paths: tuple[Path, ...]
    engine_display: str
    elapsed_seconds: float
    batch_count: int


def _executable(path: Path) -> Path | None:
    try:
        return path if path.is_file() and os.access(path, os.X_OK) else None
    except OSError:
        return None


def _direct_candidates() -> list[tuple[Path, str]]:
    candidates: list[tuple[Path, str]] = []
    configured = os.environ.get("BASE_AUDIT_SOFFICE", "").strip()
    if configured:
        candidates.append((Path(configured).expanduser(), "环境变量 BASE_AUDIT_SOFFICE"))
    configured_dir = os.environ.get("BASE_AUDIT_LIBREOFFICE_DIR", "").strip()
    if configured_dir:
        root = Path(configured_dir).expanduser()
        candidates.extend((root / name, "环境变量 BASE_AUDIT_LIBREOFFICE_DIR") for name in (
            "soffice", "libreoffice", "scalc", "program/soffice",
        ))

    # TDF DEB bundles install each major.minor release into its own /opt
    # directory and intentionally do not register a global `soffice` command.
    # Prefer the numerically newest complete versioned install so the engine
    # used by audits agrees with the version the installer has just verified.
    versioned: list[tuple[tuple[int, ...], Path]] = []
    for pattern in ("libreoffice*/program/soffice", "LibreOffice*/program/soffice"):
        try:
            paths = Path("/opt").glob(pattern)
        except OSError:
            continue
        for candidate in paths:
            match = re.fullmatch(r"(?:libreoffice|LibreOffice)([0-9]+\.[0-9]+)", candidate.parent.parent.name)
            if match:
                versioned.append((tuple(int(part) for part in match.group(1).split(".")), candidate))
    seen_versioned: set[str] = set()
    for _version, candidate in sorted(versioned, key=lambda item: item[0], reverse=True):
        key = str(candidate)
        if key not in seen_versioned:
            candidates.append((candidate, "版本化 DEB 安装目录（优先最新版本）"))
            seen_versioned.add(key)

    for command in ("soffice", "libreoffice", "scalc"):
        located = shutil.which(command)
        if located:
            candidates.append((Path(located), "PATH"))
    candidates.extend((Path(item), "系统路径") for item in (
        "/usr/bin/soffice", "/usr/bin/libreoffice",
        "/usr/lib/libreoffice/program/soffice", "/usr/lib64/libreoffice/program/soffice",
    ))
    home = Path.home()
    for root in (home / "Desktop", home / "Downloads", Path("/opt")):
        if not root.is_dir():
            continue
        for pattern in ("*/program/soffice", "libreoffice*/program/soffice", "LibreOffice*/program/soffice"):
            try:
                candidates.extend((candidate, "手动安装目录") for candidate in root.glob(pattern))
            except OSError:
                continue
    return candidates


def _find_direct_engine() -> CalcEngine | None:
    for candidate, source in _direct_candidates():
        executable = _executable(candidate)
        if executable is not None:
            return CalcEngine("LibreOffice Calc", source, (str(executable),), str(executable))
    return None


def _find_linglong_engine() -> CalcEngine | None:
    ll_cli = shutil.which("ll-cli")
    if not ll_cli:
        return None
    app_id = os.environ.get("BASE_AUDIT_LINGLONG_APP", LINGLONG_APP_DEFAULT).strip() or LINGLONG_APP_DEFAULT
    try:
        listed = subprocess.run(
            [ll_cli, "list"], stdout=subprocess.PIPE, stderr=subprocess.PIPE,
            text=True, timeout=15, check=False, env=_system_command_env(),
        )
    except (OSError, subprocess.TimeoutExpired):
        return None
    if listed.returncode != 0 or app_id not in (listed.stdout or "") + "\n" + (listed.stderr or ""):
        return None
    prefix = (ll_cli, "run", app_id, "--", "soffice")
    return CalcEngine("LibreOffice Calc（玲珑商店版）", "玲珑商店版", prefix, "ll-cli run {} -- soffice".format(app_id))


def _load_kare_config() -> dict[str, str]:
    data_dir = os.environ.get("BASE_AUDIT_DATA_DIR", "").strip()
    if not data_dir:
        return {}
    path = Path(data_dir).expanduser() / KARE_CONFIG_NAME
    try:
        if path.is_symlink() or not path.is_file():
            return {}
        value = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, ValueError, TypeError):
        return {}
    if not isinstance(value, dict):
        return {}
    environment = value.get("environment")
    app_id = value.get("appId")
    upper_root = value.get("upperRoot")
    if not all(isinstance(item, str) and item for item in (environment, app_id, upper_root)):
        return {}
    if any(character in environment + app_id for character in "\x00\n\r"):
        return {}
    return {"environment": environment, "appId": app_id, "upperRoot": upper_root}


def _find_kare_engine() -> CalcEngine | None:
    kare = shutil.which("kare")
    config = _load_kare_config()
    app_id = os.environ.get(KARE_APP_ENV, "").strip() or config.get("appId", "")
    if not kare or not app_id:
        return None

    configured_root = os.environ.get(KARE_UPPER_ENV, "").strip() or config.get("upperRoot", "")
    if configured_root:
        upper_root = Path(configured_root).expanduser().resolve()
    else:
        # A single installed KARE environment can be discovered. Multiple
        # environments require an explicit selection to avoid writing into the
        # wrong overlay layer.
        roots = sorted(Path("/opt/kare-applications").glob("*/upper"))
        if len(roots) != 1:
            return None
        upper_root = roots[0].resolve()
    if not upper_root.is_dir():
        return None

    prefix = (kare, "run", app_id)
    return CalcEngine(
        "LibreOffice Calc（KARE）", "银河麒麟 KARE", prefix,
        "kare run {}".format(app_id), upper_root,
    )


def find_calc_engine() -> CalcEngine | None:
    """Find a direct, Linglong, or explicitly configured KARE Calc entry."""
    mode = os.environ.get("BASE_AUDIT_ENGINE", "auto").strip().casefold() or "auto"
    if mode not in {"auto", "direct", "ll-cli", "kare"}:
        mode = "auto"
    if mode == "kare":
        return _find_kare_engine()
    if mode in {"auto", "direct"}:
        # An explicit KARE config opts into that engine even if a separate host
        # LibreOffice is available. The built-in installer does not create it.
        if mode == "auto" and _load_kare_config():
            kare = _find_kare_engine()
            if kare is not None:
                return kare
        direct = _find_direct_engine()
        if direct is not None:
            return direct
    if mode == "auto":
        kare = _find_kare_engine()
        if kare is not None:
            return kare
    if mode in {"auto", "ll-cli"}:
        return _find_linglong_engine()
    return None


def find_soffice() -> Path | None:
    """Backward-compatible direct-entry lookup; never returns Linglong."""
    engine = _find_direct_engine()
    return Path(engine.command_prefix[0]) if engine is not None else None


def _restore_recalc_flags(workbook_path: Path) -> None:
    """soffice 的 --convert-to 整簿回写会丢弃 calcPr 标记，此处补置。

    ``fullCalcOnLoad`` 与 ``forceFullCalc`` 由公式复制阶段设置，用于确保用户
    打开审核副本时强制重算（UOS 验收 #4 发现丢失）。

    实现上**只改 xl/workbook.xml 里 calcPr 的属性，其余 zip 条目原样复制**：
    绝不能用 openpyxl 打开再保存——那会重写工作表 XML，破坏 soffice 刚算出的
    公式缓存值，而后续“问题提取”正是读这些缓存。
    """
    import re
    import shutil
    import zipfile

    target = Path(workbook_path)
    temporary = target.with_name(target.stem + ".calcflags.tmp.xlsx")
    try:
        with zipfile.ZipFile(target, "r") as source, \
                zipfile.ZipFile(temporary, "w", zipfile.ZIP_DEFLATED) as sink:
            for item in source.infolist():
                data = source.read(item.filename)
                if item.filename == "xl/workbook.xml":
                    text = data.decode("utf-8")
                    if "<calcPr" in text:
                        # 已有 calcPr：补/覆盖两个属性，保留其它属性。
                        def _fix(match: "re.Match[str]") -> str:
                            tag = match.group(0)
                            self_closing = tag.rstrip().endswith("/>")
                            body = tag.rstrip()
                            body = body[:-2] if self_closing else body[:-1]
                            body = body.rstrip()
                            for name, value in (("fullCalcOnLoad", "1"), ("forceFullCalc", "1")):
                                if re.search(rf'\b{name}="[^"]*"', body):
                                    body = re.sub(rf'\b{name}="[^"]*"', f'{name}="{value}"', body)
                                else:
                                    body += f' {name}="{value}"'
                            return body + ("/>" if self_closing else ">")

                        text = re.sub(r"<calcPr\b[^>]*/?>", _fix, text, count=1)
                    elif "</workbook>" in text:
                        # 完全没有 calcPr：插入一个。
                        text = text.replace(
                            "</workbook>",
                            '<calcPr calcId="0" fullCalcOnLoad="1" forceFullCalc="1"/></workbook>',
                        )
                    data = text.encode("utf-8")
                sink.writestr(item, data)
        shutil.move(str(temporary), str(target))
    except Exception:
        # 标记只是“打开时重算”的提示，不是计算结果；任何失败都不应阻断审核。
        try:
            temporary.unlink(missing_ok=True)
        except OSError:
            pass


class LibreOfficeCalculator:
    """Recalculate an xlsx audit copy with a private LibreOffice profile."""

    engine_name = "LibreOffice Calc"

    def __init__(self, executable: Path | None = None, *, timeout_seconds: int = 300) -> None:
        self.executable = executable
        self.timeout_seconds = timeout_seconds

    def require_available(self) -> CalcEngine:
        if self.executable is not None:
            executable = _executable(self.executable)
            if executable is not None:
                return CalcEngine("LibreOffice Calc", "指定入口", (str(executable),), str(executable))
        engine = find_calc_engine()
        if engine is None:
            raise LibreOfficeUnavailableError(
                "未找到 LibreOffice Calc。可安装 LibreOffice、设置 BASE_AUDIT_SOFFICE 指向 soffice、"
                "设置 BASE_AUDIT_LIBREOFFICE_DIR 指向安装目录，或确认应用商店已安装 LibreOffice。"
                "麒麟 KARE 环境需设置 BASE_AUDIT_ENGINE=kare、BASE_AUDIT_KARE_APP_ID，"
                "并在无法唯一识别兼容层时设置 BASE_AUDIT_KARE_UPPER_ROOT。"
            )
        return engine

    def version(self, engine: CalcEngine | None = None) -> str:
        """Return the installed Calc version without opening a workbook."""
        selected = engine or self.require_available()
        try:
            completed = subprocess.run(
                [*selected.command_prefix, "--version"],
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                text=True,
                timeout=15,
                check=False,
                env=_system_command_env(),
            )
        except (OSError, subprocess.TimeoutExpired):
            return "版本未识别"
        text = "\n".join(item for item in (completed.stdout, completed.stderr) if item).strip()
        return text.splitlines()[0].strip() if completed.returncode == 0 and text else "版本未识别"

    def recalculate(self, workbook_path: Path) -> LibreOfficeCalculationResult:
        """Recalculate one workbook through the same isolated batch path."""
        batch = self.recalculate_many((workbook_path,), batch_size=1)
        return LibreOfficeCalculationResult(
            batch.workbook_paths[0], batch.engine_display, batch.elapsed_seconds,
        )

    def recalculate_many(
        self,
        workbook_paths: Iterable[Path],
        *,
        batch_size: int = 8,
    ) -> LibreOfficeBatchCalculationResult:
        """Recalculate XLSX files in bounded LibreOffice invocations.

        LibreOffice startup has a material fixed cost. A private profile and
        temporary staging area are shared by each bounded chunk, then every
        generated workbook is copied back only after that chunk succeeds.
        ``batch_size`` bounds temporary disk use for large audit runs.
        """
        engine = self.require_available()
        paths = tuple(Path(path).resolve() for path in workbook_paths)
        if not paths:
            raise ValueError("至少需要一个待计算工作簿")
        if len(set(paths)) != len(paths):
            raise ValueError("批量计算列表中存在重复工作簿路径")
        for workbook_path in paths:
            if not workbook_path.is_file():
                raise FileNotFoundError("待计算工作簿不存在：{}".format(workbook_path))
            if workbook_path.suffix.casefold() != ".xlsx":
                raise ValueError("LibreOffice 计算引擎当前仅支持 .xlsx 审核副本")
        try:
            chunk_size = int(batch_size)
        except (TypeError, ValueError) as exc:
            raise ValueError("LibreOffice 批量大小必须是正整数") from exc
        if chunk_size < 1:
            raise ValueError("LibreOffice 批量大小必须是正整数")

        started = time.monotonic()
        batch_count = 0
        for offset in range(0, len(paths), chunk_size):
            chunk = paths[offset:offset + chunk_size]
            self._recalculate_chunk(engine, chunk, batch_index=batch_count + 1)
            batch_count += 1
        return LibreOfficeBatchCalculationResult(
            workbook_paths=paths,
            engine_display=engine.display,
            elapsed_seconds=time.monotonic() - started,
            batch_count=batch_count,
        )

    def _recalculate_chunk(
        self,
        engine: CalcEngine,
        workbook_paths: tuple[Path, ...],
        *,
        batch_index: int,
    ) -> None:
        # 应用商店的玲珑版 LibreOffice 通常无法读取 /tmp 中的宿主文件。
        # 将临时副本放到用户 HOME 的缓存目录，采用 ASCII 文件名和 file URI，
        # 同时规避中文路径/文件名在 soffice CLI 中被错误解析为 source 的问题。
        namespace_cache_root = Path(os.environ.get("XDG_CACHE_HOME", "")).expanduser() if os.environ.get("XDG_CACHE_HOME") else Path.home() / ".cache"
        if not namespace_cache_root.is_absolute():
            namespace_cache_root = Path.home() / namespace_cache_root
        if engine.kare_upper_root is not None:
            try:
                cache_root = engine.kare_upper_root / namespace_cache_root.relative_to("/") / "base_audit_lo"
            except ValueError as exc:
                raise LibreOfficeUnavailableError(
                    "KARE 计算缓存必须位于 KARE 根文件系统可映射的绝对路径下。"
                ) from exc
        else:
            cache_root = namespace_cache_root / "base_audit_lo"
        try:
            cache_root.mkdir(parents=True, exist_ok=True)
        except OSError:
            # 开发/受限环境可没有 HOME 缓存权限；统信玲珑版仍优先走上面的
            # HOME 缓存路径，只有创建失败时才退回系统临时目录。
            cache_root = Path(tempfile.gettempdir()) / "base_audit_lo"
            cache_root.mkdir(parents=True, exist_ok=True)
        try:
            temporary_run = tempfile.TemporaryDirectory(prefix="run_", dir=cache_root)
        except OSError:
            # 某些 UOS 加固环境允许看到 HOME/.cache 却不允许在其中创建
            # 私有临时目录。此时保留隔离 profile 的前提下回退系统临时目录。
            fallback_root = Path(tempfile.gettempdir()) / "base_audit_lo"
            fallback_root.mkdir(parents=True, exist_ok=True)
            temporary_run = tempfile.TemporaryDirectory(prefix="run_", dir=fallback_root)
        with temporary_run as temp_text:
            temp_dir = Path(temp_text)
            source_dir, output_dir, profile_dir = temp_dir / "source", temp_dir / "output", temp_dir / "profile"
            source_dir.mkdir()
            output_dir.mkdir()
            profile_dir.mkdir()
            staged_paths = []
            for index, workbook_path in enumerate(workbook_paths, start=1):
                staged = source_dir / "audit_{:05d}.xlsx".format(index)
                shutil.copy2(workbook_path, staged)
                staged_paths.append(staged)
            if engine.kare_upper_root is not None:
                def _kare_namespace_path(host_path: Path) -> PurePosixPath:
                    try:
                        relative = host_path.resolve().relative_to(engine.kare_upper_root)
                    except ValueError as exc:
                        raise LibreOfficeUnavailableError(
                            "KARE 临时文件没有落在已配置的兼容层目录中，无法安全传递给计算引擎。"
                        ) from exc
                    return PurePosixPath("/") / PurePosixPath(relative.as_posix())

                guest_profile = _kare_namespace_path(profile_dir)
                guest_output = str(_kare_namespace_path(output_dir))
                profile_uri = "file://" + quote(guest_profile.as_posix(), safe="/")
                source_uris = [
                    "file://" + quote(_kare_namespace_path(staged).as_posix(), safe="/")
                    for staged in staged_paths
                ]
            else:
                guest_output = str(output_dir)
                profile_uri = profile_dir.as_uri()
                source_uris = [staged.as_uri() for staged in staged_paths]
            command = [
                *engine.command_prefix, "--headless", "--nologo", "--nodefault", "--nolockcheck",
                "--nofirststartwizard", "-env:UserInstallation=" + profile_uri,
                "--convert-to", "xlsx:Calc MS Excel 2007 XML", "--outdir", guest_output, *source_uris,
            ]
            try:
                completed = subprocess.run(
                    command, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True,
                    # Preserve the old per-file timeout budget when one process
                    # handles several files.
                    timeout=self.timeout_seconds * len(workbook_paths),
                    check=False, env=_system_command_env(),
                )
            except subprocess.TimeoutExpired as exc:
                names = "、".join(path.name for path in workbook_paths[:5])
                if len(workbook_paths) > 5:
                    names += "等 {} 份工作簿".format(len(workbook_paths))
                raise TimeoutError(
                    "LibreOffice 公式计算批次 {} 超时（预算 {} 秒）：{}".format(
                        batch_index, self.timeout_seconds * len(workbook_paths), names,
                    )
                ) from exc
            generated_paths = [output_dir / staged.name for staged in staged_paths]
            missing = [
                workbook_paths[index].name
                for index, generated in enumerate(generated_paths)
                if not generated.is_file()
            ]
            if completed.returncode != 0 or missing:
                details = "\n".join(part for part in (completed.stderr, completed.stdout) if part).strip()
                hint = ""
                if "source could not be loaded" in details.casefold():
                    hint = "；已改用用户缓存目录暂存输入文件。若仍失败，请确认当前 LibreOffice/玲珑应用可访问 {}".format(cache_root)
                if len(workbook_paths) == 1:
                    name_details = "“{}”".format(workbook_paths[0].name)
                else:
                    name_details = "{} 份工作簿（{}）".format(
                        len(workbook_paths), "、".join(missing[:5]) or "批次输出未验证",
                    )
                raise LibreOfficeUnavailableError(
                    "LibreOffice 未能计算{}（{}）：{}{}".format(
                        name_details, engine.display, (details or "无输出")[:800], hint,
                    )
                )
            # Only publish this chunk after every output was produced. This
            # prevents a failed LO command from leaving a half-written chunk.
            for workbook_path, generated in zip(workbook_paths, generated_paths):
                shutil.copy2(generated, workbook_path)
                _restore_recalc_flags(workbook_path)
