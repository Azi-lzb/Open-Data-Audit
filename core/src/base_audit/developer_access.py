"""本机开发导航开关；独立于用户设置和业务配置。"""

from __future__ import annotations

import json
from pathlib import Path


DEVELOPER_DEFAULTS = {"S1": True, "S2": True, "S3": False, "S4": False}
DEVELOPER_SYSTEMS = tuple(DEVELOPER_DEFAULTS)
DEVELOPER_PASSWORD = "HZPBC"


class DeveloperAccessStore:
    def __init__(self, path: Path) -> None:
        self.path = Path(path)
        self.message = ""
        try:
            self.path.parent.mkdir(parents=True, exist_ok=True)
            # 仅首次创建，不覆盖开发人员手工维护的文件。
            with self.path.open("x", encoding="utf-8") as stream:
                json.dump(DEVELOPER_DEFAULTS, stream, ensure_ascii=False, indent=2)
                stream.write("\n")
        except FileExistsError:
            pass
        except OSError:
            # 无写入权限仍可启动；load() 会说明读取问题并维持禁用。
            pass

    def load(self) -> dict[str, bool]:
        disabled = {system: False for system in DEVELOPER_SYSTEMS}
        self.message = ""
        try:
            payload = json.loads(self.path.read_text(encoding="utf-8-sig"))
        except (OSError, ValueError):
            self.message = "开发设置读取失败，已按禁用处理。请检查 data/开发设置.json 的格式或读取权限。"
            return disabled
        if not isinstance(payload, dict):
            self.message = "开发设置必须为 JSON 对象，已按禁用处理。请检查 data/开发设置.json。"
            return disabled
        if any(system in payload and not isinstance(payload[system], bool) for system in DEVELOPER_SYSTEMS):
            self.message = "开发设置中的 S1/S2/S3/S4 必须填写 true 或 false，非法值按禁用处理。"
        # 旧文件缺少 S1/S2 时按默认开放，不改写已有文件。
        return {system: payload.get(system, DEVELOPER_DEFAULTS[system]) is True for system in DEVELOPER_SYSTEMS}


def developer_password_matches(password: object) -> bool:
    return isinstance(password, str) and password == DEVELOPER_PASSWORD
