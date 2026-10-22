"""交付版完整界面显示口令；系统导航不再设开发禁用开关。"""

from __future__ import annotations

DEVELOPER_SYSTEMS = ("S1", "S2", "S3", "S4", "S5")
DISTRIBUTION_PASSWORD = "HZPBC790"


def distribution_password_matches(password: object) -> bool:
    return isinstance(password, str) and password == DISTRIBUTION_PASSWORD
