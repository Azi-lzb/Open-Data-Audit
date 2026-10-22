from __future__ import annotations

from dataclasses import dataclass
from datetime import date
from pathlib import Path


def workbook(value, label, *, optional=False):
    text = str(value or "").strip()
    if optional and not text:
        return None
    path = Path(text).expanduser().resolve()
    if not text or not path.is_file() or path.suffix.lower() not in {".xlsx", ".xlsm"}:
        raise ValueError(f"{label}：请选择存在的 XLSX/XLSM 文件")
    return path


def inside(path, parent):
    try:
        path.relative_to(parent)
        return True
    except ValueError:
        return False


@dataclass(frozen=True)
class S5Request:
    config_path: Path | None
    template_path: Path
    files: tuple[Path, ...]
    input_dir: Path
    output_dir: Path
    period: str
    business: str
    external_path: Path | None = None
    external_period: str = ""
    previous_files: tuple[Path, ...] = ()
    previous_period: str = ""
    identity_sheet: str = ""
    identity_cell: str = ""

    @classmethod
    def from_dict(cls, data):
        period = str(data.get("period") or "").strip()
        try:
            date.fromisoformat(period)
        except ValueError as exc:
            raise ValueError("本期报告期：请填写 YYYY-MM-DD 日期") from exc
        raw_files = data.get("files") or []
        if not isinstance(raw_files, (list, tuple)):
            raise ValueError("本期文件清单必须为文件路径列表")
        files = tuple(dict.fromkeys(workbook(p, "本期源文件") for p in raw_files))
        if not files:
            raise ValueError("请至少勾选一份本期源文件")
        previous = data.get("previous_files") or []
        if not isinstance(previous, (list, tuple)):
            raise ValueError("上期文件清单必须为文件路径列表")
        previous = tuple(dict.fromkeys(workbook(p, "上期源文件") for p in previous))
        previous_period = str(data.get("previous_period") or "").strip()
        if previous:
            try:
                date.fromisoformat(previous_period)
            except ValueError as exc:
                raise ValueError("选择上期文件时，请填写有效上期报告期") from exc
            if previous_period >= period:
                raise ValueError("上期报告期必须早于本期")
        external_period = str(data.get("external_period") or period).strip()
        if external_period != period:
            raise ValueError("外部文件报告期与本期不一致，请核实后重新选择")
        output_text = str(data.get("output_dir") or "").strip()
        if not output_text:
            raise ValueError("请选择 S5 输出目录")
        input_dir = Path(data.get("input_dir") or files[0].parent).expanduser().resolve()
        if not input_dir.is_dir() or any(not inside(p, input_dir) for p in files):
            raise ValueError("勾选的文件必须位于 S5 源数据目录内")
        return cls(
            config_path=workbook(data.get("config_path"), "S5 系统配置"),
            template_path=workbook(data.get("template_path"), "布局模板"),
            files=files, input_dir=input_dir, output_dir=Path(output_text).expanduser().resolve(),
            period=period, business=str(data.get("business") or "").strip(),
            external_path=workbook(data.get("external_path"), "外部文件", optional=True),
            external_period=external_period, previous_files=previous, previous_period=previous_period,
            identity_sheet=str(data.get("identity_sheet") or "").strip(),
            identity_cell=str(data.get("identity_cell") or "").strip(),
        )
