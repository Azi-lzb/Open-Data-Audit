"""S5 input discovery. Ambiguous inputs stay explicit instead of picking a first item."""
from __future__ import annotations

import calendar
import re
from datetime import date, datetime
from pathlib import Path

from base_audit.discovery import (
    TemplateCatalog, _all_source_sheets_in_template, _sheet_score, read_xlsx_sheet_names,
)
from s1audit_configured.source_reader import read_book, shared_readers
from s1audit_configured.workbook_config import load_configuration


_DAY = re.compile(r"(?<!\d)(20\d{2})(?:[-._年/](\d{1,2})[-._月/](\d{1,2})日?|(\d{2})(\d{2}))(?!\d)")
_MONTH = re.compile(r"(?<!\d)(20\d{2})(?:[-._年/](\d{1,2})月?|(\d{2}))(?!\d)")
_PERIOD_LABELS = {"数据日期", "报告日期", "报告期", "统计日期", "数据期", "数据截止日期", "报表日期"}


def periods_in(value):
    if isinstance(value, (date, datetime)):
        return {value.strftime("%Y-%m-%d")}
    text = str(value or "")
    if text.startswith("="):
        return set()
    found = set()
    for match in _DAY.finditer(text):
        try:
            year, month, day, compact_month, compact_day = match.groups()
            found.add(date(int(year), int(month or compact_month), int(day or compact_day)).isoformat())
        except ValueError:
            raise ValueError(f"报告期日期不正确：{match.group(0)}")
    text = _DAY.sub("", text)
    for match in _MONTH.finditer(text):
        year, month, compact_month = match.groups()
        year, month = int(year), int(month or compact_month)
        found.add(date(year, month, calendar.monthrange(year, month)[1]).isoformat())
    return found


def _period_fields(path):
    book = read_book(path, data_only=False)
    found = set()
    try:
        for sheet in book.worksheets:
            for row in range(1, min(sheet.max_row, 60) + 1):
                for col in range(1, min(sheet.max_column, 15) + 1):
                    text = str(sheet.cell(row, col).value or "").strip()
                    label = re.split(r"[:：]", text, maxsplit=1)[0].strip()
                    if label not in _PERIOD_LABELS:
                        continue
                    found.update(periods_in(text))
                    for r, c in [(row, col + 1), (row, col + 2), (row + 1, col)]:
                        found.update(periods_in(sheet.cell(r, c).value))
    finally:
        book.close()
    return found


def detect_s5_period(files, input_dir):
    values = set()
    for path in files:
        values.update(periods_in(Path(path).stem))
    source = "源文件名"
    if not values:
        for path in files:
            values.update(_period_fields(Path(path)))
        source = "源表报告期字段"
    if not values:
        folder = Path(input_dir)
        for parent in (folder, *folder.parents):
            values = periods_in(parent.name)
            if values:
                source = f"文件夹“{parent.name}”"
                if not _DAY.search(parent.name):
                    source += "（仅年月时按月末）"
                break
    if len(values) > 1:
        raise ValueError("源文件存在多个报告期：" + "、".join(sorted(values)) + "；请按报告期分目录或调整勾选文件")
    if not values:
        raise ValueError("未识别到报告期；请在源文件或所属文件夹名称中标明 YYYY-MM-DD 或 YYYYMM")
    return next(iter(values)), source


def file_period_hint(path, input_dir=""):
    """Fast file-list date hint: use names and parent folders, never workbook values."""
    path = Path(path)
    try:
        values = periods_in(path.stem)
    except ValueError:
        return "日期格式异常"
    if len(values) == 1:
        return next(iter(values))
    if len(values) > 1:
        return "多个日期"
    try:
        current = path.parent.resolve()
    except (OSError, RuntimeError):
        current = path.parent
    while True:
        try:
            values = periods_in(current.name)
        except ValueError:
            return "日期格式异常"
        if len(values) == 1:
            return next(iter(values))
        if len(values) > 1:
            return "多个日期"
        if current.parent == current:
            break
        current = current.parent
    return ""


def match_s5_template(files, directory, index_path):
    if not directory or not Path(directory).is_dir():
        raise ValueError("请选择存在的模板目录，或使用“更换”手动指定模板")
    profiles = TemplateCatalog(Path(directory), Path(index_path)).profiles()
    if not profiles:
        raise ValueError("模板目录中没有可用模板，请选择正确的模板目录或手动更换匹配模板")
    chosen = []
    for path in files:
        sheets = read_xlsx_sheet_names(Path(path))
        ranked = sorted(((_sheet_score(profile.sheet_names, sheets), profile.path) for profile in profiles), reverse=True)
        covered = bool(ranked) and _all_source_sheets_in_template(
            read_xlsx_sheet_names(ranked[0][1]), sheets
        )
        if not ranked or (ranked[0][0] < 0.9 and not covered):
            raise ValueError(f"{Path(path).name}：没有匹配到模板，请核实模板目录或手动更换模板")
        if len(ranked) > 1 and ranked[0][0] - ranked[1][0] < 0.05:
            raise ValueError(f"{Path(path).name}：有多个相近模板，请使用“更换”明确指定模板")
        chosen.append(ranked[0][1])
    if len(set(chosen)) != 1:
        raise ValueError("勾选文件匹配到多种业务模板，请一次只保留一个业务的文件")
    return chosen[0]


def match_s5_business(configuration, files, template_path):
    sheet_sets = [set(read_xlsx_sheet_names(Path(path))) for path in files]
    businesses = sorted({rule.business for rule in configuration.rules})
    candidates = []
    for business in businesses:
        sheets = {rule.default_sheet for rule in configuration.rules if rule.business == business and rule.default_sheet}
        sheets.update(metric.sheet for metric in configuration.metrics if metric.business == business and metric.sheet)
        if not sheets or all(sheets.issubset(source) for source in sheet_sets):
            candidates.append(business)
    if len(candidates) == 1:
        return candidates[0]
    text = Path(template_path).stem + " " + " ".join(str(Path(path).parent.name) + " " + Path(path).stem for path in files)
    named = [name for name in candidates if name and name in text]
    if named:
        longest = max(map(len, named))
        named = [name for name in named if len(name) == longest]
    if len(named) == 1:
        return named[0]
    if not candidates:
        raise ValueError("源文件工作表与 S5 配置中的业务不对应，请核实配置与匹配模板")
    raise ValueError("无法唯一识别业务：" + "、".join(candidates) + "；请完善配置的业务/工作表对应关系或在文件名中标明业务")


def classify_s5_files(files, *, template_dir="", template_path="", config_path="",
                      input_dir="", index_path=None, selected_files=(), business="", period=""):
    """Add S1-like per-file worksheet match, business and date hints for S5's list."""
    paths = [Path(path) for path in files]
    selected = {str(Path(path).resolve()) for path in selected_files}
    profiles = []
    chosen_path = Path(template_path).expanduser() if str(template_path or "").strip() else None
    if chosen_path and chosen_path.is_file():
        profiles = [(str(chosen_path.resolve()), read_xlsx_sheet_names(chosen_path))]
    elif template_dir and Path(template_dir).is_dir():
        index = Path(index_path) if index_path else Path(template_dir).parent / "模板索引.json"
        profiles = [(item.path, item.sheet_names)
                    for item in TemplateCatalog(Path(template_dir), index).profiles()]

    configuration = None
    if config_path and Path(config_path).is_file():
        try:
            configuration = load_configuration(config_path)
        except Exception:
            configuration = None

    rows = []
    for path in paths:
        resolved = str(path.resolve())
        try:
            source_sheets = set(read_xlsx_sheet_names(path))
        except Exception:
            source_sheets = set()
        ranked = sorted(
            ((_sheet_score(sheets, tuple(source_sheets)), profile_path, sheets)
             for profile_path, sheets in profiles),
            key=lambda item: item[0], reverse=True,
        )
        profile_path, template_sheets = "", set()
        ambiguous = False
        if ranked:
            score, candidate_path, candidate_sheets = ranked[0]
            gap = score - ranked[1][0] if len(ranked) > 1 else score
            ambiguous = len(ranked) > 1 and gap < 0.05
            # A selected template is authoritative. Without one, only use a
            # clear candidate; a non-zero overlap can still be shown as partial.
            covered = _all_source_sheets_in_template(candidate_sheets, source_sheets)
            if chosen_path or not ambiguous and (score >= 0.90 or covered):
                profile_path, template_sheets = candidate_path, set(candidate_sheets)
            elif not ambiguous and score > 0:
                profile_path, template_sheets = candidate_path, set(candidate_sheets)
            else:
                template_sheets = set(candidate_sheets)

        expected = template_sheets - {"审核规则"}
        source_data_sheets = source_sheets - {"审核规则", "兜底模板"}
        expected -= {"兜底模板"}
        common = expected & source_data_sheets
        missing = expected - source_data_sheets
        extra = source_data_sheets - expected
        if not source_sheets or not common:
            match_status = "unmatched"
        elif ambiguous or not profile_path or extra:
            match_status = "partial"
        else:
            match_status = "matched"

        row_business = ""
        if resolved in selected and business:
            row_business = str(business)
        elif configuration is not None:
            try:
                row_business = match_s5_business(configuration, [path], profile_path)
            except (ValueError, OSError):
                pass
        row_period = file_period_hint(path, input_dir)
        if not row_period and resolved in selected:
            row_period = str(period or "")
        detail = (
            f"模板：{Path(profile_path).name}" if profile_path else
            f"候选模板：{Path(ranked[0][1]).name}" if ranked and ranked[0][0] > 0 else "尚未识别模板"
        )
        if source_data_sheets:
            detail += f"；源工作表匹配 {len(common)}/{len(source_data_sheets)} 张"
        if missing:
            detail += "；模板中其他工作表（源文件未含，不影响匹配）：" + "、".join(sorted(missing))
        if extra:
            detail += "；多出：" + "、".join(sorted(extra))
        rows.append({
            "path": resolved, "name": path.name,
            "template": Path(profile_path).name if profile_path else "未识别",
            "match_status": match_status,
            "business": row_business,
            "period": row_period,
            "match_detail": detail,
            "matched_sheets": len(common),
            "source_sheet_count": len(source_data_sheets),
            "template_sheet_count": len(expected),
        })
    return rows


@shared_readers
def recognize_s5(form, data_dir):
    files = [Path(path) for path in form.get("selected_files", [])]
    result = {"errors": {}, "period_source": "", "template_source": "", "business_source": ""}
    for key in ("template_path", "period", "business"):
        if not form.get({"template_path": "template_manual", "period": "period_manual", "business": "business_manual"}[key]):
            form[key] = ""
    if not files:
        return result
    if not form.get("template_path"):
        try:
            form["template_path"] = match_s5_template(files, form.get("template_dir") or "", Path(data_dir) / "s5" / "模板索引.json")
            result["template_source"] = "模板目录自动匹配"
        except (ValueError, OSError) as exc:
            result["errors"]["template"] = str(exc)
    else:
        result["template_source"] = "手动指定模板"
    if not form.get("period"):
        try:
            form["period"], result["period_source"] = detect_s5_period(files, form.get("input_dir") or files[0].parent)
        except (ValueError, OSError) as exc:
            result["errors"]["period"] = str(exc)
    else:
        result["period_source"] = "显式报告期参数"
    if not form.get("business"):
        try:
            configuration = load_configuration(form.get("config_path") or "")
            form["business"] = match_s5_business(configuration, files, form.get("template_path") or "")
            result["business_source"] = "配置与源文件工作表"
        except Exception as exc:
            result["errors"]["business"] = str(exc)
    return result
