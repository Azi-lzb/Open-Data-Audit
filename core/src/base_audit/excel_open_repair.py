"""Excel COM 打开失败的副本修复：全角括号表名裸引用 + 冻结窗格坐标矛盾。

背景（依据 129 份真实核查表根因分析，见
test-output/s1f1_mode_parity_20261107/excel_open_root_cause.md）：

1. 部分报送文件的公式直接引用含全角括号的表名（如
   ``个人贷款发生额信息（隐藏发放）!A1``）而没有加单引号。Excel 要求含特殊
   字符的表名必须加引号，存储的公式解析不了 → 拒绝打开整个工作簿；
   WPS 解析宽松能打开。
2. 部分文件的冻结窗格 ``<pane ySplit="6" topLeftCell="C6" state="frozen"/>``
   自相矛盾（topLeftCell 的行落在冻结区内），Excel 强制要求冻结态下
   topLeftCell 位于分割线右下方 → 拒绝打开。

修复只做两类无害规范化：给强制引号表名的裸引用补单引号（公式语义不变）；
冻结窗格 topLeftCell 行号改为 ySplit+1。**只允许作用于审核副本**，调用方
必须保证传入路径不是原始报送文件；本模块不做来源检查之外的文件写保护，
但绝不主动打开或改写用户原始数据。

无问题的文件不做任何重写（先廉价预检）；异常一律吞掉并返回未修复，
不打断审核流程——WPS 引擎本来就能打开这些文件，修复只是 Excel 引擎的
前置保障。
"""

from __future__ import annotations

from html import unescape as html_unescape
import re
import time
import zipfile
from pathlib import Path
from xml.etree import ElementTree as StdET

# Excel 强制引号字符：括号（全/半角）、空格（含全角）、连字符等。
# 纯中文/字母数字表名不需要引号，不在此列。
_NEEDS_QUOTE = re.compile(r"[（）()\s　\-\+\[\]{}\\/:;\"'<>,?~！！？＠＃￥％…&*=|]")
_SHEET_TAG = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}sheet"
_F_RE = re.compile(r"(<f(?:\s[^>]*)?>)(.*?)(</f>)", re.S)
_FORMULA_RE = re.compile(r"(<formula(?:\s[^>]*)?>)(.*?)(</formula>)", re.S)
_CHART_F_RE = re.compile(r"(<c:f>)(.*?)(</c:f>)")
_DEFINED_RE = re.compile(r"(<definedNames>)(.*?)(</definedNames>)", re.S)
_PANE_RE = re.compile(r"<pane[^>]*/>")
_EXTERNAL_REFS_RE = re.compile(r"<externalReferences>.*?</externalReferences>", re.S)
_EXTERNAL_REFS_EMPTY_RE = re.compile(r"<externalReferences\s*/>")
_EXTERNAL_REL_RE = re.compile(
    r"<Relationship\b[^>]*Type=\"[^\"]*relationships/externalLink\"[^>]*/>")
_EXTERNAL_CT_RE = re.compile(
    r"<Override\b[^>]*PartName=\"/xl/externalLinks/[^\"]*\"[^>]*/>")
_XML_PART_PREFIXES = ("xl/worksheets/", "xl/charts/", "xl/drawings/")


_SHEET_NAME_RE = re.compile(r'<sheet[^>]*name="([^"]*)"')


def _sheet_names(archive: zipfile.ZipFile) -> list[str]:
    """读取工作表名；严格解析失败时回退正则（结构瑕疵但表名可读的文件仍需修复）。"""
    try:
        root = StdET.fromstring(archive.read("xl/workbook.xml"))
    except (KeyError, StdET.ParseError):
        try:
            text = archive.read("xl/workbook.xml").decode("utf-8", "ignore")
        except KeyError:
            return []
        return [html_unescape(m) for m in _SHEET_NAME_RE.findall(text)]
    return [(s.get("name") or "") for s in root.iter(_SHEET_TAG)]


def _quote_bare_refs(text: str, special: list[str]) -> tuple[str, int]:
    total = 0
    for name in special:
        quoted_name = name.replace("'", "''")
        pattern = re.compile(r"(?<!')" + re.escape(name) + r"!")
        text, count = pattern.subn("'" + quoted_name + "'!", text)
        total += count
    return text, total


def _name_flex_regex(name: str) -> str:
    """表名的编码无关正则片段：每个非 ASCII 字符同时匹配字面、十进制与十六进制 NCR。

    openpyxl 保存时会把中文等字符写成 ``&#38598;`` 实体，而 WPS/Excel 生成的文件
    通常是字面 UTF-8。修复必须两种形态都能命中，否则经 openpyxl 处理的副本漏修。
    """
    parts = []
    for ch in name:
        if ord(ch) < 128:
            parts.append(re.escape(ch))
        else:
            parts.append(
                "(?:" + re.escape(ch)
                + "|&#" + str(ord(ch)) + ";"
                + "|&#[xX]" + format(ord(ch), "x") + ";"
                + "|&#[xX]" + format(ord(ch), "X") + ";)")
    return "".join(parts)


def _strip_dangling_external_refs(text: str, local_sheets: list[str]) -> tuple[str, int]:
    """剥离公式里指向“本地已导入同名工作表”的 [N] 外部簿前缀。

    OOXML 外部引用形如 ``[1]表名!A1``，[N] 对应 workbook.xml 的第 N 个
    ``externalReference``。审核副本把外部工作簿导入为本地 sheet 后，若
    externalLink 部件缺失或索引错位，[N] 变成悬空引用——Excel 直接拒绝打开。
    仅当 [N] 后的表名存在于本包 ``<sheets>``（本地已导入）时剥掉前缀
    （公式改指本地表，语义不变）；指向真正外部簿的不动。

    覆盖带引号 ``'[表]名'!``（内部 ' 按 '' 转义）与裸 ``表名!`` 两种变体，
    ���容表名的字面与 NCR 实体两种编码。返回 (新文本, 剥离次数)。
    """
    if not local_sheets:
        return text, 0
    total = 0
    for name in local_sheets:
        flex = _name_flex_regex(name)
        escaped = _name_flex_regex(name.replace("'", "''"))
        quoted = re.compile(r"\[\d+\]'" + escaped + r"'!")
        text, count = quoted.subn("'" + name.replace("'", "''") + "'!", text)
        total += count
        bare = re.compile(r"\[\d+\]" + flex + r"!")
        text, count = bare.subn(name + "!", text)
        total += count
    return text, total


def _dangling_ref_pattern(local_sheets: list[str]):
    """悬空 [N] 预检正则：[N] 后跟本地已存在的表名（引号或裸变体，字面或 NCR）。"""
    if not local_sheets:
        return None
    alts = "|".join(_name_flex_regex(name)
                    for name in sorted(local_sheets, key=len, reverse=True))
    if len(alts) > 60000:  # 表名总量异常时弃用预检（修复本身仍按名逐个处理）
        return None
    return re.compile(r"\[\d+\]'?(?:" + alts + r")'?!")


_ANY_INDEX_PREFIX_RE = re.compile(r"\[\d+\][^\s!<>()\[\]]+!")


def _has_dangling_external_ref(archive: zipfile.ZipFile, local_sheets: list[str]) -> bool:
    """全包是否仍有任何 [N]表名! 形式的外部簿引用（判定能否安全断开外链）。

    注意这里必须检查**所有** [N] 引用而不是只查指向本地表的：外链定义一旦删除，
    指向真外部簿的 [N] 也会变成悬空引用（它原本靠 externalLink 存活）。
    只要还有任何 [N]，就不得删除外链定义。
    """
    for part in archive.namelist():
        if not part.endswith(".xml"):
            continue
        try:
            text = archive.read(part).decode("utf-8", "ignore")
        except Exception:
            continue
        if _ANY_INDEX_PREFIX_RE.search(text):
            return True
    return False


def _drop_external_link_parts(parts: dict[str, bytes]) -> int:
    """彻底断开外链：清 externalReferences、删关系与 Content_Types 声明、剔除部件。

    调用方必须确认全包已无指向本地表的 [N] 引用（否则删链接会让残留 [N] 变悬空）。
    返回移除的关系条数。
    """
    removed = 0
    workbook = parts.get("xl/workbook.xml")
    if workbook is not None:
        text = workbook.decode("utf-8", "ignore")
        text = _EXTERNAL_REFS_RE.sub("", text)
        text = _EXTERNAL_REFS_EMPTY_RE.sub("", text)
        parts["xl/workbook.xml"] = text.encode("utf-8")
    rels = parts.get("xl/_rels/workbook.xml.rels")
    if rels is not None:
        text = rels.decode("utf-8", "ignore")
        text, count = _EXTERNAL_REL_RE.subn("", text)
        removed += count
        parts["xl/_rels/workbook.xml.rels"] = text.encode("utf-8")
    content_types = parts.get("[Content_Types].xml")
    if content_types is not None:
        text = content_types.decode("utf-8", "ignore")
        text = _EXTERNAL_CT_RE.sub("", text)
        parts["[Content_Types].xml"] = text.encode("utf-8")
    for part in [n for n in parts if n.startswith("xl/externalLinks/")]:
        parts.pop(part, None)
    return removed


def _needs_repair(archive: zipfile.ZipFile, special: list[str],
                  local_sheets: list[str], *, fix_quotes: bool, fix_panes: bool,
                  fix_external_refs: bool) -> bool:
    """廉价预检：无任何待修嫌疑的文件不重写。"""
    quote_patterns = ([re.compile(r"(?<!')" + re.escape(name) + r"!") for name in special]
                      if (fix_quotes and special) else [])
    dangling_pattern = _dangling_ref_pattern(local_sheets) if fix_external_refs else None
    for part in archive.namelist():
        if not (part.endswith(".xml") and
                (part.startswith(_XML_PART_PREFIXES) or part == "xl/workbook.xml")):
            continue
        try:
            text = archive.read(part).decode("utf-8", "ignore")
        except Exception:
            continue
        for pattern in quote_patterns:
            if pattern.search(text):
                return True
        if dangling_pattern and dangling_pattern.search(text):
            return True
        if fix_external_refs and part in ("xl/workbook.xml", "[Content_Types].xml",
                                          "xl/_rels/workbook.xml.rels"):
            if "<externalReferences>" in text or "externalLink" in text:
                return True
        if fix_panes and part.startswith("xl/worksheets/") and "<pane " in text:
            for pane in _PANE_RE.findall(text):
                ys = re.search(r'ySplit="(\d+)"', pane)
                tl = re.search(r'topLeftCell="([A-Z]+)(\d+)"', pane)
                frozen = 'state="frozen"' in pane or 'state="frozenSplit"' in pane
                if ys and tl and frozen and int(tl.group(2)) <= int(ys.group(1)):
                    return True
    return False


def repair_audit_copy_for_excel(path: Path, *, attempts: int = 3,
                                fix_quotes: bool = True, fix_panes: bool = True,
                                fix_external_refs: bool = True,
                                drop_external_links: bool = True) -> dict:
    """修复审核副本使其可被 Excel 打开；返回统计。

    只处理 .xlsx/.xlsm（ZIP 容器）；其余格式原样跳过。任何异常都返回
    ``changed=False`` 且保持文件原样（先写临时文件再原子替换）。临时文件
    可能被杀毒/索引器短暂锁定，重试若干次（默认 3 次，间隔 0.5 秒）。

    三类修复（可按需关闭）：
    - fix_quotes：含全角括号等特殊字符的表名裸引用补单引号；
    - fix_panes：冻结窗格 topLeftCell 落在冻结区时行号修正；
    - fix_external_refs：公式中指向“本地已导入同名工作表”的 [N] 外部簿前缀
      剥离（externalLink 缺失/错位时 [N] 是悬空引用，Excel 拒绝打开；
      WPS/LibreOffice 打开但重算回落断链缓存）；
    - drop_external_links：剥离后若全包再无指向本地表的 [N]，彻底断开外链
      （清 externalReferences、删关系与 Content_Types 声明、剔除 externalLinks 部件），
      使副本只含工作簿内引用。仍有残留 [N] 时不动外链，避免制造新的悬空引用。
    """
    result: dict = {"changed": False, "quoted": 0, "pane_fixed": 0,
                    "stripped_refs": 0, "dropped_links": 0, "reason": ""}
    temporary = None
    last_exc: Exception | None = None
    for attempt in range(max(1, attempts)):
        try:
            with zipfile.ZipFile(path) as archive:
                if "xl/workbook.xml" not in archive.namelist():
                    result["reason"] = "非 OOXML 工作簿"
                    return result
                special = [n for n in _sheet_names(archive) if n and _NEEDS_QUOTE.search(n)]
                local_sheets = [n for n in _sheet_names(archive) if n]
                if not _needs_repair(archive, special, local_sheets,
                                     fix_quotes=fix_quotes, fix_panes=fix_panes,
                                     fix_external_refs=fix_external_refs):
                    result["reason"] = "未发现需要修复的内容"
                    return result
                temporary = path.with_name(path.stem + ".excel-repair.tmp" + path.suffix)
                with zipfile.ZipFile(temporary, "w", zipfile.ZIP_DEFLATED) as zout:
                    for part in archive.namelist():
                        data = archive.read(part)
                        if part.endswith(".xml") and (
                            part.startswith(_XML_PART_PREFIXES) or part == "xl/workbook.xml"
                        ):
                            text = data.decode("utf-8", "ignore")

                            def sub_quotes(m):
                                seg, count = _quote_bare_refs(m.group(2), special)
                                result["quoted"] += count
                                return m.group(1) + seg + m.group(3)

                            def sub_strip(m):
                                seg, count = _strip_dangling_external_refs(
                                    m.group(2), local_sheets)
                                result["stripped_refs"] += count
                                return m.group(1) + seg + m.group(3)

                            if fix_quotes and special:
                                text = _F_RE.sub(sub_quotes, text)
                                text = _FORMULA_RE.sub(sub_quotes, text)
                                text = _CHART_F_RE.sub(sub_quotes, text)
                                text = _DEFINED_RE.sub(sub_quotes, text)

                            if fix_external_refs:
                                text = _F_RE.sub(sub_strip, text)
                                text = _FORMULA_RE.sub(sub_strip, text)
                                text = _CHART_F_RE.sub(sub_strip, text)
                                text = _DEFINED_RE.sub(sub_strip, text)

                            def fix_pane(m):
                                tag = m.group(0)
                                ys = re.search(r'ySplit="(\d+)"', tag)
                                tl = re.search(r'topLeftCell="([A-Z]+)(\d+)"', tag)
                                frozen = ('state="frozen"' in tag
                                          or 'state="frozenSplit"' in tag)
                                if (fix_panes and ys and tl and frozen
                                        and int(tl.group(2)) <= int(ys.group(1))):
                                    result["pane_fixed"] += 1
                                    return tag.replace(
                                        tl.group(0),
                                        f'topLeftCell="{tl.group(1)}{int(ys.group(1)) + 1}"')
                                return tag

                            text = _PANE_RE.sub(fix_pane, text)
                            data = text.encode("utf-8")
                        zout.writestr(part, data)
                if fix_external_refs and drop_external_links:
                    # 剥离后若全包再无指向本地表的 [N]，则可安全彻底断开外链：
                    # 清 externalReferences、删关系/Content_Types、剔除 externalLinks 部件，
                    # 使副本只含工作簿内引用（Excel/WPS/LibreOffice 均按本地表重算）。
                    rebuilt = {}
                    with zipfile.ZipFile(temporary) as check:
                        still_dangling = _has_dangling_external_ref(check, local_sheets)
                        if not still_dangling:
                            try:
                                rebuilt = {n: check.read(n) for n in check.namelist()}
                            except (KeyError, zipfile.BadZipFile):
                                # 包内 rela 缺失等结构瑕疵：放弃断开外链，保留原样
                                rebuilt = {}
                                still_dangling = True
                    if not still_dangling:
                        result["dropped_links"] = _drop_external_link_parts(rebuilt)
                        with zipfile.ZipFile(temporary, "w", zipfile.ZIP_DEFLATED) as zout:
                            for name, payload in rebuilt.items():
                                zout.writestr(name, payload)
            temporary.replace(path)
            result["changed"] = True
            return result
        except Exception as exc:  # 瞬时锁/IO：重试；最终失败保持文件原样
            last_exc = exc
            result["quoted"] = 0
            result["pane_fixed"] = 0
            result["stripped_refs"] = 0
            result["dropped_links"] = 0
            result["reason"] = f"{type(exc).__name__}: {exc}"
            try:
                if temporary is not None:
                    temporary.unlink(missing_ok=True)
            except OSError:
                pass
            if attempt + 1 < max(1, attempts):
                time.sleep(0.5)
    if last_exc is not None:
        result["reason"] = f"{type(last_exc).__name__}: {last_exc}"
    return result
