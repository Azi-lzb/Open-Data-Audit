"""大集中统计系统：WorkbookRenderer 抽象层。

两套明确独立的 XLSX 渲染实现：

- ``OpenPyxlRenderer``：标准模式（现有 openpyxl 路径的包装，基线）。
- ``FastOoxmlRenderer``：高速模式（ZIP + OOXML/XML 直接处理，见
  ``fast_ooxml.py``），面向固定模板 × 多机构批量生成。

业务层只通过 ``RendererFactory.get_renderer(mode, ...)`` 拿到渲染器并调用
``render_org(...)``，不感知具体实现；**高速模式失败时直接抛出
FastOoxmlRenderError，绝不自动回退到标准模式**。
"""

from __future__ import annotations

import time
from dataclasses import dataclass, field
from pathlib import Path

from .comparison_exporter import VBA_COLOR_PALETTE
from .form_template import output_file_name  # noqa: F401  （经工厂对外复用）

RENDERER_OPENPYXL = "OPENPYXL"
RENDERER_FAST_OOXML = "FAST_OOXML"


class FastOoxmlRenderError(RuntimeError):
    """高速渲染失败；带结构化定位信息，绝不触发 openpyxl 回退。"""

    def __init__(
        self,
        message: str,
        *,
        renderer: str = RENDERER_FAST_OOXML,
        stage: str = "",
        sheet: str = "",
        part: str = "",
    ) -> None:
        super().__init__(message)
        self.renderer = renderer
        self.stage = stage
        self.sheet = sheet
        self.part = part
        self.message = message

    def to_dict(self) -> dict[str, str]:
        return {
            "renderer": self.renderer, "stage": self.stage,
            "sheet": self.sheet, "part": self.part, "message": self.message,
        }


class UnsupportedFastRenderOperation(FastOoxmlRenderError):
    """业务调用了高速模式不支持的能力；明确暴露，不做回退。"""


@dataclass
class RendererCapabilities:
    """渲染器能力声明：不支持的 Operation 由调用前检查或渲染器内部抛错。"""

    supports_value_write: bool = True
    supports_formula_write: bool = False
    supports_style_index: bool = True
    supports_row_hidden: bool = True
    supports_sheet_delete: bool = True
    supports_named_range: bool = False
    supports_conditional_format_edit: bool = False
    supports_external_link_edit: bool = False


@dataclass
class OrgRenderResult:
    output_path: Path
    filled: int
    stats: dict = field(default_factory=dict)
    messages: list[str] = field(default_factory=list)


def _empty_stats(renderer: str, template_cache_hit: bool | None = None) -> dict:
    stats = {
        "renderer": renderer,
        "template_prepare_time": 0.0,
        "render_time": 0.0,
        "zip_write_time": 0.0,
        "validation_time": 0.0,
        "total_time": 0.0,
    }
    if template_cache_hit is not None:
        stats["template_cache_hit"] = template_cache_hit
    return stats


class OpenPyxlRenderer:
    """标准模式：包装现有 openpyxl 渲染路径（逻辑不变，仅加保存与统计）。

    xlsx_render_mode = OPENPYXL 时的结果必须与包装前完全一致。
    """

    mode = RENDERER_OPENPYXL
    capabilities = RendererCapabilities()

    def __init__(self, meta, *, alerts: list[dict],
                 hide_empty_rows: bool, delete_empty_sheets: bool) -> None:
        self.meta = meta
        self.alerts = alerts
        self.hide_empty_rows = hide_empty_rows
        self.delete_empty_sheets = delete_empty_sheets
        self._prepare_seconds = 0.0

    def render_org(
        self, *, org_name: str, region_name: str, record_date: str,
        data: dict[str, tuple], output_path: Path,
    ) -> OrgRenderResult:
        from .form_template import render_org_workbook

        stats = _empty_stats(self.mode)
        started = time.monotonic()
        book, filled = render_org_workbook(
            self.meta,
            org_name=org_name, region_name=region_name,
            record_date=record_date, data=data, alerts=self.alerts,
            hide_empty_rows=self.hide_empty_rows,
            delete_empty_sheets=self.delete_empty_sheets,
        )
        stats["render_time"] = time.monotonic() - started
        stats["template_prepare_time"] = self._prepare_seconds

        started = time.monotonic()
        output_path = Path(output_path)
        book.save(output_path)
        book.close()
        stats["zip_write_time"] = time.monotonic() - started
        stats["total_time"] = sum(v for k, v in stats.items() if k != "renderer")
        return OrgRenderResult(
            output_path=output_path, filled=filled, stats=stats,
            messages=[f"已生成：{output_path.name}（回填 {filled} 处，{stats['total_time']:.1f} 秒）"],
        )


def get_renderer(
    mode: str,
    *,
    template_path: Path,
    alerts: list[dict],
    hide_empty_rows: bool,
    delete_empty_sheets: bool,
):
    """兼容入口：等价于 build_renderer。"""
    return build_renderer(
        mode, template_path=template_path, alerts=alerts,
        hide_empty_rows=hide_empty_rows,
        delete_empty_sheets=delete_empty_sheets,
    )


class RendererFactory:
    """按 xlsx_render_mode 构建渲染器；并行场景下每个工作进程构建一次。"""

    def __init__(self, template_path: Path, *, alerts: list[dict],
                 hide_empty_rows: bool, delete_empty_sheets: bool) -> None:
        self.template_path = Path(template_path)
        self.alerts = alerts
        self.hide_empty_rows = hide_empty_rows
        self.delete_empty_sheets = delete_empty_sheets

    def get_renderer(self, mode: str):
        return build_renderer(
            mode, template_path=self.template_path, alerts=self.alerts,
            hide_empty_rows=self.hide_empty_rows,
            delete_empty_sheets=self.delete_empty_sheets,
        )

    # ---- 并行 worker 装载（renderers 模块级全局，进程内只建一次） ----

    def worker_init(self, mode: str, template_path: str, alerts: list[dict],
                    hide_empty_rows: bool, delete_empty_sheets: bool) -> None:
        global _WORKER_RENDERER
        _WORKER_RENDERER = build_renderer(
            mode, template_path=Path(template_path), alerts=alerts,
            hide_empty_rows=hide_empty_rows,
            delete_empty_sheets=delete_empty_sheets,
        )

    def worker_task(self, task: dict) -> dict:
        if _WORKER_RENDERER is None:
            raise FastOoxmlRenderError(
                "工作进程未初始化渲染器", stage="worker_init",
            )
        rendered = _WORKER_RENDERER.render_org(
            org_name=task["org_name"], region_name=task["region_name"],
            record_date=task["record_date"], data=task["data"],
            output_path=Path(task["output_path"]),
        )
        return {
            "path": str(rendered.output_path), "filled": rendered.filled,
            "name": Path(rendered.output_path).name,
            "stats": rendered.stats, "messages": rendered.messages,
        }


_WORKER_RENDERER = None


def build_renderer(mode: str, *, template_path: Path, alerts: list[dict],
                   hide_empty_rows: bool, delete_empty_sheets: bool):
    """RendererFactory 入口：FAST_OOXML 编译失败直接抛错，不做 openpyxl 回退。"""
    mode = str(mode or "OPENPYXL").upper()
    if mode == RENDERER_FAST_OOXML:
        from .fast_ooxml import FastOoxmlRenderer

        return FastOoxmlRenderer(
            template_path, alerts=alerts,
            hide_empty_rows=hide_empty_rows,
            delete_empty_sheets=delete_empty_sheets,
        )
    if mode == RENDERER_OPENPYXL:
        from .form_template import TemplateMeta

        meta = TemplateMeta(template_path)
        return OpenPyxlRenderer(
            meta, alerts=alerts,
            hide_empty_rows=hide_empty_rows,
            delete_empty_sheets=delete_empty_sheets,
        )
    raise FastOoxmlRenderError(
        f"未知的 XLSX 渲染模式：{mode}", renderer=mode, stage="renderer_factory",
    )


def band_color_for(ratio: float, alerts: list[dict]) -> int:
    """警戒区间取色（与 openpyxl 渲染路径同一口径）。"""
    from .config import alert_band_for

    band = alert_band_for(ratio, alerts)
    if band is not None:
        try:
            return int(band.get("填充颜色") or 0)
        except (TypeError, ValueError):
            return 0
    return 0


def vba_fill_hex(color_index: int) -> str | None:
    hex_value = VBA_COLOR_PALETTE.get(int(color_index))
    return hex_value
