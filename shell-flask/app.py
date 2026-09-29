"""Flask 版基础数据审核工具（统一外壳：Windows + UOS/麒麟 共用一份）。

复用 core 的唯一业务核心（base_audit 后端 + frontend 前端）：
- 前端 index.html 不做修改，由本模块在 ``<head>`` 注入一个 ``bridge.api``
  兼容层，把 ``bridge.api.方法名(...)`` 代理为 ``POST /api/方法名``。
- 后端直接继承 ``base_audit.web_app.WebApi``；文件/目录选择可由高级设置切换
  Tk 对话框（Flask 首次默认）、浏览器内置目录浏览或桌面系统原生模式。
  窗口控制按钮为无操作（浏览器窗口由用户自己管理）。
- 平台差异只有“打开文件/目录”：Windows 用 os.startfile，UOS/麒麟 用
  xdg-open；计算引擎选项由 core/base_audit/engines 按操作系统提供
  （Windows：Excel/WPS COM；UOS：LibreOffice Calc）。
"""

from __future__ import annotations

import os
import queue
import subprocess
import sys
import threading
from pathlib import Path
from typing import Any

from flask import Flask, jsonify, request

# 冻结态（PyInstaller）__file__ 指向临时解包目录，须按 exe 位置定位 core。
ROOT = Path(sys.executable).resolve().parent if getattr(sys, "frozen", False) else Path(__file__).resolve().parent
# DEB 的冻结程序把可变数据放在用户目录；源码和冻结程序未指定时仍按项目布局定位。
configured_core = (
    os.environ.get("BASE_AUDIT_CORE_DIR", "").strip()
    or os.environ.get("BASE_AUDIT_CORE_ROOT", "").strip()
)
if configured_core:
    CORE = Path(configured_core).expanduser().resolve()
else:
    # 打包布局优先：exe 同级的 core/；开发布局：shell-flask 平级的 ../core。
    CORE = ROOT / "core"
    if not (CORE / "src" / "base_audit").is_dir():
        CORE = ROOT.parent / "core"
if str(CORE / "src") not in sys.path:
    sys.path.insert(0, str(CORE / "src"))

from base_audit.path_browser import browse_directory  # noqa: E402
from base_audit.web_app import WebApi  # noqa: E402


configured_project_root = os.environ.get("BASE_AUDIT_PROJECT_ROOT", "").strip()
PROJECT_ROOT = (
    Path(configured_project_root).expanduser().resolve()
    if configured_project_root else CORE
)
configured_frontend = os.environ.get("BASE_AUDIT_FRONTEND_DIR", "").strip()
FRONTEND_DIR = (
    Path(configured_frontend).expanduser().resolve()
    if configured_frontend else CORE / "frontend"
)

def _browse_dir(path: str = "", mode: str = "") -> dict[str, Any]:
    """使用 core 的受控路径浏览器，所有外壳得到相同的过滤/磁盘入口语义。"""
    return browse_directory(PROJECT_ROOT, path, mode)


def _open_with_default_app(path: Path) -> None:
    """用系统默认程序打开文件/目录：Windows 用 os.startfile，其余用 xdg-open。"""
    if sys.platform == "win32":
        os.startfile(str(path))  # noqa: S606 - 本机默认程序打开
    else:
        env = os.environ.copy()
        if "LD_LIBRARY_PATH_ORIG" in env:
            original = env.pop("LD_LIBRARY_PATH_ORIG")
            if original:
                env["LD_LIBRARY_PATH"] = original
            else:
                env.pop("LD_LIBRARY_PATH", None)
        subprocess.Popen(["xdg-open", str(path)], env=env)


class FlaskApi(WebApi):
    """Flask 外壳：首次默认使用 Tk，用户可在高级设置切换三种选择器。"""

    def __init__(self, project_root: Path) -> None:
        super().__init__(project_root, file_picker_default="Tk 对话框")
        # Flask 运行在普通浏览器中，没有可调用的宿主系统对话框。
        self.state["filePickerSystemNativeAvailable"] = False
        # Tk 有硬性线程约束：根窗口必须在创建它的线程上销毁，而 Flask 每个请求
        # 各占一个线程——实测第二次弹窗换线程重建 Tk 即触发 Tcl 跨线程销毁，
        # 进程直接非法指令退出（浏览器表现为 Failed to fetch）。因此起一个
        # 常驻专用线程持有唯一的、永不销毁的 Tk 根窗口，弹窗请求全部交给它。
        self._tk_queue: "queue.Queue[tuple] | None" = None
        self._tk_ready = threading.Event()
        self._tk_done = threading.Event()
        self._tk_result: tuple[str, ...] = ()
        self._tk_error: BaseException | None = None
        try:
            import tkinter  # noqa: F401 - 只探测目标 Python 是否带 Tcl/Tk
            tk_available = True
        except ImportError:
            tk_available = False
        if tk_available:
            worker = threading.Thread(target=self._tk_main, name="base-audit-tk", daemon=True)
            worker.start()
            self._tk_ready.wait(timeout=10)
            tk_available = self._tk_queue is not None
            if tk_available:
                import atexit
                atexit.register(self._tk_queue.put, None)
        self.state["filePickerTkAvailable"] = tk_available

        # 用户可能先运行过桌面外壳，用户设置中保留了“系统原生”。该模式在
        # Flask 中不存在，若不迁移，所有顶部“选择”都会返回 500 而没有对话框。
        selected_mode = str(self.state.get("filePickerMode") or "")
        if selected_mode == "系统原生":
            replacement = "Tk 对话框" if tk_available else "浏览器内置"
            self.state["filePickerMode"] = replacement
            self.settings.file_picker_mode = replacement
            self.settings_store.save(self.settings)
            self._log(f"Flask 已将不可用的文件选择方式切换为“{replacement}”")
        elif selected_mode == "Tk 对话框" and not tk_available:
            # Tk 仅在当前环境缺失（如源码模式未装 python3-tk）：本次会话退回
            # “浏览器内置”兜底，但不改写持久偏好——换到自带 Tk 的发行包、或
            # 装好 python3-tk 后，原生对话框自动恢复，无需用户重新设置。
            self.state["filePickerMode"] = "浏览器内置"
            self._log("当前 Python 未安装 Tcl/Tk，本次会话退回“浏览器内置”选择文件；"
                      "偏好仍保留“Tk 对话框”。")

    def browse_dir(self, path: str = "", mode: str = "") -> dict[str, Any]:
        """列出目录内容（浏览器端目录浏览的数据源；mode 过滤文件扩展名）。"""
        return _browse_dir(path, mode)

    def _tk_main(self) -> None:
        """Tk 专用线程：持有唯一根窗口并串行执行全部选择对话框。"""
        try:
            import tkinter as tk
            root = tk.Tk()
            root.withdraw()
            try:
                root.attributes("-topmost", True)
            except Exception:  # 个别平台不支持置顶，不影响弹窗
                pass
        except Exception:
            self._tk_ready.set()
            return
        self._tk_queue = queue.Queue()
        self._tk_ready.set()
        from tkinter import filedialog
        while True:
            job = self._tk_queue.get()
            if job is None:
                root.destroy()
                return
            folder, current, multiple, file_types = job
            try:
                if folder:
                    value = filedialog.askdirectory(
                        parent=root, initialdir=current, mustexist=True)
                    self._tk_result = (str(value),) if value else ()
                elif multiple:
                    self._tk_result = tuple(str(item) for item in filedialog.askopenfilenames(
                        parent=root, initialdir=current, filetypes=file_types,
                    ))
                else:
                    value = filedialog.askopenfilename(
                        parent=root, initialdir=current, filetypes=file_types)
                    self._tk_result = (str(value),) if value else ()
                self._tk_error = None
            except Exception as exc:
                self._tk_result = ()
                self._tk_error = exc
            self._tk_done.set()

    def _tk_dialog(self, *, folder: bool, current: str, multiple: bool = False,
                   file_types: tuple[tuple[str, str], ...] = ()) -> tuple[str, ...]:
        """把对话框请求转交 Tk 专用线程，等待其在属主线程上完成。"""
        if self._tk_queue is None:
            raise RuntimeError("Tk 图形组件不可用；请安装 python3-tk，或在高级设置改用“浏览器内置”")
        self._tk_done.clear()
        self._tk_queue.put((folder, current, multiple, tuple(file_types)))
        if not self._tk_done.wait(timeout=3600):
            raise RuntimeError("选择窗口长时间未响应")
        if self._tk_error is not None:
            raise RuntimeError(f"选择窗口异常：{self._tk_error}")
        return self._tk_result

    def _select_path(self, **kwargs: Any) -> tuple[str, ...]:
        if self.state.get("filePickerMode") == "系统原生":
            raise RuntimeError("系统原生选择器仅适用于桌面外壳；Flask 请改用“Tk 对话框”或“浏览器内置”")
        return super()._select_path(**kwargs)

    def _open_path(self, path: Path, what: str) -> None:
        if not path.exists():
            self.state["status"] = f"{what}不存在"
            self._log(f"未找到{what}：{path}")
            return
        try:
            _open_with_default_app(path)
        except Exception as exc:
            self.state["status"] = f"无法打开{what}"
            self._log(f"打开{what}失败：{exc}")
            return
        self._log(f"已打开{what}：{path}")

    def open_history_explanation(self) -> dict[str, Any]:
        self._open_path(self.history_path, "历史审核配置")
        return self.state

    def open_user_guide(self) -> dict[str, Any]:
        self._open_path(self.project_root / "基础数据审核工具使用说明.docx", "使用说明")
        return self.state

    def open_path(self, path: str) -> dict[str, Any]:
        self._open_path(Path(path), "路径")
        return self.state

    # ---- 路径选择（浏览器原生）：前端目录浏览模态框 + 显式提交选中的路径 ----
    # choose_* 接收预解析的路径（path 参数），只做业务联动；浏览由 browse_dir 提供。

    def choose_folder(self, field: str, path: str = "") -> str:
        if not path:
            return super().choose_folder(field)
        value = str(path or "").strip()
        if not value:
            return str(self.state.get(field, ""))
        if not Path(value).is_dir():
            self._log(f"所选目录不存在：{value}")
            return ""
        previous = self.state.get(field, "")
        self.state[field] = value
        if field == "templateDir" and value != previous:
            self.state["templateManual"] = False
        if field in {"input", "templateDir"}:
            if field == "input" and not self.state["outputPinned"]:
                self.state["output"] = str(Path(value) / "执行结果")
                self.state["outputAuto"] = True
            # 仅选择源数据目录时触发一次模板推荐。
            self._recognize(allow_template_auto=(field == "input"))
        elif field == "output":
            self.state["outputAuto"] = False
            self.state["outputPinned"] = True
        self._save_settings()
        return value

    def choose_file(self, field: str, path: str = "") -> str:
        if not path:
            return super().choose_file(field)
        value = str(path or "").strip()
        if not value:
            return str(self.state.get(field, ""))
        if not Path(value).is_file():
            self._log(f"所选文件不存在：{value}")
            return ""
        self.state[field] = value
        if field == "template":
            self.state["templateManual"] = True
        self._save_settings()
        return value

    def choose_history_config(self, path: str = "") -> str:
        if not path:
            return super().choose_history_config()
        if self.state["busy"]:
            return str(self.history_path)
        value = str(path or "").strip()
        if not value or not Path(value).is_file():
            return str(self.history_path)
        self.history_path = Path(value)
        self.state["historyConfig"] = value
        self.settings.history_config = value
        self.settings_store.save(self.settings)
        self._log(f"已选择历史审核配置：{value}")
        return value

    def create_template_merge(self, base: str = "", picked: list[str] | None = None) -> bool:
        """两步选择改由前端浏览器目录浏览完成：base=基准模板，picked=并入清单。"""
        if not base and picked is None:
            return super().create_template_merge()
        if self.state["busy"]:
            return False
        base_template = Path(str(base or "")).resolve() if base else None
        if base_template is None or not base_template.is_file():
            self._log("已取消制作联合模板：未选择基准模板")
            return False
        source_templates = [Path(p).resolve() for p in (picked or []) if p]
        if not [path for path in source_templates if path != base_template]:
            self._log("已取消制作联合模板：待复制工作簿不能只有底稿模板本身")
            return False
        self.state["busy"] = True
        self.state["status"] = "正在制作联合模板，请勿关闭页面……"
        self._log(f"开始制作联合模板：底稿“{base_template.name}”")
        self._log("提示：请将含外部依赖工作表的模板选作底稿；原始文件不会修改")
        threading.Thread(
            target=self._template_merge_worker,
            args=(base_template, source_templates),
            daemon=True,
        ).start()
        return True

    # 报表采集系统页：路径选择由前端浏览器目录浏览完成，这里只做业务联动。
    def choose_period_compare(self, kind: str, path: str = "") -> str:
        if not path:
            return super().choose_period_compare(kind)
        if self.state["busy"] or kind not in {"pcCurDir", "pcPreDir", "pcCentral", "pcConfig", "pcOutput"}:
            return ""
        value = str(path or "").strip()
        if not value:
            return str(self.state.get(kind, ""))
        if kind in {"pcCentral", "pcConfig"} and not Path(value).is_file():
            self._log(f"所选文件不存在：{value}")
            return ""
        if kind not in {"pcCentral", "pcConfig"} and not Path(value).is_dir():
            self._log(f"所选目录不存在：{value}")
            return ""
        self.state[kind] = str(Path(value).resolve())
        if kind == "pcCurDir" and self.state.get("pcOutputAuto", True):
            self.state["pcOutput"] = str(Path(value).resolve() / "执行结果")
        elif kind == "pcOutput":
            self.state["pcOutputAuto"] = False
        label = {
            "pcCurDir": "当期目录", "pcPreDir": "上期目录", "pcCentral": "大集中数据",
            "pcConfig": "跨期比较配置", "pcOutput": "输出目录",
        }[kind]
        self._log(f"跨期比较：{label}已选择")
        self._save_settings()
        self._refresh_period_pairs()
        return self.state.get(kind, "")

    def choose_central(self, kind: str, path: str = "") -> str:
        """大集中统计系统：前端浏览器目录浏览选中后提交路径，这里做联动。"""
        if not path:
            return super().choose_central(kind)
        kinds = {"centralCur", "centralPre", "centralOutput", "centralCrossCur", "centralCrossOutput",
                 "centralCompareFile", "centralExplainFile", "centralTemplate", "centralFormOutput",
                 "centralCommonConfig", "centralComparisonConfig", "centralCrossConfig", "centralFormsConfig"}
        if self.state["busy"] or kind not in kinds:
            return ""
        value = str(path or "").strip()
        if not value:
            return str(self.state.get(kind, ""))
        file_kinds = {"centralCur", "centralPre", "centralCrossCur", "centralCompareFile",
                      "centralExplainFile", "centralTemplate", "centralCommonConfig",
                      "centralComparisonConfig", "centralCrossConfig", "centralFormsConfig"}
        if kind in file_kinds and not Path(value).is_file():
            self._log(f"所选文件不存在：{value}")
            return ""
        if kind not in file_kinds and not Path(value).is_dir():
            self._log(f"所选目录不存在：{value}")
            return ""
        self.state[kind] = str(Path(value).resolve())
        label = {
            "centralCur": "本期导数文件", "centralPre": "上期导数文件",
            "centralOutput": "比较输出目录", "centralCompareFile": "比较结果",
            "centralCrossCur": "本期数值核对数据", "centralCrossOutput": "本期数值核对输出目录",
            "centralExplainFile": "说明文件", "centralTemplate": "金融表单模板", "centralFormOutput": "转表输出目录",
            "centralCommonConfig": "大集中通用配置",
            "centralComparisonConfig": "执行比较配置",
            "centralCrossConfig": "本期数值核对配置",
            "centralFormsConfig": "指标比较拆分配置",
        }[kind]
        self._log(f"大集中统计系统：{label}已选择")
        self._save_settings()
        if kind in {"centralCommonConfig", "centralComparisonConfig", "centralCrossConfig", "centralFormsConfig"}:
            self._refresh_config_issues()
        return self.state.get(kind, "")

    # 浏览器没有自绘标题栏；顶栏按钮在 Flask 版中无操作。
    def win_minimize(self) -> None:
        return None

    def win_maximize(self, restore: bool = False) -> None:
        return None

    def win_close(self) -> None:
        return None


# 把 bridge.api.method(...) 代理到 POST /api/method 的桥接层。
# 页面其余脚本一行不改；bridgeready 事件在页面加载完成后派发。
_BRIDGE_JS = """
<script>
(function () {
  function callApi(method, args) {
    return fetch('/api/' + method, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(args || [])
    }).then(function (response) {
      return response.json().then(function (data) {
        if (!response.ok || (data && data.error)) {
          // 后端把业务异常转成了 {error: "..."}；抛成 Error 让界面显示真实原因，
          // 而不是笼统的 "HTTP 500"。
          throw new Error((data && data.error) || ('HTTP ' + response.status));
        }
        return data.result;
      });
    });
  }
  var api = new Proxy({}, {
    get: function (target, name) {
      if (typeof name !== 'string') { return undefined; }
      return function () {
        return callApi(name, Array.prototype.slice.call(arguments));
      };
    }
  });
  window.bridge = { api: api };
  window.addEventListener('load', function () {
    window.dispatchEvent(new Event('bridgeready'));
  });
})();
</script>
"""



def create_app() -> Flask:
    app = Flask(__name__)
    api = FlaskApi(PROJECT_ROOT)

    def as_json(value: Any):
        # WebApi 各方法返回 dict/list/bool/str，均可 JSON 化；None 也按原样返回。
        return jsonify({"result": value})

    @app.get("/")
    def index():
        html = (FRONTEND_DIR / "web" / "index.html").read_text(encoding="utf-8")
        if "</head>" not in html:
            raise RuntimeError("本地界面文件缺失 <head>，无法注入 Flask 桥接层")
        return html.replace("</head>", _BRIDGE_JS + "</head>", 1)

    @app.post("/api/<method>")
    def call(method: str):
        handler = getattr(api, method, None)
        if handler is None or not callable(handler):
            return jsonify({"error": f"未知接口：{method}"}), 404
        args = request.get_json(silent=True) or []
        if not isinstance(args, list):
            return jsonify({"error": "请求体必须是 JSON 数组"}), 400
        # 业务异常回传带文案的 JSON（前端桥接层会把它抛成 Error.message），
        # 避免用户只看到通用 "HTTP 500" 而不知道下一步该做什么。
        try:
            return as_json(handler(*args))
        except (ValueError, RuntimeError) as exc:
            return jsonify({"error": str(exc)}), 400
        except Exception as exc:  # 未预期异常：记日志 + 回传类型与摘要
            api._log(f"接口 {method} 执行失败：{type(exc).__name__}: {exc}")
            return jsonify({"error": f"{type(exc).__name__}: {exc}"}), 500

    return app


app = create_app()


if __name__ == "__main__":
    app.run(host="127.0.0.1", port=8750, debug=False)
