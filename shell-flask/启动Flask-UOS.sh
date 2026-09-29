#!/bin/sh
# 源码目录用系统/虚拟环境 Python；DEB 安装后启动随包冻结的 Linux 程序。
set -eu

SCRIPT_DIR=$(CDPATH= cd -- "$(dirname -- "$0")" && pwd)

if [ -n "${BASE_AUDIT_INSTALL_ROOT:-}" ]; then
    RELEASE=$(CDPATH= cd -- "$BASE_AUDIT_INSTALL_ROOT" && pwd)
    DATA_HOME=${XDG_DATA_HOME:-"$HOME/.local/share"}/base-audit-v3
    PROJECT_ROOT=$DATA_HOME/core
    mkdir -p "$PROJECT_ROOT/data" "$DATA_HOME/config/默认配置"
    for source in "$RELEASE"/config/*.xlsx; do
        [ -f "$source" ] || continue
        target=$DATA_HOME/config/$(basename "$source")
        [ -e "$target" ] || cp "$source" "$target"
    done
    for source in "$RELEASE"/config/默认配置/*.xlsx; do
        [ -f "$source" ] || continue
        target=$DATA_HOME/config/默认配置/$(basename "$source")
        [ -e "$target" ] || cp "$source" "$target"
    done
    if [ -f "$RELEASE/core/基础数据审核工具使用说明.docx" ]; then
        cp "$RELEASE/core/基础数据审核工具使用说明.docx" "$PROJECT_ROOT/"
    fi
    export BASE_AUDIT_CORE_ROOT="$RELEASE/core"
    export BASE_AUDIT_PROJECT_ROOT="$PROJECT_ROOT"
    export BASE_AUDIT_FRONTEND_DIR="$RELEASE/core/frontend"
    APP_BIN=$RELEASE/app/base-audit-v3
    if [ ! -x "$APP_BIN" ]; then
        echo "[错误] DEB 冻结程序不存在或不可执行：$APP_BIN" >&2
        echo '请重新安装完整 DEB；不要单独复制 shell-flask/。' >&2
        exit 1
    fi
    exec "$APP_BIN" "$@"
else
    APP_DIR=$SCRIPT_DIR
    PYTHON=${PYTHON:-python3}
    if [ -x "$SCRIPT_DIR/.venv/bin/python" ]; then
        PYTHON=$SCRIPT_DIR/.venv/bin/python
    elif [ -x "$SCRIPT_DIR/../.venv/bin/python" ]; then
        PYTHON=$SCRIPT_DIR/../.venv/bin/python
    fi
fi

if ! command -v "$PYTHON" >/dev/null 2>&1 && [ ! -x "$PYTHON" ]; then
    echo "[错误] 未找到 Python 3：$PYTHON" >&2
    exit 1
fi
PY_VERSION=$($PYTHON -c 'import sys; print("%d.%d" % sys.version_info[:2])')
if ! "$PYTHON" -c 'import sys; raise SystemExit(0 if sys.version_info >= (3, 8) else 1)'; then
    echo "[错误] 源码启动至少需要 Python 3.8；当前为 $PY_VERSION。UOS20 可安装自包含 DEB 后启动。" >&2
    exit 1
fi
if ! "$PYTHON" -c 'import flask, openpyxl, xlrd, lxml' >/dev/null 2>&1; then
    echo "[错误] 缺少 Flask 运行依赖。源码模式请执行：python3 -m pip install -r requirements.txt" >&2
    exit 1
fi
cd "$APP_DIR"
echo "[启动] 基础数据审核工具；浏览器地址 http://127.0.0.1:8750/（端口占用时自动顺延）。"
exec "$PYTHON" run.py "$@"
