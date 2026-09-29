#!/bin/sh
# 源码目录用系统/虚拟环境 Python；DEB 安装后启动随包冻结的 Linux 程序；
# tar.gz 便携包（解压目录含 app/ 冻结程序与 core/、config/）自动按冻结模式启动。
set -eu

SCRIPT_DIR=$(CDPATH= cd -- "$(dirname -- "$0")" && pwd)
FROZEN=0
FROZEN_APP=

if [ -z "${BASE_AUDIT_INSTALL_ROOT:-}" ]; then
    # tar.gz 便携发行包：未显式指定安装根时，检测自身是否位于发行包的 shell-flask/ 内。
    PARENT_DIR=$(CDPATH= cd -- "$SCRIPT_DIR/.." && pwd)
    if [ -x "$PARENT_DIR/app/base-audit-v3" ] && [ -d "$PARENT_DIR/core/frontend" ]; then
        BASE_AUDIT_INSTALL_ROOT=$PARENT_DIR
        export BASE_AUDIT_INSTALL_ROOT
    fi
fi

if [ -n "${BASE_AUDIT_INSTALL_ROOT:-}" ]; then
    RELEASE=$(CDPATH= cd -- "$BASE_AUDIT_INSTALL_ROOT" && pwd)
    DATA_HOME=${XDG_DATA_HOME:-"$HOME/.local/share"}/base-audit-v3
    PROJECT_ROOT=$DATA_HOME/core
    APP_DIR=$DATA_HOME/shell-flask
    mkdir -p "$APP_DIR" "$PROJECT_ROOT/data" "$DATA_HOME/config/默认配置"
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
    export BASE_AUDIT_CORE_DIR=$RELEASE/core
    export BASE_AUDIT_CORE_ROOT=$RELEASE/core
    export BASE_AUDIT_PROJECT_ROOT=$PROJECT_ROOT
    export BASE_AUDIT_FRONTEND_DIR=$RELEASE/core/frontend
    export BASE_AUDIT_FROZEN=1
    FROZEN_APP=$RELEASE/app/base-audit-v3
    if [ ! -x "$FROZEN_APP" ]; then
        echo "[错误] DEB 冻结程序不存在或不可执行：$FROZEN_APP" >&2
        echo '请重新安装完整 DEB；不要单独复制 shell-flask/。' >&2
        exit 1
    fi
    FROZEN=1
    PID_FILE=$PROJECT_ROOT/data/app.pid
else
    APP_DIR=$SCRIPT_DIR
    PROJECT_ROOT=$SCRIPT_DIR/../core
    PID_FILE=$PROJECT_ROOT/data/app.pid
    PYTHON=${PYTHON:-python3}
    if [ -x "$SCRIPT_DIR/.venv/bin/python" ]; then
        PYTHON=$SCRIPT_DIR/.venv/bin/python
    elif [ -x "$SCRIPT_DIR/../.venv/bin/python" ]; then
        PYTHON=$SCRIPT_DIR/../.venv/bin/python
    fi
fi

# 桌面入口不弹终端；前台仍可直接运行本脚本查看调试输出。
if [ "${1:-}" = "--background" ]; then
    shift
    STATE_HOME=${XDG_STATE_HOME:-"$HOME/.local/state"}
    LOG_DIR=$STATE_HOME/base-audit-v3/logs
    mkdir -p "$LOG_DIR"
    LOG_FILE=$LOG_DIR/start-$(date '+%Y%m%d_%H%M%S').log
    nohup /bin/sh "$SCRIPT_DIR/启动Flask-UOS.sh" "$@" >>"$LOG_FILE" 2>&1 </dev/null &
    APP_PID=$!
    READY=0
    attempt=0
    while [ "$attempt" -lt 60 ]; do
        if grep -Fq 'Running on http://127.0.0.1:' "$LOG_FILE"; then
            READY=1
            break
        fi
        if ! kill -0 "$APP_PID" 2>/dev/null; then
            break
        fi
        attempt=$((attempt + 1))
        sleep 0.5
    done
    if [ "$READY" -eq 1 ]; then
        echo "[完成] 服务已启动；日志：$LOG_FILE"
        exit 0
    fi
    kill -TERM "$APP_PID" 2>/dev/null || true
    MESSAGE="基础数据审核工具启动失败。详细信息：$LOG_FILE"
    if command -v notify-send >/dev/null 2>&1; then
        notify-send --urgency=critical '基础数据审核工具' "$MESSAGE" >/dev/null 2>&1 || true
    fi
    if command -v xdg-open >/dev/null 2>&1 && [ -s "$LOG_FILE" ]; then
        xdg-open "$LOG_FILE" >/dev/null 2>&1 </dev/null &
    fi
    echo "[错误] $MESSAGE" >&2
    exit 1
fi

if [ "$FROZEN" -eq 0 ]; then
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
        echo '[错误] 缺少 Flask 运行依赖。源码模式请执行：python3 -m pip install -r requirements.txt' >&2
        exit 1
    fi
fi

cd "$APP_DIR"
echo '[启动] 基础数据审核工具；优先使用 http://127.0.0.1:8750/，仅在其它程序占用时顺延。'
if [ "$FROZEN" -eq 1 ]; then
    exec "$FROZEN_APP" "$@"
fi
exec "$PYTHON" run.py "$@"
