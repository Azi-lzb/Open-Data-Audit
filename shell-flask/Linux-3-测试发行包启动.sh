#!/bin/sh
# Linux 发行包本机启动冒烟（UOS20/25、麒麟 V10 等 Debian 系通用）。
#
# 用途：打包后、分发前在当前机器实测产物能否启动——解包到临时目录、启动冻结
# 服务（不弹浏览器）、校验 HTTP 200 与页面标题、关闭并确认端口释放。把产物和
# 本脚本拷到目标机直接运行，即可在安装前预检；测试机不需要 Python。
#
# 用法：
#   ./Linux-3-测试发行包启动.sh                          # 自动测 dist/ 下最新 .deb 与 .tar.gz
#   ./Linux-3-测试发行包启动.sh 产物1.deb 产物2.tar.gz    # 测试指定产物（可多个）
#
# 全部通过退出码 0；任一失败退出码 1。测试过程使用 8800 起的端口段，不影响
# 8750 上已在运行的正式实例。
set -u

SCRIPT_DIR=$(CDPATH= cd -- "$(dirname -- "$0")" && pwd)
DIST=${FLASK_DIST_DIR:-$SCRIPT_DIR/dist}

for tool in curl tar mktemp grep sed; do
    command -v "$tool" >/dev/null 2>&1 || { echo "[错误] 缺少 $tool。" >&2; exit 2; }
done

PASS_COUNT=0
FAIL_COUNT=0
NEXT_PORT=8800

if [ $# -eq 0 ]; then
    NEWEST_DEB=$(cd "$DIST" 2>/dev/null && ls -t base-audit-v3_*.deb 2>/dev/null | head -1 || true)
    NEWEST_TGZ=$(cd "$DIST" 2>/dev/null && ls -t base-audit-v3_*.tar.gz 2>/dev/null | head -1 || true)
    # set -- 为整体替换，追加用 "set -- "$@" 新项"，否则 tar.gz 会覆盖 .deb。
    [ -n "$NEWEST_DEB" ] && set -- "$DIST/$NEWEST_DEB"
    [ -n "$NEWEST_TGZ" ] && set -- "$@" "$DIST/$NEWEST_TGZ"
    if [ $# -eq 0 ]; then
        echo "[错误] $DIST 下没有 base-audit-v3_*.deb / *.tar.gz 产物；请先打包或指定产物路径。" >&2
        exit 2
    fi
fi

test_one() {  # $1=产物绝对/相对路径
    ART=$1
    LABEL=$(basename "$ART")
    T=$(mktemp -d /tmp/pkg-smoke.XXXXXX)
    echo
    echo "======== 测试 $LABEL ========"
    FAILURE=
    PID=

    # while/true 单圈循环仅作“出错即跳到收尾”的结构。
    while :; do
        if [ -f "$ART.sha256" ] && command -v sha256sum >/dev/null 2>&1; then
            if (cd "$(dirname -- "$ART")" && sha256sum --status -c "$LABEL.sha256"); then
                echo "[1/4] SHA-256 校验通过"
            else
                FAILURE="SHA-256 校验失败（$LABEL.sha256），文件可能损坏"
                break
            fi
        else
            echo "[1/4] 跳过 SHA-256（无 .sha256 文件或缺少 sha256sum）"
        fi

        echo "[2/4] 解包 ..."
        ROOT=
        case "$LABEL" in
            *.deb)
                if ! command -v dpkg-deb >/dev/null 2>&1; then
                    FAILURE="本机缺 dpkg-deb，无法解包 DEB"
                    break
                fi
                dpkg-deb -x "$ART" "$T/x" || { FAILURE="dpkg-deb 解包失败"; break; }
                ROOT=$T/x/opt/base-audit-v3
                ;;
            *.tar.gz)
                mkdir -p "$T/x"
                tar -xzf "$ART" -C "$T/x" || { FAILURE="tar 解包失败"; break; }
                FIRST=$(cd "$T/x" && ls | head -1)
                [ -n "$FIRST" ] || { FAILURE="压缩包内容为空"; break; }
                ROOT=$T/x/$FIRST
                ;;
            *)
                FAILURE="不认识的产物类型：$LABEL（需 .deb 或 .tar.gz）"
                break
                ;;
        esac
        if [ ! -x "$ROOT/app/base-audit-v3" ] || [ ! -f "$ROOT/shell-flask/Linux-1-启动审核工具.sh" ]; then
            FAILURE="包内缺少 app/base-audit-v3 或 shell-flask/Linux-1-启动审核工具.sh"
            break
        fi

        echo "[3/4] 启动冻结服务（不弹浏览器，端口 $NEXT_PORT 起）..."
        LOG=$T/server.log
        case "$LABEL" in
            *.deb)
                BASE_AUDIT_INSTALL_ROOT=$ROOT sh "$ROOT/shell-flask/Linux-1-启动审核工具.sh" \
                    --no-browser --port "$NEXT_PORT" >"$LOG" 2>&1 &
                ;;
            *)
                sh "$ROOT/shell-flask/Linux-1-启动审核工具.sh" \
                    --no-browser --port "$NEXT_PORT" >"$LOG" 2>&1 &
                ;;
        esac
        PID=$!
        PORT=
        i=0
        while [ "$i" -lt 60 ]; do
            PORT=$(sed -n 's/.*Running on http:\/\/127\.0\.0\.1:\([0-9][0-9]*\).*/\1/p' "$LOG" 2>/dev/null | head -n1)
            [ -n "$PORT" ] && break
            kill -0 "$PID" 2>/dev/null || break
            sleep 0.5
            i=$((i + 1))
        done
        NEXT_PORT=$((NEXT_PORT + 5))
        if [ -z "$PORT" ]; then
            FAILURE="服务 30 秒内未就绪；日志尾部：$(tail -5 "$LOG" 2>/dev/null | tr '\n' ' ')"
            break
        fi

        CODE=$(curl -s --max-time 10 -o "$T/page.html" -w '%{http_code}' "http://127.0.0.1:$PORT/")
        TITLE=$(grep -oE '<title>[^<]{0,60}' "$T/page.html" 2>/dev/null | head -1 | cut -c8-)
        echo "       实际端口 $PORT，HTTP $CODE，页面标题「$TITLE」"
        if [ "$CODE" != "200" ]; then
            FAILURE="首页 HTTP $CODE（预期 200）"
            break
        fi
        case "$TITLE" in
            *基础数据审核工具*) : ;;
            *) FAILURE="页面标题异常：「$TITLE」" ; break ;;
        esac

        echo "[4/4] 关闭服务并确认端口释放 ..."
        kill "$PID" 2>/dev/null || true
        i=0
        while kill -0 "$PID" 2>/dev/null && [ "$i" -lt 20 ]; do
            sleep 0.3
            i=$((i + 1))
        done
        if curl -s --max-time 2 -o /dev/null "http://127.0.0.1:$PORT/"; then
            FAILURE="进程已结束但端口 $PORT 仍可访问，疑似残留实例"
            break
        fi
        echo "       端口已释放。"
        break
    done

    # 兜底：无论成功失败，确保本测试启动的进程已退出。
    if [ -n "$PID" ]; then
        kill "$PID" 2>/dev/null || true
        i=0
        while kill -0 "$PID" 2>/dev/null && [ "$i" -lt 20 ]; do
            sleep 0.3
            i=$((i + 1))
        done
    fi

    if [ -n "$FAILURE" ]; then
        echo "[失败] $LABEL：$FAILURE"
        FAIL_COUNT=$((FAIL_COUNT + 1))
    else
        echo "[通过] $LABEL"
        PASS_COUNT=$((PASS_COUNT + 1))
    fi
    rm -rf -- "$T"
}

for art in "$@"; do
    [ -f "$art" ] || { echo "[跳过] 文件不存在：$art" >&2; FAIL_COUNT=$((FAIL_COUNT + 1)); continue; }
    test_one "$art"
done

echo
echo "======== 汇总：通过 $PASS_COUNT，失败 $FAIL_COUNT ========"
[ "$FAIL_COUNT" -eq 0 ] || exit 1
exit 0
