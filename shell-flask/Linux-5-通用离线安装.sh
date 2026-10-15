#!/bin/sh
# UOS 20/25、麒麟 V10/V11 主系统离线交付包入口。
# 请在普通用户的系统终端中运行，脚本自行调用 sudo。
# 目录布局：install.sh、packages/audit-tools/*.deb、packages/libreoffice/DEBS/*.deb
set -eu

WAIT_FOR_CLOSE=0
if [ -t 0 ] && [ -t 1 ]; then WAIT_FOR_CLOSE=1; fi

finish_terminal() {
    status=$?
    trap - EXIT
    if [ "$WAIT_FOR_CLOSE" -eq 1 ]; then
        if [ "$status" -eq 0 ]; then
            printf '\n[结束] 脚本已执行完成。按 Enter 关闭此终端。'
        else
            printf '\n[失败] 脚本退出码：%s。请先查看上方错误信息；安装未确认成功。按 Enter 关闭此终端。\n' "$status"
        fi
        IFS= read -r _ || true
    fi
    exit "$status"
}
trap finish_terminal EXIT

fail() {
    printf '[错误] %s\n' "$*" >&2
    exit 1
}

usage() {
    printf '用法：sh %s [--check]\n' "$0"
    printf '  --check  只检查文件、版本、架构和系统条件，不安装\n'
}

case "${1:-}" in
    '') CHECK_ONLY=0 ;;
    --check) CHECK_ONLY=1 ;;
    -h|--help) usage; exit 0 ;;
    *) usage >&2; exit 2 ;;
esac
[ "$#" -le 1 ] || { usage >&2; exit 2; }

ROOT=$(CDPATH= cd -- "$(dirname -- "$0")" && pwd)
APP_DIR=$ROOT/packages/audit-tools
LO_DIR=$ROOT/packages/libreoffice/DEBS
printf '审核工具与 LibreOffice 离线安装\n'
printf '安装包目录：%s\n' "$ROOT"
printf '[阶段 1/4] 正在检查离线包目录和系统安装命令……\n'
[ -d "$APP_DIR" ] && [ -d "$LO_DIR" ] || fail '交付包不完整：缺少审核工具或 LibreOffice DEBS 目录。请完整解压交付包。'
[ -f "$ROOT/SHA256SUMS" ] || fail '交付包缺少 SHA256SUMS。请完整解压交付包。'
for tool in sha256sum dpkg dpkg-deb apt-get; do
    command -v "$tool" >/dev/null 2>&1 || fail "系统缺少命令：$tool"
done

printf '[阶段 2/4] 正在核对系统架构、内核和 glibc 版本……\n'
ARCH=$(dpkg --print-architecture)
case "$ARCH" in
    amd64|arm64) : ;;
    *) fail "本安装入口仅支持 amd64 或 arm64；当前架构：$ARCH。" ;;
esac
KERNEL=$(uname -r)
KERNEL=${KERNEL%%-*}
dpkg --compare-versions "$KERNEL" ge 4.18 \
    || fail "LibreOffice 25.8.7 要求 Linux 内核至少 4.18；当前：$KERNEL。"
command -v getconf >/dev/null 2>&1 || fail '无法检查 glibc 版本：缺少 getconf。'
GLIBC=$(getconf GNU_LIBC_VERSION)
GLIBC=${GLIBC##* }
dpkg --compare-versions "$GLIBC" ge 2.27 \
    || fail "LibreOffice 25.8.7 要求 glibc 至少 2.27；当前：$GLIBC。"

printf '[阶段 3/4] 正在校验离线文件完整性（文件较多时可能需要一点时间）……\n'
(cd "$ROOT" && sha256sum -c SHA256SUMS >/dev/null) || fail '文件校验失败；请重新复制完整交付包。'

printf '[阶段 4/4] 正在核对审核工具与 LibreOffice 的版本、架构和安装文件……\n'
set --
for deb in "$APP_DIR"/*.deb; do
    [ -f "$deb" ] || fail '未找到审核工具 DEB。'
    set -- "$@" "$deb"
done
[ "$#" -eq 1 ] || fail '审核工具 DEB 必须恰好一个。'
APP_DEB=$1
APP_PACKAGE=$(dpkg-deb -f "$APP_DEB" Package)
case "$APP_PACKAGE" in
    audit-tools|audit-tools-s1-s2) : ;;
    *) fail '审核工具 DEB 包名不正确。' ;;
esac
[ "$(dpkg-deb -f "$APP_DEB" Architecture)" = "$ARCH" ] || fail "审核工具 DEB 与当前架构 $ARCH 不一致。"

LO_COUNT=0
for deb in "$LO_DIR"/*.deb; do
    [ -f "$deb" ] || fail 'LibreOffice DEBS 目录中没有 DEB。'
    package=$(dpkg-deb -f "$deb" Package) || fail "无法读取 DEB：$deb"
    version=$(dpkg-deb -f "$deb" Version) || fail "无法读取 DEB：$deb"
    arch=$(dpkg-deb -f "$deb" Architecture) || fail "无法读取 DEB：$deb"
    case "$package" in
        libobasis25.8-*|libreoffice25.8|libreoffice25.8-*) : ;;
        *) fail "混入非 LibreOffice 25.8 软件包：$package" ;;
    esac
    case "$version" in
        25.8.7*) : ;;
        *) fail "LibreOffice 版本不一致：$package $version" ;;
    esac
    case "$arch" in
        "$ARCH"|all) : ;;
        *) fail "LibreOffice 架构不一致：$package $arch" ;;
    esac
    LO_COUNT=$((LO_COUNT + 1))
    set -- "$@" "$deb"
done
[ "$LO_COUNT" -eq 42 ] || fail "LibreOffice DEB 应有 42 个，实际 $LO_COUNT 个。"
printf '[检查] 审核工具：%s；LibreOffice 25.8.7：%s 个 DEB；架构：%s；内核：%s；glibc：%s。\n' \
    "$(dpkg-deb -f "$APP_DEB" Version)" "$LO_COUNT" "$ARCH" "$KERNEL" "$GLIBC"

if [ "$CHECK_ONLY" -eq 1 ]; then
    printf '[完成] 离线包检查通过，未执行安装。\n'
    exit 0
fi

[ "$(id -u)" -ne 0 ] || fail '请以普通用户运行本脚本；脚本只在安装步骤调用 sudo。'
[ -r /etc/os-release ] || fail '无法读取系统版本。'
# shellcheck disable=SC1091
. /etc/os-release
printf '[系统] %s；版本：%s。\n' "${PRETTY_NAME:-${ID:-未知}}" "${VERSION_ID:-未知}"
if [ "${ID:-}" = kylin ]; then
    case "${VERSION_ID:-} ${VERSION:-} ${PRETTY_NAME:-}" in
        *v11*|*V11*)
            printf '\n[系统] 麒麟 V11 请先切换到维护模式并按提示重启。\n'
            command -v mm-cli >/dev/null 2>&1 || fail '未找到 mm-cli；无法确认麒麟 V11 的安装模式。'
            printf '[系统] 查询维护模式时可能需要输入管理员密码；密码输入时终端不会显示字符。\n'
            sudo mm-cli -s || fail '无法查询麒麟维护模式。'
            printf '确认当前已在维护模式，并从主系统终端运行本脚本？[y/N] '
            IFS= read -r answer || fail '未确认维护模式，已取消安装。'
            case "$answer" in
                y|Y|yes|YES|是) : ;;
                *) fail '已取消；请先切换维护模式并重启。' ;;
            esac
            ;;
    esac
fi

printf '\n[安装] 即将请求管理员权限。若终端提示输入密码，输入后按 Enter；密码字符不会显示。\n'
printf '[安装] 开始离线安装审核工具和 LibreOffice 25.8.7。安装期间请勿关闭终端；APT 若询问是否继续，请输入 y 并回车。\n'
printf '[安装] 本脚本禁止联网下载依赖；缺少系统依赖时会明确报错并停止。\n'
if ! sudo apt-get --no-download --allow-downgrades --reinstall install "$@"; then
    fail 'APT 安装失败。请查看上方缺失的系统依赖或模式错误；离线包不会从网络补装依赖。'
fi

installed=$(dpkg-query -W -f='${Version}' "$APP_PACKAGE" 2>/dev/null || true)
[ -n "$installed" ] || fail "APT 已返回成功，但主系统未登记 $APP_PACKAGE；请检查是否在 KARE 环境运行。"
printf '[核对] 审核工具已登记在主系统；版本：%s。\n' "$installed"
if command -v soffice >/dev/null 2>&1; then
    soffice --version || fail 'LibreOffice 已安装，但 soffice 无法启动。'
elif [ -x /opt/libreoffice25.8/program/soffice ]; then
    /opt/libreoffice25.8/program/soffice --version || fail 'LibreOffice 已安装，但 soffice 无法启动。'
else
    fail 'APT 已返回成功，但主系统找不到 LibreOffice soffice。'
fi
printf '[完成] 审核工具和 LibreOffice 均已安装，并且主系统可以找到 soffice。请从新版“审核工具”菜单入口验证 S1-F1。\n'
