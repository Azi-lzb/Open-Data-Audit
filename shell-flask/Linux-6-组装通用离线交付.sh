#!/bin/sh
# 在 Linux / WSL 组装审核工具 DEB + 同架构官方 LibreOffice 25.8.7 DEB 离线交付包。
# 审核工具和配置不在此脚本中重打包；配置身份从输入审核工具 DEB 的元数据读取。
# 用法：sh Linux-6-组装通用离线交付.sh 审核工具.deb LibreOffice_25.8.7_Linux_<架构>_deb.tar.gz 输出目录 [--internal-config]
set -eu

fail() { printf '[错误] %s\n' "$*" >&2; exit 1; }
[ "$#" -ge 3 ] && [ "$#" -le 4 ] || fail '用法：脚本 审核工具.deb LibreOffice_25.8.7_Linux_<架构>_deb.tar.gz 输出目录 [--internal-config]'
APP_DEB=$1
LO_TAR=$2
OUT_DIR=$3
PROFILE_ASSERTION=
if [ "$#" -eq 4 ]; then
    [ "$4" = --internal-config ] || fail '仅支持第四参数 --internal-config。'
    PROFILE_ASSERTION=internal-config
fi
[ -f "$APP_DEB" ] && [ -f "$LO_TAR" ] && [ -d "$OUT_DIR" ] || fail '输入 DEB、LibreOffice 压缩包或输出目录不存在。'
for tool in dpkg-deb sha256sum tar mktemp find sort xargs sed; do
    command -v "$tool" >/dev/null 2>&1 || fail "缺少构建命令：$tool"
done

# 校验官方发行文件本身，再从中提取 DEB。校验值来自 download.documentfoundation.org。
APP_PACKAGE=$(dpkg-deb -f "$APP_DEB" Package)
case "$APP_PACKAGE" in
    audit-tools|audit-tools-s1-s2) : ;;
    *) fail '审核工具包标识必须为 audit-tools 或 audit-tools-s1-s2。' ;;
esac
NAME_PREFIX=$(dpkg-deb -f "$APP_DEB" X-Audit-Tools-Artifact-Prefix 2>/dev/null || true)
if [ -z "$NAME_PREFIX" ]; then
    case "$APP_PACKAGE" in
        audit-tools) NAME_PREFIX=Audit-Tools ;;
        audit-tools-s1-s2) NAME_PREFIX=Audit-Tools-S1-S2 ;;
    esac
fi
case "$NAME_PREFIX" in
    ''|*[!A-Za-z0-9_-]*|[-_]* ) fail '审核工具 DEB 中的交付文件名前缀无效。' ;;
esac
ARCH=$(dpkg-deb -f "$APP_DEB" Architecture)
PROFILE=$(dpkg-deb -f "$APP_DEB" X-Audit-Tools-Config-Profile 2>/dev/null || true)
case "$PROFILE" in
    internal-config|public-config) : ;;
    '')
        [ "$PROFILE_ASSERTION" = internal-config ] || fail '审核工具 DEB 缺少配置来源标记。请用新版 Linux-2 重新打包；旧包只有在确认含真实配置时才可加 --internal-config。'
        PROFILE=internal-config
        printf '[警告] 旧版 DEB 没有配置来源标记，将按你显式指定的真实配置处理；组装脚本不会检查 DEB 内的工作簿内容。\n' >&2
        ;;
    *) fail "审核工具 DEB 的配置来源标记无效：$PROFILE。" ;;
esac
[ -z "$PROFILE_ASSERTION" ] || [ "$PROFILE" = "$PROFILE_ASSERTION" ] \
    || fail "参数声明 $PROFILE_ASSERTION，但审核工具 DEB 标记为 $PROFILE。"
printf '[配置] 沿用审核工具 DEB 内的配置身份：%s。\n' "$PROFILE"
case "$ARCH" in
    amd64)
        LO_ARCH_DIR=x86_64
        LO_ARCH_FILE=x86-64
        EXPECTED_LO_SHA=7f4d7b2e36921eec5122c655249a24cc88935ee357e8261fd3bccd15aa1f7b9f
        CPU_REQUIREMENT='x86-64-v2 CPU'
        ;;
    arm64)
        LO_ARCH_DIR=aarch64
        LO_ARCH_FILE=aarch64
        EXPECTED_LO_SHA=67e9b7dcdeae72c7aa1357345307e67376fc2b729a7f9ebfafb372b010e22ffa
        CPU_REQUIREMENT='ARM64 CPU'
        ;;
    *) fail "仅支持 amd64 或 arm64 审核工具 DEB；实际架构：$ARCH。" ;;
esac
[ "$(basename "$LO_TAR")" = "LibreOffice_25.8.7_Linux_${LO_ARCH_FILE}_deb.tar.gz" ] \
    || fail "LibreOffice 压缩包文件名与审核工具 $ARCH 架构不匹配。"
printf '%s  %s\n' "$EXPECTED_LO_SHA" "$LO_TAR" | sha256sum -c - >/dev/null \
    || fail 'LibreOffice 压缩包与官方 SHA-256 不一致。'
APP_VERSION=$(dpkg-deb -f "$APP_DEB" Version)
VERSION=${APP_VERSION#*:}
case "$VERSION" in
    26.*.*.*) : ;;
    *) fail "审核工具版本格式异常：$APP_VERSION" ;;
esac

STAGE=$(mktemp -d)
trap 'rm -rf -- "$STAGE"' EXIT HUP INT TERM
STAMP=$(date +%Y%m%d_%H%M%S)
NAME=${NAME_PREFIX}_V${VERSION}_Linux-${ARCH}_LO25.8.7_${PROFILE}_offline_${STAMP}
ROOT=$STAGE/$NAME
mkdir -p "$ROOT/packages/audit-tools" "$ROOT/packages/libreoffice/DEBS" "$STAGE/extract"
cp "$APP_DEB" "$ROOT/packages/audit-tools/"
SCRIPT_DIR=$(CDPATH= cd -- "$(dirname -- "$0")" && pwd)
sed 's/\r$//' "$SCRIPT_DIR/Linux-5-通用离线安装.sh" > "$ROOT/install.sh"
cp "$SCRIPT_DIR/../docs/deb-install-domestic-linux.md" "$ROOT/安装说明.md"
chmod 755 "$ROOT/install.sh"

tar -xzf "$LO_TAR" -C "$STAGE/extract"
FOUND=0
for deb in "$STAGE"/extract/*/DEBS/*.deb; do
    [ -f "$deb" ] || fail '官方压缩包缺少 DEBS/*.deb。'
    cp "$deb" "$ROOT/packages/libreoffice/DEBS/"
    FOUND=$((FOUND + 1))
done
[ "$FOUND" -eq 42 ] || fail "官方压缩包中预期 42 个 DEB，实际 $FOUND 个。"

CHECKSUMS=$STAGE/SHA256SUMS
(cd "$ROOT" && find install.sh 安装说明.md packages -type f -print0 | sort -z | xargs -0 sha256sum) > "$CHECKSUMS"
mv "$CHECKSUMS" "$ROOT/SHA256SUMS"

printf '[交付包] 已放入审核工具、LibreOffice DEB 和 install.sh；配置身份：%s。\n' "$PROFILE"
sh "$ROOT/install.sh" --check
OUTPUT=$OUT_DIR/$NAME.tar.gz
tar -C "$STAGE" -czf "$OUTPUT" "$NAME"
(cd "$OUT_DIR" && sha256sum "$(basename "$OUTPUT")" > "$(basename "$OUTPUT").sha256")
printf '[完成] %s\n' "$OUTPUT"
