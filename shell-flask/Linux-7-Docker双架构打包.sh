#!/bin/sh
# 用 Docker 分别构建 Linux AMD64 / ARM64 审核工具，并组装对应的 LO 25.8.7 离线包。
# 每个架构的最终目录只放审核工具 DEB、审核工具便携 TAR.GZ、审核工具+LO 离线 TAR.GZ。
set -eu

SCRIPT_DIR=$(CDPATH= cd -- "$(dirname -- "$0")" && pwd)
ROOT=$(CDPATH= cd -- "$SCRIPT_DIR/.." && pwd)
DIST_ROOT=${LINUX7_DIST_DIR:-$SCRIPT_DIR/dist}
DOCKERFILE=$SCRIPT_DIR/Dockerfile.arm64-builder
LO_DIR=$SCRIPT_DIR/LibreOfficeDEB
PROFILE_ID=full
OUTPUT_SUFFIX=
ARTIFACT_PREFIX=

fail() { printf '[错误] %s\n' "$*" >&2; exit 1; }
usage() {
    cat <<EOF
用法：sh Linux-7-Docker双架构打包.sh [--profile 方案]

使用 Docker Buildx 分别构建 linux/amd64 与 linux/arm64，调用 Linux-2 生成审核工具
DEB 和便携 TAR.GZ，再调用 Linux-6 使用本地 LibreOffice 25.8.7 包生成离线 TAR.GZ。
默认输出到：$SCRIPT_DIR/dist/Linux-amd64_V版本/ 与 Linux-arm64_V版本/
--profile 方案  从 distribution-profiles.json 读取交付范围；默认 full，也可选 s1s2。
--s1-s2-only    旧参数兼容写法，等同于 --profile s1s2。
可以通过 LINUX7_DIST_DIR 指定输出根目录。
EOF
}

for arg in "$@"; do
    case "$arg" in
        --s1-s2-only) PROFILE_ID=s1s2 ;;
        --profile) : ;;
        -h|--help) usage; exit 0 ;;
        *)
            if [ "${PREV_ARG:-}" = --profile ]; then PROFILE_ID=$arg
            else usage >&2; exit 2
            fi
            ;;
    esac
    PREV_ARG=$arg
done
case "${PREV_ARG:-}" in --profile) fail '--profile 后缺少方案标识。' ;; esac

for file in \
    "$ROOT/版本管理/VERSION_MANIFEST.json" \
    "$SCRIPT_DIR/distribution-profiles.json" \
    "$SCRIPT_DIR/distribution_profile.py" \
    "$SCRIPT_DIR/Linux-2-打包DEB与TARGZ.sh" \
    "$SCRIPT_DIR/Linux-5-通用离线安装.sh" \
    "$SCRIPT_DIR/Linux-6-组装通用离线交付.sh" \
    "$DOCKERFILE" \
    "$LO_DIR/LibreOffice_25.8.7_Linux_x86-64_deb.tar.gz" \
    "$LO_DIR/LibreOffice_25.8.7_Linux_aarch64_deb.tar.gz"; do
    [ -f "$file" ] || fail "缺少构建文件：$file"
done

command -v python3 >/dev/null 2>&1 || fail 'WSL 中未找到 python3，无法读取 distribution-profiles.json。'
profile_field() { python3 "$SCRIPT_DIR/distribution_profile.py" field "$PROFILE_ID" "$1"; }
PROFILE_NAME=$(profile_field name) || exit $?
PROFILE_ENABLED_SYSTEMS=$(profile_field enabledSystems) || exit $?
OUTPUT_SUFFIX=$(profile_field outputSuffix) || exit $?
ARTIFACT_PREFIX=$(profile_field artifactPrefix) || exit $?
printf '[交付范围] %s；开放系统：%s。\n' "$PROFILE_NAME" "$PROFILE_ENABLED_SYSTEMS"

VERSION=$(awk -F '"' '/"version"[[:space:]]*:/ { print $4; exit }' \
    "$ROOT/版本管理/VERSION_MANIFEST.json")
printf '%s\n' "$VERSION" | grep -Eq '^[0-9]+\.[0-9]+\.[0-9]+\.[0-9]+$' \
    || fail "无法从版本清单读取有效产品版本：$VERSION"

command -v docker >/dev/null 2>&1 || fail '未找到 Docker。请在 WSL 中启用 Docker Desktop 集成。'
docker buildx version >/dev/null 2>&1 || fail '当前 Docker 缺少 Buildx。请更新 Docker Desktop。'
docker info >/dev/null 2>&1 || fail 'Docker 引擎未运行或当前 WSL 发行版未连接 Docker Desktop。'

mkdir -p "$DIST_ROOT"
DIST_ROOT=$(CDPATH= cd -- "$DIST_ROOT" && pwd)
CACHE_DIR=${LINUX7_CACHE_DIR:-${XDG_CACHE_HOME:-$HOME/.cache}/base-audit-v3/docker-linux-build}
mkdir -p "$CACHE_DIR"
CACHE_DIR=$(CDPATH= cd -- "$CACHE_DIR" && pwd)
RUNTIME_CACHE_DIR=$CACHE_DIR/runtime-linux
mkdir -p "$RUNTIME_CACHE_DIR"
STAGE=$(mktemp -d "$DIST_ROOT/.linux7-stage.XXXXXX")
trap 'rm -rf -- "$STAGE"' EXIT HUP INT TERM

# Dockerfile 不复制项目内容，只构建两个平台的轻量构建环境。
BUILD_CONTEXT=$STAGE/docker-context
mkdir -p "$BUILD_CONTEXT"
cp "$DOCKERFILE" "$BUILD_CONTEXT/Dockerfile"

build_arch() {
    ARCH=$1
    case "$ARCH" in
        amd64)
            PLATFORM=linux/amd64
            LO_ARCHIVE=LibreOffice_25.8.7_Linux_x86-64_deb.tar.gz
            ;;
        arm64)
            PLATFORM=linux/arm64
            LO_ARCHIVE=LibreOffice_25.8.7_Linux_aarch64_deb.tar.gz
            ;;
        *) fail "不支持的构建架构：$ARCH" ;;
    esac

    IMAGE=audit-tools-linux-builder:glibc228-$ARCH
    BUILD_DIR=$STAGE/$ARCH
    mkdir -p "$BUILD_DIR"

    printf '\n[Docker] 构建 %s 环境（%s）……\n' "$ARCH" "$PLATFORM"
    docker buildx build --platform "$PLATFORM" --load \
        --file "$BUILD_CONTEXT/Dockerfile" --tag "$IMAGE" "$BUILD_CONTEXT"

    printf '[Linux-2] 在 %s 容器中生成审核工具 DEB 和便携 TAR.GZ……\n' "$ARCH"
    set -- --profile "$PROFILE_ID"
    docker run --rm --platform "$PLATFORM" \
        -v "$ROOT:/work/project" \
        -v "$RUNTIME_CACHE_DIR:/work/project/shell-flask/runtime/linux" \
        -v "$BUILD_DIR:/work/output" \
        -v "$CACHE_DIR:/root/.cache" \
        -e XDG_CACHE_HOME=/root/.cache \
        -e FLASK_DIST_DIR=/work/output \
        -w /work/project/shell-flask \
        "$IMAGE" sh ./Linux-2-打包DEB与TARGZ.sh "$@"

    APP_DEB=
    for candidate in "$BUILD_DIR"/*.deb; do
        [ -f "$candidate" ] || continue
        [ -z "$APP_DEB" ] || fail "$ARCH 构建产生了多个 DEB。"
        APP_DEB=$candidate
    done
    [ -n "$APP_DEB" ] || fail "$ARCH 构建未生成审核工具 DEB。"

    printf '[Linux-6] 使用本地 LO 25.8.7 组装 %s 离线包……\n' "$ARCH"
    docker run --rm --platform "$PLATFORM" \
        -v "$ROOT:/work/project" \
        -v "$BUILD_DIR:/work/output" \
        -v "$CACHE_DIR:/root/.cache" \
        -e XDG_CACHE_HOME=/root/.cache \
        -w /work/project/shell-flask \
        "$IMAGE" sh ./Linux-6-组装通用离线交付.sh \
        "/work/output/${APP_DEB##*/}" \
        "/work/project/shell-flask/LibreOfficeDEB/$LO_ARCHIVE" \
        /work/output
}

# 顺序构建，避免两个架构同时争用项目内的运行时缓存。
build_arch amd64
build_arch arm64

validate_output_dir() {
    TARGET_DIR=$1
    [ -d "$TARGET_DIR" ] || return 0
    for existing in "$TARGET_DIR"/*; do
        [ -e "$existing" ] || continue
        [ -f "$existing" ] || fail "输出目录含有非文件内容，未覆盖：$existing"
        case "${existing##*/}" in
            "$ARTIFACT_PREFIX"*.deb|"$ARTIFACT_PREFIX"*.tar.gz|*.sha256|*.build-info.txt) : ;;
            *) fail "输出目录含有非打包产物，未覆盖：$existing" ;;
        esac
    done
}

prepare_output_dir() {
    TARGET_DIR=$1
    mkdir -p "$TARGET_DIR"
    rm -f -- "$TARGET_DIR/$ARTIFACT_PREFIX"*.deb \
        "$TARGET_DIR/$ARTIFACT_PREFIX"*.tar.gz \
        "$TARGET_DIR"/*.sha256 "$TARGET_DIR"/*.build-info.txt
}

OUT_AMD64=$DIST_ROOT/Linux-amd64_V$VERSION$OUTPUT_SUFFIX
OUT_ARM64=$DIST_ROOT/Linux-arm64_V$VERSION$OUTPUT_SUFFIX
validate_output_dir "$OUT_AMD64"
validate_output_dir "$OUT_ARM64"
prepare_output_dir "$OUT_AMD64"
prepare_output_dir "$OUT_ARM64"

copy_three_packages() {
    ARCH=$1
    TARGET_DIR=$2
    BUILD_DIR=$STAGE/$ARCH
    APP_DEB=
    APP_TGZ=
    OFFLINE_TGZ=

    for candidate in "$BUILD_DIR"/*.deb; do
        [ -f "$candidate" ] || continue
        [ -z "$APP_DEB" ] || fail "$ARCH 输出中有多个 DEB。"
        APP_DEB=$candidate
    done
    for candidate in "$BUILD_DIR"/*.tar.gz; do
        [ -f "$candidate" ] || continue
        case "$candidate" in
            *_offline_*.tar.gz)
                [ -z "$OFFLINE_TGZ" ] || fail "$ARCH 输出中有多个审核工具+LO 离线包。"
                OFFLINE_TGZ=$candidate
                ;;
            *)
                [ -z "$APP_TGZ" ] || fail "$ARCH 输出中有多个审核工具便携包。"
                APP_TGZ=$candidate
                ;;
        esac
    done
    [ -n "$APP_DEB" ] && [ -n "$APP_TGZ" ] && [ -n "$OFFLINE_TGZ" ] \
        || fail "$ARCH 未得到一个 DEB 和两个 TAR.GZ。"
    cp -- "$APP_DEB" "$APP_TGZ" "$OFFLINE_TGZ" "$TARGET_DIR/"
    printf '[完成] %s：\n  %s\n  %s\n  %s\n' "$TARGET_DIR" \
        "${APP_DEB##*/}" "${APP_TGZ##*/}" "${OFFLINE_TGZ##*/}"
}

copy_three_packages amd64 "$OUT_AMD64"
copy_three_packages arm64 "$OUT_ARM64"
printf '\n[完成] 两个架构的交付目录已生成；每个目录各含一个 DEB 和两个 TAR.GZ。\n'
