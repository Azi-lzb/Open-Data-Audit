#!/bin/sh
# 复用 Linux-7；只构建 S1/S2 版的 amd64、arm64 三件套。
set -eu
SCRIPT_DIR=$(CDPATH= cd -- "$(dirname -- "$0")" && pwd)
exec /bin/sh "$SCRIPT_DIR/Linux-7-Docker双架构打包.sh" --s1-s2-only "$@"
