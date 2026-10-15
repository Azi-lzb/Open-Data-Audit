#!/bin/sh
# 仅交付 S1/S2 的 Linux 发行包入口；DEB 与 TAR.GZ 共用 Linux-2 构建流程。
set -eu

SCRIPT_DIR=$(CDPATH= cd -- "$(dirname -- "$0")" && pwd)
exec /bin/sh "$SCRIPT_DIR/Linux-2-打包DEB与TARGZ.sh" --s1-s2-only "$@"
