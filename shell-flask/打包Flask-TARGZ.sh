#!/bin/sh
# 仅生成 tar.gz 便携发行包（冻结 + glibc 审计 + tar.gz + sha256/build-info）。
# 作为 DEB 的备用方案：目标机无 root、不便安装 deb（或包管理器受限）时，
# 解压任意可写目录即可运行。与 打包Flask-DEB.sh 共用构建引擎 打包Flask-Linux.sh。
# 打包完成后建议运行 测试启动Linux包.sh 在本机实测产物。
# 参数原样传给引擎（--with-local-config、--python 等）。
exec sh "$(CDPATH= cd -- "$(dirname -- "$0")" && pwd)/打包Flask-Linux.sh" --skip-deb "$@"
