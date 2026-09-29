#!/bin/sh
# 仅生成 DEB 安装包（冻结 + glibc 审计 + DEB + sha256/build-info）。
# 与 打包Flask-TARGZ.sh（tar.gz 备用包）共用同一构建引擎 打包Flask-Linux.sh，
# 一次只做一种包，便于在不同芯片/系统的打包机上分开执行。
# 打包完成后建议运行 测试启动Linux包.sh 在本机实测产物。
# 参数原样传给引擎（--with-local-config、--python 等）。
exec sh "$(CDPATH= cd -- "$(dirname -- "$0")" && pwd)/打包Flask-Linux.sh" --skip-tgz "$@"
