#!/bin/sh
# 由 UOS .desktop 文件调用：在终端运行操作，并在结束/报错后保留输出供查看。
SCRIPT_DIR=$(CDPATH= cd -- "$(dirname -- "$0")" && pwd)
ACTION=${1:-}
shift || true

case "$ACTION" in
    start)
        sh "$SCRIPT_DIR/启动Flask-UOS.sh" "$@"
        STATUS=$?
        ;;
    deb)
        sh "$SCRIPT_DIR/打包Flask-DEB.sh" "$@"
        STATUS=$?
        ;;
    targz)
        sh "$SCRIPT_DIR/打包Flask-TARGZ.sh" "$@"
        STATUS=$?
        ;;
    test)
        sh "$SCRIPT_DIR/测试启动Linux包.sh" "$@"
        STATUS=$?
        ;;
    linux)
        sh "$SCRIPT_DIR/打包Flask-Linux.sh" "$@"
        STATUS=$?
        ;;
    source)
        sh "$SCRIPT_DIR/打包Flask-UOS.sh" "$@"
        STATUS=$?
        ;;
    *)
        echo '[错误] 未知桌面入口操作。' >&2
        STATUS=2
        ;;
esac

echo
if [ "$STATUS" -eq 0 ]; then
    printf '操作已完成。按回车关闭此窗口。'
else
    printf '操作失败（退出码 %s）。请先查看上方错误信息，按回车关闭此窗口。' "$STATUS"
fi
read -r _ || true
exit "$STATUS"
