#!/bin/sh
# 在 UOS/Linux 上构建自带 Python 运行时的 PyInstaller DEB。
# 面向旧版 UOS 时，必须在最旧的目标系统上构建，以匹配其 glibc。
set -eu

SCRIPT_DIR=$(CDPATH= cd -- "$(dirname -- "$0")" && pwd)
ROOT=$(CDPATH= cd -- "$SCRIPT_DIR/.." && pwd)
DIST=${FLASK_DIST_DIR:-$SCRIPT_DIR/dist}
RELEASE=/opt/base-audit-v3
PYTHON=${PYTHON:-}
WITH_LOCAL_CONFIG=0

usage() {
    cat <<'EOF'
用法：
  ./打包Flask-DEB-UOS.sh [--with-local-config]

默认生成清理过机构参照和审核历史的正式 DEB。
--with-local-config 仅用于自测，会把本机 config-real/ 带入包内，禁止外发。

构建机需要 Python 3.12、网络（首次安装 PyInstaller/运行依赖）、dpkg-deb。
Python 3.12 只用于构建；安装后的程序自带 Python，目标机无需安装 Python。
PYTHON=/路径/python3.12 可指定构建解释器。
EOF
}

for arg in "$@"; do
    case "$arg" in
        --with-local-config) WITH_LOCAL_CONFIG=1 ;;
        -h|--help) usage; exit 0 ;;
        *) echo "[错误] 未知参数：$arg" >&2; usage >&2; exit 2 ;;
    esac
done

for tool in dpkg-deb dpkg mktemp sha256sum; do
    command -v "$tool" >/dev/null 2>&1 || { echo "[错误] 缺少 $tool，请在 UOS/Linux 构建机安装。" >&2; exit 1; }
done
if [ -z "$PYTHON" ]; then
    for candidate in python3.12 python3 python; do
        if command -v "$candidate" >/dev/null 2>&1; then
            candidate_version=$("$candidate" -c 'import sys; print("%d.%d" % sys.version_info[:2])' 2>/dev/null || true)
            if [ "$candidate_version" = 3.12 ]; then PYTHON=$candidate; break; fi
        fi
    done
fi
[ -n "$PYTHON" ] && command -v "$PYTHON" >/dev/null 2>&1 || {
    echo "[错误] 构建机 PATH 中没有 Python 3.12。" >&2
    echo "可用 PYTHON=/完整路径/python3.12 指定。目标机不需要 Python。" >&2
    exit 1
}
[ -d "$ROOT/core/src/base_audit" ] || { echo "[错误] 缺少共享核心：$ROOT/core/src/base_audit" >&2; exit 1; }
[ -f "$ROOT/core/frontend/web/index.html" ] || { echo '[错误] 缺少前端 index.html。' >&2; exit 1; }
if [ ! -f "$ROOT/基础数据审核工具使用说明.docx" ] && [ ! -f "$ROOT/core/基础数据审核工具使用说明.docx" ]; then
    echo '[提示] 未提供使用说明 DOCX；继续构建，包内不含该文档。'
fi
[ -d "$ROOT/config" ] || { echo '[错误] 缺少正式 config/。' >&2; exit 1; }
if [ "$WITH_LOCAL_CONFIG" = 1 ] && [ ! -d "$ROOT/config-real" ]; then
    echo '[错误] --with-local-config 需要本机 config-real/。' >&2
    exit 1
fi

BASE_PY_VERSION=$("$PYTHON" -c 'import sys; print("%d.%d" % sys.version_info[:2])')
[ "$BASE_PY_VERSION" = 3.12 ] || {
    echo "[错误] 当前指定的是 Python $BASE_PY_VERSION；UOS 冻结包统一用 Python 3.12 构建。" >&2
    exit 1
}

ARCH=$(dpkg --print-architecture)
BUILD_OS=linux
if [ -r /etc/os-release ]; then
    . /etc/os-release
    BUILD_OS=${ID:-linux}-${VERSION_ID:-unknown}
fi
BUILD_OS=$(printf '%s' "$BUILD_OS" | tr -c 'A-Za-z0-9._-' '_')
DEB_VERSION=$(date +%Y.%m.%d.%H%M%S)
PACKAGE=base-audit-v3_${DEB_VERSION}_${BUILD_OS}_${ARCH}_py312.deb

# 构建 venv 放在用户缓存，不污染仓库，也不依赖系统 Python 的第三方包。
CACHE_HOME=${XDG_CACHE_HOME:-$HOME/.cache}
BUILD_ENV=$CACHE_HOME/base-audit-v3/${BUILD_OS}-${ARCH}-py312
mkdir -p "$(dirname -- "$BUILD_ENV")"
if [ ! -x "$BUILD_ENV/bin/python" ]; then
    "$PYTHON" -m venv "$BUILD_ENV" || {
        echo '[错误] 创建 Python 3.12 构建环境失败。请确认安装了 python3.12-venv。' >&2
        exit 1
    }
fi
BUILD_PY=$BUILD_ENV/bin/python
VENV_PY_VERSION=$("$BUILD_PY" -c 'import sys; print("%d.%d" % sys.version_info[:2])')
[ "$VENV_PY_VERSION" = 3.12 ] || { echo "[错误] 缓存构建环境版本异常：$VENV_PY_VERSION" >&2; exit 1; }

echo "[构建环境] $BUILD_OS / $ARCH / Python $VENV_PY_VERSION"
echo '[依赖] 安装 Flask 运行库和 PyInstaller ...'
"$BUILD_PY" -m pip install --disable-pip-version-check --upgrade pip
"$BUILD_PY" -m pip install --disable-pip-version-check -r "$SCRIPT_DIR/requirements.txt" 'pyinstaller>=6,<7'
"$BUILD_PY" -c 'import flask, openpyxl, xlrd, lxml, PyInstaller; print("[依赖] 导入检查通过")'

STAGE=$(mktemp -d "${TMPDIR:-/tmp}/base-audit-deb.XXXXXXXX")
trap 'rm -rf -- "$STAGE"' EXIT HUP INT TERM
PAYLOAD=$STAGE$RELEASE
PYI_DIST=$STAGE/pyinstaller-dist
mkdir -p "$PAYLOAD/app" "$PAYLOAD/shell-flask" "$PAYLOAD/core/frontend" \
    "$PAYLOAD/config/默认配置" "$STAGE/DEBIAN" "$STAGE/usr/bin" \
    "$STAGE/usr/share/applications" "$DIST" "$PYI_DIST"

echo '[PyInstaller] 冻结 Flask、共享核心和 Python 3.12 运行时 ...'
set -- --noconfirm --clean --onedir --name base-audit-v3 \
    --paths "$ROOT/core/src" \
    --collect-submodules base_audit \
    --collect-all flask \
    --collect-all openpyxl \
    --collect-all xlrd \
    --collect-all lxml \
    --distpath "$PYI_DIST" \
    --workpath "$STAGE/work" \
    --specpath "$STAGE/spec"
if "$BUILD_PY" -c 'import tkinter' >/dev/null 2>&1; then
    set -- "$@" --hidden-import tkinter --hidden-import tkinter.filedialog
else
    echo '[提示] 构建 Python 未带 Tk；安装后文件选择仍可用“浏览器内置”模式。'
fi
set -- "$@" "$SCRIPT_DIR/run.py"
"$BUILD_PY" -m PyInstaller "$@"
[ -x "$PYI_DIST/base-audit-v3/base-audit-v3" ] || { echo '[错误] PyInstaller 未生成 Linux 启动程序。' >&2; exit 1; }
cp -R "$PYI_DIST/base-audit-v3/." "$PAYLOAD/app/"

# 冻结代码在 app/；前端与配置为外部资源，用户数据保存在各自 HOME 中。
cp -R "$ROOT/core/frontend/." "$PAYLOAD/core/frontend/"
for doc in "$ROOT/基础数据审核工具使用说明.docx" "$ROOT/core/基础数据审核工具使用说明.docx"; do
    if [ -f "$doc" ]; then cp "$doc" "$PAYLOAD/core/"; break; fi
done
cp "$SCRIPT_DIR/启动Flask-UOS.sh" "$PAYLOAD/shell-flask/"
chmod 755 "$PAYLOAD/shell-flask/启动Flask-UOS.sh" "$PAYLOAD/app/base-audit-v3"

if [ "$WITH_LOCAL_CONFIG" = 1 ]; then
    echo '[警告] 本包带有本机 config-real/，可能含机构资料和审核历史，仅供本人验证。'
    cp -R "$ROOT/config-real/." "$PAYLOAD/config/"
else
    "$BUILD_PY" - "$ROOT/config" "$PAYLOAD/config" <<'PY'
import shutil
import sys
from pathlib import Path
from openpyxl import load_workbook

source, target = map(Path, sys.argv[1:])
names = (
    "1.逐笔统计系统_配置.xlsx", "2.报表采集系统_配置.xlsx",
    "3.0大集中通用配置.xlsx", "3.1大集中执行比较_配置.xlsx",
    "3.2大集中本期数值核对_配置.xlsx", "3.3大集中指标比较拆分_配置.xlsx",
)
private_sheets = {
    names[0]: ("核查表校验结果", "本地校验结果", "业务说明", "任务说明"),
    names[1]: ("机构参照",),
    names[2]: ("机构地区参照",),
}
for name in names:
    path = source / name
    if not path.is_file():
        raise SystemExit(f"[错误] 缺少正式配置：{path}")
    if name not in private_sheets:
        for directory in (target, target / "默认配置"):
            shutil.copy2(path, directory / name)
        continue
    book = load_workbook(path)
    try:
        for sheet_name in private_sheets[name]:
            if sheet_name not in book.sheetnames:
                raise SystemExit(f"[错误] {name} 缺少需清理的工作表：{sheet_name}")
            sheet = book[sheet_name]
            if sheet.max_row > 1:
                sheet.delete_rows(2, sheet.max_row - 1)
        for directory in (target, target / "默认配置"):
            book.save(directory / name)
    finally:
        book.close()
print("[配置] 六册配置已复制；本机审核历史和机构参照已从发行副本清理。")
PY
fi

cat > "$STAGE/usr/bin/base-audit-v3" <<'SH'
#!/bin/sh
set -eu
BASE_AUDIT_INSTALL_ROOT=/opt/base-audit-v3
export BASE_AUDIT_INSTALL_ROOT
exec "$BASE_AUDIT_INSTALL_ROOT/shell-flask/启动Flask-UOS.sh" "$@"
SH
chmod 755 "$STAGE/usr/bin/base-audit-v3"

cat > "$STAGE/usr/share/applications/base-audit-v3.desktop" <<'DESKTOP'
[Desktop Entry]
Version=1.0
Type=Application
Name=基础数据审核工具 V3
Comment=启动本机 Flask 审核工具
Exec=base-audit-v3
Terminal=true
Categories=Office;
DESKTOP

cat > "$STAGE/DEBIAN/control" <<CONTROL
Package: base-audit-v3
Version: $DEB_VERSION
Section: office
Priority: optional
Architecture: $ARCH
Depends: libc6, libstdc++6, xdg-utils
Maintainer: V3 Project
Description: V3 base audit Flask application for UOS
 Frozen Flask application with bundled Python runtime and clean configuration templates.
CONTROL

echo '[DEB] 组装安装包 ...'
dpkg-deb --build --root-owner-group "$STAGE" "$DIST/$PACKAGE"
cd "$DIST"
sha256sum "$PACKAGE" > "$PACKAGE.sha256"
dpkg-deb --info "$PACKAGE" >/dev/null
echo "[完成] $DIST/$PACKAGE"
cat "$PACKAGE.sha256"
echo '[提示] 目标机无需安装 Python；需要兼容的 glibc/架构、桌面会话、xdg-utils 和 LibreOffice Calc。'
echo '[提示] 为兼容旧 UOS，请在最旧的目标系统上构建，再到所有目标系统实机验收。'
echo '[提示] 配置和运行数据保存在 ~/.local/share/base-audit-v3/，升级不会覆盖已有用户数据。'
