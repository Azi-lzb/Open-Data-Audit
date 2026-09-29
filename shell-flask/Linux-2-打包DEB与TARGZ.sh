#!/bin/sh
# 统一 Linux/UOS 打包：PyInstaller 冻结（自带 CPython 运行时）→ DEB + tar.gz。
#
# 兼容性策略：默认不使用构建机系统 Python，而是用 python-build-standalone 的
# 自包含 CPython（按 glibc 2.17 基线构建）冻结，使在一台新系统上构建的包也能
# 运行在更老的目标系统（UOS20 glibc 2.28、麒麟 V10、UOS25 glibc 2.38）。
# 构建完成后逐个 ELF 审计 GLIBC/GLIBCXX 符号版本，实测结果写入文件名与
# build-info.txt；超过目标基线（默认 glibc 2.28 / GLIBCXX 3.4.25，即 UOS20）
# 时构建失败并列出超标文件。
#
# 首次构建需要网络（下载解释器与依赖）；解释器压缩包缓存在 shell-flask/runtime/linux/，
# 之后离线可重建。目标机不需要 Python，仍需 LibreOffice Calc。
set -eu

SCRIPT_DIR=$(CDPATH= cd -- "$(dirname -- "$0")" && pwd)
ROOT=$(CDPATH= cd -- "$SCRIPT_DIR/.." && pwd)
DIST=${FLASK_DIST_DIR:-$SCRIPT_DIR/dist}
RELEASE=/opt/base-audit-v3

# 固定解释器与工具版本，保证可复现；需要升级时改这里或用环境变量覆盖。
STANDALONE_TAG=${STANDALONE_TAG:-20260924}
STANDALONE_PY_VERSION=${STANDALONE_PY_VERSION:-3.12.14}
PYINSTALLER_PIN=${PYINSTALLER_PIN:-6.22.3}
# 国内网络优先走南大镜像，失败再回 GitHub 官方。
STANDALONE_MIRROR=${STANDALONE_MIRROR:-https://mirror.nju.edu.cn/github-release/astral-sh/python-build-standalone}
STANDALONE_FALLBACK=${STANDALONE_FALLBACK:-https://github.com/astral-sh/python-build-standalone/releases/download}
# 兼容目标基线：UOS20（Debian10，glibc 2.28 / gcc8 libstdc++ GLIBCXX 3.4.25）。
GLIBC_TARGET=${GLIBC_TARGET:-2.28}
GLIBCXX_TARGET=${GLIBCXX_TARGET:-3.4.25}

PYTHON_OVERRIDE=
WITH_LOCAL_CONFIG=0
SKIP_DEB=0
SKIP_TGZ=0

usage() {
    cat <<EOF
用法：
  ./Linux-2-打包DEB与TARGZ.sh [--with-local-config] [--python 解释器]
                       [--skip-deb] [--skip-tgz]

默认生成 DEB 与 tar.gz 两种发行包，均自带冻结 Python 运行时，目标机无需
安装 Python。tar.gz 解压即可运行（双击解压目录 shell-flask/ 里的
Linux-1-启动审核工具.desktop，或运行 shell-flask/Linux-1-启动审核工具.sh）。

--with-local-config  仅自测：把本机 config-real/（真实配置）原样带入包内，禁止外发；
                      默认使用公开 config/ 空表头配置，构建时同样做隐私清理校验。
--python 解释器      改用指定 Python（如系统 python3.12）构建。注意：系统
                      解释器按构建机 glibc 编译，产物可能无法在旧 UOS 运行；
                      脚本仍会做符号审计并在超标时报错。
--skip-deb / --skip-tgz  只生成其中一种包。

配套入口（推荐在每台打包机上按此流程）：
  Linux-3-测试发行包启动.sh  打包后在本机实测产物能否启动（解包→启动→HTTP 校验）

跨机器/芯片：amd64 与 arm64（飞腾/鲲鹏）构建机会自动下载对应架构的自包含解释器，
产物文件名带架构标记；其它架构回退该机系统 Python。适用 UOS20/UOS25、麒麟 V10
等 Debian 系（deb 安装或 tar.gz 免安装均可）。
环境变量：STANDALONE_TAG、STANDALONE_PY_VERSION、PYINSTALLER_PIN、
          STANDALONE_MIRROR、PIP_INDEX_URL、PIP_MIRROR_FALLBACK、
          GLIBC_TARGET、GLIBCXX_TARGET、FLASK_DIST_DIR。
依赖：dpkg-deb、dpkg、tar、objdump（binutils）、sha256sum、curl 或 wget。
EOF
}

for arg in "$@"; do
    case "$arg" in
        --with-local-config) WITH_LOCAL_CONFIG=1 ;;
        --skip-deb) SKIP_DEB=1 ;;
        --skip-tgz) SKIP_TGZ=1 ;;
        --python) : ;;
        *)
            case "${PREV_ARG:-}" in
                --python) PYTHON_OVERRIDE=$arg ;;
                *) echo "[错误] 未知参数：$arg" >&2; usage >&2; exit 2 ;;
            esac
            ;;
    esac
    PREV_ARG=$arg
done
[ "$SKIP_DEB" -eq 0 ] || [ "$SKIP_TGZ" -eq 0 ] || { echo '[错误] --skip-deb 与 --skip-tgz 不能同时使用。' >&2; exit 2; }

for tool in dpkg-deb dpkg mktemp sha256sum tar objdump; do
    command -v "$tool" >/dev/null 2>&1 || { echo "[错误] 缺少 $tool，请在构建机安装。" >&2; exit 1; }
done
FETCH=
for tool in curl wget; do
    if command -v "$tool" >/dev/null 2>&1; then FETCH=$tool; break; fi
done
[ -d "$ROOT/core/src/base_audit" ] || { echo "[错误] 缺少共享核心：$ROOT/core/src/base_audit" >&2; exit 1; }
[ -f "$ROOT/core/frontend/web/index.html" ] || { echo '[错误] 缺少前端 index.html。' >&2; exit 1; }
if [ ! -f "$ROOT/基础数据审核工具使用说明.docx" ] && [ ! -f "$ROOT/core/基础数据审核工具使用说明.docx" ]; then
    echo '[提示] 未提供使用说明 DOCX；继续构建，发行包不含该文档。'
fi
[ -d "$ROOT/config" ] || { echo '[错误] 缺少公开 config/。' >&2; exit 1; }
if [ "$WITH_LOCAL_CONFIG" = 1 ] && [ ! -d "$ROOT/config-real" ]; then
    echo '[错误] --with-local-config 需要本机 config-real/。' >&2
    exit 1
fi

ARCH=$(dpkg --print-architecture)
case "$ARCH" in
    amd64) STANDALONE_ARCH=x86_64 ;;
    arm64) STANDALONE_ARCH=aarch64 ;;
    *) STANDALONE_ARCH= ;;
esac
BUILD_OS=linux
if [ -r /etc/os-release ]; then
    . /etc/os-release
    BUILD_OS=${ID:-linux}-${VERSION_ID:-unknown}
fi
BUILD_OS=$(printf '%s' "$BUILD_OS" | tr -c 'A-Za-z0-9._-' '_')

CACHE_HOME=${XDG_CACHE_HOME:-$HOME/.cache}
STANDALONE_CACHE=$CACHE_HOME/base-audit-v3/standalone
RUNTIME_TARBALL_DIR=$SCRIPT_DIR/runtime/linux

# ---------------------------------------------------------------------------
# 解释器解析：--python / 环境变量 > 本地缓存的 standalone > 下载 standalone。
# ---------------------------------------------------------------------------
BUILD_PY=
INTERPRETER_SOURCE=
if [ -n "$PYTHON_OVERRIDE" ]; then
    command -v "$PYTHON_OVERRIDE" >/dev/null 2>&1 || [ -x "$PYTHON_OVERRIDE" ] || {
        echo "[错误] 指定的解释器不存在：$PYTHON_OVERRIDE" >&2; exit 1;
    }
    BUILD_PY=$(command -v "$PYTHON_OVERRIDE" 2>/dev/null || printf '%s' "$PYTHON_OVERRIDE")
    INTERPRETER_SOURCE=explicit
else
    if [ -z "$STANDALONE_ARCH" ]; then
        # 龙芯/申威等暂无 python-build-standalone 预编译包的架构：回退本机系统
        # Python（如 amd64/arm64 之外的平台），glibc 兼容性交给构建后审计判断。
        echo "[警告] 架构 $ARCH 无 python-build-standalone 预编译解释器，回退构建机系统 Python。" >&2
        echo '       产物 glibc 兼容性以构建后 ELF 审计为准；更可控的做法是用 --python 指定解释器。' >&2
        for candidate in python3.12 python3.11 python3.10 python3 python; do
            if command -v "$candidate" >/dev/null 2>&1 && "$candidate" -c 'import sys; raise SystemExit(0 if sys.version_info >= (3, 8) else 1)'; then
                BUILD_PY=$(command -v "$candidate")
                INTERPRETER_SOURCE=system-fallback
                break
            fi
        done
        [ -n "$BUILD_PY" ] || { echo '[错误] 构建机系统也没有 Python 3.8+；请安装后重试。' >&2; exit 1; }
    fi
    ASSET="cpython-${STANDALONE_PY_VERSION}+${STANDALONE_TAG}-${STANDALONE_ARCH}-unknown-linux-gnu-install_only.tar.gz"
    EXTRACT_DIR=$STANDALONE_CACHE/${ASSET%.tar.gz}
    if [ -n "$STANDALONE_ARCH" ] && [ -x "$EXTRACT_DIR/python/bin/python3" ]; then
        BUILD_PY=$EXTRACT_DIR/python/bin/python3
        INTERPRETER_SOURCE=cache
    elif [ -n "$STANDALONE_ARCH" ]; then
        TARBALL=$RUNTIME_TARBALL_DIR/$ASSET
        if [ ! -f "$TARBALL" ]; then
            echo "[解释器] 下载 python-build-standalone CPython ${STANDALONE_PY_VERSION}（glibc 2.17 基线）..."
            mkdir -p "$RUNTIME_TARBALL_DIR"
            ASSET_URL_ENC=$(printf '%s' "$ASSET" | sed 's/+/%2B/g')
            URL1=$STANDALONE_MIRROR/$STANDALONE_TAG/$ASSET_URL_ENC
            URL2=$STANDALONE_FALLBACK/$STANDALONE_TAG/$ASSET
            DOWNLOADED=0
            for url in "$URL1" "$URL2"; do
                echo "[解释器] 尝试 $url"
                if [ "$FETCH" = curl ]; then
                    curl -fL --retry 2 --connect-timeout 30 -o "$TARBALL.part" "$url" && DOWNLOADED=1
                elif [ "$FETCH" = wget ]; then
                    wget --tries=2 --timeout=60 -O "$TARBALL.part" "$url" && DOWNLOADED=1
                fi
                [ "$DOWNLOADED" -eq 1 ] && break
            done
            [ "$DOWNLOADED" -eq 1 ] || { echo '[错误] 解释器下载失败；可手工下载后放到' >&2; echo "        $TARBALL" >&2; exit 1; }
            mv "$TARBALL.part" "$TARBALL"
        fi
        echo '[解释器] 解压自包含 Python 到用户缓存 ...'
        mkdir -p "$STANDALONE_CACHE"
        rm -rf -- "$EXTRACT_DIR"
        mkdir -p "$EXTRACT_DIR"
        tar -xzf "$TARBALL" -C "$EXTRACT_DIR"
        BUILD_PY=$EXTRACT_DIR/python/bin/python3
        INTERPRETER_SOURCE=download
    fi
fi
"$BUILD_PY" -c 'import sys; raise SystemExit(0 if sys.version_info >= (3, 8) else 1)' || {
    echo "[错误] 构建解释器需要 Python 3.8+：$BUILD_PY" >&2; exit 1;
}
PY_FULL_VERSION=$("$BUILD_PY" -c 'import sys; print("%d.%d.%d" % sys.version_info[:3])')
PY_MINOR=$("$BUILD_PY" -c 'import sys; print("%d%d" % sys.version_info[:2])')
echo "[解释器] $BUILD_PY（Python $PY_FULL_VERSION，来源 $INTERPRETER_SOURCE）"
echo "[架构] $ARCH${STANDALONE_ARCH:+（standalone 目标 $STANDALONE_ARCH）}"

# ---------------------------------------------------------------------------
# 构建虚拟环境（独立于系统 Python 的第三方包，可复用）。
# ---------------------------------------------------------------------------
BUILD_ENV=$CACHE_HOME/base-audit-v3/buildenv-linux-${ARCH}-py${PY_MINOR}
VENV_EXPECT=$("$BUILD_PY" -c 'import sys; print("%d.%d" % sys.version_info[:2])')
# 环境不存在或版本不符（例如换过解释器小版本）时重建。
if [ ! -x "$BUILD_ENV/bin/python" ] || \
   [ "$("$BUILD_ENV/bin/python" -c 'import sys; print("%d.%d" % sys.version_info[:2])' 2>/dev/null || true)" != "$VENV_EXPECT" ]; then
    rm -rf -- "$BUILD_ENV"
    "$BUILD_PY" -m venv "$BUILD_ENV" || { echo '[错误] 创建构建虚拟环境失败。' >&2; exit 1; }
fi
BUILD_PY=$BUILD_ENV/bin/python

echo '[依赖] 安装 Flask 运行库和 PyInstaller（首次需网络）...'
# 官方 PyPI 直连失败时自动改用国内镜像；设 PIP_MIRROR_FALLBACK=none 可禁用。
PIP_MIRROR_FALLBACK=${PIP_MIRROR_FALLBACK:-https://pypi.tuna.tsinghua.edu.cn/simple}
if "$BUILD_PY" -m pip install --disable-pip-version-check --quiet \
        -r "$SCRIPT_DIR/requirements.txt" "pyinstaller==$PYINSTALLER_PIN"; then
    :
elif [ "$PIP_MIRROR_FALLBACK" != none ]; then
    echo "[依赖] 官方 PyPI 直连失败，改用镜像重试：$PIP_MIRROR_FALLBACK"
    "$BUILD_PY" -m pip install --disable-pip-version-check --quiet \
        -i "$PIP_MIRROR_FALLBACK" \
        -r "$SCRIPT_DIR/requirements.txt" "pyinstaller==$PYINSTALLER_PIN"
else
    echo '[错误] 依赖安装失败。' >&2
    exit 1
fi
"$BUILD_PY" -c 'import flask, openpyxl, xlrd, lxml, PyInstaller; print("[依赖] 导入检查通过")'

# ---------------------------------------------------------------------------
# PyInstaller 冻结（onedir，与 DEB 内布局一致）。
# ---------------------------------------------------------------------------
STAGE=$(mktemp -d "${TMPDIR:-/tmp}/base-audit-linux.XXXXXXXX")
trap 'rm -rf -- "$STAGE"' EXIT HUP INT TERM
PAYLOAD=$STAGE$RELEASE
PYI_DIST=$STAGE/pyinstaller-dist
mkdir -p "$PAYLOAD/app" "$PAYLOAD/shell-flask" "$PAYLOAD/core/frontend" \
    "$PAYLOAD/config/默认配置" "$STAGE/DEBIAN" "$STAGE/usr/bin" \
    "$STAGE/usr/share/applications" "$DIST" "$PYI_DIST"

echo "[PyInstaller] 冻结 Flask、共享核心和 Python $VENV_EXPECT 运行时 ..."
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
    TK_STATUS='随包（含 Tcl/Tk）'
else
    echo '[提示] 构建 Python 未带 Tk；文件选择将退回“浏览器内置”模式。'
    TK_STATUS='缺失（浏览器内置模式）'
fi
set -- "$@" "$SCRIPT_DIR/run.py"
"$BUILD_PY" -m PyInstaller "$@" >"$STAGE/pyinstaller.log" 2>&1 || {
    echo '[错误] PyInstaller 构建失败，最后 40 行日志：' >&2
    tail -40 "$STAGE/pyinstaller.log" >&2 || true
    exit 1
}
[ -x "$PYI_DIST/base-audit-v3/base-audit-v3" ] || { echo '[错误] PyInstaller 未生成 Linux 启动程序。' >&2; exit 1; }

# ---------------------------------------------------------------------------
# glibc 兼容性审计：扫描冻结目录全部 ELF 的 GLIBC/GLIBCXX 符号版本需求。
# ---------------------------------------------------------------------------
echo '[审计] 扫描冻结产物 ELF 符号版本（GLIBC / GLIBCXX）...'
FROZEN_DIR=$PYI_DIST/base-audit-v3
max_symbol_version() {  # $1=目录 $2=前缀(GLIBC/GLIBCXX)；输出最大版本号
    find "$1" -type f -exec objdump -T {} + 2>/dev/null || true
}
ALL_SYMBOLS=$(max_symbol_version "$FROZEN_DIR" all)
MIN_GLIBC=$(printf '%s\n' "$ALL_SYMBOLS" | grep -oE 'GLIBC_[0-9]+(\.[0-9]+)?' | sed 's/^GLIBC_//' | sort -uV | tail -1)
MIN_GLIBCXX=$(printf '%s\n' "$ALL_SYMBOLS" | grep -oE 'GLIBCXX_[0-9]+(\.[0-9]+)+' | sed 's/^GLIBCXX_//' | sort -uV | tail -1)
[ -n "$MIN_GLIBC" ] || { echo '[错误] 未扫描到任何 GLIBC 符号信息，无法审计。' >&2; exit 1; }
echo "[审计] 实测最低 glibc 要求：${MIN_GLIBC}；GLIBCXX：${MIN_GLIBCXX:-无}"

version_gt() {  # $1 > $2 ?
    [ "$(printf '%s\n%s\n' "$1" "$2" | sort -V | tail -1)" = "$1" ] && [ "$1" != "$2" ]
}
AUDIT_FAILED=0
if version_gt "$MIN_GLIBC" "$GLIBC_TARGET"; then AUDIT_FAILED=1; fi
if [ -n "$MIN_GLIBCXX" ] && version_gt "$MIN_GLIBCXX" "$GLIBCXX_TARGET"; then AUDIT_FAILED=1; fi
if [ "$AUDIT_FAILED" -eq 1 ]; then
    echo "[错误] 冻结产物符号版本超出目标基线（glibc<=$GLIBC_TARGET / GLIBCXX<=$GLIBCXX_TARGET）。" >&2
    echo '       超标文件（含版本号）：' >&2
    find "$FROZEN_DIR" -type f | while IFS= read -r f; do
        syms=$(objdump -T "$f" 2>/dev/null || true)
        g=$(printf '%s\n' "$syms" | grep -oE 'GLIBC_[0-9]+(\.[0-9]+)?' | sed 's/^GLIBC_//' | sort -uV | tail -1)
        gx=$(printf '%s\n' "$syms" | grep -oE 'GLIBCXX_[0-9]+(\.[0-9]+)+' | sed 's/^GLIBCXX_//' | sort -uV | tail -1)
        bad=0
        { [ -n "$g" ] && version_gt "$g" "$GLIBC_TARGET"; } && bad=1
        { [ -n "$gx" ] && version_gt "$gx" "$GLIBCXX_TARGET"; } && bad=1
        if [ "$bad" -eq 1 ]; then
            echo "       ${f#$STAGE/}  GLIBC<=$g GLIBCXX<=${gx:-无}" >&2
        fi
    done
    echo "       可升级 GLIBC_TARGET/GLIBCXX_TARGET 放宽基线，或改用更低基线构建的依赖版本。" >&2
    exit 1
fi

# ---------------------------------------------------------------------------
# 组装发行内容（DEB 与 tar.gz 共用同一 payload）。
# ---------------------------------------------------------------------------
cp -R "$FROZEN_DIR/." "$PAYLOAD/app/"
cp -R "$ROOT/core/frontend/." "$PAYLOAD/core/frontend/"
for doc in "$ROOT/基础数据审核工具使用说明.docx" "$ROOT/core/基础数据审核工具使用说明.docx"; do
    if [ -f "$doc" ]; then cp "$doc" "$PAYLOAD/core/"; break; fi
done
cp "$SCRIPT_DIR/Linux-1-启动审核工具.sh" "$PAYLOAD/shell-flask/"
cp "$SCRIPT_DIR/Linux-1-启动审核工具.desktop" "$PAYLOAD/shell-flask/"

if [ "$WITH_LOCAL_CONFIG" = 1 ]; then
    echo '[警告] 本包带有本机 config-real/，可能含机构资料和审核历史，仅供本人验证，禁止外发。'
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

DEB_VERSION=$(date +%Y.%m.%d.%H%M%S)
ARTIFACT_BASE=base-audit-v3_${DEB_VERSION}_${BUILD_OS}_py${PY_FULL_VERSION}_glibc${MIN_GLIBC}_${ARCH}

cat > "$STAGE/usr/bin/base-audit-v3" <<SH
#!/bin/sh
set -eu
BASE_AUDIT_INSTALL_ROOT=$RELEASE
export BASE_AUDIT_INSTALL_ROOT
exec "$RELEASE/shell-flask/Linux-1-启动审核工具.sh" "\$@"
SH

cat > "$STAGE/usr/share/applications/base-audit-v3.desktop" <<DESKTOP
[Desktop Entry]
Version=1.0
Type=Application
Name=基础数据审核工具 V3
Comment=后台启动本机 Flask 审核工具
Exec=base-audit-v3 --background
Terminal=false
Categories=Office;
DESKTOP

cat > "$STAGE/DEBIAN/control" <<CONTROL
Package: base-audit-v3
Version: $DEB_VERSION
Section: office
Priority: optional
Architecture: $ARCH
Depends: libc6 (>= $MIN_GLIBC), libcrypt1, libstdc++6, xdg-utils, iproute2
Maintainer: V3 Project
Description: V3 base audit Flask application for UOS
 Frozen Flask application with bundled Python runtime and clean configuration templates.
CONTROL

"$BUILD_PY" -m pip freeze > "$PAYLOAD/dependency-versions.txt"
GLIBC_VERSION=$(getconf GNU_LIBC_VERSION 2>/dev/null | sed 's/^glibc[[:space:]]*//' || true)
[ -n "$GLIBC_VERSION" ] || GLIBC_VERSION=unknown
{
    printf 'Package base: %s\n' "$ARTIFACT_BASE"
    printf 'Bundled Python: %s (python-build-standalone, %s)\n' "$PY_FULL_VERSION" "$INTERPRETER_SOURCE"
    printf 'Architecture: %s\nTk: %s\n' "$ARCH" "$TK_STATUS"
    printf 'PyInstaller: %s\n' "$PYINSTALLER_PIN"
    printf 'Build OS: %s (%s)\nBuild machine glibc: %s\n' "${PRETTY_NAME:-unknown}" "$BUILD_OS" "$GLIBC_VERSION"
    printf 'Audited min glibc of frozen app: %s (target <= %s)\n' "$MIN_GLIBC" "$GLIBC_TARGET"
    printf 'Audited max GLIBCXX of frozen app: %s (target <= %s)\n' "${MIN_GLIBCXX:-none}" "$GLIBCXX_TARGET"
    printf '\nCompatibility: runs on Linux %s with glibc >= %s (e.g. UOS20 2.28, UOS25 2.38, Kylin V10).\n' "$ARCH" "$MIN_GLIBC"
    printf '\nFrozen build environment (pip freeze):\n'
    cat "$PAYLOAD/dependency-versions.txt"
} > "$PAYLOAD/build-info.txt"

# ZIP 源目录可能没有遍历权限；统一目录与文件权限，并保证非 root 能启动。
find "$STAGE" -type d -exec chmod 755 {} +
find "$STAGE" -type f -exec chmod 644 {} +
chmod 755 "$STAGE/usr/bin/base-audit-v3" \
    "$PAYLOAD/app/base-audit-v3" \
    "$PAYLOAD/shell-flask/Linux-1-启动审核工具.sh" \
    "$PAYLOAD/shell-flask/Linux-1-启动审核工具.desktop"

FINISHED=

if [ "$SKIP_DEB" -eq 0 ]; then
    echo '[DEB] 组装安装包 ...'
    DEB_OUT=$DIST/${ARTIFACT_BASE}.deb
    dpkg-deb --build --root-owner-group "$STAGE" "$DEB_OUT"
    (cd "$DIST" && sha256sum "${DEB_OUT##*/}" > "${DEB_OUT##*/}.sha256")
    dpkg-deb --info "$DEB_OUT" >/dev/null
    cp "$PAYLOAD/build-info.txt" "$DIST/${ARTIFACT_BASE}.deb.build-info.txt"
    FINISHED="$FINISHED
$DEB_OUT"
fi

if [ "$SKIP_TGZ" -eq 0 ]; then
    echo '[TAR.GZ] 组装便携包 ...'
    TGZ_OUT=$DIST/${ARTIFACT_BASE}.tar.gz
    # 只打包 /opt/base-audit-v3 payload，不带 DEBIAN/ 与 usr/；根目录带版本号。
    tar --numeric-owner --owner=0 --group=0 \
        --transform "s|^base-audit-v3|${ARTIFACT_BASE}|" \
        -C "$STAGE/opt" -czf "$TGZ_OUT" base-audit-v3
    (cd "$DIST" && sha256sum "${TGZ_OUT##*/}" > "${TGZ_OUT##*/}.sha256")
    cp "$PAYLOAD/build-info.txt" "$DIST/${ARTIFACT_BASE}.tar.gz.build-info.txt"
    FINISHED="$FINISHED
$TGZ_OUT"
fi

echo
echo "[完成] 构建产物：$FINISHED"
echo "[审计] 冻结包实测最低 glibc ${MIN_GLIBC}（目标基线 ${GLIBC_TARGET}）；GLIBCXX ${MIN_GLIBCXX:-无}。"
echo '[提示] 目标机无需 Python；DEB 装到 /opt/base-audit-v3 并加菜单入口；tar.gz 解压后运行其中'
echo '       shell-flask/Linux-1-启动审核工具.sh 或双击同目录 Linux-1-启动审核工具.desktop。'
echo '[提示] 目标机仍需桌面会话、xdg-utils、iproute2 和 LibreOffice Calc。'
echo '[提示] 配置和运行数据保存在 ~/.local/share/base-audit-v3/，升级不覆盖已有用户数据。'
for f in "$DIST/${ARTIFACT_BASE}".*; do echo "[留档] $f"; done
