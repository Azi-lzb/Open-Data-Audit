#!/bin/sh
# 打包 UOS/麒麟 源码发行包：组装 + zip/tar.gz + SHA256。
# 产物：dist/审核工具_V<产品版本>_UOS_<日期>.zip（及 .sha256）
#
# 包内布局（app.py 同时兼容开发布局与打包布局）：
#   审核工具_V<产品版本>_UOS_<日期>/
#   ├─ app.py, run.py, requirements.txt, README.md, Linux-1-启动审核工具.sh
#   ├─ core/{src,frontend,基础数据审核工具使用说明.docx}   ← 业务核心与前端
#   └─ config/                                                ← 三册配置
#
# 配置策略：有 config-real/ 时默认打包真实配置；不存在时回退公开 config/。
#   携带真实配置的包仅供内部使用，禁止外发。
#
# 不打包：core/tests（开发基线）、core/data（本机运行数据）、Windows 专属文件
#         （build_exe.py、wheels/、*.bat、build/、dist/ 旧产物）。

set -e

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
ROOT="$(cd "$SCRIPT_DIR/.." && pwd)"
CORE="$ROOT/core"
DIST="$SCRIPT_DIR/dist"

WITH_LOCAL_CONFIG=0
[ "$1" = "--with-local-config" ] && WITH_LOCAL_CONFIG=1
if [ "$WITH_LOCAL_CONFIG" = "1" ] && [ ! -d "$ROOT/config-real" ]; then
    echo "[错误] --with-local-config 需要本机 config-real/。"; exit 1
fi
if [ -d "$ROOT/config-real" ]; then
    CONFIG_SOURCE="$ROOT/config-real"
    CONFIG_PROFILE=internal-config
else
    CONFIG_SOURCE="$ROOT/config"
    CONFIG_PROFILE=public-config
fi
[ -d "$CONFIG_SOURCE" ] || { echo "[错误] 未找到配置目录：$CONFIG_SOURCE"; exit 1; }

[ -d "$CORE/src/base_audit" ] || { echo "[错误] 未找到共享核心：$CORE"; exit 1; }
[ -f "$CORE/frontend/web/index.html" ] || { echo "[错误] 未找到共享前端 index.html"; exit 1; }
[ -f "$ROOT/版本管理/VERSION_MANIFEST.json" ] || { echo "[错误] 未找到版本清单：版本管理/VERSION_MANIFEST.json"; exit 1; }
[ -f "$ROOT/日志/更新日志.txt" ] || { echo "[错误] 未找到用户版更新日志：日志/更新日志.txt"; exit 1; }
[ -f "$ROOT/docs/常见问题.json" ] || { echo "[错误] 未找到常见问题：docs/常见问题.json"; exit 1; }
[ -s "$CORE/frontend/assets/app-icon.png" ] || { echo "[错误] 未找到应用图标：core/frontend/assets/app-icon.png"; exit 1; }
[ -s "$CORE/frontend/assets/app-icon.ico" ] || { echo "[错误] 未找到 favicon 图标：core/frontend/assets/app-icon.ico"; exit 1; }

PY=""
for candidate in python3 python; do
    if command -v "$candidate" >/dev/null 2>&1; then PY="$candidate"; break; fi
done
[ -n "$PY" ] || { echo "[错误] 未找到 python3，请先安装 Python 3。"; exit 1; }

IDENTITY_PYTHONPATH=$CORE/src${PYTHONPATH:+:$PYTHONPATH}
APP_VERSION=$(PYTHONPATH="$IDENTITY_PYTHONPATH" "$PY" -m base_audit.app_identity --version) || {
    echo "[错误] 无法从版本清单读取产品版本。"; exit 1;
}
if ! printf '%s\n' "$APP_VERSION" | grep -Eq '^[0-9]+\.[0-9]+\.[0-9]+\.[0-9]+$'; then
    echo "[错误] 版本清单中的产品版本不是 V年份.大版本.中版本.小版本：$APP_VERSION"; exit 1
fi

mkdir -p "$DIST"
PKG_NAME="审核工具_V${APP_VERSION}_UOS_$(date +%Y%m%d)"
PKG="$DIST/$PKG_NAME"
rm -rf "$PKG"
mkdir -p "$PKG"

# ---- 外壳文件（UOS 用 .sh 启动脚本；不带 Windows 的 bat/wheels/build_exe）----
cp "$SCRIPT_DIR/app.py" "$PKG/"
cp "$SCRIPT_DIR/run.py" "$PKG/"
cp "$SCRIPT_DIR/requirements.txt" "$PKG/"
cp "$SCRIPT_DIR/README.md" "$PKG/"
mkdir -p "$PKG/日志" "$PKG/docs"
cp "$ROOT/日志/更新日志.txt" "$PKG/日志/更新日志.txt"
cp "$ROOT/docs/常见问题.json" "$PKG/docs/常见问题.json"
cp "$SCRIPT_DIR/Linux-1-启动审核工具.sh" "$PKG/"
cp "$SCRIPT_DIR/Linux-4-打包UOS源码验收包.sh" "$SCRIPT_DIR/Linux-2-打包DEB与TARGZ.sh" "$PKG/"
cp "$SCRIPT_DIR/Linux-3-测试发行包启动.sh" "$PKG/"
chmod 755 "$PKG/Linux-1-启动审核工具.sh" "$PKG/Linux-4-打包UOS源码验收包.sh" "$PKG/Linux-2-打包DEB与TARGZ.sh" \
    "$PKG/Linux-3-测试发行包启动.sh"

# ---- 共享核心：后端 + 前端 + 使用说明 ----
mkdir -p "$PKG/core"
cp -r "$CORE/src" "$PKG/core/src"
cp -r "$CORE/frontend" "$PKG/core/frontend"
mkdir -p "$PKG/版本管理"
cp "$ROOT/版本管理/VERSION_MANIFEST.json" "$PKG/版本管理/VERSION_MANIFEST.json"
# 使用说明在仓库根目录（历史版本曾在 core/ 下，两处都找一次）。
for DOCX in "$ROOT/基础数据审核工具使用说明.docx" "$CORE/基础数据审核工具使用说明.docx"; do
    if [ -f "$DOCX" ]; then cp "$DOCX" "$PKG/core/"; break; fi
done
[ -f "$PKG/core/基础数据审核工具使用说明.docx" ] || echo "[提示] 未找到使用说明 docx，包内“使用说明”按钮将不可用"

# ---- 配置 ----
mkdir -p "$PKG/config"
if [ "$CONFIG_PROFILE" = "internal-config" ]; then
    echo "[警告] 携带真实配置（来源：${CONFIG_SOURCE##*/}），仅供内部使用，禁止外发。"
    cp -r "$CONFIG_SOURCE/." "$PKG/config/"
else
    cp "$CONFIG_SOURCE/"*.xlsx "$PKG/config/"
    echo "[配置] 已复制公开空表头工作簿。"
fi
printf '%s\n' "$CONFIG_PROFILE" > "$PKG/config-profile.txt"

find "$PKG" -name "__pycache__" -type d -exec rm -rf {} + 2>/dev/null || true
find "$PKG" -name "*.pyc" -delete 2>/dev/null || true

# ---- 压缩包 + SHA256 ----
cd "$DIST"
if command -v zip >/dev/null 2>&1; then
    zip -r -q "$PKG_NAME.zip" "$PKG_NAME"
else
    tar -czf "$PKG_NAME.tar.gz" "$PKG_NAME"
fi
ARTIFACT="$(ls -t "$PKG_NAME".zip "$PKG_NAME".tar.gz 2>/dev/null | head -1)"
if command -v sha256sum >/dev/null 2>&1; then
    sha256sum "$ARTIFACT" > "$ARTIFACT.sha256"
else
    "$PY" -c "import hashlib,sys;print(hashlib.sha256(open(sys.argv[1],'rb').read()).hexdigest()+'  '+sys.argv[1],end='')" "$ARTIFACT" > "$ARTIFACT.sha256"
fi

rm -rf -- "$PKG"   # 压缩与校验文件已生成，清理组装暂存目录
SIZE="$(du -sh "$ARTIFACT" | cut -f1)"
echo ""
echo "打包完成：dist/$ARTIFACT（$SIZE）"
cat "$ARTIFACT.sha256"
echo ""
[ "$CONFIG_PROFILE" = "internal-config" ] && echo "注意：本包携带真实配置，请勿外发。"
echo "目标机器要求：python3 + pip（pip3 install -r requirements.txt）、LibreOffice Calc。"
[ "$CONFIG_PROFILE" = "public-config" ] && echo "本包使用公开空表头配置；正式审核需导入机构真实配置。"
