#!/bin/sh
# 打包 UOS/麒麟 源码发行包：组装 + zip/tar.gz + SHA256。
# 产物：dist/基础数据审核工具_UOS_<日期>.zip（及 .sha256）
#
# 包内布局（app.py 同时兼容开发布局与打包布局）：
#   基础数据审核工具_UOS_<日期>/
#   ├─ app.py, run.py, requirements.txt, README.md, Linux-1-启动审核工具.sh
#   ├─ core/{src,frontend,基础数据审核工具使用说明.docx}   ← 业务核心与前端
#   └─ config/                                                ← 三册配置
#
# 配置策略：默认仅打包公开 config/ 中的空表头工作簿。
#   --with-local-config 显式读取本机 config-real/，仅供本人验证，禁止外发。
#
# 不打包：core/tests（开发基线）、core/data（本机运行数据）、Windows 专属文件
#         （build_exe.py、wheels/、*.bat、build/、dist/ 旧产物）。

set -e

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
ROOT="$(cd "$SCRIPT_DIR/.." && pwd)"
CORE="$ROOT/core"
DIST="$SCRIPT_DIR/dist"
VERSION="$(date +%Y%m%d)"
PKG_NAME="基础数据审核工具_UOS_$VERSION"
PKG="$DIST/$PKG_NAME"

WITH_LOCAL_CONFIG=0
[ "$1" = "--with-local-config" ] && WITH_LOCAL_CONFIG=1

[ -d "$CORE/src/base_audit" ] || { echo "[错误] 未找到共享核心：$CORE"; exit 1; }
[ -f "$CORE/frontend/web/index.html" ] || { echo "[错误] 未找到共享前端 index.html"; exit 1; }

PY=""
for candidate in python3 python; do
    if command -v "$candidate" >/dev/null 2>&1; then PY="$candidate"; break; fi
done
[ -n "$PY" ] || { echo "[错误] 未找到 python3，请先安装 Python 3。"; exit 1; }

rm -rf "$PKG"
mkdir -p "$PKG"

# ---- 外壳文件（UOS 用 .sh 启动脚本；不带 Windows 的 bat/wheels/build_exe）----
cp "$SCRIPT_DIR/app.py" "$PKG/"
cp "$SCRIPT_DIR/run.py" "$PKG/"
cp "$SCRIPT_DIR/requirements.txt" "$PKG/"
cp "$SCRIPT_DIR/README.md" "$PKG/"
cp "$SCRIPT_DIR/Linux-1-启动审核工具.sh" "$PKG/"
cp "$SCRIPT_DIR/Linux-4-打包UOS源码验收包.sh" "$SCRIPT_DIR/Linux-2-打包DEB与TARGZ.sh" "$PKG/"
cp "$SCRIPT_DIR/Linux-3-测试发行包启动.sh" "$PKG/"
chmod 755 "$PKG/Linux-1-启动审核工具.sh" "$PKG/Linux-4-打包UOS源码验收包.sh" "$PKG/Linux-2-打包DEB与TARGZ.sh" \
    "$PKG/Linux-3-测试发行包启动.sh"

# ---- 共享核心：后端 + 前端 + 使用说明 ----
mkdir -p "$PKG/core"
cp -r "$CORE/src" "$PKG/core/src"
cp -r "$CORE/frontend" "$PKG/core/frontend"
# 使用说明在仓库根目录（历史版本曾在 core/ 下，两处都找一次）。
for DOCX in "$ROOT/基础数据审核工具使用说明.docx" "$CORE/基础数据审核工具使用说明.docx"; do
    if [ -f "$DOCX" ]; then cp "$DOCX" "$PKG/core/"; break; fi
done
[ -f "$PKG/core/基础数据审核工具使用说明.docx" ] || echo "[提示] 未找到使用说明 docx，包内“使用说明”按钮将不可用"

# ---- 配置 ----
mkdir -p "$PKG/config"
if [ "$WITH_LOCAL_CONFIG" = "1" ]; then
    [ -d "$ROOT/config-real" ] || { echo "[错误] 未找到本机 config-real/"; exit 1; }
    echo "[警告] 携带本机 config-real/（含真实机构与历史），仅供本人验证，禁止外发。"
    cp -r "$ROOT/config-real/." "$PKG/config/"
else
    cp "$ROOT/config/"*.xlsx "$PKG/config/"
    echo "[配置] 已复制公开空表头工作簿。"
fi

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
[ "$WITH_LOCAL_CONFIG" = "1" ] && echo "注意：本包携带本机真实配置，请勿外发。"
echo "目标机器要求：python3 + pip（pip3 install -r requirements.txt）、LibreOffice Calc。"
echo "正式审核需要通过独立渠道提供真实配置。"
