#!/usr/bin/env bash
# Linux desktop entry point. It never installs packages or accesses the network.

set -u

SCRIPT_DIR=$(CDPATH= cd -- "$(dirname -- "${BASH_SOURCE[0]}")" && pwd)
RUNNER="$SCRIPT_DIR/run_full_test_linux.sh"

# Some Linux file managers launch shell scripts without a visible terminal.
# Reopen this script in the desktop's default terminal when one is available.
if [[ ! -t 1 && -n "${DISPLAY:-}" && "${ODA_LINUX_TEST_TERMINAL_CHILD:-0}" != 1 ]]; then
    if command -v x-terminal-emulator >/dev/null 2>&1; then
        ODA_LINUX_TEST_TERMINAL_CHILD=1 x-terminal-emulator -e /bin/bash "$SCRIPT_DIR/$(basename -- "${BASH_SOURCE[0]}")" "$@"
        exit $?
    fi
fi

usage() {
    cat <<'EOF'
用法：
  ./运行Linux全量测试.sh [--artifact 文件]... [--package-only|--core-only] [--no-pause]

默认行为：
  1. 自动识别当前 Linux 架构，从 shell-flask/dist、dist 或本目录查找最新的
     同架构发行包。仅在存在 dpkg-deb 时自动选择 DEB；其它系统使用 TAR.GZ。
  2. 找到完整源码、Python 3.8+ 测试解释器、pytest 和项目依赖时，运行 core 全量测试，
     并测试找到的发行包。
  3. 否则若找到发行包，直接用包内置的 Python 运行冻结程序启动冒烟，不要求目标机安装
     Python；报告会明确标记 core pytest 未执行。

本脚本不安装依赖、不访问网络。DEB 适用于有 dpkg-deb 的 Debian 系系统；其它 Linux
发行版可用同架构、且 glibc/系统运行库兼容的 TAR.GZ。发行包启动测试不受目标机 Python
版本影响。只有直接运行 core 源码 pytest 时，才需要 Python 3.8+ 和测试依赖；PyInstaller
包内的 Python 运行时供审核工具使用，并不是可直接调用的 pytest 解释器。

参数：
  --artifact 文件  指定一个 DEB 或 TAR.GZ，可重复；指定后不再自动查找发行包。
  --package-only   只测发行包启动；若未指定文件则自动查找同架构发行包。
  --core-only      只运行 core pytest，不查找或测试发行包。
  --no-pause       结束后不等待按回车，适合从终端或其他脚本调用。
  -h, --help       显示此帮助。

报告写入 test-output/full-test-reports/，名称包含 Linux 架构和时间戳。
EOF
}

MODE=auto
NO_PAUSE=0
EXPLICIT_ARTIFACTS=0
ARTIFACTS=()
while (($#)); do
    case "$1" in
        --artifact)
            if (($# < 2)); then
                echo "[错误] --artifact 后缺少文件路径。" >&2
                usage >&2
                exit 2
            fi
            ARTIFACTS+=("$2")
            EXPLICIT_ARTIFACTS=1
            shift 2
            ;;
        --package-only)
            [[ "$MODE" == auto || "$MODE" == package-only ]] || { echo '[错误] --package-only 和 --core-only 不能同时使用。' >&2; exit 2; }
            MODE=package-only
            shift
            ;;
        --core-only)
            [[ "$MODE" == auto || "$MODE" == core-only ]] || { echo '[错误] --package-only 和 --core-only 不能同时使用。' >&2; exit 2; }
            MODE=core-only
            shift
            ;;
        --no-pause)
            NO_PAUSE=1
            shift
            ;;
        -h|--help)
            usage
            exit 0
            ;;
        *)
            echo "[错误] 未知参数：$1" >&2
            usage >&2
            exit 2
            ;;
    esac
done

if [[ ! -f "$RUNNER" ]]; then
    echo "[错误] 找不到 Linux 测试引擎：$RUNNER" >&2
    exit 2
fi

ARCH_RAW=$(uname -m)
case "$ARCH_RAW" in
    aarch64|arm64)
        ARCH_LABEL=arm64
        ARTIFACT_GLOBS=("*_arm64.deb" "*_arm64.tar.gz" "*_aarch64.deb" "*_aarch64.tar.gz")
        ;;
    x86_64|amd64)
        ARCH_LABEL=amd64
        ARTIFACT_GLOBS=("*_amd64.deb" "*_amd64.tar.gz" "*_x86_64.deb" "*_x86_64.tar.gz")
        ;;
    *)
        ARCH_LABEL=$ARCH_RAW
        ARTIFACT_GLOBS=("*_${ARCH_RAW}.deb" "*_${ARCH_RAW}.tar.gz")
        ;;
esac

if ((EXPLICIT_ARTIFACTS)); then
    # Make relative paths stable after the underlying runner changes to the project root.
    CALLER_DIR=$PWD
    for i in "${!ARTIFACTS[@]}"; do
        case "${ARTIFACTS[$i]}" in
            /*) ;;
            *) ARTIFACTS[$i]="$CALLER_DIR/${ARTIFACTS[$i]}" ;;
        esac
    done
elif [[ "$MODE" != core-only ]]; then
    SEARCH_DIRS=("$SCRIPT_DIR/shell-flask/dist" "$SCRIPT_DIR/dist" "$SCRIPT_DIR")
    EXTENSIONS=(tar.gz)
    if command -v dpkg-deb >/dev/null 2>&1; then
        EXTENSIONS=(deb tar.gz)
    else
        echo '[提示] 未找到 dpkg-deb；自动发现时跳过 DEB，仅查找 TAR.GZ。'
    fi
    shopt -s nullglob
    for extension in "${EXTENSIONS[@]}"; do
        newest=""
        for search_dir in "${SEARCH_DIRS[@]}"; do
            [[ -d "$search_dir" ]] || continue
            for pattern in "${ARTIFACT_GLOBS[@]}"; do
                [[ "$pattern" == *".$extension" ]] || continue
                candidate=$(find "$search_dir" -type f -name "$pattern" -exec ls -1t -- {} + 2>/dev/null | head -n 1)
                if [[ -n "$candidate" ]]; then
                    newest=$candidate
                    break
                fi
            done
            [[ -n "$newest" ]] && break
        done
        [[ -n "$newest" ]] && ARTIFACTS+=("$newest")
    done
fi

if ((${#ARTIFACTS[@]})); then
    echo "[信息] Linux 架构：$ARCH_RAW（$ARCH_LABEL）"
    echo '[信息] 将测试以下本地发行包：'
    printf '  %s\n' "${ARTIFACTS[@]}"
else
    echo "[信息] Linux 架构：$ARCH_RAW（$ARCH_LABEL）；没有选择发行包。"
fi

PYTHON_BIN=""
SOURCE_SUITE_PRESENT=1
if [[ ! -f "$SCRIPT_DIR/core/tests/conftest.py" || ! -f "$SCRIPT_DIR/core/pytest.ini" || ! -d "$SCRIPT_DIR/core/src" ]]; then
    SOURCE_SUITE_PRESENT=0
fi
if [[ "$MODE" != package-only ]]; then
    PYTHON_CANDIDATES=()
    [[ -n "${ODA_TEST_PYTHON:-}" ]] && PYTHON_CANDIDATES+=("$ODA_TEST_PYTHON")
    PYTHON_CANDIDATES+=("$SCRIPT_DIR/.venv/bin/python")
    PYTHON_CANDIDATES+=("$SCRIPT_DIR/bin/python")
    [[ -n "${HOME:-}" ]] && PYTHON_CANDIDATES+=("$HOME/.venvs/open-data-audit-test/bin/python")
    PYTHON_CANDIDATES+=(python3 python)
    for candidate in "${PYTHON_CANDIDATES[@]}"; do
        if [[ "$candidate" == */* ]]; then
            [[ -x "$candidate" ]] || continue
        else
            command -v "$candidate" >/dev/null 2>&1 || continue
        fi
        if "$candidate" -c 'import sys, pytest, flask, openpyxl, xlrd, lxml; raise SystemExit(0 if sys.version_info >= (3, 8) else 1)' >/dev/null 2>&1; then
            PYTHON_BIN=$candidate
            break
        fi
    done
fi

if [[ "$MODE" == auto ]]; then
    if [[ -n "$PYTHON_BIN" && "$SOURCE_SUITE_PRESENT" == 1 ]]; then
        MODE=full
    elif ((${#ARTIFACTS[@]})); then
        MODE=package-only
        if [[ "$SOURCE_SUITE_PRESENT" != 1 ]]; then
            echo '[提示] 当前目录没有 core 测试源码；本次只运行发行包启动冒烟，core pytest 会标记为未执行。'
        else
            echo '[提示] 没有可用的 Python 3.8+ 源码测试环境；发行包使用自身内置 Python，本次直接运行启动冒烟，core pytest 会标记为未执行。'
        fi
    else
        echo '[错误] 没有可运行的 core 测试环境，也没有同架构发行包可做启动冒烟。' >&2
        echo '请把 Linux TAR.GZ（或 Debian 系可用的 DEB）放入 shell-flask/dist、dist 或本目录后重试；core pytest 需预先准备兼容的离线测试环境。' >&2
        exit 2
    fi
fi

if [[ "$MODE" == core-only && ( -z "$PYTHON_BIN" || "$SOURCE_SUITE_PRESENT" != 1 ) ]]; then
    echo '[错误] core 全量测试需要项目 core/tests 源码，以及 Python 3.8+、pytest、Flask、openpyxl、xlrd 和 lxml。脚本不会联网安装依赖。' >&2
    exit 2
fi
if [[ "$MODE" == package-only && ${#ARTIFACTS[@]} -eq 0 ]]; then
    echo '[错误] package-only 模式没有找到发行包；请放入同架构 DEB/TAR.GZ，或用 --artifact 指定。' >&2
    exit 2
fi
if [[ "$MODE" == package-only ]]; then
    echo '[信息] 本次只测试冻结发行包；使用包内 Python，不检查目标机 Python 版本。'
fi

RUN_ARGS=()
[[ "$MODE" == package-only ]] && RUN_ARGS+=(--package-only)
if [[ "$MODE" != core-only ]]; then
    for artifact in "${ARTIFACTS[@]}"; do
        RUN_ARGS+=(--artifact "$artifact")
    done
fi

ODA_TEST_EXECUTION_MODE="Native Linux offline/local test launch" \
ODA_TEST_PYTHON="$PYTHON_BIN" \
    bash "$RUNNER" "${RUN_ARGS[@]}"
RUN_RC=$?

if ((NO_PAUSE == 0)) && [[ -t 0 ]]; then
    echo
    read -r -p '测试结束，按回车关闭窗口。' _
fi
exit "$RUN_RC"
