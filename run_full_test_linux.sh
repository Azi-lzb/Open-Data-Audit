#!/usr/bin/env bash
# Linux full pytest runner for native amd64 and arm64 hosts.
# Optional --artifact paths are checked with the existing packaged-app smoke test.

set -uo pipefail

ROOT_DIR=$(CDPATH= cd -- "$(dirname -- "${BASH_SOURCE[0]}")" && pwd)
cd "$ROOT_DIR" || exit 2

usage() {
    cat <<'EOF'
Usage:
  bash run_full_test_linux.sh [--artifact PATH]...
  bash run_full_test_linux.sh --package-only --artifact PATH [--artifact PATH]...

Runs the complete core pytest suite and writes a timestamped report whose name
includes the detected Linux architecture. Repeat the command natively on Linux
amd64 and arm64 to produce separate reports. Python 3.8+ with project runtime
dependencies and pytest is required only for the source test suite. The packaged
application includes its own Python runtime, so package-only startup smoke does
not require or select a host Python version.

Optional --artifact PATH may be repeated to also run the existing packaged-app
startup smoke test against one or more .deb or .tar.gz files.
Use --package-only on a deployment machine that cannot run the source suite
(for example, Kylin V10 with Python 3.5). DEB extraction needs dpkg-deb;
compatible glibc Linux systems can use the matching-architecture .tar.gz package.
Alpine/musl and mismatched CPU architectures or runtime libraries are not covered.
EOF
}

ARTIFACTS=()
PACKAGE_ONLY=0
while (($#)); do
    case "$1" in
        --artifact)
            if (($# < 2)); then
                echo "[ERROR] --artifact requires a file path." >&2
                usage >&2
                exit 2
            fi
            ARTIFACTS+=("$2")
            shift 2
            ;;
        --package-only)
            PACKAGE_ONLY=1
            shift
            ;;
        -h|--help)
            usage
            exit 0
            ;;
        *)
            echo "[ERROR] Unknown argument: $1" >&2
            usage >&2
            exit 2
            ;;
    esac
done
if ((PACKAGE_ONLY)) && ((${#ARTIFACTS[@]} == 0)); then
    echo "[ERROR] --package-only requires at least one --artifact PATH." >&2
    usage >&2
    exit 2
fi

if [[ "$(uname -s)" != "Linux" ]]; then
    echo "[ERROR] This runner must be executed on Linux." >&2
    exit 2
fi

ARCH_RAW=$(uname -m)
case "$ARCH_RAW" in
    x86_64|amd64) ARCH_LABEL=amd64 ;;
    aarch64|arm64) ARCH_LABEL=arm64 ;;
    *) ARCH_LABEL=$ARCH_RAW ;;
esac

REPORT_DIR="$ROOT_DIR/test-output/full-test-reports"
if ! mkdir -p "$REPORT_DIR"; then
    echo "[ERROR] Could not create report directory: $REPORT_DIR" >&2
    exit 2
fi
RUN_STAMP=$(date '+%Y%m%d_%H%M%S')
REPORT_FILE="$REPORT_DIR/FullTest_linux-${ARCH_LABEL}_${RUN_STAMP}.txt"

PYTHON_BIN=
if ((PACKAGE_ONLY == 0)); then
    PYTHON_CANDIDATES=()
    if [[ -n "${ODA_TEST_PYTHON:-}" ]]; then
        PYTHON_CANDIDATES+=("$ODA_TEST_PYTHON")
    fi
    PYTHON_CANDIDATES+=("$ROOT_DIR/.venv/bin/python")
    PYTHON_CANDIDATES+=("$ROOT_DIR/bin/python")
    if [[ -n "${HOME:-}" ]]; then
        PYTHON_CANDIDATES+=("$HOME/.venvs/open-data-audit-test/bin/python")
    fi
    PYTHON_CANDIDATES+=(python3 python)
    for candidate in "${PYTHON_CANDIDATES[@]}"; do
        if [[ "$candidate" == */* ]]; then
            [[ -x "$candidate" ]] || continue
        else
            command -v "$candidate" >/dev/null 2>&1 || continue
        fi
        if "$candidate" -c 'import sys, pytest; raise SystemExit(0 if sys.version_info >= (3, 8) else 1)' >/dev/null 2>&1; then
            PYTHON_BIN=$candidate
            break
        fi
    done
fi

PYTEST_TMP_DIR=
if [[ -n "$PYTHON_BIN" ]]; then
    PYTEST_TMP_DIR=$(mktemp -d "${TMPDIR:-/tmp}/oda-pytest-${ARCH_LABEL}.XXXXXX" 2>/dev/null || true)
    if [[ -z "$PYTEST_TMP_DIR" ]]; then
        echo "[ERROR] Could not create an isolated pytest temporary directory under ${TMPDIR:-/tmp}." >&2
        PYTHON_BIN=
    fi
fi
SMOKE_HOME_DIR=
if ((${#ARTIFACTS[@]})); then
    SMOKE_HOME_DIR=$(mktemp -d "${TMPDIR:-/tmp}/oda-smoke-${ARCH_LABEL}.XXXXXX" 2>/dev/null || true)
fi
cleanup_pytest_tmp() {
    if [[ -n "$PYTEST_TMP_DIR" && -d "$PYTEST_TMP_DIR" ]]; then
        rm -rf -- "$PYTEST_TMP_DIR"
    fi
    if [[ -n "$SMOKE_HOME_DIR" && -d "$SMOKE_HOME_DIR" ]]; then
        rm -rf -- "$SMOKE_HOME_DIR"
    fi
}
trap cleanup_pytest_tmp EXIT

OS_NAME=Unknown
if [[ -r /etc/os-release ]]; then
    # shellcheck disable=SC1091
    . /etc/os-release
    OS_NAME=${PRETTY_NAME:-${NAME:-Unknown}}
fi

LIBC_INFO=Unknown
if command -v getconf >/dev/null 2>&1; then
    LIBC_INFO=$(getconf GNU_LIBC_VERSION 2>/dev/null || true)
fi
if [[ -z "$LIBC_INFO" ]] && command -v ldd >/dev/null 2>&1; then
    LIBC_INFO=$(ldd --version 2>/dev/null | head -n 1 || true)
fi
[[ -n "$LIBC_INFO" ]] || LIBC_INFO=Unknown

set +e
{
    echo "Open Data Audit - Linux Test Run"
    echo "Started: $(date '+%Y-%m-%d %H:%M:%S %z')"
    echo "Working directory: $ROOT_DIR"
    echo "OS: $OS_NAME"
    echo "Kernel: $(uname -sr)"
    echo "Architecture: $ARCH_RAW (report label: $ARCH_LABEL)"
    if [[ -n "${ODA_TEST_EXECUTION_MODE:-}" ]]; then
        echo "Execution mode: $ODA_TEST_EXECUTION_MODE"
    fi
    echo "glibc: $LIBC_INFO"
    if command -v soffice >/dev/null 2>&1; then
        echo "LibreOffice: $(soffice --version 2>&1 | head -n 1)"
    else
        echo "LibreOffice: not found on PATH"
    fi
    if command -v node >/dev/null 2>&1; then
        echo "Node.js: $(node --version 2>&1 | head -n 1)"
    else
        echo "Node.js: not found on PATH"
    fi
    if ((PACKAGE_ONLY)); then
        echo "Test mode: packaged-app startup smoke only"
        echo "Python interpreter: not required in package-only mode"
        echo "Command: sh shell-flask/Linux-3-测试发行包启动.sh [provided artifact paths]"
    elif [[ -n "$PYTHON_BIN" ]]; then
        echo "Test mode: full core pytest"
        echo "Python requirement for source tests: >= 3.8"
        PYTHON_INFO=$("$PYTHON_BIN" -c 'import sys; print("%s (%s)" % (sys.version.split()[0], sys.executable))')
        echo "Python interpreter: $PYTHON_INFO"
        echo "Pytest temporary directory: $PYTEST_TMP_DIR"
        echo "Command: $PYTHON_BIN -m pytest -v -rs --basetemp=$PYTEST_TMP_DIR core/tests"
    else
        echo "Test mode: full core pytest"
        echo "Python requirement for source tests: >= 3.8"
        echo "Python interpreter: no Python >= 3.8 with pytest found"
        echo "Command: not run"
    fi
    echo "Report: $REPORT_FILE"
    echo
    echo "Functional verification map and representative cases (from core/tests/TEST_GROUPS.md):"
    echo "This describes the covered behaviors; the actual case names and statuses follow below."
    echo
    if [[ ! -f "$ROOT_DIR/core/tests/TEST_GROUPS.md" ]]; then
        if ((PACKAGE_ONLY)); then
            echo "[INFO] Source test coverage map is not included in this standalone package-only kit."
        else
            echo "[ERROR] Could not locate core/tests/TEST_GROUPS.md."
        fi
    elif ((PACKAGE_ONLY)); then
        echo "[INFO] Source test coverage map is omitted in package-only mode."
    elif grep -q '^## S1' core/tests/TEST_GROUPS.md; then
        sed -n '/^## S1/,/^## 后续执行口径/{/^## 后续执行口径/d;p}' core/tests/TEST_GROUPS.md
    else
        echo "[ERROR] Could not locate S1/S2/S3 coverage details in core/tests/TEST_GROUPS.md."
    fi
    echo

    if ((PACKAGE_ONLY)); then
        echo "[NOT RUN] Core pytest intentionally omitted in package-only mode."
        CORE_RC=0
        CORE_STATUS="not run (package-only mode)"
    elif [[ -n "$PYTHON_BIN" ]]; then
        echo "Actual pytest cases and statuses:"
        echo
        PYTHONIOENCODING=utf-8 "$PYTHON_BIN" -m pytest -v -rs --basetemp="$PYTEST_TMP_DIR" core/tests
        CORE_RC=$?
        CORE_STATUS="exit code $CORE_RC"
    else
        echo "[NOT RUN] Core pytest requires Python 3.8+ and pytest."
        echo "Install the project test dependencies and pytest in a suitable environment."
        echo "Kylin V10 systems with only Python 3.5 can run package smoke with --package-only --artifact PATH."
        CORE_RC=2
        CORE_STATUS="not run (Python 3.8+ and pytest unavailable)"
    fi
    echo

    if ((${#ARTIFACTS[@]})); then
        echo "Packaged application startup smoke test:"
        echo "Command: sh shell-flask/Linux-3-测试发行包启动.sh [provided artifact paths]"
        if [[ -z "$SMOKE_HOME_DIR" ]]; then
            echo "[ERROR] Could not create an isolated temporary home for package smoke."
            SMOKE_RC=2
        else
            echo "Temporary HOME and XDG_DATA_HOME: $SMOKE_HOME_DIR"
            HOME="$SMOKE_HOME_DIR" XDG_DATA_HOME="$SMOKE_HOME_DIR/.local/share" \
                sh "$ROOT_DIR/shell-flask/Linux-3-测试发行包启动.sh" "${ARTIFACTS[@]}"
            SMOKE_RC=$?
        fi
    else
        echo "[NOT RUN] Packaged application startup smoke test."
        echo "To include it, rerun with --artifact PATH for each .deb or .tar.gz package."
        SMOKE_RC=0
    fi

    if ((CORE_RC != 0)); then
        RUN_RC=$CORE_RC
    else
        RUN_RC=$SMOKE_RC
    fi
    echo
    if ((PACKAGE_ONLY)) && ((RUN_RC == 0)); then
        echo "[PASS] Linux packaged-app smoke completed; core pytest was not run."
    elif ((RUN_RC == 0)); then
        echo "[PASS] Linux test run completed successfully."
    else
        echo "[FAIL] Linux test run exited with code $RUN_RC"
    fi
    echo "Core pytest status: $CORE_STATUS"
    echo "Package smoke exit code: $SMOKE_RC"
    echo "Exit code: $RUN_RC"
    exit "$RUN_RC"
} 2>&1 | tee "$REPORT_FILE"
RUN_PIPE_STATUS=(${PIPESTATUS[@]})
exit "${RUN_PIPE_STATUS[0]}"
