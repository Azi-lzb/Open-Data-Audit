#!/usr/bin/env bash
# Run the ARM64 core suite and release-package checks through Docker/QEMU in WSL.

set -uo pipefail

ROOT_DIR=$(CDPATH= cd -- "$(dirname -- "${BASH_SOURCE[0]}")/.." && pwd)
cd "$ROOT_DIR" || exit 2

REPORT_DIR="$ROOT_DIR/test-output/full-test-reports"
if ! mkdir -p "$REPORT_DIR"; then
    echo "[ERROR] Could not create report directory: $REPORT_DIR" >&2
    exit 2
fi
RUN_STAMP=$(date '+%Y%m%d_%H%M%S')
REPORT_FILE="$REPORT_DIR/FullTest_linux-arm64-qemu_${RUN_STAMP}.txt"
: > "$REPORT_FILE" || exit 2
exec > >(tee -a "$REPORT_FILE") 2>&1

ARM_TEST_IMAGE="open-data-audit-arm64-test:ubuntu22.04"
ARM_TEST_RUNTIME_IMAGE="open-data-audit-arm64-test:ubuntu22.04-runtime"
GLIBC223_IMAGE="ubuntu16.04-arm64-glibc223:local"
GLIBC223_ROOTFS_URL="https://cloud-images.ubuntu.com/releases/server/16.04/release-20210329.1/ubuntu-16.04-server-cloudimg-arm64-root.tar.xz"
GLIBC223_ROOTFS_SHA256="a2e16e563f0ff675799de2f8a0d2723d629e45b8c845094c3815828866aadac5"
BUILD_CONTEXT=""
ROOTFS_TEMP=""
trap 'if [[ -n "$BUILD_CONTEXT" && -d "$BUILD_CONTEXT" ]]; then rm -rf -- "$BUILD_CONTEXT"; fi; if [[ -n "$ROOTFS_TEMP" && -d "$ROOTFS_TEMP" ]]; then rm -rf -- "$ROOTFS_TEMP"; fi' EXIT

echo "Open Data Audit - ARM64 QEMU Full Test"
echo "Started: $(date '+%Y-%m-%d %H:%M:%S %z')"
echo "Working directory: $ROOT_DIR"
echo "Execution mode: Docker with QEMU ARM64 user-mode emulation in WSL"
echo "Reports: $REPORT_DIR"
echo

if ! command -v docker >/dev/null 2>&1; then
    echo "[ERROR] Docker CLI is not available in this WSL distribution."
    echo "Install/enable Docker Desktop WSL integration or select a WSL distro with Docker Engine access."
    exit 2
fi
if ! docker info >/dev/null 2>&1; then
    echo "[ERROR] Docker Engine is not reachable from this WSL distribution."
    echo "Start Docker Desktop or the distro's Docker Engine, then rerun this test."
    exit 2
fi

echo "[1/5] Registering the ARM64 QEMU handler ..."
if ! docker run --privileged --rm tonistiigi/binfmt --install arm64; then
    echo "[ERROR] Could not register QEMU binfmt for ARM64."
    exit 2
fi

echo
echo "[2/5] Preparing the ARM64 Ubuntu 22.04 full-test image ..."
BUILD_CONTEXT=$(mktemp -d "${TMPDIR:-/tmp}/oda-arm64-image.XXXXXX") || {
    echo "[ERROR] Could not create a temporary Docker build context."
    exit 2
}
cp "$ROOT_DIR/shell-flask/requirements.txt" "$BUILD_CONTEXT/requirements.txt" || exit 2
cat > "$BUILD_CONTEXT/Dockerfile" <<'EOF'
FROM ubuntu:22.04
ENV DEBIAN_FRONTEND=noninteractive
RUN apt-get update \
    && apt-get install -y --no-install-recommends \
       ca-certificates build-essential libxml2-dev libxslt1-dev zlib1g-dev \
       nodejs python3 python3-dev python3-pip python3-venv \
    && rm -rf /var/lib/apt/lists/*
COPY requirements.txt /tmp/requirements.txt
RUN python3 -m venv /opt/oda-test-venv \
    && /opt/oda-test-venv/bin/pip install --no-cache-dir \
       -r /tmp/requirements.txt pytest xlwt
ENV ODA_TEST_PYTHON=/opt/oda-test-venv/bin/python
ENV ODA_TEST_EXECUTION_MODE="QEMU ARM64 user-mode via Docker; Ubuntu 22.04"
EOF
if ! docker build --platform linux/arm64 --tag "$ARM_TEST_IMAGE" "$BUILD_CONTEXT"; then
    echo "[ERROR] Could not build the ARM64 test image."
    exit 2
fi
rm -rf -- "$BUILD_CONTEXT"
BUILD_CONTEXT=""

BUILD_CONTEXT=$(mktemp -d "${TMPDIR:-/tmp}/oda-arm64-runtime.XXXXXX") || {
    echo "[ERROR] Could not create a temporary runtime-image build context."
    exit 2
}
cat > "$BUILD_CONTEXT/Dockerfile" <<EOF
FROM $ARM_TEST_IMAGE
RUN apt-get update \\
    && apt-get install -y --no-install-recommends curl \\
    && apt-get purge -y nodejs libnode72 \\
    && rm -rf /var/lib/apt/lists/*
EOF
if ! docker build --platform linux/arm64 --tag "$ARM_TEST_RUNTIME_IMAGE" "$BUILD_CONTEXT"; then
    echo "[ERROR] Could not build the ARM64 test runtime image."
    exit 2
fi
rm -rf -- "$BUILD_CONTEXT"
BUILD_CONTEXT=""
echo "Node-dependent frontend checks are left to the explicit Node.js test environment; Jammy Node 12 is below the tests' requirement."

echo
echo "[3/5] Running ARM64 core pytest and release-package smoke tests ..."
DIST_DIR="$ROOT_DIR/shell-flask/dist"
ARM64_DIR=""
shopt -s nullglob
DIST_CANDIDATES=("$DIST_DIR"/dual-arch-glibc223-*_S1-S2_*/arm64)
if ((${#DIST_CANDIDATES[@]})); then
    ARM64_DIR=$(printf '%s\n' "${DIST_CANDIDATES[@]}" | sort -V | tail -n 1)
fi
ARTIFACT_RELATIVE_PATHS=()
ARM64_DEB_RELATIVE=""
if [[ -n "$ARM64_DIR" ]]; then
    for artifact in "$ARM64_DIR"/*_arm64.deb "$ARM64_DIR"/*_arm64.tar.gz; do
        [[ -f "$artifact" ]] || continue
        relative_path=${artifact#"$ROOT_DIR"/}
        ARTIFACT_RELATIVE_PATHS+=("$relative_path")
        if [[ "$artifact" == *.deb ]]; then
            ARM64_DEB_RELATIVE=$relative_path
        fi
    done
fi

if ((${#ARTIFACT_RELATIVE_PATHS[@]})); then
    printf 'ARM64 artifacts selected from %s:\n' "$ARM64_DIR"
    printf '  %s\n' "${ARTIFACT_RELATIVE_PATHS[@]}"
else
    echo "[INFO] No ARM64 S1/S2 glibc223 artifacts were found; core pytest will still run."
fi

ARM64_TEST_COMMAND=(docker run --rm --platform linux/arm64
    -v "$ROOT_DIR:/workspace" -w /workspace
    "$ARM_TEST_RUNTIME_IMAGE" bash run_full_test_linux.sh)
for artifact in "${ARTIFACT_RELATIVE_PATHS[@]}"; do
    ARM64_TEST_COMMAND+=(--artifact "$artifact")
done
"${ARM64_TEST_COMMAND[@]}"
ARM64_TEST_RC=$?
if ((ARM64_TEST_RC == 0)); then
    echo "[PASS] ARM64 Ubuntu 22.04 core and package tests."
else
    echo "[FAIL] ARM64 Ubuntu 22.04 tests exited with code $ARM64_TEST_RC."
fi

GLIBC223_RC=0
if [[ -z "$ARM64_DEB_RELATIVE" ]]; then
    echo "[SKIP] glibc 2.23 install/startup test: no ARM64 S1/S2 DEB artifact was found."
else
    echo
    echo "[4/5] Ensuring the ARM64 Ubuntu 16.04 / glibc 2.23 image is available ..."
    if ! docker image inspect "$GLIBC223_IMAGE" >/dev/null 2>&1; then
        if ! command -v curl >/dev/null 2>&1 || ! command -v sha256sum >/dev/null 2>&1; then
            echo "[ERROR] curl and sha256sum are required to prepare the verified Ubuntu 16.04 ARM64 rootfs."
            GLIBC223_RC=2
        else
            ROOTFS_TEMP=$(mktemp -d "${TMPDIR:-/tmp}/oda-arm64-xenial.XXXXXX") || GLIBC223_RC=2
            if ((GLIBC223_RC == 0)); then
                ROOTFS_ARCHIVE="$ROOTFS_TEMP/ubuntu-16.04-arm64-root.tar.xz"
                echo "Downloading the official ARM64 Ubuntu 16.04 rootfs ..."
                if ! curl -fL --retry 3 "$GLIBC223_ROOTFS_URL" -o "$ROOTFS_ARCHIVE"; then
                    echo "[ERROR] Could not download the official Ubuntu 16.04 ARM64 rootfs."
                    GLIBC223_RC=2
                elif ! printf '%s  %s\n' "$GLIBC223_ROOTFS_SHA256" "$ROOTFS_ARCHIVE" | sha256sum -c -; then
                    echo "[ERROR] Ubuntu 16.04 ARM64 rootfs checksum verification failed."
                    GLIBC223_RC=2
                elif ! docker import --platform linux/arm64 "$ROOTFS_ARCHIVE" "$GLIBC223_IMAGE"; then
                    echo "[ERROR] Could not import the verified Ubuntu 16.04 ARM64 rootfs."
                    GLIBC223_RC=2
                fi
                rm -rf -- "$ROOTFS_TEMP"
                ROOTFS_TEMP=""
            fi
        fi
    fi

    if ((GLIBC223_RC == 0)); then
        image_arch=$(docker image inspect "$GLIBC223_IMAGE" --format '{{.Os}}/{{.Architecture}}' 2>/dev/null || true)
        if [[ "$image_arch" != "linux/arm64" ]]; then
            echo "[ERROR] Expected $GLIBC223_IMAGE to be linux/arm64, got: ${image_arch:-unknown}"
            GLIBC223_RC=2
        fi
    fi

    if ((GLIBC223_RC == 0)); then
        echo
        echo "[5/5] Installing the ARM64 DEB and running package smoke tests on glibc 2.23 ..."
        GLIBC223_COMMAND='set -e
export DEBIAN_FRONTEND=noninteractive
echo "Container architecture: $(uname -m)"
echo "Container glibc: $(getconf GNU_LIBC_VERSION)"
apt-get update
apt-get install -y --no-install-recommends xdg-utils iproute2
INSTALL_DEB="$1"
shift
dpkg -i "$INSTALL_DEB"
dpkg -s audit-tools-s1-s2 | grep -E "^(Status|Architecture|Version):"
ARTIFACT_ARGS=(--artifact "$INSTALL_DEB")
for artifact in "$@"; do ARTIFACT_ARGS+=(--artifact "$artifact"); done
ODA_TEST_EXECUTION_MODE="QEMU ARM64 user-mode via Docker; Ubuntu 16.04 / glibc 2.23" bash run_full_test_linux.sh --package-only "${ARTIFACT_ARGS[@]}"'
        GLIBC223_ARTIFACTS=("/workspace/$ARM64_DEB_RELATIVE")
        for artifact in "${ARTIFACT_RELATIVE_PATHS[@]}"; do
            if [[ "$artifact" != "$ARM64_DEB_RELATIVE" ]]; then
                GLIBC223_ARTIFACTS+=("/workspace/$artifact")
            fi
        done
        docker run --rm --platform linux/arm64 \
            -e "ODA_TEST_EXECUTION_MODE=QEMU ARM64 user-mode via Docker; Ubuntu 16.04 / glibc 2.23" \
            -v "$ROOT_DIR:/workspace" -w /workspace \
            "$GLIBC223_IMAGE" /bin/bash -lc "$GLIBC223_COMMAND" arm64-glibc223-test \
            "${GLIBC223_ARTIFACTS[@]}"
        GLIBC223_RC=$?
        if ((GLIBC223_RC == 0)); then
            echo "[PASS] ARM64 DEB install and package smoke on glibc 2.23."
        else
            echo "[FAIL] glibc 2.23 install/package tests exited with code $GLIBC223_RC."
        fi
    fi
fi

echo
echo "ARM64 Ubuntu 22.04 test exit code: $ARM64_TEST_RC"
echo "ARM64 glibc 2.23 package test exit code: $GLIBC223_RC"
if ((ARM64_TEST_RC == 0 && GLIBC223_RC == 0)); then
    echo "[PASS] ARM64 QEMU full test completed."
    echo "Report: $REPORT_FILE"
    exit 0
fi
echo "[FAIL] ARM64 QEMU full test failed."
echo "Report: $REPORT_FILE"
exit 1
