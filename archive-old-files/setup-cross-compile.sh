#!/usr/bin/env bash
# Purpose: Install prerequisites and cross-compile archive-old-files for Windows x64 (MSVC).
# Author: Doug Hesseltine
# Created: 2026-03-28
# Modified: 2026-03-28
# Version: 1.0.0
#
# Usage (from this directory):
#   ./setup-cross-compile.sh          # install deps + build release
#   ./setup-cross-compile.sh --deps-only
#
# Requires: Homebrew (macOS), rustup, network for first-time SDK download.
# License: Using cargo-xwin implies acceptance of Microsoft’s MSVC/SDK terms
#   (see https://go.microsoft.com/fwlink/?LinkId=2086102).

set -euo pipefail
SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
cd "$SCRIPT_DIR"

DEPS_ONLY=false
if [[ "${1:-}" == "--deps-only" ]]; then
  DEPS_ONLY=true
fi

# Homebrew LLVM (clang / clang-cl used by cargo-xwin)
if [[ -d "/opt/homebrew/opt/llvm/bin" ]]; then
  export PATH="/opt/homebrew/opt/llvm/bin:$PATH"
elif [[ -d "/usr/local/opt/llvm/bin" ]]; then
  export PATH="/usr/local/opt/llvm/bin:$PATH"
fi

if ! command -v clang &>/dev/null; then
  echo "clang not found. Install LLVM via Homebrew:"
  echo "  brew install llvm"
  echo "Then re-run this script (or add brew llvm bin to PATH)."
  exit 1
fi

echo "Using clang: $(command -v clang)"

if ! command -v rustup &>/dev/null; then
  echo "rustup not found. Install from https://rustup.rs/"
  exit 1
fi

rustup target add x86_64-pc-windows-msvc
rustup component add llvm-tools

if ! command -v cargo-xwin &>/dev/null; then
  echo "Installing cargo-xwin (one-time)..."
  cargo install --locked cargo-xwin
fi

echo "cargo-xwin: $(command -v cargo-xwin)"

if [[ "$DEPS_ONLY" == true ]]; then
  echo "Dependencies ready. Build with:"
  echo "  ./setup-cross-compile.sh"
  exit 0
fi

# Pre-cache Windows SDK/CRT (speeds later builds; safe to repeat)
cargo xwin cache xwin || true

echo "Building release for x86_64-pc-windows-msvc..."
cargo xwin build --release --target x86_64-pc-windows-msvc

OUT="$SCRIPT_DIR/target/x86_64-pc-windows-msvc/release/archive-old-files.exe"
if [[ -f "$OUT" ]]; then
  echo ""
  echo "OK: $OUT"
  ls -la "$OUT"
else
  echo "Expected binary not found at $OUT"
  exit 1
fi
