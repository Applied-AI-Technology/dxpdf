#!/usr/bin/env bash
# Build the Node binding (`--features node`) and place it at node/dxpdf.node.
# A Node addon is the crate's cdylib under another name: .so on Linux, .dylib
# on macOS, .dll on Windows.
set -euo pipefail
root="$(cd "$(dirname "$0")/.." && pwd)"
profile="${1:-release}"
cargo build --manifest-path "$root/Cargo.toml" --features node \
  $([ "$profile" = release ] && echo --release)
case "$(uname -s)" in
  Darwin) lib=libdxpdf.dylib ;;
  Linux) lib=libdxpdf.so ;;
  MINGW*|MSYS*|CYGWIN*) lib=dxpdf.dll ;;
  *) echo "unsupported host: $(uname -s)" >&2; exit 1 ;;
esac
target_dir="${CARGO_TARGET_DIR:-$root/target}"
cp "$target_dir/$profile/$lib" "$root/node/dxpdf.node"
echo "built $root/node/dxpdf.node"
