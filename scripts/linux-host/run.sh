#!/bin/sh
# Run a command in the production-like Linux host (see Dockerfile beside this):
#
#   scripts/linux-host/run.sh cargo test --all
#   scripts/linux-host/run.sh cargo build --release
#   scripts/linux-host/run.sh /target/release/dxpdf output/in.docx -o output/out.pdf
#
# The checkout is mounted read-only at /src (the working directory) except
# `output/`, which is writable for rendered PDFs. Build output lives in the
# `dxpdf-linux-target` volume at /target and the cargo registry in
# `dxpdf-linux-cargo`, so a rebuild is incremental. The first build compiles
# Skia from source (no prebuilt binary matches this target), which takes a
# while; linking is memory-hungry, so jobs are capped for Docker's default VM.
set -eu
root=$(cd "$(dirname "$0")/../.." && pwd)
docker build -q -t dxpdf-linux-host "$root/scripts/linux-host" >/dev/null
mkdir -p "$root/output"
exec docker run --rm \
  -v "$root:/src:ro" -v "$root/output:/src/output" \
  -v dxpdf-linux-target:/target -v dxpdf-linux-cargo:/usr/local/cargo/registry \
  -e CARGO_TARGET_DIR=/target -e CARGO_BUILD_JOBS="${CARGO_BUILD_JOBS:-3}" \
  -w /src dxpdf-linux-host "$@"
