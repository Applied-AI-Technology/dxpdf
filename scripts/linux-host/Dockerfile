# A production-like host for dxpdf: Debian trixie, native arm64 on Apple
# silicon and Graviton alike, with only the free metric-compatible fonts a
# server image gets — Carlito for Calibri, Caladea for Cambria, Liberation for
# Arial, Times New Roman and Courier New — and none of Office's. A document
# rendered here takes the substitution path (`render::fonts::substitutes`) and
# the font backend (FreeType) a deployment does, which a macOS host with Office
# installed takes for neither. Build dependencies follow
# `.github/workflows/deb.yml`.
#
# Use `scripts/linux-host/run.sh`, which builds this image and mounts the
# checkout read-only.
FROM rust:1.95-trixie
RUN apt-get update && apt-get install -y --no-install-recommends \
      build-essential clang-19 libclang-19-dev ninja-build python3 git pkg-config \
      libfontconfig1-dev libfreetype-dev fontconfig \
      fonts-liberation fonts-crosextra-carlito fonts-crosextra-caladea fonts-dejavu-core \
    && rm -rf /var/lib/apt/lists/*
RUN rustup component add rustfmt clippy
ENV CC=clang-19 CXX=clang++-19
