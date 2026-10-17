# A production-like host for dxpdf: Debian trixie, native arm64 on Apple
# silicon and Graviton alike, with only the free metric-compatible fonts a
# server image gets — Carlito for Calibri, Caladea for Cambria, Liberation for
# Arial, Times New Roman and Courier New — and none of Office's. A document
# rendered here takes the substitution path (`render::fonts::substitutes`) and
# the font backend (FreeType) a deployment does, which a macOS host with Office
# installed takes for neither. Build dependencies follow
# `.github/workflows/deb.yml`.
#
# Two look-alikes for Office's script faces come from Google Fonts, which
# Debian does not package: Imperial Script for Edwardian Script ITC and Indie
# Flower for Bradley Hand ITC (`LOOKALIKE_SUBSTITUTIONS`). Both are SIL Open
# Font License; each is pinned to one google/fonts commit and checked against
# its digest, and ships with its licence as the OFL requires.
#
# Use `scripts/linux-host/run.sh`, which builds this image and mounts the
# checkout read-only.
FROM rust:1.95-trixie
RUN apt-get update && apt-get install -y --no-install-recommends \
      build-essential clang-19 libclang-19-dev ninja-build python3 git pkg-config \
      libfontconfig1-dev libfreetype-dev fontconfig \
      fonts-liberation fonts-crosextra-carlito fonts-crosextra-caladea fonts-dejavu-core \
    && rm -rf /var/lib/apt/lists/*
ARG GOOGLE_FONTS=https://raw.githubusercontent.com/google/fonts/0b58fb370093f9a9f4ff785d94405710b79de67c/ofl
RUN set -eu; \
    fetch() { curl -fsSL -o "$2" "$GOOGLE_FONTS/$1"; echo "$3  $2" | sha256sum -c -; }; \
    mkdir -p /usr/share/fonts/truetype/imperial-script /usr/share/fonts/truetype/indie-flower; \
    cd /usr/share/fonts/truetype/imperial-script; \
    fetch imperialscript/ImperialScript-Regular.ttf ImperialScript-Regular.ttf \
      fa040ae596cde82559c3bcf056ac216a94efa904b848382ab85740d38df4bab7; \
    fetch imperialscript/OFL.txt OFL.txt \
      42b758018a3f0fa429f9aaf7cafaa273ce922ea7dfba30f50cc828b8aa3e1717; \
    cd /usr/share/fonts/truetype/indie-flower; \
    fetch indieflower/IndieFlower-Regular.ttf IndieFlower-Regular.ttf \
      ccc94b22b156e9c5dfe50fd051f01b097600b252c24473e624bb43a143140a94; \
    fetch indieflower/OFL.txt OFL.txt \
      d9cdb075679d15a99bab2a8541b9c4fdc71347502111fbe9affd1d462219d96b; \
    fc-cache -f
RUN rustup component add rustfmt clippy
ENV CC=clang-19 CXX=clang++-19
