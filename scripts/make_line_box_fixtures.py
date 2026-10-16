#!/usr/bin/env python3
"""Build ``test-files/line-box.docx`` and ``test-files/line-box-inline.docx``,
which ask where Word puts a line's baseline inside its box and how tall the box
is. Every case sits between Calibri 11pt reference lines ``R<nn>``, whose own
metrics are known (no line gap), so the distance from one reference to the
next case line says what the case put above its baseline, and from the case to
the next reference what it put below.

``line-box.docx``: two-line paragraphs ``<case>a`` / ``<case>b`` (the second
after a ``<w:br/>``), labelled ``<font><size><spacing>``:

  fonts    CA Calibri, AR Arial, TR Times New Roman — at 11 and 22pt
  spacing  s1 single, s108 1.08, s115 1.15, s2 double, x30 exact 30pt,
           x8 exact 8pt, l30 at least 30pt

then mixed lines — ``MX1`` Calibri 22 with Arial 11 (``MY1``), ``MX2`` Times
11 with Calibri 22, ``MX3`` Arial 22 with Calibri 11, single and 1.15 — and a
table cell holding Arial 22 (``TC``).

``line-box-inline.docx``: a 40pt inline picture alone (I1), between text (I2),
a 6pt one between text (I3), alone at 1.15 (I4), exact 20pt (I5) and at least
60pt (I6), beside Arial (I7), alone over an Arial 22 paragraph mark (I8),
between text at 1.15 (I9); then Arial lines broken by a ``<w:br/>`` in a
Calibri 22 run (B1), an unformatted run (B2) and an Arial one (B3), under a
Calibri 22 paragraph mark (B4), after a 22pt first line (B5), empty paragraphs
with Arial 22 (B6) and Calibri 22 (B7) marks, and a break that ends its
paragraph (B8).

Both carry their own ``styles.xml`` (Calibri 11pt, single spacing, no
paragraph spacing) and ask for Word 2013 layout.

Regenerate and commit the result if the content changes; the build is
deterministic and needs no third-party packages.

    scripts/make_line_box_fixtures.py
"""

import pathlib
import struct
import zipfile
import zlib

ROOT = pathlib.Path(__file__).resolve().parent.parent
OUT = ROOT / "test-files"

NS = (
    'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
    'xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" '
    'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="w14"'
)
DRAWING_NS = (
    ' xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"'
    ' xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"'
    ' xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"'
    ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
)

CONTENT_TYPES = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
{png}<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>
</Types>"""
PNG_TYPE = '<Default Extension="png" ContentType="image/png"/>\n'

ROOT_RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>"""

DOC_RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rIdS" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
<Relationship Id="rIdT" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>
{image}</Relationships>"""
IMAGE_REL = '<Relationship Id="rIdImg" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/i.png"/>\n'

SETTINGS = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings {NS}><w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat></w:settings>"""

STYLES = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles {NS}>
<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:cs="Calibri"/><w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr></w:rPrDefault>
<w:pPrDefault><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>
<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
<w:style w:type="table" w:default="1" w:styleId="TableNormal"><w:name w:val="Normal Table"/><w:tblPr><w:tblInd w:w="0" w:type="dxa"/>
<w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:left w:w="108" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar></w:tblPr></w:style>
</w:styles>"""

SECT = (
    '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" '
    'w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>'
)
PAGE_BREAK = '<w:p><w:r><w:br w:type="page"/></w:r></w:p>'
BR = "<w:r><w:br/></w:r>"


def rpr(font=None, half=None) -> str:
    out = f'<w:rFonts w:ascii="{font}" w:hAnsi="{font}" w:cs="{font}"/>' if font else ""
    if half:
        out += f'<w:sz w:val="{half}"/><w:szCs w:val="{half}"/>'
    return out


def run(text: str, r: str = "") -> str:
    return f'<w:r>{f"<w:rPr>{r}</w:rPr>" if r else ""}<w:t xml:space="preserve">{text}</w:t></w:r>'


class Refs:
    """The ``R<nn>`` Calibri reference lines, numbered in document order —
    formatted explicitly, or left to the 11pt Calibri default."""

    def __init__(self, explicit: bool) -> None:
        self.n = 0
        self.rpr = rpr("Calibri", 22) if explicit else ""

    def __call__(self) -> str:
        self.n += 1
        return f"<w:p>{run(f'R{self.n:02d}', self.rpr)}</w:p>"


# --- line-box.docx ---------------------------------------------------------

FONTS = {"CA": "Calibri", "AR": "Arial", "TR": "Times New Roman"}
SPACING = {
    "s1": 'w:line="240" w:lineRule="auto"',
    "s108": 'w:line="259" w:lineRule="auto"',
    "s115": 'w:line="276" w:lineRule="auto"',
    "s2": 'w:line="480" w:lineRule="auto"',
    "x30": 'w:line="600" w:lineRule="exact"',
    "x8": 'w:line="160" w:lineRule="exact"',
    "l30": 'w:line="600" w:lineRule="atLeast"',
}


def two_lines(spacing: str, parts) -> str:
    a = "".join(run(t + "a", r) for t, r in parts)
    b = "".join(run(t + "b", r) for t, r in parts)
    return f'<w:p><w:pPr><w:spacing w:after="0" {SPACING[spacing]}/></w:pPr>{a}{BR}{b}</w:p>'


def line_box_body() -> str:
    ref = Refs(explicit=True)
    blocks, n = [], 0
    for sp in SPACING:
        for fk, font in FONTS.items():
            for size in (11, 22):
                if n and n % 6 == 0:
                    blocks.append(PAGE_BREAK)
                tag = f"{fk}{size}{sp}"
                blocks += [ref(), two_lines(sp, [(tag, rpr(font, size * 2))])]
                n += 1
    blocks.append(PAGE_BREAK)
    for sp in ("s1", "s115"):
        blocks += [ref(), two_lines(sp, [(f"MX1{sp}", rpr("Calibri", 44)), (f"MY1{sp}", rpr("Arial", 22))])]
        blocks += [ref(), two_lines(sp, [(f"MX2{sp}", rpr("Times New Roman", 22)), (f"MY2{sp}", rpr("Calibri", 44))])]
        blocks += [ref(), two_lines(sp, [(f"MX3{sp}", rpr("Arial", 44)), (f"MY3{sp}", rpr("Calibri", 22))])]
    blocks.append(ref())
    borders = "".join(
        f'<w:{s} w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        for s in ("top", "left", "bottom", "right", "insideH", "insideV")
    )
    for sp in ("s1", "s115"):
        cell = two_lines(sp, [(f"TC{sp}", rpr("Arial", 44))])
        blocks.append(
            f'<w:tbl><w:tblPr><w:tblW w:w="4320" w:type="dxa"/><w:tblBorders>{borders}</w:tblBorders></w:tblPr>'
            f'<w:tblGrid><w:gridCol w:w="4320"/></w:tblGrid><w:tr><w:tc>{cell}</w:tc></w:tr></w:tbl>'
        )
        blocks.append(ref())
    return "".join(blocks)


# --- line-box-inline.docx --------------------------------------------------

EMU = 12700


def png() -> bytes:
    """An 8×8 grey PNG — what the picture shows is beside the point."""
    raw = b"".join(b"\x00" + b"\x80\x80\x80" * 8 for _ in range(8))

    def chunk(tag: bytes, data: bytes) -> bytes:
        crc = zlib.crc32(tag + data) & 0xFFFFFFFF
        return struct.pack(">I", len(data)) + tag + data + struct.pack(">I", crc)

    header = struct.pack(">IIBBBBB", 8, 8, 8, 2, 0, 0, 0)
    return (
        b"\x89PNG\r\n\x1a\n"
        + chunk(b"IHDR", header)
        + chunk(b"IDAT", zlib.compress(raw))
        + chunk(b"IEND", b"")
    )


class Pictures:
    """Inline pictures, each with its own drawing id."""

    def __init__(self) -> None:
        self.n = 0

    def __call__(self, w_pt: float, h_pt: float) -> str:
        self.n += 1
        i = self.n
        cx, cy = int(w_pt * EMU), int(h_pt * EMU)
        return (
            f'<w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="{cx}" cy="{cy}"/>'
            f'<wp:docPr id="{i}" name="p{i}"/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">'
            f'<pic:pic><pic:nvPicPr><pic:cNvPr id="{i}" name="p{i}"/><pic:cNvPicPr/></pic:nvPicPr>'
            f'<pic:blipFill><a:blip r:embed="rIdImg"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>'
            f'<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>'
            f"</pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>"
        )


def para(content: str, spacing: str = "", mark: str = "") -> str:
    ppr = (f'<w:spacing w:after="0" {spacing}/>' if spacing else "") + (f"<w:rPr>{mark}</w:rPr>" if mark else "")
    return f'<w:p>{f"<w:pPr>{ppr}</w:pPr>" if ppr else ""}{content}</w:p>'


def brk(r: str = "") -> str:
    return f'<w:r>{f"<w:rPr>{r}</w:rPr>" if r else ""}<w:br/></w:r>'


def inline_body() -> str:
    ref, img = Refs(explicit=False), Pictures()
    ar = rpr("Arial")
    s115 = 'w:line="276" w:lineRule="auto"'
    cases = [
        para(img(40, 40)),
        para(run("I2a ") + img(40, 40) + run(" I2b")),
        para(run("I3a ") + img(6, 6) + run(" I3b")),
        para(img(40, 40), s115),
        para(img(40, 40), 'w:line="400" w:lineRule="exact"'),
        para(img(40, 40), 'w:line="1200" w:lineRule="atLeast"'),
        para(run("I7a ", ar) + img(40, 40) + run(" I7b", ar), mark=ar),
        para(img(40, 40), mark=rpr("Arial", 44)),
        para(run("I9a ") + img(40, 40) + run(" I9b"), s115),
    ]
    blocks = []
    for c in cases:
        blocks += [ref(), c]
    blocks += [ref(), PAGE_BREAK]
    breaks = [
        para(run("B1a", ar) + brk(rpr("Calibri", 44)) + run("B1b", ar), mark=ar),
        para(run("B2a", ar) + brk() + run("B2b", ar), mark=ar),
        para(run("B3a", ar) + brk(ar) + run("B3b", ar), mark=ar),
        para(run("B4a", ar) + brk(ar) + run("B4b", ar), mark=rpr("Calibri", 44)),
        para(run("B5a", ar + '<w:sz w:val="44"/>') + brk(ar) + run("B5b", ar), mark=ar),
        para("", mark=rpr("Arial", 44)),
        para("", mark=rpr("Calibri", 44)),
        para(run("B8a", ar) + brk(rpr("Calibri", 44)), mark=ar),
    ]
    for c in breaks:
        blocks += [ref(), c]
    blocks.append(ref())
    return "".join(blocks)


def document(content: str, ns: str) -> str:
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        f"<w:document {ns}><w:body>{content}{SECT}</w:body></w:document>"
    )


def write(name: str, content: str, with_image: bool) -> None:
    target = OUT / name
    parts = [
        ("[Content_Types].xml", CONTENT_TYPES.format(png=PNG_TYPE if with_image else "")),
        ("_rels/.rels", ROOT_RELS),
        ("word/_rels/document.xml.rels", DOC_RELS.format(image=IMAGE_REL if with_image else "")),
        ("word/document.xml", document(content, NS + (DRAWING_NS if with_image else ""))),
        ("word/styles.xml", STYLES),
        ("word/settings.xml", SETTINGS),
    ]
    if with_image:
        parts.append(("word/media/i.png", png()))
    # Fixed timestamps so regenerating an unchanged fixture produces identical
    # bytes and does not show up as a diff.
    with zipfile.ZipFile(target, "w", zipfile.ZIP_DEFLATED) as z:
        for part, data in parts:
            info = zipfile.ZipInfo(part, date_time=(1980, 1, 1, 0, 0, 0))
            info.compress_type = zipfile.ZIP_DEFLATED
            z.writestr(info, data)
    print(f"wrote {target.relative_to(ROOT)} ({target.stat().st_size} bytes)")


def main() -> None:
    OUT.mkdir(parents=True, exist_ok=True)
    write("line-box.docx", line_box_body(), with_image=False)
    write("line-box-inline.docx", inline_body(), with_image=True)


if __name__ == "__main__":
    main()
