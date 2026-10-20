#!/usr/bin/env python3
"""Build ``test-files/cell-float-valign.docx``, which asks how far a picture
anchored in a ``w:vAlign`` cell reaches into the cell's content, when it
wraps text.

A real-world certificate template signs in a bottom-aligned cell of a row at
least 73.8pt tall. The cell's one paragraph is empty but for a
``wrapTopAndBottom`` signature picture 33.25pt tall, anchored 25.95pt below
the paragraph's top. Word draws the picture with its bottom on the row's
bottom border, 14.6pt below where it stands if the cell's text were aligned
alone, and the row stays at its 73.8pt: the content is measured to the
picture's bottom, 59.2pt, and not to a line below the picture.
``anchor-valign.docx`` measured ``wrapNone`` and ``wrapSquare`` at offset 0.

Two tables, R1–R4 under ``REF`` on page 1 and R5–R8 under ``REF2`` on page
2, of two columns with 0.5pt borders and rows at least 1476 twips (73.8pt).
Column A holds the picture's paragraph, column B a bottom-aligned label
``R<n>B``. Pictures are 181.95pt wide; offsets and heights in points:

  R1  ``wrapTopAndBottom``, 25.95 below the paragraph, 33.25 tall —
      the template's cell
  R2  R1 with ``wrapSquare``
  R3  R1 with ``wrapNone``
  R4  R1 at offset 0, over the paragraph's line
  R5  R1 60 tall, so the picture reaches 12.15 past the row's minimum
  R6  R1 with text ``R6P`` in the picture's paragraph
  R7  R1 with both cells top-aligned
  R8  R1 with column A centred

Calibri 11pt, ``w:line`` 276 (auto), no paragraph spacing; ``MID`` and
``END`` follow the tables.

Regenerate and commit the result if the content changes; the build is
deterministic and needs no third-party packages.

    scripts/make_cell_float_valign_fixture.py
"""

import pathlib
import struct
import zipfile
import zlib

ROOT = pathlib.Path(__file__).resolve().parent.parent
OUT = ROOT / "test-files" / "cell-float-valign.docx"

EMU_PER_PT = 12700

CONTENT_TYPES = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Default Extension="png" ContentType="image/png"/>
<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
</Types>"""

ROOT_RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>"""

DOC_RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rIdS" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
<Relationship Id="rIdImg" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/block.png"/>
</Relationships>"""

W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
NS = (
    f"{W} "
    'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
    'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" '
    'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
    'xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"'
)

STYLES = (
    f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles {W}>'
    "<w:docDefaults>"
    '<w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:cs="Calibri"/>'
    '<w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr></w:rPrDefault>'
    '<w:pPrDefault><w:pPr><w:spacing w:after="0" w:line="276" w:lineRule="auto"/></w:pPr></w:pPrDefault>'
    "</w:docDefaults>"
    '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>'
    '<w:style w:type="table" w:default="1" w:styleId="TableNormal"><w:name w:val="Normal Table"/>'
    '<w:tblPr><w:tblInd w:w="0" w:type="dxa"/><w:tblCellMar><w:top w:w="0" w:type="dxa"/>'
    '<w:left w:w="108" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/><w:right w:w="108" w:type="dxa"/>'
    "</w:tblCellMar></w:tblPr></w:style>"
    "</w:styles>"
)

SECT = (
    '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" '
    'w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>'
)

BORDERS = (
    "<w:tblBorders>"
    + "".join(
        f'<w:{side} w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        for side in ("top", "left", "bottom", "right", "insideH", "insideV")
    )
    + "</w:tblBorders>"
)

WRAPS = {
    "topAndBottom": "<wp:wrapTopAndBottom/>",
    "square": '<wp:wrapSquare wrapText="bothSides"/>',
    "none": "<wp:wrapNone/>",
}


def png(width: int, height: int, rgb: tuple[int, int, int]) -> bytes:
    """A solid-colour RGB PNG, written by hand so the build needs no packages."""

    def chunk(kind: bytes, data: bytes) -> bytes:
        body = kind + data
        return struct.pack(">I", len(data)) + body + struct.pack(">I", zlib.crc32(body))

    row = b"\x00" + bytes(rgb) * width
    return (
        b"\x89PNG\r\n\x1a\n"
        + chunk(b"IHDR", struct.pack(">IIBBBBB", width, height, 8, 2, 0, 0, 0))
        + chunk(b"IDAT", zlib.compress(row * height, 9))
        + chunk(b"IEND", b"")
    )


def anchor(pic_id: int, wrap: str, offset_emu: int, height_emu: int) -> str:
    cx = 2310765
    return (
        "<w:r><w:drawing>"
        f'<wp:anchor distT="0" distB="0" distL="114300" distR="114300" simplePos="0" '
        f'relativeHeight="{pic_id}" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">'
        '<wp:simplePos x="0" y="0"/>'
        '<wp:positionH relativeFrom="column"><wp:posOffset>12700</wp:posOffset></wp:positionH>'
        f'<wp:positionV relativeFrom="paragraph"><wp:posOffset>{offset_emu}</wp:posOffset></wp:positionV>'
        f'<wp:extent cx="{cx}" cy="{height_emu}"/>'
        '<wp:effectExtent l="0" t="0" r="0" b="0"/>'
        f"{WRAPS[wrap]}"
        f'<wp:docPr id="{pic_id}" name="Picture {pic_id}"/>'
        "<wp:cNvGraphicFramePr/>"
        '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">'
        "<pic:pic>"
        f'<pic:nvPicPr><pic:cNvPr id="{pic_id}" name="block.png"/><pic:cNvPicPr/></pic:nvPicPr>'
        '<pic:blipFill><a:blip r:embed="rIdImg"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>'
        f'<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="{cx}" cy="{height_emu}"/></a:xfrm>'
        '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>'
        "</pic:pic></a:graphicData></a:graphic>"
        "</wp:anchor></w:drawing></w:r>"
    )


def run(text: str) -> str:
    return f'<w:r><w:t xml:space="preserve">{text}</w:t></w:r>'


def para(content: str = "") -> str:
    return f"<w:p>{content}</w:p>"


def cell(width: int, align: str, content: str) -> str:
    return (
        f'<w:tc><w:tcPr><w:tcW w:w="{width}" w:type="dxa"/><w:vAlign w:val="{align}"/></w:tcPr>'
        f"{content}</w:tc>"
    )


SIGNATURE_OFFSET = 329565
SIGNATURE_HEIGHT = 422275

# (wrap, offset, height, column A's alignment, column B's, text beside the picture)
ROWS = [
    ("topAndBottom", SIGNATURE_OFFSET, SIGNATURE_HEIGHT, "bottom", "bottom", ""),
    ("square", SIGNATURE_OFFSET, SIGNATURE_HEIGHT, "bottom", "bottom", ""),
    ("none", SIGNATURE_OFFSET, SIGNATURE_HEIGHT, "bottom", "bottom", ""),
    ("topAndBottom", 0, SIGNATURE_HEIGHT, "bottom", "bottom", ""),
    ("topAndBottom", SIGNATURE_OFFSET, 60 * EMU_PER_PT, "bottom", "bottom", ""),
    ("topAndBottom", SIGNATURE_OFFSET, SIGNATURE_HEIGHT, "bottom", "bottom", "R6P"),
    ("topAndBottom", SIGNATURE_OFFSET, SIGNATURE_HEIGHT, "top", "top", ""),
    ("topAndBottom", SIGNATURE_OFFSET, SIGNATURE_HEIGHT, "center", "bottom", ""),
]


def table(first: int, rows_spec) -> str:
    a, b = 4200, 2400
    rows = []
    for n, (wrap, offset, height, align_a, align_b, text) in enumerate(rows_spec, start=first):
        picture = anchor(n, wrap, offset, height) + (run(text) if text else "")
        rows.append(
            '<w:tr><w:trPr><w:trHeight w:val="1476"/></w:trPr>'
            + cell(a, align_a, para(picture))
            + cell(b, align_b, para(run(f"R{n}B")))
            + "</w:tr>"
        )
    return (
        "<w:tbl><w:tblPr>"
        f'<w:tblW w:w="{a + b}" w:type="dxa"/>{BORDERS}<w:tblLayout w:type="fixed"/>'
        f'</w:tblPr><w:tblGrid><w:gridCol w:w="{a}"/><w:gridCol w:w="{b}"/></w:tblGrid>'
        + "".join(rows)
        + "</w:tbl>"
    )


def body() -> str:
    page_break = '<w:p><w:r><w:br w:type="page"/></w:r></w:p>'
    return (
        para(run("REF"))
        + table(1, ROWS[:4])
        + para(run("MID"))
        + page_break
        + para(run("REF2"))
        + table(5, ROWS[4:])
        + para(run("END"))
    )


def main() -> None:
    document = (
        f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        f"<w:document {NS}><w:body>{body()}{SECT}</w:body></w:document>"
    )
    parts = {
        "[Content_Types].xml": CONTENT_TYPES,
        "_rels/.rels": ROOT_RELS,
        "word/_rels/document.xml.rels": DOC_RELS,
        "word/document.xml": document,
        "word/styles.xml": STYLES,
        "word/media/block.png": png(8, 8, (0x2E, 0x75, 0xB6)),
    }
    with zipfile.ZipFile(OUT, "w", zipfile.ZIP_DEFLATED) as z:
        for part, content in parts.items():
            info = zipfile.ZipInfo(part, date_time=(1980, 1, 1, 0, 0, 0))
            info.compress_type = zipfile.ZIP_DEFLATED
            data = content if isinstance(content, bytes) else content.encode("utf-8")
            z.writestr(info, data)
    print(f"wrote {OUT.relative_to(ROOT)}")


if __name__ == "__main__":
    main()
