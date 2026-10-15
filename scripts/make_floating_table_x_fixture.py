#!/usr/bin/env python3
"""Build ``test-files/floating-table-x.docx``, which asks where §17.4.57 places a floating
table horizontally. One case per page: a 2in single-cell table whose cell reads
``X<n>``, anchored at the top of the body. The cell's text x and the table's
geometry separate "the table's edge is at the position" from "the first cell's
*text* is", which differ by the cell's left margin; case 10 widens that margin
to 0.25in so the two readings cannot coincide:

  case 1 — no ``horzAnchor``, no ``tblpX``      (every corpus template)
  case 2 — ``horzAnchor="text"``,   ``tblpX="1440"``
  case 3 — ``horzAnchor="margin"``, ``tblpX="0"``
  case 4 — ``horzAnchor="margin"``, ``tblpX="1440"``
  case 5 — ``horzAnchor="page"``,   ``tblpX="0"``
  case 6 — ``horzAnchor="page"``,   ``tblpX="2160"``
  case 7 — ``horzAnchor="margin"``, ``tblpXSpec="left"``
  case 8 — ``horzAnchor="margin"``, ``tblpXSpec="center"``
  case 9 — ``horzAnchor="margin"``, ``tblpXSpec="right"``
  case 10 — case 4 with a 0.25in left cell margin
  case 11 — no ``horzAnchor``, no ``tblpX``, ``w:jc="center"``
  case 12 — ``horzAnchor="margin"``, ``tblpX="1440"``, ``w:jc="center"``
  case 13 — ``horzAnchor="margin"``, ``tblpX="1440"``, ``w:jc="right"``
  case 14 — ``horzAnchor="margin"``, ``tblpXSpec="left"``, ``w:jc="center"``

It carries its own ``styles.xml`` (Times New Roman 12pt, no paragraph
spacing, ``TableNormal`` cell margins) for the reason ``bidi-visual-table.docx``
does: an absent ``w:docDefaults`` is application-defined (§17.7.2), and it is
measured against a Word render.

Regenerate and commit the result if the content changes; the build is
deterministic and needs no third-party packages.

    scripts/make_floating_table_x_fixture.py
"""

import pathlib
import struct
import zipfile
import zlib

ROOT = pathlib.Path(__file__).resolve().parent.parent
OUT = ROOT / "test-files"

EMU_PER_PT = 12700
TWIPS_PER_PT = 20

CONTENT_TYPES = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="png" ContentType="image/png"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
</Types>
"""

ROOT_RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>
"""

DOC_RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdS" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
  <Relationship Id="rIdImg" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/block.png"/>
</Relationships>
"""

STYLES = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:docDefaults>
    <w:rPrDefault><w:rPr>
      <w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/>
      <w:sz w:val="24"/><w:szCs w:val="24"/>
    </w:rPr></w:rPrDefault>
    <w:pPrDefault><w:pPr>
      <w:spacing w:after="0" w:line="240" w:lineRule="auto"/>
    </w:pPr></w:pPrDefault>
  </w:docDefaults>
  <w:style w:type="paragraph" w:default="1" w:styleId="Normal">
    <w:name w:val="Normal"/>
  </w:style>
  <w:style w:type="table" w:default="1" w:styleId="TableNormal">
    <w:name w:val="Normal Table"/>
    <w:tblPr>
      <w:tblInd w:w="0" w:type="dxa"/>
      <w:tblCellMar>
        <w:top w:w="0" w:type="dxa"/>
        <w:left w:w="108" w:type="dxa"/>
        <w:bottom w:w="0" w:type="dxa"/>
        <w:right w:w="108" w:type="dxa"/>
      </w:tblCellMar>
    </w:tblPr>
  </w:style>
</w:styles>
"""

NS = (
    'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
    'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
    'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" '
    'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
    'xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"'
)

SECT_PR = """<w:sectPr>
      <w:pgSz w:w="12240" w:h="15840"/>
      <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"
               w:header="720" w:footer="720" w:gutter="0"/>
    </w:sectPr>"""

BORDERS = (
    "<w:tblBorders>"
    + "".join(
        f'<w:{side} w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        for side in ("top", "left", "bottom", "right", "insideH", "insideV")
    )
    + "</w:tblBorders>"
)


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


WRAP_NONE = "<wp:wrapNone/>"
WRAP_SQUARE = '<wp:wrapSquare wrapText="bothSides"/>'
WRAP_TOP_BOTTOM = "<wp:wrapTopAndBottom/>"
WRAP_TIGHT = (
    '<wp:wrapTight wrapText="bothSides"><wp:wrapPolygon edited="0">'
    '<wp:start x="0" y="0"/><wp:lineTo x="0" y="21600"/>'
    '<wp:lineTo x="21600" y="21600"/><wp:lineTo x="21600" y="0"/>'
    '<wp:lineTo x="0" y="0"/></wp:wrapPolygon></wp:wrapTight>'
)


def anchor(
    pic_id: int,
    wrap: str,
    *,
    width_pt: float = 60,
    height_pt: float = 60,
    x_pt: float = 0,
    behind: bool = False,
    in_cell: bool = True,
    align_h: str | None = None,
) -> str:
    cx, cy = int(width_pt * EMU_PER_PT), int(height_pt * EMU_PER_PT)
    pos_h = (
        f"<wp:align>{align_h}</wp:align>"
        if align_h
        else f"<wp:posOffset>{int(x_pt * EMU_PER_PT)}</wp:posOffset>"
    )
    return (
        "<w:r><w:drawing>"
        f'<wp:anchor distT="0" distB="0" distL="114300" distR="114300" simplePos="0" '
        f'relativeHeight="{pic_id}" behindDoc="{int(behind)}" locked="0" '
        f'layoutInCell="{int(in_cell)}" allowOverlap="1">'
        '<wp:simplePos x="0" y="0"/>'
        f'<wp:positionH relativeFrom="column">{pos_h}</wp:positionH>'
        '<wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV>'
        f'<wp:extent cx="{cx}" cy="{cy}"/>'
        '<wp:effectExtent l="0" t="0" r="0" b="0"/>'
        f"{wrap}"
        f'<wp:docPr id="{pic_id}" name="Picture {pic_id}"/>'
        "<wp:cNvGraphicFramePr/>"
        '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">'
        "<pic:pic>"
        f'<pic:nvPicPr><pic:cNvPr id="{pic_id}" name="block.png"/><pic:cNvPicPr/></pic:nvPicPr>'
        '<pic:blipFill><a:blip r:embed="rIdImg"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>'
        f'<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm>'
        '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>'
        "</pic:pic></a:graphicData></a:graphic>"
        "</wp:anchor></w:drawing></w:r>"
    )


def run(text: str) -> str:
    return f'<w:r><w:t xml:space="preserve">{text}</w:t></w:r>'


def para(*content: str, ppr: str = "") -> str:
    return f"<w:p>{ppr}{''.join(content)}</w:p>"


def cell(width: int, *paras: str) -> str:
    return f'<w:tc><w:tcPr><w:tcW w:w="{width}" w:type="dxa"/></w:tcPr>{"".join(paras)}</w:tc>'


def document(body: str) -> str:
    return f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document {NS}>
  <w:body>
    {body}
    {SECT_PR}
  </w:body>
</w:document>
"""

FLOAT_W = 2880  # 2in
FLOAT_ROW_H = 800  # 40pt, three rows → 120pt


def case_label(n: int) -> str:
    return para(run(f"CASE{n}"), ppr="<w:pPr><w:pageBreakBefore/></w:pPr>")


def floating_table_x_body() -> str:
    cases = [
        ("", None),
        ('w:horzAnchor="text" w:tblpX="1440"', None),
        ('w:horzAnchor="margin" w:tblpX="0"', None),
        ('w:horzAnchor="margin" w:tblpX="1440"', None),
        ('w:horzAnchor="page" w:tblpX="0"', None),
        ('w:horzAnchor="page" w:tblpX="2160"', None),
        ('w:horzAnchor="margin" w:tblpXSpec="left"', None),
        ('w:horzAnchor="margin" w:tblpXSpec="center"', None),
        ('w:horzAnchor="margin" w:tblpXSpec="right"', None),
        ('w:horzAnchor="margin" w:tblpX="1440"', 360),
        ("", "center"),
        ('w:horzAnchor="margin" w:tblpX="1440"', "center"),
        ('w:horzAnchor="margin" w:tblpX="1440"', "right"),
        ('w:horzAnchor="margin" w:tblpXSpec="left"', "center"),
    ]
    out = []
    for n, (attrs, extra) in enumerate(cases, start=1):
        if isinstance(extra, int):
            extra = f'<w:tblCellMar><w:left w:w="{extra}" w:type="dxa"/></w:tblCellMar>'
        elif extra is not None:
            extra = f'<w:jc w:val="{extra}"/>'
        cell_mar = extra or ""
        out.append(
            case_label(n)
            + "<w:tbl><w:tblPr>"
            f'<w:tblpPr w:leftFromText="180" w:rightFromText="180" '
            f'w:vertAnchor="text" w:tblpY="1" {attrs}/>'
            f'<w:tblW w:w="{FLOAT_W}" w:type="dxa"/>{BORDERS}<w:tblLayout w:type="fixed"/>'
            f"{cell_mar}"
            f'</w:tblPr><w:tblGrid><w:gridCol w:w="{FLOAT_W}"/></w:tblGrid>'
            f"<w:tr>{cell(FLOAT_W, para(run(f'X{n}')))}</w:tr></w:tbl>"
            + para()
        )
    return "".join(out)


def write(name: str, body: str) -> None:
    target = OUT / name
    # Fixed timestamps so regenerating an unchanged fixture produces identical
    # bytes and does not show up as a diff.
    with zipfile.ZipFile(target, "w", zipfile.ZIP_DEFLATED) as z:
        for part, data in (
            ("[Content_Types].xml", CONTENT_TYPES),
            ("_rels/.rels", ROOT_RELS),
            ("word/_rels/document.xml.rels", DOC_RELS),
            ("word/document.xml", document(body)),
            ("word/styles.xml", STYLES),
            ("word/media/block.png", png(8, 8, (0x2E, 0x75, 0xB6))),
        ):
            info = zipfile.ZipInfo(part, date_time=(1980, 1, 1, 0, 0, 0))
            info.compress_type = zipfile.ZIP_DEFLATED
            z.writestr(info, data)
    print(f"wrote {target.relative_to(ROOT)} ({target.stat().st_size} bytes)")


def main() -> None:
    OUT.mkdir(parents=True, exist_ok=True)
    write("floating-table-x.docx", floating_table_x_body())


if __name__ == "__main__":
    main()
