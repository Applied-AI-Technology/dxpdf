#!/usr/bin/env python3
"""Build ``test-files/inline-effect-extent.docx``, which asks what an inline
picture's ``wp:effectExtent`` adds to the line it stands on and to the width
it takes.

chazy's ``ReceiptTemplate`` centres a 65.25pt logo whose ``effectExtent``
gives 9525 EMU (0.75pt) below and to the right; everything under the logo
stands 0.75–0.95pt lower in Word than a line sized from ``wp:extent`` alone
puts it.

Each case is a paragraph ``E<n>a`` + a 36×24pt picture + ``E<n>b``, then a
paragraph ``E<n>z``, in Calibri 11pt with no paragraph spacing, so the
baselines say how tall the picture's line is and ``E<n>b``'s x how wide the
picture is. The extents, in points, left/top/right/bottom:

  E0  0/0/0/0 (the control)    E4  0/0/6/0
  E1  0/0/0/6                  E5  3/3/3/3
  E2  0/6/0/0                  E6  0/0/0.75/0.75 (the receipt's)
  E3  6/0/0/0

``M1``/``M2`` repeat E1/E2 at 1.15 (``w:line`` 276, auto) and ``X1`` E1
at an exact 40pt line. ``C1`` is the receipt's shape: a centred paragraph
holding only the picture with E1's extent, then ``C1z``.

Regenerate and commit the result if the content changes; the build is
deterministic and needs no third-party packages.

    scripts/make_inline_effect_extent_fixture.py
"""

import pathlib
import struct
import zipfile
import zlib

ROOT = pathlib.Path(__file__).resolve().parent.parent
OUT = ROOT / "test-files" / "inline-effect-extent.docx"

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
    '<w:pPrDefault><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:pPrDefault>'
    "</w:docDefaults>"
    '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>'
    "</w:styles>"
)

SECT = (
    '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" '
    'w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>'
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
        + chunk(b"IDAT",zlib.compress(row * height, 9))
        + chunk(b"IEND", b"")
    )


def emu(pt: float) -> int:
    return round(pt * EMU_PER_PT)


def inline(pic_id: int, extent: tuple[float, float, float, float]) -> str:
    left, top, right, bottom = (emu(v) for v in extent)
    cx, cy = emu(36), emu(24)
    return (
        "<w:r><w:drawing>"
        '<wp:inline distT="0" distB="0" distL="0" distR="0">'
        f'<wp:extent cx="{cx}" cy="{cy}"/>'
        f'<wp:effectExtent l="{left}" t="{top}" r="{right}" b="{bottom}"/>'
        f'<wp:docPr id="{pic_id}" name="Picture {pic_id}"/>'
        "<wp:cNvGraphicFramePr/>"
        '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">'
        "<pic:pic>"
        f'<pic:nvPicPr><pic:cNvPr id="{pic_id}" name="block.png"/><pic:cNvPicPr/></pic:nvPicPr>'
        '<pic:blipFill><a:blip r:embed="rIdImg"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>'
        f'<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm>'
        '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>'
        "</pic:pic></a:graphicData></a:graphic>"
        "</wp:inline></w:drawing></w:r>"
    )


def run(text: str) -> str:
    return f'<w:r><w:t xml:space="preserve">{text}</w:t></w:r>'


def para(content: str, ppr: str = "") -> str:
    ppr_xml = f"<w:pPr>{ppr}</w:pPr>" if ppr else ""
    return f"<w:p>{ppr_xml}{content}</w:p>"


EXTENTS = {
    "E0": (0, 0, 0, 0),
    "E1": (0, 0, 0, 6),
    "E2": (0, 6, 0, 0),
    "E3": (6, 0, 0, 0),
    "E4": (0, 0, 6, 0),
    "E5": (3, 3, 3, 3),
    "E6": (0, 0, 0.75, 0.75),
}

AUTO_115 = '<w:spacing w:line="276" w:lineRule="auto"/>'
EXACT_40 = '<w:spacing w:line="800" w:lineRule="exact"/>'


def body() -> str:
    out = [para(run("REF"))]
    pic = 0

    def case(label: str, extent, ppr: str = "") -> None:
        nonlocal pic
        pic += 1
        out.append(para(run(f"{label}a ") + inline(pic, extent) + run(f" {label}b"), ppr))
        out.append(para(run(f"{label}z")))

    for label, extent in EXTENTS.items():
        case(label, extent)
    case("M1", EXTENTS["E1"], AUTO_115)
    case("M2", EXTENTS["E2"], AUTO_115)
    case("X1", EXTENTS["E1"], EXACT_40)
    pic += 1
    out.append(para(inline(pic, EXTENTS["E1"]), '<w:jc w:val="center"/>'))
    out.append(para(run("C1z"), '<w:jc w:val="center"/>'))
    return "".join(out)


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
