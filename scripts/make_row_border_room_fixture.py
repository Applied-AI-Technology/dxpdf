#!/usr/bin/env python3
"""Build ``test-files/row-border-room.docx``, which asks which horizontal
border between two table rows takes room in the rows, and which is only drawn.

A real-world report template has a ``List Table 3`` table whose header row
carries a direct 0.5pt bottom border and whose first body row is a
``band1Horz`` row, with the style's 0.5pt top and bottom. The body row starts
below the header's border, but the row after it starts at the top of the body
row's bottom border, as if that border took no room: an edge with a border on
both sides was charged, an edge with only the upper row's bottom border was
not.

Each case is a three-row table, rows ``<case>1``–``<case>3`` in Calibri 11pt
with no cell margins and automatic heights, under a ``SEP<case>`` paragraph,
with 6pt black borders so a charge cannot hide in rounding:

  B0  no borders (the control)
  B1  row 1 ``tcBorders/bottom``
  B2  row 2 ``tcBorders/top``
  B3  both: row 1 bottom and row 2 top
  B4  ``tblBorders/insideH``
  B5  row 1 bottom, row 2 top ``nil``
  B6  row 2 bottom (the edge to row 3), as B1 one row down
  B7  a table style whose ``band1Horz`` gives rows 1 and 3 a top and bottom
      border, ``tblLook`` banding on — the report's shape
  B8  B7 with a ``firstRow`` header whose cells carry a direct bottom border

Regenerate and commit the result if the content changes; the build is
deterministic and needs no third-party packages.

    scripts/make_row_border_room_fixture.py
"""

import pathlib
import zipfile

ROOT = pathlib.Path(__file__).resolve().parent.parent
OUT = ROOT / "test-files" / "row-border-room.docx"

NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'

CONTENT_TYPES = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>
</Types>"""

ROOT_RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>"""

DOC_RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rIdS" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
<Relationship Id="rIdT" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>
</Relationships>"""

SETTINGS = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings {NS}><w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat></w:settings>"""


def border(side: str, val: str = "single") -> str:
    if val == "nil":
        return f'<w:{side} w:val="nil"/>'
    return f'<w:{side} w:val="single" w:sz="48" w:space="0" w:color="000000"/>'


STYLES = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles {NS}>
<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:cs="Calibri"/><w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr></w:rPrDefault>
<w:pPrDefault><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>
<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
<w:style w:type="table" w:default="1" w:styleId="TableNormal"><w:name w:val="Normal Table"/><w:tblPr><w:tblInd w:w="0" w:type="dxa"/>
<w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:left w:w="108" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar></w:tblPr></w:style>
<w:style w:type="table" w:styleId="Banded"><w:name w:val="Banded"/><w:basedOn w:val="TableNormal"/>
<w:tblPr><w:tblStyleRowBandSize w:val="1"/></w:tblPr>
<w:tblStylePr w:type="band1Horz"><w:tblPr/><w:tcPr><w:tcBorders>{border("top")}{border("bottom")}</w:tcBorders></w:tcPr></w:tblStylePr>
</w:style>
</w:styles>"""

SECT = (
    '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" '
    'w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>'
)


def run(text: str) -> str:
    return f'<w:r><w:t xml:space="preserve">{text}</w:t></w:r>'


def row(text: str, borders: str = "", header: bool = False) -> str:
    tc_borders = f"<w:tcBorders>{borders}</w:tcBorders>" if borders else ""
    tr_pr = "<w:trPr><w:tblHeader/></w:trPr>" if header else ""
    return (
        f'<w:tr>{tr_pr}<w:tc><w:tcPr><w:tcW w:w="2880" w:type="dxa"/>{tc_borders}</w:tcPr>'
        f"<w:p>{run(text)}</w:p></w:tc></w:tr>"
    )


def table(rows: str, tbl_borders: str = "", style: str = "", look: str = "") -> str:
    style_el = f'<w:tblStyle w:val="{style}"/>' if style else ""
    tb = f"<w:tblBorders>{tbl_borders}</w:tblBorders>" if tbl_borders else ""
    bits = int(look or "0000", 16)
    flag = lambda mask: "1" if bits & mask else "0"  # noqa: E731
    look_el = (
        f'<w:tblLook w:val="{look or "0000"}" w:firstRow="{flag(0x20)}" w:lastRow="{flag(0x40)}" '
        f'w:firstColumn="{flag(0x80)}" w:lastColumn="{flag(0x100)}" w:noHBand="{flag(0x200)}" '
        f'w:noVBand="{flag(0x400)}"/>'
    )
    return (
        f'<w:tbl><w:tblPr>{style_el}<w:tblW w:w="2880" w:type="dxa"/><w:tblLayout w:type="fixed"/>{tb}{look_el}</w:tblPr>'
        f'<w:tblGrid><w:gridCol w:w="2880"/></w:tblGrid>{rows}</w:tbl>'
    )


def case(name: str) -> str:
    r = lambda i, b="", **k: row(f"{name}{i}", b, **k)  # noqa: E731
    if name == "B0":
        t = table(r(1) + r(2) + r(3))
    elif name == "B1":
        t = table(r(1, border("bottom")) + r(2) + r(3))
    elif name == "B2":
        t = table(r(1) + r(2, border("top")) + r(3))
    elif name == "B3":
        t = table(r(1, border("bottom")) + r(2, border("top")) + r(3))
    elif name == "B4":
        t = table(r(1) + r(2) + r(3), tbl_borders=border("insideH"))
    elif name == "B5":
        t = table(r(1, border("bottom")) + r(2, border("top", "nil")) + r(3))
    elif name == "B6":
        t = table(r(1) + r(2, border("bottom")) + r(3))
    elif name == "B7":
        # noVBand set and noHBand clear: horizontal banding only.
        t = table(r(1) + r(2) + r(3), style="Banded", look="0400")
    else:  # B8: a header row, then the banded body.
        t = table(
            r(0, border("bottom"), header=True) + r(1) + r(2) + r(3),
            style="Banded",
            look="0420",
        )
    return f"<w:p>{run('SEP' + name)}</w:p>{t}"


def main() -> None:
    names = ["B0", "B1", "B2", "B3", "B4", "B5", "B6", "B7", "B8"]
    body = "".join(case(n) for n in names) + f"<w:p>{run('END')}</w:p>"
    document = (
        f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        f"<w:document {NS}><w:body>{body}{SECT}</w:body></w:document>"
    )
    parts = {
        "[Content_Types].xml": CONTENT_TYPES,
        "_rels/.rels": ROOT_RELS,
        "word/_rels/document.xml.rels": DOC_RELS,
        "word/document.xml": document,
        "word/styles.xml": STYLES,
        "word/settings.xml": SETTINGS,
    }
    with zipfile.ZipFile(OUT, "w", zipfile.ZIP_DEFLATED) as z:
        for name, text in parts.items():
            info = zipfile.ZipInfo(name, date_time=(1980, 1, 1, 0, 0, 0))
            info.compress_type = zipfile.ZIP_DEFLATED
            z.writestr(info, text.encode("utf-8"))
    print(f"wrote {OUT.relative_to(ROOT)}")


if __name__ == "__main__":
    main()
