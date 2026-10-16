#!/usr/bin/env python3
"""Build ``test-files/cell-inset-right-15.docx`` and
``test-files/cell-inset-right-14.docx``, which ask how far a cell's text stands
in from a border on each side, against the cell margin. Every table is fixed
layout, 216pt as two 108pt cells, indented 36pt; each cell holds a
left-aligned and a right-aligned label. A label is
``<side><w:sz><margin twips><cell>`` — ``RO`` steps the table's outer right
border, ``LO`` its outer left, ``IV`` the shared ``insideV`` (with both cells'
margins on that side set) — across 3, 6 and 12pt borders and 0, 2, 5.4 and
10pt margins, and the cell letters are:

  a  first cell, left-aligned      y  first cell, right-aligned
  b  second cell, left-aligned     z  second cell, right-aligned

Every other border is 0.5pt. The two files differ only in §17.15.1.21
``compatibilityMode``; Times New Roman 12pt.

Regenerate and commit the result if the content changes; the build is
deterministic and needs no third-party packages.

    scripts/make_cell_inset_right_fixtures.py
"""

import pathlib
import zipfile

ROOT = pathlib.Path(__file__).resolve().parent.parent
OUT = ROOT / "test-files"

CONTENT_TYPES = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
  <Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>
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
  <Relationship Id="rIdT" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>
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
        <w:top w:w="0" w:type="dxa"/><w:left w:w="108" w:type="dxa"/>
        <w:bottom w:w="0" w:type="dxa"/><w:right w:w="108" w:type="dxa"/>
      </w:tblCellMar>
    </w:tblPr>
  </w:style>
</w:styles>
"""

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
PAGE_BREAK = '<w:p><w:r><w:br w:type="page"/></w:r></w:p>'


def para(text: str = "", jc: str = "") -> str:
    ppr = f'<w:pPr><w:jc w:val="{jc}"/></w:pPr>' if jc else ""
    run = f'<w:r><w:t xml:space="preserve">{text}</w:t></w:r>' if text else ""
    return f"<w:p>{ppr}{run}</w:p>"


def border(side: str, eighths: int) -> str:
    return f'<w:{side} w:val="single" w:sz="{eighths}" w:space="0" w:color="000000"/>'


def margin(side: str, twips: int) -> str:
    return f'<w:{side} w:w="{twips}" w:type="dxa"/>'


def table(kind: str, eighths: int, twips: int) -> str:
    tag = f"{kind}{eighths}m{twips}"
    widths = {"left": 4, "right": 4, "insideV": 4}
    margins = ""
    if kind == "RO":
        widths["right"] = eighths
        margins = margin("right", twips)
    elif kind == "LO":
        widths["left"] = eighths
        margins = margin("left", twips)
    else:
        widths["insideV"] = eighths
        margins = margin("left", twips) + margin("right", twips)
    borders = (
        border("top", 4)
        + border("left", widths["left"])
        + border("bottom", 4)
        + border("right", widths["right"])
        + border("insideH", 4)
        + border("insideV", widths["insideV"])
    )
    cells = ""
    for left, right in (("a", "y"), ("b", "z")):
        cells += (
            '<w:tc><w:tcPr><w:tcW w:w="2160" w:type="dxa"/></w:tcPr>'
            f"{para(tag + left)}{para(tag + right, 'right')}</w:tc>"
        )
    return (
        f'<w:tbl><w:tblPr><w:tblW w:w="4320" w:type="dxa"/><w:tblInd w:w="720" w:type="dxa"/>'
        f"<w:tblBorders>{borders}</w:tblBorders><w:tblLayout w:type=\"fixed\"/>"
        f"<w:tblCellMar>{margins}</w:tblCellMar></w:tblPr>"
        '<w:tblGrid><w:gridCol w:w="2160"/><w:gridCol w:w="2160"/></w:tblGrid>'
        f"<w:tr>{cells}</w:tr></w:tbl>{para()}"
    )


def body() -> str:
    blocks, n = [para("REF")], 0
    for eighths in (24, 48, 96):
        for twips in (0, 40, 108, 200):
            for kind in ("RO", "LO", "IV"):
                if n and n % 12 == 0:
                    blocks.append(PAGE_BREAK)
                blocks.append(table(kind, eighths, twips))
                n += 1
    return "".join(blocks)


def settings(mode: int) -> str:
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        f'<w:settings xmlns:w="{W_NS}"><w:compat><w:compatSetting w:name="compatibilityMode" '
        f'w:uri="http://schemas.microsoft.com/office/word" w:val="{mode}"/></w:compat></w:settings>\n'
    )


def document(content: str) -> str:
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        f'<w:document xmlns:w="{W_NS}"><w:body>{content}'
        '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/>'
        '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" '
        'w:header="720" w:footer="720" w:gutter="0"/></w:sectPr></w:body></w:document>\n'
    )


def write(mode: int) -> None:
    target = OUT / f"cell-inset-right-{mode}.docx"
    # Fixed timestamps so regenerating an unchanged fixture produces identical
    # bytes and does not show up as a diff.
    with zipfile.ZipFile(target, "w", zipfile.ZIP_DEFLATED) as z:
        for part, data in (
            ("[Content_Types].xml", CONTENT_TYPES),
            ("_rels/.rels", ROOT_RELS),
            ("word/_rels/document.xml.rels", DOC_RELS),
            ("word/document.xml", document(body())),
            ("word/styles.xml", STYLES),
            ("word/settings.xml", settings(mode)),
        ):
            info = zipfile.ZipInfo(part, date_time=(1980, 1, 1, 0, 0, 0))
            info.compress_type = zipfile.ZIP_DEFLATED
            z.writestr(info, data)
    print(f"wrote {target.relative_to(ROOT)} ({target.stat().st_size} bytes)")


def main() -> None:
    OUT.mkdir(parents=True, exist_ok=True)
    for mode in (15, 14):
        write(mode)


if __name__ == "__main__":
    main()
