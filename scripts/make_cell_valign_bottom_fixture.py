#!/usr/bin/env python3
"""Build ``test-files/cell-valign-bottom.docx``, which asks where Word puts a
§17.4.83 ``vAlign`` bottom (or centre) cell's content in a row that one of its
other cells gives a bottom border. Every table is two cells, 108pt each, under
a ``SEP<case>`` paragraph: the first cell holds ``<case>a`` with a ``nil``
bottom border, the second ``<case>b`` with the bottom border named, both
aligned as named, in a row at least 28.8pt tall:

  V1 — bottom, 0.5pt bottom border, the table's only row
  V2 — bottom, 3pt bottom border, the table's only row
  V3 — bottom, 0.5pt bottom border, the first of two rows (``<case>r2`` below)
  V4 — centre, 3pt bottom border, the table's only row
  V5 — bottom, no bottom border on either cell (the control)

Times New Roman 12pt, ``TableNormal`` cell margins, no other borders.

Regenerate and commit the result if the content changes; the build is
deterministic and needs no third-party packages.

    scripts/make_cell_valign_bottom_fixture.py
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


def para(text: str = "") -> str:
    run = f'<w:r><w:t xml:space="preserve">{text}</w:t></w:r>' if text else ""
    return f"<w:p>{run}</w:p>"


def cell(text: str, align: str, bottom: str) -> str:
    return (
        f'<w:tc><w:tcPr><w:tcW w:w="2160" w:type="dxa"/><w:tcBorders>{bottom}</w:tcBorders>'
        f'<w:vAlign w:val="{align}"/></w:tcPr>{para(text)}</w:tc>'
    )


def table(tag: str, align: str, eighths: int, second_row: bool = False) -> str:
    rule = (
        f'<w:bottom w:val="single" w:sz="{eighths}" w:space="0" w:color="000000"/>'
        if eighths
        else '<w:bottom w:val="nil"/>'
    )
    row = (
        '<w:tr><w:trPr><w:trHeight w:hRule="atLeast" w:val="576"/></w:trPr>'
        + cell(f"{tag}a", align, '<w:bottom w:val="nil"/>')
        + cell(f"{tag}b", align, rule)
        + "</w:tr>"
    )
    if second_row:
        row += (
            '<w:tr><w:tc><w:tcPr><w:tcW w:w="2160" w:type="dxa"/></w:tcPr>'
            f"{para(f'{tag}r2')}</w:tc>"
            '<w:tc><w:tcPr><w:tcW w:w="2160" w:type="dxa"/></w:tcPr><w:p/></w:tc></w:tr>'
        )
    return (
        para(f"SEP{tag}")
        + '<w:tbl><w:tblPr><w:tblW w:w="4320" w:type="dxa"/><w:tblLayout w:type="fixed"/></w:tblPr>'
        + '<w:tblGrid><w:gridCol w:w="2160"/><w:gridCol w:w="2160"/></w:tblGrid>'
        + row
        + "</w:tbl>"
    )


CASES = [
    ("V1", "bottom", 4),
    ("V2", "bottom", 24),
    ("V3", "bottom", 4, True),
    ("V4", "center", 24),
    ("V5", "bottom", 0),
]


def body() -> str:
    return "".join(table(*case) for case in CASES) + para("SEPEND")


def document(content: str) -> str:
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        f'<w:document xmlns:w="{W_NS}"><w:body>{content}'
        '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/>'
        '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" '
        'w:header="720" w:footer="720" w:gutter="0"/></w:sectPr></w:body></w:document>\n'
    )


def write(name: str, content: str) -> None:
    target = OUT / name
    # Fixed timestamps so regenerating an unchanged fixture produces identical
    # bytes and does not show up as a diff.
    with zipfile.ZipFile(target, "w", zipfile.ZIP_DEFLATED) as z:
        for part, data in (
            ("[Content_Types].xml", CONTENT_TYPES),
            ("_rels/.rels", ROOT_RELS),
            ("word/_rels/document.xml.rels", DOC_RELS),
            ("word/document.xml", document(content)),
            ("word/styles.xml", STYLES),
        ):
            info = zipfile.ZipInfo(part, date_time=(1980, 1, 1, 0, 0, 0))
            info.compress_type = zipfile.ZIP_DEFLATED
            z.writestr(info, data)
    print(f"wrote {target.relative_to(ROOT)} ({target.stat().st_size} bytes)")


def main() -> None:
    OUT.mkdir(parents=True, exist_ok=True)
    write("cell-valign-bottom.docx", body())


if __name__ == "__main__":
    main()
