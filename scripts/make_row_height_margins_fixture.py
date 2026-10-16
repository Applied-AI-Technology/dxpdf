#!/usr/bin/env python3
"""Build ``test-files/row-height-margins.docx``, which asks how §17.4.80
``w:trHeight`` combines with the §17.4.42/§17.4.43 top and bottom cell margins
and with the table's horizontal borders. Each table is three single-cell rows
labelled ``<table>r1``..``<table>r3`` under a ``SEP<table>`` paragraph, so a
row's height is the distance between two labels:

  R1 — ``exact`` 36pt, 4.95pt top and bottom ``w:tcMar``
  R3 — ``exact`` 36pt, no margins (the control)
  R4 — ``exact`` 36pt, top margin only
  R5 — ``exact`` 36pt, bottom margin only
  A1 — ``atLeast`` 36pt, no margins
  A2 — ``atLeast`` 36pt, top margin only
  A3 — ``atLeast`` 36pt, bottom margin only
  A4 — ``atLeast`` 36pt, both margins
  A5 — ``atLeast`` 18pt with two lines of content, both margins
  A9 — A5 without the margins
  A6 — no ``w:trHeight``, both margins
  A8 — A6 without the margins
  E1 — ``exact`` 36pt, a 6pt ``insideH``, no margins
  A7 — ``atLeast`` 36pt, a 6pt ``insideH``, no margins

Every other border is 0.5pt. It carries its own ``styles.xml`` (Times New
Roman 12pt, no paragraph spacing, ``TableNormal`` cell margins) for the reason
``bidi-visual-table.docx`` does: an absent ``w:docDefaults`` is
application-defined (§17.7.2), and it is measured against a Word render.

Regenerate and commit the result if the content changes; the build is
deterministic and needs no third-party packages.

    scripts/make_row_height_margins_fixture.py
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
TOP = '<w:top w:w="99" w:type="dxa"/>'
BOTTOM = '<w:bottom w:w="99" w:type="dxa"/>'


def para(text: str = "") -> str:
    run = f'<w:r><w:t xml:space="preserve">{text}</w:t></w:r>' if text else ""
    return f"<w:p>{run}</w:p>"


def border(side: str, eighths: int) -> str:
    return f'<w:{side} w:val="single" w:sz="{eighths}" w:space="0" w:color="000000"/>'


def table(tag: str, rule: str, twips: int, tc_mar: str = "", inside_h: int = 4, lines: int = 1) -> str:
    borders = "".join(border(s, 4) for s in ("top", "left", "bottom", "right", "insideV"))
    borders += border("insideH", inside_h)
    rows = []
    for i in (1, 2, 3):
        tr_pr = f'<w:trPr><w:trHeight w:hRule="{rule}" w:val="{twips}"/></w:trPr>' if rule else ""
        tc_pr = f"<w:tcMar>{tc_mar}</w:tcMar>" if tc_mar else ""
        content = para(f"{tag}r{i}") + "".join(para(f"{tag}s{i}") for _ in range(lines - 1))
        rows.append(
            f'<w:tr>{tr_pr}<w:tc><w:tcPr><w:tcW w:w="4320" w:type="dxa"/>{tc_pr}</w:tcPr>'
            f"{content}</w:tc></w:tr>"
        )
    return (
        para(f"SEP{tag}")
        + f'<w:tbl><w:tblPr><w:tblW w:w="4320" w:type="dxa"/><w:tblBorders>{borders}</w:tblBorders>'
        + '<w:tblLayout w:type="fixed"/></w:tblPr><w:tblGrid><w:gridCol w:w="4320"/></w:tblGrid>'
        + "".join(rows)
        + "</w:tbl>"
    )


PAGE_BREAK = '<w:p><w:r><w:br w:type="page"/></w:r></w:p>'

CASES = [
    ("R1", "exact", 720, TOP + BOTTOM),
    ("R3", "exact", 720, ""),
    ("R4", "exact", 720, TOP),
    ("R5", "exact", 720, BOTTOM),
    ("A1", "atLeast", 720, ""),
    ("A2", "atLeast", 720, TOP),
    ("A3", "atLeast", 720, BOTTOM),
    ("A4", "atLeast", 720, TOP + BOTTOM),
    ("A5", "atLeast", 360, TOP + BOTTOM, 4, 2),
    ("A9", "atLeast", 360, "", 4, 2),
    ("A6", "", 0, TOP + BOTTOM),
    ("A8", "", 0, ""),
    ("E1", "exact", 720, "", 48),
    ("A7", "atLeast", 720, "", 48),
]


def body() -> str:
    parts = []
    for n, case in enumerate(CASES):
        if n and n % 4 == 0:
            parts.append(PAGE_BREAK)
        parts.append(table(*case))
    parts.append(para("SEPEND"))
    return "".join(parts)


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
    write("row-height-margins.docx", body())


if __name__ == "__main__":
    main()
