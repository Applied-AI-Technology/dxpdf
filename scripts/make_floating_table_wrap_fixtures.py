#!/usr/bin/env python3
"""Build ``test-files/floating-table-wrap-15.docx`` and
``test-files/floating-table-wrap-14.docx``, which ask where a §17.4.57
floating table with ``vertAnchor="text"`` stands and how close the text that
wraps around it comes. One case per page: a ``<case>pre`` paragraph (one line,
or three), the floating table — 216pt, three rows labelled ``<case>r1``..r3,
0.5pt borders unless stated — then a ``<case>post`` paragraph of words
``w00``..``w79`` long enough to wrap beside it and below:

  A1 — ``tblpY`` 1 twip after a one-line paragraph (what Word writes for a
       dragged table, and what the corpus uses)
  A2 — the same after a three-line paragraph
  A3 — ``tblpY`` 20pt after a one-line paragraph
  A4 — ``tblpY`` 20pt after a three-line paragraph
  A5 — ``tblpY`` −10pt after a three-line paragraph
  G1 — ``rightFromText`` 4 twips (0.2pt)
  G2 — ``rightFromText`` 20 twips (1pt)
  G3 — ``rightFromText`` 180 with a 6pt right border
  R1 — ``tblpXSpec="right"``, ``leftFromText`` 180, the wrapping paragraph
       justified so its lines end on the gap
  R2 — the same with ``leftFromText`` 0

``leftFromText``/``rightFromText`` are 180 (9pt) unless stated. The two files
differ only in §17.15.1.21 ``compatibilityMode``; Times New Roman 12pt.

Regenerate and commit the result if the content changes; the build is
deterministic and needs no third-party packages.

    scripts/make_floating_table_wrap_fixtures.py
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
WORDS = " ".join(f"w{i:02d}" for i in range(80))
# Three lines at 12pt Times across the 468pt column.
PRE3 = " ".join(f"p{i:02d}" for i in range(75))


def para(text: str, jc: str = "") -> str:
    ppr = f'<w:pPr><w:jc w:val="{jc}"/></w:pPr>' if jc else ""
    return f'<w:p>{ppr}<w:r><w:t xml:space="preserve">{text}</w:t></w:r></w:p>'


def border(side: str, eighths: int) -> str:
    return f'<w:{side} w:val="single" w:sz="{eighths}" w:space="0" w:color="000000"/>'


def table(tag: str, tblp: str, right: int = 4) -> str:
    borders = (
        border("top", 4)
        + border("left", 4)
        + border("bottom", 4)
        + border("right", right)
        + border("insideH", 4)
        + border("insideV", 4)
    )
    rows = "".join(
        '<w:tr><w:tc><w:tcPr><w:tcW w:w="4320" w:type="dxa"/></w:tcPr>'
        f"{para(f'{tag}r{i}')}</w:tc></w:tr>"
        for i in (1, 2, 3)
    )
    return (
        f'<w:tbl><w:tblPr>{tblp}<w:tblW w:w="4320" w:type="dxa"/>'
        f'<w:tblBorders>{borders}</w:tblBorders><w:tblLayout w:type="fixed"/></w:tblPr>'
        f'<w:tblGrid><w:gridCol w:w="4320"/></w:tblGrid>{rows}</w:tbl>'
    )


def tblp(y: int = 1, left: int = 180, right: int = 180, x: str = "") -> str:
    return (
        f'<w:tblpPr w:leftFromText="{left}" w:rightFromText="{right}" w:vertAnchor="text"{x} '
        f'w:tblpY="{y}"/>'
    )


RIGHT = ' w:horzAnchor="margin" w:tblpXSpec="right"'

CASES = [
    ("A1", False, tblp(1), 4, ""),
    ("A2", True, tblp(1), 4, ""),
    ("A3", False, tblp(400), 4, ""),
    ("A4", True, tblp(400), 4, ""),
    ("A5", True, tblp(-200), 4, ""),
    ("G1", False, tblp(1, right=4), 4, ""),
    ("G2", False, tblp(1, right=20), 4, ""),
    ("G3", False, tblp(1), 48, ""),
    ("R1", False, tblp(1, x=RIGHT), 4, "both"),
    ("R2", False, tblp(1, left=0, x=RIGHT), 4, "both"),
]


def body() -> str:
    parts = []
    for n, (tag, three, pos, right, jc) in enumerate(CASES):
        if n:
            parts.append(PAGE_BREAK)
        pre = f"{tag}pre " + (PRE3 if three else "")
        parts += [para(pre), table(tag, pos, right), para(f"{tag}post {WORDS}", jc)]
    return "".join(parts)


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
    target = OUT / f"floating-table-wrap-{mode}.docx"
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
