#!/usr/bin/env python3
"""Build ``test-files/table-split-head.docx``, which asks how tall a table row
is when it heads the page a table has broken onto. Each case is a table of
three single-cell rows labelled ``<case>r1``..``<case>r3``, pushed down its own
page, under a ``PAGE<case>`` line, by the ``SEP<case>`` paragraph's
``w:spacing/@w:before`` so that two rows fit and the third goes over, then an
``END<case>`` paragraph: the distance from
``<case>r3`` to ``END<case>`` is the continued row's height.

  H1 — ``exact`` 36pt, 4.95pt bottom ``w:tcMar``, 0.5pt rules
  H2 — ``exact`` 36pt, a 0.5pt top border and a 6pt ``insideH``
  H3 — ``exact`` 36pt, a 6pt top border and a 0.5pt ``insideH``
  H4 — no ``w:trHeight``, 0.5pt rules
  H5 — no ``w:trHeight``, a 6pt top border and a 0.5pt ``insideH``
  H6 — ``atLeast`` 36pt, 0.5pt rules
  H7 — ``atLeast`` 36pt, a 6pt top border and a 0.5pt ``insideH``
  H8 — ``exact`` 36pt, no top border and a 0.5pt ``insideH``

Every other border is 0.5pt. It carries its own ``styles.xml`` (Times New
Roman 12pt, no paragraph spacing, ``TableNormal`` cell margins), as
``row-height-margins.docx`` does.

Regenerate and commit the result if the content changes; the build is
deterministic and needs no third-party packages.

    scripts/make_table_split_head_fixture.py
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
BOTTOM = '<w:bottom w:w="99" w:type="dxa"/>'
# The page's 648pt of text height, less the PAGE and SEP lines, less two and a half rows:
# two rows fit under the spacing and the third does not, by half a row either way.
TEXT_HEIGHT = 648.0
LINE = 13.8


def para(text: str = "", before_twips: int = 0) -> str:
    ppr = f'<w:pPr><w:spacing w:before="{before_twips}"/></w:pPr>' if before_twips else ""
    run = f'<w:r><w:t xml:space="preserve">{text}</w:t></w:r>' if text else ""
    return f"<w:p>{ppr}{run}</w:p>"


def border(side: str, eighths: int) -> str:
    if eighths == 0:
        return f'<w:{side} w:val="nil"/>'
    return f'<w:{side} w:val="single" w:sz="{eighths}" w:space="0" w:color="000000"/>'


def table(tag: str, rule: str, twips: int, pitch: float, tc_mar: str = "", top: int = 4, inside_h: int = 4) -> str:
    borders = border("top", top) + "".join(border(s, 4) for s in ("left", "bottom", "right", "insideV"))
    borders += border("insideH", inside_h)
    rows = []
    for i in (1, 2, 3):
        tr_pr = f'<w:trPr><w:trHeight w:hRule="{rule}" w:val="{twips}"/></w:trPr>' if rule else ""
        tc_pr = f"<w:tcMar>{tc_mar}</w:tcMar>" if tc_mar else ""
        rows.append(
            f'<w:tr>{tr_pr}<w:tc><w:tcPr><w:tcW w:w="4320" w:type="dxa"/>{tc_pr}</w:tcPr>'
            f"{para(f'{tag}r{i}')}</w:tc></w:tr>"
        )
    before = round((TEXT_HEIGHT - 2 * LINE - 2.5 * pitch) * 20)
    return (
        para(f"PAGE{tag}")
        + para(f"SEP{tag}", before)
        + f'<w:tbl><w:tblPr><w:tblW w:w="4320" w:type="dxa"/><w:tblBorders>{borders}</w:tblBorders>'
        + '<w:tblLayout w:type="fixed"/></w:tblPr><w:tblGrid><w:gridCol w:w="4320"/></w:tblGrid>'
        + "".join(rows)
        + "</w:tbl>"
        + para(f"END{tag}")
    )


PAGE_BREAK = '<w:p><w:r><w:br w:type="page"/></w:r></w:p>'

CASES = [
    ("H1", "exact", 720, 41.0, BOTTOM),
    ("H2", "exact", 720, 36.0, "", 4, 48),
    ("H3", "exact", 720, 36.0, "", 48, 4),
    ("H4", "", 0, 14.3),
    ("H5", "", 0, 14.3, "", 48, 4),
    ("H6", "atLeast", 720, 36.5),
    ("H7", "atLeast", 720, 36.5, "", 48, 4),
    ("H8", "exact", 720, 36.0, "", 0, 4),
]


def body() -> str:
    parts = []
    for n, case in enumerate(CASES):
        if n:
            parts.append(PAGE_BREAK)
        parts.append(table(*case))
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
    write("table-split-head.docx", body())


if __name__ == "__main__":
    main()
