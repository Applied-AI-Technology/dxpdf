#!/usr/bin/env python3
"""Build ``test-files/table-indent-compat-{15,14,none}.docx``: the same tables
three times, differing only in §17.15.1.21 ``compatibilityMode`` (15, 14, and
no ``settings.xml`` at all), which asks where Word measures a table's §17.4.50
indent and §17.4.57 ``tblpX`` to. Every table is two cells, the first holding
``T<n>a`` and the second ``T<n>z`` right-aligned, under a ``REF`` paragraph on
the margin:

  T1  — ``pct`` 100%, no ``tblInd``
  T2  — 216pt ``dxa``, no ``tblInd``
  T3  — 216pt, ``tblInd`` 720
  T4  — T3 with 12pt left and right borders
  T5  — T1 with 12pt left and right borders
  T6  — T2 with 21.6pt left and right cell margins
  T7  — a 360pt one-cell table holding ``T7o``, then a nested 216pt table
        (``T7n``/``T7m``) and a nested 100% table (``T7p``/``T7q``)
  T10 — T2 with ``jc="center"`` (the control)
  T11 — T2 with ``tblInd`` -720
  F1  — floating, ``horzAnchor="margin"`` ``tblpX`` 1440 (its own page)
  F2  — floating, no ``tblpX`` (its own page)

Borders are 0.5pt unless stated, cell margins ``TableNormal``'s 108 twips.

Regenerate and commit the result if the content changes; the build is
deterministic and needs no third-party packages.

    scripts/make_table_indent_compat_fixtures.py
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
WORD_URI = "http://schemas.microsoft.com/office/word"

SETTINGS_CT = (
    '  <Override PartName="/word/settings.xml" '
    'ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>\n'
)
SETTINGS_REL = (
    '  <Relationship Id="rIdT" '
    'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" '
    'Target="settings.xml"/>\n'
)


def settings(mode: str) -> str:
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        f'<w:settings xmlns:w="{W_NS}"><w:compat><w:compatSetting w:name="compatibilityMode" '
        f'w:uri="{WORD_URI}" w:val="{mode}"/></w:compat></w:settings>\n'
    )


def para(text: str = "", ppr: str = "") -> str:
    run = f'<w:r><w:t xml:space="preserve">{text}</w:t></w:r>' if text else ""
    return f"<w:p>{f'<w:pPr>{ppr}</w:pPr>' if ppr else ''}{run}</w:p>"


def border(side: str, eighths: int) -> str:
    return f'<w:{side} w:val="single" w:sz="{eighths}" w:space="0" w:color="000000"/>'


RIGHT = '<w:jc w:val="right"/>'
PCT = '<w:tblW w:w="5000" w:type="pct"/>'
DXA = '<w:tblW w:w="4320" w:type="dxa"/>'


def table(cells: list, width: str, extra: str = "", outer: int = 4, grid: tuple = (2160, 2160)) -> str:
    borders = border("top", 4) + border("left", outer) + border("bottom", 4) + border("right", outer)
    borders += border("insideH", 4) + border("insideV", 4)
    cols = "".join(f'<w:gridCol w:w="{g}"/>' for g in grid)
    tcs = "".join(f"<w:tc>{c}</w:tc>" for c in cells)
    return (
        f"<w:tbl><w:tblPr>{width}<w:tblBorders>{borders}</w:tblBorders>{extra}</w:tblPr>"
        f"<w:tblGrid>{cols}</w:tblGrid><w:tr>{tcs}</w:tr></w:tbl>"
    )


def pair(n: str) -> list:
    return [para(f"T{n}a"), para(f"T{n}z", RIGHT)]


PAGE_BREAK = '<w:p><w:r><w:br w:type="page"/></w:r></w:p>'


def body() -> str:
    full = (4680, 4680)
    nested = (
        para("T7o")
        + table([para("T7n"), para("T7m", RIGHT)], DXA)
        + para()
        + table([para("T7p"), para("T7q", RIGHT)], PCT, grid=(3000, 3000))
        + para()
    )
    blocks = [
        para("REF"),
        table(pair("1"), PCT, grid=full), para(),
        table(pair("2"), DXA), para(),
        table(pair("3"), DXA, '<w:tblInd w:w="720" w:type="dxa"/>'), para(),
        table(pair("4"), DXA, '<w:tblInd w:w="720" w:type="dxa"/>', outer=96), para(),
        table(pair("5"), PCT, outer=96, grid=full), para(),
        table(pair("6"), DXA, '<w:tblCellMar><w:left w:w="432" w:type="dxa"/>'
              '<w:right w:w="432" w:type="dxa"/></w:tblCellMar>'), para(),
        table([nested], '<w:tblW w:w="7200" w:type="dxa"/>', grid=(7200,)), para(),
        table(pair("10"), DXA, '<w:jc w:val="center"/>'), para(),
        table(pair("11"), DXA, '<w:tblInd w:w="-720" w:type="dxa"/>'), para(),
        PAGE_BREAK, para("F1 before"),
        table([para("F1a"), para("F1z", RIGHT)], DXA,
              '<w:tblpPr w:horzAnchor="margin" w:tblpX="1440" w:vertAnchor="text" w:tblpY="0"/>'),
        para("F1 after"),
        PAGE_BREAK, para("F2 before"),
        table([para("F2a"), para("F2z", RIGHT)], DXA, '<w:tblpPr w:vertAnchor="text" w:tblpY="0"/>'),
        para("F2 after"),
    ]
    return "".join(blocks)


def document(content: str) -> str:
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        f'<w:document xmlns:w="{W_NS}"><w:body>{content}'
        '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/>'
        '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" '
        'w:header="720" w:footer="720" w:gutter="0"/></w:sectPr></w:body></w:document>\n'
    )


def write(name: str, mode) -> None:
    target = OUT / name
    parts = [
        ("[Content_Types].xml", CONTENT_TYPES.replace("</Types>", (SETTINGS_CT if mode else "") + "</Types>")),
        ("_rels/.rels", ROOT_RELS),
        ("word/_rels/document.xml.rels",
         DOC_RELS.replace("</Relationships>", (SETTINGS_REL if mode else "") + "</Relationships>")),
        ("word/document.xml", document(body())),
        ("word/styles.xml", STYLES),
    ]
    if mode:
        parts.append(("word/settings.xml", settings(mode)))
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
    for mode in ("15", "14", None):
        write(f"table-indent-compat-{mode or 'none'}.docx", mode)


if __name__ == "__main__":
    main()
