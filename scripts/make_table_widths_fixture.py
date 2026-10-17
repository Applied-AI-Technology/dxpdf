#!/usr/bin/env python3
"""Build ``test-files/table-widths.docx``, which asks which declared width
Word sizes a table's columns from when ``w:tcW`` and ``w:tblGrid`` disagree,
and how an autofit table shrinks columns whose content is narrower than
their ``tcW`` to make room for one whose content is wider.

albanyny's SpecialEventPermit declares ``tblLayout fixed``, a grid of
4824 + 4824 twips and every cell ``tcW`` 3024 / 6048; Word draws the columns
151.2 and 302.4pt wide — the ``tcW``. Six other templates declare autofit
tables whose ``tcW`` differs from the grid by ~120 twips, and there Word
appears to keep the grid. DonationReport (compat 14, autofit, six ``tcW``
1440) shrinks its four narrow columns by unequal amounts, in proportion to
``tcW`` minus each column's widest unbreakable content.

Each case is a bordered table in Calibri 10pt, each cell a label paragraph
``<case><row><column>`` over its content, with ``TableNormal``'s 5.4pt left
and right cell margins, under a ``SEP<case>`` paragraph. ``A`` = 4824/4824
grid; ``B`` = ``tcW`` 3024/6048; ``C`` = ``tcW`` 6048/3024.

  T1  fixed, ``tblW auto``, grid A, both rows B — the permit's shape
  T2  autofit, ``tblW auto``, grid A, both rows B
  T3  fixed, grid A, row 1 B, row 2 C — rows disagree
  T4  autofit, grid A, row 1 ``tcW`` = grid, row 2 B
  T5  fixed, ``tblW`` 9648 dxa, grid A, both rows B
  T6  fixed, six columns grid = ``tcW`` 1440, column 1 holds
      ``{#donations}{Donor}`` (~84pt, no break) — does fixed grow it?
  T7  fixed, grid 3024 x 3; row 1 a ``gridSpan`` 2 cell ``tcW`` 7000 and
      one ``tcW`` 2048, row 2 three cells ``tcW`` 3024
  T8  autofit, six ``tcW`` 1440 holding the DonationReport row
      (``{#donations}{Donor}``, ``{Email}``, ``{Amount}``, ``{Channel}``,
      ``{Method}``, ``{PaidDate}{/donations}``) — over the 468pt column
  T9  autofit, ``tblW`` 7200 dxa, six ``tcW`` 1440 holding ``{x}``
  T10 autofit, ``tblW`` 7200 dxa, six ``tcW`` 1440 holding T8's row — does a
      declared width narrower than the columns shrink them in proportion, or
      by how far each stands above its content, as T8 does?
  T11 autofit, ``tblW`` 4000 dxa, grid and ``tcW`` 2000 + 6000 holding ``L``
      and ``R`` — the same question at 1:3
  T12 autofit, ``tblW`` 8000 dxa, grid 2000 x 4, one row of two cells
      ``tcW`` 4000 — a row that stops short of its grid: do its cells take
      their ``tcW`` and the empty columns nothing, or share ``tblW`` with them?
  T13 autofit, ``tblW`` 8000 dxa, grid 2000 x 4, one cell ``gridSpan`` 3 and
      no ``tcW`` — the same without a width to go by

Regenerate and commit the result if the content changes; the build is
deterministic and needs no third-party packages.

    scripts/make_table_widths_fixture.py
"""

import pathlib
import zipfile

ROOT = pathlib.Path(__file__).resolve().parent.parent
OUT = ROOT / "test-files" / "table-widths.docx"

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

STYLES = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles {NS}>
<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:cs="Calibri"/><w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr></w:rPrDefault>
<w:pPrDefault><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>
<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
<w:style w:type="table" w:default="1" w:styleId="TableNormal"><w:name w:val="Normal Table"/><w:tblPr><w:tblInd w:w="0" w:type="dxa"/>
<w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:left w:w="108" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar></w:tblPr></w:style>
</w:styles>"""

SECT = (
    '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" '
    'w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>'
)

BORDERS = "".join(
    f'<w:{s} w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
    for s in ("top", "left", "bottom", "right", "insideH", "insideV")
)

FIXED = '<w:tblLayout w:type="fixed"/>'
AUTO = '<w:tblW w:w="0" w:type="auto"/>'
GRID_A = [4824, 4824]
B = [(3024, 1), (6048, 1)]
C = [(6048, 1), (3024, 1)]
W = "{#donations}{Donor}"


def run(text: str) -> str:
    return f'<w:r><w:t xml:space="preserve">{text}</w:t></w:r>'


def row(name: str, r: int, cells: list, texts: list) -> str:
    """``cells`` is ``(tcW, gridSpan)`` per cell — a ``tcW`` of 0 writes none —
    ``texts`` the content under each label."""
    tcs = ""
    for c, ((w, span), text) in enumerate(zip(cells, texts), 1):
        span_xml = f'<w:gridSpan w:val="{span}"/>' if span > 1 else ""
        tcs += (
            f'<w:tc><w:tcPr>{f'<w:tcW w:w="{w}" w:type="dxa"/>' if w else ""}{span_xml}</w:tcPr>'
            f'<w:p>{run(f"{name}{r}{c}")}</w:p><w:p>{run(text)}</w:p></w:tc>'
        )
    return f"<w:tr>{tcs}</w:tr>"


def table(tbl_w: str, layout: str, grid: list, rows: list) -> str:
    cols = "".join(f'<w:gridCol w:w="{g}"/>' for g in grid)
    return (
        f"<w:tbl><w:tblPr>{tbl_w}<w:tblBorders>{BORDERS}</w:tblBorders>{layout}"
        '<w:tblLook w:val="0000" w:firstRow="0" w:lastRow="0" w:firstColumn="0" w:lastColumn="0" '
        'w:noHBand="1" w:noVBand="1"/></w:tblPr>'
        f"<w:tblGrid>{cols}</w:tblGrid>{''.join(rows)}</w:tbl>"
    )


def case(name: str, tbl_w: str, layout: str, grid: list, rows: list) -> str:
    """``rows`` is ``(cells, texts)`` per row."""
    trs = [row(name, r, cells, texts) for r, (cells, texts) in enumerate(rows, 1)]
    return f"<w:p>{run('SEP' + name)}</w:p>{table(tbl_w, layout, grid, trs)}"


def main() -> None:
    six = [(1440, 1)] * 6
    report = [W, "{Email}", "{Amount}", "{Channel}", "{Method}", "{PaidDate}{/donations}"]
    body = "".join(
        [
            case("T1", AUTO, FIXED, GRID_A, [(B, ["x", "x"]), (B, ["x", "x"])]),
            case("T2", AUTO, "", GRID_A, [(B, ["x", "x"]), (B, ["x", "x"])]),
            case("T3", AUTO, FIXED, GRID_A, [(B, ["x", "x"]), (C, ["x", "x"])]),
            case("T4", AUTO, "", GRID_A, [([(4824, 1), (4824, 1)], ["x", "x"]), (B, ["x", "x"])]),
            case("T5", '<w:tblW w:w="9648" w:type="dxa"/>', FIXED, GRID_A, [(B, ["x", "x"]), (B, ["x", "x"])]),
            case("T6", AUTO, FIXED, [1440] * 6, [(six, [W] + ["{x}"] * 5)]),
            case(
                "T7",
                AUTO,
                FIXED,
                [3024] * 3,
                [([(7000, 2), (2048, 1)], ["x", "x"]), ([(3024, 1)] * 3, ["x", "x", "x"])],
            ),
            case("T8", AUTO, "", [1440] * 6, [(six, report)]),
            case("T9", '<w:tblW w:w="7200" w:type="dxa"/>', "", [1440] * 6, [(six, ["{x}"] * 6)]),
            case("T10", '<w:tblW w:w="7200" w:type="dxa"/>', "", [1440] * 6, [(six, report)]),
            case(
                "T11",
                '<w:tblW w:w="4000" w:type="dxa"/>',
                "",
                [2000, 6000],
                [([(2000, 1), (6000, 1)], ["L", "R"])],
            ),
            case(
                "T12",
                '<w:tblW w:w="8000" w:type="dxa"/>',
                "",
                [2000] * 4,
                [([(4000, 1), (4000, 1)], ["x", "x"])],
            ),
            case("T13", '<w:tblW w:w="8000" w:type="dxa"/>', "", [2000] * 4, [([(0, 3)], ["x"])]),
        ]
    )
    body += f"<w:p>{run('END')}</w:p>"
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
