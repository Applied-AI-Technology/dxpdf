#!/usr/bin/env python3
"""Build ``test-files/border-margin-room.docx``, which asks how much room a
horizontal border between two rows takes when the cells on either side have
top and bottom margins (§17.4.38, §17.4.42/§17.4.39), and whether a
paragraph's space after holds above a table.

``test-files/row-border-room.docx`` settled the border with no cell margins:
every interior border takes its width. Some report templates' tables put a
0.5pt ``tcBorders/bottom`` under their header row and 29-twip (1.45pt) top and
bottom ``tblCellMar``, and Word sets each such row about 0.5pt shorter than
margins plus border would make it. Their tables also follow a paragraph
with ``w:after="20"``, which Word appears not to keep.

Every line is Calibri 10pt on an exact 12pt line. Each case is a label
paragraph and a one-column, three-row table whose rows hold ``<case>r<n>``:
the border, where there is one, is under row 1 only, so row 2's pitch to
row 3 is the control and row 1's to row 2 differs by the border's room.

  B<m><b>  cell margins m (0 = none, 1 = 29 twips, 5 = 100 twips) top and
           bottom; b = t: a 0.5pt ``tcBorders/bottom`` on row 1's cell,
           b = T: the same at 3pt, b = i: a 0.5pt ``insideH`` on the table
           (all three of its interior edges then ruled — read row 1→2 only)
  C<m><b>  the row 1/2 edge stated as row 2's ``tcBorders/top`` instead
  F<m><b>  row 3's ``tcBorders/bottom`` — the table's foot, read to ``<case>z``
  G<m><b>  the table's own ``tblBorders/bottom`` — the same foot
  H<m><b>  the table's own ``tblBorders/top`` — read label to row 1
  S<a>     a label paragraph with ``w:after`` a (0 or 20 twips), then a
           margin-less, borderless table: where its first line sits
  K20      S20 with ``w:contextualSpacing`` on both the label and the cells
  L20      the same with it on the label only
  P1-P3    two ``w:contextualSpacing`` Normal paragraphs with no ``pStyle``
           and 1pt after, then a plain one

Regenerate and commit the result if the content changes; the build is
deterministic and needs no third-party packages.

    scripts/make_border_margin_room_fixture.py
"""

import pathlib
import zipfile

ROOT = pathlib.Path(__file__).resolve().parent.parent

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

SETTINGS = (
    f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings {NS}><w:compat>'
    '<w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/>'
    "</w:compat></w:settings>"
)

STYLES = (
    f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles {NS}>'
    "<w:docDefaults>"
    '<w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:cs="Calibri"/>'
    '<w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr></w:rPrDefault>'
    '<w:pPrDefault><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="exact"/></w:pPr></w:pPrDefault>'
    "</w:docDefaults>"
    '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>'
    "</w:styles>"
)

SECT = (
    '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" '
    'w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>'
)

MARGINS = {"0": None, "1": 29, "5": 100}


def para(text: str, ppr: str = "") -> str:
    ppr_xml = f"<w:pPr>{ppr}</w:pPr>" if ppr else ""
    return f"<w:p>{ppr_xml}<w:r><w:t>{text}</w:t></w:r></w:p>"


def table(
    name: str,
    margin: int | None,
    tbl_borders: str = "",
    cell_borders: dict[int, str] | None = None,
    cell_ppr: str = "",
) -> str:
    mar_w = 0 if margin is None else margin
    mar = (
        f'<w:tblCellMar><w:top w:w="{mar_w}" w:type="dxa"/>'
        f'<w:bottom w:w="{mar_w}" w:type="dxa"/></w:tblCellMar>'
    )
    borders = f"<w:tblBorders>{tbl_borders}</w:tblBorders>" if tbl_borders else ""
    tbl_pr = f'<w:tblPr><w:tblW w:w="4000" w:type="dxa"/><w:tblLayout w:type="fixed"/>{borders}{mar}</w:tblPr>'
    rows = []
    for n in (1, 2, 3):
        edge = (cell_borders or {}).get(n, "")
        tc_borders = f"<w:tcBorders>{edge}</w:tcBorders>" if edge else ""
        rows.append(
            f'<w:tr><w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/>{tc_borders}</w:tcPr>'
            f"{para(f'{name}r{n}', cell_ppr)}</w:tc></w:tr>"
        )
    return f'<w:tbl>{tbl_pr}<w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid>{"".join(rows)}</w:tbl>'


def edge(side: str, sz: int) -> str:
    return f'<w:{side} w:val="single" w:sz="{sz}" w:space="0" w:color="000000"/>'


SIZES = {"t": 4, "T": 24}


def body() -> str:
    out = []
    for m, margin in MARGINS.items():
        for b, sz in SIZES.items():
            name = f"B{m}{b}"
            out += [para(name), table(name, margin, cell_borders={1: edge("bottom", sz)})]
        name = f"B{m}i"
        out += [para(name), table(name, margin, tbl_borders=edge("insideH", 4))]
    for m, margin in (("1", 29), ("5", 100)):
        for b, sz in SIZES.items():
            # The same edges stated the other ways round, and the table's foot.
            cases = [
                (f"C{m}{b}", {}, {2: edge("top", sz)}),
                (f"F{m}{b}", {}, {3: edge("bottom", sz)}),
                (f"G{m}{b}", edge("bottom", sz), {}),
                (f"H{m}{b}", edge("top", sz), {}),
            ]
            for name, tbl, cells in cases:
                out += [para(name), table(name, margin, tbl_borders=tbl, cell_borders=cells), para(f"{name}z")]
    for after in (0, 20):
        name = f"S{after}"
        out += [para(name, f'<w:spacing w:after="{after}"/>'), table(name, None)]
    out += [
        para("K20", '<w:spacing w:after="20"/><w:contextualSpacing/>'),
        table("K20", None, cell_ppr="<w:contextualSpacing/>"),
        para("L20", '<w:spacing w:after="20"/><w:contextualSpacing/>'),
        table("L20", None),
        para("P1", '<w:spacing w:after="20"/><w:contextualSpacing/>'),
        para("P2", '<w:spacing w:after="20"/><w:contextualSpacing/>'),
        para("P3"),
        para("END"),
    ]
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
        "word/settings.xml": SETTINGS,
    }
    out = ROOT / "test-files" / "border-margin-room.docx"
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as z:
        for part, content in parts.items():
            info = zipfile.ZipInfo(part, date_time=(1980, 1, 1, 0, 0, 0))
            info.compress_type = zipfile.ZIP_DEFLATED
            z.writestr(info, content.encode("utf-8"))
    print(f"wrote {out.relative_to(ROOT)}")


if __name__ == "__main__":
    main()
