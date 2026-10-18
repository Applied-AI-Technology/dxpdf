#!/usr/bin/env python3
"""Build ``test-files/vmerge-excess.docx``, which asks which row grows when a
vertically merged cell holds more than the rows it spans.

albanyny's ``DogLicenseDeliquencyNotice`` (and its renewal notice, in every
tenant) closes with a table whose first column is one cell merged down three
rows of at least 945, 350 and 260 twips, holding six paragraphs; beside it,
the signature lines are the bottom border of a row-1 cell and the top border
of the row-2 cell under it, and row 3 holds a bottom-aligned italic note.
The text in the merged cell stands where Word puts it, but the note and the
table's bottom sit 0.4–0.6pt lower than Word's.

Each table ``V<n>`` has two columns and no table borders: column A one cell
merged down every row, column B a cell per row — row 1's empty with a 0.5pt
bottom border, row 2's ``V<n>B2`` with a 0.5pt top border, row 3's
``V<n>B3`` bottom-aligned. ``SEP<n>`` follows each table:

  V1  the notice's: rows at least 945/350/260, column A an empty Calibri
      11pt paragraph, two of 10pt bold and three of 9pt — ``V1A1``…``V1A5``
  V2  V1 with column A holding only ``V2A1``, shorter than the rows
  V3  V1 with column A holding ten 9pt paragraphs ``V3A1``…``V3A10``
  V4  V1 with no row heights
  V5  V1 with ``V5B3`` top-aligned

Calibri 11pt, ``w:line`` 276 (auto), no paragraph spacing.

Regenerate and commit the result if the content changes; the build is
deterministic and needs no third-party packages.

    scripts/make_vmerge_excess_fixture.py
"""

import pathlib
import zipfile

ROOT = pathlib.Path(__file__).resolve().parent.parent
OUT = ROOT / "test-files" / "vmerge-excess.docx"

W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'

CONTENT_TYPES = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
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
</Relationships>"""

STYLES = (
    f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles {W}>'
    "<w:docDefaults>"
    '<w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:cs="Calibri"/>'
    '<w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr></w:rPrDefault>'
    '<w:pPrDefault><w:pPr><w:spacing w:after="0" w:line="276" w:lineRule="auto"/></w:pPr></w:pPrDefault>'
    "</w:docDefaults>"
    '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>'
    '<w:style w:type="table" w:default="1" w:styleId="TableNormal"><w:name w:val="Normal Table"/>'
    '<w:tblPr><w:tblInd w:w="0" w:type="dxa"/><w:tblCellMar><w:top w:w="0" w:type="dxa"/>'
    '<w:left w:w="108" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/><w:right w:w="108" w:type="dxa"/>'
    "</w:tblCellMar></w:tblPr></w:style>"
    "</w:styles>"
)

SECT = (
    '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" '
    'w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>'
)

LINE = 'w:val="single" w:sz="4" w:space="0" w:color="000000"'


def para(text: str = "", half_points: int | None = None, bold: bool = False) -> str:
    props = ("<w:b/>" if bold else "") + (
        f'<w:sz w:val="{half_points}"/><w:szCs w:val="{half_points}"/>' if half_points else ""
    )
    rpr = f"<w:rPr>{props}</w:rPr>" if props else ""
    runs = f'<w:r>{rpr}<w:t xml:space="preserve">{text}</w:t></w:r>' if text else ""
    return f"<w:p><w:pPr>{rpr}</w:pPr>{runs}</w:p>"


def tc(width: int, content: str, props: str = "") -> str:
    return f'<w:tc><w:tcPr><w:tcW w:w="{width}" w:type="dxa"/>{props}</w:tcPr>{content}</w:tc>'


def merged_content(n: int, kind: str) -> str:
    if kind == "short":
        return para(f"V{n}A1", 18)
    if kind == "tall":
        return "".join(para(f"V{n}A{i}", 18) for i in range(1, 11))
    return (
        para()
        + para(f"V{n}A1", 20, bold=True)
        + para(f"V{n}A2", 20, bold=True)
        + "".join(para(f"V{n}A{i}", 18) for i in range(3, 6))
    )


def table(n: int, kind: str = "notice", heights=(945, 350, 260), b3_align: str = "bottom") -> str:
    a, b = 3600, 3060
    rows = []
    for r, height in enumerate(heights, start=1):
        tr_pr = f'<w:trPr><w:trHeight w:val="{height}"/></w:trPr>' if height else ""
        merge = '<w:vMerge w:val="restart"/>' if r == 1 else "<w:vMerge/>"
        first = tc(a, merged_content(n, kind) if r == 1 else para(), merge)
        if r == 1:
            second = tc(b, para(), f"<w:tcBorders><w:bottom {LINE}/></w:tcBorders>")
        elif r == 2:
            second = tc(b, para(f"V{n}B2"), f"<w:tcBorders><w:top {LINE}/></w:tcBorders>")
        else:
            second = tc(b, para(f"V{n}B3", bold=False), f'<w:vAlign w:val="{b3_align}"/>')
        rows.append(f"<w:tr>{tr_pr}{first}{second}</w:tr>")
    return (
        "<w:tbl><w:tblPr>"
        f'<w:tblW w:w="{a + b}" w:type="dxa"/><w:tblLayout w:type="fixed"/>'
        f'</w:tblPr><w:tblGrid><w:gridCol w:w="{a}"/><w:gridCol w:w="{b}"/></w:tblGrid>'
        + "".join(rows)
        + "</w:tbl>"
        + para(f"SEP{n}")
    )


def body() -> str:
    return (
        para("REF")
        + table(1)
        + table(2, kind="short")
        + table(3, kind="tall")
        + table(4, heights=(0, 0, 0))
        + table(5, b3_align="top")
    )


def main() -> None:
    document = (
        f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        f"<w:document {W}><w:body>{body()}{SECT}</w:body></w:document>"
    )
    parts = {
        "[Content_Types].xml": CONTENT_TYPES,
        "_rels/.rels": ROOT_RELS,
        "word/_rels/document.xml.rels": DOC_RELS,
        "word/document.xml": document,
        "word/styles.xml": STYLES,
    }
    with zipfile.ZipFile(OUT, "w", zipfile.ZIP_DEFLATED) as z:
        for part, content in parts.items():
            info = zipfile.ZipInfo(part, date_time=(1980, 1, 1, 0, 0, 0))
            info.compress_type = zipfile.ZIP_DEFLATED
            z.writestr(info, content.encode("utf-8"))
    print(f"wrote {OUT.relative_to(ROOT)}")


if __name__ == "__main__":
    main()
