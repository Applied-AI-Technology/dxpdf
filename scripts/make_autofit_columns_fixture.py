#!/usr/bin/env python3
"""Build ``test-files/autofit-columns.docx``, which asks how Word sizes the
columns of an autofit table (§17.4.52 ``tblLayout`` absent or ``autofit``)
whose content does not fit the widths its cells prefer.

A report template has six columns, each ``tcW`` 1440 twips (72pt), in a
``tblW`` ``auto`` table, and two of them hold one unbreakable merge tag wider
than 72pt. Word draws the columns 94.56, 57.12, 62.16, 62.16, 62.16 and
93.84pt wide — the tags' columns widened to hold them, the rest narrowed to
keep the 432pt total — where the declared grid would give 72pt each.

Each case is a one-row table in Calibri 10pt, each cell a label paragraph
``<case><column>`` over its content, with ``TableNormal``'s 5.4pt left
and right cell margins, every ``tcW`` 1440 twips, under a ``SEP<case>``
paragraph. ``W`` stands for a 19-character merge tag that opens a loop (about
84pt, no opportunity); ``w`` for ``{x}`` (short).

  F1  ``tblW auto``: W w w w w w — one wide column
  F2  ``tblW auto``: W w w w w W — the report's shape
  F3  ``tblW auto``: w w w w w w — the control, nothing too wide
  F4  ``tblW auto``: a four-word sentence (spaces, each word short), then
      w w w w w — wide content that can wrap
  F5  ``tblW 5000 pct``: W w w w w w — a fixed total
  F6  ``tblW auto``: W, then five cells holding w of different lengths
      (``{a}``, ``{abcdef}``, ``{abcdefghijk}``, ``{ab}``, ``{abcd}``) — how
      the narrowed columns share what is left

Regenerate and commit the result if the content changes; the build is
deterministic and needs no third-party packages.

    scripts/make_autofit_columns_fixture.py
"""

import pathlib
import zipfile

ROOT = pathlib.Path(__file__).resolve().parent.parent
OUT = ROOT / "test-files" / "autofit-columns.docx"

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

W = "{#customers}{Buyer}"
BORDERS = "".join(
    f'<w:{s} w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
    for s in ("top", "left", "bottom", "right", "insideH", "insideV")
)


def run(text: str) -> str:
    return f'<w:r><w:t xml:space="preserve">{text}</w:t></w:r>'


def table(name: str, cells: list, tbl_w: str) -> str:
    tcs = "".join(
        f'<w:tc><w:tcPr><w:tcW w:w="1440" w:type="dxa"/></w:tcPr><w:p>{run(f"{name}{i}")}</w:p><w:p>{run(t)}</w:p></w:tc>'
        for i, t in enumerate(cells, 1)
    )
    grid = "".join('<w:gridCol w:w="1440"/>' for _ in cells)
    return (
        f"<w:tbl><w:tblPr>{tbl_w}<w:tblBorders>{BORDERS}</w:tblBorders>"
        '<w:tblLook w:val="0000" w:firstRow="0" w:lastRow="0" w:firstColumn="0" w:lastColumn="0" '
        'w:noHBand="1" w:noVBand="1"/></w:tblPr>'
        f"<w:tblGrid>{grid}</w:tblGrid><w:tr>{tcs}</w:tr></w:tbl>"
    )


AUTO = '<w:tblW w:w="0" w:type="auto"/>'
PCT = '<w:tblW w:w="5000" w:type="pct"/>'
CASES = [
    ("F1", [W, "{x}", "{x}", "{x}", "{x}", "{x}"], AUTO),
    ("F2", [W, "{x}", "{x}", "{x}", "{x}", W], AUTO),
    ("F3", ["{x}"] * 6, AUTO),
    ("F4", ["Rebuy name and address", "{x}", "{x}", "{x}", "{x}", "{x}"], AUTO),
    ("F5", [W, "{x}", "{x}", "{x}", "{x}", "{x}"], PCT),
    ("F6", [W, "{a}", "{abcdef}", "{abcdefghijk}", "{ab}", "{abcd}"], AUTO),
]


def main() -> None:
    body = "".join(f"<w:p>{run('SEP' + n)}</w:p>{table(n, c, w)}" for n, c, w in CASES)
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
