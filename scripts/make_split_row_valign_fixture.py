#!/usr/bin/env python3
"""Build ``test-files/split-row-valign.docx``, which asks how Word aligns a
cell's content vertically (§17.4.84 ``w:vAlign``) in a row that breaks
across pages.

A real-world notice template has a two-cell row that Word breaks over
pages 1 and 2. Its right cell is ``vAlign="bottom"``, and Word sets the one
line of it that reaches page 2 at the top of the continuation, level with
the left cell's first line there, where bottom alignment would put it two
lines lower. The document cannot show what Word does with the first half, so
the probe asks for both halves and for centring.

Every line is Calibri 10pt on an exact 12pt line, no space between
paragraphs and no widow control, so the page holds 54 lines between its
72pt margins. Each case is on a page of its own, after 44 filler lines, so
its row starts 12 lines above the bottom margin; the bordered row's left
cell always holds 20 lines ``<case>L<n>``, the right one ``<case>R<n>``:

  V1  right 6 lines, bottom — the right cell fits in the first half
  V2  right 14 lines, bottom — it straddles the break, 2 lines carried
  V3  right 14 lines, center
  V4  left 4 lines, right 2, bottom — a row on one page, the control

Regenerate and commit the result if the content changes; the build is
deterministic and needs no third-party packages.

    scripts/make_split_row_valign_fixture.py
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
    '<w:pPrDefault><w:pPr><w:widowControl w:val="0"/>'
    '<w:spacing w:after="0" w:line="240" w:lineRule="exact"/></w:pPr></w:pPrDefault>'
    "</w:docDefaults>"
    '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>'
    "</w:styles>"
)

SECT = (
    '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" '
    'w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>'
)

BORDER = '<w:{0} w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
TBL_PR = (
    '<w:tblPr><w:tblW w:w="9360" w:type="dxa"/><w:tblLayout w:type="fixed"/><w:tblBorders>'
    + "".join(BORDER.format(s) for s in ("top", "left", "bottom", "right", "insideH", "insideV"))
    + '</w:tblBorders><w:tblCellMar><w:left w:w="108" w:type="dxa"/><w:right w:w="108" w:type="dxa"/>'
    "</w:tblCellMar></w:tblPr>"
)


def para(text: str, page_break: bool = False) -> str:
    ppr = "<w:pPr><w:pageBreakBefore/></w:pPr>" if page_break else ""
    return f"<w:p>{ppr}<w:r><w:t>{text}</w:t></w:r></w:p>"


def cell(lines: list[str], valign: str = "") -> str:
    va = f'<w:vAlign w:val="{valign}"/>' if valign else ""
    return f'<w:tc><w:tcPr><w:tcW w:w="4680" w:type="dxa"/>{va}</w:tcPr>{"".join(para(t) for t in lines)}</w:tc>'


def case(name: str, left: int, right: int, valign: str) -> str:
    filler = [para(f"{name} filler 1", page_break=True)]
    filler += [para(f"{name} filler {n}") for n in range(2, 45)]
    row = cell([f"{name}L{n}" for n in range(1, left + 1)]) + cell(
        [f"{name}R{n}" for n in range(1, right + 1)], valign
    )
    table = (
        f"<w:tbl>{TBL_PR}<w:tblGrid><w:gridCol w:w=\"4680\"/><w:gridCol w:w=\"4680\"/></w:tblGrid>"
        f"<w:tr>{row}</w:tr></w:tbl>"
    )
    return "".join(filler) + table + para(f"{name} after")


def main() -> None:
    body = "".join(
        [
            case("V1", 20, 6, "bottom"),
            case("V2", 20, 14, "bottom"),
            case("V3", 20, 14, "center"),
            case("V4", 4, 2, "bottom"),
        ]
    )
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
    out = ROOT / "test-files" / "split-row-valign.docx"
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as z:
        for part, content in parts.items():
            info = zipfile.ZipInfo(part, date_time=(1980, 1, 1, 0, 0, 0))
            info.compress_type = zipfile.ZIP_DEFLATED
            z.writestr(info, content.encode("utf-8"))
    print(f"wrote {out.relative_to(ROOT)}")


if __name__ == "__main__":
    main()
