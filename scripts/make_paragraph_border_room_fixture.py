#!/usr/bin/env python3
"""Build ``test-files/paragraph-border-room.docx``, which asks how much room a
paragraph's top and bottom borders take (§17.3.1.24, §17.3.1.42/§17.3.1.7).

dxpdf charged a paragraph's top or bottom border its ``w:space`` alone, on
the reading that only the space separates line from text. The clerk
summary reports put an empty paragraph with a 0.75pt bottom border 3pt
below its text between two tables, and Word sets the second table 0.8pt
lower than that reading does — as though the line's own width took room
as well.

Every line is Calibri 10pt on an exact 12pt line. Each case is a bordered
paragraph between two plain ones, ``<case>a`` / ``<case>`` / ``<case>z``,
so the pitch from ``a`` to the case is the top border's room plus 12, and
from the case to ``z`` the bottom's:

  T<s><w>  ``w:top`` with ``w:space`` s (0 or 3pt) and width w (6 or 24
           eighths: 0.75 or 3pt)
  B<s><w>  the same as ``w:bottom``
  E        an empty paragraph with the clerk reports' bottom border
           (``sz="6" space="3"``) between two plain ones

Regenerate and commit the result if the content changes; the build is
deterministic and needs no third-party packages.

    scripts/make_paragraph_border_room_fixture.py
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


def para(text: str, ppr: str = "") -> str:
    ppr_xml = f"<w:pPr>{ppr}</w:pPr>" if ppr else ""
    run = f"<w:r><w:t>{text}</w:t></w:r>" if text else ""
    return f"<w:p>{ppr_xml}{run}</w:p>"


def border(side: str, sz: int, space: int) -> str:
    return (
        f'<w:pBdr><w:{side} w:val="single" w:sz="{sz}" w:space="{space}" w:color="000000"/></w:pBdr>'
    )


def body() -> str:
    out = []
    for side, letter in (("top", "T"), ("bottom", "B")):
        for space in (0, 3):
            for sz in (6, 24):
                name = f"{letter}{space}{sz}"
                out += [para(f"{name}a"), para(name, border(side, sz, space)), para(f"{name}z")]
    out += [para("Ea"), para("", border("bottom", 6, 3)), para("Ez"), para("END")]
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
    out = ROOT / "test-files" / "paragraph-border-room.docx"
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as z:
        for part, content in parts.items():
            info = zipfile.ZipInfo(part, date_time=(1980, 1, 1, 0, 0, 0))
            info.compress_type = zipfile.ZIP_DEFLATED
            z.writestr(info, content.encode("utf-8"))
    print(f"wrote {out.relative_to(ROOT)}")


if __name__ == "__main__":
    main()
