#!/usr/bin/env python3
"""Build ``test-files/symbol-bullet.docx``, which asks how tall a line is when
its list label is set in Symbol or Wingdings — the faces Word's own bullet
gallery uses, whose glyphs sit at private-use codepoints that a renderer
without the face remaps to a Unicode bullet in a text face.

Calibri 11pt paragraphs at 1.15 auto spacing and no paragraph spacing, each
between Calibri reference lines ``R<n>``, so the distance from one reference to
the next is the height of the lines between them:

  S1  a Symbol bullet (U+F0B7) at 11pt, over two lines (``S1a`` / ``S1b``
      after a ``<w:br/>``) — the second carries no label
  S2  the same bullet on a one-line paragraph
  L1  a Symbol bullet at 22pt
  G1  a Wingdings bullet (U+F0A7) at 22pt
  G2  a Wingdings bullet at 11pt

Word's announcement-letter template answers the 11pt Symbol case already: each
bullet line is 16.08pt where a plain Calibri line is 15.36–15.60pt, Symbol's
ascent (2059 of 2048 units, above Calibri's 1950) over Calibri's descent.

Regenerate and commit the result if the content changes; the build is
deterministic and needs no third-party packages.

    scripts/make_symbol_bullet_fixture.py
"""

import pathlib
import zipfile

ROOT = pathlib.Path(__file__).resolve().parent.parent
OUT = ROOT / "test-files" / "symbol-bullet.docx"

NS = (
    'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
    'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
)

CONTENT_TYPES = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>
<Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>
</Types>"""

ROOT_RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>"""

DOC_RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rIdS" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
<Relationship Id="rIdT" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>
<Relationship Id="rIdN" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/>
</Relationships>"""

SETTINGS = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings {NS}><w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat></w:settings>"""

STYLES = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles {NS}>
<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:cs="Calibri"/><w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr></w:rPrDefault>
<w:pPrDefault><w:pPr><w:spacing w:after="0" w:line="276" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>
<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
</w:styles>"""

# One abstract list per label: (face, character, label size in half-points).
LABELS = [
    ("Symbol", "", 22),
    ("Symbol", "", 44),
    ("Wingdings", "", 44),
    ("Wingdings", "", 22),
]


def numbering() -> str:
    abstract, nums = [], []
    for i, (face, char, half) in enumerate(LABELS, 1):
        abstract.append(
            f'<w:abstractNum w:abstractNumId="{i}"><w:multiLevelType w:val="hybridMultilevel"/>'
            f'<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="bullet"/>'
            f'<w:lvlText w:val="{char}"/><w:lvlJc w:val="left"/>'
            f'<w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr>'
            f'<w:rPr><w:rFonts w:ascii="{face}" w:hAnsi="{face}" w:hint="default"/>'
            f'<w:sz w:val="{half}"/><w:szCs w:val="{half}"/></w:rPr></w:lvl></w:abstractNum>'
        )
        nums.append(f'<w:num w:numId="{i}"><w:abstractNumId w:val="{i}"/></w:num>')
    return (
        f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:numbering {NS}>'
        + "".join(abstract)
        + "".join(nums)
        + "</w:numbering>"
    )


SECT = (
    '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" '
    'w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>'
)


def run(text: str) -> str:
    return f'<w:r><w:t xml:space="preserve">{text}</w:t></w:r>'


def para(text: str) -> str:
    return f"<w:p>{run(text)}</w:p>"


def item(num: int, *lines: str) -> str:
    body = "<w:r><w:br/></w:r>".join(run(t) for t in lines)
    return (
        f'<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="{num}"/></w:numPr></w:pPr>'
        f"{body}</w:p>"
    )


def body() -> str:
    return "".join(
        [
            para("R0"),
            item(1, "S1a", "S1b"),
            item(1, "S2"),
            para("R1"),
            item(2, "L1"),
            para("R2"),
            item(3, "G1"),
            para("R3"),
            item(4, "G2"),
            para("R4"),
        ]
    )


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
        "word/numbering.xml": numbering(),
    }
    with zipfile.ZipFile(OUT, "w", zipfile.ZIP_DEFLATED) as z:
        for name, text in parts.items():
            info = zipfile.ZipInfo(name, date_time=(1980, 1, 1, 0, 0, 0))
            info.compress_type = zipfile.ZIP_DEFLATED
            z.writestr(info, text.encode("utf-8"))
    print(f"wrote {OUT.relative_to(ROOT)}")


if __name__ == "__main__":
    main()
