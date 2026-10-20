#!/usr/bin/env python3
"""Build ``test-files/page-bottom-spacing.docx``, which asks whether Word keeps
a line on a page when its text fits above the bottom margin but the room its
line spacing adds below the text does not.

A §17.3.1.33 ``auto`` line is its text's height scaled by ``w:line/240``, and
Word puts the baseline at the text's ascent below the line's top, so the
scaled extra is all below the text. A real-world letter template ends a page
with a Calibri 11pt line at 1.15 whose text ends at 754.87pt and whose extra
would end at 756.89pt, past the 756pt margin, and keeps it there.

Each case is one page: an ``exact`` spacer paragraph that puts the top of the
case's line ``room`` points above the bottom margin, then the line ``<case>``
(Calibri 11pt: 10.47pt up, 2.95pt down, 13.43pt of text), then ``<case>z``,
which only shows where the page after it began. Letter paper, 1in margins, so
the bottom margin is at 720pt.

  A115  1.15,   room 14.4 — the text fits, the extra (2.01pt) does not
  A15   1.5,    room 17   — the text fits, the extra (6.71pt) does not
  A2    double, room 20   — the text fits, the extra (13.43pt) does not
  A115n 1.15,   room 13   — the text does not fit (control: moves)
  S1    single, room 13   — the text does not fit (control: moves)
  S1y   single, room 13.6 — the text just fits (control: stays)
  X20   exact 20pt,    room 18 — the line's box does not fit
  L20   at least 20pt, room 18 — the line's box does not fit
  AW    1.15, room 14.4, the last of three lines of one paragraph with
        widow control off — whether a split paragraph keeps it too

Measured in Word (Word for Mac 16, PDF export, 2026-10-16): A115, A15, A2,
S1y and AW's last line stay on their page; A115n, S1, X20 and L20 move.

Regenerate and commit the result if the content changes; the build is
deterministic and needs no third-party packages.

    scripts/make_page_bottom_spacing_fixture.py
"""

import pathlib
import zipfile

ROOT = pathlib.Path(__file__).resolve().parent.parent
OUT = ROOT / "test-files" / "page-bottom-spacing.docx"

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
<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:cs="Calibri"/><w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr></w:rPrDefault>
<w:pPrDefault><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>
<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
</w:styles>"""

SECT = (
    '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" '
    'w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>'
)
TOP, BOTTOM = 72.0, 720.0

SPACING = {
    "s1": 'w:line="240" w:lineRule="auto"',
    "s115": 'w:line="276" w:lineRule="auto"',
    "s15": 'w:line="360" w:lineRule="auto"',
    "s2": 'w:line="480" w:lineRule="auto"',
    "x20": 'w:line="400" w:lineRule="exact"',
    "l20": 'w:line="400" w:lineRule="atLeast"',
}

# (case, spacing, room above the bottom margin, lines in the paragraph)
CASES = [
    ("A115", "s115", 14.4, 1),
    ("A15", "s15", 17.0, 1),
    ("A2", "s2", 20.0, 1),
    ("A115n", "s115", 13.0, 1),
    ("S1", "s1", 13.0, 1),
    ("S1y", "s1", 13.6, 1),
    ("X20", "x20", 18.0, 1),
    ("L20", "l20", 18.0, 1),
    ("AW", "s115", 14.4, 3),
]
# Calibri 11pt at 1.15: what two earlier lines of a three-line paragraph take.
LINE_115 = (1950 + 550) / 2048 * 11 * 1.15


def run(text: str) -> str:
    return f'<w:r><w:t xml:space="preserve">{text}</w:t></w:r>'


def spacer(height: float, first: bool) -> str:
    twips = round(height * 20)
    brk = "" if first else "<w:pageBreakBefore/>"
    return (
        f'<w:p><w:pPr>{brk}<w:spacing w:after="0" w:line="{twips}" w:lineRule="exact"/></w:pPr>'
        f"{run('SP')}</w:p>"
    )


def case(name: str, spacing: str, room: float, lines: int, first: bool) -> str:
    before = (lines - 1) * LINE_115
    out = [spacer(BOTTOM - TOP - room - before, first)]
    widow = '<w:widowControl w:val="0"/>' if lines > 1 else ""
    body = "<w:r><w:br/></w:r>".join(
        run(f"{name}{chr(ord('a') + i)}" if lines > 1 else name) for i in range(lines)
    )
    out.append(f'<w:p><w:pPr>{widow}<w:spacing w:after="0" {SPACING[spacing]}/></w:pPr>{body}</w:p>')
    out.append(f"<w:p>{run(name + 'z')}</w:p>")
    return "".join(out)


def main() -> None:
    body = "".join(case(*c, first=(i == 0)) for i, c in enumerate(CASES))
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
