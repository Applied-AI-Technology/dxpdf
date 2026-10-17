#!/usr/bin/env python3
"""Build ``test-files/oversized-word.docx``, which asks how Word breaks text
that has a line-break opportunity inside a word too wide for its line.

Word's ReceiptTemplate sets ``${amount}{/receiptItems}`` in a 34.2pt cell as
``${amoun`` / ``t}{/recei`` / ``ptItems}``: having cut ``${amount}``, it fills
the second line past the opportunity after ``/`` instead of breaking there,
because what follows, ``receiptItems}``, is wider than the cell and would be
cut anyway. Its ``{signature}{signature2}{/`` / ``owner}`` breaks at the same
opportunity, where what follows fits a line of its own. Two readings fit both:
Word cuts only a *piece* between opportunities that is too wide for a line, or
it cuts the whole space-delimited *word* once any of it is. These cases take
them apart, each in a fixed-layout one-cell table 72pt wide with no cell
margins, Calibri 11pt, under a ``SEP<case>`` paragraph:

  O1  ``abcd/`` then 30 letters — a short piece, then a long one
  O2  30 letters then ``/ab`` — a long piece, then a short one that fits
      beside the long piece's last line
  O3  ``abcd/efghijk/`` then 30 letters — two short pieces on one line
  O4  ``Word`` and a space, then 30 letters — a space before a long word
  O5  ``abcd-`` then 30 letters — a hyphen, not a slash
  O6  30 letters, a space, ``ab`` — a long word, then a short one
  O7  ``abcdefgh/ijklmnop/qrstuvwx`` — three pieces that each fit the line
      and a word that does not: broken at both ``/`` if Word cuts pieces,
      filled line by line if it cuts words

Measured in Word (Word for Mac 16, PDF export, 2026-10-16): O1
``abcd/ABCDEFG`` / ``HIJKLMNOPQRS`` / ``TUVWXYZabcd``; O3 fills past both
``/``; O4 ``Word`` and O5 ``abcd-`` end their lines and the long word starts
the next; O7 ``abcdefgh/ijklm`` / ``nop/qrstuvwx``. Word does not break after a
solidus; a space and a hyphen are opportunities.

Regenerate and commit the result if the content changes; the build is
deterministic and needs no third-party packages.

    scripts/make_oversized_word_fixture.py
"""

import pathlib
import zipfile

ROOT = pathlib.Path(__file__).resolve().parent.parent
OUT = ROOT / "test-files" / "oversized-word.docx"

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

# Thirty letters with no repeats closer than ten apart, so each line's last
# letter says exactly where it was cut.
LONG = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcd"

CASES = [
    ("O1", "abcd/" + LONG),
    ("O2", LONG + "/ab"),
    ("O3", "abcd/efghijk/" + LONG),
    ("O4", "Word " + LONG),
    ("O5", "abcd-" + LONG),
    ("O6", LONG + " ab"),
    ("O7", "abcdefgh/ijklmnop/qrstuvwx"),
]


def run(text: str) -> str:
    return f'<w:r><w:t xml:space="preserve">{text}</w:t></w:r>'


def table(text: str) -> str:
    return (
        '<w:tbl><w:tblPr><w:tblW w:w="1440" w:type="dxa"/><w:tblLayout w:type="fixed"/>'
        '<w:tblCellMar><w:left w:w="0" w:type="dxa"/><w:right w:w="0" w:type="dxa"/></w:tblCellMar>'
        '</w:tblPr><w:tblGrid><w:gridCol w:w="1440"/></w:tblGrid>'
        f'<w:tr><w:tc><w:tcPr><w:tcW w:w="1440" w:type="dxa"/></w:tcPr><w:p>{run(text)}</w:p></w:tc></w:tr></w:tbl>'
    )


def main() -> None:
    body = "".join(f"<w:p>{run('SEP' + name)}</w:p>{table(text)}" for name, text in CASES)
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
