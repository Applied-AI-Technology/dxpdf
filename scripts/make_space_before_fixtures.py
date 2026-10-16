#!/usr/bin/env python3
"""Build ``test-files/space-before-page-top-15.docx`` and
``test-files/space-before-page-top-14.docx``, which ask where Word keeps a
paragraph's §17.3.1.33 ``w:spacing/@w:before`` when the paragraph's first line
lands at the top of a page. Each labelled paragraph has 72pt before it, in
Times New Roman 12pt, and is followed by a plain ``<label>n`` line:

  S1 — the first paragraph of the document
  S2 — after a paragraph holding only a page break
  S4 — §17.3.1.23 ``pageBreakBefore``
  S7 — after a page break that ends the paragraph before (``S7pre``)
  S5 — carried to the next page by the flow, after 44 filler lines
  S6 — the first paragraph of a ``nextPage`` section (after ``S6pre``)

The two files differ only in §17.15.1.21 ``compatibilityMode``.

Regenerate and commit the result if the content changes; the build is
deterministic and needs no third-party packages.

    scripts/make_space_before_fixtures.py
"""

import pathlib
import zipfile

ROOT = pathlib.Path(__file__).resolve().parent.parent
OUT = ROOT / "test-files"

NS = (
    'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
    'xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" '
    'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="w14"'
)

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

STYLES = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles {NS}>
<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:cs="Calibri"/><w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr></w:rPrDefault>
<w:pPrDefault><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>
<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
<w:style w:type="table" w:default="1" w:styleId="TableNormal"><w:name w:val="Normal Table"/><w:tblPr><w:tblInd w:w="0" w:type="dxa"/>
<w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:left w:w="108" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar></w:tblPr></w:style>
</w:styles>"""

SECT = (
    '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" '
    'w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>'
)
PAGE_BREAK = '<w:p><w:r><w:br w:type="page"/></w:r></w:p>'
TIMES = '<w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:sz w:val="24"/></w:rPr>'
BEFORE = '<w:spacing w:before="1440"/>'


def para(text: str, ppr: str = "") -> str:
    return f'<w:p><w:pPr>{ppr}</w:pPr><w:r>{TIMES}<w:t xml:space="preserve">{text}</w:t></w:r></w:p>'


def body() -> str:
    b = [para("S1 first", BEFORE), para("S1n")]
    b += [PAGE_BREAK, para("S2 after break para", BEFORE), para("S2n")]
    b += [para("S4 pageBreakBefore", "<w:pageBreakBefore/>" + BEFORE), para("S4n")]
    b += [
        f'<w:p><w:r>{TIMES}<w:t>S7pre</w:t></w:r><w:r><w:br w:type="page"/></w:r></w:p>',
        para("S7 after break run", BEFORE),
        para("S7n"),
    ]
    b += [PAGE_BREAK] + [para(f"F{i:02d}") for i in range(44)]
    b += [para("S5 natural flow", BEFORE), para("S5n")]
    b += [
        f"<w:p><w:pPr>{SECT}</w:pPr><w:r>{TIMES}<w:t>S6pre</w:t></w:r></w:p>",
        para("S6 new section", BEFORE),
        para("S6n"),
    ]
    return "".join(b)


def settings(mode: int) -> str:
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        f'<w:settings {NS}><w:compat><w:compatSetting w:name="compatibilityMode" '
        f'w:uri="http://schemas.microsoft.com/office/word" w:val="{mode}"/></w:compat></w:settings>'
    )


def write(mode: int) -> None:
    target = OUT / f"space-before-page-top-{mode}.docx"
    document = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        f"<w:document {NS}><w:body>{body()}{SECT}</w:body></w:document>"
    )
    # Fixed timestamps so regenerating an unchanged fixture produces identical
    # bytes and does not show up as a diff.
    with zipfile.ZipFile(target, "w", zipfile.ZIP_DEFLATED) as z:
        for part, data in (
            ("[Content_Types].xml", CONTENT_TYPES),
            ("_rels/.rels", ROOT_RELS),
            ("word/_rels/document.xml.rels", DOC_RELS),
            ("word/document.xml", document),
            ("word/styles.xml", STYLES),
            ("word/settings.xml", settings(mode)),
        ):
            info = zipfile.ZipInfo(part, date_time=(1980, 1, 1, 0, 0, 0))
            info.compress_type = zipfile.ZIP_DEFLATED
            z.writestr(info, data)
    print(f"wrote {target.relative_to(ROOT)} ({target.stat().st_size} bytes)")


def main() -> None:
    OUT.mkdir(parents=True, exist_ok=True)
    for mode in (15, 14):
        write(mode)


if __name__ == "__main__":
    main()
