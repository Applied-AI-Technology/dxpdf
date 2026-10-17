#!/usr/bin/env python3
"""Build ``test-files/table-normal-{none,full,mar,ind,none-14}.docx``, which
ask what places a table in compatibility mode when the styles part has no
``TableNormal`` style.

``table-indent-compat-none.docx`` measured that with no ``settings.xml`` a
table's indent is to its first cell's text: the table stands one cell
margin left of the margin. Its styles part declares ``TableNormal`` with a
``tblInd`` of 0 and 108-twip margins. albanyny's CivicRecordCertificate has
no ``settings.xml`` either, but no table style at all, and Word draws its
table's left border **on** the margin with the text 5pt in — Word 2013's
placement. The files differ only in the styles part, and ``-none-14`` in
naming mode 14, so the render says which of them decides:

  none     no ``TableNormal`` — the certificate's shape
  full     ``TableNormal`` with ``tblInd`` 0 and 108-twip left/right margins
  mar      ``TableNormal`` with the margins only
  ind      ``TableNormal`` with ``tblInd`` 0 only
  none-14  ``none`` with a ``settings.xml`` naming ``compatibilityMode`` 14

Each body is Calibri 10pt, single spaced, a ``REF`` paragraph on the margin
then four two-cell tables with 0.5pt borders, each cell's text ``N<n>a``
left-aligned and ``N<n>z`` right-aligned:

  N1  216pt ``dxa``, no ``tblCellMar``
  N2  216pt, ``tblCellMar`` left/right 100 — the certificate's table
  N3  N2 with ``tblInd`` 720
  N4  ``pct`` 100% with ``tblCellMar`` left/right 100

Regenerate and commit the result if the content changes; the build is
deterministic and needs no third-party packages.

    scripts/make_table_normal_fixtures.py
"""

import pathlib
import zipfile

ROOT = pathlib.Path(__file__).resolve().parent.parent

NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'

ROOT_RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>"""


def content_types(settings: bool) -> str:
    extra = (
        '<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>'
        if settings
        else ""
    )
    return f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
{extra}</Types>"""


def doc_rels(settings: bool) -> str:
    extra = (
        '<Relationship Id="rIdT" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>'
        if settings
        else ""
    )
    return f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rIdS" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
{extra}</Relationships>"""


SETTINGS_14 = (
    f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings {NS}><w:compat>'
    '<w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="14"/>'
    "</w:compat></w:settings>"
)

IND = '<w:tblInd w:w="0" w:type="dxa"/>'
MAR = (
    '<w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:left w:w="108" w:type="dxa"/>'
    '<w:bottom w:w="0" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar>'
)


def styles(table_normal: str | None) -> str:
    table_style = (
        '<w:style w:type="table" w:default="1" w:styleId="TableNormal"><w:name w:val="Normal Table"/>'
        f"<w:tblPr>{table_normal}</w:tblPr></w:style>"
        if table_normal is not None
        else ""
    )
    return (
        f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles {NS}>'
        "<w:docDefaults>"
        '<w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:cs="Calibri"/>'
        '<w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr></w:rPrDefault>'
        '<w:pPrDefault><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:pPrDefault>'
        "</w:docDefaults>"
        '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>'
        f"{table_style}</w:styles>"
    )


SECT = (
    '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" '
    'w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>'
)


def para(text: str, ppr: str = "") -> str:
    ppr_xml = f"<w:pPr>{ppr}</w:pPr>" if ppr else ""
    return f'<w:p>{ppr_xml}<w:r><w:t xml:space="preserve">{text}</w:t></w:r></w:p>'


def table(n: int, width: str, cell_mar: bool, ind: int | None) -> str:
    edge = '<w:{0} w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
    borders = "".join(edge.format(e) for e in ("top", "left", "bottom", "right", "insideH", "insideV"))
    mar = (
        '<w:tblCellMar><w:left w:w="100" w:type="dxa"/><w:right w:w="100" w:type="dxa"/></w:tblCellMar>'
        if cell_mar
        else ""
    )
    indent = f'<w:tblInd w:w="{ind}" w:type="dxa"/>' if ind is not None else ""
    cells = (
        f'<w:tc><w:tcPr><w:tcW w:w="2160" w:type="dxa"/></w:tcPr>{para(f"N{n}a")}</w:tc>'
        f'<w:tc><w:tcPr><w:tcW w:w="2160" w:type="dxa"/></w:tcPr>{para(f"N{n}z", "<w:jc w:val=\"right\"/>")}</w:tc>'
    )
    return (
        f"<w:tbl><w:tblPr>{width}{indent}<w:tblBorders>{borders}</w:tblBorders>{mar}</w:tblPr>"
        '<w:tblGrid><w:gridCol w:w="2160"/><w:gridCol w:w="2160"/></w:tblGrid>'
        f"<w:tr>{cells}</w:tr></w:tbl>"
    )


def body() -> str:
    dxa = '<w:tblW w:w="4320" w:type="dxa"/>'
    pct = '<w:tblW w:w="5000" w:type="pct"/>'
    tables = (
        table(1, dxa, False, None),
        table(2, dxa, True, None),
        table(3, dxa, True, 720),
        table(4, pct, True, None),
    )
    return para("REF") + "".join(t + para("") for t in tables)


VARIANTS = {
    "none": (None, False),
    "full": (IND + MAR, False),
    "mar": (MAR, False),
    "ind": (IND, False),
    "none-14": (None, True),
}


def main() -> None:
    document = (
        f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        f"<w:document {NS}><w:body>{body()}{SECT}</w:body></w:document>"
    )
    for name, (table_normal, settings) in VARIANTS.items():
        parts = {
            "[Content_Types].xml": content_types(settings),
            "_rels/.rels": ROOT_RELS,
            "word/_rels/document.xml.rels": doc_rels(settings),
            "word/document.xml": document,
            "word/styles.xml": styles(table_normal),
        }
        if settings:
            parts["word/settings.xml"] = SETTINGS_14
        out = ROOT / "test-files" / f"table-normal-{name}.docx"
        with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as z:
            for part, content in parts.items():
                info = zipfile.ZipInfo(part, date_time=(1980, 1, 1, 0, 0, 0))
                info.compress_type = zipfile.ZIP_DEFLATED
                z.writestr(info, content.encode("utf-8"))
        print(f"wrote {out.relative_to(ROOT)}")


if __name__ == "__main__":
    main()
