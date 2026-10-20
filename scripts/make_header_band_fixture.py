#!/usr/bin/env python3
"""Build ``test-files/header-band.docx``, which asks whether a table's first
row takes part in its row banding (§17.7.6 ``band1Horz``/``band2Horz``) when
``tblLook`` turns ``firstRow`` on.

``row-border-room.docx`` B8 found that it does in Word under a style with no
``firstRow`` layer: the header row takes ``band1Horz``'s top border and the
body's bands start one row later. Two rules fit that: Word leaves the first
row out of the bands only where the style defines a ``firstRow`` layer, or
it always bands it and a ``firstRow`` layer merely overrides the band's
formatting. These cases tell them apart.

Each case is a four-row table between a ``SEP<case>`` and an ``END<case>``
paragraph, rows ``<case>0``-``<case>3`` in Calibri 11pt with no cell
margins, under a style whose ``band1Horz`` gives a row 6pt black top and
bottom borders (band size 1); H5 starts the second page. A row's band shows
twice: ``SEP`` to row 0 grows 6pt when row 0 is ``band1``, and row 3 to
``END`` grows 6pt when row 3 is.

  H1  ``row-border-room.docx`` B8 again: ``tblHeader`` row 0 with a direct
      bottom border, ``firstRow`` on, no ``firstRow`` layer
  H2  plain row 0, ``firstRow`` on, no ``firstRow`` layer
  H3  ``tblHeader`` row 0, ``firstRow`` on, no ``firstRow`` layer
  H4  plain row 0, ``firstRow`` on, a ``firstRow`` layer that only shades
  H5  H4 with ``tblHeader``
  H6  H4 with ``firstRow`` off — the control: every row banded from row 0
  H7  plain row 0, ``firstRow`` on, a ``firstRow`` layer that shades and sets
      its top and bottom borders ``nil``

If H4 bands row 0 and shifts the body as H2 does, a layer changes nothing;
if it starts the bands at row 1, the layer is what excludes the row. H7
separates a layer that wins row 0's borders from one that removes the row
from banding altogether.

Regenerate and commit the result if the content changes; the build is
deterministic and needs no third-party packages.

    scripts/make_header_band_fixture.py
"""

import pathlib
import zipfile

ROOT = pathlib.Path(__file__).resolve().parent.parent
OUT = ROOT / "test-files" / "header-band.docx"

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


def border(side: str, val: str = "single") -> str:
    if val == "nil":
        return f'<w:{side} w:val="nil"/>'
    return f'<w:{side} w:val="single" w:sz="48" w:space="0" w:color="000000"/>'


STYLES = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles {NS}>
<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:cs="Calibri"/><w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr></w:rPrDefault>
<w:pPrDefault><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>
<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
<w:style w:type="table" w:default="1" w:styleId="TableNormal"><w:name w:val="Normal Table"/><w:tblPr><w:tblInd w:w="0" w:type="dxa"/>
<w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:left w:w="108" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar></w:tblPr></w:style>
<w:style w:type="table" w:styleId="Banded"><w:name w:val="Banded"/><w:basedOn w:val="TableNormal"/>
<w:tblPr><w:tblStyleRowBandSize w:val="1"/></w:tblPr>
<w:tblStylePr w:type="band1Horz"><w:tblPr/><w:tcPr><w:tcBorders>{border("top")}{border("bottom")}</w:tcBorders></w:tcPr></w:tblStylePr>
</w:style>
<w:style w:type="table" w:styleId="BandedHead"><w:name w:val="Banded Head"/><w:basedOn w:val="TableNormal"/>
<w:tblPr><w:tblStyleRowBandSize w:val="1"/></w:tblPr>
<w:tblStylePr w:type="firstRow"><w:tblPr/><w:tcPr><w:shd w:val="clear" w:color="auto" w:fill="D9D9D9"/></w:tcPr></w:tblStylePr>
<w:tblStylePr w:type="band1Horz"><w:tblPr/><w:tcPr><w:tcBorders>{border("top")}{border("bottom")}</w:tcBorders></w:tcPr></w:tblStylePr>
</w:style>
<w:style w:type="table" w:styleId="BandedHeadNil"><w:name w:val="Banded Head Nil"/><w:basedOn w:val="TableNormal"/>
<w:tblPr><w:tblStyleRowBandSize w:val="1"/></w:tblPr>
<w:tblStylePr w:type="firstRow"><w:tblPr/><w:tcPr><w:tcBorders>{border("top", "nil")}{border("bottom", "nil")}</w:tcBorders><w:shd w:val="clear" w:color="auto" w:fill="D9D9D9"/></w:tcPr></w:tblStylePr>
<w:tblStylePr w:type="band1Horz"><w:tblPr/><w:tcPr><w:tcBorders>{border("top")}{border("bottom")}</w:tcBorders></w:tcPr></w:tblStylePr>
</w:style>
</w:styles>"""

SECT = (
    '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" '
    'w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>'
)


def run(text: str) -> str:
    return f'<w:r><w:t xml:space="preserve">{text}</w:t></w:r>'


def row(text: str, borders: str = "", header: bool = False) -> str:
    tc_borders = f"<w:tcBorders>{borders}</w:tcBorders>" if borders else ""
    tr_pr = "<w:trPr><w:tblHeader/></w:trPr>" if header else ""
    return (
        f'<w:tr>{tr_pr}<w:tc><w:tcPr><w:tcW w:w="2880" w:type="dxa"/>{tc_borders}</w:tcPr>'
        f"<w:p>{run(text)}</w:p></w:tc></w:tr>"
    )


def table(rows: str, tbl_borders: str = "", style: str = "", look: str = "") -> str:
    style_el = f'<w:tblStyle w:val="{style}"/>' if style else ""
    tb = f"<w:tblBorders>{tbl_borders}</w:tblBorders>" if tbl_borders else ""
    bits = int(look or "0000", 16)
    flag = lambda mask: "1" if bits & mask else "0"  # noqa: E731
    look_el = (
        f'<w:tblLook w:val="{look or "0000"}" w:firstRow="{flag(0x20)}" w:lastRow="{flag(0x40)}" '
        f'w:firstColumn="{flag(0x80)}" w:lastColumn="{flag(0x100)}" w:noHBand="{flag(0x200)}" '
        f'w:noVBand="{flag(0x400)}"/>'
    )
    return (
        f'<w:tbl><w:tblPr>{style_el}<w:tblW w:w="2880" w:type="dxa"/><w:tblLayout w:type="fixed"/>{tb}{look_el}</w:tblPr>'
        f'<w:tblGrid><w:gridCol w:w="2880"/></w:tblGrid>{rows}</w:tbl>'
    )


def case(name: str) -> str:
    def rows(head: bool, direct: bool = False) -> str:
        r0 = row(f"{name}0", border("bottom") if direct else "", header=head)
        return r0 + "".join(row(f"{name}{i}") for i in (1, 2, 3))

    if name == "H1":
        t = table(rows(True, direct=True), style="Banded", look="0420")
    elif name == "H2":
        t = table(rows(False), style="Banded", look="0420")
    elif name == "H3":
        t = table(rows(True), style="Banded", look="0420")
    elif name == "H4":
        t = table(rows(False), style="BandedHead", look="0420")
    elif name == "H5":
        t = table(rows(True), style="BandedHead", look="0420")
    elif name == "H6":
        t = table(rows(False), style="BandedHead", look="0400")
    else:  # H7
        t = table(rows(False), style="BandedHeadNil", look="0420")
    # H5 starts a page, so no table can be pushed across a page break.
    ppr = "<w:pPr><w:pageBreakBefore/></w:pPr>" if name == "H5" else ""
    return f"<w:p>{ppr}{run('SEP' + name)}</w:p>{t}<w:p>{run('END' + name)}</w:p>"


def main() -> None:
    names = ["H1", "H2", "H3", "H4", "H5", "H6", "H7"]
    body = "".join(case(n) for n in names)
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
