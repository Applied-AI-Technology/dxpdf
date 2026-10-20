#!/usr/bin/env python3
"""Build ``test-files/hanging-tab-{15,compat,legacy,legacy-15}.docx``, which
ask where a tab on the first line of a hanging-indent paragraph stops.

Two real-world notice templates open a paragraph ``w:ind left="333"
hanging="180"`` with a tab, and Word sets the text after it at the left
indent, 16.65pt, not at the 36pt default stop: a hanging indent is a tab
stop. §17.15.3's ``w:noTabHangInd`` names the behaviour by turning it off.
The probe asks how the stop orders against the paragraph's own stops, and
whether the legacy flag holds in each mode.

The four files differ only in ``settings.xml``:

  15         ``compatibilityMode`` 15 — the notices' shape
  compat     no ``settings.xml``
  legacy     ``w:noTabHangInd`` and no mode
  legacy-15  ``w:noTabHangInd`` with mode 15

Each body is Calibri 10pt, one paragraph per case, each ``<tab>`` followed by
its label:

  H1  ``left=333 hanging=180``, a tab then ``H1`` — the notices
  H2  ``left=1440 hanging=720``, a stop of its own at 1080 — between the
      first line's start and the left indent
  H3  the same with its own stop at 2160 instead — past the left indent
  H4  ``left=1440 hanging=720``, text already past the left indent first
  H5  ``left=1440 firstLine=360`` — no hanging indent, the control
  H6  ``left=1440 hanging=720``, two tabs

Regenerate and commit the result if the content changes; the build is
deterministic and needs no third-party packages.

    scripts/make_hanging_tab_fixtures.py
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


MODE_15 = '<w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/>'


def settings(no_tab_hang_ind: bool, mode_15: bool) -> str:
    # CT_Compat: the legacy flags precede the compatSetting entries.
    flag = "<w:noTabHangInd/>" if no_tab_hang_ind else ""
    mode = MODE_15 if mode_15 else ""
    return (
        f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<w:settings {NS}><w:defaultTabStop w:val="720"/><w:compat>{flag}{mode}</w:compat></w:settings>'
    )


STYLES = (
    f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles {NS}>'
    "<w:docDefaults>"
    '<w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:cs="Calibri"/>'
    '<w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr></w:rPrDefault>'
    '<w:pPrDefault><w:pPr><w:spacing w:after="120" w:line="240" w:lineRule="auto"/></w:pPr></w:pPrDefault>'
    "</w:docDefaults>"
    '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>'
    "</w:styles>"
)

SECT = (
    '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" '
    'w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>'
)

TAB = "<w:r><w:tab/></w:r>"


def text(t: str) -> str:
    return f'<w:r><w:t xml:space="preserve">{t}</w:t></w:r>'


def para(ppr: str, *runs: str) -> str:
    return f"<w:p><w:pPr>{ppr}</w:pPr>{''.join(runs)}</w:p>"


def tabs(*positions: int) -> str:
    return "<w:tabs>" + "".join(f'<w:tab w:val="left" w:pos="{p}"/>' for p in positions) + "</w:tabs>"


HANG = '<w:ind w:left="1440" w:hanging="720"/>'


def body() -> str:
    return "".join(
        [
            para('<w:ind w:left="333" w:hanging="180"/>', TAB, text("H1")),
            para(tabs(1080) + HANG, TAB, text("H2")),
            para(tabs(2160) + HANG, TAB, text("H3")),
            para(HANG, text("H4 is wider than its hanging indent"), TAB, text("H4z")),
            para('<w:ind w:left="1440" w:firstLine="360"/>', TAB, text("H5")),
            para(HANG, TAB, text("H6"), TAB, text("H6z")),
            para("", text("END")),
        ]
    )


def build(name: str, settings_xml: str | None) -> None:
    has_settings = settings_xml is not None
    document = (
        f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        f"<w:document {NS}><w:body>{body()}{SECT}</w:body></w:document>"
    )
    parts = {
        "[Content_Types].xml": content_types(has_settings),
        "_rels/.rels": ROOT_RELS,
        "word/_rels/document.xml.rels": doc_rels(has_settings),
        "word/document.xml": document,
        "word/styles.xml": STYLES,
    }
    if has_settings:
        parts["word/settings.xml"] = settings_xml
    out = ROOT / "test-files" / f"hanging-tab-{name}.docx"
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as z:
        for part, content in parts.items():
            info = zipfile.ZipInfo(part, date_time=(1980, 1, 1, 0, 0, 0))
            info.compress_type = zipfile.ZIP_DEFLATED
            z.writestr(info, content.encode("utf-8"))
    print(f"wrote {out.relative_to(ROOT)}")


def main() -> None:
    build("15", settings(no_tab_hang_ind=False, mode_15=True))
    build("compat", None)
    build("legacy", settings(no_tab_hang_ind=True, mode_15=False))
    build("legacy-15", settings(no_tab_hang_ind=True, mode_15=True))


if __name__ == "__main__":
    main()
