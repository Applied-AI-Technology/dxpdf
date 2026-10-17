#!/usr/bin/env python3
"""Build ``test-files/paragraph-defaults-{none,15,explicit}.docx``, which ask
what paragraph spacing Word gives a document whose ``w:docDefaults`` carries
an ``rPrDefault`` but no ``pPrDefault``.

albanyny's CivicRecordCertificate has exactly that ``styles.xml`` and no
``settings.xml``, and Word sets every one of its lines 1.7–3.7pt further
apart than §17.3.1.33's defaults (no space after, single spacing) would.
§17.7.2 leaves an absent ``pPrDefault`` to the application, so only a Word
render can say which defaults it supplies and whether a paragraph's own
``w:spacing`` replaces them whole or attribute by attribute. The certificate
also writes ``w:spacing`` after ``w:jc``, out of ``CT_PPr``'s sequence, and
``w:caps`` after ``w:color``, out of ``CT_RPr``'s; Word applies the caps, so
the probe asks the same of the spacing.

The three files differ only in the package around one body:

  none      ``rPrDefault`` only, no ``settings.xml`` — the certificate's shape
  15        the same with a ``settings.xml`` naming ``compatibilityMode`` 15
  explicit  the same as ``none`` with a ``pPrDefault`` of no space after and
            single spacing — the control

Each body is Calibri 10pt, groups of three one-line paragraphs ``<group><n>``:

  A  no ``w:pPr``
  B  ``w:spacing w:after="20"`` only, in sequence
  C  ``w:jc`` then ``w:spacing w:after="20"`` — out of sequence
  D  ``w:spacing w:after="0" w:line="240" w:lineRule="auto"``
  E  ``w:spacing w:line="240" w:lineRule="auto"`` — line without after
  F  one paragraph wrapping onto three lines, no ``w:pPr``
  G  ``w:caps``, ``w:smallCaps`` at 10 and 20pt, each after a plain twin, so
     the small capitals' size and the caps' widths can be read off the page

Regenerate and commit the result if the content changes; the build is
deterministic and needs no third-party packages.

    scripts/make_paragraph_defaults_fixtures.py
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


SETTINGS_15 = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings {NS}><w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat></w:settings>"""

RPR_DEFAULT = (
    '<w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:cs="Calibri"/>'
    '<w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr></w:rPrDefault>'
)
PPR_DEFAULT = (
    '<w:pPrDefault><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:pPrDefault>'
)


def styles(ppr_default: bool) -> str:
    # The certificate's styles part, byte for byte apart from the optional pPrDefault.
    return (
        f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles {NS}>'
        f"<w:docDefaults>{RPR_DEFAULT}{PPR_DEFAULT if ppr_default else ''}</w:docDefaults>"
        '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>'
        "</w:styles>"
    )


SECT = (
    '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" '
    'w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>'
)


def para(text: str, ppr: str = "", rpr: str = "") -> str:
    ppr_xml = f"<w:pPr>{ppr}</w:pPr>" if ppr else ""
    rpr_xml = f"<w:rPr>{rpr}</w:rPr>" if rpr else ""
    return f'<w:p>{ppr_xml}<w:r>{rpr_xml}<w:t xml:space="preserve">{text}</w:t></w:r></w:p>'


GROUPS = [
    ("A", ""),
    ("B", '<w:spacing w:after="20"/>'),
    ("C", '<w:jc w:val="left"/><w:spacing w:after="20"/>'),
    ("D", '<w:spacing w:after="0" w:line="240" w:lineRule="auto"/>'),
    ("E", '<w:spacing w:line="240" w:lineRule="auto"/>'),
]

WRAP = " ".join(["F wraps"] + ["onto the next line"] * 20)


def body() -> str:
    out = []
    for name, ppr in GROUPS:
        out += [para(f"{name}{n}", ppr) for n in (1, 2, 3)]
    out.append(para(WRAP))
    out += [
        para("G1 Record Of Birth"),
        para("G2 Record Of Birth", rpr="<w:caps/>"),
        para("G3 Record Of Birth", rpr="<w:smallCaps/>"),
        para("G4 Record Of Birth", rpr='<w:sz w:val="40"/>'),
        para("G5 Record Of Birth", rpr='<w:smallCaps/><w:sz w:val="40"/>'),
    ]
    out.append(para("END"))
    return "".join(out)


def build(name: str, settings: bool, ppr_default: bool) -> None:
    document = (
        f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        f"<w:document {NS}><w:body>{body()}{SECT}</w:body></w:document>"
    )
    parts = {
        "[Content_Types].xml": content_types(settings),
        "_rels/.rels": ROOT_RELS,
        "word/_rels/document.xml.rels": doc_rels(settings),
        "word/document.xml": document,
        "word/styles.xml": styles(ppr_default),
    }
    if settings:
        parts["word/settings.xml"] = SETTINGS_15
    out = ROOT / "test-files" / f"paragraph-defaults-{name}.docx"
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as z:
        for part, text in parts.items():
            info = zipfile.ZipInfo(part, date_time=(1980, 1, 1, 0, 0, 0))
            info.compress_type = zipfile.ZIP_DEFLATED
            z.writestr(info, text.encode("utf-8"))
    print(f"wrote {out.relative_to(ROOT)}")


def main() -> None:
    build("none", settings=False, ppr_default=False)
    build("15", settings=True, ppr_default=False)
    build("explicit", settings=False, ppr_default=True)


if __name__ == "__main__":
    main()
