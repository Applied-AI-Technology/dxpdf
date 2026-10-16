#!/usr/bin/env python3
"""Build ``test-files/ligatures-{15,14,14otf}.docx``, which ask when Word forms
the OpenType ligatures [MS-DOCX] ``w14:ligatures`` names. Every paragraph is
Calibri 11pt reading ``<label> office fifty fluff staff Wafflt END`` — words
Calibri ligates (``ffi``, ``fi``, ``fl``, ``ff``, ``ft``) — and differs only in
its run properties:

  L1 — ``standardContextual``, the value Word's own template writes
  L2 — none (the control)
  L3 — ``standard``
  L4 — ``contextual``
  L5 — ``standardContextual`` with ``w:spacing`` 20 (1pt)
  L6 — ``standardContextual`` with ``w:kern`` 2, and a closing ``To``
  L7 — ``of`` and ``fice END`` as two runs, both ``standardContextual``

The three documents differ only in §17.15.1.21: ``compatibilityMode`` 15; 14;
and 14 with ``enableOpenTypeFeatures`` (which carries only L1 and L2). Each
carries its own ``styles.xml`` for the reason ``bidi-visual-table.docx`` does:
an absent ``w:docDefaults`` is application-defined (§17.7.2), and they are
measured against a Word render.

Regenerate and commit the results if the content changes; the build is
deterministic and needs no third-party packages.

    scripts/make_ligatures_fixtures.py
"""

import pathlib
import zipfile

ROOT = pathlib.Path(__file__).resolve().parent.parent
OUT = ROOT / "test-files"

CONTENT_TYPES = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
</Types>
"""

ROOT_RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>
"""

DOC_RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdS" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>
"""

STYLES = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:docDefaults>
    <w:rPrDefault><w:rPr>
      <w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:cs="Calibri"/>
      <w:sz w:val="22"/><w:szCs w:val="22"/>
    </w:rPr></w:rPrDefault>
    <w:pPrDefault><w:pPr>
      <w:spacing w:after="0" w:line="240" w:lineRule="auto"/>
    </w:pPr></w:pPrDefault>
  </w:docDefaults>
  <w:style w:type="paragraph" w:default="1" w:styleId="Normal">
    <w:name w:val="Normal"/>
  </w:style>
  <w:style w:type="table" w:default="1" w:styleId="TableNormal">
    <w:name w:val="Normal Table"/>
    <w:tblPr>
      <w:tblInd w:w="0" w:type="dxa"/>
      <w:tblCellMar>
        <w:top w:w="0" w:type="dxa"/><w:left w:w="108" w:type="dxa"/>
        <w:bottom w:w="0" w:type="dxa"/><w:right w:w="108" w:type="dxa"/>
      </w:tblCellMar>
    </w:tblPr>
  </w:style>
</w:styles>
"""

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W14_NS = "http://schemas.microsoft.com/office/word/2010/wordml"
MC_NS = "http://schemas.openxmlformats.org/markup-compatibility/2006"
WORD_URI = "http://schemas.microsoft.com/office/word"

SETTINGS_CT = (
    '  <Override PartName="/word/settings.xml" '
    'ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>\n'
)
SETTINGS_REL = (
    '  <Relationship Id="rIdT" '
    'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" '
    'Target="settings.xml"/>\n'
)


def compat_setting(name: str, value: str) -> str:
    return f'<w:compatSetting w:name="{name}" w:uri="{WORD_URI}" w:val="{value}"/>'


def settings(*entries: str) -> str:
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        f'<w:settings xmlns:w="{W_NS}"><w:compat>{"".join(entries)}</w:compat></w:settings>\n'
    )


TEXT = "office fifty fluff staff Wafflt"


def run(text: str, rpr: str = "") -> str:
    props = f"<w:rPr>{rpr}</w:rPr>" if rpr else ""
    return f'<w:r>{props}<w:t xml:space="preserve">{text}</w:t></w:r>'


def para(*runs: str) -> str:
    return f"<w:p>{''.join(runs)}</w:p>"


def lig(value: str) -> str:
    return f'<w14:ligatures w14:val="{value}"/>'


SPACING = '<w:spacing w:val="20"/>'
KERN = '<w:kern w:val="2"/>'


def lines(full: bool) -> str:
    line = lambda label, rpr, tail="": para(run(f"{label} {TEXT}{tail} END", rpr))
    out = [line("L1", lig("standardContextual")), line("L2", "")]
    if full:
        out += [
            line("L3", lig("standard")),
            line("L4", lig("contextual")),
            line("L5", lig("standardContextual") + SPACING),
            line("L6", lig("standardContextual") + KERN, " To"),
            para(run("L7 of", lig("standardContextual")), run("fice END", lig("standardContextual"))),
        ]
    return "".join(out)


def document(content: str) -> str:
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        f'<w:document xmlns:w="{W_NS}" xmlns:w14="{W14_NS}" xmlns:mc="{MC_NS}" '
        f'mc:Ignorable="w14"><w:body>{content}'
        '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/>'
        '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" '
        'w:header="720" w:footer="720" w:gutter="0"/></w:sectPr></w:body></w:document>\n'
    )


def write(name: str, settings_xml: str, full: bool) -> None:
    target = OUT / name
    parts = [
        ("[Content_Types].xml", CONTENT_TYPES.replace("</Types>", SETTINGS_CT + "</Types>")),
        ("_rels/.rels", ROOT_RELS),
        ("word/_rels/document.xml.rels", DOC_RELS.replace("</Relationships>", SETTINGS_REL + "</Relationships>")),
        ("word/document.xml", document(lines(full))),
        ("word/styles.xml", STYLES),
        ("word/settings.xml", settings_xml),
    ]
    # Fixed timestamps so regenerating an unchanged fixture produces identical
    # bytes and does not show up as a diff.
    with zipfile.ZipFile(target, "w", zipfile.ZIP_DEFLATED) as z:
        for part, data in parts:
            info = zipfile.ZipInfo(part, date_time=(1980, 1, 1, 0, 0, 0))
            info.compress_type = zipfile.ZIP_DEFLATED
            z.writestr(info, data)
    print(f"wrote {target.relative_to(ROOT)} ({target.stat().st_size} bytes)")


def main() -> None:
    OUT.mkdir(parents=True, exist_ok=True)
    write("ligatures-15.docx", settings(compat_setting("compatibilityMode", "15")), True)
    write("ligatures-14.docx", settings(compat_setting("compatibilityMode", "14")), True)
    write(
        "ligatures-14otf.docx",
        settings(
            compat_setting("compatibilityMode", "14"),
            compat_setting("enableOpenTypeFeatures", "1"),
        ),
        False,
    )


if __name__ == "__main__":
    main()
