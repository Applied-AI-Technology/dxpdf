#!/usr/bin/env python3
"""Build ``test-files/kerning.docx``, which asks how Word applies §17.3.2.19
``w:kern``. Every paragraph is Calibri 11pt reading ``<label> To AVAWAY Tyre
Wolf yo. LT END`` — pairs Calibri kerns hard — and differs from ``K2`` only in
its run properties:

  K1  — ``w:kern`` 2 (1pt): kerned
  K2  — no ``w:kern`` (the control)
  K3  — ``w:kern`` 28 (14pt) at 11pt: below the threshold
  K4  — ``w:kern`` 28 at 16pt
  K4b — ``w:kern`` 28 at 14pt: on the threshold
  K4c — 14pt, no ``w:kern`` (K4b's control)
  K5  — ``T`` and ``owel To`` as two runs, both kerned: the pair across the
        run boundary
  K5b — the same with only ``T`` kerned
  K6  — ``w:kern`` 2 and ``w:spacing`` 20 (1pt)
  K6b — ``w:spacing`` 20 alone
  K12 — ``w:kern`` 0
  K7  — ``AY/AY-YA`` kerned: pairs across the line-break opportunities after
        ``/`` and ``-`` (``/A`` and ``-Y``, both kerned in Calibri)
  K8  — the same unkerned
  K9  — a 50pt one-cell table holding ``AVAVAVAVAVAVAVAVAVAV``, kerned: a word
        too wide for its cell, cut where it overflows
  K9b — the same unkerned

The document asks for ``compatibilityMode`` 15; Word kerns the same in modes 14
and absent. It carries its own ``styles.xml`` for the reason
``bidi-visual-table.docx`` does: an absent ``w:docDefaults`` is
application-defined (§17.7.2), and it is measured against a Word render.

Regenerate and commit the result if the content changes; the build is
deterministic and needs no third-party packages.

    scripts/make_kerning_fixture.py
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
SETTINGS = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
    f'<w:settings xmlns:w="{W_NS}"><w:compat><w:compatSetting w:name="compatibilityMode" '
    f'w:uri="{WORD_URI}" w:val="15"/></w:compat></w:settings>\n'
)

TEXT = "To AVAWAY Tyre Wolf yo. LT"


def run(text: str, rpr: str = "") -> str:
    props = f"<w:rPr>{rpr}</w:rPr>" if rpr else ""
    return f'<w:r>{props}<w:t xml:space="preserve">{text}</w:t></w:r>'


def para(*runs: str) -> str:
    return f"<w:p>{''.join(runs)}</w:p>"


def kern(v: int) -> str:
    return f'<w:kern w:val="{v}"/>'


SIZE = lambda half_points: f'<w:sz w:val="{half_points}"/>'
SPACING = '<w:spacing w:val="20"/>'


def body() -> str:
    line = lambda label, rpr: para(run(f"{label} {TEXT} END", rpr))
    return "".join(
        [
            line("K1", kern(2)),
            line("K2", ""),
            line("K3", kern(28)),
            line("K4", kern(28) + SIZE(32)),
            line("K4b", kern(28) + SIZE(28)),
            line("K4c", SIZE(28)),
            para(run("K5 T", kern(2)), run("owel To", kern(2)), run(" END", kern(2))),
            para(run("K5b T", kern(2)), run("owel To"), run(" END")),
            line("K6", kern(2) + SPACING),
            line("K6b", SPACING),
            line("K12", kern(0)),
            para(run("K7 AY/AY-YA END", kern(2))),
            para(run("K8 AY/AY-YA END")),
            para(run("K9", kern(2))),
            narrow_cell("AVAVAVAVAVAVAVAVAVAV", kern(2)),
            para(run("K9b")),
            narrow_cell("AVAVAVAVAVAVAVAVAVAV", ""),
            para(run("END")),
        ]
    )


def narrow_cell(text: str, rpr: str) -> str:
    """A one-cell table 1000 twips (50pt) wide, `TableNormal`'s margins."""
    return (
        '<w:tbl><w:tblPr><w:tblW w:w="1000" w:type="dxa"/><w:tblLayout w:type="fixed"/></w:tblPr>'
        '<w:tblGrid><w:gridCol w:w="1000"/></w:tblGrid><w:tr><w:tc>'
        '<w:tcPr><w:tcW w:w="1000" w:type="dxa"/></w:tcPr>'
        f"{para(run(text, rpr))}</w:tc></w:tr></w:tbl>"
    )


def document(content: str) -> str:
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        f'<w:document xmlns:w="{W_NS}"><w:body>{content}'
        '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/>'
        '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" '
        'w:header="720" w:footer="720" w:gutter="0"/></w:sectPr></w:body></w:document>\n'
    )


def write(name: str) -> None:
    target = OUT / name
    parts = [
        ("[Content_Types].xml", CONTENT_TYPES.replace("</Types>", SETTINGS_CT + "</Types>")),
        ("_rels/.rels", ROOT_RELS),
        ("word/_rels/document.xml.rels", DOC_RELS.replace("</Relationships>", SETTINGS_REL + "</Relationships>")),
        ("word/document.xml", document(body())),
        ("word/styles.xml", STYLES),
        ("word/settings.xml", SETTINGS),
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
    write("kerning.docx")


if __name__ == "__main__":
    main()
