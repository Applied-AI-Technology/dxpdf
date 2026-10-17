#!/usr/bin/env python3
"""Build ``test-files/line-grid.docx``, which asks how Word rounds the
heights it stacks a page from.

Every baseline Word writes sits on a multiple of 0.24pt — 1/300in — and its
lines are not the font's metrics rounded once at the end: a run of identical
10pt Calibri paragraphs at 1.15 steps 14.16pt (59 units) where the metrics
give 14.04, so the difference adds up down the page (albanyny's
CivicRecordCertificate is 2.3pt low by its last line). Only Word can say
which heights it rounds, and in which direction.

The page size and margins are a whole number of units (1in = 300), so a
page's first baseline is the margin plus the first line's rounded ascent.
Everything is Calibri unless named; paragraphs have no space unless named.

  pages 1–18  one font size each, s in points ``S<s>``: eight paragraphs at
              single spacing ``S<s>-<n>``, eight at 1.15 ``M<s>-<n>``, eight
              at 1.5 ``H<s>-<n>``, then a 10pt ``X<s>``
  page 19     10pt single, groups of three paragraphs whose ``w:after`` is t
              twips, ``A<t>-<n>``; then the same for ``w:before``, ``B<t>-<n>``
  page 20     10pt single between plain lines: ``w:pBdr`` bottom ``R<sz>s<sp>``
              (the certificate's is sz 8 space 1), then exact and atLeast
              lines ``E<twips>-<n>`` / ``L<twips>-<n>``
  page 21     mixed sizes on one line: ``Q<a>_<b>-<n>`` is an a-pt run then a
              b-pt run, three paragraphs each at single and 1.15 (``Q…m``)
  page 22     Times New Roman and Arial at 10 and 12pt, single and 1.15,
              ``T<s>-<n>``/``Tm<s>-<n>``, ``R…``/``Rm…`` likewise
  page 23     a table of 0.5pt borders, five rows of one 10pt single line
              (``C1``–``C5``), then five more with 1.5pt top cell margins
              (``D1``–``D5``)
  page 24     its own section with a 50pt top margin — 208.33 units — and
              a 10pt line ``G1``, then ``G2``

Regenerate and commit the result if the content changes; the build is
deterministic and needs no third-party packages.

    scripts/make_line_grid_fixture.py
"""

import pathlib
import zipfile

ROOT = pathlib.Path(__file__).resolve().parent.parent

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

SETTINGS = (
    f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings {NS}><w:compat>'
    '<w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/>'
    "</w:compat></w:settings>"
)

STYLES = (
    f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles {NS}>'
    "<w:docDefaults>"
    '<w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:cs="Calibri"/>'
    '<w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr></w:rPrDefault>'
    '<w:pPrDefault><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:pPrDefault>'
    "</w:docDefaults>"
    '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>'
    "</w:styles>"
)

MARGINS = 'w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"'

# Half-points.
SIZES = (8, 14, 16, 17, 18, 19, 20, 21, 22, 24, 26, 28, 32, 36, 40, 48, 56, 72)
AFTER = (1, 5, 10, 12, 13, 14, 15, 20, 25, 30, 40, 60, 100)


def name(half_points: int) -> str:
    return f"{half_points // 2}" + (".5" if half_points % 2 else "")


def run(text: str, half_points: int = 20, font: str = "") -> str:
    fonts = f'<w:rFonts w:ascii="{font}" w:hAnsi="{font}" w:cs="{font}"/>' if font else ""
    return (
        f'<w:r><w:rPr>{fonts}<w:sz w:val="{half_points}"/><w:szCs w:val="{half_points}"/></w:rPr>'
        f'<w:t xml:space="preserve">{text}</w:t></w:r>'
    )


def para(runs: str, ppr: str = "", page_break: bool = False) -> str:
    ppr = ("<w:pageBreakBefore/>" if page_break else "") + ppr
    ppr_xml = f"<w:pPr>{ppr}</w:pPr>" if ppr else ""
    return f"<w:p>{ppr_xml}{runs}</w:p>"


def line(rule: int) -> str:
    return f'<w:spacing w:line="{rule}" w:lineRule="auto"/>'


def size_page(half_points: int) -> list:
    s = name(half_points)
    out = []
    for prefix, rule in (("S", 240), ("M", 276), ("H", 360)):
        for n in range(1, 9):
            first = prefix == "S" and n == 1
            out.append(para(run(f"{prefix}{s}-{n}", half_points), line(rule), page_break=first))
    out.append(para(run(f"X{s}")))
    return out


def spacing_page() -> list:
    out = [para(run("A0"), page_break=True)]
    for t in AFTER:
        out += [para(run(f"A{t}-{n}"), f'<w:spacing w:after="{t}"/>') for n in range(1, 4)]
    out.append(para(run("B0")))
    for t in AFTER:
        out += [para(run(f"B{t}-{n}"), f'<w:spacing w:before="{t}"/>') for n in range(1, 4)]
    out.append(para(run("BZ")))
    return out


def border_page() -> list:
    out = [para(run("R0"), page_break=True)]
    for sz, sp in ((8, 1), (4, 0), (6, 3), (12, 2)):
        bdr = f'<w:pBdr><w:bottom w:val="single" w:sz="{sz}" w:space="{sp}" w:color="000000"/></w:pBdr>'
        out += [para(run(f"R{sz}s{sp}"), bdr), para(run(f"R{sz}s{sp}z"))]
    for twips in (240, 250, 241, 263):
        out += [
            para(run(f"E{twips}-{n}"), f'<w:spacing w:line="{twips}" w:lineRule="exact"/>')
            for n in range(1, 4)
        ]
    for twips in (250, 263, 300):
        out += [
            para(run(f"L{twips}-{n}"), f'<w:spacing w:line="{twips}" w:lineRule="atLeast"/>')
            for n in range(1, 4)
        ]
    out.append(para(run("LZ")))
    return out


def mixed_page() -> list:
    out = [para(run("Q0"), page_break=True)]
    for a, b in ((20, 24), (16, 28), (17, 22)):
        for suffix, rule in (("", 240), ("m", 276)):
            label = f"Q{name(a)}_{name(b)}{suffix}"
            out += [
                para(run(f"{label}-{n} ", a) + run("x", b), line(rule)) for n in range(1, 4)
            ]
    out.append(para(run("QZ")))
    return out


def fonts_page() -> list:
    out = [para(run("F0"), page_break=True)]
    for font, letter in (("Times New Roman", "T"), ("Arial", "R")):
        for half_points in (20, 24):
            for suffix, rule in (("", 240), ("m", 276)):
                out += [
                    para(run(f"{letter}{suffix}{name(half_points)}-{n}", half_points, font), line(rule))
                    for n in range(1, 5)
                ]
    out.append(para(run("FZ")))
    return out


def table(prefix: str, cell_top: int) -> str:
    edge = '<w:{0} w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
    borders = "".join(edge.format(e) for e in ("top", "left", "bottom", "right", "insideH", "insideV"))
    mar = f'<w:tblCellMar><w:top w:w="{cell_top}" w:type="dxa"/></w:tblCellMar>'
    rows = "".join(
        f'<w:tr><w:tc><w:tcPr><w:tcW w:w="4320" w:type="dxa"/></w:tcPr>{para(run(f"{prefix}{n}"))}</w:tc></w:tr>'
        for n in range(1, 6)
    )
    return (
        f'<w:tbl><w:tblPr><w:tblW w:w="4320" w:type="dxa"/><w:tblBorders>{borders}</w:tblBorders>{mar}'
        f'</w:tblPr><w:tblGrid><w:gridCol w:w="4320"/></w:tblGrid>{rows}</w:tbl>'
    )


def table_page() -> list:
    return [
        para(run("C0"), page_break=True),
        table("C", 0),
        para(run("D0")),
        table("D", 30),
        para(run("DZ")),
    ]


def body() -> str:
    out = []
    for half_points in SIZES:
        out += size_page(half_points)
    out += spacing_page() + border_page() + mixed_page() + fonts_page() + table_page()
    first = (
        '<w:p><w:pPr><w:sectPr><w:pgSz w:w="12240" w:h="15840"/>'
        f'<w:pgMar w:top="1440" {MARGINS}/></w:sectPr></w:pPr></w:p>'
    )
    out.append(first)
    out += [para(run("G1")), para(run("G2"))]
    return "".join(out)


SECT = f'<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1000" {MARGINS}/></w:sectPr>'


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
    }
    out = ROOT / "test-files" / "line-grid.docx"
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as z:
        for part, content in parts.items():
            info = zipfile.ZipInfo(part, date_time=(1980, 1, 1, 0, 0, 0))
            info.compress_type = zipfile.ZIP_DEFLATED
            z.writestr(info, content.encode("utf-8"))
    print(f"wrote {out.relative_to(ROOT)}")


if __name__ == "__main__":
    main()
