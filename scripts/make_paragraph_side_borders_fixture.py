#!/usr/bin/env python3
"""Build ``test-files/paragraph-side-borders.docx``, which asks how Word draws
the left and right borders of paragraphs joined into one bordered box
(§17.3.1.24).

Consecutive paragraphs with the same ``w:pBdr`` form one box: the top border
over the first, the bottom under the last. dxpdf draws each paragraph's side
borders over its own lines only, so a box whose paragraphs have space
between them is broken at every gap, and it puts the side borders on the
indents where §17.3.1.24's ``w:space`` names a distance between text and
border. The render says where each side border stands, how far it reaches,
and what joins a group.

Every line is Calibri 10pt on an exact 12pt line; borders are 1pt
(``w:sz="8"``), sides ``w:space="4"`` and top/bottom ``w:space="1"`` unless
named. Each case is between plain ``<case>a`` and ``<case>z`` paragraphs:

  S1  one paragraph boxed on all four sides
  S2  three paragraphs ``S2-<n>``, left and right only, 6pt after each
  S3  three boxed paragraphs, 6pt after and 3pt before each
  S4  S3 with a ``w:between`` border
  S5  three paragraphs, a left border and a fill, 6pt after each
  S6  two left-bordered paragraphs, the second indented 0.5in
  S7  three boxed paragraphs, the middle one with 3pt side ``w:space``
  S8  S2 with a fill — whether the last paragraph's space after is boxed
  S9  S2 with a left border only
  S10 three boxed paragraphs, 9pt before and 3pt after each
  S11 three paragraphs, left and right, 6pt before each
  S12 one paragraph, left and right, 6pt after
  S13 S8 with 6pt before instead of after — the first paragraph's space
  S14 three paragraphs, top and left, 6pt after each
  H<sz>s<sp>  one boxed paragraph, all four sides ``w:sz`` sz and side
              ``w:space`` sp — where a side border stands off the text
  HT<pt>  one boxed paragraph of ``<pt>``-point text (6 and 20)
  HI  one boxed paragraph indented 0.5in left and right
  HF  one boxed paragraph with a 0.5in first-line indent, wrapping

Regenerate and commit the result if the content changes; the build is
deterministic and needs no third-party packages.

    scripts/make_paragraph_side_borders_fixture.py
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
    '<w:pPrDefault><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="exact"/></w:pPr></w:pPrDefault>'
    "</w:docDefaults>"
    '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>'
    "</w:styles>"
)

SECT = (
    '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" '
    'w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>'
)


def para(text: str, ppr: str = "") -> str:
    ppr_xml = f"<w:pPr>{ppr}</w:pPr>" if ppr else ""
    return f'<w:p>{ppr_xml}<w:r><w:t xml:space="preserve">{text}</w:t></w:r></w:p>'


def edge(side: str, space: int, sz: int) -> str:
    return f'<w:{side} w:val="single" w:sz="{sz}" w:space="{space}" w:color="000000"/>'


def pbdr(sides: str, side_space: int = 4, sz: int = 8) -> str:
    out = ""
    for side in ("top", "left", "bottom", "right", "between"):
        if side in sides.split():
            out += edge(side, 1 if side in ("top", "bottom", "between") else side_space, sz)
    return f"<w:pBdr>{out}</w:pBdr>"


BOX = "top left bottom right"
AFTER = '<w:spacing w:after="120"/>'
BOTH = '<w:spacing w:before="60" w:after="120"/>'
FILL = '<w:shd w:val="clear" w:color="auto" w:fill="D9D9D9"/>'


def case(name: str, paragraphs: list) -> list:
    return [para(f"{name}a"), *paragraphs, para(f"{name}z")]


def body() -> str:
    out = []
    out += case("S1", [para("S1", pbdr(BOX))])
    out += case("S2", [para(f"S2-{n}", pbdr("left right") + AFTER) for n in (1, 2, 3)])
    out += case("S3", [para(f"S3-{n}", pbdr(BOX) + BOTH) for n in (1, 2, 3)])
    out += case("S4", [para(f"S4-{n}", pbdr(BOX + " between") + BOTH) for n in (1, 2, 3)])
    out += case("S5", [para(f"S5-{n}", pbdr("left") + FILL + AFTER) for n in (1, 2, 3)])
    out += case(
        "S6",
        [para("S6-1", pbdr("left")), para("S6-2", pbdr("left") + '<w:ind w:left="720"/>')],
    )
    out += case(
        "S7",
        [para("S7-1", pbdr(BOX)), para("S7-2", pbdr(BOX, 3)), para("S7-3", pbdr(BOX))],
    )
    out += case("S8", [para(f"S8-{n}", pbdr("left right") + FILL + AFTER) for n in (1, 2, 3)])
    out += case("S9", [para(f"S9-{n}", pbdr("left") + AFTER) for n in (1, 2, 3)])
    before_more = '<w:spacing w:before="180" w:after="60"/>'
    out += case("S10", [para(f"S10-{n}", pbdr(BOX) + before_more) for n in (1, 2, 3)])
    before_only = '<w:spacing w:before="120"/>'
    out += case("S11", [para(f"S11-{n}", pbdr("left right") + before_only) for n in (1, 2, 3)])
    out += case("S12", [para("S12", pbdr("left right") + AFTER)])
    for sz in (4, 24):
        for sp in (0, 10):
            out += case(f"H{sz}s{sp}", [para(f"H{sz}s{sp}", pbdr(BOX, sp, sz))])
    out += case("S13", [para(f"S13-{n}", pbdr("left right") + FILL + before_only) for n in (1, 2, 3)])
    out += case("S14", [para(f"S14-{n}", pbdr("top left") + AFTER) for n in (1, 2, 3)])
    for pt in (6, 20):
        run = f'<w:r><w:rPr><w:sz w:val="{pt * 2}"/></w:rPr><w:t>HT{pt}</w:t></w:r>'
        out += case(f"HT{pt}", [f"<w:p><w:pPr>{pbdr(BOX)}</w:pPr>{run}</w:p>"])
    out += case("HI", [para("HI", pbdr(BOX) + '<w:ind w:left="720" w:right="720"/>')])
    wrap = "HF " + " ".join("wraps" for _ in range(30))
    out += case("HF", [para(wrap, pbdr(BOX) + '<w:ind w:firstLine="720"/>')])
    out.append(para("END"))
    return "".join(out)


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
    out = ROOT / "test-files" / "paragraph-side-borders.docx"
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as z:
        for part, content in parts.items():
            info = zipfile.ZipInfo(part, date_time=(1980, 1, 1, 0, 0, 0))
            info.compress_type = zipfile.ZIP_DEFLATED
            z.writestr(info, content.encode("utf-8"))
    print(f"wrote {out.relative_to(ROOT)}")


if __name__ == "__main__":
    main()
