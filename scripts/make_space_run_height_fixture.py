#!/usr/bin/env python3
"""Build ``test-files/space-run-height.docx``, which asks whether a run that
holds only white space stands its line as tall as its font makes it.

A real-world notice template sets 10pt Times New Roman text in
a Calibri 11pt document at 1.15, with one unformatted ``" "`` run between two
words. Word sets that line as the Times New Roman lines around it; a line
sized from the space's Calibri puts it, and everything below, about 1pt
lower.

Paragraphs are one line each unless named, text in explicit 10pt Times New
Roman, the document default Calibri 11pt, Normal ``w:line`` 276 (auto) and no
paragraph spacing. ``S<n>a`` and ``S<n>b`` are the text either side of the
run under test:

  S0   no run (the control)
  S1   ``" "`` with no properties — the notice's case
  S2   ``" "`` in 20pt Times New Roman
  S3   ``"x"`` with no properties (not white space, the control for S1)
  S4   ``"   "`` with no properties
  S5   a no-break space (U+00A0) with no properties
  S6   ``" "`` with no properties after ``S6a`` at the paragraph's end
  S7   a paragraph holding only a 20pt Times New Roman ``" "``, then ``S7z``
  S8   ``" "`` in 20pt Times New Roman at the start of the line
  S9   a tab with no properties
  S10  S1 in a paragraph of several lines, the run on its last, then
       ``S10z``
  S11  text with an unformatted paragraph mark, so the mark is Calibri 11pt
       (every other paragraph's mark is the text's 10pt Times New Roman)

``T1``–``T3`` repeat S1–S3 at single spacing (``w:line`` 240).

Regenerate and commit the result if the content changes; the build is
deterministic and needs no third-party packages.

    scripts/make_space_run_height_fixture.py
"""

import pathlib
import zipfile

ROOT = pathlib.Path(__file__).resolve().parent.parent
OUT = ROOT / "test-files" / "space-run-height.docx"

W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'

CONTENT_TYPES = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
</Types>"""

ROOT_RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>"""

DOC_RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rIdS" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>"""

STYLES = (
    f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles {W}>'
    "<w:docDefaults>"
    '<w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:cs="Calibri"/>'
    '<w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr></w:rPrDefault>'
    '<w:pPrDefault><w:pPr><w:spacing w:after="0" w:line="276" w:lineRule="auto"/></w:pPr></w:pPrDefault>'
    "</w:docDefaults>"
    '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>'
    "</w:styles>"
)

SECT = (
    '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" '
    'w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>'
)

TNR = '<w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/>'


def tnr(text: str, half_points: int = 20) -> str:
    return (
        f'<w:r><w:rPr>{TNR}<w:sz w:val="{half_points}"/><w:szCs w:val="{half_points}"/></w:rPr>'
        f'<w:t xml:space="preserve">{text}</w:t></w:r>'
    )


def plain(text: str) -> str:
    return f'<w:r><w:t xml:space="preserve">{text}</w:t></w:r>'


TAB = "<w:r><w:tab/></w:r>"
# The paragraph mark in the text's own size, so only the run under test differs.
MARK = f'<w:rPr>{TNR}<w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr>'


def para(content: str, ppr: str = "") -> str:
    return f"<w:p><w:pPr>{ppr}{MARK}</w:pPr>{content}</w:p>"


SINGLE = '<w:spacing w:line="240" w:lineRule="auto"/>'

FILLER = " ".join(["filler"] * 15)


def body() -> str:
    out = [para(tnr("REF"))]

    def case(label: str, middle: str, ppr: str = "") -> None:
        out.append(para(tnr(f"{label}a") + middle + tnr(f"{label}b"), ppr))

    case("S0", tnr(" "))
    case("S1", plain(" "))
    case("S2", tnr(" ", 40))
    case("S3", plain("x"))
    case("S4", plain("   "))
    case("S5", plain(" "))
    out.append(para(tnr("S6a") + plain(" ")))
    out.append(para(tnr(" ", 40)))
    out.append(para(tnr("S7z")))
    out.append(para(tnr(" ", 40) + tnr("S8a S8b")))
    case("S9", TAB)
    out.append(para(tnr(f"S10a {FILLER} {FILLER} {FILLER} filler") + plain(" ") + tnr("S10b")))
    out.append(para(tnr("S10z")))
    out.append(f"<w:p>{tnr('S11a S11b')}</w:p>")
    case("T1", plain(" "), SINGLE)
    case("T2", tnr(" ", 40), SINGLE)
    case("T3", plain("x"), SINGLE)
    out.append(para(tnr("TZ")))
    return "".join(out)


def main() -> None:
    document = (
        f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        f"<w:document {W}><w:body>{body()}{SECT}</w:body></w:document>"
    )
    parts = {
        "[Content_Types].xml": CONTENT_TYPES,
        "_rels/.rels": ROOT_RELS,
        "word/_rels/document.xml.rels": DOC_RELS,
        "word/document.xml": document,
        "word/styles.xml": STYLES,
    }
    with zipfile.ZipFile(OUT, "w", zipfile.ZIP_DEFLATED) as z:
        for part, content in parts.items():
            info = zipfile.ZipInfo(part, date_time=(1980, 1, 1, 0, 0, 0))
            info.compress_type = zipfile.ZIP_DEFLATED
            z.writestr(info, content.encode("utf-8"))
    print(f"wrote {OUT.relative_to(ROOT)}")


if __name__ == "__main__":
    main()
