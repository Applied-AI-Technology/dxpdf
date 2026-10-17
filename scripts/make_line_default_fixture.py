#!/usr/bin/env python3
"""Build ``test-files/line-default.docx``, which asks what line spacing Word
gives a paragraph when the styles part has no ``w:pPrDefault``.

``paragraph-defaults-none.docx`` measured 8pt after, and lines that look
1.15 but are not: its 10pt Calibri lines step 14.14pt where 1.15 gives
14.04, the same body under a stated single-spaced ``pPrDefault`` matches
dxpdf to the grid, and a stated ``w:line="276"`` under one does too. Across
albanyny's CivicRecordCertificate (8–18pt) the excess scales with the size,
as a multiple near 278/240 would. The render says whether it is a multiple
at all, of what height, and whether it holds in another face.

The styles part is the certificate's: an ``rPrDefault`` of Calibri 10pt and
nothing else, no ``settings.xml``. Each case is one paragraph of eleven or
more lines, so its line pitch is read off ten intervals:

  P<s>  Calibri s pt, no ``w:pPr``                       (s = 8, 10, 12, 20)
  Q<s>  the same with ``w:spacing w:line="276" w:lineRule="auto"``
  R<s>  the same with ``w:line="278"``
  A<s>  Arial s pt, no ``w:pPr``                         (s = 10, 20)
  B<s>  Arial with ``w:line="276"``
  N1–N3 one-line 10pt Calibri paragraphs with no ``w:pPr``, for the space after

Regenerate and commit the result if the content changes; the build is
deterministic and needs no third-party packages.

    scripts/make_line_default_fixture.py
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
    f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles {NS}><w:docDefaults>'
    '<w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:cs="Calibri"/>'
    '<w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr></w:rPrDefault></w:docDefaults>'
    '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>'
    "</w:styles>"
)

SECT = (
    '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" '
    'w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>'
)


def run(text: str, half_points: int, font: str = "") -> str:
    fonts = f'<w:rFonts w:ascii="{font}" w:hAnsi="{font}" w:cs="{font}"/>' if font else ""
    return (
        f'<w:r><w:rPr>{fonts}<w:sz w:val="{half_points}"/><w:szCs w:val="{half_points}"/></w:rPr>'
        f'<w:t xml:space="preserve">{text}</w:t></w:r>'
    )


def para(runs: str, ppr: str = "") -> str:
    ppr_xml = f"<w:pPr>{ppr}</w:pPr>" if ppr else ""
    return f"<w:p>{ppr_xml}{runs}</w:p>"


def block(label: str, points: int, font: str = "", line: int | None = None) -> str:
    # About 90 characters a line at 10pt; enough words for eleven lines at
    # every size.
    words = max(12, 1400 // points)
    text = f"{label} " + " ".join("word" for _ in range(words))
    ppr = f'<w:spacing w:line="{line}" w:lineRule="auto"/>' if line else ""
    return para(run(text, points * 2, font), ppr)


def body() -> str:
    out = []
    for s in (8, 10, 12, 20):
        out += [block(f"P{s}", s), block(f"Q{s}", s, line=276), block(f"R{s}", s, line=278)]
    for s in (10, 20):
        out += [block(f"A{s}", s, "Arial"), block(f"B{s}", s, "Arial", 276)]
    out += [para(run(f"N{n}", 20)) for n in (1, 2, 3)]
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
    }
    out = ROOT / "test-files" / "line-default.docx"
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as z:
        for part, content in parts.items():
            info = zipfile.ZipInfo(part, date_time=(1980, 1, 1, 0, 0, 0))
            info.compress_type = zipfile.ZIP_DEFLATED
            z.writestr(info, content.encode("utf-8"))
    print(f"wrote {out.relative_to(ROOT)}")


if __name__ == "__main__":
    main()
