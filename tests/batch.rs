//! `convert_batch`: several documents rendered into one PDF.
//!
//! A batch is what a mail run is — the same template filled once per record —
//! and the reason it is one render rather than N renders merged afterwards is
//! the font registry: built once, it embeds each face once for the whole
//! batch, where a merge of N single-document PDFs carries N copies of every
//! face. The rest is what a merge would have done anyway, and is asserted so
//! the shared registry cannot quietly change it: the documents come out in
//! order, each paginated and numbered on its own, and a batch of one is the
//! single-document conversion byte for byte.
//!
//! Documents are told apart by page size, so the order is read off the
//! `MediaBox`es and no glyph has to be extracted.

use std::io::Write;

/// A package holding `body` under a section of `width` × `height` twips.
fn docx(width: u32, height: u32, body: &str) -> Vec<u8> {
    package(width, height, body, None)
}

/// [`docx`] with `footer` as the section's default footer.
fn docx_with_footer(width: u32, height: u32, body: &str, footer: &str) -> Vec<u8> {
    package(width, height, body, Some(footer))
}

fn package(width: u32, height: u32, body: &str, footer: Option<&str>) -> Vec<u8> {
    const R_NS: &str =
        r#"xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships""#;
    let footer_ref = match footer {
        Some(_) => format!(r#"<w:footerReference w:type="default" r:id="rIdF1" {R_NS}/>"#),
        None => String::new(),
    };
    let document = format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>{body}<w:sectPr>{footer_ref}<w:pgSz w:w="{width}" w:h="{height}"/>
    <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:footer="720"/></w:sectPr></w:body>
</w:document>"#
    );
    let mut zip = zip::ZipWriter::new(std::io::Cursor::new(Vec::new()));
    let options = zip::write::SimpleFileOptions::default()
        .compression_method(zip::CompressionMethod::Deflated);
    zip.start_file("[Content_Types].xml", options).unwrap();
    zip.write_all(
        br#"<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml"
    ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/footer1.xml"
    ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>
</Types>"#,
    )
    .unwrap();
    zip.start_file("_rels/.rels", options).unwrap();
    zip.write_all(
        br#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
    Target="word/document.xml"/>
</Relationships>"#,
    )
    .unwrap();
    zip.start_file("word/document.xml", options).unwrap();
    zip.write_all(document.as_bytes()).unwrap();
    if let Some(footer) = footer {
        zip.start_file("word/_rels/document.xml.rels", options)
            .unwrap();
        zip.write_all(
            br#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdF1"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer"
    Target="footer1.xml"/>
</Relationships>"#,
        )
        .unwrap();
        zip.start_file("word/footer1.xml", options).unwrap();
        zip.write_all(
            format!(
                r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">{footer}</w:ftr>"#
            )
            .as_bytes(),
        )
        .unwrap();
    }
    zip.finish().unwrap().into_inner()
}

fn paragraph(text: &str) -> String {
    format!(r#"<w:p><w:r><w:t xml:space="preserve">{text}</w:t></w:r></w:p>"#)
}

const PAGE_BREAK: &str = r#"<w:p><w:r><w:br w:type="page"/></w:r></w:p>"#;

/// `PAGE of NUMPAGES` as simple fields, each with a cached result no renderer
/// that evaluates them would print. Page fields are evaluated where templates
/// put them, in a header or footer.
fn page_of_pages() -> String {
    r#"<w:p><w:r><w:t xml:space="preserve">page </w:t></w:r><w:fldSimple w:instr="PAGE"><w:r><w:t>99</w:t></w:r></w:fldSimple><w:r><w:t xml:space="preserve"> of </w:t></w:r><w:fldSimple w:instr="NUMPAGES"><w:r><w:t>99</w:t></w:r></w:fldSimple></w:p>"#
        .to_string()
}

const LETTER: (u32, u32) = (12240, 15840);
const LEGAL: (u32, u32) = (12240, 20160);
const A5: (u32, u32) = (8391, 11906);

/// Each page's (width, height) in points, in page order.
fn page_sizes(pdf: &[u8]) -> Vec<(i64, i64)> {
    let doc = lopdf::Document::load_mem(pdf).expect("output parses as PDF");
    doc.get_pages()
        .values()
        .map(|&id| {
            let page = doc.get_dictionary(id).unwrap();
            let mb = page.get(b"MediaBox").unwrap().as_array().unwrap();
            let n = |o: &lopdf::Object| o.as_float().map(|f| f.round() as i64).unwrap();
            (n(&mb[2]) - n(&mb[0]), n(&mb[3]) - n(&mb[1]))
        })
        .collect()
}

fn points((w, h): (u32, u32)) -> (i64, i64) {
    (
        (w as f64 / 20.0).round() as i64,
        (h as f64 / 20.0).round() as i64,
    )
}

/// Embedded faces: one `/FontDescriptor` per face (a CID-keyed face has two
/// font dictionaries, so those would double-count).
fn embedded_faces(pdf: &[u8]) -> usize {
    let doc = lopdf::Document::load_mem(pdf).unwrap();
    doc.objects
        .values()
        .filter_map(|o| o.as_dict().ok())
        .filter(|d| {
            d.get(b"Type")
                .ok()
                .and_then(|t| t.as_name().ok())
                .is_some_and(|n| n == b"FontDescriptor")
        })
        .count()
}

/// A page's text with all white space removed: extraction breaks lines at
/// every run, so `page 1 of 2` comes back split across five.
fn page_text(pdf: &[u8], page: u32) -> String {
    lopdf::Document::load_mem(pdf)
        .unwrap()
        .extract_text(&[page])
        .unwrap()
        .split_whitespace()
        .collect()
}

#[test]
fn documents_come_out_in_order_each_paginated_on_its_own() {
    let a = docx(LETTER.0, LETTER.1, &paragraph("first"));
    let b = docx(
        LEGAL.0,
        LEGAL.1,
        &format!("{}{PAGE_BREAK}{}", paragraph("second"), paragraph("second")),
    );
    let c = docx(A5.0, A5.1, &paragraph("third"));

    let pdf = dxpdf::convert_batch(&[&a, &b, &c]).expect("the batch converts");

    assert_eq!(
        page_sizes(&pdf),
        vec![points(LETTER), points(LEGAL), points(LEGAL), points(A5)]
    );
}

#[test]
fn page_fields_count_within_their_own_document() {
    let one = docx_with_footer(LETTER.0, LETTER.1, &paragraph("one"), &page_of_pages());
    let two = docx_with_footer(
        LETTER.0,
        LETTER.1,
        &format!("{}{PAGE_BREAK}{}", paragraph("two"), paragraph("two")),
        &page_of_pages(),
    );

    let pdf = dxpdf::convert_batch(&[&one, &two]).unwrap();

    let texts: Vec<String> = (1..=3).map(|p| page_text(&pdf, p)).collect();
    assert!(texts[0].contains("page1of1"), "page 1: {:?}", texts[0]);
    assert!(texts[1].contains("page1of2"), "page 2: {:?}", texts[1]);
    assert!(texts[2].contains("page2of2"), "page 3: {:?}", texts[2]);
}

/// The point of a batch: a face the documents share is embedded once, not
/// once per document as a merge of single-document PDFs would carry it.
#[test]
fn a_face_shared_by_every_document_is_embedded_once() {
    let letter = |name: &str| docx(LETTER.0, LETTER.1, &paragraph(&format!("Dear {name},")));
    let (a, b, c) = (letter("Ada"), letter("Grace"), letter("Edsger"));

    let single = dxpdf::convert(&a).unwrap();
    let batch = dxpdf::convert_batch(&[&a, &b, &c]).unwrap();

    assert!(embedded_faces(&single) >= 1, "the control embeds its face");
    assert_eq!(embedded_faces(&batch), embedded_faces(&single));
}

#[test]
fn a_batch_of_one_is_the_single_conversion() {
    let a = docx_with_footer(LETTER.0, LETTER.1, &paragraph("only"), &page_of_pages());

    assert_eq!(
        dxpdf::convert_batch(&[&a]).unwrap(),
        dxpdf::convert(&a).unwrap()
    );
}

#[test]
fn a_document_that_fails_is_named_by_its_index() {
    let good = docx(LETTER.0, LETTER.1, &paragraph("fine"));
    let bad = b"not a zip archive".to_vec();

    let err = dxpdf::convert_batch(&[&good, &good, &bad, &good]).unwrap_err();

    match err {
        dxpdf::Error::Document { index, source } => {
            assert_eq!(index, 2);
            assert!(matches!(*source, dxpdf::Error::Parse(_)), "{source}");
        }
        other => panic!("expected the failing document's index, got {other}"),
    }
}

#[test]
fn an_empty_batch_is_refused() {
    let none: [&[u8]; 0] = [];
    assert!(matches!(
        dxpdf::convert_batch(&none),
        Err(dxpdf::Error::EmptyBatch)
    ));
}

/// The same holds for fonts a document embeds: each document of the batch
/// carries its own copy, and registering every copy would put several faces
/// under one name. `sample-docx-files-sample1.docx` embeds Ubuntu.
#[test]
fn an_embedded_face_every_document_carries_is_registered_once() {
    let sample = std::fs::read(concat!(
        env!("CARGO_MANIFEST_DIR"),
        "/test-files/sample-docx-files-sample1.docx"
    ))
    .unwrap();

    let single = dxpdf::convert(&sample).unwrap();
    let batch = dxpdf::convert_batch(&[&sample, &sample]).unwrap();

    assert_eq!(page_sizes(&batch).len(), 2 * page_sizes(&single).len());
    assert_eq!(embedded_faces(&batch), embedded_faces(&single));
}

/// Image XObjects in a PDF.
fn embedded_images(pdf: &[u8]) -> usize {
    let doc = lopdf::Document::load_mem(pdf).unwrap();
    doc.objects
        .values()
        .filter_map(|o| o.as_stream().ok())
        .filter(|s| {
            s.dict
                .get(b"Subtype")
                .ok()
                .and_then(|t| t.as_name().ok())
                .is_some_and(|n| n == b"Image")
        })
        .count()
}

/// A picture every document carries — a town's logo on every letter — is
/// embedded once. Each document holds its own copy of the bytes, so this is
/// identity by content, not by buffer.
#[test]
fn a_picture_every_document_carries_is_embedded_once() {
    let pictures = std::fs::read(concat!(
        env!("CARGO_MANIFEST_DIR"),
        "/test-files/picture-transparency.docx"
    ))
    .unwrap();

    let single = dxpdf::convert(&pictures).unwrap();
    let batch = dxpdf::convert_batch(&[&pictures, &pictures, &pictures]).unwrap();

    assert!(
        embedded_images(&single) >= 1,
        "the control embeds its pictures"
    );
    assert_eq!(embedded_images(&batch), embedded_images(&single));
}
