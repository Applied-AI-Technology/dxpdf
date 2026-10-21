//! A package handed over as its parts — what a templating engine holds after
//! filling one — renders exactly as the same package zipped.
//!
//! `convert_packages_batch_with_options` exists so a caller that has just
//! filled a template need not write a ZIP for `parse` to read straight back.
//! Everything after the unzip is shared, so the contract is byte equality with
//! `convert_batch_with_options` over the same documents, and the same failures
//! with the same indices.

use std::io::Read;

use dxpdf::docx::zip::PackageContents;
use dxpdf::{
    convert_batch_with_options, convert_packages_batch_with_options, Error, RenderOptions,
};

fn fixture(name: &str) -> Vec<u8> {
    std::fs::read(format!("{}/test-files/{name}", env!("CARGO_MANIFEST_DIR"))).unwrap()
}

/// The parts of a ZIP, as the names its entries carry.
fn parts(docx: &[u8]) -> Vec<(String, Vec<u8>)> {
    let mut archive = zip::ZipArchive::new(std::io::Cursor::new(docx)).unwrap();
    (0..archive.len())
        .filter_map(|i| {
            let mut file = archive.by_index(i).unwrap();
            if file.is_dir() {
                return None;
            }
            let mut bytes = Vec::new();
            file.read_to_end(&mut bytes).unwrap();
            Some((file.name().to_string(), bytes))
        })
        .collect()
}

fn package(docx: &[u8]) -> PackageContents {
    PackageContents::from_parts(parts(docx))
}

/// Pictures inline and floating, tables, headers and footers, fields, kerned
/// and ligated fonts: a spread of what the parse after the unzip reads.
const DOCUMENTS: &[&str] = &[
    "sample-docx-files-sample-4.docx",
    "sample-docx-files-sample1.docx",
    "picture-transparency.docx",
    "anchor-in-cell.docx",
    "floating-table-x.docx",
    "issue-159-minimal.docx",
    "kerning.docx",
    "oversized-word.docx",
];

#[test]
fn each_document_renders_as_its_zip_does() {
    let options = RenderOptions::default();
    for name in DOCUMENTS {
        let docx = fixture(name);
        let zipped = convert_batch_with_options(&[&docx], &options).unwrap();
        let unzipped = convert_packages_batch_with_options(vec![package(&docx)], &options).unwrap();
        assert!(
            zipped == unzipped,
            "{name}: the parts render differently from the ZIP"
        );
    }
}

#[test]
fn a_batch_of_packages_renders_as_the_batch_of_zips() {
    let options = RenderOptions::default();
    let docs: Vec<Vec<u8>> = [
        "sample-docx-files-sample-4.docx",
        "oversized-word.docx",
        "sample-docx-files-sample-4.docx",
    ]
    .iter()
    .map(|name| fixture(name))
    .collect();
    let zipped = convert_batch_with_options(&docs, &options).unwrap();
    let unzipped =
        convert_packages_batch_with_options(docs.iter().map(|d| package(d)).collect(), &options)
            .unwrap();
    assert!(zipped == unzipped);
}

#[test]
fn a_package_that_does_not_parse_is_named_by_its_index() {
    let good = fixture("sample-docx-files-sample-4.docx");
    let mut broken = parts(&good);
    broken.retain(|(path, _)| path != "_rels/.rels");
    let packages = vec![
        package(&good),
        PackageContents::from_parts(broken),
        package(&good),
    ];
    match convert_packages_batch_with_options(packages, &RenderOptions::default()) {
        Err(Error::Document { index, .. }) => assert_eq!(index, 1),
        other => panic!(
            "expected the second document to fail, got {:?}",
            other.map(|pdf| pdf.len())
        ),
    }
}

#[test]
fn an_empty_batch_is_refused() {
    assert!(matches!(
        convert_packages_batch_with_options(Vec::new(), &RenderOptions::default()),
        Err(Error::EmptyBatch)
    ));
}

/// A templating engine may key a part with a leading slash or another case, as
/// a relationship target can name it; a ZIP's own entry names are normalized
/// the same way.
#[test]
fn part_paths_are_normalized_as_zip_entries_are() {
    let docx = fixture("sample-docx-files-sample-4.docx");
    let shouted = parts(&docx)
        .into_iter()
        .map(|(path, bytes)| (format!("/{}", path.to_uppercase()), bytes));
    let options = RenderOptions::default();
    assert!(
        convert_batch_with_options(&[&docx], &options).unwrap()
            == convert_packages_batch_with_options(
                vec![PackageContents::from_parts(shouted)],
                &options
            )
            .unwrap()
    );
}
