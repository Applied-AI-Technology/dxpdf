//! §17.3.1.33 `w:spacing/@w:before` at the top of a page, over
//! `test-files/space-before-page-top-{15,14}.docx`
//! (`scripts/make_space_before_fixtures.py`), measured in Word (Microsoft Word
//! for Mac 16, PDF export, 2026-10-16).
//!
//! Each labelled paragraph has 72pt before it and lands at the top of a page.
//! Word keeps the space on a section's first page and drops it at the top of
//! every other, so each case's first line is either 72pt below a dropped one or
//! level with it. Word: kept at 155.04, dropped at 83.04.

use dxpdf::render::layout::draw_command::{DrawCommand, LayoutedPage};

fn pages(mode: u32) -> Vec<LayoutedPage> {
    let path = format!(
        "{}/test-files/space-before-page-top-{mode}.docx",
        env!("CARGO_MANIFEST_DIR")
    );
    let bytes = std::fs::read(&path).expect("fixture is committed");
    let doc = dxpdf::docx::parse(&bytes).expect("fixture parses");
    dxpdf::render::resolve_and_layout(doc).1
}

/// (page, baseline y) of the text command starting with `label` and a space.
fn at(pages: &[LayoutedPage], label: &str) -> (usize, f32) {
    let hits: Vec<(usize, f32)> = pages
        .iter()
        .enumerate()
        .flat_map(|(i, p)| p.commands.iter().map(move |c| (i, c)))
        .filter_map(|(i, c)| match c {
            DrawCommand::Text { text, position, .. } if text.trim_end() == label => {
                Some((i, position.y.raw()))
            }
            _ => None,
        })
        .collect();
    assert_eq!(hits.len(), 1, "exactly one {label:?}, got {hits:?}");
    hits[0]
}

/// How far below a dropped space `label`'s first line sits: 72 or 0.
fn kept(pages: &[LayoutedPage], label: &str) -> f32 {
    let (_, dropped) = at(pages, "S2");
    let (page, y) = at(pages, label);
    assert!(page > 0 || label == "S1", "{label} heads its own page");
    y - dropped
}

fn assert_kept(pages: &[LayoutedPage], label: &str, want: f32, why: &str) {
    let got = kept(pages, label);
    assert!(
        (got - want).abs() < 0.01,
        "{label}: {why} — want {want}, got {got}"
    );
}

/// The first page of a section keeps the space: the document's first
/// paragraph (S1) and a `nextPage` section's (S6).
#[test]
fn a_section_s_first_page_keeps_the_space() {
    for mode in [15, 14] {
        let p = pages(mode);
        assert_kept(&p, "S1", 72.0, "the document's first paragraph");
        assert_kept(&p, "S6", 72.0, "a new section's");
    }
}

/// Every other page top drops it: after a page-break paragraph (S2, the
/// reference), after a break at the end of the paragraph before (S7), and
/// where the flow carried the paragraph over (S5).
#[test]
fn every_other_page_top_drops_it() {
    for mode in [15, 14] {
        let p = pages(mode);
        assert_kept(&p, "S7", 0.0, "after a closing page break");
        assert_kept(&p, "S5", 0.0, "carried over by the flow");
    }
}

/// §17.3.1.23 `pageBreakBefore` is the one the modes disagree on: Word 2013
/// drops the space at the top of the page it forces, 14 keeps it.
#[test]
fn page_break_before_follows_the_compatibility_mode() {
    assert_kept(&pages(15), "S4", 0.0, "Word 2013");
    assert_kept(&pages(14), "S4", 72.0, "compatibility mode");
}
