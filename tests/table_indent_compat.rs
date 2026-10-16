//! §17.4.50 `w:tblInd` and §17.4.57 `tblpX` against §17.15.1.21
//! `compatibilityMode`, over `test-files/table-indent-compat-{15,14,none}.docx`
//! (`scripts/make_table_indent_compat_fixtures.py`) — the same tables three
//! times, differing only in the mode — measured in Word (Microsoft Word for
//! Mac 16, PDF export, 2026-10-15).
//!
//! In compatibility mode (14, or no mode at all) a top-level table's indent is
//! where its first cell's **text** goes; under Word 2013 (15) it is where the
//! table's outer edge goes, so the text sits half the leading border and one
//! cell margin further in. A nested table follows the Word 2013 rule in every
//! mode. Word's numbers are quoted per case; they sit within 0.12pt of these,
//! which is Word's own positioning grain in its PDFs.

use dxpdf::render::layout::draw_command::{DrawCommand, LayoutedPage};

fn pages(mode: &str) -> Vec<LayoutedPage> {
    let path = format!(
        "{}/test-files/table-indent-compat-{mode}.docx",
        env!("CARGO_MANIFEST_DIR")
    );
    let bytes = std::fs::read(&path).expect("fixture is committed");
    let doc = dxpdf::docx::parse(&bytes).expect("fixture parses");
    dxpdf::render::resolve_and_layout(doc).1
}

/// x of the one text command reading `label`, wherever it is.
fn x(pages: &[LayoutedPage], label: &str) -> f32 {
    let hits: Vec<f32> = pages
        .iter()
        .flat_map(|p| p.commands.iter())
        .filter_map(|c| match c {
            DrawCommand::Text { text, position, .. } if &**text == label => Some(position.x.raw()),
            _ => None,
        })
        .collect();
    assert_eq!(hits.len(), 1, "exactly one {label:?}, got {hits:?}");
    hits[0]
}

const MARGIN: f32 = 72.0;
/// `TableNormal`'s 108-twip cell margin.
const CELL: f32 = 5.4;
/// Half the 0.5pt borders, and half the 12pt ones of T4 and T5.
const HALF: f32 = 0.25;
const HALF_THICK: f32 = 6.0;

fn assert_at(pages: &[LayoutedPage], label: &str, want: f32, why: &str) {
    let got = x(pages, label);
    assert!(
        (got - want).abs() < 0.01,
        "{label}: {why} — want {want}, got {got}"
    );
}

/// Compatibility mode: the first cell's text is on the indent, full-width or
/// not, whatever the border or the margin. Word: T1–T6 at 72.0, 72.0, 108.0,
/// 108.0, 72.0, 72.0; T11 at 36.0; F1 at 144.0 and F2 at 72.0.
#[test]
fn in_compatibility_mode_the_indent_is_to_the_first_cell_s_text() {
    for mode in ["14", "none"] {
        let p = pages(mode);
        assert_at(&p, "T1a", MARGIN, "a 100% table");
        assert_at(&p, "T2a", MARGIN, "a narrower one too");
        assert_at(&p, "T3a", MARGIN + 36.0, "a 720-twip indent");
        assert_at(&p, "T4a", MARGIN + 36.0, "a 12pt border hangs outside it");
        assert_at(&p, "T5a", MARGIN, "the same at full width");
        assert_at(&p, "T6a", MARGIN, "and a 21.6pt margin hangs outside it");
        assert_at(&p, "T11a", MARGIN - 36.0, "a negative indent");
        assert_at(&p, "F1a", MARGIN + 72.0, "tblpX is an indent too");
        assert_at(&p, "F2a", MARGIN, "and so is its absence");
    }
}

/// Word 2013: the table's outer edge is on the indent and the text one border
/// half and one cell margin in from it. Word: T1 77.76, T2 77.76, T3 113.76,
/// T4 120.0, T5 84.0, T6 93.84, T11 41.76, F1 149.52, F2 77.76.
#[test]
fn under_word_2013_the_indent_is_to_the_table_s_outer_edge() {
    let p = pages("15");
    let thin = HALF + CELL;
    assert_at(&p, "T1a", MARGIN + thin, "a 100% table");
    assert_at(&p, "T2a", MARGIN + thin, "a narrower one");
    assert_at(&p, "T3a", MARGIN + 36.0 + thin, "a 720-twip indent");
    // The 12pt border's inner half outweighs the 5.4pt margin, so the text is
    // a whole border in from the indent.
    assert_at(&p, "T4a", MARGIN + 36.0 + 2.0 * HALF_THICK, "a 12pt border");
    assert_at(
        &p,
        "T5a",
        MARGIN + 2.0 * HALF_THICK,
        "the same at full width",
    );
    assert_at(&p, "T6a", MARGIN + HALF + 21.6, "a 21.6pt margin");
    assert_at(&p, "T11a", MARGIN - 36.0 + thin, "a negative indent");
    assert_at(&p, "F1a", MARGIN + 72.0 + thin, "tblpX");
    assert_at(&p, "F2a", MARGIN + thin, "no tblpX");
}

/// A nested table is measured to its edge in every mode: from the outer cell's
/// text, half a border and one cell margin in. Word: T7n and T7p at 77.76 in
/// compatibility mode (outer text at 72.0) and at 83.28 under Word 2013 (outer
/// text at 77.76).
#[test]
fn a_nested_table_is_measured_to_its_edge_in_every_mode() {
    for mode in ["15", "14", "none"] {
        let p = pages(mode);
        let outer = x(&p, "T7o");
        assert_at(&p, "T7n", outer + HALF + CELL, &format!("mode {mode}"));
        assert_at(
            &p,
            "T7p",
            outer + HALF + CELL,
            &format!("mode {mode}, 100%"),
        );
    }
}

/// A 100% table is wider than the column by both text insets in compatibility
/// mode and exactly the column under Word 2013, so its right-aligned last cell
/// ends that much further left. Word: T1z 5.55pt further left in mode 15
/// than in mode 14.
#[test]
fn a_full_width_table_spans_the_column_differently_by_mode() {
    let shift = x(&pages("15"), "T1z") - x(&pages("14"), "T1z");
    assert!(
        (shift + HALF + CELL).abs() < 0.01,
        "the last cell's text moves in by half a border and a cell margin, got {shift}"
    );
}

/// `jc="center"` places the table as a box, in both modes alike. Word: T10a at
/// 203.28 in each.
#[test]
fn a_centred_table_does_not_move_with_the_mode() {
    let a = x(&pages("15"), "T10a");
    let b = x(&pages("none"), "T10a");
    assert!((a - b).abs() < 0.01, "got {a} and {b}");
}
