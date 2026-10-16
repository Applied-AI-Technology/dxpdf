//! §17.4.83 `vAlign` in a row that one of its cells gives a bottom border,
//! over `test-files/cell-valign-bottom.docx`
//! (`scripts/make_cell_valign_bottom_fixture.py`), measured in Word
//! (Microsoft Word for Mac 16, PDF export, 2026-10-16).
//!
//! The room a bottom border drawn inside the row takes is the row's: Word sets
//! every cell's content above it, a cell whose own bottom is `nil` included,
//! so the two cells of each table share a baseline. Word's numbers are quoted
//! per case.

use dxpdf::render::layout::draw_command::{DrawCommand, LayoutedPage};

fn pages() -> Vec<LayoutedPage> {
    let path = format!(
        "{}/test-files/cell-valign-bottom.docx",
        env!("CARGO_MANIFEST_DIR")
    );
    let bytes = std::fs::read(&path).expect("fixture is committed");
    let doc = dxpdf::docx::parse(&bytes).expect("fixture parses");
    dxpdf::render::resolve_and_layout(doc).1
}

fn y(pages: &[LayoutedPage], label: &str) -> f32 {
    let hits: Vec<f32> = pages
        .iter()
        .flat_map(|p| p.commands.iter())
        .filter_map(|c| match c {
            DrawCommand::Text { text, position, .. } if &**text == label => Some(position.y.raw()),
            _ => None,
        })
        .collect();
    assert_eq!(hits.len(), 1, "exactly one {label:?}, got {hits:?}");
    hits[0]
}

fn assert_level(pages: &[LayoutedPage], case: &str, why: &str) {
    let a = y(pages, &format!("{case}a"));
    let b = y(pages, &format!("{case}b"));
    assert!((a - b).abs() < 0.01, "{case}: {why} — {a} and {b}");
}

/// Bottom-aligned, on the table's last row, where the bordered cell draws its
/// rule inside the row: Word sets both cells' text on one baseline beside a
/// 0.5pt rule (V1, 111.84 each) and a 3pt one (V2, 155.04 each).
#[test]
fn a_bottom_aligned_cell_stands_above_the_row_s_bottom_border() {
    let p = pages();
    assert_level(&p, "V1", "a 0.5pt rule");
    assert_level(&p, "V2", "a 3pt rule");
}

/// The same with the rule between two rows, where it has its own strip (V3,
/// 200.64 each), centred (V4, 249.84 each), and with no rule at all (V5,
/// 303.12 each).
#[test]
fn the_row_s_inset_holds_for_every_alignment_and_none() {
    let p = pages();
    assert_level(&p, "V3", "an interior rule");
    assert_level(&p, "V4", "centred");
    assert_level(&p, "V5", "no rule");
}
