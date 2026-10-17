//! §17.4.84 `w:vAlign` in a row that breaks across pages, and how many lines
//! the row's first half holds, over `test-files/split-row-valign.docx`
//! (`scripts/make_split_row_valign_fixture.py`), measured in Word (Word for
//! Mac 16, PDF export, 2026-10-17).
//!
//! A cell the break leaves whole stays on the first page and keeps its
//! alignment within the first half (V1). A cell the break cuts through is
//! laid out top-down on both pages, bottom- or centre-aligned alike (V2, V3)
//! — holdenma's renewal notices' right-hand cell. The first half charges the
//! 0.5pt borders it draws as well as the cell margins, so nine 12pt lines fit
//! in the 120pt left, not ten. Every line is an exact 12pt, so dxpdf's
//! positions match Word's whatever face draws them.

use dxpdf::render::layout::draw_command::{DrawCommand, LayoutedPage};

fn pages() -> Vec<LayoutedPage> {
    let dir = env!("CARGO_MANIFEST_DIR");
    let bytes = std::fs::read(format!("{dir}/test-files/split-row-valign.docx")).expect("fixture");
    let doc = dxpdf::docx::parse(&bytes).expect("fixture parses");
    dxpdf::render::resolve_and_layout(doc).1
}

/// The page and baseline of the one text run reading `label`.
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

/// The first half holds nine of the left cell's lines; the tenth starts the
/// next page.
fn assert_nine_lines(pages: &[LayoutedPage], case: &str) {
    let (p9, _) = at(pages, &format!("{case}L9"));
    let (p10, _) = at(pages, &format!("{case}L10"));
    assert_eq!(p10, p9 + 1, "{case}: L10 starts the next page");
}

#[test]
fn the_first_half_charges_its_borders() {
    let p = pages();
    for case in ["V1", "V2", "V3"] {
        assert_nine_lines(&p, case);
    }
}
