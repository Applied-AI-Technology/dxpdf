//! §17.3.1.9 `w:contextualSpacing` as Word applies it, over the S, K, L and P
//! cases of `test-files/border-margin-room.docx`
//! (`scripts/make_border_margin_room_fixture.py`), measured in Word (Word for
//! Mac 16, PDF export, 2026-10-17).
//!
//! A paragraph's own flag drops its own space on the side facing a paragraph
//! of its style; the neighbour needs only the style (P1-P3). Paragraphs with
//! no `pStyle` are both in the default style. The flag reaches across into a
//! table whose first cell opens with a paragraph of the style, flagged or not
//! (K20, L20). Every line is an exact 12pt, so the pitches are Word's
//! whatever face draws them. The unflagged control, S20, puts its table at
//! the top of the next page in Word and in dxpdf, so it measures nothing here.

use dxpdf::render::layout::draw_command::{DrawCommand, LayoutedPage};

fn pages() -> Vec<LayoutedPage> {
    let dir = env!("CARGO_MANIFEST_DIR");
    let bytes =
        std::fs::read(format!("{dir}/test-files/border-margin-room.docx")).expect("fixture");
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

fn assert_pitch(pages: &[LayoutedPage], from: &str, to: &str, expected: f32, why: &str) {
    let (pa, ya) = at(pages, from);
    let (pb, yb) = at(pages, to);
    assert!(
        pa == pb && (yb - ya - expected).abs() < 0.01,
        "{from} -> {to}: {} on pages {pa}/{pb}, expected {expected} — {why}",
        yb - ya
    );
}

#[test]
fn each_paragraphs_flag_drops_its_own_space_after() {
    let p = pages();
    assert_pitch(
        &p,
        "P1",
        "P2",
        12.0,
        "both flagged, both in the default style",
    );
    assert_pitch(&p, "P2", "P3", 12.0, "only the first flagged");
}

#[test]
fn the_flag_reaches_into_a_table_of_its_style() {
    let p = pages();
    assert_pitch(&p, "K20", "K20r1", 12.0, "the cells flagged too");
    assert_pitch(&p, "L20", "L20r1", 12.0, "the cells not flagged");
}
