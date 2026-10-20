//! How much room a paragraph's top and bottom borders take (§17.3.1.24,
//! §17.3.1.42/§17.3.1.7), over `test-files/paragraph-border-room.docx`
//! (`scripts/make_paragraph_border_room_fixture.py`), measured in Word (Word
//! for Mac 16, PDF export, 2026-10-17).
//!
//! Each border takes its `w:space` and its own width — above the first line
//! for a top border, below the last for a bottom one — at 0 and 3pt of space
//! and 0.75 and 3pt of width alike. Every line is an exact 12pt, so these are
//! Word's pitches, to its 0.24pt grid, whatever face draws them.

use dxpdf::render::layout::draw_command::{DrawCommand, LayoutedPage};

fn pages() -> Vec<LayoutedPage> {
    let dir = env!("CARGO_MANIFEST_DIR");
    let bytes =
        std::fs::read(format!("{dir}/test-files/paragraph-border-room.docx")).expect("fixture");
    let doc = dxpdf::docx::parse(&bytes).expect("fixture parses");
    dxpdf::render::resolve_and_layout(doc).1
}

/// The baseline of the one text run reading `label`, all on the first page.
fn baseline(pages: &[LayoutedPage], label: &str) -> f32 {
    let hits: Vec<f32> = pages[0]
        .commands
        .iter()
        .filter_map(|c| match c {
            DrawCommand::Text { text, position, .. } if text.trim_end() == label => {
                Some(position.y.raw())
            }
            _ => None,
        })
        .collect();
    assert_eq!(hits.len(), 1, "exactly one {label:?}, got {hits:?}");
    hits[0]
}

fn assert_pitch(pages: &[LayoutedPage], from: &str, to: &str, expected: f32) {
    let got = baseline(pages, to) - baseline(pages, from);
    assert!(
        (got - expected).abs() < 0.01,
        "{from} -> {to}: {got}, expected {expected}"
    );
}

/// `(space, sz)` of each case: the space in points, the width in eighths.
const CASES: [(u32, u32); 4] = [(0, 6), (0, 24), (3, 6), (3, 24)];

#[test]
fn a_top_border_takes_its_space_and_width_above_the_text() {
    let p = pages();
    for (space, sz) in CASES {
        let case = format!("T{space}{sz}");
        let room = space as f32 + sz as f32 / 8.0;
        assert_pitch(&p, &format!("{case}a"), &case, 12.0 + room);
        assert_pitch(&p, &case, &format!("{case}z"), 12.0);
    }
}

#[test]
fn a_bottom_border_takes_its_space_and_width_below_the_text() {
    let p = pages();
    for (space, sz) in CASES {
        let case = format!("B{space}{sz}");
        let room = space as f32 + sz as f32 / 8.0;
        assert_pitch(&p, &format!("{case}a"), &case, 12.0);
        assert_pitch(&p, &case, &format!("{case}z"), 12.0 + room);
    }
}

/// A report template's empty paragraph: its mark's line, 3pt of space
/// and the 0.75pt rule.
#[test]
fn an_empty_bordered_paragraph_takes_a_line_and_its_border() {
    let p = pages();
    assert_pitch(&p, "Ea", "Ez", 12.0 + 12.0 + 3.75);
}
