//! A picture anchored in a `w:vAlign` cell that wraps text, over
//! `test-files/cell-float-valign.docx` (`scripts/make_cell_float_valign_fixture.py`),
//! measured in Word (Word for Mac 16, PDF export, 2026-10-18).
//!
//! §20.4.2.18 `wrapTopAndBottom` keeps text off both sides of the picture,
//! but only the lines that would meet it move: an empty paragraph whose
//! picture is anchored 25.95pt below its top keeps its line above the
//! picture, so the cell's content ends at the picture's foot (59.2pt), the
//! row keeps its 73.8pt minimum, and a bottom-aligned cell stands the picture
//! on the row's bottom border — haddam's dog-licence signature (R1). Anchored
//! at the paragraph's top (R4) the line goes under the picture. The other
//! wrap modes (R2, R3) align as R1 does; a picture reaching past the minimum
//! grows the row to its foot and no further (R5).
//!
//! Word's numbers are the picture's top below its row's top rule and the
//! row's height between rules.

use dxpdf::render::layout::draw_command::{DrawCommand, LayoutedPage};

fn pages() -> Vec<LayoutedPage> {
    let path = format!(
        "{}/test-files/cell-float-valign.docx",
        env!("CARGO_MANIFEST_DIR")
    );
    let bytes = std::fs::read(&path).expect("fixture is committed");
    let doc = dxpdf::docx::parse(&bytes).expect("fixture parses");
    dxpdf::render::resolve_and_layout(doc).1
}

/// The tops of the table's horizontal rules, in page order.
fn rules(page: &LayoutedPage) -> Vec<f32> {
    let mut ys: Vec<f32> = page
        .commands
        .iter()
        .filter_map(|c| match c {
            DrawCommand::Rect { rect, .. }
                if rect.size.height.raw() < 1.5 && rect.size.width.raw() > 100.0 =>
            {
                Some(rect.origin.y.raw())
            }
            _ => None,
        })
        .collect();
    ys.sort_by(f32::total_cmp);
    ys.dedup_by(|a, b| (*a - *b).abs() < 0.01);
    ys
}

/// The pictures' tops, in page order.
fn pictures(page: &LayoutedPage) -> Vec<f32> {
    let mut ys: Vec<f32> = page
        .commands
        .iter()
        .filter_map(|c| match c {
            DrawCommand::Image { rect, .. } => Some(rect.origin.y.raw()),
            _ => None,
        })
        .collect();
    ys.sort_by(f32::total_cmp);
    ys
}

fn assert_near(got: f32, want: f32, why: &str) {
    assert!((got - want).abs() < 0.3, "{why}: want {want}, got {got}");
}

/// (picture top below the row's top rule, row height) per row, Word's.
const WORD: [[(f32, f32); 4]; 2] = [
    [(41.13, 74.4), (41.03, 74.4), (40.93, 74.16), (25.63, 74.4)],
    [
        (26.53, 86.64),
        (40.94, 74.16),
        (26.48, 74.4),
        (33.68, 74.16),
    ],
];

#[test]
fn a_cell_aligns_on_its_picture_s_foot() {
    let pages = pages();
    for (page, (word, first)) in pages.iter().zip(WORD.iter().zip([1, 5])) {
        let rules = rules(page);
        let pictures = pictures(page);
        assert_eq!(rules.len(), 5, "four rows between five rules: {rules:?}");
        for (i, &(top, height)) in word.iter().enumerate() {
            let row = format!("R{}", first + i);
            assert_near(pictures[i] - rules[i], top, &format!("{row} picture"));
            assert_near(rules[i + 1] - rules[i], height, &format!("{row} height"));
        }
    }
}

/// R6's text shares the picture's paragraph and stays on its line above the
/// picture, moved down only by the cell's alignment.
#[test]
fn a_line_above_the_picture_stays_above_it() {
    let pages = pages();
    let r6p = pages[1]
        .commands
        .iter()
        .find_map(|c| match c {
            DrawCommand::Text { text, position, .. } if &**text == "R6P" => Some(position.y.raw()),
            _ => None,
        })
        .expect("R6P is drawn");
    let picture = pictures(&pages[1])[1];
    assert!(r6p < picture, "R6P at {r6p}, its picture at {picture}");
}
