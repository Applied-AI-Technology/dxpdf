//! How much room a horizontal border between two table rows takes when the
//! cells have top and bottom margins (§17.4.38 against §17.4.42/§17.4.39),
//! over `test-files/border-margin-room.docx`
//! (`scripts/make_border_margin_room_fixture.py`), measured in Word (Word for
//! Mac 16, PDF export, 2026-10-17).
//!
//! A line a cell draws on its own bottom edge (`w:tcBorders/w:bottom`) between
//! two rows sits in that cell's bottom margin and takes only what is wider
//! than it (B). The same edge from the table's `insideH`, or from the lower
//! cell's own top (C), takes its whole width, and so does any border at the
//! table's foot, the cell's or the table's (F). With no margins every border
//! takes its width. Every line is an exact 12pt, so the pitches below are
//! Word's whatever face draws them; Word's own are these to its 0.24pt grid.

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

fn pitch(pages: &[LayoutedPage], from: &str, to: &str) -> f32 {
    let (pa, ya) = at(pages, from);
    let (pb, yb) = at(pages, to);
    assert_eq!(pa, pb, "{from} and {to} share a page");
    yb - ya
}

fn assert_pitch(pages: &[LayoutedPage], from: &str, to: &str, expected: f32) {
    let got = pitch(pages, from, to);
    assert!(
        (got - expected).abs() < 0.01,
        "{from} -> {to}: {got}, expected {expected}"
    );
}

/// The cell margins of `B0`, `B1`, `B5` and their kin: none, 29 twips, 100.
const MARGINS: [(&str, f32); 3] = [("0", 0.0), ("1", 1.45), ("5", 5.0)];
const BORDERS: [(&str, f32); 2] = [("t", 0.5), ("T", 3.0)];

#[test]
fn a_cells_own_bottom_border_sits_in_its_bottom_margin() {
    let p = pages();
    for (m, margin) in MARGINS {
        for (b, width) in BORDERS {
            let case = format!("B{m}{b}");
            let room = (width - margin).max(0.0);
            assert_pitch(
                &p,
                &format!("{case}r1"),
                &format!("{case}r2"),
                12.0 + 2.0 * margin + room,
            );
            assert_pitch(
                &p,
                &format!("{case}r2"),
                &format!("{case}r3"),
                12.0 + 2.0 * margin,
            );
        }
    }
}

#[test]
fn an_inside_h_takes_its_whole_width() {
    let p = pages();
    for (m, margin) in MARGINS {
        let case = format!("B{m}i");
        for (a, b) in [("r1", "r2"), ("r2", "r3")] {
            assert_pitch(
                &p,
                &format!("{case}{a}"),
                &format!("{case}{b}"),
                12.5 + 2.0 * margin,
            );
        }
    }
}

#[test]
fn the_lower_cells_top_border_takes_its_whole_width() {
    let p = pages();
    for (m, margin) in [("1", 1.45), ("5", 5.0)] {
        for (b, width) in BORDERS {
            let case = format!("C{m}{b}");
            assert_pitch(
                &p,
                &format!("{case}r1"),
                &format!("{case}r2"),
                12.0 + 2.0 * margin + width,
            );
        }
    }
}

#[test]
fn a_border_at_the_tables_foot_takes_its_whole_width() {
    let p = pages();
    for (m, margin) in [("1", 1.45), ("5", 5.0)] {
        for (b, width) in BORDERS {
            for kind in ["F", "G"] {
                let case = format!("{kind}{m}{b}");
                assert_pitch(
                    &p,
                    &format!("{case}r3"),
                    &format!("{case}z"),
                    12.0 + margin + width,
                );
            }
        }
    }
}
