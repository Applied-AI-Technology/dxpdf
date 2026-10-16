//! How far a cell's text stands in from the border on its right, over
//! `test-files/cell-inset-right-{15,14}.docx`
//! (`scripts/make_cell_inset_right_fixtures.py`), measured in Word (Microsoft
//! Word for Mac 16, PDF export, 2026-10-16).
//!
//! Word sizes a cell's text from its two margins alone and starts it at the
//! larger of the left margin and half the left border, so a right border takes
//! no room — right-aligned text runs into it — and a left border wider than
//! its margin shifts the text right, until the shift would carry it past the
//! right grid line, where it stops at the right border's inner edge. Every
//! comparison here is between labels of one length in one face, so only the
//! insets differ. Word's insets, from the rule's centre, are quoted per case.

use dxpdf::render::layout::draw_command::{DrawCommand, LayoutedPage};

fn pages(mode: u32) -> Vec<LayoutedPage> {
    let path = format!(
        "{}/test-files/cell-inset-right-{mode}.docx",
        env!("CARGO_MANIFEST_DIR")
    );
    let bytes = std::fs::read(&path).expect("fixture is committed");
    let doc = dxpdf::docx::parse(&bytes).expect("fixture parses");
    dxpdf::render::resolve_and_layout(doc).1
}

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

fn assert_shift(p: &[LayoutedPage], from: &str, to: &str, want: f32, why: &str) {
    let got = x(p, to) - x(p, from);
    assert!(
        (got - want).abs() < 0.01,
        "{from} → {to}: {why} — want {want}, got {got}"
    );
}

/// A table's right border takes no room from right-aligned text, whatever its
/// width: the text ends a margin from the rule's centre. Word: 2.12, 2.00 and
/// 2.12 for 3, 6 and 12pt rules over a 2pt margin; 0.12 and 0.00 for 3 and
/// 6pt over none.
#[test]
fn a_right_border_takes_no_room() {
    for mode in [15, 14] {
        let p = pages(mode);
        assert_shift(&p, "RO24m40z", "RO48m40z", 0.0, "6pt as 3pt");
        assert_shift(&p, "RO24m40z", "RO96m40z", 0.0, "12pt as 3pt");
        assert_shift(&p, "RO24m0z", "RO48m0z", 0.0, "no margin");
    }
}

/// A left border wider than its margin shifts the cell's text right: with no
/// left margin and the default 5.4pt right one, right-aligned text ends 3.9pt
/// in under a 3pt left rule and 2.4pt under a 6pt one, and under a 12pt one
/// stops on the thin right rule's inner edge. Word: 3.96, 2.52, 0.25.
///
/// The wider rule also moves the grid by the difference in its half: right
/// under Word 2013, which puts the table's outside on the indent, and left in
/// compatibility mode, which puts the first cell's text there.
#[test]
fn a_wide_left_border_shifts_the_text_right() {
    for (mode, grid) in [(15, 1.0), (14, -1.0)] {
        let p = pages(mode);
        assert_shift(&p, "LO24m0y", "LO48m0y", grid * 1.5 + 1.5, "3pt more shift");
        assert_shift(
            &p,
            "LO24m0y",
            "LO96m0y",
            grid * 4.5 + (3.9 - 0.25),
            "stopped at the rule",
        );
    }
}

/// With no margins the first cell's text is carried to the shared rule and
/// stops flush against it, so a wider rule moves it left by the difference in
/// their halves. Word: 1.51 from a 3pt rule's centre and 6.01 from a 12pt one's.
#[test]
fn text_carried_past_the_grid_line_stops_at_the_rule() {
    for mode in [15, 14] {
        let p = pages(mode);
        assert_shift(&p, "IV24m0y", "IV96m0y", -(6.0 - 1.5), "flush with each");
    }
}
