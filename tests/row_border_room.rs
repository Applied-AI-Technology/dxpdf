//! Which horizontal border between two table rows takes room in the rows, over
//! `test-files/row-border-room.docx` (`scripts/make_row_border_room_fixture.py`),
//! measured in Word (Word for Mac 16, PDF export, 2026-10-16).
//!
//! With no cell margins every interior border takes its width, whichever row
//! or the table declares it: the upper row's `w:tcBorders/w:bottom`, the lower
//! row's `w:top`, both at once (charged once, not twice), `insideH`, and a
//! banded style's `band1Horz` borders. A `nil` top on the lower row does not
//! cancel a bottom the upper one draws. Each case is three Calibri 11pt rows
//! under a label, ruled by 6pt borders; the room an edge takes is its pitch
//! less the borderless control's (B0), so no face metric is pinned.
//!
//! B8, a `w:tblHeader` row with a direct bottom border over the same banded
//! body under `firstRow`, is a banding question rather than a border one: Word
//! bands the header row too, its style defining no `firstRow` layer. It is
//! asserted as `header-band.docx`'s H1, which repeats it beside the cases that
//! settle the rule (`tests/header_band.rs`).

use dxpdf::render::layout::draw_command::{DrawCommand, LayoutedPage};

fn pages() -> Vec<LayoutedPage> {
    let dir = env!("CARGO_MANIFEST_DIR");
    let bytes = std::fs::read(format!("{dir}/test-files/row-border-room.docx")).expect("fixture");
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

/// The pitches from a case's label to its first row and between its rows.
fn pitches(pages: &[LayoutedPage], case: &str) -> [f32; 3] {
    let labels = [
        format!("SEP{case}"),
        format!("{case}1"),
        format!("{case}2"),
        format!("{case}3"),
    ];
    let at = labels.each_ref().map(|l| at(pages, l));
    std::array::from_fn(|i| {
        assert_eq!(
            at[i].0,
            at[i + 1].0,
            "{} and {} share a page",
            labels[i],
            labels[i + 1]
        );
        at[i + 1].1 - at[i].1
    })
}

/// The room the borders take on each of a case's three edges — label to row
/// 1, row 1 to 2, row 2 to 3 — against the borderless control.
fn assert_room(pages: &[LayoutedPage], case: &str, expected: [f32; 3]) {
    let control = pitches(pages, "B0");
    let got = pitches(pages, case);
    let room: [f32; 3] = std::array::from_fn(|i| got[i] - control[i]);
    assert!(
        room.iter().zip(expected).all(|(r, e)| (r - e).abs() < 0.01),
        "{case}: room {room:?}, expected {expected:?}"
    );
}

#[test]
fn an_interior_border_takes_its_width_whichever_row_declares_it() {
    let p = pages();
    assert_room(&p, "B1", [0.0, 6.0, 0.0]);
    assert_room(&p, "B2", [0.0, 6.0, 0.0]);
    assert_room(&p, "B6", [0.0, 0.0, 6.0]);
}

#[test]
fn a_border_both_rows_declare_is_charged_once() {
    assert_room(&pages(), "B3", [0.0, 6.0, 0.0]);
}

#[test]
fn an_inside_h_takes_its_width_on_every_interior_edge() {
    assert_room(&pages(), "B4", [0.0, 6.0, 6.0]);
}

#[test]
fn a_nil_top_does_not_cancel_the_bottom_above_it() {
    assert_room(&pages(), "B5", [0.0, 6.0, 0.0]);
}

/// `band1Horz` rules rows 1 and 3 top and bottom, so the table's top and both
/// interior edges carry a 6pt line.
#[test]
fn a_banded_styles_row_borders_take_their_width() {
    assert_room(&pages(), "B7", [6.0, 6.0, 6.0]);
}
