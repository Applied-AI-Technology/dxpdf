//! Whether a table's first row takes part in its row banding when `tblLook`
//! turns `firstRow` on (§17.4.55, §17.7.6), over `test-files/header-band.docx`
//! (`scripts/make_header_band_fixture.py`), measured in Word (Word for Mac 16,
//! PDF export, 2026-10-19).
//!
//! §17.7.6 orders the regions but never says whether the bands count the first
//! row. Word leaves it out only when the table's style defines a `firstRow`
//! layer (H4, H5, H7); under a style with none the first row is banded as if
//! `firstRow` were off (H1–H3 against the H6 control), whether or not it is a
//! `w:tblHeader` row. Each case is four Calibri 11pt rows under a style whose
//! `band1Horz` rules a row top and bottom at 6pt, so a `band1` row 0 adds 6pt
//! between the label and it, and a `band1` row 3 adds 6pt between it and the
//! closing paragraph. The room an edge takes is its pitch less the unruled one
//! after H6's `band2` last row, so no face metric is pinned.

use dxpdf::render::layout::draw_command::{DrawCommand, LayoutedPage};

fn pages() -> Vec<LayoutedPage> {
    let dir = env!("CARGO_MANIFEST_DIR");
    let bytes = std::fs::read(format!("{dir}/test-files/header-band.docx")).expect("fixture");
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

/// The pitches label → row 0 → … → row 3 → closing paragraph of a case.
fn pitches(pages: &[LayoutedPage], case: &str) -> [f32; 5] {
    let labels = [
        format!("SEP{case}"),
        format!("{case}0"),
        format!("{case}1"),
        format!("{case}2"),
        format!("{case}3"),
        format!("END{case}"),
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

/// The room the borders take on each of a case's five edges.
fn assert_room(pages: &[LayoutedPage], case: &str, expected: [f32; 5]) {
    let unruled = pitches(pages, "H6")[4];
    let room = pitches(pages, case).map(|p| p - unruled);
    assert!(
        room.iter().zip(expected).all(|(r, e)| (r - e).abs() < 0.01),
        "{case}: room {room:?}, expected {expected:?}"
    );
}

/// Row 0 is `band1`, so rows 0 and 2 are ruled and row 3 is not.
const FIRST_ROW_BANDED: [f32; 5] = [6.0, 6.0, 6.0, 6.0, 0.0];
/// The bands start at row 1, so rows 1 and 3 are ruled and row 0 is not.
const FIRST_ROW_LEFT_OUT: [f32; 5] = [0.0, 6.0, 6.0, 6.0, 6.0];

#[test]
fn with_first_row_off_every_row_is_banded() {
    assert_room(&pages(), "H6", FIRST_ROW_BANDED);
}

#[test]
fn a_style_with_no_first_row_layer_bands_the_first_row() {
    let p = pages();
    assert_room(&p, "H2", FIRST_ROW_BANDED);
    // A `w:tblHeader` row is no different, with a bottom border of its own
    // or without.
    assert_room(&p, "H1", FIRST_ROW_BANDED);
    assert_room(&p, "H3", FIRST_ROW_BANDED);
}

#[test]
fn a_style_with_a_first_row_layer_leaves_the_first_row_out_of_the_bands() {
    let p = pages();
    assert_room(&p, "H4", FIRST_ROW_LEFT_OUT);
    assert_room(&p, "H5", FIRST_ROW_LEFT_OUT);
    // A layer that sets the row's own borders `nil` leaves it out the same way.
    assert_room(&p, "H7", FIRST_ROW_LEFT_OUT);
}
