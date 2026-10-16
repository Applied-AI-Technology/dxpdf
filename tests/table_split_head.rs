//! The row that heads the page a table has broken onto, over
//! `test-files/table-split-head.docx` (`scripts/make_table_split_head_fixture.py`),
//! measured in Word (Microsoft Word for Mac 16, PDF export, 2026-10-16).
//!
//! Each case is three rows, the third carried onto the next page, then an
//! `END` paragraph; the distance from the third row's label to `END` is that
//! row's height there, less the restored top border that pushes the label
//! down, plus the 0.5pt bottom border. Word restores the table's top border on
//! the carried row and grows the row by it, and an `exact` row still gives up
//! the interior boundary it had above it on the page before. Word's numbers are
//! quoted per case; they sit within 0.1pt of these, Word's own positioning
//! grain in its PDFs.

use dxpdf::render::layout::draw_command::{DrawCommand, LayoutedPage};

fn pages() -> Vec<LayoutedPage> {
    let path = format!(
        "{}/test-files/table-split-head.docx",
        env!("CARGO_MANIFEST_DIR")
    );
    let bytes = std::fs::read(&path).expect("fixture is committed");
    let doc = dxpdf::docx::parse(&bytes).expect("fixture parses");
    dxpdf::render::resolve_and_layout(doc).1
}

/// (page index, baseline y) of the one text command reading `label`.
fn at(pages: &[LayoutedPage], label: &str) -> (usize, f32) {
    let hits: Vec<(usize, f32)> = pages
        .iter()
        .enumerate()
        .flat_map(|(i, p)| p.commands.iter().map(move |c| (i, c)))
        .filter_map(|(i, c)| match c {
            DrawCommand::Text { text, position, .. } if &**text == label => {
                Some((i, position.y.raw()))
            }
            _ => None,
        })
        .collect();
    assert_eq!(hits.len(), 1, "exactly one {label:?}, got {hits:?}");
    hits[0]
}

/// The carried row's label to `END`, after checking the row was carried.
fn carried(pages: &[LayoutedPage], case: &str) -> f32 {
    let (p2, _) = at(pages, &format!("{case}r2"));
    let (p3, y3) = at(pages, &format!("{case}r3"));
    let (pe, ye) = at(pages, &format!("END{case}"));
    assert_eq!(p3, p2 + 1, "{case}: the third row goes over");
    assert_eq!(pe, p3, "{case}: END follows it");
    ye - y3
}

fn assert_carried(pages: &[LayoutedPage], case: &str, want: f32, why: &str) {
    let got = carried(pages, case);
    assert!(
        (got - want).abs() < 0.01,
        "{case}: {why} — want {want}, got {got}"
    );
}

/// Distance from a carried row's label to `END`: the row's box, less the top
/// border that pushes the label down, plus the 0.5pt bottom border.
fn label_to_end(box_height: f32, top: f32) -> f32 {
    box_height - top + 0.5
}

/// An `exact` row carried over is its interior height plus the restored top
/// border: 36pt and the 4.95pt bottom margin, less the 0.5pt rule above it,
/// plus the 0.5pt top. Word: 41.04 to `END`.
#[test]
fn an_exact_row_grows_by_the_restored_top_border() {
    assert_carried(
        &pages(),
        "H1",
        label_to_end(36.0 + 4.95 - 0.5 + 0.5, 0.5),
        "a 40.95pt box",
    );
}

/// The boundary an `exact` row gives up is the one it had above it, not the
/// border it is given back. H2: 36 less a 6pt `insideH`, plus a 0.5pt top —
/// Word: 30.48. H3: 36 less a 0.5pt `insideH`, plus a 6pt top that also
/// pushes the label down — Word: 36.00. H8: no top border to restore — Word:
/// 36.00.
#[test]
fn an_exact_row_still_gives_up_the_boundary_it_had() {
    let p = pages();
    assert_carried(
        &p,
        "H2",
        label_to_end(36.0 - 6.0 + 0.5, 0.5),
        "the 6pt rule stays given up",
    );
    assert_carried(
        &p,
        "H3",
        label_to_end(36.0 - 0.5 + 6.0, 6.0),
        "the 6pt top is inside the row",
    );
    assert_carried(&p, "H8", label_to_end(36.0 - 0.5, 0.0), "nothing restored");
}

/// A row sized by its content, or by `atLeast`, grows by the restored top the
/// same way, so the label moves down with the border and the distance to `END`
/// is the row's own. Word: H4 and H5 14.40, H6 and H7 36.48.
#[test]
fn a_content_sized_row_grows_by_the_restored_top_border_too() {
    let p = pages();
    let line = carried(&p, "H4");
    assert!(
        (carried(&p, "H5") - line).abs() < 0.01,
        "a 6pt top moves the label with the row"
    );
    assert_carried(&p, "H6", label_to_end(36.0 + 0.5, 0.5), "atLeast 36");
    assert_carried(
        &p,
        "H7",
        label_to_end(36.0 + 6.0, 6.0),
        "the 6pt top moves the label with the row",
    );
}
