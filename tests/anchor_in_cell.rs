//! §20.4.2 anchored pictures in table cells, over `test-files/anchor-in-cell.docx`
//! (`scripts/make_anchor_in_cell_fixture.py`), measured in Word (Microsoft Word for Mac 16, PDF export, 2026-10-15) and asserted
//! as relations between markers on the same page, so no glyph metric is pinned.
//!
//! **`anchor-in-cell.docx`** — how much of a table row an anchored picture
//! takes. Every picture row sits between two one-line marker rows, so a row's
//! height is the distance between its neighbouring markers less one marker row.
//! Word, against a control row of one empty line (28.6pt between markers):
//!
//! | row | picture                          | Word, marker to marker |
//! |-----|----------------------------------|------------------------|
//! | 2   | `wrapNone`                       | 28.5 — no growth       |
//! | 3   | `wrapNone`, `behindDoc`          | 28.8 — no growth       |
//! | 4   | `wrapSquare`                     | 74.7 — the 60pt picture|
//! | 5   | `wrapTight`                      | 74.9 — the 60pt picture|
//! | 6   | `wrapTopAndBottom`               | 88.5 — picture + line  |
//! | 7   | `wrapNone` beside text           | the text's line only   |
//! | 8   | `wrapNone`, `layoutInCell="0"`   | 28.5 — no growth       |
//!
//! So a row is sized by what **displaces text**. A `wrapNone` picture
//! (§20.4.2.15) overlays the flow and takes no room from it, in a cell as on

use dxpdf::render::layout::draw_command::DrawCommand;
use dxpdf::render::layout::draw_command::LayoutedPage;

fn layout(name: &str) -> Vec<LayoutedPage> {
    let path = format!("{}/test-files/{name}", env!("CARGO_MANIFEST_DIR"));
    let bytes = std::fs::read(&path).expect("fixture is committed");
    let doc = dxpdf::docx::parse(&bytes).expect("fixture parses");
    dxpdf::render::resolve_and_layout(doc).1
}

/// Baseline position of the one text command reading `label` on `page`.
fn marker(page: &LayoutedPage, label: &str) -> (f32, f32) {
    let hits: Vec<(f32, f32)> = page
        .commands
        .iter()
        .filter_map(|c| match c {
            DrawCommand::Text { text, position, .. } if &**text == label => {
                Some((position.x.raw(), position.y.raw()))
            }
            _ => None,
        })
        .collect();
    assert_eq!(
        hits.len(),
        1,
        "exactly one {label:?} on the page, got {hits:?}"
    );
    hits[0]
}

const TOL: f32 = 0.05;

const PICTURE: f32 = 60.0;

/// Height of picture row `i` (1-based): the distance from `MARK{i-1}` to
/// `MARK{i}`, less the one-line marker row between them.
fn row_heights() -> (f32, Vec<f32>) {
    let pages = layout("anchor-in-cell.docx");
    let page = &pages[0];
    let ys: Vec<f32> = (0..=8)
        .map(|i| marker(page, &format!("MARK{i}")).1)
        .collect();
    // Row 1 is the control: an empty paragraph, so marker row and picture row
    // are one line each.
    let line = (ys[1] - ys[0]) / 2.0;
    let rows = (1..=8).map(|i| ys[i] - ys[i - 1] - line).collect();
    (line, rows)
}

#[test]
fn a_wrap_none_picture_does_not_grow_its_row() {
    let (line, rows) = row_heights();
    for (row, what) in [(2, "in front"), (3, "behind text"), (8, "layoutInCell=0")] {
        assert!(
            (rows[row - 1] - line).abs() < TOL,
            "row {row} (wrapNone, {what}) is one empty line ({line}) in Word, got {}",
            rows[row - 1]
        );
    }
}

#[test]
fn a_wrap_none_picture_beside_text_leaves_the_row_to_the_text() {
    let pages = layout("anchor-in-cell.docx");
    let page = &pages[0];
    let (line, _) = row_heights();
    let text = marker(page, "TEXT").1;
    let next = marker(page, "MARK7").1;
    assert!(
        (next - text - line).abs() < TOL,
        "MARK7 sits one line below TEXT ({line}), got {}",
        next - text
    );
}

/// The control for the rule above: modes that displace text still size the row.
///
/// A row's height here includes its share of the 0.5pt borders above and below
/// it, which the one-line rows carry too, so a picture row comes out at the
/// picture plus that share — Word's measures 60.4.
#[test]
fn a_wrapping_picture_still_grows_its_row_to_contain_it() {
    const BORDER_SHARE: f32 = 0.6;
    let (line, rows) = row_heights();
    for (row, what) in [(4, "wrapSquare"), (5, "wrapTight")] {
        assert!(
            rows[row - 1] >= PICTURE && rows[row - 1] - PICTURE < BORDER_SHARE,
            "row {row} ({what}) is the picture's {PICTURE}pt, got {}",
            rows[row - 1]
        );
    }
    assert!(
        (rows[5] - (PICTURE + line)).abs() < TOL,
        "row 6 (wrapTopAndBottom) is the picture plus the paragraph's line, got {}",
        rows[5]
    );
}
