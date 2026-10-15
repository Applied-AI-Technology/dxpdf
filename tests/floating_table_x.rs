//! §17.4.57 horizontal placement of a floating table, over
//! `test-files/floating-table-x.docx` (`scripts/make_floating_table_x_fixture.py`),
//! measured in Word (Microsoft Word for Mac 16, PDF export, 2026-10-15) and asserted
//! as relations between markers on the same page, so no glyph metric is pinned.

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

/// §17.4.57 horizontal placement of a floating table: Word's x for each case's
/// cell text. A 2in table, 0.075in (5.4pt) default cell margins, on a 612pt
/// page with 72pt margins.
///
/// What it settles: `tblpX` and the left/right `tblpXSpec` place the first
/// (or last) cell's **text** on the position, not the table's edge — the same
/// rule `w:tblInd` follows for a table in the flow — so the table's edge sits
/// one cell margin outside it, and in case 10 a 0.25in margin moves the edge
/// and not the text. `center` centres the box. With no `horzAnchor` the
/// position is from the text column, `margin` from the page margin and `page`
/// from the paper edge. And §17.4.28 `w:jc` is not a flow-only property: a
/// `center` or `right` one places the float as that `tblpXSpec` would, over
/// `tblpX` (cases 12, 13) and over a contrary `tblpXSpec` (case 14).
#[test]
fn a_floating_table_is_placed_by_its_cell_text() {
    let pages = layout("floating-table-x.docx");
    let expected: [(&str, f32); 14] = [
        ("no horzAnchor, no tblpX", 72.0),
        ("text, tblpX 1in", 144.0),
        ("margin, tblpX 0", 72.0),
        ("margin, tblpX 1in", 144.0),
        ("page, tblpX 0", 0.0),
        ("page, tblpX 1.5in", 108.0),
        ("margin, tblpXSpec left", 72.0),
        // (468 − 144) / 2 + 72 for the box, plus the cell margin.
        ("margin, tblpXSpec center", 239.4),
        // 540 + 5.4 − 144 for the box, plus the cell margin.
        ("margin, tblpXSpec right", 406.8),
        ("margin, tblpX 1in, 0.25in cell margin", 144.0),
        ("no horzAnchor, w:jc center", 239.4),
        ("margin, tblpX 1in, w:jc center", 239.4),
        ("margin, tblpX 1in, w:jc right", 406.8),
        ("margin, tblpXSpec left, w:jc center", 239.4),
    ];
    for (i, (what, x)) in expected.into_iter().enumerate() {
        let got = marker(&pages[i], &format!("X{}", i + 1)).0;
        assert!(
            (got - x).abs() < TOL,
            "case {} ({what}): cell text at x={x} in Word, got {got}",
            i + 1
        );
    }
}
