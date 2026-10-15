//! §17.4.83 `w:vAlign` against anchored pictures, over
//! `test-files/anchor-valign.docx` (`scripts/make_anchor_valign_fixture.py`),
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
/// Half the floating table's 0.5pt outer border. Where a table's horizontal
/// outer edge sits on its grid line is still open in this engine
/// (`table::borders` marks it as awaiting its own render), and it moves the
/// table's bottom — and so the cleared line — by up to this much.
const HALF_BORDER: f32 = 0.25;

/// Top edges of the page's anchored pictures, in reading order (row, then x).
fn picture_tops(page: &LayoutedPage) -> Vec<(f32, f32)> {
    let mut tops: Vec<(f32, f32)> = page
        .commands
        .iter()
        .filter_map(|c| match c {
            DrawCommand::Image { rect, .. } => Some((rect.origin.y.raw(), rect.origin.x.raw())),
            _ => None,
        })
        .collect();
    tops.sort_by(|a, b| (a.0.round(), a.1).partial_cmp(&(b.0.round(), b.1)).unwrap());
    tops
}

/// §17.4.83 `w:vAlign` against anchored pictures, measured in Word: a
/// paragraph-relative picture moves with its paragraph, and a `wrapNone` one
/// counts as content when the row's slack is shared out — so the labels of the
/// `wrapNone` row sit exactly where the `wrapSquare` row's do (Word: 34.8pt and
/// 69.6pt below the top-aligned label in both rows).
#[test]
fn a_cell_aligns_on_its_anchored_pictures_and_they_move_with_it() {
    let pages = layout("anchor-valign.docx");
    let page = &pages[0];
    let pictures = picture_tops(page);
    for (r, row) in [(1, &pictures[0..3]), (2, &pictures[3..6])] {
        let labels: Vec<f32> = ["T", "C", "B"]
            .iter()
            .map(|a| marker(page, &format!("R{r}{a}")).1)
            .collect();
        for i in 1..3 {
            assert!(
                ((row[i].0 - labels[i]) - (row[0].0 - labels[0])).abs() < TOL,
                "row {r}: each picture keeps its offset from its own paragraph; \
                 pictures {row:?}, labels {labels:?}"
            );
        }
    }
    for a in ["C", "B"] {
        let none = marker(page, &format!("R1{a}")).1 - marker(page, "R1T").1;
        let square = marker(page, &format!("R2{a}")).1 - marker(page, "R2T").1;
        // Row 1's top edge is the table's outer border, row 2's an interior
        // one, and the two are charged to the row differently — the open
        // question `HALF_BORDER` names — so the slack may differ by that much.
        assert!(
            (none - square).abs() < HALF_BORDER + TOL,
            "{a}: a wrapNone picture is aligned on like a wrapSquare one — \
             {none} against {square}"
        );
    }
}

/// Row 3 is the certificate's signature line: a bottom-aligned cell whose only
/// paragraph is empty but for a `wrapNone` picture anchored 40pt below it,
/// reaching past the row. Nothing is left to align, so the paragraph stays at
/// the cell's top and the picture lands 40pt below the row's top — Word: row
/// at 272.5, picture at 312.5.
#[test]
fn a_picture_reaching_past_its_row_leaves_its_paragraph_at_the_top() {
    let pages = layout("anchor-valign.docx");
    let page = &pages[0];
    let pictures = picture_tops(page);
    // Row 1's top-aligned picture sits at its paragraph's top; its label's
    // baseline is one ascent below that. Row 3's top-aligned label gives the
    // same paragraph top in row 3.
    let ascent = marker(page, "R1T").1 - pictures[0].0;
    let row3_top = marker(page, "R3T").1 - ascent;
    let signature = pictures[6].0;
    assert!(
        (signature - (row3_top + 40.0)).abs() < TOL,
        "the signature is 40pt below its row's top ({row3_top}); got {signature}"
    );
}
