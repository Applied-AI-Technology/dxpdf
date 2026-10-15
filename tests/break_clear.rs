//! §17.3.3.1 line breaks against the paragraph mark and against floats, over
//! `test-files/break-clear.docx` (`scripts/make_break_clear_fixture.py`), measured in Word (Microsoft Word for Mac 16, PDF export, 2026-10-15) and asserted
//! as relations between markers on the same page, so no glyph metric is pinned.
//!
//! A paragraph that ends in an authored `<w:br/>` is a line taller than its content:
//! the break ends its line and the paragraph mark (§17.3.1.29) takes the next one.
//! Word draws case 8 — a paragraph holding only `<w:br/>` — two lines tall.
//!
//! **`break-clear.docx`** — §17.3.3.1 `w:br/@w:clear` (§17.18.3 `ST_BrClear`)
//! beside a 120pt float. Word restarts the text after a clearing break on the
//! first line below every float on the named side, and only those:
//! `all` clears a float on either side, `left`/`right` the float on that side,
//! and the opposite side's float is left where it is. Case 2 (`none`) is the
//! control. The floating table's rows are `exact` 40pt, so a line that clears
//! it starts one row below the top of the `TBLEND` row — Word draws `AFTER`
//! 40.0pt below `TBLEND`, to the tenth.

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

const FLOAT_ROW: f32 = 40.0;
/// Half the floating table's 0.5pt outer border. Where a table's horizontal
/// outer edge sits on its grid line is still open in this engine
/// (`table::borders` marks it as awaiting its own render), and it moves the
/// table's bottom — and so the cleared line — by up to this much.
const HALF_BORDER: f32 = 0.25;
const MARGIN_LEFT: f32 = 72.0;

/// (page index, whether Word moves `AFTER` below the float)
const BESIDE_A_TABLE: [(usize, &str, bool); 6] = [
    (0, "left table, clear=all", true),
    (1, "left table, clear=none", false),
    (2, "left table, clear=left", true),
    (3, "left table, clear=right", false),
    (4, "right table, clear=right", true),
    (5, "right table, clear=left", false),
];

#[test]
fn a_clearing_break_restarts_below_the_float_on_its_side() {
    let pages = layout("break-clear.docx");
    for (page, what, clears) in BESIDE_A_TABLE {
        let page = &pages[page];
        let tbl_end = marker(page, "TBLEND").1;
        let before = marker(page, "BEFORE");
        let after = marker(page, "AFTER");
        if clears {
            assert!(
                (after.1 - (tbl_end + FLOAT_ROW)).abs() < HALF_BORDER + TOL,
                "{what}: AFTER starts on the line below the table, one {FLOAT_ROW}pt row \
                 under TBLEND ({tbl_end}); got {}",
                after.1
            );
            assert!(
                (after.0 - MARGIN_LEFT).abs() < TOL,
                "{what}: below the float the line is full width again, got x={}",
                after.0
            );
        } else {
            let line = after.1 - before.1;
            assert!(
                line > 0.0 && line < FLOAT_ROW && (after.0 - before.0).abs() < TOL,
                "{what}: AFTER stays beside the float, the line after BEFORE; \
                 BEFORE {before:?}, AFTER {after:?}"
            );
        }
    }
}

/// Case 9: the float is a `wrapSquare` picture rather than a table.
#[test]
fn a_clearing_break_clears_a_wrapping_picture() {
    let pages = layout("break-clear.docx");
    let page = &pages[8];
    let picture = page
        .commands
        .iter()
        .find_map(|c| match c {
            DrawCommand::Image { rect, .. } => Some(*rect),
            _ => None,
        })
        .expect("the picture is drawn");
    let before = marker(page, "BEFORE");
    let after = marker(page, "AFTER");
    // BEFORE's line and the picture both start at the paragraph's top, so
    // AFTER's line starting at the picture's bottom puts its baseline the same
    // distance below that bottom as BEFORE's is below the top.
    let expected =
        picture.origin.y.raw() + picture.size.height.raw() + (before.1 - picture.origin.y.raw());
    assert!(
        (after.1 - expected).abs() < TOL && (after.0 - MARGIN_LEFT).abs() < TOL,
        "AFTER restarts at the margin on the line below the picture \
         (baseline {expected}); got {after:?}"
    );
}

/// One line of the probe's body text: case 2's `AFTER` is the line after
/// `BEFORE`, beside the float, in the same face and size as every paragraph.
fn body_line(pages: &[LayoutedPage]) -> f32 {
    marker(&pages[1], "AFTER").1 - marker(&pages[1], "BEFORE").1
}

/// Case 7, the shape real report templates use after a floating table: a
/// paragraph holding only `<w:br w:clear="all"/>`. Its break line sits beside
/// the float, the clear moves the paragraph's mark to the line below the
/// float, and what follows starts one line under that — Word draws `NEXT`
/// 53.7pt below `TBLEND`: the 40pt row, then the mark's line.
#[test]
fn a_paragraph_of_only_a_clearing_break_moves_what_follows_below_the_float() {
    let pages = layout("break-clear.docx");
    let line = body_line(&pages);
    let page = &pages[6];
    let tbl_end = marker(page, "TBLEND").1;
    let next = marker(page, "NEXT");
    let expected = tbl_end + FLOAT_ROW + line;
    assert!(
        (next.1 - expected).abs() < HALF_BORDER + TOL && (next.0 - MARGIN_LEFT).abs() < TOL,
        "NEXT starts at the margin one mark line below the table ({expected}); got {next:?}"
    );
}

/// Case 8, the control for case 7: a paragraph holding only a plain `<w:br/>`
/// beside the float is **two** lines — the break's, and the mark's after it —
/// and so is `text<w:br/>¶` anywhere. Word draws `NEXT` three body lines below
/// `CASE8` (41.3pt at 13.77pt a line).
#[test]
fn a_paragraph_ending_in_a_break_gives_its_mark_a_line_of_its_own() {
    let pages = layout("break-clear.docx");
    let line = body_line(&pages);
    let page = &pages[7];
    let label = marker(page, "CASE8").1;
    let next = marker(page, "NEXT").1;
    assert!(
        (next - label - 3.0 * line).abs() < TOL,
        "CASE8, the break's line, the mark's line: NEXT is {} below CASE8, \
         three lines of {line}",
        next - label
    );
}
