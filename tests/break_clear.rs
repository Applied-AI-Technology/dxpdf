//! §17.3.3.1 line breaks against the paragraph mark and against floats, over
//! `test-files/break-clear.docx` (`scripts/make_break_clear_fixture.py`), measured in Word (Microsoft Word for Mac 16, PDF export, 2026-10-15) and asserted
//! as relations between markers on the same page, so no glyph metric is pinned.
//!
//! A paragraph that ends in an authored `<w:br/>` is a line taller than its content:
//! the break ends its line and the paragraph mark (§17.3.1.29) takes the next one.
//! Word draws case 8 — a paragraph holding only `<w:br/>` — two lines tall.

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

/// One line of the probe's body text: case 2's `AFTER` is the line after
/// `BEFORE`, beside the float, in the same face and size as every paragraph.
fn body_line(pages: &[LayoutedPage]) -> f32 {
    marker(&pages[1], "AFTER").1 - marker(&pages[1], "BEFORE").1
}

/// Case 8: a paragraph holding only a plain `<w:br/>`
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
