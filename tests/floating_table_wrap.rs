//! §17.4.57 floating tables with `vertAnchor="text"`: where they stand, and how
//! close the text wrapping around them comes, over
//! `test-files/floating-table-wrap-{15,14}.docx`
//! (`scripts/make_floating_table_wrap_fixtures.py`), measured in Word
//! (Microsoft Word for Mac 16, PDF export, 2026-10-16).
//!
//! Every table's first row and the paragraph after it are set in one face, so
//! the distance between their baselines is the table's offset from that
//! paragraph's top plus its 0.5pt top border. Word's numbers are quoted per
//! case; they sit within 0.24pt of these, Word's vertical grain in its PDFs.

use dxpdf::render::layout::draw_command::{DrawCommand, LayoutedPage};

fn pages(mode: u32) -> Vec<LayoutedPage> {
    let path = format!(
        "{}/test-files/floating-table-wrap-{mode}.docx",
        env!("CARGO_MANIFEST_DIR")
    );
    let bytes = std::fs::read(&path).expect("fixture is committed");
    let doc = dxpdf::docx::parse(&bytes).expect("fixture parses");
    dxpdf::render::resolve_and_layout(doc).1
}

/// (page, x, baseline y) of the one text command reading `label`.
fn at(pages: &[LayoutedPage], label: &str) -> (usize, f32, f32) {
    let hits: Vec<(usize, f32, f32)> = pages
        .iter()
        .enumerate()
        .flat_map(|(i, p)| p.commands.iter().map(move |c| (i, c)))
        .filter_map(|(i, c)| match c {
            DrawCommand::Text { text, position, .. } if text.trim_end() == label => {
                Some((i, position.x.raw(), position.y.raw()))
            }
            _ => None,
        })
        .collect();
    assert_eq!(hits.len(), 1, "exactly one {label:?}, got {hits:?}");
    hits[0]
}

const TOP_BORDER: f32 = 0.5;

/// The table stands `tblpY` below the top of the paragraph that follows it —
/// the cursor — however long the paragraph before it is. Word: the first row
/// 0.73 below the following line at 1 twip after one line (A1) and three (A2),
/// 20.65 at 20pt after one (A3) and 20.57 after three (A4).
#[test]
fn a_text_anchored_table_stands_tblpy_below_the_following_paragraph() {
    for mode in [15, 14] {
        let p = pages(mode);
        for (case, tblp_y) in [("A1", 0.05), ("A2", 0.05), ("A3", 20.0), ("A4", 20.0)] {
            let (pr, _, row) = at(&p, &format!("{case}r1"));
            let (pp, _, post) = at(&p, &format!("{case}post"));
            assert_eq!(pr, pp, "{case}: one page");
            let got = row - post;
            assert!(
                (got - (tblp_y + TOP_BORDER)).abs() < 0.01,
                "{case} mode {mode}: want {}, got {got}",
                tblp_y + TOP_BORDER
            );
        }
    }
}
