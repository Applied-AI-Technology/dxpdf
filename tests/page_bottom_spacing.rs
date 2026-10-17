//! Whether a line stays on a page whose bottom margin its text clears but its
//! line spacing does not, over `test-files/page-bottom-spacing.docx`
//! (`scripts/make_page_bottom_spacing_fixture.py`), measured in Word (Word for
//! Mac 16, PDF export, 2026-10-16).
//!
//! An Auto multiple's extra hangs below the text, and Word lets it hang past
//! the bottom margin: at 1.15, 1.5 and double a line whose text fits stays,
//! the last line of a split paragraph too, and the line after it starts the
//! next page. A line whose text does not fit moves, and so does an `exact` or
//! `atLeast` line whose box does not fit — those have no extra to hang.
//!
//! Laid out in `DxNoGap` as "Calibri" (13.2pt of text at 11pt, 15.18pt at
//! 1.15); the fixture's margins separate every case in it as in Calibri.

use dxpdf::model::{EmbeddedFont, EmbeddedFontVariant};
use dxpdf::render::layout::draw_command::{DrawCommand, LayoutedPage};

fn pages() -> Vec<LayoutedPage> {
    let dir = env!("CARGO_MANIFEST_DIR");
    let bytes =
        std::fs::read(format!("{dir}/test-files/page-bottom-spacing.docx")).expect("fixture");
    let mut doc = dxpdf::docx::parse(&bytes).expect("fixture parses");
    doc.embedded_fonts.push(EmbeddedFont {
        family: "Calibri".into(),
        variant: EmbeddedFontVariant::Regular,
        data: std::fs::read(format!("{dir}/test-files/fonts/DxNoGap.ttf")).expect("font"),
    });
    dxpdf::render::resolve_and_layout(doc).1
}

/// The page `label` is drawn on.
fn page_of(pages: &[LayoutedPage], label: &str) -> usize {
    let hits: Vec<usize> = pages
        .iter()
        .enumerate()
        .flat_map(|(i, p)| p.commands.iter().map(move |c| (i, c)))
        .filter_map(|(i, c)| match c {
            DrawCommand::Text { text, .. } if text.trim_end() == label => Some(i),
            _ => None,
        })
        .collect();
    assert_eq!(hits.len(), 1, "exactly one {label:?}, got {hits:?}");
    hits[0]
}

/// `case` is on the page its spacer filled (`stays`) or the next one; the
/// spacer is the page's only other text, so its page is the one before
/// `<case>z` when the case stays.
fn assert_case(pages: &[LayoutedPage], case: &str, stays: bool) {
    let line = page_of(pages, case);
    let after = page_of(pages, &format!("{case}z"));
    if stays {
        assert_eq!(
            after,
            line + 1,
            "{case} stays and {case}z starts the next page"
        );
    } else {
        assert_eq!(after, line, "{case} moves to the page {case}z is on");
        assert!(line > 0, "{case} has a page before it");
    }
}

#[test]
fn an_auto_lines_extra_hangs_past_the_bottom_margin() {
    let p = pages();
    assert_case(&p, "A115", true);
    assert_case(&p, "A15", true);
    assert_case(&p, "A2", true);
}

#[test]
fn a_line_whose_text_does_not_fit_moves() {
    let p = pages();
    assert_case(&p, "A115n", false);
    assert_case(&p, "S1", false);
    assert_case(&p, "S1y", true);
}

#[test]
fn exact_and_at_least_lines_have_nothing_to_hang() {
    let p = pages();
    assert_case(&p, "X20", false);
    assert_case(&p, "L20", false);
}

#[test]
fn the_last_line_of_a_split_paragraph_hangs_too() {
    let p = pages();
    assert_eq!(page_of(&p, "AWa"), page_of(&p, "AWc"), "AWc stays with AWa");
    assert_eq!(page_of(&p, "AWz"), page_of(&p, "AWc") + 1);
}
