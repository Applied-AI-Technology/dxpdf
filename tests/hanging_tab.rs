//! Where a tab on the first line of a hanging-indent paragraph stops, over
//! `test-files/hanging-tab-{15,compat,legacy,legacy-15}.docx`
//! (`scripts/make_hanging_tab_fixtures.py`), measured in Word (Word for Mac
//! 16, PDF export, 2026-10-17).
//!
//! A hanging indent is a tab stop at the paragraph's left indent, ordered
//! among the paragraph's own stops: one of its own before the indent comes
//! first (H2), one past it does not (H3), and a pen already past the indent
//! goes on to the default grid (H4). A first-line indent makes none (H5).
//! §17.15.3 `w:noTabHangInd` takes the stop away — in compatibility mode only;
//! under mode 15 Word ignores it.
//!
//! Laid out in `DxNoGap` as "Calibri", whose every advance is half an em; the
//! page's left margin is 72pt and the default stop every 36pt.

use dxpdf::model::{EmbeddedFont, EmbeddedFontVariant};
use dxpdf::render::layout::draw_command::{DrawCommand, LayoutedPage};

fn pages(name: &str) -> Vec<LayoutedPage> {
    let dir = env!("CARGO_MANIFEST_DIR");
    let bytes =
        std::fs::read(format!("{dir}/test-files/hanging-tab-{name}.docx")).expect("fixture");
    let mut doc = dxpdf::docx::parse(&bytes).expect("fixture parses");
    doc.embedded_fonts.push(EmbeddedFont {
        family: "Calibri".into(),
        variant: EmbeddedFontVariant::Regular,
        data: std::fs::read(format!("{dir}/test-files/fonts/DxNoGap.ttf")).expect("font"),
    });
    dxpdf::render::resolve_and_layout(doc).1
}

/// The pen x of the one text run reading `label`.
fn x(pages: &[LayoutedPage], label: &str) -> f32 {
    let hits: Vec<f32> = pages[0]
        .commands
        .iter()
        .filter_map(|c| match c {
            DrawCommand::Text { text, position, .. } if text.trim_end() == label => {
                Some(position.x.raw())
            }
            _ => None,
        })
        .collect();
    assert_eq!(hits.len(), 1, "exactly one {label:?}, got {hits:?}");
    hits[0]
}

fn assert_stops(name: &str, hanging_stop: bool) {
    let p = pages(name);
    let expect = |label: &str, at: f32, why: &str| {
        let got = x(&p, label);
        assert!(
            (got - at).abs() < 0.01,
            "{name} {label}: at {got}, expected {at} — {why}"
        );
    };
    if hanging_stop {
        expect("H1", 88.65, "the left indent, 333 twips");
        expect("H3", 144.0, "the left indent before its own later stop");
    } else {
        expect("H1", 108.0, "the first default stop");
        expect("H3", 180.0, "its own stop");
    }
    expect("H2", 126.0, "its own stop before the left indent");
    // 36 + 35 half-em advances = 211, past the indent: the default grid.
    expect("H4z", 288.0, "the default stop after the text");
    expect("H5", 180.0, "no hanging indent: the default grid");
    expect("H6", 144.0, "the left indent");
    expect("H6z", 180.0, "the default stop after it");
}

#[test]
fn a_hanging_indent_is_a_tab_stop() {
    assert_stops("15", true);
    assert_stops("compat", true);
}

#[test]
fn no_tab_hang_ind_removes_it_in_compatibility_mode() {
    assert_stops("legacy", false);
}

#[test]
fn mode_15_ignores_no_tab_hang_ind() {
    assert_stops("legacy-15", true);
}
