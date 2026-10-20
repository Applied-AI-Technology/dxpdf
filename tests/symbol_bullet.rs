//! A list label set in Symbol or Wingdings, over `test-files/symbol-bullet.docx`
//! (`scripts/make_symbol_bullet_fixture.py`): the line it leads stands as tall
//! as the face the level names makes it, whichever face draws the glyph.
//!
//! Word draws the label in its own copy of the face — Symbol, `hhea` 2059 up
//! and 450 down of 2048 units, and Wingdings, 1841 and 432 — and lays the line
//! out with it. A host without them (every Linux host, and a Mac, whose own
//! Symbol is a different face) draws the private-use character remapped to a
//! Unicode bullet in a text face instead. Measured in Word (Word for Mac 16,
//! PDF export) on real letter templates: a Calibri 11pt line at 1.15
//! with an 11pt Symbol bullet steps 16.08pt, Symbol's ascent over Calibri's
//! descent, where its unlabelled continuation steps 15.36pt.
//!
//! The label is not scaled with the text by Auto spacing. Word's render of
//! this fixture (Word for Mac 16, PDF export, 2026-10-16), from one reference
//! line to the next: R0→R1 62.88pt, R1→R2 42.48pt (a 22pt Symbol bullet: its
//! 22.12pt up over Calibri's 2.95pt down and 2.01pt extra, not 28.97pt or
//! 31.0pt), R2→R3 40.32pt (22pt Wingdings), R3→R4 30.96pt — each within the
//! 0.24pt grid Word places lines on of the rule asserted here, in Calibri.
//!
//! Laid out here in `DxNoGap` as "Calibri" (0.95em up, 0.25em down, no gap),
//! which carries no bullet at all, so the glyph comes from whatever the host
//! has and only the line metrics are asserted.

use dxpdf::model::{EmbeddedFont, EmbeddedFontVariant};
use dxpdf::render::layout::draw_command::{DrawCommand, LayoutedPage};

fn pages() -> Vec<LayoutedPage> {
    let dir = env!("CARGO_MANIFEST_DIR");
    let bytes = std::fs::read(format!("{dir}/test-files/symbol-bullet.docx")).expect("fixture");
    let mut doc = dxpdf::docx::parse(&bytes).expect("fixture parses");
    doc.embedded_fonts.push(EmbeddedFont {
        family: "Calibri".into(),
        variant: EmbeddedFontVariant::Regular,
        data: std::fs::read(format!("{dir}/test-files/fonts/DxNoGap.ttf")).expect("font"),
    });
    dxpdf::render::resolve_and_layout(doc).1
}

/// Baseline y of the one text command reading `label`.
fn at(pages: &[LayoutedPage], label: &str) -> f32 {
    let hits: Vec<f32> = pages[0]
        .commands
        .iter()
        .filter_map(|c| match c {
            DrawCommand::Text { text, position, .. } if text.trim_end() == label => {
                Some(position.y.raw())
            }
            _ => None,
        })
        .collect();
    assert_eq!(hits.len(), 1, "exactly one {label:?}, got {hits:?}");
    hits[0]
}

/// From one reference line's baseline to the next: the reference line's own
/// height plus every line between them.
fn assert_gap(pages: &[LayoutedPage], from: &str, to: &str, want: f32, why: &str) {
    let got = at(pages, to) - at(pages, from);
    assert!(
        (got - want).abs() < 0.01,
        "{from} → {to}: {why} — want {want}, got {got}"
    );
}

const SYMBOL: (f32, f32) = (2059.0 / 2048.0, 450.0 / 2048.0);
const WINGDINGS: (f32, f32) = (1841.0 / 2048.0, 432.0 / 2048.0);
const TEXT: (f32, f32) = (0.95, 0.25);

/// A 1.15 line holding 11pt text and a label of `face` at `size`: the label
/// reaches its own height, unscaled, and the text's line is scaled with its
/// extra below the text.
fn line(face: (f32, f32), size: f32) -> f32 {
    let up = (face.0 * size).max(TEXT.0 * 11.0);
    let down = (face.1 * size).max(TEXT.1 * 11.0 + 0.15 * (TEXT.0 + TEXT.1) * 11.0);
    up + down
}

fn plain() -> f32 {
    (TEXT.0 + TEXT.1) * 11.0 * 1.15
}

#[test]
fn an_11pt_symbol_bullet_raises_its_line_to_symbols_ascent() {
    // R0, S1a (labelled), S1b (not), S2 (labelled).
    let p = pages();
    let want = plain() + line(SYMBOL, 11.0) + plain() + line(SYMBOL, 11.0);
    assert_gap(&p, "R0", "R1", want, "two labelled lines and two plain");
}

#[test]
fn a_22pt_symbol_bullet_sets_its_line_in_symbol() {
    let p = pages();
    let want = plain() + line(SYMBOL, 22.0);
    assert_gap(&p, "R1", "R2", want, "Symbol's ascent and descent at 22pt");
}

#[test]
fn a_22pt_wingdings_bullet_sets_its_line_in_wingdings() {
    // Shorter than a text face of the same size would make it: Wingdings is
    // 1.11em tall, DxNoGap 1.2em.
    let p = pages();
    let want = plain() + line(WINGDINGS, 22.0);
    assert_gap(
        &p,
        "R2",
        "R3",
        want,
        "Wingdings' ascent and descent at 22pt",
    );
}

#[test]
fn an_11pt_wingdings_bullet_leaves_a_taller_text_line_alone() {
    let p = pages();
    assert_gap(
        &p,
        "R3",
        "R4",
        plain() * 2.0,
        "the text face is taller both ways",
    );
}
