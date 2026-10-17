//! The paragraph spacing Word supplies when a styles part has no
//! `w:pPrDefault`, over `test-files/paragraph-defaults-{none,15,explicit}.docx`
//! (`scripts/make_paragraph_defaults_fixtures.py`), measured in Word (Word for
//! Mac 16, PDF export, 2026-10-17).
//!
//! §17.7.2 leaves an absent `pPrDefault` to the application. Word gives it 8pt
//! after and 1.15 lines — whether or not a `settings.xml` names compatibility
//! mode 15 — and a paragraph's own `w:spacing` replaces them attribute by
//! attribute, even written out of `CT_PPr`'s sequence. A `pPrDefault` stating
//! no space and single spacing (the control) gives exactly that.
//!
//! Laid out in `DxNoGap` as "Calibri": 12pt of text at 10pt, 13.8pt at 1.15.

use dxpdf::model::{EmbeddedFont, EmbeddedFontVariant};
use dxpdf::render::layout::draw_command::{DrawCommand, LayoutedPage};

fn pages(name: &str) -> Vec<LayoutedPage> {
    let dir = env!("CARGO_MANIFEST_DIR");
    let bytes =
        std::fs::read(format!("{dir}/test-files/paragraph-defaults-{name}.docx")).expect("fixture");
    let mut doc = dxpdf::docx::parse(&bytes).expect("fixture parses");
    doc.embedded_fonts.push(EmbeddedFont {
        family: "Calibri".into(),
        variant: EmbeddedFontVariant::Regular,
        data: std::fs::read(format!("{dir}/test-files/fonts/DxNoGap.ttf")).expect("font"),
    });
    dxpdf::render::resolve_and_layout(doc).1
}

/// The baseline of the one text run reading `label`.
fn baseline(pages: &[LayoutedPage], label: &str) -> f32 {
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

/// The distance from `<group>1`'s baseline to `<group>2`'s.
fn step(pages: &[LayoutedPage], group: &str) -> f32 {
    baseline(pages, &format!("{group}2")) - baseline(pages, &format!("{group}1"))
}

fn assert_step(pages: &[LayoutedPage], group: &str, expected: f32, why: &str) {
    let got = step(pages, group);
    assert!(
        (got - expected).abs() < 0.01,
        "{group}: step {got}, expected {expected} — {why}"
    );
}

/// Word: A 22.08, B and C 15.12, D 12.24, E 20.16 (Calibri, on its 0.24pt grid).
fn assert_word_defaults(name: &str) {
    let p = pages(name);
    assert_step(&p, "A", 13.8 + 8.0, "8pt after, 1.15 lines");
    assert_step(&p, "B", 13.8 + 1.0, "its own after, the default line");
    assert_step(&p, "C", 13.8 + 1.0, "spacing after jc is still read");
    assert_step(&p, "D", 12.0, "its own after and line");
    assert_step(&p, "E", 12.0 + 8.0, "its own line, the default after");
}

#[test]
fn no_paragraph_defaults_give_8pt_after_and_115_lines() {
    assert_word_defaults("none");
}

#[test]
fn a_compatibility_mode_does_not_change_them() {
    assert_word_defaults("15");
}

/// The wrapped paragraph's lines are 1.15 apart too (Word: 14.16).
#[test]
fn a_wrapped_paragraph_takes_the_default_line() {
    let p = pages("none");
    let f = baseline(&p, "F");
    let mut lines: Vec<f32> = p[0]
        .commands
        .iter()
        .filter_map(|c| match c {
            DrawCommand::Text { position, .. } => Some(position.y.raw()),
            _ => None,
        })
        .filter(|&y| y >= f && y < baseline(&p, "G1"))
        .collect();
    lines.sort_by(f32::total_cmp);
    lines.dedup_by(|a, b| (*a - *b).abs() < 0.01);
    assert!(lines.len() >= 2, "F wraps: {lines:?}");
    for pair in lines.windows(2) {
        assert!((pair[1] - pair[0] - 13.8).abs() < 0.01, "{lines:?}");
    }
}

#[test]
fn a_stated_pprdefault_is_kept() {
    let p = pages("explicit");
    for group in ["A", "B", "D", "E"] {
        let expected = if group == "B" { 13.0 } else { 12.0 };
        assert_step(&p, group, expected, "the pPrDefault's no space, single");
    }
}
