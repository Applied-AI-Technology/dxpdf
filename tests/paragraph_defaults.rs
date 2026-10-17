//! The paragraph spacing Word supplies when a styles part has no
//! `w:pPrDefault`, over `test-files/paragraph-defaults-{none,15,explicit}.docx`
//! (`scripts/make_paragraph_defaults_fixtures.py`) and `test-files/line-default.docx`
//! (`scripts/make_line_default_fixture.py`), measured in Word (Word for Mac 16,
//! PDF export, 2026-10-17).
//!
//! §17.7.2 leaves an absent `pPrDefault` to the application. Word gives it 8pt
//! after and lines of 278/240 — not 1.15, which a stated `w:line="276"` in the
//! same document draws distinctly tighter — whether or not a `settings.xml`
//! names compatibility mode 15, and a paragraph's own `w:spacing` replaces them
//! attribute by attribute, even written out of `CT_PPr`'s sequence. A
//! `pPrDefault` stating no space and single spacing (the control) gives exactly
//! that.
//!
//! Laid out in `DxNoGap` as "Calibri": 12pt of text at 10pt, 13.9pt at 278/240.

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

/// The default line in `DxNoGap` at 10pt: 12pt of text at 278/240.
const DEFAULT_LINE: f32 = 12.0 * 278.0 / 240.0;

/// Word: A 22.08, B and C 15.12, D 12.24, E 20.16 (Calibri, its positions
/// rounded to 0.24pt).
fn assert_word_defaults(name: &str) {
    let p = pages(name);
    assert_step(&p, "A", DEFAULT_LINE + 8.0, "8pt after, the default line");
    assert_step(
        &p,
        "B",
        DEFAULT_LINE + 1.0,
        "its own after, the default line",
    );
    assert_step(
        &p,
        "C",
        DEFAULT_LINE + 1.0,
        "spacing after jc is still read",
    );
    assert_step(&p, "D", 12.0, "its own after and line");
    assert_step(&p, "E", 12.0 + 8.0, "its own line, the default after");
}

#[test]
fn no_paragraph_defaults_give_8pt_after_and_278_240_lines() {
    assert_word_defaults("none");
}

#[test]
fn a_compatibility_mode_does_not_change_them() {
    assert_word_defaults("15");
}

/// The wrapped paragraph's lines take the default line too (Word: 14.16).
#[test]
fn a_wrapped_paragraph_takes_the_default_line() {
    let p = pages("none");
    let f = baseline(&p, "wraps");
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
        assert!((pair[1] - pair[0] - DEFAULT_LINE).abs() < 0.01, "{lines:?}");
    }
}

/// The line pitch of the paragraph opening with `label` in `line-default.docx`:
/// the step between the baselines of its first and last lines on the page,
/// over the lines between.
fn line_pitch(pages: &[LayoutedPage], label: &str) -> f32 {
    let (page, top) = pages
        .iter()
        .find_map(|page| {
            page.commands.iter().find_map(|c| match c {
                DrawCommand::Text { text, position, .. } if text.trim_end() == label => {
                    Some((page, position.y.raw()))
                }
                _ => None,
            })
        })
        .unwrap_or_else(|| panic!("no {label:?}"));
    // Its lines are the baselines from its own down to the next label's.
    let mut ys: Vec<f32> = page
        .commands
        .iter()
        .filter_map(|c| match c {
            DrawCommand::Text { text, position, .. } => Some((text.trim_end(), position.y.raw())),
            _ => None,
        })
        .filter(|&(_, y)| y >= top)
        .map(|(text, y)| (text != "word" && y > top, y))
        .scan(false, |ended, (next_label, y)| {
            *ended |= next_label;
            Some((*ended, y))
        })
        .take_while(|&(ended, _)| !ended)
        .map(|(_, y)| y)
        .collect();
    ys.sort_by(f32::total_cmp);
    ys.dedup_by(|a, b| (*a - *b).abs() < 0.01);
    assert!(ys.len() >= 4, "{label}: {ys:?}");
    (ys[ys.len() - 1] - ys[0]) / (ys.len() - 1) as f32
}

/// Word draws a paragraph with no `w:pPr` under no `pPrDefault` exactly as one
/// stating `w:line="278"` (P and R, within 0.03pt at 8, 10, 12 and 20pt in
/// Calibri, and at 10 and 20pt in Arial), and both further apart than one
/// stating 276 (Q), which dxpdf matched already.
#[test]
fn the_absent_default_line_is_278_240_not_276() {
    let p = {
        let dir = env!("CARGO_MANIFEST_DIR");
        let bytes = std::fs::read(format!("{dir}/test-files/line-default.docx")).expect("fixture");
        let mut doc = dxpdf::docx::parse(&bytes).expect("fixture parses");
        doc.embedded_fonts.push(EmbeddedFont {
            family: "Calibri".into(),
            variant: EmbeddedFontVariant::Regular,
            data: std::fs::read(format!("{dir}/test-files/fonts/DxNoGap.ttf")).expect("font"),
        });
        dxpdf::render::resolve_and_layout(doc).1
    };
    for size in [8.0f32, 10.0, 12.0, 20.0] {
        let text = 1.2 * size;
        let check = |label: String, multiple: f32| {
            let got = line_pitch(&p, &label);
            let expected = text * multiple / 240.0;
            assert!(
                (got - expected).abs() < 0.01,
                "{label}: {got}, expected {expected}"
            );
        };
        check(format!("P{size}"), 278.0);
        check(format!("Q{size}"), 276.0);
        check(format!("R{size}"), 278.0);
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
