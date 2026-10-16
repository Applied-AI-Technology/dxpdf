//! §17.3.2.19 `w:kern`, over `test-files/kerning.docx`
//! (`scripts/make_kerning_fixture.py`), measured in Word (Microsoft Word for Mac
//! 16, PDF export, 2026-10-15) in Calibri: every word of every line lands within
//! 0.04pt of Word once kerned from Calibri's `kern` table.
//!
//! Calibri is not on every host, so the document is laid out here in
//! `test-files/fonts/DxKern.ttf` (`scripts/make_font_fixtures.py`), handed to it
//! as an embedded "Calibri". Every advance in that font is 500 units of a
//! 1000-unit em, so a line's width is its length plus exactly the pairs it was
//! kerned by — and its `GPOS` kerns the same pairs at other values, so the
//! assertions below also say which of the two tables the pairs came from.

use dxpdf::model::{EmbeddedFont, EmbeddedFontVariant};
use dxpdf::render::layout::draw_command::{DrawCommand, LayoutedPage};

fn pages() -> Vec<LayoutedPage> {
    let dir = env!("CARGO_MANIFEST_DIR");
    let bytes = std::fs::read(format!("{dir}/test-files/kerning.docx")).expect("fixture");
    let mut doc = dxpdf::docx::parse(&bytes).expect("fixture parses");
    doc.embedded_fonts.push(EmbeddedFont {
        family: "Calibri".into(),
        variant: EmbeddedFontVariant::Regular,
        data: std::fs::read(format!("{dir}/test-files/fonts/DxKern.ttf")).expect("font"),
    });
    dxpdf::render::resolve_and_layout(doc).1
}

/// x of `word` on the line whose first text reads `label` (with its space).
fn x_on(pages: &[LayoutedPage], label: &str, word: &str) -> f32 {
    let texts: Vec<(f32, f32, &str)> = pages[0]
        .commands
        .iter()
        .filter_map(|c| match c {
            DrawCommand::Text { text, position, .. } => {
                Some((position.y.raw(), position.x.raw(), &**text))
            }
            _ => None,
        })
        .collect();
    let line_y = texts
        .iter()
        .find(|t| t.2 == format!("{label} ") || t.2.starts_with(&format!("{label} ")))
        .unwrap_or_else(|| panic!("no line {label}"))
        .0;
    texts
        .iter()
        .find(|t| (t.0 - line_y).abs() < 0.01 && t.2.trim_end() == word)
        .unwrap_or_else(|| panic!("no {word:?} on line {label}"))
        .1
}

/// The run of the line from `To` to `END`: every kerned pair of the line, and
/// none of the label, whose width differs from line to line.
fn span(pages: &[LayoutedPage], label: &str) -> f32 {
    x_on(pages, label, "END") - x_on(pages, label, "To")
}

const TOL: f32 = 0.01;

/// `DxKern`'s `kern` table over `To AVAWAY Tyre Wolf yo. LT`: To −90, AV/VA
/// −80 each, AW/WA −60 each, AY −100, Ty −70, yo −20, LT −110. Its `GPOS`
/// says AV/VA/To −200 and K1 +40 instead, which would make this −1010.
const LINE_PAIRS: f32 = -670.0;

fn pt(units: f32, size: f32) -> f32 {
    units * size / 1000.0
}

/// A kerned line is shorter than the control by the `kern` table's pairs
/// exactly. Word: K1's span is 5.74pt shorter than K2's in Calibri, the `kern`
/// table's figure; GPOS would have made it nearly 10pt.
#[test]
fn a_kerned_line_is_kerned_by_the_kern_table() {
    let p = pages();
    let shift = span(&p, "K1") - span(&p, "K2");
    assert!(
        (shift - pt(LINE_PAIRS, 11.0)).abs() < TOL,
        "want {}, got {shift}",
        pt(LINE_PAIRS, 11.0)
    );
}

/// `w:kern` is a threshold in half-points, inclusive, and zero is off. Word:
/// K3 (14pt threshold at 11pt) and K12 (`w:val="0"`) span what K2 does, and
/// K4b (14pt at 14pt) is 7.28pt shorter than K4c.
#[test]
fn the_threshold_is_inclusive_and_zero_is_off() {
    let p = pages();
    let k2 = span(&p, "K2");
    assert!(
        (span(&p, "K3") - k2).abs() < TOL,
        "11pt under a 14pt threshold"
    );
    assert!((span(&p, "K12") - k2).abs() < TOL, "w:kern 0");
    let at = span(&p, "K4b") - span(&p, "K4c");
    assert!(
        (at - pt(LINE_PAIRS, 14.0)).abs() < TOL,
        "14pt on a 14pt threshold is kerned: want {}, got {at}",
        pt(LINE_PAIRS, 14.0)
    );
}

/// The pair a run boundary falls between is kerned when both runs kern — `T` |
/// `owel` measures as `Towel` (T·o −90, o·w −10) — and not when the second run
/// does not. Word: `To` 28.39pt after `Towel`'s start in K5, 29.50pt in K5b,
/// Calibri's T·o pair apart.
#[test]
fn a_pair_across_a_run_boundary_is_kerned_when_both_runs_kern() {
    let p = pages();
    // K5's two runs are joined into one `Towel`; K5b's stay `T` and `owel`.
    let width = |label, first| x_on(&p, label, "To") - x_on(&p, label, first);
    let kerned = width("K5", "Towel");
    let plain = width("K5b", "T");
    assert!(
        (kerned - pt(6.0 * 500.0 - 100.0, 11.0)).abs() < TOL,
        "K5: Towel and its space kerned, got {kerned}"
    );
    assert!(
        (plain - pt(6.0 * 500.0, 11.0)).abs() < TOL,
        "K5b: nothing kerned, got {plain}"
    );
}

/// Kerning and §17.3.2.35 character spacing both apply. Word: K6's span is
/// 5.74pt shorter than K6b's, the same as K1's against K2's.
#[test]
fn kerning_and_character_spacing_add() {
    let p = pages();
    let shift = span(&p, "K6") - span(&p, "K6b");
    assert!(
        (shift - pt(LINE_PAIRS, 11.0)).abs() < TOL,
        "want {}, got {shift}",
        pt(LINE_PAIRS, 11.0)
    );
}

/// The painter draws a kerned run from the pairs layout measured, so the two
/// cannot disagree: the command carries them, one per spacing unit.
#[test]
fn the_draw_command_carries_the_pairs_it_was_measured_with() {
    let p = pages();
    let pairs = p[0]
        .commands
        .iter()
        .find_map(|c| match c {
            DrawCommand::Text { text, kerning, .. } if &**text == "AVAWAY " => kerning.clone(),
            _ => None,
        })
        .expect("a kerned AVAWAY");
    let expected = [-80.0, -80.0, -60.0, -60.0, -100.0, 0.0, 0.0].map(|u| pt(u, 11.0));
    assert_eq!(pairs.len(), expected.len());
    for (got, want) in pairs.iter().zip(expected) {
        assert!((got - want).abs() < 1e-4, "{pairs:?}");
    }
}
