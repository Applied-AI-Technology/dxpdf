//! Whether white space stands a line taller, over
//! `test-files/space-run-height.docx` (`scripts/make_space_run_height_fixture.py`),
//! measured in Word (Word for Mac 16, PDF export, 2026-10-18).
//!
//! It does not. 10pt Times New Roman text at 1.15 steps 13.20pt a line in
//! Word, and keeps that step with a 20pt space run or an unformatted (Calibri
//! 11pt) one between two words (S1, S2), at the start of the line (S8), at the
//! end of a paragraph (S6), as a tab (S9) and on the last of several lines
//! (S10), and at single spacing (T1, T2). A visible `x` in Calibri (S3) and a
//! no-break space (S5) grow the line by 1.2pt. A real notice template's stray
//! space is S1.
//!
//! A paragraph holding only a 20pt space (S7) is as tall in Word as its 10pt
//! mark makes it, which this engine does not yet do: it keeps the space's
//! line when nothing else is on it (`layout::line`).

use dxpdf::render::layout::draw_command::{DrawCommand, LayoutedPage};

fn page() -> LayoutedPage {
    let path = format!(
        "{}/test-files/space-run-height.docx",
        env!("CARGO_MANIFEST_DIR")
    );
    let bytes = std::fs::read(&path).expect("fixture is committed");
    let doc = dxpdf::docx::parse(&bytes).expect("fixture parses");
    dxpdf::render::resolve_and_layout(doc).1.remove(0)
}

/// The baseline of the one text command reading `label`.
fn baseline(page: &LayoutedPage, label: &str) -> f32 {
    let hits: Vec<f32> = page
        .commands
        .iter()
        .filter_map(|c| match c {
            DrawCommand::Text { text, position, .. } if text.trim() == label => {
                Some(position.y.raw())
            }
            _ => None,
        })
        .collect();
    assert_eq!(hits.len(), 1, "exactly one {label:?}, got {hits:?}");
    hits[0]
}

fn step(page: &LayoutedPage, from: &str, to: &str) -> f32 {
    baseline(page, to) - baseline(page, from)
}

#[test]
fn white_space_in_a_larger_face_leaves_the_line_as_tall_as_its_text() {
    let p = page();
    let line = step(&p, "REF", "S0a");
    for (from, to, lines) in [
        ("S0a", "S1a", 1.0),
        ("S1a", "S2a", 1.0),
        ("S7z", "S8a", 1.0),
        ("S8a", "S9a", 1.0),
        ("S10a", "S10b", 2.0),
        ("S10b", "S10z", 1.0),
        ("S10z", "S11a", 1.0),
        ("S11a", "T1a", 1.0),
    ] {
        let got = step(&p, from, to);
        assert!(
            (got - line * lines).abs() < 0.05,
            "{from} to {to}: want {} lines of {line}, got {got}",
            lines
        );
    }
}

/// The controls: what is not white space does stand the line taller.
#[test]
fn a_visible_character_or_a_no_break_space_grows_the_line() {
    let p = page();
    let line = step(&p, "REF", "S0a");
    assert!(step(&p, "S2a", "S3a") > line + 1.0, "an unformatted x");
    assert!(
        step(&p, "S4a", "S5a") > line + 1.0,
        "an unformatted no-break space"
    );
}

/// At single spacing as at 1.15: a 20pt space beside 10pt text is a 10pt
/// line, 11.52pt in Word.
#[test]
fn single_spacing_ignores_white_space_too() {
    let p = page();
    let got = step(&p, "T1a", "T2a");
    assert!((got - 11.5).abs() < 0.05, "want 11.5, got {got}");
}
