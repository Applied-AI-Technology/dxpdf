//! Where Word breaks a word too wide for its line, over
//! `test-files/oversized-word.docx` (`scripts/make_oversized_word_fixture.py`),
//! measured in Word (Word for Mac 16, PDF export, 2026-10-16) in Calibri.
//!
//! Word does not break after a solidus. `abcdefgh/ijklmnop/qrstuvwx` in a 72pt
//! cell is one word to it, cut where the line runs out — `abcdefgh/ijklm` /
//! `nop/qrstuvwx` — and `abcd/` is followed on its line by as much of a long
//! word as fits. A space and a hyphen are opportunities: after either, a word
//! too long for any line starts the next one and is cut there.
//!
//! Laid out in `DxNoGap` as "Calibri", whose every advance is half an em, so
//! 13 characters fill the cell's 72pt at 11pt and each cut point is known.

use dxpdf::model::{EmbeddedFont, EmbeddedFontVariant};
use dxpdf::render::layout::draw_command::{DrawCommand, LayoutedPage};

fn pages() -> Vec<LayoutedPage> {
    let dir = env!("CARGO_MANIFEST_DIR");
    let bytes = std::fs::read(format!("{dir}/test-files/oversized-word.docx")).expect("fixture");
    let mut doc = dxpdf::docx::parse(&bytes).expect("fixture parses");
    doc.embedded_fonts.push(EmbeddedFont {
        family: "Calibri".into(),
        variant: EmbeddedFontVariant::Regular,
        data: std::fs::read(format!("{dir}/test-files/fonts/DxNoGap.ttf")).expect("font"),
    });
    dxpdf::render::resolve_and_layout(doc).1
}

/// The lines of the cell under `SEP<case>`, each its text in reading order.
fn lines(pages: &[LayoutedPage], case: &str) -> Vec<String> {
    let mut texts: Vec<(f32, f32, String)> = pages[0]
        .commands
        .iter()
        .filter_map(|c| match c {
            DrawCommand::Text { text, position, .. } => {
                Some((position.y.raw(), position.x.raw(), text.to_string()))
            }
            _ => None,
        })
        .collect();
    texts.sort_by(|a, b| a.0.total_cmp(&b.0).then(a.1.total_cmp(&b.1)));
    let sep = |label: &str| {
        texts
            .iter()
            .find(|t| t.2.trim_end() == label)
            .unwrap_or_else(|| panic!("no {label}"))
            .0
    };
    let top = sep(&format!("SEP{case}"));
    let next = texts
        .iter()
        .filter(|t| t.0 > top + 0.01 && (t.2.starts_with("SEP") || t.2 == "END"))
        .map(|t| t.0)
        .fold(f32::MAX, f32::min);
    let mut out: Vec<(f32, String)> = Vec::new();
    for (y, _, text) in texts.iter().filter(|t| t.0 > top + 0.01 && t.0 < next) {
        match out.last_mut() {
            Some((ly, s)) if (*ly - y).abs() < 0.01 => s.push_str(text),
            _ => out.push((*y, text.clone())),
        }
    }
    out.into_iter()
        .map(|(_, s)| s.trim_end().to_string())
        .collect()
}

#[test]
fn a_word_is_not_broken_after_a_solidus() {
    let p = pages();
    assert_eq!(lines(&p, "O7"), ["abcdefgh/ijkl", "mnop/qrstuvwx"]);
}

#[test]
fn a_long_word_after_a_solidus_fills_the_line() {
    let p = pages();
    assert_eq!(
        lines(&p, "O1"),
        ["abcd/ABCDEFGH", "IJKLMNOPQRSTU", "VWXYZabcd"]
    );
}

#[test]
fn a_long_word_after_a_space_or_hyphen_starts_a_line() {
    let p = pages();
    assert_eq!(
        lines(&p, "O4"),
        ["Word", "ABCDEFGHIJKLM", "NOPQRSTUVWXYZ", "abcd"]
    );
    assert_eq!(
        lines(&p, "O5"),
        ["abcd-", "ABCDEFGHIJKLM", "NOPQRSTUVWXYZ", "abcd"]
    );
}
