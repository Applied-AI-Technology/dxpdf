//! §17.3.2.5 `w:caps` and §17.3.2.33 `w:smallCaps`, over the G lines of
//! `test-files/paragraph-defaults-none.docx`
//! (`scripts/make_paragraph_defaults_fixtures.py`), measured in Word (Word for
//! Mac 16, PDF export, 2026-10-17).
//!
//! Word draws `w:caps` text in capitals, and extracts it so. `w:smallCaps`
//! draws a lowercase letter as its capital at 0.8 of the run's size — 8pt at
//! 10pt, 16pt at 20pt — and every space of the run at that size too, while
//! capitals and digits keep the full size. Word's pen positions for both lines
//! match these to 0.01pt in Calibri.
//!
//! Laid out in `DxNoGap` as "Calibri", whose every advance is half an em.

use dxpdf::model::{EmbeddedFont, EmbeddedFontVariant};
use dxpdf::render::layout::draw_command::{DrawCommand, LayoutedPage};

fn pages() -> Vec<LayoutedPage> {
    let dir = env!("CARGO_MANIFEST_DIR");
    let bytes =
        std::fs::read(format!("{dir}/test-files/paragraph-defaults-none.docx")).expect("fixture");
    let mut doc = dxpdf::docx::parse(&bytes).expect("fixture parses");
    doc.embedded_fonts.push(EmbeddedFont {
        family: "Calibri".into(),
        variant: EmbeddedFontVariant::Regular,
        data: std::fs::read(format!("{dir}/test-files/fonts/DxNoGap.ttf")).expect("font"),
    });
    dxpdf::render::resolve_and_layout(doc).1
}

/// The text, size and pen x of every text command on the line `label`
/// starts.
fn line(pages: &[LayoutedPage], label: &str) -> Vec<(String, f32, f32)> {
    let texts: Vec<(String, f32, f32, f32)> = pages[0]
        .commands
        .iter()
        .filter_map(|c| match c {
            DrawCommand::Text {
                text,
                position,
                font_size,
                ..
            } => Some((
                text.to_string(),
                font_size.raw(),
                position.x.raw(),
                position.y.raw(),
            )),
            _ => None,
        })
        .collect();
    let y = texts
        .iter()
        .find(|t| t.0.starts_with(label))
        .unwrap_or_else(|| panic!("no line starts {label:?}"))
        .3;
    texts
        .into_iter()
        .filter(|t| t.3 == y)
        .map(|(text, size, x, _)| (text, size, x))
        .collect()
}

fn assert_line(pages: &[LayoutedPage], label: &str, expected: &[(&str, f32, f32)]) {
    let got = line(pages, label);
    assert_eq!(got.len(), expected.len(), "{label}: {got:?}");
    for ((text, size, x), (e_text, e_size, e_x)) in got.iter().zip(expected) {
        assert_eq!(text.trim_end(), *e_text, "{label}: {got:?}");
        assert!(
            (size - e_size).abs() < 0.01 && (x - e_x).abs() < 0.01,
            "{label}: {text:?} at {size}pt x {x}, expected {e_size}pt x {e_x} — {got:?}"
        );
    }
}

#[test]
fn caps_draw_the_run_in_capitals() {
    let p = pages();
    let text: String = line(&p, "G2").into_iter().map(|t| t.0).collect();
    assert_eq!(text.trim_end(), "G2 RECORD OF BIRTH");
}

#[test]
fn small_caps_draw_lowercase_and_spaces_at_eight_tenths() {
    let p = pages();
    // 10pt: a full-size advance is 5, a reduced one 4.
    assert_line(
        &p,
        "G3",
        &[
            ("G3", 10.0, 72.0),
            ("", 8.0, 82.0),
            ("R", 10.0, 86.0),
            ("ECORD", 8.0, 91.0),
            ("O", 10.0, 115.0),
            ("F", 8.0, 120.0),
            ("B", 10.0, 128.0),
            ("IRTH", 8.0, 133.0),
        ],
    );
}
