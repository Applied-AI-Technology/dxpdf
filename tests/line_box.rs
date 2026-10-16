//! Where a line's baseline sits in its box, and how tall the box is, over
//! `test-files/line-box.docx` and `test-files/line-box-inline.docx`
//! (`scripts/make_line_box_fixtures.py`), measured in Word (Microsoft Word for
//! Mac 16, PDF export, 2026-10-16) in Calibri, Arial and Times New Roman.
//!
//! Those faces are not on every host, so the documents are laid out here in
//! two that differ only in their vertical metrics (`scripts/make_font_fixtures.py`):
//! `DxNoGap` as "Calibri" — 0.95em up, 0.25em down, no line gap, as Calibri has
//! none — and `DxLineGap` as "Arial" and "Times New Roman" — 0.9em up, 0.2em
//! down and a 0.1em gap, as both have one (0.0327em and 0.0425em). Word's own
//! numbers, in the real faces, are quoted per case.

use dxpdf::model::{EmbeddedFont, EmbeddedFontVariant};
use dxpdf::render::layout::draw_command::{DrawCommand, LayoutedPage};

fn pages(fixture: &str) -> Vec<LayoutedPage> {
    let dir = env!("CARGO_MANIFEST_DIR");
    let bytes = std::fs::read(format!("{dir}/test-files/{fixture}")).expect("fixture");
    let mut doc = dxpdf::docx::parse(&bytes).expect("fixture parses");
    for (family, font) in [
        ("Calibri", "DxNoGap.ttf"),
        ("Arial", "DxLineGap.ttf"),
        ("Times New Roman", "DxLineGap.ttf"),
    ] {
        doc.embedded_fonts.push(EmbeddedFont {
            family: family.into(),
            variant: EmbeddedFontVariant::Regular,
            data: std::fs::read(format!("{dir}/test-files/fonts/{font}")).expect("font"),
        });
    }
    dxpdf::render::resolve_and_layout(doc).1
}

/// (page, baseline y) of the one text command reading `label`.
fn at(pages: &[LayoutedPage], label: &str) -> (usize, f32) {
    let hits: Vec<(usize, f32)> = pages
        .iter()
        .enumerate()
        .flat_map(|(i, p)| p.commands.iter().map(move |c| (i, c)))
        .filter_map(|(i, c)| match c {
            DrawCommand::Text { text, position, .. } if text.trim_end() == label => {
                Some((i, position.y.raw()))
            }
            _ => None,
        })
        .collect();
    assert_eq!(hits.len(), 1, "exactly one {label:?}, got {hits:?}");
    hits[0]
}

/// Baseline-to-baseline distance from `from` down to `to`, on one page.
fn gap(pages: &[LayoutedPage], from: &str, to: &str) -> f32 {
    let (pa, a) = at(pages, from);
    let (pb, b) = at(pages, to);
    assert_eq!(pa, pb, "{from} and {to} share a page");
    b - a
}

fn assert_gap(pages: &[LayoutedPage], from: &str, to: &str, want: f32, why: &str) {
    let got = gap(pages, from, to);
    assert!(
        (got - want).abs() < 0.01,
        "{from} → {to}: {why} — want {want}, got {got}"
    );
}

/// The fixture faces' reach at `size`: (above the baseline, below it).
fn no_gap(size: f32) -> (f32, f32) {
    (0.95 * size, 0.25 * size)
}
fn line_gap(size: f32) -> (f32, f32) {
    ((0.1 + 0.9) * size, 0.2 * size)
}

/// The reference lines are 11pt "Calibri".
const REF: (f32, f32) = (0.95 * 11.0, 0.25 * 11.0);

/// A face's line gap goes above its text and nothing below it: a 22pt Arial
/// line's baseline sits the gap lower than its ascent alone puts it, and the
/// line under it starts at its descent. Word: R04 → AR22s1a 23.52, which is
/// 2.95 + 0.72 + 19.92 = 23.59; AR22s1b → R05 15.36, which is 4.66 + 10.47 =
/// 15.13; Times the same way, 23.28 and 15.60.
#[test]
fn a_face_s_line_gap_goes_above_its_text() {
    let p = pages("line-box.docx");
    let (above, below) = line_gap(22.0);
    assert_gap(&p, "R04", "AR22s1a", REF.1 + above, "the gap above");
    assert_gap(
        &p,
        "AR22s1a",
        "AR22s1b",
        above + below,
        "the line is gap, ascent, descent",
    );
    assert_gap(&p, "AR22s1b", "R05", below + REF.0, "nothing below");
}

/// A line of two faces is as tall as the larger one's line, with the larger
/// one's gap above. Word: 22pt Arial beside 11pt Calibri (MX3) sets a 25.44pt
/// line, Arial's own 25.30, 23.52 below the reference; 22pt Calibri beside
/// 11pt Arial (MX1) 26.88, Calibri's 26.86.
#[test]
fn a_line_of_two_faces_takes_the_larger_one_s_box() {
    let p = pages("line-box.docx");
    let (above, below) = line_gap(22.0);
    assert_gap(&p, "MX3s1a", "MX3s1b", above + below, "Arial's line");
    let (above, below) = no_gap(22.0);
    assert_gap(&p, "MX1s1a", "MX1s1b", above + below, "Calibri's line");
}

/// A table cell's first line takes its gap above it too. Word: the 22pt Arial
/// line in `TCs1` sits 24.00 below the reference, as a body line does.
#[test]
fn a_cell_line_takes_its_gap_above_it() {
    let p = pages("line-box.docx");
    let (above, below) = line_gap(22.0);
    assert_gap(&p, "TCs1a", "TCs1b", above + below, "the cell's line pitch");
    let r = at(&p, "R49").1;
    let a = at(&p, "TCs1a").1;
    // The reference's descent, the table's 0.5pt top border, then the line.
    assert!(
        (a - r - (REF.1 + 0.5 + above)).abs() < 0.01,
        "the first line's baseline is the gap and ascent below the cell's top: {}",
        a - r
    );
}
