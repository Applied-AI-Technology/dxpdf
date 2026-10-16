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

/// The inline pictures' (page, top, bottom), in document order.
fn pictures(pages: &[LayoutedPage]) -> Vec<(usize, f32, f32)> {
    pages
        .iter()
        .enumerate()
        .flat_map(|(i, p)| p.commands.iter().map(move |c| (i, c)))
        .filter_map(|(i, c)| match c {
            DrawCommand::Image { rect, .. } => Some((
                i,
                rect.origin.y.raw(),
                rect.origin.y.raw() + rect.size.height.raw(),
            )),
            _ => None,
        })
        .collect()
}

/// An 11pt "Calibri" line — the inline fixture's paragraphs and references.
const LINE: f32 = 0.95 * 11.0 + 0.25 * 11.0;

/// An inline picture stands on the baseline: text beside a 40pt picture sits
/// on its foot, and the line is the picture and the text's descent; a 6pt
/// picture sits on the baseline of a line it does not raise; a picture alone
/// is exactly its own height. Word: I2's text 42.96 below R02 and on the
/// picture's foot, R03 13.20 below it; I3's picture's foot on the baseline;
/// R01 to R02 53.52, the picture and the two references' reach.
#[test]
fn an_inline_picture_stands_on_the_baseline() {
    let p = pages("line-box-inline.docx");
    let pics = pictures(&p);
    assert_gap(&p, "R01", "R02", REF.1 + 40.0 + REF.0, "a picture alone");
    assert_gap(
        &p,
        "R02",
        "I2a",
        REF.1 + 40.0,
        "the text on the picture's foot",
    );
    assert_gap(&p, "I2a", "R03", REF.1 + REF.0, "then the text's descent");
    let (_, i2) = at(&p, "I2a");
    assert!(
        (pics[1].2 - i2).abs() < 0.01,
        "I2's picture ends on the baseline"
    );
    let (_, i3) = at(&p, "I3a");
    assert!(
        (pics[2].2 - i3).abs() < 0.01,
        "I3's 6pt picture ends on the baseline"
    );
    assert_gap(
        &p,
        "R03",
        "I3a",
        REF.1 + REF.0,
        "and does not raise the line",
    );
}

/// An Auto multiple scales a line's text height and adds the difference, so a
/// picture is not scaled with it. Word: a 40pt picture alone at 1.15 (I4) on a
/// 42.0pt line, 15% of the paragraph mark's 13.43pt over the picture; the
/// picture beside 11pt text (I9) on 44.89, the picture, the descent and 15%.
#[test]
fn a_multiple_adds_the_text_s_share() {
    let p = pages("line-box-inline.docx");
    let extra = 0.15 * LINE;
    assert_gap(
        &p,
        "R04",
        "R05",
        REF.1 + 40.0 + extra + REF.0,
        "a picture alone",
    );
    assert_gap(
        &p,
        "R09",
        "I9a",
        REF.1 + 40.0,
        "the text on the picture's foot",
    );
    assert_gap(
        &p,
        "I9a",
        "R10",
        REF.1 + extra + REF.0,
        "the share below it",
    );
}

/// A break's run counts for nothing on a line that holds anything else: an
/// Arial line broken by a `<w:br/>` in a 22pt Calibri run is as tall as Arial
/// makes it, and so is the empty line a closing break leaves, which takes the
/// paragraph mark's face. Word: B1a → B1b 12.48, Arial's 12.65 and not
/// Calibri 22's 26.86; B8a → R19 25.68, Arial's descent, an 11pt Arial line
/// and the reference.
#[test]
fn a_break_s_run_counts_only_on_a_line_of_its_own() {
    let p = pages("line-box-inline.docx");
    let (above, below) = line_gap(11.0);
    assert_gap(&p, "R11", "B1a", REF.1 + above, "Arial's reach above");
    assert_gap(
        &p,
        "B1a",
        "B1b",
        above + below,
        "Arial's line, not the break's",
    );
    assert_gap(
        &p,
        "B8a",
        "R19",
        below + (above + below) + REF.0,
        "the line the break leaves is the mark's",
    );
}

/// `exact` puts the baseline 80% of the way down the line whatever the face,
/// and an inline picture's foot with it. Word: every face and size at exact
/// 30pt 26.88 below the reference (2.95 + 24), at exact 8pt 9.36 (2.95 +
/// 6.4); a 40pt picture at exact 20pt with its foot 16pt into the line.
#[test]
fn an_exact_line_s_baseline_is_80_percent_down() {
    let p = pages("line-box.docx");
    for (r, case) in [("R25", "CA11x30"), ("R28", "AR22x30"), ("R29", "TR11x30")] {
        assert_gap(&p, r, &format!("{case}a"), REF.1 + 24.0, "0.8 of 30pt");
        assert_gap(
            &p,
            &format!("{case}a"),
            &format!("{case}b"),
            30.0,
            "the line",
        );
    }
    assert_gap(&p, "R32", "CA22x8a", REF.1 + 6.4, "0.8 of 8pt");
    let q = pages("line-box-inline.docx");
    let top = at(&q, "R05").1 + REF.1;
    let foot = pictures(&q)[4].2;
    assert!(
        (foot - (top + 16.0)).abs() < 0.01,
        "I5's picture foot: {}",
        foot - top
    );
}

/// `atLeast` puts what it adds above the text, so a padded line's baseline is
/// its descent up from the bottom. Word: 11pt Calibri at least 30pt 30.00
/// below the reference and the next 13.44 below it; a 40pt picture at least
/// 60pt with its foot on the line's.
#[test]
fn an_at_least_line_pads_above_its_text() {
    let p = pages("line-box.docx");
    let (_, below) = no_gap(11.0);
    assert_gap(
        &p,
        "R37",
        "CA11l30a",
        REF.1 + (30.0 - below),
        "padded above",
    );
    assert_gap(&p, "CA11l30b", "R38", below + REF.0, "nothing below");
    let (_, below22) = line_gap(22.0);
    assert_gap(&p, "R40", "AR22l30a", REF.1 + 30.0 - below22, "Arial too");
    let q = pages("line-box-inline.docx");
    let top = at(&q, "R06").1 + REF.1;
    let foot = pictures(&q)[5].2;
    assert!(
        (foot - (top + 60.0)).abs() < 0.01,
        "I6's picture foot: {}",
        foot - top
    );
}
