//! §17.3.1.24 the left and right borders of a bordered paragraph and of the
//! box consecutive bordered paragraphs share, over
//! `test-files/paragraph-side-borders.docx`
//! (`scripts/make_paragraph_side_borders_fixture.py`), measured in Word (Word
//! for Mac 16, PDF export, 2026-10-17).
//!
//! A side border stands its `w:space` and 1.35pt more off the text, its line
//! outside that, and the top and bottom lines run over the corners. A box's
//! sides run unbroken through the spacing between its paragraphs, and through
//! the first one's space before and the last one's space after unless a top or
//! bottom line closes it there or it is shaded. A `w:between` line stands its
//! space off the text on both sides and breaks the sides.
//!
//! Every line is an exact 12pt one, so nothing here depends on a face. Word's
//! numbers are quoted per case; they sit within 0.12pt of these, Word's own
//! positioning grain in its PDFs.

use dxpdf::render::layout::draw_command::{DrawCommand, LayoutedPage};

fn pages() -> Vec<LayoutedPage> {
    let path = format!(
        "{}/test-files/paragraph-side-borders.docx",
        env!("CARGO_MANIFEST_DIR")
    );
    let bytes = std::fs::read(&path).expect("fixture is committed");
    let doc = dxpdf::docx::parse(&bytes).expect("fixture parses");
    dxpdf::render::resolve_and_layout(doc).1
}

/// A painted line as the rectangle it covers: `(x0, y0, x1, y1)`.
type Ink = (f32, f32, f32, f32);

/// Every line on the page as the rectangle it covers, and whether it runs
/// down the page.
fn strokes(page: &LayoutedPage) -> Vec<(Ink, bool)> {
    page.commands
        .iter()
        .filter_map(|c| match c {
            DrawCommand::Line { line, width, .. } => {
                let (a, b, half) = (line.start, line.end, width.raw() / 2.0);
                let (x0, x1) = (a.x.raw().min(b.x.raw()), a.x.raw().max(b.x.raw()));
                let (y0, y1) = (a.y.raw().min(b.y.raw()), a.y.raw().max(b.y.raw()));
                Some(if y0 == y1 {
                    ((x0, y0 - half, x1, y0 + half), false)
                } else {
                    ((x0 - half, y0, x0 + half, y1), true)
                })
            }
            _ => None,
        })
        .collect()
}

fn lines(page: &LayoutedPage) -> Vec<Ink> {
    strokes(page).into_iter().map(|(ink, _)| ink).collect()
}

fn fills(page: &LayoutedPage) -> Vec<Ink> {
    page.commands
        .iter()
        .filter_map(|c| match c {
            DrawCommand::Rect { rect, .. } => Some((
                rect.origin.x.raw(),
                rect.origin.y.raw(),
                rect.origin.x.raw() + rect.size.width.raw(),
                rect.origin.y.raw() + rect.size.height.raw(),
            )),
            _ => None,
        })
        .collect()
}

/// The vertical runs of the side border whose inner edge is at `inner` (the
/// left one, outside it) between `from` and `to`, joined where they meet.
fn left_runs(page: &LayoutedPage, inner: f32, from: f32, to: f32) -> Vec<(f32, f32)> {
    let mut runs: Vec<(f32, f32)> = strokes(page)
        .into_iter()
        .filter(|&((_, y0, x1, y1), down)| down && near(x1, inner) && y0 >= from && y1 <= to)
        .map(|((_, y0, _, y1), _)| (y0, y1))
        .collect();
    runs.sort_by(|a, b| a.0.total_cmp(&b.0));
    let mut joined: Vec<(f32, f32)> = Vec::new();
    for (y0, y1) in runs {
        match joined.last_mut() {
            Some(last) if (y0 - last.1).abs() < 0.01 => last.1 = y1,
            _ => joined.push((y0, y1)),
        }
    }
    joined
}

fn near(got: f32, want: f32) -> bool {
    (got - want).abs() < 0.01
}

fn assert_runs(got: &[(f32, f32)], want: &[(f32, f32)], why: &str) {
    assert!(
        got.len() == want.len()
            && got
                .iter()
                .zip(want)
                .all(|(g, w)| near(g.0, w.0) && near(g.1, w.1)),
        "{why}: want {want:?}, got {got:?}"
    );
}

fn assert_line(page: &LayoutedPage, want: Ink, why: &str) {
    let all = lines(page);
    assert!(
        all.iter().any(|l| near(l.0, want.0)
            && near(l.1, want.1)
            && near(l.2, want.2)
            && near(l.3, want.3)),
        "{why}: no line {want:?} among {all:?}"
    );
}

/// The text is at 72; a 1pt side with 4pt of space is 66.65..65.65 on the
/// left and 545.35..546.35 on the right.
const LEFT_INNER: f32 = 72.0 - 4.0 - 1.35;
const RIGHT_INNER: f32 = 540.0 + 4.0 + 1.35;

/// One boxed paragraph. Word: the sides 65.76–66.72 and 545.28–546.24 from
/// 84.96 to 99.12, the top 84.0–84.96 and the bottom 99.12–100.08, both from
/// 65.76 to 546.24.
#[test]
fn a_side_border_stands_its_space_and_more_off_the_text() {
    let p = &pages()[0];
    assert_runs(
        &left_runs(p, LEFT_INNER, 80.0, 110.0),
        &[(85.0, 99.0)],
        "S1",
    );
    assert_line(p, (RIGHT_INNER, 85.0, RIGHT_INNER + 1.0, 99.0), "S1 right");
    assert_line(
        p,
        (LEFT_INNER - 1.0, 84.0, RIGHT_INNER + 1.0, 85.0),
        "S1 top, over the corners",
    );
    assert_line(
        p,
        (LEFT_INNER - 1.0, 99.0, RIGHT_INNER + 1.0, 100.0),
        "S1 bottom",
    );
}

/// The inset does not follow the line's width, the text's size or a
/// first-line indent, and is from the paragraph's own indents. Word: H4s0's
/// inner edge 70.56, H4s10's 60.72, H24s10's 60.72 (outer 57.84), HT6's and
/// HT20's 66.72, HI's 102.72 and 509.28, HF's 66.72.
#[test]
fn the_inset_is_the_same_for_every_line_and_indent() {
    let p = pages();
    let inner_edges = |page: &LayoutedPage, from: f32, to: f32| -> Vec<f32> {
        let mut xs: Vec<f32> = strokes(page)
            .into_iter()
            .filter(|&((_, y0, _, y1), down)| down && y0 >= from && y1 <= to)
            .map(|((_, _, x1, _), _)| x1)
            .filter(|&x| x < 300.0)
            .collect();
        xs.dedup_by(|a, b| near(*a, *b));
        xs
    };
    let edge = |page: &LayoutedPage, from: f32, to: f32| -> f32 {
        let xs = inner_edges(page, from, to);
        assert_eq!(xs.len(), 1, "one left side in {from}..{to}: {xs:?}");
        xs[0]
    };
    assert!(near(edge(&p[1], 296.0, 314.0), 72.0 - 1.35), "H4s0");
    assert!(near(edge(&p[1], 335.0, 353.0), 72.0 - 10.0 - 1.35), "H4s10");
    assert!(
        near(edge(&p[1], 418.0, 441.0), 72.0 - 10.0 - 1.35),
        "H24s10"
    );
    assert_line(
        &p[1],
        (72.0 - 10.0 - 1.35 - 3.0, 423.0, 72.0 - 10.0 - 1.35, 437.0),
        "H24s10's 3pt line outside its inner edge",
    );
    assert!(near(edge(&p[1], 620.0, 640.0), LEFT_INNER), "HT6");
    assert!(near(edge(&p[1], 660.0, 680.0), LEFT_INNER), "HT20");
    assert!(near(edge(&p[1], 700.0, 720.0), 108.0 - 4.0 - 1.35), "HI");
    assert_line(
        &p[1],
        (504.0 + 5.35, 703.0, 504.0 + 6.35, 717.0),
        "HI's right",
    );
    assert!(near(edge(&p[2], 95.0, 125.0), LEFT_INNER), "HF");
}

/// The box's sides run unbroken through the spacing between its paragraphs.
/// Word: S2 124.08–178.08; S3 206.16–256.08; S10 94.08–150.0 (9pt before,
/// 3pt after: the upper paragraph's run ends 3pt below its text, the lower
/// one's starts there); S11 178.08–232.08.
#[test]
fn a_box_s_sides_run_through_the_space_between_its_paragraphs() {
    let p = pages();
    // Left and right only, 6pt after each: through the last one's space after.
    assert_runs(
        &left_runs(&p[0], LEFT_INNER, 120.0, 190.0),
        &[(124.0, 178.0)],
        "S2",
    );
    // Boxed, 3pt before and 6pt after: from the top line to the bottom one.
    assert_runs(
        &left_runs(&p[0], LEFT_INNER, 200.0, 270.0),
        &[(206.0, 256.0)],
        "S3",
    );
    // 9pt before, 3pt after.
    assert_runs(
        &left_runs(&p[1], LEFT_INNER, 90.0, 160.0),
        &[(94.0, 150.0)],
        "S10",
    );
    // 6pt before each, no lines: through the first one's space before.
    assert_runs(
        &left_runs(&p[1], LEFT_INNER, 170.0, 240.0),
        &[(178.0, 232.0)],
        "S11",
    );
}

/// A lone paragraph's sides take in its space after too, and a box open at
/// its foot reaches through its last one's. Word: S12 256.08–274.08, S9
/// 654.0–708.0, S14 542.88–598.08.
#[test]
fn an_open_box_reaches_through_its_outer_space() {
    let p = pages();
    assert_runs(
        &left_runs(&p[1], LEFT_INNER, 250.0, 280.0),
        &[(256.0, 274.0)],
        "S12",
    );
    assert_runs(
        &left_runs(&p[0], LEFT_INNER, 650.0, 720.0),
        &[(654.0, 708.0)],
        "S9",
    );
    assert_runs(
        &left_runs(&p[1], LEFT_INNER, 540.0, 610.0),
        &[(543.0, 598.0)],
        "S14",
    );
    // S14 has no right border: its top line ends 1.35pt past the text.
    assert_line(&p[1], (LEFT_INNER - 1.0, 542.0, 541.35, 543.0), "S14 top");
}

/// A shaded box keeps its outer space outside, and fills between its sides
/// when it has both, or 1.35pt past the text when it has one. Word: S5's
/// fill 70.56–541.44 by 378.0–426.0, S8's 66.72–545.28 by 576.0–624.0, S13's
/// sides and fill 469.92–517.92.
#[test]
fn a_shaded_box_keeps_its_outer_space_outside() {
    let p = pages();
    assert_runs(
        &left_runs(&p[0], LEFT_INNER, 370.0, 440.0),
        &[(378.0, 426.0)],
        "S5",
    );
    assert_runs(
        &left_runs(&p[1], LEFT_INNER, 465.0, 525.0),
        &[(470.0, 518.0)],
        "S13",
    );
    let fill = |page: &LayoutedPage, from: f32, to: f32| -> Ink {
        let mut f: Vec<Ink> = fills(page)
            .into_iter()
            .filter(|r| r.1 >= from && r.3 <= to)
            .collect();
        f.sort_by(|a, b| a.1.total_cmp(&b.1));
        assert!(!f.is_empty(), "a fill in {from}..{to}");
        (f[0].0, f[0].1, f[0].2, f[f.len() - 1].3)
    };
    let s5 = fill(&p[0], 370.0, 440.0);
    assert!(
        near(s5.0, 70.65) && near(s5.2, 541.35) && near(s5.1, 378.0) && near(s5.3, 426.0),
        "S5 fill {s5:?}"
    );
    let s8 = fill(&p[0], 570.0, 640.0);
    assert!(
        near(s8.0, LEFT_INNER) && near(s8.2, RIGHT_INNER) && near(s8.1, 576.0) && near(s8.3, 624.0),
        "S8 fill {s8:?}"
    );
    // Word paints a rect per member; one rect down the box leaves no seam
    // where two would meet (`section::helpers::push_paragraph`).
    let members = fills(&p[0])
        .into_iter()
        .filter(|r| r.1 >= 570.0 && r.3 <= 640.0)
        .count();
    assert_eq!(members, 1, "S8's three members paint one fill");
}

/// `w:between` stands its 1pt of space off the text on both sides, after the
/// space that collapsed between the two paragraphs, and breaks the sides.
/// Word: the lines at 311.04–312.0 and 331.92–332.88 across the box, the sides
/// 291.12–311.04, 312.0–331.92 and 332.88–347.04.
#[test]
fn a_between_line_stands_off_both_paragraphs_and_breaks_the_sides() {
    let p = &pages()[0];
    assert_line(
        p,
        (LEFT_INNER - 1.0, 311.0, RIGHT_INNER + 1.0, 312.0),
        "S4 first between",
    );
    assert_line(
        p,
        (LEFT_INNER - 1.0, 332.0, RIGHT_INNER + 1.0, 333.0),
        "S4 second between",
    );
    assert_runs(
        &left_runs(p, LEFT_INNER, 285.0, 360.0),
        &[(291.0, 311.0), (312.0, 332.0), (333.0, 347.0)],
        "S4 sides",
    );
}
