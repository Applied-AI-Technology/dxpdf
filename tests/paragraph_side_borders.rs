//! §17.3.1.24 the left and right borders of a bordered paragraph, over
//! `test-files/paragraph-side-borders.docx`
//! (`scripts/make_paragraph_side_borders_fixture.py`), measured in Word (Word
//! for Mac 16, PDF export, 2026-10-17).
//!
//! A side border stands its `w:space` and 1.35pt more off the text, its line
//! outside that, and the top and bottom lines run over the corners. A
//! `w:between` line stands its space off the text on both sides.
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

/// `w:between` stands its 1pt of space off the text on both sides, after the
/// space that collapsed between the two paragraphs. Word: the lines at
/// 311.04–312.0 and 331.92–332.88 across the box.
#[test]
fn a_between_line_stands_off_both_paragraphs() {
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
}
