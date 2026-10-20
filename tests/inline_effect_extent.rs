//! §20.4.2.6 `wp:effectExtent` on an inline picture, over
//! `test-files/inline-effect-extent.docx`
//! (`scripts/make_inline_effect_extent_fixture.py`), measured in Word (Word
//! for Mac 16, PDF export, 2026-10-18).
//!
//! The extent grows the box the picture takes in its line on each side: the
//! line is taller by its top and bottom, the text after the picture moves by
//! its left and right, and the picture is drawn inside it — its left and top
//! in from the box, its foot raised off the baseline by the bottom. Word, for
//! the 36×24pt picture: E0 (none) to E1 (6pt below) and E2 (6pt above) each add
//! 6.0pt to the line, E3 (left) and E4 (right) move `E<n>b` 6.0pt right and
//! leave the line, E5 (3pt all round) does both, and a real template's E6
//! (0.75pt right and below) moves and grows by 0.72 on Word's grain. E3's
//! picture is drawn 6.0pt right of E0's, E1's foot 5.94pt higher above its
//! baseline.

use dxpdf::render::layout::draw_command::{DrawCommand, LayoutedPage};

fn page() -> LayoutedPage {
    let path = format!(
        "{}/test-files/inline-effect-extent.docx",
        env!("CARGO_MANIFEST_DIR")
    );
    let bytes = std::fs::read(&path).expect("fixture is committed");
    let doc = dxpdf::docx::parse(&bytes).expect("fixture parses");
    dxpdf::render::resolve_and_layout(doc).1.remove(0)
}

/// The one text command reading `label`: its x and its baseline.
fn text(page: &LayoutedPage, label: &str) -> (f32, f32) {
    let hits: Vec<(f32, f32)> = page
        .commands
        .iter()
        .filter_map(|c| match c {
            DrawCommand::Text { text, position, .. } if text.trim() == label => {
                Some((position.x.raw(), position.y.raw()))
            }
            _ => None,
        })
        .collect();
    assert_eq!(hits.len(), 1, "exactly one {label:?}, got {hits:?}");
    hits[0]
}

/// The pictures in paragraph order: left, top, right, bottom.
fn pictures(page: &LayoutedPage) -> Vec<[f32; 4]> {
    let mut rects: Vec<[f32; 4]> = page
        .commands
        .iter()
        .filter_map(|c| match c {
            DrawCommand::Image { rect, .. } => Some([
                rect.origin.x.raw(),
                rect.origin.y.raw(),
                rect.origin.x.raw() + rect.size.width.raw(),
                rect.origin.y.raw() + rect.size.height.raw(),
            ]),
            _ => None,
        })
        .collect();
    rects.sort_by(|a, b| a[1].total_cmp(&b[1]));
    rects
}

/// How far a case's line moves the next one down, beyond E0's: the step
/// from the paragraph above to `<case>a`, less the step from `REF` to `E0a`.
fn growth(page: &LayoutedPage, above: &str, case: &str) -> f32 {
    let base = text(page, "E0a").1 - text(page, "REF").1;
    text(page, &format!("{case}a")).1 - text(page, above).1 - base
}

fn assert_near(got: f32, want: f32, why: &str) {
    assert!((got - want).abs() < 0.05, "{why}: want {want}, got {got}");
}

#[test]
fn the_extent_grows_the_line_by_its_top_and_bottom() {
    let p = page();
    assert_near(growth(&p, "E0z", "E1"), 6.0, "E1, 6pt below");
    assert_near(growth(&p, "E1z", "E2"), 6.0, "E2, 6pt above");
    assert_near(growth(&p, "E2z", "E3"), 0.0, "E3, left only");
    assert_near(growth(&p, "E3z", "E4"), 0.0, "E4, right only");
    assert_near(growth(&p, "E4z", "E5"), 6.0, "E5, 3pt all round");
    assert_near(growth(&p, "E5z", "E6"), 0.75, "E6, the template's");
}

#[test]
fn the_text_after_the_picture_moves_by_its_left_and_right() {
    let p = page();
    let base = text(&p, "E0b").0;
    for (case, want) in [
        ("E1", 0.0),
        ("E2", 0.0),
        ("E3", 6.0),
        ("E4", 6.0),
        ("E5", 6.0),
        ("E6", 0.75),
    ] {
        assert_near(text(&p, &format!("{case}b")).0 - base, want, case);
    }
}

/// The picture keeps its own 36×24pt and is drawn inside the grown box: in by
/// the left extent, and raised off the baseline by the bottom one.
#[test]
fn the_picture_is_drawn_inside_its_extent() {
    let p = page();
    let pics = pictures(&p);
    for pic in &pics {
        assert_near(pic[2] - pic[0], 36.0, "width");
        assert_near(pic[3] - pic[1], 24.0, "height");
    }
    let (e0, e1, e3, e5) = (pics[0], pics[1], pics[3], pics[5]);
    assert_near(e3[0] - e0[0], 6.0, "E3 drawn its left extent in");
    assert_near(e5[0] - e0[0], 3.0, "E5 drawn its left extent in");
    let foot = |pic: [f32; 4], case: &str| text(&p, &format!("{case}a")).1 - pic[3];
    assert_near(foot(e0, "E0"), 0.0, "E0 stands on the baseline");
    assert_near(foot(e1, "E1"), 6.0, "E1 raised by its bottom extent");
}
