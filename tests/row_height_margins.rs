//! §17.4.80 `w:trHeight` against the §17.4.42/§17.4.43 top and bottom cell
//! margins and the table's horizontal borders, over
//! `test-files/row-height-margins.docx` (`scripts/make_row_height_margins_fixture.py`),
//! measured in Word (Microsoft Word for Mac 16, PDF export, 2026-10-15) in
//! compatibility modes 15 and absent, which agree. Every assertion is a
//! distance between two labels of one table, so no glyph metric is pinned; the
//! margins are 99 twips (4.95pt) and the borders 0.5pt unless a case says
//! otherwise.

use dxpdf::render::layout::draw_command::{DrawCommand, LayoutedPage};

fn pages() -> Vec<LayoutedPage> {
    let path = format!(
        "{}/test-files/row-height-margins.docx",
        env!("CARGO_MANIFEST_DIR")
    );
    let bytes = std::fs::read(&path).expect("fixture is committed");
    let doc = dxpdf::docx::parse(&bytes).expect("fixture parses");
    dxpdf::render::resolve_and_layout(doc).1
}

/// Baseline y of the one text command reading `label`, wherever it is.
fn y(pages: &[LayoutedPage], label: &str) -> f32 {
    let hits: Vec<f32> = pages
        .iter()
        .flat_map(|p| p.commands.iter())
        .filter_map(|c| match c {
            DrawCommand::Text { text, position, .. } if &**text == label => Some(position.y.raw()),
            _ => None,
        })
        .collect();
    assert_eq!(hits.len(), 1, "exactly one {label:?}, got {hits:?}");
    hits[0]
}

/// The two row pitches of table `tag`: label 1 to 2, and 2 to 3.
fn pitches(pages: &[LayoutedPage], tag: &str) -> (f32, f32) {
    let r = |i: u32| y(pages, &format!("{tag}r{i}"));
    (r(2) - r(1), r(3) - r(2))
}

const TOL: f32 = 0.01;
const MARGIN: f32 = 4.95;
const RULE: f32 = 0.5;

fn assert_pitch(pages: &[LayoutedPage], tag: &str, want: f32, why: &str) {
    let (a, b) = pitches(pages, tag);
    assert!(
        (a - want).abs() < TOL && (b - want).abs() < TOL,
        "{tag}: {why} — want {want}, got {a} and {b}"
    );
}

/// `exact` counts the top margin inside its height and the bottom one outside
/// it. Word: R3 36.0, R4 36.0, R5 41.0, R1 40.9 — the bottom margin is the
/// whole difference, and the top margin none of it. The interior 0.5pt rule
/// is inside the pitch as well.
#[test]
fn an_exact_row_adds_its_bottom_margin_and_not_its_top() {
    let p = pages();
    assert_pitch(&p, "R3", 36.0, "no margins: the declared height");
    assert_pitch(&p, "R4", 36.0, "a top margin is inside the height");
    assert_pitch(&p, "R5", 36.0 + MARGIN, "a bottom margin is outside it");
    assert_pitch(&p, "R1", 36.0 + MARGIN, "both: only the bottom one counts");
}

/// `atLeast` is a floor on the content between the margins, with the rule above
/// the row stacked on top. Word: A1 36.48, A2 41.4, A3 41.5, A4 46.4.
#[test]
fn an_at_least_row_floors_its_content_and_adds_both_margins_and_its_rule() {
    let p = pages();
    assert_pitch(&p, "A1", 36.0 + RULE, "the rule above each row is extra");
    assert_pitch(&p, "A2", 36.0 + RULE + MARGIN, "and so is the top margin");
    assert_pitch(&p, "A3", 36.0 + RULE + MARGIN, "and so is the bottom one");
    assert_pitch(&p, "A4", 36.0 + RULE + 2.0 * MARGIN, "both together");
}

/// Content taller than the floor, and a row with no rule at all, both take
/// their content plus both margins. Word: A5 37.2 against A9 27.36, A6 23.9
/// against A8 13.92 — 9.9pt apart each time, which is the two margins.
#[test]
fn content_past_the_floor_adds_both_margins_like_a_row_with_no_rule() {
    let p = pages();
    for (with, without) in [("A5", "A9"), ("A6", "A8")] {
        let (a, _) = pitches(&p, with);
        let (b, _) = pitches(&p, without);
        assert!(
            (a - b - 2.0 * MARGIN).abs() < TOL,
            "{with} against {without}: the two margins apart, got {a} and {b}"
        );
    }
}

/// The first line of a table sits below the outer top border *and* the top
/// margin, not the larger of the two. Word puts A2's first label 4.95pt lower
/// under its separator than A1's (18.96 against 13.92 below it), where the
/// larger-of-the-two reading gives 4.45pt.
#[test]
fn a_top_margin_sits_below_the_top_border_rather_than_absorbing_it() {
    let p = pages();
    let offset = |tag: &str| y(&p, &format!("{tag}r1")) - y(&p, &format!("SEP{tag}"));
    assert!(
        (offset("A2") - offset("A1") - MARGIN).abs() < TOL,
        "border and margin add: got {} and {}",
        offset("A2"),
        offset("A1")
    );
}

/// The rule between two exact rows belongs to the lower one: a 6pt `insideH`
/// puts the second label 41.5pt below the first (the whole rule at the head
/// of row two, less the 0.5pt outer top border the first label sits under) and
/// the third 36pt below the second. Under `atLeast` the rule is extra to
/// every row instead. Word: E1 41.52 and 36.0, A7 42.0 and 42.0.
#[test]
fn the_rule_between_two_rows_is_the_lower_row_s() {
    let p = pages();
    let (a, b) = pitches(&p, "E1");
    assert!(
        (a - (36.0 + 6.0 - RULE)).abs() < TOL && (b - 36.0).abs() < TOL,
        "E1: want 41.5 and 36, got {a} and {b}"
    );
    assert_pitch(&p, "A7", 36.0 + 6.0, "atLeast: the 6pt rule on every row");
}
