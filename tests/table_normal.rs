//! What places a table in compatibility mode when the styles part has no
//! `TableNormal`, over `test-files/table-normal-{none,full,mar,ind,none-14}.docx`
//! (`scripts/make_table_normal_fixtures.py`), measured in Word (Word for Mac
//! 16, PDF export, 2026-10-17). None of the files names mode 15.
//!
//! A `TableNormal` in the file is completed from Word's built-in Normal Table:
//! one that declares only a `tblInd` still gives its tables 108-twip side
//! margins (`-ind`), and one that declares only margins still gives them an
//! indent of 0, to the first cell's text (`-mar`). With no `TableNormal` at all
//! (`-none`, and `-none-14` naming mode 14) a table has no margins but its own
//! and, stating no `tblInd`, stands with its grid on the margin — its text one
//! cell margin in — while a `tblInd` it does state is still to the text.
//! A real-world certificate template is the `-none` shape.
//!
//! Word's numbers are quoted per case; they sit within 0.12pt of these, Word's
//! own positioning grain in its PDFs.

use dxpdf::render::layout::draw_command::{DrawCommand, LayoutedPage};

fn pages(name: &str) -> Vec<LayoutedPage> {
    let path = format!(
        "{}/test-files/table-normal-{name}.docx",
        env!("CARGO_MANIFEST_DIR")
    );
    let bytes = std::fs::read(&path).expect("fixture is committed");
    let doc = dxpdf::docx::parse(&bytes).expect("fixture parses");
    dxpdf::render::resolve_and_layout(doc).1
}

/// The one text command reading `label`: its x and its baseline.
fn text(pages: &[LayoutedPage], label: &str) -> (f32, f32) {
    let hits: Vec<(f32, f32)> = pages[0]
        .commands
        .iter()
        .filter_map(|c| match c {
            DrawCommand::Text { text, position, .. } if &**text == label => {
                Some((position.x.raw(), position.y.raw()))
            }
            _ => None,
        })
        .collect();
    assert_eq!(hits.len(), 1, "exactly one {label:?}, got {hits:?}");
    hits[0]
}

/// The centre of the leftmost rule beside `label`'s line: the table's grid
/// line, which its 0.5pt left border straddles.
fn left_rule(pages: &[LayoutedPage], label: &str) -> f32 {
    let (_, y) = text(pages, label);
    pages[0]
        .commands
        .iter()
        .filter_map(|c| match c {
            DrawCommand::Rect { rect, .. }
                if rect.size.height.raw() > 1.0
                    && rect.origin.y.raw() < y
                    && y < rect.origin.y.raw() + rect.size.height.raw() =>
            {
                Some(rect.origin.x.raw() + rect.size.width.raw() / 2.0)
            }
            _ => None,
        })
        .min_by(f32::total_cmp)
        .unwrap_or_else(|| panic!("no rule beside {label}"))
}

const MARGIN: f32 = 72.0;
/// Built-in Normal Table's 108-twip side margin.
const NORMAL: f32 = 5.4;
/// The 100-twip `tblCellMar` of N2–N4.
const OWN: f32 = 5.0;

fn assert_near(got: f32, want: f32, why: &str) {
    assert!((got - want).abs() < 0.01, "{why}: want {want}, got {got}");
}

/// No `TableNormal`: a table stating no `tblInd` has its grid on the margin.
/// Word: N1's rule 72.0, N2a and N4a 77.04, N2's and N4's rules 72.0.
#[test]
fn with_no_table_normal_an_unindented_table_stands_on_the_margin() {
    for name in ["none", "none-14"] {
        let p = pages(name);
        assert_near(left_rule(&p, "N1a"), MARGIN, &format!("{name}: N1 grid"));
        assert_near(left_rule(&p, "N2a"), MARGIN, &format!("{name}: N2 grid"));
        assert_near(text(&p, "N2a").0, MARGIN + OWN, &format!("{name}: N2 text"));
        assert_near(left_rule(&p, "N4a"), MARGIN, &format!("{name}: N4 grid"));
        assert_near(text(&p, "N4a").0, MARGIN + OWN, &format!("{name}: N4 text"));
    }
}

/// …and one that states a `tblInd` is still indented to its text. Word: N3a
/// 108.0, its rule 102.96.
#[test]
fn with_no_table_normal_a_stated_indent_is_still_to_the_text() {
    for name in ["none", "none-14"] {
        let p = pages(name);
        assert_near(
            text(&p, "N3a").0,
            MARGIN + 36.0,
            &format!("{name}: N3 text"),
        );
        assert_near(
            left_rule(&p, "N3a"),
            MARGIN + 36.0 - OWN,
            &format!("{name}: N3 grid"),
        );
    }
}

/// A `TableNormal` that states any part of Normal Table gets the rest from
/// Word's built-in one: an indent of 0, to the text, and 108-twip margins.
/// Word, in all three: N1a, N2a, N4a 72.0, N3a 108.0; N1's rule 66.48 (Word
/// draws 66.6 on its grain), N2's 66.96.
#[test]
fn a_table_normal_is_completed_from_word_s_own() {
    for name in ["full", "mar", "ind"] {
        let p = pages(name);
        for label in ["N1a", "N2a", "N4a"] {
            assert_near(text(&p, label).0, MARGIN, &format!("{name}: {label} text"));
        }
        assert_near(
            text(&p, "N3a").0,
            MARGIN + 36.0,
            &format!("{name}: N3 text"),
        );
        assert_near(
            left_rule(&p, "N1a"),
            MARGIN - NORMAL,
            &format!("{name}: N1 hangs Normal Table's margin"),
        );
        assert_near(
            left_rule(&p, "N2a"),
            MARGIN - OWN,
            &format!("{name}: N2 hangs its own"),
        );
    }
}
