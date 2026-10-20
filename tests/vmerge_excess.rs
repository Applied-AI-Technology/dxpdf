//! Which rows grow when a vertically merged cell outgrows them, over
//! `test-files/vmerge-excess.docx` (`scripts/make_vmerge_excess_fixture.py`),
//! measured in Word (Word for Mac 16, PDF export, 2026-10-18).
//!
//! The span's last row takes the shortfall (`issue-165-vmerge.docx`), and the
//! span already holds the room between its rows that a border in another
//! column takes: each table's 0.5pt line between rows 1 and 2 of column B
//! does not make the merged column's span 0.5pt taller than its content.
//! Two real-world notice templates close with V1.
//!
//! Asserted as Word's distances from that line to the paragraph after the
//! table, which the table's bottom decides.

use dxpdf::render::layout::draw_command::{DrawCommand, LayoutedPage};

fn page() -> LayoutedPage {
    let path = format!(
        "{}/test-files/vmerge-excess.docx",
        env!("CARGO_MANIFEST_DIR")
    );
    let bytes = std::fs::read(&path).expect("fixture is committed");
    let doc = dxpdf::docx::parse(&bytes).expect("fixture parses");
    dxpdf::render::resolve_and_layout(doc).1.remove(0)
}

fn baseline(page: &LayoutedPage, label: &str) -> f32 {
    page.commands
        .iter()
        .find_map(|c| match c {
            DrawCommand::Text { text, position, .. } if &**text == label => Some(position.y.raw()),
            _ => None,
        })
        .unwrap_or_else(|| panic!("{label} is drawn"))
}

/// The tops of the rules between rows 1 and 2, one per table.
fn rules(page: &LayoutedPage) -> Vec<f32> {
    let mut ys: Vec<f32> = page
        .commands
        .iter()
        .filter_map(|c| match c {
            DrawCommand::Rect { rect, .. } if rect.size.height.raw() < 1.5 => {
                Some(rect.origin.y.raw())
            }
            _ => None,
        })
        .collect();
    ys.sort_by(f32::total_cmp);
    ys.dedup_by(|a, b| (*a - *b).abs() < 0.01);
    ys
}

#[test]
fn a_merged_span_holds_the_room_its_rows_reserve_for_borders() {
    let p = page();
    let rules = rules(&p);
    assert_eq!(rules.len(), 5, "one rule per table: {rules:?}");
    // Word: SEP<n>'s baseline below the rule.
    let word = [44.88, 43.92, 89.76, 76.56, 44.64];
    for (n, (rule, want)) in rules.iter().zip(word).enumerate() {
        let got = baseline(&p, &format!("SEP{}", n + 1)) - rule;
        assert!(
            (got - want).abs() < 0.3,
            "V{}: want {want}, got {got}",
            n + 1
        );
    }
}
