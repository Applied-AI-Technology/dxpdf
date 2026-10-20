//! §17.4.48 `tblGrid`, §17.4.71 `tcW` and §17.4.52 `tblLayout` — which widths a
//! table's columns take, over `test-files/table-widths.docx`
//! (`scripts/make_table_widths_fixture.py`) and `test-files/autofit-columns.docx`
//! (`scripts/make_autofit_columns_fixture.py`), both measured in Word (Word for
//! Mac 16, PDF export, 2026-10-16/17) in Calibri.
//!
//! The cells' `tcW` decide the columns and the grid plays no part where they
//! state one: every grid line stands as far right as the widest cell ending on
//! it needs (T1–T3, T7). A declared `tblW` scales a fixed table's columns
//! proportionally (T5) and none grows for its content (T6). An autofit table
//! grows each column to the widest word in it, and where that is wider than it
//! may be — the text column without a `tblW`, the `tblW` with one — takes the
//! excess from each column in proportion to how far it stands above its own
//! widest word (T8, T11, F2), keeping the minimums whole when only they fit
//! (T10). Below its width it scales up proportionally (F5), and grid columns no
//! cell covers take nothing (T12, T13).
//!
//! Laid out in `DxNoGap` as "Calibri", whose every advance is half an em, so a
//! word of n characters at 10pt is 5n pt and each expectation below is the rule
//! applied to known widths rather than Word's Calibri numbers, which the module
//! docs at `build::table::autofit_widths` record.

use dxpdf::model::{EmbeddedFont, EmbeddedFontVariant};
use dxpdf::render::layout::draw_command::{DrawCommand, LayoutedPage};

fn pages(fixture: &str) -> Vec<LayoutedPage> {
    let dir = env!("CARGO_MANIFEST_DIR");
    let bytes = std::fs::read(format!("{dir}/test-files/{fixture}")).expect("fixture");
    let mut doc = dxpdf::docx::parse(&bytes).expect("fixture parses");
    doc.embedded_fonts.push(EmbeddedFont {
        family: "Calibri".into(),
        variant: EmbeddedFontVariant::Regular,
        data: std::fs::read(format!("{dir}/test-files/fonts/DxNoGap.ttf")).expect("font"),
    });
    dxpdf::render::resolve_and_layout(doc).1
}

/// The widths of a table's columns: the gaps between the vertical borders
/// drawn beside the cell labelled `label` (`T<case><row><column>` in
/// `table-widths.docx`, `F<case><column>` in `autofit-columns.docx`).
fn columns(pages: &[LayoutedPage], label: &str) -> Vec<f32> {
    let commands = &pages[0].commands;
    let label_y = |label: &str| {
        commands
            .iter()
            .find_map(|c| match c {
                DrawCommand::Text { text, position, .. } if text.trim_end() == label => {
                    Some(position.y.raw())
                }
                _ => None,
            })
            .unwrap_or_else(|| panic!("no {label}"))
    };
    let y = label_y(label);
    let mut xs: Vec<f32> = commands
        .iter()
        .filter_map(|c| match c {
            DrawCommand::Rect { rect, .. }
                if rect.size.width.raw() < 2.0
                    && rect.size.height.raw() > 2.0
                    && rect.origin.y.raw() < y
                    && rect.origin.y.raw() + rect.size.height.raw() > y =>
            {
                Some(rect.origin.x.raw() + rect.size.width.raw() / 2.0)
            }
            _ => None,
        })
        .collect();
    xs.sort_by(f32::total_cmp);
    xs.dedup_by(|a, b| (*a - *b).abs() < 0.01);
    xs.windows(2).map(|w| w[1] - w[0]).collect()
}

fn assert_columns(actual: &[f32], expected: &[f32], what: &str) {
    assert_eq!(actual.len(), expected.len(), "{what}: {actual:?}");
    for (a, e) in actual.iter().zip(expected) {
        assert!(
            (a - e).abs() < 0.05,
            "{what}: columns {actual:?}, expected {expected:?}"
        );
    }
}

/// A word of `chars` characters in a cell with `TableNormal`'s 5.4pt margins.
fn min(chars: usize) -> f32 {
    chars as f32 * 5.0 + 10.8
}

/// Word's shrink: each column keeps its minimum and its share of `fit` above
/// the minimums in proportion to how far it stood above its own.
fn shrink(grown: &[f32], minimum: &[f32], fit: f32) -> Vec<f32> {
    let floor: f32 = minimum.iter().sum();
    let total: f32 = grown.iter().sum();
    let keep = (fit - floor) / (total - floor);
    grown
        .iter()
        .zip(minimum)
        .map(|(g, m)| m + (g - m) * keep)
        .collect()
}

/// The width a full-width table's columns span under Word 2013: the 468pt text
/// column less half of each 0.5pt outer border.
const FULL: f32 = 467.5;

/// T1/T2: the grid says 241.2 + 241.2, every cell 151.2 + 302.4, and Word draws
/// the cells' widths under either layout.
#[test]
fn the_cells_widths_override_the_grid() {
    let pages = pages("table-widths.docx");
    assert_columns(&columns(&pages, "T111"), &[151.2, 302.4], "T1 fixed");
    assert_columns(&columns(&pages, "T211"), &[151.2, 302.4], "T2 autofit");
}

/// T3: rows of 151.2 + 302.4 and 302.4 + 151.2 make two 302.4pt columns — the
/// widest claim on each grid line. Word draws them off the paper; this engine
/// keeps a fixed table on the sheet (`tests/table_auto_width.rs`), so only the
/// equality survives.
#[test]
fn a_column_is_as_wide_as_its_widest_cell() {
    let pages = pages("table-widths.docx");
    let cols = columns(&pages, "T311");
    assert_eq!(cols.len(), 2, "{cols:?}");
    assert!((cols[0] - cols[1]).abs() < 0.05, "T3: {cols:?}");
}

/// T5: a fixed table's `tblW` scales the cells' 151.2 + 302.4 to 482.4.
#[test]
fn a_fixed_tables_width_scales_its_cells() {
    let pages = pages("table-widths.docx");
    assert_columns(&columns(&pages, "T511"), &[160.8, 321.6], "T5");
}

/// T6: a fixed column keeps its 72pt under a 19-character word.
#[test]
fn a_fixed_column_does_not_grow_for_its_content() {
    let pages = pages("table-widths.docx");
    assert_columns(&columns(&pages, "T611"), &[72.0; 6], "T6");
}

/// T7: a 350pt cell over two 151.2pt columns widens the second — read off
/// the second row, whose three cells draw every line.
#[test]
fn a_spanning_cell_widens_the_last_column_it_covers() {
    let pages = pages("table-widths.docx");
    assert_columns(&columns(&pages, "T721"), &[151.2, 198.8, 151.2], "T7");
}

/// T8: a report template's row. The two tags grow their columns; the table is then
/// wider than the text column and shrinks back to it, the others giving in
/// proportion to their room above their own tags.
#[test]
fn an_autofit_table_shrinks_by_each_columns_room() {
    let pages = pages("table-widths.docx");
    let minimum = [min(19), min(7), min(8), min(9), min(8), min(22)];
    let grown = [min(19), 72.0, 72.0, 72.0, 72.0, min(22)];
    assert_columns(
        &columns(&pages, "T811"),
        &shrink(&grown, &minimum, FULL),
        "T8",
    );
}

/// T9/T10/T11: a declared width narrower than the columns. Equal content takes
/// it equally (T9); content that cannot fit it keeps its minimums, wider than
/// declared, while they fit the column (T10); otherwise each column gives by
/// its room (T11, whose labels are its widest words).
#[test]
fn an_autofit_tables_width_is_taken_by_room_or_by_the_minimums() {
    let pages = pages("table-widths.docx");
    assert_columns(&columns(&pages, "T911"), &[60.0; 6], "T9");
    assert_columns(
        &columns(&pages, "T1011"),
        &[min(19), min(7), min(8), min(9), min(8), min(22)],
        "T10",
    );
    assert_columns(
        &columns(&pages, "T1111"),
        &shrink(&[100.0, 300.0], &[min(5), min(5)], 200.0),
        "T11",
    );
}

/// T12/T13: a row that stops short of its grid. The grid columns it leaves
/// empty take nothing, and its cells fill the table's 400pt `tblW` — at their
/// `tcW` (T12), or all of it for a span with none (T13).
#[test]
fn the_columns_a_short_row_leaves_empty_take_nothing() {
    let pages = pages("table-widths.docx");
    assert_columns(&columns(&pages, "T1211"), &[200.0, 200.0], "T12");
    assert_columns(&columns(&pages, "T1311"), &[400.0], "T13");
}

/// `autofit-columns.docx`: one wide tag grows its column and the table with it
/// (F1); two take the table past the column, which the narrow columns give back
/// (F2); a 100% table scales the grown columns up to its width (F5).
#[test]
fn an_autofit_column_grows_to_its_widest_word() {
    let pages = pages("autofit-columns.docx");
    assert_columns(
        &columns(&pages, "F11"),
        &[min(19), 72.0, 72.0, 72.0, 72.0, 72.0],
        "F1",
    );
    let narrow = min(3); // "F22", wider than "{x}"
    assert_columns(
        &columns(&pages, "F21"),
        &shrink(
            &[min(19), 72.0, 72.0, 72.0, 72.0, min(19)],
            &[min(19), narrow, narrow, narrow, narrow, min(19)],
            FULL,
        ),
        "F2",
    );
    let grown = [min(19), 72.0, 72.0, 72.0, 72.0, 72.0];
    let scale = FULL / grown.iter().sum::<f32>();
    assert_columns(&columns(&pages, "F51"), &grown.map(|w| w * scale), "F5");
}

/// A docx of `body` with no styles or settings part — compatibility mode, so a
/// full-width table spans the column and both of its text insets.
fn document(body: &str) -> Vec<u8> {
    use std::io::Write;
    let mut zip = zip::ZipWriter::new(std::io::Cursor::new(Vec::new()));
    let o = zip::write::SimpleFileOptions::default();
    zip.start_file("[Content_Types].xml", o).unwrap();
    zip.write_all(
        br#"<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml"
    ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>"#,
    )
    .unwrap();
    zip.start_file("_rels/.rels", o).unwrap();
    zip.write_all(
        br#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
    Target="word/document.xml"/>
</Relationships>"#,
    )
    .unwrap();
    zip.start_file("word/document.xml", o).unwrap();
    zip.write_all(
        format!(
            r#"<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>{body}</w:body></w:document>"#
        )
        .as_bytes(),
    )
    .unwrap();
    zip.finish().unwrap().into_inner()
}

/// An autofit table of `cells` (`tcW` in twips, content), 5.4pt cell margins and
/// 0.5pt borders, Calibri 10pt.
fn autofit_table(cells: &[(u32, &str)]) -> String {
    let borders: String = ["top", "left", "bottom", "right", "insideH", "insideV"]
        .iter()
        .map(|s| format!(r#"<w:{s} w:val="single" w:sz="4" w:color="000000"/>"#))
        .collect();
    let grid: String = cells
        .iter()
        .map(|(w, _)| format!(r#"<w:gridCol w:w="{w}"/>"#))
        .collect();
    let tcs: String = cells
        .iter()
        .map(|(w, content)| {
            format!(r#"<w:tc><w:tcPr><w:tcW w:w="{w}" w:type="dxa"/></w:tcPr>{content}</w:tc>"#)
        })
        .collect();
    format!(
        r#"<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/><w:tblBorders>{borders}</w:tblBorders>
<w:tblCellMar><w:left w:w="108" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar></w:tblPr>
<w:tblGrid>{grid}</w:tblGrid><w:tr>{tcs}</w:tr></w:tbl>"#
    )
}

fn para(text: &str) -> String {
    format!(
        r#"<w:p><w:r><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:sz w:val="20"/></w:rPr><w:t>{text}</w:t></w:r></w:p>"#
    )
}

/// A nested autofit table is fitted to the column its enclosing table settles
/// on, not the one it was first offered — Word's own cached grids in real
/// summary-report templates are exactly `autofit_widths` applied to the outer
/// column. Here two 300pt outer columns overflow the text column and shrink,
/// and the nested table in the first, preferring 300pt, must fit what is left
/// of it: its outer edges on the cell's text edges, inside its margins.
#[test]
fn a_nested_autofit_table_fits_its_cells_final_width() {
    let nested = autofit_table(&[(3000, &para("x")), (3000, &para("y"))]);
    let outer = autofit_table(&[
        (6000, &format!("{nested}{}", para("N"))),
        (6000, &para("abcdefghijklmnopqrst")),
    ]);
    let mut doc = dxpdf::docx::parse(&document(&outer)).expect("parses");
    let dir = env!("CARGO_MANIFEST_DIR");
    doc.embedded_fonts.push(EmbeddedFont {
        family: "Calibri".into(),
        variant: EmbeddedFontVariant::Regular,
        data: std::fs::read(format!("{dir}/test-files/fonts/DxNoGap.ttf")).expect("font"),
    });
    let pages = dxpdf::render::resolve_and_layout(doc).1;

    let mut xs: Vec<f32> = pages[0]
        .commands
        .iter()
        .filter_map(|c| match c {
            DrawCommand::Rect { rect, .. }
                if rect.size.width.raw() < 2.0 && rect.size.height.raw() > 2.0 =>
            {
                Some(rect.origin.x.raw() + rect.size.width.raw() / 2.0)
            }
            _ => None,
        })
        .collect();
    xs.sort_by(f32::total_cmp);
    xs.dedup_by(|a, b| (*a - *b).abs() < 0.01);
    // Outer lines at 0, a, b; nested lines between the first two of them.
    assert_eq!(xs.len(), 6, "three outer and three nested lines: {xs:?}");
    let (outer_left, nested_left, nested_right, outer_mid) = (xs[0], xs[1], xs[3], xs[4]);
    // A nested table's own edges straddle its border lines, half of each 0.5pt
    // border outside its box, and the box spans the cell between its margins.
    assert!(
        (nested_left - (outer_left + 5.4 + 0.25)).abs() < 0.05
            && (nested_right - (outer_mid - 5.4 - 0.25)).abs() < 0.05,
        "nested table {nested_left:.2}..{nested_right:.2} does not fill its cell \
         {outer_left:.2}..{outer_mid:.2} between the margins"
    );
}
