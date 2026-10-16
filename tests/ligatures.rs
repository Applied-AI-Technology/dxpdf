//! [MS-DOCX] `w14:ligatures`, over `test-files/ligatures-{15,14,14otf}.docx`
//! (`scripts/make_ligatures_fixtures.py`), measured in Word (Microsoft Word for
//! Mac 16, PDF export, 2026-10-15) in Calibri: every word lands within 0.05pt
//! of Word.
//!
//! Laid out here in `test-files/fonts/DxLiga.ttf` (`scripts/make_font_fixtures.py`),
//! handed to the documents as an embedded "Calibri", so no host font decides
//! the result. Every advance in it is 500 units of a 1000-unit em except its one
//! ligature, `f_i`, at 700 — so each `fi` it forms takes 300 units off a word —
//! and its `kern` table has one pair, `f_i` against `f`, at −50.

use dxpdf::model::{EmbeddedFont, EmbeddedFontVariant};
use dxpdf::render::layout::draw_command::{DrawCommand, LayoutedPage};

fn pages(name: &str) -> Vec<LayoutedPage> {
    let dir = env!("CARGO_MANIFEST_DIR");
    let bytes = std::fs::read(format!("{dir}/test-files/{name}")).expect("fixture");
    let mut doc = dxpdf::docx::parse(&bytes).expect("fixture parses");
    doc.embedded_fonts.push(EmbeddedFont {
        family: "Calibri".into(),
        variant: EmbeddedFontVariant::Regular,
        data: std::fs::read(format!("{dir}/test-files/fonts/DxLiga.ttf")).expect("font"),
    });
    dxpdf::render::resolve_and_layout(doc).1
}

/// x of `word` on the line whose first text starts with `label`.
fn x_on(pages: &[LayoutedPage], label: &str, word: &str) -> f32 {
    let texts: Vec<(f32, f32, &str)> = pages[0]
        .commands
        .iter()
        .filter_map(|c| match c {
            DrawCommand::Text { text, position, .. } => {
                Some((position.y.raw(), position.x.raw(), &**text))
            }
            _ => None,
        })
        .collect();
    let line_y = texts
        .iter()
        .find(|t| t.2.starts_with(&format!("{label} ")))
        .unwrap_or_else(|| panic!("no line {label}"))
        .0;
    texts
        .iter()
        .find(|t| (t.0 - line_y).abs() < 0.01 && t.2.trim_end() == word)
        .unwrap_or_else(|| panic!("no {word:?} on line {label}"))
        .1
}

/// `office fifty fluff staff Wafflt ` from `office` to `END`.
fn span(pages: &[LayoutedPage], label: &str) -> f32 {
    x_on(pages, label, "END") - x_on(pages, label, "office")
}

fn pt(units: f32) -> f32 {
    units * 11.0 / 1000.0
}

/// 32 letters and spaces at 500 units each.
const PLAIN: f32 = 16_000.0;
/// `office` and `fifty` each ligate one `fi`.
const LIGATED: f32 = PLAIN - 600.0;
const TOL: f32 = 0.01;

fn assert_span(pages: &[LayoutedPage], label: &str, units: f32, why: &str) {
    let got = span(pages, label);
    assert!(
        (got - pt(units)).abs() < TOL,
        "{label}: {why} — want {}, got {got}",
        pt(units)
    );
}

/// Under Word 2013 the classes asked for are formed: `standardContextual` and
/// `standard` ligate, the control does not, and `contextual` alone asks for
/// `clig`, which this font (like Calibri) does not use. Word: L1 and L3 span
/// 120.09pt, L2 and L4 121.78pt.
#[test]
fn under_word_2013_the_standard_class_ligates() {
    let p = pages("ligatures-15.docx");
    assert_span(&p, "L1", LIGATED, "standardContextual");
    assert_span(&p, "L2", PLAIN, "no w14:ligatures");
    assert_span(&p, "L3", LIGATED, "standard");
    assert_span(&p, "L4", PLAIN, "contextual: clig, not liga");
}

/// Character spacing turns them off: Word draws L5's words at their unligated
/// widths plus a point a unit.
#[test]
fn character_spacing_turns_ligatures_off() {
    let p = pages("ligatures-15.docx");
    let office = x_on(&p, "L5", "fifty") - x_on(&p, "L5", "office");
    // `office ` is seven units, each followed by the 1pt of spacing.
    assert!(
        (office - (pt(3_500.0) + 7.0)).abs() < TOL,
        "unligated and spaced, got {office}"
    );
}

/// A ligated run is kerned between the glyphs it shaped to, from the `kern`
/// table — here the one pair, `f_i` against `f`, inside `fifty`.
#[test]
fn a_ligated_run_is_kerned_between_its_glyphs() {
    let p = pages("ligatures-15.docx");
    let fifty = x_on(&p, "L6", "fluff") - x_on(&p, "L6", "fifty");
    assert!(
        (fifty - pt(2_700.0 - 50.0)).abs() < TOL,
        "f_i·f·t·y·space less the pair, got {fifty}"
    );
}

/// A ligature forms across a run boundary between runs that both take it:
/// `of` | `fice` is drawn as `office` (Word: `END` 27.29pt after `office`'s
/// start in L7, as in L1).
#[test]
fn a_ligature_forms_across_a_run_boundary() {
    let p = pages("ligatures-15.docx");
    let office = x_on(&p, "L7", "END") - x_on(&p, "L7", "office");
    assert!(
        (office - pt(3_200.0)).abs() < TOL,
        "the two runs joined and ligated, got {office}"
    );
}

/// The document's compatibility settings decide whether any of it happens:
/// mode 14 forms nothing, and mode 14 with `enableOpenTypeFeatures` forms
/// what mode 15 does. Word: L1 spans 121.78pt in mode 14 and 120.09pt with
/// the setting.
#[test]
fn ligatures_need_open_type_features_enabled() {
    assert_span(&pages("ligatures-14.docx"), "L1", PLAIN, "mode 14");
    assert_span(
        &pages("ligatures-14otf.docx"),
        "L1",
        LIGATED,
        "mode 14 + enableOpenTypeFeatures",
    );
}

/// A ligature extracts as the text it draws — copied, searched and read aloud
/// as `fi`, not as nothing. Its glyph has no `cmap` entry, and in
/// `DxLigaUnnamed.ttf` (`DxLiga` with a format 3 `post` table, as Carlito
/// ships) no name either, so the only record of its text is the run it was
/// shaped from: the PDF's `ToUnicode` map must give the ligature `fi`.
#[test]
fn a_ligature_extracts_as_the_text_it_draws() {
    let dir = env!("CARGO_MANIFEST_DIR");
    let bytes = std::fs::read(format!("{dir}/test-files/ligatures-15.docx")).expect("fixture");
    let mut doc = dxpdf::docx::parse(&bytes).expect("fixture parses");
    doc.embedded_fonts.push(EmbeddedFont {
        family: "Calibri".into(),
        variant: EmbeddedFontVariant::Regular,
        data: std::fs::read(format!("{dir}/test-files/fonts/DxLigaUnnamed.ttf")).expect("font"),
    });
    let pdf =
        dxpdf::render::render(doc, &dxpdf::render::RenderOptions::default()).expect("renders");
    let parsed = lopdf::Document::load_mem(&pdf).expect("a PDF");
    let maps_fi = parsed.objects.values().any(|object| {
        let Ok(stream) = object.as_stream() else {
            return false;
        };
        let content = stream
            .decompressed_content()
            .unwrap_or_else(|_| stream.content.clone());
        let text = String::from_utf8_lossy(&content).to_ascii_lowercase();
        // `f` `i` in UTF-16BE, as the destination of a glyph's mapping.
        text.contains("begincmap") && text.contains("<00660069>")
    });
    assert!(
        maps_fi,
        "no ToUnicode entry gives the f_i ligature its text"
    );
}
