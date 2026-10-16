//! §17.3.2.19 `w:kern`: pair kerning from the font's TrueType `kern` table.
//!
//! # Not GPOS
//!
//! An OpenType font can kern twice over: in the legacy `kern` table and in the
//! `kern` feature of its `GPOS` table, and the two need not agree — Calibri's
//! do not. A shaper (HarfBuzz) reads `GPOS` when the font has one. **Word reads
//! the `kern` table.** Measured against `test-files/kerning.docx`: every word
//! of the kerned line comes out within 0.01pt of the `kern` table's pairs and
//! up to 2pt off GPOS's — `AVAWAY` at 11pt is 40.24pt by the `kern` table,
//! 38.27pt by GPOS, and 40.25pt in Word — and `K1`, which GPOS widens, Word
//! leaves alone as the `kern` table does.
//!
//! So kerning is applied here, from that table, rather than by turning the
//! shaper on: a kerned run stays on the cmap path, measures as the sum of its
//! advances and the pairs between them, and is painted unit by unit with the
//! same pairs (`DrawCommand::Text::kerning`). The pairs travel from layout to
//! paint rather than being looked up twice, because `render::subset` rewrites
//! the typeface in between and a subsetter is free to drop the legacy table.
//!
//! A font with no `kern` table is not kerned here. Whether Word kerns such a
//! font from GPOS instead has not been measured.

use fontcull_read_fonts::tables::kern::{Kern, SubtableKind};
use fontcull_read_fonts::types::GlyphId;
use fontcull_read_fonts::{FontData, FontRead};
use skia_safe::Typeface;

/// One typeface's `kern` table, kept as bytes and read on lookup — the parse is
/// a header read, and the pair search is a binary search over the bytes.
pub struct KernTable {
    data: Vec<u8>,
    units_per_em: f32,
}

impl KernTable {
    /// `typeface`'s `kern` table, or `None` when it has none — or one with no
    /// horizontal pairs this module can read.
    pub fn of(typeface: &Typeface) -> Option<Self> {
        let data = typeface.copy_table_data(u32::from_be_bytes(*b"kern"))?;
        let units_per_em = typeface.units_per_em().filter(|&u| u > 0)? as f32;
        let table = Self {
            data: data.as_bytes().to_vec(),
            units_per_em,
        };
        table.has_pairs().then_some(table)
    }

    fn has_pairs(&self) -> bool {
        let Ok(kern) = Kern::read(FontData::new(&self.data)) else {
            return false;
        };
        kern.subtables().flatten().any(|s| {
            s.is_horizontal()
                && !s.is_cross_stream()
                && matches!(
                    s.kind(),
                    Ok(SubtableKind::Format0(_)
                        | SubtableKind::Format2(_)
                        | SubtableKind::Format3(_))
                )
        })
    }

    /// The kerning between `left` and `right`, in em units of the typeface:
    /// the sum over the table's horizontal, in-stream subtables.
    pub fn pair(&self, left: u16, right: u16) -> f32 {
        let Ok(kern) = Kern::read(FontData::new(&self.data)) else {
            return 0.0;
        };
        let (l, r) = (GlyphId::new(left.into()), GlyphId::new(right.into()));
        let total: i32 = kern
            .subtables()
            .flatten()
            .filter(|s| s.is_horizontal() && !s.is_cross_stream())
            .filter_map(|s| match s.kind().ok()? {
                SubtableKind::Format0(t) => t.kerning(l, r),
                SubtableKind::Format2(t) => t.kerning(l, r),
                SubtableKind::Format3(t) => t.kerning(l, r),
                SubtableKind::Format1(_) => None,
            })
            .sum();
        total as f32
    }

    /// `pair` scaled to `size` points.
    pub fn pair_at(&self, left: u16, right: u16, size: f32) -> f32 {
        self.pair(left, right) * size / self.units_per_em
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use skia_safe::{FontMgr, FontStyle};

    /// A host font with a `kern` table, or `None` — CI images vary, so the
    /// tests that need one return early rather than fail.
    fn kerned_typeface() -> Option<(Typeface, KernTable)> {
        let mgr = FontMgr::new();
        ["Calibri", "Times New Roman", "Arial", "Georgia", "Verdana"]
            .iter()
            .filter_map(|f| mgr.match_family_style(f, FontStyle::normal()))
            .find_map(|tf| KernTable::of(&tf).map(|k| (tf, k)))
    }

    /// `AV` is the pair every Latin kern table kerns, and kerns tighter.
    #[test]
    fn a_v_is_kerned_together() {
        let Some((tf, kern)) = kerned_typeface() else {
            return;
        };
        let a = tf.unichar_to_glyph('A' as i32);
        let v = tf.unichar_to_glyph('V' as i32);
        assert!(kern.pair(a, v) < 0.0, "{} kerns AV apart", tf.family_name());
        assert!(
            (kern.pair_at(a, v, 20.0) - kern.pair_at(a, v, 10.0) * 2.0).abs() < 1e-4,
            "the pair scales with the size"
        );
    }

    /// A pair the table does not list is not kerned.
    #[test]
    fn an_unlisted_pair_is_zero() {
        let Some((tf, kern)) = kerned_typeface() else {
            return;
        };
        let x = tf.unichar_to_glyph('x' as i32);
        assert_eq!(kern.pair(x, x), 0.0);
    }
}
