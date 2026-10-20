//! §17.3.2.19 `w:kern`: pair kerning from the font's TrueType `kern` table, or
//! from its `GPOS` pair adjustments when it has no such table.
//!
//! # The `kern` table first
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
//! # `GPOS` when there is no `kern` table
//!
//! The open metric-compatible stand-ins dxpdf substitutes for Office fonts
//! (`render::fonts::substitutes`) do not all carry a `kern` table.
//! Carlito, the stand-in for Calibri, has none — its pairs are in `GPOS` only,
//! and they are Calibri's `kern` table, pair for pair (all 607 pairs between
//! printable ASCII characters agree to the unit). So a face with no `kern`
//! table is kerned from the pair adjustments of its `GPOS` `kern` feature, and
//! a Calibri document laid out in Carlito kerns as Word kerns it in Calibri.
//! Liberation Sans and Serif carry a `kern` table that equals Arial's and
//! Times New Roman's, and are read from it like the originals.
//!
//! Only what a `kern` table could say is read from `GPOS`: the pair-positioning
//! lookups (type 2, directly or through an extension) of the `kern` feature
//! under the `latn` script, else `DFLT`, and of each pair the first glyph's
//! advance adjustment, between **two glyphs the `cmap` maps a character to**.
//! A `kern` table pairs no other glyphs — not one of the 26,706 pairs in
//! Calibri's, nor of Calibri Bold's, Times New Roman's or Arial's, names a
//! glyph the `cmap` does not reach. Contextual kerning lookups are not applied.
//! Within a lookup the first subtable that applies to the pair decides it;
//! separate lookups add — as a shaper applies them.
//!
//! # Characters, not glyphs
//!
//! Word kerns the **characters** either side of a glyph boundary, so a
//! ligature is kerned by the letters it draws: its first against what precedes
//! it and its last against what follows. Measured across 118 real-world
//! templates' Word renders in Calibri: `.`·`t_i` is set −44 units (67 times),
//! `n`·`t_i` −19 and `p`·`t_i` −9 — Calibri's `.`·`t`, `n`·`t` and `p`·`t`
//! pairs — while `t_i`·`o` takes `i`·`o`'s 0, not `t`·`o`'s −20; and in
//! `ligatures-15.docx`, `a`·`ff` is set as `a`·`f` (−12). So every caller
//! looks a pair up through the `cmap` from characters, ligated run or not,
//! and the pair a ligature glyph would form in Carlito's `GPOS` (`n`·`t_i`,
//! −19 there and in Calibri's own `GPOS`) is never asked for. That pair
//! happens to equal `n`·`t`; Word's −19 before a merge tag's `ti` is what
//! fits it in its cell in a real-world monthly report template.
//!
//! What this cannot recover is a pair the original's `kern` table omits and the
//! stand-in's `GPOS` lists anyway. Carlito Regular has none among ASCII,
//! Latin-1, the dashes and the quotes; Carlito Bold kerns 51 there that
//! Calibri Bold's table leaves at zero (`oV`, `'A`, `-T`, …), and bold Calibri
//! text laid out in it is that much tighter than Word sets it.
//!
//! # Applied here, not by the shaper
//!
//! Kerning is applied here rather than by turning the shaper on: a kerned run
//! stays on the cmap path, measures as the sum of its advances and the pairs
//! between them, and is painted unit by unit with the same pairs
//! (`DrawCommand::Text::kerning`). The pairs travel from layout to paint rather
//! than being looked up twice, because `render::subset` rewrites the typeface in
//! between and a subsetter is free to drop either table.
//!
//! Whether Word kerns a font of its own that has no `kern` table from `GPOS`
//! has not been measured; no Office font a corpus template uses is one.

use std::cell::RefCell;

use fontcull_read_fonts::tables::cmap::{Cmap, CmapIterLimits};
use fontcull_read_fonts::tables::gpos::{Gpos, PairPos, PositionSubtables, ValueContext};
use fontcull_read_fonts::tables::kern::{Kern, SubtableKind};
use fontcull_read_fonts::types::{GlyphId, GlyphId16, Tag};
use fontcull_read_fonts::{FontData, FontRead};
use rustc_hash::FxHashMap;
use skia_safe::Typeface;

/// One typeface's pair kerning, from whichever table it comes from. Kept as
/// the table's bytes and read on lookup.
pub struct PairKerning {
    source: Source,
    units_per_em: f32,
}

enum Source {
    /// The legacy `kern` table — the parse is a header read, and the pair
    /// search a binary search over the bytes.
    Legacy(Vec<u8>),
    /// `GPOS`, the indices of its `kern` feature's pair lookups in
    /// lookup-list order, and which glyphs the `cmap` reaches — the only ones
    /// a pair is read between. A pair walks several subtables, so each is
    /// looked up once and remembered.
    Gpos {
        data: Vec<u8>,
        lookups: Vec<u16>,
        mapped: Vec<bool>,
        pairs: RefCell<FxHashMap<(u16, u16), i32>>,
    },
}

const KERN_TABLE: u32 = u32::from_be_bytes(*b"kern");
const GPOS_TABLE: u32 = u32::from_be_bytes(*b"GPOS");
const CMAP_TABLE: u32 = u32::from_be_bytes(*b"cmap");
const KERN_FEATURE: Tag = Tag::new(b"kern");
/// Scripts searched for the `kern` feature, in order. Document text dxpdf
/// kerns is Latin; a font with no `latn` script lists its pairs under `DFLT`.
const SCRIPTS: [Tag; 2] = [Tag::new(b"latn"), Tag::new(b"DFLT")];

impl PairKerning {
    /// `typeface`'s pair kerning: its `kern` table when it has one with
    /// horizontal pairs this module can read, else its `GPOS` `kern` feature's
    /// pair lookups, else `None`.
    pub fn of(typeface: &Typeface) -> Option<Self> {
        let units_per_em = typeface.units_per_em().filter(|&u| u > 0)? as f32;
        let legacy = typeface
            .copy_table_data(KERN_TABLE)
            .map(|d| d.as_bytes().to_vec())
            .filter(|d| legacy_has_pairs(d));
        let source = match legacy {
            Some(data) => Source::Legacy(data),
            None => {
                let data = typeface.copy_table_data(GPOS_TABLE)?.as_bytes().to_vec();
                let lookups = gpos_kern_lookups(&data);
                if lookups.is_empty() {
                    return None;
                }
                Source::Gpos {
                    data,
                    lookups,
                    mapped: mapped_glyphs(typeface),
                    pairs: RefCell::default(),
                }
            }
        };
        Some(Self {
            source,
            units_per_em,
        })
    }

    /// The kerning between `left` and `right`, in em units of the typeface.
    pub fn pair(&self, left: u16, right: u16) -> f32 {
        match &self.source {
            Source::Legacy(data) => legacy_pair(data, left, right) as f32,
            Source::Gpos {
                data,
                lookups,
                mapped,
                pairs,
            } => {
                let reached = |g: u16| mapped.get(g as usize).copied().unwrap_or(false);
                if !(reached(left) && reached(right)) {
                    return 0.0;
                }
                *pairs
                    .borrow_mut()
                    .entry((left, right))
                    .or_insert_with(|| gpos_pair(data, lookups, left, right)) as f32
            }
        }
    }

    /// `pair` scaled to `size` points.
    pub fn pair_at(&self, left: u16, right: u16, size: f32) -> f32 {
        self.pair(left, right) * size / self.units_per_em
    }
}

fn legacy_has_pairs(data: &[u8]) -> bool {
    let Ok(kern) = Kern::read(FontData::new(data)) else {
        return false;
    };
    kern.subtables().flatten().any(|s| {
        s.is_horizontal()
            && !s.is_cross_stream()
            && matches!(
                s.kind(),
                Ok(SubtableKind::Format0(_) | SubtableKind::Format2(_) | SubtableKind::Format3(_))
            )
    })
}

/// The sum over the `kern` table's horizontal, in-stream subtables.
fn legacy_pair(data: &[u8], left: u16, right: u16) -> i32 {
    let Ok(kern) = Kern::read(FontData::new(data)) else {
        return 0;
    };
    let (l, r) = (GlyphId::new(left.into()), GlyphId::new(right.into()));
    kern.subtables()
        .flatten()
        .filter(|s| s.is_horizontal() && !s.is_cross_stream())
        .filter_map(|s| match s.kind().ok()? {
            SubtableKind::Format0(t) => t.kerning(l, r),
            SubtableKind::Format2(t) => t.kerning(l, r),
            SubtableKind::Format3(t) => t.kerning(l, r),
            SubtableKind::Format1(_) => None,
        })
        .sum()
}

/// Which of `typeface`'s glyphs its `cmap` maps a character to, by glyph id.
fn mapped_glyphs(typeface: &Typeface) -> Vec<bool> {
    let count = typeface.count_glyphs();
    let mut mapped = vec![false; count];
    let Some(data) = typeface.copy_table_data(CMAP_TABLE) else {
        return mapped;
    };
    let Ok(cmap) = Cmap::read(FontData::new(data.as_bytes())) else {
        return mapped;
    };
    if let Some((_, _, subtable)) = cmap.best_subtable() {
        let limits = CmapIterLimits {
            max_char: char::MAX as u32,
            glyph_count: count as u32,
        };
        for (_, glyph) in subtable.iter_with_limits(limits) {
            if let Some(slot) = mapped.get_mut(glyph.to_u32() as usize) {
                *slot = true;
            }
        }
    }
    mapped
}

/// The pair-positioning lookups of `GPOS`'s `kern` feature, under the first of
/// [`SCRIPTS`] the font lists, in lookup-list order.
fn gpos_kern_lookups(data: &[u8]) -> Vec<u16> {
    let Ok(gpos) = Gpos::read(FontData::new(data)) else {
        return Vec::new();
    };
    let (Ok(scripts), Ok(features), Ok(lookup_list)) =
        (gpos.script_list(), gpos.feature_list(), gpos.lookup_list())
    else {
        return Vec::new();
    };
    let Some(lang_sys) = SCRIPTS.iter().find_map(|tag| {
        let record = scripts
            .script_records()
            .iter()
            .find(|r| r.script_tag() == *tag)?;
        record
            .script(scripts.offset_data())
            .ok()?
            .default_lang_sys()?
            .ok()
    }) else {
        return Vec::new();
    };
    let records = features.feature_records();
    let mut lookups: Vec<u16> = lang_sys
        .feature_indices()
        .iter()
        .filter_map(|i| records.get(i.get() as usize))
        .filter(|r| r.feature_tag() == KERN_FEATURE)
        .filter_map(|r| r.feature(features.offset_data()).ok())
        .flat_map(|f| f.lookup_list_indices().iter().map(|i| i.get()))
        .collect();
    lookups.sort_unstable();
    lookups.dedup();
    let lookup_table = lookup_list.lookups();
    lookups.retain(|&i| {
        lookup_table
            .get(i as usize)
            .ok()
            .and_then(|l| l.subtables().ok())
            .is_some_and(|s| matches!(s, PositionSubtables::Pair(_)))
    });
    lookups
}

/// `left`·`right`'s advance adjustment over `lookups`: in each lookup, the
/// first subtable that applies to the pair; summed over the lookups.
fn gpos_pair(data: &[u8], lookups: &[u16], left: u16, right: u16) -> i32 {
    let Ok(gpos) = Gpos::read(FontData::new(data)) else {
        return 0;
    };
    let Ok(lookup_list) = gpos.lookup_list() else {
        return 0;
    };
    let lookup_table = lookup_list.lookups();
    lookups
        .iter()
        .filter_map(
            |&i| match lookup_table.get(i as usize).ok()?.subtables().ok()? {
                PositionSubtables::Pair(subtables) => subtables
                    .iter()
                    .flatten()
                    .find_map(|subtable| pair_in_subtable(&subtable, left, right)),
                _ => None,
            },
        )
        .map(i32::from)
        .sum()
}

/// The first glyph's x-advance adjustment when `subtable` applies to
/// `left`·`right` — zero when it applies and adjusts nothing — or `None` when
/// it does not apply and the next subtable is asked.
fn pair_in_subtable(subtable: &PairPos, left: u16, right: u16) -> Option<i16> {
    let (l, r) = (GlyphId16::new(left), GlyphId16::new(right));
    match subtable {
        PairPos::Format1(t) => {
            let index = t.coverage().ok()?.get(l)?;
            let set = t.pair_sets().get(index as usize).ok()?;
            let records = set.pair_value_records();
            // Records are sorted by the second glyph.
            let (mut lo, mut hi) = (0, records.len());
            while lo < hi {
                let mid = (lo + hi) / 2;
                let record = records.get(mid).ok()?;
                match record.second_glyph().cmp(&r) {
                    std::cmp::Ordering::Less => lo = mid + 1,
                    std::cmp::Ordering::Greater => hi = mid,
                    std::cmp::Ordering::Equal => {
                        return Some(record.value_record1().x_advance().unwrap_or(0));
                    }
                }
            }
            None
        }
        PairPos::Format2(t) => {
            t.coverage().ok()?.get(l)?;
            let class1 = t.class_def1().ok()?.get(l);
            let class2 = t.class_def2().ok()?.get(r);
            if class1 >= t.class1_count() || class2 >= t.class2_count() {
                return None;
            }
            let [first, _] = t.values(class1, class2, &ValueContext::default()).ok()?;
            Some(first.x_advance)
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use skia_safe::{FontMgr, FontStyle};

    /// A host font with a `kern` table, or `None` — CI images vary, so the
    /// tests that need one return early rather than fail.
    fn kerned_typeface() -> Option<(Typeface, PairKerning)> {
        let mgr = FontMgr::new();
        ["Calibri", "Times New Roman", "Arial", "Georgia", "Verdana"]
            .iter()
            .filter_map(|f| mgr.match_family_style(f, FontStyle::normal()))
            .find_map(|tf| PairKerning::of(&tf).map(|k| (tf, k)))
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

    fn fixture(name: &str) -> Typeface {
        let path = format!("{}/test-files/fonts/{name}", env!("CARGO_MANIFEST_DIR"));
        let data = skia_safe::Data::new_copy(&std::fs::read(path).expect("fixture font"));
        FontMgr::new()
            .new_from_data(&data, None)
            .expect("fixture parses")
    }

    fn pair_of(tf: &Typeface, kern: &PairKerning, a: char, b: char) -> f32 {
        kern.pair(tf.unichar_to_glyph(a as i32), tf.unichar_to_glyph(b as i32))
    }

    /// A face with both tables is kerned from the `kern` table: `DxKern`'s
    /// table says A·V −80 where its `GPOS` says −200.
    #[test]
    fn the_kern_table_wins_over_gpos() {
        let tf = fixture("DxKern.ttf");
        let kern = PairKerning::of(&tf).expect("kerned");
        assert_eq!(pair_of(&tf, &kern, 'A', 'V'), -80.0);
        assert_eq!(pair_of(&tf, &kern, 'K', '1'), 0.0, "GPOS's +40 is not read");
    }

    /// A face with no `kern` table is kerned from `GPOS`: glyph pairs (format
    /// 1), class pairs (format 2) and a pair in an extension lookup.
    /// `DxGposKern` carries `DxKern`'s table pairs that way.
    #[test]
    fn a_face_without_a_kern_table_is_kerned_from_gpos() {
        let tf = fixture("DxGposKern.ttf");
        let kern = PairKerning::of(&tf).expect("kerned from GPOS");
        assert_eq!(pair_of(&tf, &kern, 'A', 'V'), -80.0, "a glyph pair");
        assert_eq!(pair_of(&tf, &kern, 'T', 'y'), -70.0, "a class pair");
        assert_eq!(pair_of(&tf, &kern, 'A', 'Y'), -100.0, "an extension lookup");
        assert_eq!(pair_of(&tf, &kern, 'x', 'x'), 0.0, "an unlisted pair");
    }

    /// A `GPOS` pair is read only between two glyphs the `cmap` reaches:
    /// `DxGposKern` kerns its unmapped `A.alt` against `V` at −80, as it does
    /// `A`, and that pair is not read.
    #[test]
    fn a_gpos_pair_with_an_unmapped_glyph_is_not_read() {
        let tf = fixture("DxGposKern.ttf");
        let kern = PairKerning::of(&tf).expect("kerned from GPOS");
        // `A.alt` is the fixture's last glyph.
        let alt = (tf.count_glyphs() - 1) as u16;
        let v = tf.unichar_to_glyph('V' as i32);
        assert_eq!(kern.pair(alt, v), 0.0);
        assert_eq!(
            pair_of(&tf, &kern, 'A', 'V'),
            -80.0,
            "the mapped A still kerns"
        );
    }

    /// Within a lookup the first subtable that applies decides the pair: T·o
    /// is a glyph pair at −90 ahead of a class pair that would give −70.
    #[test]
    fn the_first_subtable_that_applies_decides_the_pair() {
        let tf = fixture("DxGposKern.ttf");
        let kern = PairKerning::of(&tf).expect("kerned from GPOS");
        assert_eq!(pair_of(&tf, &kern, 'T', 'o'), -90.0);
    }
}
