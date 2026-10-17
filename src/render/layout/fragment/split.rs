//! Cluster-level splitting of over-wide text fragments.
//!
//! A word wider than the space available can't be broken at a normal break
//! opportunity, so it is split into one fragment per grapheme cluster and the
//! line-fitter breaks between them. Used for narrow table cells and deep
//! indents.
//!
//! The unit is the same one spacing is inserted between
//! ([`crate::render::spacing`]) — and for the same reason. Splitting per
//! *scalar*, which this module used to do, let the line-fitter break between a
//! letter and its own combining mark and carry the accent to the next line.

use std::rc::Rc;

use super::{BreakAfter, FontProps, Fragment, TextMetrics};
use crate::render::dimension::Pt;
use crate::render::spacing;

/// Text measurement callback. Structurally identical to
/// `paragraph::MeasureTextFn`, restated here so this module doesn't depend on
/// the paragraph layer — callers in either layer pass the same closures.
pub type MeasureFn<'a> = Option<&'a dyn Fn(&str, &FontProps) -> (Pt, TextMetrics)>;

/// Which fragments belong to a word too wide for `max_width` — and so are to
/// be cut — as a mask over `fragments`.
///
/// A word is what lies between two break opportunities: a run of text
/// fragments each ending [`BreakAfter::Prohibited`], and the one after them.
/// It spans as many fragments as `<w:r>` runs split it into, and Word cuts it
/// as one: `{#owner}{signature}` `{` `signature2}{/owner}` is three runs, each
/// narrower than its 159pt line and together 191pt, and Word fills the line to
/// `…{signature2}{/` before cutting (DogLicenseTemplate_old1415). Checked one
/// fragment at a time, none was cut, and the line broke at a run boundary.
///
/// "Too wide" is measured **without the word's trailing whitespace**, the same
/// width line fitting checks overflow against: trailing space is allowed to
/// hang past the margin, so a word whose visible text fits is not over-wide
/// however much space trails it. Documents lay out with runs of spaces — one
/// in `test-cases/` pads a header with 130 of them — and UAX #14 keeps such a
/// run in a single unit, because [LB7] forbids a break *before* a space.
/// Measuring the full width would send that unit here to be cut into 130
/// one-space fragments, and the visible word ahead of them would be dragged
/// onto its own line for want of room that was never needed.
///
/// [LB7]: https://www.unicode.org/reports/tr14/#LB7
fn oversized_words(fragments: &[Fragment], max_width: Pt) -> Vec<bool> {
    let mut cut = vec![false; fragments.len()];
    for (start, end, width) in text_words(fragments) {
        if width > max_width {
            cut[start..=end].fill(true);
        }
    }
    cut
}

/// The words of `fragments`, as `oversized_words` defines them: the first and
/// last fragment of each, and its width without trailing whitespace.
fn text_words(fragments: &[Fragment]) -> Vec<(usize, usize, Pt)> {
    let text_at = |i: usize| match fragments.get(i) {
        Some(Fragment::Text {
            width,
            trimmed_width,
            break_after,
            ..
        }) => Some((*width, *trimmed_width, *break_after)),
        _ => None,
    };
    let mut words = Vec::new();
    let mut start = 0;
    while start < fragments.len() {
        if text_at(start).is_none() {
            start += 1;
            continue;
        }
        let mut end = start;
        let mut width = Pt::ZERO;
        while let Some((w, trimmed, after)) = text_at(end) {
            if after == BreakAfter::Prohibited && text_at(end + 1).is_some() {
                width += w;
                end += 1;
            } else {
                width += trimmed;
                break;
            }
        }
        words.push((start, end, width));
        start = end + 1;
    }
    words
}

/// The widths of `fragments`' words in order: each word as `oversized_words`
/// defines it, and each inline picture, fraction or emoji as a word of its
/// own. The widest is the narrowest line the paragraph can be set on without
/// cutting anything, which Word grows an autofit table's column to
/// (`build::table`'s `cell_min_width`).
pub fn word_widths(fragments: &[Fragment]) -> Vec<Pt> {
    let mut words = text_words(fragments);
    let atoms = fragments.iter().enumerate().filter_map(|(i, f)| match f {
        Fragment::Image { .. } | Fragment::MathFraction { .. } | Fragment::Emoji { .. } => {
            Some((i, i, f.width()))
        }
        _ => None,
    });
    words.extend(atoms);
    words.sort_unstable_by_key(|&(start, _, _)| start);
    words.into_iter().map(|(_, _, w)| w).collect()
}

/// Whether `fragment` has more than one cluster to cut it into. The *cluster*
/// count is what matters: a single cluster cannot be split however many
/// scalars or bytes it occupies. An earlier byte-length (`text.len() > 1`)
/// spelling of this test disagreed with the split itself for any non-ASCII
/// single character — it reported "needs split", then split nothing, and the
/// caller paid for a full clone of the fragment vector. A scalar count has
/// the same disagreement for a one-cluster accented letter.
fn splittable(fragment: &Fragment) -> bool {
    matches!(fragment, Fragment::Text { text, .. } if spacing::unit_count(text) > 1)
}

/// Split text fragments wider than `max_width` into per-cluster fragments.
///
/// Returns `None` when nothing needs splitting — the common case. That lets a
/// caller holding an owned `Vec` keep it untouched, and a caller holding a
/// slice hand back `Cow::Borrowed`; neither pays for a copy on the fast path.
///
/// Per-cluster widths come from `measure` when one is supplied. Without a
/// measurer the fragment's width is divided evenly across its clusters,
/// which is only a positioning approximation — the total is preserved.
pub fn split_oversized_fragments(
    fragments: &[Fragment],
    max_width: Pt,
    measure: MeasureFn<'_>,
) -> Option<Vec<Fragment>> {
    // A non-positive budget can't be met by any split, and dividing by it
    // below would be meaningless.
    if max_width <= Pt::ZERO {
        return None;
    }
    let cut = oversized_words(fragments, max_width);
    let to_split = |i: usize| cut[i] && splittable(&fragments[i]);
    if !(0..fragments.len()).any(to_split) {
        return None;
    }

    let mut result = Vec::with_capacity(fragments.len());
    for (index, frag) in fragments.iter().enumerate() {
        let Fragment::Text {
            text,
            width,
            font,
            color,
            shading,
            border,
            metrics,
            hyperlink_url,
            baseline_offset,
            level,
            shaped,
            kerning,
            break_after,
            ..
        } = frag
        else {
            result.push(frag.clone());
            continue;
        };
        if !to_split(index) {
            result.push(frag.clone());
            continue;
        }

        let unit_count = spacing::unit_count(text);
        let per_unit_fallback = *width / unit_count as f32;
        for (i, unit) in spacing::units(text).enumerate() {
            let (w, unit_metrics) = match measure {
                Some(m) => m(unit, font),
                None => (per_unit_fallback, *metrics),
            };
            // §17.3.2.19: the pair this unit forms with the next, which the
            // whole word was kerned by. It places the next piece when the two
            // share a line and is dropped at a line end — no pair is formed
            // across a line break — exactly as between two words
            // (`fragment::shape`'s `kern_across_breaks`). Word cuts the
            // kerned word: DogLicenseReport's `${apcp_fee}` fits `p_fee}` on
            // its second line, which the unkerned pieces overflow.
            let pair = kerning
                .as_deref()
                .and_then(|k| k.get(i))
                .copied()
                .unwrap_or(0.0);
            result.push(Fragment::Text {
                // Kept, though it buys little: a single cluster shapes to its
                // isolated form, which is what the cmap would have given
                // anyway. Cutting a word here is what costs the joining, and
                // that is already this module's documented last resort — it
                // now costs a joining script more than it costs Latin.
                shaped: *shaped,
                kerning: (pair != 0.0).then(|| Rc::from([pair].as_slice())),
                // UAX #9: a cluster of the parent's text is at the parent's
                // level. Cutting a word never crosses a level boundary —
                // `fragment::bidi` has already split at every one of those, so
                // whatever reaches here is level-uniform by the time it does.
                level: *level,
                text: Rc::from(unit),
                font: font.clone(),
                color: *color,
                shading: *shading,
                border: *border,
                // Still prohibited, even though this split exists precisely to
                // let the fitter break here. The fitter breaks at the *last*
                // opportunity before an overflow, so calling each cluster one
                // would drag the head of the word up onto the previous line —
                // "Nicht gefunden" in a narrow cell becomes "Nicht ge" /
                // "funden" — when the word would have fitted whole on the next
                // line. With no opportunity to fall back to, the fitter breaks
                // immediately before whichever cluster overflows, which is only
                // reached once the word has a line to itself and still doesn't
                // fit. That is the order Word applies too: move the word down
                // first, cut it only if it still won't fit.
                //
                // The last cluster keeps the fragment's own status: what
                // follows the word is still where the next one may begin.
                break_after: if i + 1 == unit_count {
                    *break_after
                } else {
                    BreakAfter::Prohibited
                },
                width: w + Pt::new(pair),
                trimmed_width: w,
                metrics: unit_metrics,
                hyperlink_url: hyperlink_url.clone(),
                baseline_offset: *baseline_offset,
                text_offset: Pt::ZERO,
                // Per-cluster split of an over-wide word — not a mark.
                is_footnote_ref: false,
            });
        }
    }
    Some(result)
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::render::fonts::Toggle;
    use crate::render::resolve::color::RgbColor;

    fn text_frag(text: &str, width: f32) -> Fragment {
        Fragment::Text {
            shaped: None,
            kerning: None,
            level: crate::i18n::bidi::BidiLevel::LTR,
            text: Rc::from(text),
            break_after: super::super::fixture_break_after(text),
            font: Rc::new(FontProps {
                vertical: None,
                features: Default::default(),
                effects: Default::default(),
                rtl: crate::render::fonts::Toggle::Absent,
                family: Rc::from("Test"),
                size: Pt::new(12.0),
                bold: Toggle::Absent,
                italic: Toggle::Absent,
                underline: false,
                char_spacing: Pt::ZERO,
                text_scale: 1.0,
                underline_position: Pt::ZERO,
                underline_thickness: Pt::ZERO,
            }),
            color: RgbColor::BLACK,
            width: Pt::new(width),
            trimmed_width: Pt::new(width),
            metrics: TextMetrics {
                ascent: Pt::new(10.0),
                descent: Pt::new(4.0),
                leading: Pt::ZERO,
            },
            hyperlink_url: None,
            shading: None,
            border: None,
            baseline_offset: Pt::ZERO,
            text_offset: Pt::ZERO,
            is_footnote_ref: false,
        }
    }

    fn texts(frags: &[Fragment]) -> Vec<String> {
        frags
            .iter()
            .map(|f| match f {
                Fragment::Text { text, .. } => text.to_string(),
                _ => panic!("expected Text fragment"),
            })
            .collect()
    }

    #[test]
    fn splits_into_one_fragment_per_character() {
        // "ab" at 60pt is wider than max_width=20pt → two 30pt characters
        // (uniform fallback — no measurer provided).
        let frags = vec![text_frag("ab", 60.0)];
        let result = split_oversized_fragments(&frags, Pt::new(20.0), None).expect("splits");
        assert_eq!(
            texts(&result),
            ["a", "b"],
            "one fragment per character, in order"
        );
        for frag in &result {
            let Fragment::Text { width, .. } = frag else {
                unreachable!()
            };
            assert!((width.raw() - 30.0).abs() < 1e-4, "uniform fallback 60/2");
        }
    }

    #[test]
    fn measurer_supplies_per_character_widths() {
        // A measurer that gives 'w' a wider advance than 'i' — the uniform
        // fallback would give both the same width.
        let measure = |t: &str, _: &FontProps| {
            let w = if t == "w" { 40.0 } else { 10.0 };
            (
                Pt::new(w),
                TextMetrics {
                    ascent: Pt::new(9.0),
                    descent: Pt::new(3.0),
                    leading: Pt::ZERO,
                },
            )
        };
        let frags = vec![text_frag("wi", 50.0)];
        let result =
            split_oversized_fragments(&frags, Pt::new(20.0), Some(&measure)).expect("splits");
        let widths: Vec<f32> = result
            .iter()
            .map(|f| match f {
                Fragment::Text { width, .. } => width.raw(),
                _ => unreachable!(),
            })
            .collect();
        assert_eq!(widths, [40.0, 10.0], "measured, not evenly divided");
        // Metrics come from the measurer too, not the parent fragment.
        let Fragment::Text { metrics, .. } = &result[0] else {
            unreachable!()
        };
        assert_eq!(metrics.ascent.raw(), 9.0);
    }

    /// A word is what lies between two break opportunities, however many
    /// fragments — `<w:r>` runs — it spans. Word cuts `{#owner}{signature}`
    /// `{` `signature2}{/owner}`, three runs each narrower than its line, where
    /// the line runs out (DogLicenseTemplate_old1415), not at a run boundary.
    #[test]
    fn a_word_over_several_fragments_is_cut_as_one() {
        // "ab" | "cd " with no opportunity between: 40pt together, 20pt each,
        // on a 30pt line.
        let frags = vec![text_frag("ab", 20.0), text_frag("cd ", 20.0)];
        let result = split_oversized_fragments(&frags, Pt::new(30.0), None).expect("splits");
        assert_eq!(texts(&result), ["a", "b", "c", "d", " "]);
    }

    /// The opportunity after a cut word survives the cut: it is the word's,
    /// not any one cluster's, and it is still where the next word may start.
    #[test]
    fn a_cut_word_keeps_the_opportunity_after_it() {
        let frags = vec![
            text_frag("ab", 20.0),
            text_frag("cd ", 20.0),
            text_frag("ef", 10.0),
        ];
        let result = split_oversized_fragments(&frags, Pt::new(30.0), None).expect("splits");
        assert_eq!(texts(&result), ["a", "b", "c", "d", " ", "ef"]);
        let breaks: Vec<BreakAfter> = result
            .iter()
            .map(|f| match f {
                Fragment::Text { break_after, .. } => *break_after,
                _ => unreachable!(),
            })
            .collect();
        assert_eq!(
            breaks,
            [
                BreakAfter::Prohibited,
                BreakAfter::Prohibited,
                BreakAfter::Prohibited,
                BreakAfter::Prohibited,
                BreakAfter::Opportunity,
                BreakAfter::Prohibited,
            ]
        );
    }

    #[test]
    fn nothing_to_split_returns_none() {
        let frags = vec![text_frag("hi", 10.0)];
        assert!(split_oversized_fragments(&frags, Pt::new(100.0), None).is_none());
    }

    #[test]
    fn single_character_is_never_split() {
        // Wider than max_width, but one character — nothing to break.
        let frags = vec![text_frag("M", 200.0)];
        assert!(split_oversized_fragments(&frags, Pt::new(10.0), None).is_none());
    }

    /// The character count, not the byte length, decides splittability. A
    /// byte-length test reports "needs split" for these and then splits
    /// nothing, costing the caller a pointless clone of the whole vector.
    #[test]
    fn multi_byte_single_character_is_never_split() {
        for ch in ["é", "😀", "字"] {
            let frags = vec![text_frag(ch, 200.0)];
            assert!(
                split_oversized_fragments(&frags, Pt::new(10.0), None).is_none(),
                "{ch:?} is one character regardless of its byte length"
            );
        }
    }

    /// A unit whose *visible* text fits is not over-wide, however much space
    /// trails it. Real documents pad with runs of spaces, and UAX #14 keeps a
    /// run of them in one unit ([LB7] forbids a break before a space) — so
    /// this fragment reaches the splitter looking 200pt wide when only 5pt of
    /// it will ever be drawn inside the margin.
    ///
    /// [LB7]: https://www.unicode.org/reports/tr14/#LB7
    #[test]
    fn trailing_whitespace_does_not_make_a_fragment_over_wide() {
        let mut frag = text_frag("H                    ", 200.0);
        if let Fragment::Text { trimmed_width, .. } = &mut frag {
            *trimmed_width = Pt::new(5.0);
        }
        assert!(
            split_oversized_fragments(&[frag], Pt::new(20.0), None).is_none(),
            "only the visible width counts; the spaces hang past the margin",
        );
    }

    #[test]
    fn non_positive_max_width_returns_none() {
        let frags = vec![text_frag("ab", 60.0)];
        assert!(split_oversized_fragments(&frags, Pt::ZERO, None).is_none());
        assert!(split_oversized_fragments(&frags, Pt::new(-5.0), None).is_none());
    }

    #[test]
    fn fragments_that_fit_pass_through_unchanged() {
        // Two words: "ab " ends at an opportunity, so "ok" is not part of it.
        let frags = vec![
            text_frag("ab ", 60.0), // splits
            text_frag("ok", 5.0),   // fits — must survive whole
        ];
        let result = split_oversized_fragments(&frags, Pt::new(20.0), None).expect("splits");
        assert_eq!(texts(&result), ["a", "b", " ", "ok"]);
    }

    /// A combining mark must travel with its base. Splitting per scalar put the
    /// accent in its own fragment, which the line-fitter is then free to break
    /// before — carrying a bare `U+0301` to the start of the next line.
    #[test]
    fn split_never_separates_a_combining_mark_from_its_base() {
        let frags = vec![text_frag("e\u{301}x", 60.0)];
        let result = split_oversized_fragments(&frags, Pt::new(20.0), None).expect("splits");
        assert_eq!(
            texts(&result),
            ["e\u{301}", "x"],
            "the accent stays in the same fragment as its base letter"
        );
    }

    /// The cluster count decides splittability, so a single accented letter is
    /// as unsplittable as a single plain one — `needs_split` must not report
    /// work it cannot then do.
    #[test]
    fn multi_scalar_single_cluster_is_never_split() {
        for text in ["e\u{301}", "1\u{FE0F}\u{20E3}", "\u{1F1E9}\u{1F1EA}"] {
            let frags = vec![text_frag(text, 200.0)];
            assert!(
                split_oversized_fragments(&frags, Pt::new(10.0), None).is_none(),
                "{text:?} is one grapheme cluster regardless of its scalar count"
            );
        }
    }
}
