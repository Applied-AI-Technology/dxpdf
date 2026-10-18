//! Line fitting — break fragments into lines that fit within a max width.

use crate::render::dimension::Pt;
use crate::render::layout::fragment::{BreakAfter, Fragment};

/// A fitted line — a slice of fragments that fit within the available width.
#[derive(Debug)]
pub struct FittedLine {
    /// Indices into the fragment list: [start, end).
    pub start: usize,
    pub end: usize,
    /// Total width of all fragments in this line.
    pub width: Pt,
    /// The line's natural height — `LineMetrics`: the most any fragment puts
    /// above the baseline plus the most any puts below it.
    pub height: Pt,
    /// The same for text alone. §17.3.1.33: Auto line spacing multiplies text
    /// metrics, not inline image heights.
    pub text_height: Pt,
    /// How much further below the baseline than the text something unscaled
    /// on the line reaches — the part of [`Self::height`] an Auto multiple's
    /// extra, which hangs below the text, falls inside
    /// (`paragraph::line_emit::resolve_line_height`). Zero on a line where
    /// nothing reaches below its text.
    pub overhang: Pt,
    /// How far the baseline sits below the top of the natural line: the most
    /// any fragment puts above it.
    pub ascent: Pt,
    /// Whether this line ends with an explicit line break.
    pub has_break: bool,
}

/// Break fragments into lines that fit within `max_width`.
///
/// When a line overflows, it breaks at the last fragment that reported a UAX
/// #14 break opportunity ([`BreakAfter::Opportunity`]). A single fragment wider
/// than `max_width` gets its own line (no infinite loop); over-wide *text* is
/// cut into grapheme clusters upstream by
/// [`split_oversized_fragments`](crate::render::layout::fragment::split_oversized_fragments)
/// before it reaches here.
///
/// `first_line_width`: if provided, the first line uses this narrower width
/// (e.g., to account for first-line indent). Subsequent lines use `max_width`.
pub fn fit_lines(fragments: &[Fragment], max_width: Pt) -> Vec<FittedLine> {
    fit_lines_with_first(
        fragments,
        max_width,
        max_width,
        crate::render::layout::paragraph::PTabGeometry {
            max_width,
            indent_left: Pt::ZERO,
            indent_first_line: Pt::ZERO,
            content_width: max_width,
            float_left: Pt::ZERO,
            float_right: Pt::ZERO,
        },
        false,
    )
}

/// Line fitting with separate first-line and remaining-line widths.
///
/// `ptab_geometry` is the paragraph geometry §17.3.1.30 position tabs resolve
/// against. Fitting needs it because a tab whose alignment point lies behind
/// the pen advances to the next line, and only fitting can create one.
///
/// `bidi_rtl`: whether the paragraph's base direction is right-to-left.
/// Emission mirrors a tab-bearing RTL line (§17.3.1.37 under `w:bidi`) and, in
/// that mirrored frame, swaps which physical side each float narrows before
/// resolving a `<w:ptab>` — fitting has to swap the same way, or the two can
/// reach different Placed/AdvancesToNextLine verdicts for the same tab. A
/// line's tab-placement precondition for mirroring is already guaranteed
/// once fitting has *reached* a `PTab` fragment — whatever line ends up
/// holding it will satisfy `w:bidi`'s other half — so `bidi_rtl` alone is the
/// right test here, without also threading a per-line "has a tab" flag
/// fitting cannot know in advance.
pub fn fit_lines_with_first(
    fragments: &[Fragment],
    first_line_width: Pt,
    remaining_width: Pt,
    ptab_geometry: crate::render::layout::paragraph::PTabGeometry,
    bidi_rtl: bool,
) -> Vec<FittedLine> {
    if fragments.is_empty() {
        return Vec::new();
    }

    let mut lines = Vec::new();
    let mut line_start = 0;
    let mut line_width = Pt::ZERO;
    // §17.3.1.30: where the pen actually *is*, as opposed to how much width
    // the line has accumulated. The two differ only across a position tab,
    // which jumps the pen to its anchor while contributing a nominal width.
    // Tracking it separately keeps `line_width` — and therefore every
    // non-ptab paragraph's fitting — bit-for-bit unchanged.
    let line_pen_start = |is_first_line: bool| {
        ptab_geometry.indent_left
            + if is_first_line {
                ptab_geometry.indent_first_line
            } else {
                Pt::ZERO
            }
    };
    let mut pen_x = line_pen_start(true);
    // §17.3.1.30: `relativeTo="margin"` measures against the full text area,
    // so once such a tab has placed content this line may legitimately use the
    // space a paragraph's own right indent excludes. Until then the ordinary
    // `content_width` bound applies.
    let mut margin_span_active = false;
    let mut last_break_point = None; // index after which we can break

    let mut i = 0;
    while i < fragments.len() {
        let frag = &fragments[i];

        // Explicit line break — emit current line including the break fragment.
        if frag.is_line_break() {
            let metrics = range_metrics(fragments, line_start, i + 1);
            lines.push(FittedLine {
                start: line_start,
                end: i + 1,
                width: line_width,
                height: metrics.height(),
                text_height: metrics.text_height(),
                overhang: metrics.overhang(),
                ascent: metrics.above,
                has_break: true,
            });
            line_start = i + 1;
            line_width = Pt::ZERO;
            pen_x = line_pen_start(lines.is_empty());
            margin_span_active = false;
            last_break_point = None;
            i += 1;
            continue;
        }

        // §17.3.1.30: a position tab whose alignment point lies behind the
        // pen advances to that point on the *next* line. Decided here rather
        // than at emission because it is a line break, and emission cannot
        // create one.
        if let Fragment::PTab {
            align, relative_to, ..
        } = frag
        {
            // Fitting cannot bound the zone by the line end the way emission
            // does — the lines do not exist yet — so it scans to the end of the
            // fragment list. Where the two disagree (a zone fitting then splits
            // for width) emission is authoritative for the final x; this only
            // decides whether the tab can be honoured on this line.
            let end = crate::render::layout::paragraph::zone_end(fragments, i, fragments.len());
            // Emission mirrors a tab-bearing RTL line and swaps which
            // physical side each float narrows before resolving this same
            // tab (`line_emit.rs`'s `Fragment::PTab` arm) — matched here so
            // the two agree on whether the anchor is behind the pen. See
            // `bidi_rtl`'s doc for why that flag alone is the right test.
            let (geo_float_left, geo_float_right) = if bidi_rtl {
                (ptab_geometry.float_right, ptab_geometry.float_left)
            } else {
                (ptab_geometry.float_left, ptab_geometry.float_right)
            };
            let geometry = crate::render::layout::paragraph::PTabGeometry {
                float_left: geo_float_left,
                float_right: geo_float_right,
                ..ptab_geometry
            };
            let placement = crate::render::layout::paragraph::resolve_ptab(
                *align,
                *relative_to,
                geometry,
                pen_x,
                || crate::render::layout::paragraph::zone_width(fragments, i + 1, end),
            );
            match placement {
                crate::render::layout::paragraph::PTabPlacement::Placed(at) => {
                    pen_x = at;
                    margin_span_active |=
                        matches!(relative_to, crate::model::PTabRelativeTo::Margin);
                }
                crate::render::layout::paragraph::PTabPlacement::AdvancesToNextLine { .. } => {
                    // Only break when doing so can help. A tab already first
                    // on its line would find the same anchor behind the same
                    // pen on the next one — acting on a condition the action
                    // cannot change is how this engine's pagination loops have
                    // historically become infinite.
                    if line_start < i {
                        let m = measure_range(fragments, line_start, i);
                        lines.push(FittedLine {
                            start: line_start,
                            end: i,
                            width: m.width,
                            height: m.height,
                            text_height: m.text_height,
                            overhang: m.overhang,
                            ascent: m.ascent,
                            has_break: false,
                        });
                        line_start = i;
                        line_width = Pt::ZERO;
                        pen_x = line_pen_start(lines.is_empty());
                        margin_span_active = false;
                        last_break_point = None;
                        // Re-evaluate this tab against the fresh line.
                        continue;
                    }
                }
            }
        }

        let frag_width = frag.width();
        let new_width = line_width + frag_width;

        // Use first-line width for line 0, remaining width for subsequent lines.
        let current_max = if lines.is_empty() {
            first_line_width
        } else {
            remaining_width
        };

        // For overflow checking, use trimmed width — trailing whitespace on the
        // last word is allowed to hang past the margin (standard Word behavior).
        // The check uses: previous fragments' full widths + this fragment's trimmed width.
        let check_width = line_width + frag.trimmed_width();

        // Check if adding this fragment overflows. Once a margin-relative tab
        // has placed content, the line's real right edge is the margin, and
        // the pen — not the accumulated width sum — says where we are.
        let overflows = if margin_span_active {
            pen_x + frag.trimmed_width() > ptab_geometry.max_width
        } else {
            check_width > current_max
        };
        if overflows && line_start < i {
            // Overflow — break at last break point, or before this fragment.
            let break_at = last_break_point.unwrap_or(i);
            let m = measure_range(fragments, line_start, break_at);
            lines.push(FittedLine {
                start: line_start,
                end: break_at,
                width: m.width,
                height: m.height,
                text_height: m.text_height,
                overhang: m.overhang,
                ascent: m.ascent,
                has_break: false,
            });
            line_start = break_at;
            line_width = Pt::ZERO;
            pen_x = line_pen_start(lines.is_empty());
            margin_span_active = false;
            last_break_point = None;
            // Resume *at the break*, not at the fragment that overflowed.
            //
            // `break_at` is usually `i` — no earlier opportunity, so the new
            // line starts with the fragment that did not fit and this rewinds
            // nothing. When there *was* an earlier opportunity it can be far
            // behind, and then everything between it and `i` belongs to the
            // new line and has to be measured onto it. Resuming at `i` left
            // that span inside the line's `[start, end)` range — so it painted
            // — while contributing nothing to the width, after which the line
            // could not overflow again and swallowed the rest of the
            // paragraph. That is how a Windows path in a 167.80 pt footer cell
            // became one 295 pt line running 91 pt past the edge of the page.
            //
            // Termination is unchanged: `line_start` still only ever moves
            // forward, and the next pass has `i == line_start`, where the
            // `line_start < i` guard sends an over-wide fragment down the
            // "first on the line, allow it" path instead of breaking again.
            i = break_at;
            continue;
        }

        // If this is the first fragment on the line and it overflows,
        // allow it (it will be the only fragment on this line). The
        // paragraph renderer will clip/overflow as needed.
        line_width = new_width;
        // A position tab has already jumped the pen to its anchor; every other
        // fragment advances it by its own width. Done here, after the overflow
        // check, so the check sees the pen *at* this fragment rather than past it.
        if !matches!(frag, Fragment::PTab { .. }) {
            pen_x += frag_width;
        }

        // Track break opportunity. A text fragment carries UAX #14's answer
        // (`crate::i18n::segment` computed it over the whole paragraph text,
        // across `<w:r>` boundaries); this used to re-derive it here by
        // sniffing the fragment's last character, against a list that had
        // already drifted from the one used to cut the fragments.
        let is_break_point = match frag {
            Fragment::Text { break_after, .. } | Fragment::MathFraction { break_after, .. } => {
                *break_after == BreakAfter::Opportunity
            }
            _ => true, // tabs, images, line breaks are always break points
        };
        if is_break_point {
            last_break_point = Some(i + 1);
        }

        i += 1;
    }

    // Emit remaining fragments as the last line.
    if line_start < fragments.len() {
        let metrics = range_metrics(fragments, line_start, fragments.len());
        lines.push(FittedLine {
            start: line_start,
            end: fragments.len(),
            width: line_width,
            height: metrics.height(),
            text_height: metrics.text_height(),
            overhang: metrics.overhang(),
            ascent: metrics.above,
            has_break: false,
        });
    }

    lines
}

/// What a line's fragments put above and below its baseline.
///
/// A line's box is the most anything on it reaches above the baseline plus the
/// most anything reaches below. Where one run reaches farthest both ways that
/// is its own line, which is every mixed line Word has been measured on here (a
/// 22pt run beside an 11pt one); where two runs split the reach it is taller
/// than either, as a line box is in any layout that aligns on a baseline.
/// Measured in Word (Word for Mac 16, 2026-10-16) with every height rule:
///
/// * A run reaches its font's line gap and ascent above the baseline and its
///   descent below. The line gap (`hhea` `lineGap`, which Skia reports as
///   leading) is **above** the text — a 22pt Arial line's first baseline sits
///   0.72pt lower than its ascent alone puts it, and nothing is added below.
/// * An inline object stands **on** the baseline and reaches its own height
///   above it, so text beside a 40pt picture sits on the picture's foot and the
///   line is the picture and the text's descent. A picture alone on its line
///   is exactly its own height.
/// * A break reaches nothing on a line that holds anything else. Its run's
///   font, however large, leaves an Arial line as tall as Arial makes it; a
///   line holding only the break is that font's line
///   (`test-files/break-clear.docx` case 8).
#[derive(Clone, Copy, Debug, Default)]
struct LineMetrics {
    /// The most any fragment reaches above the baseline.
    above: Pt,
    /// The most any fragment reaches below it.
    below: Pt,
    /// The same two for text alone — the base §17.3.1.33 Auto spacing scales.
    text_above: Pt,
    text_below: Pt,
    /// Fragments that carry only a height (tabs), counted toward both.
    other: Pt,
    /// A break's own line height, which only a line holding nothing else takes.
    line_break: Pt,
}

impl LineMetrics {
    fn add(&mut self, frag: &Fragment) {
        match frag {
            // A run whose line comes from a face other than the one drawing it
            // — a list label in Symbol or Wingdings — reaches its own height,
            // as a picture does, and is not the text an Auto multiple scales:
            // Word sets a 22pt Symbol bullet beside 11pt Calibri at 1.15 on
            // Symbol's reach above and Calibri's scaled line below
            // (`test-files/symbol-bullet.docx`). Measured for such labels only.
            Fragment::Text { metrics, font, .. } if font.vertical.is_some() => {
                self.above = self.above.max(metrics.leading + metrics.ascent);
                self.below = self.below.max(metrics.descent);
            }
            // The tab after such a label is set in its font and stands on its
            // line, whose reach the label has already given.
            Fragment::Tab { font, .. } if font.vertical.is_some() => {}
            Fragment::Text { metrics, .. } | Fragment::MathFraction { metrics, .. } => {
                self.add_text(metrics)
            }
            Fragment::Emoji { line_metrics, .. } => self.add_text(line_metrics),
            Fragment::Image { size, .. } => self.above = self.above.max(size.height),
            f if f.is_line_break() => self.line_break = self.line_break.max(f.height()),
            f => self.other = self.other.max(f.height()),
        }
    }

    fn add_text(&mut self, m: &crate::render::layout::fragment::TextMetrics) {
        let above = m.leading + m.ascent;
        self.above = self.above.max(above);
        self.below = self.below.max(m.descent);
        self.text_above = self.text_above.max(above);
        self.text_below = self.text_below.max(m.descent);
    }

    /// The break's height, where nothing visible stands on the line.
    fn bare_break(&self) -> Pt {
        if self.above > Pt::ZERO || self.below > Pt::ZERO {
            Pt::ZERO
        } else {
            self.line_break
        }
    }

    fn height(&self) -> Pt {
        (self.above + self.below)
            .max(self.other)
            .max(self.bare_break())
    }

    /// How far below the text's own descent anything else on the line reaches.
    fn overhang(&self) -> Pt {
        (self.below - self.text_below).max(Pt::ZERO)
    }

    fn text_height(&self) -> Pt {
        (self.text_above + self.text_below)
            .max(self.other)
            .max(self.bare_break())
    }
}

/// Measurements for a range of fragments.
struct RangeMeasure {
    width: Pt,
    height: Pt,
    text_height: Pt,
    overhang: Pt,
    ascent: Pt,
}

/// The [`LineMetrics`] of the line `fragments[start..end]`.
///
/// Spaces and tabs are left out wherever they stand: Word does not stand a
/// line taller for white space in a larger face — not a 20pt space between
/// two 10pt words, at the start of the line or at its end, nor an
/// unformatted 11pt space or tab in 10pt text (`test-files/space-run-height.docx`;
/// holdenma's delinquency notice, ShellfishTransactionReport's
/// `{ orderPaidDate}`). A no-break space is not white space to it and counts.
/// A line of nothing but white space keeps it, having nothing else to stand
/// on.
fn range_metrics(fragments: &[Fragment], start: usize, end: usize) -> LineMetrics {
    let line = &fragments[start..end];
    let blank = |f: &Fragment| match f {
        Fragment::Text { text, .. } => text.chars().all(|c| c == ' '),
        Fragment::Tab { .. } => true,
        _ => false,
    };
    let has_content = line.iter().any(|f| !blank(f) && !f.is_line_break());
    let mut metrics = LineMetrics::default();
    for frag in line {
        if has_content && blank(frag) {
            continue;
        }
        metrics.add(frag);
    }
    metrics
}

/// Measure total width and the [`LineMetrics`] of a range of fragments.
fn measure_range(fragments: &[Fragment], start: usize, end: usize) -> RangeMeasure {
    let width = fragments[start..end].iter().map(Fragment::width).sum();
    let metrics = range_metrics(fragments, start, end);
    RangeMeasure {
        width,
        height: metrics.height(),
        text_height: metrics.text_height(),
        overhang: metrics.overhang(),
        ascent: metrics.above,
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::render::fonts::Toggle;
    use crate::render::layout::fragment::{FontProps, TextMetrics};
    use crate::render::resolve::color::RgbColor;
    use std::rc::Rc;

    /// A fragment the fitter may break *after* — what a word followed by a
    /// space becomes once `crate::i18n::segment` has looked at it.
    fn text_frag(text: &str, width: f32) -> Fragment {
        frag(text, width, BreakAfter::Opportunity)
    }

    /// A fragment glued to whatever follows it: a token cut by a `<w:r>`
    /// boundary, or one UAX #14 refuses to break inside.
    fn glued_frag(text: &str, width: f32) -> Fragment {
        frag(text, width, BreakAfter::Prohibited)
    }

    /// A `Fragment::MathFraction` with a caller-chosen `break_after` — e.g.
    /// `(1/2)²`, where the fraction is a superscript's base and must glue to
    /// the exponent that follows it (`fragment::math`'s `emit_elements` sets
    /// `Prohibited` on exactly this fragment for that case).
    fn math_fraction_frag(width: f32, break_after: BreakAfter) -> Fragment {
        let row = |text: &str| crate::render::layout::fragment::MathRow {
            text: text.into(),
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
            width: Pt::new(width * 0.5),
            metrics: TextMetrics {
                ascent: Pt::new(10.0),
                descent: Pt::new(4.0),
                leading: Pt::ZERO,
            },
        };
        Fragment::MathFraction {
            num: row("1"),
            den: row("2"),
            color: RgbColor::BLACK,
            width: Pt::new(width),
            metrics: TextMetrics {
                ascent: Pt::new(10.0),
                descent: Pt::new(4.0),
                leading: Pt::ZERO,
            },
            baseline_offset: Pt::ZERO,
            break_after,
            hyperlink_url: None,
        }
    }

    fn frag(text: &str, width: f32, break_after: BreakAfter) -> Fragment {
        Fragment::Text {
            shaped: None,
            kerning: None,
            level: crate::i18n::bidi::BidiLevel::LTR,
            text: text.into(),
            break_after,
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

    /// A `<w:ptab>` fragment — `line_height`/`leader`/`color` are irrelevant
    /// to fitting, which only reads `align`/`relative_to` (plus the nominal
    /// `MIN_TAB_WIDTH` `.width()` contributes to `line_width`).
    fn ptab_frag(align: crate::model::PTabAlignment) -> Fragment {
        Fragment::PTab {
            align,
            relative_to: crate::model::PTabRelativeTo::Indent,
            leader: crate::model::TabLeader::None,
            line_height: Pt::ZERO,
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
        }
    }

    /// `frag` with an 11pt-like reach: 12pt up, 5pt down.
    fn tall(text: &str, width: f32, break_after: BreakAfter) -> Fragment {
        let mut f = frag(text, width, break_after);
        if let Fragment::Text { metrics, .. } = &mut f {
            metrics.ascent = Pt::new(12.0);
            metrics.descent = Pt::new(5.0);
        }
        f
    }

    /// A space adds nothing to its line's height, at the line's end or inside
    /// it: Word ends a 9pt line with an unformatted 11pt space run and keeps
    /// the line 9pt tall (ShellfishTransactionReport, `{ orderPaidDate}`), and
    /// sets a 20pt space between two 10pt words on a 10pt line
    /// (`test-files/space-run-height.docx`).
    #[test]
    fn a_space_adds_nothing_to_the_line() {
        let frags = vec![
            frag("word", 30.0, BreakAfter::Prohibited),
            tall(" ", 3.0, BreakAfter::Opportunity),
            frag("next", 80.0, BreakAfter::Opportunity),
        ];
        let lines = fit_lines(&frags, Pt::new(50.0));
        assert_eq!(lines.len(), 2);
        assert_eq!(
            lines[0].height.raw(),
            14.0,
            "the space's 17pt is not charged"
        );
        assert_eq!(lines[0].ascent.raw(), 10.0);

        let inside = vec![
            frag("word", 30.0, BreakAfter::Prohibited),
            tall(" ", 3.0, BreakAfter::Opportunity),
            frag("more", 10.0, BreakAfter::Opportunity),
        ];
        let lines = fit_lines(&inside, Pt::new(50.0));
        assert_eq!(lines.len(), 1);
        assert_eq!(lines[0].height.raw(), 14.0, "nor inside it");
    }

    #[test]
    fn empty_fragments_no_lines() {
        let lines = fit_lines(&[], Pt::new(100.0));
        assert!(lines.is_empty());
    }

    #[test]
    fn single_fragment_fits() {
        let frags = vec![text_frag("hello", 30.0)];
        let lines = fit_lines(&frags, Pt::new(100.0));

        assert_eq!(lines.len(), 1);
        assert_eq!(lines[0].start, 0);
        assert_eq!(lines[0].end, 1);
        assert_eq!(lines[0].width.raw(), 30.0);
        assert_eq!(lines[0].height.raw(), 14.0);
    }

    #[test]
    fn two_fragments_fit_on_one_line() {
        let frags = vec![text_frag("hello ", 35.0), text_frag("world", 30.0)];
        let lines = fit_lines(&frags, Pt::new(100.0));

        assert_eq!(lines.len(), 1);
        assert_eq!(lines[0].end, 2);
        assert_eq!(lines[0].width.raw(), 65.0);
    }

    #[test]
    fn overflow_breaks_at_boundary() {
        let frags = vec![
            text_frag("hello ", 60.0),
            text_frag("world ", 60.0),
            text_frag("end", 30.0),
        ];
        let lines = fit_lines(&frags, Pt::new(100.0));

        assert_eq!(lines.len(), 2);
        assert_eq!(lines[0].start, 0);
        assert_eq!(lines[0].end, 1); // "hello " on first line
        assert_eq!(lines[1].start, 1);
        assert_eq!(lines[1].end, 3); // "world " + "end" on second line
    }

    /// Breaking back to an earlier opportunity must re-measure everything
    /// between it and the fragment that overflowed.
    ///
    /// The fitter accumulates forward, and on overflow it rewinds to the last
    /// opportunity. That opportunity can be many fragments behind — here the
    /// space after `"Z:`, with an unbreakable run of clusters after it. The
    /// rewind moved `line_start` back but left the cursor where the overflow
    /// was found, so every fragment in between stayed inside the new line's
    /// range — painted — while contributing nothing to its width. The line
    /// then never overflowed again and swallowed the rest of the paragraph.
    ///
    /// This is issue-shaped rather than theoretical: it is why a Windows path
    /// in a 167.80 pt footer cell laid out as one 295 pt line, 91 pt past the
    /// right edge of the page.
    #[test]
    fn a_backward_break_re_measures_the_fragments_it_rewinds_over() {
        // 10 glued clusters of 10pt each behind one opportunity.
        let mut frags = vec![text_frag("head ", 10.0)];
        frags.extend((0..10).map(|_| glued_frag("x", 10.0)));

        let lines = fit_lines(&frags, Pt::new(50.0));

        // `head ` alone, then the unbreakable run cut into 50pt lines.
        assert_eq!(lines[0].start, 0);
        assert_eq!(
            lines[0].end, 1,
            "the opportunity after `head ` is the first break"
        );
        for line in &lines[1..] {
            let measured: f32 = frags[line.start..line.end]
                .iter()
                .map(|f| f.width().raw())
                .sum();
            assert!(
                measured <= 50.0,
                "line [{}..{}) holds {measured}pt of fragments in a 50pt line",
                line.start,
                line.end,
            );
            assert!(
                (line.width.raw() - measured).abs() < 0.01,
                "line [{}..{}) records width {} but its fragments measure {measured}",
                line.start,
                line.end,
                line.width.raw(),
            );
        }
        let covered: usize = lines.iter().map(|l| l.end - l.start).sum();
        assert_eq!(
            covered,
            frags.len(),
            "every fragment lands on exactly one line"
        );
    }

    /// A token UAX #14 refuses to break inside stays whole even though it is
    /// split across two fragments: `ID‑001` carries a non-breaking hyphen, so
    /// the piece before it says [`BreakAfter::Prohibited`] and the fitter has
    /// to go back to the space before `prefix`.
    #[test]
    fn a_prohibited_boundary_is_not_a_break_point() {
        let frags = vec![
            text_frag("prefix ", 45.0),
            glued_frag("ID‑", 20.0),
            glued_frag("001", 35.0),
        ];
        let lines = fit_lines(&frags, Pt::new(70.0));

        assert_eq!(lines.len(), 2);
        assert_eq!(lines[0].end, 1);
        assert_eq!(lines[1].start, 1);
        assert_eq!(lines[1].end, 3);
    }

    /// `Fragment::MathFraction` carries its own `break_after` (a fraction
    /// used as a superscript's base, `(1/2)²`, is glued to the exponent that
    /// follows it) — the fitter must honor that field exactly like a
    /// `Fragment::Text`'s, not fall into the wildcard arm that used to treat
    /// every non-text fragment as an unconditional break point regardless of
    /// what it carried. Same shape as the test above, with the middle
    /// fragment swapped for a glued `MathFraction`.
    #[test]
    fn a_prohibited_math_fraction_is_not_a_break_point() {
        let frags = vec![
            text_frag("prefix ", 45.0),
            math_fraction_frag(20.0, BreakAfter::Prohibited),
            text_frag("sup", 35.0),
        ];
        let lines = fit_lines(&frags, Pt::new(70.0));

        assert_eq!(lines.len(), 2);
        assert_eq!(lines[0].end, 1);
        assert_eq!(lines[1].start, 1);
        assert_eq!(lines[1].end, 3);
    }

    /// The fitter reads [`BreakAfter`] and nothing else. This fragment ends in
    /// a hyphen and a trailing space — every character the old rule broke on —
    /// and still does not break, because UAX #14 said not to. If this fails,
    /// the character sniff has grown back and `crate::i18n::segment` is no
    /// longer the only answer to the question.
    #[test]
    fn the_text_itself_no_longer_decides_where_a_line_breaks() {
        let frags = vec![
            glued_frag("0100- ", 40.0),
            text_frag("600", 40.0),
            text_frag("rest", 40.0),
        ];
        let lines = fit_lines(&frags, Pt::new(100.0));

        assert_eq!(lines.len(), 2, "{lines:#?}");
        assert_eq!(
            lines[0].end, 2,
            "both halves of the token stay together despite the hyphen — the \
             old character rule would have broken after it and put line 0 at 1",
        );
    }

    #[test]
    fn oversized_fragment_gets_own_line() {
        let frags = vec![text_frag("verylongword", 200.0)];
        let lines = fit_lines(&frags, Pt::new(100.0));

        assert_eq!(lines.len(), 1, "oversized fragment still produces a line");
        assert_eq!(lines[0].end, 1);
    }

    #[test]
    fn line_break_forces_new_line() {
        let frags = vec![
            text_frag("before", 30.0),
            Fragment::LineBreak {
                line_height: Pt::new(14.0),
                origin: crate::render::layout::fragment::LineBreakOrigin::Authored(
                    crate::model::BreakClear::None,
                ),
            },
            text_frag("after", 25.0),
        ];
        let lines = fit_lines(&frags, Pt::new(100.0));

        assert_eq!(lines.len(), 2);
        assert_eq!(lines[0].end, 2); // "before" + line break
        assert!(lines[0].has_break);
        assert_eq!(lines[1].start, 2);
        assert_eq!(lines[1].end, 3); // "after"
    }

    #[test]
    fn exact_fit_no_overflow() {
        let frags = vec![text_frag("a", 50.0), text_frag("b", 50.0)];
        let lines = fit_lines(&frags, Pt::new(100.0));

        assert_eq!(lines.len(), 1);
        assert_eq!(lines[0].width.raw(), 100.0);
    }

    #[test]
    fn tab_uses_min_width_for_fitting() {
        let frags = vec![
            text_frag("text", 80.0),
            Fragment::Tab {
                line_height: Pt::new(14.0),
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
                fitting_width: None,
            },
            text_frag("more", 30.0),
        ];
        // 80 + 12 (MIN_TAB_WIDTH) = 92, still fits 100
        // But + 30 = 122, doesn't fit
        let lines = fit_lines(&frags, Pt::new(100.0));

        assert_eq!(lines.len(), 2);
    }

    #[test]
    fn height_is_max_of_fragments() {
        let frags = vec![
            Fragment::Text {
                shaped: None,
                kerning: None,
                level: crate::i18n::bidi::BidiLevel::LTR,
                text: "small".into(),
                break_after: BreakAfter::Prohibited,
                font: Rc::new(FontProps {
                    vertical: None,
                    features: Default::default(),
                    effects: Default::default(),
                    rtl: crate::render::fonts::Toggle::Absent,
                    family: Rc::from("Test"),
                    size: Pt::new(10.0),
                    bold: Toggle::Absent,
                    italic: Toggle::Absent,
                    underline: false,
                    char_spacing: Pt::ZERO,
                    text_scale: 1.0,
                    underline_position: Pt::ZERO,
                    underline_thickness: Pt::ZERO,
                }),
                color: RgbColor::BLACK,
                width: Pt::new(20.0),
                trimmed_width: Pt::new(20.0),
                metrics: TextMetrics {
                    ascent: Pt::new(9.0),
                    descent: Pt::new(3.0),
                    leading: Pt::ZERO,
                },
                hyperlink_url: None,
                shading: None,
                border: None,
                baseline_offset: Pt::ZERO,
                text_offset: Pt::ZERO,
                is_footnote_ref: false,
            },
            Fragment::Text {
                shaped: None,
                kerning: None,
                level: crate::i18n::bidi::BidiLevel::LTR,
                text: "big".into(),
                break_after: BreakAfter::Prohibited,
                font: Rc::new(FontProps {
                    vertical: None,
                    features: Default::default(),
                    effects: Default::default(),
                    rtl: crate::render::fonts::Toggle::Absent,
                    family: Rc::from("Test"),
                    size: Pt::new(24.0),
                    bold: Toggle::Absent,
                    italic: Toggle::Absent,
                    underline: false,
                    char_spacing: Pt::ZERO,
                    text_scale: 1.0,
                    underline_position: Pt::ZERO,
                    underline_thickness: Pt::ZERO,
                }),
                color: RgbColor::BLACK,
                width: Pt::new(30.0),
                trimmed_width: Pt::new(30.0),
                metrics: TextMetrics {
                    ascent: Pt::new(22.0),
                    descent: Pt::new(6.0),
                    leading: Pt::ZERO,
                },
                hyperlink_url: None,
                shading: None,
                border: None,
                baseline_offset: Pt::ZERO,
                text_offset: Pt::ZERO,
                is_footnote_ref: false,
            },
        ];
        let lines = fit_lines(&frags, Pt::new(100.0));

        assert_eq!(lines.len(), 1);
        assert_eq!(lines[0].height.raw(), 28.0, "max of 12 and 28");
        assert_eq!(lines[0].ascent.raw(), 22.0, "max of 9 and 22");
    }

    #[test]
    fn multiple_overflows_produce_multiple_lines() {
        let frags = vec![
            text_frag("a ", 40.0),
            text_frag("b ", 40.0),
            text_frag("c ", 40.0),
            text_frag("d ", 40.0),
            text_frag("e", 40.0),
        ];
        // max_width=70: "a " fits (40), +"b " = 80 > 70 → break
        let lines = fit_lines(&frags, Pt::new(70.0));

        assert!(lines.len() >= 3, "should produce at least 3 lines");
        // Each line should have at most 1 fragment since 40+40=80 > 70
    }

    #[test]
    fn first_line_narrower_than_remaining() {
        // first_line_width=60, remaining_width=100.
        // "a " (40pt) fits the narrow first line alone; "b " (40pt) + "c" (40pt)
        // = 80pt fit together on the full 100pt remaining line.
        let frags = vec![
            text_frag("a ", 40.0),
            text_frag("b ", 40.0),
            text_frag("c", 40.0),
        ];
        let lines = fit_lines_with_first(
            &frags,
            Pt::new(60.0),
            Pt::new(100.0),
            crate::render::layout::paragraph::PTabGeometry {
                max_width: Pt::new(100.0),
                indent_left: Pt::ZERO,
                indent_first_line: Pt::ZERO,
                content_width: Pt::new(100.0),
                float_left: Pt::ZERO,
                float_right: Pt::ZERO,
            },
            false,
        );
        assert_eq!(lines.len(), 2);
        assert_eq!(lines[0].end, 1, "only 'a ' fits the narrow first line");
        assert_eq!(lines[1].start, 1);
        assert_eq!(
            lines[1].end, 3,
            "'b ' + 'c' both fit on the full second line"
        );
    }

    // ── PTab mirroring under `w:bidi` ────────────────────────────────────────

    #[test]
    fn bidi_rtl_ptab_swaps_floats_the_same_way_emission_does() {
        // A 60pt float narrows the line's physical left; under `w:bidi` that
        // is the line's *trailing* edge, so a `Left`/`Indent` ptab's anchor
        // must be measured from the physical *right* (an unnarrowed 0), not
        // the physical left (60) — exactly the swap emission's mirrored pen
        // already applies. Getting this wrong is invisible unless the swap
        // actually changes which line the tab lands on, which is why the
        // anchor is placed behind a 10pt fragment: 0 is behind it, 60 is not.
        let frags = vec![
            text_frag("A", 10.0),
            ptab_frag(crate::model::PTabAlignment::Left),
            text_frag("B", 10.0),
        ];
        let geometry = crate::render::layout::paragraph::PTabGeometry {
            max_width: Pt::new(100.0),
            indent_left: Pt::ZERO,
            indent_first_line: Pt::ZERO,
            content_width: Pt::new(100.0),
            float_left: Pt::new(60.0),
            float_right: Pt::ZERO,
        };
        let lines = fit_lines_with_first(&frags, Pt::new(100.0), Pt::new(100.0), geometry, true);
        assert_eq!(
            lines.len(),
            2,
            "the ptab's anchor (0, after the float swap) is behind the pen \
             (10, after 'A'), so fitting must break before it — matching \
             emission's own mirrored verdict for the same tab"
        );
        assert_eq!(lines[0].end, 1, "'A' alone, then the line breaks");
        assert_eq!(lines[1].start, 1, "the ptab restarts the next line");
        assert_eq!(lines[1].end, 3, "'B' joins the ptab on that line");
    }

    #[test]
    fn ltr_ptab_is_unaffected_by_bidi_rtl() {
        // The same fragments and geometry as the mirrored case above, but
        // `bidi_rtl: false` — the float swap must not fire, so the anchor
        // (60, unswapped) sits ahead of the pen (10) and the whole paragraph
        // fits on one line, as it always did before this fix.
        let frags = vec![
            text_frag("A", 10.0),
            ptab_frag(crate::model::PTabAlignment::Left),
            text_frag("B", 10.0),
        ];
        let geometry = crate::render::layout::paragraph::PTabGeometry {
            max_width: Pt::new(100.0),
            indent_left: Pt::ZERO,
            indent_first_line: Pt::ZERO,
            content_width: Pt::new(100.0),
            float_left: Pt::new(60.0),
            float_right: Pt::ZERO,
        };
        let lines = fit_lines_with_first(&frags, Pt::new(100.0), Pt::new(100.0), geometry, false);
        assert_eq!(lines.len(), 1, "no break: the anchor is ahead of the pen");
        assert_eq!(lines[0].end, 3);
    }
}
