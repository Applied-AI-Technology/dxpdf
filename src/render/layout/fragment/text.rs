use std::rc::Rc;

use crate::i18n::bidi::BidiLevel;
use crate::render::dimension::Pt;
use crate::render::emoji::cluster::{self, EmojiCluster, InlineCluster};
use crate::render::emoji::resolve::{EmojiFamily, EmojiTypeface};
use crate::render::layout::measurer::TextMeasurer;
use crate::render::resolve::color::RgbColor;
use crate::render::resolve::shading::ResolvedShading;

use super::{BreakAfter, FontProps, Fragment, FragmentBorder, LinkTarget, TextMetrics};

/// §17.18.40 ST_HighlightColor: map highlight enum to RGB.
/// These are the fixed palette colors defined in the OOXML spec.
///
/// Returns `None` for [`HighlightColor::None`] — the spec's explicit
/// "no highlight" override (`<w:highlight w:val="none"/>`) — so the caller
/// suppresses the background fill entirely rather than painting nothing
/// special at the wrong layer.
pub(super) fn resolve_highlight_color(hl: crate::model::HighlightColor) -> Option<RgbColor> {
    use crate::model::HighlightColor;
    Some(match hl {
        HighlightColor::None => return None,
        HighlightColor::Black => RgbColor { r: 0, g: 0, b: 0 },
        HighlightColor::Blue => RgbColor { r: 0, g: 0, b: 255 },
        HighlightColor::Cyan => RgbColor {
            r: 0,
            g: 255,
            b: 255,
        },
        HighlightColor::DarkBlue => RgbColor { r: 0, g: 0, b: 139 },
        HighlightColor::DarkCyan => RgbColor {
            r: 0,
            g: 139,
            b: 139,
        },
        HighlightColor::DarkGray => RgbColor {
            r: 169,
            g: 169,
            b: 169,
        },
        HighlightColor::DarkGreen => RgbColor { r: 0, g: 100, b: 0 },
        HighlightColor::DarkMagenta => RgbColor {
            r: 139,
            g: 0,
            b: 139,
        },
        HighlightColor::DarkRed => RgbColor { r: 139, g: 0, b: 0 },
        HighlightColor::DarkYellow => RgbColor {
            r: 139,
            g: 139,
            b: 0,
        },
        HighlightColor::Green => RgbColor { r: 0, g: 255, b: 0 },
        HighlightColor::LightGray => RgbColor {
            r: 211,
            g: 211,
            b: 211,
        },
        HighlightColor::Magenta => RgbColor {
            r: 255,
            g: 0,
            b: 255,
        },
        HighlightColor::Red => RgbColor { r: 255, g: 0, b: 0 },
        HighlightColor::White => RgbColor {
            r: 255,
            g: 255,
            b: 255,
        },
        HighlightColor::Yellow => RgbColor {
            r: 255,
            g: 255,
            b: 0,
        },
    })
}

/// Resolved styling for a single text fragment.
pub(super) struct TextRunStyle {
    pub color: RgbColor,
    pub shading: Option<ResolvedShading>,
    pub border: Option<FragmentBorder>,
    pub baseline_offset: Pt,
    pub capitals: Capitals,
}

/// §17.3.2.5 `w:caps` / §17.3.2.33 `w:smallCaps`: how a run's letters are
/// drawn. The document keeps the text in the case it was written in; Word
/// draws the capitals, and its PDF extracts them, so the fragments carry them.
#[derive(Clone, Copy, Debug, PartialEq, Eq, Default)]
pub(super) enum Capitals {
    #[default]
    AsWritten,
    /// Every letter in its capital form.
    All,
    /// Lowercase letters as capitals at [`SMALL_CAPS_SCALE`] of the run's
    /// size; capitals, digits and punctuation stay full size.
    Small,
}

/// The size Word draws small capitals at, and the spaces among them: 0.8 of
/// the run's, as `test-files/paragraph-defaults-none.docx` G3 (10pt → 8pt)
/// and G5 (20pt → 16pt) measure.
const SMALL_CAPS_SCALE: f32 = 0.8;

/// Whether small capitals draw `c` reduced: a lowercase letter, or a space —
/// Word sets the spaces of a small-caps run at the reduced size too, after a
/// capital as much as between lowercase letters (G3's "G3 ").
fn small_cap(c: char) -> bool {
    c.is_lowercase() || c == ' '
}

/// Split text into UAX #14 break units and push to the output vec.
///
/// When `measurer` is `Some`, the text is first split into grapheme clusters
/// (UAX #29) and each cluster is classified per UTS #51. Emoji clusters that
/// resolve to a host color emoji typeface become [`Fragment::Emoji`]; clusters
/// without a resolved typeface fall through to the text path with a one-time
/// warning per cluster. When `measurer` is `None` (used by unit tests that
/// don't construct a font registry), the input is passed straight to
/// [`emit_text_words`] — preserving prior behaviour.
pub(super) fn emit_text_fragments<F>(
    text: &str,
    font: &FontProps,
    style: &TextRunStyle,
    hyperlink_url: Option<&LinkTarget>,
    measure_text: &F,
    measurer: Option<&TextMeasurer<'_>>,
    fragments: &mut Vec<Fragment>,
) where
    F: Fn(&str, &FontProps) -> (Pt, TextMetrics),
{
    // §2.1 XML spec: C0 control characters (U+0000–U+001F) other than
    // HT (U+0009), LF (U+000A), CR (U+000D) are invalid in XML but some
    // producers embed LF/CR in w:t content. Strip all non-tab controls
    // so they don't render as tofu/question-mark glyphs.
    let cleaned: String = text
        .chars()
        .filter(|&c| !c.is_control() || c == '\t')
        .collect();
    if cleaned.is_empty() {
        return;
    }

    let Some(measurer) = measurer else {
        emit_text_words(
            &cleaned,
            font,
            style,
            hyperlink_url,
            measure_text,
            fragments,
        );
        return;
    };

    // Classify into clusters and route emoji clusters through the raster
    // pipeline; text spans go through the existing word-split path.
    for cluster in cluster::classify(&cleaned) {
        match cluster {
            InlineCluster::Text(span) => {
                emit_text_words(span, font, style, hyperlink_url, measure_text, fragments);
            }
            InlineCluster::Emoji(emoji) => {
                emit_emoji_or_fallback(
                    &emoji,
                    font,
                    style,
                    hyperlink_url,
                    measure_text,
                    measurer,
                    fragments,
                );
            }
        }
    }
}

/// Segment `text` per UAX #14 and emit one fragment per break unit.
///
/// For text with no surrounding context: the emoji-unavailable fallback, field
/// substitutions, and runs that [`build_inline_units`] refused to join because
/// they hold a tab or a break. Every unit — the trailing one included — takes
/// UAX #14's own answer, for the reason
/// [`JoinedTextSegment::classify`](super::segment) gives about a segment's
/// tail: whatever follows such a run is a tab, a break or a marker, and each
/// of those is already an unconditional break point.
///
/// [`build_inline_units`]: super::segment::build_inline_units
pub(super) fn emit_text_words<F>(
    text: &str,
    font: &FontProps,
    style: &TextRunStyle,
    hyperlink_url: Option<&LinkTarget>,
    measure_text: &F,
    fragments: &mut Vec<Fragment>,
) where
    F: Fn(&str, &FontProps) -> (Pt, TextMetrics),
{
    if text.is_empty() {
        return;
    }
    // Units within a run share their font properties: build one `Rc` per call
    // and hand each fragment a cheap refcount bump instead of a ~48-byte copy.
    let font = Rc::new(font.clone());
    let mut start = 0;
    for end in crate::i18n::segment::break_offsets(text) {
        push_text_fragment(
            &text[start..end],
            BreakAfter::Opportunity,
            &font,
            style,
            hyperlink_url,
            measure_text,
            fragments,
        );
        start = end;
    }
}

/// Emit one fragment for text that is *already* a single UAX #14 break unit,
/// with the break status its trailing edge earned.
///
/// The joined path calls this: `segment::JoinedTextSegment::classify` has
/// already segmented the paragraph's whole text — across `<w:r>` boundaries,
/// which is the only way a Thai word or a CJK phrase split between two runs
/// gets its boundaries right — so re-segmenting the piece here would find
/// nothing and cost a second pass.
///
/// Takes the font already shared: every unit of one `<w:r>` arrives here
/// separately and they all carry the same properties, so the caller resolves
/// and wraps them once and each fragment pays a refcount bump.
pub(super) fn emit_text_unit<F>(
    text: &str,
    break_after: BreakAfter,
    font: &Rc<FontProps>,
    style: &TextRunStyle,
    hyperlink_url: Option<&LinkTarget>,
    measure_text: &F,
    fragments: &mut Vec<Fragment>,
) where
    F: Fn(&str, &FontProps) -> (Pt, TextMetrics),
{
    if text.is_empty() {
        return;
    }
    push_text_fragment(
        text,
        break_after,
        font,
        style,
        hyperlink_url,
        measure_text,
        fragments,
    );
}

/// Measure one break unit and append its fragment.
///
/// `trimmed_width` re-measures only when the unit actually ends in whitespace:
/// trailing space is allowed to hang past the margin, so overflow checking
/// needs the width without it, and every other unit would pay a second Skia
/// call for the same number.
fn push_text_fragment<F>(
    text: &str,
    break_after: BreakAfter,
    font: &Rc<FontProps>,
    style: &TextRunStyle,
    hyperlink_url: Option<&LinkTarget>,
    measure_text: &F,
    fragments: &mut Vec<Fragment>,
) where
    F: Fn(&str, &FontProps) -> (Pt, TextMetrics),
{
    match style.capitals {
        Capitals::AsWritten => push_measured_fragment(
            text,
            break_after,
            font,
            None,
            style,
            hyperlink_url,
            measure_text,
            fragments,
        ),
        Capitals::All => push_measured_fragment(
            &text.to_uppercase(),
            break_after,
            font,
            None,
            style,
            hyperlink_url,
            measure_text,
            fragments,
        ),
        Capitals::Small => {
            // One fragment per stretch of reduced and of full-size letters,
            // nothing breakable between them: the unit is still one word.
            let small = Rc::new(FontProps {
                size: font.size * SMALL_CAPS_SCALE,
                ..(**font).clone()
            });
            let mut pieces = Vec::new();
            let mut start = 0;
            for (i, c) in text.char_indices().skip(1) {
                let previous = text[..i].chars().next_back().is_some_and(small_cap);
                if small_cap(c) != previous {
                    pieces.push(&text[start..i]);
                    start = i;
                }
            }
            pieces.push(&text[start..]);
            let last = pieces.len() - 1;
            for (n, piece) in pieces.into_iter().enumerate() {
                let reduced = piece.chars().next().is_some_and(small_cap);
                push_measured_fragment(
                    &piece.to_uppercase(),
                    if n == last {
                        break_after
                    } else {
                        BreakAfter::Prohibited
                    },
                    if reduced { &small } else { font },
                    // A line holding only small capitals is as tall as the
                    // run's own size would make it, not the reduced one.
                    reduced.then(|| measure_text("X", font).1),
                    style,
                    hyperlink_url,
                    measure_text,
                    fragments,
                );
            }
        }
    }
}

/// [`push_text_fragment`] once the text is in the case it is drawn in.
/// `line_metrics` stands in for the face's own where the line must be sized
/// by a size other than the one drawn.
#[allow(clippy::too_many_arguments)]
fn push_measured_fragment<F>(
    text: &str,
    break_after: BreakAfter,
    font: &Rc<FontProps>,
    line_metrics: Option<TextMetrics>,
    style: &TextRunStyle,
    hyperlink_url: Option<&LinkTarget>,
    measure_text: &F,
    fragments: &mut Vec<Fragment>,
) where
    F: Fn(&str, &FontProps) -> (Pt, TextMetrics),
{
    let (width, metrics) = measure_text(text, font);
    let metrics = line_metrics.unwrap_or(metrics);
    let trimmed = text.trim_end();
    let trimmed_width = if trimmed.len() < text.len() {
        measure_text(trimmed, font).0
    } else {
        width
    };
    fragments.push(Fragment::Text {
        shaped: None,
        kerning: None,
        level: BidiLevel::LTR,
        text: Rc::from(text),
        font: Rc::clone(font),
        color: style.color,
        shading: style.shading,
        border: style.border,
        break_after,
        width,
        trimmed_width,
        metrics,
        hyperlink_url: hyperlink_url.cloned(),
        baseline_offset: style.baseline_offset,
        text_offset: Pt::ZERO,
        is_footnote_ref: false,
    });
}

/// Resolve a host color emoji typeface for an emoji cluster and emit a
/// [`Fragment::Emoji`]. On `Unavailable`, log a one-time warning and route
/// the cluster through the text path so its codepoints still appear in the
/// PDF text stream. The converter never ships emoji font bytes and never
/// degrades silently: it consumes whatever color typeface the host provides,
/// and logs when there is none so an operator can install one.
pub(super) fn emit_emoji_or_fallback<F>(
    cluster: &EmojiCluster<'_>,
    font: &FontProps,
    style: &TextRunStyle,
    hyperlink_url: Option<&LinkTarget>,
    measure_text: &F,
    measurer: &TextMeasurer<'_>,
    fragments: &mut Vec<Fragment>,
) where
    F: Fn(&str, &FontProps) -> (Pt, TextMetrics),
{
    // §17.3.2.26: the run's font name acts as a hint. If it names a known
    // color emoji family, prefer it; otherwise we fall through to the
    // host-default chain (e.g. Calibri-tagged runs containing 📞 still
    // resolve to Apple Color Emoji on macOS).
    let requested = EmojiFamily::from_name_ci(&font.family);
    match measurer.resolve_emoji(requested) {
        EmojiTypeface::Resolved {
            entry: typeface, ..
        } => {
            let (advance, metrics) =
                measurer.measure_with_typeface(cluster.text, &typeface, font.size);
            // Line-height contribution uses the run's text-font metrics so
            // an inline emoji doesn't bloat the line. Color emoji typefaces
            // have ≈1.25× ascent which would otherwise stretch every line
            // that contains an emoji vs surrounding text-only lines.
            let (_, line_metrics) = measure_text("X", font);
            fragments.push(Fragment::Emoji {
                text: cluster.text.to_string(),
                typeface,
                size: font.size,
                presentation: cluster.presentation,
                structure: cluster.structure,
                advance,
                metrics,
                line_metrics,
                baseline_offset: style.baseline_offset,
            });
        }
        EmojiTypeface::Unavailable { attempted } => {
            measurer.warn_emoji_unavailable_once(cluster.text, &attempted);
            emit_text_words(
                cluster.text,
                font,
                style,
                hyperlink_url,
                measure_text,
                fragments,
            );
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::render::fonts::Toggle;

    // ── emit_text_words: UAX #14 units, and the break status each earns ──
    //
    // These were `split_into_words` tests until issue #130. They assert the
    // same divisions plus the thing that division alone never said: whether
    // the fitter may break after each piece. That answer used to be re-derived
    // in `layout::line` from the piece's last character, so a test here could
    // pass while the fitter disagreed.

    /// The text of each emitted unit with the break status it carries.
    fn units(text: &str) -> Vec<(String, BreakAfter)> {
        let mut fragments = Vec::new();
        let measure = |t: &str, _: &FontProps| {
            (
                Pt::new(t.len() as f32 * 6.0),
                TextMetrics {
                    ascent: Pt::new(10.0),
                    descent: Pt::new(2.0),
                    leading: Pt::ZERO,
                },
            )
        };
        emit_text_words(
            text,
            &font("Calibri", 12.0),
            &style(),
            None,
            &measure,
            &mut fragments,
        );
        fragments
            .iter()
            .map(|f| match f {
                Fragment::Text {
                    text, break_after, ..
                } => (text.to_string(), *break_after),
                other => panic!("expected Text, got {other:?}"),
            })
            .collect()
    }

    /// One unit, and its trailing edge is an opportunity like any other: what
    /// follows a run this entry point is used for is a tab, a break or a
    /// marker, each already an unconditional break point of its own.
    #[test]
    fn a_single_word_is_one_unit() {
        assert_eq!(units("hello"), [("hello".into(), BreakAfter::Opportunity)]);
    }

    #[test]
    fn a_space_ends_a_unit_and_opens_an_opportunity() {
        assert_eq!(
            units("hello world"),
            [
                ("hello ".into(), BreakAfter::Opportunity),
                ("world".into(), BreakAfter::Opportunity),
            ],
        );
    }

    #[test]
    fn trailing_whitespace_stays_with_the_word_it_follows() {
        assert_eq!(
            units("hello "),
            [("hello ".into(), BreakAfter::Opportunity)]
        );
    }

    #[test]
    fn a_non_breaking_hyphen_keeps_a_token_in_one_unit() {
        assert_eq!(
            units("ID\u{2011}001"),
            [("ID\u{2011}001".into(), BreakAfter::Opportunity)],
            "one unit — the hyphen is not a boundary",
        );
    }

    #[test]
    fn multiple_words_each_open_an_opportunity() {
        assert_eq!(
            units("the quick brown fox"),
            [
                ("the ".into(), BreakAfter::Opportunity),
                ("quick ".into(), BreakAfter::Opportunity),
                ("brown ".into(), BreakAfter::Opportunity),
                ("fox".into(), BreakAfter::Opportunity),
            ],
        );
    }

    #[test]
    fn empty_text_emits_nothing() {
        assert!(units("").is_empty());
    }

    // ── §17.3.2.5 / §17.3.2.33 capitals ──────────────────────────────────

    /// Each fragment `text` emits with `capitals`: its text, size, break
    /// status and ascent. The measurer's ascent is the size, so the ascent
    /// says which size the line is sized by.
    fn capitalised(text: &str, capitals: Capitals) -> Vec<(String, f32, BreakAfter, f32)> {
        let mut fragments = Vec::new();
        let measure = |t: &str, f: &FontProps| {
            (
                f.size * t.chars().count() as f32 * 0.5,
                TextMetrics {
                    ascent: f.size,
                    descent: Pt::ZERO,
                    leading: Pt::ZERO,
                },
            )
        };
        let style = TextRunStyle {
            capitals,
            ..style()
        };
        emit_text_words(
            text,
            &font("Calibri", 10.0),
            &style,
            None,
            &measure,
            &mut fragments,
        );
        fragments
            .iter()
            .map(|f| match f {
                Fragment::Text {
                    text,
                    font,
                    break_after,
                    metrics,
                    ..
                } => (
                    text.to_string(),
                    font.size.raw(),
                    *break_after,
                    metrics.ascent.raw(),
                ),
                other => panic!("expected Text, got {other:?}"),
            })
            .collect()
    }

    #[test]
    fn caps_draw_every_letter_as_a_capital() {
        use BreakAfter::Opportunity as O;
        assert_eq!(
            capitalised("G2 Record of", Capitals::All),
            [
                ("G2 ".into(), 10.0, O, 10.0),
                ("RECORD ".into(), 10.0, O, 10.0),
                ("OF".into(), 10.0, O, 10.0),
            ],
        );
    }

    /// Word's G3: lowercase letters and every space at 0.8 of the size, the
    /// capitals and digits whole, and a word still breakable only where it
    /// was — after its space.
    #[test]
    fn small_caps_reduce_lowercase_letters_and_spaces() {
        use BreakAfter::{Opportunity as O, Prohibited as P};
        assert_eq!(
            capitalised("G3 Record Of", Capitals::Small),
            [
                ("G3".into(), 10.0, P, 10.0),
                (" ".into(), 8.0, O, 10.0),
                ("R".into(), 10.0, P, 10.0),
                ("ECORD ".into(), 8.0, O, 10.0),
                ("O".into(), 10.0, P, 10.0),
                ("F".into(), 8.0, O, 10.0),
            ],
        );
    }

    #[test]
    fn small_caps_leave_a_capitalised_word_whole() {
        assert_eq!(
            capitalised("NASA42", Capitals::Small),
            [("NASA42".into(), 10.0, BreakAfter::Opportunity, 10.0)],
        );
    }

    /// The gap #130 closes, at the layer that builds the fragments: a Thai
    /// paragraph used to be a single unbreakable unit, and only wrapped
    /// because `split_oversized_fragments` cut it into clusters mid-word.
    #[test]
    fn a_space_less_script_is_divided_into_breakable_units() {
        let thai = units("ภาษาไทยเป็นภาษา");
        assert!(
            thai.len() > 1,
            "Thai must divide into more than one unit, got {thai:?}",
        );
        assert!(
            thai.iter().all(|(_, b)| *b == BreakAfter::Opportunity),
            "every unit opens an opportunity: {thai:?}",
        );
    }

    /// [`emit_text_unit`] is the joined path's entry point: it takes the
    /// break status rather than deriving one, and never re-segments.
    #[test]
    fn emit_text_unit_emits_exactly_one_fragment_with_the_given_status() {
        let mut fragments = Vec::new();
        let measure = |t: &str, _: &FontProps| {
            (
                Pt::new(t.len() as f32 * 6.0),
                TextMetrics {
                    ascent: Pt::new(10.0),
                    descent: Pt::new(2.0),
                    leading: Pt::ZERO,
                },
            )
        };
        emit_text_unit(
            "hello world",
            BreakAfter::Opportunity,
            &Rc::new(font("Calibri", 12.0)),
            &style(),
            None,
            &measure,
            &mut fragments,
        );
        assert_eq!(fragments.len(), 1, "must not re-segment: {fragments:#?}");
        match &fragments[0] {
            Fragment::Text {
                text, break_after, ..
            } => {
                assert_eq!(&**text, "hello world");
                assert_eq!(*break_after, BreakAfter::Opportunity);
            }
            other => panic!("expected Text, got {other:?}"),
        }
    }

    // ── L1, L2, L4: emoji cluster integration ────────────────────────────

    use crate::render::emoji::cluster::EmojiPresentation;
    use crate::render::fonts::FontRegistry;
    use crate::render::layout::measurer::TextMeasurer;
    use skia_safe::FontMgr;
    use std::rc::Rc;

    fn font(family: &str, size: f32) -> FontProps {
        FontProps {
            vertical: None,
            features: Default::default(),
            effects: Default::default(),
            rtl: crate::render::fonts::Toggle::Absent,
            family: Rc::from(family),
            size: Pt::new(size),
            bold: Toggle::Absent,
            italic: Toggle::Absent,
            underline: false,
            char_spacing: Pt::ZERO,
            text_scale: 1.0,
            underline_position: Pt::ZERO,
            underline_thickness: Pt::ZERO,
        }
    }

    fn style() -> TextRunStyle {
        TextRunStyle {
            color: RgbColor::BLACK,
            shading: None,
            border: None,
            baseline_offset: Pt::ZERO,
            capitals: Capitals::AsWritten,
        }
    }

    /// L1 — Run "hi 📞" produces `[Text("hi "), Emoji(...)]` when an emoji
    /// typeface is resolvable on the host. Skipped on hosts without one.
    #[test]
    fn l1_emoji_run_splits_into_text_and_emoji_fragments() {
        let registry = FontRegistry::new(FontMgr::new());
        let measurer = TextMeasurer::new(&registry);
        // Bail if the host has no color emoji — Phase 3 is platform-aware.
        use crate::render::emoji::resolve::EmojiTypeface;
        if matches!(
            measurer.resolve_emoji(None),
            EmojiTypeface::Unavailable { .. }
        ) {
            eprintln!("skipping L1: no color emoji typeface on this host");
            return;
        }
        let mut fragments = Vec::new();
        let measure = |text: &str, fp: &FontProps| measurer.measure(text, fp);
        emit_text_fragments(
            "hi \u{1F4DE}",
            &font("Calibri", 12.0),
            &style(),
            None,
            &measure,
            Some(&measurer),
            &mut fragments,
        );
        assert_eq!(
            fragments.len(),
            2,
            "expected 2 fragments (Text + Emoji), got {fragments:#?}"
        );
        match &fragments[0] {
            Fragment::Text { text, .. } => assert_eq!(&**text, "hi "),
            other => panic!("first fragment must be Text, got {other:?}"),
        }
        match &fragments[1] {
            Fragment::Emoji {
                text,
                presentation,
                advance,
                ..
            } => {
                assert_eq!(text, "\u{1F4DE}");
                assert_eq!(*presentation, EmojiPresentation::Emoji);
                // L2 — advance must be > 0 when the typeface is resolved.
                assert!(
                    advance.raw() > 0.0,
                    "advance must be positive, got {advance}"
                );
            }
            other => panic!("second fragment must be Emoji, got {other:?}"),
        }
    }

    /// L4 — When `measurer` is `None` (no emoji pipeline available), emoji
    /// codepoints flow through the existing text path unchanged. This matches
    /// the no-bundle / no-silent-degradation policy: the codepoint is still
    /// preserved in the PDF's text stream.
    #[test]
    fn l4_no_measurer_routes_emoji_through_text_path() {
        let mut fragments = Vec::new();
        let measure = |text: &str, _fp: &FontProps| {
            (
                Pt::new(text.len() as f32 * 6.0),
                TextMetrics {
                    ascent: Pt::new(10.0),
                    descent: Pt::new(2.0),
                    leading: Pt::ZERO,
                },
            )
        };
        emit_text_fragments(
            "hi \u{1F4DE}",
            &font("Calibri", 12.0),
            &style(),
            None,
            &measure,
            None,
            &mut fragments,
        );
        // No measurer → the whole input is fed to the text path. There must
        // be zero Emoji fragments and the original codepoint must appear in
        // exactly one Text fragment.
        for f in &fragments {
            assert!(
                !matches!(f, Fragment::Emoji { .. }),
                "no emoji fragments must be produced when measurer is None"
            );
        }
        let joined: String = fragments
            .iter()
            .filter_map(|f| match f {
                Fragment::Text { text, .. } => Some(&**text),
                _ => None,
            })
            .collect();
        assert!(
            joined.contains('\u{1F4DE}'),
            "emoji codepoint must survive through the text path"
        );
    }
}
