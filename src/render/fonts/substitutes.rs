//! §17.8: metric-compatible substitution, and the family-name parsing that
//! feeds it.
//!
//! Two concerns that have to live together because the second exists to serve
//! the first: [`FONT_SUBSTITUTIONS`] is keyed by *family*, so a document naming
//! a *face* only reaches the table after [`strip_weight_suffix`] has cut the
//! face name back to its family.
//!
//! This is the engine's last resort before the system default, and it is
//! deliberately the *last* place a name is parsed rather than read. Everything
//! that can be settled by reading the font's own `name` and `OS/2` tables is
//! settled before control arrives here — see the resolution chain in
//! [`super::FontRegistry::resolve`].
//!
//! "Metric-compatible" means *same advance widths*, so line breaks and page
//! counts match Word even though glyph shapes differ — why Carlito and
//! Liberation are used here rather than a generic sans/serif. Substitution is also a
//! *host-dependent* step: the same DOCX can paginate differently on macOS and
//! Linux depending on what is installed. When cross-platform output differs,
//! check this table first.
//!
//! [`LOOKALIKE_SUBSTITUTIONS`] is the weaker second tier: free faces in the
//! same *style* as a proprietary one that has no metric-compatible clone —
//! Cambria, and the script and handwriting faces Office ships. They do not share advance
//! widths, so a run set in one wraps differently; they exist so a word set in
//! Edwardian Script still reads as a script rather than as the host's default
//! sans. Kept apart from the metric table so that table's promise holds for
//! every entry in it.

/// Open-source metric-compatible substitutes for proprietary fonts. Tried
/// in order when `match_family_style` for the requested family fails.
pub(crate) const FONT_SUBSTITUTIONS: &[(&str, &[&str])] = &[
    ("Calibri", &["Carlito", "Liberation Sans", "Noto Sans"]),
    ("Arial", &["Liberation Sans", "Noto Sans", "Helvetica"]),
    (
        "Times New Roman",
        &["Liberation Serif", "Noto Serif", "Times"],
    ),
    (
        "Courier New",
        &["Liberation Mono", "Noto Sans Mono", "Courier"],
    ),
    ("Verdana", &["DejaVu Sans", "Noto Sans"]),
    ("Georgia", &["DejaVu Serif", "Noto Serif"]),
    ("Trebuchet MS", &["Ubuntu", "Noto Sans"]),
    (
        "Consolas",
        &["Inconsolata", "Liberation Mono", "Noto Sans Mono"],
    ),
    ("Segoe UI", &["Noto Sans", "Liberation Sans"]),
];

/// Strip a trailing weight word from a face-qualified family name, yielding the
/// base family: `"Segoe UI Light"` → `"Segoe UI"`. `None` when the name does not
/// end in a recognised weight word.
///
/// [`FONT_SUBSTITUTIONS`] is keyed by family, so a document naming a *face*
/// otherwise walks straight past the substitution step to the system default —
/// `"Segoe UI"` is in the table but `"Segoe UI Light"` was not reachable from
/// it. When the metadata steps also decline the name there was no path at all
/// from a face name to its family's metric-compatible substitutes.
///
/// The longest matching suffix wins, so `"Foo Extra Light"` yields `"Foo"`
/// rather than `"Foo Extra"`. A weight word must be a separate trailing word:
/// `"Highlight"` ends in `"light"` but is not face-qualified.
pub(crate) fn strip_weight_suffix(name: &str) -> Option<&str> {
    let trimmed = name.trim_end();
    let mut best: Option<&str> = None;
    for weight in (100..=900).step_by(100) {
        for suffix in canonical_weight_names(weight) {
            let Some(cut) = trimmed.len().checked_sub(suffix.len()) else {
                continue;
            };
            if cut == 0 || !trimmed.is_char_boundary(cut) {
                continue;
            }
            if !trimmed[cut..].eq_ignore_ascii_case(suffix) {
                continue;
            }
            let base = trimmed[..cut].trim_end();
            // The suffix has to be its own word, and something must precede it.
            if base.len() == cut || base.is_empty() {
                continue;
            }
            if best.is_none_or(|b| base.len() < b.len()) {
                best = Some(base);
            }
        }
    }
    best
}

/// Open-source faces that look like a proprietary face with no
/// metric-compatible clone. Tried only after [`FONT_SUBSTITUTIONS`] has no
/// answer. Each was chosen by comparing glyph shapes and widths against the
/// original on the text the documents set in it:
///
/// - Cambria → Caladea: shaped like Cambria but not cut to the widths of the
///   Cambria Office ships (6.90): 57 of the 95 printable ASCII characters
///   differ, the digits by up to 35% (Cambria's are tabular), the capitals by
///   up to 12%. Liberation Serif and Noto Serif follow it, nearer than the
///   host default.
/// - Edwardian Script ITC → Imperial Script: the same copperplate hand and
///   stroke weight, glyphs on average 12% off in width.
/// - Bradley Hand ITC → Indie Flower: no free face shares its widths (the
///   nearest are 12% off per glyph); Indie Flower is the closest in feel, a
///   thin, open, informal hand.
///
/// All three are SIL Open Font License faces. Widths cannot be
/// borrowed, but the line a face sets can: each entry carries the original's
/// `hhea` metrics, and a look-alike is measured with those
/// ([`TypefaceEntry::vertical`](super::TypefaceEntry::vertical)), so a line
/// holding one stands as tall as Word makes it and nothing below it moves.
///
/// Where a look-alike is narrower or wider than the original all through, it
/// is also stretched to the original's run of text ([`FaceWidths`],
/// [`TypefaceEntry::width_scale`](super::TypefaceEntry::width_scale)): the
/// glyphs keep their shapes, widened as §17.3.2.45 `w:w` widens them, and a
/// paragraph wraps where the original wraps it far more often than on the
/// look-alike's own widths.
/// Caladea's own line is 2% shorter than Cambria's, Imperial Script's 3%
/// taller than Edwardian's and Indie Flower's 17% taller than Bradley Hand's.
pub(crate) const LOOKALIKE_SUBSTITUTIONS: &[Lookalike] = &[
    Lookalike {
        family: "Cambria",
        substitutes: &["Caladea", "Liberation Serif", "Noto Serif"],
        // hhea 1946/455/0 of 2048 units in all four faces, Word for Mac 16's
        // bundled copy (6.90). Caladea's own is 900/250/0 of 1000.
        vertical: VerticalMetrics {
            ascent: 1946.0 / 2048.0,
            descent: 455.0 / 2048.0,
            line_gap: 0.0,
        },
        // Cambria's run of English text over Caladea's, per face, over a
        // 20,000-character sample of prose: Caladea is 4.4% narrower in
        // Regular and 1.1% in Bold Italic. Stretched by these, the words of
        // a real-world Cambria form template are 1.5% off Cambria's on
        // average rather than 4%.
        widths: FaceWidths {
            regular: 1.046,
            bold: 1.022,
            italic: 1.034,
            bold_italic: 1.011,
        },
    },
    Lookalike {
        family: "Edwardian Script ITC",
        substitutes: &["Imperial Script"],
        // hhea 1742/672/0 of 2048 units, Word for Mac 16's bundled copy.
        vertical: VerticalMetrics {
            ascent: 1742.0 / 2048.0,
            descent: 672.0 / 2048.0,
            line_gap: 0.0,
        },
        widths: FaceWidths::NATURAL,
    },
    Lookalike {
        family: "Bradley Hand ITC",
        substitutes: &["Indie Flower"],
        // hhea 1739/815/0 of 2048 units, as Word for Mac 16 embeds the face.
        vertical: VerticalMetrics {
            ascent: 1739.0 / 2048.0,
            descent: 815.0 / 2048.0,
            line_gap: 0.0,
        },
        widths: FaceWidths::NATURAL,
    },
];

/// A face with no metric-compatible clone, the free faces that look like it,
/// and the vertical metrics and widths they stand in with.
#[derive(Debug)]
pub(crate) struct Lookalike {
    pub family: &'static str,
    pub substitutes: &'static [&'static str],
    pub vertical: VerticalMetrics,
    pub widths: FaceWidths,
}

/// How much wider the original sets a run of text than its look-alike, per
/// face of the family — the factor the look-alike's advances are stretched by.
#[derive(Clone, Copy, Debug, PartialEq)]
pub(crate) struct FaceWidths {
    pub regular: f32,
    pub bold: f32,
    pub italic: f32,
    pub bold_italic: f32,
}

impl FaceWidths {
    /// The look-alike's own widths, for a face no stretch brings nearer.
    pub const NATURAL: Self = Self {
        regular: 1.0,
        bold: 1.0,
        italic: 1.0,
        bold_italic: 1.0,
    };

    /// The factor for the face resolution chose.
    pub fn of(&self, bold: bool, italic: bool) -> f32 {
        match (bold, italic) {
            (false, false) => self.regular,
            (true, false) => self.bold,
            (false, true) => self.italic,
            (true, true) => self.bold_italic,
        }
    }
}

/// A face's line metrics as fractions of its em, all positive: the `hhea`
/// ascender, descender and line gap.
#[derive(Clone, Copy, Debug, PartialEq)]
pub struct VerticalMetrics {
    pub ascent: f32,
    pub descent: f32,
    pub line_gap: f32,
}

/// The entry for `family` in `table`, falling back to its base family when
/// the name is face-qualified. Returns the matched key alongside the list so
/// the caller can log which one fired.
fn lookup_in(
    table: &'static [(&'static str, &'static [&'static str])],
    family: &str,
) -> Option<(&'static str, &'static [&'static str])> {
    let lookup = |name: &str| {
        table
            .iter()
            .find(|(key, _)| key.eq_ignore_ascii_case(name))
            .map(|(key, subs)| (*key, *subs))
    };
    lookup(family).or_else(|| lookup(strip_weight_suffix(family)?))
}

/// Metric-compatible substitutes for `family` ([`FONT_SUBSTITUTIONS`]).
pub(crate) fn substitutes_for(family: &str) -> Option<(&'static str, &'static [&'static str])> {
    lookup_in(FONT_SUBSTITUTIONS, family)
}

/// The [`LOOKALIKE_SUBSTITUTIONS`] entry for `family`, falling back to its
/// base family when the name is face-qualified.
pub(crate) fn lookalikes_for(family: &str) -> Option<&'static Lookalike> {
    let lookup = |name: &str| {
        LOOKALIKE_SUBSTITUTIONS
            .iter()
            .find(|entry| entry.family.eq_ignore_ascii_case(name))
    };
    lookup(family).or_else(|| lookup(strip_weight_suffix(family)?))
}

/// The English spellings a face of `weight` is conventionally called.
///
/// This is a *guess table*, and its existence is the reason the metadata steps
/// run first: a font's `OS/2` `usWeightClass` says 600 outright, whereas
/// matching the word "Semibold" only infers it. Kept for the step that runs
/// when a font has no readable metadata at all.
pub(crate) fn canonical_weight_names(weight: i32) -> &'static [&'static str] {
    match weight {
        100 => &["Thin", "Hairline"],
        200 => &["ExtraLight", "Extra Light", "UltraLight", "Ultra Light"],
        300 => &["Light"],
        400 => &["Regular", "Normal"],
        500 => &["Medium"],
        600 => &["Semibold", "SemiBold", "Semi Bold", "DemiBold", "Demi Bold"],
        700 => &["Bold"],
        800 => &["ExtraBold", "Extra Bold", "UltraBold", "Ultra Bold"],
        900 => &["Black", "Heavy"],
        _ => &[],
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn strips_a_trailing_weight_word_to_the_base_family() {
        assert_eq!(strip_weight_suffix("Segoe UI Light"), Some("Segoe UI"));
        assert_eq!(strip_weight_suffix("Calibri Light"), Some("Calibri"));
        assert_eq!(strip_weight_suffix("Segoe UI Semibold"), Some("Segoe UI"));
        assert_eq!(strip_weight_suffix("Arial Black"), Some("Arial"));
        // Case-insensitive, and tolerant of the spacing variants the alias
        // index already accepts.
        assert_eq!(strip_weight_suffix("Calibri LIGHT"), Some("Calibri"));
        assert_eq!(strip_weight_suffix("Foo Extra Bold"), Some("Foo"));
    }

    /// The longest suffix wins, so a two-word weight name is not left half
    /// attached.
    #[test]
    fn longest_weight_suffix_wins() {
        assert_eq!(strip_weight_suffix("Foo Extra Light"), Some("Foo"));
        assert_eq!(strip_weight_suffix("Foo Ultra Bold"), Some("Foo"));
    }

    /// A weight word must be a separate trailing word — otherwise every family
    /// ending in those letters would be silently truncated.
    #[test]
    fn weight_word_must_be_its_own_word() {
        assert_eq!(strip_weight_suffix("Highlight"), None);
        assert_eq!(strip_weight_suffix("Blackadder"), None);
        assert_eq!(strip_weight_suffix("Light"), None, "nothing precedes it");
        assert_eq!(strip_weight_suffix("Times New Roman"), None);
    }

    /// H2#4 regression: `"Segoe UI"` is in the table, so a face-qualified name
    /// built on it must reach the same substitutes. Before this, the lookup
    /// was an exact whole-name match and face names fell straight through to
    /// the system default.
    #[test]
    fn substitutes_reach_face_qualified_names_through_the_base_family() {
        let (matched, subs) = substitutes_for("Segoe UI").expect("base family is in the table");
        assert_eq!(matched, "Segoe UI");

        let (via, face_subs) =
            substitutes_for("Segoe UI Light").expect("face-qualified name must reach the family");
        assert_eq!(via, "Segoe UI", "resolved through the base family");
        assert_eq!(face_subs, subs, "and gets the same substitutes");

        assert_eq!(
            substitutes_for("Calibri Light").map(|(k, _)| k),
            Some("Calibri")
        );
        // A family that is not in the table stays absent, stripped or not.
        assert!(substitutes_for("Wingdings").is_none());
        assert!(substitutes_for("Wingdings Light").is_none());
    }

    /// A face is either metric-compatible or only a look-alike, never both:
    /// the resolver tries the metric table first, so a family in both would
    /// never reach its look-alikes, and a look-alike in the metric table would
    /// break that table's promise of matching widths.
    #[test]
    fn the_two_tables_share_no_family() {
        for entry in LOOKALIKE_SUBSTITUTIONS {
            assert!(
                substitutes_for(entry.family).is_none(),
                "'{}' is in both tables",
                entry.family
            );
        }
    }

    #[test]
    fn office_script_faces_have_look_alikes() {
        assert_eq!(
            lookalikes_for("Edwardian Script ITC").map(|e| e.substitutes),
            Some(&["Imperial Script"][..])
        );
        assert_eq!(
            lookalikes_for("bradley hand itc").map(|e| e.family),
            Some("Bradley Hand ITC"),
            "matched without regard to case"
        );
        assert!(lookalikes_for("Calibri").is_none());
    }

    /// Caladea is Cambria's look-alike, not its metric clone, and Cambria
    /// Math — a family of its own — is neither.
    #[test]
    fn cambria_has_only_a_look_alike() {
        assert!(substitutes_for("Cambria").is_none());
        assert_eq!(
            lookalikes_for("Cambria").map(|e| e.substitutes[0]),
            Some("Caladea")
        );
        assert_eq!(
            lookalikes_for("Cambria Bold").map(|e| e.family),
            Some("Cambria")
        );
        assert!(lookalikes_for("Cambria Math").is_none());
    }

    /// Caladea is stretched to Cambria's widths face by face; the script
    /// look-alikes keep their own.
    #[test]
    fn a_look_alike_is_stretched_per_face() {
        let cambria = lookalikes_for("Cambria").expect("Cambria has a look-alike");
        assert_eq!(cambria.widths.of(false, false), 1.046);
        assert_eq!(cambria.widths.of(true, false), 1.022);
        assert_eq!(cambria.widths.of(false, true), 1.034);
        assert_eq!(cambria.widths.of(true, true), 1.011);
        for script in ["Edwardian Script ITC", "Bradley Hand ITC"] {
            assert_eq!(
                lookalikes_for(script).map(|e| e.widths),
                Some(FaceWidths::NATURAL)
            );
        }
    }
}
