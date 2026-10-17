//! Document-level settings.

use crate::model::dimension::{Dimension, Twips};

use super::identifiers::RevisionSaveId;

#[derive(Clone, Debug)]
pub struct DocumentSettings {
    /// Default tab stop interval (OOXML default: 720 twips = 0.5 inch).
    pub default_tab_stop: Dimension<Twips>,
    /// Whether even/odd headers/footers are enabled.
    pub even_and_odd_headers: bool,
    /// The rsid of the original editing session that created this document.
    pub rsid_root: Option<RevisionSaveId>,
    /// All revision save IDs recorded in this document's history.
    pub rsids: Vec<RevisionSaveId>,
    /// §17.15.1.21 `w:compat`: the settings that pick which Word's layout the
    /// document asks for.
    pub compat: Compatibility,
}

/// The part of §17.15.1.21 `w:compat` that changes what a page looks like:
/// the `w:compatSetting` entries Word writes under its own URI
/// (`http://schemas.microsoft.com/office/word`), [MS-DOCX] §2.6.1, and the
/// legacy flags read so far.
///
/// The settings are `Option` because absent is its own answer, not a default
/// value: a document with no `compatibilityMode` is laid out like an old one,
/// and one with no `enableOpenTypeFeatures` takes whatever its mode implies.
#[derive(Clone, Copy, Debug, Default, PartialEq, Eq)]
pub struct Compatibility {
    /// `compatibilityMode`: the Word version whose layout the document asks
    /// for — 11 (2003), 12 (2007), 14 (2010), 15 (2013 and later).
    pub mode: Option<u32>,
    /// `enableOpenTypeFeatures`: OpenType ligatures in a document whose mode
    /// would not otherwise have them.
    pub enable_open_type_features: Option<bool>,
    /// §17.15.3 `w:noTabHangInd`: no implicit tab stop at a hanging
    /// paragraph's left indent, outside mode 15
    /// ([`Compatibility::hanging_indent_tab_stop`]).
    pub no_tab_hang_ind: bool,
}

impl Compatibility {
    /// Whether the document asks for Word 2013's layout — `compatibilityMode`
    /// 15 or later. An absent mode is not: Word opens such a document in
    /// compatibility mode, and measured against Word's own render its tables
    /// are placed the way modes 12 and 14 place them, not the way 15 does.
    pub fn is_word_2013(self) -> bool {
        self.mode.is_some_and(|m| m >= 15)
    }

    /// Whether Word applies OpenType ligatures here. Measured: `compatibilityMode`
    /// 15 does, 14 and 12 do not, and 14 with `enableOpenTypeFeatures` does; an
    /// explicit setting decides either way.
    pub fn open_type_features(self) -> bool {
        self.enable_open_type_features
            .unwrap_or_else(|| self.is_word_2013())
    }

    /// Whether a paragraph with a hanging indent has a tab stop at its left
    /// indent, ordered among its own stops. Measured
    /// (`test-files/hanging-tab-*.docx`): Word makes one unless
    /// `w:noTabHangInd` says not to, and ignores the flag in mode 15.
    pub fn hanging_indent_tab_stop(self) -> bool {
        !self.no_tab_hang_ind || self.is_word_2013()
    }
}

impl Default for DocumentSettings {
    fn default() -> Self {
        Self {
            // §17.15.1.25: when `w:defaultTabStop` is omitted the default tab
            // stop interval is 720 twips (0.5"). A derived `Default` would give
            // 0 here — wrong per spec, and it would silently collapse default
            // tabs for any consumer that reads this field.
            default_tab_stop: Dimension::new(720),
            even_and_odd_headers: false,
            rsid_root: None,
            rsids: Vec::new(),
            compat: Compatibility::default(),
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn default_tab_stop_is_720_twips_per_spec() {
        // §17.15.1.25: an omitted `w:defaultTabStop` is 720 twips (0.5"), not
        // the 0 a derived `Default` would produce. Guards against a "simplify
        // to #[derive(Default)]" regression.
        assert_eq!(DocumentSettings::default().default_tab_stop.raw(), 720);
    }

    #[test]
    fn remaining_defaults() {
        let s = DocumentSettings::default();
        assert!(!s.even_and_odd_headers);
        assert!(s.rsid_root.is_none());
        assert!(s.rsids.is_empty());
        assert_eq!(s.compat, Compatibility::default());
    }

    #[test]
    fn only_mode_15_and_later_is_word_2013() {
        let mode = |m| Compatibility {
            mode: m,
            enable_open_type_features: None,
            ..Compatibility::default()
        };
        assert!(!mode(None).is_word_2013(), "absent is compatibility mode");
        assert!(!mode(Some(14)).is_word_2013());
        assert!(mode(Some(15)).is_word_2013());
        assert!(mode(Some(16)).is_word_2013());
    }

    #[test]
    fn open_type_features_follow_the_mode_unless_set() {
        let c = |m, otf| Compatibility {
            mode: m,
            enable_open_type_features: otf,
            ..Compatibility::default()
        };
        assert!(c(Some(15), None).open_type_features());
        assert!(!c(Some(14), None).open_type_features());
        assert!(!c(None, None).open_type_features());
        assert!(c(Some(14), Some(true)).open_type_features());
        assert!(!c(Some(15), Some(false)).open_type_features());
    }

    #[test]
    fn no_tab_hang_ind_holds_only_outside_mode_15() {
        let c = |m, flag| Compatibility {
            mode: m,
            no_tab_hang_ind: flag,
            ..Compatibility::default()
        };
        assert!(c(None, false).hanging_indent_tab_stop());
        assert!(!c(None, true).hanging_indent_tab_stop());
        assert!(!c(Some(14), true).hanging_indent_tab_stop());
        assert!(c(Some(15), true).hanging_indent_tab_stop());
    }
}
