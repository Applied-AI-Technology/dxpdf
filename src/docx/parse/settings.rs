//! Parser for `word/settings.xml`.

use crate::model::Dup;
use serde::Deserialize;

use crate::docx::dimension::{Dimension, Twips};
use crate::docx::error::Result;
use crate::docx::model::{Compatibility, DocumentSettings, RevisionSaveId};
use crate::docx::parse::primitives::units::deserialize_nonnegative_dimension;
use crate::docx::parse::primitives::OnOff;
use crate::docx::parse::serde_xml::from_xml;

/// Parse `word/settings.xml`. Entry point: deserializes into an intermediate
/// schema, then maps to the model type.
pub fn parse_settings(data: &[u8]) -> Result<DocumentSettings> {
    from_xml::<SettingsXml>(data).map(Into::into)
}

#[derive(Deserialize, Default)]
struct SettingsXml {
    #[serde(rename = "defaultTabStop", default)]
    default_tab_stop: Vec<DimensionVal<Twips>>,
    #[serde(rename = "evenAndOddHeaders", default)]
    even_and_odd_headers: Vec<OnOff>,
    #[serde(default)]
    rsids: Vec<RsidsXml>,
    #[serde(default)]
    compat: Vec<CompatXml>,
}

/// §17.15.1.21 `w:compat`. Only its [MS-DOCX] §2.6.1 `w:compatSetting`
/// children are read; the legacy boolean flags beside them are not modelled.
#[derive(Deserialize, Default)]
struct CompatXml {
    #[serde(rename = "compatSetting", default)]
    settings: Vec<CompatSettingXml>,
}

#[derive(Deserialize)]
struct CompatSettingXml {
    #[serde(rename = "@name")]
    name: String,
    #[serde(rename = "@uri", default)]
    uri: String,
    #[serde(rename = "@val", default)]
    val: String,
}

/// The URI Word writes its own `w:compatSetting`s under. A setting of the same
/// name under another URI belongs to another application and says nothing
/// about Word's layout.
const WORD_COMPAT_URI: &str = "http://schemas.microsoft.com/office/word";

impl From<Vec<CompatXml>> for Compatibility {
    fn from(compat: Vec<CompatXml>) -> Self {
        let mut c = Compatibility::default();
        // Later entries win, as for every other repeated element here.
        for setting in compat
            .into_iter()
            .flat_map(|x| x.settings)
            .filter(|s| s.uri == WORD_COMPAT_URI)
        {
            match setting.name.as_str() {
                "compatibilityMode" => {
                    if let Ok(mode) = setting.val.trim().parse() {
                        c.mode = Some(mode);
                    }
                }
                "enableOpenTypeFeatures" => match setting.val.trim() {
                    "1" | "true" | "on" => c.enable_open_type_features = Some(true),
                    "0" | "false" | "off" => c.enable_open_type_features = Some(false),
                    _ => {}
                },
                _ => {}
            }
        }
        c
    }
}

#[derive(Deserialize, Default)]
struct RsidsXml {
    #[serde(rename = "rsidRoot", default)]
    rsid_root: Vec<StringVal>,
    #[serde(rename = "rsid", default)]
    rsids: Vec<StringVal>,
}

#[derive(Deserialize)]
#[serde(bound(deserialize = "U: crate::docx::dimension::Unit"))]
struct DimensionVal<U: crate::docx::dimension::Unit> {
    #[serde(
        rename = "@val",
        deserialize_with = "deserialize_nonnegative_dimension"
    )]
    val: Dimension<U>,
}

#[derive(Deserialize)]
struct StringVal {
    #[serde(rename = "@val")]
    val: String,
}

impl From<SettingsXml> for DocumentSettings {
    fn from(x: SettingsXml) -> Self {
        let mut s = DocumentSettings::default();
        if let Some(t) = Dup::from(x.default_tab_stop).into_value() {
            s.default_tab_stop = t.val;
        }
        if let Some(OnOff(on)) = Dup::from(x.even_and_odd_headers).into_value() {
            s.even_and_odd_headers = on;
        }
        if let Some(r) = Dup::from(x.rsids).into_value() {
            if let Some(root) = Dup::from(r.rsid_root).into_value() {
                s.rsid_root = RevisionSaveId::from_hex(&root.val);
            }
            s.rsids = r
                .rsids
                .into_iter()
                .filter_map(|v| RevisionSaveId::from_hex(&v.val))
                .collect();
        }
        s.compat = Compatibility::from(x.compat);
        s
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn settings(compat: &str) -> DocumentSettings {
        parse_settings(
            format!(
                r#"<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:compat>{compat}</w:compat></w:settings>"#
            )
            .as_bytes(),
        )
        .expect("parses")
    }

    const WORD: &str = "http://schemas.microsoft.com/office/word";

    #[test]
    fn reads_word_s_compatibility_mode_and_open_type_features() {
        let s = settings(&format!(
            r#"<w:compatSetting w:name="compatibilityMode" w:uri="{WORD}" w:val="15"/>
               <w:compatSetting w:name="overrideTableStyleFontSizeAndJustification" w:uri="{WORD}" w:val="1"/>
               <w:compatSetting w:name="enableOpenTypeFeatures" w:uri="{WORD}" w:val="1"/>"#
        ));
        assert_eq!(s.compat.mode, Some(15));
        assert_eq!(s.compat.enable_open_type_features, Some(true));
    }

    #[test]
    fn ignores_a_setting_under_another_uri() {
        let s = settings(
            r#"<w:compatSetting w:name="compatibilityMode" w:uri="urn:example" w:val="15"/>"#,
        );
        assert_eq!(s.compat.mode, None);
    }

    #[test]
    fn absent_compat_is_the_default() {
        let s = parse_settings(
            br#"<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>"#,
        )
        .expect("parses");
        assert_eq!(s.compat, Compatibility::default());
    }
}
