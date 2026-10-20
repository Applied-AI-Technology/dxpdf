//! §20.1.8.4 `a:alphaModFix`, over `test-files/picture-transparency.docx`
//! (`scripts/make_picture_transparency_fixture.py`).

use dxpdf::render::layout::draw_command::DrawCommand;
use dxpdf::render::layout::draw_command::LayoutedPage;

fn layout(name: &str) -> Vec<LayoutedPage> {
    let path = format!("{}/test-files/{name}", env!("CARGO_MANIFEST_DIR"));
    let bytes = std::fs::read(&path).expect("fixture is committed");
    let doc = dxpdf::docx::parse(&bytes).expect("fixture parses");
    dxpdf::render::resolve_and_layout(doc).1
}

/// §20.1.8.4 `a:alphaModFix`: a picture's opacity reaches its draw command —
/// opaque without the element, `amt / 100000` with it, inline and anchored
/// alike. Word draws a real template's background picture at `amt="5000"` as
/// a faint watermark; this engine drew it at full strength over the text.
#[test]
fn a_picture_is_drawn_at_its_alpha_mod_fix_opacity() {
    let pages = layout("picture-transparency.docx");
    let mut images: Vec<(f32, f32)> = pages[0]
        .commands
        .iter()
        .filter_map(|c| match c {
            DrawCommand::Image { rect, opacity, .. } => Some((rect.origin.y.raw(), *opacity)),
            _ => None,
        })
        .collect();
    images.sort_by(|a, b| a.0.partial_cmp(&b.0).unwrap());
    let opacities: Vec<f32> = images.iter().map(|i| i.1).collect();
    assert_eq!(
        opacities,
        vec![1.0, 0.05, 0.5],
        "control, inline at 5%, anchored at 50%"
    );
}
