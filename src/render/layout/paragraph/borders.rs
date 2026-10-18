//! Paragraph border and shading rendering.

use super::super::draw_command::DrawCommand;
use super::super::BoxConstraints;
use super::line_emit::resolve_line_height;
use super::types::ParagraphStyle;
use crate::render::dimension::Pt;
use crate::render::geometry::PtOffset;

/// §17.3.1.24: which page segment of a paragraph is being emitted, i.e. which
/// horizontal border edges and vertical spacing it owns. Side borders and
/// shading span every segment; the top edge (and space above) belongs to the
/// first segment, the bottom edge (and `space_after`) to the last.
#[derive(Clone, Copy)]
pub(super) enum SegmentEdges {
    /// A whole, un-split paragraph — both edges.
    Whole,
    /// First of several segments — top edge only.
    First,
    /// Interior segment — neither edge.
    Middle,
    /// Final segment — bottom edge only.
    Last,
}

impl SegmentEdges {
    /// The segment kind for a `[first, last)` range flagged first/last.
    pub(super) fn for_segment(is_first: bool, is_last: bool) -> Self {
        match (is_first, is_last) {
            (true, true) => Self::Whole,
            (true, false) => Self::First,
            (false, false) => Self::Middle,
            (false, true) => Self::Last,
        }
    }

    fn draw_top(self) -> bool {
        matches!(self, Self::Whole | Self::First)
    }

    fn draw_bottom(self) -> bool {
        matches!(self, Self::Whole | Self::Last)
    }
}

/// How far beyond its `w:space` Word stands a paragraph's side border off the
/// text, and how far past the text a side without one ends the box's top and
/// bottom lines. Measured against `test-files/paragraph-side-borders.docx`:
/// every side border's inner edge lies 1.32–1.40pt beyond its `w:space` —
/// at 0, 3, 4 and 10pt of space, 0.5, 1 and 3pt lines, 6, 10 and 20pt text,
/// a half-inch indent either side and a first-line indent, which it does not
/// follow — and 1.35pt is within Word's 0.24pt grain of all of them.
const SIDE_INSET: Pt = Pt::new(1.35);

/// Emit border and shading for a whole (un-split) paragraph, then advance the
/// cursor for bottom border space, `space_after`, and the empty-paragraph
/// minimum height. A thin wrapper over [`emit_segment_borders_and_shading`]
/// with both edges drawn.
pub(super) fn emit_paragraph_borders_and_shading(
    commands: &mut Vec<DrawCommand>,
    style: &ParagraphStyle,
    constraints: &BoxConstraints,
    cursor_y: Pt,
    default_line_height: Pt,
    no_lines: bool,
) -> Pt {
    emit_segment_borders_and_shading(
        commands,
        style,
        constraints,
        cursor_y,
        default_line_height,
        no_lines,
        SegmentEdges::Whole,
    )
}

/// §17.3.1.24/§17.3.1.31: emit border and shading for one page segment of a
/// paragraph, returning the cursor after the segment's trailing spacing.
///
/// `content_bottom` is the y after the segment's last line. `edges` selects
/// which horizontal border edges this segment owns (see [`SegmentEdges`]): side
/// borders and shading span every segment, so a paragraph split across pages
/// keeps a continuous box down each page, while the top/bottom edges and their
/// spacing land only on the first/last segment. `SegmentEdges::Whole`
/// reproduces the un-split geometry exactly.
pub(super) fn emit_segment_borders_and_shading(
    commands: &mut Vec<DrawCommand>,
    style: &ParagraphStyle,
    constraints: &BoxConstraints,
    content_bottom: Pt,
    default_line_height: Pt,
    no_lines: bool,
    edges: SegmentEdges,
) -> Pt {
    let draw_top = edges.draw_top();
    let draw_bottom = edges.draw_bottom();
    let borders = style.borders.as_ref();
    let top_line = borders.and_then(|b| b.top.as_ref());
    let bottom_line = borders.and_then(|b| b.bottom.as_ref());
    let left_line = borders.and_then(|b| b.left.as_ref());
    let right_line = borders.and_then(|b| b.right.as_ref());
    let between = borders.is_some_and(|b| b.joins_above) && top_line.is_some();
    let reach = borders.map(|b| b.reach).unwrap_or_default();
    let para_left = style.indent_left;
    let para_right = constraints.max_width - style.indent_right;

    // §17.3.1.24: a side border stands its `space` and `SIDE_INSET` off the
    // text, its line outside that; a side without one ends the top and bottom
    // lines `SIDE_INSET` past the text. The top and bottom lines run from the
    // outside of one side to the outside of the other, over the corners.
    let left_inner = para_left - left_line.map_or(SIDE_INSET, |l| l.space + SIDE_INSET);
    let right_inner = para_right + right_line.map_or(SIDE_INSET, |l| l.space + SIDE_INSET);
    let left_outer = left_inner - left_line.map_or(Pt::ZERO, |l| l.width);
    let right_outer = right_inner + right_line.map_or(Pt::ZERO, |l| l.width);

    // The rule above the text: a top border stands right under the space
    // before, a box's `between` line its own space below it
    // (`ParagraphStyle::top_border_room`).
    let rule_top = |line: &super::types::BorderLine| {
        style.space_before + if between { line.space } else { Pt::ZERO }
    };
    // Where the sides and shading begin when no rule closes the box above:
    // as far into the space before as the box reaches — no further than the
    // space this segment was laid out with, which is none at a page top.
    let reach_top = style.space_before - reach.above.min(style.space_before);
    let bottom_rule_top = |line: &super::types::BorderLine| content_bottom + line.space;

    // The sides' extent, in up to two runs: a `between` line breaks them.
    let mut side_runs: Vec<(Pt, Pt)> = Vec::with_capacity(2);
    let side_bottom = match bottom_line {
        Some(line) if draw_bottom => bottom_rule_top(line),
        _ if draw_bottom => content_bottom + reach.below,
        _ => content_bottom,
    };
    match top_line {
        Some(line) if draw_top && between => {
            side_runs.push((reach_top, rule_top(line)));
            side_runs.push((rule_top(line) + line.width, side_bottom));
        }
        Some(line) if draw_top => side_runs.push((rule_top(line) + line.width, side_bottom)),
        _ if draw_top => side_runs.push((reach_top, side_bottom)),
        _ => side_runs.push((Pt::ZERO, side_bottom)),
    }

    // §17.3.1.31: render paragraph shading (fills the border area) — a flat
    // colour or a §17.18.78 pattern, both spliced ahead of everything already
    // in `commands` (this segment's lines, run shading, run borders) so the
    // paragraph background sits behind all of it. Background before stripes,
    // same as `table::emit` and `line_emit`'s run shading — see
    // `shading::emit_shading_background`'s doc for why the order matters.
    //
    // Within a bordered box it fills between the side borders where both
    // sides have one and `SIDE_INSET` past the text where either does not
    // (`paragraph-side-borders.docx` S8 and S5), and from the centre of a
    // top or bottom border. A paragraph with no border fills its indents;
    // how far Word fills one has not been measured.
    if let Some(shading) = &style.shading {
        let (fill_left, fill_right) = match (borders, left_line, right_line) {
            (None, _, _) => (para_left, para_right),
            (Some(_), Some(_), Some(_)) => (left_inner, right_inner),
            (Some(_), _, _) => (para_left - SIDE_INSET, para_right + SIDE_INSET),
        };
        let fill_top = match top_line {
            Some(line) if draw_top && !between => rule_top(line) + line.width * 0.5,
            _ if draw_top => reach_top,
            _ => Pt::ZERO,
        };
        let fill_bottom = match bottom_line {
            Some(line) if draw_bottom => bottom_rule_top(line) + line.width * 0.5,
            _ => side_bottom,
        };
        let rect = crate::render::geometry::PtRect::from_xywh(
            fill_left,
            fill_top,
            fill_right - fill_left,
            fill_bottom - fill_top,
        );
        let mut shading_commands = Vec::new();
        crate::render::layout::shading::emit_shading_background(
            &mut shading_commands,
            rect,
            shading,
        );
        crate::render::layout::shading::emit_shading_stripes(&mut shading_commands, rect, shading);
        commands.splice(0..0, shading_commands);
    }

    // §17.3.1.24: render paragraph borders. Top/bottom only on the owning
    // segment; sides on every segment.
    let horizontal = |commands: &mut Vec<DrawCommand>, line: &super::types::BorderLine, top: Pt| {
        let y = top + line.width * 0.5;
        commands.push(DrawCommand::Line {
            line: crate::render::geometry::PtLineSegment::new(
                PtOffset::new(left_outer, y),
                PtOffset::new(right_outer, y),
            ),
            color: line.color,
            width: line.width,
        });
    };
    if let Some(line) = top_line.filter(|_| draw_top) {
        horizontal(commands, line, rule_top(line));
    }
    if let Some(line) = bottom_line.filter(|_| draw_bottom) {
        horizontal(commands, line, bottom_rule_top(line));
    }
    for (line, x) in [
        (
            left_line,
            left_inner - left_line.map_or(Pt::ZERO, |l| l.width * 0.5),
        ),
        (
            right_line,
            right_inner + right_line.map_or(Pt::ZERO, |l| l.width * 0.5),
        ),
    ] {
        let Some(line) = line else { continue };
        for &(top, bottom) in side_runs.iter().filter(|(top, bottom)| bottom > top) {
            commands.push(DrawCommand::Line {
                line: crate::render::geometry::PtLineSegment::new(
                    PtOffset::new(x, top),
                    PtOffset::new(x, bottom),
                ),
                color: line.color,
                width: line.width,
            });
        }
    }

    // §17.3.1.24: only the final segment carries the bottom border space and
    // space_after (§17.3.1.33).
    let mut cursor_y = content_bottom;
    if draw_bottom {
        cursor_y = content_bottom + style.bottom_border_room() + style.space_after;
    }

    // If no lines, still consume default height + spacing (whole-paragraph path).
    if no_lines {
        let line_h = resolve_line_height(
            default_line_height,
            default_line_height,
            default_line_height,
            &style.line_spacing,
            style.auto_fit,
        );
        cursor_y = style.space_before
            + style.top_border_room()
            + line_h
            + style.bottom_border_room()
            + style.space_after;
    }

    cursor_y
}
