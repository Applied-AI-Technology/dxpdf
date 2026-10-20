//! Type definitions for table layout.

use crate::render::dimension::Pt;
use crate::render::geometry::{PtEdgeInsets, PtSize};
use crate::render::resolve::color::RgbColor;

use crate::render::layout::cell::CellLayout;
use crate::render::layout::draw_command::DrawCommand;

/// §17.4.80: row height rule.
#[derive(Clone, Copy, Debug)]
pub enum RowHeightRule {
    /// Row height is at least this value; grows to fit content.
    AtLeast(Pt),
    /// Row height is exactly this value; content may clip.
    Exact(Pt),
}

impl RowHeightRule {
    /// §17.4.80 applied: the height of the row's **content box**, given the
    /// height `content` wants, `above` — the full width of the *interior*
    /// boundary above this row, `Pt::ZERO` for the first row, whose outer top
    /// is drawn inside its box and already in `content` — and `insets`.
    ///
    /// # Measured, and it has been three other ways
    ///
    /// ECMA-376 says only "the height of the current table row" and specifies no
    /// stroke geometry for table borders anywhere, so it cannot say whether a
    /// row's rules are part of its height. Word can. `atLeast` draws the full
    /// height declared and the rules outside it (`issue-157-empty-row-edge.docx`
    /// table 5). `exact` pins the row's pitch instead, and the boundary *above*
    /// the row is inside that pitch: `test-files/row-height-margins.docx` steps
    /// a 6pt `insideH` under three 36pt exact rows and Word puts the second
    /// row's text 41.5pt below the first's (0.5pt of outer top border, then the
    /// whole 6pt rule at the head of row two) and the third's 36pt below the
    /// second's. So the rule a row stands under eats into that row and not the
    /// one above it; the strip `border_gap_below` reserves is how it does,
    /// which is why this subtracts `above` from the box and nothing for below.
    /// Issue-157's table 4 drew ~37pt of cell between its 3pt rules for a 40pt
    /// row, which this reading and the two before it all fit — it measures the
    /// gap between the rules, and only the pitch separates them.
    ///
    /// The readings that died: that `exact` contains the rules on both its
    /// interior boundaries in full (refuted by issue-157's ~37pt, which would
    /// have been 34pt); that the rules stand wholly outside it as they do for
    /// `atLeast` (issue-157 again, mismeasured once as a clean 40pt); and that
    /// it is short by half of each of the two (this fixture's 41.5pt, which
    /// would have been 38.5pt). Each came from a single measurement of a
    /// single shape, which is the warning this comment keeps for whoever is
    /// next: every one of them fitted the render it was read from.
    ///
    /// # `w:val="0"` is not a height
    ///
    /// It is the marker for *unconstrained*, and Word honours it as one even
    /// under `hRule="exact"`: table 2 of that fixture declares zero and Word
    /// draws a full row of cell. [MS-OI29500] §2.4.77(c) is the corroboration —
    /// it records that Word requires `val="0"` whenever `hRule="auto"`, which is
    /// the same fact from the producing side. A literal reading draws a flat row
    /// and loses the cell entirely, which is what this engine used to do. This
    /// is decided before `above` ever applies, since an unconstrained
    /// row is not measuring a box at all.
    ///
    /// # The cell margins are outside the height, except the top one under `exact`
    ///
    /// Measured against `test-files/row-height-margins.docx`, whose rows all
    /// declare 36pt and differ only in their §17.4.42/§17.4.43 top and bottom
    /// cell margins (4.95pt each where present):
    ///
    /// | `hRule`   | no margins | top only | bottom only | both     |
    /// |-----------|------------|----------|-------------|----------|
    /// | `atLeast` | 36         | 36 + top | 36 + bottom | 36 + both |
    /// | `exact`   | 36         | 36       | 36 + bottom | 36 + bottom |
    ///
    /// So `atLeast` is a floor on the content *between* the margins — a row
    /// whose content outgrows it adds both margins to the content instead, the
    /// same as a row with no rule — and `exact` counts its top margin inside
    /// the declared height and its bottom margin outside it. The second looks
    /// arbitrary and is what Word draws in every compatibility mode; it is why
    /// the corpus's card templates (`hRule="exact"` at 3096 twips, 99-twip cell
    /// margins) stand 159.75pt apart in Word and not 154.8pt. A table's outer
    /// bottom border goes with the bottom margin, below the declared height
    /// (the last exact row of the fixture's tables ends 0.5pt lower in Word
    /// than its 36pt, by its 0.5pt rule). `insets` carries
    /// the largest of each across the row's cells, which is the row-level
    /// question a single height answers; the fixture's cells all agree, so
    /// which cell's margin wins when they do not is not something it measures.
    pub(super) fn content_height(self, content: Pt, above: Pt, insets: RowInsets) -> Pt {
        match self {
            // Zero is "unconstrained", not "flat" — see above.
            Self::Exact(h) | Self::AtLeast(h) if h <= Pt::ZERO => content,
            Self::AtLeast(min) => content.max(min + insets.vertical),
            Self::Exact(h) => (h + insets.bottom - above).max(Pt::ZERO),
        }
    }
}

/// What a row's cells put around their content, as
/// [`RowHeightRule::content_height`] needs it: the largest of each across the
/// cells that size the row.
#[derive(Clone, Copy, Debug, Default)]
pub(super) struct RowInsets {
    /// Top plus bottom: the cell margins (§17.4.42/§17.4.43) and the table's
    /// own outer borders where they are drawn inside the box.
    pub(super) vertical: Pt,
    /// The bottom cell margin and, on the last row, the outer bottom border:
    /// what `exact` adds below its height.
    pub(super) bottom: Pt,
}

/// A table row for layout.
pub struct TableRowInput {
    pub cells: Vec<TableCellInput>,
    /// §17.4.80: row height constraint.
    pub height_rule: Option<RowHeightRule>,
    /// §17.4.49: row repeats as header on each continuation page.
    ///
    /// `Option<bool>` and not `bool`, even though `Some(false)` and `None` are
    /// indistinguishable to every reader today — both mean "not a header". The
    /// tri-state is here for the cascade that is not wired yet: §17.7.6 lets a
    /// `<w:tblStylePr w:type="firstRow">` carry a `<w:trPr>`, and once a style
    /// can declare `tblHeader`, a row needs a way to say *explicitly off* and
    /// override it. That is `Some(false)`; `None` is "this level said nothing",
    /// which is what falls through to the style. Collapsing to `bool` now would
    /// erase the distinction and have to be undone then.
    pub is_header: Option<bool>,
    /// §17.4.6: if true, row cannot be split across pages. `Option<bool>` for
    /// the same reason as `is_header` above — `<w:cantSplit>` is a `<w:trPr>`
    /// child, so a table style's conditional `<w:trPr>` can declare it and a row
    /// will need `Some(false)` to turn it back off.
    pub cant_split: Option<bool>,
    /// §17.4.15: number of grid columns to skip at the row's start. The first
    /// cell's leftmost grid column is `grid_before`, not 0.
    ///
    /// There is deliberately no `grid_after` counterpart. §17.4.14 `gridAfter`
    /// is *derivable* — the row's right edge is `grid_before` plus the sum of
    /// its cells' `grid_span`s — and every consumer here works from that
    /// running grid column already, including the border resolution that
    /// decides whether the last cell sits at the table edge (`insideV`) or not.
    /// The model keeps `gridAfter` as parsed (`model::TableRowProperties`);
    /// carrying it into layout added a field nothing read.
    pub grid_before: u32,
    /// §17.4.60 `<w:tblPrEx><w:tblBorders/></w:tblPrEx>` — per-row
    /// override of the table-level `tblBorders`. When set, each side
    /// independently replaces the corresponding side of the table's
    /// `TableBorderConfig` for this row only; sides absent from the
    /// override fall through to the table's value.
    ///
    /// Note that an override side of `None` means "explicitly no
    /// border" (Word's `<w:top w:val="none"/>`) — the parser converts
    /// `val="none"` to `None` here, matching the layout's invariant
    /// that `None` is "draw nothing" everywhere else in this module.
    pub border_overrides: Option<TableBorderConfig>,
    /// §17.4.60 `<w:tblPrEx><w:bidiVisual/></w:tblPrEx>` on this one row. See
    /// [`RowBidiOverride`].
    pub bidi_override: Option<RowBidiOverride>,
}

/// §17.4.60 `<w:tblPrEx><w:bidiVisual/></w:tblPrEx>` — this one row is
/// individually right-to-left, independent of the table's own direction.
///
/// Reversing this row into visual order the way a whole-table `bidiVisual`
/// does (`build::table::mirror_columns`) is not available here: that function
/// swaps the *shared* column grid for every row at once, and a single flipped
/// row would then disagree with its neighbours about what each grid column is
/// — `mirror_columns`' own doc names this as the reason a row-level exception
/// "needs its own design rather than a flag here."
///
/// Word's render supplies that design (`test-files/issue-157-tblprex-bidi.docx`,
/// measured 2026-09-05): the row's cells keep their own declared widths rather
/// than resizing to the slot they visually land in, so its boundaries do not
/// land on the shared grid at all — the fixture's own heading predicts exactly
/// this ("the mirrored row's slot boundaries do not line up with the rows
/// around it"). So this row opts out of the grid entirely: a cell's width
/// comes from here instead of `col_widths` (`measure_table_rows`), its borders
/// are drawn as a closed frame per cell (`emit_cell_frame`) instead of through
/// the shared `super::borders::BorderPlan`, and it is excluded from every
/// cross-row border computation as if it were not part of the table's grid at
/// all — see the `bidi_override` checks in `borders.rs`.
///
/// The same render also settled where the row sits and when it paints. Its own
/// total width is right-aligned against the same "available width" a whole
/// `bidiVisual` table is (`section::helpers::table_x_offset`), as if it were,
/// by itself, a mini right-to-left table — `leading_offset` below is that.
/// And it paints *after* the row that would otherwise follow it rather than
/// between its neighbours, which `build::table` applies as a plain adjacent
/// swap of two `TableRowInput`s at the same seam as `mirror_columns`, early
/// enough that every downstream pass (measurement, border resolution,
/// splitting, painting) simply sees rows in final visual order and needs no
/// separate notion of "document order" at all.
///
/// Both are measured from the one arrangement the fixture tests — a single
/// flipped row with exactly one row after it — and are not verified for any
/// other: two flipped rows, a flipped row with nothing following it, `vMerge`
/// crossing it, or a `gridSpan` cell on it. Those are open questions this does
/// not answer, in the same way `RowHeightRule::content_height`'s doc names
/// what it does not.
pub struct RowBidiOverride {
    /// Each cell's own declared width, in **visual** order — index 0 is this
    /// row's leftmost cell as painted, which is the *last* `<w:tc>` in
    /// document order. Same length as `TableRowInput::cells`, paired with it
    /// 1:1.
    pub cell_widths: Vec<Pt>,
    /// Distance from the table's own local x = 0 to this row's own left edge,
    /// computed so its right edge reaches the *page's* content width rather
    /// than the table's own — see the struct doc above.
    pub leading_offset: Pt,
}

/// §17.4.83: cell vertical alignment.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum CellVAlign {
    Top,
    Center,
    Bottom,
}

/// A single cell for layout.
pub struct TableCellInput {
    pub blocks: Vec<crate::render::layout::section::LayoutBlock>,
    pub margins: PtEdgeInsets,
    /// Number of grid columns this cell spans (gridSpan, default 1).
    pub grid_span: u32,
    /// §17.18.78: resolved cell shading — a flat colour, or a geometric
    /// pattern with its two colours. `None` when the cell shades nothing
    /// (no `w:shd` at any §17.7.6 level, or an explicit `nil`).
    pub shading: Option<crate::render::resolve::shading::ResolvedShading>,
    /// §17.7.6: per-cell resolved borders from conditional formatting.
    pub cell_borders: Option<CellBorderConfig>,
    /// §17.4.84: vertical merge state.
    pub vertical_merge: Option<VerticalMergeState>,
    /// §17.4.83: vertical alignment of content within the cell.
    pub vertical_align: CellVAlign,
}

/// §17.4.84: vertical merge state for a cell.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum VerticalMergeState {
    /// This cell starts a new vertical merge group.
    Restart,
    /// This cell continues from the cell above (content is skipped).
    Continue,
}

/// §17.7.6: a border override for a single cell edge.
///
/// There is deliberately **no variant for `val="none"`**. Per [MS-OI29500]
/// §17.4.66 an explicit `none` is indistinguishable from an omitted edge — both
/// inherit and both yield — so `convert_cell_border_override` maps it to
/// `Option::None` and it never reaches here. Only `nil`, which suppresses,
/// needs to be represented.
#[derive(Clone, Copy, Debug)]
pub enum CellBorderOverride {
    /// §17.18.2 `val="nil"`: draw nothing on this edge *and* win the conflict
    /// against whatever the adjacent cell declares.
    Suppress,
    /// A specific border line on this edge.
    Border(TableBorderLine),
}

/// Per-cell border configuration (resolved from conditional formatting).
/// `None` = no override (use table-level default for this edge).
#[derive(Clone, Debug)]
pub struct CellBorderConfig {
    pub top: Option<CellBorderOverride>,
    pub bottom: Option<CellBorderOverride>,
    pub left: Option<CellBorderOverride>,
    pub right: Option<CellBorderOverride>,
}

/// Resolved table border configuration.
#[derive(Clone, Debug)]
pub struct TableBorderConfig {
    pub top: Option<TableBorderLine>,
    pub bottom: Option<TableBorderLine>,
    pub left: Option<TableBorderLine>,
    pub right: Option<TableBorderLine>,
    pub inside_h: Option<TableBorderLine>,
    pub inside_v: Option<TableBorderLine>,
}

/// A single table border line.
#[derive(Clone, Copy, Debug, PartialEq)]
pub struct TableBorderLine {
    pub width: Pt,
    pub color: RgbColor,
    /// §17.18.2: border style (single, double, etc.)
    pub style: TableBorderStyle,
}

/// Supported table border styles.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum TableBorderStyle {
    Single,
    Double,
}

/// One page-slice of a table.
///
/// A table laid out whole (`layout_table`) is the one-slice case, so it is the
/// same type: there is nothing a monolithic result carries that a slice does
/// not, and two names for one shape only invited the caller to guess which one
/// a given function returns.
#[derive(Debug)]
pub struct TableSlice {
    /// Draw commands positioned relative to this slice's top-left origin (0,0).
    pub commands: Vec<DrawCommand>,
    /// Size of this slice.
    pub size: PtSize,
}

/// Per-row measurement data from the table measurement phase.
/// Contains everything needed to emit draw commands for this row.
pub(super) struct MeasuredRow {
    pub(super) entries: Vec<CellLayoutEntry>,
    pub(super) borders: Vec<super::borders::CellBorders>,
    /// Total vertical space the row owns, **including** `leading_gap`.
    pub(super) height: Pt,
    /// What this row adds to `height` when it heads a continuation slice — the
    /// first row on a page the table has broken onto.
    ///
    /// The top edge that resolution gave to the row above comes back there from
    /// the table's own top border (`emit_one_row`), drawn inside the row's box,
    /// and the box grows by its width under every height rule. Nothing else
    /// changes: an `exact` row still gives up the width of the interior
    /// boundary above it (`RowHeightRule::content_height`), though that
    /// boundary is on the other page. Measured against
    /// `test-files/table-split-head.docx` (Word for Mac 16, 2026-10-16), whose
    /// 36pt `exact` rows carry onto the next page 30.5pt tall under a 6pt
    /// `insideH` and a 0.5pt top, 41.5pt under a 0.5pt `insideH` and a 6pt top,
    /// and 35.5pt with no top border at all.
    ///
    /// That is also the first row's rule, read as giving up its own top border
    /// as the boundary above it and drawing the same border back inside, which
    /// is why row 0 carries zero. So does a split half, which is not
    /// re-measured as a slice head; no render has settled one.
    pub(super) head_extra: Pt,
    /// §17.4.45: cell spacing reserved above this row's content, so the gap to
    /// the row above is exactly one `tblCellSpacing`. Zero without spacing,
    /// which is every table that does not set it.
    pub(super) leading_gap: Pt,
    /// §17.4.38: maximum bottom border width for gap between this row and the next.
    pub(super) border_gap_below: Pt,
    /// Which row of the table's [`BorderPlan`](super::borders::BorderPlan) this
    /// row's borders come from.
    ///
    /// Not the same as its index in `MeasuredTable::rows`, and that is the whole
    /// reason it is carried: a §17.4.49 repeated header appears once in the plan
    /// and many times on the page, and a split row's two halves are both the
    /// same plan row on two different pages.
    pub(super) plan_row: usize,
}

/// Result of the table measurement phase.
pub(super) struct MeasuredTable {
    pub(super) rows: Vec<MeasuredRow>,
    pub(super) table_width: Pt,
    /// §17.4.66: what stands on each line of the border grid. Table-wide and
    /// independent of pagination — only the y of each boundary is per slice.
    pub(super) plan: super::borders::BorderPlan,
    /// x of each vertical grid line, `col_widths.len() + 1` entries.
    pub(super) grid_x: Vec<Pt>,
}

/// Per-cell layout result with positioning info from pass 2.
pub(super) struct CellLayoutEntry {
    pub(super) layout: CellLayout,
    pub(super) cell_x: Pt,
    pub(super) cell_w: Pt,
    /// Starting grid column index (for vMerge neighbor lookup).
    pub(super) grid_col: usize,
    /// §17.4.66/§17.4.42: how far this cell's content starts in from its own
    /// left edge, beyond what `layout_cell` already applied for the margin.
    ///
    /// Carried rather than recomputed at emission, and that is the point. The
    /// rule is `max(border, margin)` per side, but *how much* of a border counts
    /// as inside the cell depends on whether the edge is shared — see
    /// `measure_table_rows`, which is the one place that decides it. Emission
    /// used to derive its own answer from `CellBorders`, which cannot see the
    /// shared line at all: resolution clears the losing cell's edge, so the cell
    /// on the right of every interior vertical was placed as though it had no
    /// border beside it.
    pub(super) content_dx: Pt,
    /// Whether a page break cuts through this cell's content, so that one half
    /// of its split row holds part of it and the other the rest. Word lays
    /// such a cell out top-down on both pages whatever its §17.4.84
    /// `w:vAlign`, while a cell the break leaves whole keeps its alignment
    /// within the half that holds it (`test-files/split-row-valign.docx`).
    pub(super) broken: bool,
}
