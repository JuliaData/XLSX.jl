```@meta
CurrentModule = XLSX.Charts
```

!!! note "Experimental"

    Handling of native Excel charts in XLSX.jl is experimental and a work in 
    progress. All aspects may be subject to change. Feedback is welcome!


# Formatting charts

```julia
julia> using XLSX, XLSX.Charts, Colors
```

Chart appearance is DrawingML: the same fills, outlines and text properties that
describe any shape in an Office document. This page covers reading them,
changing them, and the inheritance that decides what a chart actually looks
like.

## Fills and outlines

Every drawable part of a chart carries shape properties (`c:spPr`): the series
graphic, individual data points, the plot area, the chart area, axes, the
legend, gridlines, trendlines and error bars. The getters return a
[`DrawingShapeProps`](@ref), whose `fill` and `line` are a
[`DrawingFill`](@ref) and a [`DrawingLine`](@ref):

```julia
julia> f = XLSX.opentemplate("chart_appearance.xlsx");

julia> c = getCharts(f)[1];

julia> sp = getSeriesShapeProps(c, 1)
DrawingShapeProps(fill solid, no line, effects)

julia> sp.fill
XLSX.Charts.DrawingFill solid
  foreground: #156082

julia> sp.line.fill.kind          # an `a:ln` is present, holding <a:noFill/>
:none
```

Setting them is by property rather than by struct:

```julia
julia> setSeriesFill(c, 1, "FF0000");

julia> setSeriesLine(c, 1; color = "000000", width = 1.5, dash = :sysDot);
```

[`setSeriesLine`](@ref) applies every keyword in one rebuild of the chart part,
which is both faster and safer than a sequence of single-property calls: if one
value is rejected, nothing is written. The single-property setters —
[`setSeriesLineColor`](@ref), [`setSeriesLineWidth`](@ref),
[`setSeriesLineDash`](@ref), [`setSeriesLineCap`](@ref),
[`setSeriesLineCompound`](@ref), [`setSeriesLineJoin`](@ref) and
[`setSeriesLineMiterLimit`](@ref) — exist for when you want just one.

The same keywords set the other outlines a chart has: [`setAxisLine`](@ref) for
an axis line and [`setPlotAreaLine`](@ref) for the border round the plot area.
Like `setSeriesLine`, each also takes `:none` for a line that draws nothing and
`:inherit` to remove it:

```julia
julia> setPlotAreaLine(c; color = "000000");

julia> setAxisLine(c, getChartAxes(c, :value)[1]; color = "000000", width = 1);
```

Widths are points, as in Excel's own boxes. Dash, cap and compound take the
DrawingML name or Excel's name for the same thing, whichever you find easier to
remember:

| | DrawingML | Excel |
| --- | --- | --- |
| dash | `:sysDot`, `:sysDash`, `:lgDash`, `:lgDashDot`, `:lgDashDotDot` | `:roundDot`, `:squareDot`, `:longDash`, `:longDashDot`, `:longDashDotDot` |
| cap | `:sq`, `:rnd` | `:square`, `:round` |
| compound | `:sng`, `:dbl`, `:tri` | `:simple`, `:double`, `:triple` |

`:solid`, `:dash` and `:dashDot`; `:flat`; and `:thickThin` and `:thinThick` are
spelled the same either way.

Setters take symbols; getters report strings, because that is what the file
holds. So `dash = :sysDot` reads back as `"sysDot"`, and an Excel name you set
comes back as its DrawingML equivalent — `dash = :longDash` reads back as
`"lgDash"`.

## Three states, not two

`nothing`, `:none` and `:inherit` mean three different things, and the
difference survives a round trip:

| Written | Reads back as | Excel draws |
| --- | --- | --- |
| nothing at all | `nothing` | whatever the chart style says |
| `<a:noFill/>` | a fill with `kind === :none` | nothing |
| `<a:solidFill>` | a fill with `kind === :solid` | that color |

So `setSeriesFill(c, 1, :none)` makes a series deliberately invisible, while
`setSeriesFill(c, 1, :inherit)` removes the setting and lets the style decide.
The same three states apply to lines, markers, labels and titles.

[`hasFill`](@ref) and [`hasLine`](@ref) collapse those three cases into the
one question a call site usually asks — does this element say something is
drawn?

```julia
julia> sp = getSeriesShapeProps(c, 1);

julia> hasFill(sp), hasLine(sp)
(true, true)
```

They take a `DrawingShapeProps`, a fill or line on its own, or the
[`Effective`](@ref) a resolver returns. `false` never means Excel draws nothing:
a property written at no rung comes from the chart style, which this layer does
not read.

## Colors

Anywhere a color is set, these forms are accepted:

```julia
julia> setSeriesFill(c, 1, "FF0000");                  # hex RGB, "#RRGGBB" or AARRGGBB

julia> setSeriesFill(c, 1, "coral");                   # any Colors.jl named color

julia> setSeriesFill(c, 1, :coral);                    # the same name as a symbol

julia> setSeriesFill(c, 1, RGBA(0.2, 0.4, 0.6, 0.5)); # a Colors.jl value, e.g. computed, here half transparent

julia> setSeriesFill(c, 1, "accent2");                 # a theme color, by name

julia> setSeriesFill(c, 1, SchemeColor(:accent2));     # a theme color
```

These are the same forms cells accept (see [Specifying colors](@ref)), except that in
a chart, alpha is honored and a `SchemeColor` may use any transform.

A [`SchemeColor`](@ref) names a slot in the workbook's theme rather than a fixed
color, so it follows the theme the way Excel's own formatting does. Transforms
are percentages, applied in the order given:

```julia
julia> SchemeColor(:tx1; lumMod = 65, lumOff = 35)
SchemeColor(:tx1, [:lumMod => 65.0, :lumOff => 35.0])
```

A color read back from a file is a [`DrawingColor`](@ref), which records the
reference as authored *and* the RGB it resolves to:

```julia
julia> col = getSeriesFill(c, 1).value.fgcolor;

julia> col.kind, col.val, col.rgb
(:scheme, "accent2", "E97132")
```

So a theme color stays a theme color in the file, while `rgb` tells you what
it looks like.

## The cascade

Excel's chart formatting inherits. A data point falls back to its series, a
series to its group, a group to the plot area and then the chart space. The
resolvers walk that chain and return an [`Effective`](@ref):

```julia
julia> setSeriesFill(c, 1, "FF0000");

julia> e = getSeriesFill(c, 1);

julia> e.value.fgcolor.rgb      # the color now in force
"FF0000"

julia> e.site.level             # which rung supplied it
:series
```

`site` is the rung that answered; `level` names it — `:point`, `:series`,
`:group`, `:plotarea`, `:chartspace`, or `:style` (below). `chain` holds the rungs that exist in
the file, highest precedence first, which is how a setter knows where to write.
A data point appears in the chain only once the file has a `c:dPt` for it, so a
series whose points carry no overrides resolves through a chain of one.

`value === nothing` means no rung wrote the property at all, so Excel takes it
from the chart's style part — not that the property is off. `chain` is populated
either way.

To resolve a point rather than a series, pass `point`. The point's own
formatting wins where it has any, and falls back to the series where it hasn't:

```julia
julia> getSeriesFill(c, 1; point = 2).site.level     # this point has a c:dPt
:point

julia> getSeriesFill(c, 1; point = 3).site.level     # this one doesn't
:series
```

`point` counts from 1, while a [`ChartDataPoint`](@ref)'s `idx` is the `c:idx`
Excel wrote, counting from 0 — so a point read back at `idx == 1` is `point = 2`
here.

### The style rung of a `cx:` chart

Excel writes no fill for a `cx:` series or point until you format it, and draws
it from the chart's style and colours parts instead. [`getSeriesFill`](@ref) on
a [`ChartEx`](@ref) reads those parts as the last rung, so it reports the colour
Excel draws, with `site.level === :style`:

```julia
julia> e = getSeriesFill(box, 7);     # the 7th series of a box & whisker chart

julia> e.site.level
:style

julia> e.value.fgcolor                # accent1, darkened as Excel does for the 7th
XLSX.Charts.DrawingColor(accent1 -> #0D3A4E)
```

Series take the theme's accent colours in turn, darker or lighter ones after
the sixth. A waterfall's points take one colour each for increase, decrease and
total, and a treemap's or sunburst's points the colour of their top-level
category, so for those pass `point`; the series as a whole has no single colour
and resolves to `nothing`. A fill written in the chart part always comes first.

Only fills are read this way, and only for `cx:` charts. Lines and text, the
monochromatic colour schemes, and `c:` charts resolve to `nothing` where no rung
wrote the property; [Limitations](chartLimitations.md) says why.

## Markers

Markers belong to line, scatter and radar series, and to individual points of
them. Setting one on a bar or pie series throws, because `c:barSer` has no
`c:marker` child in the schema — which is why this section uses a line chart.

```julia
julia> k = XLSX.opentemplate("chart_kinds.xlsx");

julia> lc = getCharts(k["linemarkers"])[1];

julia> setMarker(lc, 1; symbol = :circle, size = 7, fill = "FF0000");

julia> getSeriesMarker(lc, 1)
XLSX.Charts.ChartMarker (series idx 0)
  symbol: circle
  size: 7 pt
  shape: DrawingShapeProps(fill solid, line 0.75pt, effects)

julia> setMarkerSymbol(lc, 1, :diamond; point = 3);     # just this point

julia> getDataPointMarker(lc, only(getSeriesDataPoints(lc, 1))).symbol
:diamond
```

`point` counts from 1. A point's marker overrides the series' marker for that
point; [`getMarkerFill`](@ref) resolves the two rungs the same way
[`getSeriesFill`](@ref) resolves shape properties.

The symbols are `:circle`, `:dash`, `:diamond`, `:dot`, `:none`, `:plus`,
`:square`, `:star`, `:triangle`, `:x`, plus `:auto` to let Excel choose and
`:picture` for a marker drawn from an image. Excel uses the same words, so
there are no aliases.

A marker whose `symbol` is `:none` is explicitly turned off, which is not the
same as a series with no `c:marker` element at all.

## Text

Text formatting is set field by field, because text inherits field by field — a
size from one rung and a weight from another both apply:

```julia
julia> setChartTitleTextProp(c, :size, 18);

julia> setChartTitleTextProp(c, :bold, true);

julia> setLabelTextProp(c, 1, :italic, true);          # a series' data labels
```

Each text element has its own setter: [`setChartTitleTextProp`](@ref),
[`setAxisTitleTextProp`](@ref), [`setAxisTextProp`](@ref) for an axis's tick
labels, [`setLegendTextProp`](@ref), [`setLabelTextProp`](@ref) and
[`setChartSpaceTextProp`](@ref) for the chart-wide default.

The chart-wide default only reaches elements that don't set the field
themselves, and the charts Excel creates set a typeface on every text element.
So to change the font throughout a created chart, set `:latin` on each element
rather than once on the chart space.

The fields are those of [`DrawingRunProps`](@ref): `:size`, `:bold`, `:italic`,
`:under`, `:strike`, `:caps`, `:baseline`, `:kern`, `:spacing`, `:latin`, `:ea`,
`:cs`, `:lang`, plus `:fill` and `:line` for the text's own color and outline.
Excel's names are accepted as well: `:font` for `:latin`, `:color` for `:fill`, 
`:underline` for `:under` and `:strikethrough` for `:strike`. Getters report the DrawingML names.

[`getLabelTextProp`](@ref) is the resolving counterpart. It takes a field name
and walks the chain from the series up through the group, plot area and chart
space, returning an [`Effective`](@ref) like the shape resolvers:

```julia
julia> getLabelTextProp(c, 1, :size).value             # the title's 18 doesn't apply
10.5
```

`:fill` and `:line` are composite: they resolve to a whole
[`DrawingFill`](@ref) or [`DrawingLine`](@ref), since a run's fill is written at
some rung or not at all — so `getLabelTextProp(c, 1, :fill).value.fgcolor` is a
field access on the result, not a further walk.

To read what one element says rather than what applies, the `…TextProps` getters
— [`getChartTitleTextProps`](@ref), [`getAxisTextProps`](@ref),
[`getLegendTextProps`](@ref), [`getDataLabelTextProps`](@ref) and the rest —
return a [`DrawingText`](@ref): a text body of paragraphs and runs, each
carrying its own `DrawingRunProps`. The [Chart types](../api/chartTypes.md) page
documents the structure.

## Showing data labels

[`setLabelFlag`](@ref) turns parts of a series' data labels on and off, as the
check boxes of Excel's Format Data Labels pane do:

```julia
julia> setLabelFlag(c, 1, :value, true);           # label each point with its value

julia> setLabelFlag(c, 1, :categoryName, true);    # and its category
```

The parts are `:value`, `:seriesName`, `:categoryName` (the X value on a scatter
chart), `:legendKey`, `:percent` (pie and doughnut charts) and `:bubbleSize`
(bubble charts). A series with no labels of its own is given Excel's: the same
formatting and leader lines Excel gives labels turned on with Chart Elements →
Data Labels, checked against Excel 365 for every kind. [`getLabelFlag`](@ref)
reads a part back, from the series' labels or, failing those, its group's.

The same function works on a [`ChartEx`](@ref), whose labels have three parts:
`:seriesName`, `:categoryName` and `:value`. As in Excel, histogram and Pareto
labels have no category name, and a Pareto line has no labels at all.

Formatting labels with [`setLabelTextProp`](@ref) doesn't make them appear, so
style them and show them in either order.

## Titles and legends

A title's text and its formatting are separate. [`getChartTitleText`](@ref)
gives the literal rich text, [`getChartTitleRef`](@ref) the formula when the
title is bound to a cell, and [`getChartTitleTextProps`](@ref) the formatting
that applies either way.

```julia
julia> setChartTitleText(c, "Revenue by region");

julia> setAxisTitleText(c, getChartAxes(c, :value)[1], "£m");
```

A title these create is written with `c:overlay` off, as Excel writes every
title it creates, so it sits beside the plot area rather than over it. A title
that already has `c:overlay` keeps its value. The text is formatted separately,
with [`setChartTitleTextProp`](@ref) and [`setAxisTitleTextProp`](@ref).

Multi-line titles are written as a [`DrawingText`](@ref) with several
paragraphs:

```julia
julia> setChartTitleText(c, XLSX.Charts.DrawingText("Revenue", "by region"));

julia> getChartTitle(c)
"Revenue\nby region"
```

A `"\n"` inside a single string gives the same two-line title.

A [`RichTextString`](@ref XLSX.RichTextString) keeps each run's formatting, as it
does in a cell. A run's attributes — bold, italic, underline, strikethrough,
superscript and subscript, size, font and color — are written on that run; anything
a run doesn't set is inherited from the title's own formatting. This works with
[`setChartTitleText`](@ref), [`setAxisTitleText`](@ref), [`setLabelText`](@ref)
and the `title` keyword of [`addChart`](@ref):

```julia
julia> rt = XLSX.RichTextString(XLSX.RichTextRun("CO"),
                                XLSX.RichTextRun("2", [:vertAlign => "subscript"]),
                                XLSX.RichTextRun(" emissions", [:bold => true, :color => "red"]));

julia> setChartTitleText(c, rt);
```

On Julia 1.11 and later, a styled string from
[StyledStrings.jl](https://github.com/JuliaLang/StyledStrings.jl) can be passed
instead. It is converted to a `RichTextString` as it is when written to a cell:

```julia
julia> using StyledStrings

julia> setChartTitleText(c, styled"{bold:Revenue} {(foreground=red):£m}");
```

A title can instead be bound to a cell, so it shows whatever the cell holds.
Pass a `CellRef` for a cell on the chart's own sheet, or a `SheetCellRef` to name
the sheet — which a chart on a chartsheet must, having no cells of its own. A
string is always typed text and is never read as a reference, so `"Q1"` is a
title, not cell Q1.

```julia
julia> setChartTitleText(c, CellRef("A1"));

julia> setAxisTitleText(c, getChartAxes(c, :value)[1], SheetCellRef("Data!B1"));

julia> getChartTitleRef(c)
"Sheet1!\$A\$1"
```

A bound title can't carry rich text. This is a limit of Excel's file format,
not of XLSX.jl: a title holds either a cell reference or typed runs, never both,
and a bound title has a single style, its `c:txPr`. Binding a typed title moves
its formatting there — the first paragraph's, overlaid by its first run's. If
the cell holds a `RichTextString`, formatting that all of its runs share is
applied to the title too; formatting that differs from run to run is not shown.
The cell's current text is cached in the chart, as Excel does, so
[`getChartTitle`](@ref) reads it back without opening the file in Excel. Typing
text over a bound title replaces the reference.

A [`ChartEx`](@ref) takes the same text: a string, a `RichTextString`, a styled
string or a cell reference, for its title with `setChartTitleText(c, text)` and
for an axis title with `setAxisTitleText(c, id, text)`, where `id` is one of
[`getChartAxisIds`](@ref). Excel keeps a `cx:` title's plain text in
`cx:txData/cx:v` and renders the runs in its `cx:txPr`, so typed text keeps the
title's existing run formatting, and a `RichTextString` lays each run's own
formatting over it. Binding a `cx:` title to a cell folds that formatting into
the paragraph default, as for a `c:` title, and has the same limit: no rich text.

```julia
julia> setChartTitleText(cx, rt);

julia> setAxisTitleText(cx, 1, CellRef("B1"));
```

Legends are read through [`getChartLegend`](@ref) and its companions —
[`getLegendPos`](@ref), [`getLegendOverlay`](@ref),
[`getLegendShapeProps`](@ref), [`getLegendTextProps`](@ref) — and their text is
set with [`setLegendTextProp`](@ref).

[`setLegendPos`](@ref) moves the legend to one of Excel's five positions: `:b`
(bottom), `:t` (top), `:l` (left), `:r` (right) or `:tr` (top-right). Excel's
names `:bottom`, `:top`, `:left`, `:right` and `:topRight` work too. Its
`overlay` keyword, or [`setLegendOverlay`](@ref), chooses whether the legend is
drawn over the plot area (`true`) or beside it, with the plot area shrunk to make
room (`false`). Both work on a `Chart` and a [`ChartEx`](@ref), and add a legend
if the chart has none:

```julia
julia> setLegendPos(c, :tr; overlay=true);

julia> getLegendPos(c), getLegendOverlay(c)
(:tr, true)
```

A legend shows an entry for every series, in series order. Entries can be
hidden, but not reordered. [`setLegendEntryDeleted`](@ref) hides or shows the
entries of the series at the given positions, in one rebuild:

```julia
julia> setLegendEntryDeleted(c, [1, 3], true);

julia> getLegendEntryDeleted(c, 1)
true
```

To show a legend in some other order, add series that plot zeros in the order
wanted, formatted to match, and hide the entries of the series that draw the
data. [A Likert chart](../examples.md#A-diverging-stacked-bar-chart-for-Likert-scale-survey-data)
does this.

## Axes and gridlines

Axis appearance splits the same way: [`getAxisShapeProps`](@ref) for the axis
line, [`getAxisTextProps`](@ref) for the tick labels,
[`getAxisGridlines`](@ref) for the gridlines:

```julia
julia> ax = getChartAxes(c, :value)[1];

julia> g = getAxisGridlines(c, ax);

julia> isnothing(g)          # are there gridlines at all?
false
```

Gridlines are one of several features where presence is the state: `nothing`
means the element is absent and none are drawn; an element present with no
`c:spPr` means they are drawn with inherited formatting, and comes back as a
`DrawingShapeProps` with every field absent. The same pattern covers drop lines
([`getGroupDropLines`](@ref)), high-low lines ([`getGroupHiLowLines`](@ref)),
series lines ([`getGroupSeriesLines`](@ref)) and up-down bars
([`getGroupUpDownBars`](@ref)).

[`setAxisGridlines`](@ref) shows or removes them. Since presence is the state,
`false` removes the element rather than writing an "off" value, and `true` adds
one that takes its line from the chart style:

```julia
julia> setAxisGridlines(c, ax; major = false);
```

## Axis scale and position

The axis scale itself — bounds, units, tick marks, label position, number
format — is read through the `getAxis…` accessors listed on the
[Charts](../api/charts.md) API page. `nothing` there means Excel chooses
automatically, which is the usual case. Each has a setter, and `:inherit`
returns a property to automatic:

```julia
julia> setAxisScaling(c, ax; min = -1, max = 1);      # both bounds in one call

julia> setAxisMajorUnit(c, ax, 0.25);

julia> setAxisNumberFormatCode(c, ax, "0%;0%");

julia> setAxisMajorTickMark(c, ax, :out);
```

[`setAxisScaling`](@ref) sets the orientation and both bounds in one rebuild,
checking them together, so a range can be moved past its old bounds in one
call. [`setAxisMin`](@ref), [`setAxisMax`](@ref) and
[`setAxisOrientation`](@ref) set one each.

[`setAxisNumberFormatCode`](@ref) also unlinks the format from the source cells,
without which Excel ignores it. A two-section code such as `"0%;0%"` shows
negative values without a minus sign.

Tick marks take `:out`, `:in`, `:cross` or `:none`, or Excel's `:outside` and
`:inside`. [`setAxisTickLabelPos`](@ref) places the labels `:nextTo` the axis,
at the `:low` or `:high` end of the crossing axis, or nowhere. `:low` keeps a
category axis's labels at the edge of the plot area when the values go negative,
rather than beside the zero line.

Where an axis sits is set on the *other* axis: [`setAxisCrosses`](@ref) says
where this axis crosses its partner, as a position on the partner — `:autoZero`,
`:min` or `:max` — and [`setAxisCrossesAt`](@ref) gives the position as a value
on the partner's scale. So Excel's "Horizontal axis crosses: At maximum
category", set in the category axis's pane, is written on the value axis:

```julia
julia> catax = getChartAxes(c, :category)[1];

julia> setAxisOrientation(c, catax, :maxMin);          # first category at the top

julia> setAxisCrosses(c, ax, :max);                    # keep the value axis at the bottom
```

That pair is the usual fix for a bar chart, which plots its first category at
the bottom: reversing the categories moves the value axis to the top, and
crossing at the maximum category moves it back.

## Bar spacing

The gap between bars and the overlap between the bars of a cluster belong to the
chart group rather than to a series, so they take a [`ChartGroup`](@ref). This
uses the stacked bar chart from the file opened for [Markers](#Markers), since a
chart with more than one group needs its bar group picked out first:

```julia
julia> bc = getCharts(k["stackedbar"])[1];

julia> grp = only(getChartGroups(bc));

julia> setGroupGapWidth(bc, grp, 50);                  # % of a bar's width, 0–500
```

[`setGroupOverlap`](@ref) sets the overlap, from −100 to 100. Both are
percentages, as in Excel's boxes. A stacked group needs an overlap of 100 for
each series to sit on the one before, which the stacked kinds already have.
