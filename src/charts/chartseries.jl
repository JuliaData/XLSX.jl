# Series of `c:` charts and what hangs off them: chart groups, data points,
# markers, trendlines, error bars, drop/high-low/series lines, up-down bars, and
# the series fill, line and marker setters. See chartnodes.jl for the design
# notes.

"""
    getChartSeries(c::Chart; read_cached_values=true, get_external_refs=false) -> Vector{ChartSeries}

The series of `c`, in document order across all chart groups. Each is a value identified by its idx; 
see [`ChartSeries`](@ref)."

`read_cached_values=false` reads metadata only: names, references and point
counts, without the cached values. `get_external_refs=true` replaces an external
workbook index such as `[1]` in each reference with the workbook's path.
"""
function getChartSeries(c::Chart; read_cached_values::Bool=true, get_external_refs::Bool=false)
    out = ChartSeries[]
    for (g, ser) in _series_nodes(chart_root(c))
        s = parse_chart_series(ser, Symbol(localname(g)); read_cached_values)
        push!(out, get_external_refs ? _with_external_refs(c.package, s) : s)
    end
    return out
end

# Resolve a series index against a chart, with a message that says what went
# wrong. `getChartSeries(c)[i]` would throw a BoundsError, which is correct 
# but tells an interactive user nothing about which chart or how many series 
# it has.
function _series(c::Chart, i::Integer; read_cached_values::Bool=true)
    g, ser = _series_pick(c, chart_root(c), i)
    return parse_chart_series(ser, Symbol(localname(g)); read_cached_values)
end

"""
    getChartTypes(c::Chart) -> Vector{Symbol}

The tag of each chart-type group in `c:plotArea`, in document order:
`:barChart`, `:lineChart` and so on. One entry per group, duplicates kept — a
primary and a secondary bar group both appear — and groups holding no series are
included. Contrast [`getChartType`](@ref), which names the chart as a whole.
"""
getChartTypes(c::Chart)::Vector{Symbol} =
    [Symbol(localname(g)) for g in _group_nodes(chart_root(c))]

"""
    getSeriesShapeProps(c::Chart, i::Integer) -> Union{Nothing,DrawingShapeProps}

Fill, line and effects for the graphic of series `i` (`c:ser/c:spPr`), where `i`
is a position in getChartSeries(c), not the `c:idx` value. `nothing` means no `spPr`
was written, which in DrawingML means inherit from the chart style — not "no
formatting". Contrast an explicit `<a:noFill/>`, which `hasFill` reports as
deliberately off.
"""
getSeriesShapeProps(c::Chart, i::Integer) =
    parse_drawing_shape_props(_wb(c), _series(c, i).raw)

"""
    getSeriesLabelTextProps(c::Chart, i::Integer) -> Union{Nothing,DrawingText}

Text body for the data labels of series `i` (`c:ser/c:dLbls/c:txPr`). Per-point
overrides in `c:dLbls/c:dLbl` are not consulted; this is the series-level default.
"""
getSeriesLabelTextProps(c::Chart, i::Integer) =
    parse_drawing_text(_wb(c), first_element_with_tag(_series(c, i).raw, "dLbls"))

# Shared by `getSeriesMarker` and `getDataPointMarker`. `parent` is a `c:ser` or
# `c:dPt` — anything that may carry a `c:marker` child.
#
# Note `c:marker` is ambiguous by tag: as a child of `c:lineChart` or
# `c:scatterChart` it is a boolean show-markers flag, not this element. Only
# call this with a series or data point.
function _parse_marker(wb::Workbook, parent::Union{Nothing,XML.Node},
                       sidx::Int, pidx::Union{Nothing,Int})
    m = first_element_with_tag(parent, "marker")
    isnothing(m) && return nothing
    return ChartMarker(sidx, pidx, _sym_val(m, "symbol"), _int_val(m, "size"),
                       parse_drawing_shape_props(wb, m), m)
end

"""
    getSeriesMarker(c::Chart, i::Integer) -> Union{Nothing,ChartMarker}

Marker for series `i` of a line, scatter or radar chart (`c:ser/c:marker`).
`nothing` means no `c:marker` element; a marker whose `symbol` is `:none` is a
marker explicitly turned off, which is a different thing.
"""
function getSeriesMarker(c::Chart, i::Integer)
    _, ser = _series_pick(c, chart_root(c), i)
    return _parse_marker(_wb(c), ser, _ser_idx(ser), nothing)
end

function _chart_group(el::XML.Node)::ChartGroup
    return ChartGroup(Symbol(localname(el)), _group_axids(el), el)
end

"""
    getChartGroups(c::Chart) -> Vector{ChartGroup}

The chart-type groups in `c:plotArea`, in document order. A combo chart has
several; a plain chart has one. Use [`getGroupAxes`](@ref) to find which axes a
group's series are plotted against.
"""
getChartGroups(c::Chart)::Vector{ChartGroup} =
    [_chart_group(g) for g in _group_nodes(chart_root(c))]

"""
    getSeriesGroup(c::Chart, i::Integer) -> ChartGroup

The chart-type group containing series `i`.
"""
function getSeriesGroup(c::Chart, i::Integer)::ChartGroup
    grp, _ = _series_pick(c, chart_root(c), i)
    return _chart_group(grp)
end

"""
    getGroupAxes(c::Chart, g::ChartGroup) -> Vector{ChartAxis}

The axes named by `g`'s `c:axId` children, in that order. Empty for pie and
doughnut groups, which have no axes. An id that names no axis in the plot area
is skipped rather than raising — malformed files exist.
"""
function getGroupAxes(c::Chart, g::ChartGroup)::Vector{ChartAxis}
    axes = getChartAxes(c)
    out = ChartAxis[]
    for id in g.axids
        i = findfirst(ax -> ax.axid == id, axes)
        isnothing(i) || push!(out, axes[i])
    end
    return out
end

"""
    getGroupLabelTextProps(c::Chart, g::ChartGroup) -> Union{Nothing,DrawingText}

Group-level data label text (`c:barChart/c:dLbls/c:txPr`). This is the
inheritance parent for the series-level `c:dLbls`; Excel commonly writes a group
`c:dLbls` carrying only the `show*` flags and no `txPr`, in which case this is
`nothing`.
"""
getGroupLabelTextProps(c::Chart, g::ChartGroup) =
    parse_drawing_text(_wb(c), first_element_with_tag(_node(c, g), "dLbls"))

"""
    getSeriesAxes(c::Chart, i::Integer) -> Vector{ChartAxis}

The axes series `i` is plotted against. This is how you tell a primary-axis
series from a secondary-axis one: they sit in different groups naming different
axis ids.
"""
getSeriesAxes(c::Chart, i::Integer) = getGroupAxes(c, getSeriesGroup(c, i))

"""
    getSeriesDataPoints(c::Chart, i::Integer) -> Vector{ChartDataPoint}

Per-point overrides on series `i`, in document order. Usually empty, or short —
only points the user formatted individually appear.
"""
function getSeriesDataPoints(c::Chart, i::Integer)::Vector{ChartDataPoint}
    _, ser = _series_pick(c, chart_root(c), i)
    sidx = _ser_idx(ser)
    out = ChartDataPoint[]
    for el in elements_with_tag(ser, "dPt")
        n = _int_val(el, "idx")
        isnothing(n) && continue          # a dPt with no idx applies to nothing
        push!(out, ChartDataPoint(sidx, n, _bool_val(el, "invertIfNegative"),
                                  _bool_val(el, "bubble3D"), el))
    end
    return out
end

"""
    getSeriesDataPoint(c::Chart, i::Integer, point::Integer) -> Union{Nothing,ChartDataPoint}

The override for data point `point` of series `i`, counting from 1, or
`nothing` if that point has no individual formatting.
"""
function getSeriesDataPoint(c::Chart, i::Integer, point::Integer)
    point >= 1 || throw(XLSXError("Data point positions start at 1; asked for $point."))
    dps = getSeriesDataPoints(c, i)
    j = findfirst(d -> d.idx == point - 1, dps)
    return isnothing(j) ? nothing : dps[j]
end

"""
    getDataPointShapeProps(c::Chart, d::ChartDataPoint) -> Union{Nothing,DrawingShapeProps}

Fill and line overriding the series' own, for one data point (`c:dPt/c:spPr`).
`nothing` means the point writes no `c:spPr`, so the series' formatting applies;
[`getSeriesFill`](@ref) resolves that cascade.
"""
getDataPointShapeProps(c::Chart, d::ChartDataPoint) = parse_drawing_shape_props(_wb(c), _node(c, d))

"""
    getDataPointMarker(c::Chart, d::ChartDataPoint) -> Union{Nothing,ChartMarker}

Marker override for a single data point (`c:dPt/c:marker`). `nothing` means the
point does not override the series marker — the series-level
[`getSeriesMarker`](@ref) applies.
"""
getDataPointMarker(c::Chart, d::ChartDataPoint) =
    _parse_marker(_wb(c), _node(c, d), d.series_idx, d.idx)

# --- trendlines -----------------------------------------------------------

"""
    getSeriesTrendlines(c::Chart, i::Integer) -> Vector{ChartTrendline}

Trendlines on series `i`, in document order. A series can carry several — a
linear fit and a moving average, say — so this returns a vector. Usually empty.
"""
function getSeriesTrendlines(c::Chart, i::Integer)::Vector{ChartTrendline}
    _, ser = _series_pick(c, chart_root(c), i)
    sidx = _ser_idx(ser)
    return [ChartTrendline(sidx, k, _sym_val(el, "trendlineType"), child_text(el, "name"),
                           _int_val(el, "order"), _int_val(el, "period"),
                           _num_val(el, "forward"), _num_val(el, "backward"),
                           _num_val(el, "intercept"), _bool_val(el, "dispRSqr"),
                           _bool_val(el, "dispEq"), el)
            for (k, el) in enumerate(elements_with_tag(ser, "trendline"))]
end

"""
    getTrendlineShapeProps(c::Chart, t::ChartTrendline) -> Union{Nothing,DrawingShapeProps}

Line properties of the trendline itself (`c:trendline/c:spPr`). Trendlines are
lines, so the fill is normally absent; the dash pattern lives in `line.dash`.
"""
getTrendlineShapeProps(c::Chart, t::ChartTrendline) = parse_drawing_shape_props(_wb(c), _node(c, t))

"""
    getTrendlineLabelText(c::Chart, t::ChartTrendline) -> Union{Nothing,DrawingText}

Literal text of the trendline's on-chart label (`c:trendlineLbl/c:tx/c:rich`).
`nothing` means the label shows Excel's generated equation or R² rather than
typed-over text — or that there is no label at all. Contrast
[`getTrendlineLabelTextProps`](@ref), which is the label's `c:txPr` formatting.
"""
getTrendlineLabelText(c::Chart, t::ChartTrendline) =
    parse_drawing_text(_wb(c), first_element_with_tag(_trendline_lbl(_node(c, t)), "tx"); tag="rich")

"""
    getTrendlineLabelTextProps(c::Chart, t::ChartTrendline) -> Union{Nothing,DrawingText}

Formatting of the trendline label (`c:trendlineLbl/c:txPr`).
"""
getTrendlineLabelTextProps(c::Chart, t::ChartTrendline)  = parse_drawing_text(_wb(c), _trendline_lbl(_node(c, t)))

"""
    getTrendlineLabelShapeProps(c::Chart, t::ChartTrendline) -> Union{Nothing,DrawingShapeProps}

Fill and border of the trendline label's box (`c:trendlineLbl/c:spPr`).
"""
getTrendlineLabelShapeProps(c::Chart, t::ChartTrendline) = parse_drawing_shape_props(_wb(c), _trendline_lbl(_node(c, t)))

_trendline_lbl(n::XML.Node) = first_element_with_tag(n, "trendlineLbl")

# --- error bars -----------------------------------------------------------

"""
    getSeriesErrorBars(c::Chart, i::Integer) -> Vector{ChartErrorBars}

Error bars on series `i`, in document order. A series carries at most two, one
for each of the `x` and `y` directions.
"""
function getSeriesErrorBars(c::Chart, i::Integer)::Vector{ChartErrorBars}
    _, ser = _series_pick(c, chart_root(c), i)
    sidx = _ser_idx(ser)
    return [ChartErrorBars(sidx, k, _sym_val(el, "errDir"), _sym_val(el, "errBarType"),
                           _sym_val(el, "errValType"), _num_val(el, "val"),
                           _bool_val(el, "noEndCap"), el)
            for (k, el) in enumerate(elements_with_tag(ser, "errBars"))]
end

"""
    getErrorBarsShapeProps(c::Chart, e::ChartErrorBars) -> Union{Nothing,DrawingShapeProps}

Line properties of the error bars (`c:errBars/c:spPr`).
"""
getErrorBarsShapeProps(c::Chart, e::ChartErrorBars) = parse_drawing_shape_props(_wb(c), _node(c, e))

"""
    getErrorBarsCustomRefs(c::Chart, e::ChartErrorBars) -> NamedTuple

The sheet references behind custom error bars, as `(plus, minus)`, each a
[`ChartRef`](@ref) or `nothing`.

Only `errValType="cust"` bars carry these; every other value type takes its
magnitude from [`ChartErrorBars`](@ref)`.value` instead, and this returns
`(nothing, nothing)`. Excel writes both `c:plus` and `c:minus` even for
one-sided bars, so a `nothing` here means the element was genuinely absent.

Cached values are always read, whatever `read_cached_values` is passed to
[`getChartSeries`](@ref) — these ranges are short, so there is nothing to save by
skipping them.
"""
function getErrorBarsCustomRefs(c::Chart, e::ChartErrorBars)
    n = _node(c, e)
    f(tag) = parse_chart_ref(first_element_with_tag(n, tag); read_cached_values=true)
    return (plus = f("plus"), minus = f("minus"))
end

# --- line and stock chart extras ------------------------------------------
#
# These hang off the group rather than a series: drop lines run from a point to
# the axis, hi-low lines and up-down bars span between series, so none of them
# belongs to one series. Each carries only `c:spPr`.
#
# Presence is the state. Unlike most of this file, `nothing` here means the
# feature is off rather than inherited — Excel writes the element when you turn
# the feature on and removes it when you turn it off. An element present with no
# `c:spPr` means on, drawn with inherited formatting.

"""
    getGroupDropLines(c::Chart, g::ChartGroup) -> Union{Nothing,DrawingShapeProps}

Line properties of the group's drop lines (`c:dropLines`), which run from each
data point down to the category axis. `nothing` means the group has no drop
lines. Line, area and stock charts.
"""
getGroupDropLines(c::Chart, g::ChartGroup)   = _optional_spPr(c, _node(c, g), "dropLines")

"""
    getGroupHiLowLines(c::Chart, g::ChartGroup) -> Union{Nothing,DrawingShapeProps}

Line properties of the group's high-low lines (`c:hiLowLines`), which span
between the highest and lowest series at each category. Line and stock charts.
"""
getGroupHiLowLines(c::Chart, g::ChartGroup)  = _optional_spPr(c, _node(c, g), "hiLowLines")

"""
    getGroupSeriesLines(c::Chart, g::ChartGroup) -> Union{Nothing,DrawingShapeProps}

Line properties of the group's series lines (`c:serLines`), which connect
segments across categories on a stacked bar or an of-pie chart.
"""
getGroupSeriesLines(c::Chart, g::ChartGroup) = _optional_spPr(c, _node(c, g), "serLines")

"""
    getGroupUpDownBars(c::Chart, g::ChartGroup) -> Union{Nothing,ChartUpDownBars}

The group's up-down bars (`c:upDownBars`), which span between the first and last
series at each category. `nothing` means the group has none. Line and stock
charts.

The result carries the gap width and is the handle the two formatting getters
take: [`getUpBarShapeProps`](@ref) and [`getDownBarShapeProps`](@ref).
"""
function getGroupUpDownBars(c::Chart, g::ChartGroup)
    el = first_element_with_tag(_node(c, g), "upDownBars")
    isnothing(el) && return nothing
    return ChartUpDownBars(g, _int_val(el, "gapWidth"), el)
end

"""
    getUpBarShapeProps(c::Chart, b::ChartUpDownBars) -> Union{Nothing,DrawingShapeProps}

Fill and border of the up bars (`c:upBars/c:spPr`). `nothing` means `c:upBars` 
was written with no formatting, or not written at all — Excel writes an empty 
`<c:upBars/>` for default white bars. The two cases differ only in the XML; 
both draw Excel's default bars.
"""
getUpBarShapeProps(c::Chart, b::ChartUpDownBars) =
    parse_drawing_shape_props(_wb(c), first_element_with_tag(_node(c, b), "upBars"))

"""
    getDownBarShapeProps(c::Chart, b::ChartUpDownBars) -> Union{Nothing,DrawingShapeProps}

Fill and border of the down bars (`c:downBars/c:spPr`), as
[`getUpBarShapeProps`](@ref) is for the up bars.
"""
getDownBarShapeProps(c::Chart, b::ChartUpDownBars) =
    parse_drawing_shape_props(_wb(c), first_element_with_tag(_node(c, b), "downBars"))

function _set_series_shape(c::Chart, i::Integer, f)
    root = chart_root(c)
    pfx  = ns_prefixes(root)
    steps, ser_key = _series_path(c, root, i)
    new = rebuild_path(root, [steps...; (NS_A, "spPr") => "spPr"],
                       sp -> f(sp, pfx);
                       prefixes = pfx, parent_key = ser_key)
    set_chart_root!(c, new)
    return c
end

"""
    setSeriesFill(c, i, color) -> Chart

Set the fill of series `i`'s graphic — its bars, area or pie slices.

# Color values

Every chart setter that takes a color — fills, lines, markers, the text
`:fill` field and `addSeries`' `color` keyword — accepts the same values:

- a color name from Colors.jl, as a String or Symbol: `"red"`, `:steelblue`;
- a hex string, `"RRGGBB"` or `"AARRGGBB"`, with or without a leading `#`:
  `"FF0000"`, `"#FF0000"`, `"80FF0000"` (50% transparent red);
- a `Colors.Colorant`, such as `colorant"red"` or `RGB(1, 0, 0)`;
- a theme color, named as a String or Symbol: $(_doc_syms(SCHEME_TOKENS)).
  It follows the workbook's theme, as Excel's palette colors do;
- a [`SchemeColor`](@ref), for a theme color with a tint or shade applied:
  `SchemeColor(:accent1; lumMod = 75)`;
- `:none`, for an explicit `<a:noFill/>` — deliberately invisible;
- `:inherit`, to remove the fill so the chart style applies.

`:none` and `:inherit` are different states. An absent fill inherits; `noFill`
does not.

Returns `c`. A `Chart` reads its part on every call, so `c` and any other handle
to the same chart see the change. Values read from it earlier, such as a
`ChartSeries` or a `ChartDataPoint`, describe the part as it was then.
"""
setSeriesFill(c::Chart, i::Integer, color::Union{AbstractString,Colors.Colorant,SchemeColor}) =
    _set_series_shape(c, i, (sp, pfx) -> _sp_with_fill(sp, (NS_A, "spPr"), color, pfx))

function setSeriesFill(c::Chart, i::Integer, what::Symbol)
    (what === :none || what === :inherit) &&
        return _set_series_shape(c, i, (sp, pfx) -> _sp_with_fill(sp, (NS_A, "spPr"), what, pfx))
    return setSeriesFill(c, i, String(what))
end

"""
    _set_series_line(c, i, f) -> Chart

Apply `f` to series `i`'s `a:ln`, creating one if absent, and rebuild the chart
part once. `f` takes the `a:ln` element and returns its replacement.

The part is rebuilt only after `f` returns, so a throw part-way through a
composed transform leaves the file untouched.
"""
_set_series_line(c::Chart, i::Integer, f) =
    _set_series_shape(c, i, (sp, pfx) -> _sp_with_line(sp, (NS_A, "spPr"), ln -> f(ln, pfx), pfx))

"""
    setSeriesLineColor(c, i, color) -> Chart

Set the color of series `i`'s outline — the line itself, on a line or scatter
chart. $(_DOC_COLOR_ARG)

Note `:none` here leaves the `a:ln` in place with no fill — the line exists and
draws nothing, which is what Excel writes and is distinct from
`setSeriesLine(c, i, :inherit)`, which removes the outline entirely.
"""
setSeriesLineColor(c::Chart, i::Integer, color) =
    _set_series_line(c, i, (ln, pfx) -> _ln_with_color(ln, color, pfx))

"""
    setSeriesLineWidth(c, i, points) -> Chart

Set the width of series `i`'s outline, in points — the unit Excel's width box
uses — as a `Real` from 0 to 1584. `:inherit` removes the width so the chart
style supplies it.

DrawingML caps a line width at 1584 pt; a wider one is rejected rather than
written.

One of the single-property line setters; [`setSeriesLine`](@ref) sets several at
once in a single rebuild.
"""
setSeriesLineWidth(c::Chart, i::Integer, points) =
    _set_series_line(c, i, (ln, _) -> _ln_with_width(ln, points))

"""
    setSeriesLineDash(c, i, dash) -> Chart

Set the dash pattern of series `i`'s outline to one of the DrawingML presets,
$(_doc_syms(DASH_VALUES)), or Excel's names for them,
$(_doc_aliases(DASH_ALIASES)). A String works as well as a Symbol. `:inherit`
removes the pattern.

Excel's "Round Dot" is `:sysDot` and "Square Dot" is `:sysDash`; `:dot` and
`:dash` are different patterns, which Excel's dropdown does not offer.

One of the single-property line setters; see [`setSeriesLine`](@ref).
"""
setSeriesLineDash(c::Chart, i::Integer, dash) =
    _set_series_line(c, i, (ln, pfx) -> _ln_with_dash(ln, dash, pfx))

"""
    setSeriesLineCap(c, i, cap) -> Chart

Set how series `i`'s outline ends: $(_doc_syms(CAP_VALUES)), or Excel's names
$(_doc_aliases(CAP_ALIASES)). A String works as well as a Symbol. `:inherit`
removes the setting.

One of the single-property line setters; see [`setSeriesLine`](@ref).
"""
setSeriesLineCap(c::Chart, i::Integer, cap) =
    _set_series_line(c, i, (ln, _) -> _ln_with_cap(ln, cap))

"""
    setSeriesLineCompound(c, i, cmpd) -> Chart

Set the compound line type of series `i`'s outline — how many parallel strokes
it draws: $(_doc_syms(CMPD_VALUES)), or Excel's names
$(_doc_aliases(CMPD_ALIASES)). A String works as well as a Symbol. `:inherit`
removes the setting.

One of the single-property line setters; see [`setSeriesLine`](@ref).
"""
setSeriesLineCompound(c::Chart, i::Integer, cmpd) =
    _set_series_line(c, i, (ln, _) -> _ln_with_compound(ln, cmpd))

"""
    setSeriesLineJoin(c, i, join) -> Chart

Set how the segments of series `i`'s outline meet: `:round`, `:bevel` or
`:miter` (or the same as a String). `:inherit` removes the join. Excel uses the
same three words, so there are no aliases.

A miter join takes a limit as well, set separately with
[`setSeriesLineMiterLimit`](@ref); changing the join afterwards discards it.

One of the single-property line setters; see [`setSeriesLine`](@ref).
"""
setSeriesLineJoin(c::Chart, i::Integer, join) =
    _set_series_line(c, i, (ln, pfx) -> _ln_with_join(ln, join, pfx))

"""
    setSeriesLineMiterLimit(c, i, limit) -> Chart

Set the miter limit of series `i`'s outline, as a `Real` multiple of the line
width: how far a sharp corner may extend before the join is cut off. Excel's
default is 8. `:inherit` removes the limit.

The limit belongs to the `a:miter` element, so the series' outline must already
have a miter join ([`setSeriesLineJoin`](@ref)); otherwise this throws.

One of the single-property line setters; see [`setSeriesLine`](@ref).
"""
setSeriesLineMiterLimit(c::Chart, i::Integer, limit) =
    _set_series_line(c, i, (ln, _) -> _ln_with_miter_limit(ln, limit))

"""
    setSeriesLine(c, i; color, width, dash, cap, compound, join, miterLimit) -> Chart
    setSeriesLine(c, i, :none)
    setSeriesLine(c, i, :inherit)

Set several line properties of series `i`'s outline at once, in one rebuild of
the chart part. A keyword left unspecified is left alone; pass `:inherit` to
remove one that is set.

# Keywords

These are the line keywords of every multi-property line setter —
[`setAxisLine`](@ref), [`setPlotAreaLine`](@ref) and the text `:line` field
take the same. Each matches a single-property setter, whose docstring has more.

| keyword      | value                                                        | setter                              |
|:-------------|:-------------------------------------------------------------|:------------------------------------|
| `color`      | any color value (see [`setSeriesFill`](@ref)), or `:none`    | [`setSeriesLineColor`](@ref)        |
| `width`      | points, a `Real` from 0 to 1584                              | [`setSeriesLineWidth`](@ref)        |
| `dash`       | $(_doc_syms(DASH_VALUES)); or $(_doc_aliases(DASH_ALIASES)) | [`setSeriesLineDash`](@ref)         |
| `cap`        | $(_doc_syms(CAP_VALUES)); or $(_doc_aliases(CAP_ALIASES))   | [`setSeriesLineCap`](@ref)          |
| `compound`   | $(_doc_syms(CMPD_VALUES)); or $(_doc_aliases(CMPD_ALIASES)) | [`setSeriesLineCompound`](@ref)     |
| `join`       | $(_doc_syms(JOIN_VALUES))                                    | [`setSeriesLineJoin`](@ref)         |
| `miterLimit` | a `Real` multiple of the line width; Excel's default is 8    | [`setSeriesLineMiterLimit`](@ref)   |

Every keyword also takes `:inherit`. The symbol-valued keywords accept the same
words as Strings.

`miterLimit` needs a miter join. Passing `join = :miter` and `miterLimit` in one
call works, since the join is applied first.

The symbol form acts on the whole outline: `:none` writes an `a:ln` whose fill is
`<a:noFill/>`, and `:inherit` removes the `a:ln` so the chart style supplies it.
"""
function setSeriesLine(c::Chart, i::Integer;
                       color = nothing, width = nothing, dash = nothing,
                       cap = nothing, compound = nothing,
                       join = nothing, miterLimit = nothing)
    all(isnothing, (color, width, dash, cap, compound, join, miterLimit)) && return c

    return _set_series_line(c, i, (ln, pfx) ->
        _ln_with(ln, pfx; color, width, dash, cap, compound, join, miterLimit))
end

function setSeriesLine(c::Chart, i::Integer, what::Symbol)
    what === :inherit && return _set_series_shape(c, i, (sp, _) -> remove_child(sp, "ln"))
    what === :none    && return setSeriesLineColor(c, i, :none)
    throw(XLSXError("`$what` is not a line instruction; use `:none` or `:inherit`."))
end

"""
    _marker_with_symbol(mk, symbol, pfx) -> XML.Node

Set a `c:marker`'s `c:symbol`. `:none` is a symbol in its own right — a marker
explicitly drawn as nothing — while `:inherit` removes the element so the chart
style decides. `:auto` is DrawingML's own "let Excel choose".
"""
function _marker_with_symbol(mk::XML.Node, symbol, pfx::Dict{String,String})
    symbol === :inherit && return remove_child(mk, "symbol")
    node = XML.Element(prefixed_tag(pfx[NS_C], "symbol");
                       val = String(_check(symbol, MARKER_SYMBOLS, "marker symbol")))
    return insert_child(mk, (NS_C, "marker"), node)
end

"""
    _marker_with_size(mk, size, pfx) -> XML.Node

Set a `c:marker`'s `c:size`, in points. `ST_MarkerSize` allows 2 to 72.
`:inherit` removes it; Excel's default is 5.
"""
function _marker_with_size(mk::XML.Node, size, pfx::Dict{String,String})
    size === :inherit && return remove_child(mk, "size")
    n = round(Int, size)
    2 <= n <= 72 || throw(XLSXError(
        "A marker size of $size is outside the range DrawingML allows (2 to 72)."))
    return insert_child(mk, (NS_C, "marker"),
                        XML.Element(prefixed_tag(pfx[NS_C], "size"); val = string(n)))
end

"""
    _set_marker(c, i, point, f) -> Chart

Apply `f` to a `c:marker`, creating one if absent, and rebuild the chart part
once. `point` selects a data point's marker, counting from 1; `nothing` selects
the series'.

A point with no `c:dPt` gets one, written with its `c:idx` and placed in index
order among the series' other `c:dPt` elements.
"""
function _set_marker(c::Chart, i::Integer, point::Union{Nothing,Integer}, f)
    root = chart_root(c)
    pfx  = ns_prefixes(root)
    steps, ser_key = _series_path(c, root, i)
    marker_in(parent, key, g = f) = rebuild_path(parent, [(NS_C, "marker") => "marker"],
                                                 mk -> g(mk, pfx); prefixes = pfx, parent_key = key)
    # A point given a marker of its own for the first time starts from the series'
    # marker size, as Excel writes it; otherwise Excel draws it at its default size,
    # however large the series' markers are. A size given in `f` still wins.
    function point_marker(ser, dp)
        sz = _marker_size_val(ser)
        (isnothing(sz) || !isnothing(first_element_with_tag(dp, "marker"))) && return marker_in(dp, (NS_C, "dPt"))
        return marker_in(dp, (NS_C, "dPt"), (mk, p) -> f(_marker_with_size(mk, sz, p), p))
    end
    new = rebuild_path(root, steps,
              ser -> isnothing(point) ? marker_in(ser, ser_key) :
                     _with_dpt(ser, ser_key, point, pfx, dp -> point_marker(ser, dp));
              prefixes = pfx)
    set_chart_root!(c, new)
    return c
end

# The `c:size` of a series' own `c:marker`, or `nothing` if it sets none.
function _marker_size_val(ser)
    mk = first_element_with_tag(ser, "marker")
    return isnothing(mk) ? nothing : _int_val(mk, "size")
end

"""
    setMarkerSymbol(c, i, symbol; point = nothing) -> Chart
 
Set the marker shape for series `i`, or for one of its data points when `point`
is given, counting from 1. `symbol` is one of $(_doc_syms(MARKER_SYMBOLS)), as
a Symbol or String.
 
`:none` draws no marker and `:auto` lets Excel choose; both are settings.
`:inherit` removes the element so the chart style decides.
"""
setMarkerSymbol(c::Chart, i::Integer, symbol; point::Union{Nothing,Integer} = nothing) =
    _set_marker(c, i, point, (mk, pfx) -> _marker_with_symbol(mk, symbol, pfx))

"""
    setMarkerSize(c, i, size; point = nothing) -> Chart
 
Set the marker size for series `i`, or for one of its data points when `point`
is given, counting from 1. `size` is in points, a `Real` rounded to the nearest
whole point. DrawingML allows 2 to 72; anything outside that is rejected rather
than written.
`:inherit` removes the setting so the chart style decides.
 
One of the single-property marker setters; [`setMarker`](@ref) sets several at
once in a single rebuild.
"""
setMarkerSize(c::Chart, i::Integer, size; point::Union{Nothing,Integer} = nothing) =
    _set_marker(c, i, point, (mk, pfx) -> _marker_with_size(mk, size, pfx))

_marker_with_fill(mk, color, pfx) =
    insert_child(mk, (NS_C, "marker"),
                 _sp_with_fill(_marker_sp(mk, pfx), (NS_A, "spPr"), color, pfx))

_marker_sp(mk, pfx) = something(first_element_with_tag(mk, "spPr"),
                                XML.Element(prefixed_tag(pfx[NS_C], "spPr")))

"""
    setMarkerFill(c, i, color; point = nothing) -> Chart
 
Set the fill of series `i`'s marker, or of one data point's when `point` is
given, counting from 1. $(_DOC_COLOR_ARG)
 
One of the single-property marker setters; see [`setMarker`](@ref).
"""
setMarkerFill(c::Chart, i::Integer, color; point::Union{Nothing,Integer} = nothing) =
    _set_marker(c, i, point, (mk, pfx) -> _marker_with_fill(mk, color, pfx))

"""
    _marker_with_line(mk, f, pfx) -> XML.Node

Apply `f` to the `a:ln` of a `c:marker`'s `c:spPr`, creating both if absent.
"""
_marker_with_line(mk::XML.Node, f, pfx::Dict{String,String}) =
    insert_child(mk, (NS_C, "marker"),
                 _sp_with_line(_marker_sp(mk, pfx), (NS_A, "spPr"), f, pfx))

"""
    setMarkerLineColor(c, i, color; point = nothing) -> Chart
 
Set the outline color of series `i`'s marker, or of one data point's when
`point` is given, counting from 1. $(_DOC_COLOR_ARG)
 
As with [`setSeriesLineColor`](@ref), `:none` leaves the outline in place
drawing nothing, which is not the same as removing it.
 
One of the single-property marker setters; see [`setMarker`](@ref).
"""
setMarkerLineColor(c::Chart, i::Integer, color; point::Union{Nothing,Integer} = nothing) =
    _set_marker(c, i, point, (mk, pfx) ->
        _marker_with_line(mk, ln -> _ln_with_color(ln, color, pfx), pfx))

"""
    setMarkerLineWidth(c, i, width; point = nothing) -> Chart
 
Set the outline width of series `i`'s marker, or of one data point's when
`point` is given, counting from 1. `width` is in points, a `Real` from 0 to
1584. `:inherit` removes the width.
 
DrawingML caps a line width at 1584 pt; a wider one is rejected rather than
written.
 
One of the single-property marker setters; see [`setMarker`](@ref).
"""
setMarkerLineWidth(c::Chart, i::Integer, width; point::Union{Nothing,Integer} = nothing) =
    _set_marker(c, i, point, (mk, pfx) ->
        _marker_with_line(mk, ln -> _ln_with_width(ln, width), pfx))

"""
    setMarker(c, i; point, symbol, size, fill, lineColor, lineWidth) -> Chart
 
Set several marker properties at once, in one rebuild of the chart part. Sets
the marker for series `i`, or overrides it for one of its data points when
`point` is given, counting from 1.
 
A keyword left unspecified is left alone; pass `:inherit` to remove one that is
set. Each keyword takes the values its single-property setter does:

- `symbol` — $(_doc_syms(MARKER_SYMBOLS)) ([`setMarkerSymbol`](@ref));
- `size` — points, 2 to 72 ([`setMarkerSize`](@ref));
- `fill` — a color value, as for [`setSeriesFill`](@ref) ([`setMarkerFill`](@ref));
- `lineColor` — a color value ([`setMarkerLineColor`](@ref));
- `lineWidth` — points, 0 to 1584 ([`setMarkerLineWidth`](@ref)).
 
Nothing is written if every property keyword is `nothing`.
 
# Example
 
    setMarker(c, 1; symbol = :circle, size = 7, fill = "FF0000", lineColor = :none)
"""
function setMarker(c::Chart, i::Integer;
                   point::Union{Nothing,Integer} = nothing,
                   symbol = nothing, size = nothing, fill = nothing,
                   lineColor = nothing, lineWidth = nothing)
    all(isnothing, (symbol, size, fill, lineColor, lineWidth)) && return c
    return _set_marker(c, i, point, function (mk, pfx)
        isnothing(symbol) || (mk = _marker_with_symbol(mk, symbol, pfx))
        isnothing(size)   || (mk = _marker_with_size(mk, size, pfx))
        isnothing(fill)   || (mk = _marker_with_fill(mk, fill, pfx))
        if !isnothing(lineColor) || !isnothing(lineWidth)
            mk = _marker_with_line(mk, function (ln)
                isnothing(lineColor) || (ln = _ln_with_color(ln, lineColor, pfx))
                isnothing(lineWidth) || (ln = _ln_with_width(ln, lineWidth))
                return ln
            end, pfx)
        end
        return mk
    end)
end

# ── Group setters ─────────────────────────────────────────────────────────────

function _group_path(c::Chart, root::XML.Node, g::ChartGroup)
    tag = String(g.kind)
    els = [el for el in _plotarea_children(root)
           if localname(el) == tag && _group_axids(el) == g.axids]
    length(els) == 1 || throw(XLSXError(
        "Chart `$(c.name)` has $(length(els)) `$tag` groups on axes $(g.axids); expected one."))
    el  = only(els)
    key = (NS_C, tag)
    steps = [(NS_C, "chart")    => "chart",
             (NS_C, "plotArea") => "plotArea",
             key                => (tag, n -> n === el)]
    return steps, key, el
end

function _set_group(c::Chart, g::ChartGroup, f)
    root = chart_root(c)
    pfx  = ns_prefixes(root)
    steps, key, _ = _group_path(c, root, g)
    new = rebuild_path(root, steps, el -> f(el, key, pfx); prefixes = pfx)
    set_chart_root!(c, new)
    return c
end

function _group_node(c::Chart, g::ChartGroup)
    _, _, el = _group_path(c, chart_root(c), g)
    return el
end

"""
    getGroupGapWidth(c::Chart, g::ChartGroup) -> Union{Nothing,Int}

`c:gapWidth` — the space between bars or clusters, as a percentage of a bar's
width (0–500). `nothing` means not written; Excel then uses 150. Bar and
bar-of-pie groups only.
"""
function getGroupGapWidth(c::Chart, g::ChartGroup)
    el = _group_node(c, g)
    _require_kind(el, (:barChart, :bar3DChart, :ofPieChart), "gapWidth")
    return _int_val(el, "gapWidth")
end

"""
    setGroupGapWidth(c, g, percent) -> Chart

Set the group's gap width, an `Integer` percentage of a bar's width from 0 to
500, as Excel's "Gap Width" box; or `:inherit` to remove it. `g` is a group
from [`getChartGroups`](@ref); bar and bar-of-pie groups only.
"""
function setGroupGapWidth(c::Chart, g::ChartGroup, pct)
    v = _check_pct(pct, 0, 500, "A gap width")
    return _set_group(c, g, (el, key, pfx) -> begin
        _require_kind(el, (:barChart, :bar3DChart, :ofPieChart), "gapWidth")
        _with_val(el, key, pfx, "gapWidth", v)
    end)
end

"""
    getGroupOverlap(c::Chart, g::ChartGroup) -> Union{Nothing,Int}

`c:overlap` — how far bars in a cluster overlap, as a percentage from -100
(separated by a bar's width) to 100 (fully overlapping, as stacked charts need).
`nothing` means not written. 2-D bar groups only.
"""
function getGroupOverlap(c::Chart, g::ChartGroup)
    el = _group_node(c, g)
    _require_kind(el, (:barChart,), "overlap")
    return _int_val(el, "overlap")
end

"""
    setGroupOverlap(c, g, percent) -> Chart

Set the group's overlap, an `Integer` percentage from -100 to 100, as Excel's
"Series Overlap" box; or `:inherit` to remove it. `g` is a group from
[`getChartGroups`](@ref); 2-D bar groups only. Stacked groups need 100 for each
series to sit on the one before it.
"""
function setGroupOverlap(c::Chart, g::ChartGroup, pct)
    v = _check_pct(pct, -100, 100, "An overlap")
    return _set_group(c, g, (el, key, pfx) -> begin
        _require_kind(el, (:barChart,), "overlap")
        _with_val(el, key, pfx, "overlap", v)
    end)
end

