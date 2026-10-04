# Appearance accessors for `c:` charts.
#
# This file sits above `drawingml.jl` in the layering: that file parses
# DrawingML — `spPr`, `txPr`, colours, fills, lines, text — for anything that
# carries it, knowing nothing about charts. This file knows the chart schema and
# finds the nodes: which element holds a series' fill, which axes a group plots
# against, where a per-point override lives. Nothing here parses DrawingML
# itself; it locates a node and hands it to `drawingml.jl`.
#
# Entry point. Every accessor starts from a `Chart`, a handle that reaches its
# XML through `chart_root(c)` on each call, so it never goes stale. Everything
# read from a chart — `ChartSeries`, `ChartAxis`, `ChartGroup`, data points,
# labels, trendlines, error bars, markers — is a value carrying a key: `c:idx`
# for a series or point, `c:axId` for an axis, `(kind, axids)` for a group, the
# owning series plus an ordinal for trendlines and error bars. Any function
# taking `(c, x)` finds `x`'s element by that key in the current part, getters
# and setters alike, so a value read before a write still addresses the right
# element afterwards. Its other fields describe the part as it was when read.
#
# `raw` is kept for partial modelling and is never an address. A setter
# locates its target in the root it is about to rebuild and builds its path
# predicates from nodes of that same root, never from a snapshot.
#
# Indexing is 1-based throughout, over the series in document order, as 
# getChartSeries returns them and over data points as the user sees them. 
# Excel's own identifiers — `c:idx`, `c:order`, `c:axId` — are kept on the 
# structs but are not positions: they need not be contiguous, and a file 
# where series or points were deleted will have gaps. `c:axId` is the 
# exception that is genuinely useful as a key, since `c:crossAx` and a group's 
# `c:axId` children reference it; `getChartAxis(c, axid)` looks up by it. 
# `chart_idx_gaps.xlsx` is the fixture where positions and `c:idx` differ.
#
# Absent is not explicit. Every optional value is `Union{Nothing,T}`, and
# `nothing` means the element or attribute was not written — which in this
# schema means "inherit" or "Excel decides", never "off". An explicit
# `<a:noFill/>` means deliberately off, `<c:delete val="1"/>` means deliberately
# hidden. `hasFill`/`hasLine` exist because conflating the two at a call site
# is easy and wrong. This matters for writing as much as reading: emitting a
# `val="0"` into a chart that never had one is a change to the file for no
# reason.
#
# Two ambiguities in the schema worth knowing before adding accessors here:
#
#   - `c:marker` is two different elements. Under `c:ser` or `c:dPt` it is
#     marker properties; under `c:lineChart` or `c:scatterChart` it is a boolean
#     show-markers flag. `_parse_marker` must only be called with the former.
#   - `c:dLbls` exists at three levels — series, group, and individual `c:dLbl`
#     — and they form an inheritance chain. A retyped label carries the same
#     formatting in both `c:tx/c:rich` and `c:txPr`, so a write must update both
#     or Excel shows the `rich` version and the edit appears to do nothing.
#
# Two kinds of accessor live here. Most report what one node says, and
# `nothing` means that node is silent, not that a default applies. The
# cascade resolvers — `getSeriesFill`, `getSeriesLine`, `getMarkerFill`,
# `getLabelTextProp` — walk the rungs instead and return an `Effective`
# carrying the value, the rung that answered and the whole chain.
#
# Cascade resolution deliberately stops at the chart XML. Where a property is
# written at no rung, `Effective.value` is `nothing`, which means Excel takes it
# from `xl/charts/style1.xml` and `colors1.xml`. 
# The reason is that Excel's precedence between them is underdocumented:
# style1.xml is keyed by a series-index-modulo scheme that varies by chart type,
# and colors1.xml layers on top. Pinning that down is research in its own right.
# `FormatSite.level` already accepts `:style`, so adding the rung later is
# additive, not breaking.
#
# The reason is that Excel's precedence between them is underdocumented:
# style1.xml is keyed by a series-index-modulo scheme that varies by chart type,
# and colors1.xml layers on top. Pinning that down is research, and blocking the
# write path on it would be the wrong trade. `FormatSite.level` already accepts
# `:style` so adding the rung later is additive, not breaking.

const AXIS_TAGS = ("catAx", "valAx", "dateAx", "serAx")

const AXIS_ROLES = Dict(
    :category => "catAx",
    :value    => "valAx",
    :date     => "dateAx",
    :series   => "serAx",
)

function _axis_node(c::Chart, root::XML.Node, ax::ChartAxis)
    isnothing(ax.axid) && throw(XLSXError(
        "This axis has no `c:axId`, which the schema requires, so it cannot be located."))
    for el in _plotarea_children(root)
        localname(el) in AXIS_TAGS && _axis_id(el, "axId") == ax.axid && return el
    end
    throw(XLSXError("Chart `$(c.name)` has no axis with axId $(ax.axid)."))
end

# The axis's node in the current part, for getters. Setters call _axis_node with
# the root they will rebuild instead.
_axnode(c::Chart, ax::ChartAxis) = _axis_node(c, chart_root(c), ax)

# c:axId and c:crossAx carry unsigned 32-bit values, which Excel writes near the
# top of the range. Parse to Int (64-bit on every platform we support) rather
# than Int32 — see the 32-bit portability bug in table.jl.
function _axis_id(el::Union{Nothing,XML.Node}, tag::AbstractString)
    v = _attr(first_element_with_tag(el, tag), "val")
    return isnothing(v) ? nothing : parse(Int, v)
end

function parse_chart_axis(el::XML.Node)::ChartAxis
    pos = _attr(first_element_with_tag(el, "axPos"), "val")
    del = _attr(first_element_with_tag(el, "delete"), "val")
    return ChartAxis(
        Symbol(localname(el)),
        _axis_id(el, "axId"),
        isnothing(pos) ? nothing : Symbol(pos),
        _axis_id(el, "crossAx"),
        del in ("1", "true"),
        el,
    )
end

"""
    getChartAxes(c::Chart) -> Vector{ChartAxis}
    getChartAxes(c::Chart, role::Symbol) -> Vector{ChartAxis}

Axes of `c`, in `c:plotArea` document order. `role` is `:category`, `:value`,
`:date` or `:series` and filters by axis kind; a combo chart has two value axes,
so this returns a vector rather than one axis.

Axes with `c:delete val="1"` are included. Excel does not draw them, but they
remain in the XML and remain formattable — check `ax.deleted` to skip them.
"""
getChartAxes(c::Chart)::Vector{ChartAxis} =
    [parse_chart_axis(el) for el in _plotarea_children(chart_root(c)) if localname(el) in AXIS_TAGS]

function getChartAxes(c::Chart, role::Symbol)::Vector{ChartAxis}
    haskey(AXIS_ROLES, role) || throw(XLSXError(
        "Unknown axis role `:$role`. Use one of " *
        join((":$k" for k in keys(AXIS_ROLES)), ", ") * "."))
    tag = AXIS_ROLES[role]
    return [ax for ax in getChartAxes(c) if String(ax.kind) == tag]
end

"""
    getChartAxis(c::Chart, axid::Integer) -> ChartAxis

The axis with `c:axId` equal to `axid`. This is the identifier chart groups and
`c:crossAx` use, so it is how you get from a series group, or from one axis, to
its partner.
"""
function getChartAxis(c::Chart, axid::Integer)::ChartAxis
    axes = getChartAxes(c)
    i = findfirst(ax -> ax.axid == axid, axes)
    isnothing(i) && throw(XLSXError(
        "Chart `$(c.name)` has no axis with axId $axid. Found: " *
        join((string(something(ax.axid, "?")) for ax in axes), ", ") * "."))
    return axes[i]
end

"""
    getAxisShapeProps(c::Chart, ax::ChartAxis) -> Union{Nothing,DrawingShapeProps}

Fill and line for the axis itself (`c:spPr`) — the axis line and its tick marks.
`nothing` means no `c:spPr` was written, so Excel takes the formatting from the
chart style. Gridlines are separate; see [`getAxisGridlines`](@ref).
"""
getAxisShapeProps(c::Chart, ax::ChartAxis) = parse_drawing_shape_props(_wb(c), _axnode(c, ax))

"""
    getAxisTextProps(c::Chart, ax::ChartAxis) -> Union{Nothing,DrawingText}

Text body for the axis tick labels (`c:txPr`). `nothing` means none was written.
The axis title's text is separate; see [`getAxisTitleText`](@ref).
"""
getAxisTextProps(c::Chart, ax::ChartAxis)  = parse_drawing_text(_wb(c), _axnode(c, ax))

"""
    getAxisTitleRef(c::Chart, ax::ChartAxis) -> Union{Nothing,String}

The formula behind an axis title that references a cell
(`c:title/c:tx/c:strRef/c:f`). `nothing` for a literal or absent title.
"""
function getAxisTitleRef(c::Chart, ax::ChartAxis)
    tx = first_element_with_tag(first_element_with_tag(_axnode(c, ax), "title"), "tx")
    sr = first_element_with_tag(tx, "strRef")
    return isnothing(sr) ? nothing : child_text(sr, "f")
end

"""
    getAxisTitleText(c::Chart, ax::ChartAxis) -> Union{Nothing,DrawingText}

The axis title's text body (`c:title/c:tx/c:rich`). `nothing` means no title
element, or a title that references a cell rather than carrying literal text —
`c:tx/c:strRef` — which this does not resolve.
"""
function getAxisTitleText(c::Chart, ax::ChartAxis)
    tx = first_element_with_tag(first_element_with_tag(_axnode(c, ax), "title"), "tx")
    return parse_drawing_text(_wb(c), tx; tag="rich")
end

# An element whose presence turns a feature on, carrying an optional `c:spPr`.
#
# These are the three-state cases: absent means the feature is off, present with
# formatting means on and formatted, present without means on and drawn with
# inherited formatting. Returning `nothing` for the first and a
# `DrawingShapeProps` for both others keeps "is it on?" a single `isnothing`
# check at the call site, and the returned props still distinguish an absent
# fill from an explicit `<a:noFill/>` in the usual way.
function _optional_spPr(c::Chart, parent::Union{Nothing,XML.Node}, tag::AbstractString)
    el = first_element_with_tag(parent, tag)
    isnothing(el) && return nothing
    return something(parse_drawing_shape_props(_wb(c), el),
                     DrawingShapeProps(nothing, nothing, nothing, nothing, el))
end

"""
    getAxisGridlines(c::Chart, ax::ChartAxis; minor=false) -> Union{Nothing,DrawingShapeProps}

Line properties for the axis gridlines (`c:majorGridlines`, or
`c:minorGridlines` when `minor` is true).

`nothing` means no gridlines element, so none are drawn. An element present with
no `c:spPr` means gridlines are drawn with inherited formatting, and comes back
as a `DrawingShapeProps` with every field absent — so `isnothing` answers "are
there gridlines?" and the returned value answers "how are they formatted?".
"""
getAxisGridlines(c::Chart, ax::ChartAxis; minor::Bool=false) =
    _optional_spPr(c, _axnode(c, ax), minor ? "minorGridlines" : "majorGridlines")

# Checks the kind the axis has now, not when `ax` was read: Excel keeps the axId
# when a category axis is switched to a date axis.
function _require_kind(n::XML.Node, kinds::Tuple, what::AbstractString)
    k = Symbol(localname(n))
    k in kinds || throw(XLSXError(
        "$what is only defined on $(join(kinds, " or ")); this is a $k."))
end

"""
    getAxisPartner(c::Chart, ax::ChartAxis) -> Union{Nothing,ChartAxis}

The axis that `ax` crosses (`c:crossAx`). `nothing` if unwritten or dangling.
"""
function getAxisPartner(c::Chart, ax::ChartAxis)
    cross = _axis_id(_axnode(c, ax), "crossAx")
    isnothing(cross) && return nothing
    axes = getChartAxes(c)
    i = findfirst(a -> a.axid == cross, axes)
    return isnothing(i) ? nothing : axes[i]
end

_wb(c::AbstractChart) = get_workbook(c.package)


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
    getChartLegend(c::Chart) -> Union{Nothing,XML.Node}

The `c:legend` element, or `nothing` if the chart has no legend. Returned as a
node because the legend carries no identity of its own — pass it to
[`getLegendShapeProps`](@ref) and friends, or read it directly.
"""
getChartLegend(c::Chart) =
    first_element_with_tag(first_element_with_tag(chart_root(c), "chart"), "legend")

"""
    getLegendPos(c::Chart) -> Union{Nothing,Symbol}

`:b` (bottom), `:t` (top), `:l` (left), `:r` (right) or `:tr` (top-right) from
`c:legendPos`. `nothing` means no legend,
or a legend with no explicit position.
"""
function getLegendPos(c::Chart)
    v = _attr(first_element_with_tag(getChartLegend(c), "legendPos"), "val")
    return isnothing(v) ? nothing : Symbol(v)
end

"""
    getLegendOverlay(c::Chart) -> Union{Nothing,Bool}

`c:overlay` — `true` where the legend is drawn on top of the plot area rather
than beside it. `nothing` means the chart has no legend, or the element was not
written, in which case Excel treats it as `false`.
"""
getLegendOverlay(c::Chart) = _bool_val(getChartLegend(c), "overlay")

"""
    getLegendShapeProps(c::Chart) -> Union{Nothing,DrawingShapeProps}

Fill and border of the legend box (`c:legend/c:spPr`). `nothing` means the chart
has no legend, or none was written.
"""
getLegendShapeProps(c::Chart) = parse_drawing_shape_props(_wb(c), getChartLegend(c))

"""
    getLegendTextProps(c::Chart) -> Union{Nothing,DrawingText}
    getLegendTextProps(c::ChartEx) -> Union{Nothing,DrawingText}

Text body for the legend entries (`c:legend/c:txPr`). `nothing` means the chart
has no legend, or none was written.
"""
getLegendTextProps(c::Chart)  = parse_drawing_text(_wb(c), getChartLegend(c))

_chartnode(c::Chart) = first_element_with_tag(chart_root(c), "chart")

getChartTitleNode(c::Chart) = first_element_with_tag(_chartnode(c), "title")

"""
    getChartTitleText(c::Chart) -> Union{Nothing,DrawingText}

Literal title text (`c:title/c:tx/c:rich`). `nothing` if there is no title
element, or the title references a cell rather than carrying literal text — see
[`getChartTitleRef`](@ref) — or `c:autoTitleDeleted` is 0 and Excel generates the
title from the series name, in which case no text exists in the file at all.
"""
getChartTitleText(c::Chart) =
    parse_drawing_text(_wb(c),
        first_element_with_tag(getChartTitleNode(c), "tx"); tag="rich")

"""
    getChartTitleRef(c::Chart) -> Union{Nothing,String}

The formula behind a title that references a cell (`c:title/c:tx/c:strRef/c:f`).
`nothing` for a literal or absent title.
"""
function getChartTitleRef(c::Chart)
    tx = first_element_with_tag(getChartTitleNode(c), "tx")
    sr = first_element_with_tag(tx, "strRef")
    return isnothing(sr) ? nothing : child_text(sr, "f")
end

"""
    getChartTitleShapeProps(c::Chart) -> Union{Nothing,DrawingShapeProps}

Fill and border of the title box (`c:title/c:spPr`). `nothing` means there is no
title element, or none was written.
"""
getChartTitleShapeProps(c::Chart) = parse_drawing_shape_props(_wb(c), getChartTitleNode(c))

"""
    getChartTitleTextProps(c::Chart) -> Union{Nothing,DrawingText}
    getChartTitleTextProps(c::ChartEx) -> Union{Nothing,DrawingText}

Text body for the title (`c:title/c:txPr`) — its formatting, which applies
whether the title is typed, bound to a cell or generated by Excel. Contrast
[`getChartTitleText`](@ref), which is the literal text in `c:tx/c:rich`.
"""
getChartTitleTextProps(c::Chart)  = parse_drawing_text(_wb(c), getChartTitleNode(c))

"""
    getAutoTitleDeleted(c::Chart) -> Union{Nothing,Bool}

`c:autoTitleDeleted` — `true` where the user removed the title Excel would
otherwise generate from the single series' name. `nothing` means the element was
not written, which Excel treats as `false`.
"""
getAutoTitleDeleted(c::Chart) = _bool_val(_chartnode(c), "autoTitleDeleted")

"""
    getPlotAreaShapeProps(c::Chart) -> Union{Nothing,DrawingShapeProps}

Fill and border of the plot area (`c:plotArea/c:spPr`) — the region the data is
drawn in, inside the axes. `nothing` means none was written.
"""
getPlotAreaShapeProps(c::Chart) =
    parse_drawing_shape_props(_wb(c), first_element_with_tag(_chartnode(c), "plotArea"))

"""
    getChartSpaceShapeProps(c::Chart) -> Union{Nothing,DrawingShapeProps}

Fill and border of the chart as a whole (`c:chartSpace/c:spPr`) — the outer
frame, not the plot area. `nothing` means none was written.
"""
getChartSpaceShapeProps(c::Chart) = parse_drawing_shape_props(_wb(c), chart_root(c))

"""
    getChartSpaceTextProps(c::Chart) -> Union{Nothing,DrawingText}

Text body at the chart space (`c:chartSpace/c:txPr`), the default every other
text element inherits from and the last rung of the text cascade; see
[`getLabelTextProp`](@ref). `nothing` means none was written.
"""
getChartSpaceTextProps(c::Chart)  = parse_drawing_text(_wb(c), chart_root(c))

# --- axis scalars ---------------------------------------------------------
#
# All of these are elements carrying a `val` attribute, not attributes on the
# axis — the "elements masquerading as attributes" pattern. `nothing` means the
# element was not written, which for an axis means Excel applies its own default
# rather than that the feature is off.

# Read a `val` attribute from a child element of `ax`, as a Symbol.
_axis_sym(n::XML.Node, tag::AbstractString) =
    (v = _attr(first_element_with_tag(n, tag), "val"); isnothing(v) ? nothing : Symbol(v))


# Read a `val` attribute from a child element of `ax`, as a Float64.
# Numeric axis values are xsd:double throughout; dates on a dateAx arrive as
# serial numbers, matching how the rest of the package handles date cells.
_axis_num(n::XML.Node, tag::AbstractString) =
    (v = _attr(first_element_with_tag(n, tag), "val"); isnothing(v) ? nothing : parse(Float64, v))


# As `_axis_num`, but reading from a child of `c:scaling` rather than the axis.
function _scaling_num(sc::XML.Node, tag::AbstractString)
    v = _attr(first_element_with_tag(sc, tag), "val")
    return isnothing(v) ? nothing : parse(Float64, v)
end

"""
    getAxisNumberFormatCode(c::Chart, ax::ChartAxis) -> Union{Nothing,String}

The format code from `c:numFmt/@formatCode`. Ignored by Excel when
[`getAxisNumberFormatLinked`](@ref) is `true`.
"""
getAxisNumberFormatCode(c::Chart, ax::ChartAxis) =
    _attr(first_element_with_tag(_axnode(c, ax), "numFmt"), "formatCode")

"""
    getAxisNumberFormatLinked(c::Chart, ax::ChartAxis) -> Union{Nothing,Bool}

`c:numFmt/@sourceLinked`. `true` means Excel takes the number format from the
source cells and ignores the axis's own format code. `nothing` means no
`c:numFmt` element was written.
"""
function getAxisNumberFormatLinked(c::Chart, ax::ChartAxis)
    el = first_element_with_tag(_axnode(c, ax), "numFmt")
    isnothing(el) && return nothing
    return _attr(el, "sourceLinked") in ("1", "true")
end

"""
    getAxisMajorTickMark(c::Chart, ax::ChartAxis) -> Union{Nothing,Symbol}

`c:majorTickMark` — `:cross`, `:in`, `:out` or `:none`.
"""
getAxisMajorTickMark(c::Chart, ax::ChartAxis) = _axis_sym(_axnode(c, ax), "majorTickMark")


"""
    getAxisMinorTickMark(c::Chart, ax::ChartAxis) -> Union{Nothing,Symbol}

`c:minorTickMark` — `:cross`, `:in`, `:out` or `:none`.
"""
getAxisMinorTickMark(c::Chart, ax::ChartAxis) = _axis_sym(_axnode(c, ax), "minorTickMark")

"""
    getAxisTickLabelPos(c::Chart, ax::ChartAxis) -> Union{Nothing,Symbol}

`c:tickLblPos` — `:high`, `:low`, `:nextTo` or `:none`.
"""
getAxisTickLabelPos(c::Chart, ax::ChartAxis)  = _axis_sym(_axnode(c, ax), "tickLblPos")

"""
    getAxisOrientation(c::Chart, ax::ChartAxis) -> Union{Nothing,Symbol}

`c:scaling/c:orientation` — `:minMax` for a normal axis, `:maxMin` for one
plotted in reverse order.

Throws if the axis has no `c:scaling`: the schema requires it, so its absence
means a malformed chart part rather than an inherited value.
"""
getAxisOrientation(c::Chart, ax::ChartAxis) = _axis_sym(_axis_scaling(_axnode(c, ax)), "orientation")

"""
    getAxisMin(c::Chart, ax::ChartAxis) -> Union{Nothing,Float64}

Fixed lower bound from `c:scaling/c:min`. `nothing` means Excel scales the axis
automatically, which is the usual case; a value means the user fixed the bound.
"""
getAxisMin(c::Chart, ax::ChartAxis)     = _scaling_num(_axis_scaling(_axnode(c, ax)), "min")

"""
    getAxisMax(c::Chart, ax::ChartAxis) -> Union{Nothing,Float64}

Fixed upper bound from `c:scaling/c:max`. `nothing` means automatic.
"""
getAxisMax(c::Chart, ax::ChartAxis)     = _scaling_num(_axis_scaling(_axnode(c, ax)), "max")

"""
    getAxisLogBase(c::Chart, ax::ChartAxis) -> Union{Nothing,Float64}

`c:scaling/c:logBase`. `nothing` means a linear axis.
"""
getAxisLogBase(c::Chart, ax::ChartAxis) = _scaling_num(_axis_scaling(_axnode(c, ax)), "logBase")

# c:scaling is required by the schema, so a missing one is a malformed part,
# not an absent-means-inherit case.
function _axis_scaling(n::XML.Node)
    sc = first_element_with_tag(n, "scaling")
    isnothing(sc) && throw(XLSXError(
        "Axis $(something(_axis_id(n, "axId"), "?")) ($(localname(n))) has no `c:scaling` " *
        "element, which the schema requires. The chart part is malformed."))
    return sc
end

"""
    getAxisCrosses(c::Chart, ax::ChartAxis) -> Union{Nothing,Symbol}

Where this axis crosses its partner (`c:crosses`), as a position on the
partner: `:autoZero`, `:min` or `:max`. Mutually exclusive with
[`getAxisCrossesAt`](@ref) — at most one of the two is non-`nothing`.

Excel's "Horizontal axis crosses: At maximum category", set in the category
axis's pane, is stored on the value axis.
"""
getAxisCrosses(c::Chart, ax::ChartAxis)       = _axis_sym(_axnode(c, ax), "crosses")

"""
    getAxisCrossesAt(c::Chart, ax::ChartAxis) -> Union{Nothing,Float64}

Where this axis crosses its partner, as a value on the partner's scale
(`c:crossesAt`). Mutually exclusive with [`getAxisCrosses`](@ref).
"""
getAxisCrossesAt(c::Chart, ax::ChartAxis)     = _axis_num(_axnode(c, ax), "crossesAt")

"""
    getAxisMajorUnit(c::Chart, ax::ChartAxis) -> Union{Nothing,Float64}

Interval between major ticks and gridlines (`c:majorUnit`). `nothing` means
Excel chooses automatically. Value and date axes only.
"""
function getAxisMajorUnit(c::Chart, ax::ChartAxis)
    n = _axnode(c, ax)
    _require_kind(n, (:valAx, :dateAx), "majorUnit")
    return _axis_num(n, "majorUnit")
end

"""
    getAxisMinorUnit(c::Chart, ax::ChartAxis) -> Union{Nothing,Float64}

Interval between minor ticks (`c:minorUnit`). Value and date axes only.
"""
function getAxisMinorUnit(c::Chart, ax::ChartAxis)
    n = _axnode(c, ax)
    _require_kind(n, (:valAx, :dateAx), "minorUnit")
    return _axis_num(n, "minorUnit")
end

"""
    getAxisLabelAlign(c::Chart, ax::ChartAxis) -> Union{Nothing,Symbol}

`c:lblAlgn` — `:ctr` (centre), `:l` (left) or `:r` (right). Category and date
axes only.
"""
function getAxisLabelAlign(c::Chart, ax::ChartAxis)
    n = _axnode(c, ax)
    _require_kind(n, (:catAx, :dateAx), "lblAlgn")
    return _axis_sym(n, "lblAlgn")
end

"""
    getAxisLabelOffset(c::Chart, ax::ChartAxis) -> Union{Nothing,Int}

`c:lblOffset` — distance of the labels from the axis, as a percentage between 0
and 1000. Category and date axes only.
"""
function getAxisLabelOffset(c::Chart, ax::ChartAxis)
    n = _axnode(c, ax)
    _require_kind(n, (:catAx, :dateAx), "lblOffset")
    return _int_val(n, "lblOffset")
end

"""
    getAxisMultiLevelLabels(c::Chart, ax::ChartAxis) -> Union{Nothing,Bool}

Whether multi-level category labels are shown. This inverts the XML: the
element is `c:noMultiLvlLbl`, so `val="1"` means labels are *not* multi-level
and this returns `false`. Category and date axes only.
"""
function getAxisMultiLevelLabels(c::Chart, ax::ChartAxis)
    n = _axnode(c, ax)
    _require_kind(n, (:catAx, :dateAx), "noMultiLvlLbl")
    v = _bool_val(n, "noMultiLvlLbl")
    return isnothing(v) ? nothing : !v
end

"""
    getAxisCrossBetween(c::Chart, ax::ChartAxis) -> Union{Nothing,Symbol}

`c:crossBetween` — `:between` if the partner axis crosses between categories,
`:midCat` if it crosses at their midpoints. Value axes only.
"""
function getAxisCrossBetween(c::Chart, ax::ChartAxis)
    n = _axnode(c, ax)
    _require_kind(n, (:valAx,), "crossBetween")
    return _axis_sym(n, "crossBetween")
end

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

# Read a boolean `val` attribute from a child element. Returns `nothing` when
# the element is absent, which throughout the chart schema means "inherit" or
# "Excel's default" rather than false.
_bool_val(el::Union{Nothing,XML.Node}, tag::AbstractString) =
    (v = _attr(first_element_with_tag(el, tag), "val");
     isnothing(v) ? nothing : v in ("1", "true"))

_int_val(el::Union{Nothing,XML.Node}, tag::AbstractString) =
    (v = _attr(first_element_with_tag(el, tag), "val");
     isnothing(v) ? nothing : parse(Int, v))

_num_val(el::Union{Nothing,XML.Node}, tag::AbstractString) =
    (v = _attr(first_element_with_tag(el, tag), "val");
     isnothing(v) ? nothing : parse(Float64, v))

_sym_val(el::Union{Nothing,XML.Node}, tag::AbstractString) =
    (v = _attr(first_element_with_tag(el, tag), "val");
     isnothing(v) ? nothing : Symbol(v))

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

# The `c:dLbls` container on a series, group or data point.
_dlbls(parent::Union{Nothing,XML.Node}) = first_element_with_tag(parent, "dLbls")

"""
    getSeriesDataLabels(c::Chart, i::Integer) -> Vector{ChartDataLabel}

Individual label overrides on series `i`, in document order. Usually empty.
"""
function getSeriesDataLabels(c::Chart, i::Integer)::Vector{ChartDataLabel}
    _, ser = _series_pick(c, chart_root(c), i)
    sidx = _ser_idx(ser)
    out = ChartDataLabel[]
    dl = _dlbls(ser)
    isnothing(dl) && return out
    for el in elements_with_tag(dl, "dLbl")
        n = _int_val(el, "idx")
        isnothing(n) && continue
        push!(out, ChartDataLabel(sidx, n, _bool_val(el, "delete"), el))
    end
    return out
end

"""
    getSeriesDataLabel(c::Chart, i::Integer, point::Integer) -> Union{Nothing,ChartDataLabel}

The label override for data point `point` of series `i`, counting from 1, or
`nothing` if that label is not individually formatted.
"""
function getSeriesDataLabel(c::Chart, i::Integer, point::Integer)
    point >= 1 || throw(XLSXError("Data point positions start at 1; asked for $point."))
    dls = getSeriesDataLabels(c, i)
    j = findfirst(d -> d.idx == point - 1, dls)
    return isnothing(j) ? nothing : dls[j]
end

"""
    getDataLabelTextProps(c::Chart, d::ChartDataLabel) -> Union{Nothing,DrawingText}

Text body for one label (`c:dLbl/c:txPr`) — its formatting. `nothing` means the
label writes none, so the series' [`getSeriesLabelTextProps`](@ref) applies.
Contrast [`getDataLabelText`](@ref), which is typed-over text.
"""
getDataLabelTextProps(c::Chart, d::ChartDataLabel)  = parse_drawing_text(_wb(c), _node(c, d))

"""
    getDataLabelShapeProps(c::Chart, d::ChartDataLabel) -> Union{Nothing,DrawingShapeProps}

Fill and border of one label's box (`c:dLbl/c:spPr`). `nothing` means none was
written.
"""
getDataLabelShapeProps(c::Chart, d::ChartDataLabel) = parse_drawing_shape_props(_wb(c), _node(c, d))


"""
    getDataLabelText(c::Chart, d::ChartDataLabel) -> Union{Nothing,DrawingText}

Literal replacement text for one label (`c:dLbl/c:tx/c:rich`). `nothing` means
the label shows its value rather than typed-over text. Contrast
[`getDataLabelTextProps`](@ref), which is the label's `c:txPr` formatting.
"""
getDataLabelText(c::Chart, d::ChartDataLabel) =
    parse_drawing_text(_wb(c), first_element_with_tag(_node(c, d), "tx"); tag="rich")

"""
    getDataLabelPosition(c::Chart, d::ChartDataLabel) -> Union{Nothing,Symbol}

`c:dLblPos` — `:ctr` (centre), `:inEnd` (inside-end), `:inBase` (inside-base),
`:outEnd` (outside-end), `:l` (left), `:r` (right), `:t` (above), `:b` (below),
`:bestFit`. Which values are legal depends on the chart type.
"""
getDataLabelPosition(c::Chart, d::ChartDataLabel) = _sym_val(_node(c, d), "dLblPos")

"""
    getDataLabelOffset(c::Chart, d::ChartDataLabel) -> Union{Nothing,NamedTuple}

Manual position offset (`c:dLbl/c:layout/c:manualLayout`) as `(x, y)`,
fractions of the plot area, signed. `nothing` means the label sits where Excel
puts it. Distinct from [`getDataLabelPosition`](@ref), which is the discrete
`c:dLblPos` enum — Excel writes a layout offset when you drag a label and a
position when you pick one from the menu, so a label usually has one or neither.
"""
function getDataLabelOffset(c::Chart, d::ChartDataLabel)
    ml = first_element_with_tag(first_element_with_tag(_node(c, d), "layout"), "manualLayout")
    isnothing(ml) && return nothing
    return (x = _num_val(ml, "x"), y = _num_val(ml, "y"))
end

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

    #-- Cascade resolution -------------------------------------------------------

const _RUN_PROP_TYPES = Dict{Symbol,Type}(
    f => Base.nonnothingtype(t)
    for (f, t) in zip(fieldnames(DrawingRunProps), fieldtypes(DrawingRunProps))
    if f !== :raw
)
const _RUN_PROP_FIELDS = Tuple(f for f in fieldnames(DrawingRunProps) if haskey(_RUN_PROP_TYPES, f))

# Excel's names for run properties, accepted by the text setters alongside
# the DrawingML field names. Getters report the DrawingML names.
const _RUN_PROP_ALIASES = Dict(
    :font          => :latin,     # Excel's Font box sets the Latin typeface
    :color         => :fill,      # Font Color
    :underline     => :under,
    :strikethrough => :strike,
)

"""
    _walk_fill(wb, chain) -> Effective{DrawingFill}

Return the first fill written on any rung of `chain`, with the rung that
carried it. A rung whose `spPr` exists but holds no fill element is skipped —
a series with only a line inherits its fill from the next rung down.
"""
function _walk_fill(wb::Workbook, chain::Vector{FormatSite})::Effective{DrawingFill}
    for s in chain
        isnothing(s.props) && continue
        f = parse_drawing_fill(wb, s.props)
        isnothing(f) || return Effective{DrawingFill}(f, s, chain)
    end
    Effective{DrawingFill}(nothing, nothing, chain)
end

function _walk_line(wb::Workbook, chain::Vector{FormatSite})::Effective{DrawingLine}
    for s in chain
        isnothing(s.props) && continue
        l = parse_drawing_line(wb, s.props)
        isnothing(l) || return Effective{DrawingLine}(l, s, chain)
    end
    Effective{DrawingLine}(nothing, nothing, chain)
end

"""
    _shape_chain(c, i; point=nothing) -> Vector{FormatSite}

The `spPr` cascade for the graphic of series `i`: the data point, then the
series. Chart groups carry no `spPr`, and `plotArea`/`chartSpace` `spPr`
describe their own backgrounds rather than a series, so the chain stops there
and falls through to the style part, which this layer does not read.
"""
function _shape_chain(c::Chart, i::Integer; point::Union{Nothing,Integer}=nothing)
    chain = FormatSite[]
    if !isnothing(point)
        dp = getSeriesDataPoint(c, i, point)
        isnothing(dp) || push!(chain, _site_path(:point, :shape, dp.raw, "spPr"))
    end
    push!(chain, _site_path(:series, :shape, _series(c, i).raw, "spPr"))
    return chain
end

"""
    getSeriesFill(c::Chart, i::Integer; point=nothing) -> Effective{DrawingFill}

Resolve the fill of the graphic of series `i`, or of one data point of it.

This is always `c:spPr`. On a line or scatter series what Excel's format pane
calls the point's fill lives on its marker instead — use
[`getMarkerFill`](@ref) for that.
"""
getSeriesFill(c::Chart, i::Integer; point::Union{Nothing,Integer}=nothing) =
    _walk_fill(_wb(c), _shape_chain(c, i; point))

"""
    getSeriesLine(c::Chart, i::Integer; point=nothing) -> Effective{DrawingLine}

Resolve the outline of the graphic of series `i`, or of one data point of it.
"""
getSeriesLine(c::Chart, i::Integer; point::Union{Nothing,Integer}=nothing) =
    _walk_line(_wb(c), _shape_chain(c, i; point))

"""
    getMarkerFill(c::Chart, i::Integer, point::Integer) -> Effective{DrawingFill}

Resolve the fill of a data point's marker: `c:dPt/c:marker/c:spPr`, then
`c:ser/c:marker/c:spPr`.

Both lookups are anchored on `c:ser` and `c:dPt`, where `c:marker` is marker
properties. Under a group element the same tag is a boolean show-markers flag,
so this must not be generalised to walk up from a group.
"""
function getMarkerFill(c::Chart, i::Integer, point::Integer)::Effective{DrawingFill}
    chain = FormatSite[]
    dp = getSeriesDataPoint(c, i, point)
    isnothing(dp) || push!(chain, _site_path(:point, :marker, dp.raw, "marker", "spPr"))
    push!(chain, _site_path(:series, :marker, _series(c, i).raw, "marker", "spPr"))
    _walk_fill(_wb(c), chain)
end

# Resolve one run-property field over any chain of text sites. Schema-agnostic:
# the caller decides which sites form the cascade.
function _resolve_text_field(wb, chain, field::Symbol)
    field in _RUN_PROP_FIELDS || throw(XLSXError(
        "`$field` is not a resolvable text property. Valid fields: " *
        join(_RUN_PROP_FIELDS, ", ") * "."))
    for s in chain
        isnothing(s.props) && continue
        txt = parse_drawing_text(wb, s.props)
        isnothing(txt) && continue
        rp = default_run_props(txt)
        isnothing(rp) && continue          # mixed text: no single answer at this rung
        v = getfield(rp, field)
        isnothing(v) || return Effective{_RUN_PROP_TYPES[field]}(v, s, chain)
    end
    return Effective{_RUN_PROP_TYPES[field]}(nothing, nothing, chain)
end

"""
    getLabelTextProp(c::Chart, i::Integer, field::Symbol) -> Effective

Resolve one field of the data-label text formatting for series `i`, walking
`c:txPr` from the series up through the group, plot area and chart space.

`field` is a field of `DrawingRunProps` other than `raw`. Text inherits field by
field — a size from one rung and a weight from another both apply — which is
why this resolves a named field rather than a whole `DrawingRunProps`.

Two fields are composite: `:fill` and `:line` resolve to a whole `DrawingFill`
or `DrawingLine`. They inherit as units — a run's fill is written at some rung
or it is not — so there is no cascade below them, and resolving through to
`.fgcolor` or `.width` is a field access on the result, not a further walk.

A rung whose `txPr` exists but sets nothing (an empty `a:defRPr`, which Excel
writes at the chart space) does not answer: the walk continues, and a field set
at no rung resolves to `nothing`, meaning Excel takes it from the chart style.
"""
getLabelTextProp(c::Chart, i::Integer, field::Symbol) =
    _resolve_text_field(_wb(c), _text_chain(c, i), field)


"""
    _site_path(level, kind, container, path...) -> FormatSite

Build a rung whose properties live at `path` below `container`. `container` is
the nearest element that always exists — the `c:ser`, not its `c:dLbls`, which
Excel omits when a series has no labels. A setter creates whatever the path
lacks; `props === nothing` means it lacks something.
"""
function _site_path(level::Symbol, kind::Symbol, container::XML.Node, path::AbstractString...)
    n = container
    for tag in path
        n = first_element_with_tag(n, tag)
        isnothing(n) && return FormatSite(level, kind, container, nothing)
    end
    return FormatSite(level, kind, container, n)
end

function _text_chain(c::Chart, i::Integer)
    root = chart_root(c)
    chain = FormatSite[
        _site_path(:series, :text, _series(c, i).raw,      "dLbls", "txPr"),
        _site_path(:group,  :text, getSeriesGroup(c, i).raw, "dLbls", "txPr"),
    ]
    chrt = first_element_with_tag(root, "chart")
    pa = isnothing(chrt) ? nothing : first_element_with_tag(chrt, "plotArea")
    isnothing(pa) || push!(chain, _site_path(:plotarea, :text, pa, "txPr"))
    push!(chain, _site_path(:chartspace, :text, root, "txPr"))
    return chain
end

"""
    set_chart_root!(c::AbstractChart, newroot::XML.Node) -> nothing

Write `newroot` back as the chart part's root element. `rebuild_path` returns a
new `c:chartSpace` rather than mutating, and `get_xml_data` memoizes the parsed
document, so a rebuilt root that is not written back leaves the correct tree in
memory and the old one on disk — silently, since `writexlsx` serializes what is
in `xf.data`.

Splices into the existing document so the declaration and any other top-level
nodes survive.
"""
function set_chart_root!(c::AbstractChart, newroot::XML.Node)
    doc = get_xml_data(c.package, c.path)
    old = xml_root_element(doc)
    c.package.data[c.path] = replace_child(doc, old, newroot)
    return nothing
end

#-- Write path ---------------------------------------------------------------

"""
    _ln_with(ln, pfx; color, width, dash, cap, compound, join, miterLimit) -> XML.Node

Apply the given line properties to an `a:ln`, leaving unspecified ones alone.
Join is applied before the miter limit, which lives on the `a:miter` element.
"""
function _ln_with(ln::XML.Node, pfx::Dict{String,String};
                  color = nothing, width = nothing, dash = nothing,
                  cap = nothing, compound = nothing,
                  join = nothing, miterLimit = nothing)
    isnothing(color)      || (ln = _ln_with_color(ln, color, pfx))
    isnothing(width)      || (ln = _ln_with_width(ln, width))
    isnothing(dash)       || (ln = _ln_with_dash(ln, dash, pfx))
    isnothing(cap)        || (ln = _ln_with_cap(ln, cap))
    isnothing(compound)   || (ln = _ln_with_compound(ln, compound))
    isnothing(join)       || (ln = _ln_with_join(ln, join, pfx))
    isnothing(miterLimit) || (ln = _ln_with_miter_limit(ln, miterLimit))
    return ln
end

function _series_path(c::Chart, root::XML.Node, i::Integer)
    grp, ser = _series_pick(c, root, i)
    gtag = String(localname(grp))
    skey = (NS_C, SER_TYPE[gtag])
    steps = [(NS_C, "chart")    => "chart",
             (NS_C, "plotArea") => "plotArea",
             (NS_C, gtag)       => (gtag, n -> n === grp),
             skey               => ("ser", n -> n === ser)]
    return steps, skey, ser
end

function _ser_idx(ser::XML.Node)
    v = _int_val(ser, "idx")
    isnothing(v) && throw(XLSXError("A `c:ser` has no `c:idx`, which the schema requires."))
    return v
end

function _series_by_idx(c::Chart, root::XML.Node, idx::Integer)
    for (g, s) in _series_nodes(root)
        _int_val(s, "idx") == idx && return (g, s)
    end
    throw(XLSXError("Chart `$(c.name)` has no series with idx $idx."))
end

# The `tag` child of `parent` whose c:idx is `idx`, as written (0-based).
function _idx_child(parent::Union{Nothing,XML.Node}, tag::AbstractString, idx::Integer)
    isnothing(parent) && return nothing
    for el in elements_with_tag(parent, tag)
        _int_val(el, "idx") == idx && return el
    end
    return nothing
end

# As _idx_child, for a 1-based user-facing point position.
function _point_node(parent::Union{Nothing,XML.Node}, tag::AbstractString, point::Integer)
    point >= 1 || throw(XLSXError("Data point positions start at 1; asked for $point."))
    return _idx_child(parent, tag, point - 1)
end

_dlbl_node(ser::XML.Node, point::Integer) = _point_node(first_element_with_tag(ser, "dLbls"), "dLbl", point)

_group_axids(n::XML.Node) =
    Int[parse(Int, v) for a in elements_with_tag(n, "axId") for v in (_attr(a, "val"),) if !isnothing(v)]

function _group_node(c::Chart, root::XML.Node, g::ChartGroup)
    tag  = String(g.kind)
    hits = [n for n in _group_nodes(root) if localname(n) == tag && _group_axids(n) == g.axids]
    isempty(hits) && throw(XLSXError(
        "Chart `$(c.name)` has no $(g.kind) group plotting against axes $(g.axids)."))
    length(hits) > 1 && throw(XLSXError(
        "Chart `$(c.name)` has $(length(hits)) $(g.kind) groups plotting against axes " *
        "$(g.axids), so the group cannot be identified."))
    return only(hits)
end

function _nth(parent::XML.Node, tag::AbstractString, k::Integer, what::AbstractString)
    els = collect(elements_with_tag(parent, tag))
    1 <= k <= length(els) || throw(XLSXError(
        "The series has $(length(els)) $what; asked for number $k."))
    return els[k]
end

# The element a sub-object refers to, found by its key in `root`. Getters use the
# two-argument form; setters pass the root they will rebuild.
_node(c::Chart, x) = _node(c, chart_root(c), x)

_node(c::Chart, root::XML.Node, g::ChartGroup) = _group_node(c, root, g)

function _node(c::Chart, root::XML.Node, b::ChartUpDownBars)
    n = first_element_with_tag(_group_node(c, root, b.group), "upDownBars")
    isnothing(n) && throw(XLSXError("The $(b.group.kind) group no longer has up-down bars."))
    return n
end

function _node(c::Chart, root::XML.Node, d::ChartDataPoint)
    _, ser = _series_by_idx(c, root, d.series_idx)
    n = _idx_child(ser, "dPt", d.idx)
    isnothing(n) && throw(XLSXError(
        "Series idx $(d.series_idx) no longer has formatting for point idx $(d.idx)."))
    return n
end

function _node(c::Chart, root::XML.Node, d::ChartDataLabel)
    _, ser = _series_by_idx(c, root, d.series_idx)
    n = _idx_child(first_element_with_tag(ser, "dLbls"), "dLbl", d.idx)
    isnothing(n) && throw(XLSXError(
        "Series idx $(d.series_idx) no longer has an individual label for point idx $(d.idx)."))
    return n
end

_node(c::Chart, root::XML.Node, t::ChartTrendline) =
    _nth(last(_series_by_idx(c, root, t.series_idx)), "trendline", t.ordinal, "trendlines")

_node(c::Chart, root::XML.Node, e::ChartErrorBars) =
    _nth(last(_series_by_idx(c, root, e.series_idx)), "errBars", e.ordinal, "error bar sets")

function _node(c::Chart, root::XML.Node, m::ChartMarker)
    _, ser = _series_by_idx(c, root, m.series_idx)
    parent = isnothing(m.point_idx) ? ser : _idx_child(ser, "dPt", m.point_idx)
    n = first_element_with_tag(parent, "marker")
    isnothing(n) && throw(XLSXError("That marker is no longer in the chart."))
    return n
end

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

Set the fill of series `i`'s graphic. `color` may be a color string or Symbol,
a `Colors.Colorant`, a [`SchemeColor`](@ref), `:none` for an explicit
`<a:noFill/>`, or `:inherit` to remove it so the chart style applies.

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

Set the color of series `i`'s outline. Takes the same values as
[`setSeriesFill`](@ref): a color, a [`SchemeColor`](@ref), `:none` for an
explicit `<a:noFill/>`, or `:inherit` to remove it.

Note `:none` here leaves the `a:ln` in place with no fill — the line exists and
draws nothing, which is what Excel writes and is distinct from
`setSeriesLine(c, i, :inherit)`, which removes the outline entirely.
"""
setSeriesLineColor(c::Chart, i::Integer, color) =
    _set_series_line(c, i, (ln, pfx) -> _ln_with_color(ln, color, pfx))

"""
    setSeriesLineWidth(c, i, points) -> Chart

Set the width of series `i`'s outline, in points — the unit Excel's width box
uses. `:inherit` removes the width so the chart style supplies it.

DrawingML caps a line width at 1584 pt; a wider one is rejected rather than
written.

One of the single-property line setters; [`setSeriesLine`](@ref) sets several at
once in a single rebuild.
"""
setSeriesLineWidth(c::Chart, i::Integer, points) =
    _set_series_line(c, i, (ln, _) -> _ln_with_width(ln, points))


"""
    setSeriesLineDash(c, i, dash) -> Chart

Set the dash pattern of series `i`'s outline: `"solid"`, `"dash"`, `"sysDot"`
and the other DrawingML presets, or Excel's own names for them. `:inherit`
removes the pattern.

One of the single-property line setters; see [`setSeriesLine`](@ref).
"""
setSeriesLineDash(c::Chart, i::Integer, dash) =
    _set_series_line(c, i, (ln, pfx) -> _ln_with_dash(ln, dash, pfx))

"""
    setSeriesLineCap(c, i, cap) -> Chart

Set how series `i`'s outline ends: `"rnd"`, `"sq"` or `"flat"`, or Excel's names
for them. `:inherit` removes the setting.

One of the single-property line setters; see [`setSeriesLine`](@ref).
"""
setSeriesLineCap(c::Chart, i::Integer, cap) =
    _set_series_line(c, i, (ln, _) -> _ln_with_cap(ln, cap))

"""
    setSeriesLineCompound(c, i, cmpd) -> Chart

Set the compound line type of series `i`'s outline — how many parallel strokes
it draws: `"sng"`, `"dbl"`, `"thickThin"`, `"thinThick"` or `"tri"`, or Excel's
names for them. `:inherit` removes the setting.

One of the single-property line setters; see [`setSeriesLine`](@ref).
"""
setSeriesLineCompound(c::Chart, i::Integer, cmpd) =
    _set_series_line(c, i, (ln, _) -> _ln_with_compound(ln, cmpd))

"""
    setSeriesLineJoin(c, i, join) -> Chart

Set how the segments of series `i`'s outline meet: `:round`, `:bevel` or
`:miter`. `:inherit` removes the join. Excel uses the same three words, so there
are no aliases.

A miter join takes a limit as well, set separately with
[`setSeriesLineMiterLimit`](@ref); changing the join afterwards discards it.

One of the single-property line setters; see [`setSeriesLine`](@ref).
"""
setSeriesLineJoin(c::Chart, i::Integer, join) =
    _set_series_line(c, i, (ln, pfx) -> _ln_with_join(ln, join, pfx))

"""
    setSeriesLineMiterLimit(c, i, limit) -> Chart

Set the miter limit of series `i`'s outline, as a multiple of the line width:
how far a sharp corner may extend before the join is cut off. Excel's default is
8. `:inherit` removes the limit.

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

Set several line properties at once, in one rebuild of the chart part. A keyword
left unspecified is left alone; pass `:inherit` to remove one that is set.

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
    marker_in(parent, key) = rebuild_path(parent, [(NS_C, "marker") => "marker"],
                                          mk -> f(mk, pfx); prefixes = pfx, parent_key = key)
    new = rebuild_path(root, steps,
              ser -> isnothing(point) ? marker_in(ser, ser_key) :
                     _with_dpt(ser, ser_key, point, pfx, dp -> marker_in(dp, (NS_C, "dPt")));
              prefixes = pfx)
    set_chart_root!(c, new)
    return c
end

"""
    setMarkerSymbol(c, i, symbol; point = nothing) -> Chart
 
Set the marker shape for series `i`, or for one of its data points when `point`
is given, counting from 1. One of `:circle`, `:dash`, `:diamond`, `:dot`,
`:none`, `:picture`, `:plus`, `:square`, `:star`, `:triangle`, `:x` or `:auto`.
 
`:none` draws no marker and `:auto` lets Excel choose; both are settings.
`:inherit` removes the element so the chart style decides.
"""
setMarkerSymbol(c::Chart, i::Integer, symbol; point::Union{Nothing,Integer} = nothing) =
    _set_marker(c, i, point, (mk, pfx) -> _marker_with_symbol(mk, symbol, pfx))
 
"""
    setMarkerSize(c, i, size; point = nothing) -> Chart
 
Set the marker size for series `i`, or for one of its data points when `point`
is given, counting from 1. In points, rounded to the nearest whole point.
DrawingML allows 2 to 72; anything outside that is rejected rather than written.
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
given, counting from 1. Takes the same values as [`setSeriesFill`](@ref): a
colour, a [`SchemeColor`](@ref), `:none` for an explicit `<a:noFill/>`, or
`:inherit` to remove it.
 
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
 
Set the outline colour of series `i`'s marker, or of one data point's when
`point` is given, counting from 1. Takes the same values as
[`setSeriesFill`](@ref): a colour, a [`SchemeColor`](@ref), `:none` for an
explicit `<a:noFill/>`, or `:inherit` to remove it.
 
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
`point` is given, counting from 1. In points. `:inherit` removes the width.
 
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
[`setMarkerSymbol`](@ref), [`setMarkerSize`](@ref), [`setMarkerFill`](@ref),
[`setMarkerLineColor`](@ref) and [`setMarkerLineWidth`](@ref).
 
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

"""
    _txpr_with_run_prop(el, key, field, value, pfx, ns) -> XML.Node

Set one run property on an element's `c:txPr`, creating it if absent. `key` is
`el`'s [`SchemaKey`](@ref).

`ns` is the namespace a created txPr takes, NS_C by default and NS_CX for a chartEx part.
"""
function _txpr_with_run_prop(el::XML.Node, key::SchemaKey, field::Symbol, value,
                             pfx::Dict{String,String}; ns::AbstractString = NS_C)
    txpr = something(first_element_with_tag(el, "txPr"), _new_text_body("txPr", pfx, ns))
    return insert_child(el, key, _text_with_run_prop(txpr, field, value, pfx))
end


"""
    _both_with_run_prop(el, key, field, value, pfx, ns) -> XML.Node

Set one run property on an element's `c:txPr` and, where it holds literal text,
on its `c:tx/c:rich` as well.

Excel renders the `rich` body, so writing only `c:txPr` leaves the edit
invisible. This is the shape data labels, chart titles and axis titles share.
`c:rich` is updated only where it already exists — one with no text means
nothing, and typing text is a separate operation.

`ns` is the namespace a created txPr takes, NS_C by default and NS_CX for a chartEx part.
"""
function _both_with_run_prop(el::XML.Node, key::SchemaKey, field::Symbol, value,
                             pfx::Dict{String,String}; ns::AbstractString = NS_C)
    el = _txpr_with_run_prop(el, key, field, value, pfx; ns)

    tx = first_element_with_tag(el, "tx")
    isnothing(tx) && return el
    rich = first_element_with_tag(tx, "rich")
    isnothing(rich) && return el          # a c:strRef title has no literal text

    return replace_child(el, tx,
               replace_child(tx, rich, _text_with_run_prop(rich, field, value, pfx)))
end

"""
    setLabelTextProp(c, i, field, value) -> Chart
    setLabelTextProp(c, i, point, field, value) -> Chart

Set one text property of series `i`'s data labels, or of one individual label.
`field` is a field of `DrawingRunProps` other than `raw` — the same vocabulary
[`getLabelTextProp`](@ref) resolves. `:inherit` removes it.

Creating a `c:dLbls` where none existed writes every `show*` flag as off, so
formatting a series' labels does not make them appear. A `c:dLbls` that names no
flags shows labels with Excel's own defaults — legend key, series name and value
— which would turn them on for a series that had none.

The per-point form creates the individual `c:dLbl` if the point has none, placed
in index order and carrying the same `show*` flags as its `c:dLbls`, so it
displays exactly what its sibling labels display. Formatting one label on a
series whose labels are off does not make that label appear.

Two fields take compound values. `:fill` accepts anything
[`setSeriesFill`](@ref) does — a color, a [`SchemeColor`](@ref), `:none` or
`:inherit`. `:line` accepts a color, or a `NamedTuple` of the
[`setSeriesLine`](@ref) keywords: `(color = "red", width = 1.5)`.

A label that has been deleted carries only `c:delete` and cannot be formatted —
undeleting it is a separate operation.
"""
function setLabelTextProp(c::Chart, i::Integer, field::Symbol, value)
    root = chart_root(c)
    pfx  = ns_prefixes(root)
    steps, ser_key = _series_path(c, root, i)
    new = rebuild_path(root, [steps...; (NS_C, "dLbls") => "dLbls"],
                       lbl -> _label_transform(lbl, field, value, pfx);
                       prefixes = pfx, parent_key = ser_key)
    set_chart_root!(c, new)
    return c
end

function setLabelTextProp(c::Chart, i::Integer, point::Integer, field::Symbol, value)
    root = chart_root(c)
    pfx  = ns_prefixes(root)
    steps, _, ser = _series_path(c, root, i)
    dl = _dlbl_node(ser, point)
    !isnothing(dl) && _bool_val(dl, "delete") === true && throw(XLSXError(
        "Label $point of series $i is deleted, so it has no formatting to set. " *
        "Undelete it first with `setLabelDeleted(c, $i, $point, false)`."))
    new = rebuild_path(root, [steps...; (NS_C, "dLbls") => "dLbls"],
              lbls -> _with_dlbl(_dlbls_flags_off(lbls, (NS_C, "dLbls"), pfx), point, pfx,
                                 lbl -> _label_transform(lbl, field, value, pfx));
              prefixes = pfx)
    set_chart_root!(c, new)
    return c
end

"""
    setGroupLabelTextProp(c, g::ChartGroup, field, value) -> Chart

Set one text property of a chart group's data labels — the rung every series in
the group inherits from.
"""
function setGroupLabelTextProp(c::Chart, g::ChartGroup, field::Symbol, value)
    root = chart_root(c)
    pfx  = ns_prefixes(root)
    gn   = _group_node(c, root, g)
    gtag = String(localname(gn))
    new  = rebuild_path(root,
               [(NS_C, "chart")    => "chart",
                (NS_C, "plotArea") => "plotArea",
                (NS_C, gtag)       => (gtag, n -> n === gn),
                (NS_C, "dLbls")    => "dLbls"],
               lbl -> _both_with_run_prop(lbl, (NS_C, String(localname(lbl))),
                                          field, value, pfx);
               prefixes = pfx)
    set_chart_root!(c, new)
    return c
end

"""
    setChartSpaceTextProp(c, field, value) -> Chart

Set one text property on `c:chartSpace/c:txPr` — the bottom rung of the text
cascade, which every text body in the chart inherits from unless it overrides.
"""
function setChartSpaceTextProp(c::Chart, field::Symbol, value)
    root = chart_root(c)
    pfx  = ns_prefixes(root)
    txpr = something(first_element_with_tag(root, "txPr"), _new_text_body("txPr", pfx))
    new  = insert_child(root, (NS_C, "chartSpace"),
                        _text_with_run_prop(txpr, field, value, pfx))
    set_chart_root!(c, new)
    return c
end

"""
    setChartTitleText(c, text) -> Chart

Replace the chart title's text and all its formatting. `text` may be a
`DrawingText` or a string, which becomes a one-run title inheriting its
formatting from the chart style.

This replaces rather than merges, like writing a `RichTextString` to a cell —
use [`setChartTitleTextProp`](@ref) to change one property and leave the rest.

A title bound to a cell (`c:tx/c:strRef`) is replaced by literal text; the
reference is lost. `c:autoTitleDeleted` is not touched, so a chart with the
title switched off stays that way.

A title created here gets `c:overlay val="0"`, as Excel writes on every title
it creates, so it sits beside the plot area rather than over it.
"""
function setChartTitleText(c::Chart, text::DrawingText)
    root = chart_root(c)
    pfx  = ns_prefixes(root)
    new  = rebuild_path(root,
               [(NS_C, "chart") => "chart",
                (NS_C, "title") => "title"],
               t -> _title_with_text(t, text, pfx);
               prefixes = pfx)
    set_chart_root!(c, new)
    return c
end

setChartTitleText(c::Chart, text::AbstractString) =
    setChartTitleText(c, DrawingText(text))

    # A title with no c:overlay is drawn over the plot area; Excel writes
# overlay=0 on every title it creates, so a new title gets the same.
function _title_with_text(t::XML.Node, text::DrawingText, pfx)
    tkey = (NS_C, "title")
    tx = something(first_element_with_tag(t, "tx"), XML.Element(prefixed_tag(pfx[NS_C], "tx")))
    tx = insert_child(remove_choice(tx, (NS_C, "tx"), TX_GROUP), (NS_C, "tx"),
                      _text_from(text, "rich", pfx))
    t = insert_child(t, tkey, tx)
    return isnothing(first_element_with_tag(t, "overlay")) ?
           insert_child(t, tkey, _val_el(pfx, "overlay", "0")) : t
end

"""
    setAxisTitleText(c, ax::ChartAxis, text) -> Chart

Replace an axis title's text and all its formatting. As
[`setChartTitleText`](@ref), including that a cell reference is replaced by
literal text.

A title created here gets `c:overlay val="0"`, as Excel writes on every title
it creates, so it sits beside the plot area rather than over it.
"""
function setAxisTitleText(c::Chart, ax::ChartAxis, text::DrawingText)
    root = chart_root(c)
    pfx  = ns_prefixes(root)
    steps, _, _ = _axis_path(c, root, ax)
    new = rebuild_path(root, [steps...; (NS_C, "title") => "title"],
                       t -> _title_with_text(t, text, pfx); prefixes = pfx)
    set_chart_root!(c, new)
    return c
end

setAxisTitleText(c::Chart, ax::ChartAxis, text::AbstractString) =
    setAxisTitleText(c, ax, DrawingText(text))

"""
    setLabelText(c, i, point, text) -> Chart

Replace one data label's typed-over text and all its formatting, so the label
shows `text` instead of its value.

If the point has no `c:dLbl`, one is created with the same `show*` flags as its
`c:dLbls`. Typed text is displayed whatever those flags say, so this shows the
label even on a series whose labels are off.
"""
function setLabelText(c::Chart, i::Integer, point::Integer, text::DrawingText)
    root = chart_root(c)
    pfx  = ns_prefixes(root)
    steps, _, ser = _series_path(c, root, i)
    dl = _dlbl_node(ser, point)
    !isnothing(dl) && _bool_val(dl, "delete") === true && throw(XLSXError(
        "Label $point of series $i is deleted, so it has no text to set." *
        "Undelete it first with `setLabelDeleted(c, $i, $point, false)`."))
    settx(lbl) = rebuild_path(lbl, [(NS_C, "tx") => "tx"],
                     tx -> insert_child(remove_choice(tx, (NS_C, "tx"), TX_GROUP),
                                        (NS_C, "tx"), _text_from(text, "rich", pfx));
                     prefixes = pfx, parent_key = (NS_C, "dLbl"))
    new = rebuild_path(root, [steps...; (NS_C, "dLbls") => "dLbls"],
              lbls -> _with_dlbl(_dlbls_flags_off(lbls, (NS_C, "dLbls"), pfx), point, pfx, settx);
              prefixes = pfx)
    set_chart_root!(c, new)
    return c
end

setLabelText(c::Chart, i::Integer, point::Integer, text::AbstractString) =
    setLabelText(c, i, point, DrawingText(text))

"""
    setChartTitleTextProp(c, field, value) -> Chart

Set one text property of the chart title. `field` is a field of
`DrawingRunProps` other than `raw`; `:inherit` removes it.

A title carries formatting in `c:title/c:txPr` and, when it holds literal text,
in `c:title/c:tx/c:rich` as well — Excel renders the `rich` one, so both are
written. Contrast [`setChartTitleText`](@ref), which replaces the text and all
its formatting.


"""
function setChartTitleTextProp(c::Chart, field::Symbol, value)
    root = chart_root(c)
    pfx  = ns_prefixes(root)
    new  = rebuild_path(root,
               [(NS_C, "chart") => "chart",
                (NS_C, "title") => "title"],
               t -> _both_with_run_prop(t, (NS_C, "title"), field, value, pfx);
               prefixes = pfx)
    set_chart_root!(c, new)
    return c
end

"""
    setAxisTitleTextProp(c, ax::ChartAxis, field, value) -> Chart

Set one text property of an axis title. As [`setChartTitleTextProp`](@ref).
"""
function setAxisTitleTextProp(c::Chart, ax::ChartAxis, field::Symbol, value)
    root = chart_root(c)
    pfx  = ns_prefixes(root)
    axn  = _axis_node(c, root, ax)
    atag = String(localname(axn))
    new  = rebuild_path(root,
               [(NS_C, "chart")    => "chart",
                (NS_C, "plotArea") => "plotArea",
                (NS_C, atag)       => (atag, n -> n === axn),
                (NS_C, "title")    => "title"],
               t -> _both_with_run_prop(t, (NS_C, "title"), field, value, pfx);
               prefixes = pfx)
    set_chart_root!(c, new)
    return c
end

"""
    setLegendTextProp(c, field, value) -> Chart

Set one text property of the legend (`c:legend/c:txPr`). A legend takes its text
from the series names, so there is no `c:tx` to replace — only its formatting is
settable.
"""
function setLegendTextProp(c::Chart, field::Symbol, value)
    root = chart_root(c)
    pfx  = ns_prefixes(root)
    new  = rebuild_path(root,
               [(NS_C, "chart")  => "chart",
                (NS_C, "legend") => "legend"],
               lg -> _txpr_with_run_prop(lg, (NS_C, "legend"), field, value, pfx);
               prefixes = pfx)
    set_chart_root!(c, new)
    return c
end

"""
    _dlbls_flags_off(lbl, key, pfx) -> XML.Node

Write every `show*` flag as `val="0"` on a `c:dLbls` that names none.

A `c:dLbls` with no flags shows labels with Excel's own defaults — legend key,
series name and value — so creating one to hold formatting would turn labels on
for a series that had none. Excel writes all six whenever it creates a
`c:dLbls`; matching that keeps a formatting call from changing what is
displayed. A `c:dLbls` that already names any flag is left alone.
"""
function _dlbls_flags_off(lbl::XML.Node, key::SchemaKey, pfx::Dict{String,String})
    any(f -> !isnothing(first_element_with_tag(lbl, f)), DLBLS_FLAGS) && return lbl
    for f in DLBLS_FLAGS
        lbl = insert_child(lbl, key, XML.Element(prefixed_tag(pfx[NS_C], f); val = "0"))
    end
    return lbl
end

function _label_transform(lbl, field, value, pfx)
    key = (NS_C, String(localname(lbl)))
    localname(lbl) == "dLbls" && (lbl = _dlbls_flags_off(lbl, key, pfx))
    return _both_with_run_prop(lbl, key, field, value, pfx)
end

# <c:TAG><c:idx val="idx"/></c:TAG>: the minimum a c:dPt or c:dLbl needs.
_new_idx_element(tag::AbstractString, idx::Integer, pfx) =
    XML.Element(prefixed_tag(pfx[NS_C], tag),
                XML.Element(prefixed_tag(pfx[NS_C], "idx"); val = string(idx)))

_c_idx(n::XML.Node) = _int_val(n, "idx")

function _with_dpt(ser::XML.Node, ser_key::SchemaKey, point::Integer, pfx, f)
    point >= 1 || throw(XLSXError("Data point positions start at 1; asked for $point."))
    return _with_indexed_child(ser, ser_key, "dPt", point - 1, _c_idx,
                               () -> _new_idx_element("dPt", point - 1, pfx), f)
end

# A created c:dLbl copies the show* flags of its c:dLbls, so it displays what its
# siblings display. A dLbl naming no flags would fall back to Excel's defaults and
# turn the label on, the same trap _dlbls_flags_off avoids at the dLbls level.
function _with_dlbl(lbls::XML.Node, point::Integer, pfx, f)
    point >= 1 || throw(XLSXError("Data point positions start at 1; asked for $point."))
    make() = _copy_label_flags(_new_idx_element("dLbl", point - 1, pfx), lbls, pfx)
    return _with_indexed_child(lbls, (NS_C, "dLbls"), "dLbl", point - 1, _c_idx, make, f)
end

function _copy_label_flags(lbl::XML.Node, from::XML.Node, pfx)
    for f in DLBLS_FLAGS
        src = first_element_with_tag(from, f)
        isnothing(src) && continue
        lbl = insert_child(lbl, (NS_C, "dLbl"),
                           XML.Element(prefixed_tag(pfx[NS_C], f); val = _attr(src, "val")))
    end
    return lbl
end

"""
    setLabelDeleted(c, i, point, deleted::Bool) -> Chart

Delete or undelete the data label of point `point` (counting from 1) of series
`i`.

Deleting writes a `c:dLbl` carrying only `c:idx` and `c:delete`, replacing any
individual formatting or typed text the label had: the schema makes `c:delete`
exclusive of every display property. Undeleting removes that `c:dLbl`, so the
point shows whatever its series' `c:dLbls` does. Undeleting a label that is not
deleted does nothing.
"""
function setLabelDeleted(c::Chart, i::Integer, point::Integer, deleted::Bool)
    root = chart_root(c)
    pfx  = ns_prefixes(root)
    steps, _, ser = _series_path(c, root, i)
    lbl_steps = [steps...; (NS_C, "dLbls") => "dLbls"]

    if !deleted
        dl = _dlbl_node(ser, point)
        (isnothing(dl) || _bool_val(dl, "delete") !== true) && return c
        new = rebuild_path(root, lbl_steps,
                  lbls -> _with_children(lbls, filter(k -> k !== dl, lbls.children));
                  prefixes = pfx)
    else
        new = rebuild_path(root, lbl_steps,
                  lbls -> _with_dlbl(_dlbls_flags_off(lbls, (NS_C, "dLbls"), pfx), point, pfx,
                                     lbl -> _deleted_dlbl(lbl, pfx));
                  prefixes = pfx)
    end
    set_chart_root!(c, new)
    return c
end

# <c:dLbl><c:idx/><c:delete val="1"/></c:dLbl>, keeping the label's own c:idx.
_deleted_dlbl(lbl::XML.Node, pfx) =
    _with_children(lbl, XML.Node[first_element_with_tag(lbl, "idx"),
                                 XML.Element(prefixed_tag(pfx[NS_C], "delete"); val = "1")])

# ── Axis setters ──────────────────────────────────────────────────────────────

function _axis_path(c::Chart, root::XML.Node, ax::ChartAxis)
    el  = _axis_node(c, root, ax)
    tag = String(localname(el))
    key = (NS_C, tag)
    steps = [(NS_C, "chart")    => "chart",
             (NS_C, "plotArea") => "plotArea",
             key                => (tag, n -> n === el)]
    return steps, key, el
end

"""
    _set_axis(c, ax, f) -> Chart

Apply `f(el, key, pfx)` to the axis element and rebuild the chart part once.
`key` is the axis's schema key, for `insert_child` on the axis.
"""
function _set_axis(c::Chart, ax::ChartAxis, f)
    root = chart_root(c)
    pfx  = ns_prefixes(root)
    steps, key, _ = _axis_path(c, root, ax)
    new = rebuild_path(root, steps, el -> f(el, key, pfx); prefixes = pfx)
    set_chart_root!(c, new)
    return c
end

_val_el(pfx, tag, v::AbstractString) = XML.Element(prefixed_tag(pfx[NS_C], tag); val = v)

# Set `<c:tag val=v/>` on `el`, or remove it for `:inherit`.
_with_val(el, key, pfx, tag, v) =
    v === :inherit ? remove_child(el, tag) : insert_child(el, key, _val_el(pfx, tag, v))

function _check_sym(v, allowed, what)
    v === :inherit && return :inherit
    v isa Symbol && v in allowed && return String(v)
    throw(XLSXError("`$(repr(v))` is not a valid $what; use one of " *
                    join(map(repr, allowed), ", ") * ", or `:inherit`."))
end

# xsd:double as Excel writes it: integers without a trailing `.0`.
_xsd_double(x::Float64) = isinteger(x) && abs(x) < 1e15 ? string(Int(x)) : string(x)

function _finite(v, what)
    v === :inherit && return :inherit
    x = Float64(v)
    isfinite(x) || throw(XLSXError("$what must be finite; got $v."))
    return _xsd_double(x)
end

const _AX_TICK_LBL_POS = (:high, :low, :nextTo, :none)
const _AX_ORIENTATIONS = (:minMax, :maxMin)
const _AX_CROSSES      = (:autoZero, :min, :max)

"""
    setAxisTickLabelPos(c, ax, pos) -> Chart

Set where the axis's tick labels are drawn (`c:tickLblPos`): `:nextTo` the
axis, at the `:low` or `:high` end of the crossing axis, or `:none`. `:inherit`
removes the element.

`:low` keeps a category axis's labels at the edge of the plot area when the
value axis has negative values, rather than beside the zero line.
"""
setAxisTickLabelPos(c::Chart, ax::ChartAxis, pos::Symbol) =
    _set_axis(c, ax, (el, key, pfx) ->
        _with_val(el, key, pfx, "tickLblPos", _check_sym(pos, _AX_TICK_LBL_POS, "tick label position")))

"""
    setAxisScaling(c, ax; orientation, min, max) -> Chart

Set several `c:scaling` properties in one rebuild. A keyword left unspecified is
left alone; `:inherit` removes it.

- `orientation` — `:minMax`, or `:maxMin` to plot the axis in reverse order.
- `min`, `max` — fixed bounds, in the axis's own units. `:inherit` returns a
  bound to automatic.

Throws if the resulting bounds would have `min ≥ max`. Checking once all
keywords are applied means a range can be moved in one call past its old
bounds, which separate [`setAxisMin`](@ref) and [`setAxisMax`](@ref) calls
might not allow.

Throws if the axis has no `c:scaling`; see [`getAxisOrientation`](@ref).
"""
function setAxisScaling(c::Chart, ax::ChartAxis;
                        orientation = nothing, min = nothing, max = nothing)
    all(isnothing, (orientation, min, max)) && return c
    o  = isnothing(orientation) ? nothing : _check_sym(orientation, _AX_ORIENTATIONS, "axis orientation")
    lo = isnothing(min) ? nothing : _finite(min, "An axis minimum")
    hi = isnothing(max) ? nothing : _finite(max, "An axis maximum")
    skey = (NS_C, "scaling")
    return _set_axis(c, ax, (el, key, pfx) -> begin
        sc = _axis_scaling(el)
        isnothing(o)  || (sc = _with_val(sc, skey, pfx, "orientation", o))
        isnothing(lo) || (sc = _with_val(sc, skey, pfx, "min", lo))
        isnothing(hi) || (sc = _with_val(sc, skey, pfx, "max", hi))
        a, b = _scaling_num(sc, "min"), _scaling_num(sc, "max")
        isnothing(a) || isnothing(b) || a < b || throw(XLSXError(
            "Axis minimum ($a) must be less than its maximum ($b)."))
        insert_child(el, key, sc)
    end)
end

"""
    setAxisMin(c, ax, value) -> Chart

Fix the axis's lower bound (`c:scaling/c:min`); `:inherit` returns it to
automatic. See [`setAxisScaling`](@ref) to set both bounds at once.
"""
setAxisMin(c::Chart, ax::ChartAxis, v) = setAxisScaling(c, ax; min = v)

"""
    setAxisMax(c, ax, value) -> Chart

Fix the axis's upper bound (`c:scaling/c:max`); `:inherit` returns it to
automatic. See [`setAxisScaling`](@ref) to set both bounds at once.
"""
setAxisMax(c::Chart, ax::ChartAxis, v) = setAxisScaling(c, ax; max = v)

"""
    setAxisOrientation(c, ax, orientation) -> Chart

`:minMax` for a normal axis, `:maxMin` to plot it in reverse order — Excel's
"Categories in reverse order" or "Values in reverse order". `:inherit` removes
the element.

Reversing a bar chart's category axis moves the value axis to the other end;
[`setAxisCrosses`](@ref)`(c, valax, :max)` moves it back.
"""
setAxisOrientation(c::Chart, ax::ChartAxis, o::Symbol) = setAxisScaling(c, ax; orientation = o)

"""
    setAxisCrosses(c, ax, rule) -> Chart

Set where this axis crosses its partner (`c:crosses`), as a position on the
partner: `:autoZero`, `:min` or `:max`. Removes any `c:crossesAt`, since the
schema allows one or the other. `:inherit` removes `c:crosses` only.

Excel's "Horizontal axis crosses: At maximum category", set in the category
axis's pane, is written on the *value* axis: `setAxisCrosses(c, valax, :max)`.
"""
function setAxisCrosses(c::Chart, ax::ChartAxis, rule::Symbol)
    v = _check_sym(rule, _AX_CROSSES, "crossing rule")
    return _set_axis(c, ax, (el, key, pfx) ->
        v === :inherit ? remove_child(el, "crosses") :
            _with_val(remove_child(el, "crossesAt"), key, pfx, "crosses", v))
end

"""
    setAxisCrossesAt(c, ax, value) -> Chart

Set where this axis crosses its partner as a value on the partner's scale
(`c:crossesAt`). Removes any `c:crosses`. `:inherit` removes `c:crossesAt` only.
"""
function setAxisCrossesAt(c::Chart, ax::ChartAxis, value)
    v = _finite(value, "An axis crossing")
    return _set_axis(c, ax, (el, key, pfx) ->
        v === :inherit ? remove_child(el, "crossesAt") :
            _with_val(remove_child(el, "crosses"), key, pfx, "crossesAt", v))
end

"""
    setAxisNumberFormatCode(c, ax, code) -> Chart

Set the axis's number format (`c:numFmt/@formatCode`) and unlink it from the
source cells (`sourceLinked="0"`), without which Excel ignores the code.
`:inherit` removes `c:numFmt`.

A two-section code such as `"0%;0%"` shows negative values without a minus sign.
"""
function setAxisNumberFormatCode(c::Chart, ax::ChartAxis, code)
    return _set_axis(c, ax, (el, key, pfx) ->
        code === :inherit ? remove_child(el, "numFmt") :
            insert_child(el, key, XML.Element(prefixed_tag(pfx[NS_C], "numFmt");
                                              formatCode = String(code), sourceLinked = "0")))
end

"""
    setAxisNumberFormatLinked(c, ax, linked::Bool) -> Chart

Set `c:numFmt/@sourceLinked`, keeping the current format code, or `"General"` if
none was written.
"""
function setAxisNumberFormatLinked(c::Chart, ax::ChartAxis, linked::Bool)
    return _set_axis(c, ax, (el, key, pfx) -> begin
        code = something(_attr(first_element_with_tag(el, "numFmt"), "formatCode"), "General")
        insert_child(el, key, XML.Element(prefixed_tag(pfx[NS_C], "numFmt");
                                          formatCode = code, sourceLinked = linked ? "1" : "0"))
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

function _check_pct(v, lo, hi, what)
    v === :inherit && return :inherit
    v isa Integer && lo <= v <= hi && return string(v)
    throw(XLSXError("$what must be an integer percentage from $lo to $hi, or `:inherit`; got $(repr(v))."))
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

Set the group's gap width (0–500, as Excel's "Gap Width" box), or `:inherit`
to remove it.
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

Set the group's overlap (-100 to 100), or `:inherit` to remove it. Stacked
groups need 100 for each series to sit on the one before it.
"""
function setGroupOverlap(c::Chart, g::ChartGroup, pct)
    v = _check_pct(pct, -100, 100, "An overlap")
    return _set_group(c, g, (el, key, pfx) -> begin
        _require_kind(el, (:barChart,), "overlap")
        _with_val(el, key, pfx, "overlap", v)
    end)
end

# ── Legend entries ────────────────────────────────────────────────────────────

_legend_node(root::XML.Node) =
    first_element_with_tag(first_element_with_tag(root, "chart"), "legend")

_series_idx(c::Chart, root::XML.Node, i::Integer) =
    _int_val(last(_series_pick(c, root, i)), "idx")

function _legend_entry(pfx, idx::Int, deleted::Bool)
    key = (NS_C, "legendEntry")
    e = XML.Element(prefixed_tag(pfx[NS_C], "legendEntry"))
    e = insert_child(e, key, _val_el(pfx, "idx", string(idx)))
    return insert_child(e, key, _val_el(pfx, "delete", deleted ? "1" : "0"))
end

"""
    getLegendEntryDeleted(c::Chart, i::Integer) -> Union{Nothing,Bool}

Whether series `i`'s legend entry is hidden (`c:legendEntry/c:delete`).
`nothing` means the chart has no legend, series `i` has no `c:legendEntry`, or
the entry carries text formatting instead, which the schema allows in place of
`c:delete`; in all three cases the entry is shown.

A legend entry is keyed by its series' `c:idx`, which this finds from the
position.
"""
function getLegendEntryDeleted(c::Chart, i::Integer)
    root = chart_root(c)
    lg   = _legend_node(root)
    isnothing(lg) && return nothing
    x = _series_idx(c, root, i)
    for k in XML.children(lg)
        localname(k) == "legendEntry" && _int_val(k, "idx") == x && return _bool_val(k, "delete")
    end
    return nothing
end

"""
    setLegendEntryDeleted(c, positions, deleted) -> Chart

Hide (`true`) or show (`false`) the legend entries of the series at
`positions` — a single position or any collection of them — in one rebuild of
the chart part. `:inherit` removes their `c:legendEntry` elements, so the
entries show with the legend's own formatting.

Replacing an entry discards any text formatting it carried: the schema allows
an entry either `c:delete` or `c:txPr`, not both.

Throws if the chart has no legend, rather than adding one.

Hiding entries is how to show a legend in an order other than series order:
add series with zero values in the wanted order, format them to match, and
hide the entries of the series that draw the data.
"""
function setLegendEntryDeleted(c::Chart, positions, deleted::Union{Bool,Symbol})
    deleted isa Symbol && deleted !== :inherit && throw(XLSXError(
        "`$(repr(deleted))` is not a legend entry state; use `true`, `false` or `:inherit`."))
    ps = positions isa Integer ? (positions,) : positions
    isempty(ps) && return c

    root = chart_root(c)
    pfx  = ns_prefixes(root)
    isnothing(_legend_node(root)) && throw(XLSXError(
        "Chart `$(c.name)` has no legend, so it has no legend entries to set."))
    idxs = unique(Int[_series_idx(c, root, i) for i in ps])
    lkey = (NS_C, "legend")

    new = rebuild_path(root, [(NS_C, "chart") => "chart", lkey => "legend"], lg -> begin
        kids    = isnothing(lg.children) ? XML.Node[] : lg.children
        others  = XML.Node[k for k in kids if localname(k) != "legendEntry"]
        entries = XML.Node[k for k in kids
                           if localname(k) == "legendEntry" && !(_int_val(k, "idx") in idxs)]
        deleted === :inherit || append!(entries, [_legend_entry(pfx, x, deleted) for x in idxs])
        sort!(entries; by = k -> _int_val(k, "idx"))
        foldl((l, e) -> insert_child(l, lkey, e), entries; init = _with_children(lg, others))
    end; prefixes = pfx)
    set_chart_root!(c, new)
    return c
end

const _TICK_MARKS = (:cross, :in, :out, :none)
const _TICK_MARK_ALIASES = Dict(:inside => :in, :outside => :out)   # Excel's names

_tick_mark(m::Symbol) = _check_sym(get(_TICK_MARK_ALIASES, m, m), _TICK_MARKS, "tick mark")

"""
    setAxisMajorTickMark(c, ax, mark) -> Chart

Set the axis's major tick marks (`c:majorTickMark`): `:out`, `:in`, `:cross` or
`:none`, or Excel's `:outside` and `:inside`. `:inherit` removes the element.
"""
function setAxisMajorTickMark(c::Chart, ax::ChartAxis, m::Symbol)
    v = _tick_mark(m)
    return _set_axis(c, ax, (el, key, pfx) -> _with_val(el, key, pfx, "majorTickMark", v))
end

"""
    setAxisMinorTickMark(c, ax, mark) -> Chart

As [`setAxisMajorTickMark`](@ref), for `c:minorTickMark`.
"""
function setAxisMinorTickMark(c::Chart, ax::ChartAxis, m::Symbol)
    v = _tick_mark(m)
    return _set_axis(c, ax, (el, key, pfx) -> _with_val(el, key, pfx, "minorTickMark", v))
end

"""
    setAxisGridlines(c, ax; major, minor) -> Chart

Show (`true`) or remove (`false`) the axis's major and minor gridlines. A
gridline element's presence is the switch, so `false` removes it rather than
writing an "off" value; a newly shown set takes its line from the chart style.
A keyword left unspecified is left alone.
"""
function setAxisGridlines(c::Chart, ax::ChartAxis; major = nothing, minor = nothing)
    all(isnothing, (major, minor)) && return c
    for v in (major, minor)
        v isa Union{Nothing,Bool} || throw(XLSXError("Gridlines are `true` or `false`; got $(repr(v))."))
    end
    return _set_axis(c, ax, (el, key, pfx) -> begin
        for (v, tag) in ((major, "majorGridlines"), (minor, "minorGridlines"))
            isnothing(v) && continue
            has = !isnothing(first_element_with_tag(el, tag))
            if v && !has
                el = insert_child(el, key, XML.Element(prefixed_tag(pfx[NS_C], tag)))
            elseif !v
                el = remove_child(el, tag)
            end
        end
        el
    end)
end

"""
    setAxisTextProp(c, ax, field, value) -> Chart

Set one text property of the axis's tick labels (`c:txPr` on the axis). Takes
the same fields and values as [`setLegendTextProp`](@ref).
"""
setAxisTextProp(c::Chart, ax::ChartAxis, field::Symbol, value) =
    _set_axis(c, ax, (el, key, pfx) -> _txpr_with_run_prop(el, key, field, value, pfx))

# Rebuild the spPr under the element `stepsof(root)` addresses, applying g(sp, pfx).
function _set_shape_under(c::Chart, stepsof, g)
    root = chart_root(c)
    pfx  = ns_prefixes(root)
    new  = rebuild_path(root, [stepsof(root)...; (NS_A, "spPr") => "spPr"],
                        sp -> g(sp, pfx); prefixes = pfx)
    set_chart_root!(c, new)
    return c
end

_ln_setter(kw) = (sp, pfx) -> _sp_with_line(sp, (NS_A, "spPr"), ln -> _ln_with(ln, pfx; kw...), pfx)

function _line_instruction(what::Symbol)
    what === :inherit && return (sp, _) -> remove_child(sp, "ln")
    what === :none    && return _ln_setter((; color = :none, width = nothing, dash = nothing,
                                              cap = nothing, compound = nothing,
                                              join = nothing, miterLimit = nothing))
    throw(XLSXError("`$what` is not a line instruction; use `:none` or `:inherit`."))
end

_axis_steps(c::Chart, ax::ChartAxis) = root -> first(_axis_path(c, root, ax))
_plotarea_steps(_) = [(NS_C, "chart") => "chart", (NS_C, "plotArea") => "plotArea"]

"""
    setAxisLine(c, ax; color, width, dash, cap, compound, join, miterLimit) -> Chart
    setAxisLine(c, ax, :none)
    setAxisLine(c, ax, :inherit)

Set the axis line, as [`setSeriesLine`](@ref) sets a series outline: several
properties in one rebuild, `:none` for a line that draws nothing, `:inherit` to
remove it so the chart style supplies it.
"""
function setAxisLine(c::Chart, ax::ChartAxis;
                     color = nothing, width = nothing, dash = nothing, cap = nothing,
                     compound = nothing, join = nothing, miterLimit = nothing)
    kw = (; color, width, dash, cap, compound, join, miterLimit)
    all(isnothing, kw) && return c
    return _set_shape_under(c, _axis_steps(c, ax), _ln_setter(kw))
end
setAxisLine(c::Chart, ax::ChartAxis, what::Symbol) =
    _set_shape_under(c, _axis_steps(c, ax), _line_instruction(what))

"""
    setPlotAreaLine(c; color, width, dash, cap, compound, join, miterLimit) -> Chart
    setPlotAreaLine(c, :none)
    setPlotAreaLine(c, :inherit)

Set the border of the plot area (`c:plotArea/c:spPr/a:ln`), as
[`setSeriesLine`](@ref) sets a series outline.
"""
function setPlotAreaLine(c::Chart;
                         color = nothing, width = nothing, dash = nothing, cap = nothing,
                         compound = nothing, join = nothing, miterLimit = nothing)
    kw = (; color, width, dash, cap, compound, join, miterLimit)
    all(isnothing, kw) && return c
    return _set_shape_under(c, _plotarea_steps, _ln_setter(kw))
end
setPlotAreaLine(c::Chart, what::Symbol) =
    _set_shape_under(c, _plotarea_steps, _line_instruction(what))


function _positive(v, what)
    v === :inherit && return :inherit
    x = Float64(v)
    isfinite(x) && x > 0 || throw(XLSXError("$what must be positive and finite; got $v."))
    return _xsd_double(x)
end

"""
    setAxisMajorUnit(c, ax, value) -> Chart

Set the interval between major ticks and gridlines (`c:majorUnit`), in the
axis's own units; `:inherit` returns it to automatic. Value and date axes only.

Ticks start at the axis minimum, so a minimum that is a multiple of the unit
puts a tick at zero.
"""
function setAxisMajorUnit(c::Chart, ax::ChartAxis, v)
    s = _positive(v, "A major unit")
    return _set_axis(c, ax, (el, key, pfx) -> begin
        _require_kind(el, (:valAx, :dateAx), "majorUnit")
        _with_val(el, key, pfx, "majorUnit", s)
    end)
end

"""
    setAxisMinorUnit(c, ax, value) -> Chart

As [`setAxisMajorUnit`](@ref), for the interval between minor ticks
(`c:minorUnit`).
"""
function setAxisMinorUnit(c::Chart, ax::ChartAxis, v)
    s = _positive(v, "A minor unit")
    return _set_axis(c, ax, (el, key, pfx) -> begin
        _require_kind(el, (:valAx, :dateAx), "minorUnit")
        _with_val(el, key, pfx, "minorUnit", s)
    end)
end