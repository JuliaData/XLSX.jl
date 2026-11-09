# Accessors for ChartEx: charts in the Microsoft cx: namespace
# (http://schemas.microsoft.com/office/drawing/2014/chartex).
#
# A ChartEx holds identity only; everything here reads the current part, so
# results never go stale. Series are indexed 1-based over cx:series in document
# order. A Pareto chart's paretoLine is owned by its column series and is not
# counted; it is reached with pareto = true.
#
# Discovery readers (_cx_root, _cx_chart, _cx_series_nodes, _cx_layouts,
# _cx_has_binning, _cx_refs, _cx_title) live in discovery.jl, and the creation
# helpers (_chartex_shell, _cx_add_series and friends) in creation.jl.
#

# ---- small readers ---------------------------------------------------------

# Series positions count the series a user sees: those not owned by another.
# A Pareto chart's paretoLine is owned by its column series (`ownerIdx`, the
# owner's 0-based position among the `cx:series`) and is reached with
# `pareto = true` on the owner's position.
function _cx_series_node(c::ChartEx, root::XML.Node, i::Integer; pareto::Bool = false)
    sers   = _cx_series_nodes(root)
    owners = findall(s -> isnothing(_cx_attr(s, "ownerIdx")), sers)
    n = length(owners)
    n == 0 && throw(XLSXError("Chart `$(c.name)` has no series."))
    1 <= i <= n || throw(XLSXError("Chart `$(c.name)` has $n series; asked for series $i."))
    k = owners[i]
    pareto || return sers[k]
    j = findfirst(s -> _cx_attr(s, "ownerIdx") == string(k - 1), sers)
    isnothing(j) && throw(XLSXError("Series $i of chart `$(c.name)` has no Pareto line."))
    return sers[j]
end

_cx_series_node(c::ChartEx, i::Integer; pareto::Bool = false) =
    _cx_series_node(c, _cx_root(c), i; pareto)

_cx_layoutpr(c::ChartEx, i::Integer; pareto::Bool = false) =
    first_element_with_tag(_cx_series_node(c, i; pareto), "layoutPr")
_cx_datalabels(c::ChartEx, i::Integer; pareto::Bool = false) =
    first_element_with_tag(_cx_series_node(c, i; pareto), "dataLabels")

# Attribute value as written, or `nothing` when the node or attribute is absent.
_cx_attr(node, name) = isnothing(node) ? nothing : (s = get_attr(node, name); isempty(s) ? nothing : s)

_cx_symbol(node, name) = (s = _cx_attr(node, name); isnothing(s) ? nothing : Symbol(s))

function _cx_bool(node, name)::Union{Nothing,Bool}
    s = _cx_attr(node, name)
    isnothing(s) && return nothing
    s in ("1", "true")  && return true
    s in ("0", "false") && return false
    throw(XLSXError("Attribute `$name` should be a boolean. Found \"$s\"."))
end

function _cx_num_or_auto(node, name, T::Type)
    s = _cx_attr(node, name)
    isnothing(s) && return nothing
    s == "auto" && return :auto
    return parse(T, s)
end

# cx:tx holds typed text or a cell binding. Excel writes both as cx:txData:
# a bound one has cx:f and a cached cx:v, a typed one only cx:v. cx:rich is
# allowed by the schema and read if present. Empty text is no text, as for a c:
# title: a title bound to an empty cell caches an empty cx:v and reads as nothing.
function _cx_tx_text(tx)::Union{Nothing,String}
    isnothing(tx) && return nothing
    rich = first_element_with_tag(tx, "rich")
    t = isnothing(rich) ? child_text(first_element_with_tag(tx, "txData"), "v") : text_content(rich)
    return isnothing(t) || isempty(t) ? nothing : t
end

function _cx_tx_range(xf::XLSXFile, tx)::ChartRange
    isnothing(tx) && return nothing
    fml = child_text(first_element_with_tag(tx, "txData"), "f")
    return isnothing(fml) ? nothing : resolve_chartex_ref(xf, fml)
end

_cx_title_tx(c::ChartEx) = first_element_with_tag(first_element_with_tag(_cx_chart(c), "title"), "tx")

# ---- series and data -------------------------------------------------------

"""
    getChartSeriesCount(c::ChartEx) -> Int

Number of series in the chart. A Pareto chart's Pareto line is not counted: it
belongs to its column series, and is reached with `pareto = true`.
"""
getChartSeriesCount(c::ChartEx)::Int = count(s -> isnothing(_cx_attr(s, "ownerIdx")), _cx_series_nodes(c))

"""
    getSeriesLayout(c::ChartEx, i; pareto=false) -> Symbol

The layout of series `i` as written in its `layoutId` attribute, e.g.
`:waterfall`, `:funnel`, `:treemap`, `:clusteredColumn`. With `pareto = true`,
the layout of series `i`'s Pareto line, `:paretoLine`; an error if it has none.

See also [`getChartType`](@ref), which describes the chart as a whole.
"""
function getSeriesLayout(c::ChartEx, i::Integer; pareto::Bool = false)::Symbol
    lid = get_attr(_cx_series_node(c, i; pareto), "layoutId")
    isempty(lid) && throw(XLSXError("Series $i of chart `$(c.name)` has no layoutId."))
    return Symbol(lid)
end

function _cx_dimension(xf::XLSXFile, dim::XML.Node)::ChartExDimension
    kind = localname(dim) == "numDim" ? :num : :str
    fml  = child_text(dim, "f")
    rng  = isnothing(fml) ? nothing : resolve_chartex_ref(xf, fml)
    return ChartExDimension(kind, Symbol(get_attr(dim, "type")), fml, rng)
end

"""
    getChartDataBlocks(c::ChartEx) -> Vector{ChartExData}

Every `cx:data` block in the chart, in document order, with its dimensions.
"""
function getChartDataBlocks(c::ChartEx)::Vector{ChartExData}
    out = ChartExData[]
    chartdata = first_element_with_tag(_cx_root(c), "chartData")
    isnothing(chartdata) && return out
    for d in elements_with_tag(chartdata, "data")
        dims = [_cx_dimension(c.package, dim) for dim in XML.eachelement(d)
                if localname(dim) in ("numDim", "strDim")]
        push!(out, ChartExData(parse(Int, get_attr(d, "id")), dims))
    end
    return out
end

"""
    getSeriesData(c::ChartEx, i; pareto=false) -> Union{Nothing, ChartExData}

The data block series `i` draws from, found through its `cx:dataId`. With
`pareto = true`, series `i`'s Pareto line instead; an error if it has none.
`nothing` if the series names no data block.
"""
function getSeriesData(c::ChartEx, i::Integer; pareto::Bool = false)::Union{Nothing,ChartExData}
    idnode = first_element_with_tag(_cx_series_node(c, i; pareto), "dataId")
    isnothing(idnode) && return nothing
    id = parse(Int, get_attr(idnode, "val"))
    blocks = getChartDataBlocks(c)
    k = findfirst(b -> b.id == id, blocks)
    isnothing(k) &&
        throw(XLSXError("Series $i of chart `$(c.name)` refers to data block $id, which does not exist."))
    return blocks[k]
end

"""
    getSeriesHidden(c::ChartEx, i; pareto=false) -> Union{Nothing, Bool}

The series' `hidden` attribute. With `pareto = true`, series `i`'s Pareto line
instead; an error if it has none. `nothing` if not written.
"""
getSeriesHidden(c::ChartEx, i::Integer; pareto::Bool = false) =
    _cx_bool(_cx_series_node(c, i; pareto), "hidden")

"""
    getSeriesAxisIds(c::ChartEx, i; pareto=false) -> Vector{Int}

The `id`s of the `cx:axis` elements series `i` is plotted against, as written.
These are axis ids, not positions. Empty if the series names none.

With `pareto = true`, series `i`'s Pareto line instead; an error if it has none.
"""
getSeriesAxisIds(c::ChartEx, i::Integer; pareto::Bool = false)::Vector{Int} =
    [parse(Int, get_attr(n, "val")) for n in elements_with_tag(_cx_series_node(c, i; pareto), "axisId")]

# ---- names and titles ------------------------------------------------------

"""
    getSeriesName(c::ChartEx, i; pareto=false) -> Union{Nothing, String}

The name of series `i`, typed or the cached value of a cell binding.
`nothing` if the series has no name of its own.

With `pareto = true`, series `i`'s Pareto line instead; an error if it has none.
"""
getSeriesName(c::ChartEx, i::Integer; pareto::Bool = false) =
    _cx_tx_text(first_element_with_tag(_cx_series_node(c, i; pareto), "tx"))

"""
    getSeriesNameRange(c::ChartEx, i; pareto=false) -> ChartRange

The cell a series name is bound to. `nothing` if the name is typed or absent.

With `pareto = true`, series `i`'s Pareto line instead; an error if it has none.
"""
getSeriesNameRange(c::ChartEx, i::Integer; pareto::Bool = false) =
    _cx_tx_range(c.package, first_element_with_tag(_cx_series_node(c, i; pareto), "tx"))

"""
    getChartTitleRange(c::ChartEx) -> ChartRange
    getChartTitleRange(c::Chart) -> ChartRange

The cell the chart title is bound to, resolved — through a defined name where
the formula is one. `nothing` if the title is typed or absent. See
[`getChartTitleRef`](@ref) for the formula as written.
"""
getChartTitleRange(c::ChartEx) = _cx_tx_range(c.package, _cx_title_tx(c))

getChartTitleRef(c::ChartEx) = _cx_tx_ref(_cx_title_tx(c))

# The cx:f text of a cx:tx, as written; `nothing` for typed text.
_cx_tx_ref(tx) = isnothing(tx) ? nothing : child_text(first_element_with_tag(tx, "txData"), "f")

_cx_axis_title_tx(c::ChartEx, id::Integer) =
    first_element_with_tag(first_element_with_tag(_cx_axis_node(c, id), "title"), "tx")

getAxisTitle(c::ChartEx, id::Integer)      = _cx_tx_text(_cx_axis_title_tx(c, id))
getAxisTitleRef(c::ChartEx, id::Integer)   = _cx_tx_ref(_cx_axis_title_tx(c, id))
getAxisTitleRange(c::ChartEx, id::Integer) = _cx_tx_range(c.package, _cx_axis_title_tx(c, id))

getChartTitleTextProps(c::ChartEx) =
    parse_drawing_text(_wb(c), first_element_with_tag(_cx_chart(c), "title"))

getLegendTextProps(c::ChartEx) =
    parse_drawing_text(_wb(c), first_element_with_tag(_cx_chart(c), "legend"))

# ---- layout properties -----------------------------------------------------

"""
    getSeriesSubtotals(c::ChartEx, i) -> Union{Nothing, Vector{Int}}

The 1-based points of a waterfall series drawn as totals. An empty vector means
`cx:subtotals` is present with no points; `nothing` means it is absent.
"""
function getSeriesSubtotals(c::ChartEx, i::Integer)::Union{Nothing,Vector{Int}}
    st = first_element_with_tag(_cx_layoutpr(c, i), "subtotals")
    isnothing(st) && return nothing
    return [parse(Int, get_attr(n, "val")) + 1 for n in elements_with_tag(st, "idx")]
end

"""
    getSeriesBinning(c::ChartEx, i) -> Union{Nothing, ChartExBinning}

Histogram binning for series `i`; `nothing` if the series has no `cx:binning`.
"""
function getSeriesBinning(c::ChartEx, i::Integer)::Union{Nothing,ChartExBinning}
    b = first_element_with_tag(_cx_layoutpr(c, i), "binning")
    isnothing(b) && return nothing
    return ChartExBinning(
        _cx_symbol(b, "intervalClosed"),
        _cx_num_or_auto(b, "underflow", Float64),
        _cx_num_or_auto(b, "overflow", Float64),
        _cx_num_or_auto(first_element_with_tag(b, "binSize"), "val", Float64),
        _cx_num_or_auto(first_element_with_tag(b, "binCount"), "val", Int))
end

"""
    getSeriesAggregation(c::ChartEx, i) -> Bool

Whether series `i` aggregates repeated categories (`cx:aggregation`). The
element carries no value; its presence is the setting.
"""
getSeriesAggregation(c::ChartEx, i::Integer)::Bool =
    !isnothing(first_element_with_tag(_cx_layoutpr(c, i), "aggregation"))

"""
    getSeriesQuartileMethod(c::ChartEx, i) -> Union{Nothing, Symbol}

Box & whisker quartile calculation, `:inclusive` or `:exclusive`.
"""
getSeriesQuartileMethod(c::ChartEx, i::Integer) =
    _cx_symbol(first_element_with_tag(_cx_layoutpr(c, i), "statistics"), "quartileMethod")

"""
    getSeriesParentLabelLayout(c::ChartEx, i) -> Union{Nothing, Symbol}

Treemap parent label placement, e.g. `:overlapping`, `:banner`, `:none`.
"""
getSeriesParentLabelLayout(c::ChartEx, i::Integer) =
    _cx_symbol(first_element_with_tag(_cx_layoutpr(c, i), "parentLabelLayout"), "val")

const _CX_LAYOUT_FLAGS = (:meanLine, :meanMarker, :nonoutliers, :outliers, :connectorLines)

"""
    getSeriesLayoutFlag(c::ChartEx, i, flag) -> Union{Nothing, Bool}

One attribute of `cx:layoutPr/cx:visibility`: `:meanLine`, `:meanMarker`,
`:nonoutliers`, `:outliers` (box & whisker) or `:connectorLines` (waterfall).
`nothing` if not written.
"""
function getSeriesLayoutFlag(c::ChartEx, i::Integer, flag::Symbol)::Union{Nothing,Bool}
    flag in _CX_LAYOUT_FLAGS ||
        throw(XLSXError("Unknown layout flag `$flag`. Expected one of $(join(_CX_LAYOUT_FLAGS, ", "))."))
    return _cx_bool(first_element_with_tag(_cx_layoutpr(c, i), "visibility"), String(flag))
end

# ---- data labels -----------------------------------------------------------

const _CX_LABEL_FLAGS = (:seriesName, :categoryName, :value)

"""
    getLabelFlag(c::ChartEx, i, flag; pareto=false) -> Union{Nothing, Bool}

Whether series `i`'s data labels show `:seriesName`, `:categoryName` or
`:value`, from `cx:dataLabels/cx:visibility`. With `pareto = true`, series `i`'s
Pareto line instead; an error if it has none. `nothing` if not written.
"""
function getLabelFlag(c::ChartEx, i::Integer, flag::Symbol; pareto::Bool = false)::Union{Nothing,Bool}
    flag in _CX_LABEL_FLAGS ||
        throw(XLSXError("Unknown label flag `$flag`. Expected one of $(join(_CX_LABEL_FLAGS, ", "))."))
    return _cx_bool(first_element_with_tag(_cx_datalabels(c, i; pareto), "visibility"), String(flag))
end

"""
    setLabelFlag(c::ChartEx, i, flag, on::Bool; pareto=false) -> ChartEx

Show or hide one part of series `i`'s data labels, as the Label Contains boxes
of Excel's Format Data Labels pane do: `:seriesName`, `:categoryName` or
`:value`. A series with no labels is given Excel's: just the parts, with no
position or formatting of their own, as Excel writes labels turned on with Chart
Elements → Data Labels. When two or more parts show, they are separated by
`", "`, as in Excel.

As in Excel, a histogram's or Pareto chart's labels have no category name, and a
Pareto line (`pareto = true`) has no labels at all; both are errors.
"""
function setLabelFlag(c::ChartEx, i::Integer, flag::Symbol, on::Bool; pareto::Bool = false)
    flag in _CX_LABEL_FLAGS ||
        throw(XLSXError("Unknown label flag `$flag`. Expected one of $(join(_CX_LABEL_FLAGS, ", "))."))
    pareto && throw(XLSXError(
        "Excel shows no data labels on a Pareto line. Label its bars instead, with `pareto = false`."))
    kind = getChartType(c)
    flag === :categoryName && kind in (:histogram, :pareto) && throw(XLSXError(
        "A $kind chart's data labels have no category name, as in Excel. Use `:seriesName` or `:value`."))
    return _cx_edit_series!(c, i, (ser, pfx) ->
        rebuild_path(ser, [(NS_CX, "dataLabels") => "dataLabels"], function (lbl)
            lbl = _cx_visibility_off(lbl, pfx)
            vis = first_element_with_tag(lbl, "visibility")
            new = with_attribute(vis, String(flag), on ? "1" : "0")
            lbl = replace_child(lbl, vis, new)
            shown = count(f -> _cx_bool(new, String(f)) === true, _CX_LABEL_FLAGS)
            shown >= 2 && isnothing(first_element_with_tag(lbl, "separator")) &&
                (lbl = insert_child(lbl, (NS_CX, "dataLabels"),
                                    XML.Element(prefixed_tag(pfx[NS_CX], "separator"), XML.Text(", "))))
            return lbl
        end; prefixes = pfx, parent_key = (NS_CX, "series")))
end

"""
    getLabelPosition(c::ChartEx, i; pareto=false) -> Union{Nothing, Symbol}

Data label position for series `i`, e.g. `:outEnd` (outside-end), `:inEnd`
(inside-end), `:ctr` (centre). With
`pareto = true`, series `i`'s Pareto line instead; an error if it has none.
"""
getLabelPosition(c::ChartEx, i::Integer; pareto::Bool = false) =
    _cx_symbol(_cx_datalabels(c, i; pareto), "pos")

# ---- formatting cascades ---------------------------------------------------

# The cx:dataPt for 1-based `point` of series `i`, or `nothing` if Excel never
# formatted that point. cx:dataPt idx is 0-based.
function _cx_datapt(c::ChartEx, i::Integer, point::Integer; pareto::Bool = false)
    point >= 1 || throw(XLSXError("Data point positions start at 1; asked for $point."))
    for dp in elements_with_tag(_cx_series_node(c, i; pareto), "dataPt")
        tryparse(Int, get_attr(dp, "idx")) == point - 1 && return dp
    end
    return nothing
end

# spPr cascade: data point, then series. As for c:, nothing above the series
# describes a series' graphic; the rest comes from the style part.
function _cx_shape_chain(c::ChartEx, i::Integer; point::Union{Nothing,Integer} = nothing,
                         pareto::Bool = false)
    chain = FormatSite[]
    if !isnothing(point)
        dp = _cx_datapt(c, i, point; pareto)
        isnothing(dp) || push!(chain, _site_path(:point, :shape, dp, "spPr"))
    end
    push!(chain, _site_path(:series, :shape, _cx_series_node(c, i; pareto), "spPr"))
    return chain
end

# txPr cascade for data labels: series, then chart space. cx has no chart groups.
_cx_text_chain(c::ChartEx, i::Integer; pareto::Bool = false) = FormatSite[
    _site_path(:series,     :text, _cx_series_node(c, i; pareto), "dataLabels", "txPr"),
    _site_path(:chartspace, :text, _cx_root(c),                   "txPr"),
]

"""
    getSeriesFill(c::ChartEx, i::Integer; point=nothing, pareto=false) -> Effective{DrawingFill}

Resolve the fill of series `i`, or of one data point of it, walking
`cx:dataPt/cx:spPr`, then `cx:series/cx:spPr`, then the chart's style. With
`pareto = true`, series `i`'s Pareto line instead; an error if it has none.

Excel writes no fill for a series or point it has not been asked to format, and
draws one from the chart's style and colours parts instead. That last rung
reports the same colour: `site.level === :style`, and the colour is a theme
colour such as `accent2` with Excel's variation applied (the 7th series is
`accent1` darkened by 40%), so `.value.fgcolor.rgb` is what Excel shows. It
answers for:

- box & whisker, funnel, histogram and the bars of a Pareto chart, by series;
- a waterfall's points: increase, decrease and total each have a colour, so
  `point` is needed;
- a treemap's or sunburst's points, by top-level category: each point takes the
  colour of the branch it belongs to, so `point` is needed.

It gives no answer — `value` stays `nothing` when nothing else is written — for
other layouts, for a Pareto line (a line has no fill; see
[`getSeriesLine`](@ref)), and under a monochromatic colour scheme (Chart Design →
Change Colors → Monochromatic), whose shades Excel computes in a way that is not
reproduced. The colour of a series, or of a point's branch, is checked against
Excel up to the 8th; later ones follow the variations the colours part lists.
"""
function getSeriesFill(c::ChartEx, i::Integer; point::Union{Nothing,Integer} = nothing, pareto::Bool = false)
    chain = _cx_shape_chain(c, i; point, pareto)
    site  = _cx_style_fill_site(c, i; point, pareto)
    isnothing(site) || push!(chain, site)
    return _walk_fill(_wb(c), chain)
end

"""
    getSeriesLine(c::ChartEx, i::Integer; point=nothing, pareto=false) -> Effective{DrawingLine}

Resolve the outline of series `i`, or of one data point of it, walking
`cx:dataPt/cx:spPr` then `cx:series/cx:spPr`. With `pareto = true`, series
`i`'s Pareto line instead; an error if it has none.
"""
getSeriesLine(c::ChartEx, i::Integer; point::Union{Nothing,Integer} = nothing, pareto::Bool = false) =
    _walk_line(_wb(c), _cx_shape_chain(c, i; point, pareto))

"""
    getLabelTextProp(c::ChartEx, i::Integer, field::Symbol; pareto=false) -> Effective

Resolve one field of the data-label text formatting for series `i`, walking
`cx:dataLabels/cx:txPr` then `cx:chartSpace/cx:txPr`. With `pareto = true`,
series `i`'s Pareto line instead; an error if it has none. See the `Chart`
method for how fields inherit.
"""
getLabelTextProp(c::ChartEx, i::Integer, field::Symbol; pareto::Bool = false) =
    _resolve_text_field(_wb(c), _cx_text_chain(c, i; pareto), field)

# ---- write path -------------------------------------------------------------
#
# Chart and ChartEx are both handles, so setters write the part and return `c`
# itself, which remains valid. The cx value types (`ChartExData`,
# `ChartExBinning`, …) hold no nodes: everything is addressed by position or key
# against the current part on each call.

const _CX_ROOT_KEY = (NS_CX, "chartSpace")

# rebuild_path steps from `root` to series `i`. Matches by node identity, so the
# series node must come from the same root the path is applied to.
function _cx_series_path(c::ChartEx, root::XML.Node, i::Integer; pareto::Bool = false)
    ser = _cx_series_node(c, root, i; pareto)
    return [(NS_CX, "chart")          => "chart",
            (NS_CX, "plotArea")       => "plotArea",
            (NS_CX, "plotAreaRegion") => "plotAreaRegion",
            (NS_CX, "series")         => ("series", n -> n === ser)]
end

# Apply `f(ser, pfx)` to series `i` and write the part back once. A throw inside
# `f` leaves the part untouched.
function _cx_edit_series!(c::ChartEx, i::Integer, f; pareto::Bool = false)
    root = _cx_root(c)
    pfx  = ns_prefixes(root)
    new  = rebuild_path(root, _cx_series_path(c, root, i; pareto), ser -> f(ser, pfx);
                        prefixes = pfx, parent_key = _CX_ROOT_KEY)
    set_chart_root!(c, new)
    return c
end

_cx_idx(n) = tryparse(Int, get_attr(n, "idx"))

# Apply `f` to the `tag` child of `parent` whose index is `idx`, creating it with
# `make()` if absent. A created child goes in index order among its `tag`
# siblings, or at its schema position if it is the first. `idx_of` reads an
# index from a sibling; `parent_key` is `parent`'s SchemaKey.
function _with_indexed_child(parent::XML.Node, parent_key::SchemaKey, tag::AbstractString,
                             idx::Integer, idx_of, make, f)
    kids = isnothing(parent.children) ? XML.Node[] : parent.children
    j = findfirst(k -> localname(k) == tag && idx_of(k) == idx, kids)
    if isnothing(j)
        fresh = make()
        sibs = findall(k -> localname(k) == tag, kids)
        if isempty(sibs)
            parent = insert_child(parent, parent_key, fresh)       # first one: schema position
        else
            # before the first sibling with a larger index, else after the last
            k  = findfirst(p -> something(idx_of(kids[p]), -1) > idx, sibs)
            at = isnothing(k) ? last(sibs) + 1 : sibs[k]
            new_kids = Vector{eltype(kids)}(undef, 0)
            append!(new_kids, kids[1:at-1]); push!(new_kids, fresh); append!(new_kids, kids[at:end])
            parent = _with_children(parent, new_kids)
        end
        j = findfirst(n -> n === fresh, parent.children)
    end
    node = parent.children[j]
    return replace_child(parent, node, f(node))
end

# Apply `f` to the cx:dataPt for 1-based `point`, creating it if absent. A new
# one is inserted in idx order among its siblings: before the first dataPt with a
# larger idx, or in schema position (after the last dataPt) if there is none.
function _cx_with_datapt(ser::XML.Node, point::Integer, pfx, f)
    point >= 1 || throw(XLSXError("Data point positions start at 1; asked for $point."))
    return _with_indexed_child(ser, (NS_CX, "series"), "dataPt", point - 1, _cx_idx,
        () -> XML.Element(prefixed_tag(pfx[NS_CX], "dataPt"); idx = string(point - 1)), f)
end

# Apply `f(spPr, pfx)` to the spPr of series `i`, or of one of its points.
# With `create = false`, a missing spPr (or dataPt) means there is nothing to
# change, and nothing is created: removing a fill that was never written must
# not leave an empty <cx:spPr/> behind.
function _cx_set_shape!(c::ChartEx, i::Integer, point, f; create::Bool = true, pareto::Bool = false)
    isnothing(point) || point >= 1 ||
        throw(XLSXError("Data point positions start at 1; asked for $point."))
    _cx_edit_series!(c, i, (ser, pfx) -> begin
        edit(node, key) =
            (!create && isnothing(first_element_with_tag(node, "spPr"))) ? node :
            rebuild_path(node, [(NS_A, "spPr") => "spPr"], sp -> f(sp, pfx);
                         prefixes = pfx, parent_key = key)
        isnothing(point) && return edit(ser, (NS_CX, "series"))
        (!create && isnothing(_cx_datapt_in(ser, point))) && return ser
        return _cx_with_datapt(ser, point, pfx, dp -> edit(dp, (NS_CX, "dataPt")))
    end; pareto)
end

_cx_datapt_in(ser, point) =
    findfirst(k -> localname(k) == "dataPt" && _cx_idx(k) == point - 1,
              isnothing(ser.children) ? XML.Node[] : ser.children)

"""
    setSeriesFill(c::ChartEx, i, color; point=nothing, pareto=false) -> ChartEx

Set the fill of series `i`, or of data point `point` of it, counting from 1.
$(_DOC_COLOR_ARG) With `pareto = true`, series `i`'s Pareto line instead; an
error if it has none.

Setting a point's fill creates its `cx:dataPt` if Excel never formatted that
point. `:inherit` never creates anything.

Returns `c`, which remains valid.
"""
setSeriesFill(c::ChartEx, i::Integer, color::Union{AbstractString,Colors.Colorant,SchemeColor};
              point::Union{Nothing,Integer} = nothing, pareto::Bool = false) =
    _cx_set_shape!(c, i, point, (sp, pfx) -> _sp_with_fill(sp, (NS_A, "spPr"), color, pfx); pareto)

function setSeriesFill(c::ChartEx, i::Integer, what::Symbol;
                       point::Union{Nothing,Integer} = nothing, pareto::Bool = false)
    what === :inherit &&
        return _cx_set_shape!(c, i, point, (sp, pfx) -> _sp_with_fill(sp, (NS_A, "spPr"), :inherit, pfx);
                              create = false, pareto)
    what === :none &&
        return _cx_set_shape!(c, i, point, (sp, pfx) -> _sp_with_fill(sp, (NS_A, "spPr"), :none, pfx); pareto)
    return setSeriesFill(c, i, String(what); point, pareto)
end

# Apply `f(ln, pfx)` to the a:ln of series `i`, or of one of its points,
# creating the a:ln and its cx:spPr if absent. One rebuild; a throw inside `f`
# leaves the part untouched.
_cx_set_line!(c::ChartEx, i::Integer, point, f; create::Bool = true, pareto::Bool = false) =
    _cx_set_shape!(c, i, point,
                   (sp, pfx) -> _sp_with_line(sp, (NS_A, "spPr"), ln -> f(ln, pfx), pfx);
                   create, pareto)

"""
    setSeriesLine(c::ChartEx, i; point=nothing, pareto=false, color, width, dash, cap, compound, join, miterLimit) -> ChartEx
    setSeriesLine(c::ChartEx, i, :none; point=nothing, pareto=false)
    setSeriesLine(c::ChartEx, i, :inherit; point=nothing, pareto=false)

Set several outline properties of series `i`, or of data point `point` of it,
in one rebuild of the chart part. $(_DOC_LINE_KWS) A keyword left unspecified
is left alone; pass `:inherit` to remove one that is set. With `pareto = true`,
series `i`'s Pareto line instead; an error if it has none.

The symbol form acts on the whole outline: `:none` writes an `a:ln` whose fill is
`<a:noFill/>`, and `:inherit` removes the `a:ln` so the chart style supplies it.

Returns `c`, which remains valid.
"""
function setSeriesLine(c::ChartEx, i::Integer; point::Union{Nothing,Integer} = nothing,
                       pareto::Bool = false,
                       color = nothing, width = nothing, dash = nothing,
                       cap = nothing, compound = nothing,
                       join = nothing, miterLimit = nothing)
    all(isnothing, (color, width, dash, cap, compound, join, miterLimit)) && return c
    return _cx_set_line!(c, i, point, (ln, pfx) ->
        _ln_with(ln, pfx; color, width, dash, cap, compound, join, miterLimit); pareto)
end

function setSeriesLine(c::ChartEx, i::Integer, what::Symbol;
                       point::Union{Nothing,Integer} = nothing, pareto::Bool = false)
    what === :inherit &&
        return _cx_set_shape!(c, i, point, (sp, _) -> remove_child(sp, "ln"); create = false, pareto)
    what === :none && return setSeriesLineColor(c, i, :none; point, pareto)
    throw(XLSXError("`$what` is not a line instruction; use `:none` or `:inherit`."))
end

"""
    setSeriesLineColor(c::ChartEx, i, color; point=nothing, pareto=false) -> ChartEx

Set the color of series `i`'s outline, or of data point `point` of it.
$(_DOC_COLOR_ARG) `:none` writes an outline with `<a:noFill/>`; `:inherit`
removes the color and leaves the rest of the `a:ln`. With `pareto = true`,
series `i`'s Pareto line instead; an error if it has none.
"""
setSeriesLineColor(c::ChartEx, i::Integer, color; point::Union{Nothing,Integer} = nothing,
                   pareto::Bool = false) =
    _cx_set_line!(c, i, point, (ln, pfx) -> _ln_with_color(ln, color, pfx); pareto)

"""
    setSeriesLineWidth(c::ChartEx, i, points; point=nothing, pareto=false) -> ChartEx

Set the width of series `i`'s outline, or of data point `point`'s, in points —
a `Real` from 0 to 1584 — or `:inherit` to remove it. With `pareto = true`,
series `i`'s Pareto line instead.
"""
setSeriesLineWidth(c::ChartEx, i::Integer, pts; point::Union{Nothing,Integer} = nothing,
                   pareto::Bool = false) =
    _cx_set_line!(c, i, point, (ln, _) -> _ln_with_width(ln, pts); pareto)

"""
    setSeriesLineDash(c::ChartEx, i, dash; point=nothing, pareto=false) -> ChartEx

Set the dash pattern of series `i`'s outline, or of data point `point`'s:
$(_doc_syms(DASH_VALUES)), or Excel's names $(_doc_aliases(DASH_ALIASES)); or
`:inherit` to remove it. With `pareto = true`, series `i`'s Pareto line instead.
"""
setSeriesLineDash(c::ChartEx, i::Integer, dash; point::Union{Nothing,Integer} = nothing,
                  pareto::Bool = false) =
    _cx_set_line!(c, i, point, (ln, pfx) -> _ln_with_dash(ln, dash, pfx); pareto)

"""
    setSeriesLineCap(c::ChartEx, i, cap; point=nothing, pareto=false) -> ChartEx

Set how series `i`'s outline ends, or data point `point`'s:
$(_doc_syms(CAP_VALUES)), or Excel's names $(_doc_aliases(CAP_ALIASES)); or
`:inherit` to remove it. With `pareto = true`, series `i`'s Pareto line instead.
"""
setSeriesLineCap(c::ChartEx, i::Integer, cap; point::Union{Nothing,Integer} = nothing,
                 pareto::Bool = false) =
    _cx_set_line!(c, i, point, (ln, _) -> _ln_with_cap(ln, cap); pareto)

"""
    setSeriesLineCompound(c::ChartEx, i, cmpd; point=nothing, pareto=false) -> ChartEx

Set the compound line type of series `i`'s outline, or data point `point`'s:
$(_doc_syms(CMPD_VALUES)), or Excel's names $(_doc_aliases(CMPD_ALIASES)); or
`:inherit` to remove it. With `pareto = true`, series `i`'s Pareto line instead.
"""
setSeriesLineCompound(c::ChartEx, i::Integer, cmpd; point::Union{Nothing,Integer} = nothing,
                      pareto::Bool = false) =
    _cx_set_line!(c, i, point, (ln, _) -> _ln_with_compound(ln, cmpd); pareto)

"""
    setSeriesLineJoin(c::ChartEx, i, join; point=nothing, pareto=false) -> ChartEx

Set how the segments of series `i`'s outline meet, or data point `point`'s:
$(_doc_syms(JOIN_VALUES)); or `:inherit` to remove it. Changing the join discards
any miter limit. With `pareto = true`, series `i`'s Pareto line instead.
"""
setSeriesLineJoin(c::ChartEx, i::Integer, join; point::Union{Nothing,Integer} = nothing,
                  pareto::Bool = false) =
    _cx_set_line!(c, i, point, (ln, pfx) -> _ln_with_join(ln, join, pfx); pareto)

"""
    setSeriesLineMiterLimit(c::ChartEx, i, limit; point=nothing, pareto=false) -> ChartEx

Set the miter limit of series `i`'s outline, or data point `point`'s, as a
`Real` multiple of the line width; or `:inherit` to remove it. The outline must
already have a miter join ([`setSeriesLineJoin`](@ref)). With `pareto = true`,
series `i`'s Pareto line instead.
"""
setSeriesLineMiterLimit(c::ChartEx, i::Integer, lim; point::Union{Nothing,Integer} = nothing,
                        pareto::Bool = false) =
    _cx_set_line!(c, i, point, (ln, _) -> _ln_with_miter_limit(ln, lim); pareto)

# A created cx:dataLabels showing no flags would show labels with Excel's
# defaults, as c:dLbls does, so all three are written off when creating one.
# Unlike c:, they are attributes on a single cx:visibility child.
const _CX_LABEL_FLAG_NAMES = ("seriesName", "categoryName", "value")

function _cx_visibility_off(lbl::XML.Node, pfx::Dict{String,String})
    isnothing(first_element_with_tag(lbl, "visibility")) || return lbl
    vis = XML.Element(prefixed_tag(pfx[NS_CX], "visibility");
                      seriesName = "0", categoryName = "0", value = "0")
    return insert_child(lbl, (NS_CX, "dataLabels"), vis)
end

"""
    setLabelTextProp(c::ChartEx, i, field, value; pareto=false) -> ChartEx

Set one field of the data-label text formatting for series `i`, writing it on the
series' `cx:dataLabels/cx:txPr`. With `pareto = true`, series `i`'s Pareto line
instead; an error if it has none.

$(_DOC_TEXT_FIELD_ARG)

Creating a `cx:dataLabels` to hold formatting writes `cx:visibility` with all
three flags off, since one that names no flags shows labels with Excel's
defaults rather than none.

Returns `c`, which remains valid.
"""
function setLabelTextProp(c::ChartEx, i::Integer, field::Symbol, value; pareto::Bool = false)
    get(_RUN_PROP_ALIASES, field, field) in _RUN_PROP_FIELDS || throw(XLSXError(
        "`$field` is not a text property. Valid fields: " *
        join(_RUN_PROP_FIELDS, ", ") * "; or Excel's names " *
        join(("$k (= $v)" for (k, v) in _RUN_PROP_ALIASES), ", ") * "."))
    return _cx_edit_series!(c, i, (ser, pfx) ->
        rebuild_path(ser, [(NS_CX, "dataLabels") => "dataLabels"],
                     lbl -> _cx_label_transform(lbl, field, value, pfx);
                     prefixes = pfx, parent_key = (NS_CX, "series"));
        pareto)
end

# Apply `field` to the cx:dataLabels' cx:txPr, creating the text body and the
# visibility flags if this is the first formatting written here.
function _cx_label_transform(lbl::XML.Node, field::Symbol, value, pfx::Dict{String,String})
    lbl = _cx_visibility_off(lbl, pfx)
    tx  = first_element_with_tag(lbl, "txPr")
    if isnothing(tx)
        lbl = insert_child(lbl, (NS_CX, "dataLabels"), _new_text_body("txPr", pfx, NS_CX))
        tx  = first_element_with_tag(lbl, "txPr")
    end
    return replace_child(lbl, tx, _text_with_run_prop(tx, field, value, pfx))
end

# ---- chart-level text -------------------------------------------------------

# Apply `f(node, pfx)` at the end of `steps` from the part root, and write back.
function _cx_edit_at!(c::ChartEx, steps, f)
    root = _cx_root(c)
    pfx  = ns_prefixes(root)
    new  = rebuild_path(root, steps, n -> f(n, pfx);
                        prefixes = pfx, parent_key = _CX_ROOT_KEY)
    set_chart_root!(c, new)
    return c
end

"""
    setChartTitleTextProp(c::ChartEx, field, value) -> ChartEx

Set one field of the chart title's text formatting, on `cx:title/cx:txPr`.
Creates the title if the chart has none. Returns `c`, which remains valid.

$(_DOC_TEXT_FIELD_ARG)
"""
setChartTitleTextProp(c::ChartEx, field::Symbol, value) =
    _cx_edit_at!(c, [(NS_CX, "chart") => "chart",
                     (NS_CX, "title") => "title"],
                 (t, pfx) -> _both_with_run_prop(t, (NS_CX, "title"), field, value, pfx; ns = NS_CX))

"""
    setLegendTextProp(c::ChartEx, field, value) -> ChartEx

Set one field of the legend's text formatting, on `cx:legend/cx:txPr`.
Creates the legend if the chart has none. Returns `c`, which remains valid.

$(_DOC_TEXT_FIELD_ARG)
"""
setLegendTextProp(c::ChartEx, field::Symbol, value) =
    _cx_edit_at!(c, [(NS_CX, "chart")  => "chart",
                     (NS_CX, "legend") => "legend"],
                 (lg, pfx) -> _txpr_with_run_prop(lg, (NS_CX, "legend"), field, value, pfx; ns = NS_CX))

# cx:legend positions by `pos` and `align`. Top-right is the right side,
# aligned to its start, as Excel writes it.
const _CX_LEGEND_POS = Dict("b" => ("b", "ctr"), "t" => ("t", "ctr"), "l" => ("l", "ctr"),
                            "r" => ("r", "ctr"), "tr" => ("r", "min"))

_cx_legend(c::ChartEx) = first_element_with_tag(_cx_chart(c), "legend")

function getLegendPos(c::ChartEx)
    lg  = _cx_legend(c)
    pos = _cx_attr(lg, "pos")
    isnothing(pos) && return nothing
    return pos == "r" && _cx_attr(lg, "align") == "min" ? :tr : Symbol(pos)
end

getLegendOverlay(c::ChartEx) = _cx_bool(_cx_legend(c), "overlay")

_cx_with_overlay(lg, o) = with_attribute(lg, "overlay", o === :inherit ? nothing : o)

function setLegendPos(c::ChartEx, pos::Symbol; overlay = nothing)
    p = _legend_pos(pos)
    o = isnothing(overlay) ? nothing : _legend_overlay(overlay)
    pa, al = p === :inherit ? (nothing, nothing) : _CX_LEGEND_POS[p]
    return _cx_edit_at!(c, [(NS_CX, "chart")  => "chart",
                            (NS_CX, "legend") => "legend"],
                        (lg, pfx) -> begin
        lg = with_attribute(with_attribute(lg, "pos", pa), "align", al)
        isnothing(o) ? lg : _cx_with_overlay(lg, o)
    end)
end

function setLegendOverlay(c::ChartEx, overlay::Union{Bool,Symbol})
    o = _legend_overlay(overlay)
    return _cx_edit_at!(c, [(NS_CX, "chart")  => "chart",
                            (NS_CX, "legend") => "legend"],
                        (lg, pfx) -> _cx_with_overlay(lg, o))
end

"""
    getChartAxisIds(c::ChartEx) -> Vector{Int}

The `id` of each `cx:axis`, in document order. These are the values
[`getSeriesAxisIds`](@ref) refers to.
"""
getChartAxisIds(c::ChartEx)::Vector{Int} =
    [parse(Int, get_attr(a, "id")) for a in
     elements_with_tag(first_element_with_tag(_cx_chart(c), "plotArea"), "axis")]

# Steps to the cx:axis whose id attribute is `id`.
function _cx_axis_path(c::ChartEx, id::Integer)
    ids = getChartAxisIds(c)
    id in ids || throw(XLSXError(
        "Chart `$(c.name)` has no axis with id $id. Axis ids: $(join(ids, ", "))."))
    return [(NS_CX, "chart")    => "chart",
            (NS_CX, "plotArea") => "plotArea",
            (NS_CX, "axis")     => ("axis", n -> tryparse(Int, get_attr(n, "id")) == id)]
end

"""
    setAxisTitleTextProp(c::ChartEx, id::Integer, field, value) -> ChartEx

Set one field of an axis title's text formatting, on `cx:axis/cx:title/cx:txPr`.
`id` is the axis' `id` attribute, as [`getChartAxisIds`](@ref) reports it, not a
position. Creates the title if the axis has none. Returns `c`, which remains
valid.

$(_DOC_TEXT_FIELD_ARG)
"""
setAxisTitleTextProp(c::ChartEx, id::Integer, field::Symbol, value) =
    _cx_edit_at!(c, [_cx_axis_path(c, id)...; (NS_CX, "title") => "title"],
                 (t, pfx) -> _both_with_run_prop(t, (NS_CX, "title"), field, value, pfx; ns = NS_CX))

"""
    getAxisTitleTextProps(c::ChartEx, id::Integer) -> Union{Nothing, DrawingText}

The text body of the axis title, or `nothing` if the axis has no title.
"""
getAxisTitleTextProps(c::ChartEx, id::Integer) =
    parse_drawing_text(_wb(c), first_element_with_tag(_cx_axis_node(c, id), "title"))

function _cx_axis_node(c::ChartEx, id::Integer)
    pa = first_element_with_tag(_cx_chart(c), "plotArea")
    for a in elements_with_tag(pa, "axis")
        tryparse(Int, get_attr(a, "id")) == id && return a
    end
    throw(XLSXError("Chart `$(c.name)` has no axis with id $id."))
end

# ---- text content -----------------------------------------------------------

# Replace cx:tx with a typed cx:txData/cx:v holding `text`. Any cx:f is dropped:
# typing text unbinds the element from its cell.
function _cx_tx_typed(tx::XML.Node, text::AbstractString, pfx::Dict{String,String})
    v  = XML.Element(prefixed_tag(pfx[NS_CX], "v"), XML.Text(text))
    td = XML.Element(prefixed_tag(pfx[NS_CX], "txData"), v)
    return _with_children(tx, [td])
end

# Excel writes a title's text twice: in cx:txData/cx:v and as runs in the sibling
# cx:txPr, which is what it renders. Rewrite the body's runs, one paragraph per
# line, each copying the first paragraph's pPr and endParaRPr. Every run starts
# from the old first run's a:rPr, so the words change and the appearance does
# not; a RichTextString's runs then lay their own formatting over it. Plain text
# on a title with no cx:txPr leaves it without one; rich text creates it.
function _cx_txpr_retext(el::XML.Node, text::AbstractString, pfx::Dict{String,String})
    tx = first_element_with_tag(el, "txPr")
    if isnothing(tx)
        text isa RichTextString || return el
        tx = _new_text_body("txPr", pfx, NS_CX)
    end
    kids = isnothing(tx.children) ? XML.Node[] : collect(XML.Node, tx.children)
    i    = findfirst(k -> localname(k) == "p", kids)
    p    = isnothing(i) ? XML.Element(prefixed_tag(pfx[NS_A], "p")) : kids[i]
    pk   = isnothing(p.children) ? XML.Node[] : p.children
    ppr  = filter(n -> localname(n) == "pPr", pk)
    epr  = filter(n -> localname(n) == "endParaRPr", pk)
    old  = first_element_with_tag(p, "r")
    base = isnothing(old) ? nothing : first_element_with_tag(old, "rPr")
    paras = map(_cx_lines(text)) do line
        runs = XML.Node[_cx_run(piece, props, base, pfx) for (piece, props) in line]
        _with_children(p, XML.Node[ppr; runs; epr])
    end
    body = XML.Node[k for k in kids if localname(k) != "p"]
    tx   = _with_children(tx, vcat(body, paras))
    return insert_child(el, (NS_CX, String(localname(el))), tx)
end

# `text` as lines of (piece, props) runs. A plain string's runs carry no props.
_cx_lines(text::AbstractString) =
    [isempty(l) ? Tuple{String,Any}[] : Tuple{String,Any}[(String(l), nothing)]
     for l in split(text, _NEWLINE)]
_cx_lines(text::RichTextString) =
    [Tuple{String,Any}[(r.text, r.props) for r in p.runs] for p in _as_paragraphs(text)]

# One a:r: `base` (the old first run's a:rPr, or none) with `props` laid over it.
function _cx_run(piece::AbstractString, props, base, pfx::Dict{String,String})
    rpr = base
    for (field, value) in _run_prop_pairs(props)
        rpr = _rpr_with_prop(something(rpr, XML.Element(prefixed_tag(pfx[NS_A], "rPr"))),
                             field, value, pfx)
    end
    t = XML.Element(prefixed_tag(pfx[NS_A], "t"), XML.Text(piece))
    return XML.Element(prefixed_tag(pfx[NS_A], "r"), (isnothing(rpr) ? (t,) : (rpr, t))...)
end

# A typed cx: title or axis title: cx:txData/cx:v holds the plain text, and
# cx:txPr the runs Excel renders.
function _cx_title_with_text(t::XML.Node, text::AbstractString, pfx)
    t = rebuild_path(t, [(NS_CX, "tx") => "tx"],
                     tx -> _cx_tx_typed(tx, String(text), pfx);
                     prefixes = pfx, parent_key = (NS_CX, "title"))
    return _cx_txpr_retext(t, text, pfx)
end

# A cx: title bound to a cell: cx:txData with the reference and its cached text,
# as Excel writes it. Excel writes no cx:txPr on a bound title; one already there
# loses its runs, so stale typed text cannot render, and keeps their formatting
# in the paragraph default, with whatever a rich-text cell's runs all share
# laid over it — as for a c: title.
function _cx_title_with_ref(t::XML.Node, f::AbstractString, vals::Vector, cellrich,
                            wb::Workbook, pfx)
    v   = only(vals)
    txt = ismissing(v) ? "" : string(v)
    td  = XML.Element(prefixed_tag(pfx[NS_CX], "txData"),
                      XML.Element(prefixed_tag(pfx[NS_CX], "f"), XML.Text(f)),
                      XML.Element(prefixed_tag(pfx[NS_CX], "v"), XML.Text(txt)))
    t   = rebuild_path(t, [(NS_CX, "tx") => "tx"], tx -> _with_children(tx, [td]);
                       prefixes = pfx, parent_key = (NS_CX, "title"))
    key  = (NS_CX, "title")
    txpr = first_element_with_tag(t, "txPr")
    isnothing(first_element_with_tag(txpr, "p")) ||
        (t = insert_child(t, key, _txpr_from_rich(txpr, nothing, wb, pfx; ns = NS_CX)))
    for (field, value) in _common_run_props(cellrich)
        t = _txpr_with_run_prop(t, key, field, value, pfx; ns = NS_CX)
    end
    return t
end

"""
    setChartTitleText(c::ChartEx, text) -> ChartEx
    setChartTitleText(c::ChartEx, ref::Union{CellRef,SheetCellRef}) -> ChartEx

Set the chart title to `text`, typed rather than bound to a cell. Creates the
title if the chart has none, and drops any `cx:f` binding it had. A newline
starts a new paragraph.

Excel stores a typed title both as `cx:tx/cx:txData/cx:v` and as runs in the
title's `cx:txPr`, and renders the latter, so both are updated. Run formatting
is preserved; use [`setChartTitleTextProp`](@ref) to change it. A
`RichTextString`, or a styled string from StyledStrings.jl (Julia 1.11 and
later), lays each run's own formatting over it.

Given a `CellRef` or `SheetCellRef`, the title is bound to that cell
(`cx:txData` with `cx:f` and the cell's text cached in `cx:v`). A bare `CellRef`
means the chart's own sheet; a chart on a chartsheet needs a `SheetCellRef`. A
string is always typed text, never read as a reference.

A bound title can't carry rich text, a limit of Excel's file format. Binding
keeps the title's formatting in `cx:txPr` — its first run's, folded into the
paragraph default — and, where the cell holds a `RichTextString`, lays over it
the formatting all of its runs share.

Returns `c`, which remains valid.
"""
setChartTitleText(c::ChartEx, text::AbstractString) =
    _cx_edit_at!(c, [(NS_CX, "chart") => "chart",
                     (NS_CX, "title") => "title"],
                 (t, pfx) -> _cx_title_with_text(t, text, pfx))

function setChartTitleText(c::ChartEx, ref::Union{CellRef,SheetCellRef})
    f, vals, cellrich = _title_source(c, ref)
    return _cx_edit_at!(c, [(NS_CX, "chart") => "chart",
                            (NS_CX, "title") => "title"],
                        (t, pfx) -> _cx_title_with_ref(t, f, vals, cellrich, _wb(c), pfx))
end

"""
    setAxisTitleText(c::ChartEx, id::Integer, text) -> ChartEx
    setAxisTitleText(c::ChartEx, id::Integer, ref::Union{CellRef,SheetCellRef}) -> ChartEx

Set an axis title's text, or bind it to a cell, as [`setChartTitleText`](@ref)
does for the chart title. `id` is the axis' `id` attribute, as
[`getChartAxisIds`](@ref) reports it, not a position. Creates the title if the
axis has none. Returns `c`, which remains valid.
"""
setAxisTitleText(c::ChartEx, id::Integer, text::AbstractString) =
    _cx_edit_at!(c, [_cx_axis_path(c, id)...; (NS_CX, "title") => "title"],
                 (t, pfx) -> _cx_title_with_text(t, text, pfx))

function setAxisTitleText(c::ChartEx, id::Integer, ref::Union{CellRef,SheetCellRef})
    f, vals, cellrich = _title_source(c, ref)
    return _cx_edit_at!(c, [_cx_axis_path(c, id)...; (NS_CX, "title") => "title"],
                        (t, pfx) -> _cx_title_with_ref(t, f, vals, cellrich, _wb(c), pfx))
end

"""
    setSeriesName(c::ChartEx, i, name::AbstractString; pareto=false) -> ChartEx

Set the name of series `i` to `name`, typed rather than bound to a cell.
Creates the series' `cx:tx` if it has none, and drops any `cx:f` binding.
With `pareto = true`, series `i`'s Pareto line instead; an error if it has none.

Returns `c`, which remains valid.
"""
setSeriesName(c::ChartEx, i::Integer, name::AbstractString; pareto::Bool = false) =
    _cx_edit_series!(c, i, (ser, pfx) ->
        rebuild_path(ser, [(NS_CX, "tx") => "tx"],
                     tx -> _cx_tx_typed(tx, name, pfx);
                     prefixes = pfx, parent_key = (NS_CX, "series"));
        pareto)

# ---- layout properties (write) ----------------------------------------------

# Excel refuses a file with some layout properties on the wrong kind of series
# (e.g. subtotals on a funnel), so their setters check the series' layoutId first.
function _cx_require_layout(c::ChartEx, i::Integer, layout::Symbol, what::AbstractString)
    l = getSeriesLayout(c, i)
    l === layout ||
        throw(XLSXError("$what applies only to a `$layout` series; series $i of chart `$(c.name)` is `$l`."))
    return nothing
end

# Apply `f(layoutPr, pfx)` to series `i`'s cx:layoutPr, creating it if absent.
_cx_set_layoutpr!(c::ChartEx, i::Integer, f) =
    _cx_edit_series!(c, i, (ser, pfx) ->
        rebuild_path(ser, [(NS_CX, "layoutPr") => "layoutPr"], lp -> f(lp, pfx);
                     prefixes = pfx, parent_key = (NS_CX, "series")))

"""
    setSeriesSubtotals(c::ChartEx, i, points) -> ChartEx

Mark the given points of a waterfall series as totals, replacing whatever was
marked before. `points` counts from 1, in any order; an empty collection writes
`<cx:subtotals/>`, which is what Excel writes for a waterfall with no totals.

Pass `:inherit` to remove `cx:subtotals` entirely. Any other value is an error
unless series `i` is a waterfall series (Excel refuses a file with subtotals on
other layouts).

Returns `c`, which remains valid.
"""
function setSeriesSubtotals(c::ChartEx, i::Integer, points)
    if points === :inherit
        return _cx_set_layoutpr!(c, i, (lp, _) -> remove_child(lp, "subtotals"))
    end
    _cx_require_layout(c, i, :waterfall, "Subtotals")
    idx = sort(unique(Int[p for p in points]))
    isempty(idx) || first(idx) >= 1 ||
        throw(XLSXError("Data point positions start at 1; asked for $(first(idx))."))
    return _cx_set_layoutpr!(c, i, (lp, pfx) -> begin
        st = XML.Element(prefixed_tag(pfx[NS_CX], "subtotals"),
                         (XML.Element(prefixed_tag(pfx[NS_CX], "idx"); val = string(p - 1))
                          for p in idx)...)
        return insert_child(lp, (NS_CX, "layoutPr"), st)
    end)
end

# Set or remove one attribute on a single-purpose cx:layoutPr child, creating the
# child if needed. `:inherit` removes the attribute, and the child with it when
# that leaves it empty, since these elements exist only to carry their attributes.
function _cx_set_layout_attr!(c::ChartEx, i::Integer, tag::AbstractString,
                              attr::AbstractString, value)
    return _cx_set_layoutpr!(c, i, (lp, pfx) -> begin
        if value === :inherit
            return remove_child(lp, tag)
        end
        node = something(first_element_with_tag(lp, tag),
                         XML.Element(prefixed_tag(pfx[NS_CX], tag)))
        return insert_child(lp, (NS_CX, "layoutPr"), with_attribute(node, attr, value))
    end)
end

const _CX_QUARTILE_METHODS = (:inclusive, :exclusive)

"""
    setSeriesQuartileMethod(c::ChartEx, i, method) -> ChartEx

Set the quartile calculation of a box & whisker series to `:inclusive` or
`:exclusive`, or `:inherit` to remove `cx:statistics`. Setting a method is an error
unless series `i` is a box & whisker series (Excel refuses a file with one on
other layouts).
"""
function setSeriesQuartileMethod(c::ChartEx, i::Integer, method::Symbol)
    method === :inherit || method in _CX_QUARTILE_METHODS ||
        throw(XLSXError("`$method` is not a quartile method. Use " *
                        join(_CX_QUARTILE_METHODS, " or ") * "."))
    method === :inherit || _cx_require_layout(c, i, :boxWhisker, "A quartile method")
    return _cx_set_layout_attr!(c, i, "statistics", "quartileMethod",
                                method === :inherit ? :inherit : String(method))
end

const _CX_PARENT_LABEL_LAYOUTS = (:none, :banner, :overlapping)

"""
    setSeriesParentLabelLayout(c::ChartEx, i, layout) -> ChartEx

Set a treemap series' parent label placement to `:none`, `:banner` or
`:overlapping`, or `:inherit` to remove `cx:parentLabelLayout`.
"""
function setSeriesParentLabelLayout(c::ChartEx, i::Integer, layout::Symbol)
    layout === :inherit || layout in _CX_PARENT_LABEL_LAYOUTS ||
        throw(XLSXError("`$layout` is not a parent label layout. Valid values: " *
                        join(_CX_PARENT_LABEL_LAYOUTS, ", ") * "."))
    return _cx_set_layout_attr!(c, i, "parentLabelLayout", "val",
                                layout === :inherit ? :inherit : String(layout))
end

"""
    setSeriesLayoutFlag(c::ChartEx, i, flag, value) -> ChartEx

Set one attribute of `cx:layoutPr/cx:visibility`: `:meanLine`, `:meanMarker`,
`:nonoutliers`, `:outliers` (box & whisker) or `:connectorLines` (waterfall).
`value` is a `Bool`, or `:inherit` to remove the attribute.
"""
function setSeriesLayoutFlag(c::ChartEx, i::Integer, flag::Symbol, value)
    flag in _CX_LAYOUT_FLAGS ||
        throw(XLSXError("Unknown layout flag `$flag`. Expected one of $(join(_CX_LAYOUT_FLAGS, ", "))."))
    value isa Bool || value === :inherit ||
        throw(XLSXError("A layout flag takes a Bool or `:inherit`; got `$value`."))
    return _cx_set_layoutpr!(c, i, (lp, pfx) -> begin
        vis = something(first_element_with_tag(lp, "visibility"),
                        XML.Element(prefixed_tag(pfx[NS_CX], "visibility")))
        vis = with_attribute(vis, String(flag), value === :inherit ? nothing : (value ? "1" : "0"))
        return isnothing(vis.attributes) || isempty(vis.attributes) ?
               remove_child(lp, "visibility") :
               insert_child(lp, (NS_CX, "layoutPr"), vis)
    end)
end

"""
    setSeriesAggregation(c::ChartEx, i, on::Bool) -> ChartEx

Turn category aggregation on or off for series `i`. `cx:aggregation` carries no
value, so its presence is the setting. It excludes `cx:binning`, which is
removed when this is turned on.
"""
setSeriesAggregation(c::ChartEx, i::Integer, on::Bool) =
    _cx_set_layoutpr!(c, i, (lp, pfx) -> on ?
        insert_child(remove_choice(lp, (NS_CX, "layoutPr"), AGGREGATION_GROUP),
                     (NS_CX, "layoutPr"),
                     XML.Element(prefixed_tag(pfx[NS_CX], "aggregation"))) :
        remove_child(lp, "aggregation"))

"""
    setSeriesBinning(c::ChartEx, i; intervalClosed, underflow, overflow, binSize, binCount) -> ChartEx

Set histogram binning for series `i` in one rebuild. A keyword left unspecified
is left alone; pass `:inherit` to remove one that is set.

- `intervalClosed` — `:r` (right-closed, Excel's default) or `:l` (left-closed).
- `underflow`, `overflow` — a `Real`: values below or above it are gathered in
  one end bin. `:auto` lets Excel choose.
- `binSize` — the width of each bin, a positive `Real`; or `:auto`.
- `binCount` — the number of bins, a positive `Integer`; or `:auto`.

`binSize` and `binCount` are alternatives: setting one removes the other.
`cx:binning` excludes `cx:aggregation`, which is removed if present.
"""
function setSeriesBinning(c::ChartEx, i::Integer;
                          intervalClosed = nothing, underflow = nothing, overflow = nothing,
                          binSize = nothing, binCount = nothing)
    all(isnothing, (intervalClosed, underflow, overflow, binSize, binCount)) && return c
    isnothing(binSize) || isnothing(binCount) ||
        throw(XLSXError("`binSize` and `binCount` are alternatives; set only one."))
    isnothing(intervalClosed) || intervalClosed === :inherit || intervalClosed in (:r, :l) ||
        throw(XLSXError("`intervalClosed` is `:r` or `:l`; got `$intervalClosed`."))

    fmt(v) = v === :auto ? "auto" : string(v)

    return _cx_set_layoutpr!(c, i, (lp, pfx) -> begin
        lp = remove_child(lp, "aggregation")
        b  = something(first_element_with_tag(lp, "binning"),
                       XML.Element(prefixed_tag(pfx[NS_CX], "binning")))
        isnothing(intervalClosed) ||
            (b = with_attribute(b, "intervalClosed",
                                intervalClosed === :inherit ? nothing : String(intervalClosed)))
        isnothing(underflow) ||
            (b = with_attribute(b, "underflow", underflow === :inherit ? nothing : fmt(underflow)))
        isnothing(overflow) ||
            (b = with_attribute(b, "overflow", overflow === :inherit ? nothing : fmt(overflow)))
        for (kw, tag) in ((binSize, "binSize"), (binCount, "binCount"))
            isnothing(kw) && continue
            b = remove_choice(b, (NS_CX, "binning"), BIN_GROUP)
            kw === :inherit && continue
            b = insert_child(b, (NS_CX, "binning"),
                             XML.Element(prefixed_tag(pfx[NS_CX], tag); val = fmt(kw)))
        end
        return insert_child(lp, (NS_CX, "layoutPr"), b)
    end)
end