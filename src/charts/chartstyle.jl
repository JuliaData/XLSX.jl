# The chart style part as the bottom rung of a cx: chart's fill cascade.
#
# Excel's cx: charts write no `spPr` on a series it has not been asked to
# format, so the colour it draws comes from the chart's style part
# (`xl/charts/styleN.xml`, `cs:chartStyle`) and colours part (`colorsN.xml`,
# `cs:colorStyle`). This file reads them to answer `getSeriesFill(::ChartEx, …)`
# where the chart part is silent. What follows was checked against SVG exports
# of Excel 365 charts (`test/data/chart_style_svg`), colour for colour.
#
# Three steps give a fill:
#
#   1. The object number k (0-based). What counts as an object depends on the
#      layout: a series for box & whisker, funnel, histogram and Pareto bars (its
#      position among the chart's `cx:series`); for a waterfall, increase 0,
#      decrease 1, total 2; for a treemap or sunburst, the point's top-level
#      category, numbered in order of first appearance, children taking their
#      parent's colour.
#   2. The colour for k, from the colours part. Only `meth="cycle"` is read:
#      colour `k mod n`, with variation `k ÷ n` appended as transforms. Verified
#      up to k = 7 (variation 1); beyond that the part's own list is followed.
#      The linear methods (`withinLinear` and the rest, Excel's monochromatic
#      palettes) give shades Excel does not compute by any formula we found, so
#      under them there is no answer rather than an approximate one.
#   3. The fill, from the style part's `cs:dataPoint` entry: its own `cs:spPr`
#      fill if it has one, otherwise the theme fill style its `fillRef idx`
#      names. Either way the colour replaces the `phClr` placeholder, keeping
#      any transforms written on the placeholder.
#
# Lines and text are deliberately not resolved this way: Excel draws box &
# whisker outlines darker than the style says, and picks a contrasting colour
# for labels inside funnel bars, so the style part would misreport them.
#
# `c:` charts have no style rung. A `c:` series with no `spPr` is drawn from the
# legacy `c:style` number, not from these parts: replacing the colours part
# changes nothing, and Excel drops both parts when it saves such a chart.
#
# Every step answers `nothing` when it meets a case the fixtures do not cover
# (an unknown layout, a series carrying `formatIdx`, a zero or non-numeric
# waterfall value, a blank top-level category, a numbered `cs:styleClr`), and
# the cascade then ends where it did before this rung existed.

# Layouts whose series take one colour, numbered by series position.
const _CX_STYLE_BY_SERIES = ("clusteredColumn", "funnel", "boxWhisker")

# The chart's style and colours parts, as root elements, or `nothing` for either
# that the chart does not have.
function _style_parts(c::AbstractChart)
    style = colors = nothing
    for (_, (path, rt)) in rid_to_target(c.package, c.path, (REL_CHART_STYLE, REL_CHART_COLORS))
        root = xml_root_element(get_xml_data(c.package, path))
        rt == REL_CHART_STYLE ? (style = root) : (colors = root)
    end
    return style, colors
end

# The cells of a data dimension, flattened in row order, or `nothing` where the
# dimension is not a plain range on an existing sheet. Columns are kept: a
# treemap's categories have one column per level, the leftmost the top.
function _cx_dimension_cells(c::ChartEx, i::Integer, kind::Symbol, type::Symbol)
    data = getSeriesData(c, i)
    isnothing(data) && return nothing
    k = findfirst(d -> d.kind === kind && d.type === type, data.dimensions)
    isnothing(k) && return nothing
    r = data.dimensions[k].range
    r isa SheetCellRange && hassheet(_wb(c), r.sheet) || return nothing
    v = getdata(c.package, r)
    return v isa AbstractMatrix ? v : reshape([v], 1, 1)
end

# k for a waterfall point: total 2, increase 0, decrease 1.
function _waterfall_index(c::ChartEx, i::Integer, point::Integer)
    point in something(getSeriesSubtotals(c, i), Int[]) && return 2
    vals = _cx_dimension_cells(c, i, :num, :val)
    (isnothing(vals) || !(1 <= point <= size(vals, 1))) && return nothing
    v = vals[point, 1]
    v isa Real || return nothing
    return v > 0 ? 0 : v < 0 ? 1 : nothing
end

# k for a treemap or sunburst point: its top-level category, numbered in order of
# first appearance. A blank top-level cell anywhere leaves the numbering unknown.
function _hierarchy_index(c::ChartEx, i::Integer, point::Integer)
    cats = _cx_dimension_cells(c, i, :str, :cat)
    (isnothing(cats) || !(1 <= point <= size(cats, 1))) && return nothing
    tops = cats[:, 1]
    any(t -> ismissing(t) || t == "", tops) && return nothing
    return findfirst(==(tops[point]), unique(tops)) - 1
end

# The object number k of series `i` (or one of its points), or `nothing`.
function _cx_style_index(c::ChartEx, i::Integer, point, pareto::Bool)
    root = _cx_root(c)
    ser  = _cx_series_node(c, root, i; pareto)
    isnothing(_cx_attr(ser, "formatIdx")) || return nothing
    layout = _cx_attr(ser, "layoutId")
    layout in _CX_STYLE_BY_SERIES &&
        return findfirst(s -> s === ser, _cx_series_nodes(root)) - 1
    isnothing(point) && return nothing                # the series has no single colour
    layout == "waterfall" && return _waterfall_index(c, i, point)
    layout in ("treemap", "sunburst") && return _hierarchy_index(c, i, point)
    return nothing
end

_elements(n::XML.Node) = XML.Node[k for k in XML.eachelement(n)]

# The colour element for object `k` under a `cycle` colours part: colour k mod n
# with the transforms of variation k ÷ n appended. `nothing` under any other method.
function _cycle_color(colors::XML.Node, k::Integer)::Union{Nothing,XML.Node}
    get_attr(colors, "meth") == "cycle" || return nothing
    clrs = [e for e in XML.eachelement(colors) if localname(e) in DML_COLOR_TAGS]
    vars = [e for e in XML.eachelement(colors) if localname(e) == "variation"]
    isempty(clrs) && return nothing
    base = clrs[k % length(clrs) + 1]
    isempty(vars) && return base
    return _with_children(base, vcat(_elements(base), _elements(vars[(k ÷ length(clrs)) % length(vars) + 1])))
end

# The colour a style reference (`cs:fillRef`) supplies for `phClr`: the computed
# colour for `cs:styleClr val="auto"`, its own DrawingML colour if it names one,
# `nothing` if it names none, and `:unknown` for a numbered `cs:styleClr`.
function _ref_color(ref::Union{Nothing,XML.Node}, auto::XML.Node)
    isnothing(ref) && return nothing
    sc = first_element_with_tag(ref, "styleClr")
    isnothing(sc) || return get_attr(sc, "val") == "auto" ? auto : :unknown
    return color_element(ref)
end

# `fill` with every `phClr` replaced by `clr`, the placeholder's own transforms
# applied after the colour's.
_with_placeholder(fill::XML.Node, clr::XML.Node) =
    _rewrite(fill) do e
        localname(e) == "schemeClr" && get_attr(e, "val") == "phClr" || return e
        return _with_children(clr, vcat(_elements(clr), _elements(e)))
    end

_has_placeholder(n::XML.Node) =
    (localname(n) == "schemeClr" && get_attr(n, "val") == "phClr") ||
    any(_has_placeholder, XML.eachelement(n))

# The style rung for the fill of series `i` (or one of its points): a FormatSite
# whose container is the style's `cs:dataPoint` entry and whose props are an
# `a:spPr` holding the resolved fill. `nothing` where the style gives no answer.
function _cx_style_fill_site(c::ChartEx, i::Integer; point, pareto::Bool)::Union{Nothing,FormatSite}
    style, colors = _style_parts(c)
    (isnothing(style) || isnothing(colors)) && return nothing
    k = _cx_style_index(c, i, point, pareto)
    isnothing(k) && return nothing
    clr = _cycle_color(colors, k)
    isnothing(clr) && return nothing

    entry = first_element_with_tag(style, "dataPoint")
    isnothing(entry) && return nothing
    ref    = first_element_with_tag(entry, "fillRef")
    refclr = _ref_color(ref, clr)
    refclr === :unknown && return nothing
    fill = fill_element(first_element_with_tag(entry, "spPr"))
    if isnothing(fill)
        idx  = isnothing(ref) ? nothing : tryparse(Int, get_attr(ref, "idx"))
        fill = theme_fill_style(_wb(c), idx)
        isnothing(fill) && return nothing
    end
    if isnothing(refclr)
        _has_placeholder(fill) && return nothing      # a placeholder with nothing to fill it
    else
        fill = _with_placeholder(fill, refclr)
    end
    return FormatSite(:style, :shape, entry, XML.Element("a:spPr", fill))
end
