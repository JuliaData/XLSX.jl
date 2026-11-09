# Chart-level elements of `c:` charts: the legend (position, overlay, entries),
# the plot area and the chart space. See chartnodes.jl for the design notes.

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
    getLegendPos(c::ChartEx) -> Union{Nothing,Symbol}

`:b` (bottom), `:t` (top), `:l` (left), `:r` (right) or `:tr` (top-right) from
`c:legendPos`, or from `cx:legend`'s `pos` and `align`. `nothing` means no legend,
or a legend with no explicit position. Set with [`setLegendPos`](@ref).
"""
function getLegendPos(c::Chart)
    v = _attr(first_element_with_tag(getChartLegend(c), "legendPos"), "val")
    return isnothing(v) ? nothing : Symbol(v)
end

"""
    getLegendOverlay(c::Chart) -> Union{Nothing,Bool}
    getLegendOverlay(c::ChartEx) -> Union{Nothing,Bool}

`c:overlay`, or `cx:legend`'s `overlay` attribute — `true` where the legend is drawn on top of the plot area rather
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

# ── Legend position ───────────────────────────────────────────────────────────

const _LEGEND_POS = (:b, :t, :l, :r, :tr)

const _LEGEND_POS_ALIASES = Dict(:bottom => :b, :top => :t, :left => :l, :right => :r,
                                 :topRight => :tr)

_legend_pos(p::Symbol) = _check_sym(get(_LEGEND_POS_ALIASES, p, p), _LEGEND_POS, "legend position")

function _legend_overlay(v)
    v isa Bool && return v ? "1" : "0"
    v === :inherit && return :inherit
    throw(XLSXError("`$(repr(v))` is not a legend overlay; use `true`, `false` or `:inherit`."))
end

# Apply `f(lg, key, pfx)` to `c:legend`, creating it if absent, in one rebuild.
function _set_legend(c::Chart, f)
    root = chart_root(c)
    pfx  = ns_prefixes(root)
    lkey = (NS_C, "legend")
    new  = rebuild_path(root, [(NS_C, "chart") => "chart", lkey => "legend"],
                        lg -> f(lg, lkey, pfx); prefixes = pfx)
    set_chart_root!(c, new)
    return c
end

"""
    setLegendPos(c::Chart, pos; overlay=nothing) -> Chart
    setLegendPos(c::ChartEx, pos; overlay=nothing) -> ChartEx

Set where the legend is drawn: `:b` (bottom), `:t` (top), `:l` (left), `:r`
(right) or `:tr` (top-right) — Excel's five legend positions. `:bottom`, `:top`,
`:left`, `:right` and `:topRight` are accepted too. `:inherit` removes the
position, which Excel reads as `:r`.

`overlay` also sets [`setLegendOverlay`](@ref): `true` draws the legend over the
plot area, `false` beside it, shrinking the plot area to make room. `nothing`
leaves it as it is.

Creates the legend if the chart has none. On a `Chart`, any manual layout on the
legend (`c:legend/c:layout`, written when a legend is dragged in Excel) is
removed so that the position takes effect, as picking a position in Excel does.
"""
function setLegendPos(c::Chart, pos::Symbol; overlay = nothing)
    p = _legend_pos(pos)
    o = isnothing(overlay) ? nothing : _legend_overlay(overlay)
    return _set_legend(c, (lg, key, pfx) -> begin
        lg = _with_val(remove_child(lg, "layout"), key, pfx, "legendPos", p)
        isnothing(o) ? lg : _with_val(lg, key, pfx, "overlay", o)
    end)
end

"""
    setLegendOverlay(c::Chart, overlay) -> Chart
    setLegendOverlay(c::ChartEx, overlay) -> ChartEx

`true` draws the legend over the plot area; `false` draws it beside the plot
area, which shrinks to make room — Excel's "Show the legend without overlapping
the chart". `:inherit` removes the setting, which Excel reads as `false`.
Creates the legend if the chart has none.
"""
function setLegendOverlay(c::Chart, overlay::Union{Bool,Symbol})
    o = _legend_overlay(overlay)
    return _set_legend(c, (lg, key, pfx) -> _with_val(lg, key, pfx, "overlay", o))
end

_plotarea_steps(_) = [(NS_C, "chart") => "chart", (NS_C, "plotArea") => "plotArea"]

"""
    setPlotAreaLine(c; color, width, dash, cap, compound, join, miterLimit) -> Chart
    setPlotAreaLine(c, :none)
    setPlotAreaLine(c, :inherit)

Set the border of the plot area (`c:plotArea/c:spPr/a:ln`), as
[`setSeriesLine`](@ref) sets a series outline: several properties in one
rebuild, `:none` for a border that draws nothing, `:inherit` to remove it so the
chart style supplies it.

$(_DOC_LINE_KWS) A keyword left unspecified is left alone.
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

