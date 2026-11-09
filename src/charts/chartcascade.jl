# Cascade resolution for `c:` charts.
#
# There are two kinds of accessor. Most report what one node says, and
# `nothing` means that node is silent, not that a default applies. The
# cascade resolvers — `getSeriesFill`, `getSeriesLine`, `getMarkerFill`,
# `getLabelTextProp` — walk the rungs instead and return an `Effective`
# carrying the value, the rung that answered and the whole chain.
#
# For a `c:` chart the cascade stops at the chart XML. Where a property is
# written at no rung, `Effective.value` is `nothing`, and Excel draws it from
# the legacy `c:style` number, not from `xl/charts/style1.xml` and
# `colors1.xml`: a `c:` series with no `spPr` looks the same whatever the
# colours part says, and Excel drops both parts when it saves such a chart. So
# there is no `:style` rung here. The `cx:` charts do draw from those parts, and
# chartstyle.jl adds their fill rung.

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
describe their own backgrounds rather than a series, so the chain stops there.
Excel draws an unformatted series from the chart's `c:style` number.
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

`field` is one of $(_doc_syms(_RUN_PROP_FIELDS)) — the DrawingML names in the
table at [`setChartTitleTextProp`](@ref); unlike the setters, this does not take
Excel's names. Text inherits field by
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

