# Appearance accessors for `c:` charts: the design notes, and the node locators
# and value helpers that every other chart*.jl file shares.
#
# The appearance accessors are split by object: chartnodes.jl (this file),
# charttext.jl (titles, label text, text properties), chartseries.jl (series,
# groups, data points, markers, trendlines, error bars), chartaxes.jl,
# chartlegend.jl (legend, plot area, chart space) and chartcascade.jl (the
# cascade resolvers). Together they sit above `drawingml.jl` in the layering:
# that file parses DrawingML — `spPr`, `txPr`, colors, fills, lines, text — for
# anything that carries it, knowing nothing about charts. These files know the
# chart schema and find the nodes: which element holds a series' fill, which
# axes a group plots against, where a per-point override lives. Nothing here
# parses DrawingML itself; it locates a node and hands it to `drawingml.jl`.
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

_wb(c::AbstractChart) = get_workbook(c.package)

_chartnode(c::Chart) = first_element_with_tag(chart_root(c), "chart")

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

function _check_pct(v, lo, hi, what)
    v === :inherit && return :inherit
    v isa Integer && lo <= v <= hi && return string(v)
    throw(XLSXError("$what must be an integer percentage from $lo to $hi, or `:inherit`; got $(repr(v))."))
end

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

function _positive(v, what)
    v === :inherit && return :inherit
    x = Float64(v)
    isfinite(x) && x > 0 || throw(XLSXError("$what must be positive and finite; got $v."))
    return _xsd_double(x)
end

