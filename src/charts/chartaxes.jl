# Axes of `c:` charts: lookup, titles, gridlines, scaling, ticks, number format,
# crossing and the axis line. See chartnodes.jl for the design notes.

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
    getAxisTitleRef(c::ChartEx, id::Integer) -> Union{Nothing,String}

The formula behind an axis title that references a cell, as written
(`c:title/c:tx/c:strRef/c:f`, or `cx:txData/cx:f`). `nothing` for a literal or
absent title. Excel writes a `cx:` formula as a hidden `_xlchart.*` defined
name; [`getAxisTitleRange`](@ref) resolves it.
"""
function getAxisTitleRef(c::Chart, ax::ChartAxis)
    tx = first_element_with_tag(first_element_with_tag(_axnode(c, ax), "title"), "tx")
    sr = first_element_with_tag(tx, "strRef")
    return isnothing(sr) ? nothing : child_text(sr, "f")
end

"""
    getAxisTitleRange(c::Chart, ax::ChartAxis) -> ChartRange
    getAxisTitleRange(c::ChartEx, id::Integer) -> ChartRange

The cell an axis title is bound to, resolved — through a defined name where the
formula is one. `nothing` for a typed or absent title. See
[`getAxisTitleRef`](@ref) for the formula as written.
"""
getAxisTitleRange(c::Chart, ax::ChartAxis) = _title_range(c.package, getAxisTitleRef(c, ax))

"""
    getAxisTitle(c::Chart, ax::ChartAxis) -> Union{Nothing,String}
    getAxisTitle(c::ChartEx, id::Integer) -> Union{Nothing,String}

An axis title's text as plain text, as [`getChartTitle`](@ref) gives the chart
title's: the typed text, or for a title bound to a cell, the text cached when
the file was saved. `nothing` where the axis has no title text.
"""
getAxisTitle(c::Chart, ax::ChartAxis) = parse_chart_title(_axnode(c, ax))

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

"""
    getAxisTitleTextProps(c::Chart, ax::ChartAxis) -> Union{Nothing,DrawingText}

Text body for the axis title (`c:title/c:txPr`) — its formatting, which applies
whether the title is typed or bound to a cell, and is the only formatting a
bound title has. `nothing` means the axis has no title, or the title has no
`c:txPr`. Contrast [`getAxisTitleText`](@ref), the literal text in
`c:tx/c:rich`, and [`getAxisTextProps`](@ref), the tick labels' formatting.
"""
getAxisTitleTextProps(c::Chart, ax::ChartAxis) =
    parse_drawing_text(_wb(c), first_element_with_tag(_axnode(c, ax), "title"))

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
- `min`, `max` — fixed bounds, as finite `Real`s in the axis's own units (on a
  date axis, Excel date serial numbers). `:inherit` returns a bound to automatic.

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

Fix the axis's lower bound (`c:scaling/c:min`) at `value`, a finite `Real` in
the axis's own units; `:inherit` returns it to automatic. See
[`setAxisScaling`](@ref) to set both bounds at once.
"""
setAxisMin(c::Chart, ax::ChartAxis, v) = setAxisScaling(c, ax; min = v)

"""
    setAxisMax(c, ax, value) -> Chart

Fix the axis's upper bound (`c:scaling/c:max`) at `value`, a finite `Real` in
the axis's own units; `:inherit` returns it to automatic. See
[`setAxisScaling`](@ref) to set both bounds at once.
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
(`c:crossesAt`) — `value` is a finite `Real`. Removes any `c:crosses`.
`:inherit` removes `c:crossesAt` only.
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
`code` is an Excel number format string, as for a cell: `"0.0%"`, `"#,##0"`,
`"\$#,##0.00"`, `"mmm-yy"`. `:inherit` removes `c:numFmt`.

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

Set one text property of the axis's tick labels (`c:txPr` on the axis).

$(_DOC_TEXT_FIELD_ARG)
"""
setAxisTextProp(c::Chart, ax::ChartAxis, field::Symbol, value) =
    _set_axis(c, ax, (el, key, pfx) -> _txpr_with_run_prop(el, key, field, value, pfx))

_axis_steps(c::Chart, ax::ChartAxis) = root -> first(_axis_path(c, root, ax))

"""
    setAxisLine(c, ax; color, width, dash, cap, compound, join, miterLimit) -> Chart
    setAxisLine(c, ax, :none)
    setAxisLine(c, ax, :inherit)

Set the axis line, as [`setSeriesLine`](@ref) sets a series outline: several
properties in one rebuild, `:none` for a line that draws nothing, `:inherit` to
remove it so the chart style supplies it.

$(_DOC_LINE_KWS) A keyword left unspecified is left alone.

# Example

    valax = first(getChartAxes(c, :value))
    setAxisLine(c, valax; color = "gray", width = 0.75, dash = :sysDash)
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
    setAxisMajorUnit(c, ax, value) -> Chart

Set the interval between major ticks and gridlines (`c:majorUnit`) to `value`,
a positive `Real` in the axis's own units; `:inherit` returns it to automatic.
Value and date axes only.

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

