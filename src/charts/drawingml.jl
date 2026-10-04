#
# charts/drawingml.jl
#
# Shared DrawingML parsing: the `a:` namespace elements that appear throughout
# chart parts, drawings and themes.
#
# Colours first. Every fill, line and text run in DrawingML resolves to one of
# six colour elements, optionally modified by transforms. Parsing this once and
# reusing it is what keeps the chart formatting work tractable: `spPr` alone
# appears on the chart space, plot area, every series, every data point, every
# axis, the legend, the title, data labels, gridlines, trendlines and error
# bars.
#
# It lives in Charts because nothing outside the chart code uses it yet; if
# images or other drawing objects need it, it can move to the core or become a
# sub-module of its own.

# The six colour elements. All carry a `val` attribute except `scrgbClr`, which
# uses separate r/g/b percentage attributes.
const DML_COLOR_TAGS = ("srgbClr", "schemeClr", "sysClr", "prstClr", "hslClr", "scrgbClr")

const DML_COLOR_KIND = Dict(
    "srgbClr"  => :srgb,
    "schemeClr"=> :scheme,
    "sysClr"   => :sys,
    "prstClr"  => :prst,
    "hslClr"   => :hsl,
    "scrgbClr" => :scrgb,
)

# Transforms we resolve. Others (hueMod, red, green, blue, gamma, inv, gray,
# comp) are recorded but not applied: they are vanishingly rare in chart parts,
# and applying them wrongly would be worse than leaving the base colour.
const DML_APPLIED_TRANSFORMS = (:lumMod, :lumOff, :satMod, :shade, :tint, :alpha)

"""
The DrawingML colour element among the children of `node`, or `nothing`.

A fill, line or run-property element holds at most one colour child, so this
returns the first match rather than a list.
"""
function color_element(node::Union{Nothing,XML.Node})::Union{Nothing,XML.Node}
    isnothing(node) && return nothing
    for child in XML.eachelement(node)
        localname(child) in DML_COLOR_TAGS && return child
    end
    return nothing
end

"""
Resolve a colour reference to `"RRGGBB"`, before any transforms are applied.

`schemeClr` goes through the workbook theme; `sysClr` prefers the `lastClr`
attribute Excel writes alongside the system colour name, since the name itself
depends on OS settings; `prstClr` goes through the Colors.jl named colours.
"""
function resolve_color_base(wb::Workbook, node::XML.Node)::String
    tag = localname(node)
    val = get_attr(node, "val")

    if tag == "srgbClr"
        return uppercase(val)

    elseif tag == "schemeClr"
        m = get_theme_color_map(wb)
        haskey(m, val) && return uppercase(m[val])
        # phClr is a placeholder resolved by the style that references it, and
        # has no meaning on its own. Anything else is an unknown scheme name.
        return "000000"

    elseif tag == "sysClr"
        last = get_attr(node, "lastClr")
        isempty(last) || return uppercase(last)
        return val == "window" ? "FFFFFF" : "000000"

    elseif tag == "prstClr"
        # DrawingML preset names are camelCase X11, except that dark/light/medium
        # are abbreviated, which Colors.jl does not know.
        s = lowercase(val)
        s = startswith(s, "dk")  ? "dark"   * s[3:end] :
            startswith(s, "lt")  ? "light"  * s[3:end] :
            startswith(s, "med") ? "medium" * s[4:end] : s
        c = get_colorant(replace(s, "grey" => "gray"))
        return isnothing(c) ? "000000" : uppercase(c[3:end])

    elseif tag == "hslClr"
        h = _attr_pct(node, "hue", 0.0) * 360.0 / 100.0   # hue is in 1/60000 degree
        hd = tryparse(Int, get_attr(node, "hue"))
        hue = isnothing(hd) ? 0.0 : hd / 60_000.0
        s = _attr_pct(node, "sat", 0.0)
        l = _attr_pct(node, "lum", 0.0)
        c = convert(Colors.RGB{Float64}, Colors.HSL{Float64}(hue, s, l))
        return Colors.hex(c, :RRGGBB)

    elseif tag == "scrgbClr"
        # Components are linear, in thousandths of a percent, and may lie outside
        # 0–100% (scRGB is an extended-range space); clamp to what sRGB can show.
        lin(name) = _linear_to_srgb(clamp(_attr_pct(node, name, 0.0), 0.0, 1.0))
        c = Colors.RGB{Float64}(lin("r"), lin("g"), lin("b"))
        return Colors.hex(c, :RRGGBB)
    end

    return "000000"
end

# Percentage attributes are in thousandths of a percent, as elsewhere in DrawingML.
@inline function _attr_pct(node::XML.Node, name::String, default::Float64)::Float64
    s = get_attr(node, name)
    isempty(s) && return default
    v = tryparse(Int, s)
    return isnothing(v) ? default : v / 100_000.0
end


"""
    parse_drawing_color(wb, node) -> Union{Nothing,DrawingColor}

Parse the DrawingML colour held by `node` - a fill, a line, a run property -
into a [`DrawingColor`](@ref), or `nothing` where it holds none.

Pass the colour element itself, or its parent: a parent is searched for its
colour child.
"""
function parse_drawing_color(wb::Workbook, node::Union{Nothing,XML.Node})::Union{Nothing,DrawingColor}
    isnothing(node) && return nothing
    el = localname(node) in DML_COLOR_TAGS ? node : color_element(node)
    isnothing(el) && return nothing

    kind = DML_COLOR_KIND[localname(el)]
    val  = get_attr(el, "val")
    base = resolve_color_base(wb, el)

    transforms = Pair{Symbol,Int}[]
    for child in XML.eachelement(el)
        v = tryparse(Int, get_attr(child, "val"))
        isnothing(v) && continue
        push!(transforms, Symbol(localname(child)) => v)
    end

    rgb, alpha = apply_drawingml_transforms(base, transforms)
    return DrawingColor(kind, val, transforms, rgb, alpha)
end

Base.show(io::IO, c::DrawingColor) =
    print(io, "XLSX.Charts.DrawingColor(", c.val, " -> #", c.rgb,
          c.alpha == 1.0 ? "" : ", alpha " * string(round(c.alpha; digits=3)), ")")

function Base.show(io::IO, ::MIME"text/plain", c::DrawingColor)
    println(io, "XLSX.Charts.DrawingColor ", c.kind, " ", repr(c.val))
    println(io, "  resolves to: #", c.rgb)
    c.alpha == 1.0 || println(io, "  alpha: ", round(c.alpha; digits=3))
    isempty(c.transforms) ||
        println(io, "  transforms: ",
                join(("$k=$(v/1000)%" for (k, v) in c.transforms), ", "))
end

const DML_FILL_TAGS = Dict(
    "noFill"   => :none,
    "solidFill"=> :solid,
    "gradFill" => :gradient,
    "pattFill" => :pattern,
    "blipFill" => :blip,
    "grpFill"  => :group,
)

"""
The fill element among the children of `node`, or `nothing` where there is
none. An `spPr` holds at most one.
"""
function fill_element(node::Union{Nothing,XML.Node})::Union{Nothing,XML.Node}
    isnothing(node) && return nothing
    for child in XML.eachelement(node)
        haskey(DML_FILL_TAGS, localname(child)) && return child
    end
    return nothing
end

"""
    parse_drawing_fill(wb, node) -> Union{Nothing,DrawingFill}

Parse the fill held by `node`, or `nothing` where it holds none. Pass the fill
element itself or its parent.
"""
function parse_drawing_fill(wb::Workbook, node::Union{Nothing,XML.Node})::Union{Nothing,DrawingFill}
    isnothing(node) && return nothing
    el = haskey(DML_FILL_TAGS, localname(node)) ? node : fill_element(node)
    isnothing(el) && return nothing
    kind = DML_FILL_TAGS[localname(el)]

    if kind === :solid
        return DrawingFill(kind, parse_drawing_color(wb, el), nothing, nothing, el)

    elseif kind === :pattern
        fg = parse_drawing_color(wb, first_element_with_tag(el, "fgClr"))
        bg = parse_drawing_color(wb, first_element_with_tag(el, "bgClr"))
        prst = get_attr(el, "prst")
        return DrawingFill(kind, fg, bg, isempty(prst) ? nothing : prst, el)
    end

    return DrawingFill(kind, nothing, nothing, nothing, el)
end


"""
    parse_drawing_line(wb, node) -> Union{Nothing,DrawingLine}

Parse the outline held by `node`, or `nothing` where it holds none. Pass the
`a:ln` element itself or its parent.
"""
function parse_drawing_line(wb::Workbook, node::Union{Nothing,XML.Node})::Union{Nothing,DrawingLine}
    isnothing(node) && return nothing
    el = has_localname(node, "ln") ? node : first_element_with_tag(node, "ln")
    isnothing(el) && return nothing

    dash = first_element_with_tag(el, "prstDash")

    # The join is whichever of the three is present; only miter carries a limit.
    join = first_element_with_tag(el, "round")
    isnothing(join) && (join = first_element_with_tag(el, "bevel"))
    isnothing(join) && (join = first_element_with_tag(el, "miter"))

    return DrawingLine(
        parse_drawing_fill(wb, el),
        _attr_emu(el, "w"),          # points, converted at parse time
        _attr(dash, "val"),
        _attr(el, "cap"),
        _attr(el, "cmpd"),
        isnothing(join) ? nothing : String(localname(join)),
        _attr_pct_opt(join, "lim"),  # nothing for round and bevel
        el,
    )
end

# A fill read from a file carries a DrawingColor with `rgb` resolved; one built
# by hand may carry a SchemeColor or a string, which has no resolved value until
# it is written and read back.
_show_color(c::DrawingColor) = "#" * c.rgb
_show_color(c::SchemeColor)  = string(c.token)
_show_color(c::AbstractString) = String(c)

Base.show(io::IO, fl::DrawingFill) =
    print(io, "XLSX.Charts.DrawingFill(", fl.kind,
          isnothing(fl.fgcolor) ? "" : ", " * _show_color(fl.fgcolor),
          isnothing(fl.bgcolor) ? "" : " on " * _show_color(fl.bgcolor), ")")

function Base.show(io::IO, ::MIME"text/plain", fl::DrawingFill)
    println(io, "XLSX.Charts.DrawingFill ", fl.kind)
    isnothing(fl.preset)  || println(io, "  pattern: ", fl.preset)
    isnothing(fl.fgcolor) || println(io, "  foreground: ", _show_color(fl.fgcolor))
    isnothing(fl.bgcolor) || println(io, "  background: ", _show_color(fl.bgcolor))
    fl.kind in (:gradient, :blip, :group) &&
        println(io, "  (not modelled; preserved on write)")
end

function Base.show(io::IO, ln::DrawingLine)
    parts = String[]
    if !isnothing(ln.fill)
        push!(parts, ln.fill.kind === :none ? "no stroke" :
                     isnothing(ln.fill.fgcolor) ? string(ln.fill.kind) :
                     _show_color(ln.fill.fgcolor))
    end
    isnothing(ln.width) || push!(parts, string(ln.width) * "pt")
    isnothing(ln.dash)  || push!(parts, ln.dash)
    print(io, "XLSX.Charts.DrawingLine(",
          isempty(parts) ? "inherited" : join(parts, ", "), ")")
end

function Base.show(io::IO, ::MIME"text/plain", ln::DrawingLine)
    println(io, "XLSX.Charts.DrawingLine")
    if isnothing(ln.fill)
        println(io, "  stroke: inherited")
    elseif ln.fill.kind === :none
        println(io, "  stroke: none")
    else
        println(io, "  stroke: ", ln.fill)
    end
    isnothing(ln.width) || println(io, "  width: ", ln.width, " pt")
    for (label, v) in (("dash", ln.dash), ("cap", ln.cap),
                       ("compound", ln.compound), ("join", ln.join))
        isnothing(v) || println(io, "  ", label, ": ", v)
    end
    isnothing(ln.miter_limit) ||
        println(io, "  miter limit: ", ln.miter_limit, " × width")
end

# Inheritance. Every field on DrawingRunProps is Union{Nothing,T}: absent means
# "inherit", not "default", because a run's a:rPr sets only what differs from
# what it inherits. The cascade, innermost first:
#
#   1. the run's own a:rPr
#   2. the paragraph's a:pPr/a:defRPr
#   3. the body's a:lstStyle/a:lvl<n>pPr/a:defRPr (n from a:pPr/@lvl)
#   4. the c:txPr of the enclosing element (axis, series, data labels, legend)
#   5. c:chartSpace/c:txPr
#   6. the theme font scheme for +mj-*/+mn-* typefaces — get_theme_fonts —
#      and Excel's built-in defaults for everything else
#
# Levels 4 and 5 need a parent chain that doesn't exist until stage 3, and the
# rules differ by site (a data label does not inherit like an axis title), so
# the resolver is stage 4 work. Until then nothing user-facing should promise a
# resolved value: parsers return what is written, and `default_run_props`
# returns the nearest as-written properties, not the effective ones.

# =============================================================================
# DrawingML text: CT_TextBody
#
# The same content model serves `c:txPr` (formatting-only: axes, data labels,
# legend), `c:rich` (text with content: titles), and the cx: equivalent. One
# set of parsers; the caller says which child tag to look for.
#
# Two shapes to keep in mind, because they read differently:
#
#   txPr  — no runs at all. Font lives in `a:pPr/a:defRPr`, and Excel writes an
#           empty `a:endParaRPr` after it. `default_run_props` resolves this.
#   rich  — one or more `a:r`, each with its own `a:rPr` setting only what
#           differs from `defRPr`. May be uniform or mixed; see `is_uniform`.
#
# As with `prstDash`, several things that look like attributes are elements:
# autofit (`a:noAutofit`/`a:normAutofit`/`a:spAutoFit`), line spacing
# (`a:lnSpc/a:spcPct`), and the typefaces (`a:latin`, `a:ea`, `a:cs`).
#
# Write-back note for stage 4: `a:rPr` is an xsd:sequence — ln, fill,
# effectLst, highlight, uLn*, uFill*, latin, ea, cs, sym, hlink*, rtl, extLst.
# Splice into `raw`; never regenerate.
# =============================================================================

# ---------------------------------------------------------------------------
# Attribute readers
# ---------------------------------------------------------------------------

const EMU_PER_POINT = 12700

_attr_int(node, key) = (s = _attr(node, key); s === nothing ? nothing : tryparse(Int, s))

_attr_bool(node, key) =
    (s = _attr(node, key); s === nothing ? nothing : (s == "1" || s == "true" || s == "on"))

# 1/100 pt -> pt   (sz, kern, spc)
_attr_pt(node, key) = (n = _attr_int(node, key); n === nothing ? nothing : n / 100)

# EMU -> pt        (marL, marR, indent, insets)
_attr_emu(node, key) = (n = _attr_int(node, key); n === nothing ? nothing : n / EMU_PER_POINT)

# 1/60000 deg -> deg  (rot)
_attr_deg(node, key) = (n = _attr_int(node, key); n === nothing ? nothing : n / 60_000)

# As `_attr_pct`, but returns `nothing` rather than a default when the
# attribute is absent — DrawingML text needs absent and explicit to stay
# distinct. Same units: thousandths of a percent in, fraction out (60000 -> 0.6).
_attr_pct_opt(node, key) = (n = _attr_int(node, key); n === nothing ? nothing : _pct(n))


"""
    _attr_spacing(parent, tag) -> Union{Nothing,Tuple{Symbol,Float64}}

`a:lnSpc`, `a:spcBef` and `a:spcAft` each wrap either `a:spcPct` (1/1000 %) or
`a:spcPts` (1/100 pt). Returns `(:pct, 150.0)` or `(:pts, 12.0)`.
"""
function _attr_spacing(parent::XML.Node, tag::AbstractString)
    el = first_element_with_tag(parent, tag)
    el === nothing && return nothing
    pct = first_element_with_tag(el, "spcPct")
    pct === nothing || return (:frac, _attr_pct_opt(pct, "val"))
    pts = first_element_with_tag(el, "spcPts")
    pts === nothing || return (:pts, _attr_pt(pts, "val"))
    return nothing
end

# ---------------------------------------------------------------------------
# a:rPr / a:defRPr / a:endParaRPr
# ---------------------------------------------------------------------------

"""
    parse_drawing_run_props(wb, node; tag="rPr") -> Union{Nothing,DrawingRunProps}

Parse a run-properties element. `node` may be the element itself or its parent;
`tag` selects `rPr`, `defRPr` or `endParaRPr` when searching a parent.

Fields left `nothing` are absent from the file, which means "inherit" — not
"default". Resolving the cascade is `effective_run_props`' job, not this one's.
"""
function parse_drawing_run_props(wb::Workbook, node::XML.Node; tag::AbstractString="rPr")
    el = has_localname(node, tag) ? node : first_element_with_tag(node, tag)
    el === nothing && return nothing

    return DrawingRunProps(
        _attr(el, "lang"),
        _attr_pt(el, "sz"),
        _attr_bool(el, "b"),
        _attr_bool(el, "i"),
        _attr(el, "u"),          # none | sng | dbl | heavy | dotted | ...
        _attr(el, "strike"),     # noStrike | sngStrike | dblStrike
        _attr(el, "cap"),        # none | small | all
        _attr_pct_opt(el, "baseline"),
        _attr_pt(el, "kern"),
        _attr_pt(el, "spc"),
        parse_drawing_fill(wb, el),
        parse_drawing_line(wb, el),
        _attr(first_element_with_tag(el, "latin"), "typeface"),
        _attr(first_element_with_tag(el, "ea"), "typeface"),
        _attr(first_element_with_tag(el, "cs"), "typeface"),
        el,
    )
end

# ---------------------------------------------------------------------------
# a:pPr
# ---------------------------------------------------------------------------

"""
    parse_drawing_paragraph_props(wb, node; tag="pPr") -> Union{Nothing,DrawingParaProps}

Parse paragraph properties. The nested `a:defRPr` is parsed into a full
`DrawingRunProps` — in a chart `txPr` it is the only place the font appears.
"""
function parse_drawing_paragraph_props(wb::Workbook, node::XML.Node; tag::AbstractString="pPr")
    el = has_localname(node, tag) ? node : first_element_with_tag(node, tag)
    el === nothing && return nothing

    return DrawingParaProps(
        _attr(el, "algn"),       # l | ctr | r | just | justLow | dist | thaiDist
        _attr_int(el, "lvl"),
        _attr_emu(el, "marL"),
        _attr_emu(el, "marR"),
        _attr_emu(el, "indent"),
        _attr_bool(el, "rtl"),
        _attr_spacing(el, "lnSpc"),
        _attr_spacing(el, "spcBef"),
        _attr_spacing(el, "spcAft"),
        parse_drawing_run_props(wb, el; tag="defRPr"),
        el,
    )
end

# ---------------------------------------------------------------------------
# a:bodyPr
# ---------------------------------------------------------------------------

"""
    parse_drawing_body_props(node; tag="bodyPr") -> Union{Nothing,DrawingBodyProps}

Parse text-body properties. Autofit is an element, not an attribute:
`a:noAutofit` -> `:none`, `a:normAutofit` -> `:normal` (with `fontscale` and
`linespacereduction`), `a:spAutoFit` -> `:shape`. Absent means "inherit".
"""
function parse_drawing_body_props(node::XML.Node; tag::AbstractString="bodyPr")
    el = has_localname(node, tag) ? node : first_element_with_tag(node, tag)
    el === nothing && return nothing

    autofit = nothing
    fontscale = nothing
    lnspcred = nothing
    if first_element_with_tag(el, "noAutofit") !== nothing
        autofit = :none
    elseif (na = first_element_with_tag(el, "normAutofit")) !== nothing
        autofit = :normal
        fontscale = _attr_pct_opt(na, "fontScale")
        lnspcred = _attr_pct_opt(na, "lnSpcReduction")
    elseif first_element_with_tag(el, "spAutoFit") !== nothing
        autofit = :shape
    end

    return DrawingBodyProps(
        _attr_deg(el, "rot"),
        _attr(el, "vert"),          # horz | vert | vert270 | wordArtVert | ...
        _attr(el, "wrap"),          # none | square
        _attr(el, "anchor"),        # t | ctr | b | just | dist
        _attr_bool(el, "anchorCtr"),
        _attr_bool(el, "upright"),
        _attr_bool(el, "spcFirstLastPara"),
        _attr(el, "vertOverflow"),  # overflow | ellipsis | clip
        _attr(el, "horzOverflow"),  # overflow | clip
        _attr_emu(el, "lIns"),
        _attr_emu(el, "tIns"),
        _attr_emu(el, "rIns"),
        _attr_emu(el, "bIns"),
        autofit,
        fontscale,
        lnspcred,
        el,
    )
end

"""
    parse_drawing_shape_props(wb, node; tag="spPr") -> Union{Nothing,DrawingShapeProps}

Parse a shape-properties element. `node` may be the element itself or its
parent, and may be `nothing` so that optional lookups chain without guards:

    parse_drawing_shape_props(wb, series_node)                    # c:spPr
    parse_drawing_shape_props(wb, first_element_with_tag(x, "y")) # may be nothing

Returns `nothing` when there is no `spPr` — which for a series means "inherit
from the chart style", not "no fill".
"""
function parse_drawing_shape_props(wb::Workbook, node::Union{Nothing,XML.Node};
                                   tag::AbstractString="spPr")::Union{Nothing,DrawingShapeProps}
    isnothing(node) && return nothing
    el = has_localname(node, tag) ? node : first_element_with_tag(node, tag)
    isnothing(el) && return nothing

    effects = first_element_with_tag(el, "effectLst")
    isnothing(effects) && (effects = first_element_with_tag(el, "effectDag"))

    return DrawingShapeProps(
        parse_drawing_fill(wb, el),
        parse_drawing_line(wb, el),
        effects,
        _attr(el, "bwMode"),
        el,
    )
end

# ---------------------------------------------------------------------------
# Convenience predicates
#
# These exist because `isnothing(sp.fill)` and `sp.fill.kind == :none` are easy
# to conflate at a call site, and the difference is the whole point of the type.
# ---------------------------------------------------------------------------

"""
    hasFill(x) -> Bool

Whether `x` sets a visible fill. `x` is a [`DrawingShapeProps`](@ref), a
[`DrawingFill`](@ref), or the [`Effective`](@ref) a resolver returns.

`false` covers two different cases: nothing was written, so the fill is
inherited; or `<a:noFill/>` was, so it is deliberately transparent. Read
`sp.fill`, or `e.value`, to tell those apart — `nothing` for the first, a fill
with `kind === :none` for the second.

`false` never means Excel draws nothing: a property written at no rung comes
from the chart style, which this layer does not read.
"""
hasFill(sp::DrawingShapeProps) = !isnothing(sp.fill) && sp.fill.kind !== :none
hasFill(f::DrawingFill) = f.kind !== :none
hasFill(e::Effective{DrawingFill}) = !isnothing(e.value) && hasFill(e.value)

"""
    hasLine(x) -> Bool

Whether `x` sets a visible outline. `x` is a [`DrawingShapeProps`](@ref), a
[`DrawingLine`](@ref), or the [`Effective`](@ref) a resolver returns.

`false` when `a:ln` is absent, so the outline is inherited, and also when it
contains `<a:noFill/>`, which is how Excel writes "no border". As with
[`hasFill`](@ref), read the line itself to tell those apart.

`false` never means Excel draws nothing: a property written at no rung comes
from the chart style, which this layer does not read.
"""
hasLine(sp::DrawingShapeProps) =
    !isnothing(sp.line) && !isnothing(sp.line.fill) && sp.line.fill.kind !== :none
hasLine(l::DrawingLine) = !isnothing(l.fill) && l.fill.kind !== :none
hasLine(e::Effective{DrawingLine}) = !isnothing(e.value) && hasLine(e.value)

function Base.show(io::IO, sp::DrawingShapeProps)
    parts = String[]
    if isnothing(sp.fill)
        push!(parts, "fill inherited")
    else
        push!(parts, "fill $(sp.fill.kind)")
    end
    if isnothing(sp.line)
        push!(parts, "line inherited")
    elseif !isnothing(sp.line.fill) && sp.line.fill.kind === :none
        push!(parts, "no line")
    else
        push!(parts, isnothing(sp.line.width) ? "line" : "line $(sp.line.width)pt")
    end
    isnothing(sp.effects) || push!(parts, "effects")
    print(io, "DrawingShapeProps(", join(parts, ", "), ")")
end

# ---------------------------------------------------------------------------
# a:p and its runs
# ---------------------------------------------------------------------------

"""
    parse_drawing_paragraph(wb, p) -> DrawingParagraph

`a:r`, `a:br` and `a:fld` interleave in document order, so this walks children
rather than using `elements_with_tag`. Non-element children (whitespace in a
formatted part, comments) are skipped — Excel writes chart parts unindented,
but a part that has been through a formatter or hand-edited will have them.

`a:fld` (page numbers and the like) is kept as a run with its cached `a:t` text
and `kind == :fld`; anything richer stays in `raw`.
"""
function parse_drawing_paragraph(wb::Workbook, p::XML.Node)
    runs = DrawingRun[]
    for c in XML.children(p)
        XML.nodetype(c) === XML.Element || continue
        ln = localname(c)
        if ln == "r"
            push!(runs, DrawingRun(:run, child_text(c, "t"), parse_drawing_run_props(wb, c), c))
        elseif ln == "br"
            push!(runs, DrawingRun(:br, "\n", parse_drawing_run_props(wb, c), c))
        elseif ln == "fld"
            push!(runs, DrawingRun(:fld, child_text(c, "t"), parse_drawing_run_props(wb, c), c))
        end
    end

    return DrawingParagraph(
        parse_drawing_paragraph_props(wb, p),
        runs,
        parse_drawing_run_props(wb, p; tag="endParaRPr"),
        p,
    )
end

# ---------------------------------------------------------------------------
# CT_TextBody
# ---------------------------------------------------------------------------

"""
    parse_drawing_text(wb, node; tag="txPr") -> Union{Nothing,DrawingText}

Parse a DrawingML text body. `node` may be the text body itself (`c:txPr`,
`c:rich`, `cx:txPr`) or its parent, in which case `tag` selects the child.

    parse_drawing_text(wb, axis_node)                 # c:txPr
    parse_drawing_text(wb, tx_node; tag = "rich")     # title text

`a:lstStyle` is preserved as a node only: in chart parts it is either empty or
carries list-level defaults we do not model, and dropping it would change how
Excel renders inherited text.
"""
function parse_drawing_text(wb::Workbook, node::XML.Node; tag::AbstractString="txPr")
    el = if has_localname(node, tag) || _is_text_body(node)
        node
    else
        first_element_with_tag(node, tag)
    end
    el === nothing && return nothing

    paragraphs = [parse_drawing_paragraph(wb, p) for p in elements_with_tag(el, "p")]

    return DrawingText(
        parse_drawing_body_props(el),
        first_element_with_tag(el, "lstStyle"),
        paragraphs,
        el,
    )
end
parse_drawing_text(::Workbook, ::Nothing; kw...) = nothing

_is_text_body(node::XML.Node) =
    first_element_with_tag(node, "bodyPr") !== nothing && !has_localname(node, "spPr")

# ---------------------------------------------------------------------------
# Reading the text back out
# ---------------------------------------------------------------------------

"""
    text_content(t::DrawingText) -> String

Concatenate run text, one line per paragraph. Empty for a formatting-only
`txPr`; the visible string for a `c:rich` title.

Excel splits runs on language and spellcheck boundaries, so a title typed as
one string may arrive as several runs with identical properties. This
reassembles it; editing does not — replacing whole text means collapsing to a
single run, preserving the first run's `rPr`.
"""
function text_content(t::DrawingText)
    io = IOBuffer()
    for (i, p) in enumerate(t.paragraphs)
        i > 1 && print(io, "\n")
        for r in p.runs
            r.text !== nothing && print(io, r.text)
        end
    end
    return String(take!(io))
end

"""
    text_content(node::XML.Node) -> String

The text of a DrawingML text body (`c:rich`, `cx:rich`, `a:txBody`, …) read directly
from the XML: the text of its runs and fields, with paragraphs separated by `"\\n"`.
Agrees with `text_content(::DrawingText)` for the same body, without resolving any
formatting.
"""
function text_content(node::XML.Node)::String
    io = IOBuffer()
    n  = 0
    for p in XML.eachelement(node)
        localname(p) == "p" || continue
        (n += 1) > 1 && print(io, "\n")
        for r in XML.eachelement(p)
            t = localname(r)
            t in ("r", "fld") && print(io, something(child_text(r, "t"), ""))
            t == "br"         && print(io, "\n")
        end
    end
    return String(take!(io))
end

text_runs(t::DrawingText) = [r for p in t.paragraphs for r in p.runs]

# Comparable identity for run properties, ignoring `lang` and `raw`: two runs
# differing only by spellcheck language are the same formatting.
#
# Two conventions shared with `Base.:(==)(::RichTextRun, ::RichTextRun)`:
# colours compare on their resolved value rather than how they were written
# (as that does via `get_color`), and an absent property is distinct from one
# set to a default (as that does by comparing key sets).
_color_key(c) = c === nothing ? nothing : (c.rgb, c.alpha)
_fill_key(f) = f === nothing ? nothing : (f.kind, _color_key(f.fgcolor), _color_key(f.bgcolor), f.preset)
_line_key(l) = l === nothing ? nothing : (_fill_key(l.fill), l.width, l.dash, l.cap, l.compound)

_props_key(::Nothing) = nothing
_props_key(p::DrawingRunProps) = (
    p.size, p.bold, p.italic, p.under, p.strike, p.caps,
    p.baseline, p.kern, p.spacing,
    _fill_key(p.fill), _line_key(p.line),
    p.latin, p.ea, p.cs,
)

"""
    is_uniform(t::DrawingText) -> Bool

Whether every run in `t` carries the same formatting, comparing only what is
written on each `a:rPr` (each run's rPr is overlaid on its paragraph's defRPr). 
Text with no runs is uniform by definition.

This compares *as written*, not as resolved — two runs reaching the same
appearance by different inheritance paths compare as different. That is the
conservative direction: it can report mixed where a full cascade would report
uniform, but never the reverse.

Unlike `==` on `RichTextRun`, this ignores run text entirely: the question is
whether formatting is consistent across the text, not whether the runs are
identical. Two runs reading differently but formatted alike are uniform.
"""
function is_uniform(t::DrawingText)
    key = nothing
    seen = false
    for p in t.paragraphs
        def = p.props === nothing ? nothing : p.props.defprops
        for r in p.runs
            k = _props_key(_overlay_run_props(def, r.props))
            if !seen
                key = k
                seen = true
            elseif k != key
                return false
            end
        end
    end
    return true
end

"""
    default_run_props(t::DrawingText) -> Union{Nothing,DrawingRunProps}

The run properties describing this text body's appearance as a whole: the
first paragraph's `a:defRPr` if present, else the first run's `a:rPr`, else the
first `a:endParaRPr`.

Returns `nothing` for text whose runs disagree — a mixed-format title has no
single answer, and returning the first run's properties would quietly describe
only its opening fragment. Use `first_run_props` if you want that value anyway,
or `is_uniform` to test first.
"""
function default_run_props(t::DrawingText)
    is_uniform(t) || return nothing
    return first_run_props(t)
end

# Field-wise overlay: each field of `over` that is set wins, otherwise `base`'s.
# `raw` is taken from `over` when it has one, since that is the node a caller
# would edit to change what renders.
function _overlay_run_props(base::Union{Nothing,DrawingRunProps},
                            over::Union{Nothing,DrawingRunProps})
    isnothing(base) && return over
    isnothing(over) && return base
    vals = map(fieldnames(DrawingRunProps)) do f
        f === :raw && return isnothing(over.raw) ? base.raw : over.raw
        v = getfield(over, f)
        isnothing(v) ? getfield(base, f) : v
    end
    return DrawingRunProps(vals...)
end

"""
    first_run_props(t::DrawingText) -> Union{Nothing,DrawingRunProps}

The run properties of the first text in `t`, regardless of whether later runs
agree.

Where the paragraph has runs, the first run's `a:rPr` is overlaid field by
field on the paragraph's `a:defRPr`: a field the run sets wins, a field it
omits comes from the default. An empty `a:defRPr` therefore never hides a
run's formatting.

Where it has no runs (a formatting-only `txPr`), the `a:defRPr` answers as
written, and `a:endParaRPr` only when there is no `defRPr`: the end mark
describes the paragraph end, not its text.

For mixed text this describes only the first fragment; see
[`default_run_props`](@ref).
"""
function first_run_props(t::DrawingText)
    for p in t.paragraphs
        def = p.props === nothing ? nothing : p.props.defprops
        if isempty(p.runs)
            isnothing(def) || return def
        else
            merged = _overlay_run_props(def, p.runs[1].props)
            isnothing(merged) || return merged
        end
        isnothing(p.endprops) || return p.endprops
    end
    return nothing
end

function Base.show(io::IO, t::DrawingText)
    s = text_content(t)
    np = length(t.paragraphs)
    nr = sum(length(p.runs) for p in t.paragraphs; init=0)
    if isempty(s)
        print(io, "DrawingText($np paragraph(s), formatting only)")
    else
        mixed = is_uniform(t) ? "" : ", mixed"
        print(io, "DrawingText(", repr(truncate_len(s, 40)), ", $np paragraph(s), $nr run(s)$mixed)")
    end
end

#-- Writing ------------------------------------------------------------------

"""
    _sp_with_fill(sp, key, color, pfx) -> XML.Node

Set the fill of an element that carries one — an `a:spPr`, or a run-properties
element, which share the fill group. `key` is that element's [`SchemaKey`](@ref).

`:none` writes `<a:noFill/>`; `:inherit` removes the fill so the cascade
resolves it.
"""
function _sp_with_fill(sp::XML.Node, key::SchemaKey, color, pfx::Dict{String,String})
    color === :inherit && return remove_choice(sp, key, FILL_GROUP)
    fill = color === :none       ? XML.Element(prefixed_tag(pfx[NS_A], "noFill")) :
           color isa SchemeColor ? _solid_fill_node(_scheme_color_node(color, pfx), pfx) :
                                   _solid_fill_node(_srgb_color_node(color, pfx), pfx)
    return insert_child(remove_choice(sp, key, FILL_GROUP), key, fill)
end

"""
    _sp_with_line(sp, key, f, pfx) -> XML.Node

Apply `f` to the `a:ln` of an element that carries one, creating it if absent.
`key` is that element's [`SchemaKey`](@ref) — `(NS_A, "spPr")` for shape
properties, `(NS_A, "defRPr")` for run properties.
"""
function _sp_with_line(sp::XML.Node, key::SchemaKey, f, pfx::Dict{String,String})
    ln = something(first_element_with_tag(sp, "ln"),
                   XML.Element(prefixed_tag(pfx[NS_A], "ln")))
    return insert_child(sp, key, f(ln))
end

"""
    _scheme_color_node(sc::SchemeColor, prefixes) -> XML.Node

Build the `a:schemeClr` element for `sc`, with its transforms as child elements
in the order they are held. Percentages are written back as the hundredths
DrawingML expects: `lumMod = 75` becomes `val="75000"`.
"""
function _scheme_color_node(sc::SchemeColor,
                            prefixes::Dict{String,String} = Dict(NS_A => "a"))
    pfx  = prefixes[NS_A]
    kids = XML.Node{String}[
        XML.Element(prefixed_tag(pfx, String(name)); val = string(round(Int, v * 1000)))
        for (name, v) in sc.transforms
    ]
    node = XML.Element(prefixed_tag(pfx, "schemeClr"); val = String(sc.token))
    return isempty(kids) ? node : _with_children(node, kids)
end

_solid_fill_node(color_node::XML.Node,
                 prefixes::Dict{String,String} = Dict(NS_A => "a")) =
    _with_children(XML.Element(prefixed_tag(prefixes[NS_A], "solidFill")),
                   XML.Node{String}[color_node])

"""
    _srgb_color_node(color, prefixes) -> XML.Node

Build the `a:srgbClr` element for `color`, which may be an eight-digit
`AARRGGBB` string or any Colors.jl named color — the same spellings `setFill`
accepts, via [`get_color`](@ref).

Alpha is split out into an `a:alpha` child, because DrawingML carries the color
as six digits and alpha as a transform, unlike the spreadsheet's eight-digit
form. A fully opaque color emits no `a:alpha`, since writing one Excel would
have omitted is a change for no reason.
"""
function _srgb_color_node(color::AbstractString,
                          prefixes::Dict{String,String} = Dict(NS_A => "a"))
    argb = get_color(String(color))
    pfx  = prefixes[NS_A]
    node = XML.Element(prefixed_tag(pfx, "srgbClr"); val = argb[3:end])

    a = parse(Int, argb[1:2]; base = 16)
    a == 255 && return node
    return _with_children(node, XML.Node{String}[
        XML.Element(prefixed_tag(pfx, "alpha"); val = string(round(Int, a / 255 * 100000)))])
end

_srgb_color_node(color::Symbol, prefixes::Dict{String,String} = Dict(NS_A => "a")) =
    _srgb_color_node(String(color), prefixes)

_srgb_color_node(c::Colors.Colorant, prefixes::Dict{String,String} = Dict(NS_A => "a")) =
    _srgb_color_node(Colors.hex(c, :AARRGGBB), prefixes)

"""
    _ln_with_width(ln, points, pfx) -> XML.Node

Set the `w` attribute of an `a:ln`, in points — the unit Excel's width box uses.
`:inherit` removes it.

`ST_LineWidth` caps at 20116800 EMU, so a width above 1584pt makes the part
invalid; it is rejected here rather than written.
"""
function _ln_with_width(ln::XML.Node, points)
    points === :inherit && return with_attribute(ln, "w", nothing)
    emu = round(Int, points * 12700)
    0 <= emu <= 20116800 || throw(XLSXError(
        "A line width of $points pt is outside the range DrawingML allows (0 to 1584 pt)."))
    return with_attribute(ln, "w", emu)
end

_ln_with_cap(ln, cap) =
    with_attribute(ln, "cap", cap === :inherit ? nothing :
                   String(_check(cap, CAP_VALUES, "cap", CAP_ALIASES)))

_ln_with_compound(ln, cmpd) =
    with_attribute(ln, "cmpd", cmpd === :inherit ? nothing :
                   String(_check(cmpd, CMPD_VALUES, "compound", CMPD_ALIASES)))

function _ln_with_dash(ln, dash, pfx)
    dash === :inherit && return remove_choice(ln, (NS_A, "ln"), DASH_GROUP)
    node = XML.Element(prefixed_tag(pfx[NS_A], "prstDash");
                       val = String(_check(dash, DASH_VALUES, "dash", DASH_ALIASES)))
    return insert_child(remove_choice(ln, (NS_A, "ln"), DASH_GROUP), (NS_A, "ln"), node)
end

function _ln_with_color(ln, color, pfx)
    color === :inherit && return remove_choice(ln, (NS_A, "ln"), FILL_GROUP)
    fill = color === :none      ? XML.Element(prefixed_tag(pfx[NS_A], "noFill")) :
           color isa SchemeColor ? _solid_fill_node(_scheme_color_node(color, pfx), pfx) :
                                   _solid_fill_node(_srgb_color_node(color, pfx), pfx)
    return insert_child(remove_choice(ln, (NS_A, "ln"), FILL_GROUP), (NS_A, "ln"), fill)
end

# ST_PresetLineDashVal, ST_LineCap and ST_CompoundLine (ECMA-376, dml-main.xsd).
# Vocabulary as written — the Excel UI names are aliases, not replacements.
const DASH_VALUES = (:solid, :dot, :dash, :lgDash, :dashDot, :lgDashDot,
                     :lgDashDotDot, :sysDash, :sysDot, :sysDashDot, :sysDashDotDot)

const CAP_VALUES  = (:rnd, :sq, :flat)

const CMPD_VALUES = (:sng, :dbl, :thickThin, :thinThick, :tri)

const JOIN_VALUES = (:round, :bevel, :miter)

"""
    _ln_with_join(ln, join, pfx) -> XML.Node

Set the join of an `a:ln`: `:round`, `:bevel` or `:miter`, or `:inherit` to
remove it. Excel uses the same three words, so there are no aliases.

A miter limit is an attribute of the `a:miter` element rather than a property of
the line, so it is set with [`_ln_with_miter_limit`](@ref) and is lost if the
join is later changed.
"""
function _ln_with_join(ln::XML.Node, join, pfx::Dict{String,String})
    join === :inherit && return remove_choice(ln, (NS_A, "ln"), JOIN_GROUP)
    tag  = String(_check(join, JOIN_VALUES, "join"))
    node = XML.Element(prefixed_tag(pfx[NS_A], tag))
    return insert_child(remove_choice(ln, (NS_A, "ln"), JOIN_GROUP), (NS_A, "ln"), node)
end

"""
    _ln_with_miter_limit(ln, limit, pfx) -> XML.Node

Set the `lim` attribute of an `a:ln`'s `a:miter`, as a multiple of the line
width — Excel's default is 8. `:inherit` removes it. Throws where the join is
not miter, since the attribute has no meaning on `a:round` or `a:bevel`.
"""
function _ln_with_miter_limit(ln::XML.Node, limit)
    m = first_element_with_tag(ln, "miter")
    isnothing(m) && throw(XLSXError(
        "A miter limit needs a miter join; set `join = :miter` first."))
    new = with_attribute(m, "lim", limit === :inherit ? nothing : round(Int, limit * 100000))
    return replace_child(ln, m, new)
end

"""
    _color_node_from(dc::DrawingColor, pfx) -> XML.Node

Serialize a `DrawingColor`, keeping the kind it was written as. Transforms are
emitted in the order they are held, as DrawingML applies them in sequence.

`rgb` and `alpha` are resolved values, not source, and are not written — the
file keeps the scheme reference and its transforms so the theme still applies.

`hslClr` and `scrgbClr` carry component attributes rather than a `val`, which
`DrawingColor` does not model, so they cannot be written.
"""
function _color_node_from(dc::DrawingColor, pfx::Dict{String,String})
    dc.kind in (:hsl, :scrgb) && throw(XLSXError(
        "`$(dc.kind)` colors carry component attributes rather than a `val`, " *
        "which is not modelled; write the color as srgb instead."))

    tag = dc.kind === :scheme ? "schemeClr" :
          dc.kind === :srgb   ? "srgbClr"   :
          dc.kind === :sys    ? "sysClr"    :
          dc.kind === :prst   ? "prstClr"   :
          throw(XLSXError("Unknown color kind `$(dc.kind)`."))

    kids = XML.Node{String}[_el(pfx, String(name); val = string(v))
                            for (name, v) in dc.transforms]
    return _el(pfx, tag, kids...; val = dc.val)
end
_color_node_from(sc::SchemeColor, pfx::Dict{String,String}) = _scheme_color_node(sc, pfx)
_color_node_from(s::AbstractString, pfx::Dict{String,String}) = _srgb_color_node(s, pfx)

"""
    _fill_node_from(fill::DrawingFill, pfx) -> XML.Node

Serialize a `DrawingFill`. Solid and none are built fresh; gradient, pattern,
blip and group fills are emitted from `raw`, because the struct models them
partially and reconstructing one from what it keeps would lose detail.

A fill of those kinds with no `raw` throws — it was constructed rather than
parsed, and there is nothing to write.
"""
function _fill_node_from(fill::DrawingFill, pfx::Dict{String,String})
    fill.kind === :none && return _el(pfx, "noFill")

    if fill.kind === :solid
        isnothing(fill.fgcolor) && throw(XLSXError("A solid fill needs a color."))
        return _el(pfx, "solidFill", _color_node_from(fill.fgcolor, pfx))
    end

    isnothing(fill.raw) && throw(XLSXError(
        "A $(fill.kind) fill can only be written from a parsed one; this was " *
        "constructed and carries no node."))
    return fill.raw
end

"""
    _line_node(line::DrawingLine, pfx) -> XML.Node

Serialize a `DrawingLine` by applying each property it sets to a fresh `a:ln`.
Absent fields are not written, so a line says only what it carries.
"""
function _line_node(line::DrawingLine, pfx::Dict{String,String})
    node = _el(pfx, "ln")
    isnothing(line.width)       || (node = _ln_with_width(node, line.width))
    isnothing(line.cap)         || (node = _ln_with_cap(node, Symbol(line.cap)))
    isnothing(line.compound)    || (node = _ln_with_compound(node, Symbol(line.compound)))
    isnothing(line.fill)        || (node = insert_child(node, (NS_A, "ln"),
                                               _fill_node_from(line.fill, pfx)))
    isnothing(line.dash)        || (node = _ln_with_dash(node, Symbol(line.dash), pfx))
    isnothing(line.join)        || (node = _ln_with_join(node, Symbol(line.join), pfx))
    isnothing(line.miter_limit) || (node = _ln_with_miter_limit(node, line.miter_limit))
    return node
end

#-- Text body serialization ---------------------------------------------------
#
# The inverse of parse_drawing_text. Every field that was converted at parse
# time is converted back here: points to 1/100pt or EMU, degrees to 1/60000,
# fractions to thousandths of a percent. A field left `nothing` is not written,
# which is what makes a constructed DrawingText say only what the caller asked.

_el(pfx, tag, kids...; attrs...) = XML.Element(prefixed_tag(pfx[NS_A], tag), kids...; attrs...)
#_el(pfx, tag; kw...) = XML.Element(prefixed_tag(pfx[NS_A], tag); kw...)


_emu(points) = round(Int, points * EMU_PER_POINT)

const _RUN_PROP_ATTRS = (
    :lang     => ("lang",     string),
    :size     => ("sz",       v -> string(round(Int, v * 100))),
    :bold     => ("b",        v -> v ? "1" : "0"),
    :italic   => ("i",        v -> v ? "1" : "0"),
    :under    => ("u",        string),
    :strike   => ("strike",   string),
    :caps     => ("cap",      string),
    :baseline => ("baseline", v -> string(round(Int, v * 100000))),
    :kern     => ("kern",     v -> string(round(Int, v * 100))),
    :spacing  => ("spc",      v -> string(round(Int, v * 100))),
)

# Field -> typeface element tag.
const _RUN_PROP_FONTS = Dict{Symbol,String}(
    :latin => "latin", :ea => "ea", :cs => "cs",
)

# Linear scan over ten entries — a parallel Dict for lookup would be faster and
# not worth the second structure to keep in step.
_run_prop_attr(field::Symbol) =
    (i = findfirst(p -> first(p) === field, _RUN_PROP_ATTRS);
     isnothing(i) ? nothing : last(_RUN_PROP_ATTRS[i]))

"""
    _rpr_with_prop(rpr, field, value, pfx) -> XML.Node

Set one field of a run-properties element (`a:defRPr`, `a:rPr` or
`a:endParaRPr`). `:inherit` removes it.

Most fields are attributes; `:fill` and `:line` are child elements, and the
three typefaces are `typeface` attributes on their own child elements.
`:line` takes either a color or a `NamedTuple` of line keywords — a bare value
is treated as `(color = value,)`.
"""
function _rpr_with_prop(rpr::XML.Node, field::Symbol, value, pfx::Dict{String,String})
    key = (NS_A, String(localname(rpr)))
    field = get(_RUN_PROP_ALIASES, field, field)

    spec = _run_prop_attr(field)
    if !isnothing(spec)
        name, enc = spec
        return with_attribute(rpr, name, value === :inherit ? nothing : enc(value))

    elseif haskey(_RUN_PROP_FONTS, field)
        tag = _RUN_PROP_FONTS[field]
        value === :inherit && return remove_child(rpr, tag)
        return insert_child(rpr, key,
                            XML.Element(prefixed_tag(pfx[NS_A], tag); typeface = string(value)))

    elseif field === :fill
        return _sp_with_fill(rpr, key, value, pfx)

    elseif field === :line
        value === :inherit && return remove_child(rpr, "ln")
        kw = value isa NamedTuple ? value : (color = value,)
        return _sp_with_line(rpr, key, ln -> _ln_with(ln, pfx; kw...), pfx)
    end

    throw(XLSXError("`$field` is not a run property. Valid fields: " *
                    join(_RUN_PROP_FIELDS, ", ") * "; or Excel's names " *
                    join(("$k (= $v)" for (k, v) in _RUN_PROP_ALIASES), ", ") * "."))
end

"""
    _text_with_run_prop(body, field, value, pfx) -> XML.Node

Set one run property on a text body's paragraph default (`a:defRPr`), and remove
that field from each run's `a:rPr` so the default is what renders.

A run's `a:rPr` overrides the paragraph default field by field, so writing only
`a:defRPr` leaves a body whose runs set that field unchanged on screen — which is
how Excel writes a title it has formatted. Clearing the field from the runs makes
`a:defRPr` the single source for it while leaving the rest of each run's
formatting alone. A setter that names no run is taken to mean the whole body.
"""
function _text_with_run_prop(body::XML.Node, field::Symbol, value, pfx::Dict{String,String})
    skey = (NS_A, String(localname(body)))
    body = rebuild_path(body,
        [(NS_A, "p")      => "p",
         (NS_A, "pPr")    => "pPr",
         (NS_A, "defRPr") => "defRPr"],
        rpr -> _rpr_with_prop(rpr, field, value, pfx);
        prefixes = pfx,
        parent_key = skey)
    return _clear_run_prop(body, field, pfx)
end

# Remove `field` from every a:r/a:rPr and a:endParaRPr in the body, so the
# paragraph default governs it. Runs with no a:rPr are left alone.
function _clear_run_prop(body::XML.Node, field::Symbol, pfx::Dict{String,String})
    isnothing(body.children) && return body
    kids = map(body.children) do p
        localname(p) != "p" && return p
        isnothing(p.children) && return p
        _with_children(p, map(p.children) do n
            tag = localname(n)
            if tag in ("r", "br", "fld")
                rpr = first_element_with_tag(n, "rPr")
                isnothing(rpr) ? n :
                    replace_child(n, rpr, _rpr_with_prop(rpr, field, :inherit, pfx))
            elseif tag == "endParaRPr"
                _rpr_with_prop(n, field, :inherit, pfx)
            else
                n
            end
        end)
    end
    return _with_children(body, kids)
end

"""
    _run_props_node(rp, tag, pfx) -> XML.Node

Serialize a `DrawingRunProps` as `a:rPr`, `a:defRPr` or `a:endParaRPr`, named
by `tag`. All three are `CT_TextCharacterProperties`.
"""
function _run_props_node(rp::DrawingRunProps, tag::AbstractString, pfx::Dict{String,String})
    node = _el(pfx, tag)
    for (field, (name, enc)) in _RUN_PROP_ATTRS
        v = getfield(rp, field)
        isnothing(v) || (node = with_attribute(node, name, enc(v)))
    end

    key = (NS_A, tag)
    isnothing(rp.line) || (node = insert_child(node, key, _line_node(rp.line, pfx)))
    isnothing(rp.fill) || (node = insert_child(node, key, _fill_node_from(rp.fill, pfx)))
    for (field, t) in _RUN_PROP_FONTS
        v = getfield(rp, field)
        isnothing(v) || (node = insert_child(node, key, _el(pfx, t; typeface = v)))
    end
    return node
end

"""
    _para_props_node(pp, pfx) -> XML.Node

Serialize a `DrawingParaProps` as `a:pPr`.
"""
function _para_props_node(pp::DrawingParaProps, pfx::Dict{String,String})
    node = _el(pfx, "pPr")
    isnothing(pp.align)       || (node = with_attribute(node, "algn", pp.align))
    isnothing(pp.level)       || (node = with_attribute(node, "lvl", string(pp.level)))
    isnothing(pp.marginleft)  || (node = with_attribute(node, "marL", string(_emu(pp.marginleft))))
    isnothing(pp.marginright) || (node = with_attribute(node, "marR", string(_emu(pp.marginright))))
    isnothing(pp.indent)      || (node = with_attribute(node, "indent", string(_emu(pp.indent))))
    isnothing(pp.rtl)         || (node = with_attribute(node, "rtl", pp.rtl ? "1" : "0"))

    key = (NS_A, "pPr")
    for (tag, v) in (("lnSpc", pp.linespacing), ("spcBef", pp.spacebefore),
                     ("spcAft", pp.spaceafter))
        isnothing(v) && continue
        kind, amount = v
        inner = kind === :pct ? _el(pfx, "spcPct"; val = string(round(Int, amount * 1000))) :
                                _el(pfx, "spcPts"; val = string(round(Int, amount * 100)))
        node = insert_child(node, key, XML.Element(prefixed_tag(pfx[NS_A], tag), inner))
    end
    isnothing(pp.defprops) ||
        (node = insert_child(node, key, _run_props_node(pp.defprops, "defRPr", pfx)))
    return node
end

"""
    _body_props_node(bp, pfx) -> XML.Node

Serialize a `DrawingBodyProps` as `a:bodyPr`. Autofit is one of three child
elements rather than an attribute, and `fontscale`/`linespacereduction` belong
to `a:normAutofit` alone.
"""
function _body_props_node(bp::DrawingBodyProps, pfx::Dict{String,String})
    node = _el(pfx, "bodyPr")
    isnothing(bp.rotation) || (node = with_attribute(node, "rot", string(round(Int, bp.rotation * 60000))))
    for (name, v) in (("vert", bp.vertical), ("wrap", bp.wrap), ("anchor", bp.anchor),
                      ("vertOverflow", bp.vertoverflow), ("horzOverflow", bp.horzoverflow))
        isnothing(v) || (node = with_attribute(node, name, v))
    end
    for (name, v) in (("anchorCtr", bp.anchorctr), ("upright", bp.upright),
                      ("spcFirstLastPara", bp.spcfirstlastpara))
        isnothing(v) || (node = with_attribute(node, name, v ? "1" : "0"))
    end
    for (name, v) in (("lIns", bp.insetleft), ("tIns", bp.insettop),
                      ("rIns", bp.insetright), ("bIns", bp.insetbottom))
        isnothing(v) || (node = with_attribute(node, name, string(_emu(v))))
    end

    if bp.autofit === :none
        node = insert_child(node, (NS_A, "bodyPr"), _el(pfx, "noAutofit"))
    elseif bp.autofit === :shape
        node = insert_child(node, (NS_A, "bodyPr"), _el(pfx, "spAutoFit"))
    elseif bp.autofit === :normal
        na = _el(pfx, "normAutofit")
        isnothing(bp.fontscale) ||
            (na = with_attribute(na, "fontScale", string(round(Int, bp.fontscale * 100000))))
        isnothing(bp.linespacereduction) ||
            (na = with_attribute(na, "lnSpcReduction",
                                 string(round(Int, bp.linespacereduction * 100000))))
        node = insert_child(node, (NS_A, "bodyPr"), na)
    end
    return node

end

"""
    _text_from(text, tag, pfx) -> XML.Node

Serialize a `DrawingText` as a text body element named `tag` — `txPr` or
`rich`, both `a:CT_TextBody`. The prefix comes from the tag's own namespace at
the call site; the body's children are all DrawingML.
"""
function _text_from(text::DrawingText, tag::AbstractString, pfx::Dict{String,String},
                    ns::AbstractString = NS_C)
    isempty(text.paragraphs) && throw(XLSXError(
        "A text body needs at least one paragraph; `CT_TextBody` requires it."))

    kids = XML.Node{String}[]
    # bodyPr is required even when it sets nothing, and a body without one is
    # rejected by Excel. lstStyle is optional, but Excel always writes it.
    push!(kids, isnothing(text.body) ? _el(pfx, "bodyPr") : _body_props_node(text.body, pfx))
    push!(kids, isnothing(text.liststyle) ? _el(pfx, "lstStyle") : text.liststyle)
    append!(kids, _paragraph_node(p, pfx) for p in text.paragraphs)
    return XML.Element(prefixed_tag(pfx[ns], tag), kids...)
end

"""
    _new_text_body(tag, pfx) -> XML.Node

An empty text body — `c:txPr` or `c:rich` — carrying the `a:bodyPr` that
`CT_TextBody` requires and the `a:lstStyle` Excel always writes. A body without
a `bodyPr` is rejected when the file is opened.
"""
_new_text_body(tag::AbstractString, pfx::Dict{String,String}, ns::AbstractString = NS_C) =
    XML.Element(prefixed_tag(pfx[ns], tag), _el(pfx, "bodyPr"), _el(pfx, "lstStyle"))

function _paragraph_node(p::DrawingParagraph, pfx::Dict{String,String})
    kids = XML.Node{String}[]
    isnothing(p.props) || push!(kids, _para_props_node(p.props, pfx))
    append!(kids, _run_node(r, pfx) for r in p.runs)
    isnothing(p.endprops) ||
        push!(kids, _run_props_node(p.endprops, "endParaRPr", pfx))
    return _el(pfx, "p", kids...)
end

function _run_node(r::DrawingRun, pfx::Dict{String,String})
    if r.kind === :br
        node = _el(pfx, "br")
        isnothing(r.props) || (node = insert_child(node, (NS_A, "br"),
                                        _run_props_node(r.props, "rPr", pfx)))
        return node
    end
    node = _el(pfx, r.kind === :fld ? "fld" : "r")
    isnothing(r.props) || (node = insert_child(node, (NS_A, "r"), _run_props_node(r.props, "rPr", pfx)))
    isnothing(r.text) || (node = insert_child(node, (NS_A, "r"),
                            XML.Element(prefixed_tag(pfx[NS_A], "t"), XML.Text(r.text))))
    return node
end
