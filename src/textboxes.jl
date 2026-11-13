#
# -- Text boxes and shapes: `<xdr:sp>` objects in a worksheet's drawing.
#
# Loaded after `Charts`, whose DrawingML layer writes their colours, lines and text.
#

const _DML_A = Dict(NS_A => "a")

# The preset shapes `addShape` makes, with the name Excel gives each kind.
const SHAPE_KINDS = (
    rect = "Rectangle", roundRect = "Rectangle: Rounded Corners", ellipse = "Oval",
    triangle = "Isosceles Triangle", rtTriangle = "Right Triangle", diamond = "Diamond",
    parallelogram = "Parallelogram", trapezoid = "Trapezoid", pentagon = "Pentagon",
    hexagon = "Hexagon", octagon = "Octagon", plus = "Cross", star5 = "Star: 5 Points",
    rightArrow = "Arrow: Right", leftArrow = "Arrow: Left", upArrow = "Arrow: Up",
    downArrow = "Arrow: Down", leftRightArrow = "Arrow: Left-Right", chevron = "Arrow: Chevron",
    homePlate = "Arrow: Pentagon", wedgeRectCallout = "Speech Bubble: Rectangle",
    wedgeRoundRectCallout = "Speech Bubble: Rectangle with Corners Rounded",
    wedgeEllipseCallout = "Speech Bubble: Oval", cloud = "Cloud", heart = "Heart",
    lightningBolt = "Lightning Bolt", sun = "Sun", moon = "Moon", smileyFace = "Smiley Face",
    flowChartProcess = "Flowchart: Process", flowChartDecision = "Flowchart: Decision",
    flowChartTerminator = "Flowchart: Terminator", can = "Cylinder", cube = "Cube",
)
const _SHAPE_HALIGN = (left = "l", center = "ctr", right = "r", justify = "just")
const _SHAPE_VALIGN = (top = "t", middle = "ctr", bottom = "b")
const _SHAPE_FONT_KEYS = (:bold, :italic, :under, :strike, :size, :color, :name, :vertAlign)
const _SHAPE_DEFAULT_SIZE = (192, 80)   # pixels: three default columns, four rows

"""
    addTextBox(ws::Worksheet, ref, text; kw...) -> NamedTuple

Add a text box, as Excel's *Insert → Text Box* does, floating over the cells like an
image. `ref` is a cell, where its top-left corner goes, or a range it fills (`"B2:E6"`),
as for [`addImage`](@ref). `text` is a string, in which `"\\n"` starts a new paragraph,
or a [`RichTextString`](@ref XLSX.RichTextString) for runs in different fonts.

# Keywords

- `size`: the box's `(width, height)` in pixels, from a single cell. It snaps to whole
  cells, as an image does. The default is `$(_SHAPE_DEFAULT_SIZE)`.
- `font`: a `NamedTuple` of the keywords of [`setFont`](@ref) for the whole text:
  `bold`, `italic`, `under`, `strike`, `size` (points), `color`, `name` and
  `vertAlign`. A rich text run's own settings take precedence.
- `fill`: the box's colour, or `:none`. A text box is white by default.
- `line`: the border's colour, or `:none`; `lineWidth` its width in points, and
  `lineDash` its dashes (`:solid`, `:dash`, `:dot`, `:lgDash`, `:dashDot`, `:sysDash`,
  `:sysDot`, ...). A text box has a thin grey border by default.
- `horizontal`: `:left`, `:center`, `:right` or `:justify`.
- `vertical`: `:top`, `:middle` or `:bottom`.
- `name`: the name Excel shows for the box; by default `TextBox n`, as Excel names them.

Colours can be $COLOR_FORMS_DOC

Returns `(sheet, name, from, to)`: the sheet, the box's name and the cells at its corners.

The file must be open for writing (`mode="rw"`).

# Examples
```julia
julia> XLSX.addTextBox(sheet, "B2:E6", "Quarterly results\\nAll figures in thousands";
                       font=(size=14, bold=true), fill="FFF2CC", horizontal=:center, vertical=:middle)

julia> XLSX.addTextBox(sheet, "G2", XLSX.RichTextString(XLSX.RichTextRun("Note: ", [:bold => true]),
                                                         XLSX.RichTextRun("draft figures")))
```

See also [`addShape`](@ref), [`getShapes`](@ref) and [`deleteShape`](@ref).
"""
addTextBox(ws::Worksheet, ref, text::AbstractString; kw...) = _add_sp!(ws, ref, nothing, text; kw...)

"""
    addShape(ws::Worksheet, ref, kind::Symbol; text="", kw...) -> NamedTuple

Add a shape, as Excel's *Insert → Shapes* does, floating over the cells like an image:
a rectangle, oval, arrow, callout and so on, in Excel's default style (the theme's first
accent colour, with white text). `kind` is one of
$(join((":$k" for k in keys(SHAPE_KINDS)), ", ")).

`text` is the shape's text, as for [`addTextBox`](@ref), whose keywords `addShape` takes
too. By default the text is at the top left, as Excel puts it.

# Examples
```julia
julia> XLSX.addShape(sheet, "H2:J5", :roundRect; text="Start here", horizontal=:center, vertical=:middle)

julia> XLSX.addShape(sheet, "H8", :rightArrow; fill="green", line=:none)
```

See also [`getShapes`](@ref) and [`deleteShape`](@ref).
"""
function addShape(ws::Worksheet, ref, kind::Symbol; text::AbstractString="", kw...)
    haskey(SHAPE_KINDS, kind) ||
        throw(XLSXError("Unknown shape `:$kind`. Must be one of $(join((":$k" for k in keys(SHAPE_KINDS)), ", "))."))
    return _add_sp!(ws, ref, kind, text; kw...)
end

const _SP_KEYWORDS = (:size, :font, :fill, :line, :lineWidth, :lineDash, :horizontal, :vertical, :name)

# A text box (`kind` nothing) or preset shape, anchored over `ref`.
function _add_sp!(ws::Worksheet, ref, kind::Union{Nothing,Symbol}, text::AbstractString; kw...)
    what = isnothing(kind) ? "add a text box" : "add a shape"
    _require_writable(ws, what)
    _require_worksheet(ws, what)
    for k in keys(kw)
        k in _SP_KEYWORDS || throw(XLSXError("Unknown keyword `$k`. Must be one of $(join(_SP_KEYWORDS, ", "))."))
    end
    rng = _sp_range(ref, get(kw, :size, _SHAPE_DEFAULT_SIZE))
    sp_pr_fill = _sp_fill_node(get(kw, :fill, nothing))
    # a text box is white by default, which Excel writes out; a shape takes its fill from its style
    isnothing(sp_pr_fill) && isnothing(kind) && (sp_pr_fill = Charts._solid_fill_node(XML.Element("a:schemeClr"; val="lt1"), _DML_A))
    sp_pr_line = _sp_line_node(get(kw, :line, nothing), get(kw, :lineWidth, nothing), get(kw, :lineDash, nothing),
                               isnothing(kind))
    body = _sp_text_body(text, get(kw, :font, (;)), get(kw, :horizontal, nothing), get(kw, :vertical, nothing), kind)
    name = get(kw, :name, nothing)
    isnothing(name) || (name isa AbstractString && !isempty(name)) ||
        throw(XLSXError("`name` must be a non-empty string. Found `$(repr(name))`."))

    xf = get_xlsxfile(ws)
    drawing_path = ensure_drawing!(xf, _sheet_path(ws))
    root = xml_root_element(xf.data[drawing_path])
    id = _next_shape_id(root)
    name = something(name, (isnothing(kind) ? "TextBox" : SHAPE_KINDS[kind]) * " $(id - 1)")

    c0, r0 = Int(rng.start.column_number), Int(rng.start.row_number)
    c1, r1 = Int(rng.stop.column_number), Int(rng.stop.row_number)
    emu(twips) = twips * 635
    x, y = emu(_cols_to_twips(ws, c0 - 1)), emu(_rows_to_twips(ws, r0 - 1))
    cx, cy = emu(_cols_to_twips(ws, c1)) - x, emu(_rows_to_twips(ws, r1)) - y
    sppr_kids = XML.Node[XML.Element("a:xfrm", XML.Element("a:off"; x=string(x), y=string(y)),
                                     XML.Element("a:ext"; cx=string(cx), cy=string(cy))),
                         XML.Element("a:prstGeom", XML.Element("a:avLst"); prst=String(something(kind, :rect)))]
    isnothing(sp_pr_fill) || push!(sppr_kids, sp_pr_fill)
    isnothing(sp_pr_line) || push!(sppr_kids, sp_pr_line)
    cnvsppr = isnothing(kind) ? XML.Element("xdr:cNvSpPr"; txBox="1") : XML.Element("xdr:cNvSpPr")
    sp = XML.Element("xdr:sp",
        XML.Element("xdr:nvSpPr", XML.Element("xdr:cNvPr"; id=string(id), name=String(name)), cnvsppr),
        XML.Element("xdr:spPr", sppr_kids...),
        _sp_style(isnothing(kind)),
        body; var"macro"="", textlink="")
    push!(root, build_two_cell_anchor(c0 - 1, r0 - 1, c1, r1, sp))
    return (sheet = ws.name, name = String(name), from = string(rng.start), to = string(rng.stop))
end

# The cells the object covers: a range, or from a cell for `size` pixels, snapped to
# whole default-sized cells as an image's are.
function _sp_range(ref, size)::CellRange
    ref isa CellRange && return ref
    if ref isa AbstractString
        is_valid_cellrange(ref) && return CellRange(ref)
        is_valid_cellname(ref) || throw(XLSXError("Give a cell, such as `B2`, or a range, such as `B2:E6`. Found `$ref`."))
        ref = CellRef(ref)
    end
    ref isa CellRef || throw(XLSXError("Give a cell or a range for the object. Found `$(repr(ref))`."))
    (size isa Tuple{Integer,Integer} && all(>(0), size)) ||
        throw(XLSXError("`size` must be a `(width, height)` in pixels. Found `$(repr(size))`."))
    w, h = size
    stop = CellRef(ref.row_number + max(1, round(Int, h / 20)) - 1, ref.column_number + max(1, round(Int, w / 64)) - 1)
    return CellRange(ref, stop)
end

_sp_fill_node(::Nothing) = nothing
_sp_fill_node(fill) = fill === :none ? XML.Element("a:noFill") :
    Charts._solid_fill_node(Charts._color_node(_checked_color(fill), _DML_A), _DML_A)

_checked_color(c) = (_check_color(c); c)

# A text box has a thin grey border by default, as Excel draws one; a shape's comes from its style.
function _sp_line_node(color, width, dash, textbox::Bool)
    all(isnothing, (color, width, dash)) && !textbox && return nothing
    color === :none && return XML.Element("a:ln", XML.Element("a:noFill"))
    w = something(width, 0.75)
    (w isa Real && isfinite(w) && 0 <= w <= 1584) || throw(XLSXError("`lineWidth` must be a width in points, from 0 to 1584. Found `$(repr(width))`."))
    kids = XML.Node[]
    if isnothing(color)
        textbox ? push!(kids, Charts._solid_fill_node(XML.Element("a:schemeClr", XML.Element("a:shade"; val="50000"); val="lt1"), _DML_A)) :
                  push!(kids, Charts._solid_fill_node(XML.Element("a:schemeClr", XML.Element("a:shade"; val="15000"); val="accent1"), _DML_A))
    else
        push!(kids, Charts._solid_fill_node(Charts._color_node(_checked_color(color), _DML_A), _DML_A))
    end
    if !isnothing(dash)
        dash isa Symbol && dash in Charts.DASH_VALUES ||
            throw(XLSXError("Invalid `lineDash` `$(repr(dash))`. Must be one of $(join((":$d" for d in Charts.DASH_VALUES), ", "))."))
        push!(kids, XML.Element("a:prstDash"; val=String(dash)))
    end
    return XML.Element("a:ln", kids...; w=string(round(Int, w * 12700)), cmpd="sng")
end

# Excel's style: a text box refers to nothing in the theme, a shape takes the first accent colour.
function _sp_style(textbox::Bool)
    zero(tag) = XML.Element(tag, XML.Element("a:scrgbClr"; r="0", g="0", b="0"); idx="0")
    accent(tag, idx; shade=nothing) = XML.Element(tag, isnothing(shade) ? XML.Element("a:schemeClr"; val="accent1") :
        XML.Element("a:schemeClr", XML.Element("a:shade"; val=shade); val="accent1"); idx=idx)
    textbox && return XML.Element("xdr:style", zero("a:lnRef"), zero("a:fillRef"), zero("a:effectRef"),
                                  XML.Element("a:fontRef", XML.Element("a:schemeClr"; val="dk1"); idx="minor"))
    return XML.Element("xdr:style", accent("a:lnRef", "2"; shade="15000"), accent("a:fillRef", "1"), accent("a:effectRef", "0"),
                       XML.Element("a:fontRef", XML.Element("a:schemeClr"; val="lt1"); idx="minor"))
end

# The `<xdr:txBody>`: one paragraph for each line of the text, its runs in their fonts.
function _sp_text_body(text::AbstractString, font, horizontal, vertical, kind)
    font isa NamedTuple || throw(XLSXError("`font` must be a NamedTuple such as `(size=14, bold=true)`. Found `$(repr(font))`."))
    for k in keys(font)
        k in _SHAPE_FONT_KEYS || throw(XLSXError("Unknown font setting `$k`. Must be one of $(join(_SHAPE_FONT_KEYS, ", "))."))
    end
    algn = isnothing(horizontal) ? (isnothing(kind) ? nothing : "l") :
        haskey(_SHAPE_HALIGN, horizontal) ? _SHAPE_HALIGN[horizontal] :
        throw(XLSXError("Invalid `horizontal` `$(repr(horizontal))`. Must be one of $(join((":$k" for k in keys(_SHAPE_HALIGN)), ", "))."))
    anchor = isnothing(vertical) ? "t" : haskey(_SHAPE_VALIGN, vertical) ? _SHAPE_VALIGN[vertical] :
        throw(XLSXError("Invalid `vertical` `$(repr(vertical))`. Must be one of $(join((":$k" for k in keys(_SHAPE_VALIGN)), ", "))."))

    runs = text isa RichTextString ? [(r.text, something(r.atts, Dict{Symbol,Any}())) for r in text.runs] :
                                     [(String(text), Dict{Symbol,Any}())]
    paras = [Tuple{String,Dict{Symbol,Any}}[]]
    for (t, atts) in runs
        for (i, piece) in enumerate(split(replace(t, "\r\n" => "\n", '\r' => '\n'), '\n'))
            i > 1 && push!(paras, Tuple{String,Dict{Symbol,Any}}[])
            isempty(piece) || push!(last(paras), (String(piece), atts))
        end
    end
    ppr = isnothing(algn) ? nothing : XML.Element("a:pPr"; algn=algn)
    pnodes = XML.Node[]
    for p in paras
        kids = XML.Node[]
        isnothing(ppr) || push!(kids, ppr)
        for (t, atts) in p
            push!(kids, XML.Element("a:r", _sp_run_props("a:rPr", merge(Dict{Symbol,Any}(pairs(font)), atts)), XML.Element("a:t", XML.Text(t))))
        end
        isempty(p) && push!(kids, _sp_run_props("a:endParaRPr", Dict{Symbol,Any}(pairs(font))))
        push!(pnodes, XML.Element("a:p", kids...))
    end
    battrs = ["vertOverflow" => "clip", "horzOverflow" => "clip"]
    isnothing(kind) && push!(battrs, "vert" => "horz")
    push!(battrs, "rtlCol" => "0", "anchor" => anchor)
    return XML.Element("xdr:txBody", _xml_element("a:bodyPr", battrs), XML.Element("a:lstStyle"), pnodes...)
end

# A run's `<a:rPr>` (or `<a:endParaRPr>`) from `setFont`'s settings, as Excel writes it.
function _sp_run_props(tag::String, atts::Dict{Symbol,Any})
    size = get(atts, :size, 11)
    (size isa Real && 1 <= size <= 4000) || throw(XLSXError("Invalid font `size` $(repr(size)). Must be from 1 to 4000 points."))
    vert = get(atts, :vertAlign, nothing)
    baseline = vert in (nothing, "baseline") ? nothing : vert == "superscript" ? 0.3 : vert == "subscript" ? -0.25 :
        throw(XLSXError("Invalid `vertAlign` `$(repr(vert))`. Must be \"superscript\", \"subscript\" or \"baseline\"."))
    # As in a cell: `under` is "single", "double" or "none" (or a Bool in rich text), and `strike` a Bool.
    under = get(atts, :under, nothing)
    under = under === true ? "single" : under in (false, "none") ? nothing : under
    (isnothing(under) || under in ("single", "double")) ||
        throw(XLSXError("Invalid `under` `$(repr(under))`. Must be \"single\", \"double\" or \"none\"."))
    strike = get(atts, :strike, nothing)
    (isnothing(strike) || strike isa Bool) || throw(XLSXError("`strike` must be `true` or `false`. Found `$(repr(strike))`."))
    bold, italic = get(atts, :bold, nothing), get(atts, :italic, nothing)
    for (k, v) in ((:bold, bold), (:italic, italic))
        (isnothing(v) || v isa Bool) || throw(XLSXError("`$k` must be `true` or `false`. Found `$(repr(v))`."))
    end
    name = get(atts, :name, nothing)
    (isnothing(name) || name isa AbstractString) || throw(XLSXError("The font `name` must be a string. Found `$(repr(name))`."))
    rp = Charts.DrawingRunProps(lang = "en-US", size = Float64(size), bold = bold, italic = italic,
                                under = isnothing(under) ? nothing : under == "single" ? "sng" : "dbl",
                                strike = strike === true ? "sngStrike" : nothing, baseline = baseline,
                                latin = isnothing(name) ? nothing : String(name))
    node = Charts._run_props_node(rp, tag[3:end], _DML_A)
    color = get(atts, :color, nothing)
    isnothing(color) || (node = Charts.insert_child(node, (NS_A, tag[3:end]),
        Charts._solid_fill_node(Charts._color_node(_checked_color(color), _DML_A), _DML_A)))
    return node
end

"""
    getShapes(ws::Worksheet) -> Vector{NamedTuple}

The text boxes and shapes on a worksheet, each a `NamedTuple` with fields `name`, `kind`
(`:textbox`, or the shape's kind, such as `:rect` or `:ellipse`), `from` and `to` (the
cells at its corners) and `text` (its text, with a `"\\n"` between paragraphs).

Images and charts aren't shapes: see [`getImages`](@ref) and
[`getCharts`](@ref XLSX.Charts.getCharts).

Works on files opened for reading only.
"""
function getShapes(ws::Worksheet)
    xf = get_xlsxfile(ws)
    drawing_path = _drawing_path_for_sheet(xf, _sheet_path(ws))
    out = NamedTuple{(:name, :kind, :from, :to, :text),Tuple{String,Symbol,String,String,String}}[]
    (isnothing(drawing_path) || !haskey(xf.data, drawing_path)) && return out
    for a in xml_elements(xml_root_element(get_xml_data(xf, drawing_path)))
        localname(a) == "twoCellAnchor" || continue
        sp = first_element_with_tag(a, "sp")
        isnothing(sp) && continue
        push!(out, _shape_info(a, sp))
    end
    return out
end

function _shape_info(a::XML.Node, sp::XML.Node)
    nv = first_element_with_tag(sp, "nvSpPr")
    cnv = isnothing(nv) ? nothing : first_element_with_tag(nv, "cNvPr")
    csp = isnothing(nv) ? nothing : first_element_with_tag(nv, "cNvSpPr")
    geom = first_element_with_tag(something(first_element_with_tag(sp, "spPr"), sp), "prstGeom")
    kind = !isnothing(csp) && get(csp, "txBox", "0") in ("1", "true") ? :textbox :
           isnothing(geom) ? :custom : Symbol(get(geom, "prst", "rect"))
    body = first_element_with_tag(sp, "txBody")
    text = isnothing(body) ? "" :
        join((join((_element_text(t) for r in xml_elements(p) if localname(r) in ("r", "fld")
                    for t in xml_elements(r) if localname(t) == "t"), "") for p in xml_elements(body) if localname(p) == "p"), "\n")
    return (name = isnothing(cnv) ? "" : get(cnv, "name", ""), kind = kind,
            from = something(_parse_cell_marker(a, "from"; is_to=false), ""),
            to = something(_parse_cell_marker(a, "to"; is_to=true), ""), text = text)
end

"""
    deleteShape(ws::Worksheet, ref) -> Int

Delete the text boxes and shapes whose top-left corner is in cell `ref` (`"B2"`, or a
`CellRef`), or the one named `ref`, as [`getShapes`](@ref) gives the names. Returns how
many were deleted. Images and charts are left alone.

The file must be open for writing (`mode="rw"`).
"""
function deleteShape(ws::Worksheet, ref::Union{AbstractString,CellRef})::Int
    _require_writable(ws, "delete a shape")
    _require_worksheet(ws, "delete a shape")
    xf = get_xlsxfile(ws)
    sheet_path = _sheet_path(ws)
    drawing_path = _drawing_path_for_sheet(xf, sheet_path)
    isnothing(drawing_path) && return 0
    bycell = ref isa CellRef || is_valid_cellname(ref)
    target = bycell ? string(ref isa CellRef ? ref : CellRef(ref)) : String(ref)
    return _delete_drawing_objects!(xf, sheet_path, drawing_path, a -> begin
        localname(a) == "twoCellAnchor" || return false
        sp = first_element_with_tag(a, "sp")
        isnothing(sp) && return false
        info = _shape_info(a, sp)
        bycell ? info.from == target : info.name == target
    end)
end
