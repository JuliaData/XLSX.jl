#
# creation.jl
#


# The xdr:graphicFrame for a chart, as Excel writes it: zero xfrm (the anchor
# carries the position), and for cx the frame wrapped in mc:AlternateContent with
# no Fallback, which Excel accepts. `locked` adds the graphicFrameLocks Excel
# writes for a chartsheet's chart.
function _chart_frame(rid::String, shape_id::Int, name::String;
                      cx_requires::Union{Nothing,String} = nothing,
                      locked::Bool = false)
    iscx = !isnothing(cx_requires)
    chart = XML.Element(iscx ? "cx:chart" : "c:chart")
    chart[iscx ? "xmlns:cx" : "xmlns:c"] = iscx ? NS_CX : NS_C
    chart["xmlns:r"] = NS_R
    chart["r:id"]    = rid

    data = XML.Element("a:graphicData", chart)
    data["uri"] = iscx ? "http://schemas.microsoft.com/office/drawing/2014/chartex" :
                         "http://schemas.openxmlformats.org/drawingml/2006/chart"

    cnv_frame = locked ?
        XML.Element("xdr:cNvGraphicFramePr", XML.Element("a:graphicFrameLocks"; noGrp = "1")) :
        XML.Element("xdr:cNvGraphicFramePr")

    frame = XML.Element("xdr:graphicFrame",
        XML.Element("xdr:nvGraphicFramePr",
            XML.Element("xdr:cNvPr"; id = string(shape_id), name = name),
            cnv_frame),
        XML.Element("xdr:xfrm",
            XML.Element("a:off"; x = "0", y = "0"),
            XML.Element("a:ext"; cx = "0", cy = "0")),
        XML.Element("a:graphic", data))
    frame["macro"] = ""               # `macro` is a Julia keyword, so not a kwarg

    iscx || return frame
    choice = XML.Element("mc:Choice", frame)
    choice["xmlns:$cx_requires"] = CX_NAMESPACES[cx_requires]
    choice["Requires"] = cx_requires
    mc = XML.Element("mc:AlternateContent", choice)
    mc["xmlns:mc"] = NS_MC
    return mc
end

# Add a new part at the next free `stemN.xml` in `dir`, from `xml` (a string, or an
# already-parsed document), with its content-type override. Returns its path.
function _new_part!(xf::XLSXFile, dir::String, first_name::String,
                    xml::Union{AbstractString,XML.Node}, mime::String)::String
    path = "$dir/$(_next_part_name(xf, dir, first_name))"
    xf.data[path]  = xml isa XML.Node ? xml : parse(xml, XML.Node)
    xf.files[path] = true
    register_content_type!(xf, "[Content_Types].xml";
                           tag="Override", key="PartName", val="/$path", content_type=mime)
    return path
end

"""
    _add_chart_part!(ws, xml, style_id; anchor, cx_requires=nothing) -> String

Add `xml` as a new chart part on `ws`, with its style and colour parts, and place
it in the sheet's drawing. `anchor` is a cell range for a worksheet, or `nothing`
for a chartsheet, whose chart is absolutely anchored and fills the sheet.
`cx_requires` is `nothing` for a `c:` chart, or the `Requires` prefix (`"cx1"`,
`"cx2"`) for a `cx:` one. Returns the chart part's path.

This writes the part as given: it neither checks nor rewrites what the XML
refers to.
"""
function _add_chart_part!(ws::Worksheet, xml::Union{AbstractString,XML.Node}, style_id::Integer;
                          anchor::Union{Nothing,CellRange},
                          cx_requires::Union{Nothing,String} = nothing)::String
    xf   = get_xlsxfile(ws)
    iscx = !isnothing(cx_requires)
    drawing = ensure_drawing!(xf, get_worksheet_internal_file(ws))

    path  = _new_part!(xf, "xl/charts", iscx ? "chartEx1.xml" : "chart1.xml", xml,
                       iscx ? MIME_CHARTEX : MIME_CHART)
    spath = _new_part!(xf, "xl/charts", "style1.xml",  CHART_STYLE_TEMPLATES[style_id], MIME_CHART_STYLE)
    cpath = _new_part!(xf, "xl/charts", "colors1.xml", CHART_COLOR_TEMPLATES[10],       MIME_CHART_COLORS)
    add_part_rel!(xf, path, cpath, REL_CHART_COLORS)
    add_part_rel!(xf, path, spath, REL_CHART_STYLE)

    rid  = add_part_rel!(xf, drawing, path, iscx ? REL_CHARTEX : REL_CHART)
    root = xml_root_element(xf.data[drawing])
    id   = _next_shape_id(root)
    content = _chart_frame(rid, id, "Chart $(id - 1)"; cx_requires, locked = isnothing(anchor))

    push!(root, isnothing(anchor) ?
        build_absolute_anchor(content) :
        build_two_cell_anchor(
            column_number(anchor.start) - 1, row_number(anchor.start) - 1,   # 0-based inclusive
            column_number(anchor.stop),      row_number(anchor.stop),        # 0-based exclusive
            content))
    return path
end

# Rebuild `node` bottom-up, replacing each element `e` with `f(e)` once its
# children have been rebuilt.
function _rewrite(f, node::XML.Node)
    kids = XML.children(node)
    if !isnothing(kids) && !isempty(kids)
        node = _with_children(node, [_rewrite(f, k) for k in kids])
    end
    return XML.nodetype(node) == XML.Element ? f(node) : node
end

_is_el(k, tag) = XML.nodetype(k) == XML.Element && localname(k) == tag

# The template with its series removed and fresh axis ids, as a document node.
function _chart_shell(template::AbstractString; title)
    ids  = Dict{String,String}()
    next = Ref(500_000_000)
    newid(v) = get!(() -> string(next[] += 1), ids, v)

    doc = _rewrite(parse(template, XML.Node)) do e
        t = localname(e)
        t in CHART_GROUP_TAGS && return _with_children(e, [k for k in e.children if !_is_el(k, "ser")])
        t in ("axId", "crossAx") && return with_attribute(e, "val", newid(_attr(e, "val")))
        return e
    end
    isnothing(title) && return doc

    return _rewrite(doc) do e
        localname(e) == "chart" || return e
        t = first_element_with_tag(e, "title")
        if title === false
            e   = remove_child(e, "title")
            atd = first_element_with_tag(e, "autoTitleDeleted")
            return replace_child(e, atd, with_attribute(atd, "val", "1"))
        end
        return replace_child(e, t, _title_with_text(t, title))
    end
end

# A typed title in the template's formatting: c:txPr's body, with its end-of-
# paragraph properties replaced by a run.
function _title_with_text(t::XML.Node, text::AbstractString)
    txpr = first_element_with_tag(t, "txPr")
    p    = first_element_with_tag(txpr, "p")
    run  = XML.Element("a:r", XML.Element("a:rPr"; lang = "en-US"),
                              XML.Element("a:t", XML.Text(text)))
    pk   = [k for k in p.children if !_is_el(k, "endParaRPr")]
    push!(pk, run)
    newp = _with_children(p, pk)
    rich = XML.Element("c:rich", [k === p ? newp : k for k in txpr.children]...)
    return insert_child(t, (NS_C, "title"), XML.Element("c:tx", rich))
end

"""
    addChart(ws::Worksheet, kind::Symbol; anchor, title=nothing) -> Chart

Add an empty chart of `kind` to `ws`, covering the cells in `anchor` (a range
such as `"F2:M18"`). Add data with [`addSeries`](@ref).

`kind` is one of `:column`, `:bar`, `:stackedColumn`, `:stackedBar`, `:line`,
`:lineMarkers`, `:area`, `:pie`, `:doughnut`, `:scatter`, `:bubble` or `:radar`.
`:bar` and `:stackedBar` are horizontal; `:column` and `:stackedColumn` vertical.
The chart starts as Excel's own default for that kind, formatted by the
workbook's theme.

`title` is `nothing` for Excel's automatic title, a string for typed text, or
`false` for no title.
"""
function addChart(ws::Worksheet, kind::Symbol;
                  anchor::Union{AbstractString,CellRange},
                  title::Union{Nothing,Bool,AbstractString} = nothing)::Chart
    haskey(C_KINDS, kind) || throw(XLSXError(
        "Unknown chart kind `:$kind`. Use one of " * join((":$k" for k in keys(C_KINDS)), ", ") * "."))
    title === true && throw(XLSXError("`title` is a string, `nothing` for Excel's automatic title, or `false` for none."))
    k    = C_KINDS[kind]
    path = _add_chart_part!(ws, _chart_shell(CHART_KIND_TEMPLATES[k.template]; title), k.style;
                            anchor = anchor isa CellRange ? anchor : CellRange(anchor))
    return getChart(ws, path)
end

const CHARTSHEET_TEMPLATE = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<chartsheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="xr" xmlns:xr="http://schemas.microsoft.com/office/spreadsheetml/2014/revision"><sheetPr/><sheetViews><sheetView workbookViewId="0" zoomToFit="1"/></sheetViews><pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/></chartsheet>"""

# Excel's anchor for a chartsheet's chart; with zoomToFit the size is nominal.
build_absolute_anchor(content::XML.Node)::XML.Node =
    XML.Element("xdr:absoluteAnchor",
        XML.Element("xdr:pos"; x = "0", y = "0"),
        XML.Element("xdr:ext"; cx = "9295876", cy = "6068505"),
        content,
        XML.Element("xdr:clientData"))

        function _check_chart_args(kind::Symbol, title)
    haskey(C_KINDS, kind) || throw(XLSXError(
        "Unknown chart kind `:$kind`. Use one of " * join((":$k" for k in keys(C_KINDS)), ", ") * "."))
    title === true && throw(XLSXError(
        "`title` is a string, `nothing` for Excel's automatic title, or `false` for none."))
    return C_KINDS[kind]
end

"""
    addChart(xf::XLSXFile, kind::Symbol; sheetname="", title=nothing) -> Chart

Add an empty chart of `kind` on a new chartsheet, placed after the existing
sheets. `sheetname` defaults to `Chart1`, `Chart2`, … as Excel names them. See
the worksheet method for `kind` and `title`.
"""
function addChart(xf::XLSXFile, kind::Symbol; sheetname::AbstractString = "",
                  title::Union{Nothing,Bool,AbstractString} = nothing)::Chart
    k  = _check_chart_args(kind, title)
    ws = _register_sheet!(get_workbook(xf), parse(CHARTSHEET_TEMPLATE, XML.Node), sheetname;
                          dir = "xl/chartsheets", reltype = REL_CHARTSHEET,
                          mime = MIME_CHARTSHEET, default_name = "Chart")
    path = _add_chart_part!(ws, _chart_shell(CHART_KIND_TEMPLATES[k.template]; title), k.style;
                            anchor = nothing)
    return getChart(xf, path)
end

# The c:f text for a resolved reference: sheet-qualified and absolute, as a chart
# part requires. A Worksheet supplies the sheet for a range that names none.
_chart_f(ws::Worksheet, r::CellRef)   = quoteit(ws.name) * "!" * abscell(r)
_chart_f(ws::Worksheet, r::CellRange) = quoteit(ws.name) * "!" * abscell(r.start) * ":" * abscell(r.stop)
_chart_f(::Union{XLSXFile,Worksheet}, r::SheetCellRef)   = quoteit(r.sheet) * "!" * mkabs(r)
_chart_f(::Union{XLSXFile,Worksheet}, r::SheetCellRange) = quoteit(r.sheet) * "!" * mkabs(r)

_chart_f(::Union{XLSXFile,Worksheet}, r::NonContiguousRange) =
    join((quoteit(r.sheet) * "!" *
          (x isa CellRange ? abscell(x.start) * ":" * abscell(x.stop) : abscell(x))
          for x in r.rng), ",")

_chart_f(::Union{XLSXFile,Worksheet}, r) = throw(XLSXError(
    "A chart series cannot reference a $(typeof(r)). Use a cell range such as " *
    "`\"B2:B5\"`, or a defined name instead."))

# The cached values for a reference, flattened in Excel's order: down a column,
# then across. A single cell gives one value; a NonContiguousRange gives one
# block per part, in order.
function _chart_values(x::Union{XLSXFile,Worksheet}, r)
    d = getdata(x, r)
    d isa AbstractMatrix && return vec(permutedims(d))
    d isa AbstractVector && return reduce(vcat, (vec(permutedims(b)) for b in d); init = Any[])
    return Any[d]
end

# Resolve `ref` against sheet `sheet` of `xf`, or against the workbook when `sheet`
# is nothing (a chartsheet). Returns the c:f / cx:f text and the cell values.
function _resolve_ref(xf::XLSXFile, sheet::Union{Nothing,String}, ref)
    if !(ref isa AbstractString)
        x = isnothing(sheet) ? xf : getsheet(xf, sheet)
        return (_chart_f(x, ref), _chart_values(x, ref))
    end
    g(x, r) = (_chart_f(x, r), _chart_values(x, r))
    out = isnothing(sheet) ? ref_chooser(g, xf, ref) : ref_chooser(g, getsheet(xf, sheet), ref)
    out isa Tuple || throw(XLSXError(
        "`$ref` is a defined name holding a constant, not a cell range, so a chart " *
        "series cannot plot it."))
    return out
end

# A chart on a chartsheet has no cells of its own, so its references resolve
# through the workbook: `c.sheet` names the chartsheet, not a source of data.
_data_sheet(c::AbstractChart) =
    isnothing(c.sheet) || is_chartsheet(get_workbook(c.package), c.sheet) ? nothing : c.sheet

# A bare reference (`B2:B5`, `$B$2:$B$5`, `B:B`, `2:5`) names no sheet. On a
# chartsheet that has nothing to mean, since the chartsheet holds no cells.
_is_unqualified_ref(ref) = ref isa AbstractString &&
    (r = replace(ref, '$' => "");
     is_valid_cellname(r) || is_valid_cellrange(r) || is_valid_column_range(r) || is_valid_row_range(r))

function _require_qualified(refs...)
    for ref in refs
        _is_unqualified_ref(ref) && throw(XLSXError(
            "`$ref` names no sheet. A chart on a chartsheet has no cells of its own, " *
            "so its series must name the sheet holding the data, as in `\"Sheet1!$ref\"`."))
    end
    return nothing
end

function _series_ref(c::AbstractChart, ref)
    sheet = _data_sheet(c)
    isnothing(sheet) && _require_qualified(ref)
    return _resolve_ref(c.package, sheet, ref)
end


_chart_numeric(v) = v isa Real || v isa Dates.Date || v isa Dates.DateTime || v isa Dates.Time

# (ref tag, cache tag, formatter). Excel decides by content: a numeric series
# name still gets a numRef.
function _cache_kind(vals::Vector, xf::XLSXFile)
    vs = filter(!ismissing, vals)
    isempty(vs) || all(_chart_numeric, vs) || return ("strRef", "strCache", string)
    d1904 = isdate1904(xf)
    fmt(v) = v isa Dates.Date || v isa Dates.DateTime ? string(date_to_excel_value(v, d1904)) :
             v isa Dates.Time  ? string(time_to_excel_value(v)) :
             v isa Integer     ? string(v) : string(convert(Float64, v))
    return ("numRef", "numCache", fmt)
end

# <c:numRef><c:f>…</c:f><c:numCache>…</c:numCache></c:numRef>. Missing values are
# omitted, as Excel does; ptCount still spans the whole range.
function _cache_ref(f::AbstractString, vals::Vector, xf::XLSXFile)
    reftag, cachetag, fmt = _cache_kind(vals, xf)
    cache = XML.Node[]
    reftag == "numRef" && push!(cache, XML.Element("c:formatCode", XML.Text("General")))
    push!(cache, XML.Element("c:ptCount"; val = string(length(vals))))
    for (i, v) in enumerate(vals)
        ismissing(v) && continue
        push!(cache, XML.Element("c:pt", XML.Element("c:v", XML.Text(fmt(v))); idx = string(i - 1)))
    end
    return XML.Element("c:$reftag", XML.Element("c:f", XML.Text(f)), XML.Element("c:$cachetag", cache...))
end

# A data child: <c:val><c:numRef>…</c:numRef></c:val>, and its value count.
function _data_child(c::Chart, tag::AbstractString, ref)
    f, vals = _series_ref(c, ref)
    return XML.Element("c:$tag", _cache_ref(f, vals, c.package)), length(vals)
end

"""
    _chart_kind(c::Chart) -> Symbol

The API kind of `c`, read from its part: the plot group's tag plus the
attributes that distinguish kinds sharing one tag. Works on any chart, whatever
wrote it. A `lineChart` reports `:lineMarkers` only if an existing series has a
real marker symbol, so an empty one reads as `:line`.
"""
function _chart_kind(c::Chart)::Symbol
    groups = _group_nodes(chart_root(c))
    length(groups) == 1 || throw(XLSXError(
        isempty(groups) ? "Chart `$(c.name)` has no plot group." :
        "Chart `$(c.name)` is a combo chart, with $(length(groups)) plot groups. " *
        "Adding a series to one is not supported."))
    g = only(groups)
    tag = localname(g)

    if tag == "barChart"
        stacked = _sym_val(g, "grouping") in (:stacked, :percentStacked)
        _sym_val(g, "barDir") === :bar && return :bar
        return stacked ? :stackedColumn : :column
    elseif tag == "lineChart"
        return _bool_val(g, "marker") === true ? :lineMarkers : :line
    end
    for (kind, k) in pairs(C_KINDS)
        localname(first(_group_nodes(xml_root_element(
            parse(CHART_KIND_TEMPLATES[k.template], XML.Node))))) == tag && return kind
    end
    throw(XLSXError("Adding a series to a `$tag` chart is not supported."))
end

const _SER_DATA_TAGS = ("idx", "order", "tx", "cat", "val", "xVal", "yVal", "bubbleSize", "extLst")

_accent(i::Integer) = "accent" * string(mod1(i, 6))


# A scatter series joined by a line: the template's noFill a:ln replaced with a
# solid accent1 line, which _with_accent then retargets. 1.5 pt, as Excel's
# Scatter with Straight Lines writes.
function _with_scatter_line(sp::XML.Node)
    ln  = first_element_with_tag(sp, "ln")
    new = XML.Element("a:ln",
              XML.Element("a:solidFill", XML.Element("a:schemeClr"; val = "accent1")),
              XML.Element("a:round"); w = "19050", cap = "rnd")
    return isnothing(ln) ? insert_child(sp, (NS_A, "spPr"), new) : replace_child(sp, ln, new)
end

const _SCHEME_NAMES = r"^(accent[1-6]|tx[12]|bg[12]|dk[12]|lt[12]|hlink|folHlink)$"

# The colour element for `color`, built as the series setters build it, with the
# template's child transforms (a bubble's a:alpha) carried across. A scheme name
# string becomes an a:schemeClr; anything else goes through the package's colour
# resolver, so it accepts exactly what setSeriesFill does.
function _series_color_node(color, template_clr::XML.Node, pfx)
    node = color isa SchemeColor ? _scheme_color_node(color, pfx) :
           color isa AbstractString && occursin(_SCHEME_NAMES, color) ?
               XML.Element(prefixed_tag(pfx[NS_A], "schemeClr"); val = color) :
               _srgb_color_node(color isa Symbol ? String(color) : color, pfx)
    extra = XML.children(template_clr)
    (isnothing(extra) || isempty(extra)) && return node
    return _with_children(node, vcat(something(XML.children(node), XML.Node[]), extra))
end

# Every accent1 in a template fragment, recoloured.
_with_color(n::XML.Node, color, pfx) =
    _rewrite(n) do e
        localname(e) == "schemeClr" && _attr(e, "val") == "accent1" || return e
        return _series_color_node(color, e, pfx)
    end

# The formatting children of a kind's template series: spPr, marker, smooth,
# invertIfNegative, bubble3D, and for pie and doughnut the per-point dPt list.
function _series_pattern(kind::Symbol)
    tmpl = parse(CHART_KIND_TEMPLATES[C_KINDS[kind].template], XML.Node)
    ser  = first(elements_with_tag(first(_group_nodes(xml_root_element(tmpl))), "ser"))
    return XML.Node[k for k in XML.eachelement(ser) if !(localname(k) in _SER_DATA_TAGS)]
end

_varies_by_point(kind::Symbol) = kind in (:pie, :doughnut)

# One c:dPt per point, cycling the accents, from the template's first dPt.
function _vary_color_points(proto::XML.Node, npts::Integer, pfx)
    proto = _with_children(proto, XML.Node[k for k in XML.eachelement(proto) if localname(k) != "extLst"])
    out = XML.Node[]
    for i in 1:npts
        dp  = _with_color(proto, _accent(i), pfx)
        idx = first_element_with_tag(dp, "idx")
        push!(out, replace_child(dp, idx, with_attribute(idx, "val", string(i - 1))))
    end
    return out
end

"""
    addSeries(c::Chart, values; categories=nothing, name=nothing, name_ref=nothing,
              bubble_sizes=nothing, markers=nothing, smooth=nothing, line=nothing,
              color=nothing) -> Chart

Add a series to `c`, plotting `values`. Each reference is a range string
(`"B2:B5"`), a defined name, or a reference object; an unqualified range means
the chart's own sheet, and a chart on a chartsheet must name one
(`"Data!B2:B5"`).

`categories` are the category labels, or the X values on a scatter or bubble
chart; a bubble chart requires them, and `bubble_sizes` too. `name` is literal
text for the series name, and `name_ref` a cell to take it from; pass one or
neither.

To plot a table column, pass gettablerange(t, column) for either `values` or `categories`

`markers` (line, lineMarkers, radar and scatter) is a symbol such as `:circle`,
or `false` for none. `smooth` (line, lineMarkers and scatter) curves the line.
`line` (scatter) joins the points. 

`color` overrides the series' automatic colour, and takes the same values as
[`setSeriesFill`](@ref): a colour name (`"red"` or `:red`), a hex colour
(`"FF0000"`, `"#FF0000"` or `"FFFF0000"`), a `Colors.Colorant`, a
[`SchemeColor`](@ref), or a scheme name such as `"accent4"`. It does not apply
to pie and doughnut charts, whose points are coloured individually.

The cells' current values are cached in the chart, as Excel does, so the series
reads back without opening the file in Excel. Series take the theme's accent
colours in turn; on a pie or doughnut chart each point does instead.
"""
function addSeries(c::Chart, values;
                   categories = nothing, name = nothing, name_ref = nothing,
                   bubble_sizes = nothing, markers = nothing, smooth = nothing,
                   line = nothing, color = nothing)::Chart
    isnothing(name) || isnothing(name_ref) ||
        throw(XLSXError("Pass `name` for literal text or `name_ref` for a cell, not both."))

    kind = _chart_kind(c)
    isxy = kind in (:scatter, :bubble)
    okmk = kind in (:line, :lineMarkers, :radar, :scatter)
    oksm = kind in (:line, :lineMarkers, :scatter)
    bad(opt, v, ok) = isnothing(v) || ok ||
        throw(XLSXError("`$opt` does not apply to a $kind chart."))
    bad("markers", markers, okmk)
    bad("smooth",  smooth,  oksm)
    bad("line",    line,    kind === :scatter)
    bad("bubble_sizes", bubble_sizes, kind === :bubble)
    bad("color", color, !_varies_by_point(kind))
    markers isa Union{Nothing,Bool,Symbol} ||
        throw(XLSXError("`markers` is a symbol such as `:circle`, or `true` or `false`."))
    smooth isa Union{Nothing,Bool} || throw(XLSXError("`smooth` is `true` or `false`."))
    line   isa Union{Nothing,Bool} || throw(XLSXError("`line` is `true` or `false`."))
    kind === :scatter && markers === false && line !== true && throw(XLSXError(
        "A scatter series with no markers and no line would draw nothing. " *
        "Pass `line = true` as well, or leave `markers` alone."))
    if kind === :bubble
        isnothing(bubble_sizes) && throw(XLSXError("A bubble chart series needs `bubble_sizes`."))
        isnothing(categories)   && throw(XLSXError("A bubble chart series needs `categories`, its X values."))
    end
    markers === true && (markers = :circle)
    kind === :line && !isnothing(markers) && markers !== false && (kind = :lineMarkers)

    root = chart_root(c)
    pfx  = ns_prefixes(root)
    g    = only(_group_nodes(root))
    sers = collect(elements_with_tag(g, "ser"))
    n    = length(sers)
    idx  = maximum((something(_int_val(s, "idx"),   -1) for s in sers); init = -1) + 1
    ord  = maximum((something(_int_val(s, "order"), -1) for s in sers); init = -1) + 1

    kids = XML.Node[XML.Element("c:idx"; val = string(idx)),
                    XML.Element("c:order"; val = string(ord))]
    isnothing(name) || push!(kids, XML.Element("c:tx", XML.Element("c:v", XML.Text(name))))
    if !isnothing(name_ref)
        nf, nv = _series_ref(c, name_ref)
        push!(kids, XML.Element("c:tx", _cache_ref(nf, nv, c.package)))
    end

    pattern = _series_pattern(kind)
    col     = something(color, _accent(n + 1))
    for k in pattern
        t = localname(k)
        t == "dPt" && continue                                    # rebuilt below
        if t == "marker" && !isnothing(markers)
            k = _marker_with_symbol(k, markers === false ? :none : markers, pfx)
        elseif t == "smooth" && !isnothing(smooth)
            k = with_attribute(k, "val", smooth ? "1" : "0")
        elseif t == "spPr" && line === true
            k = _with_scatter_line(k)
        end
        push!(kids, _varies_by_point(kind) ? k : _with_color(k, col, pfx))
    end

    val, npts = _data_child(c, isxy ? "yVal" : "val", values)
    if !isnothing(categories)
        cat, ncat = _data_child(c, isxy ? "xVal" : "cat", categories)
        push!(kids, cat)
        npts = max(npts, ncat)
    end
    push!(kids, val)
    isnothing(bubble_sizes) || push!(kids, first(_data_child(c, "bubbleSize", bubble_sizes)))

    if _varies_by_point(kind)
        proto = pattern[findfirst(k -> localname(k) == "dPt", pattern)]
        append!(kids, _vary_color_points(proto, npts, pfx))
    end

    ser  = XML.Element("c:ser", kids...)
    gtag = String(localname(g))
    new  = rebuild_path(root,
               [(NS_C, "chart")    => "chart",
                (NS_C, "plotArea") => "plotArea",
                (NS_C, gtag)       => (gtag, x -> x === g)],
               grp -> insert_child(grp, (NS_C, gtag), ser);
               prefixes = pfx)
    set_chart_root!(c, new)
    return c
end

# The layout's template with its data blocks and series removed. Not valid on its
# own (CT_ChartData requires a data block), so never written without a series.
function _chartex_shell(kind::Symbol; title = nothing)
    doc = _rewrite(parse(CHARTEX_KIND_TEMPLATES[kind], XML.Node)) do e
        t = localname(e)
        t == "chartData" &&
            return _with_children(e, XML.Node[k for k in XML.eachelement(e) if localname(k) != "data"])
        t == "plotAreaRegion" &&
            return _with_children(e, XML.Node[k for k in XML.eachelement(e) if localname(k) != "series"])
        t == "chart" && title === false && return remove_child(e, "title")
        return e
    end
    return doc
end

# Excel's default series for the layout, without its name, data link or identity:
# what remains is dataLabels, layoutPr and axisId. Pareto gives two: the column
# series and its owned paretoLine.
function _cx_series_pattern(kind::Symbol)
    root = xml_root_element(parse(CHARTEX_KIND_TEMPLATES[kind], XML.Node))
    clean(s) = _with_children(
        with_attribute(with_attribute(s, "uniqueId", nothing), "formatIdx", nothing),
        XML.Node[k for k in XML.eachelement(s) if !(localname(k) in ("tx", "dataId"))])
    sers = _cx_series_nodes(root)
    return kind === :pareto ? clean.(sers[1:2]) : [clean(first(sers))]
end

function _cx_data_node(id::Integer, cat_f, val_f::AbstractString, valtype::AbstractString)
    dims = XML.Node[]
    isnothing(cat_f) || push!(dims, XML.Element("cx:strDim", XML.Element("cx:f", XML.Text(cat_f)); type = "cat"))
    push!(dims, XML.Element("cx:numDim", XML.Element("cx:f", XML.Text(val_f)); type = valtype))
    return XML.Element("cx:data", dims...; id = string(id))
end

# cx:tx for a series name: txData with the reference and its cached value, or the
# value alone for literal text.
function _cx_name_tx(xf, sheet, name, name_ref)
    if !isnothing(name_ref)
        f, v = _resolve_ref(xf, sheet, name_ref)
        length(v) == 1 || throw(XLSXError("`name_ref` must be a single cell; `$name_ref` covers $(length(v))."))
        txt = ismissing(only(v)) ? "" : string(only(v))
        return XML.Element("cx:tx", XML.Element("cx:txData",
                   XML.Element("cx:f", XML.Text(f)), XML.Element("cx:v", XML.Text(txt))))
    end
    isnothing(name) && return nothing
    return XML.Element("cx:tx", XML.Element("cx:txData", XML.Element("cx:v", XML.Text(name))))
end

# Add a data block and its series to a cx:chartSpace root. Used both on a shell
# being built and on an existing chart.
function _cx_add_series(root::XML.Node, xf::XLSXFile, sheet, kind::Symbol, values;
                        categories = nothing, name = nothing, name_ref = nothing)
    haskey(_CX_SERIES_RULES, kind) || throw(XLSXError("Adding a series to a $kind chart is not supported."))
    rule = _CX_SERIES_RULES[kind]
    isnothing(name) || isnothing(name_ref) ||
        throw(XLSXError("Pass `name` for literal text or `name_ref` for a cell, not both."))
    rule.categories === :none && !isnothing(categories) &&
        throw(XLSXError("A $kind chart takes no categories."))
    rule.categories === :required && isnothing(categories) &&
        throw(XLSXError("A $kind chart needs `categories`: the columns of its hierarchy."))

    sers   = _cx_series_nodes(root)
    owners = count(s -> _attr(s, "layoutId") != "paretoLine", sers)
    rule.multi || owners == 0 ||
        throw(XLSXError("A $kind chart takes one series, and this one already has it."))

    datas = collect(elements_with_tag(first_element_with_tag(root, "chartData"), "data"))
    id    = maximum((something(tryparse(Int, something(_attr(d, "id"), "")), -1) for d in datas); init = -1) + 1
    vf    = first(_resolve_ref(xf, sheet, values))
    cf    = isnothing(categories) ? nothing : first(_resolve_ref(xf, sheet, categories))

    if owners > 0
        # A box & whisker chart's series share one category column: a later series
        # inherits the first's, and may not name different ones.
        firstid = _attr(first_element_with_tag(first(filter(s -> _attr(s, "layoutId") != "paretoLine", sers)), "dataId"), "val")
        fdata   = datas[findfirst(d -> _attr(d, "id") == firstid, datas)]
        fcat    = first_element_with_tag(fdata, "strDim")
        existing = isnothing(fcat) ? nothing : child_text(fcat, "f")
        if isnothing(cf)
            cf = existing
        elseif isnothing(existing) ||
               string(resolve_chartex_ref(xf, cf)) != string(resolve_chartex_ref(xf, existing))
            throw(XLSXError(
                "Every series in a $kind chart shares the first series' categories. " *
                (isnothing(existing) ? "The first series has none." :
                 "Omit `categories` to use them, or pass the same range.")))
        end
    end

    pattern = _cx_series_pattern(kind)
    main = pattern[1]
    tx   = _cx_name_tx(xf, sheet, name, name_ref)
    isnothing(tx) || (main = insert_child(main, (NS_CX, "series"), tx))
    main = insert_child(main, (NS_CX, "series"), XML.Element("cx:dataId"; val = string(id)))
    new_sers = XML.Node[main]
    kind === :pareto && push!(new_sers, with_attribute(pattern[2], "ownerIdx", string(length(sers))))

    pfx  = ns_prefixes(root)
    root = rebuild_path(root, [(NS_CX, "chartData") => "chartData"],
               cd -> insert_child(cd, (NS_CX, "chartData"), _cx_data_node(id, cf, vf, rule.valtype));
               prefixes = pfx, parent_key = _CX_ROOT_KEY)
    return rebuild_path(root,
               [(NS_CX, "chart") => "chart", (NS_CX, "plotArea") => "plotArea",
                (NS_CX, "plotAreaRegion") => "plotAreaRegion"],
               r -> foldl((acc, s) -> insert_child(acc, (NS_CX, "plotAreaRegion"), s), new_sers; init = r);
               prefixes = pfx, parent_key = _CX_ROOT_KEY)
end

"""
    addChartEx(ws::Worksheet, kind::Symbol, values; anchor, categories=nothing,
               name=nothing, name_ref=nothing, title=nothing) -> ChartEx
    addChartEx(xf::XLSXFile, kind::Symbol, values; sheetname="", …) -> ChartEx

Add a `chartEx` chart of `kind` — `:waterfall`, `:funnel`, `:treemap`,
`:sunburst`, `:histogram`, `:pareto` or `:boxWhisker` — plotting `values`, to a
worksheet or as a new chartsheet.

Unlike [`addChart`](@ref), this takes the first series' data: a `chartEx` part
must contain data, so an empty one cannot exist. A box & whisker chart takes
further series with [`addSeries`](@ref); every other layout takes one.

`categories` are the category labels; treemap and sunburst require them, as one
column per hierarchy level (`"A2:B8"`), and a histogram takes none. References,
`name`, `name_ref` and `title` are as for [`addChart`](@ref) and
[`addSeries`](@ref). The chart's appearance comes from Excel's style for the
layout.
"""
function addChartEx(ws::Worksheet, kind::Symbol, values; anchor::Union{AbstractString,CellRange},
                    categories = nothing, name = nothing, name_ref = nothing,
                    title::Union{Nothing,Bool,AbstractString} = nothing)::ChartEx
    doc = _chartex_doc(get_xlsxfile(ws), ws.name, kind, values; categories, name, name_ref, title)
    return _place_chartex!(ws, doc, kind, anchor isa CellRange ? anchor : CellRange(anchor), title)
end

function addChartEx(xf::XLSXFile, kind::Symbol, values; sheetname::AbstractString = "",
                    categories = nothing, name = nothing, name_ref = nothing,
                    title::Union{Nothing,Bool,AbstractString} = nothing)::ChartEx
    _require_qualified(values, categories, name_ref)        # a chartsheet holds no cells
    doc = _chartex_doc(xf, nothing, kind, values; categories, name, name_ref, title)
    ws  = _register_sheet!(get_workbook(xf), parse(CHARTSHEET_TEMPLATE, XML.Node), sheetname;
                           dir = "xl/chartsheets", reltype = REL_CHARTSHEET,
                           mime = MIME_CHARTSHEET, default_name = "Chart")
    return _place_chartex!(ws, doc, kind, nothing, title)
end

# The complete chartEx part: shell plus first series. Resolving the references
# here, before anything is added to the package, means a bad one leaves the
# workbook untouched.
function _chartex_doc(xf::XLSXFile, sheet, kind::Symbol, values;
                      categories, name, name_ref, title)
    haskey(CX_KINDS, kind) || throw(XLSXError(
        "Unknown chartEx kind `:$kind`. Use one of " * join((":$k" for k in keys(CX_KINDS)), ", ") * "."))
    title === true && throw(XLSXError("`title` is a string, `nothing` for the default, or `false` for none."))
    doc  = _chartex_shell(kind; title)
    root = xml_root_element(doc)
    return replace_child(doc, root,
               _cx_add_series(root, xf, sheet, kind, values; categories, name, name_ref))
end

function _place_chartex!(ws::Worksheet, doc::XML.Node, kind::Symbol, anchor, title)::ChartEx
    k    = CX_KINDS[kind]
    path = _add_chart_part!(ws, doc, k.style; anchor, cx_requires = k.requires)
    c    = getChart(ws, path)
    title isa AbstractString && setChartTitleText(c, title)
    return c
end

function addSeries(c::ChartEx, values; categories = nothing, name = nothing, name_ref = nothing)::ChartEx
    sheet = _data_sheet(c)                                  # nothing on a chartsheet
    isnothing(sheet) && _require_qualified(values, categories, name_ref)
    set_chart_root!(c, _cx_add_series(_cx_root(c), c.package, sheet, getChartType(c), values;
                                      categories, name, name_ref))
    return c
end