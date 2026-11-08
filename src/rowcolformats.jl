# ===========================================================================
# Formatting inserted rows and columns as Excel does
#
# Excel formats an inserted row like the row above it, and an inserted column like
# the column to its left: its default "Format Same As Above". Checked against Excel
# 365, whose before-and-after files are test/data/rowcol_inherit_*.xlsx and
# rowcol_style_*.xlsx:
#
# - each new cell takes its neighbour's style, but not its border: a cell in a new
#   row takes the border of its column's style, and one in a new column the border
#   of its row's style, or none. One left with no formatting at all is not written;
# - a new row takes its neighbour's height, row style (border removed) and outline
#   level, and a new column its width, style and outline level; neither is hidden,
#   whether or not the neighbour is, and a hidden column of width 0 passes on no
#   width;
# - a group a new row or column joins is no longer all hidden, so the collapsed
#   marker on its summary row or column goes;
# - a conditional format or data validation whose range ends at the neighbour grows
#   over the new rows (columns). Nothing else does: merged cells, tables, the
#   AutoFilter, hyperlinks and defined names stay as they were.
#
# Rows inserted at row 1, and columns at column A, have no neighbour and are left
# unformatted.
# ===========================================================================

function _inherit_formats!(ws::Worksheet, e::RowColEdit)
    (e.count > 0 && e.at > 1) || return nothing
    wb = get_workbook(ws)
    plain = styles_cell_xf(wb, 0)
    memo = Dict{Tuple{UInt32,String},UInt32}()
    # Style `s` with border `border` in place of its own, and 0 if that leaves no formatting.
    restyle(s::Integer, border::AbstractString = "0") = get!(memo, (UInt32(s), String(border))) do
        s == 0 && border == "0" && return UInt32(0)
        x = styles_cell_xf(wb, Int(s))
        get(x, "borderId", "0") == border || (x = copynode(x); x["borderId"] = border)
        x == plain && return UInt32(0)
        return UInt32(styles_add_cell_xf(wb, x).id)
    end
    news = e.at:min(e.at + e.count - 1, _edit_max(e))
    e.axis === :row ? _inherit_rows!(ws, e.at - 1, news, restyle) :
                      _inherit_columns!(ws, e.at - 1, news, restyle)
    _uncollapse_groups!(ws, e)
    _grow_format_ranges!(_sheet_root(ws), e)
    return nothing
end

_styled_empty_cell(r::Int, c::Int, s::UInt32) = Cell(CellRef(r, c), UInt64(0), s, UInt16(0), CT_EMPTY, false)

function _inherit_rows!(ws::Worksheet, src::Int, news::UnitRange{Int}, restyle::Function)
    cache = ws.cache
    colborder = _column_style_borders(ws)
    srccells = get(cache.cells, src, Dict{Int,Cell}())
    ht = get(cache.row_ht, src, nothing)
    attrs = isnothing(ws.unhandled_attributes) ? nothing : get(ws.unhandled_attributes, src, nothing)
    rowattrs = Dict{String,String}()
    if !isnothing(attrs)
        haskey(attrs, "outlineLevel") && (rowattrs["outlineLevel"] = attrs["outlineLevel"])
        s = haskey(attrs, "s") ? restyle(parse(UInt32, attrs["s"])) : UInt32(0)
        s == 0 || (rowattrs["s"] = string(s); rowattrs["customFormat"] = "1")
    end
    for r in news
        cells = Dict{Int,Cell}()
        for (col, c) in srccells
            s = restyle(c.style, get(colborder, col, "0"))
            s == 0 || (cells[col] = _styled_empty_cell(r, col, s))
        end
        isempty(cells) && isnothing(ht) && isempty(rowattrs) && continue
        cache.cells[r] = cells
        cache.row_ht[r] = ht
        push!(cache.rows_in_cache, r)
        if !isempty(rowattrs)
            isnothing(ws.unhandled_attributes) && (ws.unhandled_attributes = Dict{Int,Dict{String,String}}())
            ws.unhandled_attributes[r] = copy(rowattrs)
        end
    end
    cache.row_index = Dict{Int,Int}()
    cache.dirty = true
    return nothing
end

function _inherit_columns!(ws::Worksheet, src::Int, news::UnitRange{Int}, restyle::Function)
    uh = something(ws.unhandled_attributes, Dict{Int,Dict{String,String}}())
    rowborder(r) = (s = get(get(uh, r, Dict{String,String}()), "s", "0");
                    get(styles_cell_xf(get_workbook(ws), parse(Int, s)), "borderId", "0"))
    for row in values(ws.cache.cells), col in news
        c = get(row, src, nothing)
        isnothing(c) && continue
        s = restyle(c.style, rowborder(row_number(c.ref)))
        s == 0 || (row[col] = _styled_empty_cell(row_number(c.ref), col, s))
    end
    ws.cache.dirty = true

    # The column definition: the neighbour's, as a new column takes it. One the insert
    # fell inside (it spans the new columns) is split around them if they differ.
    root = _sheet_root(ws)
    i = findfirst(c -> localname(c) == "cols", XML.children(root))
    isnothing(i) && return nothing
    cols = XML.children(root)[i]
    defs = xml_elements(cols)
    k = findfirst(d -> parse(Int, d["min"]) <= src <= parse(Int, get(d, "max", d["min"])), defs)
    isnothing(k) && return nothing
    d = defs[k]
    hi = parse(Int, get(d, "max", d["min"]))
    new = copynode(d)
    foreach(a -> remove_attr!(new, a), ("hidden", "collapsed"))
    if get(d, "hidden", "0") in ("1", "true") && parse(Float64, get(d, "width", "0")) == 0
        foreach(a -> remove_attr!(new, a), ("width", "customWidth"))
    end
    if haskey(new, "style")
        s = restyle(parse(UInt32, new["style"]))
        s == 0 ? remove_attr!(new, "style") : (new["style"] = string(s))
    end
    # A default width alone (no customWidth) formats nothing, and Excel writes no definition for it.
    plain = all(p -> first(p) in ("min", "max", "width"), something(XML.attributes(new), Pair{String,String}[]))
    new["min"] = d["min"]
    new["max"] = string(hi)
    if new == d                                            # the same definition: it widens
        d["max"] = string(max(hi, last(news)))
        return nothing
    end
    new["min"] = string(first(news))
    new["max"] = string(last(news))
    kids = XML.children(cols)
    at = findfirst(x -> x === d, kids)
    if hi >= first(news)                                   # split the span around the new columns
        left = copynode(d)
        left["max"] = string(first(news) - 1)
        parts = XML.Node[left]
        plain || push!(parts, new)
        if hi > last(news)
            right = copynode(d)
            right["min"] = string(last(news) + 1)
            push!(parts, right)
        end
    else
        plain && return nothing
        parts = XML.Node[d, new]
    end
    splice!(kids, at, parts)
    haskey(cols, "count") && (cols["count"] = string(length(xml_elements(cols))))
    return nothing
end

# A group that a visible new row (column) joins is no longer all hidden, so its
# summary row (column), after it or before it as the sheet's outline places it, is
# no longer collapsed, as Excel writes it.
function _uncollapse_groups!(ws::Worksheet, e::RowColEdit)
    news = e.at:min(e.at + e.count - 1, _edit_max(e))
    root = _sheet_root(ws)
    pr = first_element_with_tag(first_element_with_tag(root, "sheetPr"), "outlinePr")
    after = get(something(pr, XML.Element("outlinePr")), e.axis === :row ? "summaryBelow" : "summaryRight", "1") != "0"
    # Each row's (column's) outline level, and how to clear its collapsed marker.
    if e.axis === :row
        uh = ws.unhandled_attributes
        isnothing(uh) && return nothing
        level_of = r -> parse(Int, get(get(uh, r, _EMPTY_ROW_ATTRS), "outlineLevel", "0"))
        uncollapse! = r -> (haskey(uh, r) && delete!(uh[r], "collapsed"); nothing)
    else
        i = findfirst(c -> localname(c) == "cols", XML.children(root))
        isnothing(i) && return nothing
        defs = xml_elements(XML.children(root)[i])
        def_of = c -> findfirst(d -> parse(Int, d["min"]) <= c <= parse(Int, get(d, "max", d["min"])), defs)
        level_of = c -> (k = def_of(c); isnothing(k) ? 0 : parse(Int, get(defs[k], "outlineLevel", "0")))
        uncollapse! = function (c)
            k = def_of(c)
            isnothing(k) && return nothing
            remove_attr!(defs[k], "collapsed")
            # Left with only a default width, it defines nothing, and Excel drops it.
            all(p -> first(p) in ("min", "max", "width"), XML.attributes(defs[k])) &&
                filter!(x -> x !== defs[k], XML.children(XML.children(root)[i]))
            return nothing
        end
    end
    lvl = maximum(level_of, news)
    lvl == 0 && return nothing
    step = after ? 1 : -1
    p = after ? last(news) + 1 : first(news) - 1
    while 1 <= p <= _edit_max(e) && level_of(p) >= lvl          # past the rest of the group
        p += step
    end
    1 <= p <= _edit_max(e) && uncollapse!(p)
    return nothing
end

# The conditional formats and data validations, x14 ones included, whose ranges end
# at the row (column) before the insert, grown over the new rows (columns).
function _grow_format_ranges!(root::XML.Node, e::RowColEdit)
    for c in xml_elements(root)
        t = localname(c)
        if t == "conditionalFormatting"
            haskey(c, "sqref") && (c["sqref"] = _grow_sqref(c["sqref"], e))
        elseif t == "dataValidations"
            for dv in xml_elements(c)
                haskey(dv, "sqref") && (dv["sqref"] = _grow_sqref(dv["sqref"], e))
            end
        elseif t == "extLst"
            _grow_ext_sqrefs!(c, e)
        end
    end
    return nothing
end

function _grow_ext_sqrefs!(node::XML.Node, e::RowColEdit)
    for c in xml_elements(node)
        if localname(c) in ("conditionalFormatting", "dataValidation")
            sq = findfirst(k -> localname(k) == "sqref", xml_elements(c))
            if !isnothing(sq)
                el = xml_elements(c)[sq]
                el[end] = XML.Text(_grow_sqref(something(XML.is_simple_value(el), ""), e))
                continue
            end
        end
        _grow_ext_sqrefs!(c, e)
    end
    return nothing
end

# A space-separated range list with each area that ends at row (column) `e.at - 1`
# extended to the last new row (column). Whole rows and columns are left alone.
function _grow_sqref(sqref::AbstractString, e::RowColEdit)::String
    cap = e.axis === :row ? 4 : 2
    grow(area) = begin
        parts = [match(_REF_PART, p) for p in split(area, ':')]
        (any(isnothing, parts) || any(m -> isnothing(m[cap]), parts)) && return area
        pos = [e.axis === :row ? parse(Int, m[4]) : decode_column_number(m[2]) for m in parts]
        maximum(pos) == e.at - 1 || return area
        last_new = min(e.at + e.count - 1, _edit_max(e))
        ends = length(parts) == 1 ? [(parts[1], pos[1]), (parts[1], last_new)] :   # a cell becomes a range
               [(m, i == argmax(pos) ? last_new : p) for (i, (m, p)) in enumerate(zip(parts, pos))]
        return join((_ref_part_text(m, e, p, true) for (m, p) in ends), ':')
    end
    return join((grow(a) for a in split(sqref)), ' ')
end

# The border of each styled column's style, by column number.
function _column_style_borders(ws::Worksheet)::Dict{Int,String}
    out = Dict{Int,String}()
    root = _sheet_root(ws)
    wb = get_workbook(ws)
    for c in xml_elements(root)
        localname(c) == "cols" || continue
        for d in xml_elements(c)
            haskey(d, "style") || continue
            b = get(styles_cell_xf(wb, parse(Int, d["style"])), "borderId", "0")
            for col in parse(Int, d["min"]):parse(Int, get(d, "max", d["min"]))
                out[col] = b
            end
        end
    end
    return out
end
