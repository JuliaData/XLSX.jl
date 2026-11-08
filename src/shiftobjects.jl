# ===========================================================================
# Moving a sheet's drawings and comments as rows or columns are inserted or deleted
#
# Images, charts and shapes sit in the sheet's drawing part, each placed by
# markers: a cell (0-based `row`, `col`) and an offset into it. How a marker moves
# depends on the object's `editAs`, as Excel's *Format → Properties* sets it:
#
# - "twoCell" (move and size with cells, the default): both corners move, so an
#   insert between them stretches the object;
# - "oneCell" (move but don't size): the top-left corner moves and the object keeps
#   its size;
# - "absolute" (don't move or size): nothing moves.
#
# A corner whose row (column) is deleted goes to the top (left) of the row (column)
# that follows. An object all of whose rows (columns) are deleted goes, as in Excel,
# if it moves and sizes with its cells; one that only moves, or doesn't, is kept at
# its size.
#
# Comments (notes) are in the comments part, by cell, with a box in the sheet's VML
# drawing whose `<x:Row>`/`<x:Column>` name the cell and whose `<x:Anchor>` places
# it. A comment on a deleted cell is deleted. Threaded comments (Excel 365) are
# moved, but one on a deleted cell is refused: XLSX.jl can't change them.
# ===========================================================================

# Where a marker in 1-based row (column) `p` goes: its new row (column), and
# whether its offset is kept (`false` if its row was deleted, when it goes to the
# top of the row that follows, or it is pushed off the end of the sheet).
function _shift_marker_index(p::Int, e::RowColEdit)::Tuple{Int,Bool}
    q = _shift_index(p, e)
    isnothing(q) || return (q, true)
    return e.count < 0 ? (e.at, false) : (_edit_max(e), true)
end

# The same for an object's bottom (right) corner. One at the very top (left) of row
# (column) `p`, offset 0, is the end of the row before: an insert at `p` falls
# after the object, not inside it.
function _shift_end_marker_index(p::Int, at_edge::Bool, e::RowColEdit)::Tuple{Int,Bool}
    (at_edge && p > 1) || return _shift_marker_index(p, e)
    q = _shift_index(p - 1, e)
    isnothing(q) && return (e.at, false)
    return (min(q + 1, _edit_max(e)), true)
end

# --- drawings

# Delete the "twoCell" objects whose rows (columns) edit `e` deletes, top to bottom
# (left to right), with what only they use, as Excel does; before the others move.
function _delete_drawings_gone!(xf::XLSXFile, sheet_path::String, e::RowColEdit)
    e.count < 0 || return nothing
    path = _drawing_path_for_sheet(xf, sheet_path)
    isnothing(path) && return nothing
    _delete_drawing_objects!(xf, sheet_path, path, a -> _all_deleted(a, e))
    return nothing
end

function _all_deleted(a::XML.Node, e::RowColEdit)::Bool
    localname(a) == "twoCellAnchor" && get(a, "editAs", "twoCell") == "twoCell" || return false
    from, to = first_element_with_tag(a, "from"), first_element_with_tag(a, "to")
    (isnothing(from) || isnothing(to)) && return false
    idx_tag, off_tag = e.axis === :row ? ("row", "rowOff") : ("col", "colOff")
    val(m, t) = (el = first_element_with_tag(m, t); isnothing(el) ? 0 : parse(Int, strip(something(XML.is_simple_value(el), "0"))))
    first_p = val(from, idx_tag) + 1
    last_p  = val(to, idx_tag) + 1
    val(to, off_tag) == 0 && last_p > first_p && (last_p -= 1)   # ends at the top of the row after
    return e.at <= first_p && last_p <= _edit_last_deleted(e)
end

function _shift_drawing!(xf::XLSXFile, sheet_path::String, e::RowColEdit)
    path = _drawing_path_for_sheet(xf, sheet_path)
    isnothing(path) && return nothing
    for a in xml_elements(xml_root_element(get_xml_data(xf, path)))
        name = localname(a)
        name in ("twoCellAnchor", "oneCellAnchor") || continue   # an absoluteAnchor doesn't move
        from = first_element_with_tag(a, "from")
        to = first_element_with_tag(a, "to")
        edit_as = name == "oneCellAnchor" ? "oneCell" : get(a, "editAs", "twoCell")
        edit_as == "absolute" && continue
        d = _move_marker!(from, e)
        isnothing(to) && continue
        edit_as == "oneCell" ? _move_marker!(to, e; by=d) : _move_marker!(to, e; is_end=true)
    end
    return nothing
end

# Move a drawing marker (`<xdr:from>`, `<xdr:to>`, the latter with `is_end`) for
# edit `e`, or with `by`, by that many rows (columns). Returns how far it moved.
function _move_marker!(m::Union{Nothing,XML.Node}, e::RowColEdit; by::Union{Nothing,Int}=nothing, is_end::Bool=false)::Int
    isnothing(m) && return 0
    idx_tag, off_tag = e.axis === :row ? ("row", "rowOff") : ("col", "colOff")
    idx_el = first_element_with_tag(m, idx_tag)
    off_el = first_element_with_tag(m, off_tag)
    isnothing(idx_el) && return 0
    p = parse(Int, strip(something(XML.is_simple_value(idx_el), "0"))) + 1
    if !isnothing(by)
        q, keep_off = clamp(p + by, 1, _edit_max(e)), true
    elseif is_end
        at_edge = isnothing(off_el) || strip(something(XML.is_simple_value(off_el), "0")) == "0"
        q, keep_off = _shift_end_marker_index(p, at_edge, e)
    else
        q, keep_off = _shift_marker_index(p, e)
    end
    q == p && keep_off && return 0
    idx_el[end] = XML.Text(string(q - 1))
    keep_off || isnothing(off_el) || (off_el[end] = XML.Text("0"))
    return q - p
end

# --- comments

# Comments on the cells edit `e` deletes. A threaded one is refused, before
# anything changes: XLSX.jl can't change them.
function _comments_deleted_by(ws::Worksheet, e::RowColEdit)::Vector{CellRef}
    e.count < 0 || return CellRef[]
    xf = get_xlsxfile(ws)
    sheet_path = _sheet_path(ws)
    for (_, path) in _sheet_rel_targets(xf, sheet_path, REL_THREADED_COMMENT)
        haskey(xf.data, path) || continue
        for n in elements_with_tag(xml_root_element(get_xml_data(xf, path)), "threadedComment")
            isnothing(_shift_cellref(CellRef(get_attr(n, "ref")), e)) &&
                throw(XLSXError("Cannot delete cell $(get_attr(n, "ref")) on sheet `$(ws.name)`: it has a threaded comment, which XLSX.jl cannot change. Delete it in Excel first."))
        end
    end
    comments_path = _comments_path(xf, sheet_path)
    isnothing(comments_path) && return CellRef[]
    refs = CellRef[]
    for c in _comment_nodes(xml_root_element(get_xml_data(xf, comments_path)))
        ref = CellRef(get_attr(c, "ref"))
        isnothing(_shift_cellref(ref, e)) && push!(refs, ref)
    end
    return refs
end

# Move the sheet's comments, threaded comments and their boxes for edit `e`.
# Comments on deleted cells have already gone (see `_comments_deleted_by`).
function _shift_comments!(ws::Worksheet, e::RowColEdit)
    xf = get_xlsxfile(ws)
    sheet_path = _sheet_path(ws)
    move!(c) = (new = _shift_cellref(CellRef(c["ref"]), e); isnothing(new) || (c["ref"] = string(new)))
    comments_path = _comments_path(xf, sheet_path)
    isnothing(comments_path) || foreach(move!, _comment_nodes(xml_root_element(get_xml_data(xf, comments_path))))
    for (_, path) in _sheet_rel_targets(xf, sheet_path, REL_THREADED_COMMENT)
        haskey(xf.data, path) || continue
        foreach(move!, elements_with_tag(xml_root_element(get_xml_data(xf, path)), "threadedComment"))
    end

    vml_path = _legacy_vml_path(xf, sheet_path)
    (isnothing(vml_path) || !haskey(xf.binary_data, vml_path)) && return nothing
    vml = replace(_vml_string(xf, vml_path), _VML_SHAPE_RE => s -> _shift_vml_shape(s, e))
    xf.binary_data[vml_path] = Vector{UInt8}(codeunits(vml))
    return nothing
end

const _VML_ROW = r"(<(?:\w+:)?Row>\s*)(\d+)(\s*</(?:\w+:)?Row>)"
const _VML_COLUMN = r"(<(?:\w+:)?Column>\s*)(\d+)(\s*</(?:\w+:)?Column>)"
const _VML_ANCHOR = r"(<(?:\w+:)?Anchor>)([^<]*)(</(?:\w+:)?Anchor>)"

# A VML shape (a comment's box, or a form control) moved for edit `e`: the cell a
# note is on, and the corners of its `<x:Anchor>` (0-based left column, x offset,
# top row, y offset, right column, x offset, bottom row, y offset), as for a
# "twoCell" drawing.
function _shift_vml_shape(s::AbstractString, e::RowColEdit)::String
    cell_rx = e.axis === :row ? _VML_ROW : _VML_COLUMN
    s = replace(s, cell_rx => m -> (c = match(cell_rx, m);
                                    c[1] * string(first(_shift_marker_index(parse(Int, c[2]) + 1, e)) - 1) * c[3]))
    return replace(s, _VML_ANCHOR => function (m)
        c = match(_VML_ANCHOR, m)
        v = [parse(Int, strip(x)) for x in split(c[2], ',')]
        length(v) == 8 || return m
        start, stop = e.axis === :row ? (3, 7) : (1, 5)   # each marker's cell, then its offset
        for (i, q_keep) in ((start, _shift_marker_index(v[start] + 1, e)),
                            (stop, _shift_end_marker_index(v[stop] + 1, v[stop+1] == 0, e)))
            q, keep_off = q_keep
            v[i] = q - 1
            keep_off || (v[i+1] = 0)
        end
        return c[1] * join(v, ", ") * c[3]
    end)
end
