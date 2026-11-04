# ===========================================================================
# Cell comments (Excel "Notes")
#
# A note lives in three places:
#   - xl/commentsN.xml      the text and author (parsed XML, in xf.data)
#   - a VML drawing         the pop-up box shape (raw bytes, in xf.binary_data)
#   - <legacyDrawing r:id>  in the worksheet, pointing at the VML drawing
#
# VML is kept as text rather than parsed: Excel writes it as HTML-flavoured XML
# (unclosed <br>, single-quoted attributes), and the same part may also hold
# form controls that must survive untouched. Only note shapes are added or
# removed, by string surgery.
#
# Threaded comments (Excel 365) are not written. A cell that already has one is
# refused rather than silently desynchronised from its legacy fallback note.
# ===========================================================================

const REL_COMMENTS =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments"
const REL_VML_DRAWING =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing"
const REL_THREADED_COMMENT =
    "http://schemas.microsoft.com/office/2017/10/relationships/threadedComment"
const MIME_COMMENTS =
    "application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml"
const MIME_VML =
    "application/vnd.openxmlformats-officedocument.vmlDrawing"

const CommentInfo = NamedTuple{
    (:sheet, :ref, :author, :text),
    Tuple{String, String, String, String},
}

# ===========================================================================
# Public API
# ===========================================================================

"""
    addComment(s::Worksheet, ref::AbstractString, text::AbstractString; author="", width=144, height=79, visible=false)
    addComment(s::Worksheet, row::Integer, col::Integer, text::AbstractString; kw...)

Attach a comment (an Excel "Note") to a single cell. A cell with a comment shows a
small red triangle in its top right corner, and the comment pops up when the
mouse hovers over the cell.

If the cell already has a comment, it is replaced.

Comments can only be attached to a single cell. Giving a cell range is an error.

# Arguments

`s::Worksheet`: the target worksheet.

`ref`: a single cell reference, e.g. `"B3"`, or a `CellRef`.

`text::AbstractString`: the comment text. Use `\\n` for line breaks.

# Keyword Arguments

`author::AbstractString`: the comment's author. Defaults to `""`.

`width`, `height`: the size of the pop-up box, in pixels. The box is positioned
assuming default cell sizes of 64 pixels wide and 20 pixels high.

`visible::Bool`: if `true`, the comment is always shown rather than only on hover.

# Return Value

A `NamedTuple` of `String`s: `(sheet, ref, author, text)`.

# Examples

```julia
XLSX.addComment(sheet, "B3", "Checked against the 2024 accounts"; author="Tim")
XLSX.addComment(sheet, 5, 2, "Line one\\nLine two"; visible=true)
```

!!! note

    Excel 365 also has *threaded comments*. These are not written by XLSX.jl.
    Every version of Excel shows the notes written here. Trying to replace or
    delete a threaded comment that already exists in a file throws an error.

See also [`getComment`](@ref), [`getComments`](@ref), [`deleteComment`](@ref).
"""
addComment(s::Worksheet, row::Integer, col::Integer, text::AbstractString; kw...) =
    addComment(s, CellRef(row, col), text; kw...)
addComment(s::Worksheet, ref::AbstractString, text::AbstractString; kw...) =
    addComment(s, _comment_cellref(ref), text; kw...)

function addComment(
    s::Worksheet,
    ref::CellRef,
    text::AbstractString;
    author::AbstractString = "",
    width::Integer = 144,
    height::Integer = 79,
    visible::Bool = false,
)
    _require_writable(s, "add a comment")
    _require_worksheet(s, "add a comment")
    (width > 0 && height > 0) || throw(XLSXError("Comment width and height must be positive."))

    xf         = get_xlsxfile(s)
    sheet_path = _sheet_path(s)
    r          = string(ref)
    _refuse_threaded(xf, sheet_path, r)

    comments_path = ensure_comments!(xf, sheet_path)
    _set_comment_xml!(xf, comments_path, r, author, text)

    vml_path = ensure_legacy_vml!(xf, sheet_path)
    vml      = _vml_string(xf, vml_path)
    vml      = _remove_note_shape(vml, row_number(ref) - 1, column_number(ref) - 1)
    vml      = _add_note_shape(xf, vml, ref; width, height, visible)
    xf.binary_data[vml_path] = Vector{UInt8}(codeunits(vml))

    return CommentInfo((s.name, r, String(author), String(text)))
end

"""
    getComment(s::Worksheet, ref) -> Union{Nothing, NamedTuple}
    getComment(s::Worksheet, row::Integer, col::Integer) -> Union{Nothing, NamedTuple}

Return the comment attached to the given cell as a `NamedTuple` of `String`s,
`(sheet, ref, author, text)`, or `nothing` if the cell has no comment.

The text is returned as stored. Notes written by Excel itself usually begin with
the author's name, e.g. `"Tim:\\n..."`, because Excel puts it in the note text.

For a threaded comment (Excel 365), this returns the legacy note that Excel
stores alongside it.

See also [`addComment`](@ref), [`getComments`](@ref).
"""
getComment(s::Worksheet, row::Integer, col::Integer) = getComment(s, CellRef(row, col))
getComment(s::Worksheet, ref::AbstractString) = getComment(s, _comment_cellref(ref))
function getComment(s::Worksheet, ref::CellRef)::Union{Nothing,CommentInfo}
    r  = string(ref)
    cs = getComments(s)
    i  = findfirst(c -> c.ref == r, cs)
    return isnothing(i) ? nothing : cs[i]
end

"""
    getComments(s::Worksheet) -> Vector{NamedTuple}
    getComments(xf::XLSXFile) -> Vector{NamedTuple}

Return every comment on a worksheet, or in every worksheet of a workbook, as a
vector of `NamedTuple`s of `String`s, `(sheet, ref, author, text)`.

See also [`addComment`](@ref), [`getComment`](@ref).
"""
function getComments(s::Worksheet)::Vector{CommentInfo}
    is_chartsheet(get_workbook(s), s.name) && return CommentInfo[]
    xf   = get_xlsxfile(s)
    path = _comments_path(xf, _sheet_path(s))
    isnothing(path) && return CommentInfo[]
    root    = xml_root_element(get_xml_data(xf, path))
    authors = _comment_authors(root)
    out     = CommentInfo[]
    for c in _comment_nodes(root)
        id     = tryparse(Int, get_attr(c, "authorId", "0"))
        author = isnothing(id) || id + 1 > length(authors) ? "" : authors[id + 1]
        push!(out, CommentInfo((s.name, get_attr(c, "ref"), author, _comment_text(c))))
    end
    return out
end

function getComments(xf::XLSXFile)::Vector{CommentInfo}
    wb = get_workbook(xf)
    return reduce(vcat, (getComments(getsheet(wb, i)) for i in 1:sheetcount(wb)); init=CommentInfo[])
end

"""
    deleteComment(s::Worksheet, ref)
    deleteComment(s::Worksheet, row::Integer, col::Integer)

Remove the comment attached to the given cell. Return `true` if a comment was
removed and `false` if the cell had none.

See also [`addComment`](@ref).
"""
deleteComment(s::Worksheet, row::Integer, col::Integer) = deleteComment(s, CellRef(row, col))
deleteComment(s::Worksheet, ref::AbstractString) = deleteComment(s, _comment_cellref(ref))
function deleteComment(s::Worksheet, ref::CellRef)::Bool
    _require_writable(s, "delete a comment")
    _require_worksheet(s, "delete a comment")
    xf         = get_xlsxfile(s)
    sheet_path = _sheet_path(s)
    r          = string(ref)
    _refuse_threaded(xf, sheet_path, r)

    comments_path = _comments_path(xf, sheet_path)
    isnothing(comments_path) && return false
    root = xml_root_element(get_xml_data(xf, comments_path))
    list = first_element_with_tag(root, "commentList")
    isnothing(list) && return false
    i = findfirst(c -> XML.nodetype(c) == XML.Element && localname(c) == "comment" && get_attr(c, "ref") == r,
                  XML.children(list))
    isnothing(i) && return false
    deleteat!(XML.children(list), i)

    vml_path = _legacy_vml_path(xf, sheet_path)
    if !isnothing(vml_path) && haskey(xf.binary_data, vml_path)
        vml = _remove_note_shape(_vml_string(xf, vml_path), row_number(ref) - 1, column_number(ref) - 1)
        xf.binary_data[vml_path] = Vector{UInt8}(codeunits(vml))
    end

    # The last comment gone: remove the parts, so the file looks as if none were ever added.
    if isempty(_comment_nodes(root))
        _remove_sheet_part!(xf, sheet_path, REL_COMMENTS)
        if !isnothing(vml_path) && !occursin(r"<v:shape\b", _vml_string(xf, vml_path))
            _remove_sheet_part!(xf, sheet_path, REL_VML_DRAWING; target=vml_path)
            _remove_sheet_element!(xf, sheet_path, "legacyDrawing")
        end
    end
    return true
end

# ===========================================================================
# Lookup helpers
# ===========================================================================

function _comment_cellref(ref::AbstractString)::CellRef
    is_valid_cellname(ref) && return CellRef(ref)
    is_valid_cellrange(ref) &&
        throw(XLSXError("Comments attach to a single cell, not a range: `$ref`."))
    throw(XLSXError("Invalid cell reference: `$ref`."))
end

_sheet_path(s::Worksheet) = get_relationship_target_by_id("xl", get_workbook(s), s.relationship_id)

function _require_worksheet(s::Worksheet, action::AbstractString)
    is_chartsheet(get_workbook(s), s.name) &&
        throw(XLSXError("Cannot $action: `$(s.name)` is a chartsheet."))
    return nothing
end

_sheet_rels_path(sheet_path::AbstractString) =
    (p = _split_zip_path(sheet_path); "$(p[1])/_rels/$(p[2]).rels")

# Resolved targets of the sheet's relationships of type `type`.
function _sheet_rel_targets(xf::XLSXFile, sheet_path::String, type::String)::Vector{Tuple{String,String}}
    rels_path = _sheet_rels_path(sheet_path)
    haskey(xf.data, rels_path) || return Tuple{String,String}[]
    dir = first(_split_zip_path(sheet_path))
    return [(get_attr(n, "Id"), resolve_relative_target(dir, get_attr(n, "Target")))
            for n in elements_with_tag(xml_root_element(get_xml_data(xf, rels_path)), "Relationship")
            if get_attr(n, "Type") == type]
end

function _comments_path(xf::XLSXFile, sheet_path::String)::Union{Nothing,String}
    t = _sheet_rel_targets(xf, sheet_path, REL_COMMENTS)
    return isempty(t) ? nothing : last(first(t))
end

# The VML part referenced by the sheet's <legacyDrawing>. Not just any vmlDrawing
# relationship: header/footer images use one too, through <legacyDrawingHF>.
function _legacy_vml_path(xf::XLSXFile, sheet_path::String)::Union{Nothing,String}
    el = first_element_with_tag(xml_root_element(get_xml_data(xf, sheet_path)), "legacyDrawing")
    isnothing(el) && return nothing
    rid = get_prefixed_attr(el, "id")
    for (id, target) in _sheet_rel_targets(xf, sheet_path, REL_VML_DRAWING)
        id == rid && return target
    end
    return nothing
end

function _refuse_threaded(xf::XLSXFile, sheet_path::String, ref::String)
    for (_, path) in _sheet_rel_targets(xf, sheet_path, REL_THREADED_COMMENT)
        haskey(xf.data, path) || continue
        root = xml_root_element(get_xml_data(xf, path))
        any(n -> get_attr(n, "ref") == ref, elements_with_tag(root, "threadedComment")) &&
            throw(XLSXError("Cell $ref has a threaded comment, which XLSX.jl cannot change. Edit it in Excel."))
    end
    return nothing
end

# ===========================================================================
# Comments part
# ===========================================================================

empty_comments_doc() = XML.Document(
    XML.Declaration(; version="1.0", encoding="UTF-8", standalone="yes"),
    XML.Element("comments",
        XML.Element("authors"),
        XML.Element("commentList");
        xmlns = SPREADSHEET_NAMESPACE_XPATH_ARG,
    ),
)

function ensure_comments!(xf::XLSXFile, sheet_path::String)::String
    path = _comments_path(xf, sheet_path)
    isnothing(path) || return path

    i = 1
    while haskey(xf.files, "xl/comments$i.xml") || haskey(xf.data, "xl/comments$i.xml"); i += 1; end
    path = "xl/comments$i.xml"
    xf.data[path]  = empty_comments_doc()
    xf.files[path] = true
    add_part_rel!(xf, sheet_path, path, REL_COMMENTS)
    register_content_type!(xf, "[Content_Types].xml";
                           tag="Override", key="PartName", val="/$path", content_type=MIME_COMMENTS)
    return path
end

# The prefix the part uses for the spreadsheetml namespace, read off its root.
_root_prefix(root::XML.Node) = (t = XML.tag(root); occursin(':', t) ? first(split(t, ':')) : "")

_comment_nodes(root::XML.Node) =
    (list = first_element_with_tag(root, "commentList");
     isnothing(list) ? XML.Node[] : elements_with_tag(list, "comment"))

_comment_authors(root::XML.Node) =
    (a = first_element_with_tag(root, "authors");
     isnothing(a) ? String[] : [something(text_value(n), "") for n in elements_with_tag(a, "author")])

# All the text runs of a comment, joined. Phonetic runs (<rPh>) are not text.
function _comment_text(c::XML.Node)::String
    text = first_element_with_tag(c, "text")
    isnothing(text) && return ""
    io = IOBuffer()
    for n in XML.eachelement(text)
        if has_localname(n, "t")
            print(io, something(text_value(n), ""))
        elseif has_localname(n, "r")
            t = first_element_with_tag(n, "t")
            isnothing(t) || print(io, something(text_value(t), ""))
        end
    end
    return String(take!(io))
end

function _set_comment_xml!(xf::XLSXFile, path::String, ref::String, author::AbstractString, text::AbstractString)
    root = xml_root_element(get_xml_data(xf, path))
    pfx  = _root_prefix(root)
    tag(name) = prefixed_tag(pfx, name)

    authors = first_element_with_tag(root, "authors")
    if isnothing(authors)
        authors = XML.Element(tag("authors"))
        pushfirst!(XML.children(root), authors)
    end
    names = _comment_authors(root)
    id    = findfirst(==(author), names)
    if isnothing(id)
        push!(authors, XML.Element(tag("author"), XML.Text(String(author))))
        id = length(names) + 1
    end

    list = first_element_with_tag(root, "commentList")
    if isnothing(list)
        list = XML.Element(tag("commentList"))
        j = findfirst(n -> XML.nodetype(n) == XML.Element && localname(n) == "authors", XML.children(root))
        insert!(XML.children(root), j + 1, list)
    end

    # Excel's own note font: Tahoma 9pt in the "comment text" indexed colour.
    # Child order follows CT_RPrElt.
    comment = XML.Element(tag("comment"),
        XML.Element(tag("text"),
            XML.Element(tag("r"),
                XML.Element(tag("rPr"),
                    XML.Element(tag("rFont"); val="Tahoma"),
                    XML.Element(tag("family"); val="2"),
                    XML.Element(tag("color"); indexed="81"),
                    XML.Element(tag("sz"); val="9"),
                ),
                XML.Element(tag("t"), XML.Text(String(text)); var"xml:space"="preserve"),
            ),
        );
        ref = ref, authorId = string(id - 1),
    )

    kids = XML.children(list)
    i = findfirst(c -> XML.nodetype(c) == XML.Element && localname(c) == "comment" && get_attr(c, "ref") == ref, kids)
    isnothing(i) ? push!(list, comment) : (kids[i] = comment)
    return nothing
end

# ===========================================================================
# VML part
# ===========================================================================

const _VML_SHAPETYPE_NOTE = """
 <v:shapetype id="_x0000_t202" coordsize="21600,21600" o:spt="202" path="m,l,21600r21600,l21600,xe">
  <v:stroke joinstyle="miter"/>
  <v:path gradientshapeok="t" o:connecttype="rect"/>
 </v:shapetype>
"""

# One shape element, opening tag to closing tag, or self-closed. `\b` keeps
# <v:shapetype> out.
const _VML_SHAPE_RE = r"<v:shape\b[^>]*?(?:/>|>.*?</v:shape>)"s

_vml_string(xf::XLSXFile, path::String) = String(copy(xf.binary_data[path]))

function _empty_vml(block::Int)::String
    return """
<xml xmlns:v="urn:schemas-microsoft-com:vml"
 xmlns:o="urn:schemas-microsoft-com:office:office"
 xmlns:x="urn:schemas-microsoft-com:office:excel">
 <o:shapelayout v:ext="edit">
  <o:idmap v:ext="edit" data="$block"/>
 </o:shapelayout>
$(_VML_SHAPETYPE_NOTE)</xml>
"""
end

function ensure_legacy_vml!(xf::XLSXFile, sheet_path::String)::String
    path = _legacy_vml_path(xf, sheet_path)
    isnothing(path) || return path
    isnothing(first_element_with_tag(xml_root_element(get_xml_data(xf, sheet_path)), "legacyDrawing")) ||
        throw(XLSXError("Malformed worksheet `$sheet_path`: its `<legacyDrawing>` has no matching VML relationship."))

    i = 1
    while haskey(xf.binary_data, "xl/drawings/vmlDrawing$i.vml"); i += 1; end
    path = "xl/drawings/vmlDrawing$i.vml"
    xf.binary_data[path] = Vector{UInt8}(codeunits(_empty_vml(_free_vml_block(xf))))

    rid = add_part_rel!(xf, sheet_path, path, REL_VML_DRAWING)
    register_content_type!(xf, "[Content_Types].xml";
                           tag="Default", key="Extension", val="vml", content_type=MIME_VML)

    sheet_root = xml_root_element(get_xml_data(xf, sheet_path))
    if !haskey(something(XML.attributes(sheet_root), Dict()), "xmlns:r")
        sheet_root["xmlns:r"] = NS_R
    end
    el = XML.Element(prefixed_tag(_root_prefix(sheet_root), "legacyDrawing"))
    el["r:id"] = rid
    insert!(XML.children(sheet_root), insert_index(sheet_root, "legacyDrawing", WORKSHEET_ORDER) + 1, el)
    return path
end

# Shape ids are unique across the workbook. Each VML part claims blocks of 1024
# ids in its <o:idmap data="...">: block b holds ids b*1024+1 to b*1024+1023.
_vml_blocks(vml::AbstractString)::Vector{Int} =
    (m = match(r"<o:idmap\b[^>]*\bdata\s*=\s*[\"']([^\"']*)[\"']", vml);
     isnothing(m) ? Int[] : Int[v for v in tryparse.(Int, strip.(split(m[1], ','))) if !isnothing(v)])

function _free_vml_block(xf::XLSXFile)::Int
    used = Set{Int}()
    for (k, v) in xf.binary_data
        endswith(lowercase(k), ".vml") && union!(used, _vml_blocks(String(copy(v))))
    end
    b = 1
    while b in used; b += 1; end
    return b
end

function _next_vml_shape_id(xf::XLSXFile, vml::String)::Tuple{Int,String}
    taken = Set(parse(Int, m[1]) for m in eachmatch(r"_x0000_s(\d+)", vml))
    for b in sort(_vml_blocks(vml)), id in b*1024+1:b*1024+1023
        id in taken || return id, vml
    end
    # Every block this part owns is full (or it owns none): claim a new one.
    b      = _free_vml_block(xf)
    blocks = join(vcat(_vml_blocks(vml), b), ",")
    if occursin(r"<o:idmap\b", vml)
        vml = replace(vml, r"(<o:idmap\b[^>]*\bdata\s*=\s*)[\"'][^\"']*[\"']" => SubstitutionString("\\1\"$blocks\""))
    else
        layout = """<o:shapelayout v:ext="edit"><o:idmap v:ext="edit" data="$blocks"/></o:shapelayout>"""
        vml = replace(vml, r"<xml\b[^>]*>" => SubstitutionString("\\0\n $layout"); count=1)
    end
    return b*1024 + 1, vml
end

# Remove the note shape anchored to the 0-based (row, col), if there is one.
# Form controls and other shapes in the same part are left alone.
function _remove_note_shape(vml::String, row::Int, col::Int)::String
    for m in eachmatch(_VML_SHAPE_RE, vml)
        s = m.match
        occursin(r"ObjectType\s*=\s*[\"']Note[\"']", s) || continue
        r = match(r"<(?:\w+:)?Row>\s*(\d+)\s*</(?:\w+:)?Row>", s)
        c = match(r"<(?:\w+:)?Column>\s*(\d+)\s*</(?:\w+:)?Column>", s)
        (isnothing(r) || isnothing(c)) && continue
        if parse(Int, r[1]) == row && parse(Int, c[1]) == col
            return vml[1:prevind(vml, m.offset)] * vml[m.offset+ncodeunits(s):end]
        end
    end
    return vml
end

function _add_note_shape(xf::XLSXFile, vml::String, ref::CellRef;
                         width::Integer, height::Integer, visible::Bool)::String
    occursin(r"id\s*=\s*[\"']_x0000_t202[\"']", vml) ||
        (vml = _insert_before_close(vml, _VML_SHAPETYPE_NOTE))
    id, vml = _next_vml_shape_id(xf, vml)
    z = count(_ -> true, eachmatch(_VML_SHAPE_RE, vml)) + 1
    return _insert_before_close(vml, _note_shape(id, z, row_number(ref) - 1, column_number(ref) - 1;
                                                 width, height, visible))
end

function _insert_before_close(vml::String, s::String)::String
    i = findlast("</xml>", vml)
    isnothing(i) && throw(XLSXError("Malformed VML drawing: no closing `</xml>`."))
    return vml[1:prevind(vml, first(i))] * s * vml[first(i):end]
end

# The box sits to the right of the cell and a little above it, as Excel puts
# it. The <x:Anchor> is what Excel uses to place it: left column, x offset, top
# row, y offset, right column, x offset, bottom row, y offset (0-based cells,
# pixel offsets), assuming default 64 x 20 pixel cells as `addImage` does.
function _note_shape(id::Int, z::Int, row::Int, col::Int; width::Integer, height::Integer, visible::Bool)::String
    left_col, left_off = col + 1, 15
    top_row,  top_off  = row == 0 ? (0, 2) : (row - 1, 10)
    right_col, right_off = left_col + (left_off + width) ÷ 64, (left_off + width) % 64
    bot_row,   bot_off   = top_row + (top_off + height) ÷ 20, (top_off + height) % 20
    pt(px) = string(round(px * 0.75; digits=2))
    style = "position:absolute;margin-left:$(pt(left_col*64 + left_off))pt;margin-top:$(pt(top_row*20 + top_off))pt;" *
            "width:$(pt(width))pt;height:$(pt(height))pt;z-index:$z;visibility:$(visible ? "visible" : "hidden")"
    return """
 <v:shape id="_x0000_s$id" type="#_x0000_t202" style="$style" fillcolor="#ffffe1" o:insetmode="auto">
  <v:fill color2="#ffffe1"/>
  <v:shadow on="t" color="black" obscured="t"/>
  <v:path o:connecttype="none"/>
  <v:textbox style="mso-direction-alt:auto"><div style="text-align:left"></div></v:textbox>
  <x:ClientData ObjectType="Note">
   <x:MoveWithCells/>
   <x:SizeWithCells/>
   <x:Anchor>$left_col, $left_off, $top_row, $top_off, $right_col, $right_off, $bot_row, $bot_off</x:Anchor>
   <x:AutoFill>False</x:AutoFill>
   <x:Row>$row</x:Row>
   <x:Column>$col</x:Column>$(visible ? "\n   <x:Visible/>" : "")
  </x:ClientData>
 </v:shape>
"""
end

# ===========================================================================
# Removal
# ===========================================================================

# Drop the sheet's relationship of `type` (to `target`, if given) and the part
# it points at, with anything only that part references.
function _remove_sheet_part!(xf::XLSXFile, sheet_path::String, type::String; target=nothing)
    rels_path = _sheet_rels_path(sheet_path)
    haskey(xf.data, rels_path) || return nothing
    dir  = first(_split_zip_path(sheet_path))
    root = xml_root_element(get_xml_data(xf, rels_path))
    kids = XML.children(root)
    for i in reverse(eachindex(kids))
        n = kids[i]
        XML.nodetype(n) == XML.Element && get_attr(n, "Type") == type || continue
        path = resolve_relative_target(dir, get_attr(n, "Target"))
        isnothing(target) || path == target || continue
        deleteat!(kids, i)
        delete_part_and_orphans!(xf, path)
    end
    return nothing
end

function _remove_sheet_element!(xf::XLSXFile, sheet_path::String, name::String)
    root = xml_root_element(get_xml_data(xf, sheet_path))
    i = findfirst(n -> XML.nodetype(n) == XML.Element && localname(n) == name, XML.children(root))
    isnothing(i) || deleteat!(XML.children(root), i)
    return nothing
end
