# ===========================================================================
# Threaded comments (Excel 365's "comments", as opposed to notes)
#
# A thread lives in four places:
#   - xl/threadedComments/threadedCommentN.xml   one <threadedComment> per comment
#     and reply, tied by `parentId` to the thread's first comment; the sheet points
#     at it.
#   - xl/persons/person.xml                      the people who wrote them; the
#     workbook points at it.
#   - a note in the sheet's comments part, author "tc={id}", holding the thread as
#     text for versions of Excel without threads (comments.jl writes it, and its box).
#   - that note's box in the VML drawing.
# ===========================================================================

const REL_PERSON = "http://schemas.microsoft.com/office/2017/10/relationships/person"

const ThreadedReply = NamedTuple{(:author, :text, :date), Tuple{String, String, Union{Nothing,Dates.DateTime}}}

const ThreadedCommentInfo = NamedTuple{
    (:sheet, :ref, :author, :text, :date, :resolved, :replies),
    Tuple{String, String, String, String, Union{Nothing,Dates.DateTime}, Bool, Vector{ThreadedReply}},
}

"""
    getThreadedComments(s::Worksheet) -> Vector{NamedTuple}
    getThreadedComments(xf::XLSXFile) -> Vector{NamedTuple}

Return the threaded comments (Excel 365's comments, as opposed to notes) on a
worksheet, or in every worksheet of a workbook: one `NamedTuple` per thread,
`(sheet, ref, author, text, date, resolved, replies)`. `author` is the person's
display name, `date` a `DateTime` (or `nothing` if the file gives none), and
`resolved` whether the thread is marked resolved. `replies` holds the replies, in
order, each as `(author, text, date)`.

The file doesn't need to be opened for writing. [`getComment`](@ref) returns, for
a cell with a threaded comment, the note Excel stores alongside it for older
versions of Excel.

See also [`getThreadedComment`](@ref), [`addThreadedComment`](@ref).
"""
function getThreadedComments(s::Worksheet)::Vector{ThreadedCommentInfo}
    is_chartsheet(get_workbook(s), s.name) && return ThreadedCommentInfo[]
    xf = get_xlsxfile(s)
    people = _threaded_comment_people(xf)
    out = ThreadedCommentInfo[]
    for (_, path) in _sheet_rel_targets(xf, _sheet_path(s), REL_THREADED_COMMENT)
        _part_exists(xf, path) || continue
        nodes = elements_with_tag(xml_root_element(get_xml_data(xf, path)), "threadedComment")
        replies = Dict{String,Vector{ThreadedReply}}()
        for n in nodes
            parent = get_attr(n, "parentId", "")
            isempty(parent) && continue
            push!(get!(Vector{ThreadedReply}, replies, parent),
                  ThreadedReply((get(people, get_attr(n, "personId"), ""), _threaded_text(n), _threaded_date(n))))
        end
        for n in nodes
            isempty(get_attr(n, "parentId", "")) || continue
            push!(out, ThreadedCommentInfo((s.name, get_attr(n, "ref"), get(people, get_attr(n, "personId"), ""),
                                            _threaded_text(n), _threaded_date(n), get_attr(n, "done", "0") in ("1", "true"),
                                            get(replies, get_attr(n, "id"), ThreadedReply[]))))
        end
    end
    return out
end

function getThreadedComments(xf::XLSXFile)::Vector{ThreadedCommentInfo}
    wb = get_workbook(xf)
    return reduce(vcat, (getThreadedComments(getsheet(wb, i)) for i in 1:sheetcount(wb)); init=ThreadedCommentInfo[])
end

"""
    getThreadedComment(s::Worksheet, ref) -> Union{Nothing, NamedTuple}
    getThreadedComment(s::Worksheet, row::Integer, col::Integer) -> Union{Nothing, NamedTuple}

Return the threaded comment on the given cell, as [`getThreadedComments`](@ref)
does, or `nothing` if the cell has none.
"""
getThreadedComment(s::Worksheet, row::Integer, col::Integer) = getThreadedComment(s, CellRef(row, col))
getThreadedComment(s::Worksheet, ref::AbstractString) = getThreadedComment(s, _comment_cellref(ref))
function getThreadedComment(s::Worksheet, ref::CellRef)::Union{Nothing,ThreadedCommentInfo}
    r = string(ref)
    i = findfirst(c -> c.ref == r, getThreadedComments(s))
    return isnothing(i) ? nothing : getThreadedComments(s)[i]
end

_part_exists(xf::XLSXFile, path::String) = haskey(xf.data, path) || get(xf.files, path, false)

# The workbook's people (who wrote threaded comments): id => display name.
function _threaded_comment_people(xf::XLSXFile)::Dict{String,String}
    wb = get_workbook(xf)
    has_relationship_by_type(wb, REL_PERSON) || return Dict{String,String}()
    path = get_relationship_target_by_type("xl", wb, REL_PERSON)
    _part_exists(xf, path) || return Dict{String,String}()
    return Dict(get_attr(p, "id") => get_attr(p, "displayName")
                for p in elements_with_tag(xml_root_element(get_xml_data(xf, path)), "person"))
end

function _threaded_text(n::XML.Node)::String
    t = first_element_with_tag(n, "text")
    isnothing(t) && return ""
    return something(XML.is_simple_value(t), join((something(XML.value(c), "") for c in XML.children(t)), ""))
end

function _threaded_date(n::XML.Node)::Union{Nothing,Dates.DateTime}
    d = get_attr(n, "dT", "")
    isempty(d) && return nothing
    d = rstrip(d, 'Z')
    m = match(r"^(.*T\d\d:\d\d:\d\d)(?:\.(\d+))?$", d)
    isnothing(m) && return nothing
    t = tryparse(Dates.DateTime, m[1])
    (isnothing(t) || isnothing(m[2])) && return t
    return t + Dates.Millisecond(round(Int, parse(Float64, "0." * m[2]) * 1000))
end

# ===========================================================================
# Writing
# ===========================================================================

const NS_THREADED = "http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments"
const NS_XR       = "http://schemas.microsoft.com/office/spreadsheetml/2014/revision"
const NS_MARKUP_COMPAT = "http://schemas.openxmlformats.org/markup-compatibility/2006"
const MIME_THREADED_COMMENTS = "application/vnd.ms-excel.threadedcomments+xml"

const _PERSON_PART = (rel=REL_PERSON, path="xl/persons/person.xml",
    ctype="application/vnd.ms-excel.person+xml",
    root=() -> _container("personList", ["xmlns" => NS_THREADED, "xmlns:x" => SPREADSHEET_NAMESPACE_XPATH_ARG]))

# The text Excel puts in a thread's note, for versions of Excel without threads.
const _THREAD_NOTE_HEADER = "[Threaded comment]\n\nYour version of Excel allows you to read this threaded comment; " *
    "however, any edits to it will get removed if the file is opened in a newer version of Excel. " *
    "Learn more: https://go.microsoft.com/fwlink/?linkid=870924\n"

"""
    addThreadedComment(s::Worksheet, ref, text; author, date=now(UTC))
    addThreadedComment(s::Worksheet, row::Integer, col::Integer, text; kw...)

Start a threaded comment (an Excel 365 comment, as opposed to a note) on a single
cell. Excel 365 shows it in the cell's comment thread, where people can reply to
it. Older versions of Excel show it as a note.

`author` is the name shown on the comment. Each name becomes one of the workbook's
people, shared by every comment with the same name. `date` is when it was written,
in UTC.

A cell has one thread: add to one with [`addThreadedReply`](@ref). A cell can't
have both a note and a threaded comment, so a cell with a note
([`addComment`](@ref)) is refused.

The file must be open for writing (`mode="rw"`).

# Return Value

The thread, as [`getThreadedComment`](@ref) returns it.

# Examples

```julia
XLSX.addThreadedComment(sheet, "B3", "Is this figure final?"; author="Tim")
XLSX.addThreadedReply(sheet, "B3", "Yes, checked against the accounts."; author="Ann")
XLSX.setThreadedCommentResolved(sheet, "B3", true)
```

See also [`getThreadedComments`](@ref), [`deleteThreadedComment`](@ref).
"""
addThreadedComment(s::Worksheet, row::Integer, col::Integer, text::AbstractString; kw...) =
    addThreadedComment(s, CellRef(row, col), text; kw...)
addThreadedComment(s::Worksheet, ref::AbstractString, text::AbstractString; kw...) =
    addThreadedComment(s, _comment_cellref(ref), text; kw...)
function addThreadedComment(s::Worksheet, ref::CellRef, text::AbstractString;
                            author::AbstractString, date::Dates.DateTime=Dates.now(Dates.UTC))
    xf, sheet_path = _thread_sheet(s, "add a threaded comment", text, author)
    r = string(ref)
    isempty(_thread_nodes(xf, sheet_path, r)) ||
        throw(XLSXError("Cell $r already has a threaded comment. Add to it with `addThreadedReply`."))
    isnothing(getComment(s, ref)) ||
        throw(XLSXError("Cell $r has a note. A cell can't have both: delete it first with `deleteComment`."))
    _add_thread_node!(xf, sheet_path, r, text, author, date, nothing)
    return _thread_written(s, xf, sheet_path, ref)
end

"""
    addThreadedReply(s::Worksheet, ref, text; author, date=now(UTC))
    addThreadedReply(s::Worksheet, row::Integer, col::Integer, text; kw...)

Add a reply to the end of the threaded comment on a cell (see
[`addThreadedComment`](@ref)). The cell must already have a thread.

The file must be open for writing (`mode="rw"`).

# Return Value

The thread, as [`getThreadedComment`](@ref) returns it.
"""
addThreadedReply(s::Worksheet, row::Integer, col::Integer, text::AbstractString; kw...) =
    addThreadedReply(s, CellRef(row, col), text; kw...)
addThreadedReply(s::Worksheet, ref::AbstractString, text::AbstractString; kw...) =
    addThreadedReply(s, _comment_cellref(ref), text; kw...)
function addThreadedReply(s::Worksheet, ref::CellRef, text::AbstractString;
                          author::AbstractString, date::Dates.DateTime=Dates.now(Dates.UTC))
    xf, sheet_path = _thread_sheet(s, "add a threaded reply", text, author)
    first = _thread_first(xf, sheet_path, string(ref))
    _add_thread_node!(xf, sheet_path, string(ref), text, author, date, get_attr(first, "id"))
    return _thread_written(s, xf, sheet_path, ref)
end

"""
    setThreadedCommentResolved(s::Worksheet, ref, resolved::Bool=true)
    setThreadedCommentResolved(s::Worksheet, row::Integer, col::Integer, resolved::Bool=true)

Mark the threaded comment on a cell as resolved, or (with `false`) reopen it.
Excel shows a resolved thread greyed out, and it can't be replied to there until
it's reopened.

The file must be open for writing (`mode="rw"`).

See also [`addThreadedComment`](@ref).
"""
setThreadedCommentResolved(s::Worksheet, row::Integer, col::Integer, resolved::Bool=true) =
    setThreadedCommentResolved(s, CellRef(row, col), resolved)
setThreadedCommentResolved(s::Worksheet, ref::AbstractString, resolved::Bool=true) =
    setThreadedCommentResolved(s, _comment_cellref(ref), resolved)
function setThreadedCommentResolved(s::Worksheet, ref::CellRef, resolved::Bool=true)
    _require_writable(s, "resolve a threaded comment")
    _require_worksheet(s, "resolve a threaded comment")
    first = _thread_first(get_xlsxfile(s), _sheet_path(s), string(ref))
    resolved ? (first["done"] = "1") : remove_attr!(first, "done")
    return nothing
end

"""
    deleteThreadedComment(s::Worksheet, ref)
    deleteThreadedComment(s::Worksheet, row::Integer, col::Integer)

Remove the threaded comment on a cell, with all its replies and the note Excel
keeps for it. Return `true` if there was one and `false` if the cell had none.

The file must be open for writing (`mode="rw"`).

See also [`addThreadedComment`](@ref).
"""
deleteThreadedComment(s::Worksheet, row::Integer, col::Integer) = deleteThreadedComment(s, CellRef(row, col))
deleteThreadedComment(s::Worksheet, ref::AbstractString) = deleteThreadedComment(s, _comment_cellref(ref))
function deleteThreadedComment(s::Worksheet, ref::CellRef)::Bool
    _require_writable(s, "delete a threaded comment")
    _require_worksheet(s, "delete a threaded comment")
    xf, sheet_path, r = get_xlsxfile(s), _sheet_path(s), string(ref)
    found = false
    for (_, path) in _sheet_rel_targets(xf, sheet_path, REL_THREADED_COMMENT)
        haskey(xf.data, path) || continue
        root = xml_root_element(get_xml_data(xf, path))
        kids = XML.children(root)
        n = length(kids)
        filter!(k -> !(XML.nodetype(k) == XML.Element && localname(k) == "threadedComment" && get_attr(k, "ref") == r), kids)
        found |= length(kids) < n
        isempty(elements_with_tag(root, "threadedComment")) &&
            _remove_sheet_part!(xf, sheet_path, REL_THREADED_COMMENT; target=path)
    end
    found && _remove_note!(xf, sheet_path, ref)
    return found
end

# --- helpers ---------------------------------------------------------------

# The checks every new comment or reply makes. Returns the file and the sheet's part.
function _thread_sheet(s::Worksheet, action::String, text::AbstractString, author::AbstractString)
    _require_writable(s, action)
    _require_worksheet(s, action)
    isempty(text) && throw(XLSXError("A threaded comment's text cannot be empty."))
    isempty(author) && throw(XLSXError("A threaded comment needs an `author`."))
    return get_xlsxfile(s), _sheet_path(s)
end

# The comment and replies on cell `ref`, in the file's order.
_thread_nodes(xf::XLSXFile, sheet_path::String, ref::String)::Vector{XML.Node} =
    [n for (_, path) in _sheet_rel_targets(xf, sheet_path, REL_THREADED_COMMENT) if haskey(xf.data, path)
       for n in elements_with_tag(xml_root_element(get_xml_data(xf, path)), "threadedComment")
       if get_attr(n, "ref") == ref]

# The comment that starts the thread on `ref`.
function _thread_first(xf::XLSXFile, sheet_path::String, ref::String)::XML.Node
    for n in _thread_nodes(xf, sheet_path, ref)
        isempty(get_attr(n, "parentId", "")) && return n
    end
    throw(XLSXError("Cell $ref has no threaded comment. Start one with `addThreadedComment`."))
end

# The sheet's threaded comments part, made if it has none.
function _threaded_part!(xf::XLSXFile, sheet_path::String)::XML.Node
    for (_, path) in _sheet_rel_targets(xf, sheet_path, REL_THREADED_COMMENT)
        haskey(xf.data, path) && return xml_root_element(get_xml_data(xf, path))
    end
    i = 1
    while haskey(xf.files, "xl/threadedComments/threadedComment$i.xml") ||
          haskey(xf.data, "xl/threadedComments/threadedComment$i.xml")
        i += 1
    end
    path = "xl/threadedComments/threadedComment$i.xml"
    xf.data[path]  = _part_doc(_container("ThreadedComments",
                                          ["xmlns" => NS_THREADED, "xmlns:x" => SPREADSHEET_NAMESPACE_XPATH_ARG]))
    xf.files[path] = true
    add_part_rel!(xf, sheet_path, path, REL_THREADED_COMMENT)
    register_content_type!(xf, "[Content_Types].xml";
                           tag="Override", key="PartName", val="/$path", content_type=MIME_THREADED_COMMENTS)
    return xml_root_element(get_xml_data(xf, path))
end

# The id of the person called `name`, added to the workbook's people if new. A
# person XLSX.jl adds isn't tied to an account, as for a file written offline.
function _person_id!(xf::XLSXFile, name::AbstractString)::String
    root = _workbook_part!(xf, _PERSON_PART)
    people = elements_with_tag(root, "person")
    for p in people
        get_attr(p, "displayName") == name && return get_attr(p, "id")
    end
    id = _unique_guid(xf.uuid_rng, Set(get_attr(p, "id") for p in people))
    el = _xml_element(_root_pfx(root) * "person",
                      ["displayName" => name, "id" => id, "userId" => name, "providerId" => "None"])
    _insert_before_first!(root, el, k -> localname(k) == "person" && get_attr(k, "displayName") > name)   # by name, as Excel keeps them
    return id
end

# Insert `el` among `parent`'s children before the first element for which `after`
# is true, or at the end.
function _insert_before_first!(parent::XML.Node, el::XML.Node, after::Function)
    kids = XML.children(parent)
    i = findfirst(k -> XML.nodetype(k) == XML.Element && after(k), kids)
    isnothing(i) ? push!(kids, el) : insert!(kids, i, el)
    return nothing
end

# Every comment id the workbook's threads use, so a new one is unique.
_thread_ids(xf::XLSXFile) = Set(get_attr(n, "id") for (k, doc) in xf.data
    if startswith(k, "xl/threadedComments/")
    for n in elements_with_tag(xml_root_element(doc), "threadedComment"))

# Add a first comment (`parent === nothing`) or a reply to the thread on `ref`, in
# Excel's order: threads by cell, row by row, and a reply after the thread's last.
function _add_thread_node!(xf::XLSXFile, sheet_path::String, ref::String, text::AbstractString,
                           author::AbstractString, date::Dates.DateTime, parent::Union{Nothing,String})
    root = _threaded_part!(xf, sheet_path)
    pfx  = _root_pfx(root)
    atts = ["ref" => ref, "dT" => _thread_date_string(date), "personId" => _person_id!(xf, author),
            "id" => _unique_guid(xf.uuid_rng, _thread_ids(xf))]
    isnothing(parent) || push!(atts, "parentId" => parent)
    el = _xml_element(pfx * "threadedComment", atts, [XML.Element(pfx * "text", XML.Text(String(text)))])
    c = CellRef(ref)
    later(k) = localname(k) == "threadedComment" && is_valid_cellname(get_attr(k, "ref")) &&
        (k_ref = CellRef(get_attr(k, "ref"));
         (row_number(k_ref), column_number(k_ref)) > (row_number(c), column_number(c)))
    _insert_before_first!(root, el, later)
    return nothing
end

# Excel's form: to the hundredth of a second, with no time zone.
_thread_date_string(d::Dates.DateTime) =
    Dates.format(d, "yyyy-mm-ddTHH:MM:SS") * "." * lpad(Dates.millisecond(d) ÷ 10, 2, '0')

# After a change to the thread on `ref`: rewrite its note for older versions of
# Excel, and return the thread.
function _thread_written(s::Worksheet, xf::XLSXFile, sheet_path::String, ref::CellRef)
    nodes = _thread_nodes(xf, sheet_path, string(ref))
    first = only(n for n in nodes if isempty(get_attr(n, "parentId", "")))
    io = IOBuffer()
    print(io, _THREAD_NOTE_HEADER)
    for n in nodes
        print(io, n === first ? "\nComment:\n    " : "\nReply:\n    ", _threaded_text(n))
    end
    id = get_attr(first, "id")
    _put_note!(xf, sheet_path, ref, "tc=$id", String(take!(io)); uid=id)
    return getThreadedComment(s, ref)
end

# The prefix of namespace `ns` on `root`, declared as `pfx` if it isn't, and listed
# as one a reader may ignore (`mc:Ignorable`), as Excel's extensions are.
function _ignorable_ns!(root::XML.Node, pfx::String, ns::String)::String
    p  = _declare_ns!(root, pfx, ns)
    mc = _declare_ns!(root, "mc", NS_MARKUP_COMPAT)
    ignorable = split(get_attr(root, "$mc:Ignorable", ""))
    p in ignorable || (root["$mc:Ignorable"] = join(vcat(ignorable, p), " "))
    return p
end
