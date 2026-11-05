# ===========================================================================
# Cell hyperlinks
#
# A hyperlink is a `<hyperlink>` in the worksheet's `<hyperlinks>` block. A link
# to a place in the workbook names it in its `location` attribute. A link to
# anything else (a web page, an email address, a file) holds an `r:id` naming a
# relationship of the sheet with `TargetMode="External"`, whose `Target` is the
# URL. The cell's own value and font are independent of the link: Excel shows
# the cell's text, and styles it as a link only by convention.
# ===========================================================================

const REL_HYPERLINK =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"

# Excel refuses a hyperlink address longer than this.
const EXCEL_MAX_HYPERLINK_LENGTH = 2079

# Office's default color for the "Hyperlink" cell style (the theme's `hlink`).
const HYPERLINK_FONT_COLOR = "FF0563C1"

const HyperlinkInfo = NamedTuple{
    (:sheet, :ref, :url, :location, :display, :tooltip),
    Tuple{String, String, Union{Nothing,String}, Union{Nothing,String}, Union{Nothing,String}, Union{Nothing,String}},
}

# ===========================================================================
# Public API
# ===========================================================================

"""
    addHyperlink(s::Worksheet, ref; url=nothing, location=nothing, display=nothing, tooltip=nothing)
    addHyperlink(s::Worksheet, row::Integer, col::Integer; kw...)

Make a single cell a hyperlink. Give exactly one of `url` or `location`:

- `url`: an address outside the workbook, such as a web page
  (`"https://julialang.org"`), an email address (`"mailto:someone@example.com"`)
  or a file (`"report.pdf"`, relative to the workbook's folder).
- `location`: a place in this workbook, such as a cell (`"Sheet2!A1"`, or
  `"'My Sheet'!B4"` for a sheet name with spaces) or a defined name.

If the cell is empty, its value is set to `display` (or to the `url` or
`location` if `display` is not given), and it is given a blue, underlined font,
as Excel does when you insert a link. A cell that already holds a value keeps its
value and its formatting.

`tooltip` is the text Excel shows when the mouse is over the link. `display` is
also stored with the link itself.

If the cell already has a hyperlink, it is replaced.

Hyperlinks can only be added to a single cell. Giving a cell range is an error.
Adding a hyperlink to a cell inside a range that already has one is also an
error. Delete the range's link first.

# Return Value

A `NamedTuple` `(sheet, ref, url, location, display, tooltip)`. Each field other
than `sheet` and `ref` is a `String` or `nothing`.

# Examples

```julia
XLSX.addHyperlink(sheet, "A1"; url="https://julialang.org", display="Julia")
XLSX.addHyperlink(sheet, "A2"; url="mailto:team@example.com", tooltip="Email the team")
XLSX.addHyperlink(sheet, 3, 1; location="Summary!B2", display="Go to summary")
```

See also [`getHyperlink`](@ref), [`getHyperlinks`](@ref), [`deleteHyperlink`](@ref).
"""
addHyperlink(s::Worksheet, row::Integer, col::Integer; kw...) = addHyperlink(s, CellRef(row, col); kw...)
addHyperlink(s::Worksheet, ref::AbstractString; kw...) = addHyperlink(s, _hyperlink_cellref(ref); kw...)
function addHyperlink(
    s::Worksheet,
    ref::CellRef;
    url::Union{Nothing,AbstractString} = nothing,
    location::Union{Nothing,AbstractString} = nothing,
    display::Union{Nothing,AbstractString} = nothing,
    tooltip::Union{Nothing,AbstractString} = nothing,
)
    _require_writable(s, "add a hyperlink")
    _require_worksheet(s, "add a hyperlink")
    _require_cache(s, "add a hyperlink")
    isnothing(url) == isnothing(location) &&
        throw(XLSXError("Give exactly one of `url` (an address outside the workbook) or `location` (a place in it)."))
    target = something(url, location)
    isempty(target) && throw(XLSXError("A hyperlink's `$(isnothing(url) ? "location" : "url")` cannot be empty."))
    length(target) > EXCEL_MAX_HYPERLINK_LENGTH &&
        throw(XLSXError("A hyperlink address can be at most $EXCEL_MAX_HYPERLINK_LENGTH characters long. Found $(length(target))."))

    xf         = get_xlsxfile(s)
    sheet_path = _sheet_path(s)
    r          = string(ref)

    for n in _hyperlink_nodes(xf, sheet_path)
        n_ref = get_attr(n, "ref")
        n_ref == r && continue   # replaced below
        ref in _hyperlink_range(n_ref) &&
            throw(XLSXError("Cell `$r` is inside the hyperlink on `$n_ref`. Delete that hyperlink first."))
    end
    _delete_hyperlink_node!(xf, sheet_path, r)

    sheet_root = xml_root_element(get_xml_data(xf, sheet_path))
    pfx = _root_prefix(sheet_root)
    el = XML.Element(prefixed_tag(pfx, "hyperlink"))
    el["ref"] = r
    if isnothing(url)
        el["location"] = location
    else
        _ensure_r_namespace!(sheet_root)
        el["r:id"] = add_external_rel!(xf, sheet_path, url, REL_HYPERLINK)
    end
    isnothing(tooltip) || (el["tooltip"] = tooltip)
    isnothing(display) || (el["display"] = display)
    push!(_ensure_hyperlinks_element!(sheet_root), el)

    if getcell(s, ref) isa EmptyCell
        s[ref] = something(display, target)
        setFont(s, ref; color=HYPERLINK_FONT_COLOR, under="single")
    end

    return HyperlinkInfo((s.name, r, _str_or_nothing(url), _str_or_nothing(location),
                          _str_or_nothing(display), _str_or_nothing(tooltip)))
end

"""
    getHyperlink(s::Worksheet, ref) -> Union{Nothing, NamedTuple}
    getHyperlink(s::Worksheet, row::Integer, col::Integer) -> Union{Nothing, NamedTuple}

Return the hyperlink on the given cell as a `NamedTuple`
`(sheet, ref, url, location, display, tooltip)`, or `nothing` if the cell has no
hyperlink. A link on a range of cells, which Excel can write, is returned for
any cell in the range, with the range as its `ref`.

`url` is set for a link outside the workbook and `location` for a link to a
place in it. Each is `nothing` when not used, as are `display` and `tooltip` when
the link doesn't have them.

The file doesn't need to be opened for writing.

See also [`addHyperlink`](@ref), [`getHyperlinks`](@ref).
"""
getHyperlink(s::Worksheet, row::Integer, col::Integer) = getHyperlink(s, CellRef(row, col))
getHyperlink(s::Worksheet, ref::AbstractString) = getHyperlink(s, _hyperlink_cellref(ref))
function getHyperlink(s::Worksheet, ref::CellRef)::Union{Nothing,HyperlinkInfo}
    for h in getHyperlinks(s)
        ref in _hyperlink_range(h.ref) && return h
    end
    return nothing
end

"""
    getHyperlinks(s::Worksheet) -> Vector{NamedTuple}
    getHyperlinks(xf::XLSXFile) -> Vector{NamedTuple}

Return every hyperlink on a worksheet, or on every worksheet of a workbook, as a
vector of `NamedTuple`s `(sheet, ref, url, location, display, tooltip)`, in the
order the file holds them.

The file doesn't need to be opened for writing.

See also [`addHyperlink`](@ref), [`getHyperlink`](@ref).
"""
function getHyperlinks(s::Worksheet)::Vector{HyperlinkInfo}
    is_chartsheet(get_workbook(s), s.name) && return HyperlinkInfo[]
    xf         = get_xlsxfile(s)
    sheet_path = _sheet_path(s)
    out = HyperlinkInfo[]
    for a in _hyperlink_attributes(xf, sheet_path)
        rid = _prefixed_value(a, "id")
        url = isnothing(rid) ? nothing : _external_rel_target(xf, sheet_path, rid)
        push!(out, HyperlinkInfo((s.name, get(a, "ref", ""), url, get(a, "location", nothing),
                                  get(a, "display", nothing), get(a, "tooltip", nothing))))
    end
    return out
end

function getHyperlinks(xf::XLSXFile)::Vector{HyperlinkInfo}
    wb = get_workbook(xf)
    return reduce(vcat, (getHyperlinks(getsheet(wb, i)) for i in 1:sheetcount(wb)); init=HyperlinkInfo[])
end

"""
    deleteHyperlink(s::Worksheet, ref)
    deleteHyperlink(s::Worksheet, row::Integer, col::Integer)

Remove the hyperlink on the given cell. Return `true` if a hyperlink was removed
and `false` if the cell had none. If the cell is part of a link on a range of
cells, the link is removed from the whole range.

The cell's value and formatting are left as they are.

See also [`addHyperlink`](@ref).
"""
deleteHyperlink(s::Worksheet, row::Integer, col::Integer) = deleteHyperlink(s, CellRef(row, col))
deleteHyperlink(s::Worksheet, ref::AbstractString) = deleteHyperlink(s, _hyperlink_cellref(ref))
function deleteHyperlink(s::Worksheet, ref::CellRef)::Bool
    _require_writable(s, "delete a hyperlink")
    _require_worksheet(s, "delete a hyperlink")
    xf         = get_xlsxfile(s)
    sheet_path = _sheet_path(s)
    for n in _hyperlink_nodes(xf, sheet_path)
        n_ref = get_attr(n, "ref")
        ref in _hyperlink_range(n_ref) || continue
        return _delete_hyperlink_node!(xf, sheet_path, n_ref)
    end
    return false
end

# ===========================================================================
# Helpers
# ===========================================================================

_hyperlink_cellref(ref::AbstractString)::CellRef = _single_cellref(ref, "Hyperlinks")

_hyperlink_range(ref::AbstractString)::CellRange =
    is_valid_cellname(ref) ? CellRange(CellRef(ref), CellRef(ref)) : CellRange(ref)

_str_or_nothing(x) = isnothing(x) ? nothing : String(x)

# The value of the attribute in `a` whose local name is `key` and which has a
# namespace prefix (`r:id`, whatever the prefix), or `nothing`.
function _prefixed_value(a::AbstractDict, key::AbstractString)::Union{Nothing,String}
    for (k, v) in a
        occursin(':', k) && has_localname(k, key) && return v
    end
    return nothing
end

function _ensure_r_namespace!(sheet_root::XML.Node)
    haskey(something(XML.attributes(sheet_root), Dict()), "xmlns:r") || (sheet_root["xmlns:r"] = NS_R)
    return nothing
end

# The `<hyperlink>` elements of a sheet that is open for writing.
function _hyperlink_nodes(xf::XLSXFile, sheet_path::String)::Vector{XML.Node}
    block = first_element_with_tag(xml_root_element(get_xml_data(xf, sheet_path)), "hyperlinks")
    return isnothing(block) ? XML.Node[] : elements_with_tag(block, "hyperlink")
end

function _ensure_hyperlinks_element!(sheet_root::XML.Node)::XML.Node
    block = first_element_with_tag(sheet_root, "hyperlinks")
    isnothing(block) || return block
    block = XML.Element(prefixed_tag(_root_prefix(sheet_root), "hyperlinks"))
    insert!(XML.children(sheet_root), insert_index(sheet_root, "hyperlinks", WORKSHEET_ORDER) + 1, block)
    return block
end

# Remove the `<hyperlink>` whose `ref` is exactly `ref`, the relationship it alone
# used, and the `<hyperlinks>` block if it is left empty. Return whether one was removed.
function _delete_hyperlink_node!(xf::XLSXFile, sheet_path::String, ref::String)::Bool
    sheet_root = xml_root_element(get_xml_data(xf, sheet_path))
    block = first_element_with_tag(sheet_root, "hyperlinks")
    isnothing(block) && return false
    kids = XML.children(block)
    i = findfirst(n -> XML.nodetype(n) == XML.Element && localname(n) == "hyperlink" && get_attr(n, "ref") == ref, kids)
    isnothing(i) && return false
    rid = _prefixed_value(something(XML.attributes(kids[i]), Dict{String,String}()), "id")
    deleteat!(kids, i)
    if !isnothing(rid)
        still_used = any(n -> _prefixed_value(something(XML.attributes(n), Dict{String,String}()), "id") == rid,
                         elements_with_tag(block, "hyperlink"))
        still_used || _remove_sheet_rel!(xf, sheet_path, rid)
    end
    if isempty(elements_with_tag(block, "hyperlink"))
        _remove_sheet_element!(xf, sheet_path, "hyperlinks")
    end
    return true
end

# Remove the sheet's relationship `rid`, without touching what it points at.
function _remove_sheet_rel!(xf::XLSXFile, sheet_path::String, rid::String)
    rels_path = _sheet_rels_path(sheet_path)
    haskey(xf.data, rels_path) || return nothing
    kids = XML.children(xml_root_element(get_xml_data(xf, rels_path)))
    i = findfirst(n -> XML.nodetype(n) == XML.Element && get_attr(n, "Id") == rid, kids)
    isnothing(i) || deleteat!(kids, i)
    return nothing
end

# The `Target` of the sheet's external relationship `rid`, unresolved: it is a URL,
# not a part of the package.
function _external_rel_target(xf::XLSXFile, sheet_path::String, rid::String)::Union{Nothing,String}
    rels_path = _sheet_rels_path(sheet_path)
    haskey(xf.data, rels_path) || return nothing
    for n in elements_with_tag(xml_root_element(get_xml_data(xf, rels_path)), "Relationship")
        get_attr(n, "Id") == rid && return get_attr(n, "Target")
    end
    return nothing
end

# The attributes of each `<hyperlink>` in the sheet, read without disturbing how the
# sheet is held: a sheet open only for reading is kept as raw XML (or not kept at
# all, without the cache) so its cells can be read lazily. Those are scanned with a
# cursor rather than parsed into a tree.
function _hyperlink_attributes(xf::XLSXFile, sheet_path::String)::Vector{Dict{String,String}}
    out = Dict{String,String}[]
    doc = haskey(xf.data, sheet_path) ? xf.data[sheet_path] : open_internal_file_stream(xf, sheet_path)
    if doc isa XML.Node
        block = first_element_with_tag(xml_root_element(doc), "hyperlinks")
        isnothing(block) && return out
        for n in elements_with_tag(block, "hyperlink")
            push!(out, Dict{String,String}(something(XML.attributes(n), Dict{String,String}())))
        end
        return out
    end

    lz = doc isa XML.LazyNode ? doc : parse(doc, XML.LazyNode)
    c = XML.Cursor(xml_root_element(lz))
    XML.next!(c)   # the <worksheet> element itself, depth 1
    in_block = false
    while XML.next!(c) !== nothing
        d = XML.depth(c)
        d <= 1 && break
        XML.nodetype(c) == XML.Element || continue
        if d == 2
            in_block && break   # past <hyperlinks>: there is only one
            if localname(c) == "hyperlinks"
                in_block = true
            else
                XML.skip_element!(c)
            end
        elseif d == 3 && in_block && localname(c) == "hyperlink"
            push!(out, Dict{String,String}(String(k) => String(v) for (k, v) in XML.eachattribute(XML.LazyNode(c))))
        end
    end
    return out
end

# A copied sheet's `<hyperlinks>` keeps the original's `r:id`s, but the copy has
# relationships of its own. Give each external link a relationship in the copy,
# to the same URL, and point the link at it.
function _copy_hyperlink_rels!(xf::XLSXFile, from_sheet::String, to_sheet::String)
    sheet_root = xml_root_element(get_xml_data(xf, to_sheet))
    block = first_element_with_tag(sheet_root, "hyperlinks")
    isnothing(block) && return nothing
    kids = XML.children(block)
    for i in reverse(eachindex(kids))
        n = kids[i]
        XML.nodetype(n) == XML.Element && localname(n) == "hyperlink" || continue
        a = something(XML.attributes(n), Dict{String,String}())
        ks = [k for k in keys(a) if occursin(':', k) && has_localname(k, "id")]
        isempty(ks) && continue
        url = _external_rel_target(xf, from_sheet, a[only(ks)])
        if isnothing(url)
            deleteat!(kids, i)   # dangling in the original too: a link to nowhere isn't worth copying
        else
            _ensure_r_namespace!(sheet_root)
            n[only(ks)] = add_external_rel!(xf, to_sheet, url, REL_HYPERLINK)
        end
    end
    isempty(elements_with_tag(block, "hyperlink")) && _remove_sheet_element!(xf, to_sheet, "hyperlinks")
    return nothing
end
