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

Make a cell, or a range of cells such as `"A1:C3"`, a hyperlink. Give exactly
one of `url` or `location`:

- `url`: an address outside the workbook, such as a web page
  (`"https://julialang.org"`), an email address (`"mailto:someone@example.com"`)
  or a file (`"report.pdf"`, relative to the workbook's folder).
- `location`: a place in this workbook, such as a cell (`"Sheet2!A1"`, or
  `"'My Sheet'!B4"` for a sheet name with spaces) or a defined name.

If the cell is empty, its value is set to `display` (or to the `url` or
`location` if `display` is not given), and it is given a blue, underlined font,
as Excel does when you insert a link. A cell that already holds a value keeps its
value and its formatting.

A range has one link, which works from any of its cells. As Excel does, the text
goes in the range's top-left cell, if that cell is empty, and every empty cell in
the range gets the blue, underlined font.

`tooltip` is the text Excel shows when the mouse is over the link. `display` is
also stored with the link itself.

If the cell or range already has a hyperlink of its own, it is replaced. A
hyperlink may not overlap another one: delete the other link first.

# Return Value

A `NamedTuple` `(sheet, ref, url, location, display, tooltip)`. Each field other
than `sheet` and `ref` is a `String` or `nothing`.

# Examples

```julia
XLSX.addHyperlink(sheet, "A1"; url="https://julialang.org", display="Julia")
XLSX.addHyperlink(sheet, "A2"; url="mailto:team@example.com", tooltip="Email the team")
XLSX.addHyperlink(sheet, 3, 1; location="Summary!B2", display="Go to summary")
XLSX.addHyperlink(sheet, "C1:E2"; url="https://example.com/report", display="Full report")
```

See also [`getHyperlink`](@ref), [`getHyperlinks`](@ref), [`deleteHyperlink`](@ref).
"""
addHyperlink(s::Worksheet, row::Integer, col::Integer; kw...) = addHyperlink(s, CellRef(row, col); kw...)
addHyperlink(s::Worksheet, ref::AbstractString; kw...) = addHyperlink(s, _hyperlink_ref(ref); kw...)
addHyperlink(s::Worksheet, ref::CellRef; kw...) = addHyperlink(s, CellRange(ref, ref); kw...)
function addHyperlink(
    s::Worksheet,
    rng::CellRange;
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
    r          = _hyperlink_refstring(rng)

    for n in _hyperlink_nodes(xf, sheet_path)
        n_ref = get_attr(n, "ref")
        n_ref == r && continue   # replaced below
        intersects(rng, _hyperlink_range(n_ref)) &&
            throw(XLSXError("`$r` overlaps the hyperlink on `$n_ref`. Delete that hyperlink first."))
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

    _hyperlink_cell_text!(s, rng, something(display, target))

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
function getHyperlink(s::Worksheet, ref::AbstractString)
    rng = _hyperlink_ref(ref)
    rng.start == rng.stop || throw(XLSXError("`getHyperlink` reads one cell, not a range: `$ref`. Use `getHyperlinks` for every link."))
    return getHyperlink(s, rng.start)
end
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

Remove the hyperlink on the given cell, or every hyperlink that overlaps the
given range (`"A1:C3"`). Return `true` if a hyperlink was removed and `false` if
there was none. A link on a range of cells is removed from its whole range.

The cell's value and formatting are left as they are.

See also [`addHyperlink`](@ref).
"""
deleteHyperlink(s::Worksheet, row::Integer, col::Integer) = deleteHyperlink(s, CellRef(row, col))
deleteHyperlink(s::Worksheet, ref::AbstractString) = deleteHyperlink(s, _hyperlink_ref(ref))
deleteHyperlink(s::Worksheet, ref::CellRef) = deleteHyperlink(s, CellRange(ref, ref))
function deleteHyperlink(s::Worksheet, rng::CellRange)::Bool
    _require_writable(s, "delete a hyperlink")
    _require_worksheet(s, "delete a hyperlink")
    xf         = get_xlsxfile(s)
    sheet_path = _sheet_path(s)
    refs = [get_attr(n, "ref") for n in _hyperlink_nodes(xf, sheet_path)
            if intersects(rng, _hyperlink_range(get_attr(n, "ref")))]
    foreach(r -> _delete_hyperlink_node!(xf, sheet_path, r), refs)
    return !isempty(refs)
end

# ===========================================================================
# Helpers
# ===========================================================================

# A cell (`"B2"`) or a range of cells (`"A1:C3"`), as a range.
function _hyperlink_ref(ref::AbstractString)::CellRange
    is_valid_cellname(ref) && return CellRange(CellRef(ref), CellRef(ref))
    is_valid_cellrange(ref) && return CellRange(ref)
    throw(XLSXError("Invalid cell or range reference: `$ref`."))
end

# A one-cell range is written as the cell, as Excel writes it.
_hyperlink_refstring(rng::CellRange)::String = rng.start == rng.stop ? string(rng.start) : string(rng)

# Give the link's text to the range's top-left cell, if it is empty, and the
# link font to every empty cell in the range, as Excel does. Cells holding a
# value keep it and their formatting.
function _hyperlink_cell_text!(s::Worksheet, rng::CellRange, text::AbstractString)
    font_of = Dict{UInt32,UInt32}()   # each empty cell's style, and that style with the link font
    for ref in [ref for ref in rng if getcell(s, ref) isa EmptyCell]
        s[ref] = ref == rng.start ? text : missing
        cell = getcell(s, ref)
        old = cell.style
        if haskey(font_of, old)
            cell.style = font_of[old]
        else
            setFont(s, ref; color=HYPERLINK_FONT_COLOR, under="single")
            font_of[old] = getcell(s, ref).style
        end
    end
    return nothing
end

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

# The attributes of each `<hyperlink>` in the sheet, on a sheet open for reading only too.
_hyperlink_attributes(xf::XLSXFile, sheet_path::String)::Vector{Dict{String,String}} =
    [Dict{String,String}(something(XML.attributes(n), Dict{String,String}()))
     for block in _sheet_top_elements(xf, sheet_path, ("hyperlinks",))
     for n in elements_with_tag(block, "hyperlink")]

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
