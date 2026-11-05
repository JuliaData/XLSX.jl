# ===========================================================================
# Sheet-level settings: visibility and tab colour
#
# A sheet's visibility is the `state` of its `<sheet>` in workbook.xml, held in
# `ws.visibility` and written by `update_workbook_xml!`. Its tab colour is the
# `<tabColor>` in the sheet's own `<sheetPr>`, which worksheets and chartsheets
# both have as their first element.
# ===========================================================================

const SHEET_VISIBILITIES = (:visible, :hidden, :veryHidden)

# `ws.is_hidden` was a field before `visibility` replaced it; keep it working.
function Base.getproperty(ws::Worksheet, s::Symbol)
    s === :is_hidden && return getfield(ws, :visibility) !== :visible
    return getfield(ws, s)
end
function Base.setproperty!(ws::Worksheet, s::Symbol, v)
    if s === :is_hidden
        v || return setfield!(ws, :visibility, :visible)
        getfield(ws, :visibility) === :visible && setfield!(ws, :visibility, :hidden)
        return v
    end
    return setfield!(ws, s, convert(fieldtype(Worksheet, s), v))
end

"""
    setSheetVisibility(ws::Worksheet, visibility::Symbol)

Show or hide a sheet's tab. `visibility` is one of

- `:visible`: the tab is shown;
- `:hidden`: the tab is hidden, and can be shown again from Excel's *Unhide* menu;
- `:veryHidden`: the tab is hidden and not listed in the *Unhide* menu. It can
  be shown again only from VBA, or with this function.

A workbook must keep at least one visible sheet, so hiding the last one is an
error. Hiding the sheet the workbook opens on makes the first visible sheet the
one it opens on instead.

Works on worksheets and chartsheets. The file must be open for writing
(`mode="rw"`).

# Examples
```julia
julia> XLSX.setSheetVisibility(xf["Lookups"], :hidden)

julia> XLSX.setSheetVisibility(xf["Lookups"], :visible)
```

See also [`getSheetVisibility`](@ref).
"""
function setSheetVisibility(ws::Worksheet, visibility::Symbol)
    _require_writable(ws, "change a sheet's visibility")
    visibility in SHEET_VISIBILITIES ||
        throw(XLSXError("Invalid visibility `:$visibility`. Must be one of $(join(SHEET_VISIBILITIES, ", "))."))
    visibility === ws.visibility && return nothing
    wb = get_workbook(ws)
    if visibility !== :visible && ws.visibility === :visible &&
       count(s -> s.visibility === :visible, wb.sheets) == 1
        throw(XLSXError("Cannot hide `$(ws.name)`: it is the workbook's only visible sheet, and Excel needs one."))
    end
    ws.visibility = visibility
    visibility === :visible || _move_active_tab_from!(ws)
    update_workbook_xml!(get_xlsxfile(ws))
    return nothing
end

"""
    getSheetVisibility(ws::Worksheet) -> Symbol

Whether a sheet's tab is shown: `:visible`, `:hidden` or `:veryHidden`. See
[`setSheetVisibility`](@ref).
"""
getSheetVisibility(ws::Worksheet)::Symbol = ws.visibility

# A hidden sheet cannot be the one the workbook opens on: if `ws` is, make the
# first visible sheet that one, moving the selected tab with it.
function _move_active_tab_from!(ws::Worksheet)
    xf = get_xlsxfile(ws)
    wb = get_workbook(ws)
    pos = findfirst(s -> s === ws, wb.sheets) - 1
    root = xml_root_element(xmlroot(xf, "xl/workbook.xml"))
    views = first_element_with_tag(root, "bookViews")
    view = isnothing(views) ? nothing : first_element_with_tag(views, "workbookView")
    active = isnothing(view) ? 0 : parse(Int, get(view, "activeTab", "0"))
    active == pos || return nothing

    new = findfirst(s -> s.visibility === :visible, wb.sheets) - 1
    if isnothing(view)
        pfx = get_prefix("xl/workbook.xml", xf)
        pfx = pfx == "" ? pfx : pfx * ":"
        if isnothing(views)
            views = XML.Element(pfx * "bookViews")
            insert!(XML.children(root), insert_index(root, "bookViews", WORKBOOK_ORDER) + 1, views)
        end
        view = XML.Element(pfx * "workbookView")
        push!(views, view)
    end
    view["activeTab"] = string(new)
    _set_tab_selected!(ws, false)
    _set_tab_selected!(wb.sheets[new + 1], true)
    return nothing
end

function _set_tab_selected!(ws::Worksheet, selected::Bool)
    root = xml_root_element(get_xml_data(get_xlsxfile(ws), _sheet_path(ws)))
    views = first_element_with_tag(root, "sheetViews")
    isnothing(views) && return nothing
    for v in xml_elements(views)
        localname(v) == "sheetView" || continue
        (selected || haskey(v, "tabSelected")) && (v["tabSelected"] = selected ? "1" : "0")
    end
    update_worksheets_xml!(get_xlsxfile(ws))
    return nothing
end

"""
    setTabColor(ws::Worksheet, color)

Colour a sheet's tab. The `color` can be $COLOR_FORMS_DOC

Works on worksheets and chartsheets. The file must be open for writing
(`mode="rw"`).

# Examples
```julia
julia> XLSX.setTabColor(sheet, "FF0000")

julia> XLSX.setTabColor(sheet, :accent2)
```

See also [`getTabColor`](@ref) and [`removeTabColor`](@ref).
"""
function setTabColor(ws::Worksheet, color)
    _require_writable(ws, "set a tab colour")
    atts = _color_atts(color)   # checks the colour before anything changes
    root = xml_root_element(get_xml_data(get_xlsxfile(ws), _sheet_path(ws)))
    idx = findfirst(c -> localname(c) == "sheetPr", XML.children(root))
    if isnothing(idx)   # sheetPr comes first in both worksheets and chartsheets
        pfx = get_prefix(ws)
        pushfirst!(XML.children(root), XML.Element((pfx == "" ? "" : pfx * ":") * "sheetPr"))
        idx = 1
    end
    pr = XML.children(root)[idx]
    pfx = something(_ns_prefix(XML.tag(pr)), "")
    el = XML.Element(pfx * "tabColor"; (Symbol(k) => v for (k, v) in atts)...)
    # Rebuilt rather than changed in place: a new element with no children takes none.
    others = [c for c in XML.children(pr) if localname(c) != "tabColor"]
    new_pr = XML.Element(XML.tag(pr), el, others...)   # tabColor comes first
    for (k, v) in something(XML.attributes(pr), Dict{String,String}())
        new_pr[k] = v
    end
    XML.children(root)[idx] = new_pr
    update_worksheets_xml!(get_xlsxfile(ws))
    return nothing
end

"""
    getTabColor(ws::Worksheet) -> Union{Nothing, String}

The colour of a sheet's tab as an `"FFRRGGBB"` hex string, or `nothing` if it
has none. A theme colour is given as the colour it has in the workbook's theme.

See also [`setTabColor`](@ref).
"""
function getTabColor(ws::Worksheet)::Union{Nothing,String}
    for pr in _sheet_top_elements(get_xlsxfile(ws), _sheet_path(ws), ("sheetPr",)), c in xml_elements(pr)
        localname(c) == "tabColor" || continue
        return resolveColor(get_workbook(ws), Dict{String,String}(something(XML.attributes(c), Dict{String,String}())))
    end
    return nothing
end

"""
    removeTabColor(ws::Worksheet) -> Bool

Remove a sheet's tab colour. Returns `false` if it had none.

The file must be open for writing (`mode="rw"`).
"""
function removeTabColor(ws::Worksheet)::Bool
    _require_writable(ws, "remove a tab colour")
    root = xml_root_element(get_xml_data(get_xlsxfile(ws), _sheet_path(ws)))
    pr = first_element_with_tag(root, "sheetPr")
    isnothing(pr) && return false
    idx = findfirst(c -> localname(c) == "tabColor", XML.children(pr))
    isnothing(idx) && return false
    deleteat!(XML.children(pr), idx)
    if isempty(XML.children(pr)) && isempty(something(XML.attributes(pr), Dict()))
        deleteat!(XML.children(root), findfirst(c -> c === pr, XML.children(root)))
    end
    update_worksheets_xml!(get_xlsxfile(ws))
    return true
end

# "x:" for the tag "x:sheetPr", or nothing for an unprefixed one.
_ns_prefix(tag::AbstractString) = (i = findfirst(':', tag); isnothing(i) ? nothing : tag[1:i])
