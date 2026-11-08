# ===========================================================================
# AutoFilter on a plain range
#
# A sheet has at most one `<autoFilter ref="...">`, giving the drop-down filter
# buttons on the first row of the range. Excel also keeps the range in a hidden
# name scoped to the sheet, `_xlnm._FilterDatabase`. An Excel Table has a filter
# of its own, in the table part, so a sheet filter may not overlap a table.
# ===========================================================================

const FILTER_DATABASE = "_xlnm._FilterDatabase"

"""
    setAutoFilter(ws::Worksheet, rng)
    setAutoFilter(xf::XLSXFile, "Sheet1!A1:D20")

Add filter buttons to the first row of `rng`, as Excel's *Data → Filter* does,
so its columns can be filtered and sorted in Excel. `rng` is a cell range
(`"A1:D20"`), a defined name, or `:` for the sheet's whole used range. Its first
row is the header row.

A sheet has one filter, so this replaces any filter the sheet already has,
together with the criteria set on it. The filter may not overlap an Excel Table,
which has a filter of its own (see [`addtable!`](@ref)).

The file must be open for writing (`mode="rw"`).

# Examples
```julia
julia> XLSX.setAutoFilter(sheet, "A1:F200")

julia> XLSX.setAutoFilter(sheet, :)
```

See also [`getAutoFilter`](@ref) and [`removeAutoFilter`](@ref).
"""
function setAutoFilter end
setAutoFilter(xf::XLSXFile, ref::AbstractString) = setAutoFilter(_rowcol_sheet(xf, ref)...)
setAutoFilter(ws::Worksheet, ::Colon) = setAutoFilter(ws, get_dimension(ws))
setAutoFilter(ws::Worksheet, rng::SheetCellRange) = do_sheet_names_match(ws, rng) && setAutoFilter(ws, rng.rng)
function setAutoFilter(ws::Worksheet, ref::AbstractString)
    parts = _rowcol_parts(ws, ref)
    (length(parts) == 1 && only(parts) isa CellRange) ||
        throw(XLSXError("An AutoFilter needs a single range of cells, such as `A1:D20`. Found `$ref`."))
    return setAutoFilter(ws, only(parts))
end
function setAutoFilter(ws::Worksheet, rng::CellRange)
    _require_writable(ws, "set an AutoFilter")
    _require_worksheet(ws, "set an AutoFilter")
    for t in tables(ws)
        intersects(rng, t.ref) &&
            throw(XLSXError("AutoFilter range `$rng` overlaps table `$(t.name)` (`$(t.ref)`), which has a filter of its own."))
    end

    root = xml_root_element(get_xml_data(get_xlsxfile(ws), _sheet_path(ws)))
    chn = XML.children(root)
    idx = findfirst(c -> localname(c) == "autoFilter", chn)
    if isnothing(idx) || get(chn[idx], "ref", "") != string(rng)   # the same range keeps its criteria
        isnothing(idx) || deleteat!(chn, idx)
        pfx = get_prefix(ws)
        el = XML.Element((pfx == "" ? "" : pfx * ":") * "autoFilter"; ref=string(rng))
        insert!(chn, insert_index(root, "autoFilter", WORKSHEET_ORDER) + 1, el)
    end
    get_workbook(ws).worksheet_names[(ws.sheetId, FILTER_DATABASE)] =
        DefinedNameValue(SheetCellRange(ws.name, rng), true, true)

    update_worksheets_xml!(get_xlsxfile(ws))
    return nothing
end

"""
    getAutoFilter(ws::Worksheet) -> Union{Nothing, CellRange}

The range of the sheet's filter (see [`setAutoFilter`](@ref)), or `nothing` if
it has none. The filters of Excel Tables are not included.
"""
function getAutoFilter(ws::Worksheet)::Union{Nothing,CellRange}
    is_chartsheet(get_workbook(ws), ws.name) && return nothing
    for el in _sheet_top_elements(get_xlsxfile(ws), _sheet_path(ws), ("autoFilter",))
        ref = get(el, "ref", nothing)
        isnothing(ref) || return CellRange(ref)
    end
    return nothing
end

"""
    removeAutoFilter(ws::Worksheet) -> Bool

Remove the sheet's filter and its criteria. Returns `false` if it had none.
Rows the filter hid stay hidden; show them with [`setRowHidden`](@ref).

The file must be open for writing (`mode="rw"`).
"""
function removeAutoFilter(ws::Worksheet)::Bool
    _require_writable(ws, "remove an AutoFilter")
    _require_worksheet(ws, "remove an AutoFilter")
    delete!(get_workbook(ws).worksheet_names, (ws.sheetId, FILTER_DATABASE))
    root = xml_root_element(get_xml_data(get_xlsxfile(ws), _sheet_path(ws)))
    idx = findfirst(c -> localname(c) == "autoFilter", XML.children(root))
    isnothing(idx) && return false
    deleteat!(XML.children(root), idx)
    update_worksheets_xml!(get_xlsxfile(ws))
    return true
end
