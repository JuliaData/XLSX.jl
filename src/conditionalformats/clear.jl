"""
    clearConditionalFormats(ws::Worksheet, rng) -> Int
    clearConditionalFormats(ws::Worksheet, :) -> Int

Remove every conditional format whose range falls entirely within `rng`, returning
the number of rules removed. With `:`, remove every conditional format on the sheet.

`rng` is a cell range (`"B1:B10"`), whole columns (`"B:D"`) or rows (`"2:5"`),
several of these separated by commas, or a `CellRange`. Whole columns and rows
cover the full height or width of the sheet.

A rule is removed when all of the cells it applies to lie inside `rng`. A rule
spanning a wider range is left alone, even where it overlaps.

`clearConditionalFormats` is public but not exported: call it as
`XLSX.clearConditionalFormats`.

# Examples

```julia
julia> XLSX.setConditionalFormat(sheet, "B1:B10", :cellIs; operator="greaterThan", value="5");

julia> XLSX.clearConditionalFormats(sheet, "B1:B10")
1

julia> XLSX.clearConditionalFormats(sheet, "B:B")    # every rule within column B

julia> XLSX.clearConditionalFormats(sheet, :)        # every rule on the sheet
```

See also [`setConditionalFormat`](@ref), [`getConditionalFormats`](@ref).
"""
function clearConditionalFormats end

clearConditionalFormats(ws::Worksheet, rng::AbstractString) =
    clearConditionalFormats(ws, _clear_range(ws, rng))
clearConditionalFormats(ws::Worksheet, ::Colon) =
    clearConditionalFormats(ws, CellRange(CellRef(1, 1), CellRef(EXCEL_MAX_ROWS, EXCEL_MAX_COLS)))

# The area to clear: cells and ranges as given, whole columns and rows across the
# full sheet. Several parts, comma-separated, make a NonContiguousRange.
function _clear_range(ws::Worksheet, s::AbstractString)::CfRange
    parts = Union{CellRef,CellRange}[]
    for p in split(s, ',')
        p = strip(p)
        if is_valid_column_range(p)
            c = ColumnRange(p)
            push!(parts, CellRange(CellRef(1, c.start), CellRef(EXCEL_MAX_ROWS, c.stop)))
        elseif is_valid_row_range(p)
            r = RowRange(p)
            push!(parts, CellRange(CellRef(r.start, 1), CellRef(r.stop, EXCEL_MAX_COLS)))
        elseif is_valid_cellname(p)
            c = CellRef(p)
            push!(parts, CellRange(c, c))
        elseif is_valid_cellrange(p)
            push!(parts, CellRange(p))
        else
            throw(XLSXError("`$p` is not a cell range, a column range or a row range."))
        end
    end
    return length(parts) == 1 ? only(parts) : NonContiguousRange(ws.name, parts)
end

function clearConditionalFormats(ws::Worksheet, rng::CfRange)
    sheetdoc = xmlroot(get_workbook(ws), ws.relationship_id)
    n = 0

    # 2007 blocks: direct children of <worksheet>
    ws_el = sheetdoc[find_child_index(XML.children(sheetdoc), "worksheet")]
    for i in reverse(eachindex(XML.children(ws_el)))
        c = ws_el[i]
        localname(c) == "conditionalFormatting" || continue
        haskey(c, "sqref") || continue
        if _cf_within(ws, c["sqref"], rng)
            n += length(xml_elements(c))
            deleteat!(ws_el.children, i)
        end
    end

    # x14 counterparts under <extLst>
    blk = _x14_cf_block(sheetdoc)
    if !isnothing(blk)
        for q in reverse(eachindex(XML.children(blk)))
            b = blk[q]
            els = xml_elements(b)
            s = findlast(e -> localname(e) == "sqref", els)
            isnothing(s) && continue
            if _cf_within(ws, XML.simple_value(els[s]), rng)
                deleteat!(blk.children, q)
            end
        end
        # Excel refuses to open a file with an empty `<x14:conditionalFormattings/>`.
        if isempty(xml_elements(blk))
            i, j = get_idces(sheetdoc, "worksheet", "extLst")
            ext = _x14_cf_ext(sheetdoc[i][j])
            filter!(c -> c !== blk, ext.children)
        end
    end

    update_worksheets_xml!(get_xlsxfile(ws))
    return n
end

# Every area named by `sqref` lies inside `rng`.
function _cf_within(ws::Worksheet, sqref::AbstractString, rng::CfRange)
    inner = _cf_range(ws, sqref)
    outer = [a isa CellRef ? CellRange(a, a) : a for a in _cf_areas(rng)]
    for a in _cf_areas(inner)
        r = a isa CellRef ? CellRange(a, a) : a
        any(o -> issubset(r, o), outer) || return false
    end
    return true
end
