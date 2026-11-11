#
# -- Get merged cells
#

"""
    getMergedCells(ws::Worksheet) -> Union{Vector{CellRange}, Nothing}

Return a vector of the `CellRange` of all merged cells in the specified worksheet.
Return nothing if the worksheet contains no merged cells.

The Excel file must be opened in write mode to work with merged cells.

# Examples:
```julia
julia> f = XLSX.readxlsx("test.xlsx")
XLSXFile("C:\\Users\\...\\test.xlsx") containing 1 Worksheet
            sheetname size          range
-------------------------------------------------
               Sheet1 2x2           A1:B2

julia> s = f["Sheet1"]
2×2 Worksheet: ["Sheet1"](A1:B2)

julia> XLSX.getMergedCells(s)
1-element Vector{XLSX.CellRange}:
 B1:B2
 
```
"""
function getMergedCells(ws::Worksheet)::Union{Vector{CellRange},Nothing}
    # May be better if merged cells were part of ws.cache?

    _require_writable(ws, "get merged cells")

    _require_cache(ws, "get merged cells")

    sheetdoc = xmlroot(get_workbook(ws), ws.relationship_id)
    i, j = get_idces(sheetdoc, "worksheet", "mergeCells")

    if isnothing(j) # There are no existing merged cells.
        return nothing
    end

    c = xml_elements(sheetdoc[i][j])
    if length(c) != parse(Int, sheetdoc[i][j]["count"])
        throw(XLSXError("Unexpected number of mergeCells found: $(length(c)). Expected $(sheetdoc[i][j]["count"])."))
    end

    mergedCells = Vector{CellRange}()
    for cell in c
        !haskey(cell, "ref") && throw(XLSXError("No `ref` attribute found in `mergeCell` element."))
        push!(mergedCells, CellRange(cell["ref"]))
    end

    return mergedCells
end

"""
    isMergedCell(ws::Worksheet,  cr::String) -> Bool
    isMergedCell(xf::XLSXFile,   cr::String) -> Bool

    isMergedCell(ws::Worksheet,  row::Int, col::Int) -> Bool

Return `true` if a cell is part of a merged cell range and `false` if not.

Alternatively, if you have already obtained the merged cells for the worksheet,
you can avoid repeated determinations and pass them as a keyword argument to 
the function:

    isMergedCell(ws::Worksheet, cr::String; mergedCells::Union{Vector{CellRange}, Nothing, Missing}=missing) -> Bool
    isMergedCell(xf::XLSXFile,  cr::String; mergedCells::Union{Vector{CellRange}, Nothing, Missing}=missing) -> Bool

    isMergedCell(ws::Worksheet,  row:Int, col::Int; mergedCells::Union{Vector{CellRange}, Nothing, Missing}=missing) -> Bool

The Excel file must be opened in write mode to work with merged cells.

# Examples:
```julia
julia> XLSX.isMergedCell(xf, "Sheet1!A1")

julia> XLSX.isMergedCell(sh, "A1")

julia> XLSX.isMergedCell(sh, 2, 4) # cell D2

julia> mc = XLSX.getMergedCells(sh)

julia> XLSX.isMergedCell(sh, XLSX.CellRef("A1"), mc)
 
```
"""
function isMergedCell end
isMergedCell(xl::XLSXFile, sheetcell::String; kw...)::Bool = process_get_sheetcell(isMergedCell, xl, sheetcell; kw...)
isMergedCell(ws::Worksheet, cr::String; kw...)::Bool = process_get_cellname(isMergedCell, ws, cr; kw...)
isMergedCell(ws::Worksheet, row::Integer, col::Integer; kw...) = isMergedCell(ws, CellRef(row, col); kw...)
function isMergedCell(ws::Worksheet, cellref::CellRef; mergedCells::Union{Vector{CellRange},Nothing,Missing}=missing)::Bool

    _require_writable(ws, "get merged cells")

    _require_cache(ws, "get merged cells")

    if ismissing(mergedCells) # Get mergedCells if missing
        mergedCells = getMergedCells(ws)
    end
    if isnothing(mergedCells) # No merged cells in sheet
        return false
    end
    for rng in mergedCells
        if cellref ∈ rng
            return true
        end
    end

    return false
end

"""
    getMergedBaseCell(ws::Worksheet, cr::String) -> Union{Nothing, NamedTuple{CellRef, Any}}
    getMergedBaseCell(xf::XLSXFile,  cr::String) -> Union{Nothing, NamedTuple{CellRef, Any}}

    getMergedBaseCell(ws::Worksheet, row::Int, col::Int) -> Union{Nothing, NamedTuple{CellRef, Any}}

Return the cell reference and cell value of the base cell of a merged cell range in a worksheet as a named tuple.
If the specified cell is not part of a merged cell range, return `nothing`.

The base cell is the top-left cell of the merged cell range and is the reference cell for the range.

The tuple returned contains:
- `baseCell`  : the reference (`CellRef`) of the base cell
- `baseValue` : the value of the base cell

Additionally, if you have already obtained the merged cells for the worksheet,
you can avoid repeated determinations and pass them as a keyword argument to 
the function:

    getMergedBaseCell(ws::Worksheet, cr::String; mergedCells::Union{Vector{CellRange}, Nothing, Missing}=missing) -> Union{Nothing, NamedTuple{CellRef, Any}}
    getMergedBaseCell(xf::XLSXFile,  cr::String; mergedCells::Union{Vector{CellRange}, Nothing, Missing}=missing) -> Union{Nothing, NamedTuple{CellRef, Any}}

    getMergedBaseCell(ws::Worksheet, row::Int, col::Int; mergedCells::Union{Vector{CellRange}, Nothing, Missing}=missing) -> Union{Nothing, NamedTuple{CellRef, Any}}

The Excel file must be opened in write mode to work with merged cells.

# Examples:
```julia
julia> XLSX.getMergedBaseCell(xf, "Sheet1!B2")
(baseCell = B1, baseValue = 3)

julia> XLSX.getMergedBaseCell(sh, "B2")
(baseCell = B1, baseValue = 3)

julia> XLSX.getMergedBaseCell(sh, 2, 2)
(baseCell = B1, baseValue = 3)

```
"""
function getMergedBaseCell end
getMergedBaseCell(xl::XLSXFile, sheetcell::String; kw...) = process_get_sheetcell(getMergedBaseCell, xl, sheetcell; kw...)
getMergedBaseCell(ws::Worksheet, cr::String; kw...) = process_get_cellname(getMergedBaseCell, ws, cr; kw...)
getMergedBaseCell(ws::Worksheet, row::Integer, col::Integer; kw...) = getMergedBaseCell(ws, CellRef(row, col); kw...)
function getMergedBaseCell(ws::Worksheet, cellref::CellRef; mergedCells::Union{Vector{CellRange},Nothing,Missing}=missing)

    _require_writable(ws, "get merged cells")

    _require_cache(ws, "get merged cells")

    if ismissing(mergedCells) # Get mergedCells if missing
        mergedCells = getMergedCells(ws)
    end
    if isnothing(mergedCells) # No merged cells in sheet
        return nothing
    end
    for rng in mergedCells
        if cellref ∈ rng
            return (; baseCell=rng.start, baseValue=ws[rng.start])
        end
    end
    return nothing
end

"""
    mergeCells(ws::Worksheet, cr::String) -> 0
    mergeCells(xf::XLSXFile,  cr::String) -> 0

    mergeCells(ws::Worksheet, row::Int, col::Int) -> 0

Merge the cells in the range given by `cr`. The value of the merged cell 
will be the value of the first cell in the range (the base cell) prior 
to the merge. All other cells in the range will be set to `missing`, 
reflecting the behaviour of Excel itself.

Merging is limited to the extent of the worksheet dimension.

The specified range must not overlap with any previously merged cells, nor
with an Excel Table (Excel discards a table that overlaps a merged range).

It is not possible to merge a single cell!

A non-contiguous range composed of multiple cell ranges will be processed as a 
list of separate ranges. Each range will be merged separately. No range within 
a non-contiguous range may be a single cell.

The Excel file must be opened in write mode to work with merged cells.

# Examples:
```julia
julia> XLSX.mergeCells(xf, "Sheet1!B2:D3")  # Merge a cell range.

julia> XLSX.mergeCells(sh, 1:3, :)          # Merge rows to the extent of the dimension.

julia> XLSX.mergeCells(sh, "A:D")           # Merge columns to the extent of the dimension.

```
"""
function mergeCells end
# The address methods mergeCells and removeMergedCells share.
for f in (:mergeCells, :removeMergedCells)
    @eval begin
        $f(ws::Worksheet, rng::SheetCellRange) = do_sheet_names_match(ws, rng) && $f(ws, rng.rng)
        $f(ws::Worksheet, rng::SheetColumnRange) = do_sheet_names_match(ws, rng) && $f(ws, rng.colrng)
        $f(ws::Worksheet, rng::SheetRowRange) = do_sheet_names_match(ws, rng) && $f(ws, rng.rowrng)
        $f(ws::Worksheet, colrng::ColumnRange)::Int = process_columnranges($f, ws, colrng)
        $f(ws::Worksheet, rowrng::RowRange)::Int = process_rowranges($f, ws, rowrng)
        $f(ws::Worksheet, ncrng::NonContiguousRange; kw...)::Int = process_ncranges($f, ws, ncrng; kw...)
        $f(xl::XLSXFile, sheetcell::AbstractString)::Int = process_sheetcell($f, xl, sheetcell)
        $f(ws::Worksheet, ref_or_rng::AbstractString)::Int = process_ranges($f, ws, ref_or_rng)
        $f(ws::Worksheet, row::Union{Integer,UnitRange{<:Integer}}, ::Colon) = process_colon($f, ws, row, nothing)
        $f(ws::Worksheet, ::Colon, col::Union{Integer,UnitRange{<:Integer}}) = process_colon($f, ws, nothing, col)
        $f(ws::Worksheet, ::Colon, ::Colon) = process_colon($f, ws, nothing, nothing)
        $f(ws::Worksheet, ::Colon) = process_colon($f, ws, nothing, nothing)
        $f(ws::Worksheet, row::Union{Integer,UnitRange{<:Integer}}, col::Union{Integer,UnitRange{<:Integer}}) = $f(ws, CellRange(CellRef(first(row), first(col)), CellRef(last(row), last(col))))
    end
end
function mergeCells(ws::Worksheet, cr::CellRange)
    # May be better if merged cells were part of ws.cache?

    !issubset(cr, get_dimension(ws)) && throw(XLSXError("Range `$cr` goes outside worksheet dimension."))

    length(cr) == 1 && throw(XLSXError("Cannot merge a single cell: `$cr`"))

    _require_writable(ws, "merge cells")

    _require_cache(ws, "get merged cells")

    for t in tables(ws) # Excel drops a table that overlaps a merged range
        if intersects(cr, t.ref)
            throw(XLSXError("Merged range (`$cr`) cannot overlap table `$(t.name)` (`$(t.ref)`)."))
        end
    end

    sheetdoc = xmlroot(get_workbook(ws), ws.relationship_id) # find the <mergeCells> block in the worksheet's xml file
    i, j = get_idces(sheetdoc, "worksheet", "mergeCells")

    pfx = get_prefix(ws)
    pfx = pfx == "" ? pfx : pfx * ":"
    if isnothing(j) # There are no existing merged cells. Insert a block where the schema puts it.
        root = sheetdoc[i]
        l = insert_index(root, "mergeCells", WORKSHEET_ORDER)
        insert!(root.children, l + 1, XML.Element("$(pfx)mergeCells"))
        j = l + 1
        count = 0
    else # There are already some existing merged cells
        c = xml_elements(sheetdoc[i][j])
        count = length(c)
        if count != parse(Int, sheetdoc[i][j]["count"])
            throw(XLSXError("Unexpected number of mergeCells found: $(length(c)). Expected $(sheetdoc[i][j]["count"])."))
        end
        for child in c
            if intersects(cr, CellRange(child["ref"]))
                throw(XLSXError("Merged range (`$cr`) cannot overlap with existing merged range (`" * child["ref"] * "`)."))
            end
        end
    end

    push!(sheetdoc[i][j], XML.Element("$(pfx)mergeCell", ref=string(cr))) # Add the new merged cell range.
    count += 1
    sheetdoc[i][j]["count"] = count

    # All cells except the base cell are set to missing.
    let first = true
        for cell in cr
            if first
                first = false
                continue
            else
                ws[cell] = ""
            end
        end
    end

    update_worksheets_xml!(get_xlsxfile(ws))

    return 0 # meaningless return value. Int required to comply with reference decoding structure.
end

"""
    removeMergedCells(ws::Worksheet, cr::String) -> 0
    removeMergedCells(xf::XLSXFile,  cr::String) -> 0

    removeMergedCells(ws::Worksheet, row::Int, col::Int) -> 0

    removeMergedCells(ws::Worksheet)             -> 0
    removeMergedCells(ws::Worksheet, ::Colon)    -> 0

Unmerge merged cells in a worksheet.

A merged range is unmerged if it fully encloses the range given by `cr`. Naming
any single cell of a merged range is therefore enough to unmerge the whole of
it, without needing to know its extent. A range larger than the merge encloses
nothing and is a no-op: `removeMergedCells(sh, "D47:J47")` leaves a merged
`D47:H47` intact.

Because merged ranges cannot overlap one another, a contiguous `cr` unmerges at
most one range. Use the non-contiguous form, or the no-argument form, to unmerge
several at once.

Called with no cell reference at all (or with a single `Colon`), every merged
range in the worksheet is unmerged.

Unmerging does not change any cell values, reflecting the behaviour of Excel
itself.

If nothing in the worksheet encloses `cr`, the worksheet is left untouched.
Specifying a range that is not merged is not an error it is a no-op.

A non-contiguous range composed of multiple cell ranges will be processed as a
list of separate ranges. Each range will be unmerged separately.

The Excel file must be opened in write mode to work with merged cells.

# Examples:
```julia
julia> XLSX.removeMergedCells(xf, "Sheet1!B2")     # Unmerge the range containing B2.

julia> XLSX.removeMergedCells(sh, "B2:C3")         # Unmerge a range enclosing B2:C3.

julia> XLSX.removeMergedCells(sh, 4, 4)            # Unmerge the range containing D4.

julia> XLSX.removeMergedCells(sh, :, 10)           # Unmerge a whole-column merge in J.

julia> XLSX.removeMergedCells(sh)                  # Unmerge everything in the worksheet.

```
"""
function removeMergedCells end

# Unlike `mergeCells`, a single cell is meaningful here: it identifies the merged
# range that contains it, so `process_ranges` can hand us a bare `CellRef`.
removeMergedCells(ws::Worksheet, cellref::CellRef)::Int = removeMergedCells(ws, CellRange(cellref, cellref))

removeMergedCells(ws::Worksheet)::Int = _remove_merged_cells(ws, nothing)
removeMergedCells(ws::Worksheet, cr::CellRange)::Int = _remove_merged_cells(ws, cr)

# `cr === nothing` means "every merged range in the worksheet".
function _remove_merged_cells(ws::Worksheet, cr::Union{CellRange,Nothing})::Int

    xf = get_xlsxfile(ws)

    _require_writable(xf, "remove merged cells")

    _require_cache(xf, "remove merged cells")

    if !isnothing(cr) && !issubset(cr, get_dimension(ws))
        throw(XLSXError("Range `$cr` goes outside worksheet dimension."))
    end

    sheetdoc = xmlroot(get_workbook(ws), ws.relationship_id) # find the <mergeCells> block in the worksheet's xml file
    i, j = get_idces(sheetdoc, "worksheet", "mergeCells")

    isnothing(j) && return 0 # There are no merged cells, so nothing to remove.

    block = sheetdoc[i][j]
    kids = XML.children(block)

    # Indices of the <mergeCell> children to drop, in ascending order (as `deleteat!` requires).
    drop = Int[]
    nrefs = 0
    for (k, child) in enumerate(kids)
        XML.nodetype(child) == XML.Element || continue
        nrefs += 1
        !haskey(child, "ref") && throw(XLSXError("No `ref` attribute found in `mergeCell` element."))
        if isnothing(cr) || issubset(cr, CellRange(child["ref"]))
            push!(drop, k)
        end
    end

    if nrefs != parse(Int, block["count"])
        throw(XLSXError("Unexpected number of mergeCells found: $nrefs. Expected $(block["count"])."))
    end

    isempty(drop) && return 0 # Nothing matched: leave the xml untouched.

    if length(drop) == nrefs
        # An empty <mergeCells count="0"/> is not schema-valid, so drop the whole block.
        deleteat!(sheetdoc[i].children, j)
    else
        deleteat!(kids, drop)
        block["count"] = nrefs - length(drop)
    end

    update_worksheets_xml!(xf)

    return 0 # meaningless return value. Int required to comply with reference decoding structure.
end
