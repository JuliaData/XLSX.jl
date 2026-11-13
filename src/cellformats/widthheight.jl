#
# -- Get and set column width
#

"""
    setColumnWidth(sh::Worksheet, cr::String; kw...) -> ::Int
    setColumnWidth(xf::XLSXFile,  cr::String, kw...) -> ::Int

    setColumnWidth(sh::Worksheet, row, col; kw...) -> ::Int
    setColumnWidth(sh::Worksheet, col; kw...) -> ::Int

Set the width of a column or column range.

A standard cell reference or cell range can be used to define the column range. 
The function will use the columns and ignore the rows. Named cells and named
ranges can similarly be used.
Alternatively, specify the row and column using any combination of 
Integer, UnitRange, Vector{Integer} or `:`, but only the columns will be used.
The row can be omitted and column specified alone, using an integer or colon.


The function uses one keyword used to define a column width:
- `width::Real = nothing` : Defines width in Excel's own (internal) units

When you set a column widths interactively in Excel you can see the width 
in "internal" units and in pixels. The width stored in the xlsx file is slightly 
larger than the width shown intertactively because Excel adds some cell padding. 
The method Excel uses to calculate the padding is obscure and complex. This 
function does not attempt to replicate it, but simply adds 0.71 internal units 
to the value specified. The value set is unlikely to match the value seen 
interactivley in the resultant spreadsheet, but will be close.

You can set a column width to 0.

The function returns a value of 0. The return value is meaningless but required to comply with the
range dispatch infrastructure, which expects an `Int` return from all `set*` functions.

`setColumnWidth` requires a file to be open for writing as well as reading (`mode="rw"` or open as a template) but 
it can work on empty cells.

# Examples:
```julia
julia> XLSX.setColumnWidth(xf, "Sheet1!A2"; width = 50)

julia> XLSX.setColumnWidth(sh, "F1:F5"; width = 0)

julia> XLSX.setColumnWidth(sh, "I"; width = 24.37)
 
```
"""
function setColumnWidth end
# The address methods setColumnWidth and setRowHeight share. Each also takes a
# bare column (or row) number, range or vector, below.
for f in (:setColumnWidth, :setRowHeight)
    @eval begin
        $f(ws::Worksheet, ref::SheetCellRef; kw...) = do_sheet_names_match(ws, ref) && $f(ws, ref.cellref; kw...)
        $f(ws::Worksheet, rng::SheetCellRange; kw...) = do_sheet_names_match(ws, rng) && $f(ws, rng.rng; kw...)
        $f(ws::Worksheet, rng::SheetColumnRange; kw...) = do_sheet_names_match(ws, rng) && $f(ws, rng.colrng; kw...)
        $f(ws::Worksheet, rng::SheetRowRange; kw...) = do_sheet_names_match(ws, rng) && $f(ws, rng.rowrng; kw...)
        $f(ws::Worksheet, colrng::ColumnRange; kw...)::Int = process_columnranges($f, ws, colrng; kw...)
        $f(ws::Worksheet, rowrng::RowRange; kw...)::Int = process_rowranges($f, ws, rowrng; kw...)
        $f(ws::Worksheet, ncrng::NonContiguousRange; kw...)::Int = process_ncranges($f, ws, ncrng; kw...)
        $f(ws::Worksheet, ref_or_rng::AbstractString; kw...)::Int = process_ranges($f, ws, ref_or_rng; kw...)
        $f(xl::XLSXFile, sheetcell::String; kw...)::Int = process_sheetcell($f, xl, sheetcell; kw...)
        $f(ws::Worksheet, row::Union{Integer,UnitRange{<:Integer}}, ::Colon; kw...) = process_colon($f, ws, row, nothing; kw...)
        $f(ws::Worksheet, ::Colon, col::Union{Integer,UnitRange{<:Integer}}; kw...) = process_colon($f, ws, nothing, col; kw...)
        $f(ws::Worksheet, ::Colon, ::Colon; kw...) = process_colon($f, ws, nothing, nothing; kw...)
        $f(ws::Worksheet, ::Colon; kw...) = process_colon($f, ws, nothing, nothing; kw...)
        $f(ws::Worksheet, row::Union{Vector{Int},StepRange{<:Integer}}, ::Colon; kw...) = process_veccolon($f, ws, row, nothing; kw...)
        $f(ws::Worksheet, ::Colon, col::Union{Vector{Int},StepRange{<:Integer}}; kw...) = process_veccolon($f, ws, nothing, col; kw...)
        $f(ws::Worksheet, row::Union{Integer,UnitRange{<:Integer}}, col::Union{Vector{Int},StepRange{<:Integer}}; kw...) = process_vecint($f, ws, row, col; kw...)
        $f(ws::Worksheet, row::Union{Vector{Int},StepRange{<:Integer}}, col::Union{Integer,UnitRange{<:Integer}}; kw...) = process_vecint($f, ws, row, col; kw...)
        $f(ws::Worksheet, row::Union{Vector{Int},StepRange{<:Integer}}, col::Union{Vector{Int},StepRange{<:Integer}}; kw...) = process_vecint($f, ws, row, col; kw...)
        $f(ws::Worksheet, row::Union{Integer,UnitRange{<:Integer}}, col::Union{Integer,UnitRange{<:Integer}}; kw...) = $f(ws, CellRange(CellRef(first(row), first(col)), CellRef(last(row), last(col))); kw...)
    end
end
setColumnWidth(ws::Worksheet, cr::CellRef; kw...)::Int = setColumnWidth(ws, CellRange(cr, cr); kw...)
setColumnWidth(ws::Worksheet, col::Union{Integer,UnitRange{<:Integer}}; kw...) = process_colon(setColumnWidth, ws, nothing, col; kw...)  # convenience: bare column index/range
setColumnWidth(ws::Worksheet, col::Union{Vector{Int},StepRange{<:Integer}}; kw...) = process_veccolon(setColumnWidth, ws, nothing, col; kw...)  # convenience: bare column vector/steprange
function setColumnWidth(ws::Worksheet, rng::CellRange; width::Union{Nothing,Real}=nothing)::Int
    _require_writable(ws, "set column widths")

    # Validate and resolve width before any XML work
    isnothing(width) && return 0
    width < 0 && throw(XLSXError("Invalid width: $width. Must be >= 0."))

    # The columns keep their other settings (style, hidden, ...); see `_update_col_defs!`.
    width_atts = Dict("width" => string(width + EXCEL_COLUMN_WIDTH_PADDING), "customWidth" => "1")
    _update_col_defs!(d -> merge!(d, width_atts), ws, rng.start.column_number, rng.stop.column_number)

    return 0
end

"""
    getColumnWidth(sh::Worksheet, cr::String) -> ::Union{Nothing, Real}
    getColumnWidth(xf::XLSXFile,  cr::String) -> ::Union{Nothing, Real}

    getColumnWidth(sh::Worksheet,  row::Int, col::Int) -> ::Union{Nothing, Real}

Get the width of a column defined by a cell reference or named cell.
The specified cell must be within the sheet dimension.

A standard cell reference or defined name may be used to define the column. 
The function will use the column number and ignore the row.

The function returns the value of the column width or nothing if the column 
does not have an explicitly defined width.

# Examples:
```julia
julia> XLSX.getColumnWidth(xf, "Sheet1!A2")

julia> XLSX.getColumnWidth(sh, "F1")

julia> XLSX.getColumnWidth(sh, 1, 6)
 
```
"""
function getColumnWidth end
getColumnWidth(xl::XLSXFile, sheetcell::String)::Union{Nothing,Float64} = process_get_sheetcell(getColumnWidth, xl, sheetcell)
getColumnWidth(ws::Worksheet, cr::String) = process_get_cellname(getColumnWidth, ws, cr)
getColumnWidth(ws::Worksheet, row::Integer, col::Integer) = getColumnWidth(ws, CellRef(row, col))
function getColumnWidth(ws::Worksheet, cellref::CellRef)::Union{Nothing,Real}
    # May be better if column width were part of ws.cache?

    _require_writable(ws, "get column width")

    d = get_dimension(ws)
    if cellref ∉ d
        throw(XLSXError("Cell specified is outside sheet dimension `$d`"))
    end

    sheetdoc = xmlroot(get_workbook(ws), ws.relationship_id)
    
    i, j = get_idces(sheetdoc, "worksheet", "cols")

    if isnothing(j) # There are no existing column formats defined.
        return nothing
    end
    col = cellref.column_number
    for c in xml_elements(sheetdoc[i][j])
        if parse(Int, c["min"]) <= col <= parse(Int, get(c, "max", c["min"]))   # a definition may span columns
            if haskey(c, "width")
                return parse(Float64, c["width"])
            else
                break
            end
        end
    end

    # Either the column definition was found but with no width attribute or not found at all.
    return nothing
end

#
# -- Get and set row height
#

"""
    setRowHeight(sh::Worksheet, cr::String; kw...) -> ::Int
    setRowHeight(xf::XLSXFile,  cr::String, kw...) -> ::Int

    setRowHeight(sh::Worksheet, row, col; kw...) -> ::Int

Set the height of a row or row range.

A standard cell reference or cell range must be used to define the row range. 
The function will use the rows and ignore the columns. Named cells and named
ranges can similarly be used.
Alternatively, specify the row and column using any combination of 
Integer, UnitRange, Vector{Integer} or `:`, but only the rows will be used.

The function uses one keyword used to define a row height:
- `height::Real = nothing` : Defines height in Excel's own (internal) units.

When you set row heights interactively in Excel you can see the height 
in "internal" units and in pixels. The height stored in the xlsx file is slightly 
larger than the height shown interactively because Excel adds some cell padding. 
The method Excel uses to calculate the padding is obscure and complex. This 
function does not attempt to replicate it, but simply adds 0.21 internal units 
to the value specified. The value set is unlikely to match the value seen 
interactivley in the resultant spreadsheet, but it will be close.

Rows needn't hold any data: an empty row is written with just its height.

You can set a row height to 0.

The function returns a value of 0. The return value is meaningless but required to comply with the
range dispatch infrastructure, which expects an `Int` return from all `set*` functions.

# Examples:
```julia
julia> XLSX.setRowHeight(xf, "Sheet1!A2"; height = 50)

julia> XLSX.setRowHeight(sh, "F1:F5"; height = 0)

julia> XLSX.setRowHeight(sh, "I"; height = 24.56)

```
"""
function setRowHeight end
setRowHeight(ws::Worksheet, cr::CellRef; kw...)::Int = setRowHeight(ws::Worksheet, CellRange(cr, cr); kw...)
setRowHeight(ws::Worksheet, row::Union{Integer,UnitRange{<:Integer}}; kw...) = process_colon(setRowHeight, ws, row, nothing; kw...)
setRowHeight(ws::Worksheet, row::Union{Vector{Int},StepRange{<:Integer}}; kw...) = process_veccolon(setRowHeight, ws, row, nothing; kw...)
function setRowHeight(ws::Worksheet, rng::CellRange; height::Union{Nothing,Real}=nothing)::Int

    _require_writable(ws, "set row heights")

    top = rng.start.row_number
    bottom = rng.stop.row_number
    padded_height = isnothing(height) ? -1 : height + 0.2109375 # Excel adds cell padding to a user specified width
    if !isnothing(height) && height < 0
        throw(XLSXError("Invalid value specified for height: $height. Height must be >= 0."))
    end

    if isnothing(height) # No-op
        return 0
    end

    _row_cache!(ws, "set row heights")
    for r in top:bottom
        _ensure_cache_row!(ws, r)   # an empty row is written with just its height
        ws.cache.row_ht[r] = padded_height
        _set_row_attr!(ws, Int(r), "customHeight", "1")   # set, so Excel doesn't fit it to the text
    end

    return 0

end

"""
    getRowHeight(sh::Worksheet, cr::String) -> ::Union{Nothing, Real}
    getRowHeight(xf::XLSXFile,  cr::String) -> ::Union{Nothing, Real}

    getRowHeight(sh::Worksheet,  row::Int, col::Int) -> ::Union{Nothing, Real}

Get the height of a row defined by a cell reference or named cell.
The specified cell must be within the sheet dimension.

A standard cell reference or defined name must be used to define the row. 
The function will use the row number and ignore the column.

The function returns the value of the row height or nothing if the row 
does not have an explicitly defined height.

If the row is not found (an empty row), returns -1.

# Examples:
```julia
julia> XLSX.getRowHeight(xf, "Sheet1!A2")

julia> XLSX.getRowHeight(sh, "F1")

julia> XLSX.getRowHeight(sh, 1, 6)
 
```
"""
function getRowHeight end
getRowHeight(xl::XLSXFile, sheetcell::String)::Union{Nothing,Real} = process_get_sheetcell(getRowHeight, xl, sheetcell)
getRowHeight(ws::Worksheet, cr::String) = process_get_cellname(getRowHeight, ws, cr)
getRowHeight(ws::Worksheet, row::Integer, col::Integer) = getRowHeight(ws, CellRef(row, col))
function getRowHeight(ws::Worksheet, cellref::CellRef)::Union{Nothing,Real}

    _require_writable(ws, "get row height")

    d = get_dimension(ws)
    if cellref ∉ d
        throw(XLSXError("Cell specified is outside sheet dimension `$d`"))
    end

    if isnothing(ws.cache) || !haskey(ws.cache.row_ht, cellref.row_number)
        _ = ws[cellref] # Ensure cache filled at least to include the last row in `rng`.
    end

    if haskey(ws.cache.row_ht, cellref.row_number) # Row might still be empty
        return ws.cache.row_ht[cellref.row_number]
    end

    return -1 # Row specified not found (is empty)

end
