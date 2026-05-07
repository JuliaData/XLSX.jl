# Order of worksheet elements:

const WORKSHEET_ORDER = String[
    "sheetPr",
    "dimension",
    "sheetViews",
    "sheetFormatPr",
    "cols",
    "sheetData",
    "sheetCalcPr",
    "sheetProtection",
    "protectedRanges",
    "scenarios",
    "autoFilter",
    "sortState",
    "dataConsolidate",
    "customSheetViews",
    "mergeCells",
    "phoneticPr",
    "conditionalFormatting",
    "dataValidations",
    "hyperlinks",
    "printOptions",
    "pageMargins",
    "pageSetup",
    "headerFooter",
    "rowBreaks",
    "colBreaks",
    "customProperties",
    "cellWatches",
    "ignoredErrors",
    "smartTags",
    "drawing",
    "legacyDrawing",
    "legacyDrawingHF",
    "picture",
    "oleObjects",
    "controls",
    "webPublishItems",
    "tableParts",
    "extLst"
]

function insert_index(root::XML.Node, target::String, order::Vector{String})
    # index of target in canonical order
    target_idx = findfirst(==(target), order)
    isnothing(target_idx) && throw(XLSXError("Target $target not in worksheet order list"))

    chn = XML.children(root)

    # scan backwards through the order list
    for i in target_idx:-1:1
        name = order[i]
        # find the last child with this tag
        for (j, child) in enumerate(chn)
            if localname(child) == name
                return j   # insert *after* this child
            end
        end
    end

    return 1   # nothing precedes the target → insert at top
end

function Worksheet(xf::XLSXFile, sheet_element::XML.Node)
    wb = get_workbook(xf)
    localname(sheet_element) !=   "sheet" && throw(XLSXError("Something wrong here!"))
    a = XML.attributes(sheet_element)
    sheetId = parse(Int, a["sheetId"])
    relationship_id = a["r:id"]
    name = XLSX.unescape(a["name"])
    is_hidden = haskey(a, "state") && a["state"] in ["hidden", "veryHidden"]

    return Worksheet(xf, sheetId, relationship_id, name, nothing, is_hidden)
end

function ordinal_sheet_number(wb::Workbook, name::String)
    for (i, s) in enumerate(wb.sheets)
        if s.name == name
            return i
        end
    end
    throw(XLSXError("worksheet $(ws.name) not found in workbook"))
end
function Base.axes(ws::Worksheet, d)
    dim = get_dimension(ws)
    if d == 1
        return dim.start.row_number:dim.stop.row_number
    elseif d == 2
        return dim.start.column_number:dim.stop.column_number
    else
        throw(ArgumentError("Unsupported dimension $d"))
    end
end

# 18.3.1.35 - dimension (Worksheet Dimensions). This is optional, and not required.
function read_worksheet_dimension(xf::XLSXFile, relationship_id, name)::Union{Nothing,CellRange}

    wb = get_workbook(xf)
    if hassheet(wb, name) # use worksheet cache if possible
        let ws = first(wb.sheets)
            for s in wb.sheets
                if s.name == unquoteit(name)
                    ws=s
                end
            end
            if !isnothing(ws.cache) && !isempty(ws.cache) && ws.cache.is_full
                return get_dimension(ws::Worksheet)
            end
        end
    end

    local result::Union{Nothing,CellRange} = nothing
    target_file = get_relationship_target_by_id("xl", wb, relationship_id)
    doc = open_internal_file_stream(xf, target_file)
    reader = iterate(doc)
    # Now let's look for a row element, if it exists
    while reader !== nothing # go next node
        (sheet_row, state) = reader
        if XML.nodetype(sheet_row) == XML.Element && localname(sheet_row) == "dimension"

            XML.depth(sheet_row) != 2 && throw(XLSXError("Malformed Worksheet \"$name\": unexpected node depth for dimension node: $(XML.depth(sheet_row))."))

            ref_str = XML.attributes(sheet_row)["ref"]
            if is_valid_cellname(ref_str)
                result = CellRange("$(ref_str):$(ref_str)")
            else
                result = CellRange(ref_str)
            end

            break
        end
        reader = iterate(doc, state)
    end

    return result
end

@inline isdate1904(ws::Worksheet) = isdate1904(get_workbook(ws))

# Returns the dimension of this worksheet as a CellRange.
# If the dimension is unknown, computes a dimension from cells in cache.
# If the cache is empty or is not being used, set dimension to A1:A1.
function get_dimension(ws::Worksheet)::Union{Nothing,CellRange}
    !isnothing(ws.dimension) && return ws.dimension
    if isnothing(ws.cache) || isempty(ws.cache) || !ws.cache.is_full
        set_dimension!(ws, CellRange(CellRef(1, 1), CellRef(1, 1)))
    else
        row_extr = extrema(keys(ws.cache.cells))
        row_min = first(row_extr)
        row_max = last(row_extr)
        col_extr = [extrema(y) for y in [keys(x) for x in values(ws.cache.cells)] if !isempty(y)]
        col_min = minimum([x for x in first.(col_extr)])
        col_max = maximum([x for x in last.(col_extr)])
        set_dimension!(ws, CellRange(CellRef(row_min, col_min), CellRef(row_max, col_max)))
    end
    return ws.dimension
end

function set_dimension!(ws::Worksheet, rng::CellRange)
    ws.dimension = rng
    nothing
end

function ref_chooser(f::Function, ws::Worksheet, ref::AbstractString)
    if is_worksheet_defined_name(ws, ref)
        v = get_defined_name_value(ws, ref)
        if is_defined_name_value_a_constant(v)
            return v
        elseif is_defined_name_value_a_reference(v)
            return f(ws, v)
        else
            throw(XLSXError("`$ref` is not a valid cell or range reference."))
        end
    elseif is_workbook_defined_name(get_workbook(ws), ref)
        wb = get_workbook(ws)
        v = get_defined_name_value(wb, ref)
        if is_defined_name_value_a_constant(v)
            return v
        elseif is_defined_name_value_a_reference(v)
            return f(get_xlsxfile(ws), v)
        else
            throw(XLSXError("`$ref` is not a valid cell or range reference."))
        end
    elseif is_valid_cellname(ref)
        return f(ws, CellRef(ref))
    elseif is_valid_sheet_cellname(ref)
        return f(ws, SheetCellRef(ref))
    elseif is_valid_cellrange(ref)
        return f(ws, CellRange(ref))
    elseif is_valid_column_range(ref)
        return f(ws, ColumnRange(ref))
    elseif is_valid_row_range(ref)
        return f(ws, RowRange(ref))
    elseif is_valid_non_contiguous_range(ref)
        return f(ws, NonContiguousRange(ws, ref))
    elseif is_valid_sheet_cellrange(ref)
        return f(ws, SheetCellRange(ref))
    elseif is_valid_sheet_column_range(ref)
        return f(ws, SheetColumnRange(ref))
    elseif is_valid_sheet_row_range(ref)
        return f(ws, SheetRowRange(ref))
    elseif is_valid_non_contiguous_range(ref)
        return f(ws, NonContiguousRange(ws, ref))
    end
    throw(XLSXError("`$ref` is not a valid cell or range reference."))
end

"""
    getdata(sheet, ref)
    getdata(sheet, row, column)

Returns a scalar, matrix or a vector of matrices with values from 
a spreadsheet.

`ref` can be a cell reference or a range or a valid defined name.

If `ref` is a single cell, a scalar is returned.

Most ranges are rectangular and will return a 2-D matrix 
(`Array{AbstractCell, 2}`). For row and column ranges, the 
extent of the range in the other dimension is determined by 
the worksheet's dimension.

A non-contiguous range (which may not be rectangular) will return 
a vector of `Array{AbstractCell, 2}` matrices with one element for 
each non-contiguous (comma separated) element in the range.

Indexing in a `Worksheet` will dispatch to `getdata` method.

# Example

```julia
julia> f = XLSX.readxlsx("myfile.xlsx")

julia> sheet = f["mysheet"] # Worksheet

julia> matrix = sheet["A1:B4"] # CellRange

julia> matrix = sheet["A:B"] # Column range

julia> matrix = sheet["1:4"] # Row range

julia> matrix = sheet["Contiguous"] # Named range

julia> matrix = sheet[1:30, 1] # use unit ranges to define rows and/or columns

julia> matrix = sheet[[1, 2, 3], 1] # vectors of integers to define rows and/or columns

julia> vector = sheet["A1:A4,C1:C4,G5"] # Non-contiguous range

julia> vector = sheet["Location"] # Non-contiguous named range

julia> scalar = sheet[2, 2] # Cell "B2"

```

See also [`XLSX.readdata`](@ref).
"""
function getdata(ws::Worksheet, ref::AbstractString)
    return ref_chooser(getdata, ws, ref)
end
getdata(ws::Worksheet, single::CellRef) = getdata(ws, getcell(ws, single))
getdata(ws::Worksheet, row::Integer, col::Integer) = getdata(ws, CellRef(row, col))
getdata(ws::Worksheet, row::Union{Integer,UnitRange{<:Integer}}, col::Union{Vector{Int},StepRange{<:Integer}}) = [getdata(ws, a, b) for a in row, b in col]
getdata(ws::Worksheet, row::Union{Vector{Int},StepRange{<:Integer}}, col::Union{Integer,UnitRange{<:Integer}}) = [getdata(ws, a, b) for a in row, b in col]
getdata(ws::Worksheet, row::Union{Vector{Int},StepRange{<:Integer}}, col::Union{Vector{Int},StepRange{<:Integer}}) = [getdata(ws, a, b) for a in row, b in col]
getdata(ws::Worksheet, row::Union{Integer,UnitRange{<:Integer}}, col::Union{Integer,UnitRange{<:Integer}}) = getdata(ws, CellRange(CellRef(first(row), first(col)), CellRef(last(row), last(col))))
getdata(ws::Worksheet) = getdata(ws, get_dimension(ws))
getdata(ws::Worksheet, ::Colon, ::Colon) = getdata(ws)
function getdata(ws::Worksheet, ::Colon)
    dim = get_dimension(ws)
    getdata(ws, dim)
end
function getdata(ws::Worksheet, row::Union{Integer,UnitRange{<:Integer}}, ::Colon)
    dim = get_dimension(ws)
    getdata(ws, CellRange(CellRef(first(row), dim.start.column_number), CellRef(last(row), dim.stop.column_number)))
end
function getdata(ws::Worksheet, ::Colon, col::Union{Integer,UnitRange{<:Integer}})
    dim = get_dimension(ws)
    getdata(ws, CellRange(CellRef(dim.start.row_number, first(col)), CellRef(dim.stop.row_number, last(col))))
end
function getdata(ws::Worksheet, row::Union{Vector{Int},StepRange{<:Integer}}, ::Colon)
    dim = get_dimension(ws)
    col = dim.start.column_number:dim.stop.column_number
    return getdata(ws, row, col)
end
function getdata(ws::Worksheet, ::Colon, col::Union{Vector{Int},StepRange{<:Integer}})
    dim = get_dimension(ws)
    row = dim.start.row_number:dim.stop.row_number
    return getdata(ws, row, col)
end
function getdata(ws::Worksheet, rng::CellRange)::Array{Any,2}
    result = Array{Any,2}(undef, size(rng))
    fill!(result, missing)

    top = row_number(rng.start)
    bottom = row_number(rng.stop)
    left = column_number(rng.start)
    right = column_number(rng.stop)

    for sheetrow in eachrow(ws)
        if top <= sheetrow.row && sheetrow.row <= bottom
            for column in left:right
                cell = getcell(sheetrow, column)
                if !isempty(cell)
                    (r, c) = relative_cell_position(cell, rng)
                    result[r, c] = getdata(ws, cell)
                end
            end
        end

        # don't need to read any more rows
        if sheetrow.row > bottom
            break
        end
    end

    return result
end
function getdata(ws::Worksheet, rng::ColumnRange)::Array{Any,2}
    dim = get_dimension(ws)
    start = CellRef(dim.start.row_number, rng.start)
    stop = CellRef(dim.stop.row_number, rng.stop)
    return getdata(ws, CellRange(start, stop))
end
function getdata(ws::Worksheet, rng::RowRange)::Array{Any,2}
    dim = get_dimension(ws)
    start = CellRef(rng.start, dim.start.column_number,)
    stop = CellRef(rng.stop, dim.stop.column_number)
    return getdata(ws, CellRange(start, stop))
end
function getdata(ws::Worksheet, rng::NonContiguousRange)::Vector{Array{Any,2}}
    do_sheet_names_match(ws, rng)
    results = Vector{Array{Any,2}}()
    for r in rng.rng
        if r isa CellRef
            push!(results, getdata(ws, CellRange(r, r)))
        else
            push!(results, getdata(ws, r))
        end
    end
    return results
end

# Needed for definedName references
getdata(ws::Worksheet, s::SheetCellRef) = do_sheet_names_match(ws, s) && getdata(ws, s.cellref)
getdata(ws::Worksheet, s::SheetCellRange) = do_sheet_names_match(ws, s) && getdata(ws, s.rng)
getdata(ws::Worksheet, s::SheetColumnRange) = do_sheet_names_match(ws, s) && getdata(ws, s.colrng)
getdata(ws::Worksheet, s::SheetRowRange) = do_sheet_names_match(ws, s) && getdata(ws, s.rowrng)


Base.getindex(ws::Worksheet, r) = getdata(ws, r)
Base.getindex(ws::Worksheet, r, c) = getdata(ws, r, c)
Base.getindex(ws::Worksheet, ::Colon) = getdata(ws)

function Base.show(io::IO, ws::Worksheet)
    hidden_string = ws.is_hidden ? "(hidden)" : ""
    if is_chartsheet(get_workbook(ws), ws.name)
        @printf(io, "Chartsheet: [\"%s\"] %s", ws.name, hidden_string)
        return
    end
    rg = get_dimension(ws)
    if rg !== nothing
        nrow, ncol = size(rg)
        @printf(io, "%d×%d %s: [\"%s\"](%s) %s", nrow, ncol, typeof(ws), ws.name, rg, hidden_string)
    else
        @printf(io, "%s: [\"%s\"] %s", typeof(ws), ws.name, hidden_string)
    end
end

"""
    getcell(sheet, ref)
    getcell(sheet, row, col)

Return an `AbstractCell` that represents a cell in the spreadsheet.
Return a 2-D matrix as `Array{AbstractCell, 2}` if `ref` is a 
rectangular range.
For row and column ranges, the extent of the range in the other 
dimension is determined by the worksheet's dimension.
A non-contiguous range (which may not be rectangular) will return 
a vector of `Array{AbstractCell, 2}` with one element for each 
non-contiguous (comma separated) element in the range.

If `ref` is a range, `getcell` dispatches to [`getcellrange`](@ref).

Example:

```julia
julia> xf = XLSX.readxlsx("myfile.xlsx")

julia> sheet = xf["mysheet"]

julia> cell = XLSX.getcell(sheet, "A1")

julia> cell = XLSX.getcell(sheet, 1:3, [2,4,6])

```

Other examples are as [`getdata()`](@ref).

"""
function getcell(ws::Worksheet, single::CellRef)::AbstractCell

    # if cache is in use, look-up cell direct rather than iterating
    if !isnothing(ws.cache) && is_cache_enabled(ws)
        if haskey(ws.cache.cells, single.row_number)
            if haskey(ws.cache.cells[single.row_number], single.column_number)
                return ws.cache.cells[single.row_number][single.column_number]
            end
        end
        ws.cache.is_full && return EmptyCell(single)
    end

    # If can't use cache then iterate sheetrows

    if get_xlsxfile(ws).use_cache_for_sheet_data # fill cache if active
        for sheetrow in eachrow(ws)
            if row_number(sheetrow) == row_number(single)
                return getcell(sheetrow, column_number(single))
            end
        end
    
    else
        sheetrow=match_rows(ws, [row_number(single)])
        if length(sheetrow)==1
            return getcell(sheetrow[1], column_number(single))
        end
    end
        return EmptyCell(single)
end

getcell(ws::Worksheet, s::SheetCellRef) = do_sheet_names_match(ws, s) && getcell(ws, s.cellref)
getcell(ws::Worksheet, s::SheetCellRange) = do_sheet_names_match(ws, s) && getcellrange(ws, s.rng)
getcell(ws::Worksheet, s::SheetColumnRange) = do_sheet_names_match(ws, s) && getcellrange(ws, s.colrng)
getcell(ws::Worksheet, s::SheetRowRange) = do_sheet_names_match(ws, s) && getcellrange(ws, s.rowrng)
getcell(ws::Worksheet, s::CellRange) = getcellrange(ws, s)
getcell(ws::Worksheet, s::ColumnRange) = getcellrange(ws, s)
getcell(ws::Worksheet, s::RowRange) = getcellrange(ws, s)
getcell(ws::Worksheet, s::NonContiguousRange) = getcellrange(ws, s)

getcell(ws::Worksheet, row::Integer, col::Integer) = getcell(ws, CellRef(row, col))
getcell(ws::Worksheet, row::Union{Integer,UnitRange{<:Integer}}, col::Union{Vector{Int},StepRange{<:Integer}}) = getcellrange(ws, row, col)
getcell(ws::Worksheet, row::Union{Vector{Int},StepRange{<:Integer}}, col::Union{Integer,UnitRange{<:Integer}}) = getcellrange(ws, row, col)
getcell(ws::Worksheet, row::Union{Vector{Int},StepRange{<:Integer}}, col::Union{Vector{Int},StepRange{<:Integer}}) = getcellrange(ws, row, col)
getcell(ws::Worksheet, row::Union{Integer,UnitRange{<:Integer}}, col::Union{Integer,UnitRange{<:Integer}}) = getcellrange(ws, CellRange(CellRef(first(row), first(col)), CellRef(last(row), last(col))))
function getcell(ws::Worksheet, row::Union{Integer,UnitRange{<:Integer}}, ::Colon)
    dim = get_dimension(ws)
    getcellrange(ws, CellRange(CellRef(first(row), dim.start.column_number), CellRef(last(row), dim.stop.column_number)))
end
function getcell(ws::Worksheet, ::Colon, col::Union{Integer,UnitRange{<:Integer}})
    dim = get_dimension(ws)
    getcellrange(ws, CellRange(CellRef(dim.start.row_number, first(col)), CellRef(dim.stop.row_number, last(col))))
end
function getcell(ws::Worksheet, ::Colon)
    getcellrange(ws, get_dimension(ws))
end

function getcell(ws::Worksheet, ref::AbstractString)
    return ref_chooser(getcell, ws, ref)
end

"""
    getcellrange(sheet, rng)

Return a matrix with cells as `Array{AbstractCell, 2}`.
`rng` must be a valid cell range, column range or row range,
as in `"A1:B2"`, `"A:B"` or `"1:2"`, or a non-contiguous range.
For row and column ranges, the extent of the range in the other 
dimension is determined by the worksheet's dimension.
A non-contiguous range (which may not be rectangular) will return 
a vector of `Array{AbstractCell, 2}` with one element for each 
non-contiguous (comma separated) element in the range.

Example:

```julia
julia> ncr = "B3,A1,C2" # non-contiguous range, "out of order".
"B3,A1,C2"

julia>  XLSX.getcellrange(f[1], ncr)
3-element Vector{Matrix{XLSX.AbstractCell}}:
 [XLSX.Cell(B3, 0x0000000000000018, 0x00000000, 0x0000, XLSX.CT_INT, false);;]
 [XLSX.Cell(A1, 0x0000000000000018, 0x00000000, 0x0000, XLSX.CT_INT, false);;]
 [XLSX.Cell(C2, 0x0000000000000018, 0x00000000, 0x0000, XLSX.CT_INT, false);;]

```

For other examples, see [`getcell()`](@ref) and [`getdata()`](@ref).

"""
function getcellrange(ws::Worksheet, rng::CellRange)::Array{AbstractCell,2}
    result = Array{Any,2}(undef, size(rng))
    for cell in rng # initialise with empty cells
        (r, c) = relative_cell_position(cell, rng)
        result[r, c] = EmptyCell(cell)
    end

    top = row_number(rng.start)
    bottom = row_number(rng.stop)
    left = column_number(rng.start)
    right = column_number(rng.stop)

    if is_cache_enabled(ws)
        # use cache if possible
        if !isnothing(ws.cache)
            for single in rng
                if haskey(ws.cache.cells, single.row_number)
                    if haskey(ws.cache.cells[single.row_number], single.column_number)
                        cell = ws.cache.cells[single.row_number][single.column_number]
                        (r, c) = relative_cell_position(cell, rng)
                        result[r, c] = cell
                    end
                end
            end
        else
            # If cache empty then iterate sheetrows to fill
            for sheetrow in eachrow(ws)
                if top <= sheetrow.row && sheetrow.row <= bottom
                    for column in left:right
                        cell = getcell(sheetrow, column)
                        (r, c) = relative_cell_position(cell, rng)
                        result[r, c] = cell
                    end
                end
                # don't need to read any more rows
                if sheetrow.row > bottom
                    break
                end
            end
        end
    else
        # no cache to fill - just look in file
        sheetrows = match_rows(ws, collect(top:bottom))
        for sheetrow in sheetrows
            for column in left:right
                cell = getcell(sheetrow, column)
                (r, c) = relative_cell_position(cell, rng)
                result[r, c] = cell
            end
        end
    end

    return result
end

getcellrange(ws::Worksheet, s::SheetCellRange) = do_sheet_names_match(ws, s) && getcellrange(ws, s.rng)
getcellrange(ws::Worksheet, s::SheetColumnRange) = do_sheet_names_match(ws, s) && getcellrange(ws, s.colrng)
getcellrange(ws::Worksheet, s::SheetRowRange) = do_sheet_names_match(ws, s) && getcellrange(ws, s.rowrng)

getcellrange(ws::Worksheet, row::Union{Integer,UnitRange{<:Integer}}, col::Union{Vector{Int},StepRange{<:Integer}}) = [getcell(ws, a, b) for a in row, b in col]
getcellrange(ws::Worksheet, row::Union{Vector{Int},StepRange{<:Integer}}, col::Union{Integer,UnitRange{<:Integer}}) = [getcell(ws, a, b) for a in row, b in col]
getcellrange(ws::Worksheet, row::Union{Vector{Int},StepRange{<:Integer}}, col::Union{Vector{Int},StepRange{<:Integer}}) = [getcell(ws, a, b) for a in row, b in col]
getcellrange(ws::Worksheet, row::Union{Integer,UnitRange{<:Integer}}, col::Union{Integer,UnitRange{<:Integer}}) = getcell(ws, CellRange(CellRef(first(row), first(col)), CellRef(last(row), last(col))))
getcellrange(ws::Worksheet, row::Union{Integer,UnitRange{<:Integer}}, ::Colon) = getcell(ws, row, :)
getcellrange(ws::Worksheet, ::Colon, col::Union{Integer,UnitRange{<:Integer}}) = getcell(ws, :, col)
getcellrange(ws::Worksheet, ::Colon) = getcellrange(ws, get_dimension(ws))

function getcellrange(ws::Worksheet, rng::ColumnRange)::Array{AbstractCell,2}
    dim = get_dimension(ws)
    start = CellRef(dim.start.row_number, rng.start)
    stop = CellRef(dim.stop.row_number, rng.stop)
    return getcellrange(ws, CellRange(start, stop))
end
function getcellrange(ws::Worksheet, rng::RowRange)::Array{AbstractCell,2}
    dim = get_dimension(ws)
    start = CellRef(rng.start, dim.start.column_number,)
    stop = CellRef(rng.stop, dim.stop.column_number)
    return getcellrange(ws, CellRange(start, stop))
end

function getcellrange(ws::Worksheet, rng::NonContiguousRange)::Vector{Array{AbstractCell,2}}
    # returns a simple vector because non contiguous ranges aren't rectangular
    results = Vector{Array{AbstractCell,2}}()
    for r in rng.rng
        if r isa CellRef
            push!(results, getcellrange(ws, CellRange(r, r)))
        else
            push!(results, getcellrange(ws, r))
        end
    end
    return results
end

function getcellrange(ws::Worksheet, rng::AbstractString)
    return ref_chooser(getcellrange, ws, rng)
end
