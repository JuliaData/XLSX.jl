#
# Writing tables of data: writetable! into a sheet and writetable to a new file.
#

# Given an anchor cell at (anchor_row, anchor_col).
# Returns a CellRef at:
# - (anchor_row + offset, anchol_col) if dim = 1 (operates on rows)
# - (anchor_row, anchor_col + offset) if dim = 2 (operates on cols)
function target_cell_ref_from_offset(anchor_row::Integer, anchor_col::Integer, offset::Integer, dim::Integer)::CellRef
    if dim == 1
        return CellRef(anchor_row + offset, anchor_col)
    elseif dim == 2
        return CellRef(anchor_row, anchor_col + offset)
    else
        throw(XLSXError("Invalid dimension: $dim."))
    end
end

function target_cell_ref_from_offset(anchor_cell::CellRef, offset::Integer, dim::Integer)::CellRef
    return target_cell_ref_from_offset(row_number(anchor_cell), column_number(anchor_cell), offset, dim)
end

const ALLOWED_TYPES = Union{Number,String,Bool,Dates.Date,Dates.Time,Dates.DateTime,Missing,Nothing,CellImage}
function process_vector(col) # Convert any disallowed types to strings. #239.
    if eltype(col) <: ALLOWED_TYPES
        # Case 1: All elements are of allowed types
        return col
    elseif eltype(col) <: Any && all(x -> !(typeof(x) <: ALLOWED_TYPES), col)
        # Case 2: All elements are of disallowed types
        return map(x -> string(x), col)
    else
        # Case 3: Mixed types, process each element
        return [typeof(x) <: ALLOWED_TYPES ? x : string(x) for x in col]
    end
end

"""
    writetable!(
        sheet::Worksheet,
        data,
        columnnames;
        anchor_cell::CellRef=CellRef("A1"),
        write_columnnames::Bool=true,
        as_table::Bool=false,
        table_name::AbstractString="",
        table_style::Union{AbstractString,Nothing}=nothing,
        totals::Union{Nothing,AbstractVector{<:Pair}}=nothing,
        column_widths=nothing,
        column_formats=nothing,
        freeze_header::Bool=false,
    )

Write tabular data `data` with labels given by `columnnames` to `sheet`,
starting at `anchor_cell`.

`data` must be a vector of columns.
`columnnames` must be a vector of column labels.

Column labels that are not of type `String` will be converted 
to strings before writing. Any data columns that are not of 
type `String`, `Float64`, `Int64`, `Bool`, `Date`, `Time`, 
`DateTime`, `Missing`, or `Nothing` will be converted to strings 
before writing.

Set `as_table=true` to also turn the written range into an Excel Table
(equivalent to calling [`XLSX.addtable!`](@ref) immediately afterward over
exactly the range just written). Requires `write_columnnames=true` (a table
needs a header row) and at least one data row. 
`table_name` and `table_style` are forwarded to `addtable!`'s `name` 
and `style` keywords — see its docstring for accepted values, including 
the list of Excel's built-in table style names.

Use `totals` to also set a totals row on the table in the same call —
equivalent to calling [`XLSX.settotals!`](@ref) immediately afterward.
`totals` is a vector of `"ColumnName" => value` pairs, using the same
`value` forms `settotals!` accepts (a `Symbol` for a built-in function, a
`(:custom, formula)` tuple, or a `String` label). Requires `as_table=true`.

Three keywords lay out the columns written:
- `column_widths`: one width for every column, or a vector of widths by column,
  `nothing` leaving a column's width as it is. Widths are in characters, as
  [`XLSX.setColumnWidth`](@ref) takes them.
- `column_formats`: a `Dict` from column (number from the first written, or name in
  `columnnames`) to a number format, as [`XLSX.setFormat`](@ref) takes it, applied to
  the column's data cells.
- `freeze_header`: freeze the rows down to the header row, as
  [`XLSX.freezePanes`](@ref) does. Requires `write_columnnames=true`.

# Examples
```julia
julia> using DataFrames

julia> df = DataFrame(id=[1, 2], name=["alice", "bob"], score=[10.5, 20.0])

julia> XLSX.writetable!(sheet, collect(eachcol(df)), names(df);
           as_table=true, table_name="Results", table_style="TableStyleMedium2",
           totals=["id" => "Total", "score" => :sum])
```

See also: [`XLSX.writetable`](@ref), [`XLSX.addtable!`](@ref), [`XLSX.settotals!`](@ref).
"""
function writetable!(
    sheet::Worksheet,
    data,
    columnnames;
    anchor_cell::CellRef=CellRef("A1"),
    write_columnnames::Bool=true,
    as_table::Bool=false,
    table_name::AbstractString="",
    table_style::Union{AbstractString,Nothing}=nothing,
    totals::Union{Nothing,AbstractVector{<:Pair}}=nothing,
    column_widths=nothing,
    column_formats=nothing,
    freeze_header::Bool=false,
)
    freeze_header && !write_columnnames &&
        throw(XLSXError("`freeze_header=true` requires `write_columnnames=true`: there is no header row to freeze otherwise."))
    as_table && !write_columnnames &&
        throw(XLSXError("`as_table=true` requires `write_columnnames=true` — an Excel Table needs a header row, and `writetable!` is what writes it."))
    !isnothing(totals) && !as_table &&
        throw(XLSXError("`totals` requires `as_table=true` — there's no table to set a totals row on otherwise."))

    col_count = length(data)
    row_count = _check_table_shape(data, columnnames)
    widths, formats = _column_settings([string(c) for c in columnnames], column_widths, column_formats)

    as_table && row_count <= 0 &&
        throw(XLSXError("`as_table=true` requires at least one data row — Excel does not support header-only tables."))

    anchor_row = row_number(anchor_cell)
    anchor_col = column_number(anchor_cell)
    start_from_anchor = 1

    # check before writing, so a refused table leaves the sheet untouched
    as_table && _check_table_overlap(sheet,
        CellRange(CellRef(anchor_row, anchor_col), CellRef(anchor_row + row_count, anchor_col + col_count - 1)),
        table_name)

    if write_columnnames
        for c in 1:col_count
            target_cell_ref = CellRef(anchor_row, c + anchor_col - 1)
            sheet[target_cell_ref] = string(columnnames[c])
        end
        start_from_anchor = 0
    end

    data = [process_vector(col) for col in data]
    for r in 1:row_count
        for c in 1:col_count
            target_cell_ref = CellRef(r + anchor_row - start_from_anchor, c + anchor_col - 1)
            v = data[c][r]
            sheet[target_cell_ref] = v
        end
    end

    if as_table
        table_ref = CellRange(
            CellRef(anchor_row, anchor_col),
            CellRef(anchor_row + row_count, anchor_col + col_count - 1),
        )
        t = addtable!(sheet, table_ref; name=table_name, style=table_style)
        !isnothing(totals) && settotals!(sheet, t.name, totals...)
    end

    for (c, width) in widths
        col = anchor_col + c - 1
        setColumnWidth(sheet, CellRange(CellRef(1, col), CellRef(1, col)); width)
    end
    first_data_row = anchor_row + 1 - start_from_anchor
    if row_count > 0
        for (c, fmt) in formats
            col = anchor_col + c - 1
            setUniformFormat(sheet, CellRange(CellRef(first_data_row, col), CellRef(first_data_row + row_count - 1, col)); format=fmt)
        end
    end
    freeze_header && freezePanes(sheet; nrows=anchor_row)

    return nothing
end

# The number of rows in `data`, a vector of columns named by `columnnames`, checked
# to fit in a sheet.
function _check_table_shape(data, columnnames)::Int
    col_count = length(data)
    col_count != length(columnnames) && throw(XLSXError("Column count mismatch between `data` ($col_count columns) and `columnnames` ($(length(columnnames)) columns)."))
    col_count <= 0 && throw(XLSXError("Can't write table with no columns."))
    col_count > EXCEL_MAX_COLS && throw(XLSXError("`data` contains $col_count columns, but Excel only supports up to $EXCEL_MAX_COLS; must reduce `data` size"))
    row_count = length(data[1])
    row_count > EXCEL_MAX_ROWS - 1 && throw(XLSXError("`data` contains $row_count rows, but Excel only supports up to $(EXCEL_MAX_ROWS-1); must reduce `data` size"))
    if col_count > 1
        for c in 2:col_count
            length(data[c]) != row_count && throw(XLSXError("Row count mismatch between column 1 ($row_count rows) and column $c ($(length(data[c])) rows)."))
        end
    end
    return row_count
end

#
# Helper Functions
#

"""
Attempt to derive a table name from `sheetname`, normalized into a valid
Excel Table name via the same identifier-normalization used elsewhere for
column names (spaces/invalid characters become underscores, a bad leading
character is fixed, repeated underscores collapse). Falls back to an
auto-generated name (via a bare call to addtable! with name="") only if the
normalized name still collides with an existing table name or defined
name — sheet names, unlike table names, don't share that namespace.
"""
function _table_name_from_sheet(wb::Workbook, sheet::Worksheet, sheetname::AbstractString)::String
    candidate = string(normalizename(sheetname))

    if candidate != sheetname
        @warn "Sheet name `$sheetname` isn't a valid Excel Table name; using normalized name `$candidate` instead."
    end

    existing_names = Set{String}()
    for ws in wb.sheets
        is_chartsheet(wb, ws.name) && continue
        for t in tables(ws)
            push!(existing_names, uppercase(t.name))
        end
    end

    if uppercase(candidate) ∈ existing_names ||
       is_workbook_defined_name(wb, candidate) || is_worksheet_defined_name(sheet, candidate)
        @warn "Normalized table name `$candidate` (from sheet name `$sheetname`) collides with an existing table name or defined name; using an auto-generated table name instead."
        return ""
    end
    if _is_cell_address_like(candidate)
        @warn "Table name `$candidate` (from sheet name `$sheetname`) is a cell address, which Excel does not allow; using an auto-generated table name instead."
        return ""
    end

    return candidate
end

"""
    writetable(filename, data, columnnames; [overwrite], [sheetname], [anchor_cell], [as_table], [table_name], [table_style], [totals],
               [column_widths], [column_formats], [freeze_header], [streaming], [shared_strings])

- `data` is a vector of columns.
- `columnames` is a vector of column labels.
- `overwrite` is a `Bool` to control if `filename` should be overwritten if already exists.
- `sheetname` is the name for the worksheet.
- `as_table` is a `Bool` to turn the written range into an Excel Table if true.
- `table_name` is the name of the Excel Table (if `as_table=true`)
- `table_style` is the Excel style to use for the Table (if `as_table=true`)
- `totals` defines whether and how Table totals are defined (if `as_table=true`)
- `column_widths`, `column_formats` and `freeze_header` set the columns' widths and
  number formats and freeze the header row (see [`XLSX.writetable!`](@ref)).
- `streaming=true` writes the rows straight into the file, a chunk at a time, rather
  than building the workbook in memory first, so a table of any size is written in
  about the same memory. A Tables.jl table that is read a row at a time (not by
  columns) is never held in memory as a whole. A streamed table starts in cell A1 and
  has no `totals`; strings are written inline unless `shared_strings=true`. See
  [`XLSX.streamxlsx`](@ref) and [Writing large files](@ref).

For more details on `table_name`, `table_style`, `totals`, refer to
[`XLSX.writetable!`](@ref), [`XLSX.addtable!`](@ref) and [`XLSX.settotals!`](@ref)

The other forms of `writetable`, which write several sheets, take `streaming` too.

Returns the filepath of the written file if a filename is supplied, or `nothing` if writing to an `IO`.

# Example

```julia
import XLSX
columns = [ [1, 2, 3, 4], ["Hey", "You", "Out", "There"], [10.2, 20.3, 30.4, 40.5] ]
colnames = [ "integers", "strings", "floats" ]
XLSX.writetable("table.xlsx", columns, colnames)

julia> XLSX.writetable("table.xlsx", columns, colnames; as_table=true, table_name="MyData")

julia> XLSX.writetable("table.xlsx", columns, colnames;
           as_table=true, table_name="MyData",
           totals=["integers" => :sum, "strings" => "Total"])
```

See also: [`XLSX.writetable!`](@ref), [`XLSX.addtable!`](@ref), [`XLSX.settotals!`](@ref).
"""
function writetable(filename::Union{AbstractString,IO}, data, columnnames; overwrite::Bool=false, sheetname::AbstractString="", anchor_cell::Union{AbstractString,CellRef}=CellRef("A1"),
    as_table::Bool=false, table_name::AbstractString="", table_style::Union{AbstractString,Nothing}=nothing,
    totals::Union{Nothing,AbstractVector{<:Pair}}=nothing, streaming::Bool=false, layout...
)

    if filename isa AbstractString && !overwrite
        isfile(filename) && throw(XLSXError("$filename already exists."))
    end

    if isa(anchor_cell, AbstractString)
        anchor_cell = CellRef(String(anchor_cell))
    end

    if streaming
        _check_streaming_options(anchor_cell, totals)
        return _stream_sheets(filename, [(sheetname, (data, columnnames))]; overwrite, as_table,
                              table_name, table_style, layout...)
    end

    xf = open_empty_template(sheetname)
    sheet = xf[1]

    writetable!(sheet, data, columnnames; anchor_cell=anchor_cell, as_table=as_table, table_name=table_name, table_style=table_style, totals=totals, layout...)

    writexlsx(filename, xf, overwrite=overwrite)
end

function _check_streaming_options(anchor_cell::CellRef, totals)
    anchor_cell == CellRef(1, 1) ||
        throw(XLSXError("`streaming=true` writes from cell A1: `anchor_cell` can't be used with it."))
    isnothing(totals) ||
        throw(XLSXError("`streaming=true` can't write a totals row: use `as_table=true` and add totals with `settotals!` after opening the file with `mode=\"rw\"`."))
    return nothing
end

"""
    writetable(filename::Union{AbstractString, IO}; overwrite::Bool=false, as_table::Bool=false, table_style=nothing, kw...)

Write multiple tables.

`kw` is a variable keyword argument list. Each element should be in this format: `sheetname=( data, column_names )`,
where `data` is a vector of columns and `column_names` is a vector of column labels.

Set `as_table=true` to also turn each sheet's written range into an Excel
Table. `table_style` (if given) is applied to every table the same way.

When `as_table=true`, each table's name defaults to its sheet's name. If that
isn't a valid Excel Table name it is normalized (spaces and invalid characters
become underscores) and a warning is issued. If the normalized name collides with
an existing table or defined name, a second warning is issued and an
auto-generated name (`"Table1"`, `"Table2"`, ...) is used instead — sheet names
and table names don't share the same character rules or, in the case of defined
names, the same namespace guarantees.

Example:

```julia
julia> cols_a = [[10, 20, 30], ["Fist", "Sec", "Third"]];

julia> names_a = ["COL1", "COL2"];

julia> cols_b = [["aa", "bb"], [10.1, 10.2]];

julia> names_b = ["AA", "AB"];

julia> XLSX.writetable("report.xlsx"; REPORT_A=(cols_a, names_a), REPORT_B=(cols_b, names_b))

julia> XLSX.writetable("report.xlsx"; overwrite=true, as_table=true, table_style="TableStyleMedium2",
           REPORT_A=(cols_a, names_a), REPORT_B=(cols_b, names_b))
```

!!! note
    Because sheet names are given as keywords, they must be valid Julia
    identifiers — `REPORT_A` works, `"Report A"` does not. They also share
    the keyword namespace with this method's own options so, using this method, 
    a sheet cannot be named `overwrite`, `as_table`, `table_style` or `streaming` — such a
    keyword is taken as the option, and the resulting error will not mention sheet
    names. Use the `Vector{Tuple{String, …}}` form, or `"Report A" => table` pairs,
    for names that don't satisfy this.

See also: [`XLSX.writetable!`](@ref), [`XLSX.addtable!`](@ref), [`XLSX.settotals!`](@ref).
"""
function writetable(filename::Union{AbstractString,IO}; overwrite::Bool=false,
    as_table::Bool=false, table_style::Union{AbstractString,Nothing}=nothing, streaming::Bool=false, kw...)

    for (sheetname, spec) in kw # do this before confirming isfile - ordering pinned in tests
        spec isa Tuple{Any,Any} || throw(XLSXError(
            "Keyword `$sheetname` must be a `(data, columnnames)` tuple, got `$(typeof(spec))`. " *
            "This form takes columns and labels, not a table — to write a Tables.jl source " *
            "directly, use `writetable(filename, :$sheetname => table)`."))
    end

    if filename isa AbstractString && !overwrite
        isfile(filename) && throw(XLSXError("$filename already exists."))
    end

    streaming && return _stream_sheets(filename, [(k, spec) for (k, spec) in kw];
                                       overwrite, as_table, table_style, table_name_from_sheet=true)

    xf = open_empty_template()
    wb = get_workbook(xf)
    is_first = true

    for (sheetname, spec) in kw
        sheetname_str = string(sheetname)
        data, column_names = spec

        if is_first
            sheet = xf[1]
            renamesheet!(sheet, sheetname_str)
            is_first = false
        else
            sheet = addsheet!(xf, sheetname_str)
        end

        tname = as_table ? _table_name_from_sheet(wb, sheet, sheetname_str) : ""
        writetable!(sheet, data, column_names; as_table=as_table, table_name=tname, table_style=table_style)
    end

    writexlsx(filename, xf, overwrite=overwrite)
end

"""
    writetable(filename::Union{AbstractString, IO}, tables::Vector{Tuple{String, Vector{Any}, Vector{String}}}; overwrite::Bool=false, as_table::Bool=false, table_style=nothing)

Write multiple tables.

Each element of `tables` is `(sheetname, data, column_names)`, where `data`
is a vector of columns and `column_names` is a vector of column labels.

Set `as_table=true` to also turn each sheet's written range into an Excel
Table. `table_style` (if given) is applied to every table the same way.

When `as_table=true`, each table's name defaults to its sheet's name. If that
isn't a valid Excel Table name it is normalized (spaces and invalid characters
become underscores) and a warning is issued. If the normalized name collides with
an existing table or defined name, a second warning is issued and an
auto-generated name (`"Table1"`, `"Table2"`, ...) is used instead — sheet names
and table names don't share the same character rules or, in the case of defined
names, the same namespace guarantees.

# Example
```julia
julia> XLSX.writetable("report.xlsx", [
           ("REPORT_A", columns_a, colnames_a),
           ("REPORT_B", columns_b, colnames_b),
       ]; as_table=true)
```

See also: [`XLSX.writetable!`](@ref), [`XLSX.addtable!`](@ref), [`XLSX.settotals!`](@ref).
"""
function writetable(filename::Union{AbstractString,IO}, tables::Vector{Tuple{String,S,Vector{T}}};
    overwrite::Bool=false, as_table::Bool=false, table_style::Union{AbstractString,Nothing}=nothing,
    streaming::Bool=false,
) where {S<:Vector{U} where {U},T<:Union{String,Symbol}}

    if filename isa AbstractString && !overwrite
        isfile(filename) && throw(XLSXError("$filename already exists."))
    end

    streaming && return _stream_sheets(filename, [(n, (d, c)) for (n, d, c) in tables];
                                       overwrite, as_table, table_style, table_name_from_sheet=true)

    xf = open_empty_template()
    wb = get_workbook(xf)
    is_first = true

    for (sheetname, data, column_names) in tables
        if is_first
            sheet = xf[1]
            renamesheet!(sheet, sheetname)
            is_first = false
        else
            sheet = addsheet!(xf, sheetname)
        end

        tname = as_table ? _table_name_from_sheet(wb, sheet, sheetname) : ""
        writetable!(sheet, data, column_names; as_table=as_table, table_name=tname, table_style=table_style)
    end

    writexlsx(filename, xf, overwrite=overwrite)
end
