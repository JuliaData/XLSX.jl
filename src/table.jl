
#
# Table
#

# Returns a tuple with the first and last index of the columns for a `SheetRow`.
function column_bounds(sr::SheetRow)
    isempty(sr) && throw(XLSXError("Can't get column bounds from an empty row."))

    cols = keys(sr.rowcells)
    return (minimum(cols), maximum(cols))
end

# anchor_column will be the leftmost column of the column_bounds
function last_column_index(sr::SheetRow, anchor_column::Int)::Int
    isempty(getcell(sr, anchor_column)) &&
        throw(XLSXError("Can't get column bounds based on an empty anchor cell."))

    cols = sort!(collect(keys(sr.rowcells)))
    first_i = findfirst(==(anchor_column), cols)
    first_i === nothing &&
        throw(XLSXError("Anchor column $anchor_column not present in row."))

    lastcol = anchor_column
    for c in cols[first_i+1:end]
        if c != lastcol + 1
            return lastcol
        end
        lastcol = c
    end
    return lastcol
end

function _colname_prefix_string(sheet::Worksheet, cell::Cell)
    d = getdata(sheet, cell)
    if d isa String
        return d
    else
        return string(d)
    end
end
_colname_prefix_string(::Worksheet, ::EmptyCell) = "#Empty"

# helper function to manage problematic column labels
# Empty cell -> "#Empty"
# No_unique_label -> No_unique_label_2
function push_unique!(vect::Vector{String}, sheet::Worksheet, cell::AbstractCell)
    base = _colname_prefix_string(sheet, cell)
    name = base
    i = 1
    while name in vect
        i += 1
        name = base * "_" * string(i)
    end
    push!(vect, name)
    return nothing
end

# Issue 260
const RESERVED = Set(["local", "global", "export", "let",
    "for", "struct", "while", "const", "continue", "import",
    "function", "if", "else", "try", "begin", "break", "catch",
    "return", "using", "baremodule", "macro", "finally",
    "module", "elseif", "end", "quote", "do"])
normalizename(name::Symbol) = name
function normalizename(name::String)::Symbol
    uname = strip(Unicode.normalize(name))
    id = Base.isidentifier(uname) ? uname : map(c->Base.is_id_char(c) ? c : '_', uname)
    cleansed = string((isempty(id) || !Base.is_id_start_char(id[1]) || id in RESERVED) ? "_" : "", id)
    return Symbol(replace(cleansed, r"(_)\1+"=>"_"))
end

"""
    eachtablerow(sheet, 
                [columns]; 
                [first_row], 
                [column_labels], 
                [header], 
                [stop_in_empty_row], 
                [stop_in_row_function], 
                [keep_empty_rows], 
                [normalizenames],
                [missing_strings]
    ) -> TableRowIterator

Constructs an iterator of table rows. Each element of the iterator is of type `TableRow`.

`header` is a boolean indicating whether the first row of the table is a table header.

If `header == false` and no `column_labels` were supplied, column names will be generated following the column names found in the Excel file.

The `columns` argument is a column range, as in `"B:E"`.
If `columns` is not supplied, the column range will be inferred by the non-empty contiguous cells in the first row of the table.

The user can replace column names by assigning the optional `column_labels` input variable with a `Vector{Symbol}`.

`stop_in_empty_row` is a boolean indicating whether an empty row marks the end of the table.
If `stop_in_empty_row=false`, the iterator will continue to fetch rows until there's no more rows in the Worksheet.
The default behavior is `stop_in_empty_row=true`. Empty rows may be returned by the iterator when `stop_in_empty_row=false`.

`stop_in_row_function` is a Function that receives a `TableRow` and returns a `Bool` indicating if the end of the table was reached.
The row that satisfies `stop_in_row_function` is excluded from the table.

Example for `stop_in_row_function`:

```
function stop_function(r)
    v = r[:col_label]
    return !ismissing(v) && v == "unwanted value"
end
```

`keep_empty_rows` determines whether rows where all column values are equal to `missing` are kept (`true`) or skipped (`false`) by the row iterator.
`keep_empty_rows` never affects the *bounds* of the iterator; the number of rows read from a sheet is only affected by `first_row`, `stop_in_empty_row` and `stop_in_row_function` (if specified).
`keep_empty_rows` is only checked once the first and last row of the table have been determined, to see whether to keep or drop empty rows between the first and the last row.

`normalizenames` controls whether column names will be "normalized" to valid Julia identifiers. By default, this is `false`.
If `normalizenames=true`, then column names with spaces or that start with numbers will be adjusted with underscores to become 
valid Julia identifiers. This is useful when you want to access columns via dot-access or getproperty, like `file.col1`. The 
identifier that comes after the `.` must be valid, so spaces or identifiers starting with numbers aren't allowed.
(Based on CSV.jl's `CSV.normalizename`.)

`missing_strings` can be used to specify strings that should be interpreted  
as `missing` values in the resulting table. `missing_strings` can be a single 
string or a vector of strings. The default value is `missing_strings=nothing`.

Example code:
```
for r in XLSX.eachtablerow(sheet)
    # r is a `TableRow`. Values are read using column labels or numbers.
    rn = XLSX.row_number(r) # `TableRow` row number.
    v1 = r[1] # will read value at table column 1.
    v2 = r[:COL_LABEL2] # will read value at column labeled `:COL_LABEL2`.
end
```

See also [`XLSX.gettable`](@ref).
"""
function eachtablerow(
    sheet::Worksheet,
    cols::Union{ColumnRange,AbstractString};
    first_row::Union{Nothing,Int}=nothing,
    column_labels=nothing,
    header::Bool=true,
    stop_in_empty_row::Bool=true,
    stop_in_row_function::Union{Nothing,Function}=nothing,
    keep_empty_rows::Bool=false,
    normalizenames::Bool=false,
    missing_strings::Union{AbstractString, AbstractVector{<:AbstractString}, Nothing}=nothing
)::TableRowIterator

    ms = if isnothing(missing_strings)
        Set{String}()
    elseif missing_strings isa AbstractString
        Set{String}([missing_strings])
    else
        Set{String}(missing_strings)
    end

    # Validate column_labels length early, before any work is done
    column_range = convert(ColumnRange, cols)
    if !isnothing(column_labels) && length(column_labels) != length(column_range)
        throw(XLSXError("`column_range` (length=$(length(column_range))) and `column_labels` (length=$(length(column_labels))) must have the same length."))
    end

    if isnothing(first_row)
        first_row = _find_first_row_with_data(sheet, column_range.start)
    end

    itr = eachrow(sheet)
    col_lab = Vector{String}()

    if isnothing(column_labels)
        if header
            sheet_row = find_row(itr, first_row)
            for column_index in column_range.start:column_range.stop
                cell = getcell(sheet_row, column_index)
                push_unique!(col_lab, sheet, cell)
            end
        else
            for c in column_range
                push!(col_lab, string(c))
            end
        end
    end

    column_labels = if normalizenames
        normalizename.(isnothing(column_labels) ? col_lab : column_labels)
    else
        Symbol.(isnothing(column_labels) ? col_lab : column_labels)
    end

    first_data_row = header ? first_row + 1 : first_row

    return TableRowIterator(sheet, Index(column_range, column_labels), first_data_row, stop_in_empty_row, stop_in_row_function, keep_empty_rows, ms)
end

function TableRowIterator(sheet::Worksheet, index::Index, first_data_row::Int, stop_in_empty_row::Bool=true, stop_in_row_function::Union{Nothing,Function}=nothing, keep_empty_rows::Bool=false, missing_strings::Set{String}=Set{String}())
    return TableRowIterator(eachrow(sheet), index, first_data_row, stop_in_empty_row, stop_in_row_function, keep_empty_rows, missing_strings)
end

# Detects the contiguous column range starting from `columns_ordered[ci]`
function _detect_column_range(row, columns_ordered::Vector, ci::Int)::ColumnRange
    cn_start = columns_ordered[ci]
    column_stop = cn_start
    for ci_stop in (ci+1):length(columns_ordered)
        cn_stop = columns_ordered[ci_stop]
        # Stop if the next cell is empty or there's a gap in column indices
        if ismissing(getdata(row, cn_stop)) || (cn_stop - 1 != column_stop)
            return ColumnRange(cn_start, column_stop)
        end
        column_stop = cn_stop
    end
    return ColumnRange(cn_start, column_stop)
end

function eachtablerow(
    sheet::Worksheet;
    first_row::Union{Nothing,Int}=nothing,
    column_labels=nothing,
    header::Bool=true,
    stop_in_empty_row::Bool=true,
    stop_in_row_function::Union{Nothing,Function}=nothing,
    keep_empty_rows::Bool=false,
    normalizenames::Bool=false,
    missing_strings::Union{AbstractString, AbstractVector{<:AbstractString}, Nothing}=nothing
)::TableRowIterator

    if isnothing(first_row)
        first_row = 1
    end

    # Bundle shared kwargs to avoid repetition in recursive calls
    shared_kwargs = (; column_labels, header, stop_in_empty_row, stop_in_row_function, keep_empty_rows, normalizenames, missing_strings)

    for r in eachrow(sheet)
        if row_number(r) < first_row || (isempty(r) && !keep_empty_rows)
            continue
        end

        columns_ordered = sort(collect(keys(r.rowcells)))

        # Find the first column with non-missing data
        ci = findfirst(cn -> !ismissing(getdata(r, cn)), columns_ordered)
        if isnothing(ci)
            continue
        end

        first_row = row_number(r)
        column_range = _detect_column_range(r, columns_ordered, ci)
        return eachtablerow(sheet, column_range; first_row, shared_kwargs...)
    end

    throw(XLSXError("Couldn't find a table in sheet $(sheet.name)"))
end

function _find_first_row_with_data(sheet::Worksheet, column_number::Int)
    for r in eachrow(sheet)
        if !ismissing(getdata(r, column_number))
            return row_number(r)
        end
    end
    throw(XLSXError("Column $(encode_column_number(column_number)) has no data."))
end

@inline get_worksheet(tri::TableRowIterator) = get_worksheet(tri.itr)

# Returns real sheet column numbers (based on cellref)
@inline sheet_column_numbers(i::Index) = values(i.column_map)

# Returns an iterator for table column numbers.
@inline table_column_numbers(i::Index) = eachindex(i.column_labels)
@inline table_column_numbers(r::TableRow) = table_column_numbers(r.index)

# Maps table column index (1-based) -> sheet column index (cellref based)
@inline table_column_to_sheet_column_number(index::Index, table_column_number::Int) = index.column_map[table_column_number]
@inline table_columns_count(i::Index) = length(i.column_labels)
@inline table_columns_count(itr::TableRowIterator) = table_columns_count(itr.index)
@inline table_columns_count(r::TableRow) = table_columns_count(r.index)
@inline row_number(r::TableRow) = r.row
@inline get_column_labels(index::Index) = index.column_labels
@inline get_column_labels(itr::TableRowIterator) = get_column_labels(itr.index)
@inline get_column_labels(r::TableRow) = get_column_labels(r.index)
@inline get_column_label(r::TableRow, table_column_number::Int) = get_column_labels(r)[table_column_number]

# Shared iteration logic for TableRow
function _iterate_tablerow(r::TableRow, next)
    isnothing(next) && return nothing
    col, state = next
    return r[col], state
end

Base.iterate(r::TableRow) = _iterate_tablerow(r, iterate(table_column_numbers(r)))
Base.iterate(r::TableRow, state) = _iterate_tablerow(r, iterate(table_column_numbers(r), state))

Base.getindex(r::TableRow, x) = getdata(r, x)

# Helper — apply missing_strings substitution to a single cell value
@inline function _apply_missing_strings(val, ms::Set{String})
    isempty(ms) && return val
    val isa String && val in ms && return missing
    return val
end

function TableRow(table_row::Int, index::Index, sheet_row::SheetRow,
                  missing_strings::Set{String}=Set{String}())
    ws = get_worksheet(sheet_row)

    cell_values = map(table_column_numbers(index)) do table_column_number
        sheet_column = table_column_to_sheet_column_number(index, table_column_number)
        val = getdata(ws, getcell(sheet_row, sheet_column))
        _apply_missing_strings(val, missing_strings)
    end

    return TableRow(table_row, index, cell_values)
end

getdata(r::TableRow, table_column_number::Int) = r.cell_values[table_column_number]
getdata(r::TableRow, table_column_numbers::Union{Vector{T},UnitRange{T}}) where {T<:Integer} =
    CellConcreteType[r.cell_values[x] for x in table_column_numbers]

function getdata(r::TableRow, column_label::Symbol)
    index = r.index
    if haskey(index.lookup, column_label)
        return getdata(r, index.lookup[column_label])
    else
        throw(XLSXError("Invalid column label: $column_label."))
    end
end

# Checks if there are any data inside column range (row not entirely empty)
function is_empty_table_row(itr::TableRowIterator, sheet_row::SheetRow)::Bool
    isempty(sheet_row) && return true
    ws = get_worksheet(itr)
    return all(c -> ismissing(getdata(ws, getcell(sheet_row, c))), sheet_column_numbers(itr.index))
end

Base.IteratorSize(::Type{<:TableRowIterator}) = Base.SizeUnknown()
Base.eltype(::TableRowIterator) = TableRow

# Returns true if the stop_in_row_function exists and signals a stop
@inline _should_stop(itr::TableRowIterator, row::TableRow) =
    !isnothing(itr.stop_in_row_function) && itr.stop_in_row_function(row)

# Handles a gap between expected and actual row numbers.
# Returns: (TableRow, state) if emitting a missing row, :skip to continue past gap, or nothing to stop.
function _handle_gap(itr::TableRowIterator, table_row_index::Int, col_count::Int, expected_row::Int, actual_row::Int, sheet_row, sheet_row_iterator_state)
    itr.stop_in_empty_row && return nothing
    itr.keep_empty_rows || return :skip

    table_row = TableRow(table_row_index, itr.index, fill(missing, col_count))
    _should_stop(itr, table_row) && return nothing
    newstate = TableRowIteratorState(
        table_row_index, expected_row,
        sheet_row_iterator_state,
        actual_row - expected_row - 1,
        sheet_row
    )
    return table_row, newstate
end

# Advances past empty XML rows, respecting keep_empty_rows and stop_in_empty_row.
# Returns: (sheet_row, state) on success, or nothing to stop.
function _skip_empty_rows(itr::TableRowIterator, sheet_row, sheet_row_iterator_state)
    if itr.keep_empty_rows
        is_empty_table_row(itr, sheet_row) && itr.stop_in_empty_row && return nothing
        return sheet_row, sheet_row_iterator_state
    end

    while is_empty_table_row(itr, sheet_row)
        itr.stop_in_empty_row && return nothing
        next = iterate(itr.itr, sheet_row_iterator_state)
        isnothing(next) && return nothing
        sheet_row, sheet_row_iterator_state = next
    end
    return sheet_row, sheet_row_iterator_state
end

# Constructs and returns a data TableRow and its successor state.
function _return_table_row(itr::TableRowIterator, table_row_index::Int,
                           actual_row::Int, sheet_row, sheet_row_iterator_state)
    table_row = TableRow(table_row_index, itr.index, sheet_row, itr.missing_strings)  # ← pass ms
    _should_stop(itr, table_row) && return nothing
    newstate = TableRowIteratorState(table_row_index, actual_row, sheet_row_iterator_state, 0, nothing)
    return table_row, newstate
end

function Base.iterate(itr::TableRowIterator)
    # Advance iterator to first_data_row
    next = iterate(itr.itr)
    while !isnothing(next) && row_number(next[1]) < itr.first_data_row
        next = iterate(itr.itr, next[2])
    end
    isnothing(next) && return nothing

    # Synthesize an initial state as if we just returned the row before first_data_row,
    # with the current sheet_row pending, so the stateful method handles all real logic.
    sheet_row, sheet_row_state = next
    initial_state = TableRowIteratorState(0, itr.first_data_row - 1, sheet_row_state, 0, sheet_row)
    return iterate(itr, initial_state)
end

function Base.iterate(itr::TableRowIterator, state::TableRowIteratorState)
    table_row_index = state.table_row_index + 1
    col_count = length(sheet_column_numbers(itr.index))

    # Emit any pending missing rows before advancing to the next sheet row
    if state.missing_rows > 0
        @assert itr.keep_empty_rows "Inconsistent state: missing_rows > 0 but keep_empty_rows=false"
        table_row = TableRow(table_row_index, itr.index, fill(missing, col_count))
        _should_stop(itr, table_row) && return nothing
        newstate = TableRowIteratorState(
            table_row_index,
            state.sheet_row_index + 1,
            state.sheet_row_iterator_state,
            state.missing_rows - 1,
            state.row_pending
        )
        return table_row, newstate
    end

    # Get next sheet row: from pending (gap case) or from the iterator
    local sheet_row, sheet_row_iterator_state
    if !isnothing(state.row_pending)
        sheet_row = state.row_pending
        sheet_row_iterator_state = state.sheet_row_iterator_state
    else
        next = iterate(itr.itr, state.sheet_row_iterator_state)
        isnothing(next) && return nothing
        sheet_row, sheet_row_iterator_state = next
    end

    actual_row = row_number(sheet_row)
    expected_row = state.sheet_row_index + 1

    # Handle gap between expected and actual row numbers
    if actual_row > expected_row
        result = _handle_gap(itr, table_row_index, col_count, expected_row, actual_row, sheet_row, sheet_row_iterator_state)
        result === :skip || return result  # return if nothing or (TableRow, state)
        # :skip means keep_empty_rows=false — fall through to process actual_row
    end

    # Skip over empty XML rows
    result = _skip_empty_rows(itr, sheet_row, sheet_row_iterator_state)
    isnothing(result) && return nothing
    sheet_row, sheet_row_iterator_state = result

    return _return_table_row(itr, table_row_index, row_number(sheet_row), sheet_row, sheet_row_iterator_state)
end

function infer_eltype(v::Vector{Any})
    isempty(v) && return Any
    hasmissing = false
    t = Any
    for x in v
        if ismissing(x)
            hasmissing = true
        elseif t === Any
            t = typeof(x)
        elseif typeof(x) != t
            t = promote_type(t, typeof(x))
            t === Any && return Any
        end
    end
    return hasmissing ? Union{Missing, t} : t
end

infer_eltype(v::Vector{T}) where T = T


# Address issue 225
#Tables.columnaccess(::Type{<:TableRowIterator}) = true # Not needed, it seems.
function typed_column(v::Vector{Any})
    T = XLSX.infer_eltype(v)
    result = Vector{T}(undef, length(v))
    for (i, x) in enumerate(v)
        result[i] = x
    end
    return result
end
function Tables.columns(tr::TableRowIterator)
    schema = Tables.schema(tr)
    names = schema.names
    rows = Tables.rows(tr)
    collected = collect(rows)
    if isempty(collected)
        return NamedTuple{names}(map(_ -> Any[], names))
    end
    cols = Tables.columntable(collected)
    return map(v -> typed_column(Vector{Any}(v)), cols)
end

function check_table_data_dimension(data::Vector)
    isempty(data) && return
    for (i, col) in enumerate(data)
        isa(col, Vector) || throw(XLSXError("Data type at index $i is not a vector. Found: $(typeof(col))."))
    end
    length(data) == 1 && return
    row_count = length(data[1])
    for (i, col) in enumerate(@view data[2:end])
        length(col) == row_count || throw(XLSXError("Not all columns have the same number of rows. Check column $(i+1)."))
    end
end

function gettable(itr::TableRowIterator; infer_eltypes::Bool=true)::DataTable
    column_labels = get_column_labels(itr)
    columns_count = table_columns_count(itr)
    data = Vector{Any}([Vector{Any}() for _ in 1:columns_count])

    for r in itr
        for (ci, cv) in enumerate(r)
            push!(data[ci], cv)
        end
    end

    if infer_eltypes
        for i in eachindex(data)
            col = data[i]
            T = infer_eltype(col)
            if T !== Any
                data[i] = convert(Vector{T}, col)
            end
        end
    end

    check_table_data_dimension(data)
    return DataTable(data, column_labels)
end

# Shared keyword arguments for gettable/eachtablerow dispatch
const _TABLE_KWARGS = """
    first_row::Union{Nothing,Int}=nothing,
    column_labels=nothing,
    header::Bool=true,
    infer_eltypes::Bool=true,
    stop_in_empty_row::Bool=true,
    stop_in_row_function::Union{Function,Nothing}=nothing,
    keep_empty_rows::Bool=false,
    normalizenames::Bool=false
    missing_strings::Union{AbstractString, AbstractVector{<:AbstractString}, Nothing}=nothing
"""


"""
    gettable(
        sheet,
        [columns];
        [first_row],
        [column_labels],
        [header],
        [infer_eltypes],
        [stop_in_empty_row],
        [stop_in_row_function],
        [keep_empty_rows],
        [normalizenames],
        [missing_strings]
    ) -> DataTable

Returns data from a spreadsheet as a struct `XLSX.DataTable` which
can be passed directly to any function that accepts `Tables.jl` data.
(e.g. `DataFrame` from package `DataFrames.jl`).

Use `columns` argument to specify which columns to get.
For example, `"B:D"` will select columns `B`, `C` and `D`.
If `columns` is not given, the algorithm will find the first sequence
of consecutive non-empty cells.

Use `first_row` to indicate the first row from the table.
`first_row=5` will look for a table starting at sheet row `5`.
If `first_row` is not given, the algorithm will look for the first
non-empty row in the spreadsheet.

`header` is a `Bool` indicating if the first row is a header.
If `header=true` and `column_labels` is not specified, the column labels
for the table will be read from the first row of the table.
If `header=false` and `column_labels` is not specified, the algorithm
will generate column labels. The default value is `header=true`.

Use `column_labels` as a vector of symbols to specify names for the header of the table.

Use `normalizenames=true` to normalize column names to valid Julia identifiers.

Use `missing_strings` to specify strings that should be interpreted as `missing` 
values in the resulting table. `missing_strings` can be a single string or a 
vector of strings. The default value is `missing_strings=nothing`.

Use `infer_eltypes=true` to get `data` as a `Vector{Any}` of typed vectors.
The default value is `infer_eltypes=true`.

`stop_in_empty_row` is a boolean indicating whether an empty row marks the end of the table.
If `stop_in_empty_row=false`, the `TableRowIterator` will continue to fetch rows until there's no more rows in the Worksheet.
The default behavior is `stop_in_empty_row=true`.

`stop_in_row_function` is a Function that receives a `TableRow` and returns a `Bool` indicating if the end of the table was reached.

# Example for `stop_in_row_function`

```julia
function stop_function(r)
    v = r[:col_label]
    return !ismissing(v) && v == "unwanted value"
end
```

`keep_empty_rows` determines whether rows where all column values are equal to `missing` are kept (`true`) or dropped (`false`) from the resulting table.
`keep_empty_rows` never affects the *bounds* of the table; the number of rows read from a sheet is only affected by `first_row`, `stop_in_empty_row` and `stop_in_row_function` (if specified).
`keep_empty_rows` is only checked once the first and last row of the table have been determined, to see whether to keep or drop empty rows between the first and the last row.

# Example

```julia
julia> using DataFrames, PrettyTables, XLSX

julia> df = XLSX.openxlsx("myfile.xlsx") do xf
        DataFrame(XLSX.gettable(xf["mysheet"]))
    end

julia> PrettyTable(XLSX.gettable(xf["mysheet"], "A:C"))
┌─────────┬─────────┬─────────┐
│ Header1 │ Header2 │ Header3 │
├─────────┼─────────┼─────────┤
│       1 │       2 │       3 │
│       4 │       5 │       6 │
│       7 │       8 │       9 │
└─────────┴─────────┴─────────┘
   
```

See also: [`XLSX.readtable`](@ref), [`XLSX.readto`](@ref).
"""
function gettable(sheet::Worksheet, cols::Union{ColumnRange,AbstractString};
    first_row::Union{Nothing,Int}=nothing, column_labels=nothing, header::Bool=true,
    infer_eltypes::Bool=true, stop_in_empty_row::Bool=true,
    stop_in_row_function::Union{Function,Nothing}=nothing,
    keep_empty_rows::Bool=false, normalizenames::Bool=false,
    missing_strings::Union{AbstractString, AbstractVector{<:AbstractString}, Nothing}=nothing
)

    is_chartsheet(get_workbook(sheet), sheet.name) && throw(XLSXError("Can't read a table from a chartsheet."))

    itr = eachtablerow(sheet, cols; first_row, column_labels, header,
                        stop_in_empty_row, stop_in_row_function, keep_empty_rows, normalizenames, missing_strings)
    return gettable(itr; infer_eltypes)
end

function gettable(sheet::Worksheet;
    first_row::Union{Nothing,Int}=nothing, column_labels=nothing, header::Bool=true,
    infer_eltypes::Bool=true, stop_in_empty_row::Bool=true,
    stop_in_row_function::Union{Function,Nothing}=nothing,
    keep_empty_rows::Bool=false, normalizenames::Bool=false,
    missing_strings::Union{AbstractString, AbstractVector{<:AbstractString}, Nothing}=nothing
)

    is_chartsheet(get_workbook(sheet), sheet.name) && throw(XLSXError("Can't read a table from a chartsheet."))

    itr = eachtablerow(sheet; first_row, column_labels, header,
                        stop_in_empty_row, stop_in_row_function, keep_empty_rows, normalizenames, missing_strings)
    return gettable(itr; infer_eltypes)
end

# Finds the start and stop indices of non-empty data along one dimension of a matrix.
# dim=1 scans rows, dim=2 scans columns.
function _find_data_bounds(v::AbstractMatrix, dim::Int)
    n = size(v, dim)
    start = 0
    stop = n
    for i in 1:n
        slice = dim == 1 ? v[i, :] : v[:, i]
        if all(ismissing, slice)
            start != 0 && (stop = i - 1; break)
        else
            start == 0 && (start = i)
        end
    end
    return start, stop
end

function transposetable(m::Matrix; header::Bool=true)
    v = collect(PermutedDimsArray(m, (2, 1)))

    row_start, row_stop = _find_data_bounds(v, 1)
    col_start, col_stop = _find_data_bounds(v, 2)

    row_start == 0 && throw(XLSXError("No data found in matrix."))

    if header
        headers = v[row_start, col_start:col_stop]
        cols = v[row_start+1:row_stop, col_start:col_stop]
    else
        headers = Symbol[]
        cols = v[row_start:row_stop, col_start:col_stop]
    end

    data = Vector{Any}(undef, size(cols, 2))
    for c in axes(cols, 2)
        col = cols[:, c]
        T = infer_eltype(col)
        data[c] = T === Any ? col : convert(Vector{T}, col)
    end

    return data, headers
end

# Normalises column labels to Symbols, with optional name normalisation.
function _normalise_column_labels(labels, normalizenames::Bool)
    normalizenames ? Symbol.(normalizename.(labels)) : Symbol.(labels)
end

# Validates and coerces first_column to Int or nothing.
function _parse_first_column(first_column)
    first_column isa String && return decode_column_number(first_column)
    (first_column isa Int || isnothing(first_column)) && return first_column
    throw(XLSXError("first_column must be an integer column number or a column string like \"A\", \"B\", etc."))
end

"""
    gettransposedtable(
        sheet,
        [rows];
        [first_column],
        [column_labels],
        [header],
        [normalizenames]
    ) -> DataTable

Read a transposed table from a worksheet in which data are arranged in 
rows rather than columns. For example:
```
Category    "A", "B", "C", "D"
variable 1  10,  20,  30,  40
variable 2  15,  25,  35,  40
variable 3  20,  30,  40,  50
```
Returns data from a worksheet as a struct `XLSX.DataTable` which
can be passed directly to any function that accepts `Tables.jl` data.
(e.g. `DataFrame` from package `DataFrames.jl`).

Use the `rows` argument to specify which worksheeet rows to include.
For example, `"2:7"` will select rows 2 to 7 (inclusive).
If `rows` is not given, the algorithm will find the first sequence
of consecutive non-empty cells. If `rows` includes leading or trailing 
rows that are completely empty, these rows will be omitted from the 
returned table. In any case, the table will be truncated at the first 
and last non-empty rows, even if this range is smaller than `rows`. 
A valid `sheet` must be specified when specifying `rows`.

Use `first_column` to indicate the first column of the table. May be given 
as a column number or as a string, so that `first_column="E"` and
`first_column=5` will both look for a table starting at column `5` ("E").
Any leading completely empty columns will be ignored, including 
the `first_column`. If `first_column` is not given, the algorithm will 
look for the first non-empty column in the spreadsheet.

`header` is a `Bool` indicating if the first row is a header.
If `header=true` and `column_labels` is not specified, the column labels
for the table will be read from the first column of the table.
If `header=false` and `column_labels` is not specified, the algorithm
will generate column labels. The default value is `header=true`.

Use `column_labels` as a vector of symbols to specify names for the 
header of the table. If `header=true` and `column_labels` is also given, 
column_labels will be preferred and the first column of the table will 
be ignored.

Use `normalizenames=true` to normalize column names to valid Julia identifiers. 
The default is `normalizenames=false`.

# Examples

```julia
julia> using DataFrames, PrettyTables, XLSX

julia> xf = XLSX.openxlsx("HTable.xlsx")
XLSXFile("HTable.xlsx") containing 4 Worksheets
            sheetname size          range
-------------------------------------------------
               Origin 6x10          B2:K7
               Offset 8x12          A1:L8
             Multiple 8x22          A1:V8
              Example 4x5           B2:F5
              
julia> DataFrame(XLSX.gettransposedtable(xf["Example"]))
4×4 DataFrame
 Row │ Category  Variable 1  Variable 2  Variable 3 
     │ String    Int64       Int64       Int64
─────┼──────────────────────────────────────────────
   1 │ A                 10          15          20
   2 │ B                 20          25          30
   3 │ C                 30          35          40
   4 │ D                 40          40          50

julia> PrettyTable(XLSX.gettransposedtable(xf["Example"]; normalizenames=true))
┌──────────┬────────────┬────────────┬────────────┐
│ Category │ Variable_1 │ Variable_2 │ Variable_3 │
├──────────┼────────────┼────────────┼────────────┤
│        A │         10 │         15 │         20 │
│        B │         20 │         25 │         30 │
│        C │         30 │         35 │         40 │
│        D │         40 │         40 │         50 │
└──────────┴────────────┴────────────┴────────────┘

julia> DataFrame(gettransposedtable(xf["Example"]; header=false))
5×4 DataFrame
 Row │ Col_1     Col_2       Col_3       Col_4      
     │ String    Any         Any         Any
─────┼──────────────────────────────────────────────
   1 │ Category  Variable 1  Variable 2  Variable 3
   2 │ A         10          15          20
   3 │ B         20          25          30
   4 │ C         30          35          40
   5 │ D         40          40          50

```
The worksheet `Multiple` contains two tables side by side, separated by an empty column.
Only the first table is read by default. Read the second table by additionally specifying 
the `first_column`.

```julia
julia> DataFrame(XLSX.gettransposedtable(xf["Multiple"], "2:7"))
9×6 DataFrame
 Row │ Year   Col A  Col B  Col C  Col D    Col E      
     │ Int64  Int64  Int64  Int64  Float64  Any
─────┼─────────────────────────────────────────────────
   1 │  1940      1     10    100      0.1  Hello
   2 │  1950      2     20    200      0.2  2025-12-19
   3 │  1960      3     30    300      0.3  3
   4 │  1970      4     40    400      0.4  3.33
   5 │  1980      5     50    500      0.5  Hello
   6 │  1990      6     60    600      0.6  2025-12-19
   7 │  2000      7     70    700      0.7  3
   8 │  2010      8     80    800      0.8  3.33
   9 │  2020      9     90    900      0.9  true

julia> DataFrame(XLSX.gettransposedtable(xf["Multiple"], "2:7"; first_column="M"))
9×6 DataFrame
 Row │ date   name1    name2    name3  name4     name5      
     │ Int64  Float64  Float64  Bool   Time      Any
─────┼──────────────────────────────────────────────────────
   1 │  1840     12.4    0.045   true  10:22:00  Hello
   2 │  1841     12.6    0.046   true  10:23:00  2025-12-19
   3 │  1842     12.8    0.047  false  10:24:00  3
   4 │  1843     13.0    0.048   true  10:25:00  3.33
   5 │  1844     13.2    0.049  false  10:26:00  Hello
   6 │  1845     13.4    0.05    true  10:27:00  2025-12-19
   7 │  1846     13.6    0.051   true  10:28:00  3
   8 │  1847     13.8    0.052   true  10:29:00  3.33
   9 │  1848     14.0    0.053  false  10:30:00  true

```

See also: [`XLSX.readtransposedtable`](@ref), [`XLSX.readtable`](@ref).
"""
function gettransposedtable(
    sheet::Worksheet,
    rows::Union{AbstractString,Nothing}=nothing;
    first_column=nothing,
    column_labels=nothing,
    header::Bool=true,
    normalizenames::Bool=false
)
    dim = get_dimension(sheet)

    # Resolve row range
    rng = if isnothing(rows)
        RowRange(dim.start.row_number, dim.stop.row_number)
    else
        is_valid_row_range(rows) || throw(XLSXError("Invalid row range: $rows"))
        RowRange(rows)
    end

    if rng.start < dim.start.row_number || rng.stop > dim.stop.row_number
        throw(XLSXError("Row range $rng extends outside sheet dimension ($(dim.start.row_number):$(dim.stop.row_number))"))
    end

    # Resolve and validate first_column
    first_column = _parse_first_column(first_column)

    if !isnothing(first_column) &&
        (first_column < dim.start.column_number || first_column > dim.stop.column_number)
        throw(XLSXError("First column $first_column ($(encode_column_number(first_column))) is outside of sheet dimension ($(dim.start.column_number):$(dim.stop.column_number))"))
    end

    col_start = isnothing(first_column) ? dim.start.column_number : first_column
    start = CellRef(rng.start, col_start)
    stop  = CellRef(rng.stop, dim.stop.column_number)

    # Extract and transpose data
    data, h = transposetable(sheet[CellRange(start, stop)]; header)

    # Resolve column labels
    if isnothing(column_labels)
        column_labels = header ? h : ["Col_$(i)" for i in 1:length(data)]
    end
    column_labels = _normalise_column_labels(column_labels, normalizenames)

    check_table_data_dimension(data)
    return DataTable(data, column_labels)
end

"""
    XLSXFile(table)

Take a `Tables.jl` compatible table and create a new `XLSXFile` object for writing.
Can act as a sink for functions such as `CSV.read`.

# Example
```julia
julia> using CSV, XLSX

julia> xf = CSV.read("iris.csv", XLSXFile)
XLSXFile("blank.xlsx") containing 1 Worksheet
            sheetname size          range
-------------------------------------------------
               Sheet1 151x5         A1:E151
```

"""
function XLSXFile(table)
    Tables.istable(table) || throw(XLSXError("Input must be a Tables.jl compatible table."))
    isempty(Tables.rows(table)) && throw(XLSXError("Cannot create XLSXFile from an empty table."))
    xf = newxlsx()
    writetable!(xf[1], table)
    return xf
end
