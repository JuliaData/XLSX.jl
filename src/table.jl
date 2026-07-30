
#
# Table
#

Base.show(io::IO, dt::DataTable) =
    Base.show(io, MIME"text/plain"(), dt)

Base.show(io::IO, ::MIME"text/plain", dt::DataTable) =
    print(io, "XLSX.DataTable with $(length(dt.data)) columns and $(length(dt.data[1])) rows.")

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
            sheet_row = if is_cache_enabled(sheet)
                find_row(itr, first_row)   # cheap: itr is the persistent cache
            else
                # Streaming mode: avoid restarting the whole iterator just to
                # fetch one already-known row — do a targeted single-row lookup.
                matched = match_rows(sheet, [first_row])
                isempty(matched) ? throw(XLSXError("Row $first_row not found in worksheet $(sheet.name).")) : matched[1]
            end
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

function TableRowIterator(sheet::Worksheet, index::Index, first_data_row::Int, stop_in_empty_row::Bool=true, stop_in_row_function::Union{Nothing,Function}=nothing, keep_empty_rows::Bool=false, missing_strings::Set{String}=Set{String}(), resume::Union{Nothing,Tuple}=nothing)
    return TableRowIterator(eachrow(sheet), index, first_data_row, stop_in_empty_row, stop_in_row_function, keep_empty_rows, missing_strings, resume)
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
    itr = eachrow(sheet)
    next = iterate(itr)
    while next !== nothing
        r, state = next
        if row_number(r) < first_row || (isempty(r) && !keep_empty_rows)
            next = iterate(itr, state)
            continue
        end
        columns_ordered = sort(collect(keys(r.rowcells)))
        # Find the first column with non-missing data
        ci = findfirst(cn -> !ismissing(getdata(r, cn)), columns_ordered)
        if isnothing(ci)
            next = iterate(itr, state)
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
    table_row = TableRow(table_row_index, itr.index, sheet_row, itr.missing_strings)
    _should_stop(itr, table_row) && return nothing
    newstate = TableRowIteratorState(table_row_index, actual_row, sheet_row_iterator_state, 0, nothing)
    return table_row, newstate
end

function Base.iterate(itr::TableRowIterator)
    next = if !isnothing(itr.resume)
        itr.resume  # already-fetched (row, state) — skip the expensive restart
    else
        iterate(itr.itr)
    end
    while !isnothing(next) && row_number(next[1]) < itr.first_data_row
        next = iterate(itr.itr, next[2])
    end
    isnothing(next) && return nothing

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
function typed_column(v::Vector{Any})
    T = infer_eltype(v)
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

#
# ====================================================================================== Excel Tables
#

const REL_TABLE  = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/table"
const MIME_TABLE = "application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"

# SUBTOTAL function codes in the 100s range (ignore manually-hidden rows),
# which is what Excel writes for table totals rows.
const TOTALS_ROW_FUNCTIONS = Dict{Symbol,Tuple{String,Int}}(
    :sum       => ("sum",       109),
    :average   => ("average",   101),
    :counta    => ("count",     103),  # OOXML "count" == Excel COUNTA (counts non-blank, incl. text)
    :count     => ("countNums", 102),  # OOXML "countNums" == Excel COUNT (counts numeric cells only)
    :max       => ("max",       104),
    :min       => ("min",       105),
    :stddev    => ("stdDev",    107),
    :var       => ("var",       110),
)

const TOTALS_FUNCTION_BY_NAME = Dict(first(v) => k for (k, v) in TOTALS_ROW_FUNCTIONS)

function parse_totals_settings(sheet::Worksheet, t::Table)
    t.has_totals_row || return Pair[]

    table_doc = get_xml_data(get_xlsxfile(sheet), _table_part_path(sheet, t.name))
    i, j = get_idces(table_doc, "table", "tableColumns")
    column_nodes = collect(xml_elements(table_doc[i][j]))

    totals_row = t.ref.stop.row_number
    col0 = _col_start(t)
    settings = Pair[]

    for (idx, col_name) in enumerate(t.columns)
        node  = column_nodes[idx]
        func  = get_attr(node, "totalsRowFunction", "")
        label = get_attr(node, "totalsRowLabel", "")

        if func == "custom"
            # A custom function's formula lives only in the cell, so read it
            # before the totals row is overwritten. getFormula prepends "=",
            # but setFormula (used by settotals!) wants a bare formula, so
            # strip it back off.
            fstr = getFormula(sheet, CellRef(totals_row, col0 + idx - 1))
            isnothing(fstr) && throw(XLSXError(
                "Column `$col_name` of table `$(t.name)` is marked as a custom totals function but its totals cell holds no formula."))
            push!(settings, col_name => (:custom, lstrip(fstr, '=')))
        elseif func != ""
            sym = get(TOTALS_FUNCTION_BY_NAME, func, nothing)
            isnothing(sym) && throw(XLSXError("Unrecognized totalsRowFunction `$func` on column `$col_name` of table `$(t.name)`."))
            push!(settings, col_name => sym)
        elseif label != ""
            push!(settings, col_name => label)
        end
    end
    return settings
end

# Compact one-line form — used implicitly by default Vector{Table} printing,
# error messages, and anywhere a Table appears embedded in other output.
function Base.show(io::IO, t::Table)
    print(io, "Table(id = ", t.id, ", \"", t.name, "\", ", t.ref, ", ", length(t.columns), " cols")
    t.has_totals_row && print(io, ", +totals")
    print(io, ")")
end

# Richer multi-line form — used when a single Table is the direct display
# value, e.g. at the REPL: `julia> XLSX.table(sheet, "Sales")`
function Base.show(io::IO, ::MIME"text/plain", t::Table)
    println(io, "XLSX.Table: \"", t.name, "\"", t.name == t.display_name ? "" : " (displayName: \"$(t.display_name)\")")
    println(io, "  id      : ", t.id)
    println(io, "  range   : ", t.ref)
    println(io, "  columns : ", join(t.columns, ", "))
    if isnothing(t.style)
        println(io, "  style   : none")
    else
        style_desc = something(t.style.name, "(unnamed)")
        flags = String[]
        t.style.show_first_column   && push!(flags, "first col")
        t.style.show_last_column    && push!(flags, "last col")
        t.style.show_row_stripes    && push!(flags, "row stripes")
        t.style.show_column_stripes && push!(flags, "col stripes")
        println(io, "  style   : ", style_desc, isempty(flags) ? "" : " (" * join(flags, ", ") * ")")
    end
    print(io, "  totals  : ", t.has_totals_row ? "yes" : "no")
end

# Compact form for the style struct on its own (e.g. `t.style` at the REPL)
function Base.show(io::IO, s::TableStyleInfo)
    print(io, "TableStyleInfo(", something(s.name, "none"))
    flags = String[]
    s.show_first_column   && push!(flags, "first")
    s.show_last_column    && push!(flags, "last")
    s.show_row_stripes    && push!(flags, "rows")
    s.show_column_stripes && push!(flags, "cols")
    isempty(flags) || print(io, ", ", join(flags, "+"))
    print(io, ")")
end

"""
    _find_table_part(sheet::Worksheet, name::AbstractString) -> (rid, path)

Locate the table part for the Excel Table `name` on `sheet`, by walking the
worksheet's `<tableParts>` → `r:id` → worksheet `.rels` → target path. Returns
the relationship id and the package path of `xl/tables/tableN.xml`.

Throws an `XLSXError` if `sheet` has no `<tableParts>` element, no relationship
file, or no table part whose `name` attribute matches.
"""
function _find_table_part(sheet::Worksheet, name::AbstractString)::Tuple{String,String}
    xf = get_xlsxfile(sheet)
    sheet_path = get_relationship_target_by_id("xl", get_workbook(sheet), sheet.relationship_id)
    sheet_dir, sheet_file = rsplit(sheet_path, "/"; limit=2)
    rels_path = "$sheet_dir/_rels/$sheet_file.rels"

    !haskey(xf.data, rels_path) &&
        throw(XLSXError("Internal error: `$(sheet.name)` has table `$name` but no relationship file."))

    rels_root  = root_element(xf.data[rels_path])
    sheet_root = root_element(get_xml_data(xf, sheet_path))

    tp_container_els = elements_with_tag(sheet_root, "tableParts")
    isempty(tp_container_els) &&
        throw(XLSXError("Internal error: `$(sheet.name)` has table `$name` in cache but no <tableParts> element."))
    tp_container = tp_container_els[1]

    rel_els = elements_with_tag(rels_root, "Relationship")
    for tp in elements_with_tag(tp_container, "tablePart")
        rid = get_attr(tp, "r:id")
        k = findfirst(r -> get_attr(r, "Id") == rid, rel_els)
        isnothing(k) && continue
        path = resolve_relative_target(sheet_dir, get_attr(rel_els[k], "Target"))
        get_attr(root_element(get_xml_data(xf, path)), "name") == name && return (rid, path)
    end

    throw(XLSXError("Internal error: could not locate table part for `$name`."))
end

# Convenience wrapper for callers that don't need the relationship id.
_table_part_path(sheet::Worksheet, name::AbstractString)::String = last(_find_table_part(sheet, name))

function parse_table_style_info(table_doc::XML.Node)
    i, j = get_idces(table_doc, "table", "tableStyleInfo")
    isnothing(j) && return nothing

    node = table_doc[i][j]
    attrs = XML.attributes(node)
    isnothing(attrs) && return TableStyleInfo(nothing, false, false, false, false)

    return TableStyleInfo(
        get(attrs, "name", nothing),
        get(attrs, "showFirstColumn", "0") == "1",
        get(attrs, "showLastColumn", "0") == "1",
        get(attrs, "showRowStripes", "0") == "1",
        get(attrs, "showColumnStripes", "0") == "1",
    )
end

"""
    remove_attr!(node::XML.Node, key::String)

Remove the attribute `key` from `node`, if present. No-op if `node` has no
attributes at all, or doesn't have `key`.
"""
function remove_attr!(node::XML.Node, key::String)
    attrs = XML.attributes(node)
    isnothing(attrs) && return nothing
    filter!(p -> first(p) != key, node.attributes)
    return nothing
end

"""
    _is_valid_table_display_name(name::AbstractString) -> Bool

Whether `name` is a valid Excel Table `displayName`: starts with a letter,
underscore, or other Unicode identifier-start character; contains only
identifier characters, underscores, or periods thereafter; no spaces.
Uses the same (Unicode-aware) character classification as `normalizename`,
so a name normalized via `normalizename` is always accepted here.
"""
function _is_valid_table_display_name(name::AbstractString)::Bool
    isempty(name) && return false
    Base.is_id_start_char(first(name)) || return false
    return all(c -> c == '.' || Base.is_id_char(c), name)
end

function parse_table_columns(table_doc::XML.Node)
    i, j = get_idces(table_doc, "table", "tableColumns")
    isnothing(j) && throw(XLSXError("Malformed table part: missing <tableColumns>."))

    columns = String[]
    for col_node in xml_elements(table_doc[i][j])
        localname(col_node) != "tableColumn" && continue
        attrs = XML.attributes(col_node)
        (isnothing(attrs) || !haskey(attrs, "name")) &&
            throw(XLSXError("Malformed <tableColumn>: missing required `name` attribute."))
        push!(columns, attrs["name"])
    end
    return columns
end

function parse_table_xml(table_doc::XML.Node, filename::AbstractString, sheet::Worksheet)::Table
    root_els = xml_elements(table_doc)
    isempty(root_els) && throw(XLSXError("Malformed table part $filename: no root element."))
    table_node = last(root_els)
    localname(table_node) != "table" &&
        throw(XLSXError("Malformed table part $filename: root should be <table>, got <$(localname(table_node))>."))

    attrs = XML.attributes(table_node)
    isnothing(attrs) && throw(XLSXError("Malformed table part $filename: <table> has no attributes."))
    haskey(attrs, "id")   || throw(XLSXError("<table> in $filename missing required `id` attribute."))
    haskey(attrs, "name") || throw(XLSXError("<table> in $filename missing required `name` attribute."))
    haskey(attrs, "ref")  || throw(XLSXError("<table> in $filename missing required `ref` attribute."))

    # Excel signals a totals row via EITHER `totalsRowShown="1"` or a nonzero
    # `totalsRowCount` — hand-authored/older files tend to use the former,
    # Excel's own writer favors the latter (confirmed against a real
    # Excel-saved fixture: `totalsRowCount="1"` with no `totalsRowShown` at all).
    totals_row_shown = get(attrs, "totalsRowShown", "0") == "1"
    totals_row_count = something(tryparse(Int, get(attrs, "totalsRowCount", "0")), 0)
    has_totals = totals_row_shown || totals_row_count > 0

return Table(
        parse(Int, attrs["id"]),
        attrs["name"],
        get(attrs, "displayName", attrs["name"]),
        CellRange(attrs["ref"]),
        parse_table_columns(table_doc),
        has_totals,
        parse_table_style_info(table_doc),
        sheet,
    )
end

function get_worksheet_tables(xf::XLSXFile, ws::Worksheet)::Vector{Table}
    r_ids = get_worksheet_table_rids(xf, ws)
    isempty(r_ids) && return Table[]

    tables = Table[]
    for r_id in r_ids
        target = get_worksheet_relationship_target(xf, ws, r_id)
        table_doc = xmlroot(xf, target)  # table parts are fully parsed like any other part
        push!(tables, parse_table_xml(table_doc, target, ws))
    end
    return tables
end

"""
    tables(ws::Worksheet) -> Vector{Table}

All Excel Tables defined on `ws`, in document order. Empty if none.

# Examples
```julia
julia> XLSX.tables(sheet)
2-element Vector{XLSX.Table}:
 Table(id=1, "IO_Table", A1:C8, 3 cols)
 Table(id=2, "Age_height", E1:G6, 3 cols, +totals)
```
"""
function tables(ws::Worksheet)::Vector{Table}
    if isnothing(ws.tables_cache)
        ws.tables_cache = get_worksheet_tables(get_xlsxfile(ws), ws)
    end
    return ws.tables_cache
end

"""
    table(ws::Worksheet, name::AbstractString) -> Table
    table(wb::Workbook, name::AbstractString) -> Table
    table(xf::XLSXFile, name::AbstractString) -> Table
    table(ws::Worksheet, id::Integer) -> Table
    table(wb::Workbook, id::Integer) -> Table
    table(xf::XLSXFile, id::Integer) -> Table

Look up a single table by name or workbook-scoped numeric id, searching
a single worksheet or across every worksheet in the workbook. 

Throws `KeyError` if not found.

# Examples
```julia
julia> XLSX.table(sheet, "Age_height")
XLSX.Table: "Age_height"
  id      : 2
  range   : E1:G6
  columns : name, age, height
  style   : TableStyleMedium2 (row stripes)
  totals  : yes

julia> XLSX.table(sheet, 2)  # same table, looked up by id
XLSX.Table: "Age_height"
  id      : 2
  range   : E1:G6
  columns : name, age, height
  style   : TableStyleMedium2 (row stripes)
  totals  : yes
```
"""
function table(ws::Worksheet, name::AbstractString)::Table
    idx = findfirst(t -> t.name == name, tables(ws))
    isnothing(idx) && throw(KeyError(name))
    return tables(ws)[idx]
end

function table(ws::Worksheet, id::Integer)::Table
    idx = findfirst(t -> t.id == id, tables(ws))
    isnothing(idx) && throw(KeyError(id))
    return tables(ws)[idx]
end
function table(wb::Workbook, name::AbstractString)::Table
    for ws in wb.sheets
        is_chartsheet(wb, ws.name) && continue
        idx = findfirst(t -> t.name == name, tables(ws))
        idx !== nothing && return tables(ws)[idx]
    end
    throw(KeyError(name))
end
function table(wb::Workbook, id::Integer)::Table
    for ws in wb.sheets
        is_chartsheet(wb, ws.name) && continue
        idx = findfirst(t -> t.id == id, tables(ws))
        idx !== nothing && return tables(ws)[idx]
    end
    throw(KeyError(id))
end
table(xf::XLSXFile, name::AbstractString) = table(get_workbook(xf), name)
table(xf::XLSXFile, id::Integer) = table(get_workbook(xf), id)

function build_table_xml(id::Int, name::String, display_name::String, ref::CellRange,
                          columns::Vector{String}, has_totals_row::Bool,
                          style::Union{TableStyleInfo,Nothing})::XML.Node

    buf = IOBuffer()
    print(buf, """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>""")
    print(buf, """<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" """)
    print(buf, """id="$(id)" name="$(XML.escape(name))" displayName="$(XML.escape(display_name))" ref="$(ref)\"""")
    # in build_table_xml:
    if has_totals_row
        print(buf, """ totalsRowShown="1" totalsRowCount="1\"""")
    end
    print(buf, ">")

    # autoFilter covers header+data rows only, excluding any totals row
    filter_stop_row = has_totals_row ? ref.stop.row_number - 1 : ref.stop.row_number
    filter_ref = CellRange(CellRef(ref.start.row_number, ref.start.column_number),
                            CellRef(filter_stop_row, ref.stop.column_number))
    print(buf, """<autoFilter ref="$(filter_ref)"/>""")

    print(buf, """<tableColumns count="$(length(columns))">""")
    for (i, col_name) in enumerate(columns)
        print(buf, """<tableColumn id="$(i)" name="$(XML.escape(col_name))"/>""")
    end
    print(buf, "</tableColumns>")

    if !isnothing(style)
        print(buf, "<tableStyleInfo")
        !isnothing(style.name) && print(buf, """ name="$(style.name)\"""")
        print(buf, """ showFirstColumn="$(style.show_first_column ? "1" : "0")\"""")
        print(buf, """ showLastColumn="$(style.show_last_column ? "1" : "0")\"""")
        print(buf, """ showRowStripes="$(style.show_row_stripes ? "1" : "0")\"""")
        print(buf, """ showColumnStripes="$(style.show_column_stripes ? "1" : "0")\"""")
        print(buf, "/>")
    end

    print(buf, "</table>")
    return parse(String(take!(buf)), XML.Node)
end

"""
    addtable!(
        sheet::Worksheet,
        ref::Union{CellRange,AbstractString};
        name::AbstractString="",
        style::Union{AbstractString,Nothing}=nothing,
        has_totals_row::Bool=false,
    ) -> Table

Create a new Excel Table over `ref` on `sheet`, turning an existing range of
cells into a Table object (banding, filter dropdowns, structured references,
and the `Table` metadata itself) without changing any of the underlying data.

`ref` must already contain data before calling `addtable!`: specifically, the
first row of `ref` is read as the table's header row, so every cell in that
row must already contain the column name you want (write these first, e.g.
with `sheet[...] = ...` or [`XLSX.writetable!`](@ref)). `addtable!` only wraps
existing cells in a Table; it does not write any header, data, or totals
values into the sheet itself.

`ref` must span at least two rows — a header row plus at least one data row.
Excel does not support header-only tables (a single-row `ref` is rejected).

If `name` is not given, a unique name is generated (`"Table1"`, `"Table2"`,
...). Table names are workbook-scoped and must not collide with another
table's name or with a defined name anywhere in the workbook.

`style` sets the table's visual style and should be one of Excel's built-in
table style names, matching the "Table Styles" gallery in Excel's Table
Design ribbon:
- `"TableStyleLightnn"` where nn is between 1 and 21
- `"TableStyleMediumnn"` where nn is between 1 and 28
- `"TableStyleDarknn"` where nn is between 1 and 11
- `"None"` for no style. 

`style` is not validated and any string is passed straight through as the
`tableStyleInfo`'s `name` attribute; an unrecognized name will make Excel 
fall back to its default table appearance rather than causing an error. 
If omitted, Excel treats the table as having no explicit style (its own 
default appearance applies).

`has_totals_row=true` marks the last row of `ref` as the table's totals row. 
This only sets the flag that tells Excel to *reserve and display* that row as 
a totals row; it does not populate any totals formula or label into the cells
themselves — write whatever content you want into that row's cells yourself
(or leave them blank) before or after calling `addtable!`, or use
[`XLSX.settotals!`](@ref) afterward to set per-column totals functions or
labels. If the last row of `ref` already has content when `has_totals_row=true`
is given, a warning is issued (not an error) — that row is still marked as
the totals row regardless, since pre-existing content there may be
intentional (e.g. a pre-authored totals formula or label).

# Examples
```julia
julia> sheet[1, :] = ["id", "name", "score"]
julia> sheet[2, :] = [1, "alice", 10.5]
julia> sheet[3, :] = [2, "bob", 20.0]

julia> XLSX.addtable!(sheet, "A1:C3"; name="Results", style="TableStyleMedium2")
```

See also [`XLSX.tables`](@ref), [`XLSX.table`](@ref), [`XLSX.deletetable!`](@ref), [`XLSX.settotals!`](@ref).
"""
function addtable!(sheet::Worksheet, ref::CellRange;
                    name::AbstractString="",
                    style::Union{AbstractString,Nothing}=nothing,
                    has_totals_row::Bool=false)::Table

    xf = get_xlsxfile(sheet)
    !is_writable(xf) && throw(XLSXError("XLSXFile instance is not writable. Open Excel file with `mode=\"rw\"` instead"))
    wb = get_workbook(sheet)

    # --- name uniqueness (unchanged) ---
    existing_names = Set{String}()
    for ws in wb.sheets
        is_chartsheet(wb, ws.name) && continue
        for t in tables(ws)
            push!(existing_names, t.name)
        end
    end

    if name == ""
        i = 1
        candidate = "Table$i"
        while candidate ∈ existing_names ||
              is_workbook_defined_name(wb, candidate) || is_worksheet_defined_name(sheet, candidate)
            i += 1
            candidate = "Table$i"
        end
        name = candidate
    else
        name ∈ existing_names && throw(XLSXError("Table name `$name` is already in use."))
        (is_workbook_defined_name(wb, name) || is_worksheet_defined_name(sheet, name)) &&
            throw(XLSXError("Table name `$name` collides with an existing defined name."))
    end

    display_name = name
    !_is_valid_table_display_name(display_name) &&
        throw(XLSXError("Table displayName `$display_name` is not a valid identifier (must start with a letter or underscore, contain only letters, digits, underscores, or periods thereafter, and no spaces)."))
        
    ref.stop.row_number == ref.start.row_number &&
        throw(XLSXError("Table `ref` must span at least two rows (a header row plus at least one data row) — Excel does not support header-only tables. Got `$ref`."))

    # --- derive column names from the header row (unchanged) ---
    header_row = ref.start.row_number
    columns = String[]
    for c in column_number(ref.start):column_number(ref.stop)
        v = getdata(sheet, CellRef(header_row, c))
        (ismissing(v) || v == "") &&
            throw(XLSXError("Table header cell $(CellRef(header_row, c)) is empty; every column needs a header value before calling `addtable!`."))
        push!(columns, string(v))
    end
    length(unique(columns)) != length(columns) &&
        throw(XLSXError("Table header row contains duplicate column names."))

    # --- totals row: last row of `ref`, when requested. Warn (not error) if
    # it already holds content — could be pre-authored totals, could be a
    # `writetable!` data row the caller forgot to exclude from `ref`. Either
    # way, `has_totals_row=true` always wins: that row is the totals row.
    if has_totals_row
        totals_row_num = ref.stop.row_number
        nonempty_cols = String[]
        for c in column_number(ref.start):column_number(ref.stop)
            v = getdata(sheet, CellRef(totals_row_num, c))
            (!ismissing(v) && v != "") && push!(nonempty_cols, columns[c - column_number(ref.start) + 1])
        end
        if !isempty(nonempty_cols)
            @warn "Table `$name`: last row of `ref` (row $totals_row_num) is being marked as the " *
                  "totals row, but it already has content in column(s) $(join(nonempty_cols, ", ")). " *
                  "If that row was meant to be table data (e.g. written by `writetable!`), exclude " *
                  "it from `ref` and pass `has_totals_row=false`."
        end
    end

    style_info = isnothing(style) ? nothing : TableStyleInfo(style, false, false, true, false)

    id = next_table_id!(wb)
    table_doc  = build_table_xml(id, name, display_name, ref, columns, has_totals_row, style_info)
    table_path = new_table_filename(xf)
    xf.data[table_path]  = table_doc
    xf.files[table_path] = true
    add_override!(xf, "/$table_path", MIME_TABLE)

    sheet_path = get_relationship_target_by_id("xl", wb, sheet.relationship_id)
    sheet_dir, _ = rsplit(sheet_path, "/"; limit=2)
    rels_path, rels_root = get_or_create_worksheet_rels!(xf, sheet_path)

    rid = new_relationship_id(rels_root)
    pfx_rels = get_prefix(rels_path, xf)
    push!(rels_root, XML.Element(prefixed_tag(pfx_rels, "Relationship");
        Id=rid, Type=REL_TABLE, Target=make_relative_target(sheet_dir, table_path)))

    sheet_doc  = get_xml_data(xf, sheet_path)
    sheet_root = root_element(sheet_doc)
    pfx   = get_prefix(sheet)
    pfx_c = pfx == "" ? "" : "$(pfx):"

    tp_container_els = elements_with_tag(sheet_root, "tableParts")
    if isempty(tp_container_els)
        tp_container = XML.Element("$(pfx_c)tableParts"; count="1")
        tp_node = XML.Element("$(pfx_c)tablePart")
        tp_node["r:id"] = rid
        push!(tp_container, tp_node)
        children = XML.children(sheet_root)
        ext_idx = findfirst(c -> localname(c) == "extLst", children)
        isnothing(ext_idx) ? push!(sheet_root, tp_container) : insert!(children, ext_idx, tp_container)
    else
        tp_container = tp_container_els[1]
        tp_node = XML.Element("$(pfx_c)tablePart")
        tp_node["r:id"] = rid
        push!(tp_container, tp_node)
        tp_container["count"] = string(length(XML.children(tp_container)))
    end

    sheet.tables_cache = nothing
    return table(sheet, name)
end

addtable!(sheet::Worksheet, ref::AbstractString; kw...) = addtable!(sheet, CellRange(ref); kw...)

"""
    deletetable!(sheet::Worksheet, name::AbstractString)
    deletetable!(sheet::Worksheet, id::Integer)

Delete the given Excel Table from `sheet` by `name` or by its workbook-scoped
numeric `id`.

This removes the table *object* only: its `xl/tables/tableN.xml` part, its
worksheet-level relationship, its `<tablePart>` entry, and its
`[Content_Types].xml` override. It does **not** clear, delete, or modify any
of the underlying cell data — the header row, data rows, and any totals row
are left completely untouched, still holding whatever values they held while
the table existed.

This mirrors what Excel itself does when you use **Table Design → Convert to
Range** (or right-click → Table → Convert to Range): the table's structure
(banding, filter dropdowns, structured references, and the Table object
itself) is removed, but the cells and their values remain in place as an
ordinary range. There is no single-step Excel operation that removes a table
and its data together; if you want the data gone too, clear or delete those
cells yourself as a separate step, e.g.:

```julia
julia> t = XLSX.table(sheet, "MyTable")

julia> XLSX.deletetable!(sheet, "MyTable")

julia> sheet[t.ref] = missing   # optional: also clear the data
```

Other tables on the same sheet, and tables on other sheets, are unaffected.

See also [`XLSX.addtable!`](@ref), [`XLSX.tables`](@ref), [`XLSX.table`](@ref).
"""
function deletetable!(sheet::Worksheet, name::AbstractString)
    xf = get_xlsxfile(sheet)
    !is_writable(xf) && throw(XLSXError("XLSXFile instance is not writable. Open Excel file with `mode=\"rw\"` instead"))

    table(sheet, name)  # throws KeyError if not found — fail fast

    target_rid, target_path = _find_table_part(sheet, name)

    sheet_path = get_relationship_target_by_id("xl", get_workbook(sheet), sheet.relationship_id)
    sheet_dir, sheet_file = rsplit(sheet_path, "/"; limit=2)
    rels_path  = "$sheet_dir/_rels/$sheet_file.rels"
    rels_root  = root_element(xf.data[rels_path])
    sheet_root = root_element(get_xml_data(xf, sheet_path))

    # remove <tablePart>, shrink/drop <tableParts>
    tp_container = elements_with_tag(sheet_root, "tableParts")[1]
    tp_children = XML.children(tp_container)
    deleteat!(tp_children, findfirst(tp -> get_attr(tp, "r:id") == target_rid, tp_children))
    if isempty(tp_children)
        deleteat!(XML.children(sheet_root), findfirst(c -> localname(c) == "tableParts", XML.children(sheet_root)))
    else
        tp_container["count"] = string(length(tp_children))
    end

    # remove the relationship entry itself
    deleteat!(XML.children(rels_root), findfirst(r -> get_attr(r, "Id") == target_rid, XML.children(rels_root)))

    delete_part_and_orphans!(xf, target_path)  # handles the table part + its Content_Types override

    sheet.tables_cache = nothing
    return nothing
end

deletetable!(sheet::Worksheet, id::Integer) = deletetable!(sheet, table(sheet, id).name)

"""
    settotals!(sheet::Worksheet, name::AbstractString, settings::Pair...)
    settotals!(sheet::Worksheet, id::Integer, settings::Pair...)
    settotals!(sheet::Worksheet, name::AbstractString; kwargs...)

Add or update the totals row for the Excel Table in `sheet` with the specified 
`name` or workbook-scoped numeric `id`.

Each element of `settings` is `"ColumnName" => value` (or, in the kwarg form,
`ColumnName=value` for identifier-safe column names), where `value` is one of:

- a `Symbol` naming a built-in totals function — `:sum`, `:average`, `:count`,
  `:counta`, `:max`, `:min`, `:stddev`, or `:var`. Writes both the
  `totalsRowFunction` attribute and an actual `SUBTOTAL(...)` formula into
  the totals row cell for that column; Excel performs the calculation from
  this formula exactly as it would for any other `SUBTOTAL` formula.
- a `(:custom, formula::AbstractString)` tuple, for a custom totals function
  — Excel's "More Functions..." option. `formula` is written verbatim into
  the cell (it need not be `SUBTOTAL`-based at all), and
  `totalsRowFunction="custom"` is set. XLSX.jl does not validate or evaluate
  the formula; Excel computes it on open, same as any other formula cell.
- an `AbstractString`, written as a plain text label in that column's totals
  row cell (e.g. `"Grand Total"`), with no function or formula attached.

Columns not mentioned in `settings` are left untouched: if the table already
has a totals row, their existing totals content (function, label, or blank)
is preserved as-is.

If the table does not already have a totals row, one is added by extending
the table by one row — the row immediately following its current last row.
That row must be completely empty; an `XLSXError` is thrown otherwise (clear
it first, or use [`XLSX.addtable!`](@ref) with `has_totals_row=true` if that
row was always meant to be part of the table).

As with every formula-writing path in XLSX.jl, no cached value is written
alongside a totals-row formula — Excel recalculates it on open
(`update_workbook_xml!` forces `fullCalcOnLoad="1"`). Formulas are written
via [`XLSX.setFormula`](@ref), which fully replaces the cell's formula/value
while preserving its existing style, so calling `settotals!` again on a
column that already has totals content (function, custom formula, or label)
cleanly replaces it.

# Examples
```julia
julia> XLSX.settotals!(sheet, "Sales",
           "Revenue" => :sum,
           "Notes"   => "Grand Total",
           "Margin"  => (:custom, "SUBTOTAL(109,Sales[Revenue])-SUBTOTAL(109,Sales[Cost])"),
       )

julia> XLSX.settotals!(sheet, "Sales"; Revenue=:sum, Margin=:average)

julia> XLSX.settotals!(sheet, tbl.id, "Revenue" => :sum)  # by workbook-scoped table id

julia> # Add a totals row to a table that doesn't have one yet — the row
       # immediately following the table's current last row must be empty.
       XLSX.settotals!(sheet, "Sales", "Revenue" => :sum)

julia> # Update just one column's totals function on a table that already
       # has a totals row — other columns' existing totals are untouched.
       XLSX.settotals!(sheet, "Sales", "Revenue" => :max)

julia> # Blank out a column's totals cell entirely by omitting it from
       # `settings` — no call needed; simply don't mention that column.

julia> # A custom formula MUST aggregate each column reference itself
       # (e.g. via SUBTOTAL or SUM) — a bare `Sales[Revenue]` in a totals
       # cell is rewritten by Excel to a "this row" reference, which has
       # no valid row to intersect against in the totals row and raises
       # #VALUE!.
       XLSX.settotals!(sheet, "Sales",
           "Margin" => (:custom, "SUBTOTAL(109,Sales[Revenue])-SUBTOTAL(109,Sales[Cost])"),
       )
```

See also [`XLSX.addtable!`](@ref), [`XLSX.tables`](@ref), [`XLSX.table`](@ref).
"""
function settotals!(sheet::Worksheet, name::AbstractString, settings::Pair...)
    xf = get_xlsxfile(sheet)
    !is_writable(xf) && throw(XLSXError("XLSXFile instance is not writable. Open Excel file with `mode=\"rw\"` instead"))

    t = table(sheet, name)  # throws KeyError if missing

    for (col, _) in settings
        col ∈ t.columns || throw(XLSXError("Column `$col` not found in table `$name`. Available columns: $(join(t.columns, ", "))."))
    end

    table_path = _table_part_path(sheet, name)
    table_doc  = get_xml_data(xf, table_path)
    table_root = root_element(table_doc)

    old_ref = t.ref
    local totals_row_num::Int

    if !t.has_totals_row
        new_last_row = old_ref.stop.row_number + 1
        new_last_row > EXCEL_MAX_ROWS &&
            throw(XLSXError("Cannot add a totals row to table `$name`: table already reaches the last worksheet row."))

        for c in column_number(old_ref.start):column_number(old_ref.stop)
            v = getdata(sheet, CellRef(new_last_row, c))
            (!ismissing(v) && v != "") &&
                throw(XLSXError("Cannot add a totals row to table `$name`: row $new_last_row already has content at $(CellRef(new_last_row, c)). Clear that row first, or use `addtable!` with `has_totals_row=true` if this row was always meant to be part of the table."))
        end

        new_ref = CellRange(old_ref.start, CellRef(new_last_row, old_ref.stop.column_number))
        table_root["ref"] = string(new_ref)
        table_root["totalsRowShown"] = "1"
        table_root["totalsRowCount"] = "1"

        totals_row_num = new_last_row
    else
        totals_row_num = old_ref.stop.row_number
    end

    i, j = get_idces(table_doc, "table", "tableColumns")
    column_nodes = collect(xml_elements(table_doc[i][j]))

    for (col_name, value) in settings
        col_idx  = findfirst(==(col_name), t.columns)
        col_node = column_nodes[col_idx]
        cell_ref = CellRef(totals_row_num, column_number(old_ref.start) + col_idx - 1)

        remove_attr!(col_node, "totalsRowFunction")
        remove_attr!(col_node, "totalsRowLabel")

        if value isa Symbol
            haskey(TOTALS_ROW_FUNCTIONS, value) ||
                throw(XLSXError("Unknown totals function `:$value`. Supported: $(join(sort(string.(keys(TOTALS_ROW_FUNCTIONS))), ", ")), or pass `(:custom, \"formula\")` for a custom function."))
            func_name, subtotal_code = TOTALS_ROW_FUNCTIONS[value]
            col_node["totalsRowFunction"] = func_name
            setFormula(sheet, cell_ref; val="SUBTOTAL($subtotal_code,$(name)[$col_name])")

        elseif value isa Tuple{Symbol,S} where {S<:AbstractString}
            fn, formula_str = value
            fn == :custom || throw(XLSXError("Tuple totals setting for column `$col_name` must be `(:custom, formula)`; got `(:$fn, ...)`."))
            col_node["totalsRowFunction"] = "custom"
            setFormula(sheet, cell_ref; val=formula_str)

        elseif value isa AbstractString
            col_node["totalsRowLabel"] = value
            sheet[cell_ref] = value

        else
            throw(XLSXError("Totals setting for column `$col_name` must be a Symbol (built-in function), a `(:custom, formula)` tuple, or a String (label); got $(typeof(value))."))
        end
    end

    sheet.tables_cache = nothing
    return table(sheet, name)
end

settotals!(sheet::Worksheet, id::Integer, settings::Pair...) =
    settotals!(sheet, table(sheet, id).name, settings...)

settotals!(sheet::Worksheet, name::AbstractString; kwargs...) =
    settotals!(sheet, name, (String(k) => v for (k, v) in kwargs)...)


"""
    gettable(t::Table; [infer_eltypes], [normalizenames], [missing_strings]) -> DataTable

Returns data from an Excel Table `t` (as returned by [`XLSX.table`](@ref)) as
a struct `XLSX.DataTable`, which can be passed directly to any function that
accepts `Tables.jl` data (e.g. `DataFrame` from package `DataFrames.jl`).

!!! note "Two `gettable` methods"

    Different from `XLSX.gettable(sheet, ...)`, which infers a
    table's row/column bounds heuristically from cell content on a plain
    `Worksheet`. Here, `t.ref` is authoritative, so there is no
    `columns`/`first_row`/`header`/`stop_in_empty_row`/`stop_in_row_function`/
    `keep_empty_rows` equivalent — the header row and totals row (if any)
    are always excluded, and any blank row within `t.ref` is returned as
    ordinary data.

Use `normalizenames=true` to normalize column names to valid Julia
identifiers.

Use `missing_strings` to specify strings that should be interpreted as
`missing` values in the resulting table. `missing_strings` can be a single
string or a vector of strings. The default value is `missing_strings=nothing`.

Use `infer_eltypes=true` (the default) to have each column narrowed to its
own concrete type (e.g. `Vector{Float64}` rather than `Vector{Any}`), the
same narrowing [`XLSX.eachtablerow`](@ref)/`Tables.columns` apply. Set
`infer_eltypes=false` to skip narrowing and leave every column as `Any`.

# Example
```julia
julia> using DataFrames

julia> t = XLSX.table(sheet, "Sales")

julia> df = DataFrame(XLSX.gettable(t))
```

See also: [`XLSX.table`](@ref), [`XLSX.tables`](@ref), [`XLSX.eachtablerow`](@ref), [`XLSX.readtable`](@ref).
"""
function gettable(t::Table;
    infer_eltypes::Bool=true,
    normalizenames::Bool=false,
    missing_strings::Union{AbstractString,AbstractVector{<:AbstractString},Nothing}=nothing
)::DataTable

    missing_set = missing_strings === nothing ? nothing :
                  missing_strings isa AbstractString ? Set([missing_strings]) : Set(missing_strings)

    row_range = _first_data_row(t):_last_data_row(t)
    col0 = _col_start(t)

    data = Vector{Any}(undef, length(t.columns))
    for i in eachindex(t.columns)
        col = Any[getdata(t.sheet, CellRef(r, col0 + i - 1)) for r in row_range]
        if !isnothing(missing_set)
            col = Any[(x isa AbstractString && x in missing_set) ? missing : x for x in col]
        end
        data[i] = infer_eltypes ? typed_column(col) : col
    end

    column_labels = normalizenames ? normalizename.(t.columns) : Symbol.(t.columns)

    return DataTable(data, column_labels)
end

"""
    appendtable!(sheet::Worksheet, name::AbstractString, data; [check_empty]) -> Table

Append rows to the existing Excel Table `name` on `sheet`, extending the
table's range.

`data` may be any `Tables.jl`-compatible source (e.g. an `XLSX.DataTable` or a
`DataFrame`), an `AbstractMatrix`, or a vector of row vectors/tuples.

If `data` exposes column names, columns are matched **by name** and reordered
to the table's own column order; a source missing any of the table's columns,
or carrying any column the table doesn't have, is an error. Sources without
column names (matrices, vectors of tuples/vectors) are matched
**positionally**, so their column order must match the table's.

If the table has a totals row, it moves down to remain the last row of the
table, and its content is regenerated from the table's own per-column
totals settings. Functions, custom formulas and labels are all preserved 
but values are reset to `missing`.

The rows immediately below the table must be empty; an `XLSXError` is thrown
otherwise. Pass `check_empty=false` to overwrite whatever is there.

See also [`XLSX.addtable!`](@ref), [`XLSX.settotals!`](@ref).
"""
function appendtable!(sheet::Worksheet, name::AbstractString, data; check_empty::Bool=true)
    xf = get_xlsxfile(sheet)
    !is_writable(xf) && throw(XLSXError("XLSXFile instance is not writable. Open Excel file with `mode=\"rw\"` instead"))

    t = table(sheet, name)
    ncols = length(t.columns)
    rows = _normalize_append_rows(data, t, name)
#    rows = _normalize_append_rows(data, ncols, name)
    n = length(rows)
    n == 0 && return t

    col0 = _col_start(t)
    old_stop = t.ref.stop.row_number
    new_stop = old_stop + n
    new_stop > EXCEL_MAX_ROWS &&
        throw(XLSXError("Appending $n rows to table `$name` would exceed Excel's row limit."))

    # capture totals settings before we overwrite the old totals row
    totals_settings = parse_totals_settings(sheet, t)

    if check_empty
        for r in (old_stop + 1):new_stop, c in col0:(col0 + ncols - 1)
            v = getdata(sheet, CellRef(r, c))
            (!ismissing(v) && v != "") && throw(XLSXError(
                "Cannot append to table `$name`: cell $(CellRef(r, c)) is not empty. " *
                "Clear the rows below the table, or pass `check_empty=false` to overwrite."))
        end
    end

    # first appended row lands on the old totals row position when one exists,
    # otherwise immediately below the current last data row
    first_new = t.has_totals_row ? old_stop : old_stop + 1
    for (ri, row) in enumerate(rows)
        for ci in 1:ncols
            sheet[CellRef(first_new + ri - 1, col0 + ci - 1)] = row[ci]
        end
    end

    # clear the new totals row position before settotals! repopulates it
    if t.has_totals_row
        for c in col0:(col0 + ncols - 1)
            sheet[CellRef(new_stop, c)] = missing
        end
    end

    # extend ref and autoFilter in the table part
    table_doc  = get_xml_data(xf, _table_part_path(sheet, name))
    table_root = root_element(table_doc)
    new_ref = CellRange(t.ref.start, CellRef(new_stop, t.ref.stop.column_number))
    table_root["ref"] = string(new_ref)

    af = elements_with_tag(table_root, "autoFilter")
    if !isempty(af)
        af_stop = t.has_totals_row ? new_stop - 1 : new_stop
        af[1]["ref"] = string(CellRange(t.ref.start, CellRef(af_stop, t.ref.stop.column_number)))
    end

    sheet.tables_cache = nothing

    # replay the totals row at its new position
    isempty(totals_settings) || settotals!(sheet, name, totals_settings...)

    sheet.tables_cache = nothing
    return table(sheet, name)
end

"""
Normalize whatever was passed as `data` into a vector of row-value vectors,
each of length `length(t.columns)`, ordered to match the table's columns.

For `Tables.jl` sources that expose column names (a `DataTable`, `DataFrame`,
etc.), columns are matched **by name** and reordered into the table's own
column order; both missing and extra columns are errors. Matrices, vectors of
tuples and vectors of vectors carry no names, so those are matched
**positionally**.
"""
function _normalize_append_rows(data, t::Table, name::AbstractString)
    ncols = length(t.columns)

    rows = if data isa AbstractMatrix
        [collect(data[r, :]) for r in axes(data, 1)]

    elseif data isa AbstractVector && (isempty(data) || first(data) isa Union{AbstractVector,Tuple})
        [collect(r) for r in data]

    else
        sch = Tables.schema(data)
        src_names = if !isnothing(sch) && !isnothing(sch.names)
            sch.names
        else
            # Some Tables.jl sources expose names via columnnames without
            # implementing schema (e.g. anything relying on the column-access
            # interface alone). Try that before giving up and going positional.
            try
                Tables.columnnames(Tables.columns(data))
            catch
                nothing
            end
        end
        if isnothing(src_names)
            # Unnamed Tables.jl source — fall back to positional matching.
            [collect(Tables.getcolumn(r, i) for i in 1:ncols) for r in Tables.rows(data)]
        else
            src = collect(Symbol.(src_names))
            tbl = Symbol.(t.columns)

            missing_cols = setdiff(tbl, src)
            isempty(missing_cols) || throw(XLSXError(
                "Source is missing column(s) $(join(missing_cols, ", ")) required by table `$name`. " *
                "Table columns are: $(join(t.columns, ", "))."))

            extra_cols = setdiff(src, tbl)
            isempty(extra_cols) || throw(XLSXError(
                "Source has column(s) $(join(extra_cols, ", ")) that table `$name` does not have. " *
                "Table columns are: $(join(t.columns, ", "))."))

            # Match by name, emitting values in the table's column order.
            [collect(Tables.getcolumn(r, nm) for nm in tbl) for r in Tables.rows(data)]
        end
    end

    for (ri, r) in enumerate(rows)
        length(r) == ncols || throw(XLSXError(
            "Row $ri has $(length(r)) values but table `$name` has $ncols columns."))
    end

    return rows
end
