
# Tables.jl interface

Tables.istable(::Type{<:TableRowIterator}) = true
Tables.rowaccess(::Type{<:TableRowIterator}) = true
Tables.rows(itr::TableRowIterator) = itr
Tables.schema(itr::TableRowIterator) = Tables.Schema(itr.index.column_labels, fill(Any, length(itr.index.column_labels)))
Tables.columnnames(tr::TableRow) = tr.index.column_labels
Tables.getcolumn(tr::TableRow, nm::Symbol) = getdata(tr, nm)
Tables.getcolumn(tr::TableRow, i::Integer) = getdata(tr, i)

_as_vector(y::AbstractVector) = y
_as_vector(y) = collect(y)

function _table_to_arrays(x)
    if Tables.istable(x)
            columns = Any[_as_vector(c) for c in Tables.Columns(x)]
            colnames = collect(Symbol, Tables.columnnames(x))
            return columns, colnames
    else
        throw(XLSXError("$(typeof(x)) does not implement Tables.jl interface."))
    end
end

"""
    writetable(filename, table; [overwrite], [sheetname])

Write a Tables.jl compatible `table` as an Excel file with the specified file name (and sheet name, if specified).

If a file with the given name already exists, writing will fail unless `overwrite=true` is specified, in which 
case the existing file will be overwritten.
"""
writetable(filename::Union{AbstractString, IO}, x; kw...) = writetable(filename, _table_to_arrays(x)...; kw...)

"""
    writetable(filename::Union{AbstractString, IO}, tables::Vector{Pair{String, T}}; overwrite::Bool=false)
    writetable(filename::Union{AbstractString, IO}, tables::Pair{String, Any}...; overwrite::Bool=false)
"""
function writetable(filename::Union{AbstractString, IO}, tables::Vector{<:Pair}; kw...)
    data = [(name, _table_to_arrays(x)...) for (name, x) in tables]
    return writetable(filename, data; kw...)
end

writetable(filename::Union{AbstractString, IO}, tables::Pair{<:String, <:Any}...; kw...) = writetable(filename, collect(tables); kw...)

"""
    writetable!(sheet::Worksheet, table; anchor_cell::CellRef=CellRef("A1")))

Write a Tables.jl compatible `table` to the specified sheet starting with the 
anchor cell (if given) in the top left.
"""
writetable!(sheet::Worksheet, x; kw...) = writetable!(sheet, _table_to_arrays(x)...; kw...)

#
# DataTable
#

Tables.istable(::Type{DataTable}) = true
Tables.columnaccess(::Type{DataTable}) = true
Tables.columns(dt::DataTable) = dt # DataTable implements Tables.AbstractColumns interface
Tables.schema(dt::DataTable) = Tables.Schema(dt.column_labels, nothing)
Tables.columnnames(dt::DataTable) = dt.column_labels
Tables.getcolumn(dt::DataTable, i::Int64) = dt.data[i]

function Tables.getcolumn(dt::DataTable, column_label::Symbol)
    if !haskey(dt.column_label_index, column_label)
        throw(XLSXError("Column `$column_label` not found."))
    end

    column_index = dt.column_label_index[column_label]
    return Tables.getcolumn(dt, column_index)
end

#
# ====================================================================================== Excel Tables
#

# Data rows are between the header row and (if present) the totals row —
# never include either.
_first_data_row(t::Table) = t.ref.start.row_number + 1
_last_data_row(t::Table)  = t.ref.stop.row_number - (t.has_totals_row ? 1 : 0)
_col_start(t::Table) = column_number(t.ref.start)

# Tables.jl interface — mirrors the existing TableRowIterator/TableRow pattern

Tables.istable(::Type{<:Table}) = true
Tables.istable(::Type{<:XLSXTableRowIterator}) = true
Tables.rowaccess(::Type{<:Table}) = true
Tables.rowaccess(::Type{<:XLSXTableRowIterator}) = true
Tables.columnaccess(::Type{<:Table}) = true

Tables.rows(t::Table) = XLSXTableRowIterator(t)
Tables.rows(it::XLSXTableRowIterator) = it  # identity, matching the existing TableRowIterator convention

function Tables.rowtable(t::Table)
    names = Tuple(Symbol.(t.columns))
    return [NamedTuple{names}(ntuple(i -> Tables.getcolumn(row, i), length(names))) for row in eachtablerow(t)]
end

Tables.rowtable(it::XLSXTableRowIterator) = Tables.rowtable(it.table)

Tables.schema(t::Table) = Tables.Schema(Symbol.(t.columns), nothing)
Tables.schema(it::XLSXTableRowIterator) = Tables.schema(it.table)

Tables.columnnames(t::Table) = Symbol.(t.columns)
Tables.columnnames(tr::XLSXTableRow) = Symbol.(tr.table.columns)

Tables.getcolumn(tr::XLSXTableRow, nm::Symbol) =
    getdata(tr.table.sheet, CellRef(tr.row_number, _col_start(tr.table) + findfirst(==(nm), Symbol.(tr.table.columns)) - 1))
Tables.getcolumn(tr::XLSXTableRow, i::Int64) =
    getdata(tr.table.sheet, CellRef(tr.row_number, _col_start(tr.table) + i - 1))

Base.eltype(::Type{XLSXTableRowIterator}) = XLSXTableRow
Base.length(it::XLSXTableRowIterator) = max(0, _last_data_row(it.table) - _first_data_row(it.table) + 1)
Base.IteratorSize(::Type{XLSXTableRowIterator}) = Base.HasLength()

function Base.iterate(it::XLSXTableRowIterator, state::Int = _first_data_row(it.table))
    state > _last_data_row(it.table) && return nothing
    return XLSXTableRow(it.table, state), state + 1
end
function Tables.columns(t::Table)
    row_range = _first_data_row(t):_last_data_row(t)
    col0 = _col_start(t)
    NamedTuple(
        Symbol(name) => typed_column(Any[getdata(t.sheet, CellRef(r, col0 + i - 1)) for r in row_range])
        for (i, name) in enumerate(t.columns)
    )
end

# Forwarding is sufficient — no need to also declare
# Tables.columnaccess(::Type{<:XLSXTableRowIterator}) = true; simply having
# this method defined is enough for consumers to pick it up (confirmed
# empirically for the equivalent case on the old TableRowIterator).
Tables.columns(it::XLSXTableRowIterator) = Tables.columns(it.table)

"""
    eachtablerow(t::Table) -> XLSXTableRowIterator

Iterate over the data rows of an Excel Table `t` (as returned by
[`XLSX.table`](@ref)). Each element is an `XLSXTableRow`.

!!! note "Two `eachtablerow` methods"

    Different from [`XLSX.eachtablerow(sheet, ...)`](@ref), which infers a
    table's bounds from cell content. Here, `t.ref` is authoritative: the
    header and totals row (if any) are always excluded, and any blank row
    within `t.ref` is still returned as ordinary data — there's no
    `stop_in_empty_row`/`keep_empty_rows` equivalent.

Cell values are read on demand via [`XLSX.getdata`](@ref), through the
worksheet's normal cell cache — no separate caching, so edits made before
iterating are reflected normally.

Rows conform to `Tables.jl` (`Tables.getcolumn` by name or position), and
`t` itself is directly `Tables.jl`-compatible too (`DataFrame(t)` works
without `eachtablerow`); use this for row-by-row iteration instead.

# Example
```julia
for r in XLSX.eachtablerow(t)
    v1 = Tables.getcolumn(r, 1)
    v2 = Tables.getcolumn(r, :revenue)
end
```

```julia
julia> using DataFrames

julia> DataFrame(XLSX.eachtablerow(t))  # equivalent to DataFrame(t)

julia> collect(XLSX.eachtablerow(t))    # Vector{XLSX.XLSXTableRow}
```

See also [`XLSX.table`](@ref), [`XLSX.tables`](@ref), [`XLSX.gettable`](@ref).
"""
eachtablerow(t::Table) = XLSXTableRowIterator(t)