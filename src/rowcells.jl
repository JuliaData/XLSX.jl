#
# A row's cells, as the cache and `SheetRow` hold them: a `Vector{Cell}` in column
# order, at most one cell to a column. Only these functions know how a row is stored;
# everything else reads and changes rows through them.
#
# `_rowcolumns` and `_rowcells` give a row's columns and cells, in column order.
# `_sorted_rowcolumns` gives the columns as a new vector, for code that keeps them.
#

# An empty row, with room for `n` cells.
_newrow() = RowCells()
_newrow(n::Integer) = sizehint!(RowCells(), n)

# Where column `col` is in `row`: its index if it has a cell, or else minus the index
# a cell for it would go at.
@inline function _rowindex(row::RowCells, col::Int)::Int
    lo, hi = 1, length(row)
    while lo <= hi
        mid = (lo + hi) >>> 1
        c = column_number(@inbounds row[mid])
        if c < col
            lo = mid + 1
        elseif c > col
            hi = mid - 1
        else
            return mid
        end
    end
    return -lo
end

# The cell in column `col`, or `nothing`.
@inline function _rowcell(row::RowCells, col::Integer)
    i = _rowindex(row, Int(col))
    return i > 0 ? @inbounds(row[i]) : nothing
end

@inline _hasrowcell(row::RowCells, col::Integer) = _rowindex(row, Int(col)) > 0

# Store `cell` in its column, replacing any cell there. Cells are nearly always added
# left to right, so a cell after the last is just pushed.
@inline function _setrowcell!(row::RowCells, cell::Cell)
    col = column_number(cell)
    if isempty(row) || column_number(@inbounds row[end]) < col
        push!(row, cell)
    else
        i = _rowindex(row, col)
        i > 0 ? (@inbounds row[i] = cell) : insert!(row, -i, cell)
    end
    return nothing
end

@inline function _deleterowcell!(row::RowCells, col::Integer)
    i = _rowindex(row, Int(col))
    i > 0 && deleteat!(row, i)
    return nothing
end

@inline _rowcolumns(row::RowCells) = (column_number(c) for c in row)
@inline _rowcells(row::RowCells) = row
_sorted_rowcolumns(row::RowCells) = Int[column_number(c) for c in row]

# Make room for `n` more cells in `row`.
_reserve!(row::RowCells, n::Integer) = (sizehint!(row, length(row) + n); nothing)

# Replace each cell of `row` with `f(cell)`, which must keep the cell's column.
_maprow!(f, row::RowCells) = (map!(f, row, row); nothing)

#
# `Cell` is immutable: a cell is changed by storing a changed copy in its place.
#

# A copy of `c` with the fields given changed.
@inline _withfields(c::Cell; ref::CellRef=c.ref, value=c.value, style=c.style, meta=c.meta,
                    datatype::CellValueType=c.datatype, formula::Bool=c.formula) =
    Cell(ref, UInt64(value), UInt32(style), UInt16(meta), datatype, formula)

# Store `cell`, with the fields given changed, in its place in the sheet's cache, and
# return the new cell. `cell` must be the cell the cache holds at its address: where
# anything since it was read may have changed the cell, give its address instead.
function _updatecell!(ws::Worksheet, cell::Cell; kw...)::Cell
    new = _withfields(cell; kw...)
    _setrowcell!(ws.cache.cells[row_number(new)], new)
    return new
end

# The cell the cache holds at `ref`, with the fields given changed, or `nothing` if
# there is no cell there.
function _updatecell!(ws::Worksheet, ref::CellRef; kw...)::Union{Nothing,Cell}
    cell = getcell(ws, ref)
    return cell isa Cell ? _updatecell!(ws, cell; kw...) : nothing
end
_updatecell!(xf::XLSXFile, ref::SheetCellRef; kw...) = _updatecell!(getsheet(xf, ref.sheet), ref.cellref; kw...)
