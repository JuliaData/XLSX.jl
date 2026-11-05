#=
Value rows: the row source behind `readtable` (#462).

`readtable` owns the `XLSXFile` it opens and drops it on return, so filling the
worksheet cache (a `Cell` per cell, a `Dict` per row) only to read each value back out
is wasted work. Instead, `readtable` reads the target sheet's XML in one cursor pass
into `ValueRow`s: per row, the decoded value of each cell in the requested columns, and
nothing else.

Design — one row source, one table reader:
  - Everything about the *table* stays in the shared code in table.jl: header and
    column-range detection, `first_row`, gaps, `stop_in_empty_row`, `keep_empty_rows`,
    `missing_strings`, column labels, `normalizenames`, `stop_in_row_function` and
    type inference. That code reads rows through a few generic accessors
    (`row_number`, `isempty`, `_row_value`, `_present_columns`, `_push_header_label!`)
    with one method for `SheetRow` (the cache) and one for `ValueRow` (here).
  - Cells are decoded by `_cursor_cell_contents`, as in the cache fill, with the same rules: an
    inline string's first `<is>` only, the last `<v>` wins; `<f>` is ignored here.
  - A `ValueRow` keeps an absent cell (no `<c>`) apart from a present but empty one
    (`missing`), as the cache does: header labels differ ("#Empty" vs "missing"), and
    `isempty(row)` counts any `<c>` in any column.

Differences from the cache path, by design:
  - Only the requested columns are decoded, and reading stops where the table does
    (after reading the one row that shows where it ends). So an invalid cell outside
    the area `readtable` reads (another column, or a row after that) no longer raises
    an error; one inside it raises the same error as before.
  - `enable_cache` has no effect on `readtable`.
  - Rows must be in ascending order, as the spec requires and Excel writes. If they
    aren't, `_RowsNotAscending` is thrown and `readtable` falls back to the cache path,
    which sorts and merges them. Rows the table read doesn't reach are checked too, by
    `_check_rows_ascending!`, which reads only their `r` numbers.

Reversing it: set `_READTABLE_VALUE_ROWS[] = false` (src/read.jl) and `readtable` goes
through the cache path again; or delete this file, its include, and the branches on
`_READTABLE_VALUE_ROWS` in `readtable`. The tests run every differential check with the
switch both ways (test/test_files/Differential_tests.jl), and compare `ValueRow`s with
`SheetRow`s row by row, so the two paths can't drift apart unnoticed.

`ValueRowIterator` is internal to `readtable`: its state is a mutable cursor, so a
state can be advanced only once (as `TableRowIterator` does), and each `iterate(itr)`
starts a fresh pass over the sheet XML.
=#

# Marks a column with no `<c>` in the row, as distinct from a present, empty cell.
struct _AbsentCell end
const _ABSENT = _AbsentCell()

# The formulas argument of `_cursor_cell_contents` for a read that doesn't load formulas: never written.
const _NO_FORMULAS = Dict{SheetCellRef,AbstractFormula}()

# Thrown when `<row r>` values aren't ascending; `readtable` then uses the cache path.
struct _RowsNotAscending <: Exception end

struct ValueRow <: AbstractSheetRow
    ws::Worksheet
    row::Int
    firstcol::Int              # column of values[1]
    values::Vector{Any}        # decoded values by column from `firstcol`; `_ABSENT` where no `<c>`
    has_cells::Bool            # any `<c>` at all, in any column (`!isempty` of a `SheetRow`)
end

struct ValueRowIterator <: SheetRowIterator
    ws::Worksheet
    xml::String                          # the worksheet's whole XML
    cols::Union{Nothing,UnitRange{Int}}  # columns to decode; `nothing` = every column
    sst_pfx::String
    last_state::Base.RefValue{Any}       # the state of the latest `iterate`, for `_check_rows_ascending!`
end

ValueRowIterator(ws::Worksheet, xml::String, cols::Union{Nothing,UnitRange{Int}}) =
    ValueRowIterator(ws, xml, cols, get_sst_prefix(ws), Ref{Any}(nothing))

mutable struct _ValueRowState{C}
    cursor::C
    at_row::Bool       # the cursor is on a `<row>` not yet read
    last_row::Int
    done::Bool         # the end of the sheet was reached
end

Base.IteratorSize(::Type{ValueRowIterator}) = Base.SizeUnknown()
Base.eltype(::Type{ValueRowIterator}) = ValueRow

@inline get_worksheet(itr::ValueRowIterator) = itr.ws
@inline get_worksheet(r::ValueRow) = r.ws
@inline row_number(r::ValueRow) = r.row
Base.isempty(r::ValueRow) = !r.has_cells

# The value at `column`: `missing` when absent, empty, or outside the decoded columns.
@inline function _value_or_absent(r::ValueRow, column::Int)
    i = column - r.firstcol + 1
    return 1 <= i <= length(r.values) ? r.values[i] : _ABSENT
end
function getdata(r::ValueRow, column::Int)
    v = _value_or_absent(r, column)
    return v === _ABSENT ? missing : v
end

# The same iterator, decoding only `cr`'s columns.
_narrow(itr::ValueRowIterator, cr::ColumnRange) = ValueRowIterator(itr.ws, itr.xml, cr.start:cr.stop, itr.sst_pfx, Ref{Any}(nothing))

# A state at the start of `<sheetData>`; finds it as `first_cache_fill!` does, with the
# same error if there is none.
function _value_row_start(itr::ValueRowIterator)
    sheetdata_lazy = _find_sheetdata(parse(itr.xml, XML.LazyNode), itr.ws.name)
    return _ValueRowState(XML.Cursor(sheetdata_lazy), false, 0, false)
end

Base.iterate(itr::ValueRowIterator) = iterate(itr, _value_row_start(itr))

# Check the whole sheet's rows are ascending before reading any (see `_check_rows_ascending!`).
function _check_all_rows_ascending(itr::ValueRowIterator)
    itr.last_state[] = _value_row_start(itr)
    _check_rows_ascending!(itr)
    return nothing
end

function Base.iterate(itr::ValueRowIterator, st::_ValueRowState)
    itr.last_state[] = st
    c = st.cursor
    if !st.at_row                        # find the next `<row>`
        found = false
        while XML.next!(c) !== nothing
            XML.depth(c) == 2 && XML.nodetype(c) == XML.Element || continue
            localname(c) == "row" && (found = true; break)
            XML.skip_element!(c)
        end
        found || (st.done = true; return nothing)
    end
    st.at_row = false

    ws = itr.ws
    wb = get_workbook(ws)
    r_val = XML.get(c, "r", nothing)
    r_val === nothing && throw(XLSXError("Row without 'r' attribute in worksheet $(ws.name)."))
    row = parse(Int, r_val)
    row <= st.last_row && throw(_RowsNotAscending())
    st.last_row = row

    cols = itr.cols
    firstcol = isnothing(cols) ? 1 : first(cols)
    values = isnothing(cols) ? Any[] : Any[_ABSENT for _ in cols]
    has_cells = false

    while XML.next!(c) !== nothing
        d  = XML.depth(c)
        nt = XML.nodetype(c)

        if d <= 2
            if d == 2 && nt == XML.Element
                if localname(c) == "row"
                    st.at_row = true     # the next row: read on the next `iterate`
                    break
                end
                XML.skip_element!(c)
            end
        elseif d == 3
            nt == XML.Element || continue
            if localname(c) == "c"
                has_cells = true
                attrs = _cell_attributes(XML.LazyNode(c))
                col = column_number(first(attrs))
                if isnothing(cols) || col in cols
                    # `<f>` is ignored: readtable reads values only
                    datatype, value, _ = _cursor_cell_contents(c, attrs, ws, wb, itr.sst_pfx, _NO_FORMULAS, false)
                    _set_value!(values, firstcol, col, _cell_value(ws, datatype, value))
                else
                    XML.skip_element!(c)
                end
            else
                XML.skip_element!(c)
            end
        end
    end

    return ValueRow(ws, row, firstcol, values, has_cells), st
end

# Store `v` for `col`, growing `values` (all-columns mode starts empty, from column 1).
@inline function _set_value!(values::Vector{Any}, firstcol::Int, col::Int, v)
    i = col - firstcol + 1
    if i > length(values)
        n = length(values)
        resize!(values, i)
        for k in n+1:i
            values[k] = _ABSENT
        end
    end
    values[i] = v
    return nothing
end

# The row numbered `row`, read from the start; the same error as `find_row` if absent.
function _find_value_row(itr::ValueRowIterator, row::Int)::ValueRow
    for r in itr
        row_number(r) == row && return r
        row_number(r) > row && break
    end
    _check_rows_ascending!(itr)          # a later, out-of-order row could be the one
    throw(XLSXError("Row $row not found in worksheet $(itr.ws.name)."))
end

# A table read can stop before the end of the sheet (an empty row, `stop_in_row_function`).
# The rows it didn't reach must still be in ascending order, since the cache path sorts
# them: check their `r` numbers, skipping each row's contents undecoded. Throws
# `_RowsNotAscending` (so `readtable` falls back to the cache) if they aren't. Rows
# without a readable `r` aren't validated here: they're outside the read area.
function _check_rows_ascending!(itr::ValueRowIterator)
    st = itr.last_state[]
    (st === nothing || st.done) && return nothing
    c = st.cursor
    last = st.last_row
    pending = st.at_row                  # the cursor is already on an unread `<row>`
    while true
        if !pending
            XML.next!(c) === nothing && break
            (XML.depth(c) == 2 && XML.nodetype(c) == XML.Element) || continue
            if localname(c) != "row"
                XML.skip_element!(c)
                continue
            end
        end
        pending = false
        r_val = XML.get(c, "r", nothing)
        r = isnothing(r_val) ? nothing : tryparse(Int, r_val)
        if !isnothing(r)
            r <= last && throw(_RowsNotAscending())
            last = r
        end
        XML.skip_element!(c)
    end
    st.done = true
    return nothing
end
