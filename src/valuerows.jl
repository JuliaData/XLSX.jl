#=
Value rows: the row source behind `readtable` (#462).

`readtable` owns the `XLSXFile` it opens and drops it on return, so filling the
worksheet cache (a `Cell` per cell, a `Dict` per row) only to read each value back out
is wasted work. Instead, `readtable` reads the target sheet's XML in one cursor pass
into `ValueRow`s: per row, each cell in the requested columns decoded to its type and raw
value, as a `Cell` holds them, and nothing else.

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
switch both ways (test/test_files/files/differential.jl), and compare `ValueRow`s with
`SheetRow`s row by row, so the two paths can't drift apart unnoticed.

The rows are read with an `XML.Cursor` here, or by default with the byte scanner in
src/valuescan.jl, which makes the same `ValueRow`s faster and hands anything unusual back
to XML.jl (`ValueRowIterator`'s `scan`).

`ValueRowIterator` is internal to `readtable`: its state is a mutable cursor, so a
state can be advanced only once (as `TableRowIterator` does), and each `iterate(itr)`
starts a fresh pass over the sheet XML.
=#

# Marks a column with no `<c>` in the row, as distinct from a present, empty cell.
struct _AbsentCell end
const _ABSENT = _AbsentCell()
# The same in a row's `kinds`, where a present cell has its `CellValueType`.
const _KIND_ABSENT = 0xff

# The formulas argument of `_cursor_cell_contents` for a read that doesn't load formulas: never written.
const _NO_FORMULAS = Dict{SheetCellRef,AbstractFormula}()

# Thrown when `<row r>` values aren't ascending; `readtable` then uses the cache path.
struct _RowsNotAscending <: Exception end

# Each cell is kept as a `Cell` keeps it, its `CellValueType` and raw `UInt64`, and made
# a Julia value only when read: so no value is boxed, and `gettable` pushes numbers
# straight into typed columns (`_push_row_cell!`).
struct ValueRow <: AbstractSheetRow
    ws::Worksheet
    row::Int
    firstcol::Int              # column of kinds[1] and raws[1]
    kinds::Vector{UInt8}       # each cell's `CellValueType` by column from `firstcol`; `_KIND_ABSENT` where no `<c>`
    raws::Vector{UInt64}       # each cell's raw value, as in a `Cell`
    has_cells::Bool            # any `<c>` at all, in any column (`!isempty` of a `SheetRow`)
end

struct ValueRowIterator <: SheetRowIterator
    ws::Worksheet
    xml::String                          # the worksheet's whole XML
    cols::Union{Nothing,UnitRange{Int}}  # columns to decode; `nothing` = every column
    sst_pfx::String
    last_state::Base.RefValue{Any}       # the state of the latest `iterate`, for `_check_rows_ascending!`
    scan::Bool                           # read with the byte scanner (src/valuescan.jl), not a cursor
end

ValueRowIterator(ws::Worksheet, xml::String, cols::Union{Nothing,UnitRange{Int}}, scan::Bool=_READTABLE_SCAN[]) =
    ValueRowIterator(ws, xml, cols, get_sst_prefix(ws), Ref{Any}(nothing), scan)

mutable struct _ValueRowState{C}
    cursor::C
    at_row::Bool       # the cursor is on a `<row>` not yet read
    last_row::Int
    done::Bool         # the end of the sheet was reached
    width::Int         # cells in the last row read, to size the next (all-columns mode)
end

Base.IteratorSize(::Type{ValueRowIterator}) = Base.SizeUnknown()
Base.eltype(::Type{ValueRowIterator}) = ValueRow

@inline get_worksheet(itr::ValueRowIterator) = itr.ws
@inline get_worksheet(r::ValueRow) = r.ws
@inline row_number(r::ValueRow) = r.row
Base.isempty(r::ValueRow) = !r.has_cells

# The index in `kinds` and `raws` of a present cell at `column`, or 0.
@inline function _cell_index(r::ValueRow, column::Int)
    i = column - r.firstcol + 1
    return 1 <= i <= length(r.kinds) && r.kinds[i] != _KIND_ABSENT ? i : 0
end

# The value at `column`, or `_ABSENT` when it has no `<c>` or is outside the decoded columns.
function _value_or_absent(r::ValueRow, column::Int)
    i = _cell_index(r, column)
    return i == 0 ? _ABSENT : _cell_value(r.ws, CellValueType(r.kinds[i]), r.raws[i])
end
function getdata(r::ValueRow, column::Int)
    v = _value_or_absent(r, column)
    return v === _ABSENT ? missing : v
end

# `gettable`'s push of the value at `column`, `missing_strings` applied, to the typed
# column `col`; returns the column it went into (see `_push_column!`). Each kind of value
# reaches `_push_column!` as its concrete type, so none is boxed.
function _push_row_cell!(col::Vector, ::Worksheet, r::ValueRow, column::Int, ms::Set{String})
    i = _cell_index(r, column)
    i == 0 && return _push_column!(col, missing)
    return @inline _with_cell_value(x -> _push_column!(col, _apply_missing_strings(x, ms)),
                            r.ws, CellValueType(r.kinds[i]), r.raws[i])
end

# The same iterator, decoding only `cr`'s columns.
_narrow(itr::ValueRowIterator, cr::ColumnRange) =
    ValueRowIterator(itr.ws, itr.xml, cr.start:cr.stop, itr.sst_pfx, Ref{Any}(nothing), itr.scan)

# A state at the start of `<sheetData>`; finds it as `first_cache_fill!` does, with the
# same error if there is none.
function _value_row_start(itr::ValueRowIterator)
    sheetdata_lazy = _find_sheetdata(parse(itr.xml, XML.LazyNode), itr.ws.name)
    return _ValueRowState(XML.Cursor(sheetdata_lazy), false, 0, false, 0)
end

Base.iterate(itr::ValueRowIterator) =
    itr.scan ? iterate(itr, _scan_start(itr)) : iterate(itr, _value_row_start(itr))

# Check the whole sheet's rows are ascending before reading any (see `_check_rows_ascending!`).
function _check_all_rows_ascending(itr::ValueRowIterator)
    itr.last_state[] = itr.scan ? _scan_start(itr) : _value_row_start(itr)
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
    row = _row_number(c, ws.name)
    row <= st.last_row && throw(_RowsNotAscending())
    st.last_row = row

    cols = itr.cols
    firstcol, kinds, raws = _new_row_cells(cols, st.width)
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
                attrs = _cell_attributes(c)
                col = column_number(first(attrs))
                if isnothing(cols) || col in cols
                    # `<f>` is ignored: readtable reads values only
                    datatype, value, _ = _cursor_cell_contents(c, attrs, ws, wb, itr.sst_pfx, _NO_FORMULAS, false)
                    _set_cell!(kinds, raws, firstcol, col, datatype, value)
                else
                    XML.skip_element!(c)
                end
            else
                XML.skip_element!(c)
            end
        end
    end

    st.width = length(kinds)
    return ValueRow(ws, row, firstcol, kinds, raws, has_cells), st
end

# The first column, kinds and raws of a new row: the requested columns, all absent, or
# (all-columns mode) empty from column 1 with room for `width` cells.
function _new_row_cells(cols::Union{Nothing,UnitRange{Int}}, width::Int)
    isnothing(cols) && return 1, sizehint!(UInt8[], width), sizehint!(UInt64[], width)
    return first(cols), fill(_KIND_ABSENT, length(cols)), zeros(UInt64, length(cols))
end

# Store the cell `(dt, v)` for `col`, growing the row (all-columns mode starts empty, from
# column 1).
@inline function _set_cell!(kinds::Vector{UInt8}, raws::Vector{UInt64}, firstcol::Int, col::Int,
                            dt::CellValueType, v::UInt64)
    i = col - firstcol + 1
    if i > length(kinds)
        n = length(kinds)
        resize!(kinds, i)
        resize!(raws, i)
        for k in n+1:i
            kinds[k] = _KIND_ABSENT
        end
    end
    kinds[i] = UInt8(dt)
    raws[i] = v
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
_check_rows_ascending!(itr::ValueRowIterator) = _check_rows_ascending!(itr, itr.last_state[])
_check_rows_ascending!(::ValueRowIterator, ::Nothing) = nothing
function _check_rows_ascending!(::ValueRowIterator, st::_ValueRowState)
    st.done && return nothing
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
        r_val = _cursor_attr(c, "r")
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

# `readtable` (without `table_name`) through value rows. Falls back to the cache path
# for a chartsheet, a sheet not held as XML text, or rows out of ascending order;
# results are the same as `gettable` on the cache except that invalid cells outside the
# returned area aren't decoded, so they raise no error (see src/valuerows.jl). The rows
# are read by the byte scanner, or by a cursor where it hands the sheet back (src/valuescan.jl).
function _readtable_value_rows(source, sheet, cols; infer_eltypes::Bool, kw...)
    xf = open_or_read_xlsx(source, true, true, false; target_sheet=sheet, load_formulas=false,
                           split_sheets=false)
    ws = getsheet(xf, sheet)
    wb = get_workbook(xf)
    cache_path() = isnothing(cols) ? gettable(ws; infer_eltypes, kw...) : gettable(ws, cols; infer_eltypes, kw...)
    is_chartsheet(wb, ws.name) && return cache_path()
    raw = get(xf.data, get_relationship_target_by_id("xl", wb, ws.relationship_id), nothing)
    raw isa String || return cache_path()
    # The byte scanner first (src/valuescan.jl), then the cursor for markup it hands back.
    for scan in (_READTABLE_SCAN[] ? (true, false) : (false,))
        try
            return _readtable_value_rows(ws, raw, cols, scan; infer_eltypes, kw...)
        catch e
            e isa _ScanFallback && continue
            e isa _RowsNotAscending && break
            rethrow()
        end
    end
    return cache_path()
end

# `readtable` of `ws`, held as `raw` XML, through value rows read by the byte scanner
# (`scan`) or a cursor.
function _readtable_value_rows(ws::Worksheet, raw::String, cols, scan::Bool; infer_eltypes::Bool, kw...)
    # A `stop_in_row_function` sees rows as they're read, so check the row order (and
    # that the scanner can step over every row) first: a fallback after it had been
    # called would call it on rows twice.
    get(kw, :stop_in_row_function, nothing) === nothing ||
        _check_all_rows_ascending(ValueRowIterator(ws, raw, nothing, scan))
    itr = if isnothing(cols)
        _eachtablerow(ws, ValueRowIterator(ws, raw, nothing, scan); kw...)
    else
        cr = convert(ColumnRange, cols)
        _eachtablerow(ws, ValueRowIterator(ws, raw, cr.start:cr.stop, scan), cr; kw...)
    end
    t = gettable(itr; infer_eltypes)
    _check_rows_ascending!(itr.itr)    # the rows after the table, too
    return t
end
