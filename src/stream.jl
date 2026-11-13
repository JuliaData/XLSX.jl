
#=
https://docs.julialang.org/en/v1/base/collections/#lib-collections-iteration-1

for i in iter   # or  "for i = iter"
    # body
end

is translated into:

next = iterate(iter)
while next != nothing
    (i, state) = next
    # body
    next = iterate(iter, state)
end
=#

#=
# About Iterators

* `SheetRowIterator` is an abstract iterator that has `SheetRow` as its elements. `SheetRowStreamIterator` and `WorksheetCache` implements `SheetRowIterator` interface.
* `SheetRowStreamIterator` is a dumb iterator for row elements in sheetData XML tag of a worksheet. Empty rows are not represented in the XML file so cannot be seen by the iterator.
* `WorksheetCache` has a `SheetRowStreamIterator` and caches all values read from the stream.
* `TableRowIterator` is a smart iterator that looks for tabular data, but uses a SheetRowIterator under the hood.

The implementation of `SheetRowIterator` will be chosen automatically by `eachrow` method,
based on the `enable_cache` option used in `XLSX.openxlsx` method.

=#

#=
# SheetRowIterator

It's state is the SheetRowStreamIteratorState.
The iterator element is a SheetRow.
=#

@inline get_worksheet(itr::SheetRowIterator) = itr.sheet
@inline row_number(state::SheetRowStreamIteratorState) = state.row

# Opens a file for streaming.
@inline function open_internal_file_stream(xf::XLSXFile, filename::String) :: XML.LazyNode

    !internal_xml_file_exists(xf, filename) && throw(XLSXError("Couldn't find $filename in $(xf.source)."))
    if xf.source isa IO
        seekstart(xf.source)
        zip_io = ZipArchives.ZipReader(read(xf.source))
    else
        zip_io = ZipArchives.ZipReader(FileArray(abspath(xf.source)))
    end

    return parse(strip_bom_and_lf(_readentry_string(zip_io, filename)), XML.LazyNode)

end


# Open the sheet's XML stream and return its <sheetData> node.
# Shared by the stream iterator and its tests so both navigate identically.
function _open_sheetdata(ws::Worksheet)
    target_file = get_relationship_target_by_id("xl", get_workbook(ws), ws.relationship_id)
    xf = get_xlsxfile(ws)
    doc = open_internal_file_stream(xf, target_file)
    return _find_sheetdata(doc, ws.name)
end

# Moves a cursor in `<sheetData>` to its next `<row>`; `false` at the end of the sheet.
function _next_row!(c::XML.Cursor)::Bool
    while XML.next!(c) !== nothing
        (XML.depth(c) == 2 && XML.nodetype(c) == XML.Element) || continue
        localname(c) == "row" && return true
        XML.skip_element!(c)
    end
    return false
end

# The value of the attribute `key` of the element the cursor is on, or `nothing`. The
# first read of the element's attributes is where the cursor stands, so its next
# `next!` doesn't read them again.
@inline function _cursor_attr(c::XML.Cursor, key::String)::Union{Nothing,SubString{String}}
    for (k, v) in XML.eachattribute!(c)
        k == key && return v
    end
    return nothing
end

# The `r` of the `<row>` the cursor is on.
function _row_number(c::XML.Cursor, wsname::String)::Int
    r = _cursor_attr(c, "r")
    r === nothing && throw(XLSXError("Row without 'r' attribute in worksheet $wsname."))
    return parse(Int, r)
end

# The `ht` of the `<row>` the cursor is on.
function _row_height(c::XML.Cursor)::Union{Nothing,Float64}
    ht = _cursor_attr(c, "ht")
    return isnothing(ht) ? nothing : parse(Float64, ht)
end

# Decodes the `<c>`s in `cols` (`nothing` = all) of the `<row>` the cursor is on into
# `rowcells`, by column (a repeated column: the last one), each tokenised once; the
# others are skipped after reading only their `r`. Leaves the cursor holding the node
# after the row. Returns the number of shared-string cells decoded and the first and
# last column of any `<c>` in the row (`typemax(Int)`, 0 for none).
function _cursor_rowcells!(rowcells::Dict{Int,Cell}, c::XML.Cursor, ws::Worksheet, wb::Workbook, sst_pfx::String,
                           local_formulas::Dict{SheetCellRef,AbstractFormula}, load_formulas::Bool,
                           cols::Union{Nothing,UnitRange{Int}} = nothing)::Tuple{Int,Int,Int}
    sst_count = 0
    col_min, col_max = typemax(Int), 0
    XML.@for_each_child! c child begin
        if XML.nodetype(child) == XML.Element && localname(child) == "c"
            col = isnothing(cols) ? 0 : _cell_column(child)
            if isnothing(cols) || col in cols
                cell = _cursor_cell(child, ws, wb, sst_pfx, local_formulas, load_formulas)
                col = column_number(cell)
                sst_count += cell.datatype == CT_STRING ? 1 : 0
                rowcells[col] = cell
            else
                XML.skip_element!(child)
            end
            col_min, col_max = min(col_min, col), max(col_max, col)
        else
            XML.skip_element!(child)
        end
    end
    return sst_count, col_min, col_max
end

# The column of the `<c>` the cursor is on, from its `r` attribute alone.
function _cell_column(c::XML.Cursor)::Int
    r = _cursor_attr(c, "r")
    r === nothing && throw(XLSXError("Invalid cell reference ``."))
    return _ref_column_number(r)
end

# Creates an iterator for row elements in the Worksheet's XML: one cursor walks
# `<sheetData>`, decoding each row's cells as it passes them.
function Base.iterate(itr::SheetRowStreamIterator)
    ws = get_worksheet(itr)
    state = SheetRowStreamIteratorState(XML.Cursor(_open_sheetdata(ws)), Dict{Int,Cell}(),
                                        Dict{SheetCellRef,AbstractFormula}(),
                                        isnothing(ws.dimension), typemax(Int), 0, typemax(Int), 0)
    return iterate(itr, state)
end

@inline function _track_row_bounds!(state::SheetRowStreamIteratorState, row::Int, col_min::Int, col_max::Int)
    state.row_min = min(state.row_min, row)
    state.row_max = max(state.row_max, row)
    state.col_min = min(state.col_min, col_min)
    state.col_max = max(state.col_max, col_max)
    nothing
end

function Base.iterate(itr::SheetRowStreamIterator, state::SheetRowStreamIteratorState)
    ws = get_worksheet(itr)
    wb = get_workbook(ws)
    c = state.cursor
    empty!(state.rowcells)

    if !_next_row!(c)
        # A completed pass has seen every row: record the bounds if still unknown.
        if state.track && state.col_max > 0 && isnothing(ws.dimension)
            set_dimension!(ws, CellRange(CellRef(state.row_min, state.col_min), CellRef(state.row_max, state.col_max)))
        end
        return nothing
    end

    current_row, current_row_ht = _row_number(c, ws.name), _row_height(c)
    load_formulas = get_xlsxfile(ws).load_formulas
    sst_count, col_min, col_max = _cursor_rowcells!(state.rowcells, c, ws, wb, get_sst_prefix(ws), state.local_formulas,
                                                    load_formulas, itr.cols)
    itr.sheet.sst_count += sst_count
    state.track && _track_row_bounds!(state, current_row, col_min, col_max)
    # Each row's formulas, so a pass that stops early (a `break`) loses none.
    _merge_local_formulas!(wb, state.local_formulas)

    return SheetRow(ws, current_row, current_row_ht, state.rowcells), state
end

Base.IteratorSize(::Type{<:SheetRowStreamIterator}) = Base.SizeUnknown()
Base.eltype(::Type{<:SheetRowStreamIterator}) = SheetRow

@inline function _merge_local_formulas!(wb::Workbook, local_formulas::Dict{SheetCellRef,AbstractFormula})
    isempty(local_formulas) && return nothing
    lock(wb.formulas_lock) do
        merge!(wb.formulas, local_formulas)
    end
    empty!(local_formulas)
    return nothing
end
 
#
# WorksheetCache
#

# Indicates whether worksheet cache will be fed while reading worksheet cells.
@inline is_cache_enabled(ws::Worksheet) = is_cache_enabled(get_xlsxfile(ws))
@inline is_cache_enabled(wb::Workbook) = is_cache_enabled(get_xlsxfile(wb))
@inline is_cache_enabled(xl::XLSXFile) = xl.use_cache_for_sheet_data
@inline is_cache_enabled(itr::SheetRowIterator) = is_cache_enabled(get_worksheet(itr))

@inline function push_sheetrow!(wc::WorksheetCache, sheet_row::SheetRow)
    r = row_number(sheet_row)
    if !haskey(wc.cells, r)
        # add new row to the cache
        wc.cells[r] = sheet_row.rowcells
        push!(wc.rows_in_cache, r)
        wc.row_index[r] = length(wc.rows_in_cache)
        wc.row_ht[r] = sheet_row.ht
    end
    nothing
end

# `rows_in_cache` is left unsorted while rows are added (sorting on every add is
# costly) and `dirty` is set instead. Call this before anything relies on its order.
function ensure_sorted!(wc::WorksheetCache)
    if wc.dirty
        sort!(wc.rows_in_cache)
        wc.row_index = Dict{Int, Int}(wc.rows_in_cache[i] => i for i in eachindex(wc.rows_in_cache))
        wc.dirty = false
    end
    return wc
end

#
# WorksheetCache iterator
#
# The state is the row number and a flag for if the cache is full or being filled. The element is a SheetRow.
#
function WorksheetCache(ws::Worksheet)
    itr = SheetRowStreamIterator(ws)
    return WorksheetCache(false, CellCache(), Vector{Int}(), Dict{Int, Union{Float64, Nothing}}(), Dict{Int, Int}(), itr, nothing, true)
end

@inline get_worksheet(r::SheetRow) = r.sheet
@inline get_worksheet(itr::WorksheetCache) = get_worksheet(itr.stream_iterator)

# In the WorksheetCache iterator, the element is a SheetRow, the state is the row number and a flag on whether the cache is already full or not
function Base.iterate(ws_cache::WorksheetCache, state::Union{Nothing, WorksheetCacheIteratorState}=nothing)

    isnothing(state) && (state=WorksheetCacheIteratorState(0))

    ensure_sorted!(ws_cache)

    # read from cache
    if state.row_from_last_iteration == 0 && !isempty(ws_cache.rows_in_cache)
        # the next row is in cache, and it's the first one
        current_row_number = ws_cache.rows_in_cache[1]
        current_row_ht = ws_cache.row_ht[current_row_number]
        sheet_row_cells = ws_cache.cells[current_row_number]
        state.row_from_last_iteration=current_row_number
        return SheetRow(get_worksheet(ws_cache), current_row_number, current_row_ht, sheet_row_cells), state

    elseif state.row_from_last_iteration != 0 && ws_cache.row_index[state.row_from_last_iteration] < length(ws_cache.rows_in_cache)
        # the next row is in cache
        current_row_number = ws_cache.rows_in_cache[ws_cache.row_index[state.row_from_last_iteration] + 1]
        current_row_ht = ws_cache.row_ht[current_row_number]
        sheet_row_cells = ws_cache.cells[current_row_number]
        state.row_from_last_iteration=current_row_number
        return SheetRow(get_worksheet(ws_cache), current_row_number, current_row_ht, sheet_row_cells), state

    end
end

function find_row(itr::SheetRowIterator, row::Int) :: SheetRow
    ws=get_worksheet(itr)

    # if cache is in use, look-up row direct rather than iterating
    if !isnothing(ws.cache) && is_cache_enabled(ws)
        if (c = get(ws.cache.cells, row, nothing)) !== nothing
            ht = ws.cache.row_ht[row]
            return SheetRow(ws, row, ht, c)
        end

        throw(XLSXError("Row $row not found in worksheet $(ws.name)."))

    # If can't use cache then lazily iterate sheetrows
    else
        matched = match_rows(ws, [row])
        isempty(matched) && throw(XLSXError("Row $row not found in worksheet $(ws.name)."))
        return only(matched)
    end
end

@inline row_number(sr::SheetRow) = sr.row

"""
    getcell(xlsxfile, cell_reference_name) :: AbstractCell
    getcell(worksheet, cell_reference_name) :: AbstractCell
    getcell(sheetrow, column_name) :: AbstractCell
    getcell(sheetrow, column_number) :: AbstractCell

Returns the internal representation of a worksheet cell.

Returns `XLSX.EmptyCell` if the cell has no data.
"""
function getcell(r::SheetRow, column_index::Int) :: AbstractCell
    if haskey(r.rowcells, column_index)
        return r.rowcells[column_index]
    else
        return EmptyCell(CellRef(row_number(r), column_index))
    end
end

function getcell(r::SheetRow, column_name::AbstractString)
    !is_valid_column_name(column_name) && throw(XLSXError("$column_name is not a valid column name."))
    return getcell(r, decode_column_number(column_name))
end

getdata(r::SheetRow, column::Union{Vector{T}, UnitRange{T}}) where {T<:Integer} = [getdata(get_worksheet(r), getcell(r, x)) for x in column]
getdata(r::SheetRow, column) = getdata(get_worksheet(r), getcell(r, column))
Base.getindex(r::SheetRow, x) = getdata(r, x)

Base.eachrow(ws::Worksheet) = eachrow(ws)
"""
    eachrow(sheet)

Creates a row iterator for a worksheet.

Base.eachrow(sheet::Worksheet) is defined as a synonym of XLSX.eachrow(sheet::Worksheet)

Example: Query all cells from columns 1 to 4.

```julia
left = 1  # 1st column
right = 4 # 4th column
for sheetrow in eachrow(sheet)
    for column in left:right
        cell = XLSX.getcell(sheetrow, column)

        # do something with cell
    end
end
```

!!! note

    The `eachrow` row iterator will not return any row that 
    consists entirely of `EmptyCell`s. These empty rows are not 
    represented in the .xlsx file and are therefore not seen by the 
    iterator. The `length(eachrow(sheet))` function returns 
    the number of rows that are not entirely empty and will, in any 
    case, only succeed if the worksheet cache is in use.

"""
function eachrow(ws::Worksheet) :: SheetRowIterator
    if is_cache_enabled(ws)
        if ws.cache === nothing
            target_file = get_relationship_target_by_id("xl", get_workbook(ws), ws.relationship_id)
            xf = get_xlsxfile(ws)
            raw = xf.data[target_file]
            raw isa String || throw(XLSXError("Expected raw XML string for $target_file, got parsed node."))
            lznode = parse(raw, XML.LazyNode)
            first_cache_fill!(ws, lznode)
            # swap back to the stub made at open, or make one if there isn't one
            stripped = pop!(xf.sheet_stubs, target_file, nothing)
            isnothing(stripped) && ((stripped, _) = splitNode(raw, "sheetData"))
            xf.data[target_file] = stripped
        end
        return ws.cache
    else
        return SheetRowStreamIterator(ws)
    end
end

function Base.isempty(sr::SheetRow)
    return isempty(sr.rowcells)
end

Base.length(r::WorksheetCache)=length(r.cells)

const _EMPTY_ROW_ATTRS = Dict{String,String}()

#--------------------------------------------------------------------- Fill cache on first read (multi-threaded)

function _find_sheetdata(doc::XML.LazyNode, wsname::String)::XML.LazyNode
    c = XML.Cursor(doc)
    while XML.next!(c) !== nothing
        d = XML.depth(c)
        d < 2 && continue
        d > 2 && (XML.skip_element!(c); continue)
        if XML.nodetype(c) == XML.Element && localname(c) == "sheetData"
            return XML.LazyNode(c)
        end
        XML.skip_element!(c)
    end
    throw(XLSXError("No `sheetData` node found in worksheet $wsname."))
end
# Push a row read by `first_cache_fill!` into the cache, with its unhandled
# attributes, and if `track`, widen `bounds` (row_min, row_max, col_min, col_max) to it.
@inline function _push_filled_row!(ws::Worksheet, unhandled_attributes::Dict{Int,Dict{String,String}}, row_num::Int,
                                   row_ht, rowcells::Dict{Int,Cell}, unhandled::Dict{String,String},
                                   track::Bool, bounds::NTuple{4,Int})::NTuple{4,Int}
    !isempty(unhandled) && (unhandled_attributes[row_num] = unhandled)
    push_sheetrow!(ws.cache, SheetRow(ws, row_num, row_ht, rowcells))
    track || return bounds
    row_min, row_max, col_min, col_max = bounds
    row_min = min(row_min, row_num); row_max = max(row_max, row_num)
    for col in keys(rowcells)
        col_min = min(col_min, col); col_max = max(col_max, col)
    end
    return (row_min, row_max, col_min, col_max)
end

function first_cache_fill!(ws::Worksheet, lznode::XML.LazyNode)
    handled_attributes = Set{String}(["r", "spans", "ht"])   # `customHeight` is kept as the file has it
    unhandled_attributes = Dict{Int,Dict{String,String}}()
    sst_pfx = get_sst_prefix(ws)
    wb = get_workbook(ws)
    load_formulas = get_xlsxfile(ws).load_formulas
    local_formulas = Dict{SheetCellRef, AbstractFormula}()  # ← local dict

    if ws.cache === nothing
        ws.cache = WorksheetCache(ws)
    else
        throw(XLSXError("Expecting empty cache but cache not empty!"))
    end

    sheetdata_lazy = _find_sheetdata(lznode, ws.name)

    sst_total = 0
    rowcells  = Dict{Int,Cell}()
    row_num   = nothing

    # Derive the sheet bounds from the rows as they are pushed, but only for a
    # sheet with no recorded dimension.
    track = isnothing(ws.dimension)
    bounds = (typemax(Int), 0, typemax(Int), 0)   # row_min, row_max, col_min, col_max
    row_ht    = nothing
    unhandled = _EMPTY_ROW_ATTRS

    # Pre-compute expected column count from sheet dimension for Dict sizehint
    dim = ws.dimension
    expected_cols = isnothing(dim) ? 16 :
        XLSX.column_number(dim.stop) - XLSX.column_number(dim.start) + 1

    c2 = XML.Cursor(sheetdata_lazy)
    while XML.next!(c2) !== nothing
        d  = XML.depth(c2)
        nt = XML.nodetype(c2)

        if d == 2 && nt == XML.Element && localname(c2) == "row"
            if !isnothing(row_num)
                bounds = _push_filled_row!(ws, unhandled_attributes, row_num, row_ht, rowcells, unhandled, track, bounds)
                rowcells  = Dict{Int,Cell}()
                sizehint!(rowcells, expected_cols)
                unhandled = _EMPTY_ROW_ATTRS
                row_ht    = nothing
            end
            # All of the row's attributes in one pass, where the cursor stands.
            r_val  = nothing
            ht_val = nothing
            for (k, v) in XML.eachattribute!(c2)
                if k == "r"
                    r_val = v
                elseif k == "ht"
                    ht_val = v
                elseif !(k in handled_attributes)
                    if unhandled === _EMPTY_ROW_ATTRS
                        unhandled = Dict{String,String}()
                    end
                    unhandled[String(k)] = String(v)
                end
            end
            r_val === nothing && throw(XLSXError("Row without 'r' attribute in worksheet $(ws.name)."))
            row_num = parse(Int, r_val)
            row_ht  = isnothing(ht_val) ? nothing : parse(Float64, ht_val)

        elseif d == 3 && nt == XML.Element && localname(c2) == "c"
            # Read in the same walk as its children, so its bytes are tokenized once (#462).
            cell = _cursor_cell(c2, ws, wb, sst_pfx, local_formulas, load_formulas)
            sst_total += cell.datatype == CT_STRING ? 1 : 0
            rowcells[column_number(cell)] = cell

        elseif nt == XML.Element && (d == 2 || d == 3)   # anything else in a row, or in <sheetData>
            XML.skip_element!(c2)
        end
    end

    if !isnothing(row_num)
        bounds = _push_filled_row!(ws, unhandled_attributes, row_num, row_ht, rowcells, unhandled, track, bounds)
    end

    row_min, row_max, col_min, col_max = bounds
    if track && col_max > 0 && isnothing(ws.dimension)
        set_dimension!(ws, CellRange(CellRef(row_min, col_min), CellRef(row_max, col_max)))
    end

    ws.sst_count = sst_total
    ws.unhandled_attributes = isempty(unhandled_attributes) ? nothing : unhandled_attributes
    
    # Update next_formula_id from this sheet's own formulas — all of them are in
    # `local_formulas`, so there is no need to scan the whole workbook's under the lock.
    max_id = -1
    for f in values(local_formulas)
        f isa ReferencedFormula && (max_id = max(max_id, f.id))
    end
    max_id >= ws.next_formula_id && (ws.next_formula_id = max_id + 1)

    # Merge local formulas into workbook dict under single lock
    if !isempty(local_formulas)
        lock(wb.formulas_lock) do
            merge!(wb.formulas, local_formulas)
        end
    end

    ws.cache.is_full = true
end

# Materialise specific rows from a worksheet.xml file into SheetRows
# (faster than using eachrow which materialises every row). Rows are in ascending
# order, so the pass stops after the last one wanted; a duplicated row: the first.
function match_rows(ws::Worksheet, rows_to_match::Vector{Int})::Vector{SheetRow}
    matched_rows = Vector{SheetRow}()
    wanted = sort(unique(rows_to_match))
    isempty(wanted) && return matched_rows
    wb = get_workbook(ws)
    sst_pfx = get_sst_prefix(ws)
    local_formulas = Dict{SheetCellRef,AbstractFormula}()
    load_formulas = get_xlsxfile(ws).load_formulas

    i = 1
    c = XML.Cursor(_open_sheetdata(ws))
    while i <= length(wanted) && _next_row!(c)
        row_num = _row_number(c, ws.name)
        while i <= length(wanted) && wanted[i] < row_num   # wanted rows absent from the file
            i += 1
        end
        if i <= length(wanted) && wanted[i] == row_num
            rowcells = Dict{Int,Cell}()
            ht = _row_height(c)
            _cursor_rowcells!(rowcells, c, ws, wb, sst_pfx, local_formulas, load_formulas)
            push!(matched_rows, SheetRow(ws, row_num, ht, rowcells))
            i += 1
        else
            XML.skip_element!(c)
        end
    end

    _merge_local_formulas!(wb, local_formulas)
    return matched_rows
end

# Column number from the letters of a cell reference such as "AB12".
@inline function _ref_column_number(ref::AbstractString)::Int
    n = 0
    for b in codeunits(ref)
        UInt8('A') <= b <= UInt8('Z') || break
        n = n * 26 + (b - UInt8('A') + 1)
    end
    n == 0 && throw(XLSXError("Invalid cell reference `$ref`."))
    return n
end

# Decodes the `<c>` the cursor is on, reading its children in the same walk, so its
# bytes are tokenised once. Reads cells as `Cell(::LazyNode, …)` does: an inline
# string's first `<is>` only, the last `<v>` wins, and an `<f>` is recorded (in
# `local_formulas`) when `load_formulas` is set. Leaves the cursor holding the node
# after the cell, so the caller's next `next!` yields it.
function _cursor_cell(c::XML.Cursor, ws::Worksheet, wb::Workbook, sst_pfx::String,
                      local_formulas::Dict{SheetCellRef,AbstractFormula}, load_formulas::Bool)::Cell
    attrs = _cell_attributes(c)
    datatype, value, formula = _cursor_cell_contents(c, attrs, ws, wb, sst_pfx, local_formulas, load_formulas)
    ref, _, _, style, _, meta = attrs
    return Cell(ref, value, style, meta, datatype, formula)
end

# The `(datatype, value, formula)` of the `<c>` the cursor is on, as `_cursor_cell`
# reads them, given its `_cell_attributes`; for a caller that needs no `Cell` (which
# is mutable, so would be allocated).
function _cursor_cell_contents(c::XML.Cursor, attrs::Tuple, ws::Worksheet, wb::Workbook, sst_pfx::String,
                               local_formulas::Dict{SheetCellRef,AbstractFormula}, load_formulas::Bool)
    ref, ref_str, t, _, num_style, _, vm = attrs
    datatype = CT_EMPTY
    value    = UInt64(0)
    formula  = false
    done     = false            # an inline string's `<is>` has been read: ignore the rest
    XML.@for_each_child! c child begin
        if XML.nodetype(child) == XML.Element
            tag = localname(child)
            if done
                # past an inline string's `<is>`
            elseif t == "inlineStr"
                if tag == "is"
                    r = _inline_string(wb, XML.LazyNode(child), sst_pfx)
                    isnothing(r) || ((datatype, value) = r)
                    done = true
                end
            elseif tag == "v"
                sv = XML.is_simple_value(child)
                isnothing(sv) && (sv = _v_text_fallback(XML.LazyNode(child), ref_str))
                if !isnothing(sv) && !isempty(sv)
                    datatype, value = process_tv(wb, t, sv, num_style)
                end
            elseif tag == "f"
                load_formulas && _record_formula!(wb, ws, ref, XML.LazyNode(child), local_formulas)
                formula = true
            end
            XML.skip_element!(child)
        end
    end
    datatype, value = _rich_value(datatype, value, vm)
    return datatype, value, formula
end

# Whether `x` is in `sel`: `nothing` selects everything; ranges test in O(1), and a
# sorted vector by binary search, so the test doesn't depend on the order rows arrive in.
@inline _selects(::Nothing, ::Int) = true
@inline _selects(sel::AbstractRange{<:Integer}, x::Int) = x in sel
@inline _selects(sel::AbstractVector{<:Integer}, x::Int) = insorted(x, sel)

# The last element `sel` selects, or `nothing` when it is unbounded.
@inline _last_selected(::Nothing) = nothing
@inline _last_selected(sel::AbstractVector{<:Integer}) = isempty(sel) ? 0 : last(sel)

# One streaming pass over a worksheet.xml file. Decodes the cells in `rows` × `cols`
# (`nothing` = unrestricted; otherwise a range or a sorted vector) and skims the rest,
# reading only the `r` attribute of each `<row>` and `<c>`. Returns the decoded cells
# and, when `track_bounds` is set, the bounds of every cell in the sheet (`nothing` for
# a sheet with no cells; like the row iterator, rows without cells count towards the
# row bounds). Without `track_bounds`, the pass stops at the first row after the last
# selected one, as rows are in ascending order.
function _read_cells(ws::Worksheet,
                     rows::Union{Nothing,AbstractVector{<:Integer}},
                     cols::Union{Nothing,AbstractVector{<:Integer}};
                     track_bounds::Bool)::Tuple{Vector{Cell},Union{Nothing,CellRange}}
    wb = get_workbook(ws)
    sst_pfx = get_sst_prefix(ws)
    local_formulas = Dict{SheetCellRef,AbstractFormula}()
    load_formulas = get_xlsxfile(ws).load_formulas
    stop_after = track_bounds ? nothing : _last_selected(rows)
    cells = Cell[]
    row_min = col_min = typemax(Int)
    row_max = col_max = 0
    in_rows = false

    c = XML.Cursor(_open_sheetdata(ws))
    while XML.next!(c) !== nothing
        d = XML.depth(c)
        d == 1 && continue
        if d == 2 && XML.nodetype(c) == XML.Element && localname(c) == "row"
            row = _row_number(c, ws.name)
            !isnothing(stop_after) && row > stop_after && break
            row_min = min(row_min, row); row_max = max(row_max, row)
            in_rows = _selects(rows, row)
            # Nothing to decode or measure in this row's cells.
            !in_rows && !track_bounds && XML.skip_element!(c)
        elseif d == 3 && XML.nodetype(c) == XML.Element && localname(c) == "c"
            col = _ref_column_number(something(_cursor_attr(c, "r"), ""))
            col_min = min(col_min, col); col_max = max(col_max, col)
            if in_rows && _selects(cols, col)
                push!(cells, _cursor_cell(c, ws, wb, sst_pfx, local_formulas, load_formulas))
            else
                XML.skip_element!(c)
            end
        else
            XML.skip_element!(c)
        end
    end

    _merge_local_formulas!(wb, local_formulas)
    bounds = (track_bounds && col_max > 0) ?
        CellRange(CellRef(row_min, col_min), CellRef(row_max, col_max)) : nothing
    return cells, bounds
end

# Bounds of the cells in a worksheet.xml file, read from the `r` attributes of
# `<row>` and `<c>` alone (no cells are materialised). Like the row iterator,
# rows without cells count towards the row bounds. Returns `nothing` for a sheet
# with no cells.
_scan_dimension(ws::Worksheet)::Union{Nothing,CellRange} =
    last(_read_cells(ws, 1:0, nothing; track_bounds = true))

# Single streaming pass over an uncached worksheet whose dimension is unknown.
# Materialises the cells in `rows` × `cols` (`nothing` = unrestricted) and reads
# only the `r` attribute of the rest, so the sheet bounds come from the same pass.
# Records the dimension, then returns the target range (unrestricted sides taken
# from the dimension) and the materialised cells. Returns `nothing` when the
# dimension is already known or the cache is in use, so callers keep their usual path.
function _read_unknown_dimension(ws::Worksheet,
                                 rows::Union{Nothing,AbstractUnitRange{<:Integer}},
                                 cols::Union{Nothing,AbstractUnitRange{<:Integer}})
    (isnothing(ws.dimension) && !is_cache_enabled(ws)) || return nothing
    is_chartsheet(get_workbook(ws), ws.name) && return nothing

    cells, bounds = _read_cells(ws, rows, cols; track_bounds = true)
    isnothing(bounds) || set_dimension!(ws, bounds)

    dim = something(ws.dimension, CellRange(CellRef(1, 1), CellRef(1, 1)))  # empty sheet: A1:A1, not recorded
    top    = isnothing(rows) ? dim.start.row_number    : first(rows)
    bottom = isnothing(rows) ? dim.stop.row_number     : last(rows)
    left   = isnothing(cols) ? dim.start.column_number : first(cols)
    right  = isnothing(cols) ? dim.stop.column_number  : last(cols)
    return CellRange(CellRef(top, left), CellRef(bottom, right)), cells
end
