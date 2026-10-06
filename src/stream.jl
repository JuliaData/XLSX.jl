
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

    return parse(String(ZipArchives.zip_readentry(zip_io, filename)), XML.LazyNode)

end


function _read_row_attrs(row::XML.LazyNode, wsname::String)
    current_row = nothing
    current_row_ht = nothing
    for (k, v) in XML.eachattribute(row)
        if k == "r"
            current_row = parse(Int, v)
        elseif k == "ht"
            current_row_ht = parse(Float64, v)
        end
    end
    current_row === nothing && throw(XLSXError("Row without 'r' attribute in worksheet $wsname."))
    return current_row, current_row_ht
end

# Advance `row_iter` from `next` (a raw `iterate` result) until the next
# `<row>` element. Returns `(rownode, iterator_state)` or `nothing` at EOF.
# Takes the first `iterate` result as an argument rather than a sentinel so
# it works even if the iterator's own state type is `Nothing`.
@inline function _find_next_row(row_iter, next)
    while next !== nothing
        child, st = next
        if XML.nodetype(child) == XML.Element && localname(child) == "row"
            return child, st
        end
        next = iterate(row_iter, st)
    end
    return nothing
end

# Open the sheet's XML stream and return its <sheetData> node.
# Shared by the stream iterator and its tests so both navigate identically.
function _open_sheetdata(ws::Worksheet)
    target_file = get_relationship_target_by_id("xl", get_workbook(ws), ws.relationship_id)
    xf = get_xlsxfile(ws)
    doc = open_internal_file_stream(xf, target_file)
    return _find_sheetdata(doc, ws.name)
end

# Creates an iterator for row elements in the Worksheet's XML.
# Creates an iterator for row elements in the Worksheet's XML.
function Base.iterate(itr::SheetRowStreamIterator)
    ws = get_worksheet(itr)
    xf = get_xlsxfile(ws)
    sst_pfx = get_sst_prefix(ws)
    sheetdata = _open_sheetdata(ws)
    row_iter = XML.eachchildnode(sheetdata)

    found = _find_next_row(row_iter, iterate(row_iter))
    isnothing(found) && return nothing
    rownode, row_state = found

    rowcells = Dict{Int,Cell}()
    local_formulas = Dict{SheetCellRef,AbstractFormula}()
    load_formulas = xf.load_formulas
    current_row, current_row_ht = _read_row_attrs(rownode, ws.name)
    _, sst_count = get_rowcells!(rowcells, rownode, ws, sst_pfx, local_formulas, load_formulas)
    itr.sheet.sst_count += sst_count
    _merge_local_formulas!(get_workbook(ws), local_formulas)
    state = SheetRowStreamIteratorState(row_iter, row_state, rowcells, local_formulas, 1,
                                        isnothing(ws.dimension), typemax(Int), 0, typemax(Int), 0)
    state.track && _track_row_bounds!(state, current_row, rowcells)
    return SheetRow(ws, current_row, current_row_ht, rowcells), state
end

@inline function _track_row_bounds!(state::SheetRowStreamIteratorState, row::Int, rowcells::Dict{Int,Cell})
    state.row_min = min(state.row_min, row)
    state.row_max = max(state.row_max, row)
    for col in keys(rowcells)
        state.col_min = min(state.col_min, col)
        state.col_max = max(state.col_max, col)
    end
    nothing
end

function Base.iterate(itr::SheetRowStreamIterator, state::SheetRowStreamIteratorState)
    ws = get_worksheet(itr)
    sst_pfx = get_sst_prefix(ws)
    empty!(state.rowcells)

    found = _find_next_row(state.row_iter, iterate(state.row_iter, state.row_state))
    if isnothing(found)
        # A completed pass has seen every row: record the bounds if still unknown.
        if state.track && state.col_max > 0 && isnothing(ws.dimension)
            set_dimension!(ws, CellRange(CellRef(state.row_min, state.col_min), CellRef(state.row_max, state.col_max)))
        end
        return nothing
    end
    rownode, state.row_state = found

    load_formulas = get_xlsxfile(ws).load_formulas
    current_row, current_row_ht = _read_row_attrs(rownode, ws.name)
    _, sst_count = get_rowcells!(state.rowcells, rownode, ws, sst_pfx, state.local_formulas, load_formulas)
    itr.sheet.sst_count += sst_count
    state.track && _track_row_bounds!(state, current_row, state.rowcells)

    state.rows_since_merge += 1
    if state.rows_since_merge >= 500
        _merge_local_formulas!(get_workbook(ws), state.local_formulas)
        state.rows_since_merge = 0
    end

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

    # the sorting operation is very costly when adding row and only needed if we use the row iterator
    if ws_cache.dirty
        sort!(ws_cache.rows_in_cache)
        ws_cache.row_index = Dict{Int, Int}(ws_cache.rows_in_cache[i] => i for i in 1:length(ws_cache.rows_in_cache))
        ws_cache.dirty = false
    end

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
        r = first(match_rows(ws, [row]))
        if isnothing(r)
            throw(XLSXError("Row $row not found in worksheet $(ws.name)."))
        else
            return r
        end
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
function first_cache_fill!(ws::Worksheet, lznode::XML.LazyNode)
    handled_attributes = Set{String}(["r", "spans", "ht", "customHeight"])
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

    sheetdata_lazy = nothing
    c = XML.Cursor(lznode)
    while XML.next!(c) !== nothing
        d = XML.depth(c)
        d < 2 && continue
        d > 2 && (XML.skip_element!(c); continue)
        if XML.nodetype(c) == XML.Element && localname(c) == "sheetData"
            sheetdata_lazy = XML.LazyNode(c)
            break
        end
        XML.skip_element!(c)
    end
    sheetdata_lazy === nothing && throw(XLSXError("No `sheetData` node found in worksheet"))

    sst_total = 0
    rowcells  = Dict{Int,Cell}()
    row_num   = nothing

    # Derive the sheet bounds from the rows as they are pushed, but only for a
    # sheet with no recorded dimension.
    track = isnothing(ws.dimension)
    row_min = col_min = typemax(Int)
    row_max = col_max = 0
    row_ht    = nothing
    unhandled = _EMPTY_ROW_ATTRS

    # Pre-compute expected column count from sheet dimension for Dict sizehint
    dim = ws.dimension
    expected_cols = isnothing(dim) ? 16 :
        XLSX.column_number(dim.stop) - XLSX.column_number(dim.start) + 1

    # The cell being read. A `<c>` is read in the same cursor walk as its children
    # (rather than skipped and re-read through `Cell(::LazyNode, …)`), so its bytes are
    # tokenized once (#462). It is completed when the walk leaves it (depth ≤ 3), using
    # the same helpers as `Cell(::LazyNode, …)`.
    in_cell    = false
    cell_done  = false          # an inline string's `<is>` has been read: ignore the rest
    cell_ref   = CellRef(1, 1)
    cell_rstr  = SubString("")
    cell_t     = SubString("")
    cell_style = UInt32(0)
    cell_nstyle = 0
    cell_meta  = UInt32(0)
    cell_type  = CT_EMPTY
    cell_value = UInt64(0)
    cell_formula = false

    c2 = XML.Cursor(sheetdata_lazy)
    while XML.next!(c2) !== nothing
        d  = XML.depth(c2)
        nt = XML.nodetype(c2)

        if in_cell && d <= 3
            cell = Cell(cell_ref, cell_value, cell_style, cell_meta, cell_type, cell_formula)
            sst_total += cell_type == CT_STRING ? 1 : 0
            rowcells[column_number(cell)] = cell
            in_cell = false
        end

        if d == 4 && in_cell
            nt == XML.Element || continue
            tag = localname(c2)
            if cell_done
                # past an inline string's `<is>`
            elseif cell_t == "inlineStr"
                if tag == "is"
                    r = _inline_string(wb, XML.LazyNode(c2), sst_pfx)
                    isnothing(r) || ((cell_type, cell_value) = r)
                    cell_done = true
                end
            elseif tag == "v"
                sv = XML.is_simple_value(c2)
                isnothing(sv) && (sv = _v_text_fallback(XML.LazyNode(c2), cell_rstr))
                if !isnothing(sv) && !isempty(sv)
                    cell_type, cell_value = process_tv(wb, cell_t, sv, cell_nstyle)
                end
            elseif tag == "f"
                load_formulas && _record_formula!(wb, ws, cell_ref, XML.LazyNode(c2), local_formulas)
                cell_formula = true
            end
            XML.skip_element!(c2)

        elseif d == 2 && nt == XML.Element && localname(c2) == "row"
            if !isnothing(row_num)
                sr = SheetRow(ws, row_num, row_ht, rowcells)
                !isempty(unhandled) && (unhandled_attributes[row_num] = unhandled)
                push_sheetrow!(ws.cache, sr)
                if track
                    row_min = min(row_min, row_num); row_max = max(row_max, row_num)
                    for col in keys(rowcells)
                        col_min = min(col_min, col); col_max = max(col_max, col)
                    end
                end
                rowcells  = Dict{Int,Cell}()
                sizehint!(rowcells, expected_cols)
                unhandled = _EMPTY_ROW_ATTRS
                row_ht    = nothing
            end
            r_val = XML.get(c2, "r", nothing)
            r_val === nothing && throw(XLSXError("Row without 'r' attribute in worksheet $(ws.name)."))
            row_num = parse(Int, r_val)
            ht_val  = XML.get(c2, "ht", nothing)
            row_ht  = isnothing(ht_val) ? nothing : parse(Float64, ht_val)

            let ln = XML.LazyNode(c2)
                for (k, v) in XML.eachattribute(ln)
                    k in handled_attributes && continue
                    if unhandled === _EMPTY_ROW_ATTRS
                        unhandled = Dict{String,String}()
                    end
                    unhandled[String(k)] = String(v)
                end
            end

        elseif d == 3 && nt == XML.Element && localname(c2) == "c"
            cell_ref, cell_rstr, cell_t, cell_style, cell_nstyle, cell_meta =
                _cell_attributes(XML.LazyNode(c2))
            cell_type, cell_value, cell_formula = CT_EMPTY, UInt64(0), false
            in_cell, cell_done = true, false

        elseif d == 2 && nt == XML.Element
            XML.skip_element!(c2)

        elseif d == 3 && nt == XML.Element
            XML.skip_element!(c2)

        end
    end

    if in_cell
        cell = Cell(cell_ref, cell_value, cell_style, cell_meta, cell_type, cell_formula)
        sst_total += cell_type == CT_STRING ? 1 : 0
        rowcells[column_number(cell)] = cell
    end

    if !isnothing(row_num)
        sr = SheetRow(ws, row_num, row_ht, rowcells)
        !isempty(unhandled) && (unhandled_attributes[row_num] = unhandled)
        push_sheetrow!(ws.cache, sr)
        if track
            row_min = min(row_min, row_num); row_max = max(row_max, row_num)
            for col in keys(rowcells)
                col_min = min(col_min, col); col_max = max(col_max, col)
            end
        end
    end

    if track && col_max > 0 && isnothing(ws.dimension)
        set_dimension!(ws, CellRange(CellRef(row_min, col_min), CellRef(row_max, col_max)))
    end

    ws.sst_count = sst_total
    ws.unhandled_attributes = isempty(unhandled_attributes) ? nothing : unhandled_attributes
    
    # Merge local formulas into workbook dict under single lock
    if !isempty(local_formulas)
        lock(wb.formulas_lock) do
            merge!(wb.formulas, local_formulas)
        end
    end

    # Update next_formula_id from merged formulas
      lock(wb.formulas_lock) do
          isempty(wb.formulas) && return
            ws_name = ws.name
            max_id = -1
            for (ref, f) in wb.formulas
                if ref.sheet == ws_name && f isa ReferencedFormula
                max_id = max(max_id, f.id)
            end             
        end
        if max_id >= ws.next_formula_id
            ws.next_formula_id = max_id + 1
        end
    end

    ws.cache.is_full = true
end

# Materialise specific rows from a worksheet.xml file into SheetRows
# (faster than using eachrow which materialises every row).
function match_rows(ws::Worksheet, rows_to_match::Vector{Int})::Vector{SheetRow}
    matched_rows = Vector{SheetRow}()
    sst_pfx = get_sst_prefix(ws)
    local_formulas = Dict{SheetCellRef,AbstractFormula}()
    load_formulas = get_xlsxfile(ws).load_formulas
    sort!(rows_to_match)

    target_file = get_relationship_target_by_id("xl", get_workbook(ws), ws.relationship_id)
    xf = get_xlsxfile(ws)
    doc = open_internal_file_stream(xf, target_file)
    sheetdata = _find_sheetdata(doc, ws.name)

    i = 1
    c = XML.Cursor(sheetdata)
    while XML.next!(c) !== nothing && i <= length(rows_to_match)
        XML.depth(c) == 1 && continue
        XML.depth(c) != 2 && (XML.skip_element!(c); continue)
        XML.nodetype(c) == XML.Element && localname(c) == "row" || (XML.skip_element!(c); continue)

        row_num_str = XML.get(c, "r", nothing)
        row_num_str === nothing && throw(XLSXError("Row without 'r' attribute encountered in worksheet $(ws.name)."))
        row_num = parse(Int, row_num_str)

        row_num < rows_to_match[i] && (XML.skip_element!(c); continue)
        row_num != rows_to_match[i] && (XML.skip_element!(c); continue)

        ht_str = XML.get(c, "ht", nothing)
        row_node = XML.LazyNode(c)
        rowcells = Dict{Int,Cell}()
        get_rowcells!(rowcells, row_node, ws, sst_pfx, local_formulas, load_formulas)
        push!(matched_rows, SheetRow(ws, row_num, isnothing(ht_str) ? nothing : parse(Float64, ht_str), rowcells))
        i += 1
    end

    if !isempty(local_formulas)
        wb = get_workbook(ws)
        lock(wb.formulas_lock) do
            merge!(wb.formulas, local_formulas)
        end
    end

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

# Bounds of the cells in a worksheet.xml file, read from the `r` attributes of
# `<row>` and `<c>` alone (no cells are materialised). Like the row iterator,
# rows without cells count towards the row bounds. Returns `nothing` for a sheet
# with no cells.
function _scan_dimension(ws::Worksheet)::Union{Nothing,CellRange}
    row_min = col_min = typemax(Int)
    row_max = col_max = 0
    c = XML.Cursor(_open_sheetdata(ws))
    while XML.next!(c) !== nothing
        d = XML.depth(c)
        d == 1 && continue
        if d == 2 && XML.nodetype(c) == XML.Element && localname(c) == "row"
            r_val = XML.get(c, "r", nothing)
            r_val === nothing && throw(XLSXError("Row without 'r' attribute in worksheet $(ws.name)."))
            row = parse(Int, r_val)
            row_min = min(row_min, row); row_max = max(row_max, row)
        elseif d == 3 && XML.nodetype(c) == XML.Element && localname(c) == "c"
            col = _ref_column_number(XML.get(c, "r", ""))
            col_min = min(col_min, col); col_max = max(col_max, col)
            XML.skip_element!(c)
        else
            XML.skip_element!(c)
        end
    end
    col_max == 0 && return nothing
    return CellRange(CellRef(row_min, col_min), CellRef(row_max, col_max))
end

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

    sst_pfx = get_sst_prefix(ws)
    local_formulas = Dict{SheetCellRef,AbstractFormula}()
    load_formulas = get_xlsxfile(ws).load_formulas
    cells = Cell[]
    row_min = col_min = typemax(Int)
    row_max = col_max = 0
    in_rows = false

    c = XML.Cursor(_open_sheetdata(ws))
    while XML.next!(c) !== nothing
        d = XML.depth(c)
        d == 1 && continue
        if d == 2 && XML.nodetype(c) == XML.Element && localname(c) == "row"
            r_val = XML.get(c, "r", nothing)
            r_val === nothing && throw(XLSXError("Row without 'r' attribute in worksheet $(ws.name)."))
            row = parse(Int, r_val)
            row_min = min(row_min, row); row_max = max(row_max, row)
            in_rows = isnothing(rows) || row in rows
        elseif d == 3 && XML.nodetype(c) == XML.Element && localname(c) == "c"
            col = _ref_column_number(XML.get(c, "r", ""))
            col_min = min(col_min, col); col_max = max(col_max, col)
            if in_rows && (isnothing(cols) || col in cols)
                cell = Cell(XML.LazyNode(c), ws, sst_pfx, local_formulas, load_formulas)
                cell isa Cell && push!(cells, cell)
            end
            XML.skip_element!(c)
        else
            XML.skip_element!(c)
        end
    end

    _merge_local_formulas!(get_workbook(ws), local_formulas)
    col_max > 0 && set_dimension!(ws, CellRange(CellRef(row_min, col_min), CellRef(row_max, col_max)))

    dim = something(ws.dimension, CellRange(CellRef(1, 1), CellRef(1, 1)))  # empty sheet: A1:A1, not recorded
    top    = isnothing(rows) ? dim.start.row_number    : first(rows)
    bottom = isnothing(rows) ? dim.stop.row_number     : last(rows)
    left   = isnothing(cols) ? dim.start.column_number : first(cols)
    right  = isnothing(cols) ? dim.stop.column_number  : last(cols)
    return CellRange(CellRef(top, left), CellRef(bottom, right)), cells
end
