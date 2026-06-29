
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

#Base.show(io::IO, state::SheetRowStreamIteratorState) = print(io, "SheetRowStreamIteratorState( itr = $(state.itr), itr_state = $(state.itr_state), row = $(state.row) )")

xml_elements(node) = filter(n -> XML.nodetype(n) == XML.Element, XML.children(node))

## 1. Parsed DOM document
#xml_root_element(doc::XML.Document) = XML.root(doc)
xml_root_element(doc) = last(xml_elements(doc))

# 2. Parsed DOM node
#xml_root_element(n::XML.Node) = n

# 3. LazyNode
function xml_root_element(lz::XML.LazyNode)
    c = XML.Cursor(lz)
    while XML.next!(c) !== nothing
        if XML.nodetype(c) == XML.Element
            return XML.LazyNode(c)
        end
    end
    error("No root element found")
end

# Opens a file for streaming.
#=@inline function open_internal_file_stream(xf::XLSXFile, filename::String) :: XML.LazyNode

    cached = get(xf.worksheet_xml_cache, filename, nothing)
    cached !== nothing && return cached

    !internal_xml_file_exists(xf, filename) && throw(XLSXError("Couldn't find $filename in $(xf.source)."))
    if xf.source isa IO
        seekstart(xf.source)
        zip_io = ZipArchives.ZipReader(read(xf.source))
    else
        zip_io = ZipArchives.ZipReader(FileArray(abspath(xf.source)))
    end
    doc = parse(String(ZipArchives.zip_readentry(zip_io, filename)), XML.LazyNode)
    xf.worksheet_xml_cache[filename] = doc
    return doc
end
=#
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


# Collect all row LazyNodes from a worksheet's sheetData element.
function _collect_row_nodes(doc::XML.LazyNode)
    root = xml_root_element(doc)
    localname(root) != "worksheet" && throw(XLSXError("Expecting to find a worksheet node. Found a $(localname(root))."))

    # Find sheetData
    sheetdata = nothing
    for child in XML.children(root)
        if localname(child) == "sheetData"
            sheetdata = child
            break
        end
    end
    sheetdata === nothing && throw(XLSXError("No `sheetData` node found in worksheet"))

    # Collect row nodes
    return XML.LazyNode[child for child in XML.children(sheetdata) if localname(child) == "row"]
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

# Creates an iterator for row elements in the Worksheet's XML.
# Creates an iterator for row elements in the Worksheet's XML.
function Base.iterate(itr::SheetRowStreamIterator)
    ws = get_worksheet(itr)
    target_file = get_relationship_target_by_id("xl", get_workbook(ws), ws.relationship_id)
    xf = get_xlsxfile(ws)
    doc = open_internal_file_stream(xf, target_file)
    sst_pfx = get_sst_prefix(ws)
    sheetdata = _find_sheetdata(doc, ws.name)
    row_iter = XML.eachchildnode(sheetdata)
    # Find first row
    rownode = nothing
    for child in row_iter
        if XML.nodetype(child) == XML.Element && localname(child) == "row"
            rownode = child
            break
        end
    end
    isnothing(rownode) && return nothing
    rowcells = Dict{Int,Cell}()
    local_formulas = Dict{SheetCellRef,AbstractFormula}()
    load_formulas = xf.load_formulas
    current_row, current_row_ht = _read_row_attrs(rownode, ws.name)
    _, sst_count = get_rowcells!(rowcells, rownode, ws, sst_pfx, local_formulas, load_formulas)
    itr.sheet.sst_count += sst_count
    _merge_local_formulas!(get_workbook(ws), local_formulas)
    state = SheetRowStreamIteratorState(row_iter, rowcells, local_formulas, 1)
    return SheetRow(ws, current_row, current_row_ht, rowcells), state
end

function Base.iterate(itr::SheetRowStreamIterator, state::SheetRowStreamIteratorState)
    ws = get_worksheet(itr)
    sst_pfx = get_sst_prefix(ws)
    empty!(state.rowcells)
    rownode = nothing
    for child in state.row_iter
        if XML.nodetype(child) == XML.Element && localname(child) == "row"
            rownode = child
            break
        end
    end
    isnothing(rownode) && return nothing
    load_formulas = get_xlsxfile(ws).load_formulas
    current_row, current_row_ht = _read_row_attrs(rownode, ws.name)
    _, sst_count = get_rowcells!(state.rowcells, rownode, ws, sst_pfx, state.local_formulas, load_formulas)
    itr.sheet.sst_count += sst_count

    state.rows_since_merge += 1
    if state.rows_since_merge >= 500
        _merge_local_formulas!(get_workbook(ws), state.local_formulas)
        state.rows_since_merge = 0
    end

    return SheetRow(ws, current_row, current_row_ht, state.rowcells), state
end

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

        throw(XLSXError("Row $row not found."))

    # If can't use cache then lazily iterate sheetrows
    else
        r = first(match_rows(ws, [row]))
        if isnothing(r)
            throw(XLSXError("Row $row not found."))
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
            stripped, _ = splitNode(raw, "sheetData")
            xf.data[target_file] = stripped   # swap back to stub
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
                sr = SheetRow(ws, row_num, row_ht, rowcells)
                !isempty(unhandled) && (unhandled_attributes[row_num] = unhandled)
                push_sheetrow!(ws.cache, sr)
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
                XML.foreach_attr(ln) do name_tok, val_tok
                    k = XML.XMLTokenizer.raw(name_tok, ln.data)
                    k in handled_attributes && return
                    if unhandled === _EMPTY_ROW_ATTRS
                        unhandled = Dict{String,String}()
                    end
                    unhandled[String(k)] = String(XML.XMLTokenizer.attr_value(val_tok, ln.data))
                end
            end

        elseif d == 3 && nt == XML.Element && localname(c2) == "c"
            cell_node = XML.LazyNode(c2)
            XML.skip_element!(c2)
            cell = Cell(cell_node, ws, sst_pfx, local_formulas, load_formulas)
            sst_total += cell.datatype == CT_STRING ? 1 : 0
            rowcells[column_number(cell)] = cell

        elseif d == 2 && nt == XML.Element
            XML.skip_element!(c2)

        elseif d == 3 && nt == XML.Element
            XML.skip_element!(c2)

        end
    end

    if !isnothing(row_num)
        sr = SheetRow(ws, row_num, row_ht, rowcells)
        !isempty(unhandled) && (unhandled_attributes[row_num] = unhandled)
        push_sheetrow!(ws.cache, sr)
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
    if !isempty(wb.formulas)
        ws_name = ws.name
        max_id = -1
        lock(wb.formulas_lock) do
            for (ref, f) in wb.formulas
                if ref.sheet == ws_name && f isa ReferencedFormula
                    max_id = max(max_id, f.id)
                end
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
