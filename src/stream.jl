
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

# Opens a file for streaming.
@inline function open_internal_file_stream(xf::XLSXFile, filename::String) :: XML.LazyNode

    !internal_xml_file_exists(xf, filename) && throw(XLSXError("Couldn't find $filename in $(xf.source)."))
#    if xf.use_cache_for_sheet_data || (xf.source isa IO)
    if xf.source isa IO
        seekstart(xf.source)
        zip_io = ZipArchives.ZipReader(read(xf.source))
    else
        zip_io = ZipArchives.ZipReader(FileArray(abspath(xf.source))) # FileArray is marginally slower than mmap
#       zip_io = ZipArchives.ZipReader(Mmap.mmap(abspath(xf.source))) # but Mmap is unreliable : https://discourse.julialang.org/t/struggling-to-use-mmap-with-ziparchives/129839
    end

    return XML.LazyNode(XML.Raw(ZipArchives.zip_readentry(zip_io, filename)))

end

# Creates an iterator for row elements in the Worksheet's XML.
function Base.iterate(itr::SheetRowStreamIterator)
    ws = get_worksheet(itr)
    wb = get_workbook(ws)
    target_file = get_relationship_target_by_id("xl", get_workbook(ws), ws.relationship_id)
    sheetnode = open_internal_file_stream(get_xlsxfile(ws), target_file) # worksheet target files are LazyNodes
    sst_pfx = get_sst_prefix(ws)

    length(sheetnode) <= 0 && throw(XLSXError("Couldn't open reader for Worksheet $(ws.name)."))
    localname(sheetnode[end]) !=  "worksheet"  && throw(XLSXError("Expecting to find a worksheet node: Found a $(localname(sheetnode[end]))."))

    sheetnode=XML.next(sheetnode)

    while localname(sheetnode) != "sheetData" # Check for `sheetData`
        sheetnode = XML.next(sheetnode)
        sheetnode === nothing && throw(XLSXError("No `sheetData` node found in worksheet"))
    end

    XML.depth(sheetnode) != 2 && throw(XLSXError("Malformed Worksheet \"$(ws.name)\": unexpected node depth for sheetData node: $(XML.depth(sheetnode))."))

    rownode=XML.next(sheetnode)

    while localname(rownode) != "row" # Check for at least one `row`
        rownode = XML.next(rownode)
        rownode === nothing && return nothing # no rows found
    end

    # rownode is now the first row
    a = XML.attributes(rownode) # get row number and row height (if specified)
    current_row = parse(Int, a["r"])
    current_row_ht = haskey(a, "ht") ? parse(Float64, a["ht"]) : nothing

    # collect all cells in this row
    rowcells = Dict{Int, Cell}()
    mylock=ReentrantLock()
    next_rownode, sst_count = get_rowcells!(rowcells, rownode, ws, sst_pfx; mylock) # update rowcells in place
    
    itr.sheet.sst_count += sst_count

    sheet_row = SheetRow(ws, current_row, current_row_ht, rowcells) # create the sheet_row

    # debug
#    @assert sheetnode.raw.data == next_rownode.raw.data "LazyNode data don't match"

    return sheet_row, SheetRowStreamIteratorState(next_rownode, rowcells, mylock)
end

function Base.iterate(itr::SheetRowStreamIterator, state::SheetRowStreamIteratorState)
    ws = get_worksheet(itr)
    rownode = state.next_rownode
    rowcells = state.rowcells
    mylock = state.lock
    sst_pfx = get_sst_prefix(ws)

    empty!(rowcells)

    if rownode === nothing # there is no next_rownode - all rows processed
        return nothing
    end

    # get row number and row height (if specified)
    a = XML.attributes(rownode)
    current_row = parse(Int, a["r"])
    current_row_ht = haskey(a, "ht") ? parse(Float64, a["ht"]) : nothing

    # collect all cells in this row
    next_rownode, sst_count = get_rowcells!(rowcells, rownode, ws, sst_pfx; mylock) # update rowcells in place
    itr.sheet.sst_count += sst_count

    sheet_row = SheetRow(ws, current_row, current_row_ht, rowcells) # create the sheet_row

    return sheet_row, SheetRowStreamIteratorState(next_rownode, rowcells, mylock)
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
        if ws.cache === nothing # fill cache if enabled but empty on first use of eachrow iterator
            target_file = get_relationship_target_by_id("xl", get_workbook(ws), ws.relationship_id)
            lznode = open_internal_file_stream(get_xlsxfile(ws), target_file)
            first_cache_fill!(ws, lznode, Threads.nthreads()) # eagerly fill cache
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

#--------------------------------------------------------------------- Fill cache on first read (multi-threaded)
function produce_rowchunks!(out, n, rows, chunksize)
    pos=0
    while !isnothing(n)
        if _is_tag(localname(n), "row")
            pos += 1
            rows[pos] = n
        end
        if pos >= chunksize
            put!(out, copy(rows))
            pos=0
        end
        n = XML.next(n)
    end
    if pos>0 # handle last incomplete chunk
        put!(out, copy(@view rows[1:pos]))
    end
end

function stream_rows(n::XML.LazyNode, chunksize::Int; channel_size::Int=1 << 8)
    rows = Vector{XML.LazyNode}(undef, chunksize)
    Channel{Vector{XML.LazyNode}}(channel_size) do out
        produce_rowchunks!(out, n, rows, chunksize)
    end
end

function process_row(row::XML.LazyNode, handled_attributes::Set{String}, ws::Worksheet, sst_pfx::String, mylock::ReentrantLock)
    unhandled_attributes = Dict{String,String}()
    current_row_ht = nothing   # initialise here
    row_num = nothing          # initialise here

    atts = XML.attributes(row)
    isnothing(atts) && return nothing   # skip rows with no attributes

    current_row_ht = haskey(atts, "ht") ? parse(Float64, atts["ht"]) : nothing
    row_num = haskey(atts, "r") ? parse(Int, atts["r"]) : nothing
    row_num === nothing && throw(XLSXError("Row without 'r' attribute encountered in worksheet $(ws.name)."))
    unhandled_attributes = Dict(filter(attr -> !in(first(attr), handled_attributes), atts))

    rowcells = Dict{Int,Cell}()
    _, sst_count = get_rowcells!(rowcells, row, ws, sst_pfx; mylock)

    return sst_count, SheetRow(ws, row_num, current_row_ht, rowcells), unhandled_attributes
end

function first_cache_fill!(ws::Worksheet, lznode::XML.LazyNode, nthreads::Int)
    chunksize = ROW_CHUNKSIZE
    handled_attributes = Set{String}(["r", "spans", "ht", "customHeight"])
    unhandled_attributes = Dict{Int,Dict{String,String}}()
    sst_pfx = get_sst_prefix(ws)
   
    if ws.cache === nothing
        ws.cache = WorksheetCache(ws)
    else
        throw(XLSXError("Expecting empty cache but cache not empty!"))
    end
   
    sheet_rows = Channel{Vector{Tuple{Int, SheetRow, Dict{String,String}}}}(1 << 8)
   
    consumer = @async begin
        sst_total = 0
        for rows in sheet_rows
            for (row_sst_count, sheet_row, unatt) in rows
                if !isempty(unatt)
                    unhandled_attributes[row_number(sheet_row)] = unatt
                end
                push_sheetrow!(ws.cache, sheet_row)
                sst_total += row_sst_count
            end
        end
        ws.sst_count = sst_total
        ws.unhandled_attributes = isempty(unhandled_attributes) ? nothing : unhandled_attributes
    end
   
    streamed_rows = stream_rows(lznode, chunksize)
    mylock = ReentrantLock()
   
    @sync for _ in 1:nthreads
        Threads.@spawn begin
            for rows in streamed_rows
                processed = filter!(!isnothing, [process_row(row, handled_attributes, ws, sst_pfx, mylock) for row in rows])
                put!(sheet_rows, processed)
            end
        end
    end
   
    close(sheet_rows)

    wait(consumer)

    ws.cache.is_full = true
end

# Materialise specific rows from a worksheet.xml file into SheetRows
# (faster than using eachrow which materialises every row).
function match_rows(ws::Worksheet, rows_to_match::Vector{Int})::Vector{SheetRow}
    matched_rows=Vector{SheetRow}()
    sst_pfx = get_sst_prefix(ws)

    sort!(rows_to_match)
    i=1
    l=length(rows_to_match)
    
    target_file = get_relationship_target_by_id("xl", get_workbook(ws), ws.relationship_id)
    lznode = open_internal_file_stream(get_xlsxfile(ws), target_file)

    n = XML.next(lznode)
    mylock=ReentrantLock()
    while !isnothing(n)
        if localname(n) == "row" # find each row
            atts = XML.attributes(n)
            if !isnothing(atts)
                row_num = haskey(atts, "r") ? parse(Int, atts["r"]) : nothing
            end
            row_num === nothing && throw(XLSXError("Row without 'r' attribute encountered in worksheet $(ws.name)."))
            if !isnothing(row_num) && row_num == rows_to_match[i] # process matching rows into SheetRows
                current_row_ht = haskey(atts, "ht") ? parse(Float64, atts["ht"]) : nothing

                # Process cells
                rowcells = Dict{Int,Cell}()
                n, _ = get_rowcells!(rowcells, n, ws, sst_pfx; mylock)

                sheetrow = SheetRow(ws, row_num, current_row_ht, rowcells)
                push!(matched_rows, sheetrow)
                i+=1
                i>l && break # stop once all rows matched
                continue
            end
        end
        n=XML.next(n)
    end

    return matched_rows
end
