#
# Streaming writes: rows go straight into the file a chunk at a time, never into a
# cell cache, so memory stays bounded however many rows are written.
#
# The package's small parts (styles, workbook, relationships, tables) are kept in an
# ordinary `XLSXFile` made from the blank template, with a stub for each sheet. A
# sheet's XML up to `<sheetData>` is written when it is added, then its rows, then
# the rest of it when it is closed. The other parts are written last, when what
# the rows needed (date styles, shared strings, a table's range) is known.
#
# Rows are formatted by the caller's task and compressed by another, so the two
# overlap. Compression takes about two thirds of the time and runs on one thread,
# so more formatting threads would not help.
#

"""
    StreamSheet

A worksheet being written by [`XLSX.streamxlsx`](@ref), returned by
`XLSX.addsheet!(::StreamWriter, ...)`. Write rows to it in order with
[`XLSX.writerow!`](@ref).
"""
mutable struct StreamSheet
    sheet::Worksheet
    zip::ZipArchives.ZipWriter
    shared_strings::Bool
    w::_RowWriter
    pfx::String
    io::IOBuffer
    columns::Union{Nothing,Vector{String}}
    column_styles::Vector{UInt32}   # by column; 0 for none
    table::Union{Nothing,Tuple{String,Union{Nothing,String}}}   # name and style
    row::Int                        # the last row written
    rows_in_chunk::Int
    pipe::Channel{Vector{UInt8}}
    free::Channel{Vector{UInt8}}    # buffers the writing task is done with, to reuse
    task::Task
    isopen::Bool
end

"""
    StreamWriter

A workbook being written by [`XLSX.streamxlsx`](@ref). Add sheets to it with
`XLSX.addsheet!`.
"""
mutable struct StreamWriter
    xf::XLSXFile
    zip::ZipArchives.ZipWriter
    shared_strings::Bool
    current::Union{Nothing,StreamSheet}
    streamed::Vector{Int}           # numbers of the sheets written
end

Base.show(io::IO, s::StreamWriter) = print(io, "StreamWriter(", length(s.streamed), " sheets written)")
Base.show(io::IO, sh::StreamSheet) = print(io, "StreamSheet(\"", sh.sheet.name, "\", ", sh.row, " rows written", sh.isopen ? ")" : ", closed)")

"""
    streamxlsx(f, dest; overwrite=false, shared_strings=false)

Write a new workbook to `dest`, a file name or an `IO`, one row at a time, using
memory bounded by a chunk of rows however many are written. `f` is called with a
[`StreamWriter`](@ref XLSX.StreamWriter): add sheets to it with
[`XLSX.addsheet!`](@ref) and write each sheet's rows with [`XLSX.writerow!`](@ref).
The file is complete when `f` returns.

Strings are written inline in each cell by default, which keeps memory bounded and is
fastest. With `shared_strings=true` they go in a shared string table, as Excel writes
them, which makes a file with many repeated strings a little smaller but holds every
distinct string in memory until the file is finished. Excel reads both, and stores
strings in its own shared table when it saves the file.

Values are written as by `sheet[ref] = value`: numbers, strings, `Bool`s, and dates and
times (with a default date or time format); `missing` and `nothing` leave a cell
empty, and other types are written as their `string`. If `f` throws, a file written
to a path is deleted.

A streamed sheet is written in one pass, so it can't be read or changed afterwards
within `f`, and has only what `addsheet!` sets: column widths, number formats, a
frozen header row and an Excel Table. See [`XLSX.writetable`](@ref) with
`streaming=true` to stream a whole table.

Returns the path of the file written, or `nothing` when writing to an `IO`.

# Example
```julia
julia> XLSX.streamxlsx("big.xlsx") do s
           sh = XLSX.addsheet!(s, "Data"; columns=["id", "name", "when"], freeze_header=true)
           for i in 1:1_000_000
               XLSX.writerow!(sh, (i, "item \$i", Date(2026, 1, 1) + Day(i % 365)))
           end
       end
```
"""
function streamxlsx(f::Function, dest::Union{AbstractString,IO}; overwrite::Bool=false, shared_strings::Bool=false)
    _check_overwrite(dest, overwrite)
    xf = open_empty_template()
    finished = false
    try
        ZipArchives.ZipWriter(dest) do zip
            s = StreamWriter(xf, zip, shared_strings, nothing, Int[])
            try
                f(s)
                _close_current!(s)
            finally
                _stop_current!(s)
            end
            _write_package!(zip, xf; streamed=s.streamed)
        end
        finished = true
    finally
        # a half-written file is no use; an `IO` is the caller's to clear
        !finished && dest isa AbstractString && rm(dest; force=true)
    end
    return dest isa AbstractString ? abspath(dest) : nothing
end

"""
    addsheet!(s::StreamWriter, [name=""]; columns, column_widths, column_formats,
              freeze_header=false, as_table=false, table_name="", table_style=nothing) -> StreamSheet

Add a worksheet to a workbook being written by [`XLSX.streamxlsx`](@ref), closing the
sheet added before it: sheets are written one after another. `name` defaults to the
next free `SheetN`.

- `columns`: column names, written as the first row.
- `column_widths`: one width for every column (needs `columns`), or a vector of widths
  by column, `nothing` leaving a column's width as it is. Widths are in characters, as
  [`XLSX.setColumnWidth`](@ref) takes them.
- `column_formats`: a `Dict` from column (number, or name in `columns`) to a number
  format, as [`XLSX.setFormat`](@ref) takes it, for every cell below the header.
- `freeze_header`: freeze the first row, as [`XLSX.freezePanes`](@ref) does.
- `as_table`, `table_name`, `table_style`: make the rows an Excel Table when the sheet
  is closed, as [`XLSX.addtable!`](@ref) does. Needs `columns` and at least one row
  below them.
"""
function addsheet!(s::StreamWriter, name::AbstractString=""; kw...)::StreamSheet
    return _stream_addsheet!(s, name, nothing; kw...)
end

"""
    addsheet!(s::StreamWriter, name, table; column_widths, column_formats,
              freeze_header=false, as_table=false, table_name="", table_style=nothing) -> StreamSheet

Add a worksheet to a workbook being written by [`XLSX.streamxlsx`](@ref) and write
`table`, a Tables.jl table such as a `DataFrame` or a vector of `NamedTuple`s, to it:
its column names as the header, then its rows. The keywords are those of
`addsheet!(s, name)`, but for `columns`, which come from the table. A table read a row
at a time is never held in memory as a whole.

The sheet is still open when `addsheet!` returns, so more rows can follow, with
[`XLSX.writerow!`](@ref) or [`XLSX.writerows!`](@ref). For that reason the sheet's
`<dimension>`, the range Excel records as used, is left out: Excel and XLSX.jl work
it out from the rows. To write a whole file of tables, [`XLSX.writetable`](@ref) with
`streaming=true` does the same, with the dimension.

# Example
```julia
julia> XLSX.streamxlsx("big.xlsx") do s
           XLSX.addsheet!(s, "Data", df; freeze_header=true, as_table=true)
           XLSX.addsheet!(s, "Summary", summary_df)
       end
```
"""
function addsheet!(s::StreamWriter, name::AbstractString, table; kw...)::StreamSheet
    haskey(kw, :columns) && throw(XLSXError("`columns` can't be given with a table: its column names are the header. To write a table below other column names, use `addsheet!(s, name; columns)` and then `writerows!`."))
    return _stream_rows!(s, name, table, false; kw...)
end

# `nrows`, when known, is the number of rows below the header, written as the
# sheet's `<dimension>` (which comes before the rows).
function _stream_addsheet!(s::StreamWriter, name::AbstractString, nrows::Union{Nothing,Int};
        columns=nothing, column_widths=nothing, column_formats=nothing,
        freeze_header::Bool=false, as_table::Bool=false, table_name::AbstractString="",
        table_style::Union{AbstractString,Nothing}=nothing)::StreamSheet
    _close_current!(s)
    xf = s.xf
    if isempty(s.streamed)   # the template's sheet
        sheet = xf[1]
        name == "" || renamesheet!(sheet, name)
    else
        sheet = addsheet!(xf, name)
    end
    cols = isnothing(columns) ? nothing : String[string(c) for c in columns]
    !isnothing(cols) && length(cols) > EXCEL_MAX_COLS &&
        throw(XLSXError("`columns` has $(length(cols)) names, but Excel only supports up to $EXCEL_MAX_COLS columns."))

    table = nothing
    if as_table
        isnothing(cols) && throw(XLSXError("`as_table=true` needs `columns`: an Excel Table needs a header row."))
        _check_table_columns(cols)
        dummy = CellRange(CellRef(1, 1), CellRef(2, length(cols)))   # the rows aren't known yet
        table = (_checked_new_table_name(sheet, table_name, dummy), isnothing(table_style) ? nothing : String(table_style))
    end

    widths, formats = _column_settings(cols, column_widths, column_formats)
    for (c, width) in widths
        setColumnWidth(sheet, CellRange(CellRef(1, c), CellRef(1, c)); width)
    end
    column_styles = zeros(UInt32, isempty(formats) ? 0 : maximum(first, formats))
    for (c, fmt) in formats
        column_styles[c] = id(get_num_style_index(sheet, _format_id(get_workbook(sheet), fmt)))
    end
    freeze_header && freezePanes(sheet; nrows=1)

    dimension = isnothing(nrows) || isnothing(cols) ? nothing :
        CellRange(CellRef(1, 1), CellRef(nrows + 1, max(1, length(cols))))
    head, _ = _sheet_xml_frame(sheet, dimension)
    ZipArchives.zip_newfile(s.zip, get_worksheet_internal_file(sheet); compress=true)
    write(s.zip, head)

    # The writing task: it alone touches the zip until the sheet is closed. Each chunk's
    # buffer comes back by `free` once written, so the next chunk is formatted into it
    # rather than into a new buffer grown from empty. At most four are ever made (two
    # waiting in `pipe`, one being written, one being filled), so `free` never blocks.
    pipe = Channel{Vector{UInt8}}(2)
    free = Channel{Vector{UInt8}}(4)
    zip = s.zip
    task = Threads.@spawn for bytes in pipe
        write(zip, bytes)
        put!(free, bytes)
    end
    bind(pipe, task)

    pfx = _prefix_colon(sheet)
    ncols = isnothing(cols) ? 1 : max(1, length(cols))
    sh = StreamSheet(sheet, s.zip, s.shared_strings, _RowWriter(sheet, pfx, ncols), pfx, IOBuffer(),
                     cols, column_styles, table, 0, 0, pipe, free, task, true)
    s.current = sh
    isnothing(cols) || _write_row!(sh, i -> cols[i], length(cols); styled=false)
    return sh
end

# `column_widths` and `column_formats` as `column => setting` pairs.
function _column_settings(columns, column_widths, column_formats)
    widths = Pair{Int,Float64}[]
    if column_widths isa Real
        isnothing(columns) && throw(XLSXError("A single `column_widths` applies to every column, so it needs `columns`; give a vector of widths instead."))
        append!(widths, c => Float64(column_widths) for c in eachindex(columns))
    elseif column_widths isa AbstractVector
        append!(widths, c => Float64(w) for (c, w) in enumerate(column_widths) if !isnothing(w))
    elseif !isnothing(column_widths)
        throw(XLSXError("`column_widths` must be a number or a vector of numbers, got `$(typeof(column_widths))`."))
    end

    formats = Pair{Int,String}[]
    if column_formats isa AbstractDict
        for (k, fmt) in column_formats
            push!(formats, _column_index(columns, k) => String(fmt))
        end
    elseif !isnothing(column_formats)
        throw(XLSXError("`column_formats` must be a `Dict` from columns to number formats, got `$(typeof(column_formats))`."))
    end
    return widths, formats
end

function _column_index(columns, k)::Int
    if k isa Integer
        1 <= k <= EXCEL_MAX_COLS || throw(XLSXError("Column $k is outside Excel's columns 1 to $EXCEL_MAX_COLS."))
        return Int(k)
    end
    isnothing(columns) && throw(XLSXError("Column `$k` is named, but there are no column names to find it in."))
    c = findfirst(==(string(k)), columns)
    isnothing(c) && throw(XLSXError("There is no column named `$k`."))
    return c
end

"""
    writerow!(sh::StreamSheet, row)

Write `row`, a tuple, `NamedTuple` or vector of values, as the next row of a sheet
being written by [`XLSX.streamxlsx`](@ref), starting from column A. Values are
written as described there; a row of nothing but empty values leaves the row empty.
A `NamedTuple`'s values are written in order: its names are not matched to the
sheet's `columns`.
"""
writerow!(sh::StreamSheet, row::Union{Tuple,NamedTuple,AbstractVector}) = _write_row!(sh, i -> row[i], length(row))
writerow!(::StreamSheet, row) = throw(XLSXError("`writerow!` takes a row as a tuple, `NamedTuple` or vector, got `$(typeof(row))`."))

"""
    writerows!(sh::StreamSheet, table)

Write every row of `table`, a Tables.jl table such as a `DataFrame` or a vector of
`NamedTuple`s, below the rows already written to a sheet being written by
[`XLSX.streamxlsx`](@ref), as [`XLSX.writerow!`](@ref) writes each one. A table read a
row at a time (such as a database query or a `CSV.Rows`) is never held in memory as a
whole.

The table's column names are not written, nor matched to the sheet's `columns`: its
columns are written in order from column A. To write a table with its names as the
header, use `XLSX.addsheet!(s, name, table)`. To write a table in pieces, such as the
chunks of a CSV file read with `CSV.Chunks`, call `writerows!` for each.

# Example
```julia
julia> XLSX.streamxlsx("big.xlsx") do s
           sh = XLSX.addsheet!(s, "Data"; columns=["id", "value"])
           for chunk in CSV.Chunks("big.csv"; ntasks=100)
               XLSX.writerows!(sh, chunk)
           end
       end
```
"""
function writerows!(sh::StreamSheet, table)
    rows = _table_rows(table)
    sch, rows = _rows_schema(rows)
    isnothing(sch) || _write_rows!(sh, rows, sch)
    return nothing
end

# Write the next row, whose `n` values are `value(1)`, ..., `value(n)`. Column number
# formats apply to every row but the header.
function _write_row!(sh::StreamSheet, value::F, n::Int; styled::Bool=true) where {F}
    r, start = _open_row!(sh, n)
    written = false
    for c in 1:n
        written |= _write_stream_cell!(sh, CellRef(r, c), _as_cell_value(value(c)), styled)
    end
    _close_row!(sh, r, start, written)
    return nothing
end

# Start row `sh.row + 1`, of `n` values: its number, and where it starts in `sh.io`.
function _open_row!(sh::StreamSheet, n::Int)
    sh.isopen || throw(XLSXError("Can't write to sheet `$(sh.sheet.name)`: it is closed. A streamed sheet is closed when the next sheet is added, and when `streamxlsx` returns."))
    r = sh.row + 1
    r > EXCEL_MAX_ROWS && throw(XLSXError("Can't write row $r: Excel only supports up to $EXCEL_MAX_ROWS rows."))
    n > EXCEL_MAX_COLS && throw(XLSXError("Row $r has $n values, but Excel only supports up to $EXCEL_MAX_COLS columns."))
    width = isnothing(sh.columns) ? n : length(sh.columns)
    n > width && throw(XLSXError("Row $r has $n values, more than the sheet's $width columns."))

    io, w = sh.io, sh.w
    start = position(io)
    write(io, w.row_open)
    _write_number(io, w, r)
    write(io, "\" spans=\"1:")
    _write_number(io, w, width)
    write(io, "\">")
    return r, start
end

# Finish row `r`, begun at `start`; `written` if any of its cells were.
function _close_row!(sh::StreamSheet, r::Int, start::Int, written::Bool)
    io = sh.io
    if written
        write(io, sh.w.row_close)
    else   # a row with no cells is left out, as Excel leaves it out
        truncate(io, start)
        seek(io, start)
    end
    sh.row = r
    sh.rows_in_chunk += 1
    sh.rows_in_chunk >= ROW_CHUNKSIZE && _flush_chunk!(sh)
    return nothing
end

# Write one cell; `false` if the value leaves it empty.
function _write_stream_cell!(sh::StreamSheet, ref::CellRef, v, styled::Bool)::Bool
    v === missing && return false
    v isa CellImage && throw(XLSXError("Can't put a picture in cell $ref: a streamed sheet holds values only. Write it with `writetable` without `streaming=true`."))
    io, w = sh.io, sh.w
    style = styled ? _stream_style(sh, column_number(ref), v) : UInt32(0)
    if v isa String && !sh.shared_strings
        s = strip_illegal_chars(v)
        isempty(s) && return false
        _write_cell_open!(io, w, ref, "inlineStr", style)
        write(io, "><", sh.pfx, "is>")
        write_plain_t(io, s, sh.pfx)
        write(io, "</", sh.pfx, "is>", w.c_close)
    else
        t, val = xlsx_encode(sh.sheet, v)
        t == CT_EMPTY && return false
        _write_cell_open!(io, w, ref, encode(t), style)
        write(io, ">")
        _write_cell_v!(io, w, t, val)
        write(io, w.c_close)
    end
    return true
end

# A column's number format, or else a date or time's default one.
function _stream_style(sh::StreamSheet, col::Int, v)::UInt32
    col <= length(sh.column_styles) && sh.column_styles[col] != 0 && return sh.column_styles[col]
    v isa Union{Dates.Date,Dates.DateTime,Dates.Time} || return UInt32(0)
    return id(default_cell_format(sh.sheet, v))   # cached by the workbook
end

function _flush_chunk!(sh::StreamSheet)
    sh.rows_in_chunk = 0
    position(sh.io) == 0 && return nothing
    bytes = take!(sh.io)
    if isopen(sh.pipe)
        put!(sh.pipe, bytes)
    else
        wait(sh.task)   # the writing task failed: rethrow its error
    end
    # the next chunk goes in a buffer already written, if one is back
    isready(sh.free) && (sh.io = IOBuffer(take!(sh.free); read=false, write=true, truncate=true))
    return nothing
end

# Finish the sheet being written, if any: its last rows, its table and the rest of
# its XML.
function _close_current!(s::StreamWriter)
    sh = s.current
    isnothing(sh) && return nothing
    _flush_chunk!(sh)
    sh.isopen = false
    close(sh.pipe)
    wait(sh.task)
    s.current = nothing

    if !isnothing(sh.table)
        name, style = sh.table
        sh.row < 2 && throw(XLSXError("Table `$name` needs at least one row below its header: Excel does not support header-only tables."))
        ref = CellRange(CellRef(1, 1), CellRef(sh.row, length(sh.columns)))
        _add_table_part!(sh.sheet, ref, name, sh.columns, style, false)
    end
    _, tail = _sheet_xml_frame(sh.sheet, nothing)
    write(s.zip, tail)
    push!(s.streamed, findfirst(==(sh.sheet), get_workbook(s.xf).sheets))
    return nothing
end

# After an error: stop the sheet's writing task, so the zip can be closed.
function _stop_current!(s::StreamWriter)
    sh = s.current
    isnothing(sh) && return nothing
    sh.isopen = false
    close(sh.pipe)
    istaskfailed(sh.task) || wait(sh.task)
    s.current = nothing
    return nothing
end

#
# Streaming a whole table, for `writetable(...; streaming=true)`.
#

# Stream each `(sheetname, table)` of `sheets` to a sheet of its own. A table is a
# `(data, columnnames)` tuple or a Tables.jl table. With `table_name_from_sheet`,
# each Excel Table (with `as_table`) is named after its sheet, as `writetable` names
# them when writing several sheets.
function _stream_sheets(filename::Union{AbstractString,IO}, sheets; overwrite::Bool, as_table::Bool=false,
                        shared_strings::Bool=false, table_name_from_sheet::Bool=false, kw...)
    return streamxlsx(filename; overwrite, shared_strings) do s
        for (sheetname, x) in sheets
            _close_current!(s)   # so the sheet's table is known when naming the next
            name = string(sheetname)
            tkw = as_table && table_name_from_sheet ?
                (; table_name=_table_name_from_sheet(get_workbook(s.xf), s.xf[1], name)) : (;)
            _stream_table!(s, name, x; as_table, tkw..., kw...)
        end
    end
end

_stream_table!(s::StreamWriter, name::AbstractString, x::Tuple; kw...) = _stream_columns!(s, name, x...; kw...)
_stream_table!(s::StreamWriter, name::AbstractString, x; kw...) = _stream_rows!(s, name, x, true; kw...)

# `data` is a vector of columns, as `writetable!` takes it.
function _stream_columns!(s::StreamWriter, sheetname::AbstractString, data, columnnames; kw...)
    nrows = _check_table_shape(data, columnnames)
    sh = _stream_addsheet!(s, sheetname, nrows; columns=columnnames, kw...)
    ncols = length(data)
    if ncols <= 100
        # As a column table (its names don't matter: the header is written), so that
        # `Tables.eachcolumn` reads each value with its type known, rather than boxed
        # from a `Vector{Any}` of columns. Beyond 100 columns it loops, so gains nothing.
        writerows!(sh, NamedTuple{ntuple(c -> Symbol(:c, c), ncols)}(Tuple(_as_vector(c) for c in data)))
    else
        for r in 1:nrows
            _write_row!(sh, c -> data[c][r], ncols)
        end
    end
    return nothing
end

# A Tables.jl table, its names as the header, read a row at a time so it need not be
# in memory. With `dimension`, the sheet's `<dimension>` is written when the table's
# length is known, so no more rows may follow.
function _stream_rows!(s::StreamWriter, sheetname::AbstractString, x, dimension::Bool; kw...)::StreamSheet
    rows = _table_rows(x)
    nrows = dimension && Base.IteratorSize(rows) isa Union{Base.HasLength,Base.HasShape} ? length(rows) : nothing
    !isnothing(nrows) && nrows > EXCEL_MAX_ROWS - 1 &&
        throw(XLSXError("The table has $nrows rows, but Excel only supports up to $(EXCEL_MAX_ROWS - 1) below its header."))
    sch, rows = _rows_schema(rows)
    isnothing(sch) && throw(XLSXError("Can't write a table with no rows and no column names."))
    names = String[string(n) for n in sch.names]
    isempty(names) && throw(XLSXError("Can't write a table with no columns."))
    sh = _stream_addsheet!(s, sheetname, nrows; columns=names, kw...)
    _write_rows!(sh, rows, sch)
    return sh
end

# The rows of a Tables.jl table. A table of columns is read through a `NamedTuple`
# of them (the same vectors, not copies), whose rows `Tables.eachcolumn` reads
# with each value's type known.
function _table_rows(x)
    Tables.istable(x) || throw(XLSXError("$(typeof(x)) does not implement Tables.jl interface." * _shape_hint(x)))
    if Tables.columnaccess(x)
        cols = Tables.columns(x)
        sch = Tables.schema(cols)
        # beyond this width `eachcolumn` loops rather than unrolls, so gains nothing
        !isnothing(sch) && length(sch.names) <= 100 && return Tables.rows(Tables.columntable(cols))
    end
    return Tables.rows(x)
end

# `rows`' schema, and the rows to write. Rows with no schema take their column names
# from the first row (which is still written); with no rows either, the schema is
# `nothing`.
function _rows_schema(rows)
    sch = Tables.schema(rows)
    isnothing(sch) || return sch, rows
    next = iterate(rows)
    isnothing(next) && return nothing, rows
    row, state = next
    return Tables.Schema(Tables.columnnames(row), nothing), Iterators.flatten(((row,), Iterators.rest(rows, state)))
end

function _write_rows!(sh::StreamSheet, rows, sch::Tables.Schema)
    n = length(sch.names)
    for row in rows
        _write_table_row!(sh, sch, row, n)
    end
    return nothing
end

# One row of a table, each value read by `Tables.eachcolumn`: for a schema of known
# width it unrolls the loop over the columns, so each value's type is known where
# it is written, rather than boxed.
function _write_table_row!(sh::StreamSheet, sch::Tables.Schema, row, n::Int)
    r, start = _open_row!(sh, n)
    written = Ref(false)
    Tables.eachcolumn(sch, row) do v, c, _
        written[] |= _write_stream_cell!(sh, CellRef(r, c), _as_cell_value(v), true)
    end
    _close_row!(sh, r, start, written[])
    return nothing
end
