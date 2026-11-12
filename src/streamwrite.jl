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

    # the writing task: it alone touches the zip until the sheet is closed
    pipe = Channel{Vector{UInt8}}(2)
    zip = s.zip
    task = Threads.@spawn for bytes in pipe
        write(zip, bytes)
    end
    bind(pipe, task)

    pfx = _prefix_colon(sheet)
    ncols = isnothing(cols) ? 1 : max(1, length(cols))
    sh = StreamSheet(sheet, s.zip, s.shared_strings, _RowWriter(sheet, pfx, ncols), pfx, IOBuffer(),
                     cols, column_styles, table, 0, 0, pipe, task, true)
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

# Write the next row, whose `n` values are `value(1)`, ..., `value(n)`. Column number
# formats apply to every row but the header.
function _write_row!(sh::StreamSheet, value::F, n::Int; styled::Bool=true) where {F}
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
    write(io, "\">\n")
    written = false
    for c in 1:n
        written |= _write_stream_cell!(sh, CellRef(r, c), _stream_value(value(c)), styled)
    end
    if written
        write(io, w.row_close)
    else   # a row with no cells is left out, as Excel leaves it out
        truncate(io, start)
        seek(io, start)
    end
    sh.row = r
    sh.rows_in_chunk += 1
    sh.rows_in_chunk >= ROW_CHUNKSIZE && _flush_chunk!(sh)
    return nothing
end

# Values as `setdata!` writes them.
_stream_value(::Union{Missing,Nothing}) = missing
_stream_value(v::Union{Bool,String,Dates.Date,Dates.DateTime,Dates.Time}) = v
_stream_value(v::Integer) = Int64(v)
_stream_value(v::Real) = (f = Float64(v); isfinite(f) ? f : string(v))   # Excel has no NaN or Inf
_stream_value(v::AbstractString) = String(v)
_stream_value(v) = string(v)

# Write one cell; `false` if the value leaves it empty.
function _write_stream_cell!(sh::StreamSheet, ref::CellRef, v, styled::Bool)::Bool
    v === missing && return false
    io, w = sh.io, sh.w
    style = styled ? _stream_style(sh, column_number(ref), v) : UInt32(0)
    if v isa String && !sh.shared_strings
        s = strip_illegal_chars(v)
        isempty(s) && return false
        _write_cell_open!(io, w, ref, "inlineStr", style)
        write(io, ">\n        <", sh.pfx, "is>")
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
function _stream_table!(s::StreamWriter, name::AbstractString, x; kw...)
    Tables.istable(x) && !Tables.columnaccess(x) && return _stream_rows!(s, name, x; kw...)
    return _stream_columns!(s, name, _table_to_arrays(x)...; kw...)
end

# `data` is a vector of columns, as `writetable!` takes it.
function _stream_columns!(s::StreamWriter, sheetname::AbstractString, data, columnnames; kw...)
    nrows = _check_table_shape(data, columnnames)
    sh = _stream_addsheet!(s, sheetname, nrows; columns=columnnames, kw...)
    ncols = length(data)
    for r in 1:nrows
        _write_row!(sh, c -> data[c][r], ncols)
    end
    return nothing
end

# A Tables.jl table read a row at a time, so it need not be in memory.
function _stream_rows!(s::StreamWriter, sheetname::AbstractString, x; kw...)
    rows = Tables.rows(x)
    names = _row_table_names(rows)
    nrows = Base.IteratorSize(rows) isa Union{Base.HasLength,Base.HasShape} ? length(rows) : nothing
    sh = nothing
    for row in rows
        if isnothing(sh)
            names = something(names, collect(Tables.columnnames(row)))
            sh = _stream_addsheet!(s, sheetname, nrows; columns=names, kw...)
        end
        _write_row!(sh, i -> Tables.getcolumn(row, i), length(names))
    end
    if isnothing(sh)   # no rows: just the header, if the names are known
        isnothing(names) && throw(XLSXError("Can't write a table with no rows and no column names."))
        _stream_addsheet!(s, sheetname, 0; columns=names, kw...)
    end
    return nothing
end

function _row_table_names(rows)
    sch = Tables.schema(rows)
    return isnothing(sch) ? nothing : collect(sch.names)
end
