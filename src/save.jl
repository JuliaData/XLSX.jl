#
# Saving: writing an XLSXFile to a file or IO, and the XML of its worksheets,
# shared strings, workbook and document properties.
#


"""
    savexlsx(f::XLSXFile)

Save an `XLSXFile` instance back to the file from which it was opened (given in `f.source`), 
overwriting original content.

A new `XLSXFile` created with `XLSX.newxlsx` (or using `openxlsx` without specifying a filename) will 
have `source` set to `"blank.xlsx"` and cannot be saved with this function. Use [`writexlsx`](@ref) instead 
to specify a new file name for the saved file.

An `XLSXFile` created from a native Excel template (`.xltx`) file or a macro-enabled template (`.xltm`) file
will be saved as standard worksheet (`.xlsx`) or macro-enabled worksheet (`.xlsm`) respectively, regardless 
of the extension of the original template file.

If `f` was opened from an `IO`, the `IO`'s content is replaced: it is written from the start and
truncated to the new workbook's length. The `IO` must be writable and seekable, such as an
`IOBuffer(bytes; read=true, write=true)`.

Returns the filepath of the written file if a filename is supplied, or `nothing` if writing to an `IO`.

"""
function savexlsx(f::XLSXFile)
    if f.source isa IO
        _reset_io_for_save!(f.source)
    elseif isa(f.source, AbstractString)
        if f.source == "blank.xlsx"
            throw(XLSXError("Can't save to a blank `XLSXFile` instance. Use `writexlsx` instead to specify a file name."))
        elseif f.template_type != NotATemplate
            if isa(f.source, AbstractString)
                ext = f.template_type == XLTMTemplate ? ".xlsm" : ".xlsx"
                f.source = splitext(f.source)[1] * ext
            end
        end
    end
    return writexlsx(f.source, f; overwrite=true)
end

# Empty an `IO` source so that saving replaces its content, as saving replaces a file.
function _reset_io_for_save!(io::IO)
    try
        seekstart(io)
        truncate(io, 0)
    catch e
        # The error depends on the kind of `IO` (read-only buffer or file, pipe, ...);
        # all mean it can't be rewritten. Let an interrupt through, and keep the cause.
        e isa InterruptException && rethrow()
        throw(XLSXError("Can't save back to the source `IO`: it must be writable and seekable. Use `writexlsx` to write to another `IO` or file. ($(sprint(showerror, e)))"))
    end
    return io
end



"""
    writexlsx(output_source::Union{AbstractString,IO}, xf::XLSXFile; [overwrite=false])

Write an XLSXFile given by `xf` to the IO or filepath `output_source`.

The source attribute of the `XLSXFile` will be updated to the `output_source` if it is a filepath.

Returns the filepath of the written file if a filename is supplied, or `nothing` if writing to an `IO`.

If `overwrite=true`, `output_source` (when a filepath) will be overwritten if it exists.

When `output_source` is an `IO`, it should be at its start (e.g. a new `IOBuffer()`), and
`overwrite` has no effect. To save a workbook back to the `IO` it was opened from, replacing
its content, use [`savexlsx`](@ref).

See also [`savexlsx`](@ref).

!!! note

    When XLSX.jl reads strict (ISO/IEC 29500) XLSX files, it converts them eagerly on 
    to transitional (ECMA 376) format. On write, XLSX.jl will always write in the transitional 
    format, which is the Excel default. Excel itself can convert between strict and 
    transitional formats. Use Excel directly to convert a transitional file to strict format, 
    if needed.

"""
function writexlsx(output_source::Union{AbstractString,IO}, xf::XLSXFile; overwrite::Bool=false)

    _require_writable(xf, "write the workbook")
    if !all(values(xf.files))
        throw(XLSXError("Some internal files were not loaded into memory for unknown reasons."))
    end
    _check_overwrite(output_source, overwrite)

    ZipArchives.ZipWriter(output_source) do xlsx
        _write_package!(xlsx, xf)
    end

    if !(output_source isa IO)
        (xf.source = output_source) # update source if output_source is a file path
        return abspath(xf.source)
    else
        return nothing
    end

end

_check_overwrite(::IO, ::Bool) = nothing
function _check_overwrite(output_source::AbstractString, overwrite::Bool)
    !overwrite && isfile(output_source) && throw(XLSXError("Output file $output_source already exists."))
    return nothing
end

# Write every part of `xf` into the zip `xlsx`, except the worksheets numbered in
# `streamed`, which `streamxlsx` has already written.
function _write_package!(xlsx::ZipArchives.ZipWriter, xf::XLSXFile; streamed=())
    update_workbook_xml!(xf)
    update_app_xml!(xf)

    wb = get_workbook(xf)

    # write XML files not in cache
    for f in keys(xf.files)
        if !occursin(r"^xl/worksheets/[^/]+\.xml$|^xl/sharedStrings\.xml$", f)
            ZipArchives.zip_newfile(xlsx, f; compress=true)
            xml_str = XML.write(xf.data[f]; indentsize=0)
            write(xlsx, xml_str)
        end
    end
    # write worksheet files from cache (cache must be enabled in write mode)
    for sheet_no in 1:sheetcount(wb)
        if sheet_no in streamed || is_chartsheet(wb, getsheet(wb, sheet_no).name)
            continue
        end
        pieces = update_single_sheet!(wb, sheet_no, true)
        ZipArchives.zip_newfile(xlsx, get_worksheet_internal_file(getsheet(wb, sheet_no)); compress=true)
        foreach(p -> write(xlsx, p), pieces)
    end

    # write binary files
    for f in keys(xf.binary_data)
        ZipArchives.zip_newfile(xlsx, f; compress=true)
        write(xlsx, xf.binary_data[f])
    end

    # write sharedString table
    if !isempty(get_sst(wb))
        ZipArchives.zip_newfile(xlsx, "xl/sharedStrings.xml"; compress=true)
        print(xlsx, generate_sst_xml_string(wb))
    end
    return nothing
end

get_worksheet_internal_file(ws::Worksheet) = get_relationship_target_by_id("xl", get_workbook(ws), ws.relationship_id)
function get_worksheet_xml_document(ws::Worksheet)::XML.Node
    xf = get_xlsxfile(ws)
    filename = get_worksheet_internal_file(ws)
    val = get_xml_data(xf, filename)
    return val
end
function set_worksheet_xml_document!(ws::Worksheet, xdoc::XML.Node)
    XML.nodetype(xdoc) != XML.Document && throw(XLSXError("Expected an XML Document node, got $(XML.nodetype(xdoc))."))
    xf = get_xlsxfile(ws)
    filename = get_worksheet_internal_file(ws)
    !haskey(xf.data, filename) && throw(XLSXError("Internal file not found for $(ws.name)."))
    xf.data[filename] = xdoc
end

function generate_sst_xml_string(wb::Workbook)::String
    sst = wb.sst
    !sst.is_loaded && throw(XLSXError("Can't generate XML string from a Shared String Table that is not loaded."))

    pfx = get_prefix("xl/sharedStrings.xml", get_xlsxfile(wb))
    pfx_c = isempty(pfx) ? "" : "$(pfx):"      # "x:" or ""
    c_pfx = isempty(pfx) ? "" : ":$(pfx)"      # ":x" or ""

    sst_total = sum(sheet.sst_count for sheet in wb.sheets)

    buff = IOBuffer()

    print(buff, """<?xml version="1.0" encoding="UTF-8" standalone="yes"?><$(pfx_c)sst count="$sst_total" uniqueCount="$(length(sst))" xmlns$(c_pfx)="http://schemas.openxmlformats.org/spreadsheetml/2006/main">\n""")

    for i in eachindex(sst.unformatted)
        raw = get(sst.rich, i - 1, nothing)
        if raw === nothing
            write_plain_si(buff, sst.unformatted[i], pfx_c)
        else
            print(buff, raw)
        end
    end

    print(buff, "</$(pfx_c)sst>")

    return String(take!(buff))
end

function add_node_formula!(io, ref::CellRef, ws::Worksheet, pfx)
    return add_node_formula!(io, get_formula_from_cache(ws, ref), pfx)
end
add_node_formula!(io, f::CellFormula, pfx) = add_node_formula!(io, f.value, pfx)
function add_node_formula!(io, f::Formula, pfx)
    write(io, "\n        <$(pfx)f")
    
    isnothing(f.type) || isempty(f.type) || write(io, " t=\"", f.type, "\"")
    isnothing(f.ref) || isempty(f.ref) || write(io, " ref=\"", f.ref, "\"")
    
    # Write unhandled attributes
    if !isnothing(f.unhandled)
        for (k, v) in f.unhandled
            write(io, " ", k, "=\"", XML.escape(v; attribute = true), "\"")
        end
    end
    
    write(io, ">", XML.escape(f.formula), "</$(pfx)f>")
end

function add_node_formula!(io, f::FormulaReference, pfx)
    write(io, "\n        <$(pfx)f t=\"shared\"")
    
    if !isnothing(f.unhandled)
        for (k, v) in f.unhandled
            write(io, " ", k, "=\"", XML.escape(v; attribute = true), "\"")
        end
    end
    
    write(io, " si=\"")
    print(io, f.id)
    write(io, "\"></$(pfx)f>")
end

function add_node_formula!(io, f::ReferencedFormula, pfx)
    write(io, "\n        <$(pfx)f t=\"shared\" ref=\"", f.ref, "\"")
    
    if !isnothing(f.unhandled)
        for (k, v) in f.unhandled
            write(io, " ", k, "=\"", XML.escape(v; attribute = true), "\"")
        end
    end
    
    write(io, " si=\"")
    print(io, f.id)
    write(io, "\">", XML.escape(f.formula), "</$(pfx)f>")
end

"""
    update_worksheets_xml!(xl::XLSXFile; full=false)

Update worksheet files held in `xf.data` whenever needed. These files had row and cell data 
stripped out when they were read because, when needed (ie in write mode), these data are 
stored in the cache. Worksheet xml files are fully reconstructed to include rows and cells 
from the cache only as a transient precursor to writing an xlsx file and these reconstructed 
worksheet xml files are not stored.
"""
function update_worksheets_xml!(xl::XLSXFile; full=false)
    wb = get_workbook(xl)
    for sheet_no in 1:sheetcount(wb)
        if is_chartsheet(wb, getsheet(wb, sheet_no).name)
            continue
        end
        update_single_sheet!(wb, sheet_no, full)
    end
    return nothing
end
function update_single_sheet!(wb::Workbook, sheet_no::Int, full::Bool)::Union{Nothing,Vector{Vector{UInt8}}}
    sheet = getsheet(wb, sheet_no)
    if full # need to reconstruct row and cell data from cache
        # The sheet in pieces, to be written one after another: what comes before
        # the rows, the rows from the WorksheetCache a chunk at a time, and what
        # comes after. Joining them first would copy the whole sheet again.
        head, tail = _sheet_xml_frame(sheet, get_dimension(sheet))
        pieces = [head]
        append!(pieces, get_cache_rows(sheet, _prefix_colon(sheet)))
        push!(pieces, tail)
        return pieces
    else
        set_worksheet_xml_document!(sheet, _worksheet_doc_copy(sheet)) # no need to save full reconstructed data
        return nothing
    end
end

_prefix_colon(sheet::Worksheet) = (pfx = get_prefix(sheet); pfx == "" ? pfx : pfx * ":")

# A copy of the sheet's XML (without its rows), checked to be a worksheet.
function _worksheet_doc_copy(sheet::Worksheet)::XML.Node
    doc = copynode(get_worksheet_xml_document(sheet))
    xroot = xml_root_element(doc)

    # check namespace and root node name
    ns_map = get_namespaces(xroot)
    spreadsheet_ns_declared = (SPREADSHEET_NAMESPACE_XPATH_ARG in values(ns_map))
    spreadsheet_ns_declared || throw(XLSXError("Unsupported Spreadsheet XML namespace."))
    localname(xroot) !=   "worksheet" && throw(XLSXError("Malformed Excel file. Expected root node named `worksheet` in worksheet XML file."))
    return doc
end

# The sheet's XML before its rows, up to and including `<sheetData>`, and after them,
# from `</sheetData>`. `dimension` is written as the sheet's `<dimension>`, or for
# `nothing` (a sheet streamed without knowing its size) the element is left out.
function _sheet_xml_frame(sheet::Worksheet, dimension::Union{Nothing,CellRange})::Tuple{Vector{UInt8},Vector{UInt8}}
    doc = _worksheet_doc_copy(sheet)

    # update worksheet dimension
    i, j = get_idces(doc, "worksheet", "dimension")
    if !isnothing(j)
        if isnothing(dimension)
            deleteat!(doc[i].children, j)
        else
            doc[i][j]["ref"] = string(dimension)
        end
    end

    pfx = _prefix_colon(sheet)
    empty_doc = XML.write(doc)

    idx = findfirst("<$(pfx)sheetData/>", empty_doc)
    idx === nothing && throw(XLSXError("<$(pfx)sheetData/> placeholder not found when reconstructing worksheet '$(sheet.name)'"))
    head = IOBuffer()
    print(head, empty_doc[begin:first(idx)-1])

    # create <sheetData> with any attributes
    print(head, "<$(pfx)sheetData")
    i, j = get_idces(doc, "worksheet", "sheetData")
    a = XML.attributes(doc[i][j])
    if !isnothing(a)
        for (k, v) in a
            print(head, " ",k,"=\"",v,"\"")
        end
    end
    print(head, ">\n")
    tail = Vector{UInt8}("  </$(pfx)sheetData>" * SubString(empty_doc, nextind(empty_doc, last(idx))))
    return take!(head), tail
end

const _NO_ROW_ATTRS = Dict{String,String}()

# The sheet's rows in chunks of `chunksize`, numbered in row order:
# (chunk number, [(dimension, row, row's other attributes)]).
function stream_cache_rows(sheet::Worksheet, chunksize::Int)
    Channel{Tuple{Int,Vector{Tuple{CellRange, SheetRow, Dict{String,String}}}}}(1 << 8) do out
        d = get_dimension(sheet)
        uh_all = sheet.unhandled_attributes
        chunk = Tuple{CellRange, SheetRow, Dict{String,String}}[]
        n = 0
        for r in eachrow(sheet)
            uh = isnothing(uh_all) ? _NO_ROW_ATTRS : get(uh_all, row_number(r), _NO_ROW_ATTRS)
            push!(chunk, (d, r, uh))
            if length(chunk) >= chunksize
                put!(out, (n += 1, chunk))
                chunk = Tuple{CellRange, SheetRow, Dict{String,String}}[]
            end
        end
        isempty(chunk) || put!(out, (n += 1, chunk))
    end
end

# What writing a row needs besides the row: the tags, with the namespace prefix
# `pfx` already in them, the sheet's column letters, and a buffer for numbers.
struct _RowWriter
    ws::Worksheet
    pfx::String
    row_open::String
    c_open::String
    v_open::String
    v_close::String
    c_close::String
    row_close::String
    columns::Vector{String}   # column letters, by number
    num::Vector{UInt8}        # where numbers are formatted
end
function _RowWriter(ws::Worksheet, pfx::String, ncols::Int)
    return _RowWriter(ws, pfx, "    <$(pfx)row r=\"", "      <$(pfx)c r=\"", "\n        <$(pfx)v>", "</$(pfx)v>",
                      "\n      </$(pfx)c>\n", "    </$(pfx)row>\n",
                      [encode_column_number(c) for c in 1:ncols], Vector{UInt8}(undef, 64))
end

# Write an integer without making a string. A float is printed: Julia has no public way to
# format one without making a string.
function _write_number(io::IO, w::_RowWriter, n::Integer)
    buf = w.num
    neg = n < 0
    u = neg ? unsigned(-(Int128(n))) : unsigned(n)
    i = length(buf)
    while true
        buf[i] = UInt8('0') + UInt8(u % 10)
        u ÷= 10
        i -= 1
        u == 0 && break
    end
    neg && (buf[i] = UInt8('-'); i -= 1)
    GC.@preserve buf unsafe_write(io, pointer(buf, i + 1), length(buf) - i)
    return nothing
end
_write_number(io::IO, ::_RowWriter, x::Float64) = (print(io, x); nothing)

function _write_cellref(io::IO, w::_RowWriter, c::CellRef)
    col = column_number(c)
    write(io, col <= length(w.columns) ? w.columns[col] : encode_column_number(col))
    _write_number(io, w, row_number(c))
    return nothing
end

# The sheet's rows as XML, in chunks in row order. The chunks are written in
# parallel, each into a buffer of its own.
function get_cache_rows(sheet::Worksheet, pfx::String)::Vector{Vector{UInt8}}
    cache = sheet.cache
    isnothing(cache) && throw(XLSXError("Worksheet '$(sheet.name)' cannot be written - it has no cache."))
    chunksize = ROW_CHUNKSIZE
    nworkers = clamp(cld(length(cache.rows_in_cache), chunksize), 1, max_threads())
    ncols = column_number(get_dimension(sheet).stop)
    chunks = stream_cache_rows(sheet, chunksize)
    written = Channel{Tuple{Int,Vector{UInt8}}}(1 << 8)
    workers = @async try   # `written` is read below while they run, and closed when they stop
        @sync for _ in 1:nworkers
            Threads.@spawn begin
                w = _RowWriter(sheet, pfx, ncols)
                io = IOBuffer()
                column_indexes = Int[]
                for (n, rows) in chunks
                    for row in rows
                        process_cache_row!(io, column_indexes, row, w)
                    end
                    put!(written, (n, take!(io)))
                end
            end
        end
    finally
        close(written)
    end
    out = collect(written)
    wait(workers)   # rethrows a worker's error
    sort!(out; by=first)
    return last.(out)
end

const DOCPROPS_VTYPES_NS = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"

# Regenerates the "Worksheets" and "Charts" entries in HeadingPairs, and their
# corresponding slices of TitlesOfParts, from `wb.sheets`. Any other category
# (e.g. "Named Ranges") and its titles are left untouched, in their original
# position. A recognized category is dropped entirely if the workbook no
# longer has any sheets of that kind; a missing-but-needed category is
# appended if the workbook now has sheets of a kind it didn't declare before.
#
# Note: TitlesOfParts groups titles by *category*, not by tab order — a
# workbook with tab order [Chart1, Sheet1] still lists TitlesOfParts as
# [Sheet1, Chart1], since "Worksheets" titles are grouped before "Charts"
# titles regardless of actual tab position. Confirmed against a real
# Excel-produced chartsheet fixture.
function update_app_xml!(xl::XLSXFile)
    haskey(xl.data, "docProps/app.xml") || return nothing
    wb = get_workbook(xl)

    pfx = get_prefix("docProps/app.xml", xl)
    pfx = pfx == "" ? "" : "$(pfx):"

    appdoc  = xmlroot(xl, "docProps/app.xml")
    approot = xml_root_element(appdoc)
    nss     = get_namespaces(approot)
    vt_pfx  = something(findfirst(==(DOCPROPS_VTYPES_NS), nss), "vt")
    vt(tag) = prefixed_tag(vt_pfx, tag)

    i, hj  = get_idces(appdoc, "Properties", "HeadingPairs")
    i2, tj = get_idces(appdoc, "Properties", "TitlesOfParts")
    (isnothing(hj) || isnothing(tj)) && return nothing

    heading_vector = only(xml_elements(appdoc[i][hj]))
    pairs = collect(xml_elements(heading_vector))           # flat [label, count, label, count, ...]

    titles_vector = only(xml_elements(appdoc[i2][tj]))
    title_nodes = collect(xml_elements(titles_vector))       # original <vt:lpstr> nodes, in title order

    new_names = Dict(
        "Worksheets" => [s.name for s in wb.sheets if !is_chartsheet(wb, s.name)],
        "Charts"     => [s.name for s in wb.sheets if  is_chartsheet(wb, s.name)],
    )

    new_pairs  = Vector{XML.Node}()
    new_titles = Vector{XML.Node}()
    handled    = Set{String}()
    pos = 1

    c = 1
    while c <= length(pairs) - 1
        label = text_value(only(xml_elements(pairs[c])))
        count = parse(Int, something(text_value(only(xml_elements(pairs[c+1]))), "0"))

        if haskey(new_names, label)
            names = new_names[label]
            if !isempty(names)
                push!(new_pairs, pairs[c])
                push!(new_pairs, XML.Element(vt("variant"), XML.Element(vt("i4"), XML.Text(string(length(names))))))
                append!(new_titles, [XML.Element(vt("lpstr"), XML.Text(nm)) for nm in names])
            end
            # else: no sheets of this kind remain — drop the category entirely
            push!(handled, label)
        else
            # Unrecognized category (e.g. "Named Ranges") — pass through untouched.
            push!(new_pairs, pairs[c], pairs[c+1])
            append!(new_titles, title_nodes[pos:pos+count-1])
        end
        pos += count
        c += 2
    end

    for label in ("Worksheets", "Charts")
        label in handled && continue
        names = new_names[label]
        isempty(names) && continue
        push!(new_pairs,
            XML.Element(vt("variant"), XML.Element(vt("lpstr"), XML.Text(label))),
            XML.Element(vt("variant"), XML.Element(vt("i4"), XML.Text(string(length(names))))),
        )
        append!(new_titles, [XML.Element(vt("lpstr"), XML.Text(nm)) for nm in names])
    end

    new_heading_vector = XML.Element(vt("vector"); size=string(length(new_pairs)), baseType="variant")
    for p in new_pairs; push!(new_heading_vector, p); end
    heading_el = XML.Element("$(pfx)HeadingPairs")
    push!(heading_el, new_heading_vector)
    appdoc[i][hj] = heading_el

    new_titles_vector = XML.Element(vt("vector"); size=string(length(new_titles)), baseType="lpstr")
    for t in new_titles; push!(new_titles_vector, t); end
    titles_el = XML.Element("$(pfx)TitlesOfParts")
    push!(titles_el, new_titles_vector)
    appdoc[i2][tj] = titles_el

    return nothing
end

function encode(d::CellValueType)
    if d == CT_STRING
        return "s"
    elseif d ∈ [CT_FLOAT, CT_INT, CT_DATE, CT_DATETIME, CT_TIME]
        return "n"
    elseif d == CT_BOOL
        return "b"
    elseif d == CT_ERROR
        return "e"
    elseif d == CT_EMPTY
        return ""
    else
        throw(XLSXError("Unknown CellDatatype: $d"))
    end
end
function encode(d::CellErrorType)
    if d == XL_NULL
        return "#NULL!"
    elseif d == XL_DIV0
        return "#DIV/0!"
    elseif d == XL_VALUE
        return "#VALUE!"
    elseif d == XL_REF
        return "#REF!"
    elseif d == XL_NAME
        return "#NAME?"
    elseif d == XL_NUM
        return "#NUM!"
    elseif d == XL_NA
        return "#N/A"
    elseif d == XL_SPILL
        return "#SPILL!"
    else
        throw(XLSXError("Unknown CellErrorType: $d"))
    end
end

# Write one row of the cache as XML to `row_node`. `w` holds what every row needs
# (see `_RowWriter`); integers and cell references are written without making
# strings, as this runs once for every cell of the sheet.
function process_cache_row!(row_node::IOBuffer, column_indexes::Vector{Int}, cacherow, w::_RowWriter)
    ws, pfx = w.ws, w.pfx

    d, r, unhandled_attributes = cacherow
    row_nr = row_number(r)

    sort!(append!(empty!(column_indexes), keys(r.rowcells)))   # the row's columns, reusing the vector

    write(row_node, w.row_open)
    _write_number(row_node, w, row_nr)
    write(row_node, "\" spans=\"")
    _write_number(row_node, w, column_number(d.start))
    write(row_node, ":")
    _write_number(row_node, w, column_number(d.stop))
    write(row_node, "\"")

    if !isnothing(r.ht)
        write(row_node, " ht=\"")
        _write_number(row_node, w, r.ht)
        write(row_node, "\" customHeight=\"1\"")
    end

    for (attribute, value) in unhandled_attributes
        write(row_node, " ", attribute, "=\"", XML.escape(value; attribute = true), "\"")
    end

    write(row_node, ">\n")

    for c in column_indexes
        cell = getcell(r, c)
        cell isa EmptyCell && continue

        _write_cell_open!(row_node, w, cell.ref, encode(cell.datatype), cell.style)

        if cell.meta != UInt64(0)
            write(row_node, " cm=\"")
            _write_number(row_node, w, cell.meta)
            write(row_node, "\"")
        end

        has_formula = cell.formula
        is_raw = _is_raw_value_type(cell.datatype)
        has_value = is_raw
        v_other = ""
        if cell.datatype == CT_ERROR
            v_other = encode(reinterpret(CellErrorType, cell.value))
            has_value = !isempty(v_other)
        elseif !is_raw && cell.datatype != CT_EMPTY
            v_other = getdata(ws, cell)
            has_value = !ismissing(v_other) && v_other != ""
        end

        if !has_value && !has_formula
            write(row_node, "/>\n")
        else
            write(row_node, ">")
            if has_formula
                add_node_formula!(row_node, cell.ref, ws, pfx)
            end
            if is_raw
                _write_cell_v!(row_node, w, cell.datatype, cell.value)
            elseif has_value
                write(row_node, w.v_open)
                print(row_node, v_other)
                write(row_node, w.v_close)
            end
            write(row_node, w.c_close)
        end
    end

    write(row_node, w.row_close)
    return nothing
end

# Types whose `<v>` is the cell's stored value, written as a number: a string's is its
# index in the shared string table.
_is_raw_value_type(t::CellValueType) =
    t == CT_FLOAT || t == CT_DATETIME || t == CT_TIME || t == CT_STRING || t == CT_DATE || t == CT_INT || t == CT_BOOL

# `<c r="…"`, with the cell's type `t` (`""` for none) and style, leaving the tag open.
function _write_cell_open!(io::IO, w::_RowWriter, ref::CellRef, t::AbstractString, style::Integer)
    write(io, w.c_open)
    _write_cellref(io, w, ref)
    write(io, "\"")
    isempty(t) || write(io, " t=\"", t, "\"")
    if style != 0
        write(io, " s=\"")
        _write_number(io, w, style)
        write(io, "\"")
    end
    return nothing
end

# The `<v>` of a cell whose type `_is_raw_value_type`, from its stored `value`.
function _write_cell_v!(io::IO, w::_RowWriter, t::CellValueType, value::UInt64)
    write(io, w.v_open)
    if t == CT_FLOAT || t == CT_DATETIME || t == CT_TIME
        _write_number(io, w, reinterpret(Float64, value))
    else
        _write_number(io, w, reinterpret(Int64, value))
    end
    write(io, w.v_close)
    return nothing
end

function abscell(c::CellRef)
    col = encode_column_number(column_number(c))
    row = string(row_number(c))
    return "\$" * col * "\$" * row
end

mkabs(c::SheetCellRef) = abscell(c.cellref)
mkabs(c::SheetCellRange) = abscell(c.rng.start) * ":" * abscell(c.rng.stop)
function make_absolute(dn::DefinedNameValue)
    if dn.value isa NonContiguousRange
        v = ""
        for (i, r) in enumerate(dn.value.rng)
            cr = r isa CellRange ? SheetCellRange(dn.value.sheet, r) : SheetCellRef(dn.value.sheet, r) # need to separate and handle separately
            if dn.isabs[i]
                v *= quoteit(cr.sheet) * "!" * mkabs(cr) * ","
            else
                v *= string(cr) * ","
            end
        end
        return v[begin:prevind(v, end)]
    else
        return dn.isabs ? quoteit(dn.value.sheet) * "!" * mkabs(dn.value) : string(dn.value)
    end
end

# The text of a `<definedName>`: a reference, formula text (held with a leading `=`,
# which the file omits), or a constant, a string being quoted with `"` doubled.
function _defined_name_text(v::DefinedNameValue)::String
    x = v.value
    x isa DefinedNameRangeTypes && return make_absolute(v)
    ismissing(x) && return "\"\""
    x isa Bool && return x ? "TRUE" : "FALSE"
    if x isa AbstractString
        startswith(x, '=') && return String(x[nextind(x, 1):end])
        return "\"" * replace(x, "\"" => "\"\"") * "\""
    end
    return string(x)
end

function update_workbook_xml!(xl::XLSXFile) # Need to update <sheets> and <definedNames>.
    wb = get_workbook(xl)

    pfx = get_prefix("xl/workbook.xml", xl)
    pfx = pfx == "" ? "" : "$(pfx):"

    wbdoc = xmlroot(xl, "xl/workbook.xml") # find the workbook's xml file

    #update calcPr to force update on loaded
    i, j = get_idces(wbdoc, "workbook", "calcPr")
    if !isnothing(j)
        wbdoc[i][j] = XML.Element("$(pfx)calcPr", fullCalcOnLoad="1", calcMode="auto")
    end

    #update defined names
    i, j = get_idces(wbdoc, "workbook", "definedNames")

    if isempty(wb.workbook_names) && isempty(wb.worksheet_names)
        # No defined names left: drop any existing block rather than leaving a
        # stale one behind (an empty <definedNames/> is not valid ECMA-376).
        isnothing(j) || deleteat!(wbdoc[i].children, j)
    else
        if isnothing(j)
            # there is no <definedNames> block in the workbook's xml file, so we'll need to create one
            # `wbdoc[i]`, not `wbdoc[end]`: text or comments may follow `</workbook>`.
            l = insert_index(wbdoc[i], "definedNames", WORKBOOK_ORDER)
            definedNames = XML.Element("$(pfx)definedNames")
            len = length(wbdoc[i])
            if l != len
                insert!(wbdoc[i].children, l + 1, definedNames)
            else
                push!(wbdoc[i], definedNames)
            end
            j = l + 1
        else
            definedNames = unlink(wbdoc[i][j], ("definedNames", "definedName"), pfx) # Remove old defined names
        end

        for (k, v) in wb.workbook_names
            txt = _defined_name_text(v)
            dn_node = XML.Element("$(pfx)definedName", name=k, XML.Text(txt))
            if v.hidden
                dn_node["hidden"] = "1"
            end
            push!(definedNames, dn_node)
        end

        # `n` deliberately not `i`: `i` is still needed for `wbdoc[i][j]` below.
        sheet_ordinals = Dict(s.sheetId => n for (n, s) in enumerate(wb.sheets))
        for (k, v) in wb.worksheet_names
            ordinal = get(sheet_ordinals, first(k), nothing)
            isnothing(ordinal) &&
                throw(XLSXError("Defined name `$(last(k))` is scoped to sheetId $(first(k)), which is not in the workbook."))
            txt = _defined_name_text(v)
            dn_node = XML.Element("$(pfx)definedName", name=last(k),
                                localSheetId=string(ordinal - 1), XML.Text(txt))
            if v.hidden
                dn_node["hidden"] = "1"
            end
            push!(definedNames, dn_node)
        end
        wbdoc[i][j] = definedNames # Add the new definedNames block to the workbook's xml file
    end

    #update sheets

    i, j = get_idces(wbdoc, "workbook", "sheets")
    sheets_element = XML.Element("$(pfx)sheets")
    for s in wb.sheets
        sheet_element = XML.Element("$(pfx)sheet"; name=s.name)
        sheet_element["sheetId"] = string(s.sheetId)
        s.visibility === :visible || (sheet_element["state"] = string(s.visibility))
        sheet_element["r:id"] = s.relationship_id
        push!(sheets_element, sheet_element)
    end
    wbdoc[i][j] = sheets_element

    return nothing
end

