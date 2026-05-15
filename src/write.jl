
"""
    savexlsx(f::XLSXFile)

Save an `XLSXFile` instance back to the file from which it was opened (given in `f.source`), 
overwriting original content.

A new `XLSXFile` created with `XLSX.newxlsx` (or using `openxlsx` without specifying a filename) will 
have `source` set to `"blank.xlsx"` and cannot be saved with this function. Use [`writexlsx`](@ref) instead 
to specify a new file name for the saved file.

An `XLSXFile` created from a native Excel template (`.xltx`) file cannot be saved with this function. 
Use [`writexlsx`](@ref) instead to specify a new file name for the saved file.

Returns the filepath of the written file if a filename is supplied, or `nothing` if writing to an `IO`.

"""
function savexlsx(f::XLSXFile)
    if isa(f.source, AbstractString)
        if f.source == "blank.xlsx"
            throw(XLSXError("Can't save to a blank `XLSXFile` instance. Use `writexlsx` instead to specify a file name."))
        elseif f.is_xltx
            throw(XLSXError("Can't save to a back to an Excel template file. Use `writexlsx` instead to specify a file name."))
        end
    end
    return writexlsx(f.source, f; overwrite=true)
end


"""
    writexlsx(output_source::Union{AbstractString,IO}, xf::XLSXFile; [overwrite=false])

Write an XLSXFile given by `xf` to the IO or filepath `output_source`.

The source attribute of the `XLSXFile` will be updated to the `output_source` if it is a filepath.

Returns the filepath of the written file if a filename is supplied, or `nothing` if writing to an `IO`.

If `overwrite=true`, `output_source` (when a filepath) will be overwritten if it exists.

See also [`savexlsx`](@ref).

!!! note

    When XLSX.jl reads strict (ISO/IEC 29500) XLSX files, it converts them eagerly on 
    to transitional (ECMA 376) format. On write, XLSX.jl will always write in the transitional 
    format, which is the Excel default. Excel itself can convert between strict and 
    transitional formats. Use Excel directly to convert a transitional file to strict format, 
    if needed.

"""
function writexlsx(output_source::Union{AbstractString,IO}, xf::XLSXFile; overwrite::Bool=false)

    !is_writable(xf) && throw(XLSXError("XLSXFile instance is not writable."))
    if !all(values(xf.files))
        throw(XLSXError("Some internal files were not loaded into memory for unknown reasons."))
    end
    if output_source isa AbstractString && !overwrite
        isfile(output_source) && throw(XLSXError("Output file $output_source already exists."))
    end

    update_workbook_xml!(xf)
    
    wb=get_workbook(xf)

    ZipArchives.ZipWriter(output_source) do xlsx

        # write XML files not in cache
        for f in keys(xf.files)
            if !occursin(r"^xl/worksheets/[^/]+\.xml$|^xl/sharedStrings\.xml$", f)
                ZipArchives.zip_newfile(xlsx, f; compress=true)
                xml_str = XML.write(xf.data[f])
                write(xlsx, xml_str)
            end
        end
        # write worksheet files from cache (cache must be enabled in write mode)
        for sheet_no in 1:sheetcount(wb)
            if is_chartsheet(wb, getsheet(wb, sheet_no).name)
                continue
            end
            doc = update_single_sheet!(wb, sheet_no, true)
            f = get_relationship_target_by_id("xl", wb, getsheet(wb, sheet_no).relationship_id)
            ZipArchives.zip_newfile(xlsx, f; compress=true)
            write(xlsx, doc)
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
    end

    if !(output_source isa IO)
        (xf.source = output_source) # update source if output_source is a file path
        return abspath(xf.source)
    else
        return nothing
    end

end

get_worksheet_internal_file(ws::Worksheet) = get_relationship_target_by_id("xl", get_workbook(ws), ws.relationship_id)
get_worksheet_xml_document(ws::Worksheet) = get_xlsxfile(ws).data[get_worksheet_internal_file(ws)]

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

    for s in sst.shared_strings
        print(buff, s)
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
            write(io, " ", k, "=\"", v, "\"")
        end
    end
    
    write(io, ">", XLSX.escape(f.formula), "</$(pfx)f>")
end

function add_node_formula!(io, f::FormulaReference, pfx)
    write(io, "\n        <$(pfx)f t=\"shared\"")
    
    if !isnothing(f.unhandled)
        for (k, v) in f.unhandled
            write(io, " ", k, "=\"", v, "\"")
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
            write(io, " ", k, "=\"", v, "\"")
        end
    end
    
    write(io, " si=\"")
    print(io, f.id)
    write(io, "\">", XLSX.escape(f.formula), "</$(pfx)f>")
end

function find_all_nodes(givenpath::String, doc::XML.Node)::Vector{XML.Node}
    XML.nodetype(doc) != XML.Document && throw(XLSXError("Something wrong here!"))
    found_nodes = Vector{XML.Node}()
    for xp in get_node_paths(doc)
        if xp.path == givenpath
            push!(found_nodes, xp.node)
        end
    end
    return found_nodes
end

function get_node_paths(node::XML.Node)
    XML.nodetype(node) != XML.Document && throw(XLSXError("Something wrong here!"))
    (_, default_ns) = get_default_namespace(node[end])
    xpaths = Vector{xpath}()
    get_node_paths!(xpaths, node, default_ns, "")
    return xpaths
end

function get_node_paths!(xpaths::Vector{xpath}, node::XML.Node, default_ns, path)
    for c in XML.children(node)
        if XML.nodetype(c) ∉ [XML.Declaration, XML.Comment, XML.Text]
            node_tag = localname(c)
            if !occursin(":", node_tag) && !isnothing(default_ns)
                node_tag = default_ns * ":" * node_tag
            end
            npath = path * "/" * node_tag
            push!(xpaths, xpath(c, npath))
            if length(XML.children(c)) > 0
                get_node_paths!(xpaths, c, default_ns, npath)
            end
        end
    end
    return nothing
end


# Remove all children with tag given by att[2] from a parent XML node with a tag given by att[1].
function unlink(node::XML.Node, att::Tuple{String,String}, pfx::String)
    new_node = XML.Element(pfx * first(att))
    
    atts = XML.attributes(node)
    isnothing(atts) || foreach(((k, v),) -> new_node[k] = v, atts)
    
    for child in XML.children(node)
        localname(child) != last(att) && push!(new_node, child)
    end
    
    return new_node
end

function keep_prefix(node, pfx)
    pfx && return XML.tag(node)
    return localname(node)
end

# Find the index location of child and grandchild in a parent XML node.
function find_child_index(children, target)
    pfx = occursin(":", target)
    findfirst(c -> keep_prefix(c, pfx) === target, children)
end
function get_idces(doc::XML.Node, t, b)
    chn = XML.children(doc)
    i = find_child_index(chn, t)
    i === nothing && return nothing, nothing

    j = find_child_index(XML.children(chn[i]), b)
    return i, j
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
function update_single_sheet!(wb::Workbook, sheet_no::Int, full::Bool)::Union{Nothing,Vector{UInt8}}
    sheet = getsheet(wb, sheet_no)
    doc = copynode(get_worksheet_xml_document(sheet))
    xroot = doc[end]

    # check namespace and root node name
    ns_map = get_namespaces(xroot)
    spreadsheet_ns_declared = (SPREADSHEET_NAMESPACE_XPATH_ARG in values(ns_map))
    spreadsheet_ns_declared || throw(XLSXError("Unsupported Spreadsheet XML namespace."))
    localname(xroot) !=   "worksheet" && throw(XLSXError("Malformed Excel file. Expected root node named `worksheet` in worksheet XML file."))

    if full # need to reconstruct row and cell data from cache

        # update worksheet dimension
        i, j = get_idces(doc, "worksheet", "dimension")
        if !isnothing(j)
            dimension_node = doc[i][j]
            dimension_node["ref"] = string(get_dimension(sheet))
        end

        pfx = get_prefix(sheet)
        pfx = pfx == "" ? pfx : pfx * ":"

        empty_doc = XML.write(doc)

        idx = findfirst("<$(pfx)sheetData/>", empty_doc)
        idx === nothing && throw(XLSXError("<$(pfx)sheetData/> placeholder not found when reconstructing worksheet '$(sheet.name)'"))
        new_doc=IOBuffer()
        print(new_doc, empty_doc[begin:first(idx)-1])

        # create <sheetData> with any attributes
        print(new_doc, "<$(pfx)sheetData")
        i, j = get_idces(doc, "worksheet", "sheetData")
        a = XML.attributes(doc[i][j])
        if !isnothing(a)
            for (k, v) in a
                print(new_doc, " ",k,"=\"",v,"\"")
            end
        end
        print(new_doc, ">\n")
        # iterates over WorksheetCache cells and writes the XML
        new_doc=vcat(take!(new_doc), get_cache_rows(sheet, pfx))

        new_doc=vcat(new_doc, Vector{UInt8}("  </$(pfx)sheetData>"))
        new_doc=vcat(new_doc, Vector{UInt8}(empty_doc[last(idx)+1:end]))

        return new_doc

    else
        set_worksheet_xml_document!(sheet, doc) # no need to save full reconstructed data
        return nothing
    end
end

function stream_cache_rows(sheet::Worksheet, chunksize::Int)
    cache_rows = Vector{Tuple{CellRange, SheetRow, Dict{String,String}}}(undef, chunksize)
    count=0
    Channel{Vector{Tuple{CellRange, SheetRow, Dict{String,String}}}}(1 << 8) do out
        d = get_dimension(sheet)
        for r in eachrow(sheet)
            count += 1
            rn=row_number(r)
            uh = (!isnothing(sheet.unhandled_attributes) && haskey(sheet.unhandled_attributes, rn)) ? sheet.unhandled_attributes[rn] : Dict{String,String}()
            cache_rows[count] = (d, r, uh)
            if count >= chunksize
                put!(out, copy(cache_rows))
                count=0
            end
        end
        if count>0
            put!(out, cache_rows[1:count])
        end
    end
end

function get_cache_rows(sheet::Worksheet, pfx::String)::Vector{UInt8}
    chunksize = ROW_CHUNKSIZE
    read_cache_rows = Channel{Vector{Tuple{Int64,Vector{UInt8}}}}(1 << 8)
    all_cache_rows = Vector{Tuple{Int64,Vector{UInt8}}}()
   
    consumer = @async begin
        for rows in read_cache_rows
            for row in rows
                push!(all_cache_rows, row)
            end
        end
    end
   
    cache_rows = stream_cache_rows(sheet, chunksize)
   
    @sync for _ in 1:Threads.nthreads()
        Threads.@spawn begin
            for rows in cache_rows
                processed = [process_cache_row(row, sheet, pfx) for row in rows]
                put!(read_cache_rows, processed)
            end
        end
    end
   
    close(read_cache_rows)
    wait(consumer)
   
    all_rows = last.(sort(all_cache_rows, by=first))
    return length(all_rows) == 0 ? Vector{UInt8}() : reduce(vcat, all_rows)
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

function process_cache_row(cacherow::Tuple{CellRange, SheetRow, Dict{String,String}}, ws::Worksheet, pfx::String)
    pad2 = "    "      # 4 spaces
    pad3 = "      "    # 6 spaces
    pad4 = "        "  # 8 spaces

    d, r, unhandled_attributes = cacherow
    spans_str = string(column_number(d.start), ":", column_number(d.stop))

    row_nr = row_number(r)
    ordered_column_indexes = sort!(collect(keys(r.rowcells)))

    row_node = IOBuffer()

    write(row_node, pad2, "<$(pfx)row r=\"", string(row_nr), "\"")

    if spans_str != ""
        write(row_node, " spans=\"", spans_str, "\"")
    end

    if !isnothing(r.ht)
        write(row_node, " ht=\"", string(r.ht), "\" customHeight=\"1\"")
    end

    for (attribute, value) in unhandled_attributes
        write(row_node, " ", attribute, "=\"", value, "\"")
    end

    write(row_node, ">\n")

    for c in ordered_column_indexes
        cell = getcell(r, c)

        if cell isa EmptyCell
            continue
        end

        write(row_node, pad3, "<$(pfx)c r=\"", string(cell.ref), "\"")

        if cell.datatype != CT_EMPTY
            write(row_node, " t=\"", encode(cell.datatype), "\"")
        end

        if cell.style != UInt32(0)
            write(row_node, " s=\"", string(cell.style), "\"")
        end

        if cell.meta != UInt64(0)
            write(row_node, " cm=\"", string(cell.meta), "\"")
        end

        if cell.datatype ∈ [CT_FLOAT, CT_DATETIME, CT_TIME]
            v = reinterpret(Float64, cell.value)
        elseif cell.datatype ∈ [CT_STRING, CT_DATE, CT_INT, CT_BOOL]
            v = reinterpret(Int64, cell.value)
        elseif cell.datatype == CT_ERROR
            v = encode(reinterpret(CellErrorType, cell.value))
        elseif cell.datatype == CT_EMPTY
            v = ""
        else
            v = getdata(ws, cell)
        end

        has_value = !ismissing(v) && v != ""
        has_formula = cell.formula

        if !has_value && !has_formula
            write(row_node, "/>\n")
        else
            write(row_node, ">")
            if has_formula
                add_node_formula!(row_node, cell.ref, ws, pfx)
            end
            if has_value
                write(row_node, "\n", pad4, "<$(pfx)v>", string(v), "</$(pfx)v>")
            end
            write(row_node, "\n", pad3, "</$(pfx)c>\n")
        end
    end

    write(row_node, pad2, "</$(pfx)row>\n")
    return (row_nr, take!(row_node))
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
    if length(wb.workbook_names) > 0 || length(wb.worksheet_names) > 0 # skip if no defined names present
        i, j = get_idces(wbdoc, "workbook", "definedNames")
        if isnothing(j)
            # there is no <definedNames> block in the workbook's xml file, so we'll need to create one
            l = insert_index(wbdoc[end], "definedNames", WORKBOOK_ORDER)
            definedNames = XML.Element("$(pfx)definedNames")
            len = length(wbdoc[end])
            if l != len
                insert!(wbdoc[end].children, l+1, definedNames)
            else
                push!(wbdoc[end], definedNames)
            end
            j = l + 1
        else
            definedNames = unlink(wbdoc[i][j], ("definedNames", "definedName"), pfx) # Remove old defined names
        end
        for (k, v) in wb.workbook_names
            if typeof(v.value) <: DefinedNameRangeTypes
                v = make_absolute(v)
            else
                v = string(v.value)
            end
            dn_node = XML.Element("$(pfx)definedName", name=k, XML.Text(v))
            push!(definedNames, dn_node)
        end
        for (k, v) in wb.worksheet_names
            if typeof(v.value) <: DefinedNameRangeTypes
                v = make_absolute(v)
            else
                v = string(v.value)
            end
            dn_node = XML.Element("$(pfx)definedName", name=last(k), localSheetId=first(k) - 1, XML.Text(v))
            push!(definedNames, dn_node)
        end
        wbdoc[i][j] = definedNames # Add the new definedNames block to the workbook's xml file
    end

    #update sheets
#    doc = xmlroot(xl, "xl/workbook.xml")

    i, j = get_idces(wbdoc, "workbook", "sheets")
    unlink(wbdoc[i][j], ("sheets", "sheet"), pfx)
    sheets_element = XML.Element("$(pfx)sheets")
    for s in wb.sheets
        sheet_element = XML.Element("$(pfx)sheet"; name=XLSX.escape(s.name))
        sheet_element["sheetId"] = s.sheetId
        sheet_element["r:id"] = s.relationship_id
        push!(sheets_element, sheet_element)
    end
    wbdoc[i][j] = sheets_element

    return nothing
end

function add_cell_to_worksheet_dimension!(ws::Worksheet, cell::Cell)
    # update worksheet dimension
    ws_dimension = get_dimension(ws)

    top = row_number(ws_dimension.start)
    left = column_number(ws_dimension.start)

    bottom = row_number(ws_dimension.stop)
    right = column_number(ws_dimension.stop)

    r = row_number(cell)
    c = column_number(cell)

    if r < top || c < left
        top = min(r, top)
        left = min(c, left)
        set_dimension!(ws, CellRange(top, left, bottom, right))
    elseif r > bottom || c > right
        bottom = max(r, bottom)
        right = max(c, right)
        set_dimension!(ws, CellRange(top, left, bottom, right))
    end

    nothing
end

function setdata!(ws::Worksheet, cell::Cell)
    !is_writable(get_xlsxfile(ws)) && throw(XLSXError("XLSXFile instance is not writable. Open Excel file with `mode=\"rw\"` instead"))
    ws.cache === nothing && throw(XLSXError("Can't write data to a Worksheet with empty cache."))
    cache = ws.cache

    r = row_number(cell)
    c = column_number(cell)

    if !haskey(cache.cells, r)
        push!(cache.rows_in_cache, r)
        cache.cells[r] = Dict{Int,Cell}()
        cache.row_ht[r] = nothing
        cache.dirty = true
    end
    cache.cells[r][c] = cell
    add_cell_to_worksheet_dimension!(ws, cell)

    nothing
end
 
# Issue #284
function strip_illegal_chars(s::String; strip_discouraged::Bool=true)
    # remove XML 1.0 illegal C0 controls and U+0000
    s = replace(s, r"[\x00-\x08\x0B\x0C\x0E-\x1F]" => "")
    # remove BMP non-characters U+FFFE U+FFFF
    s = replace(s, "\uFFFE" => "", "\uFFFF" => "")
    # optionally remove U+FDD0..U+FDEF and plane-ending non-characters
    if strip_discouraged
        s = replace(s, r"[\uFDD0-\uFDEF]" => "")
        for cp in (0x1FFFE:0x10000:0x10FFFF)
            # remove both XFFFE and XFFFF for each plane
            s = replace(s, Char(cp) => "")
            s = replace(s, Char(cp+1) => "")
        end
    end
    return s
end

# Returns the datatype and value for `val` to be inserted into `ws`.
function xlsx_encode(ws::Worksheet, val::AbstractString)
    if isempty(val)
        return (CT_EMPTY, UInt64(0))
    end
    sst_ind = add_shared_string!(get_workbook(ws), strip_illegal_chars(val))
    ws.sst_count+=1

    return (CT_STRING, UInt64(sst_ind))
end

xlsx_encode(::Worksheet, val::Missing) = (CT_EMPTY, UInt64(0))
xlsx_encode(::Worksheet, val::Bool) = (CT_BOOL, UInt64(val))
xlsx_encode(::Worksheet, val::Float64) = (CT_FLOAT, reinterpret(UInt64, val))
xlsx_encode(::Worksheet, val::Int64) = (CT_INT, reinterpret(UInt64, val))
xlsx_encode(ws::Worksheet, val::Dates.Date) = (CT_DATE, reinterpret(UInt64, date_to_excel_value(val, isdate1904(get_xlsxfile(ws)))))
xlsx_encode(ws::Worksheet, val::Dates.DateTime) = (CT_DATETIME, reinterpret(UInt64, datetime_to_excel_value(val, isdate1904(get_xlsxfile(ws)))))
xlsx_encode(::Worksheet, val::Dates.Time) = (CT_TIME, reinterpret(UInt64, time_to_excel_value(val)))
Base.setindex!(ws::Worksheet, v, row::Integer, col::Integer) = setdata!(ws, CellRef(row,col), v)
Base.setindex!(ws::Worksheet, v, row::Union{Integer,UnitRange{<:Integer}}, col::Union{Integer,UnitRange{<:Integer}}) = setdata!(ws, CellRange(CellRef(first(row), first(col)), CellRef(last(row), last(col))), v)
Base.setindex!(ws::Worksheet, v::AbstractVector, r::Union{Integer,UnitRange{<:Integer}}, c::UnitRange{T}) where {T<:Integer} = setdata!(ws, r, c, v)
Base.setindex!(ws::Worksheet, v::AbstractVector, r::UnitRange{T}, c::Union{Integer,UnitRange{<:Integer}}) where {T<:Integer} = setdata!(ws, r, c, v)
Base.setindex!(ws::Worksheet, v::AbstractVector, ref; dim::Integer=2) = setdata!(ws, ref, v, dim)
Base.setindex!(ws::Worksheet, v::AbstractVector, r, c; dim::Integer=2) = setdata!(ws, r, c, v, dim)
Base.setindex!(ws::Worksheet, v, ref) = setdata!(ws, ref, v)
Base.setindex!(ws::Worksheet, v, r, c) = setdata!(ws, r, c, v)
function Base.setindex!(ws::Worksheet, v, row::Union{Integer,UnitRange{<:Integer}}, col::Union{Vector{Int},StepRange{<:Integer}})
    for a in row, b in col
        setdata!(ws, CellRef(a, b), v)
    end
end
function Base.setindex!(ws::Worksheet, v, row::Union{Vector{Int},StepRange{<:Integer}}, col::Union{Integer,UnitRange{<:Integer}})
    for a in row, b in col
        setdata!(ws, CellRef(a, b), v)
    end
end
function Base.setindex!(ws::Worksheet, v, row::Union{Vector{Int},StepRange{<:Integer}}, col::Union{Vector{Int},StepRange{<:Integer}})
    for a in row, b in col
        setdata!(ws, CellRef(a, b), v)
    end
end

function setdata!(ws::Worksheet, ref::CellRef, val::CellFormula)
    c = getcell(ws, ref)
    if !(c isa EmptyCell) && c.formula && get_formula_from_cache(ws, ref) isa ReferencedFormula
        rereference_formulae(ws, c)
    end
    v = ""
    t = ""
    m = ""
    cell = Cell(get_workbook(ws), ref, t, string(id(val.styleid)), v, m, true)
    add_formula_to_cache(ws, ref, val.value)
    setdata!(ws, cell)
end
function setdata!(ws::Worksheet, ref::CellRef, val::CellValue, convert_to_string::Bool)
    # Convert Float64 values NaN, Inf and -Inf to strings. Excel can't handle them as floats
    # Addresses #179 & #342
    @assert convert_to_string == true "Converting values other than NaN, -Inf, Inf is not permitted"
    val = CellValue(string(val.value), val.styleid)
    setdata!(ws, ref, val)
end
function setdata!(ws::Worksheet, ref::CellRef, val::CellValue)
    if val.value isa Float64 && (isnan(val.value) || isinf(val.value))
        return setdata!(ws, ref, val, true)
    end
    t, v = xlsx_encode(ws, val.value)
    m = UInt64(0)
    cell = Cell(ref, v, id(val.styleid), m, t, false)
    setdata!(ws, cell)
end

# convert AbstractTypes to concrete
setdata!(ws::Worksheet, ref::CellRef, val::AbstractString) = setdata!(ws, ref, CellValue(ws, convert(String, val)))
setdata!(ws::Worksheet, ref::CellRef, val::Integer) = setdata!(ws, ref, convert(Int64, val))
setdata!(ws::Worksheet, ref::CellRef, val::Real) = setdata!(ws, ref, convert(Float64, val))

function setdata!(sheet::Worksheet, ref::CellRange, data::AbstractVector) # vector to row or column range
    rows = ref.start.row_number:ref.stop.row_number
    cols = ref.start.column_number:ref.stop.column_number
    length(rows)==1 || length(cols)==1 || throw(XLSXError("Can only write 1D data to a row or column range."))
    if length(cols)==1
        length(rows) != length(data) && throw(XLSXError("Row count mismatch between `data` ($(length(data)) rows) and row range $rows ($(length(rows)) rows)."))
        anchor_cell_ref = CellRef(first(rows), first(cols))
        setdata!(sheet, anchor_cell_ref, data, 1)
    elseif length(rows)==1
        length(cols) != length(data) && throw(XLSXError("Column count mismatch between `data` ($(length(data)) columns) and column range $cols ($(length(cols)) columns)."))
        anchor_cell_ref = CellRef(first(rows), first(cols))
        setdata!(sheet, anchor_cell_ref, data, 2)
    end
end

# convert nothing to missing when writing
setdata!(ws::Worksheet, ref::CellRef, ::Nothing) = setdata!(ws, ref, CellValue(ws, missing))

setdata!(ws::Worksheet, row::Integer, col::Integer, val::CellValue) = setdata!(ws, CellRef(row, col), val)

setdata!(ws::Worksheet, row::Union{Integer,UnitRange{<:Integer}}, col::Union{Integer,UnitRange{<:Integer}}, v) = setdata!(ws, CellRange(CellRef(first(row), first(col)), CellRef(last(row), last(col))), v)

function setdata!(ws::Worksheet, ref::CellRef, val::Union{AbstractFormula,CellConcreteType}) # use existing cell format if it exists
    c = getcell(ws, ref)
    if !(c isa EmptyCell) && c.formula && get_formula_from_cache(ws, ref) isa ReferencedFormula
        rereference_formulae(ws, c)
    end
    if c isa EmptyCell || c.style == ""
        if val isa AbstractFormula
            return setdata!(ws, ref, CellFormula(ws, val))
        else
            return setdata!(ws, ref, CellValue(ws, val))
        end
    else
        existing_style = CellDataFormat(c.style)
        isa_dt = styles_is_datetime(ws.package.workbook, existing_style)
        if val isa Dates.Date
            if isa_dt == false
                c.style = update_template_xf(ws, existing_style, ["numFmtId", "applyNumberFormat"], [string(DEFAULT_DATE_numFmtId), "1"]).id
            end
        elseif val isa Dates.Time
            if isa_dt == false
                c.style = update_template_xf(ws, existing_style, ["numFmtId", "applyNumberFormat"], [string(DEFAULT_TIME_numFmtId), "1"]).id
            end
        elseif val isa Dates.DateTime
            if isa_dt == false
                c.style = update_template_xf(ws, existing_style, ["numFmtId", "applyNumberFormat"], [string(DEFAULT_DATETIME_numFmtId), "1"]).id
            end
        elseif val isa Float64 || val isa Int
            if styles_is_float(ws.package.workbook, existing_style) == false && Int(existing_style.id) ∉ [0, 1]
                c.style = update_template_xf(ws, existing_style, ["numFmtId"], [string(DEFAULT_NUMBER_numFmtId)]).id
            end
        elseif val isa Bool # Now rerouted here rather than assigning an EmptyCellDataFormat.
            # Change any style to General (0) and retain other formatting.
            c.style = update_template_xf(ws, existing_style, ["numFmtId"], [string(DEFAULT_BOOL_numFmtId)]).id
        end
        if val isa AbstractFormula
            return setdata!(ws, ref, CellFormula(val, CellDataFormat(c.style)))
        else
            return setdata!(ws, ref, CellValue(val, CellDataFormat(c.style)))
        end
    end
end
function setdata!(ws::Worksheet, ref::AbstractString, value)
    if value isa String
        i = findfirst(!isspace, value)
        if !isnothing(i) && length(value[i:end]) > 1 && value[i] == '=' # it's a formula!
            value = Formula(last(split(value, '=')))
        end
    end
    if is_worksheet_defined_name(ws, ref)
        v = get_defined_name_value(ws, ref)
        if is_defined_name_value_a_reference(v)
            return setdata!(ws, v, value)
        else
            throw(XLSXError("`$ref` is not a valid cell or range reference."))
        end
    elseif is_workbook_defined_name(get_workbook(ws), ref)
        wb = get_workbook(ws)
        v = get_defined_name_value(wb, ref)
        if is_defined_name_value_a_reference(v)
            return setdata!(ws, v, value)
        else
            throw(XLSXError("`$ref` is not a valid cell or range reference."))
        end
    elseif is_valid_cellname(ref)
        return setdata!(ws, CellRef(ref), value)
    elseif is_valid_sheet_cellname(ref)
        return setdata!(ws, SheetCellRef(ref), value)
    elseif is_valid_cellrange(ref)
        return setdata!(ws, CellRange(ref), value)
    elseif is_valid_column_range(ref)
        return setdata!(ws, ColumnRange(ref), value)
    elseif is_valid_row_range(ref)
        return setdata!(ws, RowRange(ref), value)
    elseif is_valid_non_contiguous_range(ref)
        return setdata!(ws, NonContiguousRange(ws, ref), value)
    elseif is_valid_sheet_cellrange(ref)
        return setdata!(ws, SheetCellRange(ref), value)
    elseif is_valid_sheet_column_range(ref)
        return setdata!(ws, SheetColumnRange(ref), value)
    elseif is_valid_sheet_row_range(ref)
        return setdata!(ws, SheetRowRange(ref), value)
    elseif is_valid_non_contiguous_cellrange(ref)
        return setdata!(ws, NonContiguousRange(ws, ref), value)
    end
    throw(XLSXError("`$ref` is not a valid cell or range reference."))
end
function setdata!(ws::Worksheet, rng::CellRange, value)
    for row in rng.start.row_number:rng.stop.row_number
        for col in rng.start.column_number:rng.stop.column_number
            setdata!(ws, row, col, value)
        end
    end
end
function setdata!(ws::Worksheet, rng::RowRange, value)
    dim = get_dimension(ws)
    start = CellRef(rng.start, dim.start.column_number,)
    stop = CellRef(rng.stop, dim.stop.column_number)
    setdata!(ws, CellRange(start, stop), value)
end
function setdata!(ws::Worksheet, rng::ColumnRange, value)
    dim = get_dimension(ws)
    start = CellRef(dim.start.row_number, rng.start)
    stop = CellRef(dim.stop.row_number, rng.stop)
    setdata!(ws, CellRange(start, stop), value)
end
function setdata!(ws::Worksheet, rng::NonContiguousRange, value)
    do_sheet_names_match(ws, rng)
    for r in rng.rng
        setdata!(ws, r, value) # r may be a single Cell or a CellRange
    end
end
setdata!(ws::Worksheet, ::Colon, ::Colon, v) = setdata!(ws::Worksheet, :, v)
function setdata!(ws::Worksheet, ::Colon, v)
    dim = get_dimension(ws)
    setdata!(ws, dim, v)
end
function setdata!(ws::Worksheet, row::Union{Integer,UnitRange{<:Integer}}, ::Colon, v)
    dim = get_dimension(ws)
    setdata!(ws, CellRange(CellRef(first(row), dim.start.column_number), CellRef(last(row), dim.stop.column_number)), v)
end
function setdata!(ws::Worksheet, ::Colon, col::Union{Integer,UnitRange{<:Integer}}, v)
    dim = get_dimension(ws)
    setdata!(ws, CellRange(CellRef(dim.start.row_number, first(col)), CellRef(dim.stop.row_number, last(col))), v)
end
function setdata!(ws::Worksheet, row::Union{Vector{Int},StepRange{<:Integer}}, ::Colon, v)
    dim = get_dimension(ws)
    for a in row
        for b in dim.start.column_number:dim.stop.column_number
            setdata!(ws, CellRef(a, b), v)
        end
    end
end
function setdata!(ws::Worksheet, ::Colon, col::Union{Vector{Int},StepRange{<:Integer}}, v)
    dim = get_dimension(ws)
    for b in col
        for a in dim.start.row_number:dim.stop.row_number
            setdata!(ws, CellRef(a, b), v)
        end
    end
end
setdata!(ws::Worksheet, ref::SheetCellRef, value) = do_sheet_names_match(ws, ref) && setdata!(ws, ref.cellref, value)
setdata!(ws::Worksheet, rng::SheetCellRange, value) = do_sheet_names_match(ws, rng) && setdata!(ws, rng.rng, value)
setdata!(ws::Worksheet, rng::SheetColumnRange, value) = do_sheet_names_match(ws, rng) && setdata!(ws, rng.colrng, value)
setdata!(ws::Worksheet, rng::SheetRowRange, value) = do_sheet_names_match(ws, rng) && setdata!(ws, rng.rowrng, value)
setdata!(ws::Worksheet, row::Integer, col::Integer, data) = setdata!(ws, CellRef(row, col), data)
setdata!(ws::Worksheet, ref::CellRef, value) = throw(XLSXError("$value: Unsupported datatype $(typeof(value)) for writing data to Excel file. Supported data types are $(CellConcreteType) or $(CellValue)."))
setdata!(ws::Worksheet, row::Integer, col::Integer, data::AbstractVector, dim::Integer) = setdata!(ws, CellRef(row, col), data, dim)
function setdata!(ws::Worksheet, ref_str::AbstractString, value::AbstractVector, dim::Integer)
    if is_valid_cellrange(ref_str)
        setdata!(ws, CellRange(ref_str), value)
    elseif is_valid_cellname(ref_str)
        setdata!(ws, CellRef(ref_str), value, dim)
    else
        throw(XLSXError("Invalid cell reference or range: $ref_or_rng"))
    end
end

function setdata!(sheet::Worksheet, ref::CellRef, data::AbstractVector, dim::Integer)
    for (i, val) in enumerate(data)
        target_cell_ref = target_cell_ref_from_offset(ref, i - 1, dim)
        setdata!(sheet, target_cell_ref, val)
    end
end

Base.setindex!(ws::Worksheet, v::AbstractVector, r::Union{Colon,UnitRange{T}}, c) where {T<:Integer} = setdata!(ws, r, c, v)
Base.setindex!(ws::Worksheet, v::AbstractVector, r, c::Union{Colon,UnitRange{T}}) where {T<:Integer} = setdata!(ws, r, c, v)
setdata!(sheet::Worksheet, ::Colon, col::Integer, data::AbstractVector) = setdata!(sheet, 1, col, data, 1)
setdata!(sheet::Worksheet, row::Integer, ::Colon, data::AbstractVector) = setdata!(sheet, row, 1, data, 2)

function setdata!(sheet::Worksheet, row::Integer, cols::UnitRange{T}, data::AbstractVector) where {T<:Integer}
    length(data) != length(cols) && throw(XLSXError("Column count mismatch between `data` ($(length(data)) columns) and column range $cols ($(length(cols)) columns)."))
    anchor_cell_ref = CellRef(row, first(cols))

    # since cols is the unit range, this is a column-based operation
    setdata!(sheet, anchor_cell_ref, data, 2)
end

function setdata!(sheet::Worksheet, rows::UnitRange{T}, col::Integer, data::AbstractVector) where {T<:Integer}
    length(data) != length(rows) && throw(XLSXError("Row count mismatch between `data` ($(length(data)) rows) and row range $rows ($(length(rows)) rows)."))
    anchor_cell_ref = CellRef(first(rows), col)

    # since rows is the unit range, this is a row-based operation
    setdata!(sheet, anchor_cell_ref, data, 1)
end

function setdata!(sheet::Worksheet, ref_or_rng::AbstractString, matrix::AbstractArray{T,2}) where {T}
    if is_valid_cellrange(ref_or_rng)
        setdata!(sheet, CellRange(ref_or_rng), matrix)
    elseif is_valid_cellname(ref_or_rng)
        setdata!(sheet, CellRef(ref_or_rng), matrix)
    else
        throw(XLSXError("Invalid cell reference or range: $ref_or_rng"))
    end
end

# Generalise to AbstractArray (#158)
function setdata!(sheet::Worksheet, ref::CellRef, matrix::AbstractArray{T,2}) where {T}
    row_ax, col_ax = axes(matrix)
    anchor_row = row_number(ref)
    anchor_col = column_number(ref)

    @inbounds for c in col_ax, r in row_ax
        setdata!(sheet, anchor_row + (r - first(row_ax)), anchor_col + (c - first(col_ax)), matrix[r, c])
    end
end

function setdata!(sheet::Worksheet, rng::CellRange, matrix::AbstractArray{T,2}) where {T}
    size(rng) != size(matrix) && throw(XLSXError("Target range $rng size ($(size(rng))) must be equal to the input matrix size ($(size(matrix)))"))
    setdata!(sheet, rng.start, matrix)
end

function setdata!(sheet::Worksheet, ref::CellRef, rts::RichTextString)
    isempty(rts.text) && throw(XLSXError("Cannot set a cell value to an empty RichTextString."))
    c=getcell(sheet, ref)
    if c isa EmptyCell
        sheet[ref]=""
        c=getcell(sheet, ref)
    end

    pfx = get_prefix(sheet)
    pfx = pfx == "" ? pfx : pfx * ":"

    if length(rts.runs) > 1 # add RichTextString as a rich sharedString
        sheet.sst_count += 1
        c.value = add_formatted_string!(get_workbook(sheet), richTextStringtoXML(rts, pfx))
        c.datatype = CT_STRING
    else # add single run as a normal cell value with a cell level font style
        sheet[ref] = rts.text
        setFont(sheet, ref; pairs(rts.runs[1].atts)...)
    end     

    return sheet[ref]
end

# Given an anchor cell at (anchor_row, anchor_col).
# Returns a CellRef at:
# - (anchor_row + offset, anchol_col) if dim = 1 (operates on rows)
# - (anchor_row, anchor_col + offset) if dim = 2 (operates on cols)
function target_cell_ref_from_offset(anchor_row::Integer, anchor_col::Integer, offset::Integer, dim::Integer)::CellRef
    if dim == 1
        return CellRef(anchor_row + offset, anchor_col)
    elseif dim == 2
        return CellRef(anchor_row, anchor_col + offset)
    else
        throw(XLSXError("Invalid dimension: $dim."))
    end
end

function target_cell_ref_from_offset(anchor_cell::CellRef, offset::Integer, dim::Integer)::CellRef
    return target_cell_ref_from_offset(row_number(anchor_cell), column_number(anchor_cell), offset, dim)
end

const ALLOWED_TYPES = Union{Number,String,Bool,Dates.Date,Dates.Time,Dates.DateTime,Missing,Nothing}
function process_vector(col) # Convert any disallowed types to strings. #239.
    if eltype(col) <: ALLOWED_TYPES
        # Case 1: All elements are of allowed types
        return col
    elseif eltype(col) <: Any && all(x -> !(typeof(x) <: ALLOWED_TYPES), col)
        # Case 2: All elements are of disallowed types
        return map(x -> string(x), col)
    else
        # Case 3: Mixed types, process each element
        return [typeof(x) <: ALLOWED_TYPES ? x : string(x) for x in col]
    end
end

"""
    writetable!(
        sheet::Worksheet,
        data,
        columnnames;
        anchor_cell::CellRef=CellRef("A1"),
        write_columnnames::Bool=true,
    )

Write tabular data `data` with labels given by `columnnames` to `sheet`,
starting at `anchor_cell`.

`data` must be a vector of columns.
`columnnames` must be a vector of column labels.

Column labels that are not of type `String` will be converted 
to strings before writing. Any data columns that are not of 
type `String`, `Float64`, `Int64`, `Bool`, `Date`, `Time`, 
`DateTime`, `Missing`, or `Nothing` will be converted to strings 
before writing.


See also: [`XLSX.writetable`](@ref).
"""
function writetable!(
    sheet::Worksheet,
    data,
    columnnames;
    anchor_cell::CellRef=CellRef("A1"),
    write_columnnames::Bool=true,
)

    # read dimensions
    col_count = length(data)
    col_count != length(columnnames) && throw(XLSXError("Column count mismatch between `data` ($col_count columns) and `columnnames` ($(length(columnnames)) columns)."))
    col_count <= 0 && throw(XLSXError("Can't write table with no columns."))
    col_count > EXCEL_MAX_COLS && throw(XLSXError("`data` contains $col_count columns, but Excel only supports up to $EXCEL_MAX_COLS; must reduce `data` size"))
    row_count = length(data[1])
    row_count > EXCEL_MAX_ROWS - 1 && throw(XLSXError("`data` contains $row_count rows, but Excel only supports up to $(EXCEL_MAX_ROWS-1); must reduce `data` size"))
    if col_count > 1
        for c in 2:col_count
            length(data[c]) != row_count && throw(XLSXError("Row count mismatch between column 1 ($row_count rows) and column $c ($(length(data[c])) rows)."))
        end
    end

    anchor_row = row_number(anchor_cell)
    anchor_col = column_number(anchor_cell)
    start_from_anchor = 1

    # write table header
    if write_columnnames
        for c in 1:col_count
            target_cell_ref = CellRef(anchor_row, c + anchor_col - 1)
            sheet[target_cell_ref] = string(columnnames[c])
        end
        start_from_anchor = 0
    end

    # write table data
    data = [process_vector(col) for col in data] # Address issue #239
    for r in 1:row_count
        for c in 1:col_count # 15% speed up by inverting nesting of these for-loops!
            target_cell_ref = CellRef(r + anchor_row - start_from_anchor, c + anchor_col - 1)
            v = data[c][r]
            sheet[target_cell_ref] = v 
        end
    end
end

rename!(ws::Worksheet, name::AbstractString) = renamesheet!(ws, name)

"""
    renamesheet!(ws::Worksheet, name::AbstractString)

Rename a `Worksheet` to `name`.

See also [addsheet!](@ref), [copysheet!](@ref), [deletesheet!](@ref)

"""
function renamesheet!(ws::Worksheet, name::AbstractString)

    # no-op if the name has not changed
    if ws.name == name
        return
    end

    xf = get_xlsxfile(ws)
    !is_writable(xf) && throw(XLSXError("XLSXFile instance is not writable."))
    name ∈ sheetnames(xf) && throw(XLSXError("Sheetname $name is already in use."))

    # updates XML
    xroot = xmlroot(xf, "xl/workbook.xml")[end]
    wb = get_workbook(ws)
    for node in XML.children(xroot)
        if localname(node) == "sheets"

            for sheet_node in XML.children(node)
                if sheet_node["name"] == ws.name
                    # assign new name
                    sheet_node["name"] = name
                    break
                end
            end

            break
        end
    end

    update_formulas_renamed_sheet!(get_workbook(ws), ws.name, name)

    # updates the new name in the worksheet instance
    ws.name = name
    nothing
end

include_dependency(joinpath(@__DIR__, "data", "sheet_template.xml"))
const SHEET_TEMPLATE_XML_DATA = read(joinpath(@__DIR__, "data", "sheet_template.xml"))

"""
    addsheet!(wb::Workbook, [name::AbstractString=""]) --> ::Worksheet
    addsheet!(xf::XLSXFile, [name::AbstractString=""]) --> ::Worksheet

Create a new worksheet named `name`.
If `name` is not provided, a unique name is created.

See also [renamesheet!](@ref), [copysheet!](@ref), [deletesheet!](@ref)

"""
addsheet!(xl::XLSXFile, name::AbstractString="")::Worksheet = addsheet!(get_workbook(xl), name)::Worksheet
function addsheet!(wb::Workbook, name::AbstractString=""; sheet_template_data::Vector{UInt8}=SHEET_TEMPLATE_XML_DATA)::Worksheet
    # Copy so the global SHEET_TEMPLATE_DATA isn't accidentally changed
    bytes = copy(sheet_template_data)
    f, _ = skipNode(XML.Raw(bytes), "sheetData")
    xdoc = XML.Node(XML.Raw(f))

    new_cache = XLSX.WorksheetCache(
        true,
        Dict{Int,Dict{Int,XLSX.Cell}}(),
        Int[],
        Dict{Int,Union{Float64,Nothing}}(),
        Dict{Int,Int}(),
        SheetRowStreamIterator(get_xlsxfile(wb)[1]), # Dummy - not needed because using full cache.
        nothing,
        false
    )
    new_ws = insertsheet!(wb, xdoc, new_cache, 0, "", name)
    return new_ws
end

"""
    copysheet!(ws::Worksheet, [name::AbstractString=""]) --> ::Worksheet

Create a copy of the worksheet `ws` and add it to the end of the workbook with the 
specified worksheet name.
Return the new worksheet object.
If `name` is not provided, a new name is generated by appending " (copy)" to the original 
worksheet name, with a further numerical suffix to guarantee uniqueness if necessary.
To copy worksheets, the `XLSXFile` must be writable (opened with `mode="rw"` or as a template).
See also [`XLSX.openxlsx`](@ref) and [XLSX.opentemplate](@ref).

!!! warning "Experimental"

    This function is experimental is not guaranteed to work with all XLSX files, 
    especially those with complex features. However, cell formats, conditional formats 
    and worksheet defined names should all copy OK. Please report any issues.

See also [addsheet!](@ref), [renamesheet!](@ref), [deletesheet!](@ref) 

# Examples
```julia
julia> f=XLSX.openxlsx("general.xlsx", mode="rw")
XLSXFile("C:\\...\\general.xlsx") containing 13 Worksheets
            sheetname size          range
-------------------------------------------------
              general 10x6          A1:F10
               table3 5x6           A2:F6
               table4 4x3           E12:G15
                table 12x8          A2:H13
               table2 5x3           A1:C5
                empty 1x1           A1:A1
               table5 6x1           C3:C8
               table6 8x2           B1:C8
               table7 7x2           B2:C8
               lookup 4x9           B2:J5
         header_error 3x4           B2:E4
       named_ranges_2 4x5           A1:E4
         named_ranges 14x6          A2:F15

julia> XLSX.copysheet!(f[4])
12×8 XLSX.Worksheet: ["table (copy)"](A2:H13)

julia> f
XLSXFile("C:\\...\\general.xlsx") containing 14 Worksheets
            sheetname size          range
-------------------------------------------------
              general 10x6          A1:F10
               table3 5x6           A2:F6
               table4 4x3           E12:G15
                table 12x8          A2:H13
               table2 5x3           A1:C5
                empty 1x1           A1:A1
               table5 6x1           C3:C8
               table6 8x2           B1:C8
               table7 7x2           B2:C8
               lookup 4x9           B2:J5
         header_error 3x4           B2:E4
       named_ranges_2 4x5           A1:E4
         named_ranges 14x6          A2:F15
         table (copy) 12x8          A2:H13

```

"""
function copysheet!(ws::Worksheet, name::AbstractString="")::Worksheet
    wb = get_workbook(ws)
    is_chartsheet(wb, ws.name) && throw(XLSXError("Cannot copy a Chartsheet."))
    xl = get_xlsxfile(ws)
    !is_writable(get_xlsxfile(ws)) && throw(XLSXError("XLSXFile instance is not writable."))
    dim = get_dimension(ws)

    # make sure cache and XML are consistent
    update_worksheets_xml!(xl)

    pfx = get_prefix(ws)

    # create a copy of the XML document
    xdoc = copynode(get_worksheet_xml_document(ws))

    # if copied sheet is the currently selected sheet, do not copy this attribute over.
    # The original sheet will remain the only selected sheet.
    for c in XML.children(xdoc[end])
        if localname(c) =="sheetViews"
            for c2 in XML.children(c)
                if localname(c2) =="sheetView"
                    atts=XML.attributes(c2)
                    if haskey(atts, "tabSelected")
                        atts["tabSelected"]="0"
                    end
                end
            end
        end
    end

    # generate a new name for the copied sheet
    name = name == "" ? ws.name * " (copy)" : name

    # cache of copied sheet must be full
    ws.cache.is_full == true || throw(XLSXError("Cannot copy worksheet that does not have a full cache"))

    # copy the original worksheet cache to the new worksheet
    new_cache = WorksheetCache(
        true,
        ws.cache.cells,
        ws.cache.rows_in_cache,
        ws.cache.row_ht,
        ws.cache.row_index,
        SheetRowStreamIterator(ws), # Dummy - not needed because using full cache.
        nothing,
        ws.cache.dirty,
    )

    # insert the copied sheet into the workbook
    new_ws = insertsheet!(wb, xdoc, new_cache, ws.sst_count, pfx, name; dim)

    # Copy images if the sheet has a drawing
    sheet_path = get_relationship_target_by_id("xl", wb, ws.relationship_id)
    drawing_path = _drawing_path_for_sheet(xl, sheet_path)

    if drawing_path !== nothing
        src_drawing_file = rsplit(drawing_path, "/"; limit=2)[2]
        src_drawing_rels = "xl/drawings/_rels/$src_drawing_file.rels"

        # Pick a fresh drawing file name
        i = 1
        while haskey(xl.data, "xl/drawings/drawing$i.xml"); i += 1; end
        new_drawing_file = "drawing$i.xml"
        new_drawing_path = "xl/drawings/$new_drawing_file"
        new_drawing_rels = "xl/drawings/_rels/$new_drawing_file.rels"

        # Copy drawing XML and rels verbatim
        xl.data[new_drawing_path]  = copynode(xl.data[drawing_path])
        xl.files[new_drawing_path] = true
        xl.data[new_drawing_rels]  = copynode(xl.data[src_drawing_rels])
        xl.files[new_drawing_rels] = true

        # Register content types for drawing and any media it references
        register_content_type!(xl, "[Content_Types].xml";
                            tag="Override", key="PartName", val="/$new_drawing_path",
                            content_type=MIME_DRAWING)
        for img in _images_for_drawing(xl, drawing_path, ws.name)
            ext = detect_image_ext(xl.binary_data["xl/media/$(img.media_name)"])
            ext_no_dot = String(lstrip(ext, '.'))
            register_content_type!(xl, "[Content_Types].xml";
                                tag="Default", key="Extension", val=ext_no_dot,
                                content_type=get(EXT_MIME, ext, "image/$ext_no_dot"))
        end

        # Link drawing to new sheet using existing helpers
        new_sheet_path = get_relationship_target_by_id("xl", wb, new_ws.relationship_id)
        new_rels_path  = let (d, f) = rsplit(new_sheet_path, "/"; limit=2)
            "$d/_rels/$f.rels"
        end
        if !haskey(xl.data, new_rels_path)
            xl.data[new_rels_path]  = empty_rels_doc()
            xl.files[new_rels_path] = true
        end
        new_rels_root = root_element(xl.data[new_rels_path])
        rid = new_relationship_id(new_rels_root)
        pfx = get_prefix(new_rels_path, xl)
        push!(new_rels_root, XML.Element(prefixed_tag(pfx, "Relationship");
            Id=rid, Type=REL_DRAWING, Target="../drawings/$new_drawing_file",
        ))
        ensure_drawing_element!(xl, xl.data[new_sheet_path], new_sheet_path, rid)
    end

    # copy defined names from the original worksheet to the new worksheet
    ws_keys = [x for x in keys(wb.worksheet_names) if first(x) == ws.sheetId]
    for k in ws_keys
        val = wb.worksheet_names[k].value
        val = val isa CellRange ? string(val) :
              val isa SheetCellRange ? new_ws.name * "!" * string(val.rng) :
              val
        addDefinedName(new_ws, last(k), val; absolute=wb.worksheet_names[k].isabs)
    end

    # Copy the formula cache from the original worksheet to the new worksheet
    f_keys = [x for x in keys(wb.formulas) if x.sheet == ws.name]
    for k in f_keys
        newkey = SheetCellRef(new_ws.name, k.cellref)
        wb.formulas[newkey] = copy(wb.formulas[k])
    end

    return new_ws
end

function insertsheet!(wb::Workbook, xdoc::XML.Node, new_cache::WorksheetCache, sst_count::Int, pfx::String, name::AbstractString=""; dim=CellRange("A1:A1"))::Worksheet
    xf = get_xlsxfile(wb)
    !is_writable(xf) && throw(XLSXError("XLSXFile instance is not writable."))

    if name == ""
        new_name = ""
    else
        new_name = name
    end
    # ensure name is unique.
    i = 1
    current_sheet_names = sheetnames(wb)
    while new_name ∈ current_sheet_names || new_name == ""
        new_name = (name == "" ? "Sheet" : name * " ") * string(i)
        i += 1
    end

    new_name == "" && throw(XLSXError("Something wrong here!"))

    # checks if name is a unique sheet name
    function check_valid_sheetname(n::AbstractString)
        max_length = 31
        if length(n) > max_length
            throw(XLSXError("Invalid sheetname $n: must have at most $max_length characters. Found $(length(n))"))
        end

        if occursin(r"[:\\/\?\*\[\]]+", n)
            throw(XLSXError("Sheetname cannot contain characters: ':', '\\', '/', '?', '*', '[' or ']'."))
        end
    end

    check_valid_sheetname(new_name)

    # generate sheetId
    current_sheet_ids = [ws.sheetId for ws in wb.sheets]
    sheetId = max(current_sheet_ids...) + 1

    # generate a unique ID for the new sheet
    !haskey(xdoc[2], "xmlns:xr") && (xdoc[2]["xmlns:xr"] = "http://schemas.microsoft.com/office/spreadsheetml/2016/revision")
    xdoc[2]["xr:uid"] = "{" * uppercase(string(UUIDs.uuid4(wb.package.uuid_rng))) * "}"

    # generate a unique name for the XML
    local xml_filename::String
    i = 1
    while true
        xml_filename = "xl/worksheets/sheet" * string(i) * ".xml"
        #        if !in(xml_filename, keys(xf.files))
        if !haskey(xf.files, xml_filename)
            break
        end
        i += 1
    end

    # adds doc do XLSXFile
    xf.files[xml_filename] = true # is read
    xf.data[xml_filename] = xdoc
    xf.namespace[xml_filename] = pfx

    # adds workbook-level relationship
    # <Relationship Id="rId1" Target="worksheets/sheet1.xml" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"/>
    rId = add_relationship!(wb, xml_filename[4:end], "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet")

    # creates Worksheet instance
    ws = Worksheet(xf, sheetId, rId, new_name, dim, false)
    ws.cache = new_cache
    ws.sst_count = sst_count

    # adds the new sheet to the list of sheets in the workbook
    push!(wb.sheets, ws)

    # update [Content_Types].xml (fix for issue #275)
    add_override!(get_xlsxfile(wb), "/xl/worksheets/sheet" * string(sheetId) * ".xml", "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml")

    update_workbook_xml!(xf)

    return ws
end

add_override!(wb::Workbook, part::String, content::String) = add_override!(get_xlsxfile(wb), part, content)
function add_override!(xf::XLSXFile, part::String, content::String)
    wb = get_workbook(xf) 
    ctype_root = xmlroot(xf, "[Content_Types].xml")[end]
    localname(ctype_root) != "Types" && throw(XLSXError("Something wrong here!"))
    override_node = XML.Element("Override";
        PartName=part,
        ContentType=content
    )
    push!(ctype_root, override_node)
    return nothing
end
function renumber_files!(xf::XLSXFile, rId::String)
    wb = get_workbook(xf)
    id = parse(Int64, rId[4:end])

    # update active tab
    wbdoc = xmlroot(xf, "xl/workbook.xml")
    i, j = get_idces(wbdoc, "workbook", "bookViews")
    w = XML.children(wbdoc[i][j])
    if length(w) > 0
        for c in w
            if localname(c) == "workbookView"
                a = XML.attributes(c)
                if haskey(a, "activeTab")
                    at = parse(Int64, a["activeTab"])
                    if at >= id
                        c["activeTab"] = string(at - 1)
                    end
                end
            end
        end
    end
end


"""
    deletesheet!(ws::Worksheet) -> ::XLSXFile
    deletesheet!(wb::Workbook, name::AbstractString) -> ::XLSXFile
    deletesheet!(xf::XLSXFile, name::AbstractString) -> ::XLSXFile
    deletesheet!(xf::XLSXFile, sheetId::Integer) -> ::XLSXFile

Delete the given worksheet, the worksheet with the given name or the worksheet with the given `sheetId` from its `XLSXFile` 
(`sheetId` is a 1-based integer representing the order in which worksheet tabs are displayed in Excel).

!!! note "Caution"

    Cells in the other sheets that have references to the deleted sheet will fail when the sheet is deleted.
    The formulae are updated to contain a `#Ref!` error in place of each sheetcell reference.
    

See also [addsheet!](@ref), [renamesheet!](@ref), [copysheet!](@ref)

# Examples

```julia
julia> f = XLSX.opentemplate("general.xlsx")
XLSXFile("C:\\...\\general.xlsx") containing 13 Worksheets
            sheetname size          range
-------------------------------------------------
              general 10x6          A1:F10
               table3 5x6           A2:F6
               table4 4x3           E12:G15
                table 12x8          A2:H13
               table2 5x3           A1:C5
                empty 1x1           A1:A1        
               table5 6x1           C3:C8
               table6 8x2           B1:C8
               table7 7x2           B2:C8
               lookup 4x9           B2:J5
         header_error 3x4           B2:E4
       named_ranges_2 4x5           A1:E4
         named_ranges 14x6          A2:F15


julia> XLSX.deletesheet!(f[4])
XLSXFile("C:\\...\\general.xlsx") containing 12 Worksheets
            sheetname size          range
-------------------------------------------------
              general 10x6          A1:F10
               table3 5x6           A2:F6
               table4 4x3           E12:G15
               table2 5x3           A1:C5
                empty 1x1           A1:A1
               table5 6x1           C3:C8
               table6 8x2           B1:C8
               table7 7x2           B2:C8
               lookup 4x9           B2:J5
         header_error 3x4           B2:E4
       named_ranges_2 4x5           A1:E4
         named_ranges 14x6          A2:F15


julia> XLSX.deletesheet!(f, "table5")
XLSXFile("C:\\...\\general.xlsx") containing 11 Worksheets
            sheetname size          range
-------------------------------------------------
              general 10x6          A1:F10
               table3 5x6           A2:F6
               table4 4x3           E12:G15
               table2 5x3           A1:C5
                empty 1x1           A1:A1
               table6 8x2           B1:C8
               table7 7x2           B2:C8
               lookup 4x9           B2:J5
         header_error 3x4           B2:E4
       named_ranges_2 4x5           A1:E4
         named_ranges 14x6          A2:F15


julia> XLSX.deletesheet!(f, 1)
XLSXFile("C:\\...\\general.xlsx") containing 10 Worksheets
            sheetname size          range
-------------------------------------------------
               table3 5x6           A2:F6
               table4 4x3           E12:G15
               table2 5x3           A1:C5
                empty 1x1           A1:A1
               table6 8x2           B1:C8
               table7 7x2           B2:C8
               lookup 4x9           B2:J5
         header_error 3x4           B2:E4
       named_ranges_2 4x5           A1:E4
         named_ranges 14x6          A2:F15
```

"""
deletesheet!(ws::Worksheet) = deletesheet!(get_workbook(ws), ws.name)
deletesheet!(xl::XLSXFile, sheetId::Integer) = deletesheet!(get_workbook(xl), xl[sheetId].name)
deletesheet!(xl::XLSXFile, name::AbstractString) = deletesheet!(get_workbook(xl), name)
function deletesheet!(wb::Workbook, name::AbstractString)::XLSXFile
    hassheet(wb, name) || throw(XLSXError("Worksheet `$name` not found in workbook."))
    sheetcount(wb) > 1 || throw(XLSXError("`$name` is this workbook's only sheet. Cannot delete the only sheet!"))
    is_chartsheet(wb, name) && throw(XLSXError("Cannot delete a Chartsheet."))


    xf = get_xlsxfile(wb)

    # Worksheets and relationships
    s = (findfirst(s -> s.name == name, wb.sheets))
    sId = wb.sheets[s].sheetId
    rId = wb.sheets[s].relationship_id
    r = findfirst(y -> occursin("worksheet", y.Type) && y.Id == rId, wb.relationships)
    delete_relationships!(xf, wb.relationships[r])
    deleteat!(wb.relationships, r)
    deleteat!(wb.sheets, s)

    # Defined Names
    found_wbnames = Vector{String}()
    for (k, v) in wb.workbook_names
        wbn = v.value
        if typeof(wbn) <: DefinedNameRangeTypes
            if wbn.sheet == name
                push!(found_wbnames, k)
            end
        end
    end
    found_wsnames = Vector{Tuple{Int64,String}}()
    for (k, v) in wb.worksheet_names
        if first(k) == sId
            push!(found_wsnames, k)
        end
    end
    for key in found_wbnames
        delete!(wb.workbook_names, key)
    end
    for key in found_wsnames
        delete!(wb.worksheet_names, key)
    end
    renumber_keys = Vector{Pair{Tuple{Int64,String},Tuple{Int64,String}}}()
    for (k, _) in wb.worksheet_names
        first(k) == sId && throw(XLSXError("Something wrong here!"))
        if first(k) > sId
            push!(renumber_keys, k => (sId, last(k)))
        end
    end
    for (oldkey, newkey) in renumber_keys
        wb.worksheet_names[newkey] = wb.worksheet_names[oldkey]
        delete!(wb.worksheet_names, oldkey)
    end

    # Drawing and image cleanup
    sheet_path   = "xl/worksheets/sheet" * rId[4:end] * ".xml"
    drawing_path = _drawing_path_for_sheet(xf, sheet_path)

    if drawing_path !== nothing
        drawing_file     = rsplit(drawing_path, "/"; limit=2)[2]
        drawing_rels     = "xl/drawings/_rels/$drawing_file.rels"

        # Remove media — but only if no other sheet references it
        all_images = getImages(xf)
        deleted_images = _images_for_drawing(xf, drawing_path, name)
        deleted_names  = Set(img.media_name for img in deleted_images)
        still_used     = Set(img.media_name for img in all_images
                            if img.sheet != name)
        for media_name in deleted_names
            if media_name ∉ still_used
                delete!(xf.binary_data, "xl/media/$media_name")
            end
        end

        # Remove drawing XML and rels
        for path in (drawing_path, drawing_rels)
            delete!(xf.files, path)
            delete!(xf.data, path)
        end

        # Remove drawing Override from [Content_Types].xml
        ctype_root = xmlroot(xf, "[Content_Types].xml")[end]
        cont = XML.children(ctype_root)
        idx = findfirst(i -> haskey(cont[i], "PartName") &&
                            cont[i]["PartName"] == "/$drawing_path", eachindex(cont))
        idx !== nothing && deleteat!(cont, idx)

        # Remove sheet rels file (contains the drawing relationship)
        sheet_dir, sheet_file = rsplit(sheet_path, "/"; limit=2)
        sheet_rels = "$sheet_dir/_rels/$sheet_file.rels"
        delete!(xf.files, sheet_rels)
        delete!(xf.data,  sheet_rels)
    end

    # Files
    xml_filename = "xl/worksheets/sheet" * rId[4:end] * ".xml"
    if in(xml_filename, keys(xf.files))
        delete!(xf.files, xml_filename)
    end
    if in(xml_filename, keys(xf.data))
        delete!(xf.data, xml_filename)
    end
    if in(xml_filename, keys(xf.binary_data))
        delete!(xf.binary_data, xml_filename)
    end

    # update [Content_Types].xml
    ctype_root = xmlroot(get_xlsxfile(wb), "[Content_Types].xml")[end]
    localname(ctype_root) != "Types" && throw(XLSXError("Something wrong here!"))
    cont = XML.children(ctype_root)
    let idx = 0
        for (i, c) in enumerate(cont)
            if haskey(c, "PartName") && c["PartName"] == "/xl/worksheets/sheet" * rId[4:end] * ".xml"
                idx = i
                break
            end
        end
        if idx > 0
            deleteat!(cont, idx)
        end
    end

    update_formulas_missing_sheet!(wb, name)
    renumber_files!(xf, rId)
    update_workbook_xml!(xf)

    return xf
end

#
# Helper Functions
#

"""
    writetable(filename, data, columnnames; [overwrite], [sheetname])

- `data` is a vector of columns.
- `columnames` is a vector of column labels.
- `overwrite` is a `Bool` to control if `filename` should be overwritten if already exists.
- `sheetname` is the name for the worksheet.

Returns the filepath of the written file if a filename is supplied, or `nothing` if writing to an `IO`.

# Example

```julia
import XLSX
columns = [ [1, 2, 3, 4], ["Hey", "You", "Out", "There"], [10.2, 20.3, 30.4, 40.5] ]
colnames = [ "integers", "strings", "floats" ]
XLSX.writetable("table.xlsx", columns, colnames)
```

See also: [`XLSX.writetable!`](@ref).
"""
function writetable(filename::Union{AbstractString,IO}, data, columnnames; overwrite::Bool=false, sheetname::AbstractString="", anchor_cell::Union{String,CellRef}=CellRef("A1"))

    if filename isa AbstractString && !overwrite
        isfile(filename) && throw(XLSXError("$filename already exists."))
    end

    xf = open_empty_template(sheetname)
    sheet = xf[1]

    if isa(anchor_cell, String)
        anchor_cell = CellRef(anchor_cell)
    end

    writetable!(sheet, data, columnnames; anchor_cell=anchor_cell)

    # write output file
    writexlsx(filename, xf, overwrite=overwrite)
end

"""
    writetable(filename::Union{AbstractString, IO}; overwrite::Bool=false, kw...)
    writetable(filename::Union{AbstractString, IO}, tables::Vector{Tuple{String, Vector{Any}, Vector{String}}}; overwrite::Bool=false)

Write multiple tables.

`kw` is a variable keyword argument list. Each element should be in this format: `sheetname=( data, column_names )`,
where `data` is a vector of columns and `column_names` is a vector of column labels.

Returns the filepath of the written file if a filename is supplied, or `nothing` if writing to an `IO`.

Example:

```julia
julia> import DataFrames, XLSX

julia> df1 = DataFrames.DataFrame(COL1=[10,20,30], COL2=["Fist", "Sec", "Third"])

julia> df2 = DataFrames.DataFrame(AA=["aa", "bb"], AB=[10.1, 10.2])

julia> XLSX.writetable("report.xlsx", "REPORT_A" => df1, "REPORT_B" => df2)
```
"""
function writetable(filename::Union{AbstractString,IO}; overwrite::Bool=false, kw...)

    if filename isa AbstractString && !overwrite
        isfile(filename) && throw(XLSXError("$filename already exists."))
    end

    xf = open_empty_template()
    is_first = true

    for (sheetname, (data, column_names)) in kw
        if is_first
            # first sheet already exists in template file
            sheet = xf[1]
            renamesheet!(sheet, string(sheetname))
            writetable!(sheet, data, column_names)

            is_first = false
        else
            sheet = addsheet!(xf, string(sheetname))
            writetable!(sheet, data, column_names)
        end
    end

    # write output file
    writexlsx(filename, xf, overwrite=overwrite)

end

function writetable(filename::Union{AbstractString,IO}, tables::Vector{Tuple{String,S,Vector{T}}}; overwrite::Bool=false) where {S<:Vector{U} where {U},T<:Union{String,Symbol}}

    if filename isa AbstractString && !overwrite
        isfile(filename) && throw(XLSXError("$filename already exists."))
    end

    xf = open_empty_template()

    is_first = true
    for (sheetname, data, column_names) in tables
        if is_first
            # first sheet already exists in template file
            sheet = xf[1]
            renamesheet!(sheet, string(sheetname))
            writetable!(sheet, data, column_names)

            is_first = false
        else
            sheet = addsheet!(xf, string(sheetname))
            writetable!(sheet, data, column_names)
        end
    end

    # write output file
    writexlsx(filename, xf, overwrite=overwrite)
end
