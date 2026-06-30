#=
When a .xlsx file is read, the read may be lazy or eager and the approach differs depending on the
options selected, as follows:

- `mode="rw"`: every sheet's data fully, eagerly, parallel-cached at open.

- `mode="r"`, plain `openxlsx`/`readxlsx`: every sheet's structural XML decompressed/stripped 
    at open (cheap, scales with total file size); each sheet's actual row data lazily cache-filled only on 
    first access to that sheet. Any sheets from which cell data are never accessed do not get cached.

- `readtable`/`readtransposedtable`: only the one target worksheet's XML is ever 
    decompressed at all — structural processing for every other sheet is skipped entirely, and the 
    target sheet's row data is then cache-filled (or streamed, if enable_cache=false) exactly as in the 
    `mode="r"` case, just scoped to a single worksheet file from the very first byte read.
=#

# Name space conversion map for converting Strict OOXML files (ISO/IEC 29500) to Transitional format (ECMA-376)
const STRICT_TO_TRANSITIONAL = Dict(
    # core markup
    "http://purl.oclc.org/ooxml/spreadsheetml/main" =>
        "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "http://purl.oclc.org/ooxml/wordprocessingml/main" =>
        "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "http://purl.oclc.org/ooxml/presentationml/main" =>
        "http://schemas.openxmlformats.org/presentationml/2006/main",
    "http://purl.oclc.org/ooxml/drawingml/main" =>
        "http://schemas.openxmlformats.org/drawingml/2006/main",

    # drawingml sub-namespaces
    "http://purl.oclc.org/ooxml/drawingml/chartDrawing" =>
        "http://schemas.openxmlformats.org/drawingml/2006/chartDrawing",
    "http://purl.oclc.org/ooxml/drawingml/picture" =>
        "http://schemas.openxmlformats.org/drawingml/2006/picture",
    "http://purl.oclc.org/ooxml/drawingml/wordprocessingDrawing" =>
        "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
    "http://purl.oclc.org/ooxml/drawingml/spreadsheetDrawing" =>
        "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing",

    # officeDocument and relationships
    "http://purl.oclc.org/ooxml/officeDocument/relationships" =>
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "http://purl.oclc.org/ooxml/officeDocument/relationships/worksheet" =>
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet",
    "http://purl.oclc.org/ooxml/officeDocument/relationships/officeDocument" =>
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument",
    "http://purl.oclc.org/ooxml/officeDocument/relationships/sharedStrings" =>
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings",
    "http://purl.oclc.org/ooxml/officeDocument/relationships/styles" =>
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles",
    "http://purl.oclc.org/ooxml/officeDocument/relationships/theme" =>
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme",
    "http://purl.oclc.org/ooxml/officeDocument/relationships/chartsheet" =>
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chartsheet",
    "http://purl.oclc.org/ooxml/officeDocument/relationships/drawing" =>
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing",
    "http://purl.oclc.org/ooxml/officeDocument/relationships/extendedProperties" =>
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties",
    "http://purl.oclc.org/ooxml/officeDocument/relationships/customProperties" =>
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties",
    "http://purl.oclc.org/ooxml/officeDocument/relationships/printerSettings" =>
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/printerSettings",
    "http://purl.oclc.org/ooxml/officeDocument/relationships/calcChain" =>
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain",

    # docProps and vTypes
    "http://purl.oclc.org/ooxml/officeDocument/extendedProperties" =>
        "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties",
    "http://purl.oclc.org/ooxml/officeDocument/docPropsVTypes" =>
        "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes",
    "http://purl.oclc.org/ooxml/officeDocument/customProperties" =>
        "http://schemas.openxmlformats.org/officeDocument/2006/custom-properties",

    # customXml, math, bibliography
    "http://purl.oclc.org/ooxml/officeDocument/customXml" =>
        "http://schemas.openxmlformats.org/officeDocument/2006/customXml",
    "http://purl.oclc.org/ooxml/officeDocument/math" =>
        "http://schemas.openxmlformats.org/officeDocument/2006/math",
    "http://purl.oclc.org/ooxml/officeDocument/bibliography" =>
        "http://schemas.openxmlformats.org/officeDocument/2006/bibliography",

     # chart relationships
    "http://purl.oclc.org/ooxml/officeDocument/relationships/chart" =>
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart",
    "http://purl.oclc.org/ooxml/officeDocument/relationships/image" =>
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
    "http://purl.oclc.org/ooxml/officeDocument/relationships/hyperlink" =>
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
    "http://purl.oclc.org/ooxml/officeDocument/relationships/table" =>
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/table",
    "http://purl.oclc.org/ooxml/officeDocument/relationships/pivotTable" =>
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotTable",
    "http://purl.oclc.org/ooxml/officeDocument/relationships/pivotCacheDefinition" =>
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition",

    # chart namespace
    "http://purl.oclc.org/ooxml/drawingml/chart" =>
        "http://schemas.openxmlformats.org/drawingml/2006/chart",

    # markup compatibility
    "http://purl.oclc.org/ooxml/markup-compatibility/2006" =>
        "http://schemas.openxmlformats.org/markup-compatibility/2006",
)

@inline get_xlsxfile(wb::Workbook)::XLSXFile = wb.package
@inline get_xlsxfile(ws::Worksheet)::XLSXFile = ws.package
@inline get_workbook(ws::Worksheet)::Workbook = get_xlsxfile(ws).workbook
@inline get_workbook(xl::XLSXFile)::Workbook = xl.workbook

const ZIP_FILE_HEADER = [0x50, 0x4b, 0x03, 0x04]
const XLS_FILE_HEADER = [0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1] #[0xd0, 0xcf, 0x11, 0xe0]

function is_encrypted_xlsx(io::IO) # This function suggested by Claude AI

    # Read sector size from header (bytes 0x1E-0x1F)
    seek(io, 0x1E)
    sector_shift = read(io, UInt16)
    sector_size = 1 << sector_shift
    
    # Read the directory entries starting position (bytes 0x30-0x33)
    seek(io, 0x30)
    first_dir_sector = read(io, UInt32)
    
    # Calculate directory position
    dir_offset = 512 + first_dir_sector * sector_size
    
    # Read directory entries and look for encryption markers
    seek(io, dir_offset)
    
    # Check first several directory entries (each is 128 bytes)
    for i in 1:20
        entry_start = position(io)
        
        # Read name (64 bytes, UTF-16LE)
        name_bytes = read(io, 64)
        # Read name length in bytes (includes null terminator)
        name_length = read(io, UInt16)
        
        if name_length > 2 && name_length <= 64
            # Convert UTF-16LE to String
            # Take pairs of bytes and convert to Char
            chars = Char[]
            for j in 1:2:min(name_length-2, 64)
                if j+1 <= length(name_bytes)
                    code_point = UInt16(name_bytes[j]) | (UInt16(name_bytes[j+1]) << 8)
                    if code_point != 0
                        push!(chars, Char(code_point))
                    end
                end
            end
            name = String(chars)
            
            if occursin("EncryptionInfo", name) || occursin("EncryptedPackage", name)
                return true
            end
        end
        
        # Move to next directory entry (128 bytes total)
        seek(io, entry_start + 128)
    end
    return false
end

function check_for_xlsx_file_format(source::IO, label::AbstractString="input")
#    local header::Vector{UInt8}

    mark(source)
    header = Base.read(source, 8)
    reset(source)

    if header[1:4] == ZIP_FILE_HEADER # valid Zip file header
        return
    elseif header == XLS_FILE_HEADER # either an old XLS file or a password protected XLSX file
        if is_encrypted_xlsx(source) # Issue #251
            throw(XLSXError("`$label` looks like a password protected XLSX file. This package does not support password protected files. Consider using XLSXDecrypt.jl to decrypt the file first."))
        else
            throw(XLSXError("`$label` looks like an old XLS file (not XLSX). This package does not support XLS file format."))
        end
    else
        throw(XLSXError("`$label` is not a valid XLSX file."))
    end
end

function check_for_xlsx_file_format(filepath::AbstractString)
    !isfile(filepath) && throw(XLSXError("File $filepath not found."))

    open(filepath, "r") do io
        check_for_xlsx_file_format(io, filepath)
    end
end

@inline function localname(node)
    t = XML.tag(node)
    isnothing(t) && return nothing
    return _localname(t)
end
@inline localname(tag::AbstractString) = _localname(tag)

@inline function _localname(t::AbstractString)
    n = ncodeunits(t)
    @inbounds for i in 1:n
        codeunit(t, i) == 0x3a && return SubString(t, i+1)
    end
    return t
end

# Build a lookup dictionary for element names, qualified with the default namespace prefix if it exists.
function build_ns_dict!(xf::XLSXFile)
    ns = xf.namespace
    for (file_name, is_read) in xf.files
        is_read || continue
        haskey(ns, file_name) && continue  # ← already registered (e.g. by consumer)
        val = xf.data[file_name]
        if val isa String
            ns[file_name] = _get_ns_prefix_from_string(val)
            continue
        end
        doc = xmlroot(xf, file_name)
        els = xml_elements(doc)
        isempty(els) && continue
        xroot = last(els)
        ns[file_name] = get_default_namespace_prefix(xroot)
    end
    return nothing
end

# Extract the default namespace prefix from a raw XML string without parsing.
# Looks for xmlns:prefix="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
# or xmlns="..." to determine the prefix used for spreadsheet elements.
function _get_ns_prefix_from_string(xml_str::Union{String,Nothing})::Union{String,Nothing}
    isnothing(xml_str) && return nothing
    c = XML.Cursor(xml_str)
    # Advance to first Element — namespace is always on root element's opening tag
    while XML.next!(c) !== nothing
        XML.nodetype(c) == XML.Element || continue
        # Found root element — scan its attributes for spreadsheet namespace
        for (k, v) in XML.eachattribute(XML.LazyNode(c))
            v == "http://schemas.openxmlformats.org/spreadsheetml/2006/main" || continue
            # k is either "xmlns" (default ns) or "xmlns:prefix"
            k == "xmlns" && return nothing  # default namespace, no prefix
            return String(k)[7:end]  # strip "xmlns:" prefix
        end
        break  # only check root element
    end
    return nothing
end

function get_prefix(ws::Worksheet)
        internal_file_name = get_relationship_target_by_id("xl", get_workbook(ws), ws.relationship_id)
        pfx = get_prefix(internal_file_name, get_xlsxfile(ws))
        return something(pfx, "")
end
function get_prefix(file_name::String, xf::XLSXFile)::Union{Nothing,AbstractString}
    ns = get(xf.namespace, file_name, nothing)
    return something(ns, "")
end
# Returns the prefix (possibly "") that maps to the default/spreadsheet namespace,
# or `nothing` if there is no prefixed default. Used by `build_ns_dict!` for the
# Strict-OOXML namespace-prefix feature.
function get_default_namespace_prefix(r::XML.Node)
    nss = get_namespaces(r)
    isempty(nss) && return nothing
    if length(nss) == 1
        prefix = first(keys(nss))
        return prefix == "" ? nothing : prefix
    end
    haskey(nss, "") && return nothing
    for (k, v) in nss
        if v == SPREADSHEET_NAMESPACE_XPATH_ARG
            return k == "" ? nothing : k
        end
    end
    return nothing
end
# v0.4 contract: returns the default-namespace URI as a single `String`.
function get_default_namespace(r::XML.Node)::String
    nss = get_namespaces(r)

    # if only one namespace is defined, assume it is the default one
    # even if it has a prefix
    length(nss) == 1 && return first(values(nss))

    # otherwise, prefer the unprefixed default namespace
    haskey(nss, "") && return nss[""]

    # no unprefixed default (e.g. issues #380/#362/#267/#170): fall back to the
    # spreadsheet namespace even if it carries a prefix
    for (_, ns) in nss
        if ns == SPREADSHEET_NAMESPACE_XPATH_ARG
            return ns
        end
    end

    throw(XLSXError("No default namespace found."))
end
function get_namespaces(r::XML.Node)::Dict{String,String}
    nss = Dict{String,String}()
    atts = XML.attributes(r)
    isnothing(atts) && return nss   # XML.jl 0.4 returns `nothing` for attribute-less nodes
    for (key, value) in atts
        if startswith(key, "xmlns")
            colon_idx = findfirst(':', key)
            nss[isnothing(colon_idx) ? "" : SubString(key, colon_idx+1)] = value
        end
    end
    return nss
end
function get_sst_prefix(ws::Worksheet)::String
    sst_pfx = get_prefix("xl/SharedStrings.xml", get_xlsxfile(ws))
    if isnothing(sst_pfx) || sst_pfx == ""
        sst_pfx = ""
    else
        sst_pfx = sst_pfx*":"
    end
    return sst_pfx
end


# Determine if the file is a Strict OOXML file.
function is_strict_ooxml(xf::XLSXFile)::Bool
    wb = get_workbook(xf)
    files = xf.data

    # Primary check: conformance attribute on workbook root
    if haskey(files, "xl/workbook.xml")
        wbNode = xml_root_element(files["xl/workbook.xml"])
        attrs = XML.attributes(wbNode)
        if !isnothing(attrs)
            if get(attrs, "conformance", "") == "strict"
                return true
            end
            # Also catch strict namespace declarations on root element
            if any(occursin("purl.oclc.org/ooxml", v) for (_, v) in attrs)
                return true
            end
        end
    end

    # Fallback: check relationship types in _rels/.rels
    if haskey(files, "_rels/.rels")
        rels = xml_root_element(files["_rels/.rels"])
        for el in xml_elements(rels)
            if localname(el) == "Relationship"
                relattrs = XML.attributes(el)
                if !isnothing(relattrs) && occursin("purl.oclc.org/ooxml", get(relattrs, "Type", ""))
                    return true
                end
            end
        end
    end

    return false
end

"""
    opentemplate(source::Union{AbstractString, IO}) :: XLSXFile

Read an existing Excel (`.xlsx`) file as a template and return as a writable `XLSXFile` for editing 
and saving to another file with [XLSX.writexlsx](@ref).

A convenience function equivalent to `openxlsx(source; mode="rw", enable_cache=true)`

# Examples
```julia
julia> xf = opentemplate("myExcelFile.xlsx")
```

"""
opentemplate(source::Union{AbstractString,IO})::XLSXFile = open_or_read_xlsx(source, true, true, true)

@inline open_xlsx_template(source::Union{AbstractString,IO})::XLSXFile = open_or_read_xlsx(source, true, true, true)

"""
    newxlsx([sheetname::AbstractString]; update_timestamp::Bool) :: XLSXFile

Return an empty, writable `XLSXFile` with 1 worksheet for editing and 
subsequent saving to a file with [XLSX.writexlsx](@ref).
By default, the worksheet is `Sheet1`. Specify `sheetname` to give the worksheet a different name.

Use keyword argument `update_timestamp=false` to prevent timestamps in the file properties from being 
updated to the current date/time. This ensures bit-for-bit reproducible output when the file is written.
The file `Date` will remain as `2018-05-22T02:41:32Z`.
The default is `update_timestamp=true`, resulting in the `Date` being set to the current UTC time in the new file.

# Examples
```julia
julia> xf = XLSX.newxlsx()

julia> xf = XLSX.newxlsx("MySheet")
```

"""
newxlsx(sheetname::AbstractString=""; update_timestamp::Bool=true)::XLSXFile = open_empty_template(sheetname; update_timestamp)

function fix_datestamp!(xf::XLSXFile)
    # Fix datestamp in blank.xlsx. It is specified in the file `docProps/core.xml`.
    # These two dates dictate the created and modified dates shown in Excel file properties
    # and in Windows File Explorer.
    # The values in the file are `2018-05-22T02:41:32Z` and `2018-05-22T02:42:04Z` respectively.
    # Reset them to current date/time.
    f = "docProps/core.xml"
    time_now=Dates.now(Dates.UTC)
    date_format = Dates.dateformat"yyyy-mm-ddTHH:MM:SSZ"
    i, j = get_idces(xf.data[f], "cp:coreProperties", "dcterms:created")
    any(isnothing, [i, j]) || (xf.data[f][i][j][1]=Dates.format(time_now, date_format))
    i, j = get_idces(xf.data[f], "cp:coreProperties", "dcterms:modified")
    any(isnothing, [i, j]) || (xf.data[f][i][j][1]=Dates.format(time_now+Dates.Second(1), date_format))
    return nothing
end

include_dependency(joinpath(@__DIR__, "data", "blank.xlsx"))
const BLANK_XLSX_DATA = read(joinpath(@__DIR__, "data", "blank.xlsx"))

function open_empty_template(
    sheetname::AbstractString="";
    empty_template_data::Vector{UInt8}=BLANK_XLSX_DATA,
    update_timestamp::Bool=true
)::XLSXFile
    xf = open_xlsx_template(IOBuffer(empty_template_data))
    xf[1].cache.is_full = true

    if sheetname != ""
        renamesheet!(xf[1], sheetname)
    end
    xf.source = "blank.xlsx"
    update_timestamp && fix_datestamp!(xf) # blank.xlsx has fixed datestamp in 2018 that should be updated to now.
    return xf
end

"""
    readxlsx(source::Union{AbstractString, IO}) :: XLSXFile

Main function for reading an Excel file.
This function will read the whole Excel file into memory
and return an XLSXFile.

Functionally equivalent to ``openxlsx(source; mode="r", enable_cache=true)`

Consider using [`XLSX.openxlsx`](@ref) for lazy loading of Excel file contents.
"""
@inline readxlsx(source::Union{AbstractString,IO})::XLSXFile = open_or_read_xlsx(source, true, true, false)

"""
    openxlsx(f::F, source::Union{AbstractString, IO}; mode::AbstractString="r", enable_cache::Bool=true) where {F<:Function}

Open an XLSX file for reading and/or writing and applies the function `f` to the content.

# `Do` syntax

This function should be used with `do` syntax, like in:

```julia
XLSX.openxlsx("myfile.xlsx") do xf
    # read data from `xf`
end
```

# Filemodes

The `mode` argument controls how the file is opened. The following modes are allowed:

* `r` : read-only mode. The existing data in `source` will be accessible for reading. This is the **default** mode.

* `w` : write mode. Opens an empty file that will be written to `source`. If source already exists it will be overwritten.

* `rw` : edit mode. Opens `source` for editing. The file will be saved (overwritten) to disk when the function ends.

!!! warning

    Using do-block syntax in "rw" mode will overwrite the file you read in with the modified data when the do block ends.
    Care is needed to ensure data are not inadvertantly overwritten, especially if the xlsx file contains any elements 
    that `XLSX.jl` cannot process (such as charts, pivot tables, etc), but that would otherwise be preserved if not 
    overwritten. You may avoid this risk by choosing to open files in "rw" mode without using do-block syntax, in which 
    case it becomes necessary explicitly to write the `XLSXFile` out again, providing the option to write to another file name.
    

!!! note

    When a native Excel template (`.xltx`) file is opened in "rw" mode using do-block syntax, it will always be written 
    back out as a regular Excel file with a `.xlsx` extension at the termination of the do block. It is not written out 
    as an Excel template file.

# Arguments

* `source` is IO or the complete path to the file.

* `mode` is the file mode, as explained in the last section.

* `enable_cache`:

If `enable_cache=true` and the file is opened in read-only mode, each worksheet's cells 
are loaded and cached the first time any cell on that sheet is read — the whole sheet is 
loaded at once, not just the cell you asked for. Subsequent reads of any cell on that 
sheet then come from the cache instead of reading from disk. This lazy, per-sheet caching 
is most efficient if you only intend to access one (or a few) sheets in a multi-sheet 
workbook. If you intend to access all or most sheets in full, the parallel caching used 
in write mode may be more efficient for your use case.

If `enable_cache=true` and the file is opened in write mode, all worksheets are eagerly 
loaded into the cache in parallel as the file is opened (they will be needed for writing 
anyway).

If `enable_cache=false`, worksheet cells are read by streaming through the worksheet's 
XML directly, without building a persistent cache of parsed cell data. This avoids the 
memory and time cost of constructing that cache, at the expense of needing to re-scan 
from the current position for each row accessed. It's most useful for a single, 
sequential pass over a worksheet's rows (e.g. converting it once to a table), where 
building a cache you'll never reuse would be wasted work.

The default value is `enable_cache=true`.

# Examples

## Read from file

The following example shows how you would read worksheet cells, one row at a time,
where `myfile.xlsx` is a spreadsheet that doesn't fit into memory.

```julia
julia> XLSX.openxlsx("myfile.xlsx", enable_cache=false) do xf
          for r in eachrow(xf["mysheet"])
              # read something from row `r`
          end
       end
```

## Write a new file

```julia
XLSX.openxlsx("new.xlsx", mode="w") do xf
    sheet = xf[1]
    sheet[1, :] = [1, Date(2018, 1, 1), "test"]
end
```

## Edit an existing file

```julia
XLSX.openxlsx("edit.xlsx", mode="rw") do xf
    sheet = xf[1]
    sheet[2, :] = [2, Date(2019, 1, 1), "add new line"]
end
```

See also [`XLSX.readxlsx`](@ref).
"""
function openxlsx(f::F, source::Union{AbstractString,IO};
    mode::AbstractString="r", enable_cache::Bool=true) where {F<:Function}

    _read, _write = parse_file_mode(mode)

    if _read
        if !(source isa IO || isfile(source))
            throw(XLSXError("File $source not found."))
        end
        xf = open_or_read_xlsx(source, _read, enable_cache, _write)
    else
        xf = open_empty_template()
        xf.source = source
    end

    try
        f(xf)

    finally
        if _write
            if xf.template_type != NotATemplate
                if isa(xf.source, AbstractString)
                    ext = xf.template_type == XLTMTemplate ? ".xlsm" : ".xlsx"
                    xf.source = splitext(xf.source)[1] * ext
                end
            end
            writexlsx(xf.source, xf, overwrite=true)
        end
    end
end

"""
    openxlsx(source::Union{AbstractString, IO}; mode="r", enable_cache=true) :: XLSXFile

Supports opening an XLSX file without using do-syntax.

If opened with mode="rw" then use [`savexlsx`](@ref) to save the XLSXFile back to `source`, 
overwriting the original file.
Alternatively, use [`writexlsx`](@ref) to write the XLSXFile to a different filename.

These two invocations of `openxlsx` are functionally equivalent:
```
XLSX.openxlsx("myfile.xlsx", mode="rw") do xf
    # Do some processing on the content
end

xf = openxlsx("myfile.xlsx", mode="rw")
# Do some processing on the content
XLSX.savexlsx(xf)

```
"""
function openxlsx(source::Union{AbstractString,IO};
    mode::AbstractString="r",
    enable_cache::Bool=true)::XLSXFile

    _read, _write = parse_file_mode(mode)

    if _read
        if !(source isa IO || isfile(source))
            throw(XLSXError("File $source not found."))
        end
        return open_or_read_xlsx(source, _read, enable_cache, _write)
    else
        xf = open_empty_template()
        xf.source = source
        return xf
    end
end

function parse_file_mode(mode::AbstractString)::Tuple{Bool,Bool}
    m = lowercase(mode)
    if m == "r"
        return (true, false)
    elseif m == "w"
        return (false, true)
    elseif m == "rw" || m == "wr"
        return (true, true)
    else
        throw(XLSXError("Couldn't parse file mode $mode."))
    end
end

# Convert a strict OOXML file to transitional format in-place by remapping
# `purl.oclc.org/ooxml` namespaces and relationship types to their
# `schemas.openxmlformats.org` equivalents, and dropping the `conformance="strict"` attribute.
# Remap a single element's attributes from strict OOXML to transitional, in place.
# XML.jl 0.4 exposes attributes as a read-only `Attributes` view, so mutate the
# node directly (`node[k] = v`) and drop attributes via the backing vector.
function _strict_to_transitional_node!(node::XML.Node, filename::AbstractString)
    isnothing(node.attributes) && return nothing
    # Snapshot keys/values first since we mutate while inspecting.
    pairs = collect(node.attributes)
    for (k, v) in pairs
        if k == "conformance" && v == "strict"
            filter!(p -> first(p) != "conformance", node.attributes)
        elseif startswith(v, "http://purl.oclc.org/ooxml")
            if haskey(STRICT_TO_TRANSITIONAL, v)
                node[k] = STRICT_TO_TRANSITIONAL[v]
            else
                throw(XLSXError("Unsupported strict OOXML namespace or relationship type: \"$v\" in $filename. Please open an issue at https://github.com/JuliaData/XLSX.jl/issues"))
            end
        elseif k == "Type" && startswith(v, "http://purl.oclc.org/ooxml")
            if haskey(STRICT_TO_TRANSITIONAL, v)
                node[k] = STRICT_TO_TRANSITIONAL[v]
            else
                throw(XLSXError("Unsupported strict OOXML relationship type: \"$v\" in $filename. Please open an issue at https://github.com/JuliaData/XLSX.jl/issues"))
            end
        end
    end
    return nothing
end

function convert_strict_to_transitional!(xf::XLSXFile, pass::Int)
    for filename in keys(xf.files)
        should_process = if pass == 1
            !occursin(r"xl/worksheets/sheet\d*\.xml|xl/sharedStrings\.xml", filename)
        elseif pass == 2
            occursin(r"xl/sharedStrings\.xml", filename)
        else  # pass == 3
            occursin(r"xl/worksheets/sheet\d*\.xml", filename)
        end

        if should_process
            if occursin(r"xl/worksheets/sheet\d*\.xml", filename)
                # Worksheet data (whether the full raw or the stripped stub) is
                # always stored as a plain String. Fix root-element namespace
                # attributes with cheap string substitution rather than parsing
                # the whole tree — avoids materializing sheetData just to patch
                # a handful of xmlns attributes.
                data = xf.data[filename]
                if data isa String
                    converted = data
                    for (strict_ns, transitional_ns) in STRICT_TO_TRANSITIONAL
                        converted = replace(converted, strict_ns => transitional_ns)
                    end
                    converted = replace(converted, r"\s+conformance\s*=\s*\"strict\""=>"")
                    xf.data[filename] = converted
                end
                continue
            end

            data = xf.data[filename]
            if data isa String
                data = parse(data, XML.Node)
                xf.data[filename] = data
            end
            els = xml_elements(data)
            isempty(els) && continue
            xroot = last(els)
            _strict_to_transitional_node!(xroot, filename)
            for el in xml_elements(xroot)
                _strict_to_transitional_node!(el, filename)
            end
        end
    end
    return nothing
end
#=function convert_strict_to_transitional!(xf::XLSXFile, pass::Int)

    for filename in keys(xf.files)
        should_process = if pass == 1
            !occursin(r"xl/worksheets/sheet\d*\.xml|xl/sharedStrings\.xml", filename)
        elseif pass == 2
            occursin(r"xl/sharedStrings\.xml", filename)
        else  # pass == 3
            occursin(r"xl/worksheets/sheet\d*\.xml", filename)
        end
           
        if should_process
            data = xf.data[filename]
            # Lazy parse if stored as deferred String
            if data isa String
                data = parse(data, XML.Node)
                xf.data[filename] = data
            end
            els = xml_elements(data)            # SST/worksheet files are stored as lightweight placeholders with no
            # element children; nothing to remap there.
            isempty(els) && continue
            xroot = last(els)
            _strict_to_transitional_node!(xroot, filename)

            # For .rels files, also patch Type= on child Relationship elements
            for el in xml_elements(xroot)
                _strict_to_transitional_node!(el, filename)
            end
        end
    end

    return nothing
end
=#

function open_or_read_xlsx(source::Union{IO,AbstractString}, _read::Bool, enable_cache::Bool, _write::Bool;
                            target_sheet::Union{Nothing,AbstractString,Integer}=nothing,
                            load_formulas::Bool=true)::XLSXFile

    if _write
        !(_read && enable_cache) && throw(XLSXError("Cache must be enabled for files in `write` mode."))
    end
    xf = XLSXFile(source, enable_cache, _write, load_formulas)
    
    if source isa IO
        zip_io = ZipArchives.ZipReader(read(source))
    else
        zip_io = ZipArchives.ZipReader(FileArray(abspath(source)))
    end

    load_files!(xf, zip_io; pass=1)
    strict = is_strict_ooxml(xf)
    if strict
        convert_strict_to_transitional!(xf, 1)
    end

    build_ns_dict!(xf)
    check_minimum_requirements(xf)
    parse_relationships!(xf)
    parse_workbook!(xf)
    remove_calcChain!(xf)

    load_files!(xf, zip_io; pass=2)
    if strict
        convert_strict_to_transitional!(xf, 2)
    end

    # Resolve target_sheet to worksheet filename before pass 3
    # Only for read-only single-sheet access (readtable) — not for writable opens
    target_file = if !isnothing(target_sheet) && !_write
        wb = get_workbook(xf)
        ws = if target_sheet isa Integer
            wb.sheets[target_sheet]
        else
            s = findfirst(s -> s.name == target_sheet, wb.sheets)
            isnothing(s) ? nothing : wb.sheets[s]
        end
        isnothing(ws) ? nothing : get_relationship_target_by_id("xl", wb, ws.relationship_id)
    else
        nothing
    end

    load_files!(xf, zip_io; pass=3, target_file=target_file)
    if strict
        convert_strict_to_transitional!(xf, 3)
    end

    for sheet in get_workbook(xf).sheets
        if isnothing(sheet.dimension)
            # Only compute dimension for sheets that were loaded
            if isnothing(target_file) || 
                get_relationship_target_by_id("xl", get_workbook(xf), sheet.relationship_id) == target_file
                sheet.dimension = read_worksheet_dimension(xf, sheet.relationship_id, sheet.name)
            end
        end
    end

    return xf
end

"""
    ensure_workbook_is_xlsx!(xf::XLSXFile)

Inspect the `[Content_Types].xml` part of the package and ensure that
`/xl/workbook.xml` is marked as a regular `.xlsx` workbook. If the workbook
content type indicates a template (`.xltx`), the function converts it in-place
by rewriting the workbook `ContentType` to the standard `.xlsx` value and
updating `xf.source` to use a `.xlsx` file extension. Throws an `XLSXError`
if the workbook override is missing or has an unknown content type.
"""
function ensure_workbook_is_xlsx!(xf::XLSXFile)
    root = xml_root_element(xf.data["[Content_Types].xml"])

    workbook_override = nothing
    default_xml_type = nothing

     for child in XML.children(root)
        name = localname(child)

        if name == "Override" && lowercase(child["PartName"]) == "/xl/workbook.xml"
            workbook_override = child

        elseif name == "Default" && child["Extension"] == "xml"
            default_xml_type = child["ContentType"]
        end
    end

    ctype =
        !isnothing(workbook_override) ? workbook_override["ContentType"] :
        !isnothing(default_xml_type) ? default_xml_type :
        throw(XLSXError("Malformed XLSX: workbook.xml content type not found."))

    # Passthrough types — no conversion needed
    ctype == "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml" && return nothing
    ctype == "application/vnd.ms-excel.sheet.macroEnabled.main+xml"                       && return nothing

   # Template types — convert to their workbook equivalent
    template_conversions = (
        "application/vnd.openxmlformats-officedocument.spreadsheetml.template.main+xml" =>
            ("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml", XLTXTemplate),
        "application/vnd.ms-excel.template.macroEnabled.main+xml" =>
            ("application/vnd.ms-excel.sheet.macroEnabled.main+xml", XLTMTemplate),
    )

    for (template_ctype, (target_ctype, template_type)) in template_conversions
        ctype == template_ctype || continue

        if !isnothing(workbook_override)
            workbook_override["ContentType"] = target_ctype
        else
            for child in XML.children(root)
                if localname(child) == "Default" && child["Extension"] == "xml"
                    child["ContentType"] = target_ctype
                end
            end
        end

        xf.template_type = template_type
        return nothing
    end

    # Unknown workbook type
    throw(XLSXError("Unknown workbook content type: $ctype"))
end


# See section 12.2 - Package Structure
function check_minimum_requirements(xf::XLSXFile)
    mandatory_files = ["_rels/.rels",
        "xl/workbook.xml",
        "[Content_Types].xml",
        "xl/_rels/workbook.xml.rels"
    ]

    for f in mandatory_files
        !in(f, filenames(xf)) && throw(XLSXError("Malformed XLSX File. Couldn't find file $f in the package."))
    end

    # Further check if this is a valid `.xlsx` file.
    ensure_workbook_is_xlsx!(xf)

    return nothing
end

# Parses package level relationships defined in `_rels/.rels`.
# Parses workbook level relationships defined in `xl/_rels/workbook.xml.rels`.
function parse_relationships!(xf::XLSXFile)
    wb = get_workbook(xf)

    # package level relationships
    xroot = get_package_relationship_root(xf)
    for el in xml_elements(xroot)
        push!(xf.relationships, Relationship(wb, el))
    end
    isempty(xf.relationships) && throw(XLSXError("Relationships not found in _rels/.rels!"))

    # workbook level relationships
    xroot = get_workbook_relationship_root(xf)
    for el in xml_elements(xroot)
        push!(wb.relationships, Relationship(wb, el))
    end
    isempty(wb.relationships) && throw(XLSXError("Relationships not found in xl/_rels/workbook.xml.rels"))

    nothing
end

# Updates xf.workbook from xf.data[\"xl/workbook.xml\"]
function parse_workbook!(xf::XLSXFile)
    xroot = xml_root_element(xmlroot(xf, "xl/workbook.xml"))
    wb = get_workbook(xf)

    localname(xroot) != "workbook" && throw(XLSXError("Malformed xl/workbook.xml. Root node name should be 'workbook'. Got '$(localname(xroot))'."))

    # date1904
    wb.date1904 = false
    for node in XML.children(xroot)
        localname(node) != "workbookPr" && continue
        attrs = XML.attributes(node)
        if !isnothing(attrs) && haskey(attrs, "date1904")
            v = attrs["date1904"]
            if v ∈ ("1", "true")
                wb.date1904 = true
            elseif v ∉ ("0", "false")
                throw(XLSXError("Could not parse xl/workbook -> workbookPr -> date1904 = $v."))
            end
        end
        break
    end

    # sheets
    wb.sheets = Worksheet[]
    for node in xml_elements(xroot)
        localname(node) != "sheets" && continue
        for sheet_node in xml_elements(node)
            localname(sheet_node) != "sheet" && throw(XLSXError("Unsupported node $(localname(sheet_node)) in node $(localname(node)) in 'xl/workbook.xml'."))
            push!(wb.sheets, Worksheet(xf, sheet_node))
        end
        break
    end

    # named ranges
    for node in xml_elements(xroot)
        localname(node) != "definedNames" && continue
        for dn_node in xml_elements(node)
            localname(dn_node) != "definedName" && continue

            raw = XML.value(dn_node[1])
            name = XML.attributes(dn_node)["name"]

            defined_value, isabs = parse_defined_name_value(raw)

            attrs = XML.attributes(dn_node)
            if haskey(attrs, "localSheetId")
                localSheetId = parse(Int, attrs["localSheetId"]) + 1
                sheetId = wb.sheets[localSheetId].sheetId
                wb.worksheet_names[(sheetId, name)] = DefinedNameValue(defined_value, isabs)
            else
                wb.workbook_names[name] = DefinedNameValue(defined_value, isabs)
            end
        end
        break
    end
end

function parse_defined_name_value(s::String)::Tuple{DefinedNameValueTypes, Any}
    unquote_sheet(str) = let sp = split(str, '!')
        unquoteit(sp[1]) * "!" * sp[2]
    end

    if is_valid_non_contiguous_range(s)
        parts = split(s, ',')
        rng = [String(unquote_sheet(r)) for r in parts]
        defined_value = NonContiguousRange(join(rng, ','))
        isabs = [is_valid_fixed_sheet_cellname(d) || is_valid_fixed_sheet_cellrange(d) for d in parts]
        length(isabs) != length(defined_value.rng) && throw(XLSXError("Error parsing absolute references in non-contiguous range."))
    elseif is_valid_fixed_sheet_cellname(s)
        defined_value, isabs = SheetCellRef(unquote_sheet(s)), true
    elseif is_valid_sheet_cellname(s)
        defined_value, isabs = SheetCellRef(unquote_sheet(s)), false
    elseif is_valid_fixed_sheet_cellrange(s)
        defined_value, isabs = SheetCellRange(unquote_sheet(s)), true
    elseif is_valid_sheet_cellrange(s)
        defined_value, isabs = SheetCellRange(unquote_sheet(s)), false
    elseif startswith(s, '"') && endswith(s, '"')
        inner = String(chop(s, head=1, tail=1))
        defined_value, isabs = (isempty(inner) ? missing : inner), false
    elseif (n = tryparse(Int64, s)) !== nothing
        defined_value, isabs = n, false
    elseif (n = tryparse(Float64, s)) !== nothing
        defined_value, isabs = n, false
    elseif isempty(s)
        defined_value, isabs = missing, false
    else
        defined_value, isabs = string(s), false
    end

    return defined_value, isabs
end

# Returns a Dict mapping Workbook <externalReferences>: index => relationship id.
function get_wb_ext_refs(xf::XLSXFile)
    wb = get_workbook(xf)
    ext_refs = Dict{Int, String}()
    xroot = xmlroot(xf, "xl/workbook.xml")
    i, j = get_idces(xroot, "workbook", "externalReferences")
    if !isnothing(j)
        for (i, ref) in enumerate(xml_elements(xroot[i][j]))
            ext_refs[i] = ref["r:id"]
        end
    end
    return ext_refs
end

# delete Override PartName=calcChain since this was never loaded (#31)
function remove_calcChain!(xf::XLSXFile)
    xf.data["[Content_Types].xml"]
    ctype_root = xml_root_element(xmlroot(xf, "[Content_Types].xml"))
    for (i, c) in enumerate(XML.children(ctype_root))
        if XML.tag(c) == "Override" && haskey(c, "PartName") && c["PartName"]=="/xl/calcChain.xml"
            deleteat!(ctype_root.children, i)
            break
        end
    end
end
# Lists internal files from the XLSX package.
@inline filenames(xl::XLSXFile) = keys(xl.files)

# Returns true if the file data was read into xl.data.
@inline function internal_xml_file_isread(xl::XLSXFile, filename::String)::Bool
    return xl.files[filename]
end
@inline internal_xml_file_exists(xl::XLSXFile, filename::String)::Bool = haskey(xl.files, filename)

function internal_xml_file_add!(xl::XLSXFile, filename::String)
    !(endswith(filename, ".xml") || endswith(filename, ".rels")) && throw(XLSXError("Something wrong here!"))
    xl.files[filename] = false
    nothing
end

function strip_bom_and_lf!(bytes::Vector{UInt8})
    # Issue 243 - Need to remove BOM characters that precede the XML declaration.
    # Note: If an Excel file containing a BOM is opened in Excel itself and 
    # subsequently saved, Excel will strip the BOM out. This means the test for 
    # this issue will stop testing the fix if the file "BOM - issue243.xlsx" is 
    # opened in Excel because the offending BOM will have been removed.
    length(bytes) < 3 && return
    if bytes[1] == 0xEF && bytes[2] == 0xBB && bytes[3] == 0xBF
        if length(bytes) > 3 && bytes[4] == 0x0A
            deleteat!(bytes, 1:4)
        else
            deleteat!(bytes, 1:3)
        end
    end
end

function splitNode(xml_str::String, skipnode::String)
    c = XML.Cursor(xml_str)

    XML.next!(c)
    while !XML.eof(c) && XML.nodetype(c) != XML.Element
        XML.next!(c)
    end
    XML.eof(c) && return xml_str, ""

    target_lazy = nothing

    while XML.next!(c) !== nothing
        XML.depth(c) == 0 && break
        XML.depth(c) != 2 && (XML.skip_element!(c); continue)
        XML.nodetype(c) == XML.Element || continue
        if localname(c) == skipnode
            target_lazy = XML.LazyNode(c)
            XML.skip_element!(c)
            break
        end
        XML.skip_element!(c)
    end

    isnothing(target_lazy) && return xml_str, ""

    target_tag    = XML.tag(target_lazy)
    subtree_start = target_lazy.token.offset + 1

    # Use skip_element! for end position — raw byte scan, no sourcetext needed
    c2 = XML.Cursor(target_lazy)
    XML.next!(c2)          # land on element
    XML.skip_element!(c2)  # jump past entire subtree
    subtree_end = c2.st.state.pos   # byte just past </sheetData>

    attrs = XML.attributes(target_lazy)
    replacement = if isnothing(attrs) || isempty(attrs)
        "<$(target_tag)/>"
    else
        attr_str = join(("$(k)=\"$(v)\"" for (k,v) in attrs), " ")
        "<$(target_tag) $(attr_str)/>"
    end

    stripped_xml = xml_str[1:subtree_start-1] * replacement * xml_str[subtree_end:end]
    return stripped_xml, ""
end

const BINARY_PREFIXES = ["customxml"] # must be lowercase
is_binary_path(filename) = any(p -> startswith(lowercase(filename), p), BINARY_PREFIXES)

function stream_files(xf::XLSXFile, zip_io::ZipArchives.ZipReader; pass::Int,
                      channel_size::Int=1 << 8)

        Channel{String}(channel_size) do out
        for f in ZipArchives.zip_names(zip_io)

            # ignore xl/calcChain.xml in any case (#31)
            if f != "xl/calcChain.xml"

                if pass==1 && !is_binary_path(f) && (endswith(f, ".xml") || endswith(f, ".rels"))
                    # Identify usable xml files in XLSXFile
                    internal_xml_file_add!(xf, f)
                end
                put!(out, f)
            end
        end
    end
end

# Read xml files in three passes
# pass 1 - read all but worksheets and sharedStrings
# pass 2 - only read sharedStrings (needed before worksheets)
# pass 3 - only read worksheets
function load_files!(xf::XLSXFile, zip_io::ZipArchives.ZipReader; pass::Int,
                     target_file::Union{Nothing,String}=nothing)

    (pass < 1 || pass > 3) && throw(XLSXError("Unknown pass to read files."))
    wb = get_workbook(xf)

    read_files = Channel{ReadFile}(1 << 20)
    all_files = stream_files(xf, zip_io; pass)

    filtered_files = Channel{String}(1 << 20) do out
        for file in all_files
            is_sst = occursin(r"^xl/sharedStrings\.xml$", file)
            is_worksheet = occursin(r"^xl/worksheets/[^/]+\.xml$", file)
            should_process = if pass == 1
                !is_sst && !is_worksheet
            elseif pass == 2
                is_sst
            else  # pass == 3
                if !isnothing(target_file)
                    is_worksheet && file == target_file  # single sheet only
                else
                    is_worksheet  # all sheets
                end
            end
            if should_process
                put!(out, file)
            end
        end
    end

    consumer = @async begin
        fill_tasks = Task[]

        for file in read_files
            if !isnothing(file.node)
                xf.data[file.name] = file.node
                xf.files[file.name] = true
            end
            if xf.is_writable || pass == 2
                if occursin("xl/sharedStrings.xml", file.name)
                    if has_sst(wb)
                        xf.namespace[file.name] = _get_ns_prefix_from_string(file.raw)
                        if xf.is_writable
                            sst_load!(wb)
                        end
                    end
                end
            end
            if !isnothing(file.raw) && occursin(r"xl/worksheets/", file.name)
                xf.namespace[file.name] = _get_ns_prefix_from_string(file.raw)
                will_eager_fill = xf.use_cache_for_sheet_data && xf.is_writable
                if will_eager_fill
                    # store stub now; full raw only needed transiently for the fill below
                    xf.data[file.name] = file.node
                    xf.files[file.name] = true
                    for sheet in wb.sheets
                        target = get_relationship_target_by_id("xl", wb, sheet.relationship_id)
                        if target == file.name
                            local captured_sheet = sheet
                            local captured_raw = file.raw
                            t = Threads.@spawn begin
                                lznode = parse(captured_raw, XML.LazyNode)
                                first_cache_fill!(captured_sheet, lznode)
                            end
                            push!(fill_tasks, t)
                        end
                    end
                elseif xf.use_cache_for_sheet_data
                    # cache-on, read-only: keep full raw resident so streaming/lazy-fill/
                    # dimension lookups can use it without re-reading from ZIP.
                    # (unchanged from prior behaviour)
                    xf.data[file.name] = file.raw
                    xf.files[file.name] = true
                else
                    # enable_cache=false: do NOT retain worksheet XML in memory.
                    # Mark the file as known-present (internal_xml_file_exists checks
                    # xf.files) but leave it out of xf.data entirely, so consumers
                    # (SheetRowStreamIterator, match_rows, read_worksheet_dimension)
                    # fall through to on-demand reads via open_internal_file_stream,
                    # which re-decompresses+reparses straight from the FileArray-backed
                    # ZipReader each time, and never retains the parsed tree afterward.
                    xf.files[file.name] = true
                end
            end
            if !isnothing(file.bin)
                xf.binary_data[file.name] = file.bin
            end
        end

        for t in fill_tasks
            wait(t)
        end
    end

    @sync for _ in 1:MAX_THREADS
        Threads.@spawn begin
            for file in filtered_files
                readfile = process_file(zip_io, file)
                put!(read_files, readfile)
            end
        end
    end

    close(read_files)
    wait(consumer)
end
function process_file(zip_io::ZipArchives.ZipReader, filename::String)

    node = nothing
    raw  = nothing
    bin  = nothing

    try
        bytes = ZipArchives.zip_readentry(zip_io, filename)
        if !is_binary_path(filename) && (endswith(filename, ".xml") || endswith(filename, ".rels"))
            strip_bom_and_lf!(bytes)
            xml_str = String(bytes)
            if filename == "xl/sharedStrings.xml"
                node = XML.Element("sst")  # placeholder; SST is loaded via sst_load!
                raw  = xml_str
            elseif occursin(r"xl/worksheets/sheet\d*\.xml", filename)
                stripped_xml, _ = splitNode(xml_str, "sheetData")
                node = stripped_xml
                raw  = xml_str      # full worksheet for LazyNode construction
            else
                node = parse(xml_str, XML.Node)
            end
        else
            bin = bytes
        end
    catch err
        throw(XLSXError("Failed to parse internal XML file `$filename`"))
    end

    return ReadFile(node, raw, bin, filename)
end

function get_xml_data(xf::XLSXFile, filename::String)::XML.Node
    val = xf.data[filename]
    if val isa String
        parsed = parse(val, XML.Node)
        xf.data[filename] = parsed
        return parsed
    end
    return val::XML.Node
end
function internal_xml_file_read(xf::XLSXFile, filename::String)
    !internal_xml_file_exists(xf, filename) && throw(XLSXError("Couldn't find $filename in $(xf.source)."))
    !internal_xml_file_isread(xf, filename) && throw(XLSXError("$filename in $(xf.source) has not been read."))
    val = get_xml_data(xf,filename)
    return val::XML.Node
end
function internal_xml_file_read(xf::XLSXFile, zip_io::Union{Nothing,ZipArchives.ZipReader}, filename::String)

    !internal_xml_file_exists(xf, filename) && throw(XLSXError("Couldn't find $filename in $(xf.source)."))

    if !internal_xml_file_isread(xf, filename)

        try
            bytes = ZipArchives.zip_readentry(zip_io, filename)
            strip_bom_and_lf!(bytes)
            xml_str = String(bytes)
            if filename == "xl/sharedStrings.xml"
                xf.data[filename] = XML.Element("sst")  # placeholder; SST is loaded via sst_load!
            elseif occursin(r"xl/worksheets/sheet\d*\.xml", filename)
                xf.data[filename], _ = splitNode(xml_str, "sheetData")
            else
                xf.data[filename] = parse(xml_str, XML.Node)
            end
            xf.files[filename] = true # set file as read
        catch err
            throw(XLSXError("Failed to parse internal XML file `$filename`"))
        end

    end

    return xf.data[filename]
end

# Utility method to find the XMLDocument associated with a given package filename.
# Returns xl.data[filename] if it exists. Throws an error if it doesn't.
@inline xmldocument(xl::XLSXFile, filename::String)::XML.Node = internal_xml_file_read(xl, filename)

# Utility method to return the root element of a given XMLDocument from the package.
@inline xmlroot(xl::XLSXFile, filename::String)::XML.Node = xmldocument(xl, filename)
function xmlroot(wb::Workbook, rId::String)
    filename = get_relationship_target_by_id("xl", wb, rId)
    return xmldocument(get_xlsxfile(wb), filename)
end

#
# Helper Functions
#

"""
    readdata(source, sheet, ref)
    readdata(source, sheetref)

Return a scalar, vector or matrix with values from a spreadsheet file.
'ref' can be a defined name, a cell reference or a cell, column, row 
or non-contiguous range.


See also [`XLSX.getdata`](@ref).

# Examples

These function calls are equivalent.

```julia
julia> XLSX.readdata("myfile.xlsx", "mysheet", "A2:B4")
3×2 Array{Any,2}:
 1  "first"
 2  "second"
 3  "third"

julia> XLSX.readdata("myfile.xlsx", 1, "A2:B4")
3×2 Array{Any,2}:
 1  "first"
 2  "second"
 3  "third"

julia> XLSX.readdata("myfile.xlsx", "mysheet!A2:B4")
3×2 Array{Any,2}:
 1  "first"
 2  "second"
 3  "third"
```

Non-contiguous ranges return vectors of Array{Any, 2} with an entry for every non-contiguous (comma-separated) 
element in the range.

```julia
julia> XLSX.readdata("customXml.xlsx", "Mock-up", "Location") # `Location` is a `definedName` for a non-contiguous range
4-element Vector{Matrix{Any}}:
 ["Here";;]
 [missing;;]
 [missing;;]
 [missing;;]
```
"""
function readdata(source::Union{AbstractString,IO}, sheet::Union{AbstractString,Int}, ref)
    c = openxlsx(source, enable_cache=false) do xf
        getdata(getsheet(xf, sheet), ref)
    end
    return c
end

function readdata(source::Union{AbstractString,IO}, sheetref::AbstractString)
    c = openxlsx(source, enable_cache=false) do xf
        getdata(xf, sheetref)
    end
    return c
end

"""
    readtable(
        source,
        [sheet,
        [columns]];
        [first_row],
        [column_labels],
        [header],
        [infer_eltypes],
        [stop_in_empty_row],
        [stop_in_row_function],
        [enable_cache],
        [keep_empty_rows],
        [normalizenames],
        [missing_strings]
    ) -> DataTable

Returns tabular data from a spreadsheet as a struct `XLSX.DataTable`.
Use this function to create a `DataFrame` from package `DataFrames.jl` 
(or other `Tables.jl`` compatible object).

If `sheet` is not given, the first sheet in the `XLSXFile` will be used.

`readtable` always loads only the requested sheet, regardless of how many
other sheets exist in the workbook — other sheets are never read or
cached, so reading from a single sheet of a large multi-sheet workbook is
efficient.

Use `columns` argument to specify which columns to get.
For example, `"B:D"` will select columns `B`, `C` and `D`.
If `columns` is not given, the algorithm will find the first sequence
of consecutive non-empty cells. A valid `sheet` must be specified 
when specifying `columns`.

Use `first_row` to indicate the first row of the table.
`first_row=5` will look for a table starting at sheet row `5`.
If `first_row` is not given, the algorithm will look for the first
non-empty row in the spreadsheet.

`header` is a `Bool` indicating if the first row is a header.
If `header=true` and `column_labels` is not specified, the column labels
for the table will be read from the first row of the table.
If `header=false` and `column_labels` is not specified, the algorithm
will generate column labels. The default value is `header=true`.

Use `column_labels` to specify names for the header of the table.

Use `normalizenames=true` to normalize column names to valid Julia identifiers.

Use `missing_strings` to specify strings that should be interpreted as 
`missing` values in the resulting table. `missing_strings` can be a single 
string (e.g. `"N/A"`) or a vector of strings (e.g. `["N/A", "NULL"]`). 
The default value is `missing_strings=nothing`.

Use `infer_eltypes=true` to get `data` as a `Vector{Any}` of typed vectors.
The default value is `infer_eltypes=true`.

`stop_in_empty_row` is a boolean indicating whether an empty row marks the 
end of the table. If `stop_in_empty_row=false`, the `TableRowIterator` will 
continue to fetch rows until there's no more rows in the Worksheet or range.
The default behavior is `stop_in_empty_row=true`.

`stop_in_row_function` is a Function that receives a `TableRow` and returns
 a `Bool` indicating if the end of the table was reached.

Example for `stop_in_row_function`:

```julia
function stop_function(r)
    v = r[:col_label]
    return !ismissing(v) && v == "unwanted value"
end
```

`enable_cache` is a boolean that determines whether cell data are loaded 
into the worksheet cache on reading. Using `readtable` with `enable_cache=true` 
is faster than with `enable_cache=false` for large files, but uses more 
memory. The default behavior is `enable_cache=true`.

`keep_empty_rows` determines whether rows where all column values are equal 
to `missing` are kept (`true`) or dropped (`false`) from the resulting table. 
`keep_empty_rows` never affects the *bounds* of the table; the number of 
rows read from a sheet is only affected by `first_row`, `stop_in_empty_row` 
and `stop_in_row_function` (if specified).
`keep_empty_rows` is only checked once the first and last row of the table 
have been determined, to see whether to keep or drop empty rows between the 
first and the last row.
The default behavior is `keep_empty_rows=false`.

# Example

```julia
julia> using DataFrames, XLSX

julia> df = DataFrame(XLSX.readtable("myfile.xlsx", "mysheet"))
```

See also: [`XLSX.gettable`](@ref), [`XLSX.readto`](@ref).
"""
function readtable(source::Union{AbstractString,IO}; 
    first_row::Union{Nothing,Int}=nothing, 
    column_labels=nothing, 
    header::Bool=true, 
    infer_eltypes::Bool=true, 
    stop_in_empty_row::Bool=true, 
    stop_in_row_function::Union{Nothing,Function}=nothing, 
    enable_cache::Bool=true, 
    keep_empty_rows::Bool=false, 
    normalizenames::Bool=false, 
    missing_strings::Union{AbstractString, AbstractVector{<:AbstractString}, Nothing}=nothing
)
    if !(source isa IO || isfile(source))
        throw(XLSXError("File $source not found."))
    end
    xf = open_or_read_xlsx(source, true, enable_cache, false; target_sheet=1, load_formulas=false)
    return gettable(getsheet(xf, 1); first_row, column_labels, header, infer_eltypes, stop_in_empty_row, stop_in_row_function, keep_empty_rows, normalizenames, missing_strings)
end

function readtable(source::Union{AbstractString,IO}, sheet::Union{AbstractString,Int}; 
    first_row::Union{Nothing,Int}=nothing, 
    column_labels=nothing, 
    header::Bool=true, 
    infer_eltypes::Bool=true, 
    stop_in_empty_row::Bool=true, 
    stop_in_row_function::Union{Nothing,Function}=nothing, 
    enable_cache::Bool=true, 
    keep_empty_rows::Bool=false, 
    normalizenames::Bool=false,
    missing_strings::Union{AbstractString, AbstractVector{<:AbstractString}, Nothing}=nothing
)
    if !(source isa IO || isfile(source))
        throw(XLSXError("File $source not found."))
    end
    xf = open_or_read_xlsx(source, true, enable_cache, false; target_sheet=sheet, load_formulas=false)
    return gettable(getsheet(xf, sheet); first_row, column_labels, header, infer_eltypes, stop_in_empty_row, stop_in_row_function, keep_empty_rows, normalizenames, missing_strings)
end

function readtable(source::Union{AbstractString,IO}, sheet::Union{AbstractString,Int}, columns::ColumnRange; 
    first_row::Union{Nothing,Int}=nothing, 
    column_labels=nothing, 
    header::Bool=true, 
    infer_eltypes::Bool=true, 
    stop_in_empty_row::Bool=true, 
    stop_in_row_function::Union{Nothing,Function}=nothing, 
    enable_cache::Bool=true, 
    keep_empty_rows::Bool=false, 
    normalizenames::Bool=false,
    missing_strings::Union{AbstractString, AbstractVector{<:AbstractString}, Nothing}=nothing
)
    if !(source isa IO || isfile(source))
        throw(XLSXError("File $source not found."))
    end
    xf = open_or_read_xlsx(source, true, enable_cache, false; target_sheet=sheet, load_formulas=false)
    return gettable(getsheet(xf, sheet), columns; first_row, column_labels, header, infer_eltypes, stop_in_empty_row, stop_in_row_function, keep_empty_rows, normalizenames, missing_strings)
end

function readtable(source::Union{AbstractString,IO}, sheet::Union{AbstractString,Int}, range::AbstractString; 
    first_row::Union{Nothing,Int}=nothing, 
    column_labels=nothing, 
    header::Bool=true, 
    infer_eltypes::Bool=true, 
    stop_in_empty_row::Bool=true, 
    stop_in_row_function::Union{Nothing,Function}=nothing, 
    enable_cache::Bool=true, 
    keep_empty_rows::Bool=false, 
    normalizenames::Bool=false,
    missing_strings::Union{AbstractString, AbstractVector{<:AbstractString}, Nothing}=nothing
)
    if is_valid_column_range(range)
        range = ColumnRange(range)
    else
        throw(XLSXError("The columns argument must be a valid column range."))
    end
    return readtable(source, sheet, range; first_row, column_labels, header, infer_eltypes, stop_in_empty_row, stop_in_row_function, enable_cache, keep_empty_rows, normalizenames, missing_strings)
end

"""
    readto(
        source,
        [sheet,
        [columns]],
        sink;
        [first_row],
        [column_labels],
        [header],
        [infer_eltypes],
        [stop_in_empty_row],
        [stop_in_row_function],
        [enable_cache],
        [keep_empty_rows],
        [normalizenames],
        [missing_strings]
    ) -> sink

Read and parse an Excel worksheet, materializing directly using the 
`sink` function, which can be any `Tables.jl`-compatible function 
(e.g. `DataFrame`, `StructArray` or `TypedTable``).

Takes the same keyword arguments as [`XLSX.readtable`](@ref) 

# Example

```julia
julia> using DataFrames, StructArrays, TypedTables, XLSX

julia> df = XLSX.readto("myfile.xlsx", DataFrame)

julia> sa = XLSX.readto("myfile.xlsx", StructArray)

julia> tt = XLSX.readto("myfile.xlsx", Table) # from TypedTables.jl

julia> df = XLSX.readto("myfile.xlsx", "mysheet", DataFrame)

julia> df = XLSX.readto("myfile.xlsx", "mysheet", "A:C", DataFrame)
```

See also: [`XLSX.gettable`](@ref).
"""
function readto(source::Union{AbstractString,IO}, sheet::Union{AbstractString,Int}, range::AbstractString, sink=nothing; kw...)
    if sink === nothing
        throw(XLSXError("provide a valid sink argument, like `using DataFrames; XLSX.readto(source, sheet, columns, DataFrame)`"))
    end
    return Tables.CopiedColumns(readtable(source, sheet, range; kw...)) |> sink
end
function readto(source::Union{AbstractString,IO}, sheet::Union{AbstractString,Int}, sink=nothing; kw...)
    if sink === nothing
        throw(XLSXError("provide a valid sink argument, like `using DataFrames; XLSX.readto(source, sheet, DataFrame)`"))
    end
    return Tables.CopiedColumns(readtable(source, sheet; kw...)) |> sink
end
function readto(source::Union{AbstractString,IO}, sink=nothing; kw...)
    if sink === nothing
        throw(XLSXError("provide a valid sink argument, like `using DataFrames; XLSX.readto(source, DataFrame)`"))
    end
    return Tables.CopiedColumns(readtable(source; kw...)) |> sink
end

#---------------------------------------------------------------------------------------------- Transposed Table

"""
    readtransposedtable(
        source,
        [sheet,
        [rows]];
        [first_column],
        [column_labels],
        [header],
        [normalizenames]
    ) -> DataTable

Read a transposed table from an Excel file, `source`, in which data are arranged in 
rows rather than columns in a worksheet. For example:
```
Category      "A", "B", "C", "D"
"variable 1"  10,  20,  30,  40
"variable 2"  15,  25,  35,  40
"variable 3"  20,  30,  40,  50
```
Returns data from a worksheet as a struct `XLSX.DataTable` which
can be passed directly to any function that accepts `Tables.jl` data.
(e.g. `DataFrame` from package `DataFrames.jl`).

If `sheet` is not given, the first sheet in the `XLSXFile` will be used.

Use the `rows` argument to specify which worksheeet rows to include.
For example, `"2:7"` will select rows 2 to 7 (inclusive).
If `rows` is not given, the algorithm will find the first sequence
of consecutive non-empty cells. If `rows` includes leading or trailing 
rows that are completely empty, these rows will be omitted from the 
returned table. In any case, the table will be truncated at the first 
and last non-empty rows, even if this range is smaller than `rows`. 
A valid `sheet` must be specified when specifying `rows`.

Use `first_column` to indicate the first column of the table. May be given 
as a column number or as a string, so that `first_column="E"` and
`first_column=5` will both look for a table starting at column `5` ("E").
Any leading completely empty columns will be ignored, including 
the `first_column`. If `first_column` is not given, the algorithm will 
look for the first non-empty column in the spreadsheet.

`header` is a `Bool` indicating if the first row is a header.
If `header=true` and `column_labels` is not specified, the column labels
for the table will be read from the first column of the table.
If `header=false` and `column_labels` is not specified, the algorithm
will generate column labels. The default value is `header=true`.

Use `column_labels` as a vector of symbols to specify names for the 
header of the table. If `header=true` and `column_labels` is also given, 
column_labels will be preferred and the first column of the table will 
be ignored.

Use `normalizenames=true` to normalize column names to valid Julia identifiers. 
The default is `normalizenames=false`.

# Examples

```julia
julia> using DataFrames, XLSX, PrettyTables

julia> DataFrame(readtransposedtable("HTable.xlsx", "Example"))
4×4 DataFrame
 Row │ Category  Variable 1  Variable 2  Variable 3 
     │ String    Int64       Int64       Int64
─────┼──────────────────────────────────────────────
   1 │ A                 10          15          20
   2 │ B                 20          25          30
   3 │ C                 30          35          40
   4 │ D                 40          40          50

julia> PrettyTable(readtransposedtable("HTable.xlsx", "Multiple", "2:7"; first_column=13))
┌──────┬───────┬───────┬───────┬──────────┬────────────┐
│ date │ name1 │ name2 │ name3 │    name4 │      name5 │
├──────┼───────┼───────┼───────┼──────────┼────────────┤
│ 1840 │  12.4 │ 0.045 │  true │ 10:22:00 │      Hello │
│ 1841 │  12.6 │ 0.046 │  true │ 10:23:00 │ 2025-12-19 │
│ 1842 │  12.8 │ 0.047 │ false │ 10:24:00 │          3 │
│ 1843 │  13.0 │ 0.048 │  true │ 10:25:00 │       3.33 │
│ 1844 │  13.2 │ 0.049 │ false │ 10:26:00 │      Hello │
│ 1845 │  13.4 │  0.05 │  true │ 10:27:00 │ 2025-12-19 │
│ 1846 │  13.6 │ 0.051 │  true │ 10:28:00 │          3 │
│ 1847 │  13.8 │ 0.052 │  true │ 10:29:00 │       3.33 │
│ 1848 │  14.0 │ 0.053 │ false │ 10:30:00 │       true │
└──────┴───────┴───────┴───────┴──────────┴────────────┘
```

See also: [`XLSX.gettransposedtable`](@ref), [`XLSX.readtable`](@ref).
"""
function readtransposedtable(filename::AbstractString, sheetname::AbstractString, rows::AbstractString; first_column=nothing, column_labels=nothing, header::Bool=true, normalizenames::Bool=false)
    if !isfile(filename)
        throw(XLSXError("File $filename not found."))
    end
    xf = open_or_read_xlsx(filename, true, true, false; load_formulas=false, target_sheet=sheetname)
    hassheet(xf, sheetname) || throw(XLSX.XLSXError("Sheet $sheetname not found in file $filename"))
    return gettransposedtable(xf[sheetname], rows; first_column, column_labels, header, normalizenames)
end

function readtransposedtable(filename::AbstractString, sheetname::AbstractString; first_column=nothing, column_labels=nothing, header::Bool=true, normalizenames::Bool=false)
    if !isfile(filename)
        throw(XLSXError("File $filename not found."))
    end
    xf = open_or_read_xlsx(filename, true, true, false; load_formulas=false, target_sheet=sheetname)
    hassheet(xf, sheetname) || throw(XLSX.XLSXError("Sheet $sheetname not found in file $filename"))
    dim = get_dimension(xf[sheetname])
    return gettransposedtable(xf[sheetname], "$(dim.start.row_number):$(dim.stop.row_number)"; first_column, column_labels, header, normalizenames)
end

function readtransposedtable(filename::AbstractString; first_column=nothing, column_labels=nothing, header::Bool=true, normalizenames::Bool=false)
    if !isfile(filename)
        throw(XLSXError("File $filename not found."))
    end
    xf = open_or_read_xlsx(filename, true, true, false; target_sheet=1, load_formulas=false)
    dim = get_dimension(xf[1])
    return gettransposedtable(xf[1], "$(dim.start.row_number):$(dim.stop.row_number)"; first_column, column_labels, header, normalizenames)
end


# Hooks for FileIOExt.jl

"""
```julia
    FileIO.load(
        source::String,
        [sheet::String,
        [columns::String]];
        [first_row::Int],
        [first_column::String]
        [column_labels::Vector{String}],
        [header::Bool],
        [normalizenames::Bool],
        [transpose::Bool]
    )
```
Read tabular data from an Excel file, `source`, and return it as a `Tables.jl` compatible table.
The resulting table object can be passed directly to any function that accepts `Tables.jl` data 
(e.g. `DataFrame` from package `DataFrames.jl`).

This function requires both FileIO.jl v1.20.0 or higher to be active in the current environment and a Julia version >= v1.9.

#### Arguments:

* `source`: The name of the file to be loaded.
* `sheet`: Specifies the sheet name to be loaded. If `sheet` is not given, the first Excel sheet in the file will be used.
* `columns`: Determines which columns to read. For example, `"B:D"` will select columns B, C and D. If columns is not given, the algorithm will find the first sequence of consecutive non-empty cells. A valid sheet **must** be specified when specifying columns. If `transpose = true` or is omitted, `columns` should be used to specify rows. For example, specifying `"2:4"` with `transpose = true` will read only from these rows.

!!! note

    The file extension provided in `source` must be `.xlsx`, `.xltx`, `.xlsm`, 
    or `.xltm` for FileIO to recognize the file format as an Excel file.

#### Keywords:

* `first_row`: Indicates the first row of the data table to be read. For example, `first_row=5` will look for a table starting at sheet row 5. If first_row is not given, the algorithm will look for the first non-empty row in the sheet (ignored if `transpose = true`).
* `first_column`: Indicates the first row of the data table to be read. For example, `first_column="B"` will look for a table starting at sheet row 5. If first_row is not given, the algorithm will look for the first non-empty row in the sheet (ignored if `transpose = false` or is omitted).
* `column_labels`: Specifies column names for the header of the table. If `column_labels` are given and `header=true`, the headers given by `column_labels` will be used, and the first row of the table (containing headers) will be ignored.
* `header`: Indicates if the first row (column if `transpose = true`) is a header. If `header=true` and `column_labels` is not specified, the column labels for the table will be read from the first row (column) of the table. If `header=false` and `column_labels` is not specified, the algorithm will generate column labels. The default value is `header=true`.
* `normalizenames`: Set to `true` to normalize column names to valid Julia identifiers. Default=`false`.
* `transpose`: Set to `true` to transpose the table to read data from rows not columns.

#### Examples

```julia
julia> PrettyTable(load("HTable.xlsx", "Offset"; first_row=2))

julia> df = DataFrame(load("HTable.xlsx", "Offset", "2:7"; transpose=true, first_column="B"))

julia> df = DataFrame(load("HTable.xlsx"; normalizenames=true, transpose=true, column_labels=["Date", "Name1", "Name2", "Name3", "Name4", "Name5"]))

```
"""
function load(args...; kwargs...)
    throw(XLSXError(
        """
        load requires the FileIO.jl package.

        Please install and load it with:

            using Pkg
            Pkg.add("FileIO")
            using FileIO

        Then retry FileIO.load.
        """
    ))

    return nothing
end

"""
```julia
    FileIO.save(
        source::String,
        data;
        [sheetname::String],
        [overwrite::Bool]
    )
```
Save a `Tables.jl` compatible table to an Excel file, `source`.

This function requires both FileIO.jl v1.20.0 or higher to be active in the current environment and a Julia version >= v1.9.

#### Arguments:

* `source`: The name of the file to be created on save.
* `data`: A `Tables.jl` compatible table to be saved to the file. For example, a `DataFrame` from package `DataFrames.jl`.

!!! note

    The file extension provided in `source` must be `.xlsx`, `.xltx`, `.xlsm`, 
    or `.xltm` for FileIO to recognize the file format as an Excel file. The 
    file created will be a standard workbook (ie not an Excel template nor a 
    macro-enabled workbook) regardless of which of these four extensions is used.

#### Keywords:

* `sheetname`: Specify the sheetname to be used in the created file. By default, the sheetname will be `Sheet1`.
* `overwrite`: Set `overwrite=true` to overwite any existing file of the same name. Default = `false`.

#### Examples

```julia
julia> save("myfile.xlsx", myTable; sheetname="myname", overwrite=true)
```
"""
function save(args...; kwargs...)
    throw(XLSXError(
        """
        save requires the FileIO.jl package.

        Please install and load it with:

            using Pkg
            Pkg.add("FileIO")
            using FileIO

        Then retry FileIO.save.
        """
    ))

    return nothing
end
