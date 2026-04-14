
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
            throw(XLSXError("`$label` looks like a password protected XLSX file. This package does not support password protected files."))
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


"""
    opentemplate(source::Union{AbstractString, IO}) :: XLSXFile

Read an existing Excel (`.xlsx`) file as a template and return as a writable `XLSXFile` for editing 
and saving to another file with [XLSX.writexlsx](@ref).

A convenience function equivalent to `openxlsx(source; mode="rw", enable_cache=true)`

!!! note
    XLSX.jl only works with `.xlsx` files and cannot work with Excel `.xltx` template files. 
    Reading as a template in this package merely means opening a `.xlsx` file to edit, update 
    and then write as an updated `.xlsx` file (e.g. using `XLSX.writexlsx()`). Doing so retains 
    the formatting and layout of the opened file, but this is not the same as using a `.xltx` file.

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

# Arguments

* `source` is IO or the complete path to the file.

* `mode` is the file mode, as explained in the last section.

* `enable_cache`:

If `enable_cache=true` and the file is opened in read-only mode, all worksheet cells 
will be cached as they are read the first time. When you read a worksheet cell for the 
second (or subsequent) time it will use the cached value instead of reading from disk.
If `enable_cache=true` and the file is opened in write mode, all cells are eagerly read 
into the cache as the file is opened (they will be needed at write anyway). For very 
large files, this can take a few seconds.

If `enable_cache=false`, worksheet cells will always be read from disk.
This is useful when you want to read a spreadsheet that doesn't fit into memory.

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
            writexlsx(source, xf, overwrite=true)
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
    if mode == "r"
        return (true, false)
    elseif mode == "w"
        return (false, true)
    elseif mode == "rw" || mode == "wr"
        return (true, true)
    else
        throw(XLSXError("Couldn't parse file mode $mode."))
    end
end

function open_or_read_xlsx(source::Union{IO,AbstractString}, _read::Bool, enable_cache::Bool, _write::Bool)::XLSXFile
    # sanity check
    if _write
        !(_read && enable_cache) && throw(XLSXError("Cache must be enabled for files in `write` mode."))
    end
    xf = XLSXFile(source, enable_cache, _write)

    #if enable_cache || (source isa IO)
    if source isa IO
        zip_io = ZipArchives.ZipReader(read(source))
    else
        zip_io = ZipArchives.ZipReader(FileArray(abspath(source))) # FileArray is marginally slower than Mmap
#       zip_io = ZipArchives.ZipReader(Mmap.mmap(abspath(source))) # but Mmap is unreliable : https://discourse.julialang.org/t/struggling-to-use-mmap-with-ziparchives/129839
    end


    load_files!(xf, zip_io; pass=1) # multi-threaded file load

    check_minimum_requirements(xf)
    parse_relationships!(xf)
    parse_workbook!(xf)

    # need to remove calcChain.xml from [Content_Types].xml since file is never loaded
    remove_calcChain!(xf)

    load_files!(xf, zip_io; pass=2) # Need to load sst before worksheets
    load_files!(xf, zip_io; pass=3) # load worksheets last so inlineStrings go after existing ssts

    for sheet in get_workbook(xf).sheets
        if isnothing(sheet.dimension)
            sheet.dimension = read_worksheet_dimension(xf, sheet.relationship_id, sheet.name)
        end
    end

    return xf
end
function get_namespaces(r::XML.Node)::Dict{String,String}
    nss = Dict{String,String}()
    for (key, value) in XML.attributes(r)
        if startswith(key, "xmlns")
            colon_idx = findfirst(':', key)
            if isnothing(colon_idx)
                nss[""] = value
            else
                nss[SubString(key, colon_idx+1)] = value
            end
        end
    end
    return nss
end
function get_default_namespace(r::XML.Node)::String
    nss = get_namespaces(r)

    # in case that only one namespace is defined, assume that it is the default one
    # even if it has a prefix
    length(nss) == 1 && return first(values(nss))

    # otherwise, look for the default namespace (without prefix)
    for (prefix, ns) in nss
        if prefix == ""
            return ns
        end
    end

    throw(XLSXError("No default namespace found."))
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
    f = "[Content_Types].xml"
    content_types = XML.write(xf.data[f])

    if occursin("spreadsheetml.sheet", content_types)
        return nothing
    elseif occursin("spreadsheetml.template", content_types)
        throw(XLSXError("XLSX.jl does not support Excel template files (`.xltx` files).\nSave template as an `xlsx` file type first."))
    else
        throw(XLSXError("Unknown Excel file type."))
    end

    nothing
end

# Parses package level relationships defined in `_rels/.rels`.
# Parses workbook level relationships defined in `xl/_rels/workbook.xml.rels`.
function parse_relationships!(xf::XLSXFile)

    # package level relationships
    xroot = get_package_relationship_root(xf)
    for el in XML.children(xroot)
        push!(xf.relationships, Relationship(el))
    end
    isempty(xf.relationships) && throw(XLSXError("Relationships not found in _rels/.rels!"))

    # workbook level relationships
    wb = get_workbook(xf)
    xroot = get_workbook_relationship_root(xf)
    for el in XML.children(xroot)
        push!(wb.relationships, Relationship(el))
    end
    isempty(wb.relationships) && throw(XLSXError("Relationships not found in xl/_rels/workbook.xml.rels"))

    nothing
end

# Updates xf.workbook from xf.data[\"xl/workbook.xml\"]
function parse_workbook!(xf::XLSXFile)
    xroot = xmlroot(xf,"xl/workbook.xml")[end]
    chn = XML.children(xroot)
    XML.tag(xroot) != "workbook" && throw(XLSXError("Malformed xl/workbook.xml. Root node name should be 'workbook'. Got '$(XML.tag(xroot))'."))

    # workbook to be parsed
    workbook = get_workbook(xf)

    # workbookPr -> date1904
    # does not have attribute => is not date1904
    workbook.date1904 = false

    # changes workbook.date1904 if there is a setting in the workbookPr node
    for node in chn
        if XML.tag(node) == "workbookPr"

            # read date1904 attribute
            attributes = XML.attributes(node)
            if !isnothing(attributes)
                if haskey(attributes, "date1904")
                    attribute_value_date1904 = attributes["date1904"]
                    if attribute_value_date1904 == "1" || attribute_value_date1904 == "true"
                        workbook.date1904 = true
                    elseif attribute_value_date1904 == "0" || attribute_value_date1904 == "false"
                        workbook.date1904 = false
                    else
                        throw(XLSXError("Could not parse xl/workbook -> workbookPr -> date1904 = $(attribute_value_date1904)."))
                    end
                end
            end

            break
        end
    end

    # sheets
    sheets = Vector{Worksheet}()
    for node in chn
        if XML.tag(node) == "sheets"

            for sheet_node in XML.children(node)
                XML.tag(sheet_node) != "sheet" && throw(XLSXError("Unsupported node $(XML.tag(sheet_node)) in node $(XML.tag(node)) in 'xl/workbook.xml'."))
                worksheet = Worksheet(xf, sheet_node)
                push!(sheets, worksheet)
            end
            break
        end
    end
    workbook.sheets = sheets

    # named ranges
    for node in chn
        if XML.tag(node) == "definedNames"

            for defined_name_node in XML.children(node)

                if XML.tag(defined_name_node) == "definedName"

                    defined_value_string = XML.value(defined_name_node[1])
                    name = XML.attributes(defined_name_node)["name"]

                    local defined_value::DefinedNameValueTypes
                    if is_valid_non_contiguous_range(defined_value_string)
                        el = split(defined_value_string, ',')
                        rng = String[]
                        for rf in el
                            sp = split(rf, '!')
                            push!(rng, (unquoteit(sp[1])*"!"*sp[2]))
                        end
                        defined_value = NonContiguousRange(join(rng, ','))
                        isabs = Vector{Bool}(undef, length(defined_value.rng))
                        for (i, d) in enumerate(split(defined_value_string, ","))
                            isabs[i] = is_valid_fixed_sheet_cellname(d) || is_valid_fixed_sheet_cellrange(d)
                        end
                        length(isabs) != length(defined_value.rng) && throw(XLSXError("Error parsing absolute references in non-contiguous range."))
                    elseif is_valid_fixed_sheet_cellname(defined_value_string)
                        sp = split(defined_value_string, '!')
                        defined_value = SheetCellRef(unquoteit(sp[1])*"!"*sp[2])
                        isabs = true
                    elseif is_valid_sheet_cellname(defined_value_string)
                        sp = split(defined_value_string, '!')
                        defined_value = SheetCellRef(unquoteit(sp[1])*"!"*sp[2])
                        isabs = false
                    elseif is_valid_fixed_sheet_cellrange(defined_value_string)
                        sp = split(defined_value_string, '!')
                        defined_value = SheetCellRange(unquoteit(sp[1])*"!"*sp[2])
                        isabs = true
                    elseif is_valid_sheet_cellrange(defined_value_string)
                        sp = split(defined_value_string, '!')
                        defined_value = SheetCellRange(unquoteit(sp[1])*"!"*sp[2])
                        isabs = false
                    elseif occursin(r"^\".*\"$", defined_value_string) # is enclosed by quotes
                        defined_value = defined_value_string[nextind(defined_value_string, begin):prevind(defined_value_string, end)] # remove enclosing quotes
                        if isempty(defined_value)
                            defined_value = missing
                        end
                        isabs = false
                    elseif tryparse(Int64, defined_value_string) !== nothing
                        defined_value = parse(Int64, defined_value_string)
                        isabs = false
                    elseif tryparse(Float64, defined_value_string) !== nothing
                        defined_value = parse(Float64, defined_value_string)
                        isabs = false
                    elseif isempty(defined_value_string)
                        defined_value = missing
                        isabs = false
                    else

                        # Couldn't parse definedName. Will silently ignore it, since this is not a critical feature.
                        # Actually is just interpreted as a string anyway and added to the defined names (is this true?).
                        defined_value = string(defined_value_string)
                        isabs = false
                        #continue

                        # debug - Now more important since we are writing updated defined names to back to output file.
                        # throw(XLSXError("Could not parse value $(defined_value_string) for definedName $name."))
                    end
                    a = XML.attributes(defined_name_node)
                    if haskey(a, "localSheetId")
                        # is a Worksheet level name

                        # localSheetId is the 0-based index of the Worksheet in the order
                        # that it is displayed on screen.
                        # Which is the order of the elements under <sheets> element in workbook.xml .
                        localSheetId = parse(Int, a["localSheetId"]) + 1
                        sheetId = workbook.sheets[localSheetId].sheetId
                        workbook.worksheet_names[(sheetId, name)] = DefinedNameValue(defined_value, isabs)
                    else
                        # is a Workbook level name
                        workbook.workbook_names[name] = DefinedNameValue(defined_value, isabs)
                    end
                end

            end
            break
        end
    end

    nothing
end

# Returns a Dict mapping Workbook <externalReferences>: index => relationship id.
function get_wb_ext_refs(xf::XLSXFile)
    ext_refs = Dict{Int, String}()
    xroot = xmlroot(xf, "xl/workbook.xml")
    i, j = get_idces(xroot, "workbook", "externalReferences")
    if !isnothing(j)
        for (i, ref) in enumerate(XML.children(xroot[i][j]))
            ext_refs[i] = ref["r:id"]
        end
    end
    return ext_refs
end

# delete Override PartName=calcChain since this was never loaded (#31)
function remove_calcChain!(xf::XLSXFile)
    xf.data["[Content_Types].xml"]
    ctype_root = xmlroot(xf, "[Content_Types].xml")[end]
    for (i, c) in enumerate(XML.children(ctype_root))
        if c.tag == "Override" && haskey(c.attributes, "PartName") && c.attributes["PartName"]=="/xl/calcChain.xml"
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

function skipNode(r::XML.Raw, skipnode::String) # separate rows or ssts to speed up reading of large files
#    new = Vector{UInt8}() # original data with <sheetData> or <sst> node removed
#    skipped = Vector{UInt8}() # just the <sheetData> or <sst> node and its children
    new = IOBuffer() # original data with <sheetData> or <sst> node removed
    skipped = IOBuffer() # just the <sheetData> or <sst> node and its children
    n = XML.next(r)
    write(new, n.data[n.pos:n.pos+n.len])

    while first(XML.get_name(n.data, n.pos)) != skipnode # Retain everything before the <sheetData> or <sst> node
        n = XML.next(n)
        write(new, n.data[n.pos:n.pos+n.len])
    end

    if skipnode == "sheetData" # Add parents for <row> or <sst> elements to the excerpted data
        write(skipped, "<worksheet>")
        write(skipped, "<sheetData>")
    elseif skipnode == "sst"
        write(skipped, "<sst>")
    else
        throw(XLSXError("Unknown skipnode $skipnode."))
    end
    sdepth = n.depth
    n = XML.next(n)
    while n !== nothing && n.depth > sdepth # Put all children of <sheetData> or <sst> into the excerpted data
        write(skipped, n.data[n.pos:n.pos+n.len])
        n = XML.next(n)
    end
    while n !== nothing # Retain everything after the <sheetData> or <sst> node
        write(new, n.data[n.pos:n.pos+n.len])
        n = XML.next(n)
    end
    if skipnode == "sheetData"  # close parents for <row> or <sst> elements in the excerpted data
        write(skipped, "</sheetData>")
        write(skipped, "</worksheet>")
    elseif skipnode == "sst"
        write(skipped, "</sst>")
    end
    return take!(new), take!(skipped)
end

function stream_files(xf::XLSXFile, zip_io::ZipArchives.ZipReader; pass::Int, channel_size::Int=1 << 8)
    Channel{String}(channel_size) do out
        for f in ZipArchives.zip_names(zip_io)

            # ignore xl/calcChain.xml in any case (#31)
            if f != "xl/calcChain.xml"

                if pass==1 && !startswith(f, "customXml") && (endswith(f, ".xml") || endswith(f, ".rels"))
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
function load_files!(xf::XLSXFile, zip_io::ZipArchives.ZipReader; pass::Int)

    (pass < 1 || pass > 3) && throw(XLSXError("Unknown pass to read files."))
    wb = get_workbook(xf)

    read_files = Channel{ReadFile}(1 << 8)
    all_files = stream_files(xf, zip_io; pass)
   
    # Filter files based on pass BEFORE parallel processing
    filtered_files = Channel{String}(1 << 8) do out
        for file in all_files
            should_process = if pass == 1
                !occursin(r"xl/worksheets/sheet\d+\.xml|xl/sharedStrings\.xml", file)
            elseif pass == 2
                occursin(r"xl/sharedStrings\.xml", file)
            else  # pass == 3
                occursin(r"xl/worksheets/sheet\d+\.xml", file)
            end
           
            if should_process
                put!(out, file)
            end
        end
    end

    consumer = @async begin
        for file in read_files
            if !isnothing(file.node)
                xf.data[file.name] = file.node
                xf.files[file.name] = true
            end
            if !isnothing(file.raw)
                if xf.is_writable || pass==2
                    if occursin("xl/sharedStrings.xml", file.name)
                        if has_sst(wb)
                            sst_load!(wb)
                        end
                    elseif xf.use_cache_for_sheet_data && !occursin("xl/sharedStrings.xml", file.name)
                        rid = get_relationship_id_by_target(wb, file.name)
                        for sheet in wb.sheets
                            if sheet.relationship_id == rid
                                first_cache_fill!(sheet, XML.LazyNode(file.raw), Threads.nthreads())
                            end
                        end
                    end
                end
            end
            if !isnothing(file.bin)
                xf.binary_data[file.name] = file.bin
            end
        end
    end

    # Now workers only process relevant files
    @sync for _ in 1:Threads.nthreads()
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

        node=nothing
        raw=nothing
        bin=nothing

        try
            bytes = ZipArchives.zip_readentry(zip_io, filename)
            if !startswith(filename, "customXml") && (endswith(filename, ".xml") || endswith(filename, ".rels"))
                if occursin(r"xl/worksheets/sheet\d+\.xml|xl/sharedStrings\.xml", filename)
                    strip_bom_and_lf!(bytes)
                    skipnode = filename == "xl/sharedStrings.xml" ? "sst" : "sheetData"
                    f, s = skipNode(XML.Raw(bytes), skipnode) # <row> and <sst> elements can be very numerous in large files, so split out and keep as Raw XML data for speed
                    node = XML.Node(XML.Raw(f))
                    raw = XML.Raw(s)
                else
                    strip_bom_and_lf!(bytes)
                    node = XML.Node(XML.Raw(bytes))
                end
            else
                bin = bytes                
            end
        catch err
            throw(XLSXError("Failed to parse internal XML file `$filename`"))
        end

        return ReadFile(node, raw, bin, filename)
end

function internal_xml_file_read(xf::XLSXFile, filename::String)
    !internal_xml_file_exists(xf, filename) && throw(XLSXError("Couldn't find $filename in $(xf.source)."))
    !internal_xml_file_isread(xf, filename) && throw(XLSXError("$filename in $(xf.source) has not been read."))
    return internal_xml_file_read(xf::XLSXFile, nothing, filename::String)
end

function internal_xml_file_read(xf::XLSXFile, zip_io::Union{Nothing,ZipArchives.ZipReader}, filename::String)

    !internal_xml_file_exists(xf, filename) && throw(XLSXError("Couldn't find $filename in $(xf.source)."))

    if !internal_xml_file_isread(xf, filename)

        try
            bytes = ZipArchives.zip_readentry(zip_io, filename)
            strip_bom_and_lf!(bytes)
            if occursin(r"xl/worksheets/sheet\d+\.xml|xl/sharedStrings\.xml", filename)
                skipnode = filename == "xl/sharedStrings.xml" ? "sst" : "sheetData"
                f, _ = skipNode(XML.Raw(bytes), skipnode) # <row> and <sst> elements can be very numerous in large files, so split out and keep as Raw XML data for speed
                xf.data[filename] = XML.Node(XML.Raw(f))
            else
                xf.data[filename] = XML.Node(XML.Raw(bytes))
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
string (e.g. "N/A") or a vector of strings (e.g. ["N/A", "NULL"]). 
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
into the worksheet cache on reading.
The default behavior is `enable_cache=false`.

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
    enable_cache::Bool=false, 
    keep_empty_rows::Bool=false, 
    normalizenames::Bool=false, 
    missing_strings::Union{AbstractString, AbstractVector{<:AbstractString}, Nothing}=nothing
)
    c = openxlsx(source; enable_cache) do xf
        gettable(getsheet(xf, 1); first_row, column_labels, header, infer_eltypes, stop_in_empty_row, stop_in_row_function, keep_empty_rows, normalizenames, missing_strings)
    end
    return c
end
function readtable(source::Union{AbstractString,IO}, sheet::Union{AbstractString,Int}; 
    first_row::Union{Nothing,Int}=nothing, 
    column_labels=nothing, 
    header::Bool=true, 
    infer_eltypes::Bool=true, 
    stop_in_empty_row::Bool=true, 
    stop_in_row_function::Union{Nothing,Function}=nothing, 
    enable_cache::Bool=false, 
    keep_empty_rows::Bool=false, 
    normalizenames::Bool=false,
    missing_strings::Union{AbstractString, AbstractVector{<:AbstractString}, Nothing}=nothing
)
    c = openxlsx(source; enable_cache) do xf
        gettable(getsheet(xf, sheet); first_row, column_labels, header, infer_eltypes, stop_in_empty_row, stop_in_row_function, keep_empty_rows, normalizenames, missing_strings)
    end
    return c
end

function readtable(source::Union{AbstractString,IO}, sheet::Union{AbstractString,Int}, columns::ColumnRange; 
    first_row::Union{Nothing,Int}=nothing, 
    column_labels=nothing, 
    header::Bool=true, 
    infer_eltypes::Bool=true, 
    stop_in_empty_row::Bool=true, 
    stop_in_row_function::Union{Nothing,Function}=nothing, 
    enable_cache::Bool=false, 
    keep_empty_rows::Bool=false, 
    normalizenames::Bool=false,
    missing_strings::Union{AbstractString, AbstractVector{<:AbstractString}, Nothing}=nothing
)
    c = openxlsx(source; enable_cache) do xf
        gettable(getsheet(xf, sheet), columns; first_row, column_labels, header, infer_eltypes, stop_in_empty_row, stop_in_row_function, keep_empty_rows, normalizenames, missing_strings)
    end
    return c
end

function readtable(source::Union{AbstractString,IO}, sheet::Union{AbstractString,Int}, range::AbstractString; 
    first_row::Union{Nothing,Int}=nothing, 
    column_labels=nothing, 
    header::Bool=true, 
    infer_eltypes::Bool=true, 
    stop_in_empty_row::Bool=true, 
    stop_in_row_function::Union{Nothing,Function}=nothing, 
    enable_cache::Bool=false, 
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
    xf = XLSX.readxlsx(filename)
    XLSX.hassheet(xf, sheetname) || throw(XLSX.XLSXError("Sheet $sheetname not found in file $filename"))
    return gettransposedtable(xf[sheetname], rows; first_column, column_labels, header, normalizenames)
end
function readtransposedtable(filename::AbstractString, sheetname::AbstractString; first_column=nothing, column_labels=nothing, header::Bool=true, normalizenames::Bool=false)
    xf = XLSX.readxlsx(filename)
    XLSX.hassheet(xf, sheetname) || throw(XLSX.XLSXError("Sheet $sheetname not found in file $filename"))
    dim=XLSX.get_dimension(xf[sheetname])
    return gettransposedtable(xf[sheetname], "$(dim.start.row_number):$(dim.stop.row_number)"; first_column, column_labels, header, normalizenames)
end
function readtransposedtable(filename::AbstractString; first_column=nothing, column_labels=nothing, header::Bool=true, normalizenames::Bool=false)
    xf = XLSX.readxlsx(filename)
    dim=XLSX.get_dimension(xf[1])
    return gettransposedtable(xf[1], "$(dim.start.row_number):$(dim.stop.row_number)"; first_column, column_labels, header, normalizenames)
end

const escape_chars = ['&' => "&amp;", '<' => "&lt;", '>' => "&gt;", '"' => "&quot;", '\'' => "&apos;"]
function escape(x::AbstractString)
    result = x
    for (char, entity) in escape_chars
        result = replace(result, char => entity)
    end
    return result
end
function unescape(x::AbstractString)
    result = x
    for (char, entity) in reverse(escape_chars)
        result = replace(result, entity => char)
    end
    return result
end