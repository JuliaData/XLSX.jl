#=
const Workbook_elements = String[
    "fileVersion",
    "fileSharing",
    "workbookPr",
    "workbookProtection",
    "bookViews",
    "sheets",
    "functionGroups",
    "externalReferences",
    "definedNames",
    "calcPr",
    "oleSize",
    "customWorkbookViews",
    "pivotCaches",
    "smartTagPr",
    "smartTagTypes",
    "webPublishing",
    "fileRecoveryPr",
    "webPublishObjects",
    "extLst"
    ]
=#
    
EmptyWorkbook() = Workbook(EmptyMSOfficePackage(), Vector{Worksheet}(), false,
    Vector{Relationship}(), Dict{SheetCellRef, AbstractFormula}(), SharedStringTable(), Dict{Int,Bool}(), Dict{Int,Bool}(),
    ReentrantLock(), Dict{String,DefinedNameValueTypes}(), Dict{Tuple{Int,String},DefinedNameValueTypes}(), nothing)

#=
Indicates whether this XLSX file can be edited.
This controls if assignment to worksheet cells is allowed.
Writable XLSXFile instances are opened with `XLSX.open_xlsx_template` method.
=#
is_writable(xl::XLSXFile) = xl.is_writable

"""
    sheetnames(xl::XLSXFile)
    sheetnames(wb::Workbook)

Return a vector with Worksheet names for this Workbook.
"""
sheetnames(wb::Workbook) = [s.name for s in wb.sheets]
@inline sheetnames(xl::XLSXFile) = sheetnames(xl.workbook)

"""
    hassheet(wb::Workbook, sheetname::AbstractString)
    hassheet(xl::XLSXFile, sheetname::AbstractString)

Return `true` if `wb` contains a sheet named `sheetname`.
"""
function hassheet(wb::Workbook, sheetname::AbstractString)::Bool
    for s in wb.sheets
        if s.name == unquoteit(sheetname)
            return true
        end
    end
    return false
end

@inline hassheet(xl::XLSXFile, sheetname::AbstractString) = hassheet(xl.workbook, sheetname)

"""
    sheetcount(xlsfile) :: Int

Count the number of sheets in the Workbook.
"""
@inline sheetcount(wb::Workbook) = length(wb.sheets)
@inline sheetcount(xl::XLSXFile) = sheetcount(xl.workbook)

# Returns true if workbook follows date1904 convention.
@inline isdate1904(wb::Workbook)::Bool = wb.date1904
@inline isdate1904(xf::XLSXFile)::Bool = isdate1904(get_workbook(xf))

function getsheet(wb::Workbook, sheetname::String)::Worksheet
    for ws in wb.sheets
        if ws.name == unquoteit(sheetname)
            return ws
        end
    end
    throw(XLSXError("$(get_xlsxfile(wb).source) does not have a Worksheet named `$sheetname`."))
end

@inline getsheet(wb::Workbook, sheet_index::Int)::Worksheet = wb.sheets[sheet_index]
@inline getsheet(xl::XLSXFile, sheetname::String)::Worksheet = getsheet(xl.workbook, sheetname)
@inline getsheet(xl::XLSXFile, sheet_index::Int)::Worksheet = getsheet(xl.workbook, sheet_index)

function Base.show(io::IO, xf::XLSXFile)

    function sheetcountstr(workbook)
        sc = sheetcount(workbook)
        if sc == 1
            return "1 Worksheet"
        else
            return "$sc Worksheets"
        end
    end

    wb = xf.workbook
    print(io, "XLSXFile(\"$(xf.source)\") ",
        "containing $(sheetcountstr(wb))\n")
    @printf(io, "%21s %-13s %-13s\n", "sheetname", "size", "range")
    println(io, "-"^(21 + 1 + 13 + 1 + 13))

    for s in wb.sheets
        sheetname = s.name
        if textwidth(sheetname) > 20
            sheetname = sheetname[collect(eachindex(s.name))[1:20]] * "…"
        end

        if s.dimension !== nothing
            rg = s.dimension
            _size = size(rg) |> x -> string(x[1], "x", x[2])
            @printf(io, "%21s %-13s %-13s\n", sheetname, _size, rg)
        else
            @printf(io, "%21s size unknown\n", sheetname)
        end
    end
end

@inline Base.getindex(xl::XLSXFile, i::Integer) = getsheet(xl, i)

function Base.getindex(xl::XLSXFile, s::AbstractString)
    if hassheet(xl, s)
        return getsheet(xl, s)
    else
        return getdata(xl, s)
    end
end

function getdata(xl::XLSXFile, ref::SheetCellRef)
    !hassheet(xl, ref.sheet) && throw(XLSXError("Sheet `$(ref.sheet)` not found."))
    return getdata(getsheet(xl, ref.sheet), ref.cellref)
end

function getdata(xl::XLSXFile, rng::SheetCellRange)
    !hassheet(xl, rng.sheet) && throw(XLSXError("Sheet `$(rng.sheet)` not found."))
    return getdata(getsheet(xl, rng.sheet), rng.rng)
end

function getdata(xl::XLSXFile, rng::SheetColumnRange)
    !hassheet(xl, rng.sheet) && throw(XLSXError("Sheet `$(rng.sheet)` not found."))
    return getdata(getsheet(xl, rng.sheet), rng.colrng)
end

function getdata(xl::XLSXFile, rng::SheetRowRange)
    !hassheet(xl, rng.sheet) && throw(XLSXError("Sheet `$(rng.sheet)` not found."))
    return getdata(getsheet(xl, rng.sheet), rng.rowrng)
end

function getdata(xl::XLSXFile, rng::NonContiguousRange)
    !hassheet(xl, rng.sheet) && throw(XLSXError("Sheet `$(rng.sheet)` not found."))
    return getdata(getsheet(xl, rng.sheet), rng)
end

function getdata(xl::XLSXFile, s::AbstractString)
    if is_workbook_defined_name(xl, s)
        v = get_defined_name_value(xl.workbook, s)
        if is_defined_name_value_a_constant(v)
            return v
        elseif is_defined_name_value_a_reference(v)
            return getdata(xl, v)
        else
            throw(XLSXError("Unexpected Workbook defined name value: $v."))
        end
    elseif is_valid_sheet_cellname(s)
        return getdata(xl, SheetCellRef(s))
    elseif is_valid_sheet_cellrange(s)
        return getdata(xl, SheetCellRange(s))
    elseif is_valid_sheet_column_range(s)
        return getdata(xl, SheetColumnRange(s))
    elseif is_valid_sheet_row_range(s)
        return getdata(xl, SheetRowRange(s))
    elseif is_valid_non_contiguous_sheetcellrange(s)
        return getdata(xl, NonContiguousRange(s))
    end

    throw(XLSXError("`$s` is not a valid sheetname, workbook definedName or cell/range reference."))
end

function getcell(xl::XLSXFile, ref::SheetCellRef)
    !hassheet(xl, ref.sheet) && throw(XLSXError("Sheet `$(ref.sheet)` not found."))
    return getcell(getsheet(xl, ref.sheet), ref.cellref)
end

function getcell(xl::XLSXFile, ref_str::AbstractString)
    if is_workbook_defined_name(xl, ref_str)
        v = get_defined_name_value(xl.workbook, ref_str)
        if is_defined_name_value_a_reference(v)
            return isa(v, SheetCellRef) ? getcell(xl, v) : getcellrange(xl, v)
        else
            throw(XLSXError("`$ref_str` is not a valid Workbook definedName reference."))
        end
    elseif is_valid_sheet_cellname(ref_str)
        return getcell(xl, SheetCellRef(ref_str))
    elseif is_valid_sheet_cellrange(ref_str)
        return getcellrange(xl, SheetCellRange(ref_str))
    elseif is_valid_sheet_column_range(ref_str)
        return getcellrange(xl, SheetColumnRange(ref_str))
    elseif is_valid_sheet_row_range(ref_str)
        return getcellrange(xl, SheetRowRange(ref_str))
    elseif is_valid_non_contiguous_range(ref_str)
        return getcellrange(xl, NonContiguousRange(ref_str))
    end
    throw(XLSXError("`$ref_str` is not a valid SheetCellRef."))
end

function getcellrange(xl::XLSXFile, rng::SheetCellRange)
    !hassheet(xl, rng.sheet) && throw(XLSXError("Sheet `$(rng.sheet)` not found."))
    return getcellrange(getsheet(xl, rng.sheet), rng.rng)
end

function getcellrange(xl::XLSXFile, rng::SheetColumnRange)
    !hassheet(xl, rng.sheet) && throw(XLSXError("Sheet `$(rng.sheet)` not found."))
    return getcellrange(getsheet(xl, rng.sheet), rng.colrng)
end

function getcellrange(xl::XLSXFile, rng::SheetRowRange)
    !hassheet(xl, rng.sheet) && throw(XLSXError("Sheet `$(rng.sheet)` not found."))
    return getcellrange(getsheet(xl, rng.sheet), rng.rowrng)
end

function getcellrange(xl::XLSXFile, rng::NonContiguousRange)
    !hassheet(xl, rng.sheet) && throw(XLSXError("Sheet `$(rng.sheet)` not found."))
    return getcellrange(getsheet(xl, rng.sheet), rng)
end

function getcellrange(xl::XLSXFile, rng_str::AbstractString)
    wb = get_workbook(xl)
    if is_workbook_defined_name(wb, rng_str)
        v = get_defined_name_value(wb, rng_str)
        if is_defined_name_value_a_reference(v)
            return getcellrange(xl, v)
        else
            throw(XLSXError("`$rng_str` is not a valid Workbook definedName reference."))
        end
    elseif is_valid_sheet_cellrange(rng_str)
        return getcellrange(xl, SheetCellRange(rng_str))
    elseif is_valid_sheet_column_range(rng_str)
        return getcellrange(xl, SheetColumnRange(rng_str))
    elseif is_valid_sheet_row_range(rng_str)
        return getcellrange(xl, SheetRowRange(rng_str))
    elseif is_valid_non_contiguous_range(rng_str)
        return getcellrange(xl, NonContiguousRange(rng_str))
    end
    throw(XLSXError("`$rng_str` is not a valid SheetCellRange."))
end

# Defined names are case-insensitive in Excel. Need to check on this basis (haskey is insufficient).
@inline is_workbook_defined_name(wb::Workbook, name::AbstractString)::Bool = !isnothing(findfirst(x -> uppercase(x)==uppercase(name), collect(keys(wb.workbook_names))))
@inline is_worksheet_defined_name(wb::Workbook, sheetId::Int, name::AbstractString)::Bool = !isnothing(findfirst(x -> uppercase(last(x))==uppercase(name) && first(x)==sheetId, collect(keys(wb.worksheet_names))))
@inline is_workbook_defined_name(xl::XLSXFile, name::AbstractString)::Bool = is_workbook_defined_name(get_workbook(xl), name)
@inline is_worksheet_defined_name(ws::Worksheet, name::AbstractString)::Bool = is_worksheet_defined_name(get_workbook(ws), ws.sheetId, name)
@inline is_worksheet_defined_name(wb::Workbook, sheet_name::AbstractString, name::AbstractString)::Bool = is_worksheet_defined_name(wb, getsheet(wb, sheet_name).sheetId, name)

@inline get_defined_name_value(wb::Workbook, name::AbstractString)::DefinedNameValueTypes = wb.workbook_names[name].value

function get_defined_name_value(ws::Worksheet, name::AbstractString)::DefinedNameValueTypes
    wb = get_workbook(ws)
    sheetId = ws.sheetId
    dn = wb.worksheet_names[(sheetId, name)]
    return dn.value
end

@inline is_defined_name_value_a_reference(v::DefinedNameValueTypes) = isa(v, SheetCellRef) || isa(v, SheetCellRange) || isa(v, NonContiguousRange)
@inline is_defined_name_value_a_reference(v::Integer) = is_defined_name_value_a_reference(Int64(v))
@inline is_defined_name_value_a_constant(v::DefinedNameValueTypes) = !is_defined_name_value_a_reference(v)
@inline is_defined_name_value_a_constant(v::Integer) = is_defined_name_value_a_constant(Int64(v))

function is_valid_defined_name(name::AbstractString)::Bool
    if isempty(name)
        return false
    end
    if is_valid_cellname(name) || is_valid_cellrange(name) || is_valid_non_contiguous_cellrange(name)
        return false
    end
    if is_valid_sheet_cellname(name) || is_valid_sheet_cellrange(name) || is_valid_non_contiguous_sheetcellrange(name)
        return false
    end
    if !isletter(name[1]) && name[1] != '_'
        return false
    end
    for c in name
        if !isletter(c) && !isdigit(c) && c != '_' && c != '\\'
            return false
        end
    end
    return true
end

function addDefName(xf::XLSXFile, name::AbstractString, value::DefinedNameValueTypes; absolute=true)
    if !is_valid_defined_name(name)
        throw(XLSXError("Invalid defined name: `$name`. May only contain letters, numbers, `_` or `\\` and must start with a letter or `_`."))
    end
    if is_workbook_defined_name(xf, name)
        throw(XLSXError("Workbook already has a defined name called `$name`."))
    end
    if value isa NonContiguousRange
        abs = absolute ? fill(true, length(value.rng)) : fill(false, length(value.rng))
    else
        abs = absolute ? true : false
    end
    xf.workbook.workbook_names[name] = DefinedNameValue(value, abs)
end
addDefName(xf::XLSXFile, name::AbstractString, value::Integer; absolute=true) = addDefName(xf, name, Int64(value); absolute)
function addDefName(ws::Worksheet, name::AbstractString, value::DefinedNameValueTypes; absolute=true)
    wb = get_workbook(ws)
    if !is_valid_defined_name(name)
        throw(XLSXError("Invalid defined name: `$name`. May only contain letters, numbers, `_` or `\\` and must start with a letter or `_`."))
    end
    if is_worksheet_defined_name(ws, name)
        throw(XLSXError("Worksheet `$(ws.name)` already has a defined name called `$name`."))
    end

    if value isa NonContiguousRange || value isa SheetCellRange
        value.sheet != ws.name && throw(XLSXError("Range $value is not in the given worksheet ($(ws.name))."))
    end
    if value isa NonContiguousRange
        abs = absolute ? fill(true, length(value.rng)) : fill(false, length(value.rng))
    else
        abs = absolute ? true : false
    end
    wb.worksheet_names[(ws.sheetId, name)] = DefinedNameValue(value, abs)
end
addDefName(ws::Worksheet, name::AbstractString, value::Integer; absolute=true) = addDefName(ws, name, Int64(value); absolute)

function quoteit(x::AbstractString)
    if occursin(r"[^\w]|\s", x)
        escaped = replace(x, "'" => "''")
        return "'$escaped'"
    else
        return x
    end
end

function unquoteit(x::AbstractString)
    if startswith(x, "'") && endswith(x, "'") && ncodeunits(x) >= 2
        inner = x[2:prevind(x, end)] # First character always ASCII - "'".
        return replace(inner, "''" => "'")
    else
        return x
    end
end

"""
    addDefinedName(xf::XLSXFile,  name::AbstractString, value::Union{Int, Float64, String}; absolute=true)
    addDefinedName(xf::XLSXFile,  name::AbstractString, value::AbstractString; absolute=true)
    addDefinedName(sh::Worksheet, name::AbstractString, value::Union{Int, Float64, String}; absolute=true)
    addDefinedName(sh::Worksheet, name::AbstractString, value::AbstractString; absolute=true)

Add a defined name to the Workbook or Worksheet. If an `XLSXFile` is passed, the defined name 
is added to the Workbook. If a `Worksheet` is passed, the defined name is added to the Worksheet.

When adding defined name referring to a cell or range to a workbook, `value` must include the sheet 
name (e.g. `Sheet1!A1:B2`). 

If the new `definedName` is a cell reference or range, by default, it will be an absolute 
reference (e.g. `\$A\$1:\$C\$6`). If `absolute=false` is specified, the new `definedName` will be 
a relative reference (e.g. `A1:C6`). Any `absolute` argument specified is ignored if the 
`definedName` is not a cell reference or range.

In the context of `XLSX.jl` there is no difference between an absolute reference and a relative 
reference. However, Excel treats them differently. When `definedNames` are read in as part of 
an XLSXFile, we keep track of whether they are absolute or not. If the XLSXFile is subsequently 
written out again, the status of the `definedNames` is preserved.

# Examples
```julia
julia> XLSX.addDefinedName(sh, "ID", "C21")

julia> XLSX.addDefinedName(sh, "NEW", "A1:B2")

julia> XLSX.addDefinedName(sh, "my_name", "A1,B2,C3")

julia> XLSX.addDefinedName(xf, "New", "'Mock-up'!A1:B2")

julia> XLSX.addDefinedName(xf, "Life_the_universe_and_everything", 42)

julia> XLSX.addDefinedName(xf, "first_name", "Hello World")

```
"""
function addDefinedName end
addDefinedName(xf::XLSXFile, name::AbstractString, value::Union{Integer,Float64}; absolute=true) = addDefName(xf, name, value isa Integer ? Int64(value) : value; absolute)
addDefinedName(ws::Worksheet, name::AbstractString, value::Union{Integer,Float64}; absolute=true) = addDefName(ws, name, value isa Integer ? Int64(value) : value; absolute)
function addDefinedName(xf::XLSXFile, name::AbstractString, value::AbstractString; absolute=true)
    if value == ""
        throw(XLSXError("Defined name value cannot be an empty string."))
    end
    if is_valid_cellname(value) || is_valid_cellrange(value) || is_valid_non_contiguous_cellrange(value)
        throw(XLSXError("Workbook defined name reference `$value` incomplete. Must contain sheet name (e.g. `Sheet1!A1:B2`)."))
    elseif is_valid_sheet_cellname(value)
        return addDefName(xf, name, SheetCellRef(value); absolute)
    elseif is_valid_sheet_cellrange(value)
        return addDefName(xf, name, SheetCellRange(value); absolute)
    elseif is_valid_non_contiguous_sheetcellrange(value)
        return addDefName(xf, name, NonContiguousRange(value); absolute)
    else
        return addDefName(xf, name, value; absolute)
    end
end
function addDefinedName(ws::Worksheet, name::AbstractString, value::AbstractString; absolute=true)
    if value == ""
        throw(XLSXError("Defined name value cannot be an empty string."))
    end
    if is_valid_cellname(value)
        return addDefName(ws, name, SheetCellRef(ws.name, CellRef(value)); absolute)
    elseif is_valid_sheet_cellname(value)
        return addDefName(ws, name, SheetCellRef(value); absolute)
    elseif is_valid_cellrange(value)
        return addDefName(ws, name, SheetCellRange(ws.name, CellRange(value)); absolute)
    elseif is_valid_sheet_cellrange(value)
        return addDefName(ws, name, SheetCellRange(value); absolute)
    elseif is_valid_non_contiguous_cellrange(value)
        return addDefName(ws, name, NonContiguousRange(ws, value); absolute)
    elseif is_valid_non_contiguous_sheetcellrange(value)
        return addDefName(ws, name, NonContiguousRange(value); absolute)
    else
        return addDefName(ws, name, value; absolute)
    end
end
