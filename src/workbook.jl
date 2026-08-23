const WORKBOOK_ORDER = String[
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

EmptyWorkbook() = Workbook(EmptyMSOfficePackage(), Vector{Worksheet}(), false,
    Vector{Relationship}(), Dict{SheetCellRef, AbstractFormula}(), SharedStringTable(), Dict{Int,Bool}(), Dict{Int,Bool}(),
    ReentrantLock(), ReentrantLock(), ReentrantLock(), Dict{String,DefinedNameValueTypes}(), Dict{Tuple{Int,String},DefinedNameValueTypes}(),
    nothing, Dict{Int, CellDataFormat}(), nothing, nothing, nothing, nothing, Dict{String, Vector{XML.Node}}(), nothing)
    
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

function _print_sheet_table(io::IO, wb::Workbook)
    @printf(io, "%21s %-13s %-13s\n", "sheetname", "size", "range")
    println(io, "-"^(21 + 1 + 13 + 1 + 13))

    for s in wb.sheets
#        sheetname = length(s.name) > 20 ? first(s.name, 20) * "…" : s.name
        sheetname = truncate_len(s.name, 20)
        if s.dimension !== nothing
            rg = s.dimension
            _size = size(rg) |> x -> string(x[1], "x", x[2])
            @printf(io, "%21s %-13s %-13s\n", sheetname, _size, rg)
        elseif is_chartsheet(wb, s.name)
            @printf(io, "%21s Chartsheet\n", sheetname)
        else
            @printf(io, "%21s size unknown\n", sheetname)
        end
    end
end

function Base.show(io::IO, xf::XLSXFile)

    function sheetcountstr(workbook)
        sc = sheetcount(workbook)
        if sc == 1
            return "1 Worksheet"
        else
            return "$sc Worksheets"
        end
    end
    function source(xf)
        xf.source isa IOBuffer && return "IOBuffer"
        return "\"$(xf.source)\""
    end

    wb = xf.workbook
    print(io, "XLSXFile($(source(xf))) ",
        "containing $(sheetcountstr(wb))\n")
    _print_sheet_table(io, wb)
end

function Base.show(io::IO, wb::Workbook)
    if !isdefined(wb, :package)
        print(io, "XLSX.Workbook(<uninitialised>)")
        return
    end
    xf = get_xlsxfile(wb)
    sc = sheetcount(wb)
    src = xf.source isa IOBuffer ? "IOBuffer" : "\"$(xf.source)\""
    print(io, "Workbook($src) containing ", sc, sc == 1 ? " Worksheet\n" : " Worksheets\n")
    _print_sheet_table(io, wb)
end

@inline Base.getindex(xl::XLSXFile, i::Integer) = getsheet(xl, i)

function Base.getindex(xl::XLSXFile, s::AbstractString)
    if hassheet(xl, s)
        return getsheet(xl, s)
    else
        return getdata(xl, s)
    end
end

function ref_chooser(f::Function, xl::XLSXFile, ref_str::AbstractString)
    if is_workbook_defined_name(xl, ref_str)
        v = get_defined_name_value(xl.workbook, ref_str)
        if is_defined_name_value_a_constant(v)
            return v
        elseif is_defined_name_value_a_reference(v)
            return f(xl, v)
        else
            throw(XLSXError("Unexpected Workbook defined name value: $v."))
        end
    elseif is_valid_sheet_cellname(ref_str)
        return f(xl, SheetCellRef(ref_str))
    elseif is_valid_sheet_cellrange(ref_str)
        return f(xl, SheetCellRange(ref_str))
    elseif is_valid_sheet_column_range(ref_str)
        return f(xl, SheetColumnRange(ref_str))
    elseif is_valid_sheet_row_range(ref_str)
        return f(xl, SheetRowRange(ref_str))
    elseif is_valid_non_contiguous_sheetcellrange(ref_str)
        return f(xl, NonContiguousRange(ref_str))
    end
    throw(XLSXError("`$ref_str` is not a valid SheetCellRef."))
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
    return ref_chooser(getdata, xl, s)
end

getcell(xl::XLSXFile, rng::SheetCellRange) = getcellrange(xl::XLSXFile, rng::SheetCellRange)
getcell(xl::XLSXFile, rng::SheetRowRange) = getcellrange(xl::XLSXFile, rng::SheetRowRange)
getcell(xl::XLSXFile, rng::SheetColumnRange) = getcellrange(xl::XLSXFile, rng::SheetColumnRange)
getcell(xl::XLSXFile, rng::NonContiguousRange) = getcellrange(xl::XLSXFile, rng::NonContiguousRange)

function getcell(xl::XLSXFile, ref::SheetCellRef)
    !hassheet(xl, ref.sheet) && throw(XLSXError("Sheet `$(ref.sheet)` not found."))
    return getcell(getsheet(xl, ref.sheet), ref.cellref)
end

function getcell(xl::XLSXFile, ref_str::AbstractString)
    return ref_chooser(getcell, xl, ref_str)
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
    return ref_chooser(getcellrange, xl, rng_str)
end

# the only constructor the accessors use — keeps value and absolute paired
DefinedName(name::AbstractString, scope::Union{Nothing,AbstractString}, dnv::DefinedNameValue) =
    DefinedName(String(name), isnothing(scope) ? nothing : String(scope), dnv.value, dnv.isabs)

# ... and back, for the write path and for make_absolute
DefinedNameValue(dn::DefinedName) = DefinedNameValue(dn.value, dn.absolute)

@inline is_workbook_scoped(dn::DefinedName)::Bool = isnothing(dn.scope)
Base.:(==)(a::DefinedName, b::DefinedName) =
    a.name == b.name && a.scope == b.scope && a.value == b.value && a.absolute == b.absolute
Base.hash(dn::DefinedName, h::UInt) = hash(dn.absolute, hash(dn.value, hash(dn.scope, hash(dn.name, h))))

# make_absolute already does exactly this for the writer; reusing it means the
# displayed form cannot drift from the written form.
@inline _show_value(dn::DefinedName) =
    is_defined_name_value_a_reference(dn.value) ? make_absolute(DefinedNameValue(dn)) : repr(dn.value)
@inline _show_scope(dn::DefinedName) =
    isnothing(dn.scope) ? "Workbook" : "Worksheet: \"$(dn.scope)\""

Base.show(io::IO, dn::DefinedName) = print(io, dn.name, " = ", _show_value(dn))

function Base.show(io::IO, ::MIME"text/plain", dn::DefinedName)
    println(io, "DefinedName: ", dn.name)
    println(io, "  scope: ", _show_scope(dn))
    print(io, "  value: ", _show_value(dn))
end


function Base.show(io::IO, ::MIME"text/plain", v::AbstractVector{DefinedName})
    if isempty(v)
        print(io, "0-element Vector{DefinedName}")
        return
    end
 
    # Field widths: names and values are quoted or bracketed by the row format,
    # so the +2 on the name column leaves room for the quotes.
    namewidth, valwidth, scopewidth = 24, 24, 32
 
    mixed = !allequal(dn.scope for dn in v)
 
    println(io, length(v), "-element Vector{DefinedName}:")
    if mixed
        @printf(io, " %-26s %-24s %-32s\n", "Name", "Value", "[Scope]")
        println(io, " ", "-"^26, " ", "-"^24, " ", "-"^32)
        for dn in v
            @printf(io, " %-26s %-24s %-32s\n",
                    "\"" * truncate_len(dn.name, namewidth) * "\"",
                    truncate_len(_show_value(dn), valwidth),
                    "[" * truncate_len(_show_scope(dn), scopewidth) * "]")
        end
    else
        @printf(io, " %-26s %-24s\n", "Name", "Value")
        println(io, " ", "-"^26, " ", "-"^24)
        for dn in v
            @printf(io, " %-26s %-24s\n",
                    "\"" * truncate_len(dn.name, namewidth) * "\"",
                    truncate_len(_show_value(dn), valwidth))
        end
    end
end
 

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

# Defined names are case-insensitive in Excel. Need to check on this basis
# (haskey is insufficient). These return the key *as stored*, which is what
# `delete!` needs; the `is_*` predicates are wrappers so the matching rule
# lives in one place.
@inline function find_workbook_defined_name(wb::Workbook, name::AbstractString)::Union{Nothing,String}
    uname = uppercase(name)
    for k in keys(wb.workbook_names)
        uppercase(k) == uname && return k
    end
    return nothing
end
@inline function find_worksheet_defined_name(wb::Workbook, sheetId::Int, name::AbstractString)::Union{Nothing,Tuple{Int,String}}
    uname = uppercase(name)
    for k in keys(wb.worksheet_names)
        first(k) == sheetId && uppercase(last(k)) == uname && return k
    end
    return nothing
end
@inline find_workbook_defined_name(xl::XLSXFile, name::AbstractString) = find_workbook_defined_name(get_workbook(xl), name)
@inline find_worksheet_defined_name(ws::Worksheet, name::AbstractString) = find_worksheet_defined_name(get_workbook(ws), ws.sheetId, name)

@inline is_workbook_defined_name(wb::Workbook, name::AbstractString)::Bool =
    !isnothing(find_workbook_defined_name(wb, name))
@inline is_worksheet_defined_name(wb::Workbook, sheetId::Int, name::AbstractString)::Bool =
    !isnothing(find_worksheet_defined_name(wb, sheetId, name))

# unchanged, keep as they are:
@inline is_workbook_defined_name(xl::XLSXFile, name::AbstractString)::Bool = is_workbook_defined_name(get_workbook(xl), name)
@inline is_worksheet_defined_name(ws::Worksheet, name::AbstractString)::Bool = is_worksheet_defined_name(get_workbook(ws), ws.sheetId, name)
@inline is_worksheet_defined_name(wb::Workbook, sheet_name::AbstractString, name::AbstractString)::Bool = is_worksheet_defined_name(wb, getsheet(wb, sheet_name).sheetId, name)

function get_defined_name_value(wb::Workbook, name::AbstractString)::DefinedNameValueTypes
    uname = uppercase(name)
    for (k, v) in wb.workbook_names
        uppercase(k) == uname && return v.value
    end
    throw(XLSXError("Workbook defined name `$name` not found."))
end

function get_defined_name_value(ws::Worksheet, name::AbstractString)::DefinedNameValueTypes
    wb = get_workbook(ws)
    sheetId = ws.sheetId
    uname = uppercase(name)
    for (k, v) in wb.worksheet_names
        first(k) == sheetId && uppercase(last(k)) == uname && return v.value
    end
    throw(XLSXError("Worksheet defined name `$name` not found on this sheet."))
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
# - wb.worksheet_names[(ordinal_sheet_number(wb, ws.name), name)] = DefinedNameValue(value, abs)
    wb.worksheet_names[(ws.sheetId, name)] = DefinedNameValue(value, abs)
end
addDefName(ws::Worksheet, name::AbstractString, value::Integer; absolute=true) = addDefName(ws, name, Int64(value); absolute)

function delDefName(xf::XLSXFile, names::Vector{String})
    wb = get_workbook(xf)
    targets = Vector{String}(undef, length(names))
    for (i, name) in enumerate(names)
        isempty(name) && throw(XLSXError("Defined name cannot be an empty string."))
        k = find_workbook_defined_name(wb, name)
        isnothing(k) && throw(XLSXError("Workbook has no defined name called `$name`."))
        j = findfirst(isequal(k), view(targets, 1:i-1))
        isnothing(j) || throw(XLSXError("Defined name `$name` given more than once (also as `$(names[j])`)."))
        targets[i] = k
    end
    for k in targets
        delete!(wb.workbook_names, k)
    end
    return nothing
end

function delDefName(ws::Worksheet, names::Vector{String})
    wb = get_workbook(ws)
    targets = Vector{Tuple{Int,String}}(undef, length(names))
    for (i, name) in enumerate(names)
        isempty(name) && throw(XLSXError("Defined name cannot be an empty string."))
        k = find_worksheet_defined_name(wb, ws.sheetId, name)
        isnothing(k) && throw(XLSXError("Worksheet `$(ws.name)` has no defined name called `$name`."))
        j = findfirst(isequal(k), view(targets, 1:i-1))
        isnothing(j) || throw(XLSXError("Defined name `$name` given more than once (also as `$(names[j])`)."))
        targets[i] = k
    end
    for k in targets
        delete!(wb.worksheet_names, k)
    end
    return nothing
end

 
"""
    deleteDefinedName(xf::XLSXFile,  name::AbstractString)
    deleteDefinedName(ws::Worksheet, name::AbstractString)
    deleteDefinedName(xf::XLSXFile,  names)
    deleteDefinedName(ws::Worksheet, names)
 
Delete one or more defined names from the scope named by the first argument:
workbook scope for an `XLSXFile`, the worksheet's own scope for a `Worksheet`.
The two scopes are independent — a worksheet-scoped name is invisible to the
`XLSXFile` method and vice versa, even when both scopes hold the same name.
 
`names` may be a vector of name strings, or a vector of the `DefinedName`
values returned by [`getDefinedNames`](@ref). A `DefinedName` is checked
against the first argument rather than routed by it, so
`deleteDefinedName(ws, getDefinedNames(ws))` deletes every name scoped to `ws`,
while `deleteDefinedName(ws, getDefinedNames(xf))` throws.
 
Names are matched case-insensitively, as in Excel, so `"my_name"` deletes a name
stored as `"My_Name"`.
 
Every name is validated before any is deleted: if one is not defined in the
target scope, or is given twice, an `XLSXError` is thrown and nothing is
removed. An empty collection is a no-op.
 
Deleting a defined name does not update formulas that referred to it; Excel
shows `#NAME?` for those, exactly as it does when a name is deleted through its
own name manager.
 
To clear a whole file at both scopes at once, see [`deleteAllDefinedNames`](@ref).
 
# Examples
```julia
julia> XLSX.deleteDefinedName(xf, "Life_the_universe_and_everything")
 
julia> XLSX.deleteDefinedName(sh, ["NEW", "my_name"])
 
julia> XLSX.deleteDefinedName(sh, XLSX.getDefinedNames(sh))   # clear this sheet
 
```
See also [`addDefinedName`](@ref), [`getDefinedNames`](@ref),
[`deleteAllDefinedNames`](@ref).
"""
function deleteDefinedName end
 
deleteDefinedName(xf::XLSXFile, name::AbstractString) = delDefName(xf, [String(name)])
deleteDefinedName(ws::Worksheet, name::AbstractString) = delDefName(ws, [String(name)])
 
deleteDefinedName(xf::XLSXFile, names::AbstractVector{<:AbstractString}) = delDefName(xf, String.(names))
deleteDefinedName(ws::Worksheet, names::AbstractVector{<:AbstractString}) = delDefName(ws, String.(names))
 
deleteDefinedName(xf::XLSXFile, dns::AbstractVector{DefinedName}) =
    delDefName(xf, [_checked_name(dn, nothing, "the workbook") for dn in dns])
deleteDefinedName(ws::Worksheet, dns::AbstractVector{DefinedName}) =
    delDefName(ws, [_checked_name(dn, ws.name, "worksheet \"$(ws.name)\"") for dn in dns])

# scope is the first argument; a DefinedName from elsewhere is refused, not routed
@inline function _checked_name(dn::DefinedName, expected::Union{Nothing,String}, where_str::String)
    dn.scope == expected || throw(XLSXError(
        "Defined name `$(dn.name)` is scoped to $(_show_scope(dn)), not to $where_str."))
    return dn.name
end

"""
    deleteAllDefinedNames(xf::XLSXFile)

Delete every defined name in `xf`, workbook-scoped and worksheet-scoped alike.
This and [`getAllDefinedNames`](@ref) are the only defined name operations that
act across scopes; everything else acts solely on the scope named by its first
argument.

See also [`deleteDefinedName`](@ref).
"""
function deleteAllDefinedNames(xf::XLSXFile)
    wb = get_workbook(xf)
    empty!(wb.workbook_names)
    empty!(wb.worksheet_names)
    return nothing
end

# The ref types are immutable, so rebuild rather than mutate. The inner CellRef
# / CellRange values carry no sheet name and are reused as they are.
@inline rename_sheet(v::SheetCellRef, new_name::String)      = SheetCellRef(new_name, v.cellref)
@inline rename_sheet(v::SheetCellRange, new_name::String)    = SheetCellRange(new_name, v.rng)
@inline rename_sheet(v::NonContiguousRange, new_name::String) = NonContiguousRange(new_name, v.rng)

"""
    update_defined_names_renamed_sheet!(wb, old_name, new_name)

Rewrite every defined name value that refers to `old_name` so that it refers to
`new_name` instead. Constant-valued defined names are left alone. Must be called
*before* `ws.name` is reassigned, while `old_name` is still current.
"""
function update_defined_names_renamed_sheet!(wb::Workbook, old_name::String, new_name::String)
    for k in collect(keys(wb.workbook_names))
        dn = wb.workbook_names[k]
        dn.value isa DefinedNameRangeTypes || continue
        dn.value.sheet == old_name || continue
        wb.workbook_names[k] = DefinedNameValue(rename_sheet(dn.value, new_name), dn.isabs)
    end
    for k in collect(keys(wb.worksheet_names))
        dn = wb.worksheet_names[k]
        dn.value isa DefinedNameRangeTypes || continue
        dn.value.sheet == old_name || continue
        wb.worksheet_names[k] = DefinedNameValue(rename_sheet(dn.value, new_name), dn.isabs)
    end
    return nothing
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
addDefinedName(xf::XLSXFile, name::AbstractString, value::DefinedNameRangeTypes;
               absolute::Union{Bool,Vector{Bool}}=true) = addDefName(xf, name, value; absolute)
addDefinedName(ws::Worksheet, name::AbstractString, value::DefinedNameRangeTypes;
               absolute::Union{Bool,Vector{Bool}}=true) = addDefName(ws, name, value; absolute)
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


"""
    XLSX.getDefinedNames(xf::XLSXFile)  -> Vector{DefinedName}
    XLSX.getDefinedNames(ws::Worksheet) -> Vector{DefinedName}
 
Return the defined names in a single scope, sorted by name.
 
Given an `XLSXFile`, the workbook-scoped names are returned. Given a
`Worksheet`, the names scoped to that worksheet are returned — not the
workbook-scoped names, even though those can be used from the sheet in a
formula. The scope is the argument, so the result carries no scope of its own
to choose from.
 
The result can be passed back to `deleteDefinedName`, which takes its scope
from the same kind of argument, so `deleteDefinedName(x, getDefinedNames(x))`
deletes every defined name in the scope of `x`.
 
For every name in a file at once, use [`XLSX.getAllDefinedNames`](@ref).
 
# Examples
```julia
julia> XLSX.getDefinedNames(xf)
2-element Vector{DefinedName}:
  CONST_INT    =  100
  SINGLE_CELL  =  named_ranges!\$A\$2
 
julia> XLSX.getDefinedNames(xf["named_ranges"])
1-element Vector{DefinedName}:
  LOCAL_REF  =  named_ranges!\$A\$15:\$B\$15
 
```
 
See also [`addDefinedName`](@ref), [`XLSX.getAllDefinedNames`](@ref),
[`XLSX.DefinedName`](@ref).
"""
function getDefinedNames end
getDefinedNames(xf::XLSXFile)::Vector{DefinedName} =
    sort!([DefinedName(k, nothing, v) for (k, v) in get_workbook(xf).workbook_names],
          by=dn -> uppercase(dn.name))
function getDefinedNames(ws::Worksheet)::Vector{DefinedName}
    wb = get_workbook(ws)
    return sort!([DefinedName(last(k), ws.name, v) for (k, v) in wb.worksheet_names if first(k) == ws.sheetId],
                 by=dn -> uppercase(dn.name))
end


"""
    XLSX.getAllDefinedNames(xf::XLSXFile) -> Vector{DefinedName}
 
Return every defined name in `xf`, at every scope, sorted by scope then name.
Each result carries the scope it came from: `nothing` for a workbook-scoped
name, or the name of the worksheet it is scoped to.
 
This is a view of the whole file for inspection. Deleting names still happens
one scope at a time — pass the result of [`XLSX.getDefinedNames`](@ref) to
`deleteDefinedName`, or use `deleteAllDefinedNames` to clear the file outright.
Passing these results to a scoped delete throws if any of them belongs to
another scope.
 
# Examples
```julia
julia> XLSX.getAllDefinedNames(xf)
3-element Vector{DefinedName}:
  CONST_INT    =  100                        [workbook]
  SINGLE_CELL  =  named_ranges!\$A\$2          [workbook]
  LOCAL_REF    =  named_ranges!\$A\$15:\$B\$15   [worksheet "named_ranges"]
 
```
 
See also [`XLSX.getDefinedNames`](@ref), [`XLSX.DefinedName`](@ref).
"""
function getAllDefinedNames(xf::XLSXFile)::Vector{DefinedName}
    wb = get_workbook(xf)
    sheet_lookup = Dict(ws.sheetId => ws.name for ws in wb.sheets)

    result = DefinedName[]
    for (name, v) in wb.workbook_names
        push!(result, DefinedName(name, nothing, v))
    end
    for ((sid, name), v) in wb.worksheet_names
        haskey(sheet_lookup, sid) ||
            throw(XLSXError("Defined name `$name` is scoped to sheetId $sid, which is not in the workbook."))
        push!(result, DefinedName(name, sheet_lookup[sid], v))
    end
    return sort!(result, by = dn -> (something(dn.scope, ""), uppercase(dn.name)))
end
