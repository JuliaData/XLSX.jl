#
# -- Get and set number format attributes
#

"""
    getFormat(sh::Worksheet, cr::String) -> ::Union{Nothing, CellFormat}
    getFormat(xf::XLSXFile,  cr::String) -> ::Union{Nothing, CellFormat}

    getFormat(sh::Worksheet, row::Int, col::int) -> ::Union{Nothing, CellFormat}
   
Get the format (numFmt) used by a single cell at reference `cr` in a worksheet or XLSXFile.
A cell outside the sheet dimension is empty, so this returns `nothing`, as for any empty cell.

Return a `CellFormat` object containing:
- `numFmtId`          : a 0-based index of the formats in the workbook. Values below 164 are 
                        reserved for built-in formats. Values of 164 and over are custom formats
                        and are stored in the `styles.xml` file within the XLSXFile.
- `format`            : a dictionary of numFmt attributes: formatAttribute -> (attribute -> value)
- `applyNumberFormat` : "1" or "0", indicating whether or not the format is applied to the cell.

Return `nothing` if no cell format is found. This will occur when a cell uses a built-in format.

The function will always find any explicitly set custom format. It will also attempt to return 
the format for built-in formats, too.

# Examples:
```julia
julia> getFormat(sh, "A1")

julia> getFormat(xf, "Sheet1!A1")

julia> getFormat(sh, 1, 1)
 
```
"""
function getFormat end
getFormat(xl::XLSXFile, sheetcell::String)::Union{Nothing,CellFormat} = process_get_sheetcell(getFormat, xl, sheetcell)
getFormat(ws::Worksheet, cellref::CellRef)::Union{Nothing,CellFormat} = process_get_cellref(getFormat, ws, cellref)
getFormat(ws::Worksheet, cr::String) = process_get_cellname(getFormat, ws, cr)
getFormat(ws::Worksheet, row::Integer, col::Integer) = getFormat(ws, CellRef(row, col))

function getFormat(wb::Workbook, cell_style::XML.Node)::Union{Nothing,CellFormat}
    _require_cache(wb, "get number formats")

    haskey(cell_style, "numFmtId") || return nothing

    numfmtid_int      = parse(Int, cell_style["numFmtId"])
    applynumberformat = haskey(cell_style, "applyNumberFormat") ? cell_style["applyNumberFormat"] : "0"
    format_atts       = Dict{String,Union{Dict{String,String},Nothing}}()

    if numfmtid_int >= PREDEFINED_NUMFMT_COUNT
        # Custom format — look up in the cached numFmtId => formatCode map
        cache = get_numFmt_cache(wb)
        haskey(cache, numfmtid_int) || throw(XLSXError("No format definition found for numFmtId $numfmtid_int."))

        format_atts["numFmt"] = Dict(
            "numFmtId"   => string(numfmtid_int),
            "formatCode" => cache[numfmtid_int]
        )

    else
        # Built-in format — validate it falls in a known range
        any(numfmtid_int ∈ r for r in BUILTIN_NUMFMT_RANGES) ||
            throw(XLSXError("Expected a built-in format ID in ranges 0:22, 37:40, or 45:49. Got $numfmtid_int."))

        if haskey(builtinFormats, string(numfmtid_int))
            format_atts["numFmt"] = Dict(
                "numFmtId"   => string(numfmtid_int),
                "formatCode" => builtinFormats[string(numfmtid_int)]
            )
        end
    end

    return CellFormat(numfmtid_int, format_atts, applynumberformat)
end

"""
    setFormat(sh::Worksheet, cr::String; kw...) -> ::Int
    setFormat(xf::XLSXFile,  cr::String; kw...) -> ::Int
    
    setFormat(sh::Worksheet, row, col; kw...) -> ::Int
   
Set the number format used used by a single cell, a cell range, a column 
range or row range or a named cell or named range in a worksheet or 
XLSXFile. Alternatively, specify the row and column using any combination 
of Integer, UnitRange, Vector{Integer} or `:`.

The function uses one keyword used to define a format:
- `format::String = nothing` : Defines a built-in or custom number format

The format keyword can define some built-in formats by name:
- `General`    : specifies internal format ID  0 (General)
- `Number`     : specifies internal format ID  2 (`0.00`)
- `Currency`   : specifies internal format ID  7 (`\$#,##0.00_);(\$#,##0.00)`)
- `Percentage` : specifies internal format ID  9 (`0%`)
- `ShortDate`  : specifies internal format ID 14 (`m/d/yyyy`)
- `LongDate`   : specifies internal format ID 15 (`d-mmm-yy`)
- `Time`       : specifies internal format ID 21 (`h:mm:ss`)
- `Scientific` : specifies internal format ID 48 (`##0.0E+0`)

If `Currency` is specified, Excel will use the appropriate local currency symbol.

Alternatively, the format keyword can define built-in formats by their internal 
format ID (as a string). The following built-in format IDs are supported:
* "0" => "General"
* "1" => "0"
* "2" => "0.00",
* "3" => "#,##0"
* "4" => "#,##0.00"
* "5" => "\$#,##0_);(\$#,##0)"
* "6" => "\$#,##0_);Red"
* "7" => "\$#,##0.00_);(\$#,##0.00)"
* "8" => "\$#,##0.00_);Red"
* "9" => "0%"
* "10" => "0.00%"
* "11" => "0.00E+00"
* "12" => "# ?/?"
* "13" => "# ??/??"
* "14" => "m/d/yyyy"
* "15" => "d-mmm-yy"
* "16" => "d-mmm"
* "17" => "mmm-yy"
* "18" => "h:mm AM/PM"
* "19" => "h:mm:ss AM/PM"
* "20" => "h:mm"
* "21" => "h:mm:ss"
* "22" => "m/d/yyyy h:mm"
* "37" => "#,##0_);(#,##0)"
* "38" => "#,##0_);Red"
* "39" => "#,##0.00_);(#,##0.00)"
* "40" => "#,##0.00_);Red"
* "45" => "mm:ss"
* "46" => "[h]:mm:ss"
* "47" => "mmss.0"
* "48" => "##0.0E+0"
* "49" => "@"

Lastly, `format` can be used to specify any custom format directly. 
Only weak checks are made of custom formats specified - they are otherwise added 
to the `XLSXFile` verbatim.

Formats may need characters that must be escaped when specified (see penultimate 
example, below).

# Examples:
```julia
julia> XLSX.setFormat(sh, "D2"; format = "h:mm AM/PM")

julia> XLSX.setFormat(xf, "Sheet1!A2"; format = "# ??/??")

julia> XLSX.setFormat(sh, "F1:F5"; format = "Currency")

julia> XLSX.setFormat(sh, "named_range"; format = "Percentage")

julia> XLSX.setFormat(sh, "A2"; format = "_-£* #,##0.00_-;-£* #,##0.00_-;_-£* \\\"-\\\"??_-;_-@_-")

julia> XLSX.setFormat(sh, "named_range"; format = "39") # specifies the built-in format "#,##0.00_);(#,##0.00)"
 
```
"""
function setFormat end
define_percell_spec_methods(:setFormat, :_setFormat, :FormatSpec)

# The `numFmtId` of `format`, as `setFormat` takes it: a built-in format's id, or the id
# of a custom format, added to the workbook if new.
_format_id(wb::Workbook, format::AbstractString)::Int =
    haskey(BUILTIN_FORMAT_CODES, format) ? BUILTIN_FORMAT_CODES[format] : get_new_formatId(wb, String(format))

function _setFormat(sh::Worksheet, cellref::CellRef; spec::FormatSpec)::Int
    (; format) = spec
    _require_writable(sh, "set number format")

    wb   = get_workbook(sh)
    cell, allXfNodes = _format_target(sh, cellref, "format")

    cell_style  = styles_cell_xf(allXfNodes, cell.style)
    cell_format = getFormat(wb, cell_style)

    # No format kwarg — return existing format id, or 0 if none set
    if isnothing(format)
        return isnothing(cell_format) ? 0 : cell_format.numFmtId
    end

    new_formatid = _format_id(wb, format)

    # numFmtId == 0 is the "General" default format — applyNumberFormat is redundant for it
    atts = new_formatid == 0 ?
        (["numFmtId"],                    [string(new_formatid)]) :
        (["numFmtId", "applyNumberFormat"], [string(new_formatid), "1"])

    newstyle   = update_template_xf(sh, allXfNodes, CellDataFormat(cell.style), atts[1], atts[2]).id
    cell.style = newstyle

    return new_formatid
end

"""
    setUniformFormat(sh::Worksheet, cr::String; kw...) -> ::Int
    setUniformFormat(xf::XLSXFile,  cr::String, kw...) -> ::Int

    setUniformFormat(sh::Worksheet, rows, cols; kw...) -> ::Int

Set the number format used by a cell range, a column range or row range or a 
named range in a worksheet or XLSXFile to be to be uniformly the same format.
Alternatively, specify the rows and columns using any combination of 
Integer, UnitRange, Vector{Integer} or `:`.

First, the number format of the first cell in the range (the top-left cell) is
updated according to the given `kw...` (using `setFormat()`). The resultant format is 
then applied to each remaining cell in the range.

As a result, every cell in the range will have a uniform number format.

This is functionally equivalent to applying `setFormat()` to each cell in the range 
but may be more efficient.

Applying `setUniformFormat()` without any keyword arguments simply copies the `Format` 
attributes from the first cell specified to all the others.

The value returned is the `numfmtId` of the format uniformly applied to the cells.
If all cells in the range are `EmptyCells`, the returned value is -1.

For keyword definitions see [`setFormat()`](@ref).

# Examples:
```julia
julia> XLSX.setUniformFormat(xf, "Sheet1!A2:L6"; format = "# ??/??")

julia> XLSX.setUniformFormat(sh, "F1:F5"; format = "Currency")

julia> setUniformFormat(sh, "B2,A5:D22")                   # Copy `Format` from B2 to cells in A5:D22
 
```
"""
function setUniformFormat end
define_uniform_spec_methods(:setUniformFormat, :_setUniformFormat, :_setFormat, :FormatSpec, ["numFmtId", "applyNumberFormat"])
