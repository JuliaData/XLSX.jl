# ==========================================================================================
#
# -- Get and set font attributes
#

"""
    setFont(sh::Worksheet, cr::String; kw...) -> ::Int
    setFont(xf::XLSXFile,  cr::String, kw...) -> ::Int

    setFont(sh::Worksheet, row, col; kw...) -> ::Int

Set the font used by a single cell, a cell range, a column range or 
row range or a named cell or named range in a worksheet or XLSXFile.
Alternatively, specify the row and column using any combination of 
Integer, UnitRange, Vector{Integer} or `:`.

Font attributes are specified using keyword arguments:
- `bold::Bool = nothing`        : set to `true` to make the font bold.
- `italic::Bool = nothing`      : set to `true` to make the font italic.
- `under::String = nothing`     : set to `single`, `double` or `none`.
- `strike::Bool = nothing`      : set to `true` to strike through the font.
- `size::Real = nothing`        : set the font size (1 to 409, e.g. `11` or `10.5`).
- `color = nothing`             : set the font color.
- `name::String = nothing`      : set the font name.
- `vertAlign::String = nothing` : set to `superscript`, `subscript` or `baseline`.

Only the attributes specified will be changed. If an attribute is not specified, the current
value will be retained. These are the only attributes supported currently.

No validation of the font names specified is performed. Available fonts will depend
on what your system has installed. If you specify, for example, `name = "badFont"`,
that value will be written to the XLSXFile.

As an expedient to get fonts to work, the `scheme` attribute is simply dropped from
new font definitions.

The `color` can be $COLOR_FORMS_DOC

Font attributes cannot be set for `EmptyCell`s. Set a cell value first.
If a cell range or column range includes any `EmptyCell`s, they will be
quietly skipped and the font will be set for the remaining cells.

For single cells, the value returned is the `fontId` of the font applied to the cell.
This can be used to apply the same font to other cells or ranges.

For cell ranges, column ranges and named ranges, the value returned is -1 unless 
the range resolves to a single cell, in which case that cell's id is returned.

# Examples:
```julia
julia> setFont(sh, "A1"; bold=true, italic=true, size=12, name="Arial")          # Single cell

julia> setFont(xf, "Sheet1!A1"; bold=false, size=14, color="yellow")             # Single cell

julia> setFont(sh, "A1:B7"; name="Aptos", under="double", strike=true)           # Cell range

julia> setFont(xf, "Sheet1!A1:B7"; size=24, name="Berlin Sans FB Demi")          # Cell range

julia> setFont(sh, "A:B"; italic=true, color="green", under="single")            # Column range

julia> setFont(xf, "Sheet1!A:B"; italic=true, color="red", under="single")       # Column range

julia> setFont(xf, "Sheet1!6:12"; italic=false, color="FF8888FF", under="none")  # Row range

julia> setFont(sh, "bigred"; size=48, color="FF00FF00")                          # Named cell or range

julia> setFont(xf, "bigred"; size=48, color="magenta")                           # Named cell or range

julia> setFont(sh, 1, 2; size=48, color="magenta")                               # row and column as integers

julia> setFont(sh, 1:3, 2; size=48, color="magenta")                             # row as unit range

julia> setFont(sh, 6, [2, 3, 8, 12]; size=48, color="magenta")                   # column as vector of indices

julia> setFont(sh, :, 2:6; size=48, color="lightskyblue2")                       # all rows, columns 2 to 6

```
"""
function setFont end
define_percell_spec_methods(:setFont, :_setFont, :FontSpec)

function _setFont(sh::Worksheet, cellref::CellRef; spec::FontSpec)::Int
    (; bold, italic, under, strike, size, color, name, vertAlign) = spec
    _require_writable(sh, "set font")

    # Validate before any work
    !isnothing(under) && under ∉ ("none", "single", "double") &&
        throw(XLSXError("Invalid value for under: $under. Must be one of: `none`, `single`, `double`."))
    !isnothing(size) && (size < 1 || size > 409) &&
        throw(XLSXError("Invalid size value: $size. Must be between 1 and 409."))

    pfx = _styles_prefix(sh)

    wb   = get_workbook(sh)
    cell, allXfNodes = _format_target(sh, cellref, "font")

    cell_style    = styles_cell_xf(allXfNodes, cell.style)
    cell_font     = getFont(wb, cell_style)
    old_font_atts = cell_font.font

    new_font_atts = Dict{String,Union{Dict{String,String},Nothing}}()

    # Boolean/flag tags
    for (tag, val) in (("b", bold), ("i", italic), ("strike", strike))
        result = _merge_flag_tag(tag, val, old_font_atts)
        ismissing(result) || (new_font_atts[tag] = result)
    end

    # Single-value tags
    for (tag, val) in (("sz", size), ("name", name), ("vertAlign", vertAlign))
        result = _merge_val_tag(tag, val, old_font_atts)
        ismissing(result) || (new_font_atts[tag] = result)
    end

    # Color tag
    result = _merge_dict_tag("color", color, old_font_atts, c -> Dict(_color_atts(c)))
    ismissing(result) || (new_font_atts["color"] = result)

    # Underline: "none" explicitly removes, "single" => no attributes, "double" => val attribute
    # This doesn't fit the standard merge helpers due to the three-way new_val distinction.
    if isnothing(under)
        haskey(old_font_atts, "u") && (new_font_atts["u"] = old_font_atts["u"])
    elseif under == "single"
        new_font_atts["u"] = nothing
    elseif under == "double"
        new_font_atts["u"] = Dict("val" => "double")
    end
    # under == "none": omit "u" entirely — already absent from new_font_atts

    # Drop "scheme" tag — Excel adds it but we don't want to propagate it
    font_node  = buildNode("font", new_font_atts, pfx)
    new_fontid = styles_add_cell_attribute(wb, font_node, "fonts")

    newstyle   = update_template_xf(sh, allXfNodes, CellDataFormat(cell.style), ["fontId", "applyFont"], [string(new_fontid), "1"]).id
    cell.style = newstyle

    if cell.datatype == CT_STRING
        v = update_sharedString_font(sh, cell, spec)
        cell.value = isnothing(v) ? cell.value : reinterpret(UInt64, Int64(v))
    end

    return new_fontid
end

"""
    setUniformFont(sh::Worksheet, cr::String; kw...) -> ::Int
    setUniformFont(xf::XLSXFile,  cr::String, kw...) -> ::Int

    setUniformFont(sh::Worksheet, rows, cols; kw...) -> ::Int

Set the font used by a cell range, a column range or row range or 
a named range in a worksheet or XLSXFile to be uniformly the same font.
Alternatively, specify the rows and columns using any combination of 
Integer, UnitRange, Vector{Integer} or `:`.

First, the font attributes of the first cell in the range (the top-left cell) are
updated according to the given `kw...` (using `setFont()`). The resultant font is 
then applied to each remaining cell in the range.

As a result, every cell in the range will have a uniform font setting.

This differs from `setFont()` which merges the attributes defined by `kw...` into 
the font definition used by each cell individually. For example, if you set the 
font size to 12 for a range of cells, but these cells all use different fonts names 
or colors, etc, `setFont()` will change the font size but leave the font name and 
color unchanged for each cell individually. 

In contrast, `setUniformFont()` will set the font size to 12 for the first cell, but 
will then apply all the font attributes from the updated first cell (ie. name, color, 
etc) to all the other cells in the range.

This can be more efficient when setting the same font for a large number of cells.

Applying `setUniformFont()` without any keyword arguments simply copies the `Font` 
attributes from the first cell specified to all the others.

The value returned is the `fontId` of the font uniformly applied to the cells.
If all cells in the range are `EmptyCells` the returned value is -1.

For keyword definitions see [`setFont()`](@ref).

# Examples:
```julia
julia> setUniformFont(sh, "A1:B7"; bold=true, italic=true, size=12, name="Arial")       # Cell range

julia> setUniformFont(xf, "Sheet1!A1:B7"; size=24, name="Berlin Sans FB Demi")          # Cell range

julia> setUniformFont(sh, "A:B"; italic=true, color="FF8888FF", under="single")         # Column range

julia> setUniformFont(xf, "Sheet1!A:B"; italic=true, color="FF8888FF", under="single")  # Column range

julia> setUniformFont(sh, "33"; italic=true, color="FF8888FF", under="single")          # Row

julia> setUniformFont(sh, "bigred"; size=48, color="FF00FF00")                          # Named range

julia> setUniformFont(sh, 1, [2, 4, 6]; size=48, color="lightskyblue2")                 # vector of column indices

julia> setUniformFont(sh, "B2,A5:D22")                                                  # Copy `Font` from B2 to cells in A5:D22
 
```
"""
function setUniformFont end
define_uniform_spec_methods(:setUniformFont, :_setUniformFont, :_setFont, :FontSpec, ["fontId", "applyFont"])


"""
    getFont(sh::Worksheet, cr::String) -> ::Union{Nothing, CellFont}
    getFont(xf::XLSXFile, cr::String)  -> ::Union{Nothing, CellFont}

    getFont(sh::Worksheet, row::Int, col::Int) -> ::Union{Nothing, CellFont}

Get the font used by a single cell reference in a worksheet `sh` or XLSXFile `xf`.
A cell outside the sheet dimension is empty, so this returns `nothing`, as for any empty cell.

Return a `CellFont` object containing:
- `fontId`    : a 0-based index of the font in the workbook
- `font`      : a dictionary of font attributes: fontAttribute -> (attribute -> value)
- `applyFont` : "1" or "0", indicating whether or not the font is applied to the cell.

Return `nothing` if no cell font is found.

Excel uses several tags to define font properties in its XML structure, and `font`
can hold any of them. `setFont` sets those marked *; the others come from files
written by Excel or other tools.
- `b`         : bold *
- `i`         : italic *
- `strike`    : strikethrough *
- `u`         : underline, e.g. single or double *
- `sz`        : font size *
- `color`     : font color *
- `name`      : font name *
- `vertAlign` : superscript or subscript *
- `condense`  : condensed font spacing
- `extend`    : extended font spacing
- `outline`   : outline text
- `shadow`    : shadowed text
- `family`    : font family
- `charset`   : character set code (e.g. 204 for Cyrillic)
- `scheme`    : whether the font is the theme's major or minor font

Excel defines colors in several ways. `getFont` will return the color in any of these
e.g. `"color" => ("theme" => "1")`.

# Examples:
```julia
julia> getFont(sh, "A1")

julia> getFont(xf, "Sheet1!A1")

julia> getFont(sh, "Sheet1!A1")

julia> getFont(sh, 1, 1)
 
```
"""
function getFont end
getFont(ws::Worksheet, cr::AbstractString) = process_get_cellname(getFont, ws, cr)
getFont(xl::XLSXFile, sheetcell::AbstractString)::Union{Nothing,CellFont} = process_get_sheetcell(getFont, xl, sheetcell)
getFont(ws::Worksheet, cellref::CellRef)::Union{Nothing,CellFont} = process_get_cellref(getFont, ws, cellref)
getFont(ws::Worksheet, row::Integer, col::Integer) = getFont(ws, CellRef(row, col))
getDefaultFont(ws::Worksheet) = getFont(get_workbook(ws), styles_cell_xf(get_workbook(ws), 0))
function getFont(wb::Workbook, cell_style::XML.Node)::Union{Nothing,CellFont}

    _require_cache(wb, "get font")

    if haskey(cell_style, "fontId")
        fontid = cell_style["fontId"]
        applyfont = haskey(cell_style, "applyFont") ? cell_style["applyFont"] : "0"
        font_nodes = get_fonts_nodes(wb)
        fontid_int = parse(Int, fontid)
        current_font = font_nodes[fontid_int + 1]
        font_atts = Dict{String,Union{Dict{String,String},Nothing}}()
        for c in xml_elements(current_font)
            if isnothing(XML.attributes(c)) || length(XML.attributes(c)) == 0
                font_atts[localname(c)] = nothing
            else
                font_atts[localname(c)] = Dict(XML.attributes(c))
            end
        end
        return CellFont(fontid_int, font_atts, applyfont)
    end

    return nothing
end
