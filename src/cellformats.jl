
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
- `size::Int = nothing`         : set the font size (0 < size < 410).
- `color::String = nothing`     : set the font color.
- `name::String = nothing`      : set the font name.
- `vertAlign::String = nothing` : set to `superscript`, `subscript` or `baseline`.

Only the attributes specified will be changed. If an attribute is not specified, the current
value will be retained. These are the only attributes supported currently.

No validation of the font names specified is performed. Available fonts will depend
on what your system has installed. If you specify, for example, `name = "badFont"`,
that value will be written to the XLSXFile.

As an expedient to get fonts to work, the `scheme` attribute is simply dropped from
new font definitions.

The `color` attribute can be defined using 8-digit rgb values.
- The first two digits represent transparency (α). Excel ignores transparency.
- The next two digits give the red component.
- The next two digits give the green component.
- The next two digits give the blue component.
So, FF000000 means a fully opaque black color.

Alternatively, you can use the name of any named color from [Colors.jl](https://juliagraphics.github.io/Colors.jl/stable/namedcolors/).

Font attributes cannot be set for `EmptyCell`s. Set a cell value first.
If a cell range or column range includes any `EmptyCell`s, they will be
quietly skipped and the font will be set for the remaining cells.

For single cells, the value returned is the `fontId` of the font applied to the cell.
This can be used to apply the same font to other cells or ranges.

For cell ranges, column ranges and named ranges, the value returned is -1.

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
setFont(ws::Worksheet, ref::SheetCellRef; kw...) = do_sheet_names_match(ws, ref) && setFont(ws, ref.cellref; kw...)
setFont(ws::Worksheet, rng::SheetCellRange; kw...) = do_sheet_names_match(ws, rng) && setFont(ws, rng.rng; kw...)
setFont(ws::Worksheet, rng::SheetColumnRange; kw...) = do_sheet_names_match(ws, rng) && setFont(ws, rng.colrng; kw...)
setFont(ws::Worksheet, rng::SheetRowRange; kw...) = do_sheet_names_match(ws, rng) && setFont(ws, rng.rowrng; kw...)
setFont(ws::Worksheet, rng::CellRange; kw...)::Int = process_cellranges(setFont, ws, rng; kw...)
setFont(ws::Worksheet, colrng::ColumnRange; kw...)::Int = process_columnranges(setFont, ws, colrng; kw...)
setFont(ws::Worksheet, rowrng::RowRange; kw...)::Int = process_rowranges(setFont, ws, rowrng; kw...)
setFont(ws::Worksheet, ncrng::NonContiguousRange; kw...)::Int = process_ncranges(setFont, ws, ncrng; kw...)
setFont(ws::Worksheet, ref_or_rng::AbstractString; kw...)::Int = process_ranges(setFont, ws, ref_or_rng; kw...)
setFont(xl::XLSXFile, sheetcell::String; kw...)::Int = process_sheetcell(setFont, xl, sheetcell; kw...)
#setFont(ws::Worksheet, row::Integer, col::Integer; kw...) = setFont(ws, CellRef(row, col); kw...)
setFont(ws::Worksheet, row::Union{Integer,UnitRange{<:Integer}}, ::Colon; kw...) = process_colon(setFont, ws, row, nothing; kw...)
setFont(ws::Worksheet, ::Colon, col::Union{Integer,UnitRange{<:Integer}}; kw...) = process_colon(setFont, ws, nothing, col; kw...)
setFont(ws::Worksheet, ::Colon, ::Colon; kw...) = process_colon(setFont, ws, nothing, nothing; kw...)
setFont(ws::Worksheet, ::Colon; kw...) = process_colon(setFont, ws, nothing, nothing; kw...)
setFont(ws::Worksheet, row::Union{Vector{Int},StepRange{<:Integer}}, ::Colon; kw...) = process_veccolon(setFont, ws, row, nothing; kw...)
setFont(ws::Worksheet, ::Colon, col::Union{Vector{Int},StepRange{<:Integer}}; kw...) = process_veccolon(setFont, ws, nothing, col; kw...)
setFont(ws::Worksheet, row::Union{Integer,UnitRange{<:Integer}}, col::Union{Vector{Int},StepRange{<:Integer}}; kw...) = process_vecint(setFont, ws, row, col; kw...)
setFont(ws::Worksheet, row::Union{Vector{Int},StepRange{<:Integer}}, col::Union{Integer,UnitRange{<:Integer}}; kw...) = process_vecint(setFont, ws, row, col; kw...)
setFont(ws::Worksheet, row::Union{Vector{Int},StepRange{<:Integer}}, col::Union{Vector{Int},StepRange{<:Integer}}; kw...) = process_vecint(setFont, ws, row, col; kw...)
setFont(ws::Worksheet, row::Union{Integer,UnitRange{<:Integer}}, col::Union{Integer,UnitRange{<:Integer}}; kw...) = setFont(ws, CellRange(CellRef(first(row), first(col)), CellRef(last(row), last(col))); kw...)

function setFont(sh::Worksheet, cellref::CellRef;
    bold::Union{Nothing,Bool}=nothing,
    italic::Union{Nothing,Bool}=nothing,
    under::Union{Nothing,String}=nothing,
    strike::Union{Nothing,Bool}=nothing,
    size::Union{Nothing,Int}=nothing,
    color::Union{Nothing,String}=nothing,
    name::Union{Nothing,String}=nothing,
    vertAlign::Union{Nothing,String}=nothing
)::Int
    get_xlsxfile(sh).is_writable || throw(XLSXError("Cannot set font because XLSXFile is not writable."))

    # Validate before any work
    !isnothing(under) && under ∉ ("none", "single", "double") &&
        throw(XLSXError("Invalid value for under: $under. Must be one of: `none`, `single`, `double`."))
    !isnothing(size) && (size < 1 || size > 409) &&
        throw(XLSXError("Invalid size value: $size. Must be between 1 and 409."))

    pfx = get_prefix("xl/styles.xml", get_xlsxfile(sh))
    pfx = pfx == "" ? "" : pfx * ":"

    wb   = get_workbook(sh)
    cell = getcell(sh, cellref)
    cell isa EmptyCell && throw(XLSXError("Cannot set font for an `EmptyCell`: $(cellname(cellref)). Set the value first."))

    allXfNodes = find_all_nodes(_xpath("styleSheet", "cellXfs", "xf"), styles_xmlroot(wb))

    if cell.style == UInt32(0)
        cell.style = get_num_style_index(sh, allXfNodes, 0).id
    end

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
    result = _merge_dict_tag("color", color, old_font_atts, c -> Dict("rgb" => get_color(c)))
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
        v = update_sharedString_font(sh, cell; bold, italic, under, strike, size, color, name)
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
setUniformFont(ws::Worksheet, rng::SheetCellRange; kw...) = do_sheet_names_match(ws, rng) && setUniformFont(ws, rng.rng; kw...)
setUniformFont(ws::Worksheet, rng::SheetColumnRange; kw...) = do_sheet_names_match(ws, rng) && setUniformFont(ws, rng.colrng; kw...)
setUniformFont(ws::Worksheet, rng::SheetRowRange; kw...) = do_sheet_names_match(ws, rng) && setUniformFont(ws, rng.rowrng; kw...)
setUniformFont(ws::Worksheet, colrng::ColumnRange; kw...)::Int = process_columnranges(setUniformFont, ws, colrng; kw...)
setUniformFont(ws::Worksheet, rowrng::RowRange; kw...)::Int = process_rowranges(setUniformFont, ws, rowrng; kw...)
setUniformFont(ws::Worksheet, ncrng::NonContiguousRange; kw...)::Int = process_uniform_ncranges(setFont, ws, ncrng, ["fontId", "applyFont"]; kw...)
setUniformFont(xl::XLSXFile, sheetcell::AbstractString; kw...)::Int = process_sheetcell(setUniformFont, xl, sheetcell; kw...)
setUniformFont(ws::Worksheet, ref_or_rng::AbstractString; kw...)::Int = process_ranges(setUniformFont, ws, ref_or_rng; kw...)
setUniformFont(ws::Worksheet, row::Union{Integer,UnitRange{<:Integer}}, ::Colon; kw...) = process_colon(setFont, ws, row, nothing; kw...)
setUniformFont(ws::Worksheet, ::Colon, col::Union{Integer,UnitRange{<:Integer}}; kw...) = process_colon(setFont, ws, nothing, col; kw...)
setUniformFont(ws::Worksheet, ::Colon, ::Colon; kw...) = process_colon(setUniformFont, ws, nothing, nothing; kw...)
setUniformFont(ws::Worksheet, ::Colon; kw...) = process_colon(setUniformFont, ws, nothing, nothing; kw...)
setUniformFont(ws::Worksheet, row::Union{Vector{Int},StepRange{<:Integer}}, ::Colon; kw...) = process_uniform_veccolon(setFont, ws, row, nothing, ["fontId", "applyFont"]; kw...)
setUniformFont(ws::Worksheet, ::Colon, col::Union{Vector{Int},StepRange{<:Integer}}; kw...) = process_uniform_veccolon(setFont, ws, nothing, col, ["fontId", "applyFont"]; kw...)
setUniformFont(ws::Worksheet, row::Union{Integer,UnitRange{<:Integer}}, col::Union{Vector{Int},StepRange{<:Integer}}; kw...) = process_uniform_vecint(setFont, ws, row, col, ["fontId", "applyFont"]; kw...)
setUniformFont(ws::Worksheet, row::Union{Vector{Int},StepRange{<:Integer}}, col::Union{Integer,UnitRange{<:Integer}}; kw...) = process_uniform_vecint(setFont, ws, row, col, ["fontId", "applyFont"]; kw...)
setUniformFont(ws::Worksheet, row::Union{Vector{Int},StepRange{<:Integer}}, col::Union{Vector{Int},StepRange{<:Integer}}; kw...) = process_uniform_vecint(setFont, ws, row, col, ["fontId", "applyFont"]; kw...)
setUniformFont(ws::Worksheet, row::Union{Integer,UnitRange{<:Integer}}, col::Union{Integer,UnitRange{<:Integer}}; kw...) = setUniformFont(ws, CellRange(CellRef(first(row), first(col)), CellRef(last(row), last(col))); kw...)
setUniformFont(ws::Worksheet, rng::CellRange; kw...)::Int = process_uniform_attribute(setFont, ws, rng, ["fontId", "applyFont"]; kw...)


"""
    getFont(sh::Worksheet, cr::String) -> ::Union{Nothing, CellFont}
    getFont(xf::XLSXFile, cr::String)  -> ::Union{Nothing, CellFont}

    getFont(sh::Worksheet, row::Int, col::Int) -> ::Union{Nothing, CellFont}

Get the font used by a single cell reference in a worksheet `sh` or XLSXFile `xf`.
The specified cell must be within the sheet dimension.

Return a `CellFont` object containing:
- `fontId`    : a 0-based index of the font in the workbook
- `font`      : a dictionary of font attributes: fontAttribute -> (attribute -> value)
- `applyFont` : "1" or "0", indicating whether or not the font is applied to the cell.

Return `nothing` if no cell font is found.

Excel uses several tags to define font properties in its XML structure.
Here's a list of some common tags and their purposes (thanks to Copilot!):
- `b`        : Indicates bold font.
- `i`        : Indicates italic font.
- `strike`   : Indicates strikethrough.
- `condense` : Condenses the font spacing.
- `extend`   : Extends the font spacing.
- `outline`  : Specifies outline text.
- `shadow`   : Adds a shadow to the text.
- `u`        : Specifies underlining (e.g., single, double).
- `sz`       : Sets the font size.
- `color`    : Sets the font color using RGB values).
- `name`     : Specifies the font name.
- `family`   : Defines the font family.
- `charset`  : Specifies the character set code (e.g., 204 for Cyrillic).
- `scheme`   : Specifies whether the font is part of the major or minor theme.

Excel defines colours in several ways. `getFont` will return the colour in any of these
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
getFont(ws::Worksheet, cr::String) = process_get_cellname(getFont, ws, cr)
getFont(xl::XLSXFile, sheetcell::String)::Union{Nothing,CellFont} = process_get_sheetcell(getFont, xl, sheetcell)
getFont(ws::Worksheet, cellref::CellRef)::Union{Nothing,CellFont} = process_get_cellref(getFont, ws, cellref)
getFont(ws::Worksheet, row::Integer, col::Integer) = getFont(ws, CellRef(row, col))
getDefaultFont(ws::Worksheet) = getFont(get_workbook(ws), styles_cell_xf(get_workbook(ws), 0))
function getFont(wb::Workbook, cell_style::XML.Node)::Union{Nothing,CellFont}

    if !get_xlsxfile(wb).use_cache_for_sheet_data
        throw(XLSXError("Cannot get font because cache is not enabled."))
    end

    if haskey(cell_style, "fontId")
        fontid = cell_style["fontId"]
        applyfont = haskey(cell_style, "applyFont") ? cell_style["applyFont"] : "0"
        xroot = styles_xmlroot(wb)
        font_elements = find_all_nodes("/" * SPREADSHEET_NAMESPACE_XPATH_ARG * ":styleSheet/" * SPREADSHEET_NAMESPACE_XPATH_ARG * ":fonts", xroot)[begin]
        font_nodes = filter(n -> XML.nodetype(n) == XML.Element, XML.children(font_elements))
        if parse(Int, font_elements["count"]) != length(font_nodes)
            throw(XLSXError("Unexpected number of font definitions found: $(length(font_nodes)). Expected $(parse(Int, font_elements["count"]))"))
        end
        fontid_int = parse(Int, fontid)
        current_font = font_nodes[fontid_int + 1]  # index into filtered list, not raw children
#        current_font = XML.children(font_elements)[fontid_int+1] # Zero based!
        font_atts = Dict{String,Union{Dict{String,String},Nothing}}()
        for c in XML.children(current_font)
            if isnothing(XML.attributes(c)) || length(XML.attributes(c)) == 0
                font_atts[localname(c)] = nothing
            else
#                for (k, v) in XML.attributes(c)
                    font_atts[localname(c)] = Dict(XML.attributes(c))
#                end
            end
        end
        return CellFont(fontid_int, font_atts, applyfont)
    end

    return nothing
end

#
# -- Get and set border attributes
#

"""
    getBorder(sh::Worksheet, cr::String) -> ::Union{Nothing, CellBorder}
    getBorder(xf::XLSXFile, cr::String)  -> ::Union{Nothing, CellBorder}

    getBorder(sh::Worksheet, row::Int, col::Int) -> ::Union{Nothing, CellBorder}
   
Get the borders used by a single cell at reference in a worksheet or XLSXFile.
The specified cell must be within the sheet dimension.

Return a `CellBorder` object containing:
- `borderId`    : a 0-based index of the border in the workbook
- `border`      : a dictionary of border attributes: borderAttribute -> (attribute -> value)
- `applyBorder` : "1" or "0", indicating whether or not the border is applied to the cell.

Return `nothing` if no cell border is found.

A cell border has two attributes, `style` and `color`. A diagonal border also needs to specify 
a direction indicating whether a diagonal is needed from bottom-left to top-right ("up"), 
from top-left to bottom-left ("down") or "both"

Here's a list of the available `style`s (thanks to Copilot!):
- `none`
- `thin`
- `medium`
- `dashed`
- `dotted`
- `thick`
- `double`
- `hair`
- `mediumDashed`
- `dashDot`
- `mediumDashDot`
- `dashDotDot`
- `mediumDashDotDot`
- `slantDashDot`

A border postion element (e.g. `top` or `left`) has a `style` attribute, but `color` is a child element.
The color element has one or two attributes (e.g. `rgb`) that define the color of the border.
While the key for the `style` element will always be `style`, the keys for the `color` element
will vary depending on how the color is defined (e.g. `rgb`, `indexed`, `auto`, etc.).
Thus, for example, `"top" => Dict("style" => "thin", "rgb" => "FF000000")` would indicate a 
thin black border at the top of the cell while `"top" => Dict("style" => "thin", "auto" => "1")`
would indicate that the color is set automatically by Excel.

The `color` element can have the following attributes:
- `auto`     : Indicates that the color is automatically defined by Excel
- `indexed`  : Specifies the color using an indexed color value.
- `rgb`      : Specifies the rgb color using 8-digit hexadecimal format.
- `theme`    : Specifies the color using a theme color.
- `tint`     : Specifies the tint value to adjust the lightness or darkness of the color.

`tint` can only be used in conjunction with the theme attribute to derive different shades of the theme color.
For example: <color theme="1" tint="-0.5"/>.

# Examples:
```julia
julia> getBorder(sh, "A1")

julia> getBorder(sh, 3, 6)

julia> getBorder(xf, "Sheet1!A1")
 
```
"""
function getBorder end
getBorder(xl::XLSXFile, sheetcell::String)::Union{Nothing,CellBorder} = process_get_sheetcell(getBorder, xl, sheetcell)
getBorder(ws::Worksheet, cellref::CellRef)::Union{Nothing,CellBorder} = process_get_cellref(getBorder, ws, cellref)
getBorder(ws::Worksheet, cr::String) = process_get_cellname(getBorder, ws, cr)
getBorder(ws::Worksheet, row::Integer, col::Integer) = getBorder(ws, CellRef(row, col))
getDefaultBorders(ws::Worksheet) = getBorder(get_workbook(ws), styles_cell_xf(get_workbook(ws), 0))
function getBorder(wb::Workbook, cell_style::XML.Node)::Union{Nothing,CellBorder}

    if !get_xlsxfile(wb).use_cache_for_sheet_data
        throw(XLSXError("Cannot get border because cache is not enabled."))
    end

    if haskey(cell_style, "borderId")
        borderid = cell_style["borderId"]
        applyborder = haskey(cell_style, "applyBorder") ? cell_style["applyBorder"] : "0"
        xroot = styles_xmlroot(wb)
        border_elements = find_all_nodes("/" * SPREADSHEET_NAMESPACE_XPATH_ARG * ":styleSheet/" * SPREADSHEET_NAMESPACE_XPATH_ARG * ":borders", xroot)[begin]
        if parse(Int, border_elements["count"]) != length(XML.children(border_elements))
            throw(XLSXError("Unexpected number of border definitions found : $(length(XML.children(border_elements))). Expected $(parse(Int, border_elements["count"]))"))
        end
        current_border = XML.children(border_elements)[parse(Int, borderid)+1] # Zero based!
        diag_atts = XML.attributes(current_border)
        border_atts = Dict{String,Union{Dict{String,String},Nothing}}()
        for side in XML.children(current_border)
            if isnothing(XML.attributes(side)) || length(XML.attributes(side)) == 0
                border_atts[localname(side)] = nothing
            else
                if length(XML.attributes(side)) != 1
                    throw(XLSXError("Too many border attributes found for $(localname(side)) Expected 1, found $(length(XML.attributes(side)))."))
                end
                for (k, v) in XML.attributes(side) # style is the only possible attribute of a side
                    border_atts[localname(side)] = Dict(k => v)
                    if localname(side) == "diagonal" && !isnothing(diag_atts)
                        if haskey(diag_atts, "diagonalUp") && haskey(diag_atts, "diagonalDown")
                            border_atts[localname(side)]["direction"] = "both"
                        elseif haskey(diag_atts, "diagonalUp")
                            border_atts[localname(side)]["direction"] = "up"
                        elseif haskey(diag_atts, "diagonalDown")
                            border_atts[localname(side)]["direction"] = "down"
                        else
                            throw(XLSXError("No direction set for `diagonal` border"))
                        end
                    end
                    for subc in XML.children(side) # color is a child of a border element
                        for (k, v) in XML.attributes(subc)
                            border_atts[localname(side)][k] = v
                        end
                    end
                end
            end
        end
        return CellBorder(parse(Int, borderid), border_atts, applyborder)
    end

    return nothing
end

"""
    setBorder(sh::Worksheet, cr::String; kw...) -> ::Int
    setBorder(xf::XLSXFile, cr::String; kw...) -> ::Int

    setBorder(sh::Worksheet, row, col; kw...) -> ::Int
   
Set the borders used used by a single cell, a cell range, a column range or 
row range or a named cell or named range in a worksheet or XLSXFile.
Alternatively, specify the row and column using any combination of 
Integer, UnitRange, Vector{Integer} or `:`.

Borders are independently defined for the keywords:
- `left::Vector{Pair{String,String}} = nothing`
- `right::Vector{Pair{String,String}} = nothing`
- `top::Vector{Pair{String,String}} = nothing`
- `bottom::Vector{Pair{String,String}} = nothing`
- `diagonal::Vector{Pair{String,String}} = nothing`
- `[allsides::Vector{Pair{String,String}} = nothing]`
- `[outside::Vector{Pair{String,String}} = nothing]`

These represent each of the sides of a cell . The keyword `diagonal` defines 
diagonal lines running across the cell. These lines must share the same style 
and color in any cell.

An additional keyword, `allsides`, is provided for convenience. It can be used 
in place of the four side keywords to apply the same border setting to all four 
sides at once. It cannot be used in conjunction with any of the side-specific 
keywords or with `outside` but it can be used together with `diagonal`.

A further keyword, `outside`, can be used to set the outside border around a 
range. Any internal borders will remain unchanged. An outside border cannot be 
set for any non-contiguous/non-rectangular range, cannot be indexed with 
vectors and cannot be used in conjunction with any other keywords.

The two attributes that can be set for each keyword are `style` and `color`.
Additionally, for diagonal borders, a third keyword, `direction` can be used.

Allowed values for `style` are:
- `none`
- `thin`
- `medium`
- `dashed`
- `dotted`
- `thick`
- `double`
- `hair`
- `mediumDashed`
- `dashDot`
- `mediumDashDot`
- `dashDotDot`
- `mediumDashDotDot`
- `slantDashDot`

The `color` attribute can be set by specifying an 8-digit hexadecimal value 
in the format "FFRRGGBB". The transparency ("FF") is ignored by Excel but 
is required.
Alternatively, you can use the name of any named color from [Colors.jl](https://juliagraphics.github.io/Colors.jl/stable/namedcolors/).

Valid values for the `direction` keyword (for diagonal borders) are:
- `up`   : diagonal border runs bottom-left to top-right
- `down` : diagonal border runs top-left to bottom-right
- `both` : diagonal borders run both ways

Both diagonal borders share the same style and color.

Setting only one of the attributes leaves the other attributes unchanged for that 
side's border. Omitting one of the keywords leaves the border definition for that
side unchanged, only updating the other, specified sides.

Border attributes cannot be set for `EmptyCell`s. Set a cell value first.
If a cell range or column range includes any `EmptyCell`s, they will be
quietly skipped and the border will be set for the remaining cells.

For single cells, the value returned is the `borderId` of the borders applied to the cell.
This can be used to apply the same borders to other cells or ranges.

For cell ranges, column ranges and named ranges, the value returned is -1.

# Examples:
```julia
Julia> setBorder(sh, "D6"; allsides = ["style" => "thick"], diagonal = ["style" => "hair", "direction" => "up"])

Julia> setBorder(sh, 2:45, 2:12; outside = ["style" => "thick", "color" => "lightskyblue2"])

Julia> setBorder(xf, "Sheet1!D4"; left     = ["style" => "dotted", "color" => "FF000FF0"],
                                  right    = ["style" => "medium", "color" => "firebrick2"],
                                  top      = ["style" => "thick",  "color" => "FF230000"],
                                  bottom   = ["style" => "medium", "color" => "goldenrod3"],
                                  diagonal = ["style" => "dotted", "color" => "FF00D4D4", "direction" => "both"]
                                  )
 
```
"""
function setBorder end
setBorder(ws::Worksheet, ref::SheetCellRef; kw...) = do_sheet_names_match(ws, ref) && setBorder(ws, ref.cellref; kw...)
setBorder(ws::Worksheet, rng::SheetCellRange; kw...) = do_sheet_names_match(ws, rng) && setBorder(ws, rng.rng; kw...)
setBorder(ws::Worksheet, rng::SheetColumnRange; kw...) = do_sheet_names_match(ws, rng) && setBorder(ws, rng.colrng; kw...)
setBorder(ws::Worksheet, rng::SheetRowRange; kw...) = do_sheet_names_match(ws, rng) && setBorder(ws, rng.rowrng; kw...)
setBorder(ws::Worksheet, ncrng::NonContiguousRange; kw...)::Int = process_ncranges(setBorder, ws, ncrng; kw...)
setBorder(ws::Worksheet, ref_or_rng::AbstractString; kw...)::Int = process_ranges(setBorder, ws, ref_or_rng; kw...)
#setBorder(ws::Worksheet, row::Integer, col::Integer; kw...) = setBorder(ws, CellRef(row, col); kw...)
setBorder(ws::Worksheet, row::Union{Integer,UnitRange{<:Integer}}, ::Colon; kw...) = process_colon(setBorder, ws, row, nothing; kw...)
setBorder(ws::Worksheet, ::Colon, col::Union{Integer,UnitRange{<:Integer}}; kw...) = process_colon(setBorder, ws, nothing, col; kw...)
setBorder(ws::Worksheet, ::Colon, ::Colon; kw...) = process_colon(setBorder, ws, nothing, nothing; kw...)
setBorder(ws::Worksheet, ::Colon; kw...) = process_colon(setBorder, ws, nothing, nothing; kw...)
setBorder(ws::Worksheet, row::Union{Vector{Int},StepRange{<:Integer}}, ::Colon; kw...) = process_veccolon(setBorder, ws, row, nothing; kw...)
setBorder(ws::Worksheet, ::Colon, col::Union{Vector{Int},StepRange{<:Integer}}; kw...) = process_veccolon(setBorder, ws, nothing, col; kw...)
setBorder(ws::Worksheet, row::Union{Integer,UnitRange{<:Integer}}, col::Union{Vector{Int},StepRange{<:Integer}}; kw...) = process_vecint(setBorder, ws, row, col; kw...)
setBorder(ws::Worksheet, row::Union{Vector{Int},StepRange{<:Integer}}, col::Union{Integer,UnitRange{<:Integer}}; kw...) = process_vecint(setBorder, ws, row, col; kw...)
setBorder(ws::Worksheet, row::Union{Vector{Int},StepRange{<:Integer}}, col::Union{Vector{Int},StepRange{<:Integer}}; kw...) = process_vecint(setBorder, ws, row, col; kw...)
setBorder(ws::Worksheet, row::Union{Integer,UnitRange{<:Integer}}, col::Union{Integer,UnitRange{<:Integer}}; kw...) = setBorder(ws, CellRange(CellRef(first(row), first(col)), CellRef(last(row), last(col))); kw...)

function setBorder(ws::Worksheet, rng::CellRange;
    outside::BorderKw=nothing, allsides::BorderKw=nothing,
    left::BorderKw=nothing, right::BorderKw=nothing,
    top::BorderKw=nothing, bottom::BorderKw=nothing,
    diagonal::BorderKw=nothing
)::Int
    _check_outside_conflict(outside, left, right, top, bottom, diagonal, allsides)
    isnothing(outside) || return setOutsideBorder(ws, rng; outside)
    return process_cellranges(setBorder, ws, rng; allsides, left, right, top, bottom, diagonal)
end

function setBorder(ws::Worksheet, rng::ColumnRange;
    outside::BorderKw=nothing, allsides::BorderKw=nothing,
    left::BorderKw=nothing, right::BorderKw=nothing,
    top::BorderKw=nothing, bottom::BorderKw=nothing,
    diagonal::BorderKw=nothing
)::Int
    _check_outside_conflict(outside, left, right, top, bottom, diagonal, allsides)
    isnothing(outside) || return process_columnranges(setOutsideBorder, ws, rng; outside)
    return process_columnranges(setBorder, ws, rng; allsides, left, right, top, bottom, diagonal)
end

function setBorder(ws::Worksheet, rng::RowRange;
    outside::BorderKw=nothing, allsides::BorderKw=nothing,
    left::BorderKw=nothing, right::BorderKw=nothing,
    top::BorderKw=nothing, bottom::BorderKw=nothing,
    diagonal::BorderKw=nothing
)::Int
    _check_outside_conflict(outside, left, right, top, bottom, diagonal, allsides)
    isnothing(outside) || return process_rowranges(setOutsideBorder, ws, rng; outside)
    return process_rowranges(setBorder, ws, rng; allsides, left, right, top, bottom, diagonal)
end

function setBorder(xl::XLSXFile, sheetcell::String;
    outside::BorderKw=nothing, allsides::BorderKw=nothing,
    left::BorderKw=nothing, right::BorderKw=nothing,
    top::BorderKw=nothing, bottom::BorderKw=nothing,
    diagonal::BorderKw=nothing
)::Int
    _check_outside_conflict(outside, left, right, top, bottom, diagonal, allsides)
    if isnothing(outside)
        return process_sheetcell(setBorder, xl, sheetcell; allsides, left, right, top, bottom, diagonal)
    else
        return process_sheetcell(setOutsideBorder, xl, sheetcell; outside)
    end
end

function setBorder(sh::Worksheet, cellref::CellRef;
    allsides::BorderKw=nothing,
    left::BorderKw=nothing, right::BorderKw=nothing,
    top::BorderKw=nothing, bottom::BorderKw=nothing,
    diagonal::BorderKw=nothing
)::Int
    get_xlsxfile(sh).is_writable ||
        throw(XLSXError("Cannot set borders because XLSXFile is not writable."))

    # allsides expands to all four sides, then recurses — early return
    if !isnothing(allsides)
        !all(isnothing, [left, right, top, bottom]) &&
            throw(XLSXError("Keyword `allsides` is incompatible with `left`, `right`, `top`, or `bottom`."))
        return setBorder(sh, cellref; left=allsides, right=allsides, top=allsides, bottom=allsides, diagonal)
    end

    wb   = get_workbook(sh)
    pfx = get_prefix("xl/styles.xml", get_xlsxfile(sh))
    pfx = pfx == "" ? pfx : pfx * ":"

    cell = getcell(sh, cellref)
    cell isa EmptyCell &&
        throw(XLSXError("Cannot set border for an `EmptyCell`: $(cellname(cellref)). Set the value first."))

    allXfNodes = find_all_nodes(_xpath("styleSheet", "cellXfs", "xf"), styles_xmlroot(wb))

    if cell.style == UInt32(0)
        cell.style = get_num_style_index(sh, allXfNodes, 0).id
    end

    cell_style      = styles_cell_xf(allXfNodes, cell.style)
    cell_borders    = getBorder(wb, cell_style)
    old_border_atts = cell_borders.border

    # Convert kwargs to dicts, then merge with old attributes per side
    side_kwargs = Dict(
        s => _to_border_dict(v) for (s, v) in (
            ("left", left), ("right", right), ("top", top),
            ("bottom", bottom), ("diagonal", diagonal)
        )
    )

    new_border_atts = Dict{String,Union{Dict{String,String},Nothing}}(
        side => _merge_border_side(side, side_kwargs[side], old_border_atts)
        for side in ("left", "right", "top", "bottom", "diagonal")
    )

    border_node  = buildNode("border", new_border_atts, pfx)
    new_borderid = styles_add_cell_attribute(wb, border_node, "borders")
    newstyle     = update_template_xf(sh, allXfNodes, CellDataFormat(cell.style), ["borderId", "applyBorder"], [string(new_borderid), "1"]).id
    cell.style   = newstyle

    return new_borderid
end

"""
    setUniformBorder(sh::Worksheet, cr::String; kw...) -> ::Int
    setUniformBorder(xf::XLSXFile,  cr::String, kw...) -> ::Int

    setUniformBorder(sh::Worksheet, rows, cols; kw...) -> ::Int

Set the border used by a cell range, a column range or row range or 
a named range in a worksheet or XLSXFile to be uniformly the same border.
Alternatively, specify the rows and columns using any combination of 
Integer, UnitRange, Vector{Integer} or `:`.

First, the border attributes of the first cell in the range (the top-left cell) are
updated according to the given `kw...` (using `setBorder()`). The resultant border is 
then applied to each remaining cell in the range.

As a result, every cell in the range will have a uniform border setting.

This differs from `setBorder()` which merges the attributes defined by `kw...` into 
the border definition used by each cell individually. For example, if you set the 
border style to `thin` for a range of cells, but these cells all use different border 
colors, `setBorder()` will change the border style but leave the border color unchanged 
for each cell individually. 

In contrast, `setUniformBorder()` will set the border `style` to `thin` for the first cell,
but will then apply all the border attributes from the updated first cell (ie. both `style` 
and `color`) to all the other cells in the range.

This can be more efficient when setting the same border for a large number of cells.

Applying `setUniformBorder()` without any keyword arguments simply copies the `Border` 
attributes from the first cell specified to all the others.

The value returned is the `borderId` of the border uniformly applied to the cells.
If all cells in the range are `EmptyCells` the returned value is -1.

For keyword definitions see [`setBorder()`](@ref).

Note: `setUniformBorder` cannot be used with the `outside` keyword.

# Examples:
```julia
Julia> setUniformBorder(sh, "B2:D6"; allsides = ["style" => "thick"], diagonal = ["style" => "hair"])

Julia> setUniformBorder(sh, [1, 2, 3], [3, 5, 9]; allsides = ["style" => "thick"], diagonal = ["style" => "hair", "color" => "yellow2"])

Julia> setUniformBorder(xf, "Sheet1!A1:F20"; left     = ["style" => "dotted", "color" => "FF000FF0"],
                                             right    = ["style" => "medium", "color" => "FF765000"],
                                             top      = ["style" => "thick",  "color" => "FF230000"],
                                             bottom   = ["style" => "medium", "color" => "FF0000FF"],
                                             diagonal = ["style" => "none"]
                                             )
                                             
julia> setUniformBorder(sh, "B2,A5:D22")     # Copy `Border` from B2 to cells in A5:D22

 
```
"""
function setUniformBorder end
setUniformBorder(ws::Worksheet, rng::SheetCellRange; kw...) = do_sheet_names_match(ws, rng) && setUniformBorder(ws, rng.rng; kw...)
setUniformBorder(ws::Worksheet, rng::SheetColumnRange; kw...) = do_sheet_names_match(ws, rng) && setUniformBorder(ws, rng.colrng; kw...)
setUniformBorder(ws::Worksheet, rng::SheetRowRange; kw...) = do_sheet_names_match(ws, rng) && setUniformBorder(ws, rng.rowrng; kw...)
setUniformBorder(ws::Worksheet, colrng::ColumnRange; kw...)::Int = process_columnranges(setUniformBorder, ws, colrng; kw...)
setUniformBorder(ws::Worksheet, rowrng::RowRange; kw...)::Int = process_rowranges(setUniformBorder, ws, rowrng; kw...)
setUniformBorder(ws::Worksheet, ncrng::NonContiguousRange; kw...)::Int = process_uniform_ncranges(setBorder, ws, ncrng, ["borderId", "applyBorder"]; kw...)
setUniformBorder(xl::XLSXFile, sheetcell::AbstractString; kw...)::Int = process_sheetcell(setUniformBorder, xl, sheetcell; kw...)
setUniformBorder(ws::Worksheet, ref_or_rng::AbstractString; kw...)::Int = process_ranges(setUniformBorder, ws, ref_or_rng; kw...)
setUniformBorder(ws::Worksheet, row::Union{Integer,UnitRange{<:Integer}}, ::Colon; kw...) = process_colon(setBorder, ws, row, nothing; kw...)
setUniformBorder(ws::Worksheet, ::Colon, col::Union{Integer,UnitRange{<:Integer}}; kw...) = process_colon(setBorder, ws, nothing, col; kw...)
setUniformBorder(ws::Worksheet, ::Colon, ::Colon; kw...) = process_colon(setUniformBorder, ws, nothing, nothing; kw...)
setUniformBorder(ws::Worksheet, ::Colon; kw...) = process_colon(setUniformBorder, ws, nothing, nothing; kw...)
setUniformBorder(ws::Worksheet, row::Union{Vector{Int},StepRange{<:Integer}}, ::Colon; kw...) = process_uniform_veccolon(setBorder, ws, row, nothing, ["borderId", "applyBorder"]; kw...)
setUniformBorder(ws::Worksheet, ::Colon, col::Union{Vector{Int},StepRange{<:Integer}}; kw...) = process_uniform_veccolon(setBorder, ws, nothing, col, ["borderId", "applyBorder"]; kw...)
setUniformBorder(ws::Worksheet, row::Union{Integer,UnitRange{<:Integer}}, col::Union{Vector{Int},StepRange{<:Integer}}; kw...) = process_uniform_vecint(setBorder, ws, row, col, ["borderId", "applyBorder"]; kw...)
setUniformBorder(ws::Worksheet, row::Union{Vector{Int},StepRange{<:Integer}}, col::Union{Integer,UnitRange{<:Integer}}; kw...) = process_uniform_vecint(setBorder, ws, row, col, ["borderId", "applyBorder"]; kw...)
setUniformBorder(ws::Worksheet, row::Union{Vector{Int},StepRange{<:Integer}}, col::Union{Vector{Int},StepRange{<:Integer}}; kw...) = process_uniform_vecint(setBorder, ws, row, col, ["borderId", "applyBorder"]; kw...)
setUniformBorder(ws::Worksheet, row::Union{Integer,UnitRange{<:Integer}}, col::Union{Integer,UnitRange{<:Integer}}; kw...) = setUniformBorder(ws, CellRange(CellRef(first(row), first(col)), CellRef(last(row), last(col))); kw...)
setUniformBorder(ws::Worksheet, rng::CellRange; kw...)::Int = process_uniform_attribute(setBorder, ws, rng, ["borderId", "applyBorder"]; kw...)

"""
    setOutsideBorder(sh::Worksheet, cr::String; kw...) -> ::Int
    setOutsideBorder(xf::XLSXFile,  cr::String, kw...) -> ::Int

    setOutsideBorder(sh::Worksheet, rows, cols; kw...) -> ::Int

Set the border around the outside of a cell range, a column range or row range 
or a named range in a worksheet or XLSXFile.
Alternatively, specify the rows and columns using integers, UnitRanges or `:`.

There is one key word:
- `outside::Union{Nothing,Vector{Pair{String,String}}}=nothing`

For keyword definition see [`setBorder()`](@ref).

Only the border definitions for the sides of boundary cells that are on the 
ouside edge of the range will be set to the specified style and color. The 
borders of internal edges and any diagonals will remain unchanged. Border 
settings for all internal cells in the range will remain unchanged.

Top and bottom borders for column ranges and left and right borders for 
row ranges are taken from the worksheet `dimension`.

An outside border cannot be set for a non-contiguous range.

The value returned is -1. Since outside borders are applied across multiple edge cells, each of which
may have different pre-existing border attributes and therefore different `borderId` values,
there is no single representative `borderId` for the operation as a whole.

# Examples:
```julia
Julia> setOutsideBorder(sh, "B2:D6"; outside = ["style" => "thick"])

Julia> setOutsideBorder(xf, "Sheet1!A1:F20"; outside = ["style" => "dotted", "color" => "FF000FF0"])
```

"""
function setOutsideBorder end
setOutsideBorder(ws::Worksheet, rng::SheetCellRange; kw...) = do_sheet_names_match(ws, rng) && setOutsideBorder(ws, rng.rng; kw...)
setOutsideBorder(ws::Worksheet, rng::SheetColumnRange; kw...) = do_sheet_names_match(ws, rng) && setOutsideBorder(ws, rng.colrng; kw...)
setOutsideBorder(ws::Worksheet, rng::SheetRowRange; kw...) = do_sheet_names_match(ws, rng) && setOutsideBorder(ws, rng.rowrng; kw...)
setOutsideBorder(ws::Worksheet, colrng::ColumnRange; kw...)::Int = process_columnranges(setOutsideBorder, ws, colrng; kw...)
setOutsideBorder(ws::Worksheet, rowrng::RowRange; kw...)::Int = process_rowranges(setOutsideBorder, ws, rowrng; kw...)
setOutsideBorder(xl::XLSXFile, sheetcell::AbstractString; kw...)::Int = process_sheetcell(setOutsideBorder, xl, sheetcell; kw...)
setOutsideBorder(ws::Worksheet, ref_or_rng::AbstractString; kw...)::Int = process_ranges(setOutsideBorder, ws, ref_or_rng; kw...)
setOutsideBorder(ws::Worksheet, row::Union{Integer,UnitRange{<:Integer}}, ::Colon; kw...) = process_colon(setOutsideBorder, ws, row, nothing; kw...)
setOutsideBorder(ws::Worksheet, ::Colon, col::Union{Integer,UnitRange{<:Integer}}; kw...) = process_colon(setOutsideBorder, ws, nothing, col; kw...)
setOutsideBorder(ws::Worksheet, ::Colon, ::Colon; kw...) = process_colon(setOutsideBorder, ws, nothing, nothing; kw...)
setOutsideBorder(ws::Worksheet, ::Colon; kw...) = process_colon(setOutsideBorder, ws, nothing, nothing; kw...)
setOutsideBorder(ws::Worksheet, row::Union{Integer,UnitRange{<:Integer}}, col::Union{Integer,UnitRange{<:Integer}}; kw...) = setOutsideBorder(ws, CellRange(CellRef(first(row), first(col)), CellRef(last(row), last(col))); kw...)

function setOutsideBorder(ws::Worksheet, rng::CellRange;
    outside::Union{Nothing,Vector{Pair{String,String}}}=nothing
)::Int
    get_xlsxfile(ws).is_writable ||
        throw(XLSXError("Cannot set borders because XLSXFile is not writable."))

    topLeft     = rng.start
    topRight    = CellRef(rng.start.row_number, rng.stop.column_number)
    bottomLeft  = CellRef(rng.stop.row_number,  rng.start.column_number)
    bottomRight = rng.stop

    setBorder(ws, CellRange(topLeft,    topRight);    top=outside)
    setBorder(ws, CellRange(topLeft,    bottomLeft);  left=outside)
    setBorder(ws, CellRange(topRight,   bottomRight); right=outside)
    setBorder(ws, CellRange(bottomLeft, bottomRight); bottom=outside)

    return -1

end

#
# -- Get and set fill attributes
#

"""
    getFill(sh::Worksheet, cr::String) -> ::Union{Nothing, CellFill}
    getFill(xf::XLSXFile, cr::String)  -> ::Union{Nothing, CellFill}

    getFill(sh::Worksheet, row::Int, col::Int) -> ::Union{Nothing, CellFill}
   
Get the fill used by a single cell at reference `cr` in a worksheet or XLSXFile.
The specified cell must be within the sheet dimension.

Return a `CellFill` object containing:
- `fillId`    : a 0-based index of the fill in the workbook
- `fill`      : a dictionary of fill attributes: borderAttribute -> (attribute -> value)
- `applyFill` : "1" or "0", indicating whether or not the fill is applied to the cell.

Return `nothing` if no cell fill is found.

The `fill` element in Excel's XML schema defines the pattern and color 
properties for cell fills. The primary attributes and child elements 
in the `patternFill` element are:
- `patternType` : Specifies the type of fill pattern (see below).
- `fgColor`     : Specifies the foreground color of the pattern.
- `bgColor`     : Specifies the background color of the pattern.

The child elements `fgColor` and `bgColor` can each have one or two attributes 
of their own. These color attributes are pushed in to the `CellFill.fill` Dict 
of attributes with either `fg` or `bg` prepended to their names to support later 
reconstruction of the xml element.

Thus:
`"patternFill" => Dict("patternType" => "solid", "bgindexed" => "64", "fgtheme" => "0")`
indicates a solid fill with a foreground color defined by theme 0 (in Excel) and 
background color defined by an indexed value. In this case (solid fill), the 
background color will be ignored.

Here is a list of the `patternTypes` Excel uses (thanks to Copilot!):
- `none`
- `solid`
- `mediumGray`
- `darkGray`
- `lightGray`
- `darkHorizontal`
- `darkVertical`
- `darkDown`
- `darkUp`
- `darkGrid`
- `darkTrellis`
- `lightHorizontal`
- `lightVertical`
- `lightDown`
- `lightUp`
- `lightGrid`
- `lightTrellis`
- `gray125`
- `gray0625`

Definitions for neither `fgColor` (foreground color) nor `bgColor` (background color) 
are strictly necessary although certain pattern types are more visually meaningful 
when both are defined. These pattern types include those that create a pattern or grid 
effect, where the contrast between the foreground and background colors enhances the 
visual presentation.

These pattern types include `darkTrellis`, `darkGrid`, `darkHorizontal`, `darkVertical`,
`darkDown`, `darkUp`, `mediumGray`, `lightGray`, `lightTrellis`, `lightGrid`, `lightHorizontal`,
`lightVertical`, `lightDown`, `lightUp`, `gray125` and `gray0625`.

If `fgColor` (foreground color) and `bgColor` (background color) are specified when they aren't 
needed, they will simply be ignored by Excel, and the default appearance will be applied.

# Examples:
```julia
julia> getFill(sh, "A1")

julia> getFill(sh, 3, 4)

julia> getFill(xf, "Sheet1!A1")
 
```
"""
function getFill end
getFill(xl::XLSXFile, sheetcell::String)::Union{Nothing,CellFill} = process_get_sheetcell(getFill, xl, sheetcell)
getFill(ws::Worksheet, cellref::CellRef)::Union{Nothing,CellFill} = process_get_cellref(getFill, ws, cellref)
getFill(ws::Worksheet, cr::String) = process_get_cellname(getFill, ws, cr)
getFill(ws::Worksheet, row::Integer, col::Integer) = getFill(ws, CellRef(row, col))
getDefaultFill(ws::Worksheet) = getFill(get_workbook(ws), styles_cell_xf(get_workbook(ws), 0))

function getFill(wb::Workbook, cell_style::XML.Node)::Union{Nothing,CellFill}
    get_xlsxfile(wb).use_cache_for_sheet_data ||
        throw(XLSXError("Cannot get fill because cache is not enabled."))

    haskey(cell_style, "fillId") || return nothing

    fillid_int  = parse(Int, cell_style["fillId"])
    applyfill = haskey(cell_style, "applyFill") ? cell_style["applyFill"] : "0"
    fill_elements = find_all_nodes(_xpath("styleSheet", "fills"), styles_xmlroot(wb))[begin]
    fill_nodes    = filter(n -> XML.nodetype(n) == XML.Element, XML.children(fill_elements))

    if parse(Int, fill_elements["count"]) != length(fill_nodes)
        throw(XLSXError("Unexpected number of fill definitions found: $(length(fill_nodes)). Expected $(fill_elements["count"])."))
    end

    current_fill = fill_nodes[fillid_int + 1]  # fillId is zero-based
    fill_atts    = Dict{String,Union{Dict{String,String},Nothing}}()

    for pattern in filter(n -> XML.nodetype(n) == XML.Element, XML.children(current_fill))
        tag = localname(pattern)
        a   = XML.attributes(pattern)
        if isnothing(a) || isempty(a)
            fill_atts[tag] = nothing
        else
            fill_atts[tag] = _parse_pattern_fill(pattern)
        end
    end

    return CellFill(fillid_int, fill_atts, applyfill)
end

"""
    setFill(sh::Worksheet, cr::String; kw...) -> ::Int
    setFill(xf::XLSXFile,  cr::String; kw...) -> ::Int

    setFill(sh::Worksheet, row, col; kw...) -> ::Int

Set the fill used used by a single cell, a cell range, a column range or 
row range or a named cell or named range in a worksheet or XLSXFile.
Alternatively, specify the row and column using any combination of 
Integer, UnitRange, Vector{Integer} or `:`.

The following keywords are used to define a fill:
- `pattern::String = nothing`   : Sets the patternType for the fill.
- `fgColor::String = nothing`   : Sets the foreground color for the fill.
- `bgColor::String = nothing`   : Sets the background color for the fill.

Here is a list of the available `pattern` values (thanks to Copilot!):
- `none`
- `solid`
- `mediumGray`
- `darkGray`
- `lightGray`
- `darkHorizontal`
- `darkVertical`
- `darkDown`
- `darkUp`
- `darkGrid`
- `darkTrellis`
- `lightHorizontal`
- `lightVertical`
- `lightDown`
- `lightUp`
- `lightGrid`
- `lightTrellis`
- `gray125`
- `gray0625`

The two colors may be set by specifying an 8-digit hexadecimal value for the `fgColor`
and/or `bgColor` keywords. 
Alternatively, you can use the name of any named color from [Colors.jl](https://juliagraphics.github.io/Colors.jl/stable/namedcolors/).

Setting only one or two of the attributes leaves the other attribute(s) unchanged 
for that cell's fill.

Fill attributes cannot be set for `EmptyCell`s. Set a cell value first.
If a cell range or column range includes any `EmptyCell`s, they will be
quietly skipped and the fill will be set for the remaining cells.

For single cells, the value returned is the `fillId` of the fill applied to the cell.
This can be used to apply the same fill to other cells or ranges.

For cell ranges, column ranges and named ranges, the value returned is -1.

# Examples:
```julia
Julia> setFill(sh, "B2"; pattern="gray125", bgColor = "FF000000")

Julia> setFill(xf, "Sheet1!A1:F20"; pattern="none", fgColor = "darkseagreen3")
 
Julia> setFill(sh, "11:24"; pattern="none", fgColor = "yellow2")
 
```
"""
function setFill end
setFill(ws::Worksheet, ref::SheetCellRef; kw...) = do_sheet_names_match(ws, ref) && setFill(ws, ref.cellref; kw...)
setFill(ws::Worksheet, rng::SheetCellRange; kw...) = do_sheet_names_match(ws, rng) && setFill(ws, rng.rng; kw...)
setFill(ws::Worksheet, rng::SheetColumnRange; kw...) = do_sheet_names_match(ws, rng) && setFill(ws, rng.colrng; kw...)
setFill(ws::Worksheet, rng::SheetRowRange; kw...) = do_sheet_names_match(ws, rng) && setFill(ws, rng.rowrng; kw...)
setFill(ws::Worksheet, rng::CellRange; kw...)::Int = process_cellranges(setFill, ws, rng; kw...)
setFill(ws::Worksheet, rowrng::RowRange; kw...)::Int = process_rowranges(setFill, ws, rowrng; kw...)
setFill(ws::Worksheet, ncrng::NonContiguousRange; kw...)::Int = process_ncranges(setFill, ws, ncrng; kw...)
setFill(ws::Worksheet, colrng::ColumnRange; kw...)::Int = process_columnranges(setFill, ws, colrng; kw...)
setFill(ws::Worksheet, ref_or_rng::AbstractString; kw...)::Int = process_ranges(setFill, ws, ref_or_rng; kw...)
setFill(xl::XLSXFile, sheetcell::String; kw...)::Int = process_sheetcell(setFill, xl, sheetcell; kw...)
#setFill(ws::Worksheet, row::Integer, col::Integer; kw...) = setFill(ws, CellRef(row, col); kw...)
setFill(ws::Worksheet, row::Union{Integer,UnitRange{<:Integer}}, ::Colon; kw...) = process_colon(setFill, ws, row, nothing; kw...)
setFill(ws::Worksheet, ::Colon, ::Colon; kw...) = process_colon(setFill, ws, nothing, nothing; kw...)
setFill(ws::Worksheet, ::Colon; kw...) = process_colon(setFill, ws, nothing, nothing; kw...)
setFill(ws::Worksheet, ::Colon, col::Union{Integer,UnitRange{<:Integer}}; kw...) = process_colon(setFill, ws, nothing, col; kw...)
setFill(ws::Worksheet, row::Union{Vector{Int},StepRange{<:Integer}}, ::Colon; kw...) = process_veccolon(setFill, ws, row, nothing; kw...)
setFill(ws::Worksheet, ::Colon, col::Union{Vector{Int},StepRange{<:Integer}}; kw...) = process_veccolon(setFill, ws, nothing, col; kw...)
setFill(ws::Worksheet, row::Union{Integer,UnitRange{<:Integer}}, col::Union{Vector{Int},StepRange{<:Integer}}; kw...) = process_vecint(setFill, ws, row, col; kw...)
setFill(ws::Worksheet, row::Union{Vector{Int},StepRange{<:Integer}}, col::Union{Integer,UnitRange{<:Integer}}; kw...) = process_vecint(setFill, ws, row, col; kw...)
setFill(ws::Worksheet, row::Union{Vector{Int},StepRange{<:Integer}}, col::Union{Vector{Int},StepRange{<:Integer}}; kw...) = process_vecint(setFill, ws, row, col; kw...)
setFill(ws::Worksheet, row::Union{Integer,UnitRange{<:Integer}}, col::Union{Integer,UnitRange{<:Integer}}; kw...) = setFill(ws, CellRange(CellRef(first(row), first(col)), CellRef(last(row), last(col))); kw...)

function setFill(sh::Worksheet, cellref::CellRef;
    pattern::Union{Nothing,String}=nothing,
    fgColor::Union{Nothing,String}=nothing,
    bgColor::Union{Nothing,String}=nothing,
)::Int
    get_xlsxfile(sh).is_writable || throw(XLSXError("Cannot set fill because XLSXFile is not writable."))

    # Validate before any work
    !isnothing(pattern) && pattern ∉ VALID_FILL_PATTERNS &&
        throw(XLSXError("Invalid pattern: $pattern. Must be one of: $(join(VALID_FILL_PATTERNS, ", "))."))

    wb   = get_workbook(sh)
    pfx = get_prefix("xl/styles.xml", get_xlsxfile(sh))
    pfx = pfx == "" ? pfx : pfx * ":"

    cell = getcell(sh, cellref)
    cell isa EmptyCell && throw(XLSXError("Cannot set fill for an `EmptyCell`: $(cellname(cellref)). Set the value first."))

    allXfNodes = find_all_nodes(_xpath("styleSheet", "cellXfs", "xf"), styles_xmlroot(wb))

    if cell.style == UInt32(0)
        cell.style = get_num_style_index(sh, allXfNodes, 0).id
    end

    cell_style   = styles_cell_xf(allXfNodes, cell.style)
    cell_fill    = getFill(wb, cell_style)
    old_fill_atts = cell_fill.fill["patternFill"]

    patternFill = Dict{String,String}()

    # patternType: preserve old or apply new
    result = _merge_val_tag("patternType", pattern, old_fill_atts)
    if !ismissing(result)
        patternFill["patternType"] = result isa Dict ? result["val"] : result
    end

    # fg/bg colors: prefix-match to preserve all old fg*/bg* keys, or set new rgb
    for (prefix, new_color, key) in (("fg", fgColor, "fgrgb"), ("bg", bgColor, "bgrgb"))
        if isnothing(new_color)
            for (k, v) in old_fill_atts
                startswith(k, prefix) && (patternFill[k] = v)
            end
        else
            patternFill[key] = get_color(new_color)
        end
    end

    new_fill_atts = Dict{String,Union{Dict{String,String},Nothing}}("patternFill" => patternFill)
    fill_node  = buildNode("fill", new_fill_atts, pfx)
    new_fillid = styles_add_cell_attribute(wb, fill_node, "fills")

    newstyle   = update_template_xf(sh, allXfNodes, CellDataFormat(cell.style), ["fillId", "applyFill"], [string(new_fillid), "1"]).id
    cell.style = newstyle

    return new_fillid
end

"""
    setUniformFill(sh::Worksheet, cr::String; kw...) -> ::Int
    setUniformFill(xf::XLSXFile,  cr::String, kw...) -> ::Int

    setUniformFill(sh::Worksheet, rows, cols; kw...) -> ::Int

Set the fill used by a cell range, a column range or row range or a 
named range in a worksheet or XLSXFile to be uniformly the same fill.
Alternatively, specify the rows and columns using any combination of 
Integer, UnitRange, Vector{Integer} or `:`.

First, the fill attributes of the first cell in the range (the top-left cell) are
updated according to the given `kw...` (using `setFill()`). The resultant fill is 
then applied to each remaining cell in the range.

As a result, every cell in the range will have a uniform fill setting.

This differs from `setFill()` which merges the attributes defined by `kw...` into 
the fill definition used by each cell individually. For example, if you set the 
fill `patern` to `darkGrid` for a range of cells, but these cells all use different fill  
`color`s, `setFill()` will change the fill `pattern` but leave the fill `color` unchanged 
for each cell individually. 

In contrast, `setUniformFill()` will set the fill `pattern` to `darkGrid` for the first cell,
but will then apply all the fill attributes from the updated first cell (ie. `pattern` 
and both foreground and background colors) to all the other cells in the range.

This can be more efficient when setting the same fill for a large number of cells.

Applying `setUniformFill()` without any keyword arguments simply copies the `Fill` 
attributes from the first cell specified to all the others.

The value returned is the `fillId` of the fill uniformly applied to the cells.
If all cells in the range are `EmptyCells` the returned value is -1.

For keyword definitions see [`setFill()`](@ref).

# Examples:
```julia
Julia> setUniformFill(sh, "B2:D4"; pattern="gray125", bgColor = "FF000000")

Julia> setUniformFill(xf, "Sheet1!A1:F20"; pattern="none", fgColor = "darkseagreen3")

julia> setUniformFill(sh, "B2,A5:D22")               # Copy `Fill` from B2 to cells in A5:D22
 
```
"""
function setUniformFill end
setUniformFill(ws::Worksheet, rng::SheetCellRange; kw...) = do_sheet_names_match(ws, rng) && setUniformFill(ws, rng.rng; kw...)
setUniformFill(ws::Worksheet, rng::SheetColumnRange; kw...) = do_sheet_names_match(ws, rng) && setUniformFill(ws, rng.colrng; kw...)
setUniformFill(ws::Worksheet, rng::SheetRowRange; kw...) = do_sheet_names_match(ws, rng) && setUniformFill(ws, rng.rowrng; kw...)
setUniformFill(ws::Worksheet, colrng::ColumnRange; kw...)::Int = process_columnranges(setUniformFill, ws, colrng; kw...)
setUniformFill(ws::Worksheet, rowrng::RowRange; kw...)::Int = process_rowranges(setUniformFill, ws, rowrng; kw...)
setUniformFill(ws::Worksheet, ncrng::NonContiguousRange; kw...)::Int = process_uniform_ncranges(setFill, ws, ncrng, ["fillId", "applyFill"]; kw...)
setUniformFill(xl::XLSXFile, sheetcell::AbstractString; kw...)::Int = process_sheetcell(setUniformFill, xl, sheetcell; kw...)
setUniformFill(ws::Worksheet, ref_or_rng::AbstractString; kw...)::Int = process_ranges(setUniformFill, ws, ref_or_rng; kw...)
setUniformFill(ws::Worksheet, row::Union{Integer,UnitRange{<:Integer}}, ::Colon; kw...) = process_colon(setFill, ws, row, nothing; kw...)
setUniformFill(ws::Worksheet, ::Colon, col::Union{Integer,UnitRange{<:Integer}}; kw...) = process_colon(setFill, ws, nothing, col; kw...)
setUniformFill(ws::Worksheet, ::Colon, ::Colon; kw...) = process_colon(setUniformFill, ws, nothing, nothing; kw...)
setUniformFill(ws::Worksheet, ::Colon; kw...) = process_colon(setUniformFill, ws, nothing, nothing; kw...)
setUniformFill(ws::Worksheet, row::Union{Vector{Int},StepRange{<:Integer}}, ::Colon; kw...) = process_uniform_veccolon(setFill, ws, row, nothing, ["fillId", "applyFill"]; kw...)
setUniformFill(ws::Worksheet, ::Colon, col::Union{Vector{Int},StepRange{<:Integer}}; kw...) = process_uniform_veccolon(setFill, ws, nothing, col, ["fillId", "applyFill"]; kw...)
setUniformFill(ws::Worksheet, row::Union{Integer,UnitRange{<:Integer}}, col::Union{Vector{Int},StepRange{<:Integer}}; kw...) = process_uniform_vecint(setFill, ws, row, col, ["fillId", "applyFill"]; kw...)
setUniformFill(ws::Worksheet, row::Union{Vector{Int},StepRange{<:Integer}}, col::Union{Integer,UnitRange{<:Integer}}; kw...) = process_uniform_vecint(setFill, ws, row, col, ["fillId", "applyFill"]; kw...)
setUniformFill(ws::Worksheet, row::Union{Vector{Int},StepRange{<:Integer}}, col::Union{Vector{Int},StepRange{<:Integer}}; kw...) = process_uniform_vecint(setFill, ws, row, col, ["fillId", "applyFill"]; kw...)
setUniformFill(ws::Worksheet, row::Union{Integer,UnitRange{<:Integer}}, col::Union{Integer,UnitRange{<:Integer}}; kw...) = setUniformFill(ws, CellRange(CellRef(first(row), first(col)), CellRef(last(row), last(col))); kw...)
setUniformFill(ws::Worksheet, rng::CellRange; kw...)::Int = process_uniform_attribute(setFill, ws, rng, ["fillId", "applyFill"]; kw...)

#
# -- Get and set alignment attributes
#

"""
    getAlignment(sh::Worksheet, cr::String) -> ::Union{Nothing, CellAlignment}
    getAlignment(xf::XLSXFile,  cr::String) -> ::Union{Nothing, CellAlignment}

    getAlignment(sh::Worksheet, row::Int, col::Int) -> ::Union{Nothing, CellAlignment}
   
Get the alignment used by a single cell at reference `cr` in a worksheet or XLSXFile.
The specified cell must be within the sheet dimension.

Return a `CellAlignment` object containing:
- `alignment`      : a dictionary of alignment attributes: alignmentAttribute -> (attribute -> value)
- `applyAlignment` : "1" or "0", indicating whether or not the alignment is applied to the cell.

Return `nothing` if no cell alignment is found.

The `alignment` element in Excel's XML schema defines the following attributes:
- `horizontal`     : Specifies the horizontal alignment of the cell.
- `vertical`       : Specifies the vertical alignment of the cell.
- `wrapText`       : Specifies whether ("1") or not ("0") the cell content wraps
                     in the cell.
- `shrinkToFit`    : Indicates whether ("1") or not ("0") the text should shrink to fit the cell.
- `indent`         : Specifies the number of spaces by which to indent the text (always from the left).
- `textRotation`   : Specifies the rotation angle of the text in a range -90 to 90 (positive values 
                     rotate the text counterclockwise).


Excel supports the following values for the horizontal alignment:
- `left`             : Aligns the text to the left of the cell.
- `center`           : Centers the text within the cell.
- `right`            : Aligns the text to the right of the cell.
- `fill`             : Repeats the text to fill the entire width of the cell.
- `justify`          : Justifies the text, spacing it out so that it spans the entire width of the cell.
- `centerContinuous` : Centers the text across multiple cells (specifically the currrent cell and all 
                       empty cells to the right) as if the text were in a merged cell.
- `distributed`      : Distributes the text evenly across the width of the cell.

Excel supports the following values for the vertical alignment:
- `top`              : Aligns the text to the top of the cell.
- `center`           : Centers the text vertically within the cell.
- `bottom`           : Aligns the text to the bottom of the cell.
- `justify`          : Justifies the text vertically, spreading it out evenly within the cell.
- `distributed`      : Distributes the text evenly from top to bottom in the cell.

# Examples:
```julia
julia> getAlignment(sh, "A1")

julia> getAlignment(sh, 2, 5) # Cell E2

julia> getAlignment(xf, "Sheet1!A1")
 
```
"""
function getAlignment end
getAlignment(xl::XLSXFile, sheetcell::String)::Union{Nothing,CellAlignment} = process_get_sheetcell(getAlignment, xl, sheetcell)
getAlignment(ws::Worksheet, cellref::CellRef)::Union{Nothing,CellAlignment} = process_get_cellref(getAlignment, ws, cellref)
getAlignment(ws::Worksheet, cr::String) = process_get_cellname(getAlignment, ws, cr)
getAlignment(ws::Worksheet, row::Integer, col::Integer) = getAlignment(ws, CellRef(row, col))

function getAlignment(wb::Workbook, cell_style::XML.Node)::Union{Nothing,CellAlignment}
    get_xlsxfile(wb).use_cache_for_sheet_data ||
        throw(XLSXError("Cannot get alignment because cache is not enabled."))

    # alignment is an optional child element of the cell xf node
    children = filter(n -> XML.nodetype(n) == XML.Element, XML.children(cell_style))
    alignment_node = findfirst(n -> localname(n) == "alignment", children)
    isnothing(alignment_node) && return nothing

    atts = Dict{String,String}(XML.attributes(children[alignment_node]))
    alignment_atts = Dict{String,Union{Dict{String,String},Nothing}}("alignment" => atts)
    applyalignment = haskey(cell_style, "applyAlignment") ? cell_style["applyAlignment"] : "0"

    return CellAlignment(alignment_atts, applyalignment)
end

"""
    setAlignment(sh::Worksheet, cr::String; kw...) -> ::Int}
    setAlignment(xf::XLSXFile,  cr::String; kw...) -> ::Int}

    setAlignment(sh::Worksheet, row, col; kw...) -> ::Int}

   
Set the alignment used used by a single cell, a cell range, a column range or 
row range or a named cell or named range in a worksheet or XLSXFile.
Alternatively, specify the row and column using any combination of 
Integer, UnitRange, Vector{Integer} or `:`.

The following keywords are used to define an alignment:
- `horizontal::String = nothing` : Sets the horizontal alignment.
- `vertical::String = nothing`   : Sets the vertical alignment.
- `wrapText::Bool = nothing`     : Determines whether the cell content wraps within the cell.
- `shrink::Bool = nothing`       : Indicates whether the text should shrink to fit the cell.
- `indent::Int = nothing`        : Specifies the number of spaces by which to indent the text 
                                   (always from the left).
- `rotation::Int = nothing`      : Specifies the rotation angle of the text in the range -90 to 90 
                                   (positive values rotate the text counterclockwise), 

Here are the possible values for the `horizontal` alignment:
- `left`             : Aligns the text to the left of the cell.
- `center`           : Centers the text within the cell.
- `right`            : Aligns the text to the right of the cell.
- `fill`             : Repeats the text to fill the entire width of the cell.
- `justify`          : Justifies the text, spacing it out so that it spans the entire 
                       width of the cell.
- `centerContinuous` : Centers the text across multiple cells (specifically the currrent 
                       cell and all empty cells to the right) as if the text were in 
                       a merged cell.
- `distributed`      : Distributes the text evenly across the width of the cell.

Here are the possible values for the `vertical` alignment:
- `top`              : Aligns the text to the top of the cell.
- `center`           : Centers the text vertically within the cell.
- `bottom`           : Aligns the text to the bottom of the cell.
- `justify`          : Justifies the text vertically, spreading it out evenly within the cell.
- `distributed`      : Distributes the text evenly from top to bottom in the cell.

For single cells, the value returned is the `styleId` of the cell.

For cell ranges, column ranges and named ranges, the value returned is -1.

# Examples:
```julia
julia> setAlignment(sh, "D18"; horizontal="center", wrapText=true)

julia> setAlignment(xf, "sheet1!D18"; horizontal="right", vertical="top", wrapText=true)

julia> setAlignment(sh, "L6"; horizontal="center", rotation="90", shrink=true, indent="2")

julia> setAlignment(sh, 1:3, 3:6; horizontal="center", rotation="90", shrink=true, indent="2")
 
```
"""
function setAlignment end
setAlignment(ws::Worksheet, ref::SheetCellRef; kw...) = do_sheet_names_match(ws, ref) && setAlignment(ws, ref.cellref; kw...)
setAlignment(ws::Worksheet, rng::SheetCellRange; kw...) = do_sheet_names_match(ws, rng) && setAlignment(ws, rng.rng; kw...)
setAlignment(ws::Worksheet, rng::SheetColumnRange; kw...) = do_sheet_names_match(ws, rng) && setAlignment(ws, rng.colrng; kw...)
setAlignment(ws::Worksheet, rng::SheetRowRange; kw...) = do_sheet_names_match(ws, rng) && setAlignment(ws, rng.rowrng; kw...)
setAlignment(ws::Worksheet, rng::CellRange; kw...)::Int = process_cellranges(setAlignment, ws, rng; kw...)
setAlignment(ws::Worksheet, colrng::ColumnRange; kw...)::Int = process_columnranges(setAlignment, ws, colrng; kw...)
setAlignment(ws::Worksheet, rowrng::RowRange; kw...)::Int = process_rowranges(setAlignment, ws, rowrng; kw...)
setAlignment(ws::Worksheet, ncrng::NonContiguousRange; kw...)::Int = process_ncranges(setAlignment, ws, ncrng; kw...)
setAlignment(ws::Worksheet, ref_or_rng::AbstractString; kw...)::Int = process_ranges(setAlignment, ws, ref_or_rng; kw...)
setAlignment(xl::XLSXFile, sheetcell::String; kw...)::Int = process_sheetcell(setAlignment, xl, sheetcell; kw...)
#setAlignment(ws::Worksheet, row::Integer, col::Integer; kw...) = setAlignment(ws, CellRef(row, col); kw...)
setAlignment(ws::Worksheet, row::Union{Integer,UnitRange{<:Integer}}, ::Colon; kw...) = process_colon(setAlignment, ws, row, nothing; kw...)
setAlignment(ws::Worksheet, ::Colon, ::Colon; kw...) = process_colon(setAlignment, ws, nothing, nothing; kw...)
setAlignment(ws::Worksheet, ::Colon; kw...) = process_colon(setAlignment, ws, nothing, nothing; kw...)
setAlignment(ws::Worksheet, ::Colon, col::Union{Integer,UnitRange{<:Integer}}; kw...) = process_colon(setAlignment, ws, nothing, col; kw...)
setAlignment(ws::Worksheet, row::Union{Vector{Int},StepRange{<:Integer}}, ::Colon; kw...) = process_veccolon(setAlignment, ws, row, nothing; kw...)
setAlignment(ws::Worksheet, ::Colon, col::Union{Vector{Int},StepRange{<:Integer}}; kw...) = process_veccolon(setAlignment, ws, nothing, col; kw...)
setAlignment(ws::Worksheet, row::Union{Integer,UnitRange{<:Integer}}, col::Union{Vector{Int},StepRange{<:Integer}}; kw...) = process_vecint(setAlignment, ws, row, col; kw...)
setAlignment(ws::Worksheet, row::Union{Vector{Int},StepRange{<:Integer}}, col::Union{Integer,UnitRange{<:Integer}}; kw...) = process_vecint(setAlignment, ws, row, col; kw...)
setAlignment(ws::Worksheet, row::Union{Vector{Int},StepRange{<:Integer}}, col::Union{Vector{Int},StepRange{<:Integer}}; kw...) = process_vecint(setAlignment, ws, row, col; kw...)
setAlignment(ws::Worksheet, row::Union{Integer,UnitRange{<:Integer}}, col::Union{Integer,UnitRange{<:Integer}}; kw...) = setAlignment(ws, CellRange(CellRef(first(row), first(col)), CellRef(last(row), last(col))); kw...)

function setAlignment(sh::Worksheet, cellref::CellRef;
    horizontal::Union{Nothing,String}=nothing,
    vertical::Union{Nothing,String}=nothing,
    wrapText::Union{Nothing,Bool}=nothing,
    shrink::Union{Nothing,Bool}=nothing,
    indent::Union{Nothing,Int}=nothing,
    rotation::Union{Nothing,Int}=nothing
)::Int
    get_xlsxfile(sh).is_writable ||
        throw(XLSXError("Cannot set alignment because XLSXFile is not writable."))

    # Validate before any work
    !isnothing(horizontal) && horizontal ∉ VALID_HORIZONTAL_ALIGNMENTS &&
        throw(XLSXError("Invalid horizontal alignment: $horizontal. Must be one of: $(join(VALID_HORIZONTAL_ALIGNMENTS, ", "))."))
    !isnothing(vertical) && vertical ∉ VALID_VERTICAL_ALIGNMENTS &&
        throw(XLSXError("Invalid vertical alignment: $vertical. Must be one of: $(join(VALID_VERTICAL_ALIGNMENTS, ", "))."))
    !isnothing(indent) && indent < 0 &&
        throw(XLSXError("Invalid indent value: $indent. Must be a positive integer."))
    !isnothing(rotation) && rotation ∉ -90:90 &&
        throw(XLSXError("Invalid rotation value: $rotation. Must be an integer between -90 and 90."))

    wb   = get_workbook(sh)
    pfx = get_prefix("xl/styles.xml", get_xlsxfile(sh))
    pfx = pfx == "" ? pfx : pfx * ":"
    
    cell = getcell(sh, cellref)
    cell isa EmptyCell &&
        throw(XLSXError("Cannot set alignment for an `EmptyCell`: $(cellname(cellref)). Set the value first."))

    allXfNodes = find_all_nodes(_xpath("styleSheet", "cellXfs", "xf"), styles_xmlroot(wb))

    if cell.style == UInt32(0)
        cell.style = get_num_style_index(sh, allXfNodes, 0).id
    end

    cell_style     = styles_cell_xf(allXfNodes, cell.style)
    cell_alignment = getAlignment(wb, cell_style)
    old_atts       = isnothing(cell_alignment) ? Dict{String,String}() : cell_alignment.alignment["alignment"]

    atts = OrderedDict{String,String}()
    _merge_alignment_att(atts, "horizontal",   horizontal, old_atts)
    _merge_alignment_att(atts, "vertical",     vertical,   old_atts)
    _merge_alignment_att(atts, "wrapText",     wrapText,   old_atts, b -> b ? "1" : "0")
    _merge_alignment_att(atts, "shrinkToFit",  shrink,     old_atts, b -> b ? "1" : "0")
    _merge_alignment_att(atts, "indent",       indent,     old_atts)
    _merge_alignment_att(atts, "textRotation", rotation,   old_atts)

    alignment_node = XML.Node(XML.Element, "$(pfx)alignment", atts, nothing, nothing)
    newstyle       = update_template_xf(sh, allXfNodes, CellDataFormat(cell.style), alignment_node).id
    cell.style     = newstyle

    return Int(newstyle)
end

"""
    setUniformAlignment(sh::Worksheet, cr::String; kw...) -> ::Int
    setUniformAlignment(xf::XLSXFile,  cr::String, kw...) -> ::Int

    setUniformAlignment(sh::Worksheet, rows, cols; kw...) -> ::Int

Set the alignment used by a cell range, a column range or row range or a 
named range in a worksheet or XLSXFile to be uniformly the same alignment.
Alternatively, specify the rows and columns using any combination of 
Integer, UnitRange, Vector{Integer} or `:`.

First, the alignment attributes of the first cell in the range (the top-left cell) are
updated according to the given `kw...` (using `setAlignment()`). The resultant alignment 
is then applied to each remaining cell in the range.

As a result, every cell in the range will have a uniform alignment setting.

This differs from `setAlignment()` which merges the attributes defined by `kw...` into 
the alignment definition used by each cell individually. For example, if you set the 
`horizontal` alignment to `left` for a range of cells, but these cells all use different 
`vertical` alignment or `wrapText`, `setAlignment()` will change the horizontal alignment but 
leave the `vertical` alignment and `wrapText` unchanged for each cell individually. 

In contrast, `setUniformAlignment()` will set the `horizontal` alignment to `left` for  
the first cell, but will then apply all the alignment attributes from the updated first  
cell to all the other cells in the range.

This can be more efficient when setting the same alignment for a large number of cells.

Applying `setUniformAlignment()` without any keyword arguments simply copies the `Alignment` 
attributes from the first cell specified to all the others.

The value returned is the `styleId` of the reference (top-left) cell, from which the 
alignment uniformly applied to the cells was taken.
If all cells in the range are `EmptyCells`, the returned value is -1.

For keyword definitions see [`setAlignment()`](@ref).

# Examples:
```julia
Julia> setUniformAlignment(sh, "B2:D4"; horizontal="center", wrap = true)

Julia> setUniformAlignment(xf, "Sheet1!A1:F20"; horizontal="center", vertical="top")

Julia> setUniformAlignment(sh, :, 1:24; horizontal="center", vertical="top")

julia> setUniformAlignment(sh, "B2,A5:D22")                # Copy `Alignment` from B2 to cells in A5:D22
 
```
"""
function setUniformAlignment end
setUniformAlignment(ws::Worksheet, rng::SheetCellRange; kw...) = do_sheet_names_match(ws, rng) && setUniformAlignment(ws, rng.rng; kw...)
setUniformAlignment(ws::Worksheet, rng::SheetColumnRange; kw...) = do_sheet_names_match(ws, rng) && setUniformAlignment(ws, rng.colrng; kw...)
setUniformAlignment(ws::Worksheet, rng::SheetRowRange; kw...) = do_sheet_names_match(ws, rng) && setUniformAlignment(ws, rng.rowrng; kw...)
setUniformAlignment(ws::Worksheet, colrng::ColumnRange; kw...)::Int = process_columnranges(setUniformAlignment, ws, colrng; kw...)
setUniformAlignment(ws::Worksheet, rowrng::RowRange; kw...)::Int = process_rowranges(setUniformAlignment, ws, rowrng; kw...)
setUniformAlignment(ws::Worksheet, ncrng::NonContiguousRange; kw...)::Int = process_uniform_ncranges(setAlignment, ws, ncrng; kw...)
setUniformAlignment(xl::XLSXFile, sheetcell::AbstractString; kw...)::Int = process_sheetcell(setUniformAlignment, xl, sheetcell; kw...)
setUniformAlignment(ws::Worksheet, ref_or_rng::AbstractString; kw...)::Int = process_ranges(setUniformAlignment, ws, ref_or_rng; kw...)
setUniformAlignment(ws::Worksheet, row::Union{Integer,UnitRange{<:Integer}}, ::Colon; kw...) = process_colon(setAlignment, ws, row, nothing; kw...)
setUniformAlignment(ws::Worksheet, ::Colon, col::Union{Integer,UnitRange{<:Integer}}; kw...) = process_colon(setAlignment, ws, nothing, col; kw...)
setUniformAlignment(ws::Worksheet, ::Colon, ::Colon; kw...) = process_colon(setUniformAlignment, ws, nothing, nothing; kw...)
setUniformAlignment(ws::Worksheet, ::Colon; kw...) = process_colon(setUniformAlignment, ws, nothing, nothing; kw...)
setUniformAlignment(ws::Worksheet, row::Union{Vector{Int},StepRange{<:Integer}}, ::Colon; kw...) = process_uniform_veccolon(setAlignment, ws, row, nothing; kw...)
setUniformAlignment(ws::Worksheet, ::Colon, col::Union{Vector{Int},StepRange{<:Integer}}; kw...) = process_uniform_veccolon(setAlignment, ws, nothing, col; kw...)
setUniformAlignment(ws::Worksheet, row::Union{Integer,UnitRange{<:Integer}}, col::Union{Vector{Int},StepRange{<:Integer}}; kw...) = process_uniform_vecint(setAlignment, ws, row, col; kw...)
setUniformAlignment(ws::Worksheet, row::Union{Vector{Int},StepRange{<:Integer}}, col::Union{Integer,UnitRange{<:Integer}}; kw...) = process_uniform_vecint(setAlignment, ws, row, col; kw...)
setUniformAlignment(ws::Worksheet, row::Union{Vector{Int},StepRange{<:Integer}}, col::Union{Vector{Int},StepRange{<:Integer}}; kw...) = process_uniform_vecint(setAlignment, ws, row, col; kw...)
setUniformAlignment(ws::Worksheet, row::Union{Integer,UnitRange{<:Integer}}, col::Union{Integer,UnitRange{<:Integer}}; kw...) = setUniformAlignment(ws, CellRange(CellRef(first(row), first(col)), CellRef(last(row), last(col))); kw...)
setUniformAlignment(ws::Worksheet, rng::CellRange; kw...)::Int = process_uniform_attribute(setAlignment, ws, rng; kw...)

#
# -- Get and set number format attributes
#

"""
    getFormat(sh::Worksheet, cr::String) -> ::Union{Nothing, CellFormat}
    getFormat(xf::XLSXFile,  cr::String) -> ::Union{Nothing, CellFormat}

    getFormat(sh::Worksheet, row::Int, col::int) -> ::Union{Nothing, CellFormat}
   
Get the format (numFmt) used by a single cell at reference `cr` in a worksheet or XLSXFile.
The specified cell must be within the sheet dimension.

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
    get_xlsxfile(wb).use_cache_for_sheet_data ||
        throw(XLSXError("Cannot get number formats because cache is not enabled."))

    haskey(cell_style, "numFmtId") || return nothing

    numfmtid_int      = parse(Int, cell_style["numFmtId"])
    applynumberformat = haskey(cell_style, "applyNumberFormat") ? cell_style["applyNumberFormat"] : "0"
    format_atts       = Dict{String,Union{Dict{String,String},Nothing}}()

    if numfmtid_int >= PREDEFINED_NUMFMT_COUNT
        # Custom format — look up in the numFmts table
        format_elements = find_all_nodes(_xpath("styleSheet", "numFmts"), styles_xmlroot(wb))[begin]
        format_nodes    = filter(n -> XML.nodetype(n) == XML.Element, XML.children(format_elements))

        if parse(Int, format_elements["count"]) != length(format_nodes)
            throw(XLSXError("Unexpected number of format definitions found: $(length(format_nodes)). Expected $(format_elements["count"])."))
        end

        idx = findfirst(n -> haskey(n, "numFmtId") && n["numFmtId"] == string(numfmtid_int), format_nodes)
        isnothing(idx) && throw(XLSXError("No format definition found for numFmtId $numfmtid_int."))
        current_format = format_nodes[idx]

        atts = Dict{String,String}(k => XLSX.unescape(v) for (k, v) in XML.attributes(current_format))
        format_atts[localname(current_format)] = atts

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
setFormat(ws::Worksheet, ref::SheetCellRef; kw...) = do_sheet_names_match(ws, ref) && setFormat(ws, ref.cellref; kw...)
setFormat(ws::Worksheet, rng::SheetCellRange; kw...) = do_sheet_names_match(ws, rng) && setFormat(ws, rng.rng; kw...)
setFormat(ws::Worksheet, rng::SheetColumnRange; kw...) = do_sheet_names_match(ws, rng) && setFormat(ws, rng.colrng; kw...)
setFormat(ws::Worksheet, rng::SheetRowRange; kw...) = do_sheet_names_match(ws, rng) && setFormat(ws, rng.rowrng; kw...)
setFormat(ws::Worksheet, rng::CellRange; kw...)::Int = process_cellranges(setFormat, ws, rng; kw...)
setFormat(ws::Worksheet, colrng::ColumnRange; kw...)::Int = process_columnranges(setFormat, ws, colrng; kw...)
setFormat(ws::Worksheet, ncrng::NonContiguousRange; kw...)::Int = process_ncranges(setFormat, ws, ncrng; kw...)
setFormat(ws::Worksheet, rowrng::RowRange; kw...)::Int = process_rowranges(setFormat, ws, rowrng; kw...)
setFormat(ws::Worksheet, ref_or_rng::AbstractString; kw...)::Int = process_ranges(setFormat, ws, ref_or_rng; kw...)
setFormat(xl::XLSXFile, sheetcell::String; kw...)::Int = process_sheetcell(setFormat, xl, sheetcell; kw...)
#setFormat(ws::Worksheet, row::Integer, col::Integer; kw...) = setFormat(ws, CellRef(row, col); kw...)
setFormat(ws::Worksheet, row::Union{Integer,UnitRange{<:Integer}}, ::Colon; kw...) = process_colon(setFormat, ws, row, nothing; kw...)
setFormat(ws::Worksheet, ::Colon, col::Union{Integer,UnitRange{<:Integer}}; kw...) = process_colon(setFormat, ws, nothing, col; kw...)
setFormat(ws::Worksheet, ::Colon, ::Colon; kw...) = process_colon(setFormat, ws, nothing, nothing; kw...)
setFormat(ws::Worksheet, ::Colon; kw...) = process_colon(setFormat, ws, nothing, nothing; kw...)
setFormat(ws::Worksheet, row::Union{Vector{Int},StepRange{<:Integer}}, ::Colon; kw...) = process_veccolon(setFormat, ws, row, nothing; kw...)
setFormat(ws::Worksheet, ::Colon, col::Union{Vector{Int},StepRange{<:Integer}}; kw...) = process_veccolon(setFormat, ws, nothing, col; kw...)
setFormat(ws::Worksheet, row::Union{Integer,UnitRange{<:Integer}}, col::Union{Vector{Int},StepRange{<:Integer}}; kw...) = process_vecint(setFormat, ws, row, col; kw...)
setFormat(ws::Worksheet, row::Union{Vector{Int},StepRange{<:Integer}}, col::Union{Integer,UnitRange{<:Integer}}; kw...) = process_vecint(setFormat, ws, row, col; kw...)
setFormat(ws::Worksheet, row::Union{Vector{Int},StepRange{<:Integer}}, col::Union{Vector{Int},StepRange{<:Integer}}; kw...) = process_vecint(setFormat, ws, row, col; kw...)
setFormat(ws::Worksheet, row::Union{Integer,UnitRange{<:Integer}}, col::Union{Integer,UnitRange{<:Integer}}; kw...) = setFormat(ws, CellRange(CellRef(first(row), first(col)), CellRef(last(row), last(col))); kw...)

function setFormat(sh::Worksheet, cellref::CellRef;
    format::Union{Nothing,String}=nothing
)::Int
    get_xlsxfile(sh).is_writable ||
        throw(XLSXError("Cannot set number format because XLSXFile is not writable."))

    wb   = get_workbook(sh)
    cell = getcell(sh, cellref)
    cell isa EmptyCell &&
        throw(XLSXError("Cannot set format for an `EmptyCell`: $(cellname(cellref)). Set the value first."))

    allXfNodes = find_all_nodes(_xpath("styleSheet", "cellXfs", "xf"), styles_xmlroot(wb))

    if cell.style == UInt32(0)
        cell.style = get_num_style_index(sh, allXfNodes, 0).id
    end

    cell_style  = styles_cell_xf(allXfNodes, cell.style)
    cell_format = getFormat(wb, cell_style)

    # No format kwarg — return existing format id, or 0 if none set
    if isnothing(format)
        return isnothing(cell_format) ? 0 : cell_format.numFmtId
    end

    # Resolve new format id: built-in lookup or next available custom id
    new_formatid = if haskey(BUILTIN_FORMAT_CODES, format)
        BUILTIN_FORMAT_CODES[format]
    else
        get_new_formatId(wb, format)
    end

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
setUniformFormat(ws::Worksheet, rng::SheetCellRange; kw...) = do_sheet_names_match(ws, rng) && setUniformFormat(ws, rng.rng; kw...)
setUniformFormat(ws::Worksheet, rng::SheetColumnRange; kw...) = do_sheet_names_match(ws, rng) && setUniformFormat(ws, rng.colrng; kw...)
setUniformFormat(ws::Worksheet, rng::SheetRowRange; kw...) = do_sheet_names_match(ws, rng) && setUniformFormat(ws, rng.rowrng; kw...)
setUniformFormat(ws::Worksheet, colrng::ColumnRange; kw...)::Int = process_columnranges(setUniformFormat, ws, colrng; kw...)
setUniformFormat(ws::Worksheet, rowrng::RowRange; kw...)::Int = process_rowranges(setUniformFormat, ws, rowrng; kw...)
setUniformFormat(ws::Worksheet, ncrng::NonContiguousRange; kw...)::Int = process_uniform_ncranges(setFormat, ws, ncrng, ["numFmtId", "applyNumberFormat"]; kw...)
setUniformFormat(xl::XLSXFile, sheetcell::AbstractString; kw...)::Int = process_sheetcell(setUniformFormat, xl, sheetcell; kw...)
setUniformFormat(ws::Worksheet, ref_or_rng::AbstractString; kw...)::Int = process_ranges(setUniformFormat, ws, ref_or_rng; kw...)
setUniformFormat(ws::Worksheet, row::Union{Integer,UnitRange{<:Integer}}, ::Colon; kw...) = process_colon(setFormat, ws, row, nothing; kw...)
setUniformFormat(ws::Worksheet, ::Colon, col::Union{Integer,UnitRange{<:Integer}}; kw...) = process_colon(setFormat, ws, nothing, col; kw...)
setUniformFormat(ws::Worksheet, ::Colon, ::Colon; kw...) = process_colon(setUniformFormat, ws, nothing, nothing; kw...)
setUniformFormat(ws::Worksheet, ::Colon; kw...) = process_colon(setUniformFormat, ws, nothing, nothing; kw...)
setUniformFormat(ws::Worksheet, row::Union{Vector{Int},StepRange{<:Integer}}, ::Colon; kw...) = process_uniform_veccolon(setFormat, ws, row, nothing, ["numFmtId", "applyNumberFormat"]; kw...)
setUniformFormat(ws::Worksheet, ::Colon, col::Union{Vector{Int},StepRange{<:Integer}}; kw...) = process_uniform_veccolon(setFormat, ws, nothing, col, ["numFmtId", "applyNumberFormat"]; kw...)
setUniformFormat(ws::Worksheet, row::Union{Integer,UnitRange{<:Integer}}, col::Union{Vector{Int},StepRange{<:Integer}}; kw...) = process_uniform_vecint(setFormat, ws, row, col, ["numFmtId", "applyNumberFormat"]; kw...)
setUniformFormat(ws::Worksheet, row::Union{Vector{Int},StepRange{<:Integer}}, col::Union{Integer,UnitRange{<:Integer}}; kw...) = process_uniform_vecint(setFormat, ws, row, col, ["numFmtId", "applyNumberFormat"]; kw...)
setUniformFormat(ws::Worksheet, row::Union{Vector{Int},StepRange{<:Integer}}, col::Union{Vector{Int},StepRange{<:Integer}}; kw...) = process_uniform_vecint(setFormat, ws, row, col, ["numFmtId", "applyNumberFormat"]; kw...)
setUniformFormat(ws::Worksheet, row::Union{Integer,UnitRange{<:Integer}}, col::Union{Integer,UnitRange{<:Integer}}; kw...) = setUniformFormat(ws, CellRange(CellRef(first(row), first(col)), CellRef(last(row), last(col))); kw...)
setUniformFormat(ws::Worksheet, rng::CellRange; kw...)::Int = process_uniform_attribute(setFormat, ws, rng, ["numFmtId", "applyNumberFormat"]; kw...)

#
# -- Set uniform styles
#

"""
    setUniformStyle(sh::Worksheet, cr::String) -> ::Int
    setUniformStyle(xf::XLSXFile,  cr::String) -> ::Int

    setUniformStyle(sh::Worksheet, rows, cols) -> ::Int

Set the cell `style` used by a cell range, a column range or row range 
or a named range in a worksheet or XLSXFile to be the same as that of 
the first cell in the range that is not an `EmptyCell`.
Alternatively, specify the rows and columns using any combination of 
Integer, UnitRange, Vector{Integer} or `:`.

As a result, every cell in the range will have a uniform `style`.

A cell `style` consists of the collection of `format`, `alignment`, `border`, 
`font` and `fill`.

If the first cell has no defined `style` (`s=""`), all cells will be given the 
same undefined `style`.

The value returned is the `styleId` of the `style` uniformly applied to the cells or 
`nothing` if the style is undefined.
If all cells in the range are `EmptyCells`, the returned value is -1.

# Examples:
```julia
julia> XLSX.setUniformStyle(xf, "Sheet1!A2:L6")

julia> XLSX.setUniformStyle(sh, "F1:F5")

julia> XLSX.setUniformStyle(sh, 2:5, 5)

julia> XLSX.setUniformStyle(sh, 2, :)
 
```
"""
function setUniformStyle end
setUniformStyle(ws::Worksheet, rng::SheetCellRange) = do_sheet_names_match(ws, rng) && setUniformStyle(ws, rng.rng)
setUniformStyle(ws::Worksheet, rng::SheetColumnRange) = do_sheet_names_match(ws, rng) && setUniformStyle(ws, rng.colrng)
setUniformStyle(ws::Worksheet, rng::SheetRowRange) = do_sheet_names_match(ws, rng) && setUniformStyle(ws, rng.rowrng)
setUniformStyle(ws::Worksheet, colrng::ColumnRange) = process_columnranges(setUniformStyle, ws, colrng)
setUniformStyle(ws::Worksheet, rowrng::RowRange) = process_rowranges(setUniformStyle, ws, rowrng)
setUniformStyle(ws::Worksheet, ncrng::NonContiguousRange) = process_uniform_ncranges(ws, ncrng)
setUniformStyle(xl::XLSXFile, sheetcell::AbstractString) = process_sheetcell(setUniformStyle, xl, sheetcell)
setUniformStyle(ws::Worksheet, ref_or_rng::AbstractString) = process_ranges(setUniformStyle, ws, ref_or_rng)
setUniformStyle(ws::Worksheet, row::Union{Integer,UnitRange{<:Integer}}, ::Colon) = process_colon(ws, row, nothing)
setUniformStyle(ws::Worksheet, ::Colon, col::Union{Integer,UnitRange{<:Integer}}) = process_colon(ws, nothing, col)
setUniformStyle(ws::Worksheet, ::Colon, ::Colon) = process_colon(ws, nothing, nothing)
setUniformStyle(ws::Worksheet, ::Colon) = process_colon(ws, nothing, nothing)
setUniformStyle(ws::Worksheet, row::Union{Vector{Int},StepRange{<:Integer}}, ::Colon) = process_uniform_veccolon(ws, row, nothing)
setUniformStyle(ws::Worksheet, ::Colon, col::Union{Vector{Int},StepRange{<:Integer}}) = process_uniform_veccolon(ws, nothing, col)
setUniformStyle(ws::Worksheet, row::Union{Integer,UnitRange{<:Integer}}, col::Union{Vector{Int},StepRange{<:Integer}}) = process_uniform_vecint(ws, row, col)
setUniformStyle(ws::Worksheet, row::Union{Vector{Int},StepRange{<:Integer}}, col::Union{Integer,UnitRange{<:Integer}}) = process_uniform_vecint(ws, row, col)
setUniformStyle(ws::Worksheet, row::Union{Vector{Int},StepRange{<:Integer}}, col::Union{Vector{Int},StepRange{<:Integer}}) = process_uniform_vecint(ws, row, col)
setUniformStyle(ws::Worksheet, row::Union{Integer,UnitRange{<:Integer}}, col::Union{Integer,UnitRange{<:Integer}}) = setUniformStyle(ws, CellRange(CellRef(first(row), first(col)), CellRef(last(row), last(col))))

function setUniformStyle(ws::Worksheet, rng::CellRange)::Union{Nothing,Int}
    get_xlsxfile(ws).is_writable ||
        throw(XLSXError("Cannot set styles because XLSXFile is not writable."))

    newid::Union{Nothing,Int} = nothing
    is_first   = true
    first_font = nothing

    for cellref in rng
        newid, is_first, first_font = process_uniform_core(ws, cellref, newid, is_first, first_font)
    end

    # -1 signals that all cells in the range were EmptyCells
    is_first && return -1

    return newid
end

#
# -- Get and set column width
#

"""
    setColumnWidth(sh::Worksheet, cr::String; kw...) -> ::Int
    setColumnWidth(xf::XLSXFile,  cr::String, kw...) -> ::Int

    setColumnWidth(sh::Worksheet, row, col; kw...) -> ::Int
    setColumnWidth(sh::Worksheet, col; kw...) -> ::Int

Set the width of a column or column range.

A standard cell reference or cell range can be used to define the column range. 
The function will use the columns and ignore the rows. Named cells and named
ranges can similarly be used.
Alternatively, specify the row and column using any combination of 
Integer, UnitRange, Vector{Integer} or `:`, but only the columns will be used.
The row can be omitted and column specified alone, using an integer or colon.


The function uses one keyword used to define a column width:
- `width::Real = nothing` : Defines width in Excel's own (internal) units

When you set a column widths interactively in Excel you can see the width 
in "internal" units and in pixels. The width stored in the xlsx file is slightly 
larger than the width shown intertactively because Excel adds some cell padding. 
The method Excel uses to calculate the padding is obscure and complex. This 
function does not attempt to replicate it, but simply adds 0.71 internal units 
to the value specified. The value set is unlikely to match the value seen 
interactivley in the resultant spreadsheet, but will be close.

You can set a column width to 0.

The function returns a value of 0. The return value is meaningless but required to comply with the
range dispatch infrastructure, which expects an `Int` return from all `set*` functions.

`setColumnWidth` requires a file to be open for writing as well as reading (`mode="rw"` or open as a template) but 
it can work on empty cells.

# Examples:
```julia
julia> XLSX.setColumnWidth(xf, "Sheet1!A2"; width = 50)

julia> XLSX.seColumnWidth(sh, "F1:F5"; width = 0)

julia> XLSX.setColumnWidth(sh, "I"; width = 24.37)
 
```
"""
function setColumnWidth end
setColumnWidth(ws::Worksheet, ref::SheetCellRef; kw...) = do_sheet_names_match(ws, ref) && setColumnWidth(ws, ref.cellref; kw...)
setColumnWidth(ws::Worksheet, rng::SheetCellRange; kw...) = do_sheet_names_match(ws, rng) && setColumnWidth(ws, rng.rng; kw...)
setColumnWidth(ws::Worksheet, rng::SheetColumnRange; kw...) = do_sheet_names_match(ws, rng) && setColumnWidth(ws, rng.colrng; kw...)
setColumnWidth(ws::Worksheet, rng::SheetRowRange; kw...) = do_sheet_names_match(ws, rng) && setColumnWidth(ws, rng.rowrng; kw...)
setColumnWidth(ws::Worksheet, colrng::ColumnRange; kw...)::Int = process_columnranges(setColumnWidth, ws, colrng; kw...)
setColumnWidth(ws::Worksheet, rowrng::RowRange; kw...)::Int = process_rowranges(setColumnWidth, ws, rowrng; kw...)
setColumnWidth(ws::Worksheet, ncrng::NonContiguousRange; kw...)::Int = process_ncranges(setColumnWidth, ws, ncrng; kw...)
setColumnWidth(ws::Worksheet, ref_or_rng::AbstractString; kw...)::Int = process_ranges(setColumnWidth, ws, ref_or_rng; kw...)
setColumnWidth(xl::XLSXFile, sheetcell::String; kw...)::Int = process_sheetcell(setColumnWidth, xl, sheetcell; kw...)
setColumnWidth(ws::Worksheet, cr::CellRef; kw...)::Int = setColumnWidth(ws, CellRange(cr, cr); kw...)
#setColumnWidth(ws::Worksheet, row::Integer, col::Integer; kw...) = setColumnWidth(ws, CellRef(row, col); kw...)
setColumnWidth(ws::Worksheet, row::Union{Integer,UnitRange{<:Integer}}, ::Colon; kw...) = process_colon(setColumnWidth, ws, row, nothing; kw...)
setColumnWidth(ws::Worksheet, col::Union{Integer,UnitRange{<:Integer}}; kw...) = process_colon(setColumnWidth, ws, nothing, col; kw...)  # convenience: bare column index/range
setColumnWidth(ws::Worksheet, ::Colon, col::Union{Integer,UnitRange{<:Integer}}; kw...) = process_colon(setColumnWidth, ws, nothing, col; kw...)
setColumnWidth(ws::Worksheet, ::Colon, ::Colon; kw...) = process_colon(setColumnWidth, ws, nothing, nothing; kw...)
setColumnWidth(ws::Worksheet, ::Colon; kw...) = process_colon(setColumnWidth, ws, nothing, nothing; kw...)
setColumnWidth(ws::Worksheet, row::Union{Vector{Int},StepRange{<:Integer}}, ::Colon; kw...) = process_veccolon(setColumnWidth, ws, row, nothing; kw...)
setColumnWidth(ws::Worksheet, col::Union{Vector{Int},StepRange{<:Integer}}; kw...) = process_veccolon(setColumnWidth, ws, nothing, col; kw...)  # convenience: bare column vector/steprange
setColumnWidth(ws::Worksheet, ::Colon, col::Union{Vector{Int},StepRange{<:Integer}}; kw...) = process_veccolon(setColumnWidth, ws, nothing, col; kw...)
setColumnWidth(ws::Worksheet, row::Union{Integer,UnitRange{<:Integer}}, col::Union{Vector{Int},StepRange{<:Integer}}; kw...) = process_vecint(setColumnWidth, ws, row, col; kw...)
setColumnWidth(ws::Worksheet, row::Union{Vector{Int},StepRange{<:Integer}}, col::Union{Integer,UnitRange{<:Integer}}; kw...) = process_vecint(setColumnWidth, ws, row, col; kw...)
setColumnWidth(ws::Worksheet, row::Union{Vector{Int},StepRange{<:Integer}}, col::Union{Vector{Int},StepRange{<:Integer}}; kw...) = process_vecint(setColumnWidth, ws, row, col; kw...)
setColumnWidth(ws::Worksheet, row::Union{Integer,UnitRange{<:Integer}}, col::Union{Integer,UnitRange{<:Integer}}; kw...) = setColumnWidth(ws, CellRange(CellRef(first(row), first(col)), CellRef(last(row), last(col))); kw...)
function setColumnWidth(ws::Worksheet, rng::CellRange; width::Union{Nothing,Real}=nothing)::Int
    get_xlsxfile(ws).is_writable ||
        throw(XLSXError("Cannot set column widths: `XLSXFile` is not writable."))

    # Validate and resolve width before any XML work
    isnothing(width) && return 0
    width < 0 && throw(XLSXError("Invalid width: $width. Must be >= 0."))

    wb=get_workbook(ws)
    padded_width = width + EXCEL_COLUMN_WIDTH_PADDING

    left  = rng.start.column_number
    right = rng.stop.column_number

    # Find the <cols> block in the worksheet XML, inserting one before <sheetData> if absent
    sheetdoc = xmlroot(wb, ws.relationship_id)
    pfx = get_prefix(ws)
    pfx = pfx == "" ? pfx : pfx * ":"

    i, j = get_idces(sheetdoc, "worksheet", "cols")

    if isnothing(j)
        k, l = get_idces(sheetdoc, "worksheet", "sheetData")
        i != k && throw(XLSXError("Unexpected worksheet structure: <cols> parent index does not match <worksheet> index."))
        insert!(sheetdoc[k].children, l, XML.Element("$(pfx)Cols"))
        j = l
    end

    # Build a map of existing column definitions keyed by min (column index string)
    col_defs = Dict{String,Dict{String,String}}()
    for c in XML.children(sheetdoc[i][j])
        col_defs[c["min"]] = Dict{String,String}(XML.attributes(c))
    end

    # Update or insert column definitions for each column in the range
    for col in left:right
        scol = string(col)
        if haskey(col_defs, scol)
            col_defs[scol]["width"]       = string(padded_width)
            col_defs[scol]["customWidth"] = "1"
        else
            col_defs[scol] = Dict(
                "min"         => scol,
                "max"         => scol,
                "width"       => string(padded_width),
                "customWidth" => "1"
            )
        end
    end

    # Rebuild the <cols> node from the updated definitions
    new_cols = unlink(sheetdoc[i][j], ("cols", "col"), pfx)
    for atts in values(col_defs)
        new_col = XML.Element("$(pfx)col")
        for (k, v) in atts
            new_col[k] = v
        end
        push!(new_cols, new_col)
    end

    sheetdoc[i][j] = new_cols

    # Working on worksheet XML directly requires updating the file cache
    update_worksheets_xml!(get_xlsxfile(ws))

    return 0
end

"""
    getColumnWidth(sh::Worksheet, cr::String) -> ::Union{Nothing, Real}
    getColumnWidth(xf::XLSXFile,  cr::String) -> ::Union{Nothing, Real}

    getColumnWidth(sh::Worksheet,  row::Int, col::Int) -> ::Union{Nothing, Real}

Get the width of a column defined by a cell reference or named cell.
The specified cell must be within the sheet dimension.

A standard cell reference or defined name may be used to define the column. 
The function will use the column number and ignore the row.

The function returns the value of the column width or nothing if the column 
does not have an explicitly defined width.

# Examples:
```julia
julia> XLSX.getColumnWidth(xf, "Sheet1!A2")

julia> XLSX.getColumnWidth(sh, "F1")

julia> XLSX.getColumnWidth(sh, 1, 6)
 
```
"""
function getColumnWidth end
getColumnWidth(xl::XLSXFile, sheetcell::String)::Union{Nothing,Float64} = process_get_sheetcell(getColumnWidth, xl, sheetcell)
getColumnWidth(ws::Worksheet, cr::String) = process_get_cellname(getColumnWidth, ws, cr)
getColumnWidth(ws::Worksheet, row::Integer, col::Integer) = getColumnWidth(ws, CellRef(row, col))
function getColumnWidth(ws::Worksheet, cellref::CellRef)::Union{Nothing,Real}
    # May be better if column width were part of ws.cache?

    if !get_xlsxfile(ws).is_writable
        throw(XLSXError("Cannot get column width: `XLSXFile` is not writable."))
    end

    d = get_dimension(ws)
    if cellref ∉ d
        throw(XLSXError("Cell specified is outside sheet dimension `$d`"))
    end

    sheetdoc = xmlroot(get_workbook(ws), ws.relationship_id)
    
    i, j = get_idces(sheetdoc, "worksheet", "cols")

    if isnothing(j) # There are no existing column formats defined.
        return nothing
    end
    for c in XML.children(sheetdoc[i][j])
        if c["min"] == string(cellref.column_number)
            if haskey(c, "width")
                return parse(Float64, c["width"])
            else
                break
            end
        end
    end

    # Either the column definition was found but with no width attribute or not found at all.
    return nothing
end

#
# -- Get and set row height
#

"""
    setRowHeight(sh::Worksheet, cr::String; kw...) -> ::Int
    setRowHeight(xf::XLSXFile,  cr::String, kw...) -> ::Int

    setRowHeight(sh::Worksheet, row, col; kw...) -> ::Int

Set the height of a row or row range.

A standard cell reference or cell range must be used to define the row range. 
The function will use the rows and ignore the columns. Named cells and named
ranges can similarly be used.
Alternatively, specify the row and column using any combination of 
Integer, UnitRange, Vector{Integer} or `:`, but only the rows will be used.

The function uses one keyword used to define a row height:
- `height::Real = nothing` : Defines height in Excel's own (internal) units.

When you set row heights interactively in Excel you can see the height 
in "internal" units and in pixels. The height stored in the xlsx file is slightly 
larger than the height shown interactively because Excel adds some cell padding. 
The method Excel uses to calculate the padding is obscure and complex. This 
function does not attempt to replicate it, but simply adds 0.21 internal units 
to the value specified. The value set is unlikely to match the value seen 
interactivley in the resultant spreadsheet, but it will be close.

Row height cannot be set for empty rows, which will quietly be skipped.
A row must have at least one cell containing a value before its height can be set.

You can set a row height to 0.

The function returns a value of 0 unless all rows are empty, in which case 
it returns a value of -1.

# Examples:
```julia
julia> XLSX.setRowHeight(xf, "Sheet1!A2"; height = 50)

julia> XLSX.setRowHeight(sh, "F1:F5"; height = 0)

julia> XLSX.setRowHeight(sh, "I"; height = 24.56)

```
"""
function setRowHeight end
setRowHeight(ws::Worksheet, ref::SheetCellRef; kw...) = do_sheet_names_match(ws, ref) && setRowHeight(ws, ref.cellref; kw...)
setRowHeight(ws::Worksheet, rng::SheetCellRange; kw...) = do_sheet_names_match(ws, rng) && setRowHeight(ws, rng.rng; kw...)
setRowHeight(ws::Worksheet, rng::SheetColumnRange; kw...) = do_sheet_names_match(ws, rng) && setRowHeight(ws, rng.colrng; kw...)
setRowHeight(ws::Worksheet, rng::SheetRowRange; kw...) = do_sheet_names_match(ws, rng) && setRowHeight(ws, rng.rowrng; kw...)
setRowHeight(ws::Worksheet, colrng::ColumnRange; kw...)::Int = process_columnranges(setRowHeight, ws, colrng; kw...)
setRowHeight(ws::Worksheet, rowrng::RowRange; kw...)::Int = process_rowranges(setRowHeight, ws, rowrng; kw...)
setRowHeight(ws::Worksheet, ncrng::NonContiguousRange; kw...)::Int = process_ncranges(setRowHeight, ws, ncrng; kw...)
setRowHeight(ws::Worksheet, ref_or_rng::AbstractString; kw...)::Int = process_ranges(setRowHeight, ws, ref_or_rng; kw...)
setRowHeight(xl::XLSXFile, sheetcell::String; kw...)::Int = process_sheetcell(setRowHeight, xl, sheetcell; kw...)
setRowHeight(ws::Worksheet, cr::CellRef; kw...)::Int = setRowHeight(ws::Worksheet, CellRange(cr, cr); kw...)
#setRowHeight(ws::Worksheet, row::Integer, col::Integer; kw...) = setRowHeight(ws, CellRef(row, col); kw...)
setRowHeight(ws::Worksheet, row::Union{Integer,UnitRange{<:Integer}}; kw...) = process_colon(setRowHeight, ws, row, nothing; kw...)
setRowHeight(ws::Worksheet, row::Union{Integer,UnitRange{<:Integer}}, ::Colon; kw...) = process_colon(setRowHeight, ws, row, nothing; kw...)
setRowHeight(ws::Worksheet, ::Colon, col::Union{Integer,UnitRange{<:Integer}}; kw...) = process_colon(setRowHeight, ws, nothing, col; kw...)
setRowHeight(ws::Worksheet, ::Colon, ::Colon; kw...) = process_colon(setRowHeight, ws, nothing, nothing; kw...)
setRowHeight(ws::Worksheet, ::Colon; kw...) = process_colon(setRowHeight, ws, nothing, nothing; kw...)
setRowHeight(ws::Worksheet, row::Union{Vector{Int},StepRange{<:Integer}}; kw...) = process_veccolon(setRowHeight, ws, row, nothing; kw...)
setRowHeight(ws::Worksheet, row::Union{Vector{Int},StepRange{<:Integer}}, ::Colon; kw...) = process_veccolon(setRowHeight, ws, row, nothing; kw...)
setRowHeight(ws::Worksheet, ::Colon, col::Union{Vector{Int},StepRange{<:Integer}}; kw...) = process_veccolon(setRowHeight, ws, nothing, col; kw...)
setRowHeight(ws::Worksheet, row::Union{Integer,UnitRange{<:Integer}}, col::Union{Vector{Int},StepRange{<:Integer}}; kw...) = process_vecint(setRowHeight, ws, row, col; kw...)
setRowHeight(ws::Worksheet, row::Union{Vector{Int},StepRange{<:Integer}}, col::Union{Integer,UnitRange{<:Integer}}; kw...) = process_vecint(setRowHeight, ws, row, col; kw...)
setRowHeight(ws::Worksheet, row::Union{Vector{Int},StepRange{<:Integer}}, col::Union{Vector{Int},StepRange{<:Integer}}; kw...) = process_vecint(setRowHeight, ws, row, col; kw...)
setRowHeight(ws::Worksheet, row::Union{Integer,UnitRange{<:Integer}}, col::Union{Integer,UnitRange{<:Integer}}; kw...) = setRowHeight(ws, CellRange(CellRef(first(row), first(col)), CellRef(last(row), last(col))); kw...)
function setRowHeight(ws::Worksheet, rng::CellRange; height::Union{Nothing,Real}=nothing)::Int

    if !get_xlsxfile(ws).is_writable
        throw(XLSXError("Cannot get row height: `XLSXFile` is not writable."))
    end

    top = rng.start.row_number
    bottom = rng.stop.row_number
    padded_height = isnothing(height) ? -1 : height + 0.2109375 # Excel adds cell padding to a user specified width
    if !isnothing(height) && height < 0
        throw(XLSXError("Invalid value specified for height: $height. Height must be >= 0."))
    end

    if isnothing(height) # No-op
        return 0
    end

    if isnothing(ws.cache) || !haskey(ws.cache.row_ht, rng.stop.row_number)
        _ = ws[rng.stop] # Ensure cache filled at least to include the last row in `rng`.
    end

    first = true

    for r in top:bottom
        if haskey(ws.cache.row_ht, r) # may still be missing if row is entirely empty.
            ws.cache.row_ht[r] = padded_height
            first=false
        end
    end

    if first == true
        return -1
    end

    return 0

end

"""
    getRowHeight(sh::Worksheet, cr::String) -> ::Union{Nothing, Real}
    getRowHeight(xf::XLSXFile,  cr::String) -> ::Union{Nothing, Real}

    getRowHeight(sh::Worksheet,  row::Int, col::Int) -> ::Union{Nothing, Real}

Get the height of a row defined by a cell reference or named cell.
The specified cell must be within the sheet dimension.

A standard cell reference or defined name must be used to define the row. 
The function will use the row number and ignore the column.

The function returns the value of the row height or nothing if the row 
does not have an explicitly defined height.

If the row is not found (an empty row), returns -1.

# Examples:
```julia
julia> XLSX.getRowHeight(xf, "Sheet1!A2")

julia> XLSX.getRowHeight(sh, "F1")

julia> XLSX.getRowHeight(sh, 1, 6)
 
```
"""
function getRowHeight end
getRowHeight(xl::XLSXFile, sheetcell::String)::Union{Nothing,Real} = process_get_sheetcell(getRowHeight, xl, sheetcell)
getRowHeight(ws::Worksheet, cr::String) = process_get_cellname(getRowHeight, ws, cr)
getRowHeight(ws::Worksheet, row::Integer, col::Integer) = getRowHeight(ws, CellRef(row, col))
function getRowHeight(ws::Worksheet, cellref::CellRef)::Union{Nothing,Real}

    if !get_xlsxfile(ws).is_writable
        throw(XLSXError("Cannot get row height: `XLSXFile` is not writable."))
    end

    d = get_dimension(ws)
    if cellref ∉ d
        throw(XLSXError("Cell specified is outside sheet dimension `$d`"))
    end

    if isnothing(ws.cache) || !haskey(ws.cache.row_ht, cellref.row_number)
        _ = ws[cellref] # Ensure cache filled at least to include the last row in `rng`.
    end

    if haskey(ws.cache.row_ht, cellref.row_number) # Row might still be empty
        return ws.cache.row_ht[cellref.row_number]
    end

    return -1 # Row specified not found (is empty)

end

#
# -- Get merged cells
#

"""
    getMergedCells(ws::Worksheet) -> Union{Vector{CellRange}, Nothing}

Return a vector of the `CellRange` of all merged cells in the specified worksheet.
Return nothing if the worksheet contains no merged cells.

The Excel file must be opened in write mode to work with merged cells.

# Examples:
```julia
julia> f = XLSX.readxlsx("test.xlsx")
XLSXFile("C:\\Users\\...\\test.xlsx") containing 1 Worksheet
            sheetname size          range
-------------------------------------------------
               Sheet1 2x2           A1:B2

julia> s = f["Sheet1"]
2×2 XLSX.Worksheet: ["Sheet1"](A1:B2)

julia> XLSX.getMergedCells(s)
1-element Vector{XLSX.CellRange}:
 B1:B2
 
```
"""
function getMergedCells(ws::Worksheet)::Union{Vector{CellRange},Nothing}
    # May be better if merged cells were part of ws.cache?

    if !get_xlsxfile(ws).is_writable
        throw(XLSXError("Cannot get merged cells: `XLSXFile` is not writable."))
    end

    if !get_xlsxfile(ws).use_cache_for_sheet_data
        throw(XLSXError("Cannot get merged cells because cache is not enabled."))
    end

    sheetdoc = xmlroot(get_workbook(ws), ws.relationship_id)
    i, j = get_idces(sheetdoc, "worksheet", "mergeCells")

    if isnothing(j) # There are no existing merged cells.
        return nothing
    end

    c = XML.children(sheetdoc[i][j])
    if length(c) != parse(Int, sheetdoc[i][j]["count"])
        throw(XLSXError("Unexpected number of mergeCells found: $(length(c)). Expected $(sheetdoc[i][j]["count"])."))
    end

    mergedCells = Vector{CellRange}()
    for cell in c
        !haskey(cell, "ref") && throw(XLSXError("No `ref` attribute found in `mergeCell` element."))
        push!(mergedCells, CellRange(cell["ref"]))
    end

    return mergedCells
end

"""
    isMergedCell(ws::Worksheet,  cr::String) -> Bool
    isMergedCell(xf::XLSXFile,   cr::String) -> Bool

    isMergedCell(ws::Worksheet,  row::Int, col::Int) -> Bool

Return `true` if a cell is part of a merged cell range and `false` if not.
The specified cell must be within the sheet dimension.

Alternatively, if you have already obtained the merged cells for the worksheet,
you can avoid repeated determinations and pass them as a keyword argument to 
the function:

    isMergedCell(ws::Worksheet, cr::String; mergedCells::Union{Vector{CellRange}, Nothing, Missing}=missing) -> Bool
    isMergedCell(xf::XLSXFile,  cr::String; mergedCells::Union{Vector{CellRange}, Nothing, Missing}=missing) -> Bool

    isMergedCell(ws::Worksheet,  row:Int, col::Int; mergedCells::Union{Vector{CellRange}, Nothing, Missing}=missing) -> Bool

The Excel file must be opened in write mode to work with merged cells.

# Examples:
```julia
julia> XLSX.isMergedCell(xf, "Sheet1!A1")

julia> XLSX.isMergedCell(sh, "A1")

julia> XLSX.isMergedCell(sh, 2, 4) # cell D2

julia> mc = XLSX.getMergedCells(sh)

julia> XLSX.isMergedCell(sh, XLSX.CellRef("A1"), mc)
 
```
"""
function isMergedCell end
isMergedCell(xl::XLSXFile, sheetcell::String; kw...)::Bool = process_get_sheetcell(isMergedCell, xl, sheetcell; kw...)
isMergedCell(ws::Worksheet, cr::String; kw...)::Bool = process_get_cellname(isMergedCell, ws, cr; kw...)
isMergedCell(ws::Worksheet, row::Integer, col::Integer; kw...) = isMergedCell(ws, CellRef(row, col); kw...)
function isMergedCell(ws::Worksheet, cellref::CellRef; mergedCells::Union{Vector{CellRange},Nothing,Missing}=missing)::Bool

    if !get_xlsxfile(ws).is_writable
        throw(XLSXError("Cannot get merged cells: `XLSXFile` is not writable."))
    end

    if !get_xlsxfile(ws).use_cache_for_sheet_data
        throw(XLSXError("Cannot get merged cells because cache is not enabled."))
    end

    d = get_dimension(ws)
    if cellref ∉ d
        throw(XLSXError("Cell specified is outside sheet dimension `$d`"))
    end

    if ismissing(mergedCells) # Get mergedCells if missing
        mergedCells = getMergedCells(ws)
    end
    if isnothing(mergedCells) # No merged cells in sheet
        return false
    end
    for rng in mergedCells
        if cellref ∈ rng
            return true
        end
    end

    return false
end

"""
    getMergedBaseCell(ws::Worksheet, cr::String) -> Union{Nothing, NamedTuple{CellRef, Any}}
    getMergedBaseCell(xf::XLSXFile,  cr::String) -> Union{Nothing, NamedTuple{CellRef, Any}}

    getMergedBaseCell(ws::Worksheet, row::Int, col::Int) -> Union{Nothing, NamedTuple{CellRef, Any}}

Return the cell reference and cell value of the base cell of a merged cell range in a worksheet as a named tuple.
The specified cell must be within the sheet dimension.
If the specified cell is not part of a merged cell range, return `nothing`.

The base cell is the top-left cell of the merged cell range and is the reference cell for the range.

The tuple returned contains:
- `baseCell`  : the reference (`CellRef`) of the base cell
- `baseValue` : the value of the base cell

Additionally, if you have already obtained the merged cells for the worksheet,
you can avoid repeated determinations and pass them as a keyword argument to 
the function:

    getMergedBaseCell(ws::Worksheet, cr::String; mergedCells::Union{Vector{CellRange}, Nothing, Missing}=missing) -> Union{Nothing, NamedTuple{CellRef, Any}}
    getMergedBaseCell(xf::XLSXFile,  cr::String; mergedCells::Union{Vector{CellRange}, Nothing, Missing}=missing) -> Union{Nothing, NamedTuple{CellRef, Any}}

    getMergedBaseCell(ws::Worksheet, row::Int, col::Int; mergedCells::Union{Vector{CellRange}, Nothing, Missing}=missing) -> Union{Nothing, NamedTuple{CellRef, Any}}

The Excel file must be opened in write mode to work with merged cells.

# Examples:
```julia
julia> XLSX.getMergedBaseCell(xf, "Sheet1!B2")
(baseCell = B1, baseValue = 3)

julia> XLSX.getMergedBaseCell(sh, "B2")
(baseCell = B1, baseValue = 3)

julia> XLSX.getMergedBaseCell(sh, 2, 2)
(baseCell = B1, baseValue = 3)

```
"""
function getMergedBaseCell end
getMergedBaseCell(xl::XLSXFile, sheetcell::String; kw...) = process_get_sheetcell(getMergedBaseCell, xl, sheetcell; kw...)
getMergedBaseCell(ws::Worksheet, cr::String; kw...) = process_get_cellname(getMergedBaseCell, ws, cr; kw...)
getMergedBaseCell(ws::Worksheet, row::Integer, col::Integer; kw...) = getMergedBaseCell(ws, CellRef(row, col); kw...)
function getMergedBaseCell(ws::Worksheet, cellref::CellRef; mergedCells::Union{Vector{CellRange},Nothing,Missing}=missing)

    if !get_xlsxfile(ws).is_writable
        throw(XLSXError("Cannot get merged cells: `XLSXFile` is not writable."))
    end

    if !get_xlsxfile(ws).use_cache_for_sheet_data
        throw(XLSXError("Cannot get merged cells because cache is not enabled."))
    end

    d = get_dimension(ws)
    if cellref ∉ d
        throw(XLSXError("Cell specified is outside sheet dimension `$d`"))
    end

    if ismissing(mergedCells) # Get mergedCells if missing
        mergedCells = getMergedCells(ws)
    end
    if isnothing(mergedCells) # No merged cells in sheet
        return nothing
    end
    for rng in mergedCells
        if cellref ∈ rng
            return (; baseCell=rng.start, baseValue=ws[rng.start])
        end
    end
    return nothing
end

"""
    mergeCells(ws::Worksheet, cr::String) -> 0
    mergeCells(xf::XLSXFile,  cr::String) -> 0

    mergeCells(ws::Worksheet, row::Int, col::Int) -> 0

Merge the cells in the range given by `cr`. The value of the merged cell 
will be the value of the first cell in the range (the base cell) prior 
to the merge. All other cells in the range will be set to `missing`, 
reflecting the behaviour of Excel itself.

Merging is limited to the extent of the worksheet dimension.

The specified range must not overlap with any previously merged cells.

It is not possible to merge a single cell!

A non-contiguous range composed of multiple cell ranges will be processed as a 
list of separate ranges. Each range will be merged separately. No range within 
a non-contiguous range may be a single cell.

The Excel file must be opened in write mode to work with merged cells.

# Examples:
```julia
julia> XLSX.mergeCells(xf, "Sheet1!B2:D3")  # Merge a cell range.

julia> XLSX.mergeCells(sh, 1:3, :)          # Merge rows to the extent of the dimension.

julia> XLSX.mergeCells(sh, "A:D")           # Merge columns to the extent of the dimension.

```
"""
function mergeCells end
mergeCells(ws::Worksheet, rng::SheetCellRange) = do_sheet_names_match(ws, rng) && mergeCells(ws, rng.rng)
mergeCells(ws::Worksheet, rng::SheetColumnRange) = do_sheet_names_match(ws, rng) && mergeCells(ws, rng.colrng)
mergeCells(ws::Worksheet, rng::SheetRowRange) = do_sheet_names_match(ws, rng) && mergeCells(ws, rng.rowrng)
mergeCells(ws::Worksheet, colrng::ColumnRange)::Int = process_columnranges(mergeCells, ws, colrng)
mergeCells(ws::Worksheet, rowrng::RowRange)::Int = process_rowranges(mergeCells, ws, rowrng)
mergeCells(ws::Worksheet, ncrng::NonContiguousRange; kw...)::Int = process_ncranges(mergeCells, ws, ncrng; kw...)
mergeCells(xl::XLSXFile, sheetcell::AbstractString)::Int = process_sheetcell(mergeCells, xl, sheetcell)
mergeCells(ws::Worksheet, ref_or_rng::AbstractString)::Int = process_ranges(mergeCells, ws, ref_or_rng)
mergeCells(ws::Worksheet, row::Union{Integer,UnitRange{<:Integer}}, ::Colon) = process_colon(mergeCells, ws, row, nothing)
mergeCells(ws::Worksheet, ::Colon, col::Union{Integer,UnitRange{<:Integer}}) = process_colon(mergeCells, ws, nothing, col)
mergeCells(ws::Worksheet, ::Colon, ::Colon) = process_colon(mergeCells, ws, nothing, nothing)
mergeCells(ws::Worksheet, ::Colon) = process_colon(mergeCells, ws, nothing, nothing)
mergeCells(ws::Worksheet, row::Union{Integer,UnitRange{<:Integer}}, col::Union{Integer,UnitRange{<:Integer}}) = mergeCells(ws, CellRange(CellRef(first(row), first(col)), CellRef(last(row), last(col))))
function mergeCells(ws::Worksheet, cr::CellRange)
    # May be better if merged cells were part of ws.cache?

    #    !is_valid_cell_range(cr) && throw(XLSXError("\"$cr\" is not a valid cell range."))

    !issubset(cr, get_dimension(ws)) && throw(XLSXError("Range `$cr` goes outside worksheet dimension."))

    length(cr) == 1 && throw(XLSXError("Cannot merge a single cell: `$cr`"))

    if !get_xlsxfile(ws).is_writable
        throw(XLSXError("Cannot merge cells: `XLSXFile` is not writable."))
    end

    if !get_xlsxfile(ws).use_cache_for_sheet_data
        throw(XLSXError("Cannot get merged cells because cache is not enabled."))
    end

    sheetdoc = xmlroot(get_workbook(ws), ws.relationship_id)
    pfx = get_prefix(ws)
    pfx = pfx == "" ? pfx : pfx * ":"

    l = insert_index(sheetdoc[end], "mergeCells", WORKSHEET_ORDER)

    if localname(sheetdoc[end][l]) != "mergeCells" # There are no existing merged cells. Insert immediately after the <sheetData> block and push everything else down one.
        len = length(sheetdoc[end])
        merge_node=XML.Node(XML.Element, "$(pfx)mergeCells", OrderedDict{String,String}(), nothing, Vector{XML.Node}())
        if l != len
            insert!(sheetdoc[end].children, l+1, merge_node)
        else
            push!(sheetdoc[end], merge_node)
        end
        l += 1
        count = 0
    else # There are already some existing merged cells
        c = XML.children(sheetdoc[end][l])
        count = length(c)
        if count != parse(Int, sheetdoc[end][l]["count"])
            throw(XLSXError("Unexpected number of mergeCells found: $(length(c)). Expected $(sheetdoc[i][j]["count"])."))
        end
        for child in c
            if intersects(cr, CellRange(child["ref"]))
                throw(XLSXError("Merged range (`$cr`) cannot overlap with existing merged range (`" * child["ref"] * "`)."))
            end
        end
    end

    push!(sheetdoc[end][l], XML.Element("$(pfx)mergeCell", ref=string(cr))) # Add the new merged cell range.
    count += 1
    sheetdoc[end][l]["count"] = count

    # All cells except the base cell are set to missing.
    let first = true
        for cell in cr
            if first
                first = false
                continue
            else
                ws[cell] = ""
            end
        end
    end

    update_worksheets_xml!(get_xlsxfile(ws))

    return 0 # meaningless return value. Int required to comply with reference decoding structure.
end
