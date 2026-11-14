#
# -- Get and set border attributes
#

"""
    getBorder(sh::Worksheet, cr::String) -> ::Union{Nothing, CellBorder}
    getBorder(xf::XLSXFile, cr::String)  -> ::Union{Nothing, CellBorder}

    getBorder(sh::Worksheet, row::Int, col::Int) -> ::Union{Nothing, CellBorder}
   
Get the borders used by a single cell at reference in a worksheet or XLSXFile.
A cell outside the sheet dimension is empty, so this returns `nothing`, as for any empty cell.

Return a `CellBorder` object containing:
- `borderId`    : a 0-based index of the border in the workbook
- `border`      : a dictionary of border attributes: borderAttribute -> (attribute -> value)
- `applyBorder` : "1" or "0", indicating whether or not the border is applied to the cell.

Return `nothing` if no cell border is found.

A cell border has two attributes, `style` and `color`. A diagonal border also needs to specify 
a direction indicating whether a diagonal is needed from bottom-left to top-right ("up"), 
from top-left to bottom-left ("down") or "both"

The available `style`s are:
$(_doc_list(VALID_BORDER_STYLES))

A border position element (e.g. `top` or `left`) has a `style` attribute, but `color` is a child element.
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
getBorder(xl::XLSXFile, sheetcell::AbstractString)::Union{Nothing,CellBorder} = process_get_sheetcell(getBorder, xl, sheetcell)
getBorder(ws::Worksheet, cellref::CellRef)::Union{Nothing,CellBorder} = process_get_cellref(getBorder, ws, cellref)
getBorder(ws::Worksheet, cr::AbstractString) = process_get_cellname(getBorder, ws, cr)
getBorder(ws::Worksheet, row::Integer, col::Integer) = getBorder(ws, CellRef(row, col))
getDefaultBorders(ws::Worksheet) = getBorder(get_workbook(ws), styles_cell_xf(get_workbook(ws), 0))
function getBorder(wb::Workbook, cell_style::XML.Node)::Union{Nothing,CellBorder}

    _require_cache(wb, "get border")

    if haskey(cell_style, "borderId")
        borderid = cell_style["borderId"]
        applyborder = haskey(cell_style, "applyBorder") ? cell_style["applyBorder"] : "0"
        border_nodes = get_borders_nodes(wb)
        current_border = border_nodes[parse(Int, borderid)+1] # Zero based!
        diag_atts = XML.attributes(current_border)
        border_atts = Dict{String,Union{Dict{String,String},Nothing}}()
        for side in xml_elements(current_border)
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
                    for subc in xml_elements(side) # color is a child of a border element
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
- `left::Vector{<:Pair{String}} = nothing`
- `right::Vector{<:Pair{String}} = nothing`
- `top::Vector{<:Pair{String}} = nothing`
- `bottom::Vector{<:Pair{String}} = nothing`
- `diagonal::Vector{<:Pair{String}} = nothing`
- `[allsides::Vector{<:Pair{String}} = nothing]`
- `[outside::Vector{<:Pair{String}} = nothing]`

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
$(_doc_list(VALID_BORDER_STYLES))

The `color` attribute can be $COLOR_FORMS_DOC

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

For cell ranges, column ranges and named ranges, the value returned is -1 unless 
the range resolves to a single cell, in which case that cell's id is returned.

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
define_percell_spec_methods(:setBorder, :_setBorder, :BorderSpec; ranges=false)

# The range forms take `outside`, which is set around the range as a whole.
function _setBorder(ws::Worksheet, rng::CellRange; spec::BorderSpec)::Int
    _check_outside_conflict(spec)
    isnothing(spec.outside) || return setOutsideBorder(ws, rng; spec.outside)
    return process_cellranges(_setBorder, ws, rng; spec)
end

function _setBorder(ws::Worksheet, rng::ColumnRange; spec::BorderSpec)::Int
    _check_outside_conflict(spec)
    isnothing(spec.outside) || return process_columnranges(setOutsideBorder, ws, rng; spec.outside)
    return process_columnranges(_setBorder, ws, rng; spec)
end

function _setBorder(ws::Worksheet, rng::RowRange; spec::BorderSpec)::Int
    _check_outside_conflict(spec)
    isnothing(spec.outside) || return process_rowranges(setOutsideBorder, ws, rng; spec.outside)
    return process_rowranges(_setBorder, ws, rng; spec)
end

function _setBorder(xl::XLSXFile, sheetcell::AbstractString; spec::BorderSpec)::Int
    _check_outside_conflict(spec)
    if isnothing(spec.outside)
        return process_sheetcell(_setBorder, xl, sheetcell; spec)
    else
        return process_sheetcell(setOutsideBorder, xl, sheetcell; spec.outside)
    end
end

function _setBorder(sh::Worksheet, cellref::CellRef; spec::BorderSpec)::Int
    (; outside, allsides, left, right, top, bottom, diagonal) = spec
    _require_writable(sh, "set borders")

    isnothing(outside) ||
        throw(XLSXError("Keyword `outside` sets the border around a range, so cannot be used for the single cell $(cellname(cellref)) or with `setUniformBorder`."))

    # allsides expands to all four sides, then recurses — early return
    if !isnothing(allsides)
        !all(isnothing, [left, right, top, bottom]) &&
            throw(XLSXError("Keyword `allsides` is incompatible with `left`, `right`, `top`, or `bottom`."))
        return _setBorder(sh, cellref; spec=BorderSpec(nothing, nothing, allsides, allsides, allsides, allsides, diagonal))
    end

    wb   = get_workbook(sh)
    pfx = _styles_prefix(sh)

    cell, allXfNodes = _format_target(sh, cellref, "border")

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
define_uniform_spec_methods(:setUniformBorder, :_setUniformBorder, :_setBorder, :BorderSpec, ["borderId", "applyBorder"])

"""
    setOutsideBorder(sh::Worksheet, cr::String; kw...) -> ::Int
    setOutsideBorder(xf::XLSXFile,  cr::String, kw...) -> ::Int

    setOutsideBorder(sh::Worksheet, rows, cols; kw...) -> ::Int

Set the border around the outside of a cell range, a column range or row range 
or a named range in a worksheet or XLSXFile.
Alternatively, specify the rows and columns using integers, UnitRanges or `:`.

There is one key word:
- `outside::Union{Nothing,Vector{<:Pair{String}}}=nothing`

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
    outside::BorderKw=nothing
)::Int
    _require_writable(ws, "set borders")

    topLeft     = rng.start
    topRight    = CellRef(rng.start.row_number, rng.stop.column_number)
    bottomLeft  = CellRef(rng.stop.row_number,  rng.start.column_number)
    bottomRight = rng.stop

    _setBorder(ws, CellRange(topLeft,    topRight);    spec=BorderSpec(:setOutsideBorder; top=outside))
    _setBorder(ws, CellRange(topLeft,    bottomLeft);  spec=BorderSpec(:setOutsideBorder; left=outside))
    _setBorder(ws, CellRange(topRight,   bottomRight); spec=BorderSpec(:setOutsideBorder; right=outside))
    _setBorder(ws, CellRange(bottomLeft, bottomRight); spec=BorderSpec(:setOutsideBorder; bottom=outside))

    return -1

end
