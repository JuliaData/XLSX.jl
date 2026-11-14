#
# -- Get and set fill attributes
#

"""
    getFill(sh::Worksheet, cr::String) -> ::Union{Nothing, CellFill}
    getFill(xf::XLSXFile, cr::String)  -> ::Union{Nothing, CellFill}

    getFill(sh::Worksheet, row::Int, col::Int) -> ::Union{Nothing, CellFill}
   
Get the fill used by a single cell at reference `cr` in a worksheet or XLSXFile.
A cell outside the sheet dimension is empty, so this returns `nothing`, as for any empty cell.

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

The `patternType`s Excel uses are:
$(_doc_list(VALID_FILL_PATTERNS))

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
getFill(xl::XLSXFile, sheetcell::AbstractString)::Union{Nothing,CellFill} = process_get_sheetcell(getFill, xl, sheetcell)
getFill(ws::Worksheet, cellref::CellRef)::Union{Nothing,CellFill} = process_get_cellref(getFill, ws, cellref)
getFill(ws::Worksheet, cr::AbstractString) = process_get_cellname(getFill, ws, cr)
getFill(ws::Worksheet, row::Integer, col::Integer) = getFill(ws, CellRef(row, col))
getDefaultFill(ws::Worksheet) = getFill(get_workbook(ws), styles_cell_xf(get_workbook(ws), 0))

function getFill(wb::Workbook, cell_style::XML.Node)::Union{Nothing,CellFill}
    _require_cache(wb, "get fill")

    haskey(cell_style, "fillId") || return nothing

    fillid_int  = parse(Int, cell_style["fillId"])
    applyfill = haskey(cell_style, "applyFill") ? cell_style["applyFill"] : "0"
    fill_nodes = get_fills_nodes(wb)
    current_fill = fill_nodes[fillid_int + 1]  # fillId is zero-based
    fill_atts    = Dict{String,Union{Dict{String,String},Nothing}}()

    for pattern in xml_elements(current_fill)
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
- `fgColor = nothing`           : Sets the foreground color for the fill.
- `bgColor = nothing`           : Sets the background color for the fill.

The available `pattern` values are:
$(_doc_list(VALID_FILL_PATTERNS))

The two colors may each be $COLOR_FORMS_DOC

Setting only one or two of the attributes leaves the other attribute(s) unchanged 
for that cell's fill.

Fill attributes cannot be set for `EmptyCell`s. Set a cell value first.
If a cell range or column range includes any `EmptyCell`s, they will be
quietly skipped and the fill will be set for the remaining cells.

For single cells, the value returned is the `fillId` of the fill applied to the cell.
This can be used to apply the same fill to other cells or ranges.

For cell ranges, column ranges and named ranges, the value returned is -1 unless 
the range resolves to a single cell, in which case that cell's id is returned.

# Examples:
```julia
Julia> setFill(sh, "B2"; pattern="gray125", bgColor = "FF000000")

Julia> setFill(xf, "Sheet1!A1:F20"; pattern="none", fgColor = "darkseagreen3")
 
Julia> setFill(sh, "11:24"; pattern="none", fgColor = "yellow2")
 
```
"""
function setFill end
define_percell_spec_methods(:setFill, :_setFill, :FillSpec)

function _setFill(sh::Worksheet, cellref::CellRef; spec::FillSpec)::Int
    (; pattern, fgColor, bgColor) = spec
    _require_writable(sh, "set fill")

    # Validate before any work
    !isnothing(pattern) && pattern ∉ VALID_FILL_PATTERNS &&
        throw(XLSXError("Invalid pattern: $pattern. Must be one of: $(join(VALID_FILL_PATTERNS, ", "))."))

    wb   = get_workbook(sh)
    pfx = _styles_prefix(sh)

    cell, allXfNodes = _format_target(sh, cellref, "fill")

    cell_style   = styles_cell_xf(allXfNodes, cell.style)
    cell_fill    = getFill(wb, cell_style)
    old_fill_atts = cell_fill.fill["patternFill"]

    patternFill = Dict{String,String}()

    # patternType: preserve old or apply new
    result = _merge_val_tag("patternType", pattern, old_fill_atts)
    if !ismissing(result)
        patternFill["patternType"] = result isa Dict ? result["val"] : result
    end

    # fg/bg colors: prefix-match to preserve all old fg*/bg* keys, or set the new color
    for (prefix, new_color) in (("fg", fgColor), ("bg", bgColor))
        if isnothing(new_color)
            for (k, v) in old_fill_atts
                startswith(k, prefix) && (patternFill[k] = v)
            end
        else
            for (k, v) in _color_atts(new_color)
                patternFill[prefix * k] = v
            end
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
define_uniform_spec_methods(:setUniformFill, :_setUniformFill, :_setFill, :FillSpec, ["fillId", "applyFill"])
