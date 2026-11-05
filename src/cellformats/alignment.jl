#
# -- Get and set alignment attributes
#

"""
    getAlignment(sh::Worksheet, cr::String) -> ::Union{Nothing, CellAlignment}
    getAlignment(xf::XLSXFile,  cr::String) -> ::Union{Nothing, CellAlignment}

    getAlignment(sh::Worksheet, row::Int, col::Int) -> ::Union{Nothing, CellAlignment}
   
Get the alignment used by a single cell at reference `cr` in a worksheet or XLSXFile.
A cell outside the sheet dimension is empty, so this returns `nothing`, as for any empty cell.

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
    _require_cache(wb, "get alignment")

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

For cell ranges, column ranges and named ranges, the value returned is -1 unless 
the range resolves to a single cell, in which case that cell's id is returned.

# Examples:
```julia
julia> setAlignment(sh, "D18"; horizontal="center", wrapText=true)

julia> setAlignment(xf, "sheet1!D18"; horizontal="right", vertical="top", wrapText=true)

julia> setAlignment(sh, "L6"; horizontal="center", rotation="90", shrink=true, indent="2")

julia> setAlignment(sh, 1:3, 3:6; horizontal="center", rotation="90", shrink=true, indent="2")
 
```
"""
function setAlignment end
define_percell_spec_methods(:setAlignment, :_setAlignment, :AlignmentSpec)

function _setAlignment(sh::Worksheet, cellref::CellRef; spec::AlignmentSpec)::Int
    (; horizontal, vertical, wrapText, shrink, indent, rotation) = spec
    _require_writable(sh, "set alignment")

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
    cell, allXfNodes = _format_target(sh, cellref, "alignment")

    cell_style     = styles_cell_xf(allXfNodes, cell.style)
    cell_alignment = getAlignment(wb, cell_style)
    old_atts       = isnothing(cell_alignment) ? Dict{String,String}() : cell_alignment.alignment["alignment"]

    atts = Dict{String,String}()
    _merge_alignment_att(atts, "horizontal",   horizontal, old_atts)
    _merge_alignment_att(atts, "vertical",     vertical,   old_atts)
    _merge_alignment_att(atts, "wrapText",     wrapText,   old_atts, b -> b ? "1" : "0")
    _merge_alignment_att(atts, "shrinkToFit",  shrink,     old_atts, b -> b ? "1" : "0")
    _merge_alignment_att(atts, "indent",       indent,     old_atts)
    _merge_alignment_att(atts, "textRotation", rotation,   old_atts)

    alignment_node = XML.Node{String}(XML.Element, "alignment", Pair{String,String}[k => v for (k,v) in atts], nothing, nothing)
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
define_uniform_spec_methods(:setUniformAlignment, :_setUniformAlignment, :_setAlignment, :AlignmentSpec, nothing)
