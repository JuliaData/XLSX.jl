#
# -- Cell protection: locked and hidden
#
# Part of a cell's format: the `<protection locked hidden>` child of its cellXfs
# `<xf>`. Every cell is locked unless its format says otherwise, but locking only
# takes effect on a protected sheet (see `protectSheet`).
#

"""
    setCellProtection(ws::Worksheet, ref; locked=nothing, hidden=nothing)
    setCellProtection(xf::XLSXFile, "Sheet1!B2:D10"; locked=nothing, hidden=nothing)
    setCellProtection(ws::Worksheet, row, col; locked=nothing, hidden=nothing)

Set whether cells are locked, and whether their formulas are hidden, as the
*Protection* tab of Excel's *Format Cells* dialog does. These take effect only
when the sheet is protected with [`protectSheet`](@ref).

- `locked`: whether the cells can be changed on a protected sheet. Every cell is
  locked unless its format says otherwise, so pass `locked=false` to let users
  type into some cells of a protected sheet.
- `hidden`: whether a cell's formula is hidden from the formula bar on a
  protected sheet. Its value still shows.

Only the keywords given are changed, and the cells' other formatting is kept.

`ref` is a cell (`"B2"`), a range (`"B2:D10"`), several of these separated by
commas, a defined name, a `CellRef` or a `CellRange`. Empty cells are included:
an empty cell is given the format without a value, so that it can be filled in
later on the protected sheet.

The file must be open for writing (`mode="rw"`).

# Examples
```julia
julia> XLSX.setCellProtection(sheet, "B2:B20"; locked=false)   # input cells

julia> XLSX.setCellProtection(sheet, "D2:D20"; hidden=true)    # hide the formulas

julia> XLSX.protectSheet(sheet)
```

See also [`getCellProtection`](@ref).
"""
function setCellProtection end
setCellProtection(xf::XLSXFile, ref::AbstractString; kw...) = setCellProtection(_rowcol_sheet(xf, ref)...; kw...)
setCellProtection(ws::Worksheet, row::Integer, col::Integer; kw...) = setCellProtection(ws, CellRef(row, col); kw...)
setCellProtection(ws::Worksheet, ref::CellRef; kw...) = setCellProtection(ws, CellRange(ref, ref); kw...)
setCellProtection(ws::Worksheet, rng::CellRange; kw...) = _set_cell_protection!(ws, Any[rng]; kw...)
setCellProtection(ws::Worksheet, ref::AbstractString; kw...) = _set_cell_protection!(ws, _rowcol_parts(ws, ref); kw...)

function _set_cell_protection!(ws::Worksheet, parts::Vector; locked::Union{Nothing,Bool}=nothing, hidden::Union{Nothing,Bool}=nothing)
    _require_writable(ws, "set cell protection")
    _require_worksheet(ws, "set cell protection")
    _require_cache(ws, "set cell protection")
    for p in parts
        p isa Union{CellRef,CellRange} ||
            throw(XLSXError("Give the cells to protect, such as `B2:D10`. Found `$p`, which is a whole row or column range."))
    end
    isnothing(locked) && isnothing(hidden) && return nothing

    wb = get_workbook(ws)
    allXfNodes = get_cellXfs_nodes(wb)
    new_style = Dict{Int,Int}()   # each style used, and its protected version
    for p in parts
        rng = p isa CellRef ? CellRange(p, p) : p
        for r in row_number(rng.start):row_number(rng.stop), c in column_number(rng.start):column_number(rng.stop)
            ref = CellRef(r, c)
            cell = getcell(ws, ref)
            if cell isa EmptyCell   # a cell with a format and no value
                ws[ref] = missing
                cell = getcell(ws, ref)
            end
            cell.style = get!(new_style, Int(cell.style)) do
                _protected_style(ws, allXfNodes, Int(cell.style), locked, hidden)
            end
        end
    end
    return nothing
end

function _protected_style(ws::Worksheet, allXfNodes::Vector{XML.Node}, style::Int,
                          locked::Union{Nothing,Bool}, hidden::Union{Nothing,Bool})::Int
    style == 0 && (style = Int(get_num_style_index(ws, allXfNodes, 0).id))
    old = _xf_protection(styles_cell_xf(allXfNodes, style))
    l = something(locked, old.locked)
    h = something(hidden, old.hidden)
    node = XML.Node{String}(XML.Element, "protection", ["locked" => (l ? "1" : "0"), "hidden" => (h ? "1" : "0")], nothing, nothing)
    return Int(update_template_xf(ws, allXfNodes, CellDataFormat(style), node; apply="applyProtection").id)
end

# The protection of a cellXfs `<xf>`: locked and not hidden unless it says otherwise.
function _xf_protection(xf::XML.Node)
    for c in xml_elements(xf)
        localname(c) == "protection" || continue
        a = Dict{String,String}(something(XML.attributes(c), Dict{String,String}()))
        return (locked = _xml_bool(a, "locked", true), hidden = _xml_bool(a, "hidden", false))
    end
    return (locked = true, hidden = false)
end

"""
    getCellProtection(ws::Worksheet, ref) -> NamedTuple
    getCellProtection(xf::XLSXFile, "Sheet1!B2") -> NamedTuple
    getCellProtection(ws::Worksheet, row, col) -> NamedTuple

Whether a cell is locked, and whether its formula is hidden, on a protected
sheet: a `NamedTuple` `(locked, hidden)` of `Bool`s. A cell is locked and not
hidden unless its format says otherwise. `ref` is a single cell.

See also [`setCellProtection`](@ref).
"""
function getCellProtection end
getCellProtection(xf::XLSXFile, ref::AbstractString) = getCellProtection(_rowcol_sheet(xf, ref)...)
getCellProtection(ws::Worksheet, row::Integer, col::Integer) = getCellProtection(ws, CellRef(row, col))
function getCellProtection(ws::Worksheet, ref::AbstractString)
    parts = _rowcol_parts(ws, ref)
    (length(parts) == 1 && only(parts) isa CellRef) ||
        throw(XLSXError("Give a single cell, such as `B2`. Found `$ref`."))
    return getCellProtection(ws, only(parts))
end
function getCellProtection(ws::Worksheet, ref::CellRef)
    _require_worksheet(ws, "get cell protection")
    cell = getcell(ws, ref)
    (cell isa EmptyCell || cell.style == 0) && return (locked = true, hidden = false)
    return _xf_protection(styles_cell_xf(get_workbook(ws), Int(cell.style)))
end
