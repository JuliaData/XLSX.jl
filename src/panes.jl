#=
panes.jl — freezePanes / splitFreeze / splitPanes / removePanes for XLSX.jl
=#

@enum PaneKind FROZEN FROZEN_SPLIT SPLIT

#-----------------------------------------------------------------------------
# Column-width / row-height -> twips (only exercised by plain `splitPanes`)
#-----------------------------------------------------------------------------

const DEFAULT_MDW = 7              # Calibri 11 / Arial 10 @ 96dpi -- approximation,
                                   # see caveat in splitPanes docstring below.
const DEFAULT_COL_WIDTH = 8.43     # Excel's absolute fallback (character units)
const DEFAULT_ROW_HEIGHT = 15.0    # points, Excel's absolute fallback

# OOXML native column "width" (character units) -> pixels, per ECMA-376 §18.3.1.13
_charwidth_to_pixels(width::Real, mdw::Int=DEFAULT_MDW) =
    Int(floor(((256 * width + floor(128 / mdw)) / 256) * mdw))

_pixels_to_twips(px::Real) = round(Int, px * 15)    # 96dpi: 1 inch = 96px = 1440 twips
_points_to_twips(pts::Real) = round(Int, pts * 20)  # 1 point = 20 twips, always

function _sheet_format_default(ws::Worksheet, attr::String)::Union{Nothing,Float64}
    sheetdoc = xmlroot(get_workbook(ws), ws.relationship_id)
    i, j = get_idces(sheetdoc, "worksheet", "sheetFormatPr")
    isnothing(j) && return nothing
    node = sheetdoc[i][j]
    haskey(node, attr) || return nothing
    return parse(Float64, node[attr])
end

function _col_width_chars(ws::Worksheet, col::Int)::Float64
    w = getColumnWidth(ws, CellRef(1, col))   # nothing if no explicit width for this column
    !isnothing(w) && return Float64(w)
    d = _sheet_format_default(ws, "defaultColWidth")
    !isnothing(d) && return d
    return DEFAULT_COL_WIDTH
end

function _row_height_pts(ws::Worksheet, row::Int)::Float64
    h = getRowHeight(ws, CellRef(row, 1))     # nothing (unset) or -1 (empty row)
    !isnothing(h) && h >= 0 && return Float64(h)
    d = _sheet_format_default(ws, "defaultRowHeight")
    !isnothing(d) && return d
    return DEFAULT_ROW_HEIGHT
end

_cols_to_twips(ws::Worksheet, ncols::Int)::Int =
    ncols <= 0 ? 0 : sum(_pixels_to_twips(_charwidth_to_pixels(_col_width_chars(ws, c))) for c in 1:ncols)

_rows_to_twips(ws::Worksheet, nrows::Int)::Int =
    nrows <= 0 ? 0 : sum(_points_to_twips(_row_height_pts(ws, r)) for r in 1:nrows)

#-----------------------------------------------------------------------------
# Selection-cell placement
#-----------------------------------------------------------------------------

# Which pane names get a <selection>, and which cell each should show.
# corner_selectable=true only for plain (non-frozen) splits -- the top-left
# pane is a real interactive pane there; for frozen/frozenSplit it's locked
# and Excel doesn't give it a <selection> at all.
function _selection_cells(nrows::Int, ncols::Int, topLeftCell::CellRef; corner_selectable::Bool)
    sels = Dict{String,CellRef}()
    if ncols > 0 && nrows > 0
        sels["topRight"]    = CellRef(1, topLeftCell.column_number)
        sels["bottomLeft"]  = CellRef(topLeftCell.row_number, 1)
        sels["bottomRight"] = topLeftCell
        corner_selectable && (sels["topLeft"] = CellRef(1, 1))
    elseif ncols > 0
        sels["topRight"] = topLeftCell
        corner_selectable && (sels["topLeft"] = CellRef(1, 1))
    elseif nrows > 0
        sels["bottomLeft"] = topLeftCell
        corner_selectable && (sels["topLeft"] = CellRef(1, 1))
    end
    return sels
end

#-----------------------------------------------------------------------------
# Core implementation shared by all four public functions
#-----------------------------------------------------------------------------

function _apply_pane!(ws::Worksheet, kind::PaneKind, nrows::Int, ncols::Int)

    if !get_xlsxfile(ws).is_writable
        throw(XLSXError("Cannot apply pane: `XLSXFile` is not writable."))
    end

    (nrows < 0 || ncols < 0) && throw(XLSXError("nrows and ncols must be non-negative."))

    doc   = get_worksheet_xml_document(ws)
    pfx   = get_prefix(ws)
    pfx_c = isempty(pfx) ? "" : "$(pfx):"

    wroot_idx, sv_idx = get_idces(doc, "worksheet", "sheetViews")
    wroot_idx === nothing && throw(XLSXError("Malformed worksheet: no <worksheet> root element found."))
    xroot = doc[wroot_idx]
    isnothing(xroot.children) && throw(XLSXError("Malformed worksheet: <worksheet> element has no children."))
    root_children = XML.children(xroot)

    # Preserve any existing <sheetView>'s own attributes (tabSelected, zoomScale,
    # showGridLines, etc.) and locate where it sits so we can replace it in
    # place rather than mutate it.
    existing_attrs = Pair{String,String}[]
    existing_sv_idx = nothing
    sheetViews_children = nothing

    if sv_idx !== nothing
        sheetViews_node = root_children[sv_idx]
        sheetViews_children = sheetViews_node.children  # may be `nothing` if self-closed <sheetViews/>
        if !isnothing(sheetViews_children)
            existing_sv_idx = find_child_index(sheetViews_children, "sheetView")
            if existing_sv_idx !== nothing
                existing_sheetview = sheetViews_children[existing_sv_idx]
                existing_attrs = isnothing(existing_sheetview.attributes) ? Pair{String,String}[] : copy(existing_sheetview.attributes)
            end
        end
    end

    attrs_dict = Dict(existing_attrs)
    haskey(attrs_dict, "workbookViewId") || (attrs_dict["workbookViewId"] = "0")
    sheetview_kwargs = (Symbol(k) => v for (k, v) in attrs_dict)

    # Build the new pane/selection children (empty for removePanes).
    new_children = XML.Node{String}[]
    if !(nrows == 0 && ncols == 0)
        topLeftCell = CellRef(nrows + 1, ncols + 1)
        activePane  = nrows > 0 && ncols > 0 ? "bottomRight" :
                      ncols > 0              ? "topRight"    : "bottomLeft"
        xSplit, ySplit = kind == SPLIT ? (_cols_to_twips(ws, ncols), _rows_to_twips(ws, nrows)) : (ncols, nrows)

        pane_kw = Dict{Symbol,Any}(:activePane => activePane, :topLeftCell => string(topLeftCell))
        ncols > 0 && (pane_kw[:xSplit] = string(xSplit))
        nrows > 0 && (pane_kw[:ySplit] = string(ySplit))
        kind != SPLIT && (pane_kw[:state] = kind == FROZEN ? "frozen" : "frozenSplit")

        push!(new_children, XML.Element(pfx_c * "pane"; pane_kw...))

        corner_selectable = (kind == SPLIT)
        for (pane_name, cell) in _selection_cells(nrows, ncols, topLeftCell; corner_selectable)
            push!(new_children, XML.Element(pfx_c * "selection"; pane=pane_name, activeCell=string(cell), sqref=string(cell)))
        end
    end

    new_sheetView = XML.Element(pfx_c * "sheetView", new_children...; sheetview_kwargs...)

    if sv_idx === nothing
        # No <sheetViews> at all yet: build it together with its <sheetView>
        # in one call, so it's never an empty node we'd need to mutate later.
        new_sheetViews = XML.Element(pfx_c * "sheetViews", new_sheetView)
        insertpos = length(root_children) + 1
        for name in ("sheetFormatPr", "cols", "sheetData")
            j = find_child_index(root_children, name)
            if j !== nothing
                insertpos = j
                break
            end
        end
        insert!(root_children, insertpos, new_sheetViews)
    elseif isnothing(sheetViews_children)
        # <sheetViews/> exists but is self-closed (no children vector to splice
        # into) -- replace the whole node, built together with its <sheetView>.
        root_children[sv_idx] = XML.Element(pfx_c * "sheetViews", new_sheetView)
    elseif existing_sv_idx !== nothing
        # <sheetViews> and <sheetView> both exist already: replace the sheetView
        # in place within its parent's (real, mutable) children vector.
        sheetViews_children[existing_sv_idx] = new_sheetView
    else
        # <sheetViews> exists with children, but none of them is a <sheetView>
        # (schema-invalid in practice, but handle it rather than error out).
        push!(sheetViews_children, new_sheetView)
    end

    set_worksheet_xml_document!(ws, doc)
    return ws
end

_anchor_to_counts(anchor_cell::AbstractString) = begin
    is_valid_cellname(anchor_cell) || throw(XLSXError("`$anchor_cell` is not a valid cell reference."))
    cr = CellRef(anchor_cell)
    (row_number(cr) - 1, column_number(cr) - 1)
end

#-----------------------------------------------------------------------------
# Public API
#-----------------------------------------------------------------------------

"""
    freezePanes(ws::Worksheet; nrows::Int=1, ncols::Int=0)
    freezePanes(ws::Worksheet, anchor_cell::AbstractString)

Freeze the first `nrows` rows and/or `ncols` columns of `ws` so they stay
visible while scrolling, or equivalently freeze everything above/left of
`anchor_cell` (the first cell of the scrolling region). Calling this again
replaces any existing frozen/split panes.

# Examples
```julia
julia> XLSX.freezePanes(sh)                       # freeze row 1 (the default)
julia> XLSX.freezePanes(sh; nrows=2, ncols=1)     # freeze first 2 rows and column A
julia> XLSX.freezePanes(sh, "B2")                 # equivalent to nrows=1, ncols=1
```

The Excel file must be opened in write mode to work with panes.

"""
freezePanes(ws::Worksheet; ncols::Int=0, nrows::Int=1) = _apply_pane!(ws, FROZEN, nrows, ncols)
freezePanes(ws::Worksheet, anchor_cell::AbstractString) = _apply_pane!(ws, FROZEN, _anchor_to_counts(anchor_cell)...)

"""
    splitFreeze(ws::Worksheet; nrows::Int=1, ncols::Int=0)
    splitFreeze(ws::Worksheet, anchor_cell::AbstractString)

Like [`freezePanes`](@ref), but marks the pane as originating from a split
(`state="frozenSplit"`) rather than a plain freeze (`state="frozen"`). This
mirrors what Excel writes when a user freezes panes while a draggable split
already exists. Functionally very similar to `freezePanes` for a freshly
created pane; the distinction mainly affects how Excel behaves if the pane
is later unfrozen interactively.

The Excel file must be opened in write mode to work with panes.

"""
splitFreeze(ws::Worksheet; ncols::Int=0, nrows::Int=1) = _apply_pane!(ws, FROZEN_SPLIT, nrows, ncols)
splitFreeze(ws::Worksheet, anchor_cell::AbstractString) = _apply_pane!(ws, FROZEN_SPLIT, _anchor_to_counts(anchor_cell)...)

"""
    splitPanes(ws::Worksheet; nrows::Int=1, ncols::Int=0)
    splitPanes(ws::Worksheet, anchor_cell::AbstractString)

Split `ws` into a draggable (non-frozen) pane layout at the given row/column
boundary.

!!! note "Approximate divider placement"
    The requested boundary is converted to a twips position by summing actual
    column widths / row heights (falling back to `sheetFormatPr` defaults,
    then Excel's built-in defaults). The column-width-to-pixel step assumes a
    Maximum Digit Width of $(DEFAULT_MDW) (Calibri 11 / Arial 10 @ 96dpi).
    The divider will be very close to the requested cell boundary but not 
    necessarily exact.

The Excel file must be opened in write mode to work with panes.

"""
splitPanes(ws::Worksheet; ncols::Int=0, nrows::Int=1) = _apply_pane!(ws, SPLIT, nrows, ncols)
splitPanes(ws::Worksheet, anchor_cell::AbstractString) = _apply_pane!(ws, SPLIT, _anchor_to_counts(anchor_cell)...)

"""
    removePanes(ws::Worksheet)

Remove any frozen or split panes from `ws`, restoring a plain single-pane
view. Equivalent to `freezePanes(ws; nrows=0, ncols=0)`.

The Excel file must be opened in write mode to work with panes.

"""
removePanes(ws::Worksheet) = _apply_pane!(ws, SPLIT, 0, 0)
