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
_charwidth_to_pixels(width::Real) =
    Int(floor(((256 * width + floor(128 / DEFAULT_MDW)) / 256) * DEFAULT_MDW))

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

function _default_col_width_chars(ws::Worksheet)::Float64
    d = _sheet_format_default(ws, "defaultColWidth")
    !isnothing(d) && return d

    # defaultColWidth is often absent; Excel then derives it from
    # baseColWidth (a raw character count with no padding yet applied) via
    # the same char-count -> stored-width formula used for explicit <col>
    # widths (ECMA-376 18.3.1.13): width = floor((NoC*MDW+5)/MDW*256)/256.
    base = _sheet_format_default(ws, "baseColWidth")
    if !isnothing(base)
        return floor(((base * DEFAULT_MDW + 5) / DEFAULT_MDW) * 256) / 256
    end

    return DEFAULT_COL_WIDTH
end

function _col_width_chars(ws::Worksheet, col::Int)::Float64
    cr = CellRef(1, col)
    d = get_dimension(ws)
    if !isnothing(d) && cr ∈ d
        w = getColumnWidth(ws, cr)
        !isnothing(w) && return Float64(w)
    end
    return _default_col_width_chars(ws)
end

function _row_height_pts(ws::Worksheet, row::Int)::Float64
    cr = CellRef(row, 1)
    d = get_dimension(ws)
    if !isnothing(d) && cr ∈ d
        h = getRowHeight(ws, cr)
        !isnothing(h) && h >= 0 && return Float64(h)
    end
    d2 = _sheet_format_default(ws, "defaultRowHeight")
    !isnothing(d2) && return d2
    return DEFAULT_ROW_HEIGHT
end

_cols_to_twips(ws::Worksheet, ncols::Int)::Int =
    ncols <= 0 ? 0 : sum(_pixels_to_twips(_charwidth_to_pixels(_col_width_chars(ws, c))) for c in 1:ncols)

_rows_to_twips(ws::Worksheet, nrows::Int)::Int =
    nrows <= 0 ? 0 : sum(_points_to_twips(_row_height_pts(ws, r)) for r in 1:nrows)

#-----------------------------------------------------------------------------
# Selection-cell placement
#-----------------------------------------------------------------------------

# Which pane names get a <selection>, and which cell each should show, in the
# order Excel writes them. corner_selectable=true only for plain (non-frozen)
# splits -- the top-left pane is a real interactive pane there; for
# frozen/frozenSplit it's locked and Excel doesn't give it a <selection> at all.
function _selection_cells(nrows::Int, ncols::Int, topLeftCell::CellRef; corner_selectable::Bool)
    sels = Pair{String,CellRef}[]
    corner_selectable && (ncols > 0 || nrows > 0) && push!(sels, "topLeft" => CellRef(1, 1))
    if ncols > 0 && nrows > 0
        push!(sels, "topRight" => CellRef(1, topLeftCell.column_number),
                    "bottomLeft" => CellRef(topLeftCell.row_number, 1),
                    "bottomRight" => topLeftCell)
    elseif ncols > 0
        push!(sels, "topRight" => topLeftCell)
    elseif nrows > 0
        push!(sels, "bottomLeft" => topLeftCell)
    end
    return sels
end

#-----------------------------------------------------------------------------
# Core implementation shared by all four public functions
#-----------------------------------------------------------------------------

function _apply_pane!(ws::Worksheet, kind::PaneKind, nrows::Int, ncols::Int)

    _require_writable(ws, "apply pane")

    (nrows < 0 || ncols < 0) && throw(XLSXError("nrows and ncols must be non-negative."))

    pfx_c = _sheet_view_prefix(ws)

    # Build the new pane/selection children (empty for removePanes).
    new_children = XML.Node[]
    if !(nrows == 0 && ncols == 0)
        topLeftCell = CellRef(nrows + 1, ncols + 1)
        activePane = nrows > 0 && ncols > 0 ? "bottomRight" :
                     ncols > 0 ? "topRight" : "bottomLeft"
        xSplit, ySplit = kind == SPLIT ? (_cols_to_twips(ws, ncols), _rows_to_twips(ws, nrows)) : (ncols, nrows)

        pane_atts = Pair{String,String}[]   # in the order Excel writes them
        ncols > 0 && push!(pane_atts, "xSplit" => string(xSplit))
        nrows > 0 && push!(pane_atts, "ySplit" => string(ySplit))
        push!(pane_atts, "topLeftCell" => string(topLeftCell), "activePane" => activePane)
        kind != SPLIT && push!(pane_atts, "state" => kind == FROZEN ? "frozen" : "frozenSplit")

        push!(new_children, _xml_element(pfx_c * "pane", pane_atts))

        corner_selectable = (kind == SPLIT)
        for (pane_name, cell) in _selection_cells(nrows, ncols, topLeftCell; corner_selectable)
            push!(new_children, XML.Element(pfx_c * "selection"; pane=pane_name, activeCell=string(cell), sqref=string(cell)))
        end
    end

    # The panes count from row 1 and column 1, so the view starts there: Excel
    # freezes at the cell in view, so a view scrolled to row 10 would freeze rows 10 on.
    _update_sheet_view!(ws) do atts, _
        return _merge_atts(atts, ["topLeftCell" => nothing]), new_children
    end
    return ws
end

# "x:" for a sheet whose elements are prefixed, otherwise "".
_sheet_view_prefix(ws::Worksheet) = (pfx = get_prefix(ws); isempty(pfx) ? "" : "$(pfx):")

# Replace the sheet's first `<sheetView>` with one whose attributes and children are
# `f(atts, kids)`, given the old ones (none if there is no `<sheetView>` yet). The
# attributes are written in the schema's order, as Excel writes them.
function _update_sheet_view!(f, ws::Worksheet)
    doc = get_worksheet_xml_document(ws)
    pfx_c = _sheet_view_prefix(ws)

    wroot_idx, sv_idx = get_idces(doc, "worksheet", "sheetViews")
    wroot_idx === nothing && throw(XLSXError("Malformed worksheet: no <worksheet> root element found."))
    xroot = doc[wroot_idx]
    isnothing(xroot.children) && throw(XLSXError("Malformed worksheet: <worksheet> element has no children."))
    root_children = XML.children(xroot)

    existing_attrs = Pair{String,String}[]
    existing_kids = XML.Node[]
    existing_sv_idx = nothing
    sheetViews_children = nothing

    if sv_idx !== nothing
        sheetViews_node = root_children[sv_idx]
        sheetViews_children = sheetViews_node.children  # may be `nothing` if self-closed <sheetViews/>
        if !isnothing(sheetViews_children)
            existing_sv_idx = find_child_index(sheetViews_children, "sheetView")
            if existing_sv_idx !== nothing
                existing_sheetview = sheetViews_children[existing_sv_idx]
                existing_attrs = _atts_vector(existing_sheetview)
                existing_kids = XML.Node[c for c in XML.children(existing_sheetview)]
            end
        end
    end

    atts, kids = f(existing_attrs, existing_kids)
    atts = _merge_atts(atts, any(p -> first(p) == "workbookViewId", atts) ? () : ["workbookViewId" => "0"])
    sort!(atts; by = p -> something(findfirst(==(first(p)), _SHEET_VIEW_ATT_ORDER), length(_SHEET_VIEW_ATT_ORDER) + 1))
    new_sheetView = _xml_element(pfx_c * "sheetView", atts, kids)

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
    return nothing
end

# The attributes of `<sheetView>` in the schema's order.
const _SHEET_VIEW_ATT_ORDER = ("windowProtection", "showFormulas", "showGridLines", "showRowColHeaders",
    "showZeros", "rightToLeft", "tabSelected", "showRuler", "showOutlineSymbols", "defaultGridColor",
    "showWhiteSpace", "view", "topLeftCell", "colorId", "zoomScale", "zoomScaleNormal",
    "zoomScaleSheetLayoutView", "zoomScalePageLayoutView", "workbookViewId")

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
replaces any existing frozen/split panes. The panes count from row 1 and column A, so
a view scrolled elsewhere (see [`setSheetView`](@ref)) is scrolled back to the top.

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

#-----------------------------------------------------------------------------
# Sheet view: zoom, gridlines, headings, zeros, direction, scroll and selection
#-----------------------------------------------------------------------------

# The on/off settings of a sheet's view and their defaults, which Excel leaves unwritten.
const _SHEET_VIEW_FLAGS = (showGridLines = true, showRowColHeaders = true, showZeros = true, rightToLeft = false)
# Each view keeps its own zoom, and the zoom Excel gives a view seen for the first time.
const _SHEET_VIEW_ZOOMS = (normal = ("zoomScaleNormal", "100"), pageLayout = ("zoomScalePageLayoutView", "100"),
                           pageBreakPreview = ("zoomScaleSheetLayoutView", "60"))
const _SHEET_VIEW_KEYWORDS = (:zoom, :view, keys(_SHEET_VIEW_FLAGS)..., :topLeftCell, :activeCell, :selection)

"""
    setSheetView(ws::Worksheet; zoom, view, showGridLines, showRowColHeaders, showZeros,
                 rightToLeft, topLeftCell, activeCell, selection)

Set how a worksheet looks when Excel opens it, as Excel's *View* tab does. Only the
keywords given are changed.

- `zoom`: the zoom, in percent, from 10 to 400. Excel keeps a zoom for each `view`;
  this sets the zoom of the view the sheet is in.
- `view`: `:normal`, `:pageLayout` or `:pageBreakPreview`. A view takes back the zoom it
  had before.
- `showGridLines`, `showRowColHeaders`: show the gridlines, and the row and column
  headings (`1, 2, 3...` and `A, B, C...`), on screen. For printing them, see
  [`setPrintOptions`](@ref).
- `showZeros`: show a cell whose value is zero as `0`, rather than empty.
- `rightToLeft`: show column A at the right, for right-to-left languages.
- `topLeftCell`: scroll so that this cell is at the top left of the window. With frozen
  panes, it is the top-left cell of the part that scrolls, so it must be below and to
  the right of the frozen rows and columns.
- `selection`: the selected cells: a cell or range (`"B2:D5"`), several separated by
  commas (`"B2:D5,F7"`), or a defined name for them.
- `activeCell`: the cell with the cursor, which must be in the `selection`. Given alone,
  it is selected on its own; a `selection` alone has its first cell active. With frozen
  or split panes, the selection is the active pane's, as in Excel.

Cells are given as a string (`"C10"`) or a `CellRef`.

The file must be open for writing (`mode="rw"`).

# Examples
```julia
julia> XLSX.setSheetView(sheet; zoom=150, showGridLines=false)

julia> XLSX.setSheetView(sheet; selection="B2:D5", activeCell="C3", topLeftCell="A1")
```

See also [`getSheetView`](@ref) and [`freezePanes`](@ref).
"""
function setSheetView(ws::Worksheet; kw...)
    _require_writable(ws, "set the sheet view")
    _require_worksheet(ws, "set the sheet view")
    isempty(kw) && throw(XLSXError("Give the view settings to change as keywords, such as `zoom=150`."))
    for k in keys(kw)
        k in _SHEET_VIEW_KEYWORDS || throw(XLSXError("Unknown keyword `$k`. Must be one of $(join(_SHEET_VIEW_KEYWORDS, ", "))."))
    end
    for k in keys(_SHEET_VIEW_FLAGS)
        haskey(kw, k) && !(kw[k] isa Bool) && throw(XLSXError("`$k` must be `true` or `false`. Found `$(repr(kw[k]))`."))
    end
    if haskey(kw, :view)
        (kw[:view] isa Symbol && haskey(_SHEET_VIEW_ZOOMS, kw[:view])) ||
            throw(XLSXError("Invalid `view` `$(repr(kw[:view]))`. Must be one of $(join((":$k" for k in keys(_SHEET_VIEW_ZOOMS)), ", "))."))
    end
    if haskey(kw, :zoom)
        (kw[:zoom] isa Integer && 10 <= kw[:zoom] <= 400) ||
            throw(XLSXError("Invalid `zoom` `$(repr(kw[:zoom]))`. Must be a whole percentage from 10 to 400."))
    end
    topleft = haskey(kw, :topLeftCell) ? _view_cell(kw[:topLeftCell], :topLeftCell) : nothing
    sel = haskey(kw, :selection) || haskey(kw, :activeCell) ?
          _view_selection(ws, get(kw, :selection, nothing), get(kw, :activeCell, nothing)) : nothing

    _update_sheet_view!(ws) do atts, kids
        old = Dict(atts)
        changes = Pair{String,Union{Nothing,String}}[]
        for (k, d) in pairs(_SHEET_VIEW_FLAGS)
            haskey(kw, k) && push!(changes, String(k) => kw[k] == d ? nothing : kw[k] ? "1" : "0")
        end
        view = haskey(kw, :view) ? kw[:view] : Symbol(get(old, "view", "normal"))
        haskey(_SHEET_VIEW_ZOOMS, view) || (view = :normal)
        haskey(kw, :view) && push!(changes, "view" => view == :normal ? nothing : String(view))
        zoom_att, zoom_default = _SHEET_VIEW_ZOOMS[view]
        if haskey(kw, :zoom)   # the view's own zoom is written once set, as Excel does
            zoom = string(kw[:zoom])
            push!(changes, "zoomScale" => (zoom == "100" ? nothing : zoom), zoom_att => zoom)
        elseif haskey(kw, :view)
            zoom = get(old, zoom_att, zoom_default)
            push!(changes, "zoomScale" => (zoom == "100" ? nothing : zoom))
        end

        kids = XML.Node[k for k in kids]
        ip = findfirst(c -> localname(c) == "pane", kids)
        pane = isnothing(ip) ? nothing : kids[ip]
        if !isnothing(topleft)
            if isnothing(pane)
                push!(changes, "topLeftCell" => topleft == CellRef(1, 1) ? nothing : string(topleft))
            else
                _check_scroll_cell(pane, CellRef(get(old, "topLeftCell", "A1")), topleft)
                kids[ip] = _xml_element(XML.tag(pane), _merge_atts(_atts_vector(pane), ["topLeftCell" => string(topleft)]))
            end
        end
        isnothing(sel) || _set_view_selection!(kids, _active_pane(pane), _sheet_view_prefix(ws), sel...)
        return _merge_atts(atts, changes), kids
    end
    return nothing
end

_view_cell(c::CellRef, ::Symbol) = c
_view_cell(c, k::Symbol) = (c isa AbstractString && is_valid_cellname(c)) ? CellRef(c) :
    throw(XLSXError("`$k` must be a single cell, such as `\"C10\"`. Found `$(repr(c))`."))

# The selection's active cell and ranges.
function _view_selection(ws::Worksheet, selection, active)
    ac = isnothing(active) ? nothing : _view_cell(active, :activeCell)
    rngs = CellRange[]
    if isnothing(selection)
        push!(rngs, CellRange(ac, ac))
    else
        parts = selection isa Union{CellRef,CellRange} ? Any[selection] :
                selection isa AbstractString ? _rowcol_parts(ws, selection) :
                throw(XLSXError("`selection` must be cells, such as `\"B2:D5\"` or `\"B2:D5,F7\"`. Found `$(repr(selection))`."))
        for p in parts
            p isa Union{CellRef,CellRange} || throw(XLSXError("A selection is made of cells, such as `B2:D5`. Found `$p`."))
            push!(rngs, p isa CellRef ? CellRange(p, p) : p)
        end
    end
    ac = something(ac, first(rngs).start)
    any(r -> ac ∈ r, rngs) || throw(XLSXError("The `activeCell` $ac must be in the `selection`."))
    return ac, rngs
end

# A frozen pane's scrolling part starts below and right of the frozen rows and columns.
function _check_scroll_cell(pane::XML.Node, view_topleft::CellRef, cell::CellRef)
    get(pane, "state", "split") in ("frozen", "frozenSplit") || return nothing
    first_row = view_topleft.row_number + parse(Int, get(pane, "ySplit", "0"))
    first_col = view_topleft.column_number + parse(Int, get(pane, "xSplit", "0"))
    (cell.row_number >= first_row && cell.column_number >= first_col) ||
        throw(XLSXError("The `topLeftCell` $cell is in the frozen panes: with these panes it must be at or below and right of $(CellRef(first_row, first_col))."))
    return nothing
end

# Excel keeps the selection of each pane in a `<selection>`, and shows the active pane's.
# A sheet without panes has one `<selection>`, with no `pane`.
_active_pane(pane) = isnothing(pane) ? nothing : get(pane, "activePane", "topLeft")
_selection_pane(s::XML.Node) = get(s, "pane", "topLeft")

function _set_view_selection!(kids::Vector, pane_name, pfx_c::String, ac::CellRef, rngs::Vector{CellRange})
    atts = Pair{String,String}[]
    isnothing(pane_name) || push!(atts, "pane" => pane_name)
    push!(atts, "activeCell" => string(ac), "sqref" => join(_sqref_part.(rngs), ' '))
    el = _xml_element(pfx_c * "selection", atts)
    target = something(pane_name, "topLeft")
    i = findfirst(c -> localname(c) == "selection" && _selection_pane(c) == target, kids)
    if !isnothing(i)
        kids[i] = el
    else   # after the pane and any other selections, before anything else
        j = findlast(c -> localname(c) in ("pane", "selection"), kids)
        insert!(kids, something(j, 0) + 1, el)
    end
    return kids
end
_sqref_part(r::CellRange) = r.start == r.stop ? string(r.start) : string(r)

"""
    getSheetView(ws::Worksheet) -> NamedTuple

How a worksheet looks when Excel opens it: a `NamedTuple` with fields `zoom`, `view`,
`showGridLines`, `showRowColHeaders`, `showZeros`, `rightToLeft`, `topLeftCell`,
`activeCell` and `selection`, described in [`setSheetView`](@ref). `topLeftCell` and
`activeCell` are `CellRef`s, and `selection` is a string such as `"B2:D5,F7"`. A setting
the sheet doesn't set has Excel's default.

Works on files opened for reading only.

# Examples
```julia
julia> XLSX.getSheetView(sheet).zoom
150
```
"""
function getSheetView(ws::Worksheet)
    sv = nothing
    for views in _sheet_top_elements(get_xlsxfile(ws), _sheet_path(ws), ("sheetViews",))
        svs = [c for c in xml_elements(views) if localname(c) == "sheetView"]
        isempty(svs) || (sv = first(svs); break)
    end
    atts = isnothing(sv) ? Dict{String,String}() : Dict(_atts_vector(sv))
    kids = isnothing(sv) ? XML.Node[] : collect(xml_elements(sv))
    flags = NamedTuple{keys(_SHEET_VIEW_FLAGS)}(map(k -> _xml_bool(atts, String(k), _SHEET_VIEW_FLAGS[k]), keys(_SHEET_VIEW_FLAGS)))
    ip = findfirst(c -> localname(c) == "pane", kids)
    pane = isnothing(ip) ? nothing : kids[ip]
    topleft = get(atts, "topLeftCell", "A1")
    isnothing(pane) || (topleft = get(pane, "topLeftCell", topleft))
    target = something(_active_pane(pane), "topLeft")
    i = findfirst(c -> localname(c) == "selection" && _selection_pane(c) == target, kids)
    s = isnothing(i) ? nothing : kids[i]
    active = isnothing(s) ? "A1" : get(s, "activeCell", "A1")
    sqref = isnothing(s) ? active : get(s, "sqref", active)
    return (zoom = parse(Int, get(atts, "zoomScale", "100")), view = Symbol(get(atts, "view", "normal")), flags...,
            topLeftCell = CellRef(topleft), activeCell = CellRef(active), selection = join(split(sqref), ','))
end
