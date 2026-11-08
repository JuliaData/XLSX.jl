# ===========================================================================
# Inserting and deleting rows and columns
#
# Each edit is a `RowColEdit` (see shiftrefs.jl) applied to everything that refers
# to a place on the sheet: its cells and their row attributes, every formula in
# the workbook that refers to it, its merged cells, conditional formats, data
# validations, hyperlinks, AutoFilter, panes and selections, column definitions,
# page breaks, defined names and chart series.
# ===========================================================================

"""
    insertRows(ws::Worksheet, at; count=1)
    insertRows(xf::XLSXFile, "Sheet1!5"; count=1)

Insert `count` empty rows before row `at`, as Excel's *Insert Sheet Rows* does.
Rows from `at` down move down, and everything that refers to them follows: formulas
anywhere in the workbook, merged cells, conditional formats, data validations,
hyperlinks, the AutoFilter, panes, defined names (including the print area and
titles), chart series, and the Excel Tables, images, charts and comments on the
sheet. A range that spans row `at` grows to include the new rows.

`at` is a row number or a string naming one (`"5"`). The new rows are empty, and
formatted like the row above as Excel formats them; see [How new rows and columns
are formatted](@ref). Inserting rows that would push data past the last row Excel
allows is an error, as in Excel.

The file must be open for writing (`mode="rw"`).

# Examples
```julia
julia> XLSX.insertRows(sheet, 5)            # a new row 5

julia> XLSX.insertRows(sheet, 2; count=3)   # three new rows, 2 to 4
```

See also [`deleteRows`](@ref), [`insertColumns`](@ref) and [`deleteColumns`](@ref).
"""
function insertRows end

"""
    insertColumns(ws::Worksheet, at; count=1)
    insertColumns(xf::XLSXFile, "Sheet1!C"; count=1)

Insert `count` empty columns before column `at`, as Excel's *Insert Sheet Columns*
does. `at` is a column number or letters (`"C"`). Everything that refers to the
moved columns follows, as for [`insertRows`](@ref).

# Examples
```julia
julia> XLSX.insertColumns(sheet, "C")           # a new column C

julia> XLSX.insertColumns(sheet, 2; count=2)    # new columns B and C
```

See also [`deleteColumns`](@ref) and [`insertRows`](@ref).
"""
function insertColumns end

"""
    deleteRows(ws::Worksheet, rows)
    deleteRows(xf::XLSXFile, "Sheet1!3:5")

Delete rows, as Excel's *Delete Sheet Rows* does. Rows below move up, and
everything that refers to them follows, as for [`insertRows`](@ref). A reference
to a deleted cell becomes `#REF!`; a range that loses some of its rows shrinks,
and one that loses all of them goes too (a merged range, a conditional format, a
data validation, a hyperlink, the AutoFilter, an Excel Table), or becomes `#REF!`
in a formula or defined name. An image or chart set to move and size with its
cells (Excel's default for a chart) goes too when all its rows are deleted; one
that only moves, or doesn't, is kept. Deleting a table's header row, its totals
row or all its data rows, without the rest of it, is an error.

`rows` is a row number, a range or vector of row numbers, or a string naming
rows (`"3:5"`, `"3:5,9"`), cells (whose rows are used) or a defined name.

The file must be open for writing (`mode="rw"`).

# Examples
```julia
julia> XLSX.deleteRows(sheet, 5)

julia> XLSX.deleteRows(sheet, "3:5,9")
```

See also [`insertRows`](@ref) and [`deleteColumns`](@ref).
"""
function deleteRows end

"""
    deleteColumns(ws::Worksheet, cols)
    deleteColumns(xf::XLSXFile, "Sheet1!B:C")

Delete columns, as Excel's *Delete Sheet Columns* does. Columns to the right move
left, and everything that refers to them follows, as for [`deleteRows`](@ref).

`cols` is a column number, a range or vector of column numbers, or a string naming
columns (`"B"`, `"B:C"`, `"B:C,F"`), cells (whose columns are used) or a defined name.

# Examples
```julia
julia> XLSX.deleteColumns(sheet, "B:C")

julia> XLSX.deleteColumns(sheet, [2, 5])
```

See also [`insertColumns`](@ref) and [`deleteRows`](@ref).
"""
function deleteColumns end

for f in (:insertRows, :insertColumns, :deleteRows, :deleteColumns)
    @eval $f(xf::XLSXFile, ref::AbstractString; kw...) = $f(_rowcol_sheet(xf, ref)...; kw...)
end

insertRows(ws::Worksheet, at::Integer; count::Integer=1) = _insert!(ws, :row, at, count)
insertRows(ws::Worksheet, at::AbstractString; kw...) = insertRows(ws, _single(_row_spans(ws, at), "row"); kw...)
insertColumns(ws::Worksheet, at::Integer; count::Integer=1) = _insert!(ws, :column, at, count)
insertColumns(ws::Worksheet, at::AbstractString; kw...) = insertColumns(ws, _single(_col_spans(ws, at), "column"); kw...)
deleteRows(ws::Worksheet, rows) = _delete!(ws, :row, _row_spans(ws, rows))
deleteColumns(ws::Worksheet, cols) = _delete!(ws, :column, _col_spans(ws, cols))

function _insert!(ws::Worksheet, axis::Symbol, at::Integer, count::Integer)
    count >= 1 || throw(XLSXError("`count` must be at least 1. Found $count."))
    _edit!(ws, RowColEdit(ws.name, axis, at, count))
    return nothing
end

function _delete!(ws::Worksheet, axis::Symbol, spans::Vector{UnitRange{Int}})
    for s in reverse(spans)   # from the bottom (right), so the others stay where they are
        _edit!(ws, RowColEdit(ws.name, axis, first(s), -length(s)))
    end
    return nothing
end

# ===========================================================================
# Applying an edit
# ===========================================================================

function _edit!(ws::Worksheet, e::RowColEdit)
    what = e.count > 0 ? "insert $(e.axis)s" : "delete $(e.axis)s"
    _require_writable(ws, what)
    _require_worksheet(ws, what)
    _require_cache(ws, what)
    isnothing(ws.cache) && eachrow(ws)
    e.count > 0 && _check_room(ws, e)
    # Refuse what can't be done before anything changes.
    dead_comments = _comments_deleted_by(ws, e)
    table_plan = _plan_table_edits(ws, e)

    xf = get_xlsxfile(ws)
    wb = get_workbook(ws)
    dead_tables = Dict{String,Union{Nothing,Set{String}}}()   # tables, or their columns, deleted
    _delete_tables!(table_plan, dead_tables)
    foreach(r -> deleteComment(ws, r), dead_comments)
    _shift_comments!(ws, e)
    _delete_drawings_gone!(xf, _sheet_path(ws), e)
    _shift_drawing!(xf, _sheet_path(ws), e)
    _expand_shared_formulas!(wb, e)
    _shift_cells!(ws, e)
    _shift_cell_formulas!(wb, e)
    _shift_sheet_xml!(ws, e)
    _inherit_formats!(ws, e)
    _move_tables!(ws, table_plan, e, dead_tables)
    for s in wb.sheets   # formulas in conditional formats and data validations, on every sheet
        is_chartsheet(wb, s.name) && continue
        _repoint_sheet_xml_formulas!(s, r"(?s)\A.+\z" => (t -> _shift_formula(t, e; own_sheet=s.name)))
    end
    _shift_defined_names!(wb, e)
    for mime in (MIME_CHART, MIME_CHARTEX), path in parts_with_content_type(xf, mime)
        _rewrite_chart_formulas!(s -> _shift_formula(s, e), xml_root_element(get_xml_data(xf, path)))
    end
    _rewrite_table_formulas!((f, own) -> _shift_formula(f, e; own_sheet=own), xf)
    _kill_structured_refs!(xf, dead_tables)

    _set_dimension_from_cache!(ws)
    isempty(ws.cache) && set_dimension!(ws, CellRange("A1:A1"))
    update_worksheets_xml!(xf)
    return nothing
end

# An insert may not push data off the end of the sheet.
function _check_room(ws::Worksheet, e::RowColEdit)
    last_kept = _edit_max(e) - e.count
    for (r, row) in ws.cache.cells
        far = e.axis === :row ? r : maximum(keys(row); init=0)
        far > last_kept &&
            throw(XLSXError("Cannot insert $(e.count) $(e.axis)s: it would push data in $(e.axis) $far past the last $(e.axis) of the sheet."))
    end
    return nothing
end

# --- cells

function _shift_cells!(ws::Worksheet, e::RowColEdit)
    cache = ws.cache
    cells = typeof(cache.cells)()
    row_ht = typeof(cache.row_ht)()
    uh = ws.unhandled_attributes
    new_uh = Dict{Int,Dict{String,String}}()
    strings_gone = 0
    for (r, row) in cache.cells
        if e.axis === :row
            r2 = _shift_index(r, e)
            if isnothing(r2)
                strings_gone += count(c -> c.datatype == CT_STRING, values(row))
                continue
            end
            for c in values(row)
                c.ref = CellRef(r2, column_number(c.ref))
            end
            cells[r2] = row
        else
            r2 = r
            new_row = Dict{Int,Cell}()
            for (col, c) in row
                col2 = _shift_index(col, e)
                if isnothing(col2)
                    strings_gone += c.datatype == CT_STRING
                    continue
                end
                c.ref = CellRef(r, col2)
                new_row[col2] = c
            end
            cells[r] = new_row
        end
        row_ht[r2] = cache.row_ht[r]
        isnothing(uh) || !haskey(uh, r) || (new_uh[r2] = uh[r])
    end
    cache.cells = cells
    cache.row_ht = row_ht
    cache.rows_in_cache = collect(keys(cells))
    cache.row_index = Dict{Int,Int}()
    cache.dirty = true
    ws.unhandled_attributes = isempty(new_uh) ? nothing : new_uh
    ws.sst_count -= strings_gone
    return nothing
end

# --- cell formulas

# A shared formula is held once, in its first cell (the master), and worked out for
# the others from their offsets. That stays right after an edit only if the edit
# moves the whole block alike, and each cell's formula, moved, is still what the
# moved master gives for it: not so if, say, the cells refer across the edited row,
# or a delete makes some of them #REF! and not others. A group the edit could
# change is checked, and one that fails becomes a plain formula in each cell. The
# others stay shared; `_shift_cell_formulas!` moves their master and block.
function _expand_shared_formulas!(wb::Workbook, e::RowColEdit)
    groups = Dict{Tuple{String,Int},Tuple{SheetCellRef,ReferencedFormula}}()
    for (k, f) in wb.formulas
        f isa ReferencedFormula || continue
        (k.sheet == e.sheet || _formula_refers_to(f.formula, e.sheet, k.sheet)) && (groups[(k.sheet, f.id)] = (k, f))
    end
    isempty(groups) && return nothing
    members = Dict{Tuple{String,Int},Vector{SheetCellRef}}(g => SheetCellRef[] for g in keys(groups))
    for (k, f) in wb.formulas
        f isa FormulaReference && haskey(members, (k.sheet, f.id)) && push!(members[(k.sheet, f.id)], k)
    end
    for (g, (mk, master)) in groups
        _stays_shared(mk, master, members[g], e) && continue
        for k in members[g]
            text = shift_excel_references(master.formula, cell_offset(mk.cellref, k.cellref))
            wb.formulas[k] = Formula(text, nothing, nothing, wb.formulas[k].unhandled)
        end
        wb.formulas[mk] = Formula(master.formula, nothing, nothing, master.unhandled)
    end
    return nothing
end

# Whether the shared formula with master `mk` and cells `ks` is still right as a
# shared formula after edit `e`, with its master moved.
function _stays_shared(mk::SheetCellRef, master::ReferencedFormula, ks::Vector{SheetCellRef}, e::RowColEdit)::Bool
    if mk.sheet == e.sheet   # the block must move whole, keeping its shape
        block = CellRange(master.ref)
        moved = _shift_cellrange(block, e)
        (isnothing(moved) || size(moved) != size(block) || _shift_cellref(block.start, e) != moved.start) && return false
    end
    new_master = _shift_formula(master.formula, e; own_sheet=mk.sheet)
    for k in ks
        o = cell_offset(mk.cellref, k.cellref)
        _shift_formula(shift_excel_references(master.formula, o), e; own_sheet=mk.sheet) ==
            shift_excel_references(new_master, o) || return false
    end
    return true
end

# Every cell formula's references moved, and the edited sheet's formulas moved with
# their cells, or dropped with deleted ones.
function _shift_cell_formulas!(wb::Workbook, e::RowColEdit)
    moved = Dict{SheetCellRef,AbstractFormula}()
    for (k, f) in collect(wb.formulas)
        f isa ExplicitFormula && (f.formula = _shift_formula(f.formula, e; own_sheet=k.sheet))
        k.sheet == e.sheet || continue
        delete!(wb.formulas, k)
        c = _shift_cellref(k.cellref, e)
        isnothing(c) && continue
        if f isa Formula && !isnothing(f.ref) && !isempty(f.ref)   # an array or spill range
            f.ref = something(_shift_ref_body(f.ref, e), string(c))
        elseif f isa ReferencedFormula                              # a shared formula's block, moved whole
            f.ref = string(_shift_cellrange(CellRange(f.ref), e))
        end
        moved[SheetCellRef(k.sheet, c)] = f
    end
    merge!(wb.formulas, moved)
    return nothing
end

# --- defined names

function _shift_defined_names!(wb::Workbook, e::RowColEdit)
    for names in (wb.workbook_names, wb.worksheet_names), (k, v) in collect(names)
        names[k] = _shift_name_value(v, e)
    end
    return nothing
end

const _REF_NAME = "=#REF!"

# A range shrunk to one cell is written as the cell, as Excel writes it.
_one_cell_as_ref(r) = r isa CellRange && r.start == r.stop ? r.start : r

function _shift_name_value(v::DefinedNameValue, e::RowColEdit)::DefinedNameValue
    x = v.value
    if x isa String
        startswith(x, '=') || return v
        return DefinedNameValue(_shift_formula(x, e), v.isabs, v.hidden)
    end
    x isa DefinedNameRangeTypes && _same_sheetname(x.sheet, e.sheet) || return v
    if x isa SheetCellRef
        c = _shift_cellref(x.cellref, e)
        return isnothing(c) ? DefinedNameValue(_REF_NAME, false, v.hidden) : DefinedNameValue(SheetCellRef(x.sheet, c), v.isabs, v.hidden)
    elseif x isa SheetCellRange
        r = _one_cell_as_ref(_shift_cellrange(x.rng, e))
        isnothing(r) && return DefinedNameValue(_REF_NAME, false, v.hidden)
        return DefinedNameValue(r isa CellRef ? SheetCellRef(x.sheet, r) : SheetCellRange(x.sheet, r), v.isabs, v.hidden)
    end
    parts = Union{CellRef,CellRange}[]
    isabs = Bool[]
    for (p, a) in zip(x.rng, v.isabs)
        q = p isa CellRef ? _shift_cellref(p, e) : _one_cell_as_ref(_shift_cellrange(p, e))
        isnothing(q) && continue
        push!(parts, q)
        push!(isabs, a)
    end
    isempty(parts) && return DefinedNameValue(_REF_NAME, false, v.hidden)
    return DefinedNameValue(NonContiguousRange(x.sheet, parts), isabs, v.hidden)
end

# --- the sheet's own XML

function _shift_sheet_xml!(ws::Worksheet, e::RowColEdit)
    xf = get_xlsxfile(ws)
    root = _sheet_root(ws)
    _shift_ranges_in!(root, "mergeCells", "mergeCell", "ref", e; drop = r -> !occursin(':', r) || _is_one_cell(r))
    for tag in ("conditionalFormatting",)
        _filter_elements!(root, c -> localname(c) != tag || _shift_attr!(c, "sqref", e))
    end
    _shift_ranges_in!(root, "dataValidations", "dataValidation", "sqref", e)
    _shift_ranges_in!(root, "ignoredErrors", "ignoredError", "sqref", e)
    _shift_ranges_in!(root, "protectedRanges", "protectedRange", "sqref", e)
    _shift_hyperlinks!(xf, ws, root, e)
    _shift_autofilter!(ws, root, e)
    _shift_breaks!(root, e)
    e.axis === :column && _shift_ranges_in!(root, "cols", "col", "", e)
    for ext in XML.Node[c for c in XML.children(root) if localname(c) == "extLst"]
        _shift_ext_sqrefs!(ext, e)
    end
    frozen = _frozen_panes(root)
    for views in XML.Node[c for c in XML.children(root) if localname(c) == "sheetViews"], v in xml_elements(views)
        localname(v) == "sheetView" && _shift_sheet_view!(v, e)
    end
    isnothing(frozen) || _refreeze!(ws, frozen, e)
    return nothing
end

_is_one_cell(r::AbstractString) = (p = split(r, ':'); length(p) == 2 && p[1] == p[2])

# Move attribute `att` of `el`, a range list; `false` if nothing of it is left.
function _shift_attr!(el::XML.Node, att::String, e::RowColEdit)::Bool
    v = get(el, att, nothing)
    isnothing(v) && return true
    new = _shift_sqref(v, e)
    isnothing(new) && return false
    new == v || (el[att] = new)
    return true
end

# Remove the children of `parent` for which `keep` is false.
_filter_elements!(keep::Function, parent::XML.Node) = _filter_elements!(parent, keep)
function _filter_elements!(parent::XML.Node, keep::Function)
    kids = XML.children(parent)
    kids isa Vector || return nothing
    filter!(c -> XML.nodetype(c) != XML.Element || keep(c), kids)
    return nothing
end

# Move the ranges in attribute `att` of each `<item>` in the root's `<block>`,
# dropping items left with none (or, with `drop`, whose new range it rejects), and
# the block if it is left empty. `att=""` means a column span, `min`/`max`.
function _shift_ranges_in!(root::XML.Node, block::String, item::String, att::String, e::RowColEdit; drop=_ -> false)
    idx = findfirst(c -> localname(c) == block, XML.children(root))
    isnothing(idx) && return nothing
    b = XML.children(root)[idx]
    _filter_elements!(b) do c
        localname(c) == item || return true
        att == "" && return _shift_col_span!(c, e)
        _shift_attr!(c, att, e) || return false
        return !drop(c[att])
    end
    n = length(xml_elements(b))
    if n == 0
        deleteat!(XML.children(root), idx)
    elseif haskey(b, "count")
        b["count"] = string(n)
    end
    return nothing
end

function _shift_col_span!(c::XML.Node, e::RowColEdit)::Bool
    lo = parse(Int, c["min"])
    span = _shift_span(lo, parse(Int, get(c, "max", c["min"])), e)
    isnothing(span) && return false
    c["min"] = string(span[1])
    c["max"] = string(span[2])
    return true
end

# The x14 conditional formats and data validations in an `<extLst>`, each with its
# range in an `<xm:sqref>`: moved, or removed with nothing left.
function _shift_ext_sqrefs!(node::XML.Node, e::RowColEdit)
    removed = false
    _filter_elements!(node) do c
        sq = findfirst(k -> localname(k) == "sqref", xml_elements(c))
        if isnothing(sq)   # a container: kept unless the edit empties it
            had = !isempty(xml_elements(c))
            _shift_ext_sqrefs!(c, e)
            return !had || !isempty(xml_elements(c))
        end
        el = xml_elements(c)[sq]
        new = _shift_sqref(something(XML.is_simple_value(el), ""), e)
        isnothing(new) && (removed = true; return false)
        el[end] = XML.Text(new)
        return true
    end
    removed && haskey(node, "count") && (node["count"] = string(length(xml_elements(node))))
    return nothing
end

function _shift_hyperlinks!(xf::XLSXFile, ws::Worksheet, root::XML.Node, e::RowColEdit)
    idx = findfirst(c -> localname(c) == "hyperlinks", XML.children(root))
    isnothing(idx) && return nothing
    block = XML.children(root)[idx]
    _filter_elements!(block) do h
        _shift_attr!(h, "ref", e) && return true
        rid = get_prefixed_attr(h, "id")
        isnothing(rid) || _remove_sheet_rel!(xf, _sheet_path(ws), rid)
        return false
    end
    isempty(xml_elements(block)) && deleteat!(XML.children(root), idx)
    return nothing
end

function _shift_autofilter!(ws::Worksheet, root::XML.Node, e::RowColEdit)
    idx = findfirst(c -> localname(c) == "autoFilter", XML.children(root))
    isnothing(idx) && return nothing
    af = XML.children(root)[idx]
    old = CellRange(af["ref"])
    new = _shift_cellrange(old, e)
    if isnothing(new)
        deleteat!(XML.children(root), idx)
        delete!(get_workbook(ws).worksheet_names, (ws.sheetId, FILTER_DATABASE))
        return nothing
    end
    af["ref"] = string(new)
    if e.axis === :column   # a filter column's colId counts from the filter's first column
        _filter_elements!(af) do fc
            localname(fc) == "filterColumn" || return true
            col = _shift_index(column_number(old.start) + parse(Int, fc["colId"]), e)
            isnothing(col) && return false
            fc["colId"] = string(col - column_number(new.start))
            return true
        end
    end
    for s in xml_elements(af)   # a sort, with its ranges
        localname(s) == "sortState" || continue
        _shift_attr!(s, "ref", e)
        _filter_elements!(s, c -> localname(c) != "sortCondition" || _shift_attr!(c, "ref", e))
    end
    return nothing
end

# Page breaks: a `<brk id>` is the row (column) after which a page ends.
function _shift_breaks!(root::XML.Node, e::RowColEdit)
    block = e.axis === :row ? "rowBreaks" : "colBreaks"
    idx = findfirst(c -> localname(c) == block, XML.children(root))
    isnothing(idx) && return nothing
    b = XML.children(root)[idx]
    _filter_elements!(b) do brk
        p = _shift_index(parse(Int, brk["id"]), e)
        isnothing(p) && return false
        brk["id"] = string(p)
        return true
    end
    n = length(xml_elements(b))
    n == 0 && return deleteat!(XML.children(root), idx)
    for att in ("count", "manualBreakCount")
        haskey(b, att) && (b[att] = string(n))
    end
    return nothing
end

# A cell after the edit; if it was deleted, the cell now where it was.
function _shift_or_stay(ref::CellRef, e::RowColEdit)::CellRef
    c = _shift_cellref(ref, e)
    isnothing(c) || return c
    return e.axis === :row ? CellRef(min(e.at, EXCEL_MAX_ROWS), column_number(ref)) : CellRef(row_number(ref), min(e.at, EXCEL_MAX_COLS))
end

function _shift_sheet_view!(v::XML.Node, e::RowColEdit)
    for c in xml_elements(v)
        if localname(c) == "pane"
            haskey(c, "topLeftCell") && (c["topLeftCell"] = string(_shift_or_stay(CellRef(c["topLeftCell"]), e)))
        elseif localname(c) == "selection"
            active = haskey(c, "activeCell") ? _shift_or_stay(CellRef(c["activeCell"]), e) : nothing
            isnothing(active) || (c["activeCell"] = string(active))
            if haskey(c, "sqref") && !_shift_attr!(c, "sqref", e)
                c["sqref"] = string(something(active, CellRef("A1")))
            end
        end
    end
    return nothing
end

# The sheet's frozen panes, as (kind, rows frozen, columns frozen), or `nothing`
# if it has none.
function _frozen_panes(root::XML.Node)
    for views in xml_elements(root)
        localname(views) == "sheetViews" || continue
        for v in xml_elements(views), p in xml_elements(v)
            localname(p) == "pane" || continue
            state = get(p, "state", "split")
            state in ("frozen", "frozenSplit") || return nothing
            n(att) = round(Int, parse(Float64, get(p, att, "0")))
            return (state == "frozen" ? FROZEN : FROZEN_SPLIT, n("ySplit"), n("xSplit"))
        end
    end
    return nothing
end

# Freeze the sheet again after edit `e`: an edit among the frozen rows (columns)
# freezes more or fewer, and with none left frozen, the panes go.
function _refreeze!(ws::Worksheet, (kind, nrows, ncols), e::RowColEdit)
    kept(n) = n == 0 ? 0 : (s = _shift_span(1, n, e); isnothing(s) ? 0 : s[2])
    e.axis === :row ? (nrows = kept(nrows)) : (ncols = kept(ncols))
    nrows == 0 && ncols == 0 ? removePanes(ws) : _apply_pane!(ws, kind, nrows, ncols)
    return nothing
end
