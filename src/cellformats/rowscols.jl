# ===========================================================================
# Hide and group (outline) rows and columns
#
# A column's `hidden`, `outlineLevel` and `collapsed` are attributes of the
# `<col>` definition spanning it, in the worksheet's `<cols>` block. A row's are
# attributes of its `<row>`, kept in `ws.unhandled_attributes` between reading
# and writing the sheet, like the row's other attributes.
#
# A group is a run of rows (or columns) one outline level deeper than their
# neighbours. Its summary row is the one just below it (column: just to the
# right), unless the sheet's `<outlinePr>` says otherwise. A collapsed group has
# its members hidden and `collapsed="1"` on its summary row.
# ===========================================================================

const EXCEL_MAX_OUTLINE_LEVEL = 7

"""
    setColumnHidden(ws::Worksheet, cols; hidden=true)
    setColumnHidden(xf::XLSXFile, "Sheet1!B:D"; hidden=true)

Hide the columns `cols`, or show them again with `hidden=false`.

`cols` may be a column number, a range or vector of column numbers, or a string
naming columns (`"B"`, `"B:D"`), cells (`"B2:D5"`, whose columns are used) or a
defined name, or several of these separated by commas (`"B:C,F"`). Columns
needn't hold any data.

The file must be open for writing (`mode="rw"`).

# Examples
```julia
julia> XLSX.setColumnHidden(sheet, "C:D")

julia> XLSX.setColumnHidden(sheet, [2, 5, 6])

julia> XLSX.setColumnHidden(sheet, 3; hidden=false)
```

See also [`setRowHidden`](@ref), [`isColumnHidden`](@ref) and [`groupColumns`](@ref).
"""
function setColumnHidden end

"""
    setRowHidden(ws::Worksheet, rows; hidden=true)
    setRowHidden(xf::XLSXFile, "Sheet1!3:5"; hidden=true)

Hide the rows `rows`, or show them again with `hidden=false`.

`rows` may be a row number, a range or vector of row numbers, or a string naming
rows (`"3"`, `"3:5"`), cells (`"A3:D5"`, whose rows are used) or a defined name,
or several of these separated by commas (`"3:4,8"`). Rows needn't hold any data.

The file must be open for writing (`mode="rw"`).

# Examples
```julia
julia> XLSX.setRowHidden(sheet, 2:4)

julia> XLSX.setRowHidden(sheet, "7"; hidden=false)
```

See also [`setColumnHidden`](@ref), [`isRowHidden`](@ref) and [`groupRows`](@ref).
"""
function setRowHidden end

"""
    groupColumns(ws::Worksheet, cols; collapsed=false)
    groupColumns(xf::XLSXFile, "Sheet1!B:D"; collapsed=false)

Group the columns `cols`, as Excel's *Data → Group* does, adding the outline
control that hides and shows them. `cols` is given as for [`setColumnHidden`](@ref);
each run of adjacent columns in it becomes a group.

Grouping columns that are already grouped puts them one outline level deeper,
making a group within a group. Excel allows up to $EXCEL_MAX_OUTLINE_LEVEL levels.

With `collapsed=true` the group is shown collapsed: its columns are hidden and
the column after it, which carries the outline button, is marked collapsed.

The file must be open for writing (`mode="rw"`).

# Examples
```julia
julia> XLSX.groupColumns(sheet, "B:D")

julia> XLSX.groupColumns(sheet, "C:D"; collapsed=true)   # a collapsed group within B:D
```

See also [`ungroupColumns`](@ref), [`getColumnOutlineLevel`](@ref) and [`groupRows`](@ref).
"""
function groupColumns end

"""
    groupRows(ws::Worksheet, rows; collapsed=false)
    groupRows(xf::XLSXFile, "Sheet1!3:5"; collapsed=false)

Group the rows `rows`, as Excel's *Data → Group* does, adding the outline
control that hides and shows them. `rows` is given as for [`setRowHidden`](@ref);
each run of adjacent rows in it becomes a group.

Grouping rows that are already grouped puts them one outline level deeper,
making a group within a group. Excel allows up to $EXCEL_MAX_OUTLINE_LEVEL levels.

With `collapsed=true` the group is shown collapsed: its rows are hidden and the
row below it, which carries the outline button, is marked collapsed.

The file must be open for writing (`mode="rw"`).

# Examples
```julia
julia> XLSX.groupRows(sheet, 2:10)

julia> XLSX.groupRows(sheet, 3:5; collapsed=true)   # a collapsed group within 2:10
```

See also [`ungroupRows`](@ref), [`getRowOutlineLevel`](@ref) and [`groupColumns`](@ref).
"""
function groupRows end

"""
    ungroupColumns(ws::Worksheet, cols)
    ungroupColumns(xf::XLSXFile, "Sheet1!B:D")

Move the columns `cols` out one outline level, undoing [`groupColumns`](@ref).
Columns not in a group are left alone. Columns hidden by a collapsed group stay
hidden: show them with [`setColumnHidden`](@ref)`(ws, cols; hidden=false)`.

The file must be open for writing (`mode="rw"`).
"""
function ungroupColumns end

"""
    ungroupRows(ws::Worksheet, rows)
    ungroupRows(xf::XLSXFile, "Sheet1!3:5")

Move the rows `rows` out one outline level, undoing [`groupRows`](@ref). Rows not
in a group are left alone. Rows hidden by a collapsed group stay hidden: show
them with [`setRowHidden`](@ref)`(ws, rows; hidden=false)`.

The file must be open for writing (`mode="rw"`).
"""
function ungroupRows end

for f in (:setColumnHidden, :setRowHidden, :groupColumns, :groupRows, :ungroupColumns, :ungroupRows)
    @eval $f(xf::XLSXFile, ref::AbstractString; kw...) = $f(_rowcol_sheet(xf, ref)...; kw...)
end

setColumnHidden(ws::Worksheet, cols; hidden::Bool=true) = (_set_col_attr!(ws, _col_spans(ws, cols), "hidden", hidden ? "1" : nothing); nothing)
setRowHidden(ws::Worksheet, rows; hidden::Bool=true) = (_set_row_attr!(ws, _row_spans(ws, rows), "hidden", hidden ? "1" : nothing); nothing)
groupColumns(ws::Worksheet, cols; collapsed::Bool=false) = _outline!(ws, _col_spans(ws, cols), :col, +1, collapsed)
groupRows(ws::Worksheet, rows; collapsed::Bool=false) = _outline!(ws, _row_spans(ws, rows), :row, +1, collapsed)
ungroupColumns(ws::Worksheet, cols) = _outline!(ws, _col_spans(ws, cols), :col, -1, false)
ungroupRows(ws::Worksheet, rows) = _outline!(ws, _row_spans(ws, rows), :row, -1, false)

"""
    isColumnHidden(ws::Worksheet, col) -> Bool
    isColumnHidden(xf::XLSXFile, "Sheet1!C") -> Bool

Whether column `col` is hidden. `col` is a column number or a string naming a
column (`"C"`) or a cell in it (`"C5"`).

See also [`setColumnHidden`](@ref).
"""
function isColumnHidden end

"""
    isRowHidden(ws::Worksheet, row) -> Bool
    isRowHidden(xf::XLSXFile, "Sheet1!5") -> Bool

Whether row `row` is hidden. `row` is a row number or a string naming a row
(`"5"`) or a cell in it (`"C5"`).

Rows can be read from a file opened only for reading if the cell cache is
enabled, as it is by default.

See also [`setRowHidden`](@ref).
"""
function isRowHidden end

"""
    getColumnOutlineLevel(ws::Worksheet, col) -> Int
    getColumnOutlineLevel(xf::XLSXFile, "Sheet1!C") -> Int

The outline level of column `col`: 0 if it isn't grouped, 1 in a group, 2 in a
group within a group, and so on. `col` is given as for [`isColumnHidden`](@ref).

See also [`groupColumns`](@ref).
"""
function getColumnOutlineLevel end

"""
    getRowOutlineLevel(ws::Worksheet, row) -> Int
    getRowOutlineLevel(xf::XLSXFile, "Sheet1!5") -> Int

The outline level of row `row`: 0 if it isn't grouped, 1 in a group, 2 in a
group within a group, and so on. `row` is given as for [`isRowHidden`](@ref).

See also [`groupRows`](@ref).
"""
function getRowOutlineLevel end

for f in (:isColumnHidden, :isRowHidden, :getColumnOutlineLevel, :getRowOutlineLevel)
    @eval $f(xf::XLSXFile, ref::AbstractString) = $f(_rowcol_sheet(xf, ref)...)
end

isColumnHidden(ws::Worksheet, col) = _col_attr(ws, _single(_col_spans(ws, col), "column"), "hidden") in ("1", "true")
isRowHidden(ws::Worksheet, row) = _row_attr(ws, _single(_row_spans(ws, row), "row"), "hidden") in ("1", "true")
getColumnOutlineLevel(ws::Worksheet, col)::Int = _level(_col_attr(ws, _single(_col_spans(ws, col), "column"), "outlineLevel"))
getRowOutlineLevel(ws::Worksheet, row)::Int = _level(_row_attr(ws, _single(_row_spans(ws, row), "row"), "outlineLevel"))

# ===========================================================================
# Which rows and columns
# ===========================================================================

# The worksheet and the reference in `"Sheet1!B:D"`.
function _rowcol_sheet(xf::XLSXFile, ref::AbstractString)
    occursin('!', ref) || throw(XLSXError("Give the sheet as well as the reference, e.g. `\"Sheet1!B:D\"`. Found `$ref`."))
    sheet, r = rsplit(ref, '!'; limit=2)
    sheet = replace(strip(sheet, '\''), "''" => "'")
    hassheet(xf, sheet) || throw(XLSXError("No sheet named `$sheet` in the workbook."))
    return xf[sheet], String(r)
end

# Sorted, disjoint runs of the numbers in `v`.
function _runs(v::AbstractVector{<:Integer})::Vector{UnitRange{Int}}
    out = UnitRange{Int}[]
    for x in sort!(unique(Int.(v)))
        if !isempty(out) && x == last(out[end]) + 1
            out[end] = first(out[end]):x
        else
            push!(out, x:x)
        end
    end
    return out
end

function _check_spans(spans::Vector{UnitRange{Int}}, maxn::Int, what::String)
    isempty(spans) && throw(XLSXError("No $(what)s given."))
    for s in spans
        (first(s) >= 1 && last(s) <= maxn) ||
            throw(XLSXError("Invalid $what number: $(first(s) < 1 ? first(s) : last(s)). Must be from 1 to $maxn."))
    end
    return spans
end

_col_spans(ws::Worksheet, c) = _check_spans(_col_runs(ws, c), EXCEL_MAX_COLS, "column")
_row_spans(ws::Worksheet, r) = _check_spans(_row_runs(ws, r), EXCEL_MAX_ROWS, "row")

_col_runs(::Worksheet, c::Integer) = [Int(c):Int(c)]
_col_runs(::Worksheet, c::AbstractVector{<:Integer}) = _runs(c)
_col_runs(::Worksheet, c::ColumnRange) = [c.start:c.stop]
_col_runs(::Worksheet, c::CellRange) = [column_number(c.start):column_number(c.stop)]
_col_runs(::Worksheet, c::CellRef) = [column_number(c):column_number(c)]
_col_runs(::Worksheet, r::RowRange) = throw(XLSXError("Expected columns, found the rows `$r`."))
_col_runs(ws::Worksheet, s::AbstractString) = _runs(reduce(vcat, (collect(only(_col_runs(ws, p))) for p in _rowcol_parts(ws, s))))

_row_runs(::Worksheet, r::Integer) = [Int(r):Int(r)]
_row_runs(::Worksheet, r::AbstractVector{<:Integer}) = _runs(r)
_row_runs(::Worksheet, r::RowRange) = [r.start:r.stop]
_row_runs(::Worksheet, r::CellRange) = [row_number(r.start):row_number(r.stop)]
_row_runs(::Worksheet, r::CellRef) = [row_number(r):row_number(r)]
_row_runs(::Worksheet, c::ColumnRange) = throw(XLSXError("Expected rows, found the columns `$c`."))
_row_runs(ws::Worksheet, s::AbstractString) = _runs(reduce(vcat, (collect(only(_row_runs(ws, p))) for p in _rowcol_parts(ws, s))))

# The parts of a reference string, each a CellRef, CellRange, ColumnRange or
# RowRange. Defined names are resolved; a part on another sheet is an error.
function _rowcol_parts(ws::Worksheet, s::AbstractString)
    parts = Any[]
    for p in split(s, ',')
        p = strip(p)
        if is_worksheet_defined_name(ws, p) || is_workbook_defined_name(get_workbook(ws), p)
            v = is_worksheet_defined_name(ws, p) ? get_defined_name_value(ws, p) : get_defined_name_value(get_workbook(ws), p)
            is_defined_name_value_a_reference(v) || throw(XLSXError("Defined name `$p` is not a reference to cells."))
            for q in split(string(v), ',')
                if occursin('!', q)
                    sheet, q = rsplit(q, '!'; limit=2)
                    sheet = replace(strip(sheet, '\''), "''" => "'")
                    _same_sheetname(sheet, ws.name) || throw(XLSXError("Defined name `$p` refers to sheet `$sheet`, not `$(ws.name)`."))
                end
                push!(parts, _rowcol_part(replace(q, "\$" => "")))
            end
        else
            push!(parts, _rowcol_part(p))
        end
    end
    return parts
end

function _rowcol_part(p::AbstractString)
    is_valid_cellname(p)       && return CellRef(p)
    is_valid_cellrange(p)      && return CellRange(p)
    is_valid_column_range(p)   && return ColumnRange(p)
    is_valid_row_range(p)      && return RowRange(p)
    throw(XLSXError("Invalid row or column reference: `$p`."))
end

function _single(spans::Vector{UnitRange{Int}}, what::String)::Int
    (length(spans) == 1 && length(only(spans)) == 1) ||
        throw(XLSXError("Give a single $what, not $(join(spans, ", "))."))
    return first(only(spans))
end

_level(v) = isnothing(v) ? 0 : parse(Int, v)

# ===========================================================================
# Columns
# ===========================================================================

# Apply `f!` to the attributes of each `<col>` definition covering columns
# `left:right`, splitting definitions at the edges so the part outside keeps its
# settings and no two overlap (Excel repairs a file where they do). Columns with
# no definition get one, with the sheet's default width.
function _update_col_defs!(f!::Function, ws::Worksheet, left::Integer, right::Integer)
    left, right = Int(left), Int(right)
    sheetdoc = xmlroot(get_workbook(ws), ws.relationship_id)
    pfx = get_prefix(ws)
    pfx = pfx == "" ? pfx : pfx * ":"

    i, j = get_idces(sheetdoc, "worksheet", "cols")
    if isnothing(j)
        k, l = get_idces(sheetdoc, "worksheet", "sheetData")
        i != k && throw(XLSXError("Unexpected worksheet structure: <cols> parent index does not match <worksheet> index."))
        insert!(sheetdoc[k].children, l, XML.Element("$(pfx)cols"))
        j = l
    end

    covered = falses(right - left + 1)
    col_defs = Dict{String,String}[]
    for c in xml_elements(sheetdoc[i][j])
        atts = Dict{String,String}(XML.attributes(c))
        lo, hi = parse(Int, atts["min"]), parse(Int, get(atts, "max", atts["min"]))
        for (a, b, inside) in ((lo, min(hi, left - 1), false),
                               (max(lo, left), min(hi, right), true),
                               (max(lo, right + 1), hi, false))
            a <= b || continue
            d = merge(atts, Dict("min" => string(a), "max" => string(b)))
            inside && (f!(d); covered[(a:b) .- (left - 1)] .= true)
            push!(col_defs, d)
        end
    end

    k = 1
    while k <= length(covered)
        if covered[k]
            k += 1
            continue
        end
        e = something(findnext(covered, k), length(covered) + 1) - 1
        d = Dict("min" => string(left + k - 1), "max" => string(left + e - 1),
                 "width" => string(_default_col_width(sheetdoc[i])))
        f!(d)
        push!(col_defs, d)
        k = e + 1
    end

    # Rebuild <cols> in column order, as the schema requires, joining adjacent
    # definitions with the same settings.
    sort!(col_defs; by = d -> parse(Int, d["min"]))
    same(a, b) = all(k in ("min", "max") || get(b, k, nothing) == v for (k, v) in a) && length(a) == length(b)
    merged = Dict{String,String}[]
    for d in col_defs
        if !isempty(merged) && parse(Int, merged[end]["max"]) + 1 == parse(Int, d["min"]) && same(merged[end], d)
            merged[end]["max"] = d["max"]
        else
            push!(merged, d)
        end
    end
    col_defs = merged
    new_cols = unlink(sheetdoc[i][j], ("cols", "col"), pfx)
    for atts in col_defs
        new_col = XML.Element("$(pfx)col")
        new_col["min"] = atts["min"]
        new_col["max"] = atts["max"]
        for (k, v) in atts
            k in ("min", "max") || (new_col[k] = v)
        end
        push!(new_cols, new_col)
    end
    sheetdoc[i][j] = new_cols

    update_worksheets_xml!(get_xlsxfile(ws))
    return nothing
end

# The width of a column with no definition. Excel's is `defaultColWidth` if the
# sheet gives one, else `baseColWidth` (8 by default) plus the cell padding,
# 9.140625 for the default font.
function _default_col_width(root::XML.Node)::Float64
    for c in xml_elements(root)
        localname(c) == "sheetFormatPr" || continue
        w = get(c, "defaultColWidth", nothing)
        isnothing(w) || return parse(Float64, w)
        return parse(Float64, get(c, "baseColWidth", "8")) + 1.140625
    end
    return 9.140625
end

function _set_col_attr!(ws::Worksheet, spans, key::String, value::Union{Nothing,String})
    _require_writable(ws, "change columns")
    _require_worksheet(ws, "change columns")
    for s in spans
        _update_col_defs!(d -> isnothing(value) ? delete!(d, key) : (d[key] = value), ws, first(s), last(s))
    end
end

# The value of attribute `key` on the definition of column `col`, if any.
function _col_attr(ws::Worksheet, col::Int, key::String)
    _require_worksheet(ws, "read columns")
    for cols in _sheet_top_elements(get_xlsxfile(ws), _sheet_path(ws), ("cols",)), c in xml_elements(cols)
        parse(Int, c["min"]) <= col <= parse(Int, get(c, "max", c["min"])) && return get(c, key, nothing)
    end
    return nothing
end

function _col_levels(ws::Worksheet)
    out = Int[]
    for cols in _sheet_top_elements(get_xlsxfile(ws), _sheet_path(ws), ("cols",)), c in xml_elements(cols)
        push!(out, _level(get(c, "outlineLevel", nothing)))
    end
    return out
end

# ===========================================================================
# Rows
# ===========================================================================

# Put row `r` in the cache with no cells, if it isn't there already.
function _ensure_cache_row!(ws::Worksheet, r::Integer)
    r = Int(r)
    cache = ws.cache
    if !haskey(cache.cells, r)
        push!(cache.rows_in_cache, r)
        cache.cells[r] = Dict{Int,Cell}()
        cache.row_ht[r] = nothing
        cache.dirty = true
    end
    return nothing
end

# The cache, filled.
function _row_cache!(ws::Worksheet, action::String)
    _require_worksheet(ws, action)
    _require_cache(ws, action)
    isnothing(ws.cache) && eachrow(ws)
    return ws.cache
end

function _set_row_attr!(ws::Worksheet, spans, key::String, value::Union{Nothing,String})
    _require_writable(ws, "change rows")
    _row_cache!(ws, "change rows")
    for s in spans, r in s
        _set_row_attr!(ws, r, key, value)
    end
end

function _set_row_attr!(ws::Worksheet, r::Int, key::String, value::Union{Nothing,String})
    uh = ws.unhandled_attributes
    if isnothing(value)
        (isnothing(uh) || !haskey(uh, r)) && return
        delete!(uh[r], key)
        isempty(uh[r]) && (delete!(uh, r); _drop_row_if_empty!(ws, r))
    else
        isnothing(uh) && (uh = ws.unhandled_attributes = Dict{Int,Dict{String,String}}())
        _ensure_cache_row!(ws, r)
        get!(Dict{String,String}, uh, r)[key] = value
    end
    return nothing
end

# Remove a row with no cells, height or other attributes, so that showing or
# ungrouping a row leaves nothing behind.
function _drop_row_if_empty!(ws::Worksheet, r::Int)
    cache = ws.cache
    (haskey(cache.cells, r) && isempty(cache.cells[r]) && isnothing(cache.row_ht[r])) || return
    delete!(cache.cells, r)
    delete!(cache.row_ht, r)
    deleteat!(cache.rows_in_cache, findfirst(==(r), cache.rows_in_cache))
    cache.dirty = true
    return nothing
end

function _row_attr(ws::Worksheet, r::Int, key::String)
    _row_cache!(ws, "read rows")
    uh = ws.unhandled_attributes
    (isnothing(uh) || !haskey(uh, r)) && return nothing
    return get(uh[r], key, nothing)
end

_row_levels(ws::Worksheet) =
    isnothing(ws.unhandled_attributes) ? Int[] : [_level(get(a, "outlineLevel", nothing)) for a in values(ws.unhandled_attributes)]

# ===========================================================================
# Outlines
# ===========================================================================

# Group (`step=+1`) or ungroup (`step=-1`) each span of rows or columns.
function _outline!(ws::Worksheet, spans::Vector{UnitRange{Int}}, kind::Symbol, step::Int, collapsed::Bool)
    what = kind === :row ? "row" : "column"
    _require_writable(ws, "group $(what)s")
    _require_worksheet(ws, "group $(what)s")
    kind === :row && _row_cache!(ws, "group rows")
    level(i) = _level(kind === :row ? _row_attr(ws, i, "outlineLevel") : _col_attr(ws, i, "outlineLevel"))
    set!(spn, key, v) = kind === :row ? _set_row_attr!(ws, [spn], key, v) : _set_col_attr!(ws, [spn], key, v)
    maxn = kind === :row ? EXCEL_MAX_ROWS : EXCEL_MAX_COLS

    if step > 0
        for s in spans, i in s
            level(i) >= EXCEL_MAX_OUTLINE_LEVEL &&
                throw(XLSXError("Cannot group $what $i: it is already at outline level $EXCEL_MAX_OUTLINE_LEVEL, the most Excel allows."))
        end
    end

    after = _summary_after(ws, kind)
    for s in spans
        # Change each run of equal levels together, so columns are split only where needed.
        i = first(s)
        while i <= last(s)
            l = level(i)
            j = i
            while j < last(s) && level(j + 1) == l
                j += 1
            end
            n = max(l + step, 0)
            n == l || set!(i:j, "outlineLevel", n == 0 ? nothing : string(n))
            i = j + 1
        end

        collapsed && set!(s, "hidden", "1")
        summary = after ? last(s) + 1 : first(s) - 1
        1 <= summary <= maxn || continue       # a group at the edge of the sheet has no summary row
        if collapsed
            set!(summary:summary, "collapsed", "1")
        elseif step < 0 && level(after ? last(s) : first(s)) <= level(summary)
            # No group is left beside the summary row, so nothing for it to collapse.
            set!(summary:summary, "collapsed", nothing)
        end
    end

    _update_outline_levels!(ws)
    return nothing
end

# Whether summary rows are below their groups (summary columns to the right),
# from the sheet's `<sheetPr><outlinePr>`; Excel's default is yes.
function _summary_after(ws::Worksheet, kind::Symbol)::Bool
    key = kind === :row ? "summaryBelow" : "summaryRight"
    for pr in _sheet_top_elements(get_xlsxfile(ws), _sheet_path(ws), ("sheetPr",)), c in xml_elements(pr)
        localname(c) == "outlinePr" && return get(c, key, "1") ∉ ("0", "false")
    end
    return true
end

# Record the deepest row and column outline levels in `<sheetFormatPr>`, which
# Excel uses to draw the outline buttons.
function _update_outline_levels!(ws::Worksheet)
    root = xml_root_element(get_xml_data(get_xlsxfile(ws), _sheet_path(ws)))
    idx = findfirst(c -> localname(c) == "sheetFormatPr", XML.children(root))
    if isnothing(idx)
        pfx = get_prefix(ws)
        el = XML.Element((pfx == "" ? "" : pfx * ":") * "sheetFormatPr"; defaultRowHeight="15")
        idx = insert_index(root, "sheetFormatPr", WORKSHEET_ORDER) + 1
        insert!(XML.children(root), idx, el)
    end
    el = XML.children(root)[idx]
    atts = Dict{String,String}(something(XML.attributes(el), Dict{String,String}()))
    for (key, levels) in (("outlineLevelRow", _row_levels(ws)), ("outlineLevelCol", _col_levels(ws)))
        m = maximum(levels; init=0)
        m == 0 ? delete!(atts, key) : (atts[key] = string(m))
    end
    new_el = XML.Element(XML.tag(el))
    for (k, v) in atts
        new_el[k] = v
    end
    XML.children(root)[idx] = new_el
    update_worksheets_xml!(get_xlsxfile(ws))
    return nothing
end
