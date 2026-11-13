# ===========================================================================
# AutoFilter on a plain range
#
# A sheet has at most one `<autoFilter ref="...">`, giving the drop-down filter
# buttons on the first row of the range. Excel also keeps the range in a hidden
# name scoped to the sheet, `_xlnm._FilterDatabase`. An Excel Table has a filter
# of its own, in the table part, so a sheet filter may not overlap a table.
#
# A filter's criteria are its `<filterColumn>`s, and a sort its `<sortState>`, which
# Excel writes both in the `<autoFilter>` and after it. A sheet whose filter has
# criteria is marked `filterMode="1"` in its `<sheetPr>`. Criteria and sorting are
# in filtersort.jl.
# ===========================================================================

const FILTER_DATABASE = "_xlnm._FilterDatabase"

"""
    setAutoFilter(ws::Worksheet, rng; filters=nothing, sort=nothing)
    setAutoFilter(xf::XLSXFile, "Sheet1!A1:D20"; kw...)

Add filter buttons to the first row of `rng`, as Excel's *Data → Filter* does,
so its columns can be filtered and sorted in Excel. `rng` is a cell range
(`"A1:D20"`), a defined name, or `:` for the sheet's whole used range. Its first
row is the header row.

A sheet has one filter, so this replaces any filter the sheet already has. The
filter may not overlap an Excel Table, which has a filter of its own (see
[`addtable!`](@ref)). Setting the filter again on the same range keeps its
criteria and sort, unless new ones are given.

# Filtering

`filters` gives criteria for some of the columns, as pairs of a column and a
criterion. A column is given by its header text, its column letter or its position
in the range (`1` for the first). Rows a criterion doesn't show are hidden, and the
others shown, so the file opens filtered, as Excel would have saved it. A row is
shown if it meets the criteria of every filtered column. A criterion is one of:

- a value, or a vector of values, to show: text (in any case, as the cell shows
  it), numbers, dates (by the day), `true`/`false`, and `missing` for empty cells.
- a condition, `op => value`, where `op` is `"=="`, `"!="`, `">"`, `">="`, `"<"`
  or `"<="` and `value` a number, a date or text. Text may use Excel's wildcards:
  `*` for any characters, `?` for any one, and `~` before one of them for itself,
  so `"==" => "A*"` shows text beginning with A.
- two conditions joined by `:and` or `:or`: `(">=" => 10, :and, "<" => 20)`.
- `(:top, n)` or `(:bottom, n)` for the `n` largest or smallest numbers, or
  `(:top, n, :percent)` for the top `n` percent.
- `:aboveAverage` or `:belowAverage`.
- a date relative to today: `:today`, `:yesterday`, `:tomorrow`, `:thisWeek`,
  `:lastWeek`, `:nextWeek`, `:thisMonth`, `:lastMonth`, `:nextMonth`, `:thisQuarter`,
  `:lastQuarter`, `:nextQuarter`, `:thisYear`, `:lastYear`, `:nextYear` or
  `:yearToDate`; or a quarter or month of any year, `:Q1` to `:Q4` or `:M1` to `:M12`.

`filters=[]` removes the criteria and shows every row of the range.

# Sorting

`sort` gives a column, `column => :ascending` or `column => :descending`, or a
vector of these, the first sorting first. The rows below the header are put in
that order, as Excel's sort puts them: numbers and dates, then text (in any case),
then `TRUE`/`FALSE`, then errors, and empty cells last in either order. Each row's
cells in the range move with their formats, notes and hyperlinks, and formulas
move as Excel moves them, as if copied, so `=B5*2` in row 5 becomes `=B2*2` in row
2. Rows a filter hides are sorted too, and stay hidden. Cells outside the range's
columns stay where they are. A range with merged
cells, or with array formulas or shared formulas reaching outside it, can't be
sorted. The sort is recorded, so Excel shows the sort arrow on the header.

The file must be open for writing (`mode="rw"`).

# Examples
```julia
julia> XLSX.setAutoFilter(sheet, "A1:F200")

julia> XLSX.setAutoFilter(sheet, :)

julia> XLSX.setAutoFilter(sheet, "A1:D50";
           filters = ["Colour" => ["Red", "Blue"], "Amount" => (">" => 100)],
           sort = "Amount" => :descending)
```

See also [`getAutoFilter`](@ref) and [`removeAutoFilter`](@ref).
"""
function setAutoFilter end
setAutoFilter(xf::XLSXFile, ref::AbstractString; kw...) = setAutoFilter(_rowcol_sheet(xf, ref)...; kw...)
setAutoFilter(ws::Worksheet, ::Colon; kw...) = setAutoFilter(ws, get_dimension(ws); kw...)
setAutoFilter(ws::Worksheet, rng::SheetCellRange; kw...) = do_sheet_names_match(ws, rng) && setAutoFilter(ws, rng.rng; kw...)
function setAutoFilter(ws::Worksheet, ref::AbstractString; kw...)
    parts = _rowcol_parts(ws, ref)
    (length(parts) == 1 && only(parts) isa CellRange) ||
        throw(XLSXError("An AutoFilter needs a single range of cells, such as `A1:D20`. Found `$ref`."))
    return setAutoFilter(ws, only(parts); kw...)
end
function setAutoFilter(ws::Worksheet, rng::CellRange; filters=nothing, sort=nothing)
    _require_writable(ws, "set an AutoFilter")
    _require_worksheet(ws, "set an AutoFilter")
    for t in tables(ws)
        intersects(rng, t.ref) &&
            throw(XLSXError("AutoFilter range `$rng` overlaps table `$(t.name)` (`$(t.ref)`), which has a filter of its own."))
    end
    formats  = Dict{String,Union{Nothing,Vector{_FmtSection}}}()
    criteria = isnothing(filters) ? nothing : _parse_filters(ws, rng, filters, formats)
    keys     = isnothing(sort) ? nothing : _parse_sort(ws, rng, sort, formats)
    isnothing(keys) || _sort_rows!(ws, rng, keys, formats)

    root = _sheet_root(ws)
    chn  = XML.children(root)
    idx  = findfirst(c -> localname(c) == "autoFilter", chn)
    old  = isnothing(idx) ? nothing : chn[idx]
    same = !isnothing(old) && get(old, "ref", "") == string(rng)   # the same range keeps what isn't given
    pfx  = something(_ns_prefix(XML.tag(root)), "")

    columns = isnothing(criteria) ?
        (same ? [c for c in xml_elements(old) if localname(c) == "filterColumn"] : XML.Node[]) :
        _apply_filters!(ws, rng, criteria, formats, pfx)
    state = isnothing(keys) ?
        (same ? first_element_with_tag(old, "sortState") : nothing) :
        _sort_state(rng, keys, pfx)

    el = _container(pfx * "autoFilter", same ? _atts_vector(old) : ["ref" => string(rng)])
    foreach(c -> push!(el, c), columns)
    isnothing(state) || push!(el, state)
    same && foreach(c -> localname(c) in ("filterColumn", "sortState") || push!(el, c), xml_elements(old))   # an extLst
    isnothing(idx) ? insert!(chn, insert_index(root, "autoFilter", WORKSHEET_ORDER) + 1, el) : (chn[idx] = el)
    if !(same && isnothing(keys))   # Excel writes the sort after the filter too
        _remove_sheet_element!(get_xlsxfile(ws), _sheet_path(ws), "sortState")
        isnothing(state) || insert!(chn, insert_index(root, "sortState", WORKSHEET_ORDER) + 1, copynode(state))
    end
    _set_filter_mode!(ws, !isempty(columns))
    get_workbook(ws).worksheet_names[(ws.sheetId, FILTER_DATABASE)] =
        DefinedNameValue(SheetCellRange(ws.name, rng), true, true)

    update_worksheets_xml!(get_xlsxfile(ws))
    return nothing
end

# Hide the data rows of `rng` the criteria don't show, and show the others. Returns
# the `<filterColumn>`s.
function _apply_filters!(ws::Worksheet, rng::CellRange, criteria, formats::Dict, pfx::String)::Vector{XML.Node}
    r1, c1 = row_number(rng.start), column_number(rng.start)
    shows = trues(row_number(rng.stop) - r1)
    columns = XML.Node[]
    for (id, crit) in criteria
        s, child = _apply_criterion(ws, crit, _filter_column_cells(ws, rng, c1 + id, formats), pfx)
        shows .&= s
        push!(columns, _xml_element(pfx * "filterColumn", ["colId" => string(id)], [child]))
    end
    rows = r1+1:row_number(rng.stop)
    hidden = [r for (r, s) in zip(rows, shows) if !s]
    isempty(hidden) || setRowHidden(ws, hidden; hidden=true)
    shown = [r for (r, s) in zip(rows, shows) if s && isRowHidden(ws, r)]
    isempty(shown) || setRowHidden(ws, shown; hidden=false)
    return columns
end

# The `<sortState>` for sorting the data rows of `rng` by `keys`.
function _sort_state(rng::CellRange, keys::Vector{Tuple{Int,Bool}}, pfx::String)::XML.Node
    r1, r2, c1 = row_number(rng.start) + 1, row_number(rng.stop), column_number(rng.start)
    data = CellRange(CellRef(r1, c1), CellRef(max(r1, r2), column_number(rng.stop)))
    conditions = [_xml_element(pfx * "sortCondition",
                               vcat(desc ? ["descending" => "1"] : Pair{String,String}[],
                                    ["ref" => string(CellRange(CellRef(r1, c1 + k), CellRef(max(r1, r2), c1 + k)))]))
                  for (k, desc) in keys]
    return _xml_element(pfx * "sortState", ["ref" => string(data)], conditions)
end

# Mark the sheet as filtered (`<sheetPr filterMode="1">`), or not.
function _set_filter_mode!(ws::Worksheet, on::Bool)
    root = _sheet_root(ws)
    chn  = XML.children(root)
    idx  = findfirst(c -> localname(c) == "sheetPr", chn)
    if on
        if isnothing(idx)
            pushfirst!(chn, _xml_element(something(_ns_prefix(XML.tag(root)), "") * "sheetPr", ["filterMode" => "1"]))
        else
            chn[idx]["filterMode"] = "1"
        end
    elseif !isnothing(idx)
        pr = chn[idx]
        remove_attr!(pr, "filterMode")
        isempty(_atts_vector(pr)) && isempty(xml_elements(pr)) && deleteat!(chn, idx)
    end
    return nothing
end

"""
    getAutoFilter(ws::Worksheet) -> Union{Nothing, CellRange}

The range of the sheet's filter (see [`setAutoFilter`](@ref)), or `nothing` if
it has none. The filters of Excel Tables are not included.
"""
function getAutoFilter(ws::Worksheet)::Union{Nothing,CellRange}
    is_chartsheet(get_workbook(ws), ws.name) && return nothing
    for el in _sheet_top_elements(get_xlsxfile(ws), _sheet_path(ws), ("autoFilter",))
        ref = get(el, "ref", nothing)
        isnothing(ref) || return CellRange(ref)
    end
    return nothing
end

"""
    removeAutoFilter(ws::Worksheet) -> Bool

Remove the sheet's filter, its criteria and its sort. Returns `false` if it had
none. Rows the filter hid stay hidden; show them with [`setRowHidden`](@ref).

The file must be open for writing (`mode="rw"`).
"""
function removeAutoFilter(ws::Worksheet)::Bool
    _require_writable(ws, "remove an AutoFilter")
    _require_worksheet(ws, "remove an AutoFilter")
    delete!(get_workbook(ws).worksheet_names, (ws.sheetId, FILTER_DATABASE))
    root = _sheet_root(ws)
    idx = findfirst(c -> localname(c) == "autoFilter", XML.children(root))
    isnothing(idx) && return false
    deleteat!(XML.children(root), idx)
    _remove_sheet_element!(get_xlsxfile(ws), _sheet_path(ws), "sortState")
    _set_filter_mode!(ws, false)
    update_worksheets_xml!(get_xlsxfile(ws))
    return true
end
