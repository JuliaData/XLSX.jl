# ===========================================================================
# AutoFilter criteria and sorting
#
# Excel doesn't apply a filter or a sort as it opens a file: it shows the rows as
# the file has them. So, as Excel does when it saves, XLSX.jl hides the rows a
# filter doesn't show (the rows' `hidden`) and puts the rows in sorted order,
# besides writing the criteria (`<filterColumn>`) and the sort (`<sortState>`).
#
# Each column's criterion is matched against what each cell holds and shows, as
# Excel matches it (checked against Excel 365 by COM):
#   - a list of values: the cell's shown text, in any case; a date by its day; and
#     the empty cells if `missing` is listed.
#   - conditions (`">" => 5`): numbers with numbers, text with text, and `*`, `?`
#     and `~` as wildcards in text.
#   - top or bottom items or percent: the numbers at or beyond the n-th.
#   - above or below average: strictly, against the numbers' mean.
#   - relative dates: whole days, weeks starting on Sunday, from today's date.
# ===========================================================================

const _FILTER_OPERATORS = Dict("==" => "equal", "=" => "equal", "!=" => "notEqual", "<>" => "notEqual",
    ">" => "greaterThan", ">=" => "greaterThanOrEqual", "<" => "lessThan", "<=" => "lessThanOrEqual")

const _DATE_PERIOD_FILTERS = (:today, :yesterday, :tomorrow, :thisWeek, :lastWeek, :nextWeek,
    :thisMonth, :lastMonth, :nextMonth, :thisQuarter, :lastQuarter, :nextQuarter,
    :thisYear, :lastYear, :nextYear, :yearToDate)
const _DATE_MONTH_FILTERS = Tuple(vcat([Symbol("Q$i") for i in 1:4], [Symbol("M$i") for i in 1:12]))
const _DYNAMIC_FILTERS = (:aboveAverage, :belowAverage, _DATE_PERIOD_FILTERS..., _DATE_MONTH_FILTERS...)

# What a cell holds, for filtering and sorting: its kind, in Excel's sort order (a
# number or date, text, a logical, an error, empty), its number (a date's serial,
# a logical's 0 or 1), and the text it shows.
struct _FilterCell
    kind::Int
    num::Float64
    text::String
end
const _NUMBER, _TEXT, _LOGICAL, _ERROR, _BLANK = 1, 2, 3, 4, 5

function _filter_cell(ws::Worksheet, cell::AbstractCell, formats::Dict)::_FilterCell
    cell isa EmptyCell && return _FilterCell(_BLANK, 0.0, "")
    dt = cell.datatype
    dt == CT_EMPTY && return _FilterCell(_BLANK, 0.0, "")
    (dt == CT_ERROR || dt == CT_RICHVALUE) && return _FilterCell(_ERROR, 0.0, _cell_text(ws, cell; formats))
    dt == CT_BOOL && return _FilterCell(_LOGICAL, Float64(cell.value != 0), _cell_text(ws, cell; formats))
    if dt == CT_STRING
        text = _cell_text(ws, cell; formats)
        return _FilterCell(isempty(text) ? _BLANK : _TEXT, 0.0, text)
    end
    num = (dt == CT_INT || dt == CT_DATE) ? Float64(reinterpret(Int64, cell.value)) : reinterpret(Float64, cell.value)
    return _FilterCell(_NUMBER, num, _cell_text(ws, cell; formats))
end

# The cells of one column of the data rows (all but the header row) of `rng`.
_filter_column_cells(ws::Worksheet, rng::CellRange, col::Int, formats::Dict) =
    [_filter_cell(ws, getcell(ws, r, col), formats) for r in row_number(rng.start)+1:row_number(rng.stop)]

# --- criteria ---------------------------------------------------------------

struct _ValueList
    texts::Vector{String}
    numbers::Vector{Float64}
    days::Vector{Dates.Date}
    blank::Bool
end
struct _Conditions
    conds::Vector{Tuple{String,Union{Float64,String}}}   # Excel's operator name and the value
    and::Bool
end
struct _TopItems
    top::Bool
    percent::Bool
    n::Int
end
struct _DynamicFilter
    type::Symbol
end

const _FILTER_HELP = "Give a value or a vector of values, a condition such as `\">\" => 5`, two joined by " *
    ":and or :or, `(:top, n)`, `(:bottom, n, :percent)`, or one of :aboveAverage, :belowAverage, :today, " *
    ":thisMonth and the other relative dates."

function _parse_criterion(x, is1904::Bool)
    x isa Symbol && return x in _DYNAMIC_FILTERS ? _DynamicFilter(x) :
        throw(XLSXError("Unknown filter `:$x`. " * _FILTER_HELP))
    x isa Pair && return _Conditions([_parse_condition(x, is1904)], false)
    if x isa Tuple && !isempty(x) && first(x) in (:top, :bottom)
        (length(x) == 2 || (length(x) == 3 && x[3] === :percent)) && x[2] isa Integer ||
            throw(XLSXError("A top or bottom filter is `(:top, n)` or `(:top, n, :percent)`. Found `$x`."))
        percent = length(x) == 3
        1 <= x[2] <= (percent ? 100 : 500) ||
            throw(XLSXError("A top or bottom filter takes 1 to $(percent ? "100 percent" : "500 items"). Found $(x[2])."))
        return _TopItems(first(x) === :top, percent, x[2])
    end
    if x isa Tuple && length(x) == 3 && x[2] in (:and, :or)
        return _Conditions([_parse_condition(x[1], is1904), _parse_condition(x[3], is1904)], x[2] === :and)
    end
    x isa Tuple && throw(XLSXError("Unknown filter `$x`. " * _FILTER_HELP))
    vals = x isa AbstractVector ? x : [x]
    isempty(vals) && throw(XLSXError("A list of values to show cannot be empty."))
    texts, numbers, days, blank = String[], Float64[], Dates.Date[], false
    for v in vals
        if ismissing(v)
            blank = true
        elseif v isa Bool
            push!(texts, v ? "TRUE" : "FALSE")
        elseif v isa Real
            push!(numbers, Float64(v))
        elseif v isa Union{Dates.Date,Dates.DateTime}
            push!(days, Dates.Date(v))
        elseif v isa AbstractString
            push!(texts, String(v))
        else
            throw(XLSXError("Cannot filter by a `$(typeof(v))`: `$v`. " * _FILTER_HELP))
        end
    end
    return _ValueList(texts, numbers, days, blank)
end

function _parse_condition(p, is1904::Bool)
    p isa Pair && first(p) isa AbstractString && haskey(_FILTER_OPERATORS, first(p)) ||
        throw(XLSXError("A condition is an operator and a value, such as `\">\" => 5`; the operators are " *
                        "$(join(sort(collect(keys(_FILTER_OPERATORS))), ", ")). Found `$p`."))
    v = last(p)
    val = v isa Bool ? throw(XLSXError("Use a list of values, `[true]`, to filter by a logical value.")) :
          v isa Real ? Float64(v) :
          v isa Dates.Date ? Float64(date_to_excel_value(v, is1904)) :
          v isa Dates.DateTime ? datetime_to_excel_value(v, is1904) :
          v isa AbstractString ? String(v) :
          throw(XLSXError("Cannot compare with a `$(typeof(v))`: `$v`."))
    return (_FILTER_OPERATORS[first(p)], val)
end

# Whether `pattern`, with Excel's wildcards (`*`, `?`, and `~` before one taken as
# itself), matches all of `text`, in any case.
function _wildcard_match(pattern::String, text::String)::Bool
    io = IOBuffer()
    escaped = false
    for ch in pattern
        if !escaped && ch == '~'
            escaped = true
            continue
        end
        if !escaped && ch == '*'
            print(io, ".*")
        elseif !escaped && ch == '?'
            print(io, ".")
        else
            ch in "\\^\$.|?*+()[]{}" && print(io, '\\')
            print(io, ch)
        end
        escaped = false
    end
    escaped && print(io, "~")
    return occursin(Regex("^" * String(take!(io)) * "\$", "is"), text)
end

function _condition_holds(c::_FilterCell, op::String, v::Union{Float64,String})::Bool
    if v isa Float64
        c.kind == _NUMBER || return op == "notEqual"
        op == "equal"              && return c.num == v
        op == "notEqual"           && return c.num != v
        op == "greaterThan"        && return c.num > v
        op == "greaterThanOrEqual" && return c.num >= v
        op == "lessThan"           && return c.num < v
        return c.num <= v
    end
    op == "equal"    && return _wildcard_match(v, c.text)
    op == "notEqual" && return !_wildcard_match(v, c.text)
    c.kind == _TEXT || return false
    a, b = lowercase(c.text), lowercase(v)
    op == "greaterThan"        && return a > b
    op == "greaterThanOrEqual" && return a >= b
    op == "lessThan"           && return a < b
    return a <= b
end

# The first and last-plus-one serial days of a relative date filter, from today.
function _date_period(type::Symbol, today::Dates.Date)
    week = today - Dates.Day(Dates.dayofweek(today) % 7)   # Sunday
    month, quarter, year = Dates.firstdayofmonth(today), Dates.firstdayofquarter(today), Dates.firstdayofyear(today)
    type === :today       && return today, today + Dates.Day(1)
    type === :yesterday   && return today - Dates.Day(1), today
    type === :tomorrow    && return today + Dates.Day(1), today + Dates.Day(2)
    type === :thisWeek    && return week, week + Dates.Week(1)
    type === :lastWeek    && return week - Dates.Week(1), week
    type === :nextWeek    && return week + Dates.Week(1), week + Dates.Week(2)
    type === :thisMonth   && return month, month + Dates.Month(1)
    type === :lastMonth   && return month - Dates.Month(1), month
    type === :nextMonth   && return month + Dates.Month(1), month + Dates.Month(2)
    type === :thisQuarter && return quarter, quarter + Dates.Month(3)
    type === :lastQuarter && return quarter - Dates.Month(3), quarter
    type === :nextQuarter && return quarter + Dates.Month(3), quarter + Dates.Month(6)
    type === :thisYear    && return year, year + Dates.Year(1)
    type === :lastYear    && return year - Dates.Year(1), year
    type === :nextYear    && return year + Dates.Year(1), year + Dates.Year(2)
    return year, today + Dates.Day(1)   # :yearToDate
end

_filter_number(x::Float64) = isinteger(x) && abs(x) < 1e15 ? string(Int64(x)) : string(x)

_filter_mean(cells) = (nums = [c.num for c in cells if c.kind == _NUMBER]; isempty(nums) ? nothing : sum(nums) / length(nums))

# The rows of one column a criterion shows, and its `<filterColumn>` child, as
# Excel writes it.
function _apply_criterion(ws::Worksheet, crit::_ValueList, cells::Vector{_FilterCell}, pfx::String)
    is1904 = isdate1904(get_workbook(ws))
    texts = Set(lowercase.(crit.texts))
    day(c) = excel_value_to_date(floor(Int, c.num), is1904)
    shows = [c.kind == _BLANK ? crit.blank :
             lowercase(c.text) in texts ||
             (c.kind == _NUMBER && (c.num in crit.numbers || (!isempty(crit.days) && c.num >= 0 && day(c) in crit.days)))
             for c in cells]
    # Excel lists numbers by the text they show, so the number's cells give it
    vals = copy(crit.texts)
    for n in crit.numbers
        shown = unique(c.text for c in cells if c.kind == _NUMBER && c.num == n)
        append!(vals, isempty(shown) ? [_filter_number(n)] : shown)
    end
    vals = sort(unique(lowercase, vals); by=lowercase)
    kids = XML.Node[_xml_element(pfx * "filter", ["val" => v]) for v in vals]
    for d in sort(unique(crit.days))
        push!(kids, _xml_element(pfx * "dateGroupItem", ["year" => string(Dates.year(d)), "month" => string(Dates.month(d)),
                                                        "day" => string(Dates.day(d)), "dateTimeGrouping" => "day"]))
    end
    return shows, _xml_element(pfx * "filters", crit.blank ? ["blank" => "1"] : Pair{String,String}[], kids)
end

function _apply_criterion(ws::Worksheet, crit::_Conditions, cells::Vector{_FilterCell}, pfx::String)
    holds(c) = (crit.and ? all : any)(_condition_holds(c, op, v) for (op, v) in crit.conds)
    kids = XML.Node[_xml_element(pfx * "customFilter",
                                 vcat(op == "equal" ? Pair{String,String}[] : ["operator" => op],
                                      ["val" => v isa Float64 ? _filter_number(v) : v]))
                    for (op, v) in crit.conds]
    return [holds(c) for c in cells], _xml_element(pfx * "customFilters", crit.and ? ["and" => "1"] : Pair{String,String}[], kids)
end

function _apply_criterion(ws::Worksheet, crit::_TopItems, cells::Vector{_FilterCell}, pfx::String)
    nums = sort!([c.num for c in cells if c.kind == _NUMBER]; rev=crit.top)
    atts = Pair{String,String}[]
    crit.top || push!(atts, "top" => "0")
    crit.percent && push!(atts, "percent" => "1")
    push!(atts, "val" => string(crit.n))
    isempty(nums) && return falses(length(cells)), _xml_element(pfx * "top10", atts)
    k = crit.percent ? clamp(floor(Int, length(nums) * crit.n / 100), 1, length(nums)) : min(crit.n, length(nums))
    edge = nums[k]
    push!(atts, "filterVal" => _filter_number(edge))
    return [c.kind == _NUMBER && (crit.top ? c.num >= edge : c.num <= edge) for c in cells], _xml_element(pfx * "top10", atts)
end

function _apply_criterion(ws::Worksheet, crit::_DynamicFilter, cells::Vector{_FilterCell}, pfx::String)
    type = crit.type
    atts = ["type" => string(type)]
    if type in (:aboveAverage, :belowAverage)
        avg = _filter_mean(cells)
        isnothing(avg) && return falses(length(cells)), _xml_element(pfx * "dynamicFilter", atts)
        push!(atts, "val" => _filter_number(avg))
        return [c.kind == _NUMBER && (type === :aboveAverage ? c.num > avg : c.num < avg) for c in cells],
               _xml_element(pfx * "dynamicFilter", atts)
    end
    is1904 = isdate1904(get_workbook(ws))
    if type in _DATE_MONTH_FILTERS
        name = string(type)
        i = parse(Int, name[2:end])
        months = name[1] == 'Q' ? (3i - 2:3i) : (i:i)
        return [c.kind == _NUMBER && c.num >= 0 && Dates.month(excel_value_to_date(floor(Int, c.num), is1904)) in months for c in cells],
               _xml_element(pfx * "dynamicFilter", atts)
    end
    lo, hi = (date_to_excel_value(d, is1904) for d in _date_period(type, Dates.today()))
    append!(atts, ["val" => string(lo), "maxVal" => string(hi)])
    return [c.kind == _NUMBER && lo <= c.num < hi for c in cells], _xml_element(pfx * "dynamicFilter", atts)
end

# --- columns ----------------------------------------------------------------

# The 0-based position in `rng` of column `col`: its position (1 for the first
# column), the text of its header cell, or its letter.
function _filter_column_id(ws::Worksheet, rng::CellRange, col, formats::Dict)::Int
    c1, c2 = column_number(rng.start), column_number(rng.stop)
    if col isa Integer
        1 <= col <= c2 - c1 + 1 || throw(XLSXError("Column $col is outside the filter's $(c2 - c1 + 1) columns (`$rng`)."))
        return col - 1
    end
    col isa AbstractString || throw(XLSXError("A filter column is given by its position, its header or its letter. Found `$col`."))
    hdr = row_number(rng.start)
    for c in c1:c2
        _cell_text(ws, getcell(ws, hdr, c); formats) == col && return c - c1
    end
    if occursin(r"^[A-Za-z]{1,3}$", col)
        c = decode_column_number(uppercase(col))
        c1 <= c <= c2 && return c - c1
    end
    throw(XLSXError("No column `$col` in `$rng`: it is neither a header nor a column letter in the range."))
end

_criteria_pairs(f::Pair) = [f]
_criteria_pairs(f::AbstractDict) = collect(f)
_criteria_pairs(f::AbstractVector) = (all(p -> p isa Pair, f) ||
    throw(XLSXError("Give filters as pairs of a column and its criterion, such as `\"Colour\" => [\"Red\"]`.")); f)
_criteria_pairs(f) = throw(XLSXError("Give filters as pairs of a column and its criterion, such as `\"Colour\" => [\"Red\"]`."))

# Each filtered column's 0-based position and criterion, in column order.
function _parse_filters(ws::Worksheet, rng::CellRange, filters, formats::Dict)
    out = Tuple{Int,Any}[]
    is1904 = isdate1904(get_workbook(ws))
    for (col, x) in _criteria_pairs(filters)
        id = _filter_column_id(ws, rng, col, formats)
        any(o -> first(o) == id, out) && throw(XLSXError("Column `$col` has more than one filter."))
        push!(out, (id, _parse_criterion(x, is1904)))
    end
    return sort!(out; by=first)
end

# Each sort key's 0-based column and whether it is descending.
function _parse_sort(ws::Worksheet, rng::CellRange, sort, formats::Dict)::Vector{Tuple{Int,Bool}}
    keys = sort isa AbstractVector ? sort : [sort]
    isempty(keys) && throw(XLSXError("`sort` needs at least one column."))
    length(keys) > 64 && throw(XLSXError("Excel sorts by at most 64 columns. Found $(length(keys))."))
    out = Tuple{Int,Bool}[]
    for k in keys
        col, order = k isa Pair ? (first(k), last(k)) : (k, :ascending)
        order in (:ascending, :descending) ||
            throw(XLSXError("A sort order is :ascending or :descending. Found `$(repr(order))`."))
        push!(out, (_filter_column_id(ws, rng, col, formats), order === :descending))
    end
    return out
end

# --- sorting ----------------------------------------------------------------

# Excel's order: numbers, text (in any case), logicals, errors, and empty cells
# last, whichever the direction.
function _sort_before(a::_FilterCell, b::_FilterCell, descending::Bool)::Union{Nothing,Bool}
    if a.kind != b.kind
        a.kind == _BLANK && return false
        b.kind == _BLANK && return true
        return descending ? a.kind > b.kind : a.kind < b.kind
    end
    a.kind in (_ERROR, _BLANK) && return nothing
    if a.kind == _TEXT
        x, y = lowercase(a.text), lowercase(b.text)
        x == y && return nothing
        return descending ? x > y : x < y
    end
    a.num == b.num && return nothing
    return descending ? a.num > b.num : a.num < b.num
end

# Put the data rows of `rng` in the order of `keys`, moving each row's cells, with
# their formats, formulas (as Excel moves them, as if copied), hyperlinks and
# notes. Cells outside the range's columns stay.
function _sort_rows!(ws::Worksheet, rng::CellRange, keys::Vector{Tuple{Int,Bool}}, formats::Dict)
    r1, r2 = row_number(rng.start) + 1, row_number(rng.stop)
    c1, c2 = column_number(rng.start), column_number(rng.stop)
    r2 > r1 || return nothing
    area = CellRange(CellRef(r1, c1), CellRef(r2, c2))
    _check_sortable(ws, area)

    cols = Dict(k => _filter_column_cells(ws, rng, c1 + k, formats) for (k, _) in keys)
    function before(i, j)
        for (k, desc) in keys
            b = _sort_before(cols[k][i], cols[k][j], desc)
            isnothing(b) || return b
        end
        return false
    end
    perm = sortperm(1:r2-r1+1; lt=before)   # stable: equal rows keep their order
    perm == 1:r2-r1+1 && return nothing
    new_row = Dict(r1 + perm[i] - 1 => r1 + i - 1 for i in eachindex(perm))

    # Take every cell and formula out, each formula as its own text (a shared one
    # read before any of its cells goes), then put them back in their new rows
    wb = get_workbook(ws)
    taken = Tuple{Cell,Union{Nothing,AbstractFormula},String}[]
    for r in r1:r2, c in c1:c2
        cell = getcell(ws, r, c)
        cell isa Cell || continue
        f = cell.formula ? get_formula_from_cache(ws, cell.ref) : nothing
        text = isnothing(f) ? "" : f isa FormulaReference ? get_referenced_formula(ws, cell.ref) : f.formula
        push!(taken, (cell, f, text))
    end
    for (cell, f, _) in taken
        isnothing(f) || delete!(wb.formulas, SheetCellRef(ws.name, cell.ref))
        delete!(ws.cache.cells[row_number(cell.ref)], column_number(cell.ref))
    end
    was_hidden = Set(r for r in r1:r2 if isRowHidden(ws, r))   # a row's filter state goes with it
    for (cell, f, text) in taken
        old = cell.ref
        ref = CellRef(new_row[row_number(old)], column_number(old))
        setdata!(ws, Cell(ref, cell.value, cell.style, cell.meta, cell.datatype, cell.formula))
        isnothing(f) && continue
        moved = shift_excel_references(text, (row_number(ref) - row_number(old), 0))
        add_formula_to_cache(ws, ref, f isa Formula ?
            Formula(moved, f.type, isnothing(f.ref) ? nothing : string(ref), f.unhandled) : Formula(moved))
    end
    hide = [new_row[r] for r in was_hidden]
    show = [r for r in r1:r2 if r in was_hidden && !(r in hide)]
    isempty(hide) || setRowHidden(ws, hide; hidden=true)
    isempty(show) || setRowHidden(ws, show; hidden=false)
    _move_sorted_hyperlinks!(ws, area, new_row)
    _move_sorted_notes!(ws, area, new_row)
    return nothing
end

# Refuse what a sort can't move: merged cells, formulas over several cells (array
# formulas, and shared formulas reaching outside the rows), and hyperlinks over
# several rows. Excel refuses the first too.
function _check_sortable(ws::Worksheet, area::CellRange)
    for m in something(getMergedCells(ws), CellRange[])
        intersects(m, area) && throw(XLSXError("Cannot sort `$area`: it has merged cells (`$m`)."))
    end
    for r in row_number(area.start):row_number(area.stop), c in column_number(area.start):column_number(area.stop)
        cell = getcell(ws, r, c)
        cell isa Cell && cell.formula || continue
        f = get_formula_from_cache(ws, cell.ref)
        spans = f isa Formula && !isnothing(f.ref) && f.ref != string(cell.ref) && size(CellRange(f.ref)) != (1, 1)
        spans && throw(XLSXError("Cannot sort `$area`: cell $(cell.ref) has an array formula over `$(f.ref)`."))
        f isa ReferencedFormula && !issubset(CellRange(f.ref), area) &&
            throw(XLSXError("Cannot sort `$area`: the formula in $(cell.ref) is shared with cells outside it (`$(f.ref)`)."))
    end
    for n in _hyperlink_nodes(get_xlsxfile(ws), _sheet_path(ws))
        h = _hyperlink_range(get_attr(n, "ref"))
        intersects(h, area) || continue
        row_number(h.start) == row_number(h.stop) && issubset(h, CellRange(CellRef(row_number(h.start), column_number(area.start)),
                                                                        CellRef(row_number(h.start), column_number(area.stop)))) ||
            throw(XLSXError("Cannot sort `$area`: the hyperlink on `$(get_attr(n, "ref"))` covers cells that would move apart."))
    end
    return nothing
end

# A reference `ref` in the sorted `area`, in its new row.
function _sorted_ref(ref::AbstractString, area::CellRange, new_row::Dict{Int,Int})::String
    rng = _hyperlink_range(ref)
    issubset(rng, area) || return String(ref)
    d = new_row[row_number(rng.start)] - row_number(rng.start)
    moved = CellRange(CellRef(row_number(rng.start) + d, column_number(rng.start)), CellRef(row_number(rng.stop) + d, column_number(rng.stop)))
    return _hyperlink_refstring(moved)
end

function _move_sorted_hyperlinks!(ws::Worksheet, area::CellRange, new_row::Dict{Int,Int})
    for n in _hyperlink_nodes(get_xlsxfile(ws), _sheet_path(ws))
        n["ref"] = _sorted_ref(get_attr(n, "ref"), area, new_row)
    end
end

# Notes, threads and the notes' boxes move with their cells, as in Excel.
function _move_sorted_notes!(ws::Worksheet, area::CellRange, new_row::Dict{Int,Int})
    xf, sheet_path = get_xlsxfile(ws), _sheet_path(ws)
    move!(n) = (n["ref"] = _sorted_ref(get_attr(n, "ref"), area, new_row))
    comments_path = _comments_path(xf, sheet_path)
    isnothing(comments_path) || foreach(move!, _comment_nodes(xml_root_element(get_xml_data(xf, comments_path))))
    for (_, path) in _sheet_rel_targets(xf, sheet_path, REL_THREADED_COMMENT)
        haskey(xf.data, path) && foreach(move!, elements_with_tag(xml_root_element(get_xml_data(xf, path)), "threadedComment"))
    end
    vml_path = _legacy_vml_path(xf, sheet_path)
    (isnothing(vml_path) || !haskey(xf.binary_data, vml_path)) && return nothing
    vml = replace(_vml_string(xf, vml_path), _VML_SHAPE_RE => function (s)
        occursin(r"ObjectType\s*=\s*[\"']Note[\"']", s) || return s
        r, c = match(_VML_ROW, s), match(_VML_COLUMN, s)
        (isnothing(r) || isnothing(c)) && return s
        old = parse(Int, r[2]) + 1
        CellRef(old, parse(Int, c[2]) + 1) in area || return s
        d = new_row[old] - old
        s = replace(s, _VML_ROW => m -> (x = match(_VML_ROW, m); x[1] * string(old + d - 1) * x[3]))
        return replace(s, _VML_ANCHOR => function (m)
            x = match(_VML_ANCHOR, m)
            v = [parse(Int, strip(t)) for t in split(x[2], ',')]
            length(v) == 8 || return m
            v[3] += d; v[7] += d
            return x[1] * join(v, ", ") * x[3]
        end)
    end)
    xf.binary_data[vml_path] = Vector{UInt8}(codeunits(vml))
    return nothing
end
