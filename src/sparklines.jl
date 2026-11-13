#
# -- Sparklines: small charts in single cells, kept in an x14 extension of the worksheet
#

const X14_NS = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"
const XM_NS = "http://schemas.microsoft.com/office/excel/2006/main"
const SPARKLINE_EXT_URI = "{05C60535-1F16-4fd2-B633-F4F36F0B64E0}"

# The worksheet extensions in the order Excel writes them in `<extLst>`.
const WORKSHEET_EXT_ORDER = uppercase.([
    X14_CF_EXT_URI,                                  # x14 conditional formats
    "{CCE6A557-97BC-4b89-ADB6-D9C93CAAB3DF}",        # x14 data validations
    SPARKLINE_EXT_URI,
    "{A8765BA9-456A-4dab-B4F3-ACF838C121DE}",        # slicers
    "{FC87AEE6-9EDD-4A0A-B7FB-166176984837}",        # protected ranges
    "{01252117-D84E-4E92-8308-4BE1C098FCBB}",        # ignored errors
    "{F7C9EE02-42E1-4005-9D12-6889AFFD525C}",        # web extensions
    "{3A4CF648-6AED-40f4-86FF-DC5316D8AED3}",        # x15 slicers
    "{7E03D99C-DC04-49d9-9315-930204A7B6E9}",        # timelines
])

# The worksheet's `<ext>` with `uri`, or `nothing`.
function _worksheet_ext(root::XML.Node, uri::AbstractString)
    i = findfirst(c -> localname(c) == "extLst", xml_elements(root))
    isnothing(i) && return nothing
    for e in xml_elements(xml_elements(root)[i])
        localname(e) == "ext" && uppercase(get(e, "uri", "")) == uppercase(uri) && return e
    end
    return nothing
end

# As `_worksheet_ext`, making `<extLst>` and the `<ext>` if needed, the `<ext>` among the
# others in the order Excel writes them.
function _worksheet_ext!(root::XML.Node, uri::AbstractString)
    ext = _worksheet_ext(root, uri)
    isnothing(ext) || return ext
    pfx = something(_ns_prefix(XML.tag(root)), "")
    chn = XML.children(root)
    i = findfirst(c -> XML.nodetype(c) == XML.Element && localname(c) == "extLst", chn)
    if isnothing(i)
        push!(chn, XML.Element(pfx * "extLst"))   # `extLst` is last in `WORKSHEET_ORDER`
        i = length(chn)
    end
    extlst = chn[i]
    ext = XML.Element(pfx * "ext")
    ext["uri"] = uri          # in the order Excel writes them
    ext["xmlns:x14"] = X14_NS
    rank(e) = findfirst(==(uppercase(get(e, "uri", ""))), WORKSHEET_EXT_ORDER)
    mine = rank(ext)
    kids = XML.children(extlst)
    j = isnothing(mine) ? nothing :
        findfirst(c -> XML.nodetype(c) == XML.Element && (r = rank(c); !isnothing(r) && r > mine), kids)
    isnothing(j) ? push!(extlst, ext) : insert!(kids, j, ext)
    return ext
end

# The sparkline types, as `type` gives them and as the file keeps them.
const SPARKLINE_TYPES = (line = nothing, column = "column", winloss = "stacked")
# How empty cells are shown; Excel writes `gap`, its own default, and leaves out `zero`, the schema's.
const _SPARKLINE_EMPTY = (gap = "gap", zero = nothing, span = "span")
# The group's colours: keyword, element and Excel's default.
const _SPARKLINE_COLORS = (seriesColor = ("colorSeries", "FF376092"), negativeColor = ("colorNegative", "FFD00000"),
                           axisColor = ("colorAxis", "FF000000"), markersColor = ("colorMarkers", "FFD00000"),
                           firstColor = ("colorFirst", "FFD00000"), lastColor = ("colorLast", "FFD00000"),
                           highColor = ("colorHigh", "FFD00000"), lowColor = ("colorLow", "FFD00000"))
# The on/off options: keyword and attribute.
const _SPARKLINE_FLAGS = (markers = "markers", high = "high", low = "low", first = "first", last = "last",
                          negative = "negative", axis = "displayXAxis", showHidden = "displayHidden",
                          rightToLeft = "rightToLeft")
const _SPARKLINE_KEYWORDS = (:type, :emptyCells, :lineWeight, :min, :max, :dateRange,
                             keys(_SPARKLINE_FLAGS)..., keys(_SPARKLINE_COLORS)...)
# The group's attributes in the schema's order.
const _SPARKLINE_GROUP_ATT_ORDER = ("manualMax", "manualMin", "lineWeight", "type", "dateAxis", "displayEmptyCellsAs",
    "markers", "high", "low", "first", "last", "negative", "displayXAxis", "displayHidden",
    "minAxisType", "maxAxisType", "rightToLeft")

"""
    addSparklines(ws::Worksheet, location, data; type=:line, kw...)

Add a group of sparklines, small charts each in one cell, as Excel's *Insert → Sparklines*
does. `location` is the cell, or the row or column of cells, to hold them, and `data` the
cells they chart, on this sheet or another (`"Data!B2:M13"`):

- one cell: one sparkline of a row or column of data;
- a column of cells: one sparkline for each row of `data`, which has as many rows;
- a row of cells: one sparkline for each column of `data`, which has as many columns.

A sparkline already in one of the cells is replaced. The sparklines of one call form a
group, sharing these options:

| Keyword | Meaning | Default |
|:--------|:--------|:--------|
| `type` | `:line`, `:column` or `:winloss` | `:line` |
| `markers` | mark every point (lines only) | `false` |
| `high`, `low`, `first`, `last` | mark the highest, lowest, first and last points | `false` |
| `negative` | mark negative points | `true` for `:winloss` |
| `axis` | show the horizontal axis | `false` |
| `emptyCells` | show empty cells as a `:gap`, as `:zero`, or `:span` them with a line | `:gap` |
| `showHidden` | chart the data in hidden rows and columns | `false` |
| `lineWeight` | the line's width in points | `0.75` |
| `min`, `max` | the vertical axis's limits: `:individual` for each sparkline, `:group` for all alike, or a number | `:individual` |
| `rightToLeft` | plot from right to left | `false` |
| `dateRange` | cells of dates to space the points by, as `data` is given | none |
| `seriesColor`, `negativeColor`, `axisColor`, `markersColor`, `firstColor`, `lastColor`, `highColor`, `lowColor` | colours | Excel's |

The colours can be $COLOR_FORMS_DOC

The file must be open for writing (`mode="rw"`).

# Examples
```julia
julia> XLSX.addSparklines(sheet, "N2:N13", "B2:M13")                 # one per row

julia> XLSX.addSparklines(sheet, "B14:M14", "B2:M13"; type=:column, high=true, highColor="red")

julia> XLSX.addSparklines(sheet, "P2", "Data!A1:A50"; markers=true, emptyCells=:span)
```

See also [`getSparklines`](@ref) and [`deleteSparklines`](@ref).
"""
function addSparklines(ws::Worksheet, location::Union{AbstractString,CellRef,CellRange},
                       data::Union{AbstractString,CellRange}; kw...)
    _require_writable(ws, "add sparklines")
    _require_worksheet(ws, "add sparklines")
    for k in keys(kw)
        k in _SPARKLINE_KEYWORDS || throw(XLSXError("Unknown keyword `$k`. Must be one of $(join(_SPARKLINE_KEYWORDS, ", "))."))
    end
    cells = _sparkline_cells(ws, location)
    sheet, rng = _sparkline_data(ws, data, "data")
    pairs = _sparkline_pairs(cells, rng, location)
    group = _sparkline_group(ws, [c => _sparkline_ref(sheet, r) for (c, r) in pairs]; kw...)

    root = _sheet_root(ws)
    _remove_sparklines!(root, Set(cells), false)
    ext = _worksheet_ext!(root, SPARKLINE_EXT_URI)
    block = findfirst(c -> localname(c) == "sparklineGroups", xml_elements(ext))
    if isnothing(block)
        b = XML.Element("x14:sparklineGroups", group)
        b["xmlns:xm"] = XM_NS
        push!(ext, b)
    else
        b = xml_elements(ext)[block]
        kids = XML.children(b)
        j = findfirst(c -> XML.nodetype(c) == XML.Element, kids)
        insert!(kids, something(j, length(kids) + 1), group)   # Excel puts a new group first
    end
    update_worksheets_xml!(get_xlsxfile(ws))
    return nothing
end

# The cells of `location`: one cell, or one row or column of them.
function _sparkline_cells(ws::Worksheet, location)::Vector{CellRef}
    rngs = _cell_region(ws, location isa CellRange ? location : location isa CellRef ? CellRange(location, location) : location)
    length(rngs) == 1 || throw(XLSXError("Give the sparklines' cells as one cell, or one row or column of cells. Found `$location`."))
    r = only(rngs)
    (r.start.row_number == r.stop.row_number || r.start.column_number == r.stop.column_number) ||
        throw(XLSXError("Sparklines go in one row or one column of cells, not `$r`."))
    return vec(collect(r))
end

# The sheet and range of a data reference such as `"B2:M13"` or `"'My Data'!B2:M13"`.
function _sparkline_data(ws::Worksheet, data, what::String)::Tuple{String,CellRange}
    data isa CellRange && return (ws.name, data)
    s = strip(replace(data, "\$" => ""))
    sheet = ws.name
    i = findlast('!', s)
    if !isnothing(i)
        sheet = String(s[1:prevind(s, i)])
        startswith(sheet, '\'') && endswith(sheet, '\'') && (sheet = replace(sheet[2:end-1], "''" => "'"))
        s = s[nextind(s, i):end]
        hassheet(get_workbook(ws), sheet) || throw(XLSXError("The sparklines' $what refers to sheet `$sheet`, which isn't in the workbook."))
        sheet = getsheet(get_workbook(ws), sheet).name
    end
    is_valid_cellname(s) && return (sheet, CellRange(CellRef(s), CellRef(s)))
    is_valid_cellrange(s) || throw(XLSXError("The sparklines' $what must be a range of cells, such as `B2:M13`. Found `$data`."))
    return (sheet, CellRange(s))
end

_sparkline_ref(sheet::AbstractString, r::CellRange) = quoteit(sheet) * "!" * string(r)

# Each sparkline's cell and data, as Excel pairs them.
function _sparkline_pairs(cells::Vector{CellRef}, rng::CellRange, location)::Vector{Pair{CellRef,CellRange}}
    nr = rng.stop.row_number - rng.start.row_number + 1
    nc = rng.stop.column_number - rng.start.column_number + 1
    rows() = [CellRange(CellRef(r, rng.start.column_number), CellRef(r, rng.stop.column_number)) for r in rng.start.row_number:rng.stop.row_number]
    cols() = [CellRange(CellRef(rng.start.row_number, c), CellRef(rng.stop.row_number, c)) for c in rng.start.column_number:rng.stop.column_number]
    n = length(cells)
    if n == 1
        (nr == 1 || nc == 1) ||
            throw(XLSXError("One sparkline charts one row or column of data, not `$rng`. Give a cell for each row or column."))
        return [only(cells) => rng]
    end
    vertical = first(cells).column_number == last(cells).column_number
    parts = vertical ? (nr == n ? rows() : nc == n ? cols() : nothing) :
                       (nc == n ? cols() : nr == n ? rows() : nothing)
    isnothing(parts) && throw(XLSXError(
        "The data `$rng` has $nr rows and $nc columns, but there are $n cells in `$location`: give one row or column of data for each sparkline."))
    return [c => p for (c, p) in zip(cells, parts)]
end

# A `<x14:sparklineGroup>` of these sparklines (cell => data reference) with the given options.
function _sparkline_group(ws::Worksheet, sparks::Vector{Pair{CellRef,String}};
                          type::Symbol=:line, emptyCells::Symbol=:gap, lineWeight::Real=0.75,
                          min::Union{Symbol,Real}=:individual, max::Union{Symbol,Real}=:individual,
                          dateRange=nothing, kw...)
    haskey(SPARKLINE_TYPES, type) || throw(XLSXError("Invalid sparkline `type` `:$type`. Must be one of $(join((":$k" for k in keys(SPARKLINE_TYPES)), ", "))."))
    haskey(_SPARKLINE_EMPTY, emptyCells) || throw(XLSXError("Invalid `emptyCells` `:$emptyCells`. Must be one of $(join((":$k" for k in keys(_SPARKLINE_EMPTY)), ", "))."))
    (isfinite(lineWeight) && 0 < lineWeight <= 1584) || throw(XLSXError("Invalid `lineWeight` $lineWeight. Must be a width in points, more than 0."))
    atts = Dict{String,String}()
    isnothing(SPARKLINE_TYPES[type]) || (atts["type"] = SPARKLINE_TYPES[type])
    isnothing(_SPARKLINE_EMPTY[emptyCells]) || (atts["displayEmptyCellsAs"] = _SPARKLINE_EMPTY[emptyCells])
    lineWeight == 0.75 || (atts["lineWeight"] = string(Float64(lineWeight)))
    for (k, v, manual, axistype) in ((:min, min, "manualMin", "minAxisType"), (:max, max, "manualMax", "maxAxisType"))
        if v isa Real
            isfinite(v) || throw(XLSXError("`$k` must be a finite number, `:individual` or `:group`."))
            atts[axistype] = "custom"
            atts[manual] = string(v)
        elseif v === :group
            atts[axistype] = "group"
        elseif v !== :individual
            throw(XLSXError("Invalid `$k` `:$v`. Must be `:individual`, `:group` or a number."))
        end
    end
    for (k, att) in pairs(_SPARKLINE_FLAGS)
        v = get(kw, k, k === :negative && type === :winloss)
        v isa Bool || throw(XLSXError("`$k` must be `true` or `false`. Found `$(repr(v))`."))
        v && (atts[att] = "1")
    end
    datef = nothing
    if !isnothing(dateRange)
        sheet, r = _sparkline_data(ws, dateRange, "dateRange")
        atts["dateAxis"] = "1"
        datef = XML.Element("xm:f", XML.Text(_sparkline_ref(sheet, r)))
    end
    kids = XML.Node[]
    for (k, (tag, default)) in pairs(_SPARKLINE_COLORS)
        push!(kids, _xml_element("x14:" * tag, _color_atts(get(kw, k, default))))
    end
    isnothing(datef) || push!(kids, datef)
    push!(kids, XML.Element("x14:sparklines",
        (XML.Element("x14:sparkline", XML.Element("xm:f", XML.Text(f)), XML.Element("xm:sqref", XML.Text(string(c))))
         for (c, f) in sparks)...))
    ordered = sort!(collect(atts); by = p -> findfirst(==(first(p)), _SPARKLINE_GROUP_ATT_ORDER))
    return _xml_element("x14:sparklineGroup", ordered, kids)
end

# Remove the sparklines in `cells`, or with `whole_groups` the groups holding any of them,
# and whatever is left empty. Returns whether any was removed.
function _remove_sparklines!(root::XML.Node, cells::Set{CellRef}, whole_groups::Bool)::Bool
    ext = _worksheet_ext(root, SPARKLINE_EXT_URI)
    isnothing(ext) && return false
    removed = false
    for block in xml_elements(ext)
        localname(block) == "sparklineGroups" || continue
        _filter_elements!(block) do g
            localname(g) == "sparklineGroup" || return true
            sl = findfirst(c -> localname(c) == "sparklines", xml_elements(g))
            isnothing(sl) && return false
            sparks = xml_elements(g)[sl]
            hit(s) = (q = _spark_child(s, "sqref"); !isnothing(q) && _spark_cells(q) ⊆ cells)
            if whole_groups
                any(hit, xml_elements(sparks)) && (removed = true; return false)
                return true
            end
            _filter_elements!(s -> localname(s) != "sparkline" || !hit(s) || (removed = true; false), sparks)
            return !isempty(xml_elements(sparks))
        end
    end
    _filter_elements!(b -> localname(b) != "sparklineGroups" || !isempty(xml_elements(b)), ext)
    if isempty(xml_elements(ext))
        extlst = xml_elements(root)[findfirst(c -> localname(c) == "extLst", xml_elements(root))]
        _filter_elements!(e -> e !== ext, extlst)
        isempty(xml_elements(extlst)) && _filter_elements!(c -> c !== extlst, root)
    end
    return removed
end

_spark_child(s::XML.Node, name::String) = (i = findfirst(c -> localname(c) == name, xml_elements(s)); isnothing(i) ? nothing : something(XML.is_simple_value(xml_elements(s)[i]), ""))
_spark_cells(sqref::AbstractString) = Set(c for t in split(sqref) for c in (is_valid_cellname(t) ? (CellRef(t),) : is_valid_cellrange(t) ? CellRange(t) : ()))

"""
    getSparklines(ws::Worksheet) -> Vector{NamedTuple}

A worksheet's sparkline groups, each a `NamedTuple` with fields:

- `sparklines`: each sparkline's cell and data, as `"F2" => "Sheet1!A2:E2"`. A sparkline
  whose data was deleted, with its rows, columns or sheet, has `""`, as Excel keeps it;
- `type`, `markers`, `high`, `low`, `first`, `last`, `negative`, `axis`, `emptyCells`,
  `showHidden`, `lineWeight`, `min`, `max`, `rightToLeft` and `dateRange` (`nothing` if
  none), as the keywords of [`addSparklines`](@ref);
- `colors`: a `NamedTuple` of the colours, `seriesColor` and so on, each an `"FFRRGGBB"`
  hex string or `nothing`.

Works on files opened for reading only.

# Examples
```julia
julia> g = only(XLSX.getSparklines(sheet));

julia> g.type, g.sparklines
(:line, ["F1" => "Sheet1!A1:E1"])
```
"""
function getSparklines(ws::Worksheet)
    wb = get_workbook(ws)
    out = []
    for extlst in _sheet_top_elements(get_xlsxfile(ws), _sheet_path(ws), ("extLst",)), ext in xml_elements(extlst)
        (localname(ext) == "ext" && uppercase(get(ext, "uri", "")) == uppercase(SPARKLINE_EXT_URI)) || continue
        for block in xml_elements(ext), g in xml_elements(block)
            localname(g) == "sparklineGroup" && push!(out, _read_sparkline_group(wb, g))
        end
    end
    return [x for x in out]
end

function _read_sparkline_group(wb::Workbook, g::XML.Node)
    atts = Dict(_atts_vector(g))
    kids = xml_elements(g)
    type = something(findfirst(==(get(atts, "type", nothing)), SPARKLINE_TYPES), :line)
    empty = something(findfirst(==(get(atts, "displayEmptyCellsAs", nothing)), _SPARKLINE_EMPTY), :zero)
    function limit(manual, axistype)
        t = get(atts, axistype, "individual")
        t == "custom" && return something(tryparse(Float64, get(atts, manual, "")), :individual)
        return t == "group" ? :group : :individual
    end
    colors = map(_SPARKLINE_COLORS) do (tag, _)
        i = findfirst(c -> localname(c) == tag, kids)
        isnothing(i) ? nothing : resolveColor(wb, Dict(_atts_vector(kids[i])))
    end
    di = findfirst(c -> localname(c) == "f", kids)
    sl = findfirst(c -> localname(c) == "sparklines", kids)
    sparks = isnothing(sl) ? Pair{String,String}[] :
        [something(_spark_child(s, "sqref"), "") => something(_spark_child(s, "f"), "")
         for s in xml_elements(kids[sl]) if localname(s) == "sparkline"]
    flags = map(att -> _xml_bool(atts, att, false), _SPARKLINE_FLAGS)
    return (sparklines = sparks, type = type, flags..., emptyCells = empty,
            lineWeight = something(tryparse(Float64, get(atts, "lineWeight", "")), 0.75),
            min = limit("manualMin", "minAxisType"), max = limit("manualMax", "maxAxisType"),
            dateRange = isnothing(di) || !_xml_bool(atts, "dateAxis", false) ? nothing : something(XML.is_simple_value(kids[di]), nothing),
            colors = colors)
end

"""
    deleteSparklines(ws::Worksheet, ref; group=false) -> Bool

Delete the sparklines in the cells of `ref` (a cell, a range, several separated by
commas, whole columns or rows), as Excel's *Clear Selected Sparklines* does; with
`group=true`, delete every group with a sparkline there, as *Clear Selected Sparkline
Groups* does. Returns `false` if there were none.

The file must be open for writing (`mode="rw"`).
"""
function deleteSparklines(ws::Worksheet, ref::Union{AbstractString,CellRef,CellRange}; group::Bool=false)::Bool
    _require_writable(ws, "delete sparklines")
    _require_worksheet(ws, "delete sparklines")
    region = _cell_region(ws, ref isa CellRef ? CellRange(ref, ref) : ref)
    cells = Set{CellRef}()
    for g in getSparklines(ws), (c, _) in g.sparklines
        is_valid_cellname(c) && any(r -> CellRef(c) ∈ r, region) && push!(cells, CellRef(c))
    end
    isempty(cells) && return false
    _remove_sparklines!(_sheet_root(ws), cells, group)
    update_worksheets_xml!(get_xlsxfile(ws))
    return true
end

# A sparkline whose data has been deleted, with its rows, columns or sheet, keeps its cell
# but loses its data, as Excel writes it, rather than charting `#REF!`.
function _drop_dead_sparkline_data!(root::XML.Node)
    ext = _worksheet_ext(root, SPARKLINE_EXT_URI)
    isnothing(ext) && return nothing
    for block in xml_elements(ext), g in xml_elements(block), sl in xml_elements(g)
        localname(sl) == "sparklines" || continue
        for s in xml_elements(sl)
            f = _spark_child(s, "f")
            isnothing(f) || !occursin("#REF!", f) || _filter_elements!(c -> localname(c) != "f", s)
        end
    end
    return nothing
end

# "Format Same As Above": rows (or columns) inserted below a sparkline get one each, in
# its group, charting the cells as far below its data, as Excel gives them. Called once
# the edit's formulas are moved.
function _inherit_sparklines!(ws::Worksheet, e::RowColEdit)
    (e.count > 0 && e.at > 1) || return nothing
    ext = _worksheet_ext(_sheet_root(ws), SPARKLINE_EXT_URI)
    isnothing(ext) && return nothing
    news = e.at:min(e.at + e.count - 1, _edit_max(e))
    for block in xml_elements(ext), g in xml_elements(block)
        sl = findfirst(c -> localname(c) == "sparklines", xml_elements(g))
        isnothing(sl) && continue
        sparks = xml_elements(g)[sl]
        for s in collect(xml_elements(sparks))
            q, f = _spark_child(s, "sqref"), _spark_child(s, "f")
            (isnothing(q) || isnothing(f) || !is_valid_cellname(q)) && continue
            c = CellRef(q)
            (e.axis === :row ? c.row_number : c.column_number) == e.at - 1 || continue
            i = findfirst(x -> x === s, XML.children(sparks))
            tag(name) = XML.tag(xml_elements(s)[findfirst(x -> localname(x) == name, xml_elements(s))])
            for (k, n) in enumerate(news)
                offset = e.axis === :row ? (k, 0) : (0, k)
                cell = e.axis === :row ? CellRef(n, c.column_number) : CellRef(c.row_number, n)
                el = XML.Element(XML.tag(s), XML.Element(tag("f"), XML.Text(shift_excel_references(f, offset))),
                                 XML.Element(tag("sqref"), XML.Text(string(cell))))
                insert!(XML.children(sparks), i + k, el)
            end
        end
    end
    return nothing
end
