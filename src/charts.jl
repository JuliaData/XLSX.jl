#
# charts.jl
#
# Read chart metadata and the cached data Excel stores inside chart parts
# (JuliaData/XLSX.jl#263).
#
# Excel writes a snapshot of every series' source data into the chart part
# itself (`c:numCache` / `c:strCache` / `c:multiLvlStrCache`). That cache is what
# makes a chart render when its source is unavailable - a deleted sheet, or an
# external workbook that isn't to hand - and it is what this file exposes. It is
# never re-read from the worksheet, so it reflects the values as of the last time
# Excel saved the file.
#

const REL_CHART =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart"

const CT_CHART = "application/vnd.openxmlformats-officedocument.drawingml.chart+xml"

const CT_CHARTEX = "application/vnd.ms-office.chartex+xml"


# The <c:plotArea> children that group series. Series live one level below these.
const CHART_GROUP_TAGS = Set([
    "areaChart", "area3DChart", "lineChart", "line3DChart", "stockChart",
    "radarChart", "scatterChart", "pieChart", "pie3DChart", "doughnutChart",
    "barChart", "bar3DChart", "ofPieChart", "surfaceChart", "surface3DChart",
    "bubbleChart",
])

# chart part path => (sheet name, from, to, rId)
const ChartAnchor = NamedTuple{
    (:sheet, :from, :to, :rId),
    Tuple{String,Union{Nothing,String},Union{Nothing,String},String},
}

const ChartRange = Union{Nothing,SheetCellRef,SheetCellRange,SheetRowRange,SheetColumnRange,NonContiguousRange}

const ChartRanges = @NamedTuple{
    idx::Int,
    name::Union{Nothing,String},
    categories::ChartRange,
    values::ChartRange,
    bubble_sizes::ChartRange,
}

#=
# ===========================================================================
# Traversal helpers
# ===========================================================================

# First-match, `nothing`-tolerant counterpart of `elements_with_tag`, which
# collects every match. Iterates `XML.eachelement` (a lazy filter) rather than
# `element_children`, so nothing is allocated to find one child. The manual loop
# in `_parse_cell_marker` (images.jl) could be replaced by this.
function first_element_with_tag(node::Union{Nothing,XML.Node}, tag::String)::Union{Nothing,XML.Node}
    isnothing(node) && return nothing
    for n in XML.eachelement(node)
        localname(n) == tag && return n
    end
    return nothing
end

# Text of a simple child element, e.g. <c:f>Sheet1!$A$1</c:f>. `XML.is_simple_value`
# returns the text of an element with no attributes and a single Text/CData child,
# or `nothing`, which is precisely the shape of every element read here.
function child_text(node::Union{Nothing,XML.Node}, tag::String)::Union{Nothing,String}
    el = first_element_with_tag(node, tag)
    isnothing(el) && return nothing
    v = XML.is_simple_value(el)
    return isnothing(v) ? nothing : String(v)
end

# `val` attribute of an optional child element, e.g. <c:order val="1"/>.
function child_val(node::Union{Nothing,XML.Node}, tag::String, default::Int)::Int
    el = first_element_with_tag(node, tag)
    isnothing(el) && return default
    return something(tryparse(Int, get_attr(el, "val")), default)
end

# Attribute lookup ignoring any namespace prefix (`r:id`, `id`, ...).
function get_prefixed_attr(node::XML.Node, key::AbstractString)::Union{Nothing,String}
    atts = XML.attributes(node)
    isnothing(atts) && return nothing
    for (k, v) in atts
        occursin(':', k) || continue     # unprefixed => no namespace, never a match
        localname(k) == key && return v
    end
    return nothing
end
=#

"""
Id => resolved target path for the relationships of `part_path`, filtered by
relationship type. Targets are resolved against the part's own directory, so
`../charts/chart1.xml` from `xl/drawings/drawing1.xml` gives `xl/charts/chart1.xml`.
"""
function rid_to_target(xf::XLSXFile, part_path::String, reltype::String)::Dict{String,String}
    dir, fname = _split_zip_path(part_path)
    rels_path = isempty(dir) ? "_rels/$fname.rels" : "$dir/_rels/$fname.rels"
    targets = Dict{String,String}()
    haskey(xf.data, rels_path) || return targets
    for n in elements_with_tag(xml_root_element(xf.data[rels_path]), "Relationship")
        get_attr(n, "Type") == reltype && get_attr(n, "TargetMode") != "External" || continue
        id = get_attr(n, "Id")
        isempty(id) && continue
        targets[id] = resolve_relative_target(dir, get_attr(n, "Target"))
    end
    return targets
end

# ===========================================================================
# Cache parsing
# ===========================================================================

# A cached point is one of: a number, a known Excel error, or (rarely) text that
# is neither. Errors are recorded separately and stored as `missing`.
function parse_cached_point!(errors::Dict{Int,UInt64}, i::Int, s::AbstractString, numeric::Bool)
    if haskey(ERROR_STRING_TO_CODE, s)
        errors[i] = ERROR_STRING_TO_CODE[s]
        return missing
    end
    numeric || return String(s)
    v = tryparse(Float64, s)
    return isnothing(v) ? String(s) : v
end

"Dense vector of the `c:pt` children of a cache, literal or level node."
function parse_cached_points(node::XML.Node, numeric::Bool)
    n = child_val(node, "ptCount", 0)
    values = Vector{Any}(missing, n)
    errors = Dict{Int,UInt64}()
    for pt in XML.eachelement(node)
        localname(pt) == "pt" || continue
        i = tryparse(Int, get_attr(pt, "idx"))
        isnothing(i) && continue
        s = child_text(pt, "v")
        isnothing(s) && continue
        i += 1
        i < 1 && continue
        i > length(values) && append!(values, fill(missing, i - length(values)))  # tolerate a bad ptCount
        values[i] = parse_cached_point!(errors, i, s, numeric)
    end
    return (isempty(values) ? values : identity.(values)), errors
end

cache_ptcount(cache::Union{Nothing,XML.Node})::Int = child_val(cache, "ptCount", 0)

"""
Parse a series-role container (`c:tx`, `c:cat`, `c:val`, `c:xVal`, `c:yVal`,
`c:bubbleSize`) into a `ChartRef`, or `nothing` if it holds no cache.
"""
function parse_chart_ref(container::Union{Nothing,XML.Node}; cache::Bool=true)::Union{Nothing,ChartRef}
    isnothing(container) && return nothing
    for node in XML.eachelement(container)
        tag = localname(node)
        if tag == "numRef" || tag == "strRef"
            numeric = tag == "numRef"
            cnode = first_element_with_tag(node, numeric ? "numCache" : "strCache")
            pts, errs = (cache && !isnothing(cnode)) ? parse_cached_points(cnode, numeric) : (Any[], Dict{Int,UInt64}())
            return ChartRef(numeric ? :num : :str, child_text(node, "f"), child_text(cnode, "formatCode"),
                            cache_ptcount(cnode), pts, errs)
        elseif tag == "multiLvlStrRef"
            cnode = first_element_with_tag(node, "multiLvlStrCache")
            levels = Any[]
            if cache && !isnothing(cnode)
                for lvl in XML.eachelement(cnode)
                    localname(lvl) == "lvl" || continue
                    pts, _ = parse_cached_points(lvl, false)
                    push!(levels, pts)
                end
            end
            return ChartRef(:multiLvlStr, child_text(node, "f"), nothing, cache_ptcount(cnode), levels, Dict{Int,UInt64}())
        elseif tag == "numLit" || tag == "strLit"
            numeric = tag == "numLit"
            pts, errs = cache ? parse_cached_points(node, numeric) : (Any[], Dict{Int,UInt64}())
            return ChartRef(numeric ? :numLit : :strLit, nothing,
                            numeric ? child_text(node, "formatCode") : nothing,
                            cache_ptcount(node), pts, errs)
        elseif tag == "v"
            # <c:tx><c:v>Literal name</c:v></c:tx>
            return ChartRef(:strLit, nothing, nothing, 1,
                            Any[something(XML.is_simple_value(node), "")], Dict{Int,UInt64}())
        end
    end
    return nothing
end

function first_cached_string(r::Union{Nothing,ChartRef})::Union{Nothing,String}
    (isnothing(r) || isempty(r.data)) && return nothing
    v = first(r.data)
    (ismissing(v) || v isa AbstractVector) && return nothing
    return string(v)
end

# ===========================================================================
# External references
# ===========================================================================

# Replace the `[n]` index of an external workbook reference with the workbook
# path recorded in xl/externalLinks, as `getFormula(...; get_external_refs=true)`
# does for formulas.
function materialise_external_ref(xf::XLSXFile, ref::Union{Nothing,String})::Union{Nothing,String}
    isnothing(ref) && return nothing
    occursin('[', ref) || return ref
    out = ref
    for e in get_ext_refs(ref)
        path = try
            get_external_workbook_path(xf, e.index)
        catch err
            err isa XLSXError || rethrow()
            continue    # leave `[n]` unresolved rather than lose the chart
        end
        out = replace(out, "[" * string(e.index) * "]" => "[" * path * "]")
    end
    return out
end

function materialise(xf::XLSXFile, r::Union{Nothing,ChartRef})::Union{Nothing,ChartRef}
    isnothing(r) && return nothing
    isnothing(r.ref) && return r
    return ChartRef(r.kind, materialise_external_ref(xf, r.ref), r.format_code, r.ptCount, r.data, r.errors)
end

# ===========================================================================
# Chart part parsing
# ===========================================================================

chart_name(path::AbstractString) = first(splitext(last(_split_zip_path(String(path)))))

# Part names declared in [Content_Types].xml with the given content type.
function parts_with_content_type(xf::XLSXFile, ctype::String)::Vector{String}
    paths = String[]
    haskey(xf.data, "[Content_Types].xml") || return paths
    for n in elements_with_tag(xml_root_element(xf.data["[Content_Types].xml"]), "Override")
        get_attr(n, "ContentType") == ctype || continue
        push!(paths, String(lstrip(get_attr(n, "PartName"), '/')))
    end
    return sort!(paths)
end

chart_parts(xf::XLSXFile) = filter(p -> haskey(xf.data, p), parts_with_content_type(xf, CT_CHART))
chartex_parts(xf::XLSXFile) = parts_with_content_type(xf, CT_CHARTEX)

function chartex_note(xf::XLSXFile)::String
    n = length(chartex_parts(xf))
    n == 0 && return ""
    return " The file also contains $n chart$(n == 1 ? "" : "s") using the `chartEx` schema " *
           "(waterfall, funnel, treemap, sunburst, histogram, box & whisker), which XLSX.jl cannot read."
end

function parse_chart_series(ser::XML.Node, charttype::Symbol; cache::Bool=true)::ChartSeries
    # The series name always needs its cache, even under `cache=false`: it is metadata.
    name_ref = parse_chart_ref(first_element_with_tag(ser, "tx"); cache=true)

    categories = parse_chart_ref(first_element_with_tag(ser, "cat"); cache=cache)
    isnothing(categories) && (categories = parse_chart_ref(first_element_with_tag(ser, "xVal"); cache=cache))

    values = parse_chart_ref(first_element_with_tag(ser, "val"); cache=cache)
    isnothing(values) && (values = parse_chart_ref(first_element_with_tag(ser, "yVal"); cache=cache))

    return ChartSeries(
        child_val(ser, "idx", -1),
        child_val(ser, "order", -1),
        charttype,
        first_cached_string(name_ref),
        name_ref,
        categories,
        values,
        parse_chart_ref(first_element_with_tag(ser, "bubbleSize"); cache=cache),
    )
end

function parse_chart_title(chartnode::Union{Nothing,XML.Node})::Union{Nothing,String}
    tx = first_element_with_tag(first_element_with_tag(chartnode, "title"), "tx")
    isnothing(tx) && return nothing

    rich = first_element_with_tag(tx, "rich")
    if !isnothing(rich)
        buf = IOBuffer()
        for p in XML.eachelement(rich)
            localname(p) == "p" || continue
            for run in XML.eachelement(p)
                localname(run) in ("r", "fld") || continue
                print(buf, something(child_text(run, "t"), ""))
            end
        end
        title = String(take!(buf))
        return isempty(title) ? nothing : title
    end

    return first_cached_string(parse_chart_ref(tx; cache=true))
end

function parse_chart_part(
    xf::XLSXFile,
    path::String;
    cache::Bool=true,
    get_external_refs::Bool=false,
    rId::Union{Nothing,String}=nothing,
    sheet::Union{Nothing,String}=nothing,
    from::Union{Nothing,String}=nothing,
    to::Union{Nothing,String}=nothing,
)::Chart

    haskey(xf.data, path) || throw(XLSXError("Chart part `$path` not found in the package."))

    chartspace = xml_root_element(xf.data[path])
    localname(chartspace) != "chartSpace" &&
        throw(XLSXError("Malformed chart part $path. Root node name should be `chartSpace`. Found $(localname(chartspace))."))

    chartnode = first_element_with_tag(chartspace, "chart")
    plotarea = first_element_with_tag(chartnode, "plotArea")

    charttypes = Symbol[]
    series = ChartSeries[]
    for group in (isnothing(plotarea) ? () : XML.eachelement(plotarea))
        tag = localname(group)
        tag in CHART_GROUP_TAGS || continue
        charttype = Symbol(tag)
        push!(charttypes, charttype)
        for ser in elements_with_tag(group, "ser")
            s = parse_chart_series(ser, charttype; cache=cache)
            if get_external_refs
                s = ChartSeries(s.idx, s.order, s.charttype, s.name,
                                materialise(xf, s.name_ref),
                                materialise(xf, s.categories),
                                materialise(xf, s.values),
                                materialise(xf, s.bubble_sizes))
            end
            push!(series, s)
        end
    end

    _, fname = _split_zip_path(path)
    return Chart(path, first(splitext(fname)), rId, sheet, from, to,
                 parse_chart_title(chartnode), charttypes, series)
end

# ===========================================================================
# Discovery: sheet -> drawing -> chart
# ===========================================================================

# Chart parts in document order, each with its anchor when a drawing references it.
# Sheet-anchored charts come first (sheet order, then anchor order); parts the
# package declares but no drawing references follow.
function chart_positions(xf::XLSXFile)::Vector{Pair{String,Union{Nothing,ChartAnchor}}}
    out = Pair{String,Union{Nothing,ChartAnchor}}[]
    seen = Set{String}()
    for (path, a) in chart_anchors(xf)
        (path in seen || !haskey(xf.data, path)) && continue
        push!(seen, path)
        push!(out, path => a)
    end
    for path in chart_parts(xf)
        path in seen && continue
        push!(seen, path)
        push!(out, path => nothing)
    end
    return out
end

function chart_positions(ws::Worksheet)::Vector{Pair{String,Union{Nothing,ChartAnchor}}}
    xf = get_xlsxfile(ws)
    anchors = Pair{String,ChartAnchor}[]
    sheet_path = get_relationship_target_by_id("xl", get_workbook(ws), ws.relationship_id)
    charts_for_sheet!(anchors, xf, sheet_path, ws.name)
    out = Pair{String,Union{Nothing,ChartAnchor}}[]
    seen = Set{String}()
    for (path, a) in anchors
        path in seen && continue
        push!(seen, path)
        push!(out, path => a)
    end
    return out
end

parse_chart_at(xf::XLSXFile, path::String, a::Nothing; kw...) = parse_chart_part(xf, path; kw...)
parse_chart_at(xf::XLSXFile, path::String, a::ChartAnchor; kw...) = parse_chart_part(xf, path; kw..., a...)

# The chart reference sits at a fixed depth in a drawing anchor:
#   <xdr:*Anchor><xdr:graphicFrame><a:graphic><a:graphicData><c:chart r:id="..."/>
# Walking it explicitly is both cheaper and more precise than a recursive search
# for a "chart" element, which could match elsewhere in the anchor.
function anchor_chart_element(anchor::XML.Node)::Union{Nothing,XML.Node}
    frame = first_element_with_tag(anchor, "graphicFrame")
    graphic = first_element_with_tag(frame, "graphic")
    graphicdata = first_element_with_tag(graphic, "graphicData")
    return first_element_with_tag(graphicdata, "chart")
end

function charts_for_sheet!(anchors::Vector{Pair{String,ChartAnchor}}, xf::XLSXFile, sheet_path::String, sheet_name::String)
    drawing_path = _drawing_path_for_sheet(xf, sheet_path)
    isnothing(drawing_path) && return anchors
    haskey(xf.data, drawing_path) || return anchors

    rid_to_chart = rid_to_target(xf, drawing_path, REL_CHART)
    isempty(rid_to_chart) && return anchors

    for anchor in XML.eachelement(xml_root_element(xf.data[drawing_path]))
        endswith(localname(anchor), "Anchor") || continue
        chart_el = anchor_chart_element(anchor)
        isnothing(chart_el) && continue
        rId = get_prefixed_attr(chart_el, "id")
        isnothing(rId) && continue
        chart_path = get(rid_to_chart, rId, nothing)
        isnothing(chart_path) && continue
        push!(anchors, chart_path => (
            sheet=sheet_name,
            from=_parse_cell_marker(anchor, "from"; is_to=false),
            to=_parse_cell_marker(anchor, "to"; is_to=true),
            rId=rId,
        ))
    end

    return anchors
end

function chart_anchors(xf::XLSXFile)::Vector{Pair{String,ChartAnchor}}
    wb = get_workbook(xf)
    anchors = Pair{String,ChartAnchor}[]
    for sheet in wb.sheets
        sheet_path = get_relationship_target_by_id("xl", wb, sheet.relationship_id)
        charts_for_sheet!(anchors, xf, sheet_path, sheet.name)
    end
    return anchors
end

# ===========================================================================
# Public API
# ===========================================================================

"""
    getCharts(xf::XLSXFile; cache=true, get_external_refs=false) -> Vector{Chart}
    getCharts(ws::Worksheet; cache=true, get_external_refs=false) -> Vector{Chart}

Return every chart in the file, or every chart anchored to `ws`, together with
the data Excel cached inside each chart part.

Pass `cache=false` to read metadata only - title, chart types, series names,
source formulas, format codes and point counts - and skip the cached values,
which is the expensive part for a large chart.

A chart may reference an external workbook, in which case its source formula
takes the form `[1]Sheet1!\$A\$1:\$A\$10`, where `[1]` indexes the workbook's
external references. Use `get_external_refs=true` to substitute the workbook
path, as [`XLSX.getFormula`](@ref) does.

# Examples
```julia
julia> f = XLSX.readxlsx("sales.xlsx");

julia> c = XLSX.getCharts(f["Summary"])[1];

julia> c.title
"Revenue by region"

julia> c.series[1].values.ref
"Summary!\$B\$2:\$B\$5"

julia> XLSX.getChartData(c)
```

!!! note
    The values returned are Excel's cache, written when the file was last saved
    by Excel. They may be stale relative to the source, and a file written by a
    tool that does not populate the cache will return empty series.

!!! note
    Charts using the newer `chartEx` schema (waterfall, funnel, treemap,
    sunburst, histogram, box & whisker) are stored under a different namespace
    and are not read.

See also [`XLSX.getChart`](@ref), [`XLSX.getChartData`](@ref).
"""
function getCharts(x::Union{Worksheet,XLSXFile}; cache::Bool=true, get_external_refs::Bool=false)::Vector{Chart}
    xf = get_xlsxfile(x)
    charts = [parse_chart_at(xf, path, a; cache=cache, get_external_refs=get_external_refs)
              for (path, a) in chart_positions(x)]
    if isempty(charts) && !isempty(chartex_parts(xf))
        @warn "No readable charts found." * chartex_note(xf) maxlog=1
    end
    return charts
end

"""
    getChart(ws::Worksheet, name; cache=true, get_external_refs=false) -> Chart
    getChart(xf::XLSXFile, name; cache=true, get_external_refs=false) -> Chart

Return a single chart. `name` may be the part name (`"chart1"` or
`"chart1.xml"`), the full package path, or the chart's relationship id within its
drawing part (`"rId1"`).

See also [`XLSX.getCharts`](@ref).
"""
function getChart(x::Union{Worksheet,XLSXFile}, name::AbstractString;
                  cache::Bool=true, get_external_refs::Bool=false)::Chart
    xf = get_xlsxfile(x)
    positions = chart_positions(x)
    stem = chart_name(name)
    for (path, a) in positions
        path == name || chart_name(path) == stem ||
            (!isnothing(a) && a.rId == name) || continue
        return parse_chart_at(xf, path, a; cache=cache, get_external_refs=get_external_refs)
    end
    throw(XLSXError("No chart matching `$name`. Found: $(join(chart_name.(first.(positions)), ", "))." * chartex_note(get_xlsxfile(x))))
end


# ===========================================================================
# Errors in cached values
# ===========================================================================

"""
    iserror(r::ChartRef) -> Vector{Bool}
    iserror(r::ChartRef, i::Integer) -> Bool

Report which cached chart values are Excel error values. When Excel writes source data 
to a chart cache, all error values are written as simple zeros except for `#N/A`. 
Therefore, when  operating on a chart cache, only `#N/A` values will return `true`. 
All other error values will return `false`, and are indistinguisable from genuine 
zero values.

See also [`XLSX.geterror`](@ref).
"""
iserror(r::ChartRef)::Vector{Bool} = Bool[haskey(r.errors, i) for i in 1:length(r.data)]
iserror(r::ChartRef, i::Integer)::Bool = haskey(r.errors, Int(i))

"""
    geterror(r::ChartRef) -> Vector{String}
    geterror(r::ChartRef, i::Integer) -> String

Resolve cached chart `#N/A`error values to their Excel strings (`"#N/A"`). All other 
error values are written by Excel as simple zeros in the chart cache, and return "".

See also [`XLSX.iserror`](@ref).
"""
geterror(r::ChartRef)::Vector{String} = String[geterror(r, i) for i in 1:length(r.data)]
geterror(r::ChartRef, i::Integer)::String =
    haskey(r.errors, Int(i)) ? get_error_string(r.errors[Int(i)]) : ""

# ===========================================================================
# Cached data as a table
# ===========================================================================

function unique_label!(labels::Vector{Symbol}, name::AbstractString)::Symbol
    base = Symbol(isempty(name) ? "column" : name)
    label = base
    n = 1
    while label in labels
        n += 1
        label = Symbol(base, "_", n)
    end
    push!(labels, label)
    return label
end

pad_to(v::AbstractVector, n::Int) =
    length(v) >= n ? collect(v) : vcat(collect(v), fill(missing, n - length(v)))

# A ref read with `cache=false` keeps its declared ptCount but no values. An
# absent cache gives ptCount == 0; a cache of blanks gives a full-length vector
# of `missing`. So this combination is unambiguous.
no_cached_values(r::ChartRef)::Bool = r.ptCount > 0 && isempty(r.data)

"""
    getChartData(c::Chart) -> DataTable
    getChartData(ws::Worksheet, name) -> DataTable
    getChartData(xf::XLSXFile, name) -> DataTable

Return the cached data of a chart as a `DataTable`, ready for `DataFrame(...)` or
any other Tables.jl sink.

Categories become the leading column(s). Where every series shares one category
reference a single `categories` column is emitted; otherwise each series
contributes its own `<series>_x` column, which is the usual layout for scatter
and bubble charts. Multi-level categories give one column per level, and bubble
charts add a `<series>_size` column. Series of unequal length are padded with
`missing`.

Series with no name in the file - Excel shows these as "Series1", "Series2" in
the legend - are labelled by position. No category column is produced when the
chart has no `c:cat` at all: Excel is plotting against an implicit index in that
case, and nothing is cached for it.

# Examples
```julia
julia> using DataFrames

julia> DataFrame(XLSX.getChartData(f["Summary"], "chart1"))
4×3 DataFrame
 Row │ categories  2024      2025
     │ String      Float64   Float64
```

See also [`XLSX.getCharts`](@ref), [`XLSX.gettable`](@ref).
"""
function getChartData(c::Chart)::DataTable
    isempty(c.series) && return DataTable(Any[], Symbol[])

    for s in c.series, r in (s.categories, s.values, s.bubble_sizes)
        isnothing(r) && continue
        no_cached_values(r) && throw(XLSXError(
            "Chart `$(c.name)` was read with `cache=false`, so its cached values are not available. Read it again with `cache=true`."))
    end
    catrefs = [s.categories for s in c.series]
    shared = !isnothing(first(catrefs)) && !isnothing(first(catrefs).ref) &&
             all(r -> !isnothing(r) && r.ref == first(catrefs).ref, catrefs)

    n = 0
    for s in c.series, r in (s.categories, s.values)
        isnothing(r) && continue
        n = max(n, r.ptCount, length(r.data))
    end

    columns = Any[]
    labels = Symbol[]

    function push_categories!(r::ChartRef, prefix::AbstractString)
        if r.kind == :multiLvlStr
            for (i, level) in enumerate(r.data)
                unique_label!(labels, "$(prefix)_$(i)")
                push!(columns, pad_to(level, n))
            end
        else
            unique_label!(labels, prefix)
            push!(columns, pad_to(r.data, n))
        end
    end

    shared && push_categories!(first(catrefs), "categories")

    for (i, s) in enumerate(c.series)
        name = something(s.name, "Series$(i)")
        shared || isnothing(s.categories) || push_categories!(s.categories, "$(name)_x")
        unique_label!(labels, name)
        push!(columns, isnothing(s.values) ? Vector{Any}(missing, n) : pad_to(s.values.data, n))
        if !isnothing(s.bubble_sizes)
            unique_label!(labels, "$(name)_size")
            push!(columns, pad_to(s.bubble_sizes.data, n))
        end
    end

    return DataTable(columns, labels)
end

getChartData(x::Union{Worksheet,XLSXFile}, name::AbstractString; kw...)::DataTable =
    getChartData(getChart(x, name; kw...))

"""
    chart_range(r) -> Union{Nothing,SheetCellRef,SheetCellRange,SheetRowRange,SheetColumnRange,NonContiguousRange}

The source range of a `ChartRef`, or `nothing` when it has no addressable one:
literal series, external-workbook references, and defined names.
"""
function chart_range(r::Union{Nothing,ChartRef})
    (isnothing(r) || isnothing(r.ref)) && return nothing
    s = strip(r.ref)
    occursin('[', s) && return nothing                  # external workbook
    if startswith(s, '(') && endswith(s, ')')           # multi-area
        s = s[nextind(s, firstindex(s)):prevind(s, lastindex(s))]
    end
    occursin(',', s) && return NonContiguousRange(String(s))
    (is_valid_fixed_sheet_cellrange(s) || is_valid_sheet_cellrange(s)) && return SheetCellRange(s)
    (is_valid_fixed_sheet_cellname(s)  || is_valid_sheet_cellname(s))  && return SheetCellRef(s)
    (is_valid_fixed_sheet_column_range(s) || is_valid_sheet_column_range(s)) && return SheetColumnRange(s)
    (is_valid_fixed_sheet_row_range(s)    || is_valid_sheet_row_range(s))    && return SheetRowRange(s)
    return nothing                                      # defined name, or unrecognised
end

"""
    getChartRanges(c::Chart) -> Vector{ChartRanges}
    getChartRanges(ws::Worksheet, name) -> Vector{ChartRanges}
    getChartRanges(xf::XLSXFile, name) -> Vector{ChartRanges}
    getChartRanges(ws::Worksheet) -> Vector{@NamedTuple{chart::String, ranges::Vector{ChartRanges}}}
    getChartRanges(xf::XLSXFile) -> Vector{@NamedTuple{chart::String, ranges::Vector{ChartRanges}}}

The worksheet ranges of the source data a chart plots from.

Given a `Chart`, or a chart `name` in any of the forms [`XLSX.getChart`](@ref)
accepts, return one entry per series in document order, parallel to `c.series`.
Each entry carries the series `idx` and `name` alongside its `categories`,
`values` and `bubble_sizes` ranges.

Given no name, return the ranges of every chart on the worksheet or in the
workbook, each paired with its chart name, following [`XLSX.getCharts`](@ref).

`categories` holds `c:cat` or `c:xVal` and `values` holds `c:val` or `c:yVal`, so
the two mean the same thing whatever the chart type, as in [`XLSX.ChartSeries`](@ref).
`bubble_sizes` is `nothing` for every chart type but bubble.

A range is `nothing` wherever the series has no addressable source: a literal
series (`c:numLit`/`c:strLit`), a reference to an external workbook, a defined
name.


# Examples
```julia
julia> f = XLSX.readxlsx("sales.xlsx");

julia> r = XLSX.getChartRanges(f["Summary"], "chart1");

julia> r[1].name, r[1].values
("2024", Summary!B2:B5)

julia> XLSX.getdata(f, r[1].values)      # read the live source cells, not the cache
4-element Vector{Any}:
 1250.0
 1310.0
 ⋮

julia> [(x.chart, length(x.ranges)) for x in XLSX.getChartRanges(f)]
2-element Vector{Tuple{String, Int64}}:
 ("chart1", 3)
 ("chart2", 1)
```

!!! note
    A range records where the chart says its source data came from, which is not
    necessarily where the values in [`XLSX.getChartData`](@ref) came from: the
    cache is a snapshot from the last save, and the cells may have changed
    since, or the source sheet may have been deleted entirely.

See also [`XLSX.getChart`](@ref), [`XLSX.getCharts`](@ref), [`XLSX.getChartData`](@ref).
"""
getChartRanges(c::Chart)::Vector{ChartRanges} =
    [(idx = s.idx,
      name = s.name,
      categories = chart_range(s.categories),
      values = chart_range(s.values),
      bubble_sizes = chart_range(s.bubble_sizes))
     for s in c.series]

getChartRanges(x::Union{Worksheet,XLSXFile}, name::AbstractString)::Vector{ChartRanges} =
    getChartRanges(getChart(x, name; cache=false))

getChartRanges(x::Union{Worksheet,XLSXFile}) =
    [(chart = c.name, ranges = getChartRanges(c)) for c in getCharts(x; cache=false)]

# ===========================================================================
# Display
# ===========================================================================

function Base.show(io::IO, c::Chart)
    print(io, "XLSX.Chart(\"", c.name, "\"",
          isnothing(c.sheet) ? "" : ", \"" * c.sheet * "\"",
          isnothing(c.from) ? "" : ", " * c.from,
          ", ", join(string.(c.charttypes), "+"),
          ", ", length(c.series), " series)")
end

Base.show(io::IO, s::ChartSeries) =
    print(io, "XLSX.ChartSeries(", something(s.name, "<unnamed>"), ", ", s.charttype, ")")

Base.show(io::IO, r::ChartRef) =
    print(io, "XLSX.ChartRef(", something(r.ref, "<literal>"), ", ", r.ptCount, " pts)")
    
function Base.show(io::IO, ::MIME"text/plain", c::Chart)
    print(io, "XLSX.Chart \"", c.name, "\"")
    isnothing(c.sheet) || print(io, " on sheet \"", c.sheet, "\"")
    isnothing(c.from) || print(io, " at ", c.from, isnothing(c.to) ? "" : ":" * c.to)
    println(io)
    isnothing(c.title) || println(io, "  title: ", repr(c.title))
    println(io, "  type: ", join(string.(c.charttypes), ", "))
    println(io, "  series: ", length(c.series))
    for (i, s) in enumerate(c.series)
        print(io, "    [", s.order, "] ", something(s.name, "Series$(i)"))
        isnothing(s.values) ||
            print(io, " - ", something(s.values.ref, "<literal>"), " (", s.values.ptCount, " pts)")
        println(io)
    end
end

function Base.show(io::IO, ::MIME"text/plain", s::ChartSeries)
    println(io, "XLSX.ChartSeries ", something(s.name, "<unnamed>"), " (", s.charttype, ")")
    for (label, r) in (("categories", s.categories), ("values", s.values), ("sizes", s.bubble_sizes))
        isnothing(r) && continue
        println(io, "  ", label, ": ", something(r.ref, "<literal>"), " - ", r.ptCount, " pts, ", r.kind)
    end
end

function Base.show(io::IO, ::MIME"text/plain", r::ChartRef)
    println(io, "XLSX.ChartRef ", something(r.ref, "<literal>"), " (", r.kind, ", ", r.ptCount, " pts)")
    isnothing(r.format_code) || println(io, "  format: ", r.format_code)
    isempty(r.errors) || println(io, "  errors at: ", join(sort(collect(keys(r.errors))), ", "))
    isempty(r.data) || println(io, "  data: ", r.data)
end

