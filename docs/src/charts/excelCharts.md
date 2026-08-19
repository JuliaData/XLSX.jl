# Charts

(This guide is a Work in Progress)

## What XLSX.jl can do with charts

XLSX.jl can *read* the charts in a workbook: where each chart lives, what type it
is, what its series are called, which cell ranges each series plots from, and the
data Excel cached inside the chart itself.

It cannot create, modify or delete charts. A chart present in a file that is read
and written again is preserved unchanged.

## The chart cache

Every chart part in an `.xlsx` file carries a snapshot of the data it plots,
written by Excel each time the file is saved. This is what makes a chart render
even when its source is unavailable — the source sheet has been deleted, or the
data lives in another workbook that isn't to hand.

That cache is what XLSX.jl reads. It is never recomputed from the worksheet, so
two things follow, and both matter:

- The values reflect the state of the source cells **at the last save by Excel**.
  If the source has changed since, the cache is stale.
- A file written by a tool that doesn't populate the cache will have charts with
  no cached values at all, even though the chart is perfectly valid and will
  display correctly once Excel opens and re-saves it.

Where a chart's series refers to a real range, [`XLSX.getChartRanges`](@ref) gives
you that range so you can read the live cells instead. See
[Where the data came from](#Where-the-data-came-from) below.

## Finding the charts

[`XLSX.getCharts`](@ref) lists every chart anchored to a worksheet, or every chart
in the file:

```julia
julia> using XLSX

julia> f = XLSX.readxlsx("chart_basic.xlsx")
XLSXFile("chart_basic.xlsx") containing 1 Worksheet
            sheetname size          range
-------------------------------------------------
                 Data 5x3           A1:C5

julia> XLSX.getCharts(f["Data"])
1-element Vector{XLSX.Chart}:
 XLSX.Chart("chart1", "Data", G9, barChart, 2 series)

julia> XLSX.getCharts(f)          # the whole workbook
1-element Vector{XLSX.Chart}:
 XLSX.Chart("chart1", "Data", G9, barChart, 2 series)
```

`getCharts(ws)` returns the charts *anchored to that sheet*, in the order the
drawing declares them. `getCharts(xf)` returns those for every sheet in sheet
order, followed by any chart part the package declares but no drawing references
— an orphan left behind by an edit elsewhere. Both return an empty vector when
there is nothing to find.

[`XLSX.getChart`](@ref) fetches a single chart. It accepts the part name, with or
without its extension, the full package path, or the chart's relationship id
within its drawing part:

```julia
julia> c = XLSX.getChart(f["Data"], "chart1")
XLSX.Chart "chart1" on sheet "Data" at G9:N23
  title: "Revenue by Region"
  type: barChart
  series: 2
    [0] 2024 - Data!$B$2:$B$5 (4 pts)
    [1] 2025 - Data!$C$2:$C$5 (4 pts)

julia> XLSX.getChart(f, "xl/charts/chart1.xml") == c
true

julia> XLSX.getChart(f["Data"], "rId1") == c
true
```

Asking for a chart that isn't there throws an `XLSXError` listing the charts that
are.

## Anatomy of a chart

Three types carry the result, nested one inside the next.

[`XLSX.Chart`](@ref) is one chart part:

| Field | Meaning |
| --- | --- |
| `path` | package path, e.g. `"xl/charts/chart1.xml"` |
| `name` | part name without extension, e.g. `"chart1"` |
| `rId` | relationship id within the drawing part, if resolved |
| `sheet` | sheet the chart is anchored to, if resolved |
| `from`, `to` | anchor cells as strings, following [`XLSX.getImages`](@ref) |
| `title` | title text, or `nothing` when auto-generated or deleted |
| `charttypes` | `[:barChart]`, or several for a combo chart |
| `series` | `Vector{ChartSeries}` in document order |

[`XLSX.ChartSeries`](@ref) is one series within it, holding its `idx` and `order`
as Excel recorded them, its `charttype`, its `name`, and up to four references:
`name_ref`, `categories`, `values` and `bubble_sizes`.

```julia
julia> s = c.series[1]
XLSX.ChartSeries 2024 (barChart)
  categories: Data!$A$2:$A$5 - 4 pts, str
  values: Data!$B$2:$B$5 - 4 pts, num
```

!!! note
    `categories` holds `c:cat` for category charts and `c:xVal` for scatter and
    bubble charts; `values` holds `c:val` or `c:yVal` correspondingly. The two
    fields therefore mean the same thing whatever the chart type, and you never
    need to branch on `charttype` to find the data.

[`XLSX.ChartRef`](@ref) is a single reference — the formula it came from, the
number format Excel recorded for it, and the cached points themselves:

```julia
julia> s.values
XLSX.ChartRef Data!$B$2:$B$5 (num, 4 pts)
  format: General
  data: [10.0, 20.0, 15.0, 5.0]

julia> s.values.ref
"Data!\$B\$2:\$B\$5"

julia> s.values.kind
:num
```

`kind` is one of `:num`, `:str`, `:multiLvlStr`, `:numLit` or `:strLit`. The two
`Lit` kinds are series typed directly into the chart rather than read from cells;
they have no `ref`.

## Getting the cached data

[`XLSX.getChartData`](@ref) flattens a chart's series into an
[`XLSX.DataTable`](@ref), ready for `DataFrame` or any other Tables.jl sink:

```julia
julia> using DataFrames

julia> DataFrame(XLSX.getChartData(c))
4×3 DataFrame
 Row │ categories  2024     2025
     │ String      Float64  Float64
─────┼──────────────────────────────
   1 │ North          10.0     12.0
   2 │ South          20.0     18.0
   3 │ East           15.0     25.0
   4 │ West            5.0      9.0
```

It also takes a chart name directly, saving the intermediate `getChart` call:

```julia
julia> dt = XLSX.getChartData(f["Data"], "chart1");
```

Series that have no name in the file — the ones Excel labels "Series1",
"Series2" in the legend — are labelled by position. Duplicate names are
disambiguated with a numeric suffix. Series of unequal length are padded with
`missing`.

### How the columns are laid out

Where every series shares one category reference, a single `categories` column
leads the table, as above. Where they don't — the usual case for scatter and
bubble charts, where each series carries its own x values — each series
contributes its own `<series>_x` column immediately before its values:

```julia
julia> DataFrame(XLSX.getChartData(XLSX.readxlsx("chart_scatter.xlsx"), "chart1"))
4×4 DataFrame
 Row │ Series1_x  Series1  Series2_x  Series2
     │ Float64    Float64  Float64    Float64
─────┼──────────────────────────────────────────
   1 │       1.0     10.0        2.0      5.0
   2 │       2.0     20.0        4.0     15.0
   3 │       3.0     30.0        6.0     25.0
   4 │       4.0     40.0        8.0     35.0
```

Multi-level categories give one column per level, in the order Excel writes them
(innermost first):

```julia
julia> DataFrame(XLSX.getChartData(XLSX.readxlsx("chart_multilevel.xlsx"), "chart1"))
4×3 DataFrame
 Row │ categories_1  categories_2  Sales
     │ String        String        Float64
─────┼───────────────────────────────────────
   1 │ Q1            North             …
   2 │ Q2            North             …
   3 │ Q1            South             …
   4 │ Q2            South             …
```

A bubble chart adds a `<series>_size` column after each series' values.

No category column is produced when the chart has no `c:cat` at all: Excel is
plotting against an implicit index in that case, and caches nothing for it.

## Where the data came from

The cache tells you what the chart *displayed*; [`XLSX.getChartRanges`](@ref)
tells you where it says the data *came from*, as a range object you can hand
straight back to [`XLSX.getdata`](@ref):

```julia
julia> r = XLSX.getChartRanges(f["Data"], "chart1");

julia> r[1].name, r[1].categories, r[1].values
("2024", Data!A2:A5, Data!B2:B5)

julia> XLSX.getdata(f, r[1].values)      # the live cells, not the cache
4×1 Matrix{Any}:
 10
 20
 15
  5
```

One entry is returned per series, in document order, parallel to `c.series`. Each
carries the series `idx` and `name` alongside its `categories`, `values` and
`bubble_sizes` ranges. `bubble_sizes` is `nothing` for every chart type but
bubble.

Called with no name, `getChartRanges` covers every chart on the sheet or in the
workbook, each paired with its chart name:

```julia
julia> [(x.chart, length(x.ranges)) for x in XLSX.getChartRanges(f)]
1-element Vector{Tuple{String, Int64}}:
 ("chart1", 2)
```

A range is `nothing` wherever the series has no addressable source: a literal
series, a reference to an external workbook, or a defined name.

!!! note
    A chart may plot from a range whose sheet no longer exists, or which has since
    been overwritten with something else entirely. The range records the chart's
    claim about its source, not a guarantee about the current contents of those
    cells. Where the two disagree, the cache is the older of the pair.

## Blanks and errors

Excel writes the cache as a sparse list of points against a declared `ptCount`.
A blank cell in the source is simply omitted from that list, so it arrives as
`missing` in `data`.

Errors are a narrower story. `#N/A` is cached as the literal string `#N/A`, and
XLSX.jl records it in `errors` and stores `missing` in its place — so that it is
not mistaken for a number, and can be told apart from a blank. **Every other error
value is written into the cache as `0`.** By the time the file is on disk, a
`#DIV/0!` and a genuine zero are the same three bytes, and nothing can recover the
difference.

[`XLSX.iserror`](@ref) and [`XLSX.geterror`](@ref) report what the cache preserved,
which is `#N/A` and nothing else:

```julia
julia> g = XLSX.getChart(XLSX.readxlsx("chart_gaps.xlsx"), "chart1");

julia> v = g.series[1].values
XLSX.ChartRef Data!$B$2:$B$6 (num, 5 pts)
  format: General
  errors at: 3
  data: Union{Missing, Float64}[1.0, missing, missing, 0.0, 5.0]

julia> XLSX.iserror(v)
5-element Vector{Bool}:
 0
 0
 1
 0
 0

julia> XLSX.geterror(v)
5-element Vector{String}:
 ""
 ""
 "#N/A"
 ""
 ""
```

The source column behind this chart is worth setting out in full, because it shows
all three outcomes at once:

| Cell | Source | In the cache | In `data` |
| --- | --- | --- | --- |
| `B2` | `1` | `<c:pt idx="0">1</c:pt>` | `1.0` |
| `B3` | *(blank)* | omitted | `missing` |
| `B4` | `=NA()` → `#N/A` | `<c:pt idx="2">#N/A</c:pt>` | `missing`, flagged in `errors` |
| `B5` | `=1/0` → `#DIV/0!` | `<c:pt idx="3">0</c:pt>` | `0.0` |
| `B6` | `5` | `<c:pt idx="4">5</c:pt>` | `5.0` |

So point 2 was blank, point 3 was `#N/A` — and point 4 reads as a plain zero even
though the source cell holds `#DIV/0!`.

!!! warning
    A zero in cached chart data may be a genuine zero or may be any error other
    than `#N/A`. This is a limitation of the file format, not of XLSX.jl: Excel
    discards the distinction when it writes the cache. Where it matters, read the
    source cells through [`XLSX.getChartRanges`](@ref) and
    [`XLSX.getdata`](@ref), which see the real cell values and report every error
    type.

!!! note
    [`XLSX.iserror`](@ref) and [`XLSX.geterror`](@ref) on a cell range report every
    error value. On a `ChartRef` they can only report what survived into the cache.
    The functions behave the same way; the data they are given does not.

## Combo charts

A chart whose plot area holds more than one group — a bar series and a line
series sharing an axis, say — reports every group in `charttypes`, and each
series remembers which group it belongs to:

```julia
julia> cc = XLSX.getChart(XLSX.readxlsx("chart_combo.xlsx"), "chart1");

julia> cc.charttypes
2-element Vector{Symbol}:
 :barChart
 :lineChart

julia> [(s.name, s.charttype) for s in cc.series]
2-element Vector{Tuple{Union{Nothing, String}, Symbol}}:
 ("2024", :barChart)
 ("2025", :lineChart)
```

[`XLSX.getChartData`](@ref) is indifferent to this: series from every group land in the same
table, in document order.

## Chartsheets

A chart on its own chartsheet is found the same way, through the sheet it lives
on:

```julia
julia> f = XLSX.readxlsx("chart_chartsheet.xlsx")
XLSXFile("chart_chartsheet.xlsx") containing 2 Worksheets
            sheetname size          range
-------------------------------------------------
             TheChart Chartsheet
                 Data 5x3           A1:C5

julia> c = XLSX.getCharts(f["TheChart"])[1]
XLSX.Chart "chart1" on sheet "TheChart"
  title: "Revenue by Region"
  type: barChart
  series: 2
    [0] 2024 - Data!$B$2:$B$5 (4 pts)
    [1] 2025 - Data!$C$2:$C$5 (4 pts)
```

Note the missing anchor: a chartsheet positions its chart absolutely rather than
against a cell grid, so `from` and `to` are `nothing`. The chart's `sheet` is the
chartsheet it occupies, while its series still refer to the worksheet holding the
source data.

## Charts referring to another workbook

A series may plot from a range in a different workbook, in which case its `ref`
begins with a bracketed index into the workbook's external references:

```julia
julia> e = XLSX.getChart(XLSX.readxlsx("chart_external.xlsx"), "chart1");

julia> e.series[1].values.ref
"[1]Feuil1!\$B\$1:\$B\$10"
```

Pass `get_external_refs=true` to substitute the recorded workbook path, exactly
as [`XLSX.getFormula`](@ref) does for formulas:

```julia
julia> e = XLSX.getChart(XLSX.readxlsx("chart_external.xlsx"), "chart1";
                         get_external_refs=true);

julia> e.series[1].values.ref
"[Test2.xlsx]Feuil1!\$B\$1:\$B\$10"
```

The cached values are available either way — that is the whole point of the cache
— but `getChartRanges` returns `nothing` for such a series, since the range is not
addressable within this workbook.

## Reading metadata only

Parsing the cached points is the expensive part of reading a large chart. Where
you only need the shape of a chart — its title, types, series names, source
formulas, format codes and point counts — pass `read_cached_values=false`:

```julia
julia> c = XLSX.getChart(f["Data"], "chart1"; read_cached_values=false);

julia> c.series[1].values
XLSX.ChartRef Data!$B$2:$B$5 (num, 4 pts)
  format: General

julia> c.series[1].values.data
Any[]
```

`ptCount` is still populated; only `data` is left empty. Series names are always
read, since they are metadata rather than plotted values.

Calling [`XLSX.getChartData`](@ref) on a chart read this way throws an `XLSXError` rather than
silently returning an empty table. [`XLSX.getChartRanges`](@ref) uses `read_cached_values=false` 
internally, as it never needs the values.

## What is not supported

- **Creating or editing charts.** Reading only, for now.
- **`chartEx` charts.** Waterfall, funnel, treemap, sunburst, histogram and
  box & whisker charts use a newer schema under a different namespace and are not
  read. Where a file contains them and nothing else readable, `getCharts` warns
  and returns an empty vector; where a lookup fails, the error message says how
  many `chartEx` charts were found.
- **Chart appearance.** Colours, fonts, axis scales, gridlines, data labels,
  trendlines and legends are all left alone on read and preserved on write, but
  are not exposed.
- **Live recomputation.** Values come from the cache. Use
  [`XLSX.getChartRanges`](@ref) with [`XLSX.getdata`](@ref) to read the source
  cells as they stand now.
