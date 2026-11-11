```@meta
CurrentModule = XLSX.Charts
```

!!! note "Experimental"

    Handling of native Excel charts in XLSX.jl is experimental and a work in 
    progress. All aspects may be subject to change. Feedback is welcome!


# Charts

!!! tip "Start here: create a chart"

    ```julia
    using XLSX, XLSX.Charts

    f = XLSX.newxlsx()
    ws = f[1]
    ws["A1"] = ["Region", "2024"]
    ws["A2:A5"] = ["North", "South", "East", "West"]
    ws["B2:B5"] = [10.0, 20.0, 15.0, 5.0]

    c = addChart(ws, :column; anchor = "D2:K17", title = "Revenue")
    addSeries(c, "B2:B5"; categories = "A2:A5", name_ref = "B1")

    XLSX.writexlsx("revenue.xlsx", f)
    ```

    [Creating charts](creatingCharts.md) explains each step and lists the chart kinds;
    [Creating chartEx charts](creatingChartEx.md) covers waterfall, funnel, treemap,
    histogram and the other newer kinds; [Formatting charts](formattingCharts.md)
    changes their colors, titles, legends and axes. The rest of this page explains the
    ideas the chart API is built on.

Excel charts come in two schemas, and XLSX.jl reads, formats and creates both.

- The **`c:` schema** (ECMA-376) covers bar, column, line, pie, doughnut, area,
  scatter, bubble, radar and stock charts — everything Excel has had for a long
  time. These are [`Chart`](@ref).
- The **`cx:` schema** ("chartEx", a Microsoft extension) covers waterfall,
  funnel, treemap, sunburst, histogram, Pareto, box & whisker and region map
  (Excel's filled map). These are [`ChartEx`](@ref). All are read; all but
  region map can be created.

Chart support lives in its own sub-module. Qualify the functions as
`XLSX.Charts.getCharts`, or bring them into scope:

```julia
julia> using XLSX, XLSX.Charts
```

The rest of this guide uses the unqualified form. The handle and value types are
not exported, so they stay qualified: `XLSX.Charts.ChartSeries`.

This page covers the ideas the rest of the guide assumes. The other pages are
[Reading charts](readingCharts.md), [Formatting charts](formattingCharts.md),
[Creating charts](creatingCharts.md), [Creating chartEx charts](creatingChartEx.md),
[Sheets with charts](chartSheets.md) and [Limitations](chartLimitations.md).

## Finding the charts

[`getCharts`](@ref) lists the charts anchored to a worksheet, or every chart in
the workbook:

```julia
julia> f = XLSX.readxlsx("chart_kinds.xlsx");

julia> getCharts(f["line"])          # this sheet
1-element Vector{XLSX.Charts.AbstractChart}:
 XLSX.Charts.Chart("chart8", "line", F9, lineChart, 3 series)

julia> getCharts(f)                  # the whole workbook
12-element Vector{XLSX.Charts.AbstractChart}:
 XLSX.Charts.Chart("chart1", "radar", F9, radarChart, 3 series)
 XLSX.Charts.Chart("chart2", "bubble", F9, bubbleChart, 1 series)
 XLSX.Charts.Chart("chart3", "scatter", F9, scatterChart, 2 series)
 XLSX.Charts.Chart("chart4", "doughnut", F9, doughnutChart, 1 series)
 XLSX.Charts.Chart("chart5", "pie", F9, pieChart, 1 series)
 XLSX.Charts.Chart("chart6", "area", F9, areaChart, 3 series)
 XLSX.Charts.Chart("chart7", "linemarkers", F9, lineChart, 3 series)
 XLSX.Charts.Chart("chart8", "line", F9, lineChart, 3 series)
 XLSX.Charts.Chart("chart9", "stacked", F9, barChart, 3 series)
 XLSX.Charts.Chart("chart10", "bar", F9, barChart, 3 series)
 XLSX.Charts.Chart("chart11", "column", F9, barChart, 3 series)
 XLSX.Charts.Chart("chart12", "stackedbar", F19, barChart, 3 series)
```

`getCharts(ws)` returns them in the order the drawing declares them.
`getCharts(xf)` returns those for every sheet in sheet order, followed by any
chart part the package declares that no drawing references — an orphan left
behind by an edit elsewhere. Both give an empty vector when there is nothing to
find.

[`getChart`](@ref) fetches one, by part name with or without its extension, by
package path, or by relationship id within the drawing:

```julia
julia> c = getChart(f["line"], "chart8")
XLSX.Charts.Chart "chart8" on sheet "line" at F9:L23
  type: lineChart
  series: 3
    [1] Alpha - line!$B$2:$B$5 (4 pts)
    [2] Beta - line!$C$2:$C$5 (4 pts)
    [3] Gamma - line!$D$2:$D$5 (4 pts)

julia> getChart(f, "xl/charts/chart8.xml") == c
true

julia> c.rId
"rId1"

julia> getChart(f["line"], "rId1") == c
true
```

Asking for a chart that isn't there throws an `XLSXError` listing the ones that
are.

## Which schema is this?

A workbook can hold both kinds, so `getCharts` returns a vector that may mix
`Chart` and `ChartEx`. [`getChartSchema`](@ref) tells them apart, and
[`getChartType`](@ref) names the kind:

```julia
julia> getChartSchema(c), XLSX.Charts.getChartType(c)
(:c, :lineChart)
```

Most of the API takes either. Where it doesn't, it's because the schemas
genuinely differ — a `cx:` chart has no cached values, for instance — and the
guide says so.

## Handles never go stale

A `Chart` is a handle: it holds the package and the part's path, and nothing
else. Every accessor reads the part as it stands at that moment, and every
setter returns the handle unchanged.

```julia
julia> setChartTitleText(c, "Revenue 2024-25") === c
true
```

So a handle read before a write still works afterwards, and setters chain
naturally. You never need to re-fetch a chart after modifying it.

## Values carry keys

Everything read *from* a chart — a series, an axis, a group, a data point, a
label, a trendline, error bars, a marker — is a value rather than a handle. Each
carries the key that identifies its element:

| Value | Key |
| --- | --- |
| [`ChartSeries`](@ref XLSX.Charts.ChartSeries) | `c:idx` |
| [`ChartAxis`](@ref XLSX.Charts.ChartAxis) | `c:axId` |
| [`ChartGroup`](@ref XLSX.Charts.ChartGroup) | its tag and axis ids |
| [`ChartDataPoint`](@ref XLSX.Charts.ChartDataPoint), [`ChartDataLabel`](@ref XLSX.Charts.ChartDataLabel) | owning series and point index |
| [`ChartTrendline`](@ref XLSX.Charts.ChartTrendline), [`ChartErrorBars`](@ref XLSX.Charts.ChartErrorBars) | owning series and ordinal |

Any function taking `(c, x)` finds `x`'s element by that key in the current
part. So a value read before a write still addresses the right element
afterwards, even though its other fields describe the part as it was when read:

```julia
julia> ax = getChartAxes(c, :value)[1];

julia> setSeriesFill(c, 1, "FF0000");        # unrelated write

julia> getAxisTickLabelPos(c, ax)            # `ax` still resolves
:nextTo
```

Re-read the value when you want its own fields refreshed; keep using it as an
argument either way.

## Positions count from 1

A series position, and a data point position, count from 1 as everywhere else in
the package. Excel's own identifiers are something else:

```julia
julia> [s.idx for s in getChartSeries(c)]    # c:idx, as the file records it
3-element Vector{Int64}:
 0
 1
 2
```

They usually agree, because Excel numbers series from 0 in order. They need not:
a file where series were deleted can have gaps. Positions are resolved in
document order on each call, so they are always what you see in
`getChartSeries(c)`.

## Absent is not the same as off

Optional values are `Union{Nothing,T}`, and `nothing` means the file says
nothing — so Excel inherits from the chart style. It never means "off". Turning
something off is a setting in its own right:

| You want | You write | Reads back as |
| --- | --- | --- |
| this color | `setSeriesFill(c, 1, "FF0000")` | a fill with that color |
| deliberately invisible | `setSeriesFill(c, 1, :none)` | a fill with `kind === :none` |
| let the style decide | `setSeriesFill(c, 1, :inherit)` | `nothing` |

The same three states apply to lines, markers, labels and titles.

## Formatting resolves up a cascade

Excel's chart formatting inherits: a data point falls back to its series, a
series to its group, a group to the plot area and the chart space. The resolvers
walk that chain and return an [`Effective`](@ref XLSX.Charts.Effective), which
carries the value, the rung that supplied it, and the whole chain:

```julia
julia> e = getSeriesFill(c, 1);

julia> e.value.fgcolor.rgb
"FF0000"

julia> e.site.level
:series
```

A `value` of `nothing` means no rung wrote the property, so Excel decides it
from the chart's style. For the fills of a `cx:` chart that style is read as a
last rung, so `value` is the colour Excel draws. [Formatting
charts](formattingCharts.md) covers this in full.

## Created charts follow the workbook's theme

Charts made by [`addChart`](@ref) and [`addChartEx`](@ref) start from Excel's own
default chart parts, with the colors written as theme references. So a new
chart looks like one Excel would have made in that workbook — which means a
chart created in a workbook from `XLSX.newxlsx()` shows that template's older
Office colors rather than current Excel's.