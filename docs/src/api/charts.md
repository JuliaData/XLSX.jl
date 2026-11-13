# Charts

```@meta
CurrentModule = XLSX.Charts
```

Chart support lives in the `XLSX.Charts` sub-module. `using XLSX.Charts` brings
the functions below into scope; the handle and value types are public but not
exported, so refer to them as `XLSX.Charts.ChartSeries` and so on. They are 
documented on the [Chart types](chartTypes.md) page.

The names released before the sub-module existed — `getCharts`, `getChart`,
`getChartData`, `getChartRanges`, `chartType`, `chartSchema`, `AbstractChart`,
`Chart`, `ChartEx`, `ChartRef` and `ChartSeries` — are also reachable as
`XLSX.getCharts` and so on.

```@docs
XLSX.Charts
```

Guides: start at the [Charts overview](../charts/excelCharts.md), or go straight to
[Creating charts](../charts/creatingCharts.md) or [Formatting charts](../charts/formattingCharts.md).

## Conventions

The setters share their argument conventions, so the docstrings below don't
repeat them.

| argument     | meaning                                                                                   |
|:-------------|:------------------------------------------------------------------------------------------|
| `c`          | a chart, from [`getCharts`](@ref), [`getChart`](@ref), [`addChart`](@ref) or [`addChartEx`](@ref) |
| `i`          | a series' position, counting from 1, in the order [`getChartSeries`](@ref) returns them   |
| `point`      | a data point's position in its series, counting from 1                                    |
| `ax`         | a [`ChartAxis`](@ref), from [`getChartAxes`](@ref) or [`getChartAxis`](@ref)              |
| `g`          | a [`ChartGroup`](@ref), from [`getChartGroups`](@ref)                                      |
| `id`         | on a `ChartEx`, an axis's `id`, from [`getChartAxisIds`](@ref) — not a position           |
| `pareto`     | on a `ChartEx`, `true` to act on a Pareto chart's line rather than its bars               |

Every setter returns `c`. Most take `:inherit` in place of a value, to remove
the setting so that it is inherited from the chart style or from the level
above; each docstring says where it does. That is different from `:none`, which
some setters take to mean "draw nothing".

Three vocabularies are shared by many setters, and each is listed in full once:

- **Color values** — at [`setSeriesFill`](@ref): color names, hex strings,
  `Colors.Colorant`s, theme colors such as `"accent2"`, and
  [`SchemeColor`](@ref)s.
- **Line keywords** (`color`, `width`, `dash`, `cap`, `compound`, `join`,
  `miterLimit`) — at [`setSeriesLine`](@ref).
- **Text-property fields** (`:size`, `:bold`, `:fill`, `:font`, …) — at
  [`setChartTitleTextProp`](@ref).

## Finding and reading charts

```@docs
getCharts
getChart
getChartSchema
getChartType
getChartTitle
getChartData
getChartRanges
```

## The chart as a whole

```@docs
getChartTypes
getChartGroups
getChartLegend
getLegendPos
setLegendPos
getLegendOverlay
setLegendOverlay
getLegendShapeProps
getLegendTextProps
setLegendTextProp
getLegendEntryDeleted
setLegendEntryDeleted
getChartTitleText
getChartTitleRef
getChartTitleShapeProps
getChartTitleTextProps
getAutoTitleDeleted
setChartTitleText
setChartTitleTextProp
getPlotAreaShapeProps
getChartSpaceShapeProps
getChartSpaceTextProps
setChartSpaceTextProp
setPlotAreaLine
```

## Series

```@docs
getChartSeries
getSeriesGroup
getSeriesAxes
getSeriesShapeProps
getSeriesFill
getSeriesLine
getSeriesMarker
getSeriesLabelTextProps
setSeriesFill
setSeriesLine
setSeriesLineColor
setSeriesLineWidth
setSeriesLineDash
setSeriesLineCap
setSeriesLineCompound
setSeriesLineJoin
setSeriesLineMiterLimit
hasFill
hasLine
```

## Chart groups

```@docs
getGroupAxes
getGroupLabelTextProps
getGroupDropLines
getGroupHiLowLines
getGroupSeriesLines
getGroupUpDownBars
getUpBarShapeProps
getDownBarShapeProps
setGroupLabelTextProp
getGroupGapWidth
setGroupGapWidth
getGroupOverlap
setGroupOverlap
```

## Creating charts

```@docs
addChart
addChartEx
addSeries
makeCombo
deleteChart
```


## Sparklines

Sparklines are part of `XLSX` itself, not `XLSX.Charts`. Guide: [Sparklines](../charts/sparklines.md).

```@docs
XLSX.addSparklines
XLSX.getSparklines
XLSX.deleteSparklines
```
