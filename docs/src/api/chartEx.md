# chartEx charts

```@meta
CurrentModule = XLSX.Charts
```


Charts in the Microsoft `cx:` namespace — waterfall, funnel, treemap, sunburst,
histogram, Pareto, box & whisker and region map. Many of the functions above
also take a [`ChartEx`](@ref); these are specific to it. Guide: [Creating chartEx charts](../charts/creatingChartEx.md).

```@docs
getChartSeriesCount
getChartDataBlocks
getChartAxisIds
getChartTitleRange
getSeriesLayout
getSeriesLayoutFlag
getSeriesData
getSeriesName
getSeriesNameRange
getSeriesHidden
getSeriesAxisIds
getSeriesSubtotals
getSeriesBinning
getSeriesAggregation
getSeriesQuartileMethod
getSeriesParentLabelLayout
getLabelFlag(::ChartEx, ::Integer, ::Symbol)
setLabelFlag(::ChartEx, ::Integer, ::Symbol, ::Bool)
getLabelPosition
setSeriesName
setSeriesLayoutFlag
setSeriesSubtotals
setSeriesBinning
setSeriesAggregation
setSeriesQuartileMethod
setSeriesParentLabelLayout
```
