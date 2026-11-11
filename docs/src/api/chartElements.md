# Chart axes and elements

```@meta
CurrentModule = XLSX.Charts
```

Guide: [Formatting charts](../charts/formattingCharts.md), which covers axes, gridlines,
markers and data labels.

## Axes

```@docs
getChartAxes
getChartAxis
getAxisPartner
getAxisShapeProps
getAxisTextProps
getAxisGridlines
getAxisTitle
getAxisTitleText
getAxisTitleRef
getAxisTitleRange
getAxisTitleTextProps
setAxisTitleText
setAxisTitleTextProp
getAxisNumberFormatCode
getAxisNumberFormatLinked
getAxisMajorTickMark
getAxisMinorTickMark
getAxisTickLabelPos
getAxisOrientation
getAxisMin
getAxisMax
getAxisLogBase
getAxisCrosses
getAxisCrossesAt
getAxisCrossBetween
getAxisMajorUnit
getAxisMinorUnit
getAxisLabelAlign
getAxisLabelOffset
getAxisMultiLevelLabels
setAxisTickLabelPos
setAxisScaling
setAxisMin
setAxisMax
setAxisOrientation
setAxisCrosses
setAxisCrossesAt
setAxisNumberFormatCode
setAxisNumberFormatLinked
setAxisMajorTickMark
setAxisMinorTickMark
setAxisGridlines
setAxisTextProp
setAxisLine
setAxisMajorUnit
setAxisMinorUnit
```

## Data points, markers and labels

```@docs
getSeriesDataPoints
getSeriesDataPoint
getDataPointShapeProps
getDataPointMarker
getMarkerFill
setMarker
setMarkerSymbol
setMarkerSize
setMarkerFill
setMarkerLineColor
setMarkerLineWidth
getSeriesDataLabels
getSeriesDataLabel
getDataLabelText
getDataLabelTextProps
getDataLabelShapeProps
getDataLabelPosition
getDataLabelOffset
getLabelTextProp
setLabelText
setLabelTextProp
setLabelDeleted
getLabelFlag(::Chart, ::Integer, ::Symbol)
setLabelFlag(::Chart, ::Integer, ::Symbol, ::Bool)
```

## Trendlines and error bars

```@docs
getSeriesTrendlines
getTrendlineShapeProps
getTrendlineLabelText
getTrendlineLabelTextProps
getTrendlineLabelShapeProps
getSeriesErrorBars
getErrorBarsShapeProps
getErrorBarsCustomRefs
```

