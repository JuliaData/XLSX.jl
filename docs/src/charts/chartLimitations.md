```@meta
CurrentModule = XLSX.Charts
```

!!! note "Experimental"

    Handling of native Excel charts in XLSX.jl is experimental and a work in 
    progress. All aspects may be subject to change. Feedback is welcome!


# Limitations

What chart support does not do, and why. Most of these are deliberate stopping
points rather than oversights, and each says what to do instead.

## Chart kinds with no template

A created chart starts from a chart part Excel itself wrote, kept verbatim in
the package, so the result is what Excel would have produced. That means a kind
with no template cannot be created at all, rather than being approximated:

- 3-D charts of any kind
- stock, surface and of-pie charts
- combo charts other than `:combo` (column and line on secondary axes)
- region maps (the `cx:` filled map), whose geography Excel resolves online

All of these are *read* normally, including their series and formatting and,
for the `c:` kinds, their cached data. The restriction is on creating them.

[`addSeries`](@ref) adds to the column or line group of any combo chart, through
its `group` keyword, but not to a combo's other groups, such as an area group.
A group can't be moved between the primary and secondary axes. [`makeCombo`](@ref)
makes a combo of a `:column`, `:stackedColumn` or `:line` chart only, the kinds
checked against Excel's own Change Chart Type.

## Style and color parts are read for `cx:` fills only

Excel keeps a chart's default appearance in two parts beside the chart itself,
`style1.xml` and `colors1.xml`. A `cx:` chart relies on them: Excel writes no
fill for a series or point until you format it. [`getSeriesFill`](@ref) on a
[`ChartEx`](@ref) reads them as the last rung of its cascade (`site.level ===
:style`), and reports the colour Excel draws — checked colour for colour against
Excel 365 for every layout below. It answers for box & whisker, funnel,
histogram and Pareto bars by series, and for waterfall, treemap and sunburst
points.

Everything else written at no rung still resolves to `nothing`, meaning "Excel
decides", not "this is off":

- **Lines and text** of a `cx:` chart. Excel does not draw them as the style
  part says: it draws box & whisker outlines darker than the style's colour,
  and picks a contrasting colour of its own for labels inside funnel bars.
- **Monochromatic colour schemes** (Chart Design → Change Colors →
  Monochromatic). No formula was found that reproduces Excel's shades for
  these exactly, so no answer is given rather than an approximate one.
- **Region maps and Pareto lines**, and series written with a `formatIdx`
  attribute, which Excel itself has not been seen to write.
- **`c:` charts.** These do not use the style parts at all: a `c:` series
  written without formatting is drawn from the chart's built-in style number
  (`c:style`), and Excel deletes the two parts when it saves such a chart.
  Excel itself writes a `c:` series' formatting in full, so this matters
  only for files from other programs.

The parts are preserved on write either way.

## chartEx charts carry no cached values

Excel writes no value cache into a `cx:` chart part, so
[`getChartData`](@ref) throws for a [`ChartEx`](@ref) and a chart created with
[`addChartEx`](@ref) shows nothing until Excel opens the file and computes it.
Read the source cells instead, through [`getChartRanges`](@ref) and
[`XLSX.getdata`](@ref).

## Errors in cached data

Excel caches `#N/A` as itself and every other error value as `0`, so a
`#DIV/0!` in the source is indistinguishable from a genuine zero once the file
is written. This is a property of the format, not of the package. Where the
distinction matters, read the source cells — see
[Blanks and errors](readingCharts.md#Blanks-and-errors).

## Values are the ones Excel last wrote

Nothing is recomputed. A chart's cached values are those of the last save by
Excel, and a series added by [`addSeries`](@ref) caches the cells as they stand
at that moment. Changing a cell afterwards does not update any chart that plots
it; Excel does that when it next opens the file.

The same goes for formulas. A cell written by [`XLSX.setFormula`](@ref) has no
value until Excel calculates it, so a series plotting it is cached with no
values, and [`getChartData`](@ref) finds none. Excel recalculates on opening
and draws the chart from the results, so the chart looks right there. To read
the values in Julia, compute them there too, or read them back from a file Excel
has saved.

## Legend entries follow series order

Excel draws a legend entry for each series, in series order, and can hide
entries but not reorder them. So [`setLegendEntryDeleted`](@ref) hides entries,
and there is no setter for their order. To show a legend in another order, add
series that plot zeros, in the order wanted and formatted to match, and hide the
entries of the series that draw the data;
[A Likert chart](../examples.md#A-diverging-stacked-bar-chart-for-Likert-scale-survey-data)
shows how.

## Lines and dashes: symbols in, strings out

Dash, cap, compound and join are set with symbols (`:sysDot`, `:rnd`) and read
back as strings (`"sysDot"`, `"rnd"`), because a getter reports what the file
holds. An Excel name you set comes back as its DrawingML equivalent.

## Markers belong to some chart kinds only

A marker is a property of a line, scatter or radar series. Setting one on a bar
or pie series throws, because the schema gives `c:barSer` no `c:marker` child.
The same applies to the options [`addSeries`](@ref) takes: `smooth` is for line
and scatter series, `line` for scatter, and `color` has no meaning on a pie or
doughnut chart, whose points are colored individually.

## Resolving text properties

[`getLabelTextProp`](@ref) resolves one run property up the cascade for a
series' data labels. There is no equivalent for a chart title, an axis title or
the legend — only the plural `…TextProps` getters, which report what one element
says. Setting is symmetrical; the gap is on the reading side.