#
# Tests for the chart appearance accessors in `src/charts/chartnodes.jl` and the other chart*.jl files it lists — stage 3.
#
# These test node-finding, not DrawingML parsing: that a given accessor reaches
# the right element and hands it to the right parser. DrawingML parsing itself
# is covered in drawings/drawingml.jl, and chart discovery and data extraction in
# charts/charts.jl. A failure here should mean "the accessor looked in the wrong
# place", not "the colour resolved wrongly".
#
# A few theme-colour values are asserted anyway, because a colour reached
# through an accessor exercises the whole chain, and a regression in either
# layer should fail in both files for different reasons.
#
# ---------------------------------------------------------------------------
# Fixture inventory
#
# chart_appearance.xlsx — built in Excel, one sheet, 3x3 of small integers.
# Every element below is deliberate; if the file is ever regenerated, it needs
# all of this or the tests lose their point.
#
#   Combo chart, two groups:
#     c:barChart   series 1 "2024", series 2 "2025"   axIds 612078287, 460195247
#     c:lineChart  series 3 "2026"                    axIds 1926562432, 1773317264
#
#   Four axes, exercising role filtering, id lookup and the deleted case:
#     catAx 612078287  @b  title "Horizontal"  spPr with noFill + real a:ln
#     valAx 460195247  @l  title "Primary"     majorGridlines with spPr
#     valAx 1773317264 @r  title "Secondary"   majorTickMark="out", crosses="max"
#     catAx 1926562432 @b  delete="1"          no spPr, no txPr, no title
#
#   No axis title is bound to a cell, so getAxisTitleRef is nothing throughout —
#   as is getChartTitleRef. Bound titles are written and read back in the
#   "rich text and cell-bound titles" testset, which builds its own workbook.
#
#   Series 1: spPr accent1 (a:ln present but noFill — hasLine false because the
#     line has no fill, not because the element is absent). Series-level
#     c:dLbls with txPr, sz 1050, accent1 lumMod 75000 -> 104862. One c:dPt at
#     idx 1 filled plain red FF0000. Three c:dLbl —
#       idx 0  retyped "Best", carries BOTH c:tx/c:rich and c:txPr in 00B0F0,
#              and an a:r/a:rPr overriding the paragraph defRPr
#       idx 1  dragged: c:layout/c:manualLayout, no c:dLblPos
#       idx 2  c:delete val="1" and nothing else
#     Plus a linear c:trendline, dispRSqr and dispEq both 1, sysDot dash at
#     1.5pt, with a c:trendlineLbl carrying spPr and txPr but no c:tx.
#
#   Series 2: spPr accent2. Series-level c:dLbls with showVal="1" and a txPr at
#     sz 900, tx1 lumMod 75000 / lumOff 25000 -> 404040 — a different level from
#     series 1, so the two labelled series distinguish rather than duplicate.
#     Fixed-value error bars, c:val val="0.5", no c:errDir (Excel omits it on a
#     bar chart).
#
#   Series 3: line, spPr with a:ln only and no fill element at all — the mirror
#     of series 1. c:marker diamond size 9 with its own spPr. One c:dPt at idx 2
#     whose point-level spPr is a 2.25pt line with no fill (the line segment),
#     and whose c:marker/c:spPr carries fill FF0000 and line 00B050 — so the
#     "fill of a data point" lives on the marker, not the point.
#
#   Chart space: solid bg1 fill and a real line at 0.75pt — the only element in
#   the file with both. Legend @b, txPr sz 900. Chart title with spPr and txPr
#   (sz 1400, tx1 lumMod 65000 / lumOff 35000 -> 595959) but no c:tx, so Excel
#   generates the text. c:chartSpace/c:txPr present but empty, which is the top
#   rung of the inheritance cascade saying nothing.
#
#   Not present, so untested here: c:dropLines, c:hiLowLines, c:upDownBars,
#   c:serLines, c:bandFmts; a c:majorGridlines with no c:spPr (the three-state
#   "on but unformatted" case); a polynomial or moving-average trendline, which
#   would exercise c:order and c:period; custom error bars with c:plus/c:minus.
#
# chart_basic.xlsx        — 2-series bar chart, title "Revenue by Region" as
#                           literal c:rich text, so getChartTitleText and
#                           parse_chart_title can be cross-checked.
# chart_theme_colors.xlsx — 6 series: accent1 (156082) and five variants,
#                           series 5 being accent1 lumMod 75000 -> 104862, the
#                           same value series 1's labels carry in the fixture
#                           above by a different route.
# chart_mixed.xlsx        — a c: chart and a cx: chart on one sheet; the c:
#                           accessors are typed to Chart and must not accept a
#                           ChartEx.
# ---------------------------------------------------------------------------

@testset "ChartProps" begin

    f = XLSX.readxlsx(joinpath(data_directory, "chart_appearance.xlsx"))
    c = XLSX.Charts.getCharts(f)[1]

    @testset "reaching the XML" begin
        root = XLSX.Charts.chart_root(c)
        @test XLSX.localname(root) == "chartSpace"

        # Node identity is the write handle: the accessor must return the node
        # that lives in the tree writexlsx serializes, not a copy. Everything
        # in stage 4 depends on this.
        a = XLSX.Charts.getSeriesShapeProps(c, 1)
        b = XLSX.Charts.getSeriesShapeProps(c, 1)
        @test a.raw === b.raw
        @test a.raw === XLSX.first_element_with_tag(XLSX.Charts.getChartSeries(c)[1].raw, "spPr")

        # Two Chart objects from two getCharts calls share their nodes.
        c2 = XLSX.Charts.getCharts(f)[1]
        @test XLSX.Charts.getSeriesShapeProps(c2, 1).raw === a.raw

        @test_throws XLSX.XLSXError XLSX.Charts.getSeriesShapeProps(c, 4)
        @test_throws XLSX.XLSXError XLSX.Charts.getSeriesShapeProps(c, 0)
    end

    @testset "series appearance" begin
        @test length(XLSX.Charts.getChartSeries(c)) == 3

        sp1 = XLSX.Charts.getSeriesShapeProps(c, 1)
        @test !isnothing(sp1)
        @test sp1.fill.kind == :solid
        @test sp1.fill.fgcolor.rgb == "156082"          # accent1, verified in Excel
        @test XLSX.Charts.hasFill(sp1)
        @test !XLSX.Charts.hasLine(sp1)                        # a:ln holds only a:noFill
        @test !isnothing(sp1.line)
        @test sp1.line.fill.kind == :none

        # Series 3 is a line: a:ln present, no fill element. The mirror image.
        sp3 = XLSX.Charts.getSeriesShapeProps(c, 3)
        @test XLSX.Charts.hasLine(sp3)
        @test !XLSX.Charts.hasFill(sp3)
        @test isnothing(sp3.fill)

        # hasFill / hasLine agree across every argument type
        for i in 1:3
            sp = XLSX.Charts.getSeriesShapeProps(c, i)
            f, l = XLSX.Charts.getSeriesFill(c, i), XLSX.Charts.getSeriesLine(c, i)
            @test XLSX.Charts.hasFill(f) == XLSX.Charts.hasFill(sp)
            @test XLSX.Charts.hasLine(l) == XLSX.Charts.hasLine(sp)
            isnothing(f.value) || @test XLSX.Charts.hasFill(f.value) == XLSX.Charts.hasFill(sp)
            isnothing(l.value) || @test XLSX.Charts.hasLine(l.value) == XLSX.Charts.hasLine(sp)
        end
    end

    @testset "series data labels" begin
        tx1 = XLSX.Charts.getSeriesLabelTextProps(c, 1)
        @test !isnothing(tx1)
        rp = XLSX.Charts.default_run_props(tx1)
        @test rp.size == 10.5                            # sz="1050", hundredths -> points
        @test rp.fill.fgcolor.rgb == "104862"            # accent1 + lumMod 75000

        # Series 2 has its own series-level dLbls; series 3 has none.
        tx2 = XLSX.Charts.getSeriesLabelTextProps(c, 2)
        @test !isnothing(tx2)
        rp2 = XLSX.Charts.default_run_props(tx2)
        @test rp2.size == 9.0
        @test rp2.fill.fgcolor.rgb == "404040"   # tx1 + lumMod 75000 / lumOff 25000
        @test isnothing(XLSX.Charts.getSeriesLabelTextProps(c, 3))
    end

    @testset "series markers" begin
        @test isnothing(XLSX.Charts.getSeriesMarker(c, 1))        # bar series
        @test isnothing(XLSX.Charts.getSeriesMarker(c, 2))

        mk = XLSX.Charts.getSeriesMarker(c, 3)
        @test mk.symbol == :diamond
        @test mk.size == 9
        @test !isnothing(mk.shape)
        @test XLSX.Charts.hasFill(mk.shape)
    end

    @testset "chart groups" begin
        gs = XLSX.Charts.getChartGroups(c)
        @test length(gs) == 2
        @test [g.kind for g in gs] == [:barChart, :lineChart]
        @test gs[1].axids == [612078287, 460195247]
        @test gs[2].axids == [1926562432, 1773317264]

        # Group-level dLbls carry only the show* flags here, no txPr.
        @test isnothing(XLSX.Charts.getGroupLabelTextProps(c, gs[1]))
        @test isnothing(XLSX.Charts.getGroupLabelTextProps(c, gs[2]))

        @test [a.kind for a in XLSX.Charts.getGroupAxes(c, gs[1])] == [:catAx, :valAx]
        @test [a.axid for a in XLSX.Charts.getGroupAxes(c, gs[2])] == [1926562432, 1773317264]

        # The group is what ties a series to its axes.
        @test XLSX.Charts.getSeriesGroup(c, 1).kind == :barChart
        @test XLSX.Charts.getSeriesGroup(c, 3).kind == :lineChart
        @test [a.pos for a in XLSX.Charts.getSeriesAxes(c, 1)] == [:b, :l]
        @test [a.pos for a in XLSX.Charts.getSeriesAxes(c, 3)] == [:b, :r]
    end

    @testset "axes: identity and lookup" begin
        axes = XLSX.Charts.getChartAxes(c)
        @test length(axes) == 4
        @test [a.kind for a in axes] == [:catAx, :valAx, :valAx, :catAx]
        @test [a.axid for a in axes] == [612078287, 460195247, 1773317264, 1926562432]
        @test [a.pos for a in axes] == [:b, :l, :r, :b]

        # A combo chart has two value axes, which is why this returns a vector.
        @test length(XLSX.Charts.getChartAxes(c, :value)) == 2
        @test length(XLSX.Charts.getChartAxes(c, :category)) == 2
        @test isempty(XLSX.Charts.getChartAxes(c, :date))
        @test_throws XLSX.XLSXError XLSX.Charts.getChartAxes(c, :nonsense)

        @test XLSX.Charts.getChartAxis(c, 1773317264).pos == :r
        @test_throws XLSX.XLSXError XLSX.Charts.getChartAxis(c, 1)

        # crossAx forms two closed pairs.
        @test XLSX.Charts.getAxisPartner(c, axes[1]).axid == 460195247
        @test XLSX.Charts.getAxisPartner(c, axes[2]).axid == 612078287
        @test XLSX.Charts.getAxisPartner(c, axes[3]).axid == 1926562432
        @test XLSX.Charts.getAxisPartner(c, axes[4]).axid == 1773317264
    end

    @testset "axes: deleted" begin
        del = XLSX.Charts.getChartAxis(c, 1926562432)
        @test del.deleted
        @test all(!a.deleted for a in XLSX.Charts.getChartAxes(c) if a.axid != 1926562432)

        # Excel strips formatting from a deleted axis but keeps its structure.
        @test isnothing(XLSX.Charts.getAxisShapeProps(c, del))
        @test isnothing(XLSX.Charts.getAxisTextProps(c, del))
        @test isnothing(XLSX.Charts.getAxisTitleText(c, del))
        @test isnothing(XLSX.Charts.getAxisGridlines(c, del))
        @test XLSX.Charts.getAxisLabelOffset(c, del) == 100      # scalars survive
        @test !isnothing(XLSX.Charts.getAxisPartner(c, del))
    end

    @testset "axes: appearance" begin
        cat, pri, sec = XLSX.Charts.getChartAxis(c, 612078287),
        XLSX.Charts.getChartAxis(c, 460195247),
        XLSX.Charts.getChartAxis(c, 1773317264)

        @test XLSX.Charts.text_content(XLSX.Charts.getAxisTitleText(c, cat)) == "Horizontal"
        @test XLSX.Charts.text_content(XLSX.Charts.getAxisTitleText(c, pri)) == "Primary"
        @test XLSX.Charts.text_content(XLSX.Charts.getAxisTitleText(c, sec)) == "Secondary"

        # No fixture has a title bound to a cell.
        @test isnothing(XLSX.Charts.getAxisTitleRef(c, cat))

        # noFill plus a real line, versus noFill on both: same has_* answers
        # from different XML, and both explicit rather than absent.
        spc = XLSX.Charts.getAxisShapeProps(c, cat)
        @test !XLSX.Charts.hasFill(spc) && XLSX.Charts.hasLine(spc)
        @test spc.fill.kind == :none                     # explicit <a:noFill/>

        spp = XLSX.Charts.getAxisShapeProps(c, pri)
        @test !XLSX.Charts.hasFill(spp) && !XLSX.Charts.hasLine(spp)

        @test !isnothing(XLSX.Charts.getAxisGridlines(c, pri))
        @test isnothing(XLSX.Charts.getAxisGridlines(c, cat))
        @test isnothing(XLSX.Charts.getAxisGridlines(c, pri; minor=true))
    end

    @testset "axes: scalars" begin
        cat, pri, sec = XLSX.Charts.getChartAxis(c, 612078287),
        XLSX.Charts.getChartAxis(c, 460195247),
        XLSX.Charts.getChartAxis(c, 1773317264)

        @test XLSX.Charts.getAxisNumberFormatCode(c, cat) == "General"
        @test XLSX.Charts.getAxisNumberFormatLinked(c, cat) === true

        @test XLSX.Charts.getAxisMajorTickMark(c, cat) == :none
        @test XLSX.Charts.getAxisMajorTickMark(c, sec) == :out
        @test XLSX.Charts.getAxisMinorTickMark(c, cat) == :none
        @test XLSX.Charts.getAxisTickLabelPos(c, cat) == :nextTo

        @test XLSX.Charts.getAxisOrientation(c, cat) == :minMax
        @test isnothing(XLSX.Charts.getAxisMin(c, pri))           # automatic scaling
        @test isnothing(XLSX.Charts.getAxisMax(c, pri))
        @test isnothing(XLSX.Charts.getAxisLogBase(c, pri))

        @test XLSX.Charts.getAxisCrosses(c, pri) == :autoZero
        @test XLSX.Charts.getAxisCrosses(c, sec) == :max
        @test isnothing(XLSX.Charts.getAxisCrossesAt(c, pri))    # mutually exclusive

        @test isnothing(XLSX.Charts.getAxisMajorUnit(c, pri))
        @test isnothing(XLSX.Charts.getAxisMinorUnit(c, pri))

        @test XLSX.Charts.getAxisLabelAlign(c, cat) == :ctr
        @test XLSX.Charts.getAxisLabelOffset(c, cat) == 100
        @test XLSX.Charts.getAxisMultiLevelLabels(c, cat) === true   # noMultiLvlLbl="0"
        @test XLSX.Charts.getAxisCrossBetween(c, pri) == :between

        # Kind-specific accessors throw rather than returning nothing, so that
        # nothing keeps meaning "not written".
        @test_throws XLSX.XLSXError XLSX.Charts.getAxisCrossBetween(c, cat)
        @test_throws XLSX.XLSXError XLSX.Charts.getAxisLabelOffset(c, pri)
        @test_throws XLSX.XLSXError XLSX.Charts.getAxisLabelAlign(c, pri)
        @test_throws XLSX.XLSXError XLSX.Charts.getAxisMajorUnit(c, cat)
    end

    @testset "chart space level" begin
        # Title: txPr present, no c:tx at all — Excel generates the text.
        @test !isnothing(XLSX.Charts.getChartTitleNode(c))
        @test isnothing(XLSX.Charts.getChartTitleText(c))
        @test isnothing(XLSX.Charts.getChartTitleRef(c))
        @test XLSX.Charts.getAutoTitleDeleted(c) === false
        ttp = XLSX.Charts.getChartTitleTextProps(c)
        @test XLSX.Charts.default_run_props(ttp).size == 14.0
        @test XLSX.Charts.default_run_props(ttp).fill.fgcolor.rgb == "595959"  # tx1 +lumMod/lumOff

        @test XLSX.Charts.getLegendPos(c) == :b
        @test XLSX.Charts.getLegendOverlay(c) === false
        @test XLSX.Charts.default_run_props(XLSX.Charts.getLegendTextProps(c)).size == 9.0

        pa = XLSX.Charts.getPlotAreaShapeProps(c)
        @test !XLSX.Charts.hasFill(pa) && !XLSX.Charts.hasLine(pa)

        # The only element in the file with both a real fill and a real line.
        cs = XLSX.Charts.getChartSpaceShapeProps(c)
        @test XLSX.Charts.hasFill(cs) && XLSX.Charts.hasLine(cs)
        @test cs.fill.fgcolor.rgb == "FFFFFF"            # bg1 -> lt1
        @test cs.line.width == 0.75                      # w="9525" EMU -> points

        # The top of the cascade: present, but says nothing.
        cst = XLSX.Charts.getChartSpaceTextProps(c)
        @test !isnothing(cst)
        @test isnothing(XLSX.Charts.default_run_props(cst).size)
    end

    @testset "data points" begin
        @test length(XLSX.Charts.getSeriesDataPoints(c, 1)) == 1
        @test isempty(XLSX.Charts.getSeriesDataPoints(c, 2))
        @test length(XLSX.Charts.getSeriesDataPoints(c, 3)) == 1

        d1 = XLSX.Charts.getSeriesDataPoint(c, 1, 2)             # 1-based position
        @test !isnothing(d1)
        @test d1.idx == 1                                # c:idx is 0-based
        @test d1.invert_if_negative === false            # written by Excel
        @test XLSX.Charts.getDataPointShapeProps(c, d1).fill.fgcolor.rgb == "FF0000"
        @test isnothing(XLSX.Charts.getDataPointMarker(c, d1))

        @test isnothing(XLSX.Charts.getSeriesDataPoint(c, 1, 1))
        @test isnothing(XLSX.Charts.getSeriesDataPoint(c, 1, 3))
        @test_throws XLSX.XLSXError XLSX.Charts.getSeriesDataPoint(c, 1, 0)

        # A point that overrides only its marker: the point-level spPr is the
        # line segment, which has no fill; the colour lives on the marker.
        d3 = XLSX.Charts.getSeriesDataPoint(c, 3, 3)
        @test d3.idx == 2
        @test isnothing(d3.invert_if_negative)           # meaningless on a line
        sp = XLSX.Charts.getDataPointShapeProps(c, d3)
        @test isnothing(sp.fill)                         # absent, not noFill
        @test sp.line.width == 2.25
        mk = XLSX.Charts.getDataPointMarker(c, d3)
        @test mk.symbol == :diamond && mk.size == 9
        @test mk.shape.fill.fgcolor.rgb == "FF0000"

        # idx round-trips through the 1-based lookup to the same node.
        for i in 1:length(XLSX.Charts.getChartSeries(c)), d in XLSX.Charts.getSeriesDataPoints(c, i)
            @test XLSX.Charts.getSeriesDataPoint(c, i, d.idx + 1).raw === d.raw
        end
    end

    @testset "individual data labels" begin
        dls = XLSX.Charts.getSeriesDataLabels(c, 1)
        @test length(dls) == 3
        @test [d.idx for d in dls] == [0, 1, 2]
        @test isempty(XLSX.Charts.getSeriesDataLabels(c, 2))

        # Retyped label: literal text, and formatting in both c:rich and c:txPr.
        r = XLSX.Charts.getSeriesDataLabel(c, 1, 1)
        @test XLSX.Charts.text_content(XLSX.Charts.getDataLabelText(c, r)) == "Best"
        @test !isnothing(XLSX.Charts.getDataLabelTextProps(c, r))
        @test XLSX.Charts.default_run_props(XLSX.Charts.getDataLabelText(c, r)).fill.fgcolor.rgb == "00B0F0"
        @test isnothing(XLSX.Charts.getDataLabelOffset(c, r))
        @test isnothing(XLSX.Charts.getDataLabelPosition(c, r))

        # Dragged label: a manual layout offset, no dLblPos.
        m = XLSX.Charts.getSeriesDataLabel(c, 1, 2)
        off = XLSX.Charts.getDataLabelOffset(c, m)
        @test off.x ≈ -0.038888888888888994
        @test off.y ≈ -0.06481481481481485
        @test isnothing(XLSX.Charts.getDataLabelPosition(c, m))
        @test isnothing(XLSX.Charts.getDataLabelText(c, m))

        # Deleted label: c:delete and nothing else.
        x = XLSX.Charts.getSeriesDataLabel(c, 1, 3)
        @test x.delete === true
        @test isnothing(XLSX.Charts.getDataLabelText(c, x))
        @test isnothing(XLSX.Charts.getDataLabelTextProps(c, x))

        # Absent delete means shown, and is distinct from an explicit false.
        @test r.delete === nothing
        @test m.delete === nothing
        @test length(filter(d -> d.delete !== true, dls)) == 2
    end

    @testset "trendlines" begin
        @test isempty(XLSX.Charts.getSeriesTrendlines(c, 2))
        ts = XLSX.Charts.getSeriesTrendlines(c, 1)
        @test length(ts) == 1

        t = ts[1]
        @test t.kind == :linear
        @test t.disp_rsqr === true
        @test t.disp_eq === true
        # A plain linear fit writes almost nothing else.
        @test isnothing(t.name)
        @test isnothing(t.order)
        @test isnothing(t.period)
        @test isnothing(t.forward)
        @test isnothing(t.backward)
        @test isnothing(t.intercept)

        sp = XLSX.Charts.getTrendlineShapeProps(c, t)
        @test sp.line.width == 1.5                       # w="19050"
        @test sp.line.dash == "sysDot"                   # element, not attribute
        @test isnothing(sp.fill)

        # The label carries formatting but no typed text.
        @test isnothing(XLSX.Charts.getTrendlineLabelText(c, t))
        @test !isnothing(XLSX.Charts.getTrendlineLabelTextProps(c, t))
        @test !isnothing(XLSX.Charts.getTrendlineLabelShapeProps(c, t))
    end

    @testset "error bars" begin
        @test isempty(XLSX.Charts.getSeriesErrorBars(c, 1))
        es = XLSX.Charts.getSeriesErrorBars(c, 2)
        @test length(es) == 1

        e = es[1]
        @test isnothing(e.direction)                     # omitted on a bar chart
        @test e.bar_type == :both
        @test e.value_type == :fixedVal
        @test e.value == 0.5
        @test e.no_end_cap === false

        sp = XLSX.Charts.getErrorBarsShapeProps(c, e)
        @test XLSX.Charts.hasLine(sp)
        @test sp.fill.kind == :none                      # explicit noFill

        # Only cust bars carry plus/minus references.
        @test XLSX.Charts.getErrorBarsCustomRefs(c, e) == (plus=nothing, minus=nothing)
    end

    @testset "titles, groups and theme colors in other fixtures" begin
        fb = XLSX.readxlsx(joinpath(data_directory, "chart_basic.xlsx"))
        cb = XLSX.Charts.getCharts(fb)[1]

        # A literal title, reached two ways: the accessor and parse_chart_title.
        @test XLSX.Charts.text_content(XLSX.Charts.getChartTitleText(cb)) == "Revenue by Region"
        @test XLSX.Charts.getChartTitle(cb) == "Revenue by Region"
        @test length(XLSX.Charts.getChartGroups(cb)) == 1
        @test isempty(XLSX.Charts.getSeriesDataPoints(cb, 1))

        # Six theme variants, all reached through the accessor.
        ft = XLSX.readxlsx(joinpath(data_directory, "chart_theme_colors.xlsx"))
        ct = XLSX.Charts.getCharts(ft)[1]
        @test length(XLSX.Charts.getChartSeries(ct)) == 6
        @test XLSX.Charts.getSeriesShapeProps(ct, 1).fill.fgcolor.rgb == "156082"
        @test XLSX.Charts.getSeriesShapeProps(ct, 5).fill.fgcolor.rgb == "104862"

        # A cx: chart is a ChartEx, and the c: accessors are not defined for it.
        fm = XLSX.readxlsx(joinpath(data_directory, "chart_mixed.xlsx"))
        cms = XLSX.Charts.getCharts(fm)
        @test any(x -> x isa XLSX.Charts.ChartEx, cms)
        cx = first(x for x in cms if x isa XLSX.Charts.ChartEx)
        @test_throws MethodError XLSX.Charts.getSeriesShapeProps(cx, 1)
    end

    @testset "fill cascade" begin
        # Series 3: spPr with a:ln only, no fill element — must not stop the walk.
        e = XLSX.Charts.getSeriesFill(c, 3)
        @test isnothing(e.value) && isnothing(e.site)
        @test length(e.chain) == 1 && !isnothing(e.chain[1].props)   # rung exists, fill doesn't

        # Series 1: solid accent1 on the series.
        e = XLSX.Charts.getSeriesFill(c, 1)
        @test e.site.level === :series && e.value.kind === :solid

        # Series 1 point idx 1 (position 2): plain red, overrides the series.
        e = XLSX.Charts.getSeriesFill(c, 1, point=2)
        @test e.site.level === :point && e.value.fgcolor.rgb == "FF0000"

        # Series 3 point idx 2 (position 3): point spPr is the line segment, no
        # fill — the fill Excel shows is on the marker.
        e = XLSX.Charts.getSeriesFill(c, 3, point=3)
        @test isnothing(e.value)
        @test length(e.chain) == 2 && all(s -> s.kind === :shape, e.chain)

        m = XLSX.Charts.getMarkerFill(c, 3, 3)
        @test m.site.level === :point && m.site.kind === :marker && m.value.fgcolor.rgb == "FF0000"

        # Series 3's own marker (diamond, size 9) has its own spPr, so it is the
        # second rung and answers when the point has none.
        @test length(m.chain) == 2
    end

    @testset "line cascade" begin
        # Series 1: a:ln present but noFill — hasLine false. The line element IS
        # written, so the walk stops here; the resolved DrawingLine has a :none fill.
        e = XLSX.Charts.getSeriesLine(c, 1)
        @test e.site.level === :series
        @test !isnothing(e.value) && e.value.fill.kind === :none
        @test !XLSX.Charts.hasLine(XLSX.Charts.getSeriesShapeProps(c, 1))
    end
    @testset "empty txPr does not answer" begin
        e = XLSX.Charts.getLabelTextProp(c, 3, :size)          # series 3 has no dLbls at all
        @test isnothing(e.value) && isnothing(e.site)
        @test length(e.chain) == 4                    # series, group, plotarea, chartspace
        @test e.chain[1].level === :series && isnothing(e.chain[1].props)
        @test e.chain[end].level === :chartspace && !isnothing(e.chain[end].props)  # present, empty
    end

    @testset "text cascade resolves at the series" begin
        e = XLSX.Charts.getLabelTextProp(c, 1, :size)
        @test e.site.level === :series && e.value == 10.5
        e = XLSX.Charts.getLabelTextProp(c, 2, :size)
        @test e.site.level === :series && e.value == 9.0
    end

    @testset "setSeriesFill" begin
        # Setters mutate xf.data, so work on a copy rather than the tracked fixture.
        src = joinpath(data_directory, "chart_appearance.xlsx")
        tmp = "appearance_setseriesfill.xlsx"
        cp(src, tmp; force=true)

        xf = XLSX.openxlsx(tmp; mode="rw")
        local c = first(XLSX.Charts.getCharts(xf[1]))
        wb = XLSX.get_workbook(xf[1])

        # Series 3 has an spPr with a line and no fill — the empty-slot case.
        @test isnothing(XLSX.Charts.getSeriesFill(c, 3).value)

        c = XLSX.Charts.setSeriesFill(c, 3, "red")
        e = XLSX.Charts.getSeriesFill(c, 3)
        @test e.site.level === :series
        @test e.value.kind === :solid
        @test e.value.fgcolor.rgb == "FF0000"

        # the line is untouched
        @test !isnothing(XLSX.Charts.getSeriesLine(c, 3).value)

        # replacing an existing fill does not throw and does not duplicate
        c = XLSX.Charts.setSeriesFill(c, 3, "blue")
        sp = XLSX.Charts.getSeriesShapeProps(c, 3).raw
        @test count(k -> XLSX.localname(k) in ("solidFill", "noFill", "gradFill", "pattFill",
            "blipFill", "grpFill"), XML.children(sp)) == 1
        @test XLSX.Charts.getSeriesFill(c, 3).value.fgcolor.rgb == "0000FF"

        # a theme color with transforms
        c = XLSX.Charts.setSeriesFill(c, 3, XLSX.Charts.SchemeColor(:accent1; lumMod=75))
        e = XLSX.Charts.getSeriesFill(c, 3)
        @test e.value.fgcolor.val == "accent1"
        @test e.value.fgcolor.transforms == [:lumMod => 75000]

        # :none is explicit, and resolves to a fill rather than to nothing
        c = XLSX.Charts.setSeriesFill(c, 3, :none)
        e = XLSX.Charts.getSeriesFill(c, 3)
        @test e.value.kind === :none && !isnothing(e.site)

        # :inherit removes it, so the cascade finds nothing
        c = XLSX.Charts.setSeriesFill(c, 3, :inherit)
        e = XLSX.Charts.getSeriesFill(c, 3)
        @test isnothing(e.value) && isnothing(e.site)
        @test length(e.chain) == 1                 # rung still reported for a setter

        # a Symbol that is a Colors.jl name is a color, not an instruction
        c = XLSX.Charts.setSeriesFill(c, 3, :red)
        @test XLSX.Charts.getSeriesFill(c, 3).value.fgcolor.rgb == "FF0000"

        # a theme slot named as a string or Symbol is a scheme color, as in addSeries
        for name in ("accent4", :accent4)
            c = XLSX.Charts.setSeriesFill(c, 3, name)
            fg = XLSX.Charts.getSeriesFill(c, 3).value.fgcolor
            @test fg.kind === :scheme && fg.val == "accent4"
        end
        c = XLSX.Charts.setSeriesLineColor(c, 3, :tx1)
        fg = XLSX.Charts.getSeriesLine(c, 3).value.fill.fgcolor
        @test fg.kind === :scheme && fg.val == "tx1"
        c = XLSX.Charts.setSeriesLineColor(c, 3, "dk2")      # an alias, kept as written
        @test XLSX.Charts.getSeriesLine(c, 3).value.fill.fgcolor.val == "dk2"

        # series 1 is in the bar group — the other branch of SER_TYPE
        c = XLSX.Charts.setSeriesFill(c, 1, "green")
        @test XLSX.Charts.getSeriesFill(c, 1).value.fgcolor.rgb == "008000"

        # the survivor test: writing and reopening keeps the change
        XLSX.writexlsx(tmp, xf; overwrite=true)
        xf2 = XLSX.openxlsx(tmp)
        c2 = first(XLSX.Charts.getCharts(xf2[1]))
        @test XLSX.Charts.getSeriesFill(c2, 1).value.fgcolor.rgb == "008000"
        SAVE_FILES && save_outfile(xf)
        isfile(tmp) && rm(tmp)
    end
    @testset "setSeriesLine" begin
        # Setters mutate xf.data, so work on a copy rather than the tracked fixture.
        tmp = "appearance_setseriesline.xlsx"
        cp(joinpath(data_directory, "chart_appearance.xlsx"), tmp; force=true)

        xf = XLSX.openxlsx(tmp; mode="rw")
        local c = first(XLSX.Charts.getCharts(xf[1]))

        # All three series have an a:ln written with noFill inside — the line exists
        # and draws nothing. Remove it to exercise creation from scratch.
        @test !isnothing(XLSX.Charts.getSeriesLine(c, 2).value)
        @test !XLSX.Charts.hasLine(XLSX.Charts.getSeriesShapeProps(c, 2))

        c = XLSX.Charts.setSeriesLine(c, 2, :inherit)
        @test isnothing(XLSX.Charts.getSeriesLine(c, 2).value)

        c = XLSX.Charts.setSeriesLineColor(c, 2, "red")
        e = XLSX.Charts.getSeriesLine(c, 2)
        @test e.site.level === :series
        @test e.value.fill.fgcolor.rgb == "FF0000"
        @test XLSX.Charts.hasLine(XLSX.Charts.getSeriesShapeProps(c, 2))
        # the fill is untouched
        @test !isnothing(XLSX.Charts.getSeriesFill(c, 2).value)

        # width in points, as Excel's box takes it
        c = XLSX.Charts.setSeriesLineWidth(c, 2, 2.25)
        @test XLSX.Charts.getSeriesLine(c, 2).value.width ≈ 2.25
        @test_throws XLSX.XLSXError XLSX.Charts.setSeriesLineWidth(c, 2, 2000)

        # dash by either vocabulary
        c = XLSX.Charts.setSeriesLineDash(c, 2, :roundDot)
        @test XLSX.Charts.getSeriesLine(c, 2).value.dash == "sysDot"
        c = XLSX.Charts.setSeriesLineDash(c, 2, :sysDash)
        @test XLSX.Charts.getSeriesLine(c, 2).value.dash == "sysDash"

        c = XLSX.Charts.setSeriesLineCap(c, 2, :round)
        @test XLSX.Charts.getSeriesLine(c, 2).value.cap == "rnd"

        c = XLSX.Charts.setSeriesLineCompound(c, 2, :double)
        @test XLSX.Charts.getSeriesLine(c, 2).value.compound == "dbl"

        c = XLSX.Charts.setSeriesLineJoin(c, 2, :miter)
        @test XLSX.Charts.getSeriesLine(c, 2).value.join == "miter"
        c = XLSX.Charts.setSeriesLineMiterLimit(c, 2, 8)
        @test XLSX.Charts.getSeriesLine(c, 2).value.miter_limit ≈ 8.0

        # everything set at once is still in schema order
        ln = XLSX.first_element_with_tag(XLSX.Charts.getSeriesShapeProps(c, 2).raw, "ln")
        @test issorted([findfirst(==(XLSX.localname(k)), XLSX.Charts.CHILD_ORDER[(XLSX.NS_A, "ln")])
                        for k in XML.eachelement(ln)])

        # :inherit removes one property and leaves the rest
        c = XLSX.Charts.setSeriesLineDash(c, 2, :inherit)
        e = XLSX.Charts.getSeriesLine(c, 2)
        @test isnothing(e.value.dash) && e.value.cap == "rnd" && e.value.width ≈ 2.25

        # the sugar is one rebuild, and a bad value leaves nothing applied
        c = XLSX.Charts.setSeriesLine(c, 3; color="blue", width=1.5, dash=:dash)
        e = XLSX.Charts.getSeriesLine(c, 3)
        @test e.value.fill.fgcolor.rgb == "0000FF" && e.value.width ≈ 1.5 && e.value.dash == "dash"

        before = XML.write(XLSX.Charts.getSeriesShapeProps(c, 3).raw)
        @test_throws XLSX.XLSXError XLSX.Charts.setSeriesLine(c, 3; width=3, dash=:nonsense)
        @test XML.write(XLSX.Charts.getSeriesShapeProps(c, 3).raw) == before      # atomic

        # no keywords is a no-op returning the same chart
        @test XLSX.Charts.setSeriesLine(c, 3) === c

        # whole-line :none and :inherit
        c = XLSX.Charts.setSeriesLine(c, 3, :none)
        @test !XLSX.Charts.hasLine(XLSX.Charts.getSeriesLine(c, 3))
        @test !XLSX.Charts.hasLine(XLSX.Charts.getSeriesLine(c, 3).value)
        @test !XLSX.Charts.hasLine(XLSX.Charts.getSeriesShapeProps(c, 3))          # element present, fill off
        @test !isnothing(XLSX.first_element_with_tag(XLSX.Charts.getSeriesShapeProps(c, 3).raw, "ln"))
        c = XLSX.Charts.setSeriesLine(c, 3, :inherit)
        @test isnothing(XLSX.first_element_with_tag(XLSX.Charts.getSeriesShapeProps(c, 3).raw, "ln"))

        # survives a write and reopen
        c = XLSX.Charts.setSeriesLineColor(c, 1, "green")
        XLSX.writexlsx(tmp, xf; overwrite=true)
        xf2 = XLSX.openxlsx(tmp)
        c2 = first(XLSX.Charts.getCharts(xf2[1]))
        @test XLSX.Charts.getSeriesLine(c2, 1).value.fill.fgcolor.rgb == "008000"
        SAVE_FILES && save_outfile(xf)

        isfile(tmp) && rm(tmp)

        @testset "an unknown symbol is rejected" begin
            local xf = XLSX.openxlsx(joinpath(data_directory, "chart_basic.xlsx"); mode="rw")
            local c = XLSX.getChart(xf["Data"], "chart1")
            @test_throws XLSX.XLSXError XLSX.Charts.setSeriesLine(c, 1, :bogus)       # not :none/:inherit
            SAVE_FILES && save_outfile(xf)
        end
    end

    @testset "setMarker" begin
        tmp = "appearance_setmarker.xlsx"
        cp(joinpath(data_directory, "chart_appearance.xlsx"), tmp; force=true)

        xf = XLSX.openxlsx(tmp; mode="rw")
        local c = first(XLSX.Charts.getCharts(xf[1]))

        # Series 3 is the line series: c:marker diamond size 9 with its own spPr.
        m = XLSX.Charts.getSeriesMarker(c, 3)
        @test m.symbol === :diamond && m.size == 9

        c = XLSX.Charts.setMarkerSymbol(c, 3, :circle)
        @test XLSX.Charts.getSeriesMarker(c, 3).symbol === :circle
        c = XLSX.Charts.setMarkerSize(c, 3, 12)
        @test XLSX.Charts.getSeriesMarker(c, 3).size == 12

        @test_throws XLSX.XLSXError XLSX.Charts.setMarkerSize(c, 3, 100)     # ST_MarkerSize is 2..72
        @test_throws XLSX.XLSXError XLSX.Charts.setMarkerSize(c, 3, 1)
        @test_throws XLSX.XLSXError XLSX.Charts.setMarkerSymbol(c, 3, :hexagon)

        # :none is a symbol; :inherit removes the element
        c = XLSX.Charts.setMarkerSymbol(c, 3, :none)
        @test XLSX.Charts.getSeriesMarker(c, 3).symbol === :none
        c = XLSX.Charts.setMarkerSymbol(c, 3, :inherit)
        @test isnothing(XLSX.Charts.getSeriesMarker(c, 3).symbol)
        @test XLSX.Charts.getSeriesMarker(c, 3).size == 12              # size survives

        # marker fill and line, on the series marker
        c = XLSX.Charts.setMarkerFill(c, 3, "red")
        @test XLSX.Charts.getMarkerFill(c, 3, 1).value.fgcolor.rgb == "FF0000"
        c = XLSX.Charts.setMarkerLineColor(c, 3, "blue")
        c = XLSX.Charts.setMarkerLineWidth(c, 3, 1.5)
        ln = XLSX.first_element_with_tag(XLSX.Charts.getSeriesMarker(c, 3).shape.raw, "ln")
        @test !isnothing(ln)

        # schema order inside c:marker
        mk = XLSX.Charts.getSeriesMarker(c, 3).raw
        @test issorted([findfirst(==(XLSX.localname(k)), XLSX.Charts.CHILD_ORDER[(XLSX.Charts.NS_C, "marker")])
                        for k in XML.eachelement(mk)])

        # the point marker: series 3's dPt at idx 2 is point 3
        c = XLSX.Charts.setMarkerFill(c, 3, "green"; point=3)
        e = XLSX.Charts.getMarkerFill(c, 3, 3)
        @test e.site.level === :point && e.value.fgcolor.rgb == "008000"

        # Point 1 of series 3 had no c:dPt; the setter now creates one.
        XLSX.Charts.setMarkerFill(c, 3, "red"; point=1)
        d = XLSX.Charts.getSeriesDataPoint(c, 3, 1)
        @test !isnothing(d)
        @test XLSX.Charts.getDataPointMarker(c, d).shape.fill.fgcolor.rgb == "FF0000"

        # the sugar is one rebuild
        c = XLSX.Charts.setMarker(c, 3; symbol=:square, size=7, fill="yellow")
        m = XLSX.Charts.getSeriesMarker(c, 3)
        @test m.symbol === :square && m.size == 7
        @test XLSX.Charts.setMarker(c, 3) === c                          # no keywords, no-op

        # survives a write and reopen
        XLSX.writexlsx(tmp, xf; overwrite=true)
        c2 = first(XLSX.Charts.getCharts(XLSX.openxlsx(tmp)[1]))
        @test XLSX.Charts.getSeriesMarker(c2, 3).symbol === :square
        SAVE_FILES && save_outfile(xf)
        isfile(tmp) && rm(tmp)
    end

    @testset "setLabelTextProp" begin
        tmp = "appearance_setlabeltextprop.xlsx"
        cp(joinpath(data_directory, "chart_appearance.xlsx"), tmp; force=true)

        xf = XLSX.openxlsx(tmp; mode="rw")
        wb = XLSX.get_workbook(xf[1])
        local c = first(XLSX.Charts.getCharts(xf[1]))

        # Series 1's labels are sz 1050 accent1+lumMod; series 2's are sz 900.
        @test XLSX.Charts.getLabelTextProp(c, 1, :size).value ≈ 10.5
        @test XLSX.Charts.getLabelTextProp(c, 2, :size).value ≈ 9.0

        c = XLSX.Charts.setLabelTextProp(c, 1, :size, 14)
        e = XLSX.Charts.getLabelTextProp(c, 1, :size)
        @test e.value ≈ 14.0 && e.site.level === :series
        @test XLSX.Charts.getLabelTextProp(c, 2, :size).value ≈ 9.0      # series 2 untouched

        # other fields are independent
        c = XLSX.Charts.setLabelTextProp(c, 1, :bold, true)
        @test XLSX.Charts.getLabelTextProp(c, 1, :bold).value === true
        @test XLSX.Charts.getLabelTextProp(c, 1, :size).value ≈ 14.0

        # :inherit removes one field and leaves the rest
        c = XLSX.Charts.setLabelTextProp(c, 1, :size, :inherit)
        @test isnothing(XLSX.Charts.getLabelTextProp(c, 1, :size).value)
        @test XLSX.Charts.getLabelTextProp(c, 1, :bold).value === true

        # compound fields
        c = XLSX.Charts.setLabelTextProp(c, 1, :fill, "red")
        @test XLSX.Charts.getLabelTextProp(c, 1, :fill).value.fgcolor.rgb == "FF0000"
        c = XLSX.Charts.setLabelTextProp(c, 1, :fill, XLSX.Charts.SchemeColor(:accent1; lumMod=75))
        @test XLSX.Charts.getLabelTextProp(c, 1, :fill).value.fgcolor.val == "accent1"
        c = XLSX.Charts.setLabelTextProp(c, 1, :line, (color="blue", width=1.5))
        l = XLSX.Charts.getLabelTextProp(c, 1, :line).value
        @test l.fill.fgcolor.rgb == "0000FF" && l.width ≈ 1.5

        @test_throws XLSX.XLSXError XLSX.Charts.setLabelTextProp(c, 1, :nonsense, 1)

        # a scheme name works for a text fill too
        c = XLSX.Charts.setLabelTextProp(c, 1, :color, "accent2")
        @test XLSX.Charts.getLabelTextProp(c, 1, :fill).value.fgcolor.val == "accent2"

        # underline, strike and caps take their DrawingML vocabularies — as Symbol
        # or String — plus Excel's names and, for the first two, a Bool. Anything
        # else would make a part Excel refuses, so it throws.
        for (field, v, want) in ((:under, :dbl, "dbl"), (:underline, "wavy", "wavy"),
                                 (:under, :single, "sng"), (:under, true, "sng"),
                                 (:under, false, "none"),
                                 (:strike, "dblStrike", "dblStrike"), (:strikethrough, true, "sngStrike"),
                                 (:strike, :none, "noStrike"),
                                 (:caps, :small, "small"), (:caps, "all", "all"))
            c = XLSX.Charts.setLabelTextProp(c, 1, field, v)
            f = get(XLSX.Charts._RUN_PROP_ALIASES, field, field)
            @test XLSX.Charts.getLabelTextProp(c, 1, f).value == want
        end
        @test_throws XLSX.XLSXError XLSX.Charts.setLabelTextProp(c, 1, :under, "single-ish")
        @test_throws XLSX.XLSXError XLSX.Charts.setLabelTextProp(c, 1, :strike, :sng)
        @test_throws XLSX.XLSXError XLSX.Charts.setLabelTextProp(c, 1, :caps, true)
        @test_throws XLSX.XLSXError XLSX.Charts.setLabelTextProp(c, 1, :under, 1)
        c = XLSX.Charts.setLabelTextProp(c, 1, :under, :inherit)
        @test isnothing(XLSX.Charts.getLabelTextProp(c, 1, :under).value)

        # Series 3 has no c:dLbls at all — creating one writes formatting only.
        @test isnothing(XLSX.Charts.getLabelTextProp(c, 3, :size).value)
        c = XLSX.Charts.setLabelTextProp(c, 3, :size, 11)
        @test XLSX.Charts.getLabelTextProp(c, 3, :size).value ≈ 11.0

        # Series 3 has no c:dLbls at all. Creating one to hold formatting must not
        # turn labels on: a c:dLbls naming no show* flags shows labels with Excel's
        # defaults, so all six are written explicitly off, after c:txPr in schema order.
        dl = XLSX.first_element_with_tag(XLSX.Charts._series(c, 3).raw, "dLbls")
        @test !isnothing(dl)
        flags = ["showLegendKey", "showVal", "showCatName", "showSerName", "showPercent", "showBubbleSize"]
        @test XLSX.localname.(collect(XML.eachelement(dl))) == ["txPr"; flags]
        for f in flags
            @test XML.attributes(XLSX.first_element_with_tag(dl, f))["val"] == "0"
        end

        # An individual label. Series 1's dLbl at idx 0 is point 1, retyped, and
        # carries formatting in BOTH c:tx/c:rich and c:txPr — writing one alone
        # would leave the edit invisible in Excel.
        c = XLSX.Charts.setLabelTextProp(c, 1, 1, :size, 20)
        lbl = XLSX.Charts.getSeriesDataLabel(c, 1, 1).raw
        for body in (XLSX.first_element_with_tag(lbl, "txPr"),
            XLSX.first_element_with_tag(XLSX.first_element_with_tag(lbl, "tx"), "rich"))
            rp = XLSX.Charts.default_run_props(XLSX.Charts.parse_drawing_text(wb, body))
            @test rp.size ≈ 20.0
        end

        # Point 1 of series 2 had no individual c:dLbl; the setter now creates one.
        XLSX.Charts.setLabelTextProp(c, 2, 1, :size, 12)
        dl = XLSX.Charts.getSeriesDataLabel(c, 2, 1)
        @test !isnothing(dl)
        @test XLSX.Charts.default_run_props(XLSX.Charts.getDataLabelTextProps(c, dl)).size ≈ 12.0

        # the group rung. Series 1's txPr sets i="0" explicitly — Excel writes the
        # full attribute set — so the group cannot be reached until that is removed.
        c = XLSX.Charts.setLabelTextProp(c, 1, :italic, :inherit)
        g = XLSX.Charts.getChartGroups(c)[1]
        c = XLSX.Charts.setGroupLabelTextProp(c, g, :italic, true)
        e = XLSX.Charts.getLabelTextProp(c, 1, :italic)
        @test e.value === true && e.site.level === :group

        # the chart space rung
        c = XLSX.Charts.setChartSpaceTextProp(c, :caps, "all")
        e = XLSX.Charts.getLabelTextProp(c, 3, :caps)
        @test e.value == "all" && e.site.level === :chartspace

        # survives a write and reopen
        XLSX.writexlsx(tmp, xf; overwrite=true)
        c2 = first(XLSX.Charts.getCharts(XLSX.openxlsx(tmp)[1]))
        @test XLSX.Charts.getLabelTextProp(c2, 3, :size).value ≈ 11.0
        SAVE_FILES && save_outfile(xf)
        isfile(tmp) && rm(tmp)
    end

    @testset "run properties written in hundredths" begin
        local xf = XLSX.openxlsx(joinpath(data_directory, "chart_basic.xlsx"); mode="rw")
        local c = XLSX.getChart(xf["Data"], "chart1")
        for (field, v) in ((:kern, 12.0), (:spacing, 1.5), (:baseline, 0.3))
            XLSX.Charts.setLabelTextProp(c, 1, field, v)
            @test XLSX.Charts.getLabelTextProp(c, 1, field).value ≈ v
        end
        SAVE_FILES && save_outfile(xf)
    end
    @testset "title and legend text" begin
        tmp = "appearance_title_and_legend_text.xlsx"
        cp(joinpath(data_directory, "chart_appearance.xlsx"), tmp; force=true)
        xf = XLSX.openxlsx(tmp; mode="rw")
        local c = first(XLSX.Charts.getCharts(xf[1]))
        wb = XLSX.get_workbook(xf[1])

        # The fixture's chart title has spPr and txPr but no c:tx — Excel generates
        # the text. Setting a property touches txPr only.
        @test isnothing(XLSX.Charts.getChartTitleText(c))
        c = XLSX.Charts.setChartTitleTextProp(c, :bold, true)
        @test XLSX.Charts.default_run_props(XLSX.Charts.getChartTitleTextProps(c)).bold === true
        @test isnothing(XLSX.Charts.getChartTitleText(c))        # still no literal text

        # Giving it literal text, then changing a property, must update both bodies.
        c = XLSX.Charts.setChartTitleText(c, "Revenue")
        @test XLSX.Charts.text_content(XLSX.Charts.getChartTitleText(c)) == "Revenue"
        c = XLSX.Charts.setChartTitleTextProp(c, :size, 18)
        @test XLSX.Charts.default_run_props(XLSX.Charts.getChartTitleTextProps(c)).size ≈ 18.0
        @test XLSX.Charts.default_run_props(XLSX.Charts.getChartTitleText(c)).size ≈ 18.0

        # A DrawingText replaces text and formatting wholesale.
        c = XLSX.Charts.setChartTitleText(c, XLSX.Charts.DrawingText(XLSX.Charts.DrawingParagraph(
            XLSX.Charts.DrawingRun("Q4", props=XLSX.Charts.DrawingRunProps(size=24.0, italic=true)))))
        g = XLSX.Charts.getChartTitleText(c)
        @test XLSX.Charts.text_content(g) == "Q4"
        @test XLSX.Charts.first_run_props(g).size ≈ 24.0 && XLSX.Charts.first_run_props(g).italic === true

        # Axis titles. The catAx at 612078287 is titled "Horizontal".
        ax = XLSX.Charts.getChartAxis(c, 612078287)
        @test XLSX.Charts.text_content(XLSX.Charts.getAxisTitleText(c, ax)) == "Horizontal"
        c = XLSX.Charts.setAxisTitleText(c, ax, "Quarter")
        ax = XLSX.Charts.getChartAxis(c, 612078287)              # the Chart is fresh; re-fetch
        @test XLSX.Charts.text_content(XLSX.Charts.getAxisTitleText(c, ax)) == "Quarter"

        c = XLSX.Charts.setAxisTitleTextProp(c, ax, :bold, true)
        ax = XLSX.Charts.getChartAxis(c, 612078287)
        @test XLSX.Charts.default_run_props(XLSX.Charts.getAxisTitleText(c, ax)).bold === true
        # The title's c:txPr carries it too, and is read by getAxisTitleTextProps.
        @test XLSX.Charts.default_run_props(XLSX.Charts.getAxisTitleTextProps(c, ax)).bold === true
        # The deleted catAx has no title at all.
        @test isnothing(XLSX.Charts.getAxisTitleTextProps(c, XLSX.Charts.getChartAxis(c, 1926562432)))

        # The legend has txPr and no c:tx, so only formatting is settable.
        c = XLSX.Charts.setLegendTextProp(c, :size, 11)
        @test XLSX.Charts.default_run_props(XLSX.Charts.getLegendTextProps(c)).size ≈ 11.0

        # Schema order survives creating a title from nothing.
        t = XLSX.first_element_with_tag(XLSX.first_element_with_tag(XLSX.Charts.chart_root(c), "chart"), "title")
        @test issorted([XLSX.Charts._slot(XLSX.Charts.CHILD_ORDER[(XLSX.Charts.NS_C, "title")], XLSX.localname(k))
                        for k in XML.eachelement(t)])

        # survives a write and reopen
        XLSX.writexlsx(tmp, xf; overwrite=true)
        c2 = first(XLSX.Charts.getCharts(XLSX.openxlsx(tmp)[1]))
        @test XLSX.Charts.text_content(XLSX.Charts.getChartTitleText(c2)) == "Q4"
        ax2 = XLSX.Charts.getChartAxis(c2, 612078287)
        @test XLSX.Charts.text_content(XLSX.Charts.getAxisTitleText(c2, ax2)) == "Quarter"
        SAVE_FILES && save_outfile(xf)
        isfile(tmp) && rm(tmp)
    end

    @testset "created spPr takes the chart prefix" begin
        tmp = "prefix.xlsx"
        cp(joinpath(data_directory, "chart_appearance.xlsx"), tmp; force=true)
        xf = XLSX.openxlsx(tmp; mode="rw")
        local c = only(filter(x -> x isa XLSX.Charts.Chart, XLSX.Charts.getCharts(xf)))

        # strip series 1's spPr so the setter has to create one
        root = XLSX.Charts.chart_root(c)
        new = XLSX.Charts.rebuild_path(root, XLSX.Charts._series_path(c, root, 1)[1],
            s -> XLSX.Charts.remove_child(s, "spPr");
            prefixes=XLSX.ns_prefixes(root))
        XLSX.Charts.set_chart_root!(c, new)
        @test XLSX.Charts.setSeriesFill(c, 1, "FF00B0F0") === c
        @test XLSX.XML.tag(XLSX.first_element_with_tag(XLSX.Charts._series(c, 1).raw, "spPr")) == "c:spPr"
        SAVE_FILES && save_outfile(xf)
        isfile(tmp) && rm(tmp)
    end

    @testset "title run properties yield to the paragraph default" begin
        tmp = "c_runclear.xlsx"
        cp(joinpath(data_directory, "chart_basic.xlsx"), tmp; force=true)
        xf = XLSX.openxlsx(tmp; mode="rw")
        local c = only(filter(x -> x isa XLSX.Charts.Chart, XLSX.Charts.getCharts(xf)))
        c = XLSX.Charts.setChartTitleTextProp(c, :size, 18)
        @test XLSX.Charts.default_run_props(XLSX.Charts.getChartTitleTextProps(c)).size ≈ 18.0
        rich = XLSX.first_element_with_tag(
            XLSX.first_element_with_tag(XLSX.Charts.getChartTitleNode(c), "tx"), "rich")
        for r in XLSX.elements_with_tag(XLSX.first_element_with_tag(rich, "p"), "r")
            rpr = XLSX.first_element_with_tag(r, "rPr")
            isnothing(rpr) || @test isempty(XLSX.get_attr(rpr, "sz"))
        end
        SAVE_FILES && save_outfile(xf)
        isfile(tmp) && rm(tmp)
    end

    @testset "Chart handles are durable" begin
        path = "chart_durable.xlsx"
        out = "chart_durable_out.xlsx"
        cp(joinpath(data_directory, "chart_basic.xlsx"), path; force=true)
        xf = XLSX.openxlsx(path; mode="rw")
        local c = XLSX.Charts.getChart(xf, "chart1")
        c2 = XLSX.Charts.getChart(xf, "chart1")          # obtained before the write

        n = length(XLSX.Charts.getChartSeries(c))
        @test XLSX.Charts.setSeriesFill(c, 1, "FFFF0000") === c

        # Both handles see the write; neither needed refreshing.
        @test XLSX.Charts.getSeriesFill(c, 1).value.fgcolor.rgb == "FF0000"
        @test XLSX.Charts.getSeriesFill(c2, 1).value.fgcolor.rgb == "FF0000"
        @test length(XLSX.Charts.getChartSeries(c)) == n
        @test XLSX.Charts.getChartTypes(c) == [:barChart]
        @test occursin("barChart", sprint(show, c))

        # A second write through the same handle, then a round trip.
        XLSX.Charts.setSeriesLineWidth(c, 2, 2.5)
        SAVE_FILES && save_outfile(xf)
        XLSX.writexlsx(out, xf, overwrite=true)
        c3 = XLSX.Charts.getChart(XLSX.readxlsx(out), "chart1")
        @test XLSX.Charts.getSeriesFill(c3, 1).value.fgcolor.rgb == "FF0000"
        @test XLSX.Charts.getSeriesLine(c3, 2).value.width ≈ 2.5

        isfile(path) && rm(path)
        isfile(out) && rm(out)
    end

    @testset "data point setters address the same point as the getters" begin
        path = "chart_appearance_dpt.xlsx"
        cp(joinpath(data_directory, "chart_appearance.xlsx"), path; force=true)
        xf = XLSX.openxlsx(path; mode="rw")

        hit = nothing
        for c in XLSX.Charts.getCharts(xf)
            c isa XLSX.Charts.Chart || continue
            for i in 1:length(XLSX.Charts.getChartSeries(c))
                dps = XLSX.Charts.getSeriesDataPoints(c, i)
                isempty(dps) || (hit=(c, i, first(dps).idx + 1); break)
            end
            isnothing(hit) || break
        end
        @test !isnothing(hit)                      # the fixture must have a c:dPt
        c, i, point = hit

        XLSX.Charts.setMarkerSymbol(c, i, :diamond; point=point)
        SAVE_FILES && save_outfile(xf)
        d = XLSX.Charts.getSeriesDataPoint(c, i, point)
        @test XLSX.Charts.getDataPointMarker(c, d).symbol === :diamond

        isfile(path) && rm(path)
    end

    @testset "axis handles survive writes" begin
        path = "chart_axis_durable.xlsx"
        cp(joinpath(data_directory, "chart_basic.xlsx"), path; force=true)
        xf = XLSX.openxlsx(path; mode="rw")
        local c = XLSX.Charts.getChart(xf, "chart1")
        ax = only(XLSX.Charts.getChartAxes(c, :value))

        @test isnothing(XLSX.Charts.getAxisTitleText(c, ax))
        @test XLSX.Charts.setAxisTitleText(c, ax, "Revenue") === c
        @test XLSX.Charts.text_content(XLSX.Charts.getAxisTitleText(c, ax)) == "Revenue"
        XLSX.Charts.setAxisTitleTextProp(c, ax, :bold, true)
        @test XLSX.Charts.default_run_props(XLSX.Charts.getAxisTitleText(c, ax)).bold === true
        @test XLSX.Charts.getAxisCrossBetween(c, ax) === :between           # scalars still read
        SAVE_FILES && save_outfile(xf)

        isfile(path) && rm(path)
    end

    @testset "creating c:dPt and c:dLbl" begin
        path = "chart_create_points.xlsx"
        cp(joinpath(data_directory, "chart_gaps.xlsx"), path; force=true)
        xf = XLSX.openxlsx(path; mode="rw")
        local c = XLSX.Charts.getCharts(xf)[1]
        @test isempty(XLSX.Charts.getSeriesDataPoints(c, 1))

        # Created out of order; stored in c:idx order. Points 1 and 4 both plot.
        XLSX.Charts.setMarkerSymbol(c, 1, :diamond; point=4)
        XLSX.Charts.setMarkerSymbol(c, 1, :square; point=1)
        @test [d.idx for d in XLSX.Charts.getSeriesDataPoints(c, 1)] == [0, 3]
        @test XLSX.Charts.getDataPointMarker(c, XLSX.Charts.getSeriesDataPoint(c, 1, 4)).symbol === :diamond
        @test XLSX.Charts.getDataPointMarker(c, XLSX.Charts.getSeriesDataPoint(c, 1, 1)).symbol === :square

        # A second write to an existing point edits it rather than adding another.
        XLSX.Charts.setMarkerSize(c, 1, 9; point=4)
        @test length(XLSX.Charts.getSeriesDataPoints(c, 1)) == 2
        @test XLSX.Charts.getDataPointMarker(c, XLSX.Charts.getSeriesDataPoint(c, 1, 4)).size == 9

        # A created label is formatted but not switched on. Point 1 has a value,
        # so if the flags were wrong the label would be visible.
        XLSX.Charts.setLabelTextProp(c, 1, 1, :bold, true)
        dl = XLSX.Charts.getSeriesDataLabel(c, 1, 1)
        @test !isnothing(dl)
        @test XLSX.Charts.default_run_props(XLSX.Charts.getDataLabelTextProps(c, dl)).bold === true
        n = XLSX.Charts._node(c, dl)
        @test all(f -> XLSX.Charts._bool_val(n, f) === false, XLSX.Charts.DLBLS_FLAGS)

        # And one with typed text.
        XLSX.Charts.setLabelText(c, 1, 5, "peak")
        @test XLSX.Charts.text_content(XLSX.Charts.getDataLabelText(c, XLSX.Charts.getSeriesDataLabel(c, 1, 5))) == "peak"

        # Deleting the typed label at e removes its text; undeleting removes the override.
        XLSX.Charts.setLabelDeleted(c, 1, 5, true)
        dl = XLSX.Charts.getSeriesDataLabel(c, 1, 5)
        @test dl.delete === true
        @test isnothing(XLSX.Charts.getDataLabelText(c, dl))
        @test_throws XLSX.XLSXError XLSX.Charts.setLabelTextProp(c, 1, 5, :size, 10)

        XLSX.Charts.setLabelDeleted(c, 1, 5, false)
        @test isnothing(XLSX.Charts.getSeriesDataLabel(c, 1, 5))
        XLSX.Charts.setLabelDeleted(c, 1, 5, false)                  # undeleting again is a no-op
        @test isnothing(XLSX.Charts.getSeriesDataLabel(c, 1, 5))

        # Delete it once more so the saved file shows the deleted form.
        XLSX.Charts.setLabelDeleted(c, 1, 5, true)

        SAVE_FILES && save_outfile(xf)
        isfile(path) && rm(path)
    end
    @testset "ChartMarker" begin
        plain(x) = sprint(show, MIME("text/plain"), x)

        xf = XLSX.opentemplate(joinpath(data_directory, "chart_kinds.xlsx"))
        local c = XLSX.Charts.getCharts(xf["linemarkers"])[1]

        XLSX.Charts.setMarker(c, 1; symbol=:circle, size=7)
        m = XLSX.Charts.getSeriesMarker(c, 1)
        @test repr(m) == "XLSX.Charts.ChartMarker(circle, size 7)"
        s = plain(m)
        @test occursin("series idx $(m.series_idx)", s)
        @test occursin("symbol: circle", s)
        @test occursin("size: 7 pt", s)
        @test !occursin("Element", s)
        @test !occursin("point idx", s)         # the series' own marker

        XLSX.Charts.setMarkerSymbol(c, 1, :diamond; point=2)
        p = XLSX.Charts.getDataPointMarker(c, XLSX.Charts.getSeriesDataPoints(c, 1)[1])
        @test occursin("point idx 1", repr(p))  # c:idx, one less than the position
        @test occursin("point idx 1", plain(p))

        SAVE_FILES && save_outfile(xf)
    end

    @testset "axis, group and legend setters" begin
        tmp = "chart_setters_stackedbar.xlsx"
        cp(joinpath(data_directory, "chart_kinds.xlsx"), tmp; force=true)
        try
            xf = XLSX.openxlsx(tmp; mode="rw")
            local c = first(XLSX.Charts.getCharts(xf["stackedbar"]))
            axs = XLSX.Charts.getChartAxes(c)
            catax = only(filter(a -> a.kind === :catAx, axs))
            valax = only(filter(a -> a.kind === :valAx, axs))
            g = only(XLSX.Charts.getChartGroups(c))
            snapshot(c) = XML.write(XLSX.Charts.chart_root(c))

            # tick label position
            @test XLSX.Charts.getAxisTickLabelPos(c, catax) === :nextTo
            c = XLSX.Charts.setAxisTickLabelPos(c, catax, :low)
            @test XLSX.Charts.getAxisTickLabelPos(c, catax) === :low
            before = snapshot(c)
            @test_throws XLSX.XLSXError XLSX.Charts.setAxisTickLabelPos(c, catax, :left)
            @test snapshot(c) == before                                          # atomic
            c = XLSX.Charts.setAxisTickLabelPos(c, catax, :inherit)
            @test isnothing(XLSX.Charts.getAxisTickLabelPos(c, catax))
            c = XLSX.Charts.setAxisTickLabelPos(c, catax, :low)

            # bounds: absent by default, set together, checked after all keywords apply
            @test isnothing(XLSX.Charts.getAxisMin(c, valax))
            @test isnothing(XLSX.Charts.getAxisMax(c, valax))
            c = XLSX.Charts.setAxisScaling(c, valax; min=-0.8, max=0.8)
            @test XLSX.Charts.getAxisMin(c, valax) ≈ -0.8
            @test XLSX.Charts.getAxisMax(c, valax) ≈ 0.8
            before = snapshot(c)
            @test_throws XLSX.XLSXError XLSX.Charts.setAxisMin(c, valax, 1)        # past the max
            @test snapshot(c) == before
            c = XLSX.Charts.setAxisScaling(c, valax; min=1, max=2)            # one call moves both
            @test XLSX.Charts.getAxisMin(c, valax) ≈ 1.0
            @test XLSX.Charts.getAxisMax(c, valax) ≈ 2.0
            @test_throws XLSX.XLSXError XLSX.Charts.setAxisScaling(c, valax; min=3, max=2)
            @test_throws XLSX.XLSXError XLSX.Charts.setAxisMax(c, valax, Inf)
            @test XLSX.Charts.setAxisScaling(c, valax) === c                       # no keywords
            c = XLSX.Charts.setAxisMax(c, valax, :inherit)
            @test isnothing(XLSX.Charts.getAxisMax(c, valax))
            @test XLSX.Charts.getAxisMin(c, valax) ≈ 1.0

            # orientation, and c:scaling stays in schema order
            c = XLSX.Charts.setAxisOrientation(c, catax, :maxMin)
            @test XLSX.Charts.getAxisOrientation(c, catax) === :maxMin
            @test_throws XLSX.XLSXError XLSX.Charts.setAxisOrientation(c, catax, :reverse)
            c = XLSX.Charts.setAxisScaling(c, valax; orientation=:maxMin, min=-1, max=1)
            sc = XLSX.first_element_with_tag(XLSX.Charts._axnode(c, valax), "scaling")
            @test issorted([findfirst(==(XLSX.localname(k)), XLSX.Charts.CHILD_ORDER[(XLSX.Charts.NS_C, "scaling")])
                            for k in XML.eachelement(sc)])

            # crosses and crossesAt replace each other; :inherit removes only its own
            @test XLSX.Charts.getAxisCrosses(c, valax) === :autoZero
            c = XLSX.Charts.setAxisCrossesAt(c, valax, 2.5)
            @test XLSX.Charts.getAxisCrossesAt(c, valax) ≈ 2.5
            @test isnothing(XLSX.Charts.getAxisCrosses(c, valax))
            c = XLSX.Charts.setAxisCrosses(c, valax, :max)
            @test XLSX.Charts.getAxisCrosses(c, valax) === :max
            @test isnothing(XLSX.Charts.getAxisCrossesAt(c, valax))
            @test_throws XLSX.XLSXError XLSX.Charts.setAxisCrosses(c, valax, :middle)
            c = XLSX.Charts.setAxisCrossesAt(c, valax, :inherit)                  # absent: no-op
            @test XLSX.Charts.getAxisCrosses(c, valax) === :max
            c = XLSX.Charts.setAxisCrosses(c, valax, :inherit)
            @test isnothing(XLSX.Charts.getAxisCrosses(c, valax))

            # number format: setting a code unlinks it from the source cells
            @test XLSX.Charts.getAxisNumberFormatLinked(c, valax) === true
            c = XLSX.Charts.setAxisNumberFormatCode(c, valax, "0%;0%")
            @test XLSX.Charts.getAxisNumberFormatCode(c, valax) == "0%;0%"
            @test XLSX.Charts.getAxisNumberFormatLinked(c, valax) === false
            c = XLSX.Charts.setAxisNumberFormatLinked(c, valax, true)
            @test XLSX.Charts.getAxisNumberFormatCode(c, valax) == "0%;0%"
            c = XLSX.Charts.setAxisNumberFormatCode(c, valax, :inherit)
            @test isnothing(XLSX.Charts.getAxisNumberFormatCode(c, valax))
            @test isnothing(XLSX.Charts.getAxisNumberFormatLinked(c, valax))
            c = XLSX.Charts.setAxisNumberFormatLinked(c, valax, false)
            @test XLSX.Charts.getAxisNumberFormatCode(c, valax) == "General"
            escaped = "0\" pts & more\";0"
            c = XLSX.Charts.setAxisNumberFormatCode(c, catax, escaped)            # checked after reopen

            # group gap width and overlap
            @test XLSX.Charts.getGroupGapWidth(c, g) == 150
            @test XLSX.Charts.getGroupOverlap(c, g) == 100
            c = XLSX.Charts.setGroupGapWidth(c, g, 50)
            @test XLSX.Charts.getGroupGapWidth(c, g) == 50
            @test_throws XLSX.XLSXError XLSX.Charts.setGroupGapWidth(c, g, 501)
            @test_throws XLSX.XLSXError XLSX.Charts.setGroupGapWidth(c, g, 1.5)
            @test_throws XLSX.XLSXError XLSX.Charts.setGroupOverlap(c, g, -101)
            c = XLSX.Charts.setGroupOverlap(c, g, :inherit)
            @test isnothing(XLSX.Charts.getGroupOverlap(c, g))
            c = XLSX.Charts.setGroupOverlap(c, g, 100)

            # legend entries: keyed by series, kept sorted, one call for several
            @test isnothing(XLSX.Charts.getLegendEntryDeleted(c, 1))
            c = XLSX.Charts.setLegendEntryDeleted(c, [3, 1], true)
            @test XLSX.Charts.getLegendEntryDeleted(c, 1) === true
            @test XLSX.Charts.getLegendEntryDeleted(c, 3) === true
            @test isnothing(XLSX.Charts.getLegendEntryDeleted(c, 2))
            c = XLSX.Charts.setLegendEntryDeleted(c, 2, false)
            @test XLSX.Charts.getLegendEntryDeleted(c, 2) === false
            lg = XLSX.Charts.getChartLegend(c)
            entries = [k for k in XML.eachelement(lg) if XLSX.localname(k) == "legendEntry"]
            @test [XLSX.Charts._int_val(k, "idx") for k in entries] == [0, 1, 2]
            @test issorted([findfirst(==(XLSX.localname(k)), XLSX.Charts.CHILD_ORDER[(XLSX.Charts.NS_C, "legend")])
                            for k in XML.eachelement(lg)])
            @test_throws XLSX.XLSXError XLSX.Charts.setLegendEntryDeleted(c, 1, :hide)
            @test XLSX.Charts.setLegendEntryDeleted(c, Int[], true) === c
            c = XLSX.Charts.setLegendEntryDeleted(c, 1:3, :inherit)
            @test all(i -> isnothing(XLSX.Charts.getLegendEntryDeleted(c, i)), 1:3)
            c = XLSX.Charts.setLegendEntryDeleted(c, 1, true)

            # legend position and overlay: the five positions, Excel's names as aliases
            for p in (:b, :t, :l, :r, :tr)
                c = XLSX.Charts.setLegendPos(c, p)
                @test XLSX.Charts.getLegendPos(c) === p
            end
            for (alias, p) in (:bottom => :b, :top => :t, :left => :l, :right => :r, :topRight => :tr)
                c = XLSX.Charts.setLegendPos(c, alias)
                @test XLSX.Charts.getLegendPos(c) === p
            end
            before = snapshot(c)
            @test_throws XLSX.XLSXError XLSX.Charts.setLegendPos(c, :bl)
            @test_throws XLSX.XLSXError XLSX.Charts.setLegendPos(c, :t; overlay=:yes)
            @test snapshot(c) == before
            c = XLSX.Charts.setLegendPos(c, :t; overlay=true)
            @test XLSX.Charts.getLegendOverlay(c) === true
            c = XLSX.Charts.setLegendPos(c, :l)                                     # overlay untouched
            @test XLSX.Charts.getLegendOverlay(c) === true
            c = XLSX.Charts.setLegendOverlay(c, false)
            @test XLSX.Charts.getLegendOverlay(c) === false
            c = XLSX.Charts.setLegendOverlay(c, :inherit)
            @test isnothing(XLSX.Charts.getLegendOverlay(c))
            c = XLSX.Charts.setLegendPos(c, :inherit)
            @test isnothing(XLSX.Charts.getLegendPos(c))
            # a dragged legend's manual layout gives way to the position
            lkey = (XLSX.Charts.NS_C, "legend")
            pfx = XLSX.Charts.ns_prefixes(XLSX.Charts.chart_root(c))
            c = XLSX.Charts._set_legend(c, (lg, key, _) -> XLSX.Charts.insert_child(lg, key,
                    XML.Element(XLSX.Charts.prefixed_tag(pfx[XLSX.Charts.NS_C], "layout"))))
            @test !isnothing(XLSX.first_element_with_tag(XLSX.Charts.getChartLegend(c), "layout"))
            c = XLSX.Charts.setLegendPos(c, :tr; overlay=true)
            lg = XLSX.Charts.getChartLegend(c)
            @test isnothing(XLSX.first_element_with_tag(lg, "layout"))
            @test issorted([findfirst(==(XLSX.localname(k)), XLSX.Charts.CHILD_ORDER[lkey])
                            for k in XML.eachelement(lg)])
            @test XLSX.Charts.getLegendEntryDeleted(c, 1) === true                 # entries kept

            # survives a write and reopen
            SAVE_FILES && save_outfile(xf)
            XLSX.writexlsx(tmp, xf; overwrite=true)
            xf2 = XLSX.openxlsx(tmp; mode="rw")
            SAVE_FILES && save_outfile(xf2)
            c2 = first(XLSX.Charts.getCharts(xf2["stackedbar"]))
            ax2 = XLSX.Charts.getChartAxes(c2)
            cat2 = only(filter(a -> a.kind === :catAx, ax2))
            val2 = only(filter(a -> a.kind === :valAx, ax2))
            g2 = only(XLSX.Charts.getChartGroups(c2))
            @test XLSX.Charts.getAxisTickLabelPos(c2, cat2) === :low
            @test XLSX.Charts.getAxisOrientation(c2, cat2) === :maxMin
            @test XLSX.Charts.getAxisMin(c2, val2) ≈ -1.0
            @test XLSX.Charts.getAxisMax(c2, val2) ≈ 1.0
            @test XLSX.Charts.getAxisNumberFormatCode(c2, cat2) == escaped
            @test XLSX.Charts.getGroupGapWidth(c2, g2) == 50
            @test XLSX.Charts.getLegendEntryDeleted(c2, 1) === true
            @test XLSX.Charts.getLegendPos(c2) === :tr
            @test XLSX.Charts.getLegendOverlay(c2) === true
        finally
            isfile(tmp) && rm(tmp)
        end

    end

    @testset "axis styling, units and title overlay" begin
        tmp = "chart_styling_stackedbar.xlsx"
        cp(joinpath(data_directory, "chart_kinds.xlsx"), tmp; force=true)
        try
            xf = XLSX.openxlsx(tmp; mode="rw")
            local c = first(XLSX.Charts.getCharts(xf["stackedbar"]))
            axs = XLSX.Charts.getChartAxes(c)
            catax = only(filter(a -> a.kind === :catAx, axs))
            valax = only(filter(a -> a.kind === :valAx, axs))
            snapshot(c) = XML.write(XLSX.Charts.chart_root(c))
            axorder(c, ax) = issorted([findfirst(==(XLSX.localname(k)),
                XLSX.Charts.CHILD_ORDER[(XLSX.Charts.NS_C, String(ax.kind))])
                                       for k in XML.eachelement(XLSX.Charts._axnode(c, ax))])

            # tick marks, in either vocabulary
            @test XLSX.Charts.getAxisMajorTickMark(c, catax) === :none
            c = XLSX.Charts.setAxisMajorTickMark(c, catax, :cross)
            @test XLSX.Charts.getAxisMajorTickMark(c, catax) === :cross
            c = XLSX.Charts.setAxisMajorTickMark(c, valax, :outside)
            @test XLSX.Charts.getAxisMajorTickMark(c, valax) === :out
            c = XLSX.Charts.setAxisMinorTickMark(c, valax, :inside)
            @test XLSX.Charts.getAxisMinorTickMark(c, valax) === :in
            before = snapshot(c)
            @test_throws XLSX.XLSXError XLSX.Charts.setAxisMajorTickMark(c, valax, :sideways)
            @test snapshot(c) == before                                          # atomic
            c = XLSX.Charts.setAxisMinorTickMark(c, valax, :inherit)
            @test isnothing(XLSX.Charts.getAxisMinorTickMark(c, valax))

            # gridlines: presence is the switch
            has(c, ax, tag) = !isnothing(XLSX.first_element_with_tag(XLSX.Charts._axnode(c, ax), tag))
            @test has(c, valax, "majorGridlines")
            c = XLSX.Charts.setAxisGridlines(c, valax; major=false)
            @test !has(c, valax, "majorGridlines")
            c = XLSX.Charts.setAxisGridlines(c, valax; major=true, minor=true)
            @test has(c, valax, "majorGridlines") && has(c, valax, "minorGridlines")
            @test axorder(c, valax)
            c = XLSX.Charts.setAxisGridlines(c, valax; major=true)                # already there: no-op
            @test count(k -> XLSX.localname(k) == "majorGridlines",
                XML.eachelement(XLSX.Charts._axnode(c, valax))) == 1
            @test_throws XLSX.XLSXError XLSX.Charts.setAxisGridlines(c, valax; minor=:off)
            @test XLSX.Charts.setAxisGridlines(c, valax) === c
            c = XLSX.Charts.setAxisGridlines(c, valax; major=false, minor=false)

            # tick label text
            c = XLSX.Charts.setAxisTextProp(c, catax, :size, 14)
            c = XLSX.Charts.setAxisTextProp(c, catax, :latin, "Comic Sans MS")
            txpr = XLSX.first_element_with_tag(XLSX.Charts._axnode(c, catax), "txPr")
            p = XLSX.first_element_with_tag(txpr, "p")
            rpr = XLSX.first_element_with_tag(XLSX.first_element_with_tag(p, "pPr"), "defRPr")
            @test XLSX.Charts._attr(rpr, "sz") == "1400"
            @test XLSX.Charts._attr(XLSX.first_element_with_tag(rpr, "latin"), "typeface") == "Comic Sans MS"

            # Excel's names are accepted as aliases for the DrawingML fields
            c = XLSX.Charts.setAxisTextProp(c, catax, :font, "Arial")
            rpr = XLSX.first_element_with_tag(
                XLSX.first_element_with_tag(
                    XLSX.first_element_with_tag(
                        XLSX.first_element_with_tag(XLSX.Charts._axnode(c, catax), "txPr"),
                        "p"), "pPr"), "defRPr")
            @test XLSX.Charts._attr(XLSX.first_element_with_tag(rpr, "latin"), "typeface") == "Arial"
            c = XLSX.Charts.setAxisTextProp(c, catax, :color, "FF0000")
            @test XLSX.Charts.getAxisTextProps(c, catax) isa XLSX.Charts.DrawingText
            @test_throws XLSX.XLSXError XLSX.Charts.setAxisTextProp(c, catax, :typeface, "Arial")

            # axis line and plot area border
            @test !XLSX.Charts.hasLine(XLSX.Charts.getAxisShapeProps(c, valax))  # template: noFill
            c = XLSX.Charts.setAxisLine(c, valax; color="000000", width=1)
            @test XLSX.Charts.hasLine(XLSX.Charts.getAxisShapeProps(c, valax))
            @test XLSX.Charts.setAxisLine(c, valax) === c
            c = XLSX.Charts.setAxisLine(c, valax, :none)
            @test !XLSX.Charts.hasLine(XLSX.Charts.getAxisShapeProps(c, valax))
            c = XLSX.Charts.setAxisLine(c, valax, :inherit)
            @test isnothing(XLSX.first_element_with_tag(XLSX.Charts.getAxisShapeProps(c, valax).raw, "ln"))
            @test_throws XLSX.XLSXError XLSX.Charts.setAxisLine(c, valax, :dotted)

            @test !XLSX.Charts.hasLine(XLSX.Charts.getPlotAreaShapeProps(c))
            c = XLSX.Charts.setPlotAreaLine(c; color="000000")
            @test XLSX.Charts.hasLine(XLSX.Charts.getPlotAreaShapeProps(c))
            c = XLSX.Charts.setPlotAreaLine(c, :inherit)
            @test isnothing(XLSX.first_element_with_tag(XLSX.Charts.getPlotAreaShapeProps(c).raw, "ln"))
            c = XLSX.Charts.setPlotAreaLine(c; color="000000")

            # major and minor units: value and date axes only
            @test isnothing(XLSX.Charts.getAxisMajorUnit(c, valax))
            c = XLSX.Charts.setAxisMajorUnit(c, valax, 0.2)
            @test XLSX.Charts.getAxisMajorUnit(c, valax) ≈ 0.2
            c = XLSX.Charts.setAxisMinorUnit(c, valax, 0.05)
            @test XLSX.Charts.getAxisMinorUnit(c, valax) ≈ 0.05
            @test axorder(c, valax)
            @test_throws XLSX.XLSXError XLSX.Charts.setAxisMajorUnit(c, valax, 0)
            @test_throws XLSX.XLSXError XLSX.Charts.setAxisMajorUnit(c, valax, -1)
            before = snapshot(c)
            @test_throws XLSX.XLSXError XLSX.Charts.setAxisMajorUnit(c, catax, 1)
            @test snapshot(c) == before
            c = XLSX.Charts.setAxisMinorUnit(c, valax, :inherit)
            @test isnothing(XLSX.Charts.getAxisMinorUnit(c, valax))

            # a title created by a setter is not drawn over the plot area
            title(c, ax) = XLSX.first_element_with_tag(XLSX.Charts._axnode(c, ax), "title")
            @test isnothing(title(c, valax))
            c = XLSX.Charts.setAxisTitleText(c, valax, "Proportion of responses")
            @test XLSX.Charts._bool_val(title(c, valax), "overlay") === false
            @test [XLSX.localname(k) for k in XML.eachelement(title(c, valax))] == ["tx", "overlay"]
            c = XLSX.Charts.setAxisTitleText(c, valax, "Share of responses")       # replaced, not duplicated
            @test count(k -> XLSX.localname(k) == "overlay", XML.eachelement(title(c, valax))) == 1
            @test axorder(c, valax)
            # the chart title's own overlay is kept as it was
            ov = XLSX.Charts._bool_val(XLSX.Charts.getChartTitleNode(c), "overlay")
            c = XLSX.Charts.setChartTitleText(c, "Course feedback")
            @test XLSX.Charts._bool_val(XLSX.Charts.getChartTitleNode(c), "overlay") === ov

            # survives a write and reopen
            SAVE_FILES && save_outfile(xf)
            XLSX.writexlsx(tmp, xf; overwrite=true)
            xf2 = XLSX.openxlsx(tmp; mode="rw")
            SAVE_FILES && save_outfile(xf2)
            c2 = first(XLSX.Charts.getCharts(xf2["stackedbar"]))
            ax2 = XLSX.Charts.getChartAxes(c2)
            cat2 = only(filter(a -> a.kind === :catAx, ax2))
            val2 = only(filter(a -> a.kind === :valAx, ax2))
            @test XLSX.Charts.getAxisMajorTickMark(c2, cat2) === :cross
            @test XLSX.Charts.getAxisMajorUnit(c2, val2) ≈ 0.2
            @test XLSX.Charts.hasLine(XLSX.Charts.getPlotAreaShapeProps(c2))
            @test XLSX.Charts._bool_val(title(c2, val2), "overlay") === false
            @test isnothing(XLSX.first_element_with_tag(XLSX.Charts._axnode(c2, val2), "majorGridlines"))
        finally
            isfile(tmp) && rm(tmp)
        end
    end

    @testset "rich text, cell-bound titles and part prefixes" begin
        # A sheet with a typed, a numeric and a rich-text cell, and one column chart.
        function rich_fixture()
            xf = XLSX.newxlsx()
            ws = xf[1]
            ws["A1"] = "Q1"; ws["A2"] = 10; ws["A3"] = 20; ws["B1"] = 2026
            ws["C1"] = XLSX.RichTextString(
                XLSX.RichTextRun("Big", [:bold => true, :color => "red"]),
                XLSX.RichTextRun(" red", [:bold => true, :color => "red", :italic => true]))
            c = XLSX.Charts.addChart(ws, :column; anchor = "D2:K16")
            XLSX.Charts.addSeries(c, "A2:A3")
            return xf, c
        end
        title_node(c) = XLSX.first_element_with_tag(
            XLSX.first_element_with_tag(XLSX.Charts.chart_root(c), "chart"), "title")
        run_props(atts) = XLSX.Charts._rich_run_props(atts)

        @testset "attribute mapping" begin
            @test isnothing(run_props(nothing))
            for v in (true, 1)
                @test run_props(Dict{Symbol,Any}(:bold => v)).bold === true
                @test run_props(Dict{Symbol,Any}(:italic => v)).italic === true
                @test run_props(Dict{Symbol,Any}(:under => v)).under == "sng"
                @test run_props(Dict{Symbol,Any}(:strike => v)).strike == "sngStrike"
            end
            for v in (false, 0)                                  # explicit off, not absent
                @test run_props(Dict{Symbol,Any}(:bold => v)).bold === false
                @test run_props(Dict{Symbol,Any}(:italic => v)).italic === false
                @test run_props(Dict{Symbol,Any}(:under => v)).under == "none"
                @test run_props(Dict{Symbol,Any}(:strike => v)).strike == "noStrike"
            end
            @test_throws XLSX.XLSXError run_props(Dict{Symbol,Any}(:bold => "yes"))
            @test run_props(Dict{Symbol,Any}(:vertAlign => "superscript")).baseline ≈ 0.30
            @test run_props(Dict{Symbol,Any}(:vertAlign => "subscript")).baseline ≈ -0.25
            @test run_props(Dict{Symbol,Any}(:vertAlign => "baseline")).baseline == 0.0
            @test_throws XLSX.XLSXError run_props(Dict{Symbol,Any}(:vertAlign => "raised"))
            @test run_props(Dict{Symbol,Any}(:size => 14)).size === 14.0
            @test run_props(Dict{Symbol,Any}(:name => "Arial")).latin == "Arial"
            for col in ("red", :red, "FF0000", "#FF0000", "FFFF0000", "#ffff0000")
                f = run_props(Dict{Symbol,Any}(:color => col)).fill
                @test f.kind === :solid && f.fgcolor == "FFFF0000"
            end
            @test_throws XLSX.XLSXError run_props(Dict{Symbol,Any}(:color => "notacolour"))
            p = run_props(Dict{Symbol,Any}(:bold => true))          # unset fields inherit
            @test isnothing(p.italic) && isnothing(p.size) && isnothing(p.fill)

            # Written as Excel writes them: baseline in thousandths of a percent.
            xf, c = rich_fixture()
            rts = XLSX.RichTextString(XLSX.RichTextRun("x"),
                                      XLSX.RichTextRun("2", [:vertAlign => "superscript"]),
                                      XLSX.RichTextRun("i", [:vertAlign => "subscript"]))
            XLSX.Charts.setChartTitleText(c, rts)
            rprs = collect(XLSX.elements_with_tag(only(XLSX.elements_with_tag(XLSX.first_element_with_tag(
                       XLSX.first_element_with_tag(title_node(c), "tx"), "rich"), "p")), "r"))
            @test isnothing(XLSX.first_element_with_tag(rprs[1], "rPr"))   # no attributes, no rPr
            @test [XLSX.get_attr(XLSX.first_element_with_tag(r, "rPr"), "baseline") for r in rprs[2:3]] ==
                  ["30000", "-25000"]
            SAVE_FILES && save_outfile(xf)
        end

        @testset "newlines" begin
            rts = XLSX.RichTextString(XLSX.RichTextRun("one\ntwo", [:bold => true]),
                                      XLSX.RichTextRun(" three\r\n", [:italic => true]),
                                      XLSX.RichTextRun("\nfour"))
            # In a text body a newline starts a paragraph.
            t = XLSX.Charts.DrawingText(rts)
            @test length(t.paragraphs) == 4
            @test [r.text for r in t.paragraphs[1].runs] == ["one"]
            @test [r.text for r in t.paragraphs[2].runs] == ["two", " three"]
            @test t.paragraphs[2].runs[1].props.bold === true
            @test t.paragraphs[2].runs[2].props.italic === true
            @test isempty(t.paragraphs[3].runs)
            @test isnothing(t.paragraphs[3].endprops)            # its newline is in an unformatted run
            @test [r.text for r in t.paragraphs[4].runs] == ["four"]
            @test XLSX.Charts.text_content(t) == "one\ntwo three\n\nfour"
            # A blank line keeps the formatting of the run its newline is in.
            t2 = XLSX.Charts.DrawingText(XLSX.RichTextString(XLSX.RichTextRun("a\n\nb", [:size => 9])))
            @test t2.paragraphs[2].endprops.size == 9.0

            # In a paragraph it is an a:br, carrying the run's formatting.
            p = XLSX.Charts.DrawingParagraph(rts)
            @test [r.kind for r in p.runs] == [:run, :br, :run, :run, :br, :br, :run]
            @test p.runs[2].props.bold === true
            @test p.runs[5].props.italic === true
            @test isnothing(p.runs[6].props)
        end

        @testset "setters keep run formatting" begin
            xf, c = rich_fixture()
            rts = XLSX.RichTextString(XLSX.RichTextRun("Revenue", [:bold => true, :size => 16]),
                                      XLSX.RichTextRun(" £m", [:color => "blue", :italic => true]))
            XLSX.Charts.setChartTitleText(c, rts)
            g = XLSX.Charts.getChartTitleText(c)
            @test XLSX.Charts.text_content(g) == "Revenue £m"
            r1, r2 = only(g.paragraphs).runs
            @test r1.props.bold === true && r1.props.size ≈ 16.0
            @test r2.props.italic === true && r2.props.fill.fgcolor.val == "0000FF"
            @test XLSX.Charts._bool_val(title_node(c), "overlay") === false

            ax = first(XLSX.Charts.getChartAxes(c, :value))
            XLSX.Charts.setAxisTitleText(c, ax, rts)
            ax = first(XLSX.Charts.getChartAxes(c, :value))
            ga = XLSX.Charts.getAxisTitleText(c, ax)
            @test only(ga.paragraphs).runs[2].props.italic === true

            XLSX.Charts.setLabelText(c, 1, 2, rts)
            gl = XLSX.Charts.getDataLabelText(c, XLSX.Charts.getSeriesDataLabel(c, 1, 2))
            @test XLSX.Charts.text_content(gl) == "Revenue £m"
            @test only(gl.paragraphs).runs[1].props.size ≈ 16.0
            SAVE_FILES && save_outfile(xf)
        end

        @testset "addChart with a RichTextString title" begin
            xf, _ = rich_fixture()
            ws = xf[1]
            rts = XLSX.RichTextString(XLSX.RichTextRun("Sales", [:bold => true]),
                                      XLSX.RichTextRun(" 2026\nby region", [:color => "red"]))
            c = XLSX.Charts.addChart(ws, :line; anchor = "D20:K34", title = rts)
            g = XLSX.Charts.getChartTitleText(c)
            @test XLSX.Charts.text_content(g) == "Sales 2026\nby region"
            @test length(g.paragraphs) == 2
            # The template's paragraph defaults are kept on every line ...
            for p in g.paragraphs
                @test p.props.defprops.size ≈ 14.0
                @test p.props.defprops.bold === false
            end
            # ... and the runs carry only their own attributes, with no lang.
            r1, r2 = g.paragraphs[1].runs
            @test r1.props.bold === true && isnothing(r1.props.size) && isnothing(r1.props.lang)
            @test r2.props.fill.fgcolor.val == "FF0000" && isnothing(r2.props.bold)
            @test g.paragraphs[2].runs[1].props.fill.fgcolor.val == "FF0000"

            # A plain string title keeps the template's formatting; its run has no a:rPr.
            p = XLSX.Charts.addChart(ws, :bar; anchor = "D40:K54", title = "Plain")
            gp = XLSX.Charts.getChartTitleText(p)
            @test XLSX.Charts.text_content(gp) == "Plain"
            @test only(gp.paragraphs).props.defprops.size ≈ 14.0
            @test isnothing(only(only(gp.paragraphs).runs).props)

            # A newline in a plain title starts a paragraph, as setChartTitleText does.
            m = XLSX.Charts.addChart(ws, :bar; anchor = "D60:K74", title = "Two\nlines")
            gm = XLSX.Charts.getChartTitleText(m)
            @test [only(p.runs).text for p in gm.paragraphs] == ["Two", "lines"]
            @test all(p -> p.props.defprops.size ≈ 14.0, gm.paragraphs)
            SAVE_FILES && save_outfile(xf)
        end

        @testset "binding a title to a cell" begin
            xf, c = rich_fixture()
            tmp = "chart_bound_titles.xlsx"
            try
                # A typed title, formatted, becomes bound: the formatting moves to c:txPr.
                XLSX.Charts.setChartTitleText(c, XLSX.Charts.DrawingText(XLSX.Charts.DrawingParagraph(
                    XLSX.Charts.DrawingRun("Typed", props = XLSX.Charts.DrawingRunProps(italic = true, size = 20.0)))))
                atd = XLSX.Charts.getAutoTitleDeleted(c)
                XLSX.Charts.setChartTitleText(c, XLSX.CellRef("A1"))
                @test XLSX.Charts.getChartTitleRef(c) == "Sheet1!\$A\$1"
                @test XLSX.Charts.getChartTitleRange(c).cellref == XLSX.CellRef("A1")
                @test XLSX.Charts.getChartTitle(c) == "Q1"
                @test isnothing(XLSX.Charts.getChartTitleText(c))
                d = XLSX.Charts.default_run_props(XLSX.Charts.getChartTitleTextProps(c))
                @test d.italic === true && d.size ≈ 20.0
                @test !isnothing(d.fill)                         # the template's colour survives
                @test XLSX.Charts.getAutoTitleDeleted(c) === atd
                @test XLSX.Charts._bool_val(title_node(c), "overlay") === false

                # A numeric cell still gives a string cache: c:tx allows only c:strRef.
                XLSX.Charts.setChartTitleText(c, XLSX.SheetCellRef("Sheet1!B1"))
                tx = XLSX.first_element_with_tag(title_node(c), "tx")
                @test !isnothing(XLSX.first_element_with_tag(tx, "strRef"))
                @test isnothing(XLSX.first_element_with_tag(tx, "numRef"))
                @test XLSX.Charts.getChartTitle(c) == "2026"

                # An empty cell binds, caching no value: ptCount 1, no c:pt.
                XLSX.Charts.setChartTitleText(c, XLSX.CellRef("Z99"))
                sr = XLSX.first_element_with_tag(XLSX.first_element_with_tag(title_node(c), "tx"), "strRef")
                cache = XLSX.first_element_with_tag(sr, "strCache")
                @test XLSX.get_attr(XLSX.first_element_with_tag(cache, "ptCount"), "val") == "1"
                @test isnothing(XLSX.first_element_with_tag(cache, "pt"))
                @test isnothing(XLSX.Charts.getChartTitle(c))             # as for a cx: title

                # A rich-text cell lends the formatting all its runs share: bold and red,
                # not italic, which only one run has.
                XLSX.Charts.setChartTitleText(c, "Plain typed")
                XLSX.Charts.setChartTitleText(c, XLSX.SheetCellRef("Sheet1!C1"))
                d = XLSX.Charts.default_run_props(XLSX.Charts.getChartTitleTextProps(c))
                @test d.bold === true && d.fill.fgcolor.val == "FF0000"
                @test d.italic === true                          # kept from the earlier typed title
                @test XLSX.Charts.getChartTitle(c) == "Big red"
                # On a title with no italic, the cell's one italic run adds none.
                b0 = XLSX.Charts.addChart(xf[1], :column; anchor = "D60:K74", title = XLSX.SheetCellRef("Sheet1!C1"))
                d0 = XLSX.Charts.default_run_props(XLSX.Charts.getChartTitleTextProps(b0))
                @test d0.bold === true && d0.italic === false   # the template's i="0"

                # Bound back to typed: c:rich returns and the reference is gone.
                XLSX.Charts.setChartTitleText(c, "Typed again")
                @test isnothing(XLSX.Charts.getChartTitleRef(c))
                @test XLSX.Charts.text_content(XLSX.Charts.getChartTitleText(c)) == "Typed again"
                tx = XLSX.first_element_with_tag(title_node(c), "tx")
                @test isnothing(XLSX.first_element_with_tag(tx, "strRef"))

                # Axis titles, created from nothing.
                ax = first(XLSX.Charts.getChartAxes(c, :category))
                XLSX.Charts.setAxisTitleText(c, ax, XLSX.CellRef("B1"))
                ax = first(XLSX.Charts.getChartAxes(c, :category))
                @test XLSX.Charts.getAxisTitleRef(c, ax) == "Sheet1!\$B\$1"
                @test isnothing(XLSX.Charts.getAxisTitleText(c, ax))
                @test XLSX.Charts.getAxisTitle(c, ax) == "2026"           # the cached text
                r = XLSX.Charts.getAxisTitleRange(c, ax)
                @test r isa XLSX.SheetCellRef && r.sheet == "Sheet1" && r.cellref == XLSX.CellRef("B1")
                vax = first(XLSX.Charts.getChartAxes(c, :value))
                XLSX.Charts.setAxisTitleText(c, vax, "Units")
                vax = first(XLSX.Charts.getChartAxes(c, :value))
                @test XLSX.Charts.getAxisTitle(c, vax) == "Units"
                @test isnothing(XLSX.Charts.getAxisTitleRange(c, vax))
                XLSX.Charts.setAxisTitleText(c, vax, XLSX.SheetCellRef("Sheet1!C1"))
                vax = first(XLSX.Charts.getChartAxes(c, :value))
                @test XLSX.Charts.getAxisTitleRef(c, vax) == "Sheet1!\$C\$1"

                # addChart binds too; a string is never read as a reference.
                b = XLSX.Charts.addChart(xf[1], :column; anchor = "D20:K34", title = XLSX.CellRef("A1"))
                @test XLSX.Charts.getChartTitleRef(b) == "Sheet1!\$A\$1"
                @test XLSX.Charts.default_run_props(XLSX.Charts.getChartTitleTextProps(b)).size ≈ 14.0
                s = XLSX.Charts.addChart(xf[1], :column; anchor = "D40:K54", title = "A1")
                @test isnothing(XLSX.Charts.getChartTitleRef(s))
                @test XLSX.Charts.text_content(XLSX.Charts.getChartTitleText(s)) == "A1"
                # A bad reference fails before the chart is added.
                n = length(XLSX.Charts.getCharts(xf[1]))
                @test_throws XLSX.XLSXError XLSX.Charts.addChart(xf[1], :column; anchor = "D80:K94",
                                                                 title = XLSX.SheetCellRef("Nope!A1"))
                @test length(XLSX.Charts.getCharts(xf[1])) == n

                # Survives a write and reopen.
                SAVE_FILES && save_outfile(xf)
                XLSX.writexlsx(tmp, xf; overwrite = true)
                cs = XLSX.Charts.getCharts(XLSX.openxlsx(tmp)[1])
                c2 = cs[findfirst(x -> XLSX.Charts.getChartTitle(x) == "Typed again", cs)]
                ax2 = first(XLSX.Charts.getChartAxes(c2, :category))
                @test XLSX.Charts.getAxisTitleRef(c2, ax2) == "Sheet1!\$B\$1"
                vax2 = first(XLSX.Charts.getChartAxes(c2, :value))
                @test XLSX.Charts.getAxisTitleRef(c2, vax2) == "Sheet1!\$C\$1"
                d2 = XLSX.Charts.default_run_props(XLSX.Charts.getAxisTitleTextProps(c2, vax2))
                @test d2.bold === true && d2.fill.fgcolor.val == "FF0000"
                b2 = cs[findfirst(x -> XLSX.Charts.getChartTitleRef(x) == "Sheet1!\$A\$1", cs)]
                @test XLSX.Charts.getChartTitle(b2) == "Q1"
            finally
                isfile(tmp) && rm(tmp)
            end
        end

        @testset "rich titles survive a write and reopen" begin
            xf, c = rich_fixture()
            tmp = "chart_rich_titles.xlsx"
            try
                rts = XLSX.RichTextString(XLSX.RichTextRun("H"),
                                          XLSX.RichTextRun("2", [:vertAlign => "subscript", :name => "Arial"]),
                                          XLSX.RichTextRun("O", [:under => true, :strike => false]))
                XLSX.Charts.setChartTitleText(c, rts)
                SAVE_FILES && save_outfile(xf)
                XLSX.writexlsx(tmp, xf; overwrite = true)
                c2 = first(XLSX.Charts.getCharts(XLSX.openxlsx(tmp)[1]))
                r = only(XLSX.Charts.getChartTitleText(c2).paragraphs).runs
                @test [x.text for x in r] == ["H", "2", "O"]
                @test isnothing(r[1].props)
                @test r[2].props.baseline ≈ -0.25 && r[2].props.latin == "Arial"
                @test r[3].props.under == "sng" && r[3].props.strike == "noStrike"
            finally
                isfile(tmp) && rm(tmp)
            end
        end

        @static if VERSION >= v"1.11-"
            @testset "styled strings as chart text" begin
                xf, c = rich_fixture()
                ss = styled"{bold:Revenue} {(foreground=#ff0000):£m}"
                check(t) = begin
                    r = only(t.paragraphs).runs
                    @test [x.text for x in r] == ["Revenue", " ", "£m"]
                    @test r[1].props.bold === true
                    @test isnothing(r[2].props)                  # unstyled: inherits
                    @test r[3].props.fill.fgcolor.val == "FF0000"
                end
                XLSX.Charts.setChartTitleText(c, ss)
                check(XLSX.Charts.getChartTitleText(c))
                ax = first(XLSX.Charts.getChartAxes(c, :value))
                XLSX.Charts.setAxisTitleText(c, ax, ss)
                check(XLSX.Charts.getAxisTitleText(c, first(XLSX.Charts.getChartAxes(c, :value))))
                XLSX.Charts.setLabelText(c, 1, 1, ss)
                check(XLSX.Charts.getDataLabelText(c, XLSX.Charts.getSeriesDataLabel(c, 1, 1)))
                a = XLSX.Charts.addChart(xf[1], :column; anchor = "D20:K34", title = ss)
                check(XLSX.Charts.getChartTitleText(a))
                @test only(XLSX.Charts.getChartTitleText(a).paragraphs).props.defprops.size ≈ 14.0
                # A substring of a styled string is a styled string too.
                XLSX.Charts.setChartTitleText(c, SubString(ss, 1, 7))
                @test only(only(XLSX.Charts.getChartTitleText(c).paragraphs).runs).props.bold === true
                # An empty one is plain empty text.
                XLSX.Charts.setChartTitleText(c, styled"")
                @test XLSX.Charts.text_content(XLSX.Charts.getChartTitleText(c)) == ""
                SAVE_FILES && save_outfile(xf)
            end
        end

        @testset "new elements take the part's own prefixes" begin
            # A chart part may bind the chart and DrawingML namespaces to any prefix;
            # Excel's `c` and `a` are a convention. Rename both, then add to the part.
            # Renamed in the tree, not via text: XML.write indents, and whitespace
            # between elements is not what this is testing.
            function reprefixed(c)
                ren(t) = startswith(t, "c:") ? "ch:" * t[3:end] : startswith(t, "a:") ? "dr:" * t[3:end] : t
                root = XLSX.Charts._rewrite(XLSX.Charts.chart_root(c)) do e
                    typeof(e)(XML.nodetype(e), ren(XML.tag(e)), e.attributes, XML.value(e), e.children)
                end
                for (old, new, ns) in (("c", "ch", XLSX.Charts.NS_C), ("a", "dr", XLSX.NS_A))
                    root = XLSX.with_attribute(XLSX.with_attribute(root, "xmlns:$old", nothing), "xmlns:$new", ns)
                end
                XLSX.Charts.set_chart_root!(c, root)
                return c
            end
            # Every element in the namespaces the part renamed, by prefix as written.
            function prefixes_used(n, acc = Set{String}())
                if XML.nodetype(n) == XML.Element
                    t = XML.tag(n)
                    occursin(':', t) && push!(acc, first(split(t, ':')))
                    for k in something(XML.children(n), ())
                        prefixes_used(k, acc)
                    end
                end
                return acc
            end
            xf, c = rich_fixture()
            ws = xf[1]
            tmp = "chart_prefixes.xlsx"
            try
                reprefixed(c)
                @test XLSX.ns_prefixes(XLSX.Charts.chart_root(c))[XLSX.Charts.NS_C] == "ch"
                XLSX.Charts.addSeries(c, "A2:A3"; name_ref = "B1")
                XLSX.Charts.setChartTitleText(c, XLSX.CellRef("C1"))
                XLSX.Charts.setAxisTitleText(c, first(XLSX.Charts.getChartAxes(c, :value)), XLSX.CellRef("B1"))
                used = prefixes_used(XLSX.Charts.chart_root(c))
                @test !("c" in used) && !("a" in used)
                @test length(XLSX.Charts.getChartSeries(c)) == 2
                @test XLSX.Charts.getChartTitleRef(c) == "Sheet1!\$C\$1"

                # A scatter series' line is built from the template, then renamed.
                s = XLSX.Charts.addChart(ws, :scatter; anchor = "D20:K34")
                XLSX.Charts.addSeries(s, "A2:A3"; categories = "A2:A3")
                reprefixed(s)
                XLSX.Charts.addSeries(s, "A2:A3"; categories = "A2:A3", line = true)
                used = prefixes_used(XLSX.Charts.chart_root(s))
                @test !("c" in used) && !("a" in used)
                @test XLSX.Charts.getSeriesLine(s, 2).value.width ≈ 1.5           # 19050 EMU

                SAVE_FILES && save_outfile(xf)
                XLSX.writexlsx(tmp, xf; overwrite = true)
                cs = XLSX.Charts.getCharts(XLSX.openxlsx(tmp)[1])
                c2 = cs[findfirst(x -> XLSX.Charts.getChartTitleRef(x) == "Sheet1!\$C\$1", cs)]
                @test XLSX.Charts.getChartTitle(c2) == "Big red"
                @test length(XLSX.Charts.getChartSeries(c2)) == 2
            finally
                isfile(tmp) && rm(tmp)
            end
        end

        @testset "an indented chart part can be edited" begin
            # Excel writes chart parts without whitespace; another tool may indent
            # them. XML.write indents, so a round trip through it gives such a part.
            xf, c = rich_fixture()
            tmp = "chart_indented.xlsx"
            try
                XLSX.Charts.set_chart_root!(c, XLSX.xml_root_element(
                    parse(XML.write(XLSX.Charts.chart_root(c)), XML.Node)))
                chartnode = XLSX.first_element_with_tag(XLSX.Charts.chart_root(c), "chart")
                @test any(k -> XML.nodetype(k) == XML.Text, XML.children(chartnode))   # it is indented
                ax = first(XLSX.Charts.getChartAxes(c, :value))
                XLSX.Charts.setSeriesFill(c, 1, "FF0000")
                XLSX.Charts.setChartTitleText(c, "Typed")
                XLSX.Charts.setChartTitleText(c, XLSX.CellRef("A1"))
                XLSX.Charts.setAxisTitleText(c, ax, "Units")
                XLSX.Charts.setLabelText(c, 1, 1, "Hi")
                XLSX.Charts.addSeries(c, "A2:A3")
                t = title_node(c)
                @test issorted([XLSX.Charts._slot(XLSX.Charts.CHILD_ORDER[(XLSX.Charts.NS_C, "title")], XLSX.localname(k))
                                for k in XML.eachelement(t)])
                @test XLSX.Charts.getChartTitleRef(c) == "Sheet1!\$A\$1"
                @test XLSX.Charts.getSeriesFill(c, 1).value.fgcolor.val == "FF0000"
                SAVE_FILES && save_outfile(xf)
                XLSX.writexlsx(tmp, xf; overwrite = true)
                c2 = only(XLSX.Charts.getCharts(XLSX.openxlsx(tmp)[1]))
                @test XLSX.Charts.getChartTitle(c2) == "Q1"
                @test length(XLSX.Charts.getChartSeries(c2)) == 2
            finally
                isfile(tmp) && rm(tmp)
            end
        end

        @testset "a new chart's anchor takes the drawing's prefixes" begin
            drawing_path(xf) = XLSX._drawing_path_for_sheet(xf, XLSX.get_worksheet_internal_file(xf[1]))
            drawing_root(xf) = XLSX.xml_root_element(xf.data[drawing_path(xf)])
            function set_drawing_root!(xf, new)
                doc = xf.data[drawing_path(xf)]
                xf.data[drawing_path(xf)] = XLSX.Charts.replace_child(doc, XLSX.xml_root_element(doc), new)
            end
            function element_prefixes(n, acc = Set{String}())
                if XML.nodetype(n) == XML.Element
                    t = XML.tag(n)
                    occursin(':', t) && push!(acc, first(split(t, ':')))
                    foreach(k -> element_prefixes(k, acc), something(XML.children(n), ()))
                end
                return acc
            end
            xf, _ = rich_fixture()
            ws = xf[1]
            tmp = "chart_drawing_prefixes.xlsx"
            try
                # Rename the drawing's xdr and a to wd and da, as another tool might.
                ren(t) = startswith(t, "xdr:") ? "wd:" * t[5:end] : startswith(t, "a:") ? "da:" * t[3:end] : t
                root = XLSX.Charts._rewrite(drawing_root(xf)) do e
                    typeof(e)(XML.nodetype(e), ren(XML.tag(e)), e.attributes, XML.value(e), e.children)
                end
                for (old, new, ns) in (("xdr", "wd", XLSX.NS_XDR), ("a", "da", XLSX.NS_A))
                    root = XLSX.with_attribute(XLSX.with_attribute(root, "xmlns:$old", nothing), "xmlns:$new", ns)
                end
                set_drawing_root!(xf, root)

                XLSX.Charts.addChart(ws, :line; anchor = "D20:K34")
                XLSX.Charts.addChartEx(ws, :waterfall, "A2:A3"; anchor = "D40:K54")
                used = element_prefixes(drawing_root(xf))
                @test !("xdr" in used) && !("a" in used)
                @test "wd" in used && "da" in used
                SAVE_FILES && save_outfile(xf)
                XLSX.writexlsx(tmp, xf; overwrite = true)
                @test length(XLSX.Charts.getCharts(XLSX.openxlsx(tmp)[1])) == 3

                # A drawing that doesn't declare DrawingML at all: the anchor declares it.
                root = drawing_root(xf)
                set_drawing_root!(xf, XLSX.with_attribute(root, "xmlns:da", nothing))
                XLSX.Charts.addChart(ws, :bar; anchor = "D60:K74")
                anchor = last(collect(XML.eachelement(drawing_root(xf))))
                @test XLSX.get_attr(anchor, "xmlns:a") == XLSX.NS_A
            finally
                isfile(tmp) && rm(tmp)
            end
        end

        @testset "a bound title on a chartsheet must name its sheet" begin
            xf, _ = rich_fixture()
            @test_throws XLSX.XLSXError XLSX.Charts.addChart(xf, :column; title = XLSX.CellRef("A1"))
            @test XLSX.sheetnames(xf) == ["Sheet1"]              # nothing half-built
            err = try
                XLSX.Charts.addChart(xf, :column; title = XLSX.CellRef("A1"))
            catch e
                e
            end
            @test occursin("SheetCellRef", err.msg) && !occursin("series", err.msg)

            c = XLSX.Charts.addChart(xf, :column; title = XLSX.SheetCellRef("Sheet1!A1"))
            @test XLSX.Charts.getChartTitleRef(c) == "Sheet1!\$A\$1"
            @test XLSX.Charts.getChartTitle(c) == "Q1"
            @test_throws XLSX.XLSXError XLSX.Charts.setChartTitleText(c, XLSX.CellRef("B1"))
            XLSX.Charts.addSeries(c, "Sheet1!A2:A3")
            SAVE_FILES && save_outfile(xf)
        end
    end


    @testset "chart_extras.xlsx: elements with no other fixture" begin
        xf = XLSX.readxlsx(joinpath(data_directory, "chart_extras.xlsx"))
        chart(sheet) = only(XLSX.Charts.getCharts(xf[sheet]))

        # Drop lines, on a line chart's group.
        c = chart("droplines")
        g = only(XLSX.Charts.getChartGroups(c))
        dl = XLSX.Charts.getGroupDropLines(c, g)
        @test !isnothing(dl)
        @test dl.line.width ≈ 0.75                                   # 9525 EMU
        @test isnothing(XLSX.Charts.getGroupHiLowLines(c, g))

        # High-low lines and up/down bars together.
        c = chart("hilolines")
        g = only(XLSX.Charts.getChartGroups(c))
        @test XLSX.Charts.getGroupHiLowLines(c, g).line.width ≈ 0.75
        @test isnothing(XLSX.Charts.getGroupDropLines(c, g))
        b = XLSX.Charts.getGroupUpDownBars(c, g)
        @test b.gap_width == 150
        @test XLSX.Charts.hasFill(XLSX.Charts.getUpBarShapeProps(c, b))
        @test XLSX.Charts.hasFill(XLSX.Charts.getDownBarShapeProps(c, b))

        # Polynomial and moving-average trendlines, one per series.
        c = chart("trends")
        tp = only(XLSX.Charts.getSeriesTrendlines(c, 1))
        @test tp.kind === :poly
        @test tp.order == 2
        @test isnothing(tp.period)
        @test tp.disp_rsqr === false
        tm = only(XLSX.Charts.getSeriesTrendlines(c, 2))
        @test tm.kind === :movingAvg
        @test tm.period == 2
        @test isnothing(tm.order)
        ln = XLSX.Charts.getTrendlineShapeProps(c, tp).line
        @test ln.width ≈ 1.5
        @test ln.dash == "sysDot"

        # Custom error bars with cell references for both directions.
        c = chart("errbars")
        e = only(XLSX.Charts.getSeriesErrorBars(c, 1))
        @test e.bar_type === :both
        @test e.value_type === :cust
        @test isnothing(e.direction)                                 # no c:errDir written
        @test e.no_end_cap === false
        refs = XLSX.Charts.getErrorBarsCustomRefs(c, e)
        @test refs.plus.ref == "errbars!\$C\$2:\$C\$5"
        @test refs.minus.ref == "errbars!\$D\$2:\$D\$5"
        @test refs.plus.data == [12, 18, 25, 9]

        # A title bound to a cell: a reference and its cached string, no rich text.
        c = chart("boundtitle")
        @test XLSX.Charts.getChartTitleRef(c) == "boundtitle!\$F\$1"
        r = XLSX.Charts.getChartTitleRange(c)                         # Excel's own bound title
        @test r isa XLSX.SheetCellRef && r.sheet == "boundtitle" && r.cellref == XLSX.CellRef("F1")
        @test XLSX.Charts.getChartTitle(c) == "My Bound Title"
        @test isnothing(XLSX.Charts.getChartTitleText(c))

        # The a:ln join group.
        c = chart("join")
        l = XLSX.Charts.getSeriesLine(c, 1).value
        @test l.width ≈ 3.0
        @test l.join == "bevel"
    end

    @testset "gridlines written without c:spPr are still gridlines" begin
        path = "chart_gridlines_bare.xlsx"
        cp(joinpath(data_directory, "chart_basic.xlsx"), path; force=true)
        xf = XLSX.openxlsx(path; mode="rw")
        c = XLSX.Charts.getChart(xf, "chart1")
        ax = only(XLSX.Charts.getChartAxes(c, :value))
        root = XLSX.Charts.chart_root(c)
        axn = XLSX.Charts._axis_node(c, root, ax)
        new = XLSX.Charts.rebuild_path(root,
            [(XLSX.Charts.NS_C, "chart") => "chart",
                (XLSX.Charts.NS_C, "plotArea") => "plotArea",
                (XLSX.Charts.NS_C, "valAx") => ("valAx", n -> n === axn),
                (XLSX.Charts.NS_C, "majorGridlines") => "majorGridlines"],
            gl -> XLSX.Charts.remove_child(gl, "spPr");
            prefixes=XLSX.ns_prefixes(root))
        XLSX.Charts.set_chart_root!(c, new)

        gl = XLSX.Charts.getAxisGridlines(c, ax)
        @test !isnothing(gl)                                         # still there
        @test isnothing(gl.fill) && isnothing(gl.line)               # but unformatted
        isfile(path) && rm(path)
        SAVE_FILES && save_outfile(xf)
    end

    @testset "Non-sequential series c:idx" begin
        # chart_idx_gaps is chart_basic with the two series' `c:idx` changed to 3 and 7,
        # as Excel leaves them after series have been deleted. A series *position* counts
        # from 1 in document order; a value's `series_idx` is the `c:idx`. The two differ
        # here, so anything that confuses them fails.
        xf = XLSX.opentemplate(joinpath(data_directory, "chart_idx_gaps.xlsx"))
        c = XLSX.Charts.getCharts(xf)[1]

        ss = XLSX.Charts.getChartSeries(c)
        @test length(ss) == 2
        @test [s.idx for s in ss] == [3, 7]
        @test [s.order for s in ss] == [0, 1]
        @test [s.name for s in ss] == ["2024", "2025"]        # document order, not idx order
        @test length(XLSX.Charts.getChartRanges(c)) == 2

        # A setter taking a position writes to that series, not to the one whose idx matches.
        was = XLSX.Charts.getSeriesFill(c, 1).value
        XLSX.Charts.setSeriesFill(c, 2, "FF0000")
        @test XLSX.Charts.getSeriesFill(c, 2).value.fgcolor.rgb == "FF0000"
        @test XLSX.Charts.getSeriesFill(c, 1).value == was      # series 1 untouched

        # A value carries the owning series' `c:idx`, and is resolved by it.
        XLSX.Charts.setLabelDeleted(c, 2, 1, true)              # creates a c:dLbl
        d = only(XLSX.Charts.getSeriesDataLabels(c, 2))
        @test d.series_idx == 7                                 # the key, not the position
        @test d.idx == 0                                        # c:idx of the label, 0-based
        @test d.delete === true

        # The value still addresses the same label after an unrelated write.
        XLSX.Charts.setSeriesLineColor(c, 2, "0000FF")
        @test XLSX.Charts.getSeriesDataLabel(c, 2, 1).idx == d.idx
        @test XLSX.Charts.getSeriesDataLabel(c, 2, 1).series_idx == 7

        SAVE_FILES && save_outfile(xf)
    end

    @testset "setLabelFlag and getLabelFlag" begin
        C = XLSX.Charts
        # Excel's labels, turned on with Chart Elements → Data Labels, for one chart of
        # each kind (sheets named after the kinds), and then turned off again.
        excel(file) = Dict(c.sheet => c for c in C.getCharts(XLSX.readxlsx(joinpath(data_directory, file))))
        on, off = excel("chart_labels_on.xlsx"), excel("chart_labels_off.xlsx")
        lbls(c, i) = (s = C._series_nodes(C.chart_root(c))[i][2]; n = XLSX.first_element_with_tag(s, "dLbls");
                      isnothing(n) ? nothing : XLSX.XML.write(n))

        xf = XLSX.newxlsx()
        ws = xf[1]
        ws["A1"] = ["" "North" "South"; 2 4 7; 3 5 8; 4 6 9; 5 7 10]
        for (row, kind) in enumerate(keys(C.C_KINDS))
            kind === :combo && continue
            c = C.addChart(ws, kind; anchor = "F$(20row):M$(20row + 15)")
            if kind === :bubble
                C.addSeries(c, "B2:B5"; categories = "A2:A5", bubble_sizes = "C2:C5")
            else
                C.addSeries(c, "B2:B5"; categories = "A2:A5", name_ref = "B1")
                kind in (:pie, :doughnut) || C.addSeries(c, "C2:C5"; categories = "A2:A5", name_ref = "C1")
            end
            n = length(C.getChartSeries(c))
            @test all(isnothing(C.getLabelFlag(c, i, :value)) || C.getLabelFlag(c, i, :value) === false for i in 1:n)
            foreach(i -> C.setLabelFlag(c, i, :value, true), 1:n)
            @test [lbls(c, i) for i in 1:n] == [lbls(on[String(kind)], i) for i in 1:n]   # as Excel writes them
            @test C.getLabelFlag(c, 1, :value) === true && C.getLabelFlag(c, 1, :seriesName) === false
        end

        # The flags are set one at a time, and the labels stay when all are off.
        c = C.getCharts(ws)[1]
        C.setLabelFlag(c, 1, :seriesName, true)
        @test C.getLabelFlag(c, 1, :seriesName) === true && C.getLabelFlag(c, 1, :value) === true
        C.setLabelFlag(c, 1, :value, false)
        C.setLabelFlag(c, 1, :seriesName, false)
        @test C.getLabelFlag(c, 1, :value) === false && !isnothing(lbls(c, 1))

        # Labels a series has deleted are made again; ones of single points are kept.
        g = XLSX.opentemplate(joinpath(data_directory, "chart_labels_off.xlsx"))
        col = only(x for x in C.getCharts(g) if x.sheet == "column")
        @test C.getLabelFlag(col, 1, :value) === false                    # c:delete
        C.setLabelFlag(col, 1, :value, true)
        @test lbls(col, 1) == lbls(on["column"], 1)
        C.setLabelFlag(col, 2, :value, true)
        C.setLabelText(col, 2, 3, "three")
        C.setLabelFlag(col, 2, :categoryName, true)
        @test occursin(">three<", lbls(col, 2))

        # Parts a kind doesn't have are refused.
        @test_throws XLSX.XLSXError C.setLabelFlag(c, 1, :percent, true)
        @test_throws XLSX.XLSXError C.setLabelFlag(c, 1, :bubbleSize, true)
        @test_throws XLSX.XLSXError C.setLabelFlag(c, 1, :values, true)
        @test_throws XLSX.XLSXError C.getLabelFlag(c, 1, :percent)
        pie = only(x for x in C.getCharts(ws) if C.getChartType(x) === :pieChart)
        C.setLabelFlag(pie, 1, :percent, true)
        @test C.getLabelFlag(pie, 1, :percent) === true
        SAVE_FILES && save_outfile(xf)
    end
end
