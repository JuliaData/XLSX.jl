@testset "ChartExProps" begin
    @testset "series, layout and data blocks" begin
        f = XLSX.readxlsx(joinpath(data_directory, "chart_ex.xlsx"))
        cx = only(filter(x -> x isa XLSX.Charts.ChartEx, XLSX.Charts.getCharts(f)))

        @test XLSX.Charts.getChartSeriesCount(cx) == 1
        @test XLSX.Charts.getSeriesLayout(cx, 1) == :waterfall
        @test_throws XLSX.XLSXError XLSX.Charts.getSeriesLayout(cx, 2)

        blocks = XLSX.Charts.getChartDataBlocks(cx)
        @test length(blocks) == 1
        b = only(blocks)
        @test b.id == 0
        @test [d.kind for d in b.dimensions] == [:str, :num]
        @test [d.type for d in b.dimensions] == [:cat, :val]
        @test all(d -> startswith(d.formula, "_xlchart."), b.dimensions)
        @test all(d -> d.range isa XLSX.SheetCellRange, b.dimensions)

        @test XLSX.Charts.getSeriesData(cx, 1) == b
        @test isnothing(XLSX.Charts.getSeriesName(cx, 1))
        @test isnothing(XLSX.Charts.getSeriesNameRange(cx, 1))
    end

    @testset "layouts fixture" begin
        f = XLSX.readxlsx(joinpath(data_directory, "chartex_layouts.xlsx"))
        cxs = filter(x -> x isa XLSX.Charts.ChartEx, XLSX.Charts.getCharts(f))
        chart(sheet) = only(filter(x -> x.sheet == sheet, cxs))

        @testset "waterfall" begin
            c = chart("waterfall")
            @test XLSX.Charts.getSeriesLayout(c, 1) == :waterfall
            @test XLSX.Charts.getSeriesSubtotals(c, 1) == [2, 5]
            @test XLSX.Charts.getLabelPosition(c, 1) == :outEnd
            @test XLSX.Charts.getLabelFlag(c, 1, :value) === true
            @test XLSX.Charts.getLabelFlag(c, 1, :seriesName) === false
            @test_throws XLSX.XLSXError XLSX.Charts.getLabelFlag(c, 1, :nonsense)
            @test isnothing(XLSX.Charts.getChartTitle(c))
            @test isnothing(XLSX.Charts.getSeriesName(c, 1))
        end

        @testset "funnel" begin
            c = chart("funnel")
            @test XLSX.Charts.getSeriesLayout(c, 1) == :funnel
            @test only(only(XLSX.Charts.getChartDataBlocks(c)).dimensions).formula == "_xlchart.v2.3"
            @test isnothing(XLSX.Charts.getSeriesSubtotals(c, 1))
            @test isnothing(XLSX.Charts.getLabelPosition(c, 1))
            @test isnothing(XLSX.Charts.getLabelFlag(c, 1, :value))
        end

        @testset "treemap and sunburst" begin
            t = chart("treemap")
            @test [d.type for d in only(XLSX.Charts.getChartDataBlocks(t)).dimensions] == [:cat, :size]
            @test XLSX.Charts.getSeriesName(t, 1) == "Value"
            @test !isnothing(XLSX.Charts.getSeriesNameRange(t, 1))
            @test XLSX.Charts.getSeriesParentLabelLayout(t, 1) == :overlapping
            @test XLSX.Charts.getLabelPosition(t, 1) == :inEnd
            @test XLSX.Charts.getLabelFlag(t, 1, :categoryName) === true

            s = chart("sunburst")
            @test XLSX.Charts.getSeriesLayout(s, 1) == :sunburst
            @test isnothing(XLSX.Charts.getSeriesParentLabelLayout(s, 1))   # no cx:layoutPr
            @test XLSX.Charts.getLabelPosition(s, 1) == :ctr
        end

        @testset "histogram" begin
            c = chart("histogram")
            @test XLSX.Charts.getChartType(c) == :histogram                    # binning read inside layoutPr
            @test XLSX.Charts.getSeriesLayout(c, 1) == :clusteredColumn
            @test XLSX.Charts.getSeriesBinning(c, 1) == XLSX.Charts.ChartExBinning(:r, 0.0, 100.0, 10.0, nothing)
            @test XLSX.Charts.getChartTitle(c) == "Chart Title"                 # typed: txData without f
            @test isnothing(XLSX.Charts.getChartTitleRange(c))
            @test XLSX.Charts.getSeriesName(c, 1) == "Values"
        end

        @testset "pareto" begin
            c = chart("pareto")
            @test XLSX.Charts.getChartSeriesCount(c) == 2
            @test [XLSX.Charts.getSeriesLayout(c, i) for i in 1:2] == [:clusteredColumn, :clusteredColumn]
            @test [XLSX.Charts.getSeriesLayout(c, i; pareto=true) for i in 1:2] == [:paretoLine, :paretoLine]
            @test [XLSX.Charts.getSeriesHidden(c, i) for i in 1:2] == [nothing, true]
            @test [XLSX.Charts.getSeriesHidden(c, i; pareto=true) for i in 1:2] == [nothing, nothing]
            @test [XLSX.Charts.getSeriesAxisIds(c, i) for i in 1:2] == [[1], [1]]
            @test [XLSX.Charts.getSeriesAxisIds(c, i; pareto=true) for i in 1:2] == [[2], [2]]
            @test [XLSX.Charts.getSeriesAggregation(c, i) for i in 1:2] == [true, true]
            @test isnothing(XLSX.Charts.getSeriesData(c, 1; pareto=true))
            blocks = XLSX.Charts.getChartDataBlocks(c)
            @test length(blocks) == 2
            @test blocks[1].dimensions[1].formula == blocks[2].dimensions[1].formula  # shared name
            @test XLSX.Charts.getSeriesData(c, 2) == blocks[2]
            @test XLSX.Charts.getChartType(c) == :pareto
        end

        @testset "box and whisker" begin
            c = chart("boxwhisker")
            @test XLSX.Charts.getChartSeriesCount(c) == 3
            @test [XLSX.Charts.getSeriesName(c, i) for i in 1:3] == ["Group A", "Group B", "Group C"]
            @test [XLSX.Charts.getSeriesQuartileMethod(c, i) for i in 1:3] == [:exclusive, :inclusive, :exclusive]
            @test [XLSX.Charts.getSeriesLayoutFlag(c, i, :meanMarker) for i in 1:3] == [true, false, true]
            @test [XLSX.Charts.getSeriesLayoutFlag(c, i, :outliers) for i in 1:3] == [true, false, true]
            @test all(i -> XLSX.Charts.getSeriesLayoutFlag(c, i, :meanLine) === false, 1:3)
            @test [XLSX.Charts.getSeriesData(c, i).id for i in 1:3] == [0, 1, 2]
            @test XLSX.Charts.getChartType(c) == :boxWhisker
        end

        @testset "bound title and series name" begin
            c = chart("bound")
            @test XLSX.Charts.getChartTitle(c) == "Title here"
            @test !isnothing(XLSX.Charts.getChartTitleRange(c))
            @test XLSX.Charts.getSeriesName(c, 1) == "Series name"
            @test !isnothing(XLSX.Charts.getSeriesNameRange(c, 1))
            @test XLSX.Charts.getSeriesSubtotals(c, 1) == Int[]              # present, empty
        end
    end

    @testset "formatting cascade" begin
        f = XLSX.readxlsx(joinpath(data_directory, "chartex_formatted.xlsx"))
        c = only(filter(x -> x isa XLSX.Charts.ChartEx, XLSX.Charts.getCharts(f)))

        # Series fill: theme colour with a transform.
        e = XLSX.Charts.getSeriesFill(c, 1)
        @test e.site.level == :series
        @test e.value.fgcolor.val == "accent2"

        # Point 5 (idx 4) has its own fill: explicit colour with transforms.
        e5 = XLSX.Charts.getSeriesFill(c, 1; point=5)
        @test e5.site.level == :point
        @test [s.level for s in e5.chain] == [:point, :series]
        @test e5.value.fgcolor.kind == :srgb
        @test e5.value.fgcolor.val == "4EA72E"                                  # as written
        @test e5.value.fgcolor.transforms == [:lumMod => 60000, :lumOff => 40000]
        @test e5.value.fgcolor.rgb == "8ED973"                                  # transforms applied

        # Point 1 has no dataPt: no point rung, resolves at the series.
        e1 = XLSX.Charts.getSeriesFill(c, 1; point=1)
        @test e1.site.level == :series
        @test [s.level for s in e1.chain] == [:series]
        @test_throws XLSX.XLSXError XLSX.Charts.getSeriesFill(c, 1; point=0)

        # Label text: set on dataLabels/txPr defRPr.
        s = XLSX.Charts.getLabelTextProp(c, 1, :size)
        @test s.value ≈ 11.0
        @test s.site.level == :series
        @test [x.level for x in s.chain] == [:series, :chartspace]
        @test XLSX.Charts.getLabelTextProp(c, 1, :bold).value === true
        @test XLSX.Charts.getLabelTextProp(c, 1, :fill).value.fgcolor.val == "7030A0"
        @test_throws XLSX.XLSXError XLSX.Charts.getLabelTextProp(c, 1, :nonsense)

        # Title: formatting on the run, not defRPr.
        tp = XLSX.Charts.default_run_props(XLSX.Charts.getChartTitleTextProps(c))
        @test tp.size ≈ 18.0
        @test tp.latin == "Comic Sans MS"

        # Legend: no run; defRPr sets only the colour, size is on endParaRPr.
        lg = XLSX.Charts.default_run_props(XLSX.Charts.getLegendTextProps(c))
        @test lg.fill.fgcolor.val == "accent3"
        @test isnothing(lg.size)

        # Nothing written anywhere: no value, no site, but a full chain.
        g = XLSX.readxlsx(joinpath(data_directory, "chartex_layouts.xlsx"))
        w = only(filter(x -> x isa XLSX.Charts.ChartEx && x.sheet == "waterfall", XLSX.Charts.getCharts(g)))
        u = XLSX.Charts.getSeriesFill(w, 1)
        @test isnothing(u.value) && isnothing(u.site)
        @test length(u.chain) == 1
        @test isnothing(only(u.chain).props)
        @test isnothing(XLSX.Charts.getLabelTextProp(w, 1, :size).value)

    end

    @testset "setSeriesFill" begin
        tmp = joinpath(mktempdir(), "cx_fill.xlsx")
        cp(joinpath(data_directory, "chartex_formatted.xlsx"), tmp)
        xf = XLSX.openxlsx(tmp; mode="rw")
        c = only(filter(x -> x isa XLSX.Charts.ChartEx, XLSX.Charts.getCharts(xf)))

        # series fill, replacing a theme colour
        @test XLSX.Charts.setSeriesFill(c, 1, "red") === c
        @test XLSX.Charts.getSeriesFill(c, 1).value.fgcolor.rgb == "FF0000"   # same handle, still valid

        # an existing dataPt (point 5, idx 4) is edited in place
        XLSX.Charts.setSeriesFill(c, 1, XLSX.Charts.SchemeColor(:accent1); point=5)
        @test XLSX.Charts.getSeriesFill(c, 1; point=5).value.fgcolor.val == "accent1"

        # new dataPts: one before idx 4, one after, kept in idx order
        XLSX.Charts.setSeriesFill(c, 1, "blue"; point=2)
        XLSX.Charts.setSeriesFill(c, 1, "green"; point=7)
        ser = XLSX.Charts._cx_series_node(c, 1)
        @test [XLSX.Charts._cx_idx(n) for n in XLSX.elements_with_tag(ser, "dataPt")] == [1, 4, 6]
        @test XLSX.Charts.getSeriesFill(c, 1; point=2).site.level == :point

        # child order of the series follows the schema
        @test [String(XLSX.localname(n)) for n in XLSX.XML.eachelement(ser)] ==
            ["spPr", "dataPt", "dataPt", "dataPt", "dataLabels", "dataId", "layoutPr"]

        # :none is an explicit setting; :inherit on an unformatted point creates nothing
        XLSX.Charts.setSeriesFill(c, 1, :none; point=2)
        @test XLSX.Charts.getSeriesFill(c, 1; point=2).value.kind === :none
        XLSX.Charts.setSeriesFill(c, 1, :inherit; point=3)
        @test length(XLSX.elements_with_tag(XLSX.Charts._cx_series_node(c, 1), "dataPt")) == 3

        @test_throws XLSX.XLSXError XLSX.Charts.setSeriesFill(c, 1, "red"; point=0)
        @test_throws XLSX.XLSXError XLSX.Charts.setSeriesFill(c, 2, "red")   # only one series

        @test all(n -> XLSX.XML.tag(XLSX.first_element_with_tag(n, "spPr")) == "cx:spPr",
            XLSX.elements_with_tag(XLSX.Charts._cx_series_node(c, 1), "dataPt"))

        # the changes reach disk
        out = joinpath(mktempdir(), "cx_fill_out.xlsx")
        XLSX.writexlsx(out, xf, overwrite=true)
        f = XLSX.readxlsx(out)
        d = only(filter(x -> x isa XLSX.Charts.ChartEx, XLSX.Charts.getCharts(f)))
        @test XLSX.Charts.getSeriesFill(d, 1).value.fgcolor.rgb == "FF0000"
        @test XLSX.Charts.getSeriesFill(d, 1; point=7).value.fgcolor.rgb == "008000"
        @test XLSX.Charts.getSeriesFill(d, 1; point=2).value.kind === :none
        SAVE_FILES && save_outfile(xf)
    end

    @testset "setSeriesFill :inherit writes nothing where nothing was set" begin
        tmp = joinpath(mktempdir(), "cx_inherit.xlsx")
        cp(joinpath(data_directory, "chartex_layouts.xlsx"), tmp)
        xf = XLSX.openxlsx(tmp; mode="rw")
        w = only(filter(x -> x isa XLSX.Charts.ChartEx && x.sheet == "waterfall", XLSX.Charts.getCharts(xf)))
        XLSX.Charts.setSeriesFill(w, 1, :inherit)
        @test isnothing(XLSX.first_element_with_tag(XLSX.Charts._cx_series_node(w, 1), "spPr"))
        SAVE_FILES && save_outfile(xf)
    end

    @testset "setSeriesLine" begin
        tmp = joinpath(mktempdir(), "cx_line.xlsx")
        cp(joinpath(data_directory, "chartex_formatted.xlsx"), tmp)
        xf = XLSX.openxlsx(tmp; mode="rw")
        c = only(filter(x -> x isa XLSX.Charts.ChartEx, XLSX.Charts.getCharts(xf)))

        # several properties at once, on a series with no a:ln
        XLSX.Charts.setSeriesLine(c, 1; color="FFC00000", width=2.25, dash=:dash, cap=:round)
        l = XLSX.Charts.getSeriesLine(c, 1).value
        @test l.fill.fgcolor.rgb == "C00000"
        @test l.width ≈ 2.25
        @test l.dash == "dash"
        @test l.cap == "rnd"

        # an Excel UI alias resolves to the DrawingML spelling
        XLSX.Charts.setSeriesLineDash(c, 1, :roundDot)
        @test XLSX.Charts.getSeriesLine(c, 1).value.dash == "sysDot"

        # one property back to inherit, leaving the rest
        XLSX.Charts.setSeriesLine(c, 1; width=:inherit)
        l = XLSX.Charts.getSeriesLine(c, 1).value
        @test isnothing(l.width) && l.fill.fgcolor.rgb == "C00000"

        # a created point gets its own outline
        XLSX.Charts.setSeriesLine(c, 1; point=2, color="FF0070C0", width=1.5)
        @test XLSX.Charts.getSeriesLine(c, 1; point=2).site.level == :point
        dp = only(filter(n -> XLSX.Charts._cx_idx(n) == 1,
            XLSX.elements_with_tag(XLSX.Charts._cx_series_node(c, 1), "dataPt")))
        @test XLSX.XML.tag(XLSX.first_element_with_tag(dp, "spPr")) == "cx:spPr"

        # :none is explicit, :inherit removes the a:ln
        XLSX.Charts.setSeriesLine(c, 1, :none)
        @test XLSX.Charts.getSeriesLine(c, 1).value.fill.kind === :none
        XLSX.Charts.setSeriesLine(c, 1, :inherit)
        @test isnothing(XLSX.first_element_with_tag(
            XLSX.first_element_with_tag(XLSX.Charts._cx_series_node(c, 1), "spPr"), "ln"))

        # no keywords is a no-op; a bad symbol throws
        @test XLSX.Charts.setSeriesLine(c, 1) === c
        @test_throws XLSX.XLSXError XLSX.Charts.setSeriesLine(c, 1, :dotted)

        # reaches disk
        out = joinpath(mktempdir(), "cx_line_out.xlsx")
        XLSX.writexlsx(out, xf, overwrite=true)
        d = only(filter(x -> x isa XLSX.Charts.ChartEx, XLSX.Charts.getCharts(XLSX.readxlsx(out))))
        @test XLSX.Charts.getSeriesLine(d, 1; point=2).value.width ≈ 1.5
        SAVE_FILES && save_outfile(xf)
    end

    @testset "setLabelTextProp" begin
        tmp = joinpath(mktempdir(), "cx_label.xlsx")
        cp(joinpath(data_directory, "chartex_layouts.xlsx"), tmp)
        xf = XLSX.openxlsx(tmp; mode="rw")
        # the funnel has no cx:dataLabels at all
        c = only(filter(x -> x isa XLSX.Charts.ChartEx && x.sheet == "funnel", XLSX.Charts.getCharts(xf)))

        @test isnothing(XLSX.Charts.getLabelTextProp(c, 1, :size).value)
        XLSX.Charts.setLabelTextProp(c, 1, :size, 11)
        @test XLSX.Charts.getLabelTextProp(c, 1, :size).value ≈ 11.0

        # a created cx:dataLabels must not turn labels on
        lbl = XLSX.first_element_with_tag(XLSX.Charts._cx_series_node(c, 1), "dataLabels")
        @test XLSX.XML.tag(lbl) == "cx:dataLabels"
        @test [String(XLSX.localname(n)) for n in XLSX.XML.eachelement(lbl)] == ["txPr", "visibility"]
        @test all(f -> XLSX.Charts.getLabelFlag(c, 1, f) === false,
            (:seriesName, :categoryName, :value))
        @test XLSX.XML.tag(XLSX.first_element_with_tag(lbl, "txPr")) == "cx:txPr"

        # more fields, and a composite one
        XLSX.Charts.setLabelTextProp(c, 1, :bold, true)
        XLSX.Charts.setLabelTextProp(c, 1, :fill, "FF7030A0")
        @test XLSX.Charts.getLabelTextProp(c, 1, :bold).value === true
        @test XLSX.Charts.getLabelTextProp(c, 1, :fill).value.fgcolor.rgb == "7030A0"
        @test XLSX.Charts.getLabelTextProp(c, 1, :size).value ≈ 11.0     # earlier field kept

        XLSX.Charts.setLabelTextProp(c, 1, :size, :inherit)
        @test isnothing(XLSX.Charts.getLabelTextProp(c, 1, :size).value)
        @test_throws XLSX.XLSXError XLSX.Charts.setLabelTextProp(c, 1, :nonsense, 1)

        # an existing cx:dataLabels keeps the flags Excel wrote
        w = only(filter(x -> x isa XLSX.Charts.ChartEx && x.sheet == "waterfall", XLSX.Charts.getCharts(xf)))
        XLSX.Charts.setLabelTextProp(w, 1, :size, 9)
        @test XLSX.Charts.getLabelFlag(w, 1, :value) === true
        @test XLSX.Charts.getLabelTextProp(w, 1, :size).value ≈ 9.0

        out = joinpath(mktempdir(), "cx_label_out.xlsx")
        XLSX.writexlsx(out, xf, overwrite=true)
        f = XLSX.readxlsx(out)
        d = only(filter(x -> x isa XLSX.Charts.ChartEx && x.sheet == "funnel", XLSX.Charts.getCharts(f)))
        @test XLSX.Charts.getLabelTextProp(d, 1, :bold).value === true
        SAVE_FILES && save_outfile(xf)
    end

    @testset "title, legend and axis text" begin
        tmp = joinpath(mktempdir(), "cx_text.xlsx")
        cp(joinpath(data_directory, "chartex_layouts.xlsx"), tmp)
        xf = XLSX.openxlsx(tmp; mode="rw")
        w = only(filter(x -> x isa XLSX.Charts.ChartEx && x.sheet == "waterfall", XLSX.Charts.getCharts(xf)))

        # title: cx:title exists but has no txPr
        XLSX.Charts.setChartTitleTextProp(w, :size, 16)
        XLSX.Charts.setChartTitleTextProp(w, :fill, "FF7030A0")
        tp = XLSX.Charts.default_run_props(XLSX.Charts.getChartTitleTextProps(w))
        @test tp.size ≈ 16.0 && tp.fill.fgcolor.rgb == "7030A0"
        @test XLSX.XML.tag(XLSX.first_element_with_tag(
            XLSX.first_element_with_tag(XLSX.Charts._cx_chart(w), "title"), "txPr")) == "cx:txPr"

        # legend
        XLSX.Charts.setLegendTextProp(w, :bold, true)
        @test XLSX.Charts.default_run_props(XLSX.Charts.getLegendTextProps(w)).bold === true

        # axes: by id, not position
        @test XLSX.Charts.getChartAxisIds(w) == [0, 1]
        XLSX.Charts.setAxisTitleTextProp(w, 1, :size, 12)
        @test XLSX.Charts.default_run_props(XLSX.Charts.getAxisTitleTextProps(w, 1)).size ≈ 12.0
        @test isnothing(XLSX.Charts.getAxisTitleTextProps(w, 0))          # untouched axis
        @test_throws XLSX.XLSXError XLSX.Charts.setAxisTitleTextProp(w, 7, :size, 12)

        # the legend is created where there is none (the funnel has no cx:legend)
        fn = only(filter(x -> x isa XLSX.Charts.ChartEx && x.sheet == "funnel", XLSX.Charts.getCharts(xf)))
        XLSX.Charts.setLegendTextProp(fn, :size, 9)
        lg = XLSX.first_element_with_tag(XLSX.Charts._cx_chart(fn), "legend")
        @test XLSX.XML.tag(lg) == "cx:legend"
        @test XLSX.Charts.default_run_props(XLSX.Charts.getLegendTextProps(fn)).size ≈ 9.0

        out = joinpath(mktempdir(), "cx_text_out.xlsx")
        XLSX.writexlsx(out, xf, overwrite=true)
        d = only(filter(x -> x isa XLSX.Charts.ChartEx && x.sheet == "waterfall",
            XLSX.Charts.getCharts(XLSX.readxlsx(out))))
        @test XLSX.Charts.default_run_props(XLSX.Charts.getChartTitleTextProps(d)).size ≈ 16.0
         
        x = XLSX.getCharts(XLSX.readxlsx(joinpath(data_directory, "chartex_kinds.xlsx")))[1]
        @test_throws XLSX.XLSXError XLSX.Charts.getAxisTitleTextProps(x, 999)     # no such axis id
        SAVE_FILES && save_outfile(xf)
    
    end

    @testset "setChartTitleText and setSeriesName" begin
        tmp = joinpath(mktempdir(), "cx_text_set.xlsx")
        cp(joinpath(data_directory, "chartex_layouts.xlsx"), tmp)
        xf = XLSX.openxlsx(tmp; mode="rw")

        # histogram: a typed title, with the text also in the txPr runs
        h = only(filter(x -> x isa XLSX.Charts.ChartEx && x.sheet == "histogram", XLSX.Charts.getCharts(xf)))
        @test XLSX.Charts.getChartTitle(h) == "Chart Title"
        XLSX.Charts.setChartTitleText(h, "Distribution of values")
        @test XLSX.Charts.getChartTitle(h) == "Distribution of values"
        @test XLSX.Charts.text_content(XLSX.Charts.getChartTitleTextProps(h)) == "Distribution of values"
        @test XLSX.Charts.default_run_props(XLSX.Charts.getChartTitleTextProps(h)).size ≈ 14.0  # formatting kept

        # bound: typing replaces the binding
        b = only(filter(x -> x isa XLSX.Charts.ChartEx && x.sheet == "bound", XLSX.Charts.getCharts(xf)))
        @test !isnothing(XLSX.Charts.getChartTitleRange(b))
        XLSX.Charts.setChartTitleText(b, "Typed over")
        @test XLSX.Charts.getChartTitle(b) == "Typed over"
        @test isnothing(XLSX.Charts.getChartTitleRange(b))

        # waterfall: no title text at all, so this creates it
        w = only(filter(x -> x isa XLSX.Charts.ChartEx && x.sheet == "waterfall", XLSX.Charts.getCharts(xf)))
        @test isnothing(XLSX.Charts.getChartTitle(w))
        XLSX.Charts.setChartTitleText(w, "Cash flow")
        @test XLSX.Charts.getChartTitle(w) == "Cash flow"

        # series names
        @test XLSX.Charts.getSeriesName(b, 1) == "Series name"
        XLSX.Charts.setSeriesName(b, 1, "Renamed")
        @test XLSX.Charts.getSeriesName(b, 1) == "Renamed"
        @test isnothing(XLSX.Charts.getSeriesNameRange(b, 1))
        @test isnothing(XLSX.Charts.getSeriesName(w, 1))
        XLSX.Charts.setSeriesName(w, 1, "Movement")
        @test XLSX.Charts.getSeriesName(w, 1) == "Movement"

        out = joinpath(mktempdir(), "cx_text_set_out.xlsx")
        XLSX.writexlsx(out, xf, overwrite=true)
        f = XLSX.readxlsx(out)
        d = only(filter(x -> x isa XLSX.Charts.ChartEx && x.sheet == "histogram", XLSX.Charts.getCharts(f)))
        @test XLSX.Charts.getChartTitle(d) == "Distribution of values"
        SAVE_FILES && save_outfile(xf)
    end

    @testset "setSeriesSubtotals" begin
        tmp = joinpath(mktempdir(), "cx_subtotals.xlsx")
        cp(joinpath(data_directory, "chartex_layouts.xlsx"), tmp)
        xf = XLSX.openxlsx(tmp; mode="rw")
        w = only(filter(x -> x isa XLSX.Charts.ChartEx && x.sheet == "waterfall", XLSX.Charts.getCharts(xf)))

        @test XLSX.Charts.getSeriesSubtotals(w, 1) == [2, 5]     # what Excel wrote
        XLSX.Charts.setSeriesSubtotals(w, 1, [5, 3])
        @test XLSX.Charts.getSeriesSubtotals(w, 1) == [3, 5]     # sorted, replaced not appended
        XLSX.Charts.setSeriesSubtotals(w, 1, Int[])
        @test XLSX.Charts.getSeriesSubtotals(w, 1) == Int[]
        @test !isnothing(XLSX.first_element_with_tag(
            XLSX.Charts._cx_layoutpr(w, 1), "subtotals"))  # present but empty
        XLSX.Charts.setSeriesSubtotals(w, 1, :inherit)
        @test isnothing(XLSX.Charts.getSeriesSubtotals(w, 1))
        @test_throws XLSX.XLSXError XLSX.Charts.setSeriesSubtotals(w, 1, [0, 2])

        # a series with no cx:layoutPr at all: the funnel
        fn = only(filter(x -> x isa XLSX.Charts.ChartEx && x.sheet == "funnel", XLSX.Charts.getCharts(xf)))
        @test isnothing(XLSX.Charts._cx_layoutpr(fn, 1))
        XLSX.Charts.setSeriesSubtotals(fn, 1, [2])
        @test XLSX.Charts.getSeriesSubtotals(fn, 1) == [2]
        @test XLSX.XML.tag(XLSX.Charts._cx_layoutpr(fn, 1)) == "cx:layoutPr"

        out = joinpath(mktempdir(), "cx_subtotals_out.xlsx")
        XLSX.writexlsx(out, xf, overwrite=true)
        d = only(filter(x -> x isa XLSX.Charts.ChartEx && x.sheet == "waterfall",
            XLSX.Charts.getCharts(XLSX.readxlsx(out))))
        @test isnothing(XLSX.Charts.getSeriesSubtotals(d, 1))
        SAVE_FILES && save_outfile(xf)
    end

    @testset "layoutPr setters" begin
        tmp = joinpath(mktempdir(), "cx_layoutpr.xlsx")
        cp(joinpath(data_directory, "chartex_layouts.xlsx"), tmp)
        xf = XLSX.openxlsx(tmp; mode="rw")
        chart(sheet) = only(filter(x -> x isa XLSX.Charts.ChartEx && x.sheet == sheet,
            XLSX.Charts.getCharts(xf)))

        @testset "quartile method and parent label layout" begin
            b = chart("boxwhisker")
            @test XLSX.Charts.getSeriesQuartileMethod(b, 1) == :exclusive
            XLSX.Charts.setSeriesQuartileMethod(b, 1, :inclusive)
            @test XLSX.Charts.getSeriesQuartileMethod(b, 1) == :inclusive
            XLSX.Charts.setSeriesQuartileMethod(b, 1, :inherit)
            @test isnothing(XLSX.Charts.getSeriesQuartileMethod(b, 1))
            @test_throws XLSX.XLSXError XLSX.Charts.setSeriesQuartileMethod(b, 1, :median)

            t = chart("treemap")
            @test XLSX.Charts.getSeriesParentLabelLayout(t, 1) == :overlapping
            XLSX.Charts.setSeriesParentLabelLayout(t, 1, :banner)
            @test XLSX.Charts.getSeriesParentLabelLayout(t, 1) == :banner
            @test_throws XLSX.XLSXError XLSX.Charts.setSeriesParentLabelLayout(t, 1, :sideways)

            # a series with no cx:layoutPr: the sunburst
            s = chart("sunburst")
            @test isnothing(XLSX.Charts._cx_layoutpr(s, 1))
            XLSX.Charts.setSeriesParentLabelLayout(s, 1, :none)
            @test XLSX.Charts.getSeriesParentLabelLayout(s, 1) == :none
        end

        @testset "layout flags" begin
            b = chart("boxwhisker")
            @test XLSX.Charts.getSeriesLayoutFlag(b, 2, :outliers) === false
            XLSX.Charts.setSeriesLayoutFlag(b, 2, :outliers, true)
            @test XLSX.Charts.getSeriesLayoutFlag(b, 2, :outliers) === true
            # the other attributes on the same cx:visibility are untouched
            @test XLSX.Charts.getSeriesLayoutFlag(b, 2, :meanMarker) === false
            XLSX.Charts.setSeriesLayoutFlag(b, 2, :outliers, :inherit)
            @test isnothing(XLSX.Charts.getSeriesLayoutFlag(b, 2, :outliers))
            @test XLSX.Charts.getSeriesLayoutFlag(b, 2, :meanMarker) === false

            # a waterfall flag on a series whose cx:visibility does not exist
            w = chart("waterfall")
            @test isnothing(XLSX.Charts.getSeriesLayoutFlag(w, 1, :connectorLines))
            XLSX.Charts.setSeriesLayoutFlag(w, 1, :connectorLines, false)
            @test XLSX.Charts.getSeriesLayoutFlag(w, 1, :connectorLines) === false
            # removing the only attribute removes the element
            XLSX.Charts.setSeriesLayoutFlag(w, 1, :connectorLines, :inherit)
            @test isnothing(XLSX.first_element_with_tag(XLSX.Charts._cx_layoutpr(w, 1), "visibility"))

            @test_throws XLSX.XLSXError XLSX.Charts.setSeriesLayoutFlag(w, 1, :nonsense, true)
            @test_throws XLSX.XLSXError XLSX.Charts.setSeriesLayoutFlag(w, 1, :connectorLines, 1)
        end

        @testset "binning and aggregation" begin
            h = chart("histogram")
            @test XLSX.Charts.getSeriesBinning(h, 1) ==
                XLSX.Charts.ChartExBinning(:r, 0.0, 100.0, 10.0, nothing)

            XLSX.Charts.setSeriesBinning(h, 1; binSize=20, overflow=80, intervalClosed=:l)
            @test XLSX.Charts.getSeriesBinning(h, 1) ==
                XLSX.Charts.ChartExBinning(:l, 0.0, 80.0, 20.0, nothing)

            # binCount replaces binSize
            XLSX.Charts.setSeriesBinning(h, 1; binCount=7)
            b = XLSX.Charts.getSeriesBinning(h, 1)
            @test isnothing(b.binSize) && b.binCount == 7

            XLSX.Charts.setSeriesBinning(h, 1; underflow=:auto, overflow=:inherit)
            b = XLSX.Charts.getSeriesBinning(h, 1)
            @test b.underflow === :auto && isnothing(b.overflow)

            @test XLSX.Charts.setSeriesBinning(h, 1) === h        # no keywords is a no-op
            @test_throws XLSX.XLSXError XLSX.Charts.setSeriesBinning(h, 1; binSize=5, binCount=5)
            @test_throws XLSX.XLSXError XLSX.Charts.setSeriesBinning(h, 1; intervalClosed=:x)

            # aggregation and binning exclude each other
            p = chart("pareto")
            @test XLSX.Charts.getSeriesAggregation(p, 1) === true
            XLSX.Charts.setSeriesBinning(p, 1; binSize=5)
            @test XLSX.Charts.getSeriesAggregation(p, 1) === false
            XLSX.Charts.setSeriesAggregation(p, 1, true)
            @test isnothing(XLSX.Charts.getSeriesBinning(p, 1))
            XLSX.Charts.setSeriesAggregation(p, 1, false)
            @test XLSX.Charts.getSeriesAggregation(p, 1) === false
        end

        @testset "reaches disk" begin
            out = joinpath(mktempdir(), "cx_layoutpr_out.xlsx")
            XLSX.writexlsx(out, xf, overwrite=true)
            f = XLSX.readxlsx(out)
            h = only(filter(x -> x isa XLSX.Charts.ChartEx && x.sheet == "histogram",
                XLSX.Charts.getCharts(f)))
            @test XLSX.Charts.getSeriesBinning(h, 1).binCount == 7
        end
        SAVE_FILES && save_outfile(xf)
    end
    @testset "every line setter reaches a ChartEx" begin
        xf = XLSX.openxlsx(joinpath(data_directory, "chartex_kinds.xlsx"); mode="rw")
        c = XLSX.getCharts(xf)[1]
        XLSX.Charts.setSeriesLineCap(c, 1, :round)
        XLSX.Charts.setSeriesLineCompound(c, 1, :double)
        XLSX.Charts.setSeriesLineJoin(c, 1, :miter)
        XLSX.Charts.setSeriesLineMiterLimit(c, 1, 8)
        ln = XLSX.Charts.getSeriesLine(c, 1).value
        @test ln.cap == "rnd"
        @test ln.compound == "dbl"
        @test ln.join == "miter"
        @test ln.miter_limit ≈ 8.0

        XLSX.Charts.setSeriesFill(c, 1, :red)                       # a colour name as a Symbol
        @test XLSX.Charts.getSeriesFill(c, 1).value.fgcolor.rgb == "FF0000"
        XLSX.Charts.setSeriesFill(c, 1, :inherit)
        @test isnothing(XLSX.Charts.getSeriesFill(c, 1).value)
        SAVE_FILES && save_outfile(xf)
    end

end