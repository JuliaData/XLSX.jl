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
        @test [s.level for s in e5.chain] == [:point, :series, :style]    # style rung last, not reached
        @test e5.value.fgcolor.kind == :srgb
        @test e5.value.fgcolor.val == "4EA72E"                                  # as written
        @test e5.value.fgcolor.transforms == [:lumMod => 60000, :lumOff => 40000]
        @test e5.value.fgcolor.rgb == "8ED973"                                  # transforms applied

        # Point 1 has no dataPt: no point rung, resolves at the series.
        e1 = XLSX.Charts.getSeriesFill(c, 1; point=1)
        @test e1.site.level == :series
        @test [s.level for s in e1.chain] == [:series, :style]
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

    # The expected colours are the fills in Excel 365's SVG exports of these
    # charts, test/data/chart_style_svg/cx_*.svg, under the workbook's theme
    # (accent1 156082, accent2 E97132, accent3 196B24, ...).
    @testset "style rung: the colours Excel draws" begin
        f  = XLSX.readxlsx(joinpath(data_directory, "chartex_colors.xlsx"))
        cs = Dict(basename(c.path) => c for c in XLSX.Charts.getCharts(f) if c isa XLSX.Charts.ChartEx)
        rgb(c, i = 1; kw...) = (e = XLSX.Charts.getSeriesFill(c, i; kw...);
                                isnothing(e.value) ? nothing : e.value.fgcolor.rgb)
        cycle8  = ["156082", "E97132", "196B24", "0F9ED5", "A02B93", "4EA72E", "0D3A4E", "994010"]
        regions = ["156082", "156082", "156082", "E97132", "E97132", "196B24", "196B24", "196B24"]

        # cx_1: Start, Sales, Returns, Costs, Grants, End (set as total)
        w = cs["chartEx1.xml"]
        @test [rgb(w; point = p) for p in 1:6] == ["156082", "156082", "E97132", "E97132", "156082", "196B24"]
        @test isnothing(rgb(w))                         # increase, decrease and total differ
        @test isnothing(rgb(w; point = 7))              # beyond the data
        @test rgb(cs["chartEx2.xml"]) == "156082"                                   # cx_2 funnel
        @test [rgb(cs["chartEx3.xml"]; point = p) for p in 1:8] == cycle8           # cx_3 treemap, A..H
        @test [rgb(cs["chartEx4.xml"]; point = p) for p in 1:8] == regions          # cx_4 sunburst, by region
        @test rgb(cs["chartEx5.xml"]) == "156082"                                   # cx_5 histogram
        @test rgb(cs["chartEx6.xml"]) == "156082"                                   # cx_6 Pareto bars
        @test isnothing(rgb(cs["chartEx6.xml"]; pareto = true))                     # a line has no fill
        @test [rgb(cs["chartEx7.xml"], i) for i in 1:8] == cycle8                   # cx_7 box & whisker
        @test [rgb(cs["chartEx8.xml"]; point = p) for p in 1:8] == regions          # cx_8 two-level treemap

        # The rung: a :style site on the style's dataPoint entry, with the colour
        # as the theme colour and Excel's variation.
        e = XLSX.Charts.getSeriesFill(cs["chartEx7.xml"], 7)
        @test e.site.level === :style && e.site.kind === :shape
        @test XLSX.localname(e.site.container) == "dataPoint"
        @test [s.level for s in e.chain] == [:series, :style]
        @test e.value.kind === :solid
        @test e.value.fgcolor.val == "accent1"
        @test e.value.fgcolor.transforms == [:lumMod => 60000]

        # Lines and text get no style rung.
        l = XLSX.Charts.getSeriesLine(cs["chartEx7.xml"], 1)
        @test isnothing(l.value) && [s.level for s in l.chain] == [:series]
        @test all(s -> s.level !== :style, XLSX.Charts.getLabelTextProp(w, 1, :size).chain)
    end

    @testset "style rung: no answer where Excel was not checked" begin
        nostyle(e) = isnothing(e.value) && all(s -> s.level !== :style, e.chain)

        # Monochromatic Palette 1 (withinLinear) is not reproduced.
        m = XLSX.readxlsx(joinpath(data_directory, "chartex_mono.xlsx"))
        @test all(c -> nostyle(XLSX.Charts.getSeriesFill(c, 1; point = 1)), XLSX.Charts.getCharts(m))

        xf  = XLSX.openxlsx(joinpath(data_directory, "chartex_colors.xlsx"); mode = "rw")
        box = only(filter(c -> basename(c.path) == "chartEx7.xml", XLSX.Charts.getCharts(xf)))

        # A series carrying formatIdx, which no fixture has.
        ser = XLSX.Charts._cx_series_node(box, 2)
        ser["formatIdx"] = "5"
        @test nostyle(XLSX.Charts.getSeriesFill(box, 2))
        @test XLSX.Charts.getSeriesFill(box, 1).value.fgcolor.rgb == "156082"   # others still answer

        # Without its style part the chart ends its cascade where it did before.
        rels = xf.data["xl/charts/_rels/chartEx7.xml.rels"]
        kids = XLSX.XML.children(XLSX.xml_root_element(rels))
        filter!(n -> XLSX.XML.nodetype(n) != XLSX.XML.Element ||
                     XLSX.get_attr(n, "Type") != XLSX.Charts.REL_CHART_STYLE, kids)
        @test nostyle(XLSX.Charts.getSeriesFill(box, 1))
    end

    @testset "style rung: charts made by addChartEx" begin
        xf = XLSX.newxlsx()
        ws = xf[1]
        ws["A1"] = ["Item" "Amount" "s1" "s2" "s3"
                    "a"     3       1    4    7
                    "b"    -1       2    5    8
                    "c"     2       3    6    9]
        w = XLSX.Charts.addChartEx(ws, :waterfall, "B2:B4"; anchor = "G2:N16", categories = "A2:A4")
        @test [XLSX.Charts.getSeriesFill(w, 1; point = p).value.fgcolor.val for p in 1:3] ==
              ["accent1", "accent2", "accent1"]
        b = XLSX.Charts.addChartEx(ws, :boxWhisker, "C2:C4"; anchor = "G20:N34")
        XLSX.Charts.addSeries(b, "D2:D4")
        XLSX.Charts.addSeries(b, "E2:E4")
        @test [XLSX.Charts.getSeriesFill(b, i).value.fgcolor.val for i in 1:3] == ["accent1", "accent2", "accent3"]
        XLSX.Charts.setSeriesFill(b, 2, "red")                              # an explicit fill still wins
        @test XLSX.Charts.getSeriesFill(b, 2).site.level === :series
        SAVE_FILES && save_outfile(xf)
    end

    @testset "setSeriesFill" begin
        tmp = "cx_fill.xlsx"
        cp(joinpath(data_directory, "chartex_formatted.xlsx"), tmp; force=true)
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
        out = "cx_fill_out.xlsx"
        XLSX.writexlsx(out, xf, overwrite=true)
        f = XLSX.readxlsx(out)
        d = only(filter(x -> x isa XLSX.Charts.ChartEx, XLSX.Charts.getCharts(f)))
        @test XLSX.Charts.getSeriesFill(d, 1).value.fgcolor.rgb == "FF0000"
        @test XLSX.Charts.getSeriesFill(d, 1; point=7).value.fgcolor.rgb == "008000"
        @test XLSX.Charts.getSeriesFill(d, 1; point=2).value.kind === :none
        SAVE_FILES && save_outfile(xf)
        isfile(out) && rm(out)
        isfile(tmp) && rm(tmp)
    end

    @testset "setSeriesFill :inherit writes nothing where nothing was set" begin
        tmp = "cx_inherit.xlsx"
        cp(joinpath(data_directory, "chartex_layouts.xlsx"), tmp; force=true)
        xf = XLSX.openxlsx(tmp; mode="rw")
        w = only(filter(x -> x isa XLSX.Charts.ChartEx && x.sheet == "waterfall", XLSX.Charts.getCharts(xf)))
        XLSX.Charts.setSeriesFill(w, 1, :inherit)
        @test isnothing(XLSX.first_element_with_tag(XLSX.Charts._cx_series_node(w, 1), "spPr"))
        SAVE_FILES && save_outfile(xf)
        isfile(tmp) && rm(tmp)
    end

    @testset "setSeriesLine" begin
        tmp = "cx_line.xlsx"
        cp(joinpath(data_directory, "chartex_formatted.xlsx"), tmp; force=true)
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
        out = "cx_line_out.xlsx"
        XLSX.writexlsx(out, xf, overwrite=true)
        d = only(filter(x -> x isa XLSX.Charts.ChartEx, XLSX.Charts.getCharts(XLSX.readxlsx(out))))
        @test XLSX.Charts.getSeriesLine(d, 1; point=2).value.width ≈ 1.5
        SAVE_FILES && save_outfile(xf)
        isfile(out) && rm(out)
        isfile(tmp) && rm(tmp)
    end

    @testset "every fill and line setter reaches a ChartEx" begin
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

    @testset "setLabelTextProp" begin
        tmp = "cx_label.xlsx"
        cp(joinpath(data_directory, "chartex_layouts.xlsx"), tmp; force=true)
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

        # Excel's names for fields work here as they do on a Chart
        XLSX.Charts.setLabelTextProp(c, 1, :font, "Arial")
        XLSX.Charts.setLabelTextProp(c, 1, :color, "accent3")
        XLSX.Charts.setLabelTextProp(c, 1, :underline, :dbl)
        @test XLSX.Charts.getLabelTextProp(c, 1, :latin).value == "Arial"
        @test XLSX.Charts.getLabelTextProp(c, 1, :fill).value.fgcolor.val == "accent3"
        @test XLSX.Charts.getLabelTextProp(c, 1, :under).value == "dbl"

        # an existing cx:dataLabels keeps the flags Excel wrote
        w = only(filter(x -> x isa XLSX.Charts.ChartEx && x.sheet == "waterfall", XLSX.Charts.getCharts(xf)))
        XLSX.Charts.setLabelTextProp(w, 1, :size, 9)
        @test XLSX.Charts.getLabelFlag(w, 1, :value) === true
        @test XLSX.Charts.getLabelTextProp(w, 1, :size).value ≈ 9.0

        out = "cx_label_out.xlsx"
        XLSX.writexlsx(out, xf, overwrite=true)
        f = XLSX.readxlsx(out)
        d = only(filter(x -> x isa XLSX.Charts.ChartEx && x.sheet == "funnel", XLSX.Charts.getCharts(f)))
        @test XLSX.Charts.getLabelTextProp(d, 1, :bold).value === true
        SAVE_FILES && save_outfile(xf)
        isfile(out) && rm(out)
        isfile(tmp) && rm(tmp)
    end

    @testset "title, legend and axis text" begin
        tmp = "cx_text.xlsx"
        cp(joinpath(data_directory, "chartex_layouts.xlsx"), tmp; force=true)
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

        # legend position: cx:legend's pos and align attributes, and overlay
        @test XLSX.Charts.getLegendPos(w) === :t                          # pos="t" align="ctr"
        @test XLSX.Charts.getLegendOverlay(w) === false
        for (p, pa, al) in ((:b, "b", "ctr"), (:t, "t", "ctr"), (:l, "l", "ctr"),
                            (:r, "r", "ctr"), (:tr, "r", "min"))
            XLSX.Charts.setLegendPos(w, p)
            @test XLSX.Charts.getLegendPos(w) === p
            lg = XLSX.first_element_with_tag(XLSX.Charts._cx_chart(w), "legend")
            @test XLSX.get_attr(lg, "pos") == pa && XLSX.get_attr(lg, "align") == al
        end
        XLSX.Charts.setLegendPos(w, :bottom; overlay=true)
        @test XLSX.Charts.getLegendPos(w) === :b
        @test XLSX.Charts.getLegendOverlay(w) === true
        @test_throws XLSX.XLSXError XLSX.Charts.setLegendPos(w, :br)
        XLSX.Charts.setLegendOverlay(w, :inherit)
        @test isnothing(XLSX.Charts.getLegendOverlay(w))
        XLSX.Charts.setLegendPos(w, :inherit)
        @test isnothing(XLSX.Charts.getLegendPos(w))
        XLSX.Charts.setLegendPos(w, :topRight; overlay=true)
        @test XLSX.Charts.default_run_props(XLSX.Charts.getLegendTextProps(w)).bold === true  # txPr kept
        XLSX.Charts.setLegendPos(fn, :l)                                   # funnel's created legend
        @test XLSX.Charts.getLegendPos(fn) === :l

        out = "cx_text_out.xlsx"
        XLSX.writexlsx(out, xf, overwrite=true)
        d = only(filter(x -> x isa XLSX.Charts.ChartEx && x.sheet == "waterfall",
            XLSX.Charts.getCharts(XLSX.readxlsx(out))))
        @test XLSX.Charts.default_run_props(XLSX.Charts.getChartTitleTextProps(d)).size ≈ 16.0
        @test XLSX.Charts.getLegendPos(d) === :tr
        @test XLSX.Charts.getLegendOverlay(d) === true
         
        x = XLSX.getCharts(XLSX.readxlsx(joinpath(data_directory, "chartex_kinds.xlsx")))[1]
        @test_throws XLSX.XLSXError XLSX.Charts.getAxisTitleTextProps(x, 999)     # no such axis id
        SAVE_FILES && save_outfile(xf)
    
        isfile(out) && rm(out)
        isfile(tmp) && rm(tmp)
    end

    @testset "legend positions as Excel writes them" begin
        # Made by Excel 16: a column chart and a treemap for each Legend.Position,
        # in the order bottom, corner (top-right, overlaid), left, right, top.
        # Excel writes top-right as c:legendPos="tr" and as cx:legend pos="r" align="min".
        xf = XLSX.readxlsx(joinpath(data_directory, "legend_positions.xlsx"))
        cs = XLSX.Charts.getCharts(xf["Sheet1"])
        @test map(typeof, cs) == repeat([XLSX.Charts.Chart, XLSX.Charts.ChartEx], 5)
        @test map(XLSX.Charts.getLegendPos, cs) == repeat([:b, :tr, :l, :r, :t]; inner=2)
        @test map(XLSX.Charts.getLegendOverlay, cs) == repeat([false, true, false, false, false]; inner=2)

        # The setters write the same XML
        tmp = "legend_positions_set.xlsx"
        cp(joinpath(data_directory, "legend_positions.xlsx"), tmp; force=true)
        xf = XLSX.openxlsx(tmp; mode="rw")
        cs = XLSX.Charts.getCharts(xf["Sheet1"])
        cx = cs[2]                                                        # bottom treemap
        XLSX.Charts.setLegendPos(cx, :tr; overlay=true)
        @test XLSX.Charts._cx_legend(cx).attributes == XLSX.Charts._cx_legend(cs[4]).attributes
        c = XLSX.Charts.setLegendPos(cs[1], :tr; overlay=true)
        val(n, t) = XLSX.get_attr(XLSX.first_element_with_tag(XLSX.Charts.getChartLegend(n), t), "val")
        @test val(c, "legendPos") == val(cs[3], "legendPos") == "tr"
        @test val(c, "overlay") == val(cs[3], "overlay") == "1"
        SAVE_FILES && save_outfile(xf)
        isfile(tmp) && rm(tmp)
    end

    @testset "setChartTitleText and setSeriesName" begin
        tmp = "cx_text_set.xlsx"
        cp(joinpath(data_directory, "chartex_layouts.xlsx"), tmp; force=true)
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

        out = "cx_text_set_out.xlsx"
        XLSX.writexlsx(out, xf, overwrite=true)
        f = XLSX.readxlsx(out)
        d = only(filter(x -> x isa XLSX.Charts.ChartEx && x.sheet == "histogram", XLSX.Charts.getCharts(f)))
        @test XLSX.Charts.getChartTitle(d) == "Distribution of values"
        SAVE_FILES && save_outfile(xf)
        isfile(out) && rm(out)
        isfile(tmp) && rm(tmp)
    end

    @testset "rich text and cell-bound titles" begin
        tmp = "cx_rich_titles.xlsx"
        cp(joinpath(data_directory, "chartex_layouts.xlsx"), tmp; force=true)
        xf = XLSX.openxlsx(tmp; mode="rw")
        cx(sheet) = only(filter(x -> x isa XLSX.Charts.ChartEx && x.sheet == sheet, XLSX.Charts.getCharts(xf)))
        cx_title(c) = XLSX.first_element_with_tag(XLSX.Charts._cx_chart(c), "title")
        cx_axis_title(c, id) = XLSX.first_element_with_tag(XLSX.Charts._cx_axis_node(c, id), "title")
        cx_f(t) = XLSX.child_text(XLSX.first_element_with_tag(XLSX.first_element_with_tag(t, "tx"), "txData"), "f")
        runs(t) = [r for p in t.paragraphs for r in p.runs]
        rts = XLSX.RichTextString(XLSX.RichTextRun("Bold", [:bold => true]),
                                  XLSX.RichTextRun(" red\nline 2", [:color => "red"]))
        try
            # A typed title (Excel's own): each run starts from the old first run's
            # rPr — which carries Excel's Aptos Narrow — with its own attributes on top.
            h = cx("histogram")
            XLSX.Charts.setChartTitleText(h, rts)
            @test XLSX.Charts.getChartTitle(h) == "Bold red\nline 2"          # cx:v, plain
            body = XLSX.Charts.getChartTitleTextProps(h)
            @test length(body.paragraphs) == 2
            r = runs(body)
            @test [x.text for x in r] == ["Bold", " red", "line 2"]
            @test r[1].props.bold === true && r[2].props.bold === false       # the old run's b="0"
            @test r[2].props.fill.fgcolor.val == "FF0000" && r[3].props.fill.fgcolor.val == "FF0000"
            @test all(x -> x.props.latin == "Aptos Narrow", r)
            @test all(x -> x.props.size ≈ 14.0, r)
            @test body.paragraphs[2].props.align == "ctr"                     # pPr copied

            # A plain string with a newline gives two paragraphs too.
            XLSX.Charts.setChartTitleText(h, "One\nTwo")
            @test [x.text for x in runs(XLSX.Charts.getChartTitleTextProps(h))] == ["One", "Two"]
            @test XLSX.Charts.getChartTitle(h) == "One\nTwo"

            # A title with no cx:txPr: rich text creates one, runs carrying only their own.
            w = cx("waterfall")
            @test isnothing(XLSX.Charts.getChartTitleTextProps(w))
            XLSX.Charts.setChartTitleText(w, "Plain")                         # plain: still none
            @test isnothing(XLSX.Charts.getChartTitleTextProps(w))
            XLSX.Charts.setChartTitleText(w, rts)
            r = runs(XLSX.Charts.getChartTitleTextProps(w))
            @test r[1].props.bold === true && isnothing(r[1].props.size)
            @test isnothing(r[2].props.bold)

            # Binding: cx:f and the cached text, the runs folded into the paragraph default.
            XLSX.Charts.setChartTitleText(h, rts)
            XLSX.Charts.setChartTitleText(h, XLSX.CellRef("A1"))
            @test cx_f(cx_title(h)) == "histogram!\$A\$1"
            @test !isnothing(XLSX.Charts.getChartTitleRange(h))
            @test XLSX.Charts.getChartTitle(h) == string(xf["histogram"]["A1"])
            body = XLSX.Charts.getChartTitleTextProps(h)
            @test isempty(runs(body))                                         # no stale text
            d = XLSX.Charts.default_run_props(body)
            @test d.bold === true && d.latin == "Aptos Narrow" && d.size ≈ 14.0

            # A numeric cell is cached as text; a rich-text cell lends what its runs share.
            ws = xf["histogram"]
            ws["Z1"] = 2026
            ws["Z2"] = XLSX.RichTextString(XLSX.RichTextRun("Big", [:italic => true, :color => "blue"]),
                                           XLSX.RichTextRun(" blue", [:italic => true, :color => "blue", :bold => false]))
            XLSX.Charts.setChartTitleText(h, XLSX.SheetCellRef("histogram!Z1"))
            @test XLSX.Charts.getChartTitle(h) == "2026"
            XLSX.Charts.setChartTitleText(h, XLSX.CellRef("Z2"))
            d = XLSX.Charts.default_run_props(XLSX.Charts.getChartTitleTextProps(h))
            @test d.italic === true && d.fill.fgcolor.val == "0000FF"
            @test d.bold === true                                             # not shared, so kept
            @test XLSX.Charts.getChartTitle(h) == "Big blue"

            # Typing again unbinds it; the new runs inherit the folded formatting.
            XLSX.Charts.setChartTitleText(h, "Typed")
            @test isnothing(XLSX.Charts.getChartTitleRange(h))
            @test isnothing(only(runs(XLSX.Charts.getChartTitleTextProps(h))).props)

            # Excel's own bound title: the formula is a hidden _xlchart name, which the
            # Range getter resolves. It can be rebound, and typed over with rich text.
            b = cx("bound")
            @test XLSX.Charts.getChartTitleRef(b) == "_xlchart.v1.31"
            r = XLSX.Charts.getChartTitleRange(b)
            @test r isa XLSX.SheetCellRef && r.sheet == "bound" && r.cellref == XLSX.CellRef("D1")
            XLSX.Charts.setChartTitleText(b, XLSX.CellRef("A1"))
            @test XLSX.Charts.getChartTitleRef(b) == "bound!\$A\$1"
            @test cx_f(cx_title(b)) == "bound!\$A\$1"
            XLSX.Charts.setChartTitleText(b, rts)
            @test isnothing(XLSX.Charts.getChartTitleRange(b))
            @test runs(XLSX.Charts.getChartTitleTextProps(b))[1].props.bold === true

            # Axis titles, by axis id, created where there are none.
            XLSX.Charts.setAxisTitleText(w, 1, rts)
            at = XLSX.Charts.getAxisTitleTextProps(w, 1)
            @test XLSX.Charts.text_content(at) == "Bold red\nline 2"
            @test runs(at)[1].props.bold === true
            @test XLSX.Charts.getAxisTitle(w, 1) == "Bold red\nline 2"
            @test isnothing(XLSX.Charts.getAxisTitleRef(w, 1))
            @test isnothing(XLSX.Charts.getAxisTitle(w, 0))                   # no title at all
            xf["waterfall"]["Z1"] = "Movement"
            xf["waterfall"]["Z2"] = "Stage"
            XLSX.Charts.setAxisTitleText(w, 0, XLSX.SheetCellRef("waterfall!Z1"))
            @test cx_f(cx_axis_title(w, 0)) == "waterfall!\$Z\$1"
            @test XLSX.Charts.getAxisTitleRef(w, 0) == "waterfall!\$Z\$1"
            r = XLSX.Charts.getAxisTitleRange(w, 0)
            @test r isa XLSX.SheetCellRef && r.sheet == "waterfall" && r.cellref == XLSX.CellRef("Z1")
            @test XLSX.Charts.getAxisTitle(w, 0) == "Movement"
            XLSX.Charts.setAxisTitleText(w, 1, XLSX.CellRef("Z3"))          # empty: cx:v is ""
            @test isnothing(XLSX.Charts.getAxisTitle(w, 1))                   # as for a c: title
            XLSX.Charts.setAxisTitleText(w, 1, XLSX.CellRef("Z2"))
            @test cx_f(cx_axis_title(w, 1)) == "waterfall!\$Z\$2"
            @test isempty(runs(XLSX.Charts.getAxisTitleTextProps(w, 1)))
            @test_throws XLSX.XLSXError XLSX.Charts.setAxisTitleText(w, 7, "No such axis")

            # Survives a write and reopen.
            SAVE_FILES && save_outfile(xf)
            out = "cx_rich_titles_out.xlsx"
            XLSX.writexlsx(out, xf; overwrite=true)
            try
                yf = XLSX.readxlsx(out)
                cy(sheet) = only(filter(x -> x isa XLSX.Charts.ChartEx && x.sheet == sheet, XLSX.Charts.getCharts(yf)))
                @test XLSX.Charts.getChartTitle(cy("histogram")) == "Typed"
                w2 = cy("waterfall")
                @test runs(XLSX.Charts.getChartTitleTextProps(w2))[1].props.bold === true
                @test XLSX.Charts.getAxisTitleRef(w2, 0) == "waterfall!\$Z\$1"
                @test XLSX.Charts.getAxisTitle(w2, 1) == "Stage"
                @test XLSX.Charts.getChartTitleRange(cy("bound")) === nothing
            finally
                isfile(out) && rm(out)
            end
        finally
            isfile(tmp) && rm(tmp)
        end
    end

    @testset "addChartEx titles" begin
        xf = XLSX.newxlsx()
        ws = xf[1]
        ws["A1"] = "Flow"; ws["A2"] = "a"; ws["A3"] = "b"; ws["B2"] = 3; ws["B3"] = -1
        rts = XLSX.RichTextString(XLSX.RichTextRun("Cash", [:bold => true]), XLSX.RichTextRun(" flow"))
        c = XLSX.Charts.addChartEx(ws, :waterfall, "B2:B3"; anchor = "D2:K16", categories = "A2:A3", title = rts)
        @test XLSX.Charts.getChartTitle(c) == "Cash flow"
        r = only(XLSX.Charts.getChartTitleTextProps(c).paragraphs).runs
        @test r[1].props.bold === true
        b = XLSX.Charts.addChartEx(ws, :waterfall, "B2:B3"; anchor = "D20:K34", title = XLSX.CellRef("A1"))
        @test XLSX.Charts.getChartTitle(b) == "Flow"
        @test !isnothing(XLSX.Charts.getChartTitleRange(b))
        s = XLSX.Charts.addChartEx(ws, :waterfall, "B2:B3"; anchor = "D40:K54", title = "A1")
        @test isnothing(XLSX.Charts.getChartTitleRange(s))                   # a string is literal

        # A bad reference fails before anything is added.
        n = length(XLSX.Charts.getCharts(ws))
        @test_throws XLSX.XLSXError XLSX.Charts.addChartEx(ws, :waterfall, "B2:B3"; anchor = "D60:K74",
                                                           title = XLSX.SheetCellRef("Nope!A1"))
        @test length(XLSX.Charts.getCharts(ws)) == n
        @test_throws XLSX.XLSXError XLSX.Charts.addChartEx(xf, :waterfall, "Sheet1!B2:B3"; title = XLSX.CellRef("A1"))
        @test XLSX.sheetnames(xf) == ["Sheet1"]
        cs = XLSX.Charts.addChartEx(xf, :waterfall, "Sheet1!B2:B3"; title = XLSX.SheetCellRef("Sheet1!A1"))
        @test XLSX.Charts.getChartTitle(cs) == "Flow"
        @test_throws XLSX.XLSXError XLSX.Charts.setChartTitleText(cs, XLSX.CellRef("A1"))

        @static if VERSION >= v"1.11-"
            ss = styled"{bold:Styled} {(foreground=#ff0000):title}"
            check(t) = begin
                r = only(t.paragraphs).runs
                @test [x.text for x in r] == ["Styled", " ", "title"]
                @test r[1].props.bold === true && r[3].props.fill.fgcolor.val == "FF0000"
            end
            t = XLSX.Charts.addChartEx(ws, :waterfall, "B2:B3"; anchor = "D80:K94", title = ss)
            check(XLSX.Charts.getChartTitleTextProps(t))
            XLSX.Charts.setChartTitleText(c, ss)
            check(XLSX.Charts.getChartTitleTextProps(c))
            XLSX.Charts.setAxisTitleText(c, first(XLSX.Charts.getChartAxisIds(c)), ss)
            check(XLSX.Charts.getAxisTitleTextProps(c, first(XLSX.Charts.getChartAxisIds(c))))
        end
        SAVE_FILES && save_outfile(xf)
    end

    @testset "setSeriesSubtotals" begin
        tmp = "cx_subtotals.xlsx"
        cp(joinpath(data_directory, "chartex_layouts.xlsx"), tmp; force=true)
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

        # Excel refuses a file with subtotals or a quartile method on a funnel series.
        fn = only(filter(x -> x isa XLSX.Charts.ChartEx && x.sheet == "funnel", XLSX.Charts.getCharts(xf)))
        @test isnothing(XLSX.Charts._cx_layoutpr(fn, 1))
        err = try XLSX.Charts.setSeriesSubtotals(fn, 1, [2]); nothing catch e; e end
        @test err isa XLSX.XLSXError && occursin("only to a `waterfall` series", err.msg)
        @test_throws XLSX.XLSXError XLSX.Charts.setSeriesQuartileMethod(fn, 1, :inclusive)
        XLSX.Charts.setSeriesSubtotals(fn, 1, :inherit)          # removing is always allowed
        XLSX.Charts.setSeriesQuartileMethod(fn, 1, :inherit)
        @test isnothing(XLSX.Charts.getSeriesSubtotals(fn, 1))

        out = "cx_subtotals_out.xlsx"
        XLSX.writexlsx(out, xf, overwrite=true)
        d = only(filter(x -> x isa XLSX.Charts.ChartEx && x.sheet == "waterfall",
            XLSX.Charts.getCharts(XLSX.readxlsx(out))))
        @test isnothing(XLSX.Charts.getSeriesSubtotals(d, 1))
        SAVE_FILES && save_outfile(xf)
        isfile(out) && rm(out)
        isfile(tmp) && rm(tmp)
    end

    @testset "layoutPr setters" begin
        tmp = "cx_layoutpr.xlsx"
        cp(joinpath(data_directory, "chartex_layouts.xlsx"), tmp; force=true)
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
            out = "cx_layoutpr_out.xlsx"
            XLSX.writexlsx(out, xf, overwrite=true)
            f = XLSX.readxlsx(out)
            h = only(filter(x -> x isa XLSX.Charts.ChartEx && x.sheet == "histogram",
                XLSX.Charts.getCharts(f)))
            @test XLSX.Charts.getSeriesBinning(h, 1).binCount == 7
            isfile(out) && rm(out)
        end
        SAVE_FILES && save_outfile(xf)
        isfile(tmp) && rm(tmp)
    end

    @testset "setting a run property clears it from the runs" begin
        tmp = "runclear.xlsx"
        cp(joinpath(data_directory, "chartex_layouts.xlsx"), tmp; force=true)
        xf = XLSX.openxlsx(tmp; mode="rw")
        h = only(filter(x -> x isa XLSX.Charts.ChartEx && x.sheet == "histogram", XLSX.Charts.getCharts(xf)))

        # Excel wrote this title with sz on both a:defRPr and the run's a:rPr
        XLSX.Charts.setChartTitleTextProp(h, :size, 20)
        tx = XLSX.first_element_with_tag(
            XLSX.first_element_with_tag(XLSX.Charts._cx_chart(h), "title"), "txPr")
        p = XLSX.first_element_with_tag(tx, "p")
        def = XLSX.first_element_with_tag(XLSX.first_element_with_tag(p, "pPr"), "defRPr")
        rpr = XLSX.first_element_with_tag(XLSX.first_element_with_tag(p, "r"), "rPr")

        @test XLSX.get_attr(def, "sz") == "2000"
        @test isempty(XLSX.get_attr(rpr, "sz"))          # cleared, so the default renders
        @test XLSX.get_attr(rpr, "lang") == "en-GB"      # the run keeps everything else

        # what the reader resolves now matches what Excel will render
        @test XLSX.Charts.default_run_props(XLSX.Charts.getChartTitleTextProps(h)).size ≈ 20.0

        # a composite field: the run's own solidFill must go too
        XLSX.Charts.setChartTitleTextProp(h, :fill, "FFC00000")
        rpr = XLSX.first_element_with_tag(XLSX.first_element_with_tag(
            XLSX.first_element_with_tag(
                XLSX.first_element_with_tag(XLSX.Charts._cx_chart(h), "title"), "txPr"), "p"), "r")
        @test isnothing(XLSX.first_element_with_tag(
            XLSX.first_element_with_tag(rpr, "rPr"), "solidFill"))
        @test XLSX.Charts.default_run_props(XLSX.Charts.getChartTitleTextProps(h)).fill.fgcolor.rgb == "C00000"
        SAVE_FILES && save_outfile(xf)
        isfile(tmp) && rm(tmp)
    end

    @testset "Pareto line is reached with pareto = true" begin
        kinds = XLSX.getCharts(XLSX.readxlsx(joinpath(data_directory, "chartex_kinds.xlsx")))
        p = kinds[6]                                        # Pareto
        w = kinds[1]                                        # waterfall: no Pareto line

        # One series, as in Excel's Select Data; the line belongs to it
        @test XLSX.Charts.getChartType(p) === :pareto
        @test XLSX.Charts.getChartSeriesCount(p) == 1
        @test XLSX.Charts.getSeriesLayout(p, 1) === :clusteredColumn
        @test XLSX.Charts.getSeriesLayout(p, 1; pareto=true) === :paretoLine
        @test_throws XLSX.XLSXError XLSX.Charts.getSeriesLayout(p, 2)
        # the line is drawn against the secondary (percentage) axis
        @test XLSX.Charts.getSeriesAxisIds(p, 1; pareto=true) != XLSX.Charts.getSeriesAxisIds(p, 1)

        # pareto = true on a chart with no Pareto line is an error, not the series itself
        @test_throws XLSX.XLSXError XLSX.Charts.getSeriesLayout(w, 1; pareto=true)
        @test_throws XLSX.XLSXError XLSX.Charts.getSeriesLine(w, 1; pareto=true)

        # A setter with pareto = true formats the line and leaves the columns alone
        xf = XLSX.openxlsx(joinpath(data_directory, "chartex_kinds.xlsx"); mode="rw")
        c = XLSX.getCharts(xf)[6]
        cols_before = XLSX.Charts.getSeriesLine(c, 1)
        line_before = XLSX.Charts.getSeriesLine(c, 1; pareto=true)
        XLSX.Charts.setSeriesLineColor(c, 1, "FF0000"; pareto=true)
        XLSX.Charts.setSeriesLineWidth(c, 1, 2.5; pareto=true)
        line_after = XLSX.Charts.getSeriesLine(c, 1; pareto=true)
        @test line_after != line_before
        line_after = XLSX.Charts.getSeriesLine(c, 1; pareto=true)
        @test line_after.value != line_before.value
        @test XLSX.Charts.getSeriesLine(c, 1).value == cols_before.value
        @test XLSX.Charts.getChartSeriesCount(c) == 1         # formatting adds no series

        f = "pareto_line.xlsx"
        XLSX.writexlsx(f, xf; overwrite=true)
        SAVE_FILES && save_outfile(xf)
        yf = XLSX.openxlsx(f; mode="rw")
        d = XLSX.getCharts(yf)[6]
        @test XLSX.Charts.getSeriesLine(d, 1; pareto=true).value == line_after.value
        @test XLSX.Charts.getSeriesLine(d, 1).value == cols_before.value
        @test XLSX.Charts.getChartSeriesCount(d) == 1
        SAVE_FILES && save_outfile(yf)
        isfile(f) && rm(f)
    end
end
