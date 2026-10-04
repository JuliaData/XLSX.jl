
#
# Charts_tests.jl - cached chart data and chart metadata (issue #263)
#
# Fixtures live in test/data. Each testset names the fixture it needs in its
# first line so a missing file is obvious.
#
# Assertions marked `# FIXTURE:` encode a property of the file as built. If you
# rebuild a fixture differently, these are the lines to revisit.
#

@testset "Charts" begin

    @testset "basic chart" begin  # chart_basic.xlsx
        f = XLSX.readxlsx(joinpath(data_directory, "chart_basic.xlsx"))

        c = XLSX.Charts.getChart(f["Data"], "chart1")
        @test occursin("chart1", repr(c))
        @test occursin("series", repr(MIME"text/plain"(), c))
        @test occursin("ChartSeries", repr(XLSX.Charts.getChartSeries(c)[1]))
        @test occursin("pts", repr(MIME"text/plain"(), XLSX.Charts.getChartSeries(c)[1].values))
        @test occursin("pts", sprint(show, XLSX.Charts.getChartSeries(c)[1].values))
        @test occursin("categories", repr(MIME"text/plain"(), XLSX.Charts.getChartSeries(c)[1]))

        charts = XLSX.Charts.getCharts(f)
        @test length(charts) == 1

        c = charts[1]
        @test c.name == "chart1"
        @test c.path == "xl/charts/chart1.xml"
        @test c.sheet == "Data"
        @test XLSX.Charts.getChartTitle(c) == "Revenue by Region"
        @test XLSX.Charts.getChartTypes(c) == [:barChart]
        @test !isnothing(c.rId)

        # Anchored to cells, so both markers parse. The exact cells depend on
        # where the chart was dropped, so only their form is asserted.
        @test !isnothing(c.from)
        @test !isnothing(c.to)
        @test XLSX.CellRef(c.from) isa XLSX.CellRef
        @test XLSX.CellRef(c.to) isa XLSX.CellRef

        @test length(XLSX.Charts.getChartSeries(c)) == 2
        @test [s.name for s in XLSX.Charts.getChartSeries(c)] == ["2024", "2025"]
        @test [s.order for s in XLSX.Charts.getChartSeries(c)] == [0, 1]
        @test [s.idx for s in XLSX.Charts.getChartSeries(c)] == [0, 1]
        @test all(s -> s.charttype == :barChart, XLSX.Charts.getChartSeries(c))
        @test all(s -> isnothing(s.bubble_sizes), XLSX.Charts.getChartSeries(c))

        s1, s2 = XLSX.Charts.getChartSeries(c)
        @test s1.categories.kind == :str
        @test s1.categories.ref == "Data!\$A\$2:\$A\$5"
        @test s1.categories.data == ["North", "South", "East", "West"]
        @test s1.categories.ptCount == 4

        @test s1.values.kind == :num
        @test s1.values.ref == "Data!\$B\$2:\$B\$5"
        @test s1.values.data == [10.0, 20.0, 15.0, 5.0]
        @test s1.values.format_code == "General"
        @test isempty(s1.values.errors)

        @test s2.values.ref == "Data!\$C\$2:\$C\$5"
        @test s2.values.data == [12.0, 18.0, 25.0, 9.0]

        # Both series share one category reference, so one `categories` column.
        dt = XLSX.Charts.getChartData(c)
        @test dt.column_labels == [:categories, Symbol("2024"), Symbol("2025")]
        @test dt.data[1] == ["North", "South", "East", "West"]
        @test dt.data[2] == [10.0, 20.0, 15.0, 5.0]
        @test dt.data[3] == [12.0, 18.0, 25.0, 9.0]

        df = DataFrames.DataFrame(dt)
        @test size(df) == (4, 3)
        @test names(df) == ["categories", "2024", "2025"]

        # A DataTable must report real column types, or PrettyTables trips over
        # `Tables.Schema(names, nothing)`.
        sch = Tables.schema(dt)
        @test !isnothing(sch.types)
        @test sch.types[2] == Float64
    end

    @testset "chart lookup" begin  # chart_basic.xlsx
        f = XLSX.readxlsx(joinpath(data_directory, "chart_basic.xlsx"))
        c = XLSX.Charts.getCharts(f)[1]

        @test XLSX.Charts.getChart(f, "chart1").path == c.path
        @test XLSX.Charts.getChart(f, "chart1.xml").path == c.path
        @test XLSX.Charts.getChart(f, "xl/charts/chart1.xml").path == c.path
        @test XLSX.Charts.getChart(f, c.rId).path == c.path
        @test_throws XLSX.XLSXError XLSX.Charts.getChart(f, "nosuchchart")

        # Worksheet-scoped discovery walks the sheet's own drawing.
        ws_charts = XLSX.Charts.getCharts(f["Data"])
        @test length(ws_charts) == 1
        @test ws_charts[1].path == c.path

        @test XLSX.Charts.getChartData(f, "chart1").column_labels == XLSX.Charts.getChartData(c).column_labels
    end

    @testset "metadata only" begin  # chart_basic.xlsx
        f = XLSX.readxlsx(joinpath(data_directory, "chart_basic.xlsx"))
        c = XLSX.Charts.getCharts(f)[1]

        # Metadata survives; values do not.
        @test XLSX.Charts.getChartTitle(c) == "Revenue by Region"
        @test [s.name for s in XLSX.Charts.getChartSeries(c; read_cached_values=false)] == ["2024", "2025"]   # names come from c:tx regardless
        @test XLSX.Charts.getChartSeries(c; read_cached_values=false)[1].values.ref == "Data!\$B\$2:\$B\$5"
        @test XLSX.Charts.getChartSeries(c; read_cached_values=false)[1].values.format_code == "General"
        @test XLSX.Charts.getChartSeries(c; read_cached_values=false)[1].values.ptCount == 4
        @test isempty(XLSX.Charts.getChartSeries(c; read_cached_values=false)[1].values.data)
        @test isempty(XLSX.Charts.getChartSeries(c; read_cached_values=false)[1].categories.data)

        # Reading options belong to the accessors, not the chart, so the data table
        # is always available.
        @test XLSX.Charts.getChartData(c) isa XLSX.DataTable
        @test XLSX.Charts.getChartData(f, "chart1") isa XLSX.DataTable
    end

    @testset "gaps and errors" begin  # chart_gaps.xlsx
        f = XLSX.readxlsx(joinpath(data_directory, "chart_gaps.xlsx"))
        c = XLSX.Charts.getCharts(f)[1]

        @test XLSX.Charts.getChartTypes(c) == [:lineChart]
        @test length(XLSX.Charts.getChartSeries(c; read_cached_values=false)) == 1

        v = XLSX.Charts.getChartSeries(c)[1].values
        @test v.ptCount == 5
        @test length(v.data) == 5

        # Row 3 is blank: Excel omits the point entirely, leaving a gap in `idx`.
        # Rows 4 and 5 are errors: written as text, stored as `missing`, recorded
        # in `errors`. Both read back as `missing`; only `errors` tells them apart.
        @test v.data[1] == 1.0
        @test ismissing(v.data[2])
        @test ismissing(v.data[3])
        @test v.data[4] == 0.0        # #DIV/0! in the source; Excel caches it as 0
        @test v.data[5] == 5.0

        @test XLSX.iserror(v) == [false, false, true, false, false]

        @test XLSX.geterror(v, 3) == "#N/A"
        @test XLSX.geterror(v, 4) == ""
        @test XLSX.geterror(v, 1) == ""     # #DIV/0! resolves to no error
        @test XLSX.geterror(v, 2) == ""     # blank resolves to no error
        @test XLSX.geterror(v) == ["", "", "#N/A", "", ""]

        @test length(XLSX.geterror(v)) == length(v.data)

        # Categories are complete even where values are not.
        @test XLSX.Charts.getChartSeries(c)[1].categories.data == ["a", "b", "c", "d", "e"]
    end

    @testset "combo chart" begin  # chart_combo.xlsx
        f = XLSX.readxlsx(joinpath(data_directory, "chart_combo.xlsx"))
        c = XLSX.Charts.getCharts(f)[1]

        # Two group elements in one plotArea. Excel writes them in schema order,
        # so compare as a set rather than pinning the sequence.
        @test length(XLSX.Charts.getChartTypes(c)) == 2
        @test issetequal(XLSX.Charts.getChartTypes(c), [:barChart, :lineChart])

        # Series are collected across both groups, with `order` continuing rather
        # than restarting per group.
        @test length(XLSX.Charts.getChartSeries(c; read_cached_values=false)) == 2
        @test sort([s.order for s in XLSX.Charts.getChartSeries(c; read_cached_values=false)]) == [0, 1]
        @test issetequal([s.charttype for s in XLSX.Charts.getChartSeries(c; read_cached_values=false)], [:barChart, :lineChart])

        dt = XLSX.Charts.getChartData(c)
        @test length(dt.column_labels) == 3   # shared categories + two series
    end

    @testset "scatter chart" begin  # chart_scatter.xlsx
        f = XLSX.readxlsx(joinpath(data_directory, "chart_scatter.xlsx"))
        c = XLSX.Charts.getCharts(f)[1]

        @test XLSX.Charts.getChartTypes(c) == [:scatterChart]
        @test length(XLSX.Charts.getChartSeries(c; read_cached_values=false)) == 2

        # c:xVal folds into `categories`, c:yVal into `values`.
        s1, s2 = XLSX.Charts.getChartSeries(c)
        @test s1.categories.data == [1.0, 2.0, 3.0, 4.0]
        @test s1.values.data == [10.0, 20.0, 30.0, 40.0]
        @test s2.categories.data == [2.0, 4.0, 6.0, 8.0]
        @test s2.values.data == [5.0, 15.0, 25.0, 35.0]

        # Differing x references, so no shared `categories` column: each series
        # brings its own.
        @test s1.categories.ref != s2.categories.ref
        dt = XLSX.Charts.getChartData(c)
        @test length(dt.column_labels) == 4
        @test endswith(String(dt.column_labels[1]), "_x")
        @test endswith(String(dt.column_labels[3]), "_x")
    end

    @testset "bubble chart and document order" begin  # chart_bubble.xlsx
        f = XLSX.readxlsx(joinpath(data_directory, "chart_bubble.xlsx"))
        charts = XLSX.Charts.getCharts(f)
        @test length(charts) == 2

        # FIXTURE: the pie chart was inserted first but positioned to the right
        # of the bubble chart. Anchor order is creation order, not reading order,
        # so the pie comes back first.
        @test XLSX.Charts.getChartTypes(charts[1]) == [:pieChart]
        @test XLSX.Charts.getChartTypes(charts[2]) == [:bubbleChart]

        bub = charts[2]
        s = XLSX.Charts.getChartSeries(bub)[1]
        @test !isnothing(s.bubble_sizes)
        @test length(s.bubble_sizes.data) == length(s.values.data)

        dt = XLSX.Charts.getChartData(bub)
        @test any(l -> endswith(String(l), "_size"), dt.column_labels)
    end

    @testset "multi-level categories" begin  # chart_multilevel.xlsx
        f = XLSX.readxlsx(joinpath(data_directory, "chart_multilevel.xlsx"))
        c = XLSX.Charts.getCharts(f)[1]

        cat = XLSX.Charts.getChartSeries(c)[1].categories
        @test cat.kind == :multiLvlStr
        @test cat.ptCount == 4
        @test length(cat.data) == 2          # two levels, each a vector

        # FIXTURE: `c:lvl` elements are stored in document order. Confirm which
        # way round Excel wrote them by unzipping the fixture and reading
        # xl/charts/chart1.xml before trusting this - and correct the ChartRef
        # docstring if it is the other way about.
        levels = cat.data
        @test issetequal(levels[1], ["Q1", "Q2"])
        @test issetequal(levels[2], ["North", "South"])
        @test levels[1] == ["Q1", "Q2", "Q1", "Q2"]
        @test levels[2] == ["North", "North", "South", "South"]

        dt = XLSX.Charts.getChartData(c)
        @test dt.column_labels[1] == :categories_1
        @test dt.column_labels[2] == :categories_2
        @test length(dt.column_labels) == 3   # two levels + one series
    end

    @testset "chartsheet" begin  # chart_chartsheet.xlsx
        f = XLSX.readxlsx(joinpath(data_directory, "chart_chartsheet.xlsx"))
        charts = XLSX.Charts.getCharts(f)
        @test length(charts) == 1

        c = charts[1]
        @test c.sheet == "TheChart"

        # A chartsheet drawing uses an absoluteAnchor with pos/ext, not cell
        # markers, so there is nothing to resolve into a CellRef.
        @test isnothing(c.from)
        @test isnothing(c.to)

        # The data still reads exactly as when the chart was on a worksheet.
        @test length(XLSX.Charts.getChartSeries(c)) == 2
        @test XLSX.Charts.getChartSeries(c)[1].values.data == [10.0, 20.0, 15.0, 5.0]
    end

    @testset "strict OOXML" begin  # chart_strict.xlsx
        f = XLSX.readxlsx(joinpath(data_directory, "chart_strict.xlsx"))
        charts = XLSX.Charts.getCharts(f)

        # Strict files are remapped to transitional at read, so the transitional
        # chart relationship type still resolves and the chart is found.
        @test length(charts) == 1
        c = charts[1]
        @test c.sheet == "Data"
        @test XLSX.Charts.getChartTypes(c) == [:barChart]
        @test XLSX.Charts.getChartSeries(c)[1].categories.data == ["North", "South", "East", "West"]
        @test XLSX.Charts.getChartSeries(c)[1].values.data == [10.0, 20.0, 15.0, 5.0]
    end

    @testset "chartEx is read for discovery" begin
        f = XLSX.readxlsx(joinpath(data_directory, "chart_ex.xlsx"))   # your path

        charts = @test_nowarn XLSX.Charts.getCharts(f)
        @test length(charts) == 1

        c = charts[1]
        @test c isa XLSX.Charts.ChartEx
        @test XLSX.Charts.getChartSchema(c) === :cx
        @test XLSX.Charts.getChartType(c) === :waterfall
        @test c.name == "chartEx1"
        @test c.sheet == "Data"
        @test c.from == "G9"
        @test length(XLSX.Charts._cx_refs(c)) == 2
        @test XLSX.Charts._cx_refs(c) == ["_xlchart.v1.0", "_xlchart.v1.1"]
        ranges = XLSX.Charts.getChartRanges(c)
        @test length(ranges) == 2
        @test all(!isnothing, ranges)
        @test string(ranges[1]) == "Data!A2:A5"
        @test string(ranges[2]) == "Data!B2:B5"

        # Ranges resolve; cached values do not exist.
        ranges = XLSX.Charts.getChartRanges(c)
        @test length(ranges) == 2
        @test all(!isnothing, ranges)

        err = @test_throws XLSX.XLSXError XLSX.Charts.getChartData(c)
        @test occursin("waterfall", err.value.msg)
        @test occursin("chartEx", err.value.msg)

        # Reachable by name and by rId, like a c: chart.
        @test XLSX.Charts.getChart(f, "chartEx1") isa XLSX.Charts.ChartEx
        @test XLSX.Charts.getChart(f, "chartEx1.xml") isa XLSX.Charts.ChartEx

        # Sheet-scoped discovery finds it too.
        @test length(XLSX.Charts.getCharts(f["Data"])) == 1

        rm_xf = XLSX.readxlsx(joinpath(data_directory, "chartex_regionmap.xlsx"))
        rmc = only(XLSX.getCharts(rm_xf))
        @test rmc isa XLSX.Charts.ChartEx
        @test XLSX.Charts.getChartType(rmc) === :regionMap
        # survives a round trip
        f = "regionmap_rt.xlsx"
        XLSX.writexlsx(f, XLSX.openxlsx(joinpath(data_directory, "chartex_regionmap.xlsx"); mode="rw"); overwrite=true)
        SAVE_FILES && save_outfile(f)
        @test only(XLSX.getCharts(XLSX.readxlsx(f))) isa XLSX.Charts.ChartEx
        isfile(f) && rm(f)
    end
    @testset "external reference" begin  # chart_external.xlsx
        f = XLSX.readxlsx(joinpath(data_directory, "chart_external.xlsx"))
        c = XLSX.Charts.getCharts(f)[1]

        @test XLSX.Charts.getChartTypes(c) == [:lineChart]
        @test length(XLSX.Charts.getChartSeries(c)) == 2

        # Source lives in another workbook: the `[1]` indexes the workbook's
        # external references.
        @test XLSX.Charts.getChartSeries(c)[1].values.ref == "[1]Feuil1!\$A\$1:\$A\$10"
        @test XLSX.Charts.getChartSeries(c)[2].values.ref == "[1]Feuil1!\$B\$1:\$B\$10"

        # No c:cat at all - Excel plots against an implicit index and caches
        # nothing for it - and no c:tx, so the series are unnamed.
        @test all(s -> isnothing(s.categories), XLSX.Charts.getChartSeries(c; read_cached_values=false))
        @test all(s -> isnothing(s.name), XLSX.Charts.getChartSeries(c; read_cached_values=false))

        @test XLSX.Charts.getChartSeries(c)[1].values.data == collect(1.0:10.0)
        @test XLSX.Charts.getChartSeries(c)[2].values.data == [1.0, 10.0, 5.0, 2.0, 3.0, 45.0, 6.0, 8.0, 7.0, 2.0]

        # No categories, so no leading column, and positional series labels.
        dt = XLSX.Charts.getChartData(c)
        @test dt.column_labels == [:Series1, :Series2]
        @test length(dt.data[1]) == 10

        # Materialising the reference resolves `[1]` through xl/externalLinks and
        # its relationships (regression test for get_external_workbook_path,
        # which previously ignored the externalBook r:id).
        c = XLSX.Charts.getCharts(f)[1]
        @test XLSX.Charts.getChartSeries(c; get_external_refs=true)[1].values.ref == "[Test2.xlsx]Feuil1!\$A\$1:\$A\$10"
        @test XLSX.Charts.getChartSeries(c; get_external_refs=true)[1].values.data == XLSX.Charts.getChartSeries(c)[1].values.data   # unchanged by materialising
    end

    @testset "chart cache agrees with external link cache" begin  # chart_external.xlsx
        # The workbook carries two independently written copies of the same
        # numbers: the chart's c:numCache, and the externalLink's sheetDataSet.
        # Parsing the second by hand checks the first without a reference
        # implementation.
        f = XLSX.readxlsx(joinpath(data_directory, "chart_external.xlsx"))

        function external_column(xf, col::String)
            root = XLSX.xml_root_element(xf.data["xl/externalLinks/externalLink1.xml"])
            book = XLSX.first_element_with_tag(root, "externalBook")
            dataset = XLSX.first_element_with_tag(book, "sheetDataSet")
            sheet = XLSX.first_element_with_tag(dataset, "sheetData")
            out = Float64[]
            for row in XML.children(sheet)
                XML.nodetype(row) == XML.Element && XLSX.localname(row) == "row" || continue
                for cell in XML.children(row)
                    XML.nodetype(cell) == XML.Element && XLSX.localname(cell) == "cell" || continue
                    startswith(XLSX.get_attr(cell, "r"), col) || continue
                    push!(out, parse(Float64, XLSX.child_text(cell, "v")))
                end
            end
            return out
        end

        c = XLSX.Charts.getCharts(f)[1]
        @test XLSX.Charts.getChartSeries(c)[1].values.data == external_column(f, "A")
        @test XLSX.Charts.getChartSeries(c)[2].values.data == external_column(f, "B")
    end

    @testset "round trip" begin  # chart_basic.xlsx
        # Reading charts must not disturb them: the chart part should survive a
        # save byte for byte, and read back identically.
        original = joinpath(data_directory, "chart_basic.xlsx")
        f = XLSX.openxlsx(original; mode="rw")
        before = XLSX.Charts.getChartData(XLSX.Charts.getCharts(f)[1])

        tmp = "chart_basic_roundtrip.xlsx"
        XLSX.writexlsx(tmp, f; overwrite=true)

        part = "xl/charts/chart1.xml"
        # XML.jl pretty-prints where Excel writes a single line; the CRLF after the
        # declaration differs too. Both are cosmetic — XML ignores inter-tag
        # whitespace — so compare with it collapsed rather than byte for byte.
        strip_ws(s) = replace(s, r">\s+<" => "><")
        @test strip_ws(String(zip_readentry(ZipReader(read(original)), part))) ==
            strip_ws(String(zip_readentry(ZipReader(read(tmp)), part)))

        g = XLSX.readxlsx(tmp)
        c = XLSX.Charts.getCharts(g)[1]
        @test c.sheet == "Data"
        @test XLSX.Charts.getChartTitle(c) == "Revenue by Region"
        after = XLSX.Charts.getChartData(c)
        @test after.column_labels == before.column_labels
        @test after.data == before.data

        SAVE_FILES && save_outfile(tmp)
        isfile(tmp) && rm(tmp)
    end

    @testset "chart_range" begin
        cr(ref) = XLSX.Charts.chart_range(XLSX.Charts.ChartRef(:num, ref, nothing, 0, Any[], Dict{Int,UInt64}()))

        @test cr(nothing) === nothing
        @test XLSX.Charts.chart_range(nothing) === nothing
        @test cr("[1]Sheet1!\$A\$1:\$A\$5") === nothing        # external
        @test cr("MyDefinedName") === nothing                  # defined name
        @test cr("Sheet1!\$A\$1:\$A\$5") isa XLSX.SheetCellRange
        @test cr("Sheet1!A1:A5") isa XLSX.SheetCellRange
        @test cr("Sheet1!\$A\$1") isa XLSX.SheetCellRef
        @test cr("Sheet1!A:C") isa XLSX.SheetColumnRange
        @test cr("(Sheet1!\$A\$1:\$A\$3,Sheet1!\$C\$1:\$C\$3)") isa XLSX.NonContiguousRange
        @test cr("Sheet1!\$A\$1:\$A\$3,Sheet1!\$C\$1:\$C\$3") isa XLSX.NonContiguousRange
    end
    @testset "row-range source (ChartRange union)" begin
        rr(ref) = XLSX.Charts.ChartRef(:num, ref, nothing, 0, Any[], Dict{Int,UInt64}())

        @test XLSX.Charts.chart_range(rr("Sheet1!\$2:\$5")) isa XLSX.SheetRowRange
        @test XLSX.Charts.chart_range(rr("Sheet1!2:5")) isa XLSX.SheetRowRange
        @test XLSX.Charts.chart_range(rr("Sheet1!\$A:\$C")) isa XLSX.SheetColumnRange

        # the conversion that used to throw
        s = XLSX.Charts.ChartSeries(0, 0, :barChart, "S", nothing,
            rr("Sheet1!\$2:\$2"), rr("Sheet1!\$3:\$3"), nothing,
            XML.Element("c:ser"))
        ranges = XLSX.Charts._chart_ranges([s])
        @test ranges[1].categories isa XLSX.SheetRowRange
        @test ranges[1].values isa XLSX.SheetRowRange
    end
    @testset "ChartRange union covers chart_range" begin
        @test XLSX.SheetCellRef <: XLSX.Charts.ChartRange
        @test XLSX.SheetCellRange <: XLSX.Charts.ChartRange
        @test XLSX.SheetColumnRange <: XLSX.Charts.ChartRange
        @test XLSX.SheetRowRange <: XLSX.Charts.ChartRange
        @test XLSX.NonContiguousRange <: XLSX.Charts.ChartRange
        @test Nothing <: XLSX.Charts.ChartRange
    end

    @testset "literal series" begin
        lit(tag, pts) = XLSX.XML.parse("""<c:val xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:$tag><c:ptCount val="$(length(pts))"/>""" *
            join("<c:pt idx=\"$(i-1)\"><c:v>$p</c:v></c:pt>" for (i, p) in enumerate(pts)) *
            "</c:$tag></c:val>", XLSX.XML.Node)
        root(d) = first(XLSX.XML.eachelement(d))

        n = XLSX.Charts.parse_chart_ref(root(lit("numLit", [1, 2.5])))
        @test n.kind === :numLit
        @test isnothing(n.ref)
        @test n.data == [1.0, 2.5]
        @test isnothing(XLSX.Charts.chart_range(n))                 # literal: no addressable range

        s = XLSX.Charts.parse_chart_ref(root(lit("strLit", ["a", "b"])))
        @test s.kind === :strLit
        @test s.data == ["a", "b"]

        m = XLSX.Charts.parse_chart_ref(root(lit("numLit", [1, 2])); read_cached_values=false)
        @test isempty(m.data)
        @test m.ptCount == 2
    end
    @testset "unique labels" begin
        labels = Symbol[]
        @test XLSX.Charts.unique_label!(labels, "Sales") === :Sales
        @test XLSX.Charts.unique_label!(labels, "Sales") === :Sales_2
        @test XLSX.Charts.unique_label!(labels, "Sales") === :Sales_3
        @test XLSX.Charts.unique_label!(labels, "") === :column
    end

    @testset "getChartRanges dispatch" begin
        f = XLSX.readxlsx(joinpath(data_directory, "chart_bubble.xlsx"))

        # workbook-wide form: Vector{@NamedTuple{chart::String, ranges::Vector{ChartRanges}}}
        all_f = XLSX.Charts.getChartRanges(f)
        @test all_f isa Vector
        @test length(all_f) == 2
        @test all(x -> x isa NamedTuple{(:chart, :ranges)}, all_f)
        @test all(x -> x.chart isa String, all_f)
        @test all(x -> x.ranges isa Vector{XLSX.Charts.ChartRanges}, all_f)

        # follows getCharts, per the docstring
        @test [x.chart for x in all_f] == [c.name for c in XLSX.Charts.getCharts(f)]

        # identify the two charts by type rather than by part name
        charts = XLSX.Charts.getCharts(f)
        bub = charts[findfirst(c -> :bubbleChart in XLSX.Charts.getChartTypes(c), charts)]
        pie = charts[findfirst(c -> :pieChart in XLSX.Charts.getChartTypes(c), charts)]

        # --- (x, name) form -----------------------------------------------------
        rb = XLSX.Charts.getChartRanges(f, bub.name)
        rp = XLSX.Charts.getChartRanges(f, pie.name)
        @test rb isa Vector{XLSX.Charts.ChartRanges}
        @test rp isa Vector{XLSX.Charts.ChartRanges}

        # parallel to XLSX.Charts.getChartSeries(bub), document order
        @test length(rb) == length(XLSX.Charts.getChartSeries(bub))
        @test [x.idx for x in rb] == [s.idx for s in XLSX.Charts.getChartSeries(bub)]
        @test [x.name for x in rb] == [s.name for s in XLSX.Charts.getChartSeries(bub)]

        # every field is a member of the declared union
        for x in vcat(rb, rp), fld in (:categories, :values, :bubble_sizes)
            @test getfield(x, fld) isa XLSX.Charts.ChartRange
        end

        # the docstring's specific claim: bubble_sizes only on bubble charts
        @test any(!isnothing(x.bubble_sizes) for x in rb)
        @test all(isnothing(x.bubble_sizes) for x in rp)

        # bubble uses xVal/yVal, which land in categories/values
        @test all(!isnothing(x.categories) for x in rb)
        @test all(!isnothing(x.values) for x in rb)

        # name forms getChart accepts
        @test XLSX.Charts.getChartRanges(f, bub.name * ".xml") == rb

        @test_throws XLSX.XLSXError XLSX.Charts.getChartRanges(f, "nosuchchart")

        # --- worksheet form agrees with the workbook form ------------------------
        ws = f[bub.sheet]
        all_ws = XLSX.Charts.getChartRanges(ws)
        @test all(x -> x isa NamedTuple{(:chart, :ranges)}, all_ws)
        @test issubset(Set(x.chart for x in all_ws), Set(x.chart for x in all_f))

        i = findfirst(x -> x.chart == bub.name, all_f)
        @test !isnothing(i)
        @test all_f[i].ranges == rb

    end

    @testset "mixed c: and cx: on one sheet" begin
        f = XLSX.readxlsx(joinpath(data_directory, "chart_mixed.xlsx"))
        charts = XLSX.Charts.getCharts(f)

        @test length(charts) == 2
        @test count(c -> c isa XLSX.Charts.Chart, charts) == 1
        @test count(c -> c isa XLSX.Charts.ChartEx, charts) == 1
        @test issetequal(XLSX.Charts.getChartSchema.(charts), [:c, :cx])

        # Each chart keeps its own anchor — the risk is one anchor's from/to
        # being attributed to the other chart's frame.
        @test all(c -> !isnothing(c.sheet), charts)

        c = only(filter(x -> x isa XLSX.Charts.Chart, charts))
        cx = only(filter(x -> x isa XLSX.Charts.ChartEx, charts))

        @test cx.from == "F1"
        @test c.from == "F17"
        @test XLSX.Charts.getChartType(c) === :barChart      # adjust to what you inserted
        @test XLSX.Charts.getChartType(cx) === :waterfall

        # Data works for one, throws for the other; ranges work for both.
        @test XLSX.Charts.getChartData(c) isa XLSX.DataTable
        @test_throws XLSX.XLSXError XLSX.Charts.getChartData(cx)
        @test !isempty(XLSX.Charts.getChartRanges(c))
        @test all(!isnothing, XLSX.Charts.getChartRanges(cx))

        # Lookup by name reaches both.
        @test XLSX.Charts.getChart(f, c.name) isa XLSX.Charts.Chart
        @test XLSX.Charts.getChart(f, cx.name) isa XLSX.Charts.ChartEx

        # Sheet-scoped discovery finds both.
        @test length(XLSX.Charts.getCharts(f[c.sheet])) == 2

        # Document order within the drawing, not schema order.
        @test charts[1] isa XLSX.Charts.ChartEx
        @test charts[2] isa XLSX.Charts.Chart
    end
    @testset "c: chart kind templates match Excel" begin
        fx = joinpath(data_directory, "chart_kinds.xlsx")
        raw = XLSX.ZipArchives.ZipReader(read(fx))
        bysheet = Dict(c.sheet => c for c in XLSX.Charts.getCharts(XLSX.readxlsx(fx)))
        for (kind, t) in pairs(XLSX.Charts.C_KINDS)
            c = bysheet[String(t.template)]
            @test XLSX.ZipArchives.zip_readentry(raw, c.path, String) == XLSX.Charts.CHART_KIND_TEMPLATES[t.template]
            @test haskey(XLSX.Charts.CHART_STYLE_TEMPLATES, t.style)
        end
    end
    @testset "getChartData with per-series x values" begin
        # chart_scatter_xy has two scatter series on different x ranges, so the
        # categories can't collapse into one column.
        xf = XLSX.readxlsx(joinpath(data_directory, "chart_scatter_xy.xlsx"))
        c = XLSX.Charts.getChart(xf, "chart1")

        r = XLSX.Charts.getChartRanges(c)
        @test r[1].categories != r[2].categories        # the premise of the fixture

        dt = XLSX.Charts.getChartData(c)
        @test dt.column_labels == [:y1_x, :y1, :y2_x, :y2]
        @test dt.data[1] == [1.0, 2.0, 3.0, 4.0]      # B1_x
        @test dt.data[2] == [10.0, 20.0, 30.0, 40.0]  # B1
        @test dt.data[3] == [2.0, 4.0, 6.0, 8.0]      # D1_x
        @test dt.data[4] == [5.0, 15.0, 25.0, 35.0]   # D1    
    end
    @testset "cx: chart kind templates match Excel" begin
        fx = joinpath(data_directory, "chartex_kinds.xlsx")
        raw = XLSX.ZipArchives.ZipReader(read(fx))
        for c in XLSX.Charts.getCharts(XLSX.readxlsx(fx))
            c isa XLSX.Charts.ChartEx || continue
            k = XLSX.Charts.getChartType(c)
            @test haskey(XLSX.Charts.CX_KINDS, k)
            @test XLSX.ZipArchives.zip_readentry(raw, c.path, String) == XLSX.Charts.CHARTEX_KIND_TEMPLATES[k]
        end
    end
    @testset "chart plumbing: a verbatim template on a new sheet" begin
        path = "chart_plumbing.xlsx"
        xf = XLSX.newxlsx("column")
        ws = xf["column"]
        for (r, row) in enumerate((("Region", "Alpha", "Beta", "Gamma"),
            ("North", 10, 12, 14), ("South", 20, 18, 22),
            ("East", 15, 25, 19), ("West", 5, 9, 11)))
            for (col, v) in enumerate(row)
                ws[r, col] = v
            end
        end

        p = XLSX.Charts._add_chart_part!(ws, XLSX.Charts.CHART_KIND_TEMPLATES[:column], 201;
            anchor=XLSX.CellRange("F2:M18"))
        @test p == "xl/charts/chart1.xml"

        c = only(XLSX.Charts.getCharts(xf))
        @test c isa XLSX.Charts.Chart
        @test c.sheet == "column"
        @test XLSX.Charts.getChartTypes(c) == [:barChart]
        @test [s.name for s in XLSX.Charts.getChartSeries(c)] == ["Alpha", "Beta", "Gamma"]
        @test XLSX.Charts.getChartData(c).data[2] == [10, 20, 15, 5]

        XLSX.writexlsx(path, xf, overwrite=true)
        c2 = only(XLSX.Charts.getCharts(XLSX.readxlsx(path)))
        @test XLSX.Charts.getChartTypes(c2) == [:barChart]
        @test c2.from == "F2"

        SAVE_FILES && save_outfile(xf)
        isfile(path) && rm(path)
    end

    @testset "addChart on a worksheet" begin
        path = "chart_add.xlsx"
        xf = XLSX.newxlsx("data")
        ws = xf["data"]

        c = XLSX.Charts.addChart(ws, :column; anchor="F2:M18", title="Revenue")
        @test c isa XLSX.Charts.Chart
        @test XLSX.Charts.getChartTypes(c) == [:barChart]
        @test isempty(XLSX.Charts.getChartSeries(c))
        @test XLSX.Charts.getChartTitle(c) == "Revenue"

        # Fresh axis ids, consistently wired.
        axes = XLSX.Charts.getChartAxes(c)
        @test length(axes) == 2
        @test only(XLSX.Charts.getChartGroups(c)).axids == [ax.axid for ax in axes]
        @test XLSX.Charts.getAxisPartner(c, axes[1]).axid == axes[2].axid

        # Every kind builds, and the title options work.
        for (i, kind) in enumerate(keys(XLSX.Charts.C_KINDS))
            k = XLSX.Charts.addChart(ws, kind; anchor="O$(20i):V$(20i + 15)", title=i == 1 ? false : nothing)
            @test isempty(XLSX.Charts.getChartSeries(k))
        end
        @test length(XLSX.Charts.getCharts(ws)) == 1 + length(XLSX.Charts.C_KINDS)

        XLSX.writexlsx(path, xf, overwrite=true)
        @test length(XLSX.Charts.getCharts(XLSX.readxlsx(path))) == 1 + length(XLSX.Charts.C_KINDS)

        SAVE_FILES && save_outfile(xf)
        isfile(path) && rm(path)
    end

    @testset "addChart as a chartsheet" begin
        path = "chart_chartsheet_add.xlsx"
        xf = XLSX.newxlsx("data")

        c = XLSX.Charts.addChart(xf, :line; sheetname="Trend", title="Trend")
        @test c isa XLSX.Charts.Chart
        @test XLSX.sheetnames(xf) == ["data", "Trend"]
        @test XLSX.is_chartsheet(XLSX.get_workbook(xf), "Trend")
        @test c.sheet == "Trend"
        @test isnothing(c.from)
        @test XLSX.Charts.getChartTypes(c) == [:lineChart]

        d = XLSX.Charts.addChart(xf, :pie)
        @test d.sheet == "Chart1"

        # A chartsheet chart's series take their data from a worksheet.
        xf["data"]["B2:B5"] = [8, 7, 3, 9]
        xf["data"]["A2:A5"] = ["A", "B", "C", "D"]
        XLSX.Charts.addSeries(c, "data!B2:B5"; categories="data!A2:A5", name="y")
        s = only(XLSX.Charts.getChartSeries(c))
        @test s.name == "y"
        @test s.values.ref == "data!\$B\$2:\$B\$5"
        @test length(s.values.data) == 4

        # An unqualified reference has no sheet to mean.
        @test_throws XLSX.XLSXError XLSX.Charts.addSeries(c, "B2:B5")

        # A $-fixed reference is just as unqualified.
        @test_throws XLSX.XLSXError XLSX.Charts.addSeries(c, "\$B\$2:\$B\$5")

        # chartEx on a chartsheet: qualified references resolve through the workbook,
        # unqualified ones are refused before anything is written.
        @test_throws XLSX.XLSXError XLSX.Charts.addChartEx(xf, :boxWhisker, "B2:B5"; sheetname="Bad")
        @test !("Bad" in XLSX.sheetnames(xf))
        xf["data"]["C2:C5"] = [5, 6, 2, 4]
        bw = XLSX.Charts.addChartEx(xf, :boxWhisker, "data!B2:B5"; sheetname="Box")
        XLSX.Charts.addSeries(bw, "data!C2:C5")
        @test XLSX.Charts.getChartSeriesCount(bw) == 2
        @test_throws XLSX.XLSXError XLSX.Charts.addSeries(bw, "C2:C5")
        @test XLSX.Charts.getChartSeriesCount(bw) == 2

        XLSX.writexlsx(path, xf, overwrite=true)
        g = XLSX.readxlsx(path)
        @test XLSX.sheetnames(g) == ["data", "Trend", "Chart1", "Box"]
        @test length(XLSX.Charts.getCharts(g)) == 3

        SAVE_FILES && save_outfile(xf)
        isfile(path) && rm(path)
    end

    @testset "a new sheet's content type names its own part" begin
        path = "chart_chartsheet_ct.xlsx"
        cp(joinpath(data_directory, "chart_chartsheet.xlsx"), path; force=true)
        xf = XLSX.openxlsx(path; mode="rw")        # sheetIds 2 and 1, one worksheet file
        ws = XLSX.addsheet!(xf, "More")
        part = XLSX.get_worksheet_internal_file(ws)
        @test ws.sheetId == 3
        @test part == "xl/worksheets/sheet2.xml"
        @test XLSX.content_type_for_part(xf, part) == XLSX.MIME_WORKSHEET

        SAVE_FILES && save_outfile(xf)
        isfile(path) && rm(path)
    end

    @testset "addSeries" begin
        path = "chart_addseries.xlsx"
        xf = XLSX.newxlsx("data")
        ws = xf["data"]
        ws["A1"] = "Region"
        ws["B1"] = "Alpha"
        ws["C1"] = "Beta"
        for (i, r) in enumerate(("North", "South", "East", "West"))
            ws[i+1, 1] = r
            ws[i+1, 2] = 10i
            ws[i+1, 3] = 5i
        end

        c = XLSX.Charts.addChart(ws, :column; anchor="F2:M18")
        XLSX.Charts.addSeries(c, "B2:B5"; categories="A2:A5", name_ref="B1")
        XLSX.Charts.addSeries(c, "C2:C5"; categories="A2:A5", name_ref="C1")

        s = XLSX.Charts.getChartSeries(c)
        @test [x.name for x in s] == ["Alpha", "Beta"]
        @test [x.idx for x in s] == [0, 1]
        @test s[1].values.data == [10, 20, 30, 40]
        @test s[1].categories.data == ["North", "South", "East", "West"]
        @test XLSX.Charts.getSeriesFill(c, 1).value.fgcolor.val == "accent1"
        @test XLSX.Charts.getSeriesFill(c, 2).value.fgcolor.val == "accent2"
        @test XLSX.Charts.getChartData(c).data[2] == [10, 20, 30, 40]

        XLSX.writexlsx(path, xf, overwrite=true)
        @test length(XLSX.Charts.getChartSeries(only(XLSX.Charts.getCharts(XLSX.readxlsx(path))))) == 2

        SAVE_FILES && save_outfile(xf)
        isfile(path) && rm(path)
    end
    @testset "addSeries for every kind" begin
        path = "chart_addseries_kinds.xlsx"
        xf = XLSX.newxlsx("data")
        ws = xf["data"]
        ws["A1"] = "Region"
        ws["B1"] = "Alpha"
        ws["C1"] = "Beta"
        ws["D1"] = "Size"
        for (i, r) in enumerate(("North", "South", "East", "West"))
            ws[i+1, 1] = r
            ws[i+1, 2] = 10i
            ws[i+1, 3] = 5i
            ws[i+1, 4] = i
        end

        row = 1
        for kind in keys(XLSX.Charts.C_KINDS)
            c = XLSX.Charts.addChart(ws, kind; anchor="F$row:M$(row + 15)", title=string(kind))
            row += 16
            if kind === :bubble
                XLSX.Charts.addSeries(c, "C2:C5"; categories="B2:B5", bubble_sizes="D2:D5", name="Bubbles")
            else
                XLSX.Charts.addSeries(c, "B2:B5"; categories="A2:A5", name_ref="B1")
                kind in (:pie, :doughnut) ||
                    XLSX.Charts.addSeries(c, "C2:C5"; categories="A2:A5", name_ref="C1")
            end
            s = XLSX.Charts.getChartSeries(c)
            @test s[1].values.data == [10, 20, 30, 40] || kind === :bubble
            @test !isnothing(s[1].categories)
            @test length(s) == (kind in (:pie, :doughnut, :bubble) ? 1 : 2)
        end

        # Per-point colours on a pie, one per category.
        pie = XLSX.Charts.getCharts(ws)[findfirst(x -> XLSX.Charts.getChartType(x) === :pieChart, XLSX.Charts.getCharts(ws))]
        @test length(XLSX.Charts.getSeriesDataPoints(pie, 1)) == 4

        # Rejected options.
        col = XLSX.Charts.getCharts(ws)[1]
        @test_throws XLSX.XLSXError XLSX.Charts.addSeries(col, "B2:B5"; smooth=true)
        @test_throws XLSX.XLSXError XLSX.Charts.addSeries(col, "B2:B5"; bubble_sizes="D2:D5")
        @test_throws XLSX.XLSXError XLSX.Charts.addSeries(col, "B2:B5"; name="x", name_ref="B1")

        XLSX.writexlsx(path, xf, overwrite=true)
        @test length(XLSX.Charts.getCharts(XLSX.readxlsx(path))) == length(XLSX.Charts.C_KINDS)

        SAVE_FILES && save_outfile(xf)
        isfile(path) && rm(path)
    end

    @testset "addSeries options" begin
        path = "chart_addseries_options.xlsx"
        xf = XLSX.newxlsx("data")
        ws = xf["data"]
        ws["A1"] = "Region"
        ws["B1"] = "Alpha"
        ws["C1"] = "Beta"
        for (i, r) in enumerate(("North", "South", "East", "West"))
            ws[i+1, 1] = r
            ws[i+1, 2] = 10i
            ws[i+1, 3] = 5i
        end
        smooth_of(c, i) = XLSX.Charts._bool_val(last(XLSX.Charts._series_nodes(XLSX.Charts.chart_root(c))[i]), "smooth")

        # Line: markers bring in the lineMarkers pattern; the default stays plain.
        c = XLSX.Charts.addChart(ws, :line; anchor="F1:M16")
        XLSX.Charts.addSeries(c, "B2:B5"; categories="A2:A5", markers=:diamond, smooth=true)
        XLSX.Charts.addSeries(c, "C2:C5"; categories="A2:A5")
        @test XLSX.Charts.getSeriesMarker(c, 1).symbol === :diamond
        @test !isnothing(XLSX.Charts.getSeriesMarker(c, 1).shape)
        @test smooth_of(c, 1) === true
        @test XLSX.Charts.getSeriesMarker(c, 2).symbol === :none
        @test smooth_of(c, 2) === false

        # Scatter: a line, and no markers only when there is a line.
        s = XLSX.Charts.addChart(ws, :scatter; anchor="F18:M33")
        XLSX.Charts.addSeries(s, "C2:C5"; categories="B2:B5", line=true)
        @test XLSX.Charts.getSeriesLine(s, 1).value.fill.kind === :solid
        @test_throws XLSX.XLSXError XLSX.Charts.addSeries(s, "C2:C5"; categories="B2:B5", markers=false)
        XLSX.Charts.addSeries(s, "C2:C5"; categories="B2:B5", markers=false, line=true)
        @test XLSX.Charts.getSeriesMarker(s, 2).symbol === :none

        # Radar markers.
        r = XLSX.Charts.addChart(ws, :radar; anchor="F35:M50")
        XLSX.Charts.addSeries(r, "B2:B5"; categories="A2:A5", markers=:square)
        @test XLSX.Charts.getSeriesMarker(r, 1).symbol === :square

        # Colour, in each accepted form.
        col = XLSX.Charts.addChart(ws, :column; anchor="O1:V16")
        XLSX.Charts.addSeries(col, "B2:B5"; categories="A2:A5", color="FFFF0000")
        XLSX.Charts.addSeries(col, "C2:C5"; categories="A2:A5", color="accent6")
        @test XLSX.Charts.getSeriesFill(col, 1).value.fgcolor.rgb == "FF0000"

        # Rejected.
        pie = XLSX.Charts.addChart(ws, :pie; anchor="O18:V33")
        @test_throws XLSX.XLSXError XLSX.Charts.addSeries(pie, "B2:B5"; color="FF0000")
        @test_throws XLSX.XLSXError XLSX.Charts.addSeries(c, "B2:B5"; markers=:nonsense)
        @test_throws XLSX.XLSXError XLSX.Charts.addSeries(c, "B2:B5"; smooth="yes")
        @test_throws XLSX.XLSXError XLSX.Charts.addSeries(col, "B2:B5"; color="notacolor")

        XLSX.writexlsx(path, xf, overwrite=true)
        SAVE_FILES && save_outfile(xf)
        isfile(path) && rm(path)
    end

    @testset "addChartEx" begin
        path = "chart_addchartex.xlsx"
        xf = XLSX.newxlsx("data")
        ws = xf["data"]
        ws["A1"] = "Region"
        ws["B1"] = "Item"
        ws["C1"] = "Amount"
        ws["D1"] = "Beta"
        for (i, row) in enumerate((("Europe", "France", 30, 12), ("Europe", "Spain", 20, 15),
            ("Asia", "Japan", 50, 17), ("Asia", "India", 35, 19),
            ("Americas", "USA", 60, 21)))
            for (j, v) in enumerate(row); ws[i+1, j] = v; end
        end

        row = 1
        for kind in keys(XLSX.Charts.CX_KINDS)
            cats = kind in (:treemap, :sunburst) ? "A2:B6" :
                   kind === :histogram ? nothing :
                   kind === :boxWhisker ? "A2:A6" : "B2:B6"
            c = XLSX.Charts.addChartEx(ws, kind, "C2:C6"; anchor="F$row:M$(row + 15)",
                categories=cats, name_ref="C1")
            row += 16
            @test c isa XLSX.Charts.ChartEx
            @test XLSX.Charts.getChartType(c) === kind
            @test XLSX.Charts.getChartSeriesCount(c) == 1
            kind === :pareto && @test XLSX.Charts.getSeriesLayout(c, 1; pareto=true) === :paretoLine
            @test XLSX.Charts.getSeriesName(c, 1) == "Amount"
        end

        charts = XLSX.Charts.getCharts(ws)
        bw = charts[findfirst(c -> XLSX.Charts.getChartType(c) === :boxWhisker, charts)]
        XLSX.Charts.addSeries(bw, "D2:D6"; name_ref="D1")
        @test XLSX.Charts.getChartSeriesCount(bw) == 2

        wf = charts[findfirst(c -> XLSX.Charts.getChartType(c) === :waterfall, charts)]
        @test_throws XLSX.XLSXError XLSX.Charts.addSeries(wf, "D2:D6")
        @test_throws XLSX.XLSXError XLSX.Charts.addChartEx(ws, :histogram, "C2:C6"; anchor="O1:V16", categories="B2:B6")
        @test_throws XLSX.XLSXError XLSX.Charts.addChartEx(ws, :treemap, "C2:C6"; anchor="O1:V16")

        XLSX.writexlsx(path, xf, overwrite=true)
        @test length(XLSX.Charts.getCharts(XLSX.readxlsx(path))) == length(XLSX.Charts.CX_KINDS)

        SAVE_FILES && save_outfile(xf)
        isfile(path) && rm(path)
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
    @testset "a failed addChartEx leaves no chartsheet behind" begin
        xf = XLSX.newxlsx("data")
        @test_throws XLSX.XLSXError XLSX.Charts.addChartEx(xf, :waterfall, "B2:B6")          # unqualified
        @test_throws XLSX.XLSXError XLSX.Charts.addChartEx(xf, :waterfall, "Nowhere!B2:B6")  # no such sheet
        @test XLSX.sheetnames(xf) == ["data"]
        @test isempty(XLSX.Charts.getCharts(xf))
        SAVE_FILES && save_outfile(xf)
    end

    @testset "chart values compare by content" begin
        f = XLSX.readxlsx(joinpath(data_directory, "chart_gaps.xlsx"))          # has error cells
        c = XLSX.Charts.getCharts(f)[1]
        a, b = XLSX.Charts.getChartSeries(c), XLSX.Charts.getChartSeries(c)
        @test a == b                                                             # missing in the cache
        @test hash(a) == hash(b)
        @test length(Set(vcat(a, b))) == length(a)

        fa = XLSX.readxlsx(joinpath(data_directory, "chart_appearance.xlsx"))
        ca = only(filter(x -> x isa XLSX.Charts.Chart, XLSX.Charts.getCharts(fa)))
        @test XLSX.Charts.getChartAxes(ca) == XLSX.Charts.getChartAxes(ca)
        @test XLSX.Charts.getSeriesShapeProps(ca, 1) == XLSX.Charts.getSeriesShapeProps(ca, 1)
        @test XLSX.Charts.getSeriesMarker(ca, 1) == XLSX.Charts.getSeriesMarker(ca, 1)
    end

    @testset "chart handles compare and hash by part" begin
        f = XLSX.readxlsx(joinpath(data_directory, "chart_basic.xlsx"))
        a = XLSX.getChart(f["Data"], "chart1")
        b = only(XLSX.getCharts(f))
        @test a == b
        @test hash(a) == hash(b)
        @test length(Set([a, b])) == 1

        g = XLSX.readxlsx(joinpath(data_directory, "chartex_kinds.xlsx"))
        x, y = XLSX.getCharts(g)[1], XLSX.getCharts(g)[1]
        @test x == y
        @test hash(x) == hash(y)
        @test x != XLSX.getCharts(g)[2]
        # the same part in a different file is a different chart
        @test x != XLSX.getCharts(XLSX.readxlsx(joinpath(data_directory, "chartex_kinds.xlsx")))[1]
    end

    @testset "ChartEx show" begin
        x = XLSX.getCharts(XLSX.readxlsx(joinpath(data_directory, "chartex_kinds.xlsx")))[1]
        s = sprint(show, MIME"text/plain"(), x)
        @test startswith(s, "XLSX.Charts.ChartEx \"chartEx1\"")
        @test occursin("type: waterfall", s)
        @test occursin("series: 1", s)
        @test occursin("refs: ", s)
        @test !isempty(sprint(show, x))
    end

    @testset "a value read before a write differs from one read after" begin
        path = "chart_equality_write.xlsx"
        cp(joinpath(data_directory, "chart_basic.xlsx"), path; force=true)
        xf = XLSX.openxlsx(path; mode="rw")
        c = XLSX.Charts.getChart(xf, "chart1")
        before = XLSX.Charts.getSeriesShapeProps(c, 1)
        XLSX.Charts.setSeriesFill(c, 1, "red")
        @test XLSX.Charts.getSeriesShapeProps(c, 1) != before
        isfile(path) && rm(path)
        SAVE_FILES && save_outfile(xf)
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

    @testset "XLSX-level chart names" begin
        for n in (:AbstractChart, :Chart, :ChartEx, :ChartRef, :ChartSeries,
            :getChartSchema, :getChartType, :getCharts, :getChart, :getChartData, :getChartRanges)
            @test getglobal(XLSX, n) === getglobal(XLSX.Charts, n)
            VERSION >= v"1.11" && @test Base.ispublic(XLSX, n)
        end
    end

    @testset "v0.12 API still works, deprecated" begin
        c = XLSX.getChart(XLSX.readxlsx(joinpath(data_directory, "chart_basic.xlsx"))["Data"], "chart1")
        @test (@test_deprecated c.title) == XLSX.Charts.getChartTitle(c)
        @test (@test_deprecated c.charttypes) == XLSX.Charts.getChartTypes(c)
        @test (@test_deprecated c.series) == XLSX.Charts.getChartSeries(c)
        @test (@test_deprecated c.series[1].values.data) == [10.0, 20.0, 15.0, 5.0]
        @test (@test_deprecated XLSX.chartType(c)) === XLSX.Charts.getChartType(c)
        @test (@test_deprecated XLSX.chartSchema(c)) === XLSX.Charts.getChartSchema(c)
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

end