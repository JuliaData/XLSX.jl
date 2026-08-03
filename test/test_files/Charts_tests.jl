
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

        charts = XLSX.getCharts(f)
        @test length(charts) == 1

        c = charts[1]
        @test c.name == "chart1"
        @test c.path == "xl/charts/chart1.xml"
        @test c.sheet == "Data"
        @test c.title == "Revenue by Region"
        @test c.charttypes == [:barChart]
        @test !isnothing(c.rId)

        # Anchored to cells, so both markers parse. The exact cells depend on
        # where the chart was dropped, so only their form is asserted.
        @test !isnothing(c.from)
        @test !isnothing(c.to)
        @test XLSX.CellRef(c.from) isa XLSX.CellRef
        @test XLSX.CellRef(c.to) isa XLSX.CellRef

        @test length(c.series) == 2
        @test [s.name for s in c.series] == ["2024", "2025"]
        @test [s.order for s in c.series] == [0, 1]
        @test [s.idx for s in c.series] == [0, 1]
        @test all(s -> s.charttype == :barChart, c.series)
        @test all(s -> isnothing(s.bubble_sizes), c.series)

        s1, s2 = c.series
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
        dt = XLSX.getChartData(c)
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
        c = XLSX.getCharts(f)[1]

        @test XLSX.getChart(f, "chart1").path == c.path
        @test XLSX.getChart(f, "chart1.xml").path == c.path
        @test XLSX.getChart(f, "xl/charts/chart1.xml").path == c.path
        @test XLSX.getChart(f, c.rId).path == c.path
        @test_throws XLSX.XLSXError XLSX.getChart(f, "nosuchchart")

        # Worksheet-scoped discovery walks the sheet's own drawing.
        ws_charts = XLSX.getCharts(f["Data"])
        @test length(ws_charts) == 1
        @test ws_charts[1].path == c.path

        @test XLSX.getChartData(f, "chart1").column_labels == XLSX.getChartData(c).column_labels
    end

    @testset "metadata only" begin  # chart_basic.xlsx
        f = XLSX.readxlsx(joinpath(data_directory, "chart_basic.xlsx"))
        c = XLSX.getCharts(f; cache=false)[1]

        # Metadata survives; values do not.
        @test c.title == "Revenue by Region"
        @test [s.name for s in c.series] == ["2024", "2025"]   # names come from c:tx regardless
        @test c.series[1].values.ref == "Data!\$B\$2:\$B\$5"
        @test c.series[1].values.format_code == "General"
        @test c.series[1].values.ptCount == 4
        @test isempty(c.series[1].values.data)
        @test isempty(c.series[1].categories.data)

        # A table of `missing` would be a silent lie.
        @test_throws XLSX.XLSXError XLSX.getChartData(c)
        @test_throws XLSX.XLSXError XLSX.getChartData(f, "chart1"; cache=false)
    end

    @testset "gaps and errors" begin  # chart_gaps.xlsx
        f = XLSX.readxlsx(joinpath(data_directory, "chart_gaps.xlsx"))
        c = XLSX.getCharts(f)[1]

        @test c.charttypes == [:lineChart]
        @test length(c.series) == 1

        v = c.series[1].values
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
        @test c.series[1].categories.data == ["a", "b", "c", "d", "e"]
    end

    @testset "combo chart" begin  # chart_combo.xlsx
        f = XLSX.readxlsx(joinpath(data_directory, "chart_combo.xlsx"))
        c = XLSX.getCharts(f)[1]

        # Two group elements in one plotArea. Excel writes them in schema order,
        # so compare as a set rather than pinning the sequence.
        @test length(c.charttypes) == 2
        @test issetequal(c.charttypes, [:barChart, :lineChart])

        # Series are collected across both groups, with `order` continuing rather
        # than restarting per group.
        @test length(c.series) == 2
        @test sort([s.order for s in c.series]) == [0, 1]
        @test issetequal([s.charttype for s in c.series], [:barChart, :lineChart])

        dt = XLSX.getChartData(c)
        @test length(dt.column_labels) == 3   # shared categories + two series
    end

    @testset "scatter chart" begin  # chart_scatter.xlsx
        f = XLSX.readxlsx(joinpath(data_directory, "chart_scatter.xlsx"))
        c = XLSX.getCharts(f)[1]

        @test c.charttypes == [:scatterChart]
        @test length(c.series) == 2

        # c:xVal folds into `categories`, c:yVal into `values`.
        s1, s2 = c.series
        @test s1.categories.data == [1.0, 2.0, 3.0, 4.0]
        @test s1.values.data == [10.0, 20.0, 30.0, 40.0]
        @test s2.categories.data == [2.0, 4.0, 6.0, 8.0]
        @test s2.values.data == [5.0, 15.0, 25.0, 35.0]

        # Differing x references, so no shared `categories` column: each series
        # brings its own.
        @test s1.categories.ref != s2.categories.ref
        dt = XLSX.getChartData(c)
        @test length(dt.column_labels) == 4
        @test endswith(String(dt.column_labels[1]), "_x")
        @test endswith(String(dt.column_labels[3]), "_x")
    end

    @testset "bubble chart and document order" begin  # chart_bubble.xlsx
        f = XLSX.readxlsx(joinpath(data_directory, "chart_bubble.xlsx"))
        charts = XLSX.getCharts(f)
        @test length(charts) == 2

        # FIXTURE: the pie chart was inserted first but positioned to the right
        # of the bubble chart. Anchor order is creation order, not reading order,
        # so the pie comes back first.
        @test charts[1].charttypes == [:pieChart]
        @test charts[2].charttypes == [:bubbleChart]

        bub = charts[2]
        s = bub.series[1]
        @test !isnothing(s.bubble_sizes)
        @test length(s.bubble_sizes.data) == length(s.values.data)

        dt = XLSX.getChartData(bub)
        @test any(l -> endswith(String(l), "_size"), dt.column_labels)
    end

    @testset "multi-level categories" begin  # chart_multilevel.xlsx
        f = XLSX.readxlsx(joinpath(data_directory, "chart_multilevel.xlsx"))
        c = XLSX.getCharts(f)[1]

        cat = c.series[1].categories
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

        dt = XLSX.getChartData(c)
        @test dt.column_labels[1] == :categories_1
        @test dt.column_labels[2] == :categories_2
        @test length(dt.column_labels) == 3   # two levels + one series
    end

    @testset "chartsheet" begin  # chart_chartsheet.xlsx
        f = XLSX.readxlsx(joinpath(data_directory, "chart_chartsheet.xlsx"))
        charts = XLSX.getCharts(f)
        @test length(charts) == 1

        c = charts[1]
        @test c.sheet == "TheChart"

        # A chartsheet drawing uses an absoluteAnchor with pos/ext, not cell
        # markers, so there is nothing to resolve into a CellRef.
        @test isnothing(c.from)
        @test isnothing(c.to)

        # The data still reads exactly as when the chart was on a worksheet.
        @test length(c.series) == 2
        @test c.series[1].values.data == [10.0, 20.0, 15.0, 5.0]
    end

    @testset "strict OOXML" begin  # chart_strict.xlsx
        f = XLSX.readxlsx(joinpath(data_directory, "chart_strict.xlsx"))
        charts = XLSX.getCharts(f)

        # Strict files are remapped to transitional at read, so the transitional
        # chart relationship type still resolves and the chart is found.
        @test length(charts) == 1
        c = charts[1]
        @test c.sheet == "Data"
        @test c.charttypes == [:barChart]
        @test c.series[1].categories.data == ["North", "South", "East", "West"]
        @test c.series[1].values.data == [10.0, 20.0, 15.0, 5.0]
    end

    @testset "chartEx is not read" begin  # chart_ex.xlsx
        f = XLSX.readxlsx(joinpath(data_directory, "chart_ex.xlsx"))

        # The part exists but uses the cx: schema, so nothing is returned - and
        # the user is told why rather than left with a bare empty vector.
        charts = @test_logs (:warn, r"chartEx") match_mode = :any XLSX.getCharts(f)
        @test isempty(charts)
    end

    @testset "external reference" begin  # chart_external.xlsx
        f = XLSX.readxlsx(joinpath(data_directory, "chart_external.xlsx"))
        c = XLSX.getCharts(f)[1]

        @test c.charttypes == [:lineChart]
        @test length(c.series) == 2

        # Source lives in another workbook: the `[1]` indexes the workbook's
        # external references.
        @test c.series[1].values.ref == "[1]Feuil1!\$A\$1:\$A\$10"
        @test c.series[2].values.ref == "[1]Feuil1!\$B\$1:\$B\$10"

        # No c:cat at all - Excel plots against an implicit index and caches
        # nothing for it - and no c:tx, so the series are unnamed.
        @test all(s -> isnothing(s.categories), c.series)
        @test all(s -> isnothing(s.name), c.series)

        @test c.series[1].values.data == collect(1.0:10.0)
        @test c.series[2].values.data == [1.0, 10.0, 5.0, 2.0, 3.0, 45.0, 6.0, 8.0, 7.0, 2.0]

        # No categories, so no leading column, and positional series labels.
        dt = XLSX.getChartData(c)
        @test dt.column_labels == [:Series1, :Series2]
        @test length(dt.data[1]) == 10

        # Materialising the reference resolves `[1]` through xl/externalLinks and
        # its relationships (regression test for get_external_workbook_path,
        # which previously ignored the externalBook r:id).
        cm = XLSX.getCharts(f; get_external_refs=true)[1]
        @test cm.series[1].values.ref == "[Test2.xlsx]Feuil1!\$A\$1:\$A\$10"
        @test cm.series[1].values.data == c.series[1].values.data   # unchanged by materialising
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

        c = XLSX.getCharts(f)[1]
        @test c.series[1].values.data == external_column(f, "A")
        @test c.series[2].values.data == external_column(f, "B")
    end

    @testset "round trip" begin  # chart_basic.xlsx
        # Reading charts must not disturb them: the chart part should survive a
        # save byte for byte, and read back identically.
        original = joinpath(data_directory, "chart_basic.xlsx")
        f = XLSX.openxlsx(original; mode="rw")
        before = XLSX.getChartData(XLSX.getCharts(f)[1])

        tmp = joinpath(tempdir(), "chart_basic_roundtrip.xlsx")
        isfile(tmp) && rm(tmp; force=true)
        XLSX.writexlsx(tmp, f; overwrite=true)

        part = "xl/charts/chart1.xml"
        # XML.jl pretty-prints where Excel writes a single line; the CRLF after the
        # declaration differs too. Both are cosmetic — XML ignores inter-tag
        # whitespace — so compare with it collapsed rather than byte for byte.
        strip_ws(s) = replace(s, r">\s+<" => "><")
        @test strip_ws(String(zip_readentry(ZipReader(read(original)), part))) ==
              strip_ws(String(zip_readentry(ZipReader(read(tmp)), part)))

        g = XLSX.readxlsx(tmp)
        c = XLSX.getCharts(g)[1]
        @test c.sheet == "Data"
        @test c.title == "Revenue by Region"
        after = XLSX.getChartData(c)
        @test after.column_labels == before.column_labels
        @test after.data == before.data

        SAVE_FILES && save_outfile(tmp)
        rm(tmp; force=true)
    end

end