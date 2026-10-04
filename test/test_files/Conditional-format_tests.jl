@testset "Conditional Formats" begin

    @testset "DataBar" begin

        f = XLSX.newxlsx()
        s = f[1]
        for i in 1:5, j in 1:5
            s[i, j] = i + j
        end

        @test_throws XLSX.XLSXError XLSX.setConditionalFormat(s, "A1:A5", :dataBar; notAKeyword="x")
        @test_throws XLSX.XLSXError XLSX.setConditionalFormat(s, "A1:A5", :dataBar; priority=1) # priority is not a valid keyword argument
        @test_throws XLSX.XLSXError XLSX.setConditionalFormat(s, 1:5, 1, :dataBar; priority=1)
        @test_throws XLSX.XLSXError XLSX.setConditionalFormat(s, 1, 1:3:7, :dataBar) # out of range
        @test XLSX.setConditionalFormat(s, 2, :, :dataBar; databar="greengrad") == 0
        @test XLSX.setConditionalFormat(s, 3, 1:5, :dataBar;
            min_type="least",
            min_val="green", #should be ignored because type=least
            max_type="percentile",
            max_val="50",
        ) == 0
        @test XLSX.setConditionalFormat(s, "Sheet1!A4:E4", :dataBar;
            min_type="automatic",
            max_type="automatic",
        ) == 0
        @test XLSX.setConditionalFormat(f, "Sheet1!A5:E5", :dataBar;
            min_type="num",
            min_val="\$A\$1",
            max_type="formula",
            max_val="\$A\$2"
        ) == 0
        @test sort!(XLSX.getConditionalFormats(s), by=x -> x.second.priority, rev=true) == [
            XLSX.CellRange("A5:E5") => (type="dataBar", priority=4),
            XLSX.CellRange("A4:E4") => (type="dataBar", priority=3),
            XLSX.CellRange("A3:E3") => (type="dataBar", priority=2),
            XLSX.CellRange("A2:E2") => (type="dataBar", priority=1)
        ]

        @test XLSX.setConditionalFormat(s, "A1", :dataBar) == 0
        @test XLSX.setConditionalFormat(s, "A1:C3", :dataBar) == 0
        @test XLSX.setConditionalFormat(s, "Sheet1!A1", :dataBar) == 0
        @test XLSX.setConditionalFormat(s, "Sheet1!A1:A2", :dataBar) == 0
        @test XLSX.setConditionalFormat(s, "Sheet1!1:2", :dataBar) == 0
        @test XLSX.setConditionalFormat(s, "2:4", :dataBar) == 0
        @test XLSX.setConditionalFormat(s, "A:C", :dataBar) == 0
        @test XLSX.setConditionalFormat(s, "Sheet1!A:C", :dataBar) == 0
        @test XLSX.setConditionalFormat(f, "Sheet1!1:2", :dataBar) == 0
        @test XLSX.setConditionalFormat(f, "Sheet1!A:C", :dataBar) == 0
        @test XLSX.setConditionalFormat(s, :, 1:3, :dataBar) == 0
        @test XLSX.setConditionalFormat(s, 1:3, :, :dataBar) == 0
        @test XLSX.setConditionalFormat(s, "2:4", :dataBar) == 0
        @test XLSX.setConditionalFormat(s, "A:C", :dataBar) == 0
        @test XLSX.setConditionalFormat(s, "Sheet1!A:C", :dataBar) == 0
        @test XLSX.setConditionalFormat(s, :, :dataBar) == 0
        @test XLSX.setConditionalFormat(s, :, :, :dataBar) == 0
        @test length(XLSX.getConditionalFormats(s)) == 21
        @test sort!(XLSX.getConditionalFormats(s), by=x -> x.second.priority, rev=true) == [
            XLSX.CellRange("A1:E5") => (type="dataBar", priority=21),
            XLSX.CellRange("A1:E5") => (type="dataBar", priority=20),
            XLSX.CellRange("A1:C5") => (type="dataBar", priority=19),
            XLSX.CellRange("A1:C5") => (type="dataBar", priority=18),
            XLSX.CellRange("A2:E4") => (type="dataBar", priority=17),
            XLSX.CellRange("A1:E3") => (type="dataBar", priority=16),
            XLSX.CellRange("A1:C5") => (type="dataBar", priority=15),
            XLSX.CellRange("A1:C5") => (type="dataBar", priority=14),
            XLSX.CellRange("A1:E2") => (type="dataBar", priority=13),
            XLSX.CellRange("A1:C5") => (type="dataBar", priority=12),
            XLSX.CellRange("A1:C5") => (type="dataBar", priority=11),
            XLSX.CellRange("A2:E4") => (type="dataBar", priority=10),
            XLSX.CellRange("A1:E2") => (type="dataBar", priority=9),
            XLSX.CellRange("A1:A2") => (type="dataBar", priority=8),
            XLSX.CellRange("A1:A1") => (type="dataBar", priority=7),
            XLSX.CellRange("A1:C3") => (type="dataBar", priority=6),
            XLSX.CellRange("A1:A1") => (type="dataBar", priority=5),
            XLSX.CellRange("A5:E5") => (type="dataBar", priority=4),
            XLSX.CellRange("A4:E4") => (type="dataBar", priority=3),
            XLSX.CellRange("A3:E3") => (type="dataBar", priority=2),
            XLSX.CellRange("A2:E2") => (type="dataBar", priority=1)
        ]

        SAVE_FILES && save_outfile(f)

        f = XLSX.newxlsx()
        s = f[1]
        for i in 1:5, j in 1:5
            s[i, j] = i + j
        end
        @test XLSX.setConditionalFormat(s, "A1:A5", :dataBar) == 0
        @test XLSX.setConditionalFormat(s, :, 2, :dataBar; databar="orange") == 0
        @test XLSX.setConditionalFormat(s, "Sheet1!E:E", :dataBar; databar="purplegrad") == 0
        @test XLSX.setConditionalFormat(s, 1:5, 3:4, :dataBar;
            borders="false",
            min_type="percentile",
            min_val="25",
            max_type="percentile",
            max_val="75"
        ) == 0
        @test sort!(XLSX.getConditionalFormats(s), by=x -> x.second.priority, rev=true) == [
            XLSX.CellRange("C1:D5") => (type="dataBar", priority=4),
            XLSX.CellRange("E1:E5") => (type="dataBar", priority=3),
            XLSX.CellRange("B1:B5") => (type="dataBar", priority=2),
            XLSX.CellRange("A1:A5") => (type="dataBar", priority=1)
        ]
        SAVE_FILES && save_outfile(f)

        f = XLSX.newxlsx()
        s = f[1]
        for i in 1:5, j in 1:5
            s[i, j] = i + j
        end

        @test XLSX.setConditionalFormat(s, :, 1:4, :dataBar;
            databar="red",
            borders="true",
            fill_col="blue",
            border_col="yellow",
            neg_fill_col="magenta",
            neg_border_col="green",
            axis_col="cyan"
        ) == 0
        SAVE_FILES && save_outfile(f)

        f = XLSX.newxlsx()
        s = f[1]
        for i in 1:5, j in 1:5
            s[i, j] = i + j
        end
        XLSX.addDefinedName(s, "myRange", "A1:B5")
        @test XLSX.setConditionalFormat(s, "myRange", :dataBar;
            showVal="false",
            direction="leftToRight",
            borders="true",
            sameNegBorders="false"
        ) == 0
        XLSX.addDefinedName(s, "myNCRange", "C1:C5,D1:D5")
        @test XLSX.setConditionalFormat(s, "myNCRange", :dataBar; showVal="false",
            direction="leftToRight", borders="true", sameNegBorders="false") == 0
        @test_throws XLSX.XLSXError XLSX.setConditionalFormat(s, "A1:A2", :dataBar;
            databar="rainbow"
        )
        SAVE_FILES && save_outfile(f)

        f = XLSX.newxlsx()
        s = f[1]
        for i in 1:5, j in 1:12
            s[i, j] = i + j
        end
        s[1, 13] = 5

        @test XLSX.setConditionalFormat(s, :, 1, :dataBar;
            databar="orange",
            sameNegFill="true",
            sameNegBorders="true"
        ) == 0
        @test XLSX.setConditionalFormat(s, :, 2, :dataBar;
            databar="orange",
            axis_pos="none"
        ) == 0
        @test XLSX.setConditionalFormat(s, :, 3, :dataBar;
            databar="orange",
            axis_pos="middle"
        ) == 0
        @test XLSX.setConditionalFormat(s, :, 4, :dataBar;
            databar="orange",
            min_type="num",
            min_val="Sheet1!\$M\$1"
        ) == 0
        @test XLSX.setConditionalFormat(s, :, 5, :dataBar;
            databar="orange",
            showVal="false",
            direction="rightToLeft",
            borders="true",
            sameNegBorders="false",
            sameNegFill="false"
        ) == 0

        @test_throws XLSX.XLSXError XLSX.setConditionalFormat(s, :, 4, :dataBar;
            axis_pos="nonsense",
            databar="orange",
            min_type="num",
            min_val="Sheet1!\$M\$1"
        )
        @test_throws XLSX.XLSXError XLSX.setConditionalFormat(s, :, 4, :dataBar;
            borders="nonsense",
            databar="orange",
            min_type="num",
            min_val="Sheet1!\$M\$1"
        )
        @test_throws XLSX.XLSXError XLSX.setConditionalFormat(s, :, 4, :dataBar;
            fill_col="nonsense",
            databar="orange",
            min_type="num",
            min_val="Sheet1!\$M\$1"
        )
        @test_throws XLSX.XLSXError XLSX.setConditionalFormat(s, :, 4, :dataBar;
            sameNegFill="nonsense",
            databar="orange",
            min_type="num",
            min_val="Sheet1!\$M\$1"
        )
        @test_throws XLSX.XLSXError XLSX.setConditionalFormat(s, :, 4, :dataBar;
            databar="orange",
            min_type="num",
            min_val="Sheet2!\$M\$1"
        )
        SAVE_FILES && save_outfile(f)

        f = XLSX.newxlsx()
        s = f[1]
        for i in 1:5, j in 1:12
            s[i, j] = i + j
        end
        for (j, k) in enumerate(keys(XLSX.databars))
            @test XLSX.setConditionalFormat(s, :, j, :dataBar; databar=k) == 0
        end
        SAVE_FILES && save_outfile(f)
    end

    @testset "colorScale" begin

        f = XLSX.newxlsx()
        s = f[1]
        for i in 1:5, j in 1:5
            s[i, j] = i + j
        end

        @test_throws XLSX.XLSXError XLSX.setConditionalFormat(s, "A1:A5", :colorScale; notAKeyword="x")
        @test_throws XLSX.XLSXError XLSX.setConditionalFormat(s, "A1,A3", :wrongOne)
        @test_throws XLSX.XLSXError XLSX.setConditionalFormat(s, 1, 2, :wrongOne)

        @test XLSX.setConditionalFormat(s, "A1,A3", :colorScale) == 0
        @test XLSX.setConditionalFormat(s, [1], 1, :colorScale) == 0
        @test_throws XLSX.XLSXError XLSX.setConditionalFormat(s, 1, 1:3:7, :colorScale) # out of range
        @test XLSX.setConditionalFormat(s, "1:1", :colorScale) == 0
        @test XLSX.setConditionalFormat(s, 2, :, :colorScale; colorscale="redwhiteblue") == 0
        @test XLSX.setConditionalFormat(s, 3, 1:5, :colorScale;
            min_type="min",
            min_col="green",
            mid_type="percentile",
            mid_val="50",
            mid_col="red",
            max_type="max",
            max_col="blue"
        ) == 0
        @test XLSX.setConditionalFormat(s, "Sheet1!A4:E4", :colorScale;
            min_type="min",
            min_col="tomato",
            max_type="max",
            max_col="gold4"
        ) == 0
        @test XLSX.setConditionalFormat(f, "Sheet1!A5:E5", :colorScale;
            min_type="min",
            min_col="yellow",
            max_type="max",
            max_col="darkgreen"
        ) == 0
        @test sort!(XLSX.getConditionalFormats(s), by=x -> x.second.priority, rev=true) == [
            XLSX.CellRange("A5:E5") => (type="colorScale", priority=7),
            XLSX.CellRange("A4:E4") => (type="colorScale", priority=6),
            XLSX.CellRange("A3:E3") => (type="colorScale", priority=5),
            XLSX.CellRange("A2:E2") => (type="colorScale", priority=4),
            XLSX.CellRange("A1:E1") => (type="colorScale", priority=3),
            XLSX.CellRange("A1:A1") => (type="colorScale", priority=2),
            XLSX.NonContiguousRange("Sheet1!A1,Sheet1!A3") => (type="colorScale", priority=1)
        ]

        @test XLSX.setConditionalFormat(s, "A1", :colorScale) == 0
        @test XLSX.setConditionalFormat(s, "A1:C3", :colorScale) == 0
        @test XLSX.setConditionalFormat(s, "Sheet1!A1", :colorScale) == 0
        @test XLSX.setConditionalFormat(s, "Sheet1!A1:A2", :colorScale) == 0
        @test XLSX.setConditionalFormat(s, "Sheet1!1:2", :colorScale) == 0
        @test XLSX.setConditionalFormat(s, "2:4", :colorScale) == 0
        @test XLSX.setConditionalFormat(s, "A:C", :colorScale) == 0
        @test XLSX.setConditionalFormat(s, "Sheet1!A:C", :colorScale) == 0
        @test XLSX.setConditionalFormat(f, "Sheet1!1:2", :colorScale) == 0
        @test XLSX.setConditionalFormat(f, "Sheet1!A:C", :colorScale) == 0
        @test XLSX.setConditionalFormat(s, :, 1:3, :colorScale) == 0
        @test XLSX.setConditionalFormat(s, 1:3, :, :colorScale) == 0
        @test XLSX.setConditionalFormat(s, "2:4", :colorScale) == 0
        @test XLSX.setConditionalFormat(s, "A:C", :colorScale) == 0
        @test XLSX.setConditionalFormat(s, "Sheet1!A:C", :colorScale) == 0
        @test XLSX.setConditionalFormat(s, :, :colorScale) == 0
        @test XLSX.setConditionalFormat(s, :, :, :colorScale) == 0
        @test length(XLSX.getConditionalFormats(s)) == 24
        @test sort!(XLSX.getConditionalFormats(s), by=x -> x.second.priority, rev=true) == [
            XLSX.CellRange("A1:E5") => (type="colorScale", priority=24),
            XLSX.CellRange("A1:E5") => (type="colorScale", priority=23),
            XLSX.CellRange("A1:C5") => (type="colorScale", priority=22),
            XLSX.CellRange("A1:C5") => (type="colorScale", priority=21),
            XLSX.CellRange("A2:E4") => (type="colorScale", priority=20),
            XLSX.CellRange("A1:E3") => (type="colorScale", priority=19),
            XLSX.CellRange("A1:C5") => (type="colorScale", priority=18),
            XLSX.CellRange("A1:C5") => (type="colorScale", priority=17),
            XLSX.CellRange("A1:E2") => (type="colorScale", priority=16),
            XLSX.CellRange("A1:C5") => (type="colorScale", priority=15),
            XLSX.CellRange("A1:C5") => (type="colorScale", priority=14),
            XLSX.CellRange("A2:E4") => (type="colorScale", priority=13),
            XLSX.CellRange("A1:E2") => (type="colorScale", priority=12),
            XLSX.CellRange("A1:A2") => (type="colorScale", priority=11),
            XLSX.CellRange("A1:A1") => (type="colorScale", priority=10),
            XLSX.CellRange("A1:C3") => (type="colorScale", priority=9),
            XLSX.CellRange("A1:A1") => (type="colorScale", priority=8),
            XLSX.CellRange("A5:E5") => (type="colorScale", priority=7),
            XLSX.CellRange("A4:E4") => (type="colorScale", priority=6),
            XLSX.CellRange("A3:E3") => (type="colorScale", priority=5),
            XLSX.CellRange("A2:E2") => (type="colorScale", priority=4),
            XLSX.CellRange("A1:E1") => (type="colorScale", priority=3),
            XLSX.CellRange("A1:A1") => (type="colorScale", priority=2),
            XLSX.NonContiguousRange("Sheet1!A1,Sheet1!A3") => (type="colorScale", priority=1)
        ]

        SAVE_FILES && save_outfile(f)

        f = XLSX.newxlsx()
        s = f[1]
        for i in 1:5, j in 1:5
            s[i, j] = i + j
        end
        @test XLSX.setConditionalFormat(s, "A1:A5", :colorScale) == 0
        @test XLSX.setConditionalFormat(s, :, 2, :colorScale; colorscale="redwhiteblue") == 0
        @test XLSX.setConditionalFormat(s, "Sheet1!E:E", :colorScale; colorscale="greenwhitered") == 0
        @test XLSX.setConditionalFormat(s, 1:5, 3:4, :colorScale;
            min_type="min",
            min_col="green",
            mid_type="percentile",
            mid_val="50",
            mid_col="red",
            max_type="max",
            max_col="blue"
        ) == 0
        @test sort!(XLSX.getConditionalFormats(s), by=x -> x.second.priority, rev=true) == [
            XLSX.CellRange("C1:D5") => (type="colorScale", priority=4),
            XLSX.CellRange("E1:E5") => (type="colorScale", priority=3),
            XLSX.CellRange("B1:B5") => (type="colorScale", priority=2),
            XLSX.CellRange("A1:A5") => (type="colorScale", priority=1)
        ]
        SAVE_FILES && save_outfile(f)

        f = XLSX.newxlsx()
        s = f[1]
        for i in 1:5, j in 1:5
            s[i, j] = i + j
        end

        @test XLSX.setConditionalFormat(s, :, 1:4, :colorScale;
            min_type="min",
            min_col="green",
            mid_type="percentile",
            mid_val="\$E\$4",
            mid_col="red",
            max_type="max",
            max_col="blue"
        ) == 0
        @test XLSX.setConditionalFormat(s, :, 5, :colorScale;
            min_type="min",
            min_col="green",
            mid_type="percentile",
            mid_val="Sheet1!\$E\$4",
            mid_col="red",
            max_type="max",
            max_col="blue"
        ) == 0
        @test_throws XLSX.XLSXError XLSX.setConditionalFormat(s, :, 5, :colorScale;
            min_type="min",
            min_col="green",
            mid_type="percentile",
            mid_val="Sheet2!\$E\$4",
            mid_col="red",
            max_type="max",
            max_col="blue"
        )
        SAVE_FILES && save_outfile(f)

        f = XLSX.newxlsx()
        s = f[1]
        for i in 1:5, j in 1:5
            s[i, j] = i + j
        end
        XLSX.addDefinedName(s, "myRange", "A1:B5")
        @test XLSX.setConditionalFormat(s, "myRange", :colorScale;
            min_type="min",
            min_col="green",
            mid_type="percentile",
            mid_val="50",
            mid_col="red",
            max_type="max",
            max_col="blue"
        ) == 0
        XLSX.addDefinedName(s, "myNCRange", "C1:C5,D1:D5")
        @test XLSX.setConditionalFormat(s, "myNCRange", :colorScale;
            min_type="min",
            min_col="green",
            mid_type="percentile",
            mid_val="50",
            mid_col="red",
            max_type="max",
            max_col="blue"
        ) == 0
        @test_throws XLSX.XLSXError XLSX.setConditionalFormat(s, "A1:A2", :colorScale;
            colorscale="rainbow"
        )
        SAVE_FILES && save_outfile(f)

        f = XLSX.newxlsx()
        s = f[1]
        for i in 1:5, j in 1:12
            s[i, j] = i + j
        end
        for (j, k) in enumerate(keys(XLSX.colorscales))
            @test XLSX.setConditionalFormat(s, :, j, :colorScale; colorscale=k) == 0
        end
        SAVE_FILES && save_outfile(f)
    end

    @testset "iconSet" begin
        f = XLSX.newxlsx()
        s = f[1]
        for i in 1:5, j in 1:5
            s[i, j] = i + j
        end
        @test XLSX.setConditionalFormat(s, "A1,A3", :iconSet) == 0
        @test XLSX.setConditionalFormat(s, [1], 1, :iconSet) == 0 # Vectors may be non-contiguous
        @test_throws XLSX.XLSXError XLSX.setConditionalFormat(s, 1, 1:3:7, :iconSet) # out of range
        @test XLSX.setConditionalFormat(s, "1:1", :iconSet) == 0
        @test XLSX.setConditionalFormat(s, 2, :, :iconSet; iconset="3Arrows") == 0
        @test XLSX.setConditionalFormat(s, 3, 1:5, :iconSet;
            min_type="percent",
            min_val="20",
            max_type="num",
            max_val="4"
        ) == 0
        @test XLSX.setConditionalFormat(s, "Sheet1!A4:E4", :iconSet;
            min_type="percentile",
            min_val="10",
            max_type="num",
            max_val="\$C\$4"
        ) == 0
        @test XLSX.setConditionalFormat(f, "Sheet1!A5:E5", :iconSet;
            min_type="percentile",
            min_val="\$D\$5",
            max_type="percent",
            max_val="95"
        ) == 0
        @test sort!(XLSX.getConditionalFormats(s), by=x -> x.second.priority, rev=true) == [
            XLSX.CellRange("A5:E5") => (type="iconSet", priority=7),
            XLSX.CellRange("A4:E4") => (type="iconSet", priority=6),
            XLSX.CellRange("A3:E3") => (type="iconSet", priority=5),
            XLSX.CellRange("A2:E2") => (type="iconSet", priority=4),
            XLSX.CellRange("A1:E1") => (type="iconSet", priority=3),
            XLSX.CellRange("A1:A1") => (type="iconSet", priority=2),
            XLSX.NonContiguousRange("Sheet1!A1,Sheet1!A3") => (type="iconSet", priority=1)
        ]

        @test XLSX.setConditionalFormat(s, "A1", :iconSet) == 0
        @test XLSX.setConditionalFormat(s, "A1:C3", :iconSet) == 0
        @test XLSX.setConditionalFormat(s, "Sheet1!A1", :iconSet) == 0
        @test XLSX.setConditionalFormat(s, "Sheet1!A1:A2", :iconSet) == 0
        @test XLSX.setConditionalFormat(s, "Sheet1!1:2", :iconSet) == 0
        @test XLSX.setConditionalFormat(s, "2:4", :iconSet) == 0
        @test XLSX.setConditionalFormat(s, "A:C", :iconSet) == 0
        @test XLSX.setConditionalFormat(s, "Sheet1!A:C", :iconSet) == 0
        @test XLSX.setConditionalFormat(f, "Sheet1!1:2", :iconSet) == 0
        @test XLSX.setConditionalFormat(f, "Sheet1!A:C", :iconSet) == 0
        @test XLSX.setConditionalFormat(s, :, 1:3, :iconSet) == 0
        @test XLSX.setConditionalFormat(s, 1:3, :, :iconSet) == 0
        @test XLSX.setConditionalFormat(s, "2:4", :iconSet) == 0
        @test XLSX.setConditionalFormat(s, "A:C", :iconSet) == 0
        @test XLSX.setConditionalFormat(s, "Sheet1!A:C", :iconSet) == 0
        @test XLSX.setConditionalFormat(s, :, :iconSet) == 0
        @test XLSX.setConditionalFormat(s, :, :, :iconSet) == 0
        @test length(XLSX.getConditionalFormats(s)) == 24

        @test sort!(XLSX.getConditionalFormats(s), by=x -> x.second.priority, rev=true) == [
            XLSX.CellRange("A1:E5") => (type="iconSet", priority=24),
            XLSX.CellRange("A1:E5") => (type="iconSet", priority=23),
            XLSX.CellRange("A1:C5") => (type="iconSet", priority=22),
            XLSX.CellRange("A1:C5") => (type="iconSet", priority=21),
            XLSX.CellRange("A2:E4") => (type="iconSet", priority=20),
            XLSX.CellRange("A1:E3") => (type="iconSet", priority=19),
            XLSX.CellRange("A1:C5") => (type="iconSet", priority=18),
            XLSX.CellRange("A1:C5") => (type="iconSet", priority=17),
            XLSX.CellRange("A1:E2") => (type="iconSet", priority=16),
            XLSX.CellRange("A1:C5") => (type="iconSet", priority=15),
            XLSX.CellRange("A1:C5") => (type="iconSet", priority=14),
            XLSX.CellRange("A2:E4") => (type="iconSet", priority=13),
            XLSX.CellRange("A1:E2") => (type="iconSet", priority=12),
            XLSX.CellRange("A1:A2") => (type="iconSet", priority=11),
            XLSX.CellRange("A1:A1") => (type="iconSet", priority=10),
            XLSX.CellRange("A1:C3") => (type="iconSet", priority=9),
            XLSX.CellRange("A1:A1") => (type="iconSet", priority=8),
            XLSX.CellRange("A5:E5") => (type="iconSet", priority=7),
            XLSX.CellRange("A4:E4") => (type="iconSet", priority=6),
            XLSX.CellRange("A3:E3") => (type="iconSet", priority=5),
            XLSX.CellRange("A2:E2") => (type="iconSet", priority=4),
            XLSX.CellRange("A1:E1") => (type="iconSet", priority=3),
            XLSX.CellRange("A1:A1") => (type="iconSet", priority=2),
            XLSX.NonContiguousRange("Sheet1!A1,Sheet1!A3") => (type="iconSet", priority=1)
        ]

        SAVE_FILES && save_outfile(f)

        f = XLSX.newxlsx()
        s = f[1]

        XLSX.writetable!(s, [collect(1:10), collect(1:10), collect(1:10), collect(1:10), collect(1:10), collect(1:10)],
            ["normal", "showVal=\"false\"", "reverse=\"true\"", "min_gte=\"false\"", "extra1", "extra2"])
        s["G1"] = 3
        s["G4"] = "y"

        @test XLSX.setConditionalFormat(s, "A2:A11", :iconSet;
            min_type="num", max_type="formula",
            min_val="3", max_val="if(\$G\$4=\"y\", \$G\$1+5, 10)") == 0

        @test XLSX.setConditionalFormat(s, "A2:A11", :iconSet;
            min_type="num", max_type="num",
            min_val="3", max_val="8") == 0

        @test XLSX.setConditionalFormat(s, "B2:B11", :iconSet; iconset="4TrafficLights",
            min_type="num", mid_type="percent", max_type="num",
            min_val="3", mid_val="50", max_val="8",
            showVal="false") == 0

        @test XLSX.setConditionalFormat(s, "C2:C11", :iconSet; iconset="3Symbols2",
            min_type="num", mid_type="percentile", max_type="num",
            min_val="3", mid_val="50", max_val="8",
            reverse="true") == 0

        @test XLSX.setConditionalFormat(s, "D2:D11", :iconSet; iconset="5Arrows",
            min_type="num", mid_type="percentile", mid2_type="percentile", max_type="num",
            min_val="3", mid_val="45", mid2_val="65", max_val="8",
            min_gte="false", max_gte="false") == 0

        @test XLSX.setConditionalFormat(s, "E2:E11", :iconSet; iconset="3Stars",
            reverse="true",
            showVal="false",
            min_type="num", mid_type="percentile", mid2_type="percentile", max_type="num",
            min_val="3", mid_val="45", mid2_val="65", max_val="8",
            min_gte="false", max_gte="false") == 0

        @test XLSX.setConditionalFormat(s, "F2:F11", :iconSet; iconset="5Boxes",
            reverse="true",
            showVal="false",
            min_type="num", mid_type="percentile", mid2_type="percentile", max_type="num",
            min_val="3", mid_val="45", mid2_val="65", max_val="8",
            min_gte="false", mid_gte="false", mid2_gte="false", max_gte="false") == 0

        @test sort!(XLSX.getConditionalFormats(s), by=x -> x.second.priority, rev=true) == [
            XLSX.CellRange("F2:F11") => (type="iconSet", priority=7),
            XLSX.CellRange("E2:E11") => (type="iconSet", priority=6),
            XLSX.CellRange("D2:D11") => (type="iconSet", priority=5),
            XLSX.CellRange("C2:C11") => (type="iconSet", priority=4),
            XLSX.CellRange("B2:B11") => (type="iconSet", priority=3),
            XLSX.CellRange("A2:A11") => (type="iconSet", priority=2),
            XLSX.CellRange("A2:A11") => (type="iconSet", priority=1),
        ]
        SAVE_FILES && save_outfile(f)

        f = XLSX.newxlsx()
        s = f[1]
        for i = 0:3
            for j = 1:13
                s[i+1, j] = i * 13 + j
            end
        end
        for j = 1:13
            @test XLSX.setConditionalFormat(s, 1:4, j, :iconSet; # Create a custom 4-icon set in each column.
                iconset="Custom",
                icon_list=[j, 13 + j, 26 + j, 39 + j],
                min_type="percent", mid_type="percent", max_type="percent",
                min_val="25", mid_val="50", max_val="75"
            ) == 0
        end

        @test XLSX.setConditionalFormat(s, 1:4, 1, :iconSet;
            iconset="Custom",
            icon_list=[1, 2, 3, 4, 5],
            min_type="percent", max_type="percent",
            min_val="25", max_val="75",
            min_gte="false", max_gte="false"
        ) == 0
        @test XLSX.setConditionalFormat(s, 1:4, 1, :iconSet;
            iconset="Custom",
            showVal="false",
            icon_list=[1, 2, 3, 4, 5],
            min_type="percent", mid_type="percent", max_type="percent",
            min_val="25", mid_val="50", max_val="75"
        ) == 0
        @test XLSX.setConditionalFormat(s, 1:4, 1, :iconSet;
            iconset="Custom",
            reverse="true",
            icon_list=[1, 2, 3, 4, 5],
            min_type="percent", mid_type="percent", mid2_type="percentile", max_type="percent",
            min_val="25", mid_val="50", mid2_val="60", max_val="75"
        ) == 0

        @test XLSX.setConditionalFormat(s, "A2:M2", :iconSet;
            iconset="Custom",
            icon_list=[31, 24, 11],
            min_type="num", max_type="formula",
            min_val="3", max_val="if(\$G\$4=\"y\", \$G\$1+5, 10)") == 0

        @test_throws XLSX.XLSXError XLSX.setConditionalFormat(s, 1:4, 1, :iconSet;
            iconset="Custom",
            icon_list=[1, 2, 3, 4, 5],
            min_type="percent", mid_type="madeUp", mid2_type="percentile", max_type="num",
            min_val="25", mid_val="50", mid2_val="60", max_val="75"
        )
        @test_throws XLSX.XLSXError XLSX.setConditionalFormat(s, 1:4, 1, :iconSet;
            iconset="Custom",
            icon_list=[99, 2, 3, 4, 5],
            min_type="percent", mid_type="percent", mid2_type="percentile", max_type="num",
            min_val="25", mid_val="50", mid2_val="60", max_val="75"
        )
        @test_throws XLSX.XLSXError XLSX.setConditionalFormat(s, 1:4, 1, :iconSet;
            iconset="Custom",
            min_type="percent", mid_type="percent", max_type="percent",
            min_val="25", mid_val="50", max_val="75"
        )
        @test_throws XLSX.XLSXError XLSX.setConditionalFormat(s, 1:4, 1, :iconSet;
            iconset="Custom",
            icon_list=[],
            min_type="percent", mid_type="percent", max_type="percent",
            min_val="25", mid_val="50", max_val="75"
        )
        @test_throws XLSX.XLSXError XLSX.setConditionalFormat(s, 1:4, 1, :iconSet;
            iconset="Custom",
            icon_list=[1, 13, 26],
            min_type="percent", mid_type="percent", max_type="percent",
            min_val="25", mid_val="50", max_val="75"
        )
        @test_throws XLSX.XLSXError XLSX.setConditionalFormat(s, 1:4, 1, :iconSet;
            iconset="Custom",
            icon_list=[1, 13, 26, 39],
            min_type="percent", max_type="percent",
            min_val="25"
        ) == 0
        @test_throws XLSX.XLSXError XLSX.setConditionalFormat(s, 1:4, 1, :iconSet;
            iconset="Custom",
            icon_list=[1, 13, 26, 39],
            min_type="percent",
            min_val="25", max_val="75"
        ) == 0
        @test_throws XLSX.XLSXError XLSX.setConditionalFormat(s, 1:4, 1, :iconSet;
            iconset="Custom",
            icon_list=[1, 13, 26, 39]
        ) == 0
        @test_throws XLSX.XLSXError XLSX.setConditionalFormat(s, 1:4, 1, :iconSet;
            iconset="10ThousandManiacs",
        ) == 0


        @test sort!(XLSX.getConditionalFormats(s), by=x -> x.second.priority) == [
            XLSX.CellRange("A1:A4") => (type="iconSet", priority=1),
            XLSX.CellRange("B1:B4") => (type="iconSet", priority=2),
            XLSX.CellRange("C1:C4") => (type="iconSet", priority=3),
            XLSX.CellRange("D1:D4") => (type="iconSet", priority=4),
            XLSX.CellRange("E1:E4") => (type="iconSet", priority=5),
            XLSX.CellRange("F1:F4") => (type="iconSet", priority=6),
            XLSX.CellRange("G1:G4") => (type="iconSet", priority=7),
            XLSX.CellRange("H1:H4") => (type="iconSet", priority=8),
            XLSX.CellRange("I1:I4") => (type="iconSet", priority=9),
            XLSX.CellRange("J1:J4") => (type="iconSet", priority=10),
            XLSX.CellRange("K1:K4") => (type="iconSet", priority=11),
            XLSX.CellRange("L1:L4") => (type="iconSet", priority=12),
            XLSX.CellRange("M1:M4") => (type="iconSet", priority=13),
            XLSX.CellRange("A1:A4") => (type="iconSet", priority=14),
            XLSX.CellRange("A1:A4") => (type="iconSet", priority=15),
            XLSX.CellRange("A1:A4") => (type="iconSet", priority=16),
            XLSX.CellRange("A2:M2") => (type="iconSet", priority=17)
        ]

        XLSX.addDefinedName(s, "myRange", "A1:B2")
        XLSX.addDefinedName(s, "myNCRange", "C1:C5,D1:D5")
        @test XLSX.setConditionalFormat(s, "myRange", :iconSet) == 0
        @test_throws XLSX.XLSXError XLSX.setConditionalFormat(s, "myNCRange", :iconSet)
        SAVE_FILES && save_outfile(f)

        f = XLSX.newxlsx()
        s = f[1]
        for i in 1:5, j in 1:21
            s[i, j] = i + j
        end
        for (j, k) in enumerate(keys(XLSX.iconsets))
            if k == "Custom"
                @test XLSX.setConditionalFormat(s, :, j, :iconSet;
                    iconset=k,
                    icon_list=[1, 2, 3, 4, 5],
                    min_type="num", mid_type="num", mid2_type="num", max_type="num",
                    min_val="8", mid_val="12", mid2_val="15", max_val="18",
                ) == 0
            else
                @test XLSX.setConditionalFormat(s, :, j, :iconSet; iconset=k) == 0
            end
        end
        SAVE_FILES && save_outfile(f)

        f = XLSX.newxlsx()
        s = f[1]
        for i in 1:3, j in 1:21
            s[i, j] = i + j
        end
        @test XLSX.setConditionalFormat(s, "Sheet1!A1:E1", :iconSet;
            min_type="percentile",
            min_val="10",
            max_type="num",
            max_val="Sheet1!\$C\$4"
        ) == 0
        @test XLSX.setConditionalFormat(f, "Sheet1!A2:E2", :iconSet;
            min_type="percentile",
            min_val="Sheet1!\$D\$5",
            max_type="percent",
            max_val="95"
        ) == 0
        @test_throws XLSX.XLSXError XLSX.setConditionalFormat(s, "Sheet1!A1:E1", :iconSet;
            min_type="percentile",
            min_val="10",
            max_type="num",
            max_val="Sheet2!\$C\$4"
        )
        @test_throws XLSX.XLSXError XLSX.setConditionalFormat(f, "Sheet1!A5:E5", :iconSet;
            min_type="percentile",
            min_val="Sheet2!\$D\$5",
            max_type="percent",
            max_val="95"
        )
        @test XML.tag(XLSX.get_x14_icon("3Triangles")) == "x14:cfRule"
        @test XML.attributes(XLSX.get_x14_icon("3Stars")) == OrderedDict("type" => "iconSet", "priority" => "1", "id" => "XXXX-xxxx-XXXX")
        @test length(XML.children(XLSX.get_x14_icon("5Boxes"))) == 1
        @test typeof(XLSX.get_x14_icon("Custom")) == XML.Node{String}
        SAVE_FILES && save_outfile(f)
    end

    @testset "cellIs" begin
        f = XLSX.newxlsx()
        s = f[1]
        for i in 1:5, j in 1:5
            s[i, j] = i + j
        end
        @test_throws XLSX.XLSXError XLSX.setConditionalFormat(s, "A1:A5", :cellIs; notAKeyword="x")
        @test XLSX.setConditionalFormat(s, "A1,A3", :cellIs) == 0 # Non-contiguous ranges not allowed
        @test XLSX.setConditionalFormat(s, [1], 1, :cellIs) == 0 # Vectors may be non-contiguous
        @test_throws XLSX.XLSXError XLSX.setConditionalFormat(s, 1, 1:3:7, :cellIs) # out of range
        @test_throws XLSX.XLSXError XLSX.setConditionalFormat(s, "A1:A3", :cellIs; dxStyle="madeUp") # dxStyle invalid
        @test XLSX.setConditionalFormat(s, "1:1", :cellIs) == 0
        @test XLSX.setConditionalFormat(s, 2, :, :cellIs; dxStyle="greenfilltext") == 0
        @test XLSX.setConditionalFormat(s, 3, 1:5, :cellIs;
            operator="between",
            value="2",
            value2="3",
            stopIfTrue="true",
            fill=["pattern" => "none", "bgColor" => "FFFFC7CE"],
            format=["format" => "0.00%"],
            font=["color" => "blue", "bold" => "true"]
        ) == 0

        @test XLSX.setConditionalFormat(s, "Sheet1!A4:E4", :cellIs;
            operator="greaterThan",
            value="4",
            fill=["pattern" => "none", "bgColor" => "green"],
            format=["format" => "0.0"],
            font=["color" => "red", "italic" => "true"]
        ) == 0
        @test XLSX.setConditionalFormat(f, "Sheet1!A5:E5", :cellIs;
            operator="lessThan",
            value="2",
            fill=["pattern" => "none", "bgColor" => "yellow"],
            format=["format" => "0.0"],
            font=["color" => "green"],
            border=["style" => "thick", "color" => "coral"]
        ) == 0
        @test sort!(XLSX.getConditionalFormats(s), by=x -> x.second.priority, rev=true) == [
            XLSX.CellRange("A5:E5") => (type="cellIs", priority=7),
            XLSX.CellRange("A4:E4") => (type="cellIs", priority=6),
            XLSX.CellRange("A3:E3") => (type="cellIs", priority=5),
            XLSX.CellRange("A2:E2") => (type="cellIs", priority=4),
            XLSX.CellRange("A1:E1") => (type="cellIs", priority=3),
            XLSX.CellRange("A1:A1") => (type="cellIs", priority=2),
            XLSX.NonContiguousRange("Sheet1!A1,Sheet1!A3") => (type="cellIs", priority=1)
        ]

        @test XLSX.setConditionalFormat(s, "A1", :cellIs) == 0
        @test XLSX.setConditionalFormat(s, "A1:C3", :cellIs) == 0
        @test XLSX.setConditionalFormat(s, "Sheet1!A1", :cellIs) == 0
        @test XLSX.setConditionalFormat(s, "Sheet1!A1:A2", :cellIs) == 0
        @test XLSX.setConditionalFormat(s, "Sheet1!1:2", :cellIs) == 0
        @test XLSX.setConditionalFormat(s, "2:4", :cellIs) == 0
        @test XLSX.setConditionalFormat(s, "A:C", :cellIs) == 0
        @test XLSX.setConditionalFormat(s, "Sheet1!A:C", :cellIs) == 0
        @test XLSX.setConditionalFormat(f, "Sheet1!1:2", :cellIs) == 0
        @test XLSX.setConditionalFormat(f, "Sheet1!A:C", :cellIs) == 0
        @test XLSX.setConditionalFormat(s, :, 1:3, :cellIs) == 0
        @test XLSX.setConditionalFormat(s, 1:3, :, :cellIs) == 0
        @test XLSX.setConditionalFormat(s, "2:4", :cellIs) == 0
        @test XLSX.setConditionalFormat(s, "A:C", :cellIs) == 0
        @test XLSX.setConditionalFormat(s, "Sheet1!A:C", :cellIs) == 0
        @test XLSX.setConditionalFormat(s, :, :cellIs) == 0
        @test XLSX.setConditionalFormat(s, :, :, :cellIs) == 0
        @test length(XLSX.getConditionalFormats(s)) == 24
        @test sort!(XLSX.getConditionalFormats(s), by=x -> x.second.priority, rev=true) == [
            XLSX.CellRange("A1:E5") => (type="cellIs", priority=24),
            XLSX.CellRange("A1:E5") => (type="cellIs", priority=23),
            XLSX.CellRange("A1:C5") => (type="cellIs", priority=22),
            XLSX.CellRange("A1:C5") => (type="cellIs", priority=21),
            XLSX.CellRange("A2:E4") => (type="cellIs", priority=20),
            XLSX.CellRange("A1:E3") => (type="cellIs", priority=19),
            XLSX.CellRange("A1:C5") => (type="cellIs", priority=18),
            XLSX.CellRange("A1:C5") => (type="cellIs", priority=17),
            XLSX.CellRange("A1:E2") => (type="cellIs", priority=16),
            XLSX.CellRange("A1:C5") => (type="cellIs", priority=15),
            XLSX.CellRange("A1:C5") => (type="cellIs", priority=14),
            XLSX.CellRange("A2:E4") => (type="cellIs", priority=13),
            XLSX.CellRange("A1:E2") => (type="cellIs", priority=12),
            XLSX.CellRange("A1:A2") => (type="cellIs", priority=11),
            XLSX.CellRange("A1:A1") => (type="cellIs", priority=10),
            XLSX.CellRange("A1:C3") => (type="cellIs", priority=9),
            XLSX.CellRange("A1:A1") => (type="cellIs", priority=8),
            XLSX.CellRange("A5:E5") => (type="cellIs", priority=7),
            XLSX.CellRange("A4:E4") => (type="cellIs", priority=6),
            XLSX.CellRange("A3:E3") => (type="cellIs", priority=5),
            XLSX.CellRange("A2:E2") => (type="cellIs", priority=4),
            XLSX.CellRange("A1:E1") => (type="cellIs", priority=3),
            XLSX.CellRange("A1:A1") => (type="cellIs", priority=2),
            XLSX.NonContiguousRange("Sheet1!A1,Sheet1!A3") => (type="cellIs", priority=1)
        ]

        @test_throws XLSX.XLSXError XLSX.setConditionalFormat(s, "Sheet1!A4:E4", :cellIs;
            operator="madeUp",
            value="4",
            fill=["pattern" => "none", "bgColor" => "green"],
            format=["format" => "0.0"],
            font=["color" => "red", "italic" => "true"]
        )
        SAVE_FILES && save_outfile(f)

        f = XLSX.newxlsx()
        s = f[1]
        for i in 1:5, j in 1:5
            s[i, j] = i + j
        end
        XLSX.setConditionalFormat(s, "A1:A5", :cellIs)
        XLSX.setConditionalFormat(s, :, 2, :cellIs; dxStyle="redborder")
        XLSX.setConditionalFormat(s, "Sheet1!E:E", :cellIs; dxStyle="redfilltext")
        XLSX.setConditionalFormat(s, 1:5, 3:4, :cellIs;
            operator="between",
            value="2",
            value2="4",
            fill=["pattern" => "none", "bgColor" => "yellow"],
            format=["format" => "0.0"],
            font=["color" => "green"],
            border=["style" => "thick", "color" => "coral"]
        ) == 0

        @test sort!(XLSX.getConditionalFormats(s), by=x -> x.second.priority, rev=true) == [
            XLSX.CellRange("C1:D5") => (type="cellIs", priority=4),
            XLSX.CellRange("E1:E5") => (type="cellIs", priority=3),
            XLSX.CellRange("B1:B5") => (type="cellIs", priority=2),
            XLSX.CellRange("A1:A5") => (type="cellIs", priority=1)
        ]
        SAVE_FILES && save_outfile(f)

        f = XLSX.newxlsx()
        s = f[1]
        for i in 1:5, j in 1:5
            s[i, j] = i + j
        end

        @test XLSX.setConditionalFormat(s, :, 1:4, :cellIs;
            operator="lessThan",
            value="\$E\$4",
            fill=["pattern" => "none", "bgColor" => "yellow"],
            format=["format" => "0.0"],
            font=["color" => "green", "under" => "double"],
            border=["style" => "thin", "color" => "coral"]
        ) == 0
        SAVE_FILES && save_outfile(f)

        f = XLSX.newxlsx()
        s = f[1]
        for i in 1:5, j in 1:5
            s[i, j] = i + j
        end
        XLSX.addDefinedName(s, "myRange", "A1:B5")
        @test XLSX.setConditionalFormat(s, "myRange", :cellIs;
            operator="lessThan",
            value="2",
            fill=["pattern" => "none", "bgColor" => "yellow"],
            format=["format" => "0.0"],
            font=["color" => "green"],
            border=["style" => "hair", "color" => "cyan"]
        ) == 0
        XLSX.addDefinedName(s, "myNCRange", "C1:C5,D1:D5")
        @test XLSX.setConditionalFormat(s, "myNCRange", :cellIs; # Non-contiguous ranges not allowed
            operator="lessThan",
            value="2",
            fill=["pattern" => "none", "bgColor" => "yellow"],
            format=["format" => "0.0"],
            font=["color" => "green"],
            border=["style" => "hair", "color" => "cyan"]
        ) == 0
        SAVE_FILES && save_outfile(f)

        f = XLSX.newxlsx()
        s = f[1]
        for i in 1:5, j in 1:6
            s[i, j] = i + j
        end
        for (j, k) in enumerate(keys(XLSX.highlights))
            @test XLSX.setConditionalFormat(s, :, j, :cellIs; dxStyle=k) == 0
        end
        SAVE_FILES && save_outfile(f)
    end


    @testset "containsText" begin
        f = XLSX.newxlsx()
        s = f[1]
        s["A1:E1"] = "Hello World"
        s["A2:E2"] = "Life the universe and everything"
        s["A3:E3"] = "Once upon a time"
        s["A4:E4"] = "In America"
        s["A5:E5"] = "a"
        @test_throws XLSX.XLSXError XLSX.setConditionalFormat(s, "A1:A5", :containsText; notAKeyword="x")
        @test XLSX.setConditionalFormat(s, "A1,A3", :containsText; value="a") == 0
        @test XLSX.setConditionalFormat(s, [1], 1, :containsText; value="a") == 0
        @test_throws XLSX.XLSXError XLSX.setConditionalFormat(s, 1, 1:3:7, :containsText; value="a") # out of range
        @test_throws XLSX.XLSXError XLSX.setConditionalFormat(s, "1:1", :containsText) # value must be defined
        @test XLSX.setConditionalFormat(s, "1:1", :containsText; value="a") == 0
        @test XLSX.setConditionalFormat(s, 2, :, :containsText; value="a", dxStyle="greenfilltext") == 0
        @test XLSX.setConditionalFormat(s, 3, 1:5, :containsText;
            operator="notContainsText",
            value="a",
            fill=["pattern" => "none", "bgColor" => "FFFFC7CE"],
            format=["format" => "0.00%"],
            font=["color" => "blue", "bold" => "true"]
        ) == 0

        @test XLSX.setConditionalFormat(s, "Sheet1!A4:E4", :containsText;
            operator="notContainsText",
            value="a",
            fill=["pattern" => "none", "bgColor" => "green"],
            format=["format" => "0.0"],
            font=["color" => "red", "italic" => "true"]
        ) == 0
        @test XLSX.setConditionalFormat(f, "Sheet1!A5:E5", :containsText;
            operator="beginsWith",
            value="a",
            fill=["pattern" => "none", "bgColor" => "yellow"],
            format=["format" => "0.0"],
            font=["color" => "green"],
            border=["style" => "thick", "color" => "coral"]
        ) == 0
        @test sort!(XLSX.getConditionalFormats(s), by=x -> x.second.priority, rev=true) == [
            XLSX.CellRange("A5:E5") => (type="beginsWith", priority=7),
            XLSX.CellRange("A4:E4") => (type="notContainsText", priority=6),
            XLSX.CellRange("A3:E3") => (type="notContainsText", priority=5),
            XLSX.CellRange("A2:E2") => (type="containsText", priority=4),
            XLSX.CellRange("A1:E1") => (type="containsText", priority=3),
            XLSX.CellRange("A1:A1") => (type="containsText", priority=2),
            XLSX.NonContiguousRange("Sheet1!A1,Sheet1!A3") => (type="containsText", priority=1)
        ]

        #        @test XLSX.getConditionalFormats(s) == [XLSX.CellRange("A5:E5") => (type = "containsText", priority = 5), XLSX.CellRange("A4:E4") => (type = "containsText", priority = 4), XLSX.CellRange("A3:E3") => (type = "containsText", priority = 3), XLSX.CellRange("A2:E2") => (type = "containsText", priority = 2), XLSX.CellRange("A1:E1") => (type = "containsText", priority = 1)]
        @test XLSX.setConditionalFormat(s, "A1", :containsText; value="a") == 0
        @test XLSX.setConditionalFormat(s, "A1:C3", :containsText; value="a") == 0
        @test XLSX.setConditionalFormat(s, "Sheet1!A1", :containsText; value="a") == 0
        @test XLSX.setConditionalFormat(s, "Sheet1!A1:A2", :containsText; value="a") == 0
        @test XLSX.setConditionalFormat(s, "Sheet1!1:2", :containsText; value="a") == 0
        @test XLSX.setConditionalFormat(s, "2:4", :containsText; value="a") == 0
        @test XLSX.setConditionalFormat(s, "A:C", :containsText; value="a") == 0
        @test XLSX.setConditionalFormat(s, "Sheet1!A:C", :containsText; value="a") == 0
        @test XLSX.setConditionalFormat(f, "Sheet1!1:2", :containsText; value="a") == 0
        @test XLSX.setConditionalFormat(f, "Sheet1!A:C", :containsText; value="a") == 0
        @test XLSX.setConditionalFormat(s, :, 1:3, :containsText; value="a") == 0
        @test XLSX.setConditionalFormat(s, 1:3, :, :containsText; value="a") == 0
        @test XLSX.setConditionalFormat(s, "2:4", :containsText; value="a") == 0
        @test XLSX.setConditionalFormat(s, "A:C", :containsText; value="a") == 0
        @test XLSX.setConditionalFormat(s, "Sheet1!A:C", :containsText; value="a") == 0
        @test XLSX.setConditionalFormat(s, :, :containsText; value="a") == 0
        @test XLSX.setConditionalFormat(s, :, :, :containsText; value="a") == 0
        @test length(XLSX.getConditionalFormats(s)) == 24
        @test sort!(XLSX.getConditionalFormats(s), by=x -> x.second.priority, rev=true) == [
            XLSX.CellRange("A1:E5") => (type="containsText", priority=24),
            XLSX.CellRange("A1:E5") => (type="containsText", priority=23),
            XLSX.CellRange("A1:C5") => (type="containsText", priority=22),
            XLSX.CellRange("A1:C5") => (type="containsText", priority=21),
            XLSX.CellRange("A2:E4") => (type="containsText", priority=20),
            XLSX.CellRange("A1:E3") => (type="containsText", priority=19),
            XLSX.CellRange("A1:C5") => (type="containsText", priority=18),
            XLSX.CellRange("A1:C5") => (type="containsText", priority=17),
            XLSX.CellRange("A1:E2") => (type="containsText", priority=16),
            XLSX.CellRange("A1:C5") => (type="containsText", priority=15),
            XLSX.CellRange("A1:C5") => (type="containsText", priority=14),
            XLSX.CellRange("A2:E4") => (type="containsText", priority=13),
            XLSX.CellRange("A1:E2") => (type="containsText", priority=12),
            XLSX.CellRange("A1:A2") => (type="containsText", priority=11),
            XLSX.CellRange("A1:A1") => (type="containsText", priority=10),
            XLSX.CellRange("A1:C3") => (type="containsText", priority=9),
            XLSX.CellRange("A1:A1") => (type="containsText", priority=8),
            XLSX.CellRange("A5:E5") => (type="beginsWith", priority=7),
            XLSX.CellRange("A4:E4") => (type="notContainsText", priority=6),
            XLSX.CellRange("A3:E3") => (type="notContainsText", priority=5),
            XLSX.CellRange("A2:E2") => (type="containsText", priority=4),
            XLSX.CellRange("A1:E1") => (type="containsText", priority=3),
            XLSX.CellRange("A1:A1") => (type="containsText", priority=2),
            XLSX.NonContiguousRange("Sheet1!A1,Sheet1!A3") => (type="containsText", priority=1)
        ]

        SAVE_FILES && save_outfile(f)

        f = XLSX.newxlsx()
        s = f[1]
        s["A1:E1"] = "Hello World"
        s["A2:E2"] = "Life the universe and everything"
        s["A3:E3"] = "Once upon a time"
        s["A4:E4"] = "In America"
        s["A5:E5"] = "a"
        XLSX.setConditionalFormat(s, "A1:A5", :containsText; value="a")
        XLSX.setConditionalFormat(s, :, 2, :containsText; value="a", dxStyle="redborder")
        XLSX.setConditionalFormat(s, "Sheet1!E:E", :containsText; value="a", dxStyle="redfilltext")
        XLSX.setConditionalFormat(s, 1:5, 3:4, :containsText;
            operator="endsWith",
            value="a",
            fill=["pattern" => "none", "bgColor" => "yellow"],
            format=["format" => "0.0"],
            font=["color" => "green"],
            border=["style" => "thick", "color" => "coral"]
        ) == 0

        @test sort!(XLSX.getConditionalFormats(s), by=x -> x.second.priority, rev=true) == [
            XLSX.CellRange("C1:D5") => (type="endsWith", priority=4),
            XLSX.CellRange("E1:E5") => (type="containsText", priority=3),
            XLSX.CellRange("B1:B5") => (type="containsText", priority=2),
            XLSX.CellRange("A1:A5") => (type="containsText", priority=1)
        ]
        SAVE_FILES && save_outfile(f)

        f = XLSX.newxlsx()
        s = f[1]
        s["A1:E1"] = "Hello World"
        s["A2:E2"] = "Life the universe and everything"
        s["A3:E3"] = "Once upon a time"
        s["A4:E4"] = "In America"
        s["A5:E5"] = "a"

        @test XLSX.setConditionalFormat(s, :, 1:4, :containsText;
            operator="containsText",
            value="Sheet1!\$E\$5",
            stopIfTrue="true",
            fill=["pattern" => "none", "bgColor" => "yellow"],
            format=["format" => "0.0"],
            font=["color" => "green", "under" => "double"],
            border=["style" => "thin", "color" => "coral"]
        ) == 0
        SAVE_FILES && save_outfile(f)

        f = XLSX.newxlsx()
        s = f[1]
        s["A1:E1"] = "Hello World"
        s["A2:E2"] = "Life the universe and everything"
        s["A3:E3"] = "Once upon a time"
        s["A4:E4"] = "In America"
        s["A5:E5"] = "a"
        XLSX.addDefinedName(s, "myRange", "A1:B5")
        @test XLSX.setConditionalFormat(s, "myRange", :containsText;
            operator="notContainsText",
            value="a",
            fill=["pattern" => "none", "bgColor" => "yellow"],
            format=["format" => "0.0"],
            font=["color" => "green"],
            border=["style" => "hair", "color" => "cyan"]
        ) == 0
        @test_throws XLSX.XLSXError XLSX.setConditionalFormat(s, "myRange", :containsText;
            operator="madeUp",
            value="a",
            fill=["pattern" => "none", "bgColor" => "yellow"],
            format=["format" => "0.0"],
            font=["color" => "green"],
            border=["style" => "hair", "color" => "cyan"]
        ) == 0
        XLSX.addDefinedName(s, "myNCRange", "C1:C5,D1:D5")
        @test XLSX.setConditionalFormat(s, "myNCRange", :containsText;
            operator="beginsWith",
            value="a",
            fill=["pattern" => "none", "bgColor" => "yellow"],
            format=["format" => "0.0"],
            font=["color" => "green"],
            border=["style" => "hair", "color" => "cyan"]
        ) == 0
        SAVE_FILES && save_outfile(f)

    end

    @testset "top10" begin
        f = XLSX.newxlsx()
        s = f[1]
        for i = 1:10
            for j = 1:10
                s[i, j] = i * j
            end
        end
        @test_throws XLSX.XLSXError XLSX.setConditionalFormat(s, "A1:A5", :top10; notAKeyword="x")
        @test XLSX.setConditionalFormat(s, "A1,A3", :top10) == 0
        @test XLSX.setConditionalFormat(s, [1], 1, :top10) == 0 # Vectors may be non-contiguous
        @test XLSX.setConditionalFormat(s, 1, 1:3:7, :top10) == 0 # StepRange is non-contiguous
        @test XLSX.setConditionalFormat(s, "1:1", :top10) == 0
        @test XLSX.setConditionalFormat(s, 2, :, :top10; dxStyle="greenfilltext") == 0
        @test XLSX.setConditionalFormat(s, 1:10, 1:10, :top10;
            operator="topN",
            value="5",
            stopIfTrue="true",
            fill=["pattern" => "none", "bgColor" => "green"],
            border=["style" => "thick", "color" => "coral"],
            font=["color" => "blue", "bold" => "true", "strike" => "true"]
        ) == 0
        @test XLSX.setConditionalFormat(s, 1:10, 1:10, :top10;
            operator="bottomN",
            value="5",
            stopIfTrue="true",
            fill=["pattern" => "lightVertical", "fgColor" => "grey", "bgColor" => "FFFFC7CE"],
            border=["style" => "thick", "color" => "coral"],
            font=["color" => "blue", "bold" => "true", "strike" => "true"]
        ) == 0
        @test XLSX.setConditionalFormat(s, 1:10, 1:10, :top10;
            operator="topN%",
            value="20",
            fill=["pattern" => "none", "bgColor" => "FFFFC7CE"],
            border=["style" => "thick", "color" => "coral"],
            font=["color" => "blue", "bold" => "true", "strike" => "true"]
        ) == 0
        @test XLSX.setConditionalFormat(s, 1:10, 1:10, :top10;
            operator="bottomN%",
            value="30",
            fill=["pattern" => "none", "bgColor" => "pink"],
            border=["style" => "thick", "color" => "coral"],
            font=["color" => "blue", "bold" => "true", "italic" => "true"]
        ) == 0
        @test sort!(XLSX.getConditionalFormats(s), by=x -> x.second.priority) == [
            XLSX.NonContiguousRange("Sheet1!A1,Sheet1!A3") => (type="top10", priority=1),
            XLSX.CellRange("A1:A1") => (type="top10", priority=2),
            XLSX.NonContiguousRange("Sheet1!A1,Sheet1!D1,Sheet1!G1") => (type="top10", priority=3),
            XLSX.CellRange("A1:J1") => (type="top10", priority=4),
            XLSX.CellRange("A2:J2") => (type="top10", priority=5),
            XLSX.CellRange("A1:J10") => (type="top10", priority=6),
            XLSX.CellRange("A1:J10") => (type="top10", priority=7),
            XLSX.CellRange("A1:J10") => (type="top10", priority=8),
            XLSX.CellRange("A1:J10") => (type="top10", priority=9)
        ]


        @test XLSX.setConditionalFormat(s, "A1", :top10) == 0
        @test XLSX.setConditionalFormat(s, "A1:C3", :top10) == 0
        @test XLSX.setConditionalFormat(s, "Sheet1!A1", :top10) == 0
        @test XLSX.setConditionalFormat(s, "Sheet1!A1:A2", :top10) == 0
        @test XLSX.setConditionalFormat(s, "Sheet1!1:2", :top10) == 0
        @test XLSX.setConditionalFormat(s, "2:4", :top10) == 0
        @test XLSX.setConditionalFormat(s, "A:C", :top10) == 0
        @test XLSX.setConditionalFormat(s, "Sheet1!A:C", :top10) == 0
        @test XLSX.setConditionalFormat(f, "Sheet1!1:2", :top10) == 0
        @test XLSX.setConditionalFormat(f, "Sheet1!A:C", :top10) == 0
        @test XLSX.setConditionalFormat(s, :, 1:3, :top10) == 0
        @test XLSX.setConditionalFormat(s, 1:3, :, :top10) == 0
        @test XLSX.setConditionalFormat(s, "2:4", :top10) == 0
        @test XLSX.setConditionalFormat(s, "A:C", :top10) == 0
        @test XLSX.setConditionalFormat(s, "Sheet1!A:C", :top10) == 0
        @test XLSX.setConditionalFormat(s, :, :top10) == 0
        @test XLSX.setConditionalFormat(s, :, :, :top10) == 0
        @test length(XLSX.getConditionalFormats(s)) == 26
        @test sort!(XLSX.getConditionalFormats(s), by=x -> x.second.priority) == [
            XLSX.NonContiguousRange("Sheet1!A1,Sheet1!A3") => (type="top10", priority=1),
            XLSX.CellRange("A1:A1") => (type="top10", priority=2),
            XLSX.NonContiguousRange("Sheet1!A1,Sheet1!D1,Sheet1!G1") => (type="top10", priority=3),
            XLSX.CellRange("A1:J1") => (type="top10", priority=4),
            XLSX.CellRange("A2:J2") => (type="top10", priority=5),
            XLSX.CellRange("A1:J10") => (type="top10", priority=6),
            XLSX.CellRange("A1:J10") => (type="top10", priority=7),
            XLSX.CellRange("A1:J10") => (type="top10", priority=8),
            XLSX.CellRange("A1:J10") => (type="top10", priority=9),
            XLSX.CellRange("A1:A1") => (type="top10", priority=10),
            XLSX.CellRange("A1:C3") => (type="top10", priority=11),
            XLSX.CellRange("A1:A1") => (type="top10", priority=12),
            XLSX.CellRange("A1:A2") => (type="top10", priority=13),
            XLSX.CellRange("A1:J2") => (type="top10", priority=14),
            XLSX.CellRange("A2:J4") => (type="top10", priority=15),
            XLSX.CellRange("A1:C10") => (type="top10", priority=16),
            XLSX.CellRange("A1:C10") => (type="top10", priority=17),
            XLSX.CellRange("A1:J2") => (type="top10", priority=18),
            XLSX.CellRange("A1:C10") => (type="top10", priority=19),
            XLSX.CellRange("A1:C10") => (type="top10", priority=20),
            XLSX.CellRange("A1:J3") => (type="top10", priority=21),
            XLSX.CellRange("A2:J4") => (type="top10", priority=22),
            XLSX.CellRange("A1:C10") => (type="top10", priority=23),
            XLSX.CellRange("A1:C10") => (type="top10", priority=24),
            XLSX.CellRange("A1:J10") => (type="top10", priority=25),
            XLSX.CellRange("A1:J10") => (type="top10", priority=26)
        ]

        SAVE_FILES && save_outfile(f)

        f = XLSX.newxlsx()
        s = f[1]
        for i = 1:10
            for j = 1:10
                s[i, j] = i * j
            end
        end
        XLSX.setConditionalFormat(s, "A1:A5", :top10)
        XLSX.setConditionalFormat(s, :, 2, :top10; dxStyle="redborder")
        XLSX.setConditionalFormat(s, "Sheet1!E:E", :top10; dxStyle="redfilltext")
        XLSX.setConditionalFormat(s, 1:5, 3:4, :top10;
            operator="topN%",
            value="20",
            fill=["pattern" => "none", "bgColor" => "yellow"],
            format=["format" => "0.0"],
            font=["color" => "green"],
            border=["style" => "thick", "color" => "coral"]
        ) == 0

        @test sort!(XLSX.getConditionalFormats(s), by=x -> x.second.priority, rev=true) == [
            XLSX.CellRange("C1:D5") => (type="top10", priority=4),
            XLSX.CellRange("E1:E10") => (type="top10", priority=3),
            XLSX.CellRange("B1:B10") => (type="top10", priority=2),
            XLSX.CellRange("A1:A5") => (type="top10", priority=1)]
        SAVE_FILES && save_outfile(f)

        f = XLSX.newxlsx()
        s = f[1]
        for i = 1:10
            for j = 1:10
                s[i, j] = i * j
            end
        end

        @test XLSX.setConditionalFormat(s, :, 1:4, :top10;
            operator="bottomN",
            value="\$E\$4",
            fill=["pattern" => "none", "bgColor" => "yellow"],
            format=["format" => "0.0"],
            font=["color" => "green", "under" => "double"],
            border=["style" => "thin", "color" => "coral"]
        ) == 0
        SAVE_FILES && save_outfile(f)

        f = XLSX.newxlsx()
        s = f[1]
        for i = 1:10
            for j = 1:10
                s[i, j] = i * j
            end
        end
        XLSX.addDefinedName(s, "myRange", "A1:E5")
        @test XLSX.setConditionalFormat(s, "myRange", :top10;
            operator="topN%",
            value="2",
            fill=["pattern" => "none", "bgColor" => "yellow"],
            format=["format" => "0.0"],
            font=["color" => "green"],
            border=["style" => "medium", "color" => "cyan"]
        ) == 0
        XLSX.addDefinedName(s, "myNCRange", "C1:C5,D1:D5")
        @test XLSX.setConditionalFormat(s, "myNCRange", :top10;
            operator="bottomN%",
            value="2",
            fill=["pattern" => "none", "bgColor" => "yellow"],
            format=["format" => "0.0"],
            font=["color" => "green"],
            border=["style" => "hair", "color" => "cyan"]
        ) == 0
        @test_throws XLSX.XLSXError XLSX.setConditionalFormat(s, "myRange", :top10;
            operator="madeUp",
            value="2",
            fill=["pattern" => "none", "bgColor" => "yellow"],
            format=["format" => "0.0"],
            font=["color" => "green"],
            border=["style" => "hair", "color" => "cyan"]
        )
        SAVE_FILES && save_outfile(f)

    end

    @testset "aboveAverage" begin
        f = XLSX.newxlsx()
        s = f[1]
        d = Dist.Normal()
        columns = [rand(d, 1000), rand(d, 1000), rand(d, 1000)]
        XLSX.writetable!(s, columns, ["normal1", "normal2", "normal3"])
        @test XLSX.setConditionalFormat(s, "A2:A1001,C1:C1000", :aboveAverage) == 0
        @test XLSX.setConditionalFormat(s, [2, 3, 19], 1:3, :aboveAverage) == 0 # Vectors may be non-contiguous
        @test_throws XLSX.XLSXError XLSX.setConditionalFormat(s, "A1:A5", :aboveAverage; notAKeyword="x")
        @test_throws XLSX.XLSXError XLSX.setConditionalFormat(s, 2, 1:3:7, :aboveAverage) # out of range
        @test XLSX.setConditionalFormat(s, "2:2", :aboveAverage) == 0
        @test XLSX.setConditionalFormat(s, 2, :, :aboveAverage; dxStyle="greenfilltext") == 0
        @test XLSX.setConditionalFormat(s, 2:10, 1:3, :aboveAverage;
            operator="plus3StdDev",
            stopIfTrue="true",
            fill=["pattern" => "none", "bgColor" => "FFFFC7CE"],
            border=["style" => "thick", "color" => "coral"],
            font=["color" => "blue", "bold" => "true", "strike" => "true"]
        ) == 0
        @test XLSX.setConditionalFormat(s, 2:1001, 1:3, :aboveAverage;
            operator="minus3StdDev",
            stopIfTrue="true",
            fill=["pattern" => "lightVertical", "fgColor" => "grey", "bgColor" => "FFFFC7CE"],
            border=["style" => "thick", "color" => "coral"],
            font=["color" => "blue", "bold" => "true", "strike" => "true"]
        ) == 0
        @test XLSX.setConditionalFormat(s, 2:1001, 1:3, :aboveAverage;
            operator="plus2StdDev",
            stopIfTrue="true",
            fill=["pattern" => "none", "bgColor" => "FFFFC7CE"],
            border=["style" => "thick", "color" => "coral"],
            font=["color" => "blue", "bold" => "true", "strike" => "true"]
        ) == 0
        @test XLSX.setConditionalFormat(s, 2:1001, 1:3, :aboveAverage;
            operator="minus2StdDev",
            stopIfTrue="true",
            fill=["pattern" => "none", "bgColor" => "pink"],
            border=["style" => "thick", "color" => "coral"],
            font=["color" => "blue", "bold" => "true", "italic" => "true"]
        ) == 0
        @test XLSX.setConditionalFormat(s, 2:1001, 1:3, :aboveAverage;
            operator="plus1StdDev",
            stopIfTrue="true",
            fill=["pattern" => "none", "bgColor" => "FFFFCFCE"],
            border=["style" => "thick", "color" => "coral"],
            font=["color" => "yellow", "bold" => "true", "strike" => "true"]
        ) == 0
        @test XLSX.setConditionalFormat(s, 1:1001, 1:3, :aboveAverage;
            operator="minus1StdDev",
            stopIfTrue="true",
            fill=["pattern" => "none", "bgColor" => "yellow"],
            border=["style" => "thick", "color" => "coral"],
            font=["color" => "green", "bold" => "true", "italic" => "true"]
        ) == 0
        @test XLSX.setConditionalFormat(s, 1:1001, 1:3, :aboveAverage;
            operator="aboveAverage",
            fill=["pattern" => "none", "bgColor" => "FFFFCFCE"],
            border=["style" => "thick", "color" => "gray"],
            font=["color" => "yellow", "bold" => "true", "strike" => "true"]
        ) == 0
        @test XLSX.setConditionalFormat(s, 1:1001, 1:3, :aboveAverage;
            operator="belowAverage",
            fill=["pattern" => "none", "bgColor" => "yellow"],
            border=["style" => "thick", "color" => "green"],
            font=["color" => "green", "bold" => "true", "italic" => "true"]
        ) == 0
        @test sort!(XLSX.getConditionalFormats(s), by=x -> x.second.priority) == [
            XLSX.NonContiguousRange("Sheet1!A2:A1001,Sheet1!C1:C1000") => (type="aboveAverage", priority=1),
            XLSX.NonContiguousRange("Sheet1!A2:C3,Sheet1!A19:C19") => (type="aboveAverage", priority=2),
            XLSX.CellRange("A2:C2") => (type="aboveAverage", priority=3),
            XLSX.CellRange("A2:C2") => (type="aboveAverage", priority=4),
            XLSX.CellRange("A2:C10") => (type="aboveAverage", priority=5),
            XLSX.CellRange("A2:C1001") => (type="aboveAverage", priority=6),
            XLSX.CellRange("A2:C1001") => (type="aboveAverage", priority=7),
            XLSX.CellRange("A2:C1001") => (type="aboveAverage", priority=8),
            XLSX.CellRange("A2:C1001") => (type="aboveAverage", priority=9),
            XLSX.CellRange("A1:C1001") => (type="aboveAverage", priority=10),
            XLSX.CellRange("A1:C1001") => (type="aboveAverage", priority=11),
            XLSX.CellRange("A1:C1001") => (type="aboveAverage", priority=12)
        ]

        @test XLSX.setConditionalFormat(s, "A1", :aboveAverage) == 0
        @test XLSX.setConditionalFormat(s, "A1:C3", :aboveAverage) == 0
        @test XLSX.setConditionalFormat(s, "Sheet1!A1", :aboveAverage) == 0
        @test XLSX.setConditionalFormat(s, "Sheet1!A1:A2", :aboveAverage) == 0
        @test XLSX.setConditionalFormat(s, "Sheet1!1:2", :aboveAverage) == 0
        @test XLSX.setConditionalFormat(s, "2:4", :aboveAverage) == 0
        @test XLSX.setConditionalFormat(s, "A:C", :aboveAverage) == 0
        @test XLSX.setConditionalFormat(s, "Sheet1!A:C", :aboveAverage) == 0
        @test XLSX.setConditionalFormat(f, "Sheet1!1:2", :aboveAverage) == 0
        @test XLSX.setConditionalFormat(f, "Sheet1!A:C", :aboveAverage) == 0
        @test XLSX.setConditionalFormat(s, :, 1:3, :aboveAverage) == 0
        @test XLSX.setConditionalFormat(s, 1:3, :, :aboveAverage) == 0
        @test XLSX.setConditionalFormat(s, "2:4", :aboveAverage) == 0
        @test XLSX.setConditionalFormat(s, "A:C", :aboveAverage) == 0
        @test XLSX.setConditionalFormat(s, "Sheet1!A:C", :aboveAverage) == 0
        @test XLSX.setConditionalFormat(s, :, :aboveAverage) == 0
        @test XLSX.setConditionalFormat(s, :, :, :aboveAverage) == 0
        @test length(XLSX.getConditionalFormats(s)) == 29
        @test sort!(XLSX.getConditionalFormats(s), by=x -> x.second.priority) == [
            XLSX.NonContiguousRange("Sheet1!A2:A1001,Sheet1!C1:C1000") => (type="aboveAverage", priority=1),
            XLSX.NonContiguousRange("Sheet1!A2:C3,Sheet1!A19:C19") => (type="aboveAverage", priority=2),
            XLSX.CellRange("A2:C2") => (type="aboveAverage", priority=3),
            XLSX.CellRange("A2:C2") => (type="aboveAverage", priority=4),
            XLSX.CellRange("A2:C10") => (type="aboveAverage", priority=5),
            XLSX.CellRange("A2:C1001") => (type="aboveAverage", priority=6),
            XLSX.CellRange("A2:C1001") => (type="aboveAverage", priority=7),
            XLSX.CellRange("A2:C1001") => (type="aboveAverage", priority=8),
            XLSX.CellRange("A2:C1001") => (type="aboveAverage", priority=9),
            XLSX.CellRange("A1:C1001") => (type="aboveAverage", priority=10),
            XLSX.CellRange("A1:C1001") => (type="aboveAverage", priority=11),
            XLSX.CellRange("A1:C1001") => (type="aboveAverage", priority=12),
            XLSX.CellRange("A1:A1") => (type="aboveAverage", priority=13),
            XLSX.CellRange("A1:C3") => (type="aboveAverage", priority=14),
            XLSX.CellRange("A1:A1") => (type="aboveAverage", priority=15),
            XLSX.CellRange("A1:A2") => (type="aboveAverage", priority=16),
            XLSX.CellRange("A1:C2") => (type="aboveAverage", priority=17),
            XLSX.CellRange("A2:C4") => (type="aboveAverage", priority=18),
            XLSX.CellRange("A1:C1001") => (type="aboveAverage", priority=19),
            XLSX.CellRange("A1:C1001") => (type="aboveAverage", priority=20),
            XLSX.CellRange("A1:C2") => (type="aboveAverage", priority=21),
            XLSX.CellRange("A1:C1001") => (type="aboveAverage", priority=22),
            XLSX.CellRange("A1:C1001") => (type="aboveAverage", priority=23),
            XLSX.CellRange("A1:C3") => (type="aboveAverage", priority=24),
            XLSX.CellRange("A2:C4") => (type="aboveAverage", priority=25),
            XLSX.CellRange("A1:C1001") => (type="aboveAverage", priority=26),
            XLSX.CellRange("A1:C1001") => (type="aboveAverage", priority=27),
            XLSX.CellRange("A1:C1001") => (type="aboveAverage", priority=28),
            XLSX.CellRange("A1:C1001") => (type="aboveAverage", priority=29)
        ]


        SAVE_FILES && save_outfile(f)

        f = XLSX.newxlsx()
        s = f[1]
        for i = 1:10
            for j = 1:10
                s[i, j] = i * j
            end
        end
        @test XLSX.setConditionalFormat(s, "A1:A5", :aboveAverage) == 0
        @test XLSX.setConditionalFormat(s, :, 2, :aboveAverage; dxStyle="redborder") == 0
        @test XLSX.setConditionalFormat(s, "Sheet1!E:E", :aboveAverage; dxStyle="redfilltext") == 0
        @test XLSX.setConditionalFormat(s, 1:5, 3:4, :aboveAverage;
            operator="aboveEqAverage",
            fill=["pattern" => "none", "bgColor" => "yellow"],
            format=["format" => "0.0"],
            font=["color" => "green"],
            border=["style" => "thick", "color" => "coral"]
        ) == 0
        @test_throws XLSX.XLSXError XLSX.setConditionalFormat(s, 1:5, 3:4, :aboveAverage;
            operator="madeup",
            fill=["pattern" => "none", "bgColor" => "yellow"],
            format=["format" => "0.0"],
            font=["color" => "green"],
            border=["style" => "thick", "color" => "coral"]
        ) == 0

        @test sort!(XLSX.getConditionalFormats(s), by=x -> x.second.priority, rev=true) == [
            XLSX.CellRange("C1:D5") => (type="aboveAverage", priority=4),
            XLSX.CellRange("E1:E10") => (type="aboveAverage", priority=3),
            XLSX.CellRange("B1:B10") => (type="aboveAverage", priority=2),
            XLSX.CellRange("A1:A5") => (type="aboveAverage", priority=1)
        ]
        SAVE_FILES && save_outfile(f)

        f = XLSX.newxlsx()
        s = f[1]
        for i = 1:10
            for j = 1:10
                s[i, j] = i * j
            end
        end

        @test XLSX.setConditionalFormat(s, :, 1:4, :aboveAverage;
            operator="belowEqAverage",
            fill=["pattern" => "none", "bgColor" => "yellow"],
            format=["format" => "0.0"],
            font=["color" => "green", "under" => "double"],
            border=["style" => "thin", "color" => "coral"]
        ) == 0
        SAVE_FILES && save_outfile(f)

        f = XLSX.newxlsx()
        s = f[1]
        for i = 1:10
            for j = 1:10
                s[i, j] = i * j
            end
        end
        XLSX.addDefinedName(s, "myRange", "A1:E5")
        @test XLSX.setConditionalFormat(s, "myRange", :aboveAverage;
            operator="aboveEqAverage",
            fill=["pattern" => "none", "bgColor" => "yellow"],
            format=["format" => "0.0"],
            font=["color" => "green"],
            border=["style" => "medium", "color" => "cyan"]
        ) == 0
        XLSX.addDefinedName(s, "myNCRange", "C1:C5,D1:D5")
        @test XLSX.setConditionalFormat(s, "myNCRange", :aboveAverage;
            operator="belowEqAverage",
            fill=["pattern" => "none", "bgColor" => "yellow"],
            format=["format" => "0.0"],
            font=["color" => "green"],
            border=["style" => "hair", "color" => "cyan"]
        ) == 0
        SAVE_FILES && save_outfile(f)

    end

    @testset "timePeriod" begin
        f = XLSX.newxlsx()
        s = f[1]
        todaynow = Dates.today()
        s[1, 1:10] = todaynow - Dates.Year(1)
        s[2, 1:10] = todaynow - Dates.Month(1)
        s[3, 1:10] = todaynow - Dates.Day(14)
        s[4, 1:10] = todaynow - Dates.Day(5)
        s[5, 1:10] = todaynow - Dates.Day(1)
        s[6, 1:10] = todaynow
        s[7, 1:10] = todaynow + Dates.Day(1)
        s[8, 1:10] = todaynow + Dates.Day(14)
        s[9, 1:10] = todaynow + Dates.Month(1)
        s[10, 1:10] = todaynow + Dates.Year(1)

        @test_throws XLSX.XLSXError XLSX.setConditionalFormat(s, "A1:A5", :timePeriod; notAKeyword="x")
        @test XLSX.setConditionalFormat(s, "A1:A5,C1:C5", :timePeriod) == 0
        @test XLSX.setConditionalFormat(s, [2, 3, 8], 1:3, :timePeriod) == 0 # Vectors may be non-contiguous
        @test XLSX.setConditionalFormat(s, 2, 1:3:7, :timePeriod) == 0 # StepRange is non-contiguous
        @test XLSX.setConditionalFormat(s, "2:2", :timePeriod) == 0
        @test XLSX.setConditionalFormat(s, 2, :, :timePeriod; dxStyle="greenfilltext") == 0
        @test_throws XLSX.XLSXError XLSX.setConditionalFormat(s, 1:10, 1:3, :timePeriod;
            operator="madeUp",
            stopIfTrue="true",
            fill=["pattern" => "none", "bgColor" => "FFFFC7CE"],
            border=["style" => "thick", "color" => "coral"],
            font=["color" => "blue", "bold" => "true", "strike" => "true"]
        )
        @test XLSX.setConditionalFormat(s, 1:10, 1:3, :timePeriod;
            operator="today",
            stopIfTrue="true",
            fill=["pattern" => "none", "bgColor" => "FFFFC7CE"],
            border=["style" => "thick", "color" => "coral"],
            font=["color" => "blue", "bold" => "true", "strike" => "true"]
        ) == 0
        @test XLSX.setConditionalFormat(s, 1:10, 1:3, :timePeriod;
            operator="yesterday",
            stopIfTrue="true",
            fill=["pattern" => "lightVertical", "fgColor" => "grey", "bgColor" => "FFFFC7CE"],
            border=["style" => "thick", "color" => "coral"],
            font=["color" => "blue", "bold" => "true", "strike" => "true"]
        ) == 0
        @test XLSX.setConditionalFormat(s, 1:10, 1:3, :timePeriod;
            operator="tomorrow",
            stopIfTrue="true",
            fill=["pattern" => "none", "bgColor" => "FFFFC7CE"],
            border=["style" => "thick", "color" => "coral"],
            font=["color" => "blue", "bold" => "true", "strike" => "true"]
        ) == 0
        @test XLSX.setConditionalFormat(s, 1:10, 1:3, :timePeriod;
            operator="lastMonth",
            stopIfTrue="true",
            fill=["pattern" => "none", "bgColor" => "pink"],
            border=["style" => "thick", "color" => "coral"],
            font=["color" => "blue", "bold" => "true", "italic" => "true"]
        ) == 0
        @test XLSX.setConditionalFormat(s, 1:10, 1:3, :timePeriod;
            operator="thisMonth",
            stopIfTrue="true",
            fill=["pattern" => "none", "bgColor" => "FFCC4411"],
            border=["style" => "thick", "color" => "coral"],
            font=["color" => "yellow", "bold" => "true", "strike" => "true"]
        ) == 0
        @test XLSX.setConditionalFormat(s, 1:10, 1:3, :timePeriod;
            operator="nextMonth",
            stopIfTrue="true",
            fill=["pattern" => "none", "bgColor" => "FFFFCFCE"],
            border=["style" => "thick", "color" => "coral"],
            font=["color" => "yellow", "bold" => "true", "strike" => "true"]
        ) == 0
        @test XLSX.setConditionalFormat(s, 1:10, 1:3, :timePeriod;
            operator="last7Days",
            stopIfTrue="true",
            fill=["pattern" => "none", "bgColor" => "yellow"],
            border=["style" => "thick", "color" => "coral"],
            font=["color" => "green", "bold" => "true", "italic" => "true"]
        ) == 0
        @test sort!(XLSX.getConditionalFormats(s), by=x -> x.second.priority) == [
            XLSX.NonContiguousRange("Sheet1!A1:A5,Sheet1!C1:C5") => (type="timePeriod", priority=1),
            XLSX.NonContiguousRange("Sheet1!A2:C3,Sheet1!A8:C8") => (type="timePeriod", priority=2),
            XLSX.NonContiguousRange("Sheet1!A2,Sheet1!D2,Sheet1!G2") => (type="timePeriod", priority=3),
            XLSX.CellRange("A2:J2") => (type="timePeriod", priority=4),
            XLSX.CellRange("A2:J2") => (type="timePeriod", priority=5),
            XLSX.CellRange("A1:C10") => (type="timePeriod", priority=6),
            XLSX.CellRange("A1:C10") => (type="timePeriod", priority=7),
            XLSX.CellRange("A1:C10") => (type="timePeriod", priority=8),
            XLSX.CellRange("A1:C10") => (type="timePeriod", priority=9),
            XLSX.CellRange("A1:C10") => (type="timePeriod", priority=10),
            XLSX.CellRange("A1:C10") => (type="timePeriod", priority=11),
            XLSX.CellRange("A1:C10") => (type="timePeriod", priority=12)
        ]


        @test XLSX.setConditionalFormat(s, "A1", :timePeriod) == 0
        @test XLSX.setConditionalFormat(s, "A1:C3", :timePeriod) == 0
        @test XLSX.setConditionalFormat(s, "Sheet1!A1", :timePeriod) == 0
        @test XLSX.setConditionalFormat(s, "Sheet1!A1:A2", :timePeriod) == 0
        @test XLSX.setConditionalFormat(s, "Sheet1!1:2", :timePeriod) == 0
        @test XLSX.setConditionalFormat(s, "2:4", :timePeriod) == 0
        @test XLSX.setConditionalFormat(s, "A:C", :timePeriod) == 0
        @test XLSX.setConditionalFormat(s, "Sheet1!A:C", :timePeriod) == 0
        @test XLSX.setConditionalFormat(f, "Sheet1!1:2", :timePeriod) == 0
        @test XLSX.setConditionalFormat(f, "Sheet1!A:C", :timePeriod) == 0
        @test XLSX.setConditionalFormat(s, :, 1:3, :timePeriod) == 0
        @test XLSX.setConditionalFormat(s, 1:3, :, :timePeriod) == 0
        @test XLSX.setConditionalFormat(s, "2:4", :timePeriod) == 0
        @test XLSX.setConditionalFormat(s, "A:C", :timePeriod) == 0
        @test XLSX.setConditionalFormat(s, "Sheet1!A:C", :timePeriod) == 0
        @test XLSX.setConditionalFormat(s, :, :timePeriod) == 0
        @test XLSX.setConditionalFormat(s, :, :, :timePeriod) == 0
        @test length(XLSX.getConditionalFormats(s)) == 29
        @test sort!(XLSX.getConditionalFormats(s), by=x -> x.second.priority) == [
            XLSX.NonContiguousRange("Sheet1!A1:A5,Sheet1!C1:C5") => (type="timePeriod", priority=1),
            XLSX.NonContiguousRange("Sheet1!A2:C3,Sheet1!A8:C8") => (type="timePeriod", priority=2),
            XLSX.NonContiguousRange("Sheet1!A2,Sheet1!D2,Sheet1!G2") => (type="timePeriod", priority=3),
            XLSX.CellRange("A2:J2") => (type="timePeriod", priority=4),
            XLSX.CellRange("A2:J2") => (type="timePeriod", priority=5),
            XLSX.CellRange("A1:C10") => (type="timePeriod", priority=6),
            XLSX.CellRange("A1:C10") => (type="timePeriod", priority=7),
            XLSX.CellRange("A1:C10") => (type="timePeriod", priority=8),
            XLSX.CellRange("A1:C10") => (type="timePeriod", priority=9),
            XLSX.CellRange("A1:C10") => (type="timePeriod", priority=10),
            XLSX.CellRange("A1:C10") => (type="timePeriod", priority=11),
            XLSX.CellRange("A1:C10") => (type="timePeriod", priority=12),
            XLSX.CellRange("A1:A1") => (type="timePeriod", priority=13),
            XLSX.CellRange("A1:C3") => (type="timePeriod", priority=14),
            XLSX.CellRange("A1:A1") => (type="timePeriod", priority=15),
            XLSX.CellRange("A1:A2") => (type="timePeriod", priority=16),
            XLSX.CellRange("A1:J2") => (type="timePeriod", priority=17),
            XLSX.CellRange("A2:J4") => (type="timePeriod", priority=18),
            XLSX.CellRange("A1:C10") => (type="timePeriod", priority=19),
            XLSX.CellRange("A1:C10") => (type="timePeriod", priority=20),
            XLSX.CellRange("A1:J2") => (type="timePeriod", priority=21),
            XLSX.CellRange("A1:C10") => (type="timePeriod", priority=22),
            XLSX.CellRange("A1:C10") => (type="timePeriod", priority=23),
            XLSX.CellRange("A1:J3") => (type="timePeriod", priority=24),
            XLSX.CellRange("A2:J4") => (type="timePeriod", priority=25),
            XLSX.CellRange("A1:C10") => (type="timePeriod", priority=26),
            XLSX.CellRange("A1:C10") => (type="timePeriod", priority=27),
            XLSX.CellRange("A1:J10") => (type="timePeriod", priority=28),
            XLSX.CellRange("A1:J10") => (type="timePeriod", priority=29)
        ]



        SAVE_FILES && save_outfile(f)

        f = XLSX.newxlsx()
        s = f[1]
        s[1, 1:10] = todaynow - Dates.Year(1)
        s[2, 1:10] = todaynow - Dates.Month(1)
        s[3, 1:10] = todaynow - Dates.Day(14)
        s[4, 1:10] = todaynow - Dates.Day(5)
        s[5, 1:10] = todaynow - Dates.Day(1)
        s[6, 1:10] = todaynow
        s[7, 1:10] = todaynow + Dates.Day(1)
        s[8, 1:10] = todaynow + Dates.Day(14)
        s[9, 1:10] = todaynow + Dates.Month(1)
        s[10, 1:10] = todaynow + Dates.Year(1)
        XLSX.setConditionalFormat(s, "A1:A5", :timePeriod)
        XLSX.setConditionalFormat(s, :, 2, :timePeriod; dxStyle="redborder")
        XLSX.setConditionalFormat(s, "Sheet1!E:E", :timePeriod; dxStyle="redfilltext")
        XLSX.setConditionalFormat(s, 1:5, 3:4, :timePeriod;
            operator="lastWeek",
            fill=["pattern" => "none", "bgColor" => "yellow"],
            format=["format" => "0.0"],
            font=["color" => "green"],
            border=["style" => "thick", "color" => "coral"]
        ) == 0

        @test sort!(XLSX.getConditionalFormats(s), by=x -> x.second.priority, rev=true) == [
            XLSX.CellRange("C1:D5") => (type="timePeriod", priority=4),
            XLSX.CellRange("E1:E10") => (type="timePeriod", priority=3),
            XLSX.CellRange("B1:B10") => (type="timePeriod", priority=2),
            XLSX.CellRange("A1:A5") => (type="timePeriod", priority=1)
        ]
        SAVE_FILES && save_outfile(f)

        f = XLSX.newxlsx()
        s = f[1]
        s[1, 1:10] = todaynow - Dates.Year(1)
        s[2, 1:10] = todaynow - Dates.Month(1)
        s[3, 1:10] = todaynow - Dates.Day(14)
        s[4, 1:10] = todaynow - Dates.Day(5)
        s[5, 1:10] = todaynow - Dates.Day(1)
        s[6, 1:10] = todaynow
        s[7, 1:10] = todaynow + Dates.Day(1)
        s[8, 1:10] = todaynow + Dates.Day(14)
        s[9, 1:10] = todaynow + Dates.Month(1)
        s[10, 1:10] = todaynow + Dates.Year(1)

        @test XLSX.setConditionalFormat(s, :, 1:4, :timePeriod;
            operator="thisWeek",
            fill=["pattern" => "none", "bgColor" => "yellow"],
            format=["format" => "0.0"],
            font=["color" => "green", "under" => "double"],
            border=["style" => "thin", "color" => "coral"]
        ) == 0
        SAVE_FILES && save_outfile(f)

        f = XLSX.newxlsx()
        s = f[1]
        s[1, 1:10] = todaynow - Dates.Year(1)
        s[2, 1:10] = todaynow - Dates.Month(1)
        s[3, 1:10] = todaynow - Dates.Day(14)
        s[4, 1:10] = todaynow - Dates.Day(5)
        s[5, 1:10] = todaynow - Dates.Day(1)
        s[6, 1:10] = todaynow
        s[7, 1:10] = todaynow + Dates.Day(1)
        s[8, 1:10] = todaynow + Dates.Day(14)
        s[9, 1:10] = todaynow + Dates.Month(1)
        s[10, 1:10] = todaynow + Dates.Year(1)
        XLSX.addDefinedName(s, "myRange", "A1:E5")
        @test XLSX.setConditionalFormat(s, "myRange", :timePeriod;
            operator="nextWeek",
            fill=["pattern" => "none", "bgColor" => "yellow"],
            format=["format" => "0.0"],
            font=["color" => "green"],
            border=["style" => "medium", "color" => "cyan"]
        ) == 0
        XLSX.addDefinedName(s, "myNCRange", "C1:C5,D1:D5")
        @test XLSX.setConditionalFormat(s, "myNCRange", :timePeriod; # Non-contiguous ranges not allowed
            operator="lastWeek",
            fill=["pattern" => "none", "bgColor" => "yellow"],
            format=["format" => "0.0"],
            font=["color" => "green"],
            border=["style" => "hair", "color" => "cyan"]
        ) == 0
        SAVE_FILES && save_outfile(f)

    end

    @testset "expression" begin
        f = XLSX.newxlsx()
        s = f[1]
        for i = 1:10
            for j = 1:10
                s[i, j] = i * j
            end
        end

        @test_throws XLSX.XLSXError XLSX.setConditionalFormat(s, "A1:A5", :expression; notAKeyword="x")
        @test XLSX.setConditionalFormat(s, "A1:A5,C1:C5", :expression; formula="A1>3") == 0
        @test XLSX.setConditionalFormat(s, [2, 3, 8], 1:3, :expression; formula="A1 > 11") == 0
        @test XLSX.setConditionalFormat(s, 2, 1:3:7, :expression; formula="A1 < 7") == 0
        @test XLSX.setConditionalFormat(s, "2:2", :expression; formula="A1 = 16") == 0
        @test XLSX.setConditionalFormat(s, 2, :, :expression; formula="A1 < 16", dxStyle="greenfilltext") == 0
        @test_throws XLSX.XLSXError XLSX.setConditionalFormat(s, 1:10, 1:3, :expression;
            stopIfTrue="true",
            fill=["pattern" => "none", "bgColor" => "FFFFC7CE"],
            border=["style" => "thick", "color" => "coral"],
            font=["color" => "blue", "bold" => "true", "strike" => "true"]
        ) == 0
        @test XLSX.setConditionalFormat(s, 1:10, 1:3, :expression;
            formula="A1 > 15",
            stopIfTrue="true",
            fill=["pattern" => "none", "bgColor" => "FFFFC7CE"],
            border=["style" => "thick", "color" => "coral"],
            font=["color" => "blue", "bold" => "true", "strike" => "true"]
        ) == 0
        @test XLSX.setConditionalFormat(s, 1:10, 1:3, :expression;
            formula="iseven(A1)",
            stopIfTrue="true",
            fill=["pattern" => "lightVertical", "fgColor" => "grey", "bgColor" => "FFFFC7CE"],
            border=["style" => "thick", "color" => "coral"],
            font=["color" => "blue", "bold" => "true", "strike" => "true"]
        ) == 0
        @test XLSX.setConditionalFormat(s, 1:10, 1:3, :expression;
            formula="A1 < 10",
            stopIfTrue="true",
            fill=["pattern" => "none", "bgColor" => "FFFFC7CE"],
            border=["style" => "thick", "color" => "coral"],
            font=["color" => "blue", "bold" => "true", "strike" => "true"]
        ) == 0
        @test XLSX.setConditionalFormat(s, 1:10, 1:3, :expression;
            formula="A1 < 5",
            stopIfTrue="true",
            fill=["pattern" => "none", "bgColor" => "pink"],
            border=["style" => "thick", "color" => "coral"],
            font=["color" => "blue", "bold" => "true", "italic" => "true"]
        ) == 0
        @test sort!(XLSX.getConditionalFormats(s), by=x -> x.second.priority) == [
            XLSX.NonContiguousRange("Sheet1!A1:A5,Sheet1!C1:C5") => (type="expression", priority=1),
            XLSX.NonContiguousRange("Sheet1!A2:C3,Sheet1!A8:C8") => (type="expression", priority=2),
            XLSX.NonContiguousRange("Sheet1!A2,Sheet1!D2,Sheet1!G2") => (type="expression", priority=3),
            XLSX.CellRange("A2:J2") => (type="expression", priority=4),
            XLSX.CellRange("A2:J2") => (type="expression", priority=5),
            XLSX.CellRange("A1:C10") => (type="expression", priority=6),
            XLSX.CellRange("A1:C10") => (type="expression", priority=7),
            XLSX.CellRange("A1:C10") => (type="expression", priority=8),
            XLSX.CellRange("A1:C10") => (type="expression", priority=9)
        ]


        @test XLSX.setConditionalFormat(s, "A1", :expression; formula="iseven(A1)") == 0
        @test XLSX.setConditionalFormat(s, "A1:C3", :expression; formula="iseven(A1)") == 0
        @test XLSX.setConditionalFormat(s, "Sheet1!A1", :expression; formula="iseven(A1)") == 0
        @test XLSX.setConditionalFormat(s, "Sheet1!A1:A2", :expression; formula="iseven(A1)") == 0
        @test XLSX.setConditionalFormat(s, "Sheet1!1:2", :expression; formula="iseven(A1)") == 0
        @test XLSX.setConditionalFormat(s, "2:4", :expression; formula="iseven(A1)") == 0
        @test XLSX.setConditionalFormat(s, "A:C", :expression; formula="iseven(A1)") == 0
        @test XLSX.setConditionalFormat(s, "Sheet1!A:C", :expression; formula="iseven(A1)") == 0
        @test XLSX.setConditionalFormat(f, "Sheet1!1:2", :expression; formula="iseven(A1)") == 0
        @test XLSX.setConditionalFormat(f, "Sheet1!A:C", :expression; formula="iseven(A1)") == 0
        @test XLSX.setConditionalFormat(s, :, 1:3, :expression; formula="iseven(A1)") == 0
        @test XLSX.setConditionalFormat(s, 1:3, :, :expression; formula="iseven(A1)") == 0
        @test XLSX.setConditionalFormat(s, "2:4", :expression; formula="iseven(A1)") == 0
        @test XLSX.setConditionalFormat(s, "A:C", :expression; formula="iseven(A1)") == 0
        @test XLSX.setConditionalFormat(s, "Sheet1!A:C", :expression; formula="iseven(A1)") == 0
        @test XLSX.setConditionalFormat(s, :, :expression; formula="iseven(A1)") == 0
        @test XLSX.setConditionalFormat(s, :, :, :expression; formula="iseven(A1)") == 0
        @test length(XLSX.getConditionalFormats(s)) == 26
        @test sort!(XLSX.getConditionalFormats(s), by=x -> x.second.priority, rev=true) == [
            XLSX.CellRange("A1:J10") => (type="expression", priority=26),
            XLSX.CellRange("A1:J10") => (type="expression", priority=25),
            XLSX.CellRange("A1:C10") => (type="expression", priority=24),
            XLSX.CellRange("A1:C10") => (type="expression", priority=23),
            XLSX.CellRange("A2:J4") => (type="expression", priority=22),
            XLSX.CellRange("A1:J3") => (type="expression", priority=21),
            XLSX.CellRange("A1:C10") => (type="expression", priority=20),
            XLSX.CellRange("A1:C10") => (type="expression", priority=19),
            XLSX.CellRange("A1:J2") => (type="expression", priority=18),
            XLSX.CellRange("A1:C10") => (type="expression", priority=17),
            XLSX.CellRange("A1:C10") => (type="expression", priority=16),
            XLSX.CellRange("A2:J4") => (type="expression", priority=15),
            XLSX.CellRange("A1:J2") => (type="expression", priority=14),
            XLSX.CellRange("A1:A2") => (type="expression", priority=13),
            XLSX.CellRange("A1:A1") => (type="expression", priority=12),
            XLSX.CellRange("A1:C3") => (type="expression", priority=11),
            XLSX.CellRange("A1:A1") => (type="expression", priority=10),
            XLSX.CellRange("A1:C10") => (type="expression", priority=9),
            XLSX.CellRange("A1:C10") => (type="expression", priority=8),
            XLSX.CellRange("A1:C10") => (type="expression", priority=7),
            XLSX.CellRange("A1:C10") => (type="expression", priority=6),
            XLSX.CellRange("A2:J2") => (type="expression", priority=5),
            XLSX.CellRange("A2:J2") => (type="expression", priority=4),
            XLSX.NonContiguousRange("Sheet1!A2,Sheet1!D2,Sheet1!G2") => (type="expression", priority=3),
            XLSX.NonContiguousRange("Sheet1!A2:C3,Sheet1!A8:C8") => (type="expression", priority=2),
            XLSX.NonContiguousRange("Sheet1!A1:A5,Sheet1!C1:C5") => (type="expression", priority=1)
        ]

        SAVE_FILES && save_outfile(f)

        f = XLSX.newxlsx()
        s = f[1]
        for i = 1:10
            for j = 1:10
                s[i, j] = i * j
            end
        end
        XLSX.setConditionalFormat(s, "A1:A5", :expression; formula="A1=1")
        XLSX.setConditionalFormat(s, :, 2, :expression; formula="A1=1", dxStyle="redborder")
        XLSX.setConditionalFormat(s, "Sheet1!E:E", :expression; formula="A1=1", dxStyle="redfilltext")
        XLSX.setConditionalFormat(s, 1:5, 3:4, :expression;
            formula="A1=1",
            fill=["pattern" => "none", "bgColor" => "yellow"],
            format=["format" => "0.0"],
            font=["color" => "green"],
            border=["style" => "thick", "color" => "coral"]
        ) == 0

        @test sort!(XLSX.getConditionalFormats(s), by=x -> x.second.priority, rev=true) == [
            XLSX.CellRange("C1:D5") => (type="expression", priority=4),
            XLSX.CellRange("E1:E10") => (type="expression", priority=3),
            XLSX.CellRange("B1:B10") => (type="expression", priority=2),
            XLSX.CellRange("A1:A5") => (type="expression", priority=1)
        ]
        SAVE_FILES && save_outfile(f)

        f = XLSX.newxlsx()
        s = f[1]
        for i = 1:10
            for j = 1:10
                s[i, j] = i * j
            end
        end
        @test XLSX.setConditionalFormat(s, :, 1:4, :expression;
            formula="A1 > \$E\$3",
            fill=["pattern" => "none", "bgColor" => "yellow"],
            format=["format" => "0.0"],
            font=["color" => "green", "under" => "double"],
            border=["style" => "thin", "color" => "coral"]
        ) == 0
        SAVE_FILES && save_outfile(f)

        f = XLSX.newxlsx()
        s = f[1]
        for i = 1:10
            for j = 1:10
                s[i, j] = i * j
            end
        end
        XLSX.addDefinedName(f, "myTest", "Sheet1!L11")
        s["L11"] = 70
        XLSX.addDefinedName(s, "myRange", "F6:J10")

        @test XLSX.setConditionalFormat(s, "myRange", :expression;
            formula="E5 > myTest",
            fill=["pattern" => "none", "bgColor" => "yellow"],
            format=["format" => "0.0"],
            font=["color" => "green"],
            border=["style" => "medium", "color" => "cyan"]
        ) == 0
        XLSX.addDefinedName(s, "myNCRange", "C1:C5,D1:D5")
        @test XLSX.setConditionalFormat(s, "myNCRange", :expression; # Non-contiguous ranges not allowed
            formula="C4 < myTest",
            fill=["pattern" => "none", "bgColor" => "yellow"],
            format=["format" => "0.0"],
            font=["color" => "green"],
            border=["style" => "hair", "color" => "cyan"]
        ) == 0
        SAVE_FILES && save_outfile(f)

    end

    @testset "containsErrors" begin
        f = XLSX.newxlsx()
        s = f[1]
        for i = 1:10
            for j = 1:10
                s[i, j] = i * j
            end
        end

        @test_throws XLSX.XLSXError XLSX.setConditionalFormat(s, "A1:A5", :containsErrors; notAKeyword="x")
        @test XLSX.setConditionalFormat(s, "A1:A5,C1:C5", :containsErrors) == 0
        @test XLSX.setConditionalFormat(s, [2, 3, 8], 1:3, :containsErrors) == 0
        @test XLSX.setConditionalFormat(s, 2, 1:3:7, :containsErrors) == 0
        @test XLSX.setConditionalFormat(s, "2:2", :containsErrors) == 0
        @test XLSX.setConditionalFormat(s, 2, :, :containsErrors; dxStyle="greenfilltext") == 0
        @test XLSX.setConditionalFormat(s, 1:10, 1:3, :containsErrors;
            stopIfTrue="true",
            fill=["pattern" => "none", "bgColor" => "FFFFC7CE"],
            border=["style" => "thick", "color" => "coral"],
            font=["color" => "blue", "bold" => "true", "strike" => "true"]
        ) == 0
        @test XLSX.setConditionalFormat(s, 1:10, 1:3, :notContainsErrors;
            stopIfTrue="true",
            fill=["pattern" => "lightVertical", "fgColor" => "grey", "bgColor" => "FFFFC7CE"],
            border=["style" => "thick", "color" => "coral"],
            font=["color" => "blue", "bold" => "true", "strike" => "true"]
        ) == 0
        @test XLSX.setConditionalFormat(s, 1:10, 1:3, :containsBlanks;
            stopIfTrue="true",
            fill=["pattern" => "none", "bgColor" => "FFFFC7CE"],
            border=["style" => "thick", "color" => "coral"],
            font=["color" => "blue", "bold" => "true", "strike" => "true"]
        ) == 0
        @test XLSX.setConditionalFormat(s, 1:10, 1:3, :notContainsBlanks;
            stopIfTrue="true",
            fill=["pattern" => "none", "bgColor" => "pink"],
            border=["style" => "thick", "color" => "coral"],
            font=["color" => "blue", "bold" => "true", "italic" => "true"]
        ) == 0
        @test XLSX.setConditionalFormat(s, 1:10, 1:3, :uniqueValues;
            stopIfTrue="true",
            fill=["pattern" => "none", "bgColor" => "FFFFCFCE"],
            border=["style" => "thick", "color" => "coral"],
            font=["color" => "yellow", "bold" => "true", "strike" => "true"]
        ) == 0
        @test XLSX.setConditionalFormat(s, 1:10, 1:3, :duplicateValues;
            stopIfTrue="true",
            fill=["pattern" => "none", "bgColor" => "yellow"],
            border=["style" => "thick", "color" => "coral"],
            font=["color" => "green", "bold" => "true", "italic" => "true"]
        ) == 0
        @test sort!(XLSX.getConditionalFormats(s), by=x -> x.second.priority) == [
            XLSX.NonContiguousRange("Sheet1!A1:A5,Sheet1!C1:C5") => (type="containsErrors", priority=1),
            XLSX.NonContiguousRange("Sheet1!A2:C3,Sheet1!A8:C8") => (type="containsErrors", priority=2),
            XLSX.NonContiguousRange("Sheet1!A2,Sheet1!D2,Sheet1!G2") => (type="containsErrors", priority=3),
            XLSX.CellRange("A2:J2") => (type="containsErrors", priority=4),
            XLSX.CellRange("A2:J2") => (type="containsErrors", priority=5),
            XLSX.CellRange("A1:C10") => (type="containsErrors", priority=6),
            XLSX.CellRange("A1:C10") => (type="notContainsErrors", priority=7),
            XLSX.CellRange("A1:C10") => (type="containsBlanks", priority=8),
            XLSX.CellRange("A1:C10") => (type="notContainsBlanks", priority=9),
            XLSX.CellRange("A1:C10") => (type="uniqueValues", priority=10),
            XLSX.CellRange("A1:C10") => (type="duplicateValues", priority=11)
        ]


        @test XLSX.setConditionalFormat(s, "A1", :containsErrors) == 0
        @test XLSX.setConditionalFormat(s, "A1:C3", :notContainsErrors) == 0
        @test XLSX.setConditionalFormat(s, "Sheet1!A1", :containsBlanks) == 0
        @test XLSX.setConditionalFormat(s, "Sheet1!A1:A2", :notContainsBlanks) == 0
        @test XLSX.setConditionalFormat(s, "Sheet1!1:2", :uniqueValues) == 0
        @test XLSX.setConditionalFormat(s, "2:4", :duplicateValues) == 0
        @test XLSX.setConditionalFormat(s, "A:C", :containsErrors) == 0
        @test XLSX.setConditionalFormat(s, "Sheet1!A:C", :containsErrors) == 0
        @test XLSX.setConditionalFormat(f, "Sheet1!1:2", :containsErrors) == 0
        @test XLSX.setConditionalFormat(f, "Sheet1!A:C", :containsErrors) == 0
        @test XLSX.setConditionalFormat(s, :, 1:3, :containsErrors) == 0
        @test XLSX.setConditionalFormat(s, 1:3, :, :notContainsErrors) == 0
        @test XLSX.setConditionalFormat(s, "2:4", :containsBlanks) == 0
        @test XLSX.setConditionalFormat(s, "A:C", :notContainsBlanks) == 0
        @test XLSX.setConditionalFormat(s, "Sheet1!A:C", :containsErrors) == 0
        @test XLSX.setConditionalFormat(s, :, :uniqueValues) == 0
        @test XLSX.setConditionalFormat(s, :, :, :duplicateValues) == 0
        @test length(XLSX.getConditionalFormats(s)) == 28
        @test sort!(XLSX.getConditionalFormats(s), by=x -> x.second.priority, rev=true) == [
            XLSX.CellRange("A1:J10") => (type="duplicateValues", priority=28),
            XLSX.CellRange("A1:J10") => (type="uniqueValues", priority=27),
            XLSX.CellRange("A1:C10") => (type="containsErrors", priority=26),
            XLSX.CellRange("A1:C10") => (type="notContainsBlanks", priority=25),
            XLSX.CellRange("A2:J4") => (type="containsBlanks", priority=24),
            XLSX.CellRange("A1:J3") => (type="notContainsErrors", priority=23),
            XLSX.CellRange("A1:C10") => (type="containsErrors", priority=22),
            XLSX.CellRange("A1:C10") => (type="containsErrors", priority=21),
            XLSX.CellRange("A1:J2") => (type="containsErrors", priority=20),
            XLSX.CellRange("A1:C10") => (type="containsErrors", priority=19),
            XLSX.CellRange("A1:C10") => (type="containsErrors", priority=18),
            XLSX.CellRange("A2:J4") => (type="duplicateValues", priority=17),
            XLSX.CellRange("A1:J2") => (type="uniqueValues", priority=16),
            XLSX.CellRange("A1:A2") => (type="notContainsBlanks", priority=15),
            XLSX.CellRange("A1:A1") => (type="containsBlanks", priority=14),
            XLSX.CellRange("A1:C3") => (type="notContainsErrors", priority=13),
            XLSX.CellRange("A1:A1") => (type="containsErrors", priority=12),
            XLSX.CellRange("A1:C10") => (type="duplicateValues", priority=11),
            XLSX.CellRange("A1:C10") => (type="uniqueValues", priority=10),
            XLSX.CellRange("A1:C10") => (type="notContainsBlanks", priority=9),
            XLSX.CellRange("A1:C10") => (type="containsBlanks", priority=8),
            XLSX.CellRange("A1:C10") => (type="notContainsErrors", priority=7),
            XLSX.CellRange("A1:C10") => (type="containsErrors", priority=6),
            XLSX.CellRange("A2:J2") => (type="containsErrors", priority=5),
            XLSX.CellRange("A2:J2") => (type="containsErrors", priority=4),
            XLSX.NonContiguousRange("Sheet1!A2,Sheet1!D2,Sheet1!G2") => (type="containsErrors", priority=3),
            XLSX.NonContiguousRange("Sheet1!A2:C3,Sheet1!A8:C8") => (type="containsErrors", priority=2),
            XLSX.NonContiguousRange("Sheet1!A1:A5,Sheet1!C1:C5") => (type="containsErrors", priority=1)
        ]

        SAVE_FILES && save_outfile(f)

        f = XLSX.newxlsx()
        s = f[1]
        for i = 1:10
            for j = 1:10
                s[i, j] = i * j
            end
        end
        XLSX.setConditionalFormat(s, "A1:A5", :containsErrors)
        XLSX.setConditionalFormat(s, :, 2, :notContainsErrors; dxStyle="redborder")
        XLSX.setConditionalFormat(s, "Sheet1!E:E", :containsBlanks; dxStyle="redfilltext")
        XLSX.setConditionalFormat(s, 1:5, 3:4, :uniqueValues;
            fill=["pattern" => "none", "bgColor" => "yellow"],
            format=["format" => "0.0"],
            font=["color" => "green"],
            border=["style" => "thick", "color" => "coral"]
        ) == 0

        @test sort!(XLSX.getConditionalFormats(s), by=x -> x.second.priority, rev=true) == [
            XLSX.CellRange("C1:D5") => (type="uniqueValues", priority=4),
            XLSX.CellRange("E1:E10") => (type="containsBlanks", priority=3),
            XLSX.CellRange("B1:B10") => (type="notContainsErrors", priority=2),
            XLSX.CellRange("A1:A5") => (type="containsErrors", priority=1)
        ]
        SAVE_FILES && save_outfile(f)

        f = XLSX.newxlsx()
        s = f[1]
        for i = 1:10
            for j = 1:10
                s[i, j] = i * j
            end
        end

        @test XLSX.setConditionalFormat(s, :, 1:4, :containsErrors;
            fill=["pattern" => "none", "bgColor" => "yellow"],
            format=["format" => "0.0"],
            font=["color" => "green", "under" => "double"],
            border=["style" => "thin", "color" => "coral"]
        ) == 0
        SAVE_FILES && save_outfile(f)

        f = XLSX.newxlsx()
        s = f[1]
        for i = 1:10
            for j = 1:10
                s[i, j] = i * j
            end
        end
        XLSX.addDefinedName(s, "myRange", "A1:E5")
        @test XLSX.setConditionalFormat(s, "myRange", :containsErrors;
            fill=["pattern" => "none", "bgColor" => "yellow"],
            format=["format" => "0.0"],
            font=["color" => "green"],
            border=["style" => "medium", "color" => "cyan"]
        ) == 0
        XLSX.addDefinedName(s, "myNCRange", "C1:C5,D1:D5")
        @test XLSX.setConditionalFormat(s, "myNCRange", :containsErrors; # Non-contiguous ranges not allowed
            fill=["pattern" => "none", "bgColor" => "yellow"],
            format=["format" => "0.0"],
            font=["color" => "green"],
            border=["style" => "hair", "color" => "cyan"]
        ) == 0
        SAVE_FILES && save_outfile(f)

    end

    @testset "getConditionalFormats does not break subsequent cell reads (issue #425)" begin

        function build_cf_workbook()
            path = tempname() * ".xlsx"
            f = XLSX.newxlsx()
            s = f[1]
            for i in 1:5, j in 1:5
                s[i, j] = i + j
            end
            XLSX.setConditionalFormat(s, "A1:E1", :dataBar)
            XLSX.writexlsx(path, f)
            SAVE_FILES && save_outfile(f)
            return path
        end

        @testset "getConditionalFormats then getdata (originally-failing order)" begin
            path = build_cf_workbook()
            XLSX.openxlsx(path) do xf
                ws = xf[1]
                cfs = XLSX.getConditionalFormats(ws)
                @test length(cfs) == 1

                data = XLSX.getdata(ws)
                @test data[1, 1] == 2
                @test size(data) == (5, 5)
            end
            rm(path; force=true)
        end

        @testset "getdata then getConditionalFormats (originally-working order, must not regress)" begin
            path = build_cf_workbook()
            XLSX.openxlsx(path) do xf
                ws = xf[1]
                data = XLSX.getdata(ws)
                @test data[1, 1] == 2

                cfs = XLSX.getConditionalFormats(ws)
                @test length(cfs) == 1
            end
            rm(path; force=true)
        end

        @testset "getConditionalFormats then eachrow" begin
            path = build_cf_workbook()
            XLSX.openxlsx(path) do xf
                ws = xf[1]
                XLSX.getConditionalFormats(ws)

                rows = collect(XLSX.eachrow(ws))
                @test length(rows) == 5
                @test XLSX.getdata(rows[1], 1) == 2
            end
            rm(path; force=true)
        end

        @testset "getConditionalFormats then readtable" begin
            path = build_cf_workbook()
            XLSX.openxlsx(path) do xf
                ws = xf[1]
                XLSX.getConditionalFormats(ws)

                tbl = XLSX.readtable(path, ws.name; header=false)
                @test length(tbl.data) == 5          # 5 columns
                @test length(tbl.data[1]) == 5       # 5 rows
            end
            rm(path; force=true)
        end

        @testset "readxlsx: getConditionalFormats then getdata" begin
            path = build_cf_workbook()
            xf = XLSX.readxlsx(path)
            ws = xf[1]
            cfs = XLSX.getConditionalFormats(ws)
            @test length(cfs) == 1
            data = XLSX.getdata(ws)
            @test data[1, 1] == 2
            rm(path; force=true)
        end

        @testset "getConditionalExtFormats (allExtCfs) does not break subsequent reads" begin
            path = tempname() * ".xlsx"
            f = XLSX.newxlsx()
            s = f[1]
            for i in 1:5, j in 1:5
                s[i, j] = i + j
            end
            # icon sets / top10 / etc. route through the x14 extension block —
            # use whichever of these your API exposes as the ext-format path.
            XLSX.setConditionalFormat(s, "A1:E1", :iconSet; iconset="5Boxes")
            XLSX.writexlsx(path, f)

            XLSX.openxlsx(path) do xf2
                ws = xf2[1]
                XLSX.getConditionalFormats(ws)
                data = XLSX.getdata(ws)
                @test data[1, 1] == 2
            end
            rm(path; force=true)
            SAVE_FILES && save_outfile(f)
        end

    end

    @testset "Non-contiguous conditional formats" begin

        C(r, c) = XLSX.CellRef(r, c)
        expand(areas) = Set(vcat([a isa XLSX.CellRef ? [a] : collect(a) for a in areas]...))

        @testset "_compress" begin
            a = XLSX._compress([C(1, 1)])
            @test length(a) == 1 && a[1] isa XLSX.CellRef

            @test XLSX._compress([C(1, 1), C(2, 1), C(3, 1)]) == [XLSX.CellRange("A1:A3")]
            @test length(XLSX._compress([C(1, 1), C(2, 1), C(5, 1)])) == 2

            block = vec([C(r, c) for r in 1:4, c in 1:3])
            @test XLSX._compress(block) == [XLSX.CellRange("A1:C4")]

            # ragged columns must not merge horizontally
            @test length(XLSX._compress([C(1, 1), C(2, 1), C(1, 2)])) == 2

            cb = [C(r, c) for r in 1:4 for c in 1:4 if iseven(r + c)]
            @test length(XLSX._compress(cb)) == length(cb)

            @test XLSX._compress([C(1, 1), C(1, 1), C(2, 1)]) == [XLSX.CellRange("A1:A2")]
            @test isempty(XLSX._compress(XLSX.CellRef[]))

            # compression preserves the covered cell set
            scattered = [C(r, c) for r in 1:8 for c in 1:5 if (r * c) % 3 == 0]
            @test expand(XLSX._compress(scattered)) == Set(scattered)
        end

        @testset "_band_colors" begin
            @test XLSX._band_colors(["green", "orange", "red"], 3) == ["FF008000", "FFFFA500", "FFFF0000"]
            @test length(XLSX._band_colors(["green", "red"], 5)) == 5
            @test XLSX._band_colors(["green", "red"], 2) == ["FF008000", "FFFF0000"]
            @test XLSX._band_colors("red", 1) == ["FFFF0000"]
            @test XLSX._band_colors(:red, 1) == ["FFFF0000"]

            interp = XLSX._band_colors(["FF008000", "FFFF0000"], 5)
            @test interp[1] == "FF008000" && interp[end] == "FFFF0000"
            @test all(c -> occursin(r"^FF[0-9A-F]{6}$", c), interp)

            @test_throws XLSX.XLSXError XLSX._band_colors(["red", "green", "blue"], 5)
            @test_throws XLSX.XLSXError XLSX._band_colors(["notacolour", "red"], 3)
        end

        @testset "_cf_sqref" begin
            @test XLSX._cf_sqref(XLSX.CellRange("A1:B2")) == "A1:B2"
            xf = XLSX.newxlsx()
            s = xf[1]
            s["A1"] = 1
            s["A3"] = 3
            @test XLSX._cf_sqref(XLSX.NonContiguousRange(s, "A1,A3")) == "A1 A3"
            SAVE_FILES && save_outfile(xf)
        end

        @testset "partition" begin
            xf = XLSX.newxlsx()
            s = xf[1]
            for i in 1:10
                s[i, 1] = i
            end
            s["A11"] = "text"
            s["A12"] = NaN

            p = XLSX.partition(s, "A1:A12", [3, 7])
            @test length(p) == 3
            @test first.(p) == [1, 2, 3]
            @test XLSX._cf_sqref(p[1][2]) == "A1:A2"
            @test XLSX._cf_sqref(p[2][2]) == "A3:A6"
            @test XLSX._cf_sqref(p[3][2]) == "A7:A10"

            all_sq = join((XLSX._cf_sqref(r) for (_, r) in p), " ")
            @test !occursin("A11", all_sq)          # text excluded
            @test !occursin("A12", all_sq)          # NaN excluded

            @test XLSX._cf_sqref(XLSX.partition(s, "A1:A12", [3, 7]; gte=false)[1][2]) == "A1:A3"
            @test XLSX._cf_sqref(XLSX.partition(s, "A1:A12", [3]; gte=[false])[1][2]) == "A1:A3"

            p = XLSX.partition(s, "A1:A10", [3, 7]; labels=[:lo, :mid, :hi])
            @test first.(p) == [:lo, :mid, :hi]
            @test eltype(p) <: Pair{Symbol}

            @test length(XLSX.partition(s, "A1:A10", [100, 200])) == 1   # empty bands dropped

            # key-function method: encounter order, not sorted
            p = XLSX.partition(s, "A1:A10", v -> v isa Real ? (iseven(v) ? :even : :odd) : nothing)
            @test first.(p) == [:odd, :even]

            p = XLSX.partition(s, "A1:A10", v -> v isa Real ? :all : nothing; compress=false)
            @test length(XLSX._cf_areas(p[1][2])) == 10

            @test_throws XLSX.XLSXError XLSX.partition(s, "A1:A10", [7, 3])
            @test_throws XLSX.XLSXError XLSX.partition(s, "A1:A10", [3, 7]; labels=[:a, :b])
            @test_throws XLSX.XLSXError XLSX.partition(s, "A1:A10", [3, 7]; labels=[:a, :b, :a])
            @test_throws XLSX.XLSXError XLSX.partition(s, "A1:A10", [3, 7]; gte=[true])
            SAVE_FILES && save_outfile(xf)
        end

        @testset "setConditionalFormat on a non-contiguous range" begin
            xf = XLSX.newxlsx()
            s = xf[1]
            s["A1"] = 1
            s["A2"] = 2
            s["A3"] = 3

            @test XLSX.setConditionalFormat(s, "A1,A3", :dataBar) == 0
            @test XLSX.setConditionalFormat(s, "A1,A3", :cellIs; operator="greaterThan", value="1") == 0
            @test_throws XLSX.XLSXError XLSX.setConditionalFormat(s, "myNCRange", :cellIs; operator="greaterThan", value="1") # out of range

            cfs = XLSX.getConditionalFormats(s)
            rng = first(first(cfs))
            @test rng isa XLSX.NonContiguousRange
            @test XLSX._cf_sqref(rng) == "A1 A3"

            # a second rule on the same range joins the existing block
            XLSX.setConditionalFormat(s, "A1,A3", :cellIs; operator="greaterThan", value="1", fill=["pattern" => "solid", "bgColor" => "FFFFC7CE"])
            @test length(XLSX.getConditionalFormats(s)) == 3   # 2007 x2 + ext databar

            f = tempname() * ".xlsx"
            XLSX.writexlsx(f, xf; overwrite=true)
            xf2 = XLSX.opentemplate(f)
            rng2 = first(first(XLSX.getConditionalFormats(xf2[1])))
            @test rng2 isa XLSX.NonContiguousRange
            @test XLSX._cf_sqref(rng2) == "A1 A3"
            SAVE_FILES && save_outfile(xf2)
            SAVE_FILES && save_outfile(xf)
        end

        @testset "setColoredDataBars" begin
            xf = XLSX.newxlsx()
            s = xf[1]
            for i in 1:20
                s[i, 1] = (i * 7) % 20 + 1
            end

            p = XLSX.setColoredDataBars(s, "A1:A20"; bands=4)
            @test length(p) == 4
            @test first.(p) == [1, 2, 3, 4]

            f = tempname() * ".xlsx"
            XLSX.writexlsx(f, xf; overwrite=true)
            x = ZipArchives.zip_readentry(ZipArchives.ZipReader(read(f)),
                "xl/worksheets/sheet1.xml", String)

            @test count("<conditionalFormatting ", x) == 4
            @test count("<x14:cfRule ", x) == 4
            @test count("gradient=\"0\"", x) == 4
            @test !occursin("border=\"1\"", x)

            # extLst must follow </dataBar> inside <cfRule>, not sit inside <dataBar>
            @test occursin(r"</dataBar>\s*<extLst>\s*<ext uri=", x)
            @test !occursin(r"<dataBar>\s*<ext ", x)

            # every 2007 rule has an id, and each is matched on the x14 side
            ids = [m.captures[1] for m in eachmatch(r"<x14:id>(\{[^}]+\})</x14:id>", x)]
            @test length(ids) == 4
            @test all(i -> occursin("id=\"$i\"", x), ids)

            # shared axis: exactly one min and one max cfvo value across all bands
            @test length(Set(m.match for m in eachmatch(r"<cfvo type=\"num\" val=\"[^\"]+\"/>", x))) == 2

            # sqrefs are space-separated
            @test !occursin(r"sqref=\"[^\"]*,", x)
            @test occursin(r"sqref=\"[^\"]* [^\"]*\"", x)
            SAVE_FILES && save_outfile(xf)
        end

        @testset "setColoredDataBars options" begin
            xf = XLSX.newxlsx()
            s = xf[1]
            for i in 1:10
                s[i, 1] = i
            end

            # bands=1 degenerates to one contiguous rule
            p = XLSX.setColoredDataBars(s, "A1:A10"; bands=1, colors="steelblue")
            @test length(p) == 1
            @test p[1][2] isa XLSX.CellRange

            # breaks overrides bands
            @test length(XLSX.setColoredDataBars(s, "A1:A10"; bands=9, breaks=[5], colors=["green", "red"])) == 2

            # explicit colours, one per band
            @test length(XLSX.setColoredDataBars(s, "A1:A10"; bands=3, colors=["green", "orange", "red"])) == 3

            # passthrough kwargs and explicit axis
            @test length(XLSX.setColoredDataBars(s, "A1:A10"; bands=2, min_val="0", max_val="100",
                showVal="false", direction="rightToLeft")) == 2

            # gte forwarding
            @test length(XLSX.setColoredDataBars(s, "A1:A10"; breaks=[5], gte=false, colors=["green", "red"])) == 2

            # alternative range argument forms
            @test length(XLSX.setColoredDataBars(s, XLSX.CellRange("A1:A10"); bands=2)) == 2
            @test length(XLSX.setColoredDataBars(s, XLSX.SheetCellRange("Sheet1!A1:A10"); bands=2)) == 2
            SAVE_FILES && save_outfile(xf)
        end

        @testset "setColoredDataBars errors" begin
            xf = XLSX.newxlsx()
            s = xf[1]
            for i in 1:10
                s[i, 1] = i
            end
            s["C1"] = 7
            s["C2"] = 7
            s["C3"] = 7
            s["D1"] = "text"

            @test_throws XLSX.XLSXError XLSX.setColoredDataBars(s, "A1:A10"; bands=0)
            @test_throws XLSX.XLSXError XLSX.setColoredDataBars(s, "A1:A10"; bands=3, min_type="percentile")
            @test_throws XLSX.XLSXError XLSX.setColoredDataBars(s, "A1:A10"; bands=3, max_type="highest")
            @test_throws XLSX.XLSXError XLSX.setColoredDataBars(s, "A1:A10"; bands=3, colors=["red", "green", "blue", "cyan"])
            @test_throws XLSX.XLSXError XLSX.setColoredDataBars(s, "D1:D1")          # no numeric values
            @test_throws XLSX.XLSXError XLSX.setColoredDataBars(s, "C1:C3")          # all values equal
            @test_throws XLSX.XLSXError XLSX.setColoredDataBars(s, "A1:A1000")       # outside dimension
            SAVE_FILES && save_outfile(xf)
        end
        @testset "vector and step-range dispatch" begin
            xf = XLSX.newxlsx()
            s = xf[1]
            for i in 1:6, j in 1:3
                s[i, j] = i * j
            end

            sq(ws) = XLSX._cf_sqref(first(first(XLSX.getConditionalFormats(ws))))

            # rows as a vector, single column
            xf1 = XLSX.newxlsx()
            s1 = xf1[1]
            for i in 1:6; s1[i, 1] = i; end
            @test XLSX.setConditionalFormat(s1, [1, 3, 5], 1, :dataBar) == 0
            @test sq(s1) == "A1 A3 A5"

            # contiguous StepRange collapses to a single area
            xf2 = XLSX.newxlsx()
            s2 = xf2[1]
            for i in 1:6; s2[i, 1] = i; end
            @test XLSX.setConditionalFormat(s2, 1:1:6, 1, :dataBar) == 0
            @test sq(s2) == "A1:A6"

            # vector of columns, contiguous rows -> merges into per-column blocks
            xf3 = XLSX.newxlsx()
            s3 = xf3[1]
            for i in 1:3, j in 1:3; s3[i, j] = i; end
            @test XLSX.setConditionalFormat(s3, 1:3, [1, 3], :dataBar) == 0
            @test sq(s3) == "A1:A3 C1:C3"

            # both vectors
            xf4 = XLSX.newxlsx()
            s4 = xf4[1]
            for i in 1:3, j in 1:3; s4[i, j] = i; end
            @test XLSX.setConditionalFormat(s4, [1, 3], [1, 3], :dataBar) == 0
            @test sq(s4) == "A1 A3 C1 C3"

            # colon forms
            xf5 = XLSX.newxlsx()
            s5 = xf5[1]
            for i in 1:4, j in 1:2; s5[i, j] = i; end
            @test XLSX.setConditionalFormat(s5, [1, 3], :, :dataBar) == 0
            @test sq(s5) == "A1:B1 A3:B3"

            xf6 = XLSX.newxlsx()
            s6 = xf6[1]
            for i in 1:2, j in 1:4; s6[i, j] = i; end
            @test XLSX.setConditionalFormat(s6, :, [1, 3], :dataBar) == 0
            @test sq(s6) == "A1:A2 C1:C2"

            # blanks inside the range stay in the sqref (unlike the style setters)
            xf7 = XLSX.newxlsx()
            s7 = xf7[1]
            s7["A1"] = 1
            s7["A3"] = 3
            s7["A5"] = 5   # A2, A4 empty
            @test XLSX.setConditionalFormat(s7, 1:5, 1, :dataBar) == 0
            @test sq(s7) == "A1:A5"

            # works for a dxf-bearing type too, not just dataBar
            xf8 = XLSX.newxlsx()
            s8 = xf8[1]
            for i in 1:5; s8[i, 1] = i; end
            @test XLSX.setConditionalFormat(s8, [1, 3, 5], 1, :cellIs;
                operator="greaterThan", value="1") == 0
            @test sq(s8) == "A1 A3 A5"

            # cells absent from sheetData are still covered: CF applies to a region,
            # so the veccolon path must not filter on cell existence
            xf9 = XLSX.newxlsx()
            s9 = xf9[1]
            s9["C3"] = 1                       # dimension A1:C3; A1..C2 absent
            @test XLSX.setConditionalFormat(s9, [1, 3], :, :dataBar) == 0
            @test sq(s9) == "A1:C1 A3:C3"
            SAVE_FILES && save_outfile(xf9)
            SAVE_FILES && save_outfile(xf8)
            SAVE_FILES && save_outfile(xf7)
            SAVE_FILES && save_outfile(xf6)
            SAVE_FILES && save_outfile(xf5)
            SAVE_FILES && save_outfile(xf4)
            SAVE_FILES && save_outfile(xf3)
            SAVE_FILES && save_outfile(xf2)
            SAVE_FILES && save_outfile(xf1)
            SAVE_FILES && save_outfile(xf)

        end

        @testset "vector dispatch across all CF types" begin
            for (t, kw) in ((:dataBar, ()), (:cellIs, (operator="greaterThan", value="1")),
                (:colorScale, ()), (:iconSet, ()), (:top10, ()),
                (:aboveAverage, ()), (:expression, (formula="A1>1",)),
                (:containsErrors, ()), (:containsText, (value="a",)),
                (:timePeriod, ()))
                xf = XLSX.newxlsx()
                s = xf[1]
                for i in 1:4, j in 1:4; s[i, j] = i*j; end
                @test XLSX.setConditionalFormat(s, [1, 3], :, t; kw...) == 0
                @test XLSX.setConditionalFormat(s, :, [1, 3], t; kw...) == 0
                @test XLSX.setConditionalFormat(s, [1, 3], [1, 3], t; kw...) == 0
                SAVE_FILES && save_outfile(xf)
            end
        end
        @testset "formula anchoring on a non-contiguous range" begin
            xf = XLSX.newxlsx()
            s = xf[1]
            for i in 1:9, j in 1:3
                s[i, j] = "x"
            end

            # __CR__ anchors to the first area's top-left, not the lowest cell
            @test XLSX.setConditionalFormat(s, "C5:C9,A1:A3", :containsText; value="a") == 0

            f = tempname() * ".xlsx"
            XLSX.writexlsx(f, xf; overwrite=true)
            x = ZipArchives.zip_readentry(ZipArchives.ZipReader(read(f)),
                "xl/worksheets/sheet1.xml", String)
            @test occursin("C5", x)
            @test occursin(r"sqref=\"C5:C9 A1:A3\"", x)
            SAVE_FILES && save_outfile(xf)
        end

        @testset "_check_cf_range errors" begin
            xf = XLSX.newxlsx()
            s = xf[1]
            s["A1"] = 1
            s["A3"] = 3
            XLSX.addsheet!(xf, "Other")
            other = xf["Other"]
            other["A1"] = 1

            ncr = XLSX.NonContiguousRange(s, "A1,A3")

            # range belongs to a different sheet
            @test_throws XLSX.XLSXError XLSX.setCfDataBar(other, ncr; allkws=Dict{Symbol,Any}())

            # an area outside the worksheet dimension
            @test_throws XLSX.XLSXError XLSX.setConditionalFormat(s, "A1,Z99", :dataBar)

            # sqref too long: many scattered single cells, uncompressed
            xf2 = XLSX.newxlsx()
            s2 = xf2[1]
            for i in 1:1200; s2[i, 1] = i; end
            cells = [XLSX.CellRef(i, 1) for i in 1:2:1200]
            big = XLSX.NonContiguousRange(s2.name, XLSX.NCArea[c for c in cells])
            @test_throws XLSX.XLSXError XLSX.setCfDataBar(s2, big; allkws=Dict{Symbol,Any}())
            SAVE_FILES && save_outfile(xf2)
            SAVE_FILES && save_outfile(xf)
        end

        @testset "defined names resolve to one rule" begin
            xf = XLSX.newxlsx()
            s = xf[1]
            for i in 1:5; s[i, 1] = i; end
            XLSX.addDefinedName(xf, "ncr", "Sheet1!\$A\$1,Sheet1!\$A\$3")

            @test XLSX.setConditionalFormat(s, "ncr", :dataBar) == 0
            @test XLSX.setConditionalFormat(s, "A1,A3", :dataBar) == 0

            f = tempname() * ".xlsx"
            XLSX.writexlsx(f, xf; overwrite=true)
            x = ZipArchives.zip_readentry(ZipArchives.ZipReader(read(f)),
                "xl/worksheets/sheet1.xml", String)

            # named and literal spellings produce the same sqref, one block each
            @test count("sqref=\"A1 A3\"", x) == 1   # both rules join one block
            @test count("<cfRule ", x) == 2
            SAVE_FILES && save_outfile(xf)
        end

        @testset "single-cell sqref round-trips" begin
            xf = XLSX.newxlsx()
            s = xf[1]
            s["A1"] = 1
            @test XLSX.setConditionalFormat(s, "A1", :dataBar) == 0

            f = tempname() * ".xlsx"
            XLSX.writexlsx(f, xf; overwrite=true)
            xf2 = XLSX.opentemplate(f)
            rng = first(first(XLSX.getConditionalFormats(xf2[1])))
            @test XLSX._cf_sqref(rng) == "A1:A1"
            SAVE_FILES && save_outfile(xf2)
            SAVE_FILES && save_outfile(xf)
        end

        @testset "setColoredDataBars CellRef form" begin
            xf = XLSX.newxlsx()
            s = xf[1]
            s["A1"] = 5
            p = XLSX.setColoredDataBars(s, XLSX.CellRef("A1"); bands=1)
            @test XLSX._band_colors(["green", "red"], 1) == ["FF008000"]
            @test length(p) == 1
            SAVE_FILES && save_outfile(xf)
        end
        @testset "invalid keyword arguments" begin
            xf = XLSX.newxlsx()
            s = xf[1]
            for i in 1:5, j in 1:3
                s[i, j] = i * j
            end

            for (t, kw) in ((:cellIs, (operator="greaterThan", value="1")),
                (:containsText, (value="a",)),
                (:top10, ()),
                (:aboveAverage, ()),
                (:timePeriod, ()),
                (:containsErrors, ()),
                (:expression, (formula="A1>1",)),
                (:colorScale, ()),
                (:iconSet, ()),
                (:dataBar, ()))
                @test_throws XLSX.XLSXError XLSX.setConditionalFormat(s, "A1:C5", t; kw..., notAKeyword="x")
            end
            SAVE_FILES && save_outfile(xf)
        end

        @testset "colorScale numeric bounds and mid_col default" begin
            xf = XLSX.newxlsx()
            s = xf[1]
            for i in 1:10, j in 1:3
                s[i, j] = i * j
            end

            # explicit colours
            @test XLSX.setConditionalFormat(s, "A1:C10", :colorScale;
                min_type="num", min_val="2", min_col="tomato",
                mid_type="num", mid_val="6", mid_col="lawngreen",
                max_type="num", max_val="10", max_col="cadetblue") == 0

            # mid_col omitted: defaults to white rather than throwing
            @test XLSX.setConditionalFormat(s, "A1:C10", :colorScale;
                min_type="num", min_val="2",
                mid_type="num", mid_val="6",
                max_type="num", max_val="10") == 0

            f = tempname() * ".xlsx"
            XLSX.writexlsx(f, xf; overwrite=true)
            x = ZipArchives.zip_readentry(ZipArchives.ZipReader(read(f)),
                "xl/worksheets/sheet1.xml", String)

            # "num" types mean the values survive the `== "min"` / `== "max"` guards
            @test occursin("<cfvo type=\"num\" val=\"2\"/>", x)
            @test occursin("<cfvo type=\"num\" val=\"6\"/>", x)
            @test occursin("<cfvo type=\"num\" val=\"10\"/>", x)
            @test occursin("rgb=\"FFFCFCFF\"", x)
            SAVE_FILES && save_outfile(xf)
        end

        @testset "cellIs default value averages the cells of a non-contiguous range" begin
            xf = XLSX.newxlsx()
            s = xf[1]
            s["A1"] = 2
            s["A2"] = 100
            s["A3"] = 4
            s["A4"] = "text"
            XLSX.setConditionalFormat(s, "A1,A3", :cellIs)            # the average of A1 and A3 only
            XLSX.setConditionalFormat(s, "A1:A4", :cellIs)            # contiguous, with a text cell
            f = tempname() * ".xlsx"
            XLSX.writexlsx(f, xf; overwrite=true)
            r = XLSX.ZipArchives.ZipReader(read(f))
            xml = XLSX.ZipArchives.zip_readentry(r, "xl/worksheets/sheet1.xml", String)
            fmls = [m[1] for m in eachmatch(r"<formula>([^<]*)</formula>", xml)]
            @test "3.0" in fmls                                       # (2 + 4) / 2
            @test any(x -> parse(Float64, x) ≈ 106 / 3, fmls)         # (2 + 100 + 4) / 3; text ignored
            @test !any(x -> occursin(r"\[|;;", x), fmls)              # never a printed array
            rm(f)
            SAVE_FILES && save_outfile(xf)
        end

    end
    @testset "clear replaces previous bands" begin
        xf = XLSX.newxlsx()
        s = xf[1]
        for i in 1:10; s[i, 1] = i; end

        XLSX.setColoredDataBars(s, "A1:A10"; bands=1, colors="steelblue")
        XLSX.setColoredDataBars(s, "A1:A10"; bands=3)

        f = tempname() * ".xlsx"
        XLSX.writexlsx(f, xf; overwrite=true)
        x = ZipArchives.zip_readentry(ZipArchives.ZipReader(read(f)),
            "xl/worksheets/sheet1.xml", String)

        @test count("<conditionalFormatting ", x) == 3
        @test count("<x14:cfRule ", x) == 3
        @test count("<x14:conditionalFormattings", x) == 1   # no empty stray block

        # re-banding after the data changes drops empty bands and leaves no orphans
        xf2 = XLSX.newxlsx()
        s2 = xf2[1]
        for i in 1:20; s2[i, 1] = i; end
        XLSX.setColoredDataBars(s2, "A1:A20"; bands=5)
        for i in 1:20; s2[i, 1] = 1 + (i % 2); end
        p = XLSX.setColoredDataBars(s2, "A1:A20"; bands=5)

        @test length(p) == 2
        @test first.(p) == [1, 5]                     # labels keep their place on the scale
        @test length(XLSX.getConditionalFormats(s2)) == length(p)

        # clear=false layers instead
        XLSX.setColoredDataBars(s2, "A1:A20"; bands=5, clear=false)
        @test length(XLSX.getConditionalFormats(s2)) == 2 * length(p)
        SAVE_FILES && save_outfile(xf2)
        SAVE_FILES && save_outfile(xf)
    end
    @testset "coverage: clearConditionalFormats with a string range" begin
        xf = XLSX.newxlsx()
        ws = xf[1]
        ws["A1:A5"] = reshape(collect(1:5), 5, 1)
        XLSX.setConditionalFormat(ws, "A1:A5", :colorScale)
        @test !isempty(XLSX.getConditionalFormats(ws))
        XLSX.clearConditionalFormats(ws, "A1:A5")
        @test isempty(XLSX.getConditionalFormats(ws))
        SAVE_FILES && save_outfile(xf)
    end
end
