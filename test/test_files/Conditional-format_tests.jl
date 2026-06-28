@testset "Conditional Formats" begin

    @testset "DataBar" begin

        f = XLSX.newxlsx()
        s = f[1]
        for i in 1:5, j in 1:5
            s[i, j] = i + j
        end

        @test_throws MethodError XLSX.setConditionalFormat(s, "A1,A3", :dataBar) # Non-contiguous ranges not allowed
        @test_throws MethodError XLSX.setConditionalFormat(s, [1], 1, :dataBar) # Vectors may be non-contiguous
        @test_throws MethodError XLSX.setConditionalFormat(s, 1, 1:3:7, :dataBar) # StepRange is non-contiguous
        @test XLSX.setConditionalFormat(s, "1:1", :dataBar) == 0
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
        @test sort!(XLSX.getConditionalFormats(s), by = x -> x.second.priority, rev=true) == [
            XLSX.CellRange("A5:E5") => (type="dataBar", priority=5), 
            XLSX.CellRange("A4:E4") => (type="dataBar", priority=4), 
            XLSX.CellRange("A3:E3") => (type="dataBar", priority=3), 
            XLSX.CellRange("A2:E2") => (type="dataBar", priority=2), 
            XLSX.CellRange("A1:E1") => (type="dataBar", priority=1),
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
        @test length(XLSX.getConditionalFormats(s)) == 22
        @test sort!(XLSX.getConditionalFormats(s), by = x -> x.second.priority, rev=true) == [
            XLSX.CellRange("A1:E5") => (type="dataBar", priority=22),
            XLSX.CellRange("A1:E5") => (type="dataBar", priority=21),
            XLSX.CellRange("A1:C5") => (type="dataBar", priority=20),
            XLSX.CellRange("A1:C5") => (type="dataBar", priority=19),
            XLSX.CellRange("A2:E4") => (type="dataBar", priority=18),
            XLSX.CellRange("A1:E3") => (type="dataBar", priority=17),
            XLSX.CellRange("A1:C5") => (type="dataBar", priority=16),
            XLSX.CellRange("A1:C5") => (type="dataBar", priority=15),
            XLSX.CellRange("A1:E2") => (type="dataBar", priority=14),
            XLSX.CellRange("A1:C5") => (type="dataBar", priority=13),
            XLSX.CellRange("A1:C5") => (type="dataBar", priority=12),
            XLSX.CellRange("A2:E4") => (type="dataBar", priority=11),
            XLSX.CellRange("A1:E2") => (type="dataBar", priority=10),
            XLSX.CellRange("A1:A2") => (type="dataBar", priority=9),
            XLSX.CellRange("A1:A1") => (type="dataBar", priority=8),
            XLSX.CellRange("A1:C3") => (type="dataBar", priority=7),
            XLSX.CellRange("A1:A1") => (type="dataBar", priority=6),
            XLSX.CellRange("A5:E5") => (type="dataBar", priority=5),
            XLSX.CellRange("A4:E4") => (type="dataBar", priority=4),
            XLSX.CellRange("A3:E3") => (type="dataBar", priority=3),
            XLSX.CellRange("A2:E2") => (type="dataBar", priority=2),
            XLSX.CellRange("A1:E1") => (type="dataBar", priority=1)
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
        @test sort!(XLSX.getConditionalFormats(s), by = x -> x.second.priority, rev=true) == [
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
        @test_throws MethodError XLSX.setConditionalFormat(s, "myNCRange", :dataBar; # Non-contiguous ranges not allowed
            showVal="false",
            direction="leftToRight",
            borders="true",
            sameNegBorders="false"
        )
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

        @test_throws XLSX.XLSXError XLSX.setConditionalFormat(s, "A1,A3", :wrongOne)
        @test_throws XLSX.XLSXError XLSX.setConditionalFormat(s, 1, 2, :wrongOne)

        @test_throws MethodError XLSX.setConditionalFormat(s, "A1,A3", :colorScale) # Non-contiguous ranges not allowed
        @test_throws MethodError XLSX.setConditionalFormat(s, [1], 1, :colorScale) # Vectors may be non-contiguous
        @test_throws MethodError XLSX.setConditionalFormat(s, 1, 1:3:7, :colorScale) # StepRange is non-contiguous
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
        @test sort!(XLSX.getConditionalFormats(s), by = x -> x.second.priority, rev=true) == [
            XLSX.CellRange("A5:E5") => (type="colorScale", priority=5), 
            XLSX.CellRange("A4:E4") => (type="colorScale", priority=4), 
            XLSX.CellRange("A3:E3") => (type="colorScale", priority=3), 
            XLSX.CellRange("A2:E2") => (type="colorScale", priority=2), 
            XLSX.CellRange("A1:E1") => (type="colorScale", priority=1)
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
        @test length(XLSX.getConditionalFormats(s)) == 22
        @test sort!(XLSX.getConditionalFormats(s), by = x -> x.second.priority, rev=true) == [
            XLSX.CellRange("A1:E5") => (type="colorScale", priority=22),
            XLSX.CellRange("A1:E5") => (type="colorScale", priority=21),
            XLSX.CellRange("A1:C5") => (type="colorScale", priority=20),
            XLSX.CellRange("A1:C5") => (type="colorScale", priority=19),
            XLSX.CellRange("A2:E4") => (type="colorScale", priority=18),
            XLSX.CellRange("A1:E3") => (type="colorScale", priority=17),
            XLSX.CellRange("A1:C5") => (type="colorScale", priority=16),
            XLSX.CellRange("A1:C5") => (type="colorScale", priority=15),
            XLSX.CellRange("A1:E2") => (type="colorScale", priority=14),
            XLSX.CellRange("A1:C5") => (type="colorScale", priority=13),
            XLSX.CellRange("A1:C5") => (type="colorScale", priority=12),
            XLSX.CellRange("A2:E4") => (type="colorScale", priority=11),
            XLSX.CellRange("A1:E2") => (type="colorScale", priority=10),
            XLSX.CellRange("A1:A2") => (type="colorScale", priority=9),
            XLSX.CellRange("A1:A1") => (type="colorScale", priority=8),
            XLSX.CellRange("A1:C3") => (type="colorScale", priority=7),
            XLSX.CellRange("A1:A1") => (type="colorScale", priority=6),
            XLSX.CellRange("A5:E5") => (type="colorScale", priority=5),
            XLSX.CellRange("A4:E4") => (type="colorScale", priority=4),
            XLSX.CellRange("A3:E3") => (type="colorScale", priority=3),
            XLSX.CellRange("A2:E2") => (type="colorScale", priority=2),
            XLSX.CellRange("A1:E1") => (type="colorScale", priority=1)
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
        @test sort!(XLSX.getConditionalFormats(s), by = x -> x.second.priority, rev=true) == [
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
        @test_throws MethodError XLSX.setConditionalFormat(s, "myNCRange", :colorScale; # Non-contiguous ranges not allowed
            min_type="min",
            min_col="green",
            mid_type="percentile",
            mid_val="50",
            mid_col="red",
            max_type="max",
            max_col="blue"
        )
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
        @test_throws MethodError XLSX.setConditionalFormat(s, "A1,A3", :iconSet) # Non-contiguous ranges not allowed
        @test_throws MethodError XLSX.setConditionalFormat(s, [1], 1, :iconSet) # Vectors may be non-contiguous
        @test_throws MethodError XLSX.setConditionalFormat(s, 1, 1:3:7, :iconSet) # StepRange is non-contiguous
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
        @test sort!(XLSX.getConditionalFormats(s), by = x -> x.second.priority, rev=true) == [
            XLSX.CellRange("A5:E5") => (type="iconSet", priority=5), 
            XLSX.CellRange("A4:E4") => (type="iconSet", priority=4), 
            XLSX.CellRange("A3:E3") => (type="iconSet", priority=3), 
            XLSX.CellRange("A2:E2") => (type="iconSet", priority=2), 
            XLSX.CellRange("A1:E1") => (type="iconSet", priority=1)
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
        @test length(XLSX.getConditionalFormats(s)) == 22

        @test sort!(XLSX.getConditionalFormats(s), by = x -> x.second.priority, rev=true) == [
            XLSX.CellRange("A1:E5") => (type="iconSet", priority=22),
            XLSX.CellRange("A1:E5") => (type="iconSet", priority=21),
            XLSX.CellRange("A1:C5") => (type="iconSet", priority=20),
            XLSX.CellRange("A1:C5") => (type="iconSet", priority=19),
            XLSX.CellRange("A2:E4") => (type="iconSet", priority=18),
            XLSX.CellRange("A1:E3") => (type="iconSet", priority=17),
            XLSX.CellRange("A1:C5") => (type="iconSet", priority=16),
            XLSX.CellRange("A1:C5") => (type="iconSet", priority=15),
            XLSX.CellRange("A1:E2") => (type="iconSet", priority=14),
            XLSX.CellRange("A1:C5") => (type="iconSet", priority=13),
            XLSX.CellRange("A1:C5") => (type="iconSet", priority=12),
            XLSX.CellRange("A2:E4") => (type="iconSet", priority=11),
            XLSX.CellRange("A1:E2") => (type="iconSet", priority=10),
            XLSX.CellRange("A1:A2") => (type="iconSet", priority=9),
            XLSX.CellRange("A1:A1") => (type="iconSet", priority=8),
            XLSX.CellRange("A1:C3") => (type="iconSet", priority=7),
            XLSX.CellRange("A1:A1") => (type="iconSet", priority=6),
            XLSX.CellRange("A5:E5") => (type="iconSet", priority=5),
            XLSX.CellRange("A4:E4") => (type="iconSet", priority=4),
            XLSX.CellRange("A3:E3") => (type="iconSet", priority=3),
            XLSX.CellRange("A2:E2") => (type="iconSet", priority=2),
            XLSX.CellRange("A1:E1") => (type="iconSet", priority=1)
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

        @test sort!(XLSX.getConditionalFormats(s), by = x -> x.second.priority, rev=true) == [
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


        @test sort!(XLSX.getConditionalFormats(s), by = x -> x.second.priority) == [
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
        @test_throws MethodError XLSX.setConditionalFormat(s, "myNCRange", :iconSet)
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
        @test_throws MethodError XLSX.setConditionalFormat(s, "A1,A3", :cellIs) # Non-contiguous ranges not allowed
        @test_throws MethodError XLSX.setConditionalFormat(s, [1], 1, :cellIs) # Vectors may be non-contiguous
        @test_throws MethodError XLSX.setConditionalFormat(s, 1, 1:3:7, :cellIs) # StepRange is non-contiguous
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
        @test sort!(XLSX.getConditionalFormats(s), by = x -> x.second.priority, rev=true) == [
            XLSX.CellRange("A5:E5") => (type="cellIs", priority=5), 
            XLSX.CellRange("A4:E4") => (type="cellIs", priority=4), 
            XLSX.CellRange("A3:E3") => (type="cellIs", priority=3), 
            XLSX.CellRange("A2:E2") => (type="cellIs", priority=2), 
            XLSX.CellRange("A1:E1") => (type="cellIs", priority=1)
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
        @test length(XLSX.getConditionalFormats(s)) == 22
        @test sort!(XLSX.getConditionalFormats(s), by = x -> x.second.priority, rev=true) == [
            XLSX.CellRange("A1:E5") => (type="cellIs", priority=22),
            XLSX.CellRange("A1:E5") => (type="cellIs", priority=21),
            XLSX.CellRange("A1:C5") => (type="cellIs", priority=20),
            XLSX.CellRange("A1:C5") => (type="cellIs", priority=19),
            XLSX.CellRange("A2:E4") => (type="cellIs", priority=18),
            XLSX.CellRange("A1:E3") => (type="cellIs", priority=17),
            XLSX.CellRange("A1:C5") => (type="cellIs", priority=16),
            XLSX.CellRange("A1:C5") => (type="cellIs", priority=15),
            XLSX.CellRange("A1:E2") => (type="cellIs", priority=14),
            XLSX.CellRange("A1:C5") => (type="cellIs", priority=13),
            XLSX.CellRange("A1:C5") => (type="cellIs", priority=12),
            XLSX.CellRange("A2:E4") => (type="cellIs", priority=11),
            XLSX.CellRange("A1:E2") => (type="cellIs", priority=10),
            XLSX.CellRange("A1:A2") => (type="cellIs", priority=9),
            XLSX.CellRange("A1:A1") => (type="cellIs", priority=8),
            XLSX.CellRange("A1:C3") => (type="cellIs", priority=7),
            XLSX.CellRange("A1:A1") => (type="cellIs", priority=6),
            XLSX.CellRange("A5:E5") => (type="cellIs", priority=5),
            XLSX.CellRange("A4:E4") => (type="cellIs", priority=4),
            XLSX.CellRange("A3:E3") => (type="cellIs", priority=3),
            XLSX.CellRange("A2:E2") => (type="cellIs", priority=2),
            XLSX.CellRange("A1:E1") => (type="cellIs", priority=1)
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

        @test sort!(XLSX.getConditionalFormats(s), by = x -> x.second.priority, rev=true) == [
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
        @test_throws MethodError XLSX.setConditionalFormat(s, "myNCRange", :cellIs; # Non-contiguous ranges not allowed
            operator="lessThan",
            value="2",
            fill=["pattern" => "none", "bgColor" => "yellow"],
            format=["format" => "0.0"],
            font=["color" => "green"],
            border=["style" => "hair", "color" => "cyan"]
        )
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
        @test_throws MethodError XLSX.setConditionalFormat(s, "A1,A3", :containsText; value="a") # Non-contiguous ranges not allowed
        @test_throws MethodError XLSX.setConditionalFormat(s, [1], 1, :containsText; value="a") # Vectors may be non-contiguous
        @test_throws MethodError XLSX.setConditionalFormat(s, 1, 1:3:7, :containsText; value="a") # StepRange is non-contiguous
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
        @test sort!(XLSX.getConditionalFormats(s), by = x -> x.second.priority, rev=true) == [
            XLSX.CellRange("A5:E5") => (type="beginsWith", priority=5), 
            XLSX.CellRange("A4:E4") => (type="notContainsText", priority=4), 
            XLSX.CellRange("A3:E3") => (type="notContainsText", priority=3), 
            XLSX.CellRange("A2:E2") => (type="containsText", priority=2), 
            XLSX.CellRange("A1:E1") => (type="containsText", priority=1)
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
        @test length(XLSX.getConditionalFormats(s)) == 22
        @test sort!(XLSX.getConditionalFormats(s), by = x -> x.second.priority, rev=true) == [
            XLSX.CellRange("A1:E5") => (type="containsText", priority=22),
            XLSX.CellRange("A1:E5") => (type="containsText", priority=21),
            XLSX.CellRange("A1:C5") => (type="containsText", priority=20),
            XLSX.CellRange("A1:C5") => (type="containsText", priority=19),
            XLSX.CellRange("A2:E4") => (type="containsText", priority=18),
            XLSX.CellRange("A1:E3") => (type="containsText", priority=17),
            XLSX.CellRange("A1:C5") => (type="containsText", priority=16),
            XLSX.CellRange("A1:C5") => (type="containsText", priority=15),
            XLSX.CellRange("A1:E2") => (type="containsText", priority=14),
            XLSX.CellRange("A1:C5") => (type="containsText", priority=13),
            XLSX.CellRange("A1:C5") => (type="containsText", priority=12),
            XLSX.CellRange("A2:E4") => (type="containsText", priority=11),
            XLSX.CellRange("A1:E2") => (type="containsText", priority=10),
            XLSX.CellRange("A1:A2") => (type="containsText", priority=9),
            XLSX.CellRange("A1:A1") => (type="containsText", priority=8),
            XLSX.CellRange("A1:C3") => (type="containsText", priority=7),
            XLSX.CellRange("A1:A1") => (type="containsText", priority=6),
            XLSX.CellRange("A5:E5") => (type="beginsWith", priority=5),
            XLSX.CellRange("A4:E4") => (type="notContainsText", priority=4),
            XLSX.CellRange("A3:E3") => (type="notContainsText", priority=3),
            XLSX.CellRange("A2:E2") => (type="containsText", priority=2),
            XLSX.CellRange("A1:E1") => (type="containsText", priority=1)
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

        @test sort!(XLSX.getConditionalFormats(s), by = x -> x.second.priority, rev=true) == [
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
        @test_throws MethodError XLSX.setConditionalFormat(s, "myNCRange", :containsText; # Non-contiguous ranges not allowed
            operator="beginsWith",
            value="a",
            fill=["pattern" => "none", "bgColor" => "yellow"],
            format=["format" => "0.0"],
            font=["color" => "green"],
            border=["style" => "hair", "color" => "cyan"]
        )
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
        @test_throws MethodError XLSX.setConditionalFormat(s, "A1,A3", :top10) # Non-contiguous ranges not allowed
        @test_throws MethodError XLSX.setConditionalFormat(s, [1], 1, :top10) # Vectors may be non-contiguous
        @test_throws MethodError XLSX.setConditionalFormat(s, 1, 1:3:7, :top10) # StepRange is non-contiguous
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
        @test sort!(XLSX.getConditionalFormats(s), by = x -> x.second.priority) == [
            XLSX.CellRange("A1:J1") => (type="top10", priority=1),
            XLSX.CellRange("A2:J2") => (type="top10", priority=2),
            XLSX.CellRange("A1:J10") => (type="top10", priority=3), 
            XLSX.CellRange("A1:J10") => (type="top10", priority=4), 
            XLSX.CellRange("A1:J10") => (type="top10", priority=5), 
            XLSX.CellRange("A1:J10") => (type="top10", priority=6), 
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
        @test length(XLSX.getConditionalFormats(s)) == 23
        @test sort!(XLSX.getConditionalFormats(s), by = x -> x.second.priority) == [
            XLSX.CellRange("A1:J1") => (type="top10", priority=1),
            XLSX.CellRange("A2:J2") => (type="top10", priority=2),
            XLSX.CellRange("A1:J10") => (type="top10", priority=3),
            XLSX.CellRange("A1:J10") => (type="top10", priority=4),
            XLSX.CellRange("A1:J10") => (type="top10", priority=5),
            XLSX.CellRange("A1:J10") => (type="top10", priority=6),
            XLSX.CellRange("A1:A1") => (type="top10", priority=7),
            XLSX.CellRange("A1:C3") => (type="top10", priority=8),
            XLSX.CellRange("A1:A1") => (type="top10", priority=9),
            XLSX.CellRange("A1:A2") => (type="top10", priority=10),
            XLSX.CellRange("A1:J2") => (type="top10", priority=11),
            XLSX.CellRange("A2:J4") => (type="top10", priority=12),
            XLSX.CellRange("A1:C10") => (type="top10", priority=13),
            XLSX.CellRange("A1:C10") => (type="top10", priority=14),
            XLSX.CellRange("A1:J2") => (type="top10", priority=15),
            XLSX.CellRange("A1:C10") => (type="top10", priority=16),
            XLSX.CellRange("A1:C10") => (type="top10", priority=17),
            XLSX.CellRange("A1:J3") => (type="top10", priority=18),
            XLSX.CellRange("A2:J4") => (type="top10", priority=19),
            XLSX.CellRange("A1:C10") => (type="top10", priority=20),
            XLSX.CellRange("A1:C10") => (type="top10", priority=21),
            XLSX.CellRange("A1:J10") => (type="top10", priority=22),
            XLSX.CellRange("A1:J10") => (type="top10", priority=23),
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

        @test sort!(XLSX.getConditionalFormats(s), by = x -> x.second.priority, rev=true) == [
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
        @test_throws MethodError XLSX.setConditionalFormat(s, "myNCRange", :top10; # Non-contiguous ranges not allowed
            operator="bottomN%",
            value="2",
            fill=["pattern" => "none", "bgColor" => "yellow"],
            format=["format" => "0.0"],
            font=["color" => "green"],
            border=["style" => "hair", "color" => "cyan"]
        )
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
        @test_throws MethodError XLSX.setConditionalFormat(s, "A2:A1001,C1:C1000", :aboveAverage) # Non-contiguous ranges not allowed
        @test_throws MethodError XLSX.setConditionalFormat(s, [2, 3, 19], 1:3, :aboveAverage) # Vectors may be non-contiguous
        @test_throws MethodError XLSX.setConditionalFormat(s, 2, 1:3:7, :aboveAverage) # StepRange is non-contiguous
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
        @test sort!(XLSX.getConditionalFormats(s), by = x -> x.second.priority) == [
            XLSX.CellRange("A2:C2") => (type="aboveAverage", priority=1),
            XLSX.CellRange("A2:C2") => (type="aboveAverage", priority=2),
            XLSX.CellRange("A2:C10") => (type="aboveAverage", priority=3),
            XLSX.CellRange("A2:C1001") => (type="aboveAverage", priority=4),
            XLSX.CellRange("A2:C1001") => (type="aboveAverage", priority=5),
            XLSX.CellRange("A2:C1001") => (type="aboveAverage", priority=6),
            XLSX.CellRange("A2:C1001") => (type="aboveAverage", priority=7),
            XLSX.CellRange("A1:C1001") => (type="aboveAverage", priority=8),
            XLSX.CellRange("A1:C1001") => (type="aboveAverage", priority=9),
            XLSX.CellRange("A1:C1001") => (type="aboveAverage", priority=10)
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
        @test length(XLSX.getConditionalFormats(s)) == 27
        @test sort!(XLSX.getConditionalFormats(s), by = x -> x.second.priority) == [
            XLSX.CellRange("A2:C2") => (type="aboveAverage", priority=1),
            XLSX.CellRange("A2:C2") => (type="aboveAverage", priority=2),
            XLSX.CellRange("A2:C10") => (type="aboveAverage", priority=3),
            XLSX.CellRange("A2:C1001") => (type="aboveAverage", priority=4),
            XLSX.CellRange("A2:C1001") => (type="aboveAverage", priority=5),
            XLSX.CellRange("A2:C1001") => (type="aboveAverage", priority=6),
            XLSX.CellRange("A2:C1001") => (type="aboveAverage", priority=7),
            XLSX.CellRange("A1:C1001") => (type="aboveAverage", priority=8),
            XLSX.CellRange("A1:C1001") => (type="aboveAverage", priority=9),
            XLSX.CellRange("A1:C1001") => (type="aboveAverage", priority=10),
            XLSX.CellRange("A1:A1") => (type="aboveAverage", priority=11),
            XLSX.CellRange("A1:C3") => (type="aboveAverage", priority=12),
            XLSX.CellRange("A1:A1") => (type="aboveAverage", priority=13),
            XLSX.CellRange("A1:A2") => (type="aboveAverage", priority=14),
            XLSX.CellRange("A1:C2") => (type="aboveAverage", priority=15),
            XLSX.CellRange("A2:C4") => (type="aboveAverage", priority=16),
            XLSX.CellRange("A1:C1001") => (type="aboveAverage", priority=17),
            XLSX.CellRange("A1:C1001") => (type="aboveAverage", priority=18),
            XLSX.CellRange("A1:C2") => (type="aboveAverage", priority=19),
            XLSX.CellRange("A1:C1001") => (type="aboveAverage", priority=20),
            XLSX.CellRange("A1:C1001") => (type="aboveAverage", priority=21),
            XLSX.CellRange("A1:C3") => (type="aboveAverage", priority=22),
            XLSX.CellRange("A2:C4") => (type="aboveAverage", priority=23),
            XLSX.CellRange("A1:C1001") => (type="aboveAverage", priority=24),
            XLSX.CellRange("A1:C1001") => (type="aboveAverage", priority=25),
            XLSX.CellRange("A1:C1001") => (type="aboveAverage", priority=26),
            XLSX.CellRange("A1:C1001") => (type="aboveAverage", priority=27),
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

        @test sort!(XLSX.getConditionalFormats(s), by = x -> x.second.priority, rev=true) == [
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
        @test_throws MethodError XLSX.setConditionalFormat(s, "myNCRange", :aboveAverage; # Non-contiguous ranges not allowed
            operator="belowEqAverage",
            fill=["pattern" => "none", "bgColor" => "yellow"],
            format=["format" => "0.0"],
            font=["color" => "green"],
            border=["style" => "hair", "color" => "cyan"]
        )
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

        @test_throws MethodError XLSX.setConditionalFormat(s, "A1:A5,C1:C5", :timePeriod) # Non-contiguous ranges not allowed
        @test_throws MethodError XLSX.setConditionalFormat(s, [2, 3, 8], 1:3, :timePeriod) # Vectors may be non-contiguous
        @test_throws MethodError XLSX.setConditionalFormat(s, 2, 1:3:7, :timePeriod) # StepRange is non-contiguous
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
        @test sort!(XLSX.getConditionalFormats(s), by = x -> x.second.priority) == [
            XLSX.CellRange("A2:J2") => (type="timePeriod", priority=1),
            XLSX.CellRange("A2:J2") => (type="timePeriod", priority=2),
            XLSX.CellRange("A1:C10") => (type="timePeriod", priority=3),
            XLSX.CellRange("A1:C10") => (type="timePeriod", priority=4),
            XLSX.CellRange("A1:C10") => (type="timePeriod", priority=5),
            XLSX.CellRange("A1:C10") => (type="timePeriod", priority=6),
            XLSX.CellRange("A1:C10") => (type="timePeriod", priority=7),
            XLSX.CellRange("A1:C10") => (type="timePeriod", priority=8),
            XLSX.CellRange("A1:C10") => (type="timePeriod", priority=9)
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
        @test length(XLSX.getConditionalFormats(s)) == 26

        @test sort!(XLSX.getConditionalFormats(s), by = x -> x.second.priority) == [
            XLSX.CellRange("A2:J2") => (type="timePeriod", priority=1),
            XLSX.CellRange("A2:J2") => (type="timePeriod", priority=2),
            XLSX.CellRange("A1:C10") => (type="timePeriod", priority=3),
            XLSX.CellRange("A1:C10") => (type="timePeriod", priority=4),
            XLSX.CellRange("A1:C10") => (type="timePeriod", priority=5),
            XLSX.CellRange("A1:C10") => (type="timePeriod", priority=6),
            XLSX.CellRange("A1:C10") => (type="timePeriod", priority=7),
            XLSX.CellRange("A1:C10") => (type="timePeriod", priority=8),
            XLSX.CellRange("A1:C10") => (type="timePeriod", priority=9),
            XLSX.CellRange("A1:A1") => (type="timePeriod", priority=10),
            XLSX.CellRange("A1:C3") => (type="timePeriod", priority=11),
            XLSX.CellRange("A1:A1") => (type="timePeriod", priority=12),
            XLSX.CellRange("A1:A2") => (type="timePeriod", priority=13),
            XLSX.CellRange("A1:J2") => (type="timePeriod", priority=14),
            XLSX.CellRange("A2:J4") => (type="timePeriod", priority=15),
            XLSX.CellRange("A1:C10") => (type="timePeriod", priority=16),
            XLSX.CellRange("A1:C10") => (type="timePeriod", priority=17),
            XLSX.CellRange("A1:J2") => (type="timePeriod", priority=18),
            XLSX.CellRange("A1:C10") => (type="timePeriod", priority=19),
            XLSX.CellRange("A1:C10") => (type="timePeriod", priority=20),
            XLSX.CellRange("A1:J3") => (type="timePeriod", priority=21),
            XLSX.CellRange("A2:J4") => (type="timePeriod", priority=22),
            XLSX.CellRange("A1:C10") => (type="timePeriod", priority=23),
            XLSX.CellRange("A1:C10") => (type="timePeriod", priority=24),
            XLSX.CellRange("A1:J10") => (type="timePeriod", priority=25),
            XLSX.CellRange("A1:J10") => (type="timePeriod", priority=26)
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

        @test sort!(XLSX.getConditionalFormats(s), by = x -> x.second.priority, rev=true) == [
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
        @test_throws MethodError XLSX.setConditionalFormat(s, "myNCRange", :timePeriod; # Non-contiguous ranges not allowed
            operator="lastWeek",
            fill=["pattern" => "none", "bgColor" => "yellow"],
            format=["format" => "0.0"],
            font=["color" => "green"],
            border=["style" => "hair", "color" => "cyan"]
        )
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

        @test_throws MethodError XLSX.setConditionalFormat(s, "A1:A5,C1:C5", :expression; formula="A1>3") # Non-contiguous ranges not allowed
        @test_throws MethodError XLSX.setConditionalFormat(s, [2, 3, 8], 1:3, :expression; formula="A1 > 11") # Vectors may be non-contiguous
        @test_throws MethodError XLSX.setConditionalFormat(s, 2, 1:3:7, :expression; formula="A1 < 7") # StepRange is non-contiguous
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
        @test sort!(XLSX.getConditionalFormats(s), by = x -> x.second.priority) == [
            XLSX.CellRange("A2:J2") => (type="expression", priority=1),
            XLSX.CellRange("A2:J2") => (type="expression", priority=2),
            XLSX.CellRange("A1:C10") => (type="expression", priority=3),
            XLSX.CellRange("A1:C10") => (type="expression", priority=4),
            XLSX.CellRange("A1:C10") => (type="expression", priority=5),
            XLSX.CellRange("A1:C10") => (type="expression", priority=6),
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
        @test length(XLSX.getConditionalFormats(s)) == 23
        @test sort!(XLSX.getConditionalFormats(s), by = x -> x.second.priority, rev=true) == [
            XLSX.CellRange("A1:J10") => (type="expression", priority=23),
            XLSX.CellRange("A1:J10") => (type="expression", priority=22),
            XLSX.CellRange("A1:C10") => (type="expression", priority=21),
            XLSX.CellRange("A1:C10") => (type="expression", priority=20),
            XLSX.CellRange("A2:J4") => (type="expression", priority=19),
            XLSX.CellRange("A1:J3") => (type="expression", priority=18),
            XLSX.CellRange("A1:C10") => (type="expression", priority=17),
            XLSX.CellRange("A1:C10") => (type="expression", priority=16),
            XLSX.CellRange("A1:J2") => (type="expression", priority=15),
            XLSX.CellRange("A1:C10") => (type="expression", priority=14),
            XLSX.CellRange("A1:C10") => (type="expression", priority=13),
            XLSX.CellRange("A2:J4") => (type="expression", priority=12),
            XLSX.CellRange("A1:J2") => (type="expression", priority=11),
            XLSX.CellRange("A1:A2") => (type="expression", priority=10),
            XLSX.CellRange("A1:A1") => (type="expression", priority=9),
            XLSX.CellRange("A1:C3") => (type="expression", priority=8),
            XLSX.CellRange("A1:A1") => (type="expression", priority=7),
            XLSX.CellRange("A1:C10") => (type="expression", priority=6),
            XLSX.CellRange("A1:C10") => (type="expression", priority=5),
            XLSX.CellRange("A1:C10") => (type="expression", priority=4),
            XLSX.CellRange("A1:C10") => (type="expression", priority=3),
            XLSX.CellRange("A2:J2") => (type="expression", priority=2),
            XLSX.CellRange("A2:J2") => (type="expression", priority=1),
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

        @test sort!(XLSX.getConditionalFormats(s), by = x -> x.second.priority, rev=true) == [
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
        @test_throws MethodError XLSX.setConditionalFormat(s, "myNCRange", :expression; # Non-contiguous ranges not allowed
            formula="C4 < myTest",
            fill=["pattern" => "none", "bgColor" => "yellow"],
            format=["format" => "0.0"],
            font=["color" => "green"],
            border=["style" => "hair", "color" => "cyan"]
        )
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
        @test_throws MethodError XLSX.setConditionalFormat(s, "A1:A5,C1:C5", :containsErrors) # Non-contiguous ranges not allowed
        @test_throws MethodError XLSX.setConditionalFormat(s, [2, 3, 8], 1:3, :containsErrors) # Vectors may be non-contiguous
        @test_throws MethodError XLSX.setConditionalFormat(s, 2, 1:3:7, :containsErrors) # StepRange is non-contiguous
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
        @test sort!(XLSX.getConditionalFormats(s), by = x -> x.second.priority) == [
            XLSX.CellRange("A2:J2") => (type="containsErrors", priority=1),
            XLSX.CellRange("A2:J2") => (type="containsErrors", priority=2),
            XLSX.CellRange("A1:C10") => (type="containsErrors", priority=3),
            XLSX.CellRange("A1:C10") => (type="notContainsErrors", priority=4),
            XLSX.CellRange("A1:C10") => (type="containsBlanks", priority=5),
            XLSX.CellRange("A1:C10") => (type="notContainsBlanks", priority=6),
            XLSX.CellRange("A1:C10") => (type="uniqueValues", priority=7),
            XLSX.CellRange("A1:C10") => (type="duplicateValues", priority=8),
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
        @test length(XLSX.getConditionalFormats(s)) == 25
        @test sort!(XLSX.getConditionalFormats(s), by = x -> x.second.priority, rev=true) == [
            XLSX.CellRange("A1:J10") => (type="duplicateValues", priority=25),
            XLSX.CellRange("A1:J10") => (type="uniqueValues", priority=24),
            XLSX.CellRange("A1:C10") => (type="containsErrors", priority=23),
            XLSX.CellRange("A1:C10") => (type="notContainsBlanks", priority=22),
            XLSX.CellRange("A2:J4") => (type="containsBlanks", priority=21),
            XLSX.CellRange("A1:J3") => (type="notContainsErrors", priority=20),
            XLSX.CellRange("A1:C10") => (type="containsErrors", priority=19),
            XLSX.CellRange("A1:C10") => (type="containsErrors", priority=18),
            XLSX.CellRange("A1:J2") => (type="containsErrors", priority=17),
            XLSX.CellRange("A1:C10") => (type="containsErrors", priority=16),
            XLSX.CellRange("A1:C10") => (type="containsErrors", priority=15),
            XLSX.CellRange("A2:J4") => (type="duplicateValues", priority=14),
            XLSX.CellRange("A1:J2") => (type="uniqueValues", priority=13),
            XLSX.CellRange("A1:A2") => (type="notContainsBlanks", priority=12),
            XLSX.CellRange("A1:A1") => (type="containsBlanks", priority=11),
            XLSX.CellRange("A1:C3") => (type="notContainsErrors", priority=10),
            XLSX.CellRange("A1:A1") => (type="containsErrors", priority=9),
            XLSX.CellRange("A1:C10") => (type="duplicateValues", priority=8),
            XLSX.CellRange("A1:C10") => (type="uniqueValues", priority=7),
            XLSX.CellRange("A1:C10") => (type="notContainsBlanks", priority=6),
            XLSX.CellRange("A1:C10") => (type="containsBlanks", priority=5),
            XLSX.CellRange("A1:C10") => (type="notContainsErrors", priority=4),
            XLSX.CellRange("A1:C10") => (type="containsErrors", priority=3),
            XLSX.CellRange("A2:J2") => (type="containsErrors", priority=2),
            XLSX.CellRange("A2:J2") => (type="containsErrors", priority=1),
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

        @test sort!(XLSX.getConditionalFormats(s), by = x -> x.second.priority, rev=true) == [
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
        @test_throws MethodError XLSX.setConditionalFormat(s, "myNCRange", :containsErrors; # Non-contiguous ranges not allowed
            fill=["pattern" => "none", "bgColor" => "yellow"],
            format=["format" => "0.0"],
            font=["color" => "green"],
            border=["style" => "hair", "color" => "cyan"]
        )
        SAVE_FILES && save_outfile(f)

    end

end
