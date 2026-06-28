@testset "Ranges" begin
    @testset "Range intersect" begin
        @test XLSX.intersects(XLSX.CellRange(XLSX.CellRef("A3"), XLSX.CellRef("E4")), XLSX.CellRange(XLSX.CellRef("C1"), XLSX.CellRef("D6")))
        @test XLSX.intersects(XLSX.CellRange(XLSX.CellRef("C1"), XLSX.CellRef("D6")), XLSX.CellRange(XLSX.CellRef("A3"), XLSX.CellRef("E4")))
        @test XLSX.intersects(XLSX.CellRange(XLSX.CellRef("A3"), XLSX.CellRef("D4")), XLSX.CellRange(XLSX.CellRef("C1"), XLSX.CellRef("E6")))
        @test XLSX.intersects(XLSX.CellRange(XLSX.CellRef("C1"), XLSX.CellRef("E6")), XLSX.CellRange(XLSX.CellRef("A3"), XLSX.CellRef("D4")))
        @test XLSX.intersects(XLSX.CellRange(XLSX.CellRef("C3"), XLSX.CellRef("E4")), XLSX.CellRange(XLSX.CellRef("C1"), XLSX.CellRef("G6")))
        @test XLSX.intersects(XLSX.CellRange(XLSX.CellRef("C1"), XLSX.CellRef("G6")), XLSX.CellRange(XLSX.CellRef("C3"), XLSX.CellRef("E4")))
        @test XLSX.intersects(XLSX.CellRange(XLSX.CellRef("A1"), XLSX.CellRef("E4")), XLSX.CellRange(XLSX.CellRef("C3"), XLSX.CellRef("G6")))
        @test XLSX.intersects(XLSX.CellRange(XLSX.CellRef("C3"), XLSX.CellRef("G6")), XLSX.CellRange(XLSX.CellRef("A1"), XLSX.CellRef("E4")))
        @test XLSX.intersects(XLSX.CellRange(XLSX.CellRef("C3"), XLSX.CellRef("D4")), XLSX.CellRange(XLSX.CellRef("A1"), XLSX.CellRef("E6")))
    end

    @testset "Column Range" begin

        @testset "Single Column" begin
            c = XLSX.ColumnRange("C")
            @test c.start == 3
            @test c.stop == 3
            show(IOBuffer(), c)

            c = XLSX.ColumnRange("AA")
            @test c.start == 27
            @test c.stop == 27
        end

        @testset "Multiple Columns" begin
            c = XLSX.ColumnRange("A:Z")
            @test c.start == 1
            @test c.stop == 26

            c = XLSX.ColumnRange("A:AA")
            @test c.start == 1
            @test c.stop == 27

            cr = XLSX.ColumnRange("B:D")
            @test string(cr) == "B:D"
            @test cr.start == 2
            @test cr.stop == 4
            @test length(cr) == 3

            cr = XLSX.ColumnRange("B")
            @test string(cr) == "B:B"
            @test cr.start == 2
            @test cr.stop == 2
            @test length(cr) == 1

            @test XLSX.ColumnRange("B1:D3") == XLSX.ColumnRange("B:D")

            @test_throws XLSX.XLSXError XLSX.ColumnRange("D:A")
            @test collect(XLSX.ColumnRange("B:D")) == ["B", "C", "D"]
            @test XLSX.ColumnRange("B:D") == XLSX.ColumnRange("B:D")
            @test hash(XLSX.ColumnRange("B:D")) == hash(XLSX.ColumnRange("B:D"))
        end
    end

    @testset "Row Range" begin # Issue #150
        cr = XLSX.RowRange("2:5")
        @test string(cr) == "2:5"
        @test cr.start == 2
        @test cr.stop == 5
        @test length(cr) == 4
        @test collect(cr) == ["2", "3", "4", "5"]

        cr = XLSX.RowRange("2")
        @test string(cr) == "2:2"
        @test cr.start == 2
        @test cr.stop == 2
        @test length(cr) == 1
        @test collect(cr) == ["2"]

        @test XLSX.RowRange("B1:D3") == XLSX.RowRange("1:3")
        @test_throws XLSX.XLSXError XLSX.RowRange("5:2")
        @test XLSX.RowRange("2:5") == XLSX.RowRange("2:5")
        @test hash(XLSX.RowRange("2:5")) == hash(XLSX.RowRange("2:5"))
    end

    @testset "Non-contiguous Range" begin
        cr = XLSX.NonContiguousRange("Sheet1!D1:D3,Sheet1!B1:B3")
        @test string(cr) == "Sheet1!D1:D3,Sheet1!B1:B3"
        @test cr.sheet == "Sheet1"
        @test cr.rng == [XLSX.CellRange("D1:D3"), XLSX.CellRange("B1:B3")]
        @test length(cr) == 6
        @test length(XLSX.NonContiguousRange("Sheet1!B1:B1,Sheet1!B1")) == 1
        @test collect(cr.rng) == [XLSX.CellRange("D1:D3"), XLSX.CellRange("B1:B3")]
        @test XLSX.NonContiguousRange("Sheet1!D1:D3,Sheet1!B1:B3") == XLSX.NonContiguousRange("Sheet1!D1:D3,Sheet1!B1:B3")
        @test hash(XLSX.NonContiguousRange("Sheet1!D1:D3,Sheet1!B1:B3")) == hash(XLSX.NonContiguousRange("Sheet1!D1:D3,Sheet1!B1:B3"))

        f = XLSX.newxlsx("Sheet 1")
        s = f["Sheet 1"]
        for cell in XLSX.CellRange("A1:D6")
            s[cell] = ""
        end
        cr = XLSX.NonContiguousRange(s, "D1:D3,A2,B1:B3")
        @test string(cr) == "'Sheet 1'!D1:D3,'Sheet 1'!A2,'Sheet 1'!B1:B3"
        @test cr.sheet == "Sheet 1"
        @test cr.rng == [XLSX.CellRange("D1:D3"), XLSX.CellRef("A2"), XLSX.CellRange("B1:B3")]
        @test length(cr) == 7
        @test length(XLSX.NonContiguousRange(s, "B1:B1,B1")) == 1
        @test collect(cr.rng) == [XLSX.CellRange("D1:D3"), XLSX.CellRef("A2"), XLSX.CellRange("B1:B3")]
        @test XLSX.NonContiguousRange(s, "D1:D3,A2,B1:B3") == XLSX.NonContiguousRange(s, "D1:D3,A2,B1:B3")
        @test hash(XLSX.NonContiguousRange(s, "D1:D3,A2,B1:B3")) == hash(XLSX.NonContiguousRange(s, "D1:D3,A2,B1:B3"))

        @test_throws XLSX.XLSXError XLSX.NonContiguousRange("Sheet1!D1:D3,B1:B3")
        @test_throws XLSX.XLSXError XLSX.NonContiguousRange("Sheet1!D1:D3,Sheet2!B1:B3")
        @test_throws XLSX.XLSXError XLSX.NonContiguousRange("B1:D3")
        @test_throws XLSX.XLSXError XLSX.NonContiguousRange("2:5")

        SAVE_FILES && save_outfile(f)
    end

    @testset "CellRange iterator" begin
        rng = XLSX.CellRange("A2:C4")
        @test collect(rng) == [XLSX.CellRef("A2"), XLSX.CellRef("B2"), XLSX.CellRef("C2"), XLSX.CellRef("A3"), XLSX.CellRef("B3"), XLSX.CellRef("C3"), XLSX.CellRef("A4"), XLSX.CellRef("B4"), XLSX.CellRef("C4")]
    end
end