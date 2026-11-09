@testset "getcell" begin
    f = XLSX.newxlsx()
    s = f[1]
    for i in 1:3
        for j in 1:3
            s[i, j] = i + j
        end
    end
    @test XLSX.getcell(s, "A1") == XLSX.Cell(XLSX.get_workbook(f), XLSX.CellRef("A1"), "", "", "2", "", false)
    @test XLSX.getcell(s, "Sheet1!A1") == XLSX.Cell(XLSX.get_workbook(f), XLSX.CellRef("A1"), "", "", "2", "", false)
    @test XLSX.getcell(f, "Sheet1!A1") == XLSX.Cell(XLSX.get_workbook(f), XLSX.CellRef("A1"), "", "", "2", "", false)
    @test XLSX.getcell(s, XLSX.SheetCellRef("Sheet1!A1")) == XLSX.Cell(XLSX.get_workbook(f), XLSX.CellRef("A1"), "", "", "2", "", false)
    @test XLSX.getcell(f, XLSX.SheetCellRef("Sheet1!A1")) == XLSX.Cell(XLSX.get_workbook(f), XLSX.CellRef("A1"), "", "", "2", "", false)
    @test XLSX.getcell(s, "B1:B3") == [XLSX.Cell(XLSX.get_workbook(f), XLSX.CellRef("B1"), "", "", "3", "", false); XLSX.Cell(XLSX.get_workbook(f), XLSX.CellRef("B2"), "", "", "4", "", false); XLSX.Cell(XLSX.get_workbook(f), XLSX.CellRef("B3"), "", "", "5", "", false);;]
    @test XLSX.getcell(s, "Sheet1!B1:B3") == [XLSX.Cell(XLSX.get_workbook(f), XLSX.CellRef("B1"), "", "", "3", "", false); XLSX.Cell(XLSX.get_workbook(f), XLSX.CellRef("B2"), "", "", "4", "", false); XLSX.Cell(XLSX.get_workbook(f), XLSX.CellRef("B3"), "", "", "5", "", false);;]
    @test XLSX.getcell(f, "Sheet1!B1:B3") == [XLSX.Cell(XLSX.get_workbook(f), XLSX.CellRef("B1"), "", "", "3", "", false); XLSX.Cell(XLSX.get_workbook(f), XLSX.CellRef("B2"), "", "", "4", "", false); XLSX.Cell(XLSX.get_workbook(f), XLSX.CellRef("B3"), "", "", "5", "", false);;]
    @test XLSX.getcell(s, XLSX.SheetCellRange("Sheet1!B1:B3")) == [XLSX.Cell(XLSX.get_workbook(f), XLSX.CellRef("B1"), "", "", "3", "", false); XLSX.Cell(XLSX.get_workbook(f), XLSX.CellRef("B2"), "", "", "4", "", false); XLSX.Cell(XLSX.get_workbook(f), XLSX.CellRef("B3"), "", "", "5", "", false);;]
    @test XLSX.getcell(s, "B1,B3") == [[XLSX.Cell(XLSX.get_workbook(f), XLSX.CellRef("B1"), "", "", "3", "", false);;], [XLSX.Cell(XLSX.get_workbook(f), XLSX.CellRef("B3"), "", "", "5", "", false);;]]
    @test XLSX.getcell(s, "Sheet1!B1,Sheet1!B3") == [[XLSX.Cell(XLSX.get_workbook(f), XLSX.CellRef("B1"), "", "", "3", "", false);;], [XLSX.Cell(XLSX.get_workbook(f), XLSX.CellRef("B3"), "", "", "5", "", false);;]]
    @test XLSX.getcell(f, "Sheet1!B1,Sheet1!B3") == [[XLSX.Cell(XLSX.get_workbook(f), XLSX.CellRef("B1"), "", "", "3", "", false);;], [XLSX.Cell(XLSX.get_workbook(f), XLSX.CellRef("B3"), "", "", "5", "", false);;]]
    @test XLSX.getcell(s, "B:B") == [XLSX.Cell(XLSX.get_workbook(f), XLSX.CellRef("B1"), "", "", "3", "", false); XLSX.Cell(XLSX.get_workbook(f), XLSX.CellRef("B2"), "", "", "4", "", false); XLSX.Cell(XLSX.get_workbook(f), XLSX.CellRef("B3"), "", "", "5", "", false);;]
    @test XLSX.getcell(s, "Sheet1!B:B") == [XLSX.Cell(XLSX.get_workbook(f), XLSX.CellRef("B1"), "", "", "3", "", false); XLSX.Cell(XLSX.get_workbook(f), XLSX.CellRef("B2"), "", "", "4", "", false); XLSX.Cell(XLSX.get_workbook(f), XLSX.CellRef("B3"), "", "", "5", "", false);;]
    @test XLSX.getcell(f, "Sheet1!B:B") == [XLSX.Cell(XLSX.get_workbook(f), XLSX.CellRef("B1"), "", "", "3", "", false); XLSX.Cell(XLSX.get_workbook(f), XLSX.CellRef("B2"), "", "", "4", "", false); XLSX.Cell(XLSX.get_workbook(f), XLSX.CellRef("B3"), "", "", "5", "", false);;]
    @test XLSX.getcell(s, XLSX.SheetColumnRange("Sheet1!B:B")) == [XLSX.Cell(XLSX.get_workbook(f), XLSX.CellRef("B1"), "", "", "3", "", false); XLSX.Cell(XLSX.get_workbook(f), XLSX.CellRef("B2"), "", "", "4", "", false); XLSX.Cell(XLSX.get_workbook(f), XLSX.CellRef("B3"), "", "", "5", "", false);;]
    @test XLSX.getcell(s, "Sheet1!2:2") == [XLSX.Cell(XLSX.get_workbook(f), XLSX.CellRef("A2"), "", "", "3", "", false) XLSX.Cell(XLSX.get_workbook(f), XLSX.CellRef("B2"), "", "", "4", "", false) XLSX.Cell(XLSX.get_workbook(f), XLSX.CellRef("C2"), "", "", "5", "", false)]
    @test XLSX.getcell(f, "Sheet1!2:2") == [XLSX.Cell(XLSX.get_workbook(f), XLSX.CellRef("A2"), "", "", "3", "", false) XLSX.Cell(XLSX.get_workbook(f), XLSX.CellRef("B2"), "", "", "4", "", false) XLSX.Cell(XLSX.get_workbook(f), XLSX.CellRef("C2"), "", "", "5", "", false)]
    @test XLSX.getcell(s, XLSX.SheetRowRange("Sheet1!2:2")) == [XLSX.Cell(XLSX.get_workbook(f), XLSX.CellRef("A2"), "", "", "3", "", false) XLSX.Cell(XLSX.get_workbook(f), XLSX.CellRef("B2"), "", "", "4", "", false) XLSX.Cell(XLSX.get_workbook(f), XLSX.CellRef("C2"), "", "", "5", "", false)]
    @test XLSX.getcell(s, "2:2") == [XLSX.Cell(XLSX.get_workbook(f), XLSX.CellRef("A2"), "", "", "3", "", false) XLSX.Cell(XLSX.get_workbook(f), XLSX.CellRef("B2"), "", "", "4", "", false) XLSX.Cell(XLSX.get_workbook(f), XLSX.CellRef("C2"), "", "", "5", "", false)]
    @test XLSX.getcell(s, :, 2) == [XLSX.Cell(XLSX.get_workbook(f), XLSX.CellRef("B1"), "", "", "3", "", false); XLSX.Cell(XLSX.get_workbook(f), XLSX.CellRef("B2"), "", "", "4", "", false); XLSX.Cell(XLSX.get_workbook(f), XLSX.CellRef("B3"), "", "", "5", "", false);;]
    @test XLSX.getcell(s, 2, :) == [XLSX.Cell(XLSX.get_workbook(f), XLSX.CellRef("A2"), "", "", "3", "", false) XLSX.Cell(XLSX.get_workbook(f), XLSX.CellRef("B2"), "", "", "4", "", false) XLSX.Cell(XLSX.get_workbook(f), XLSX.CellRef("C2"), "", "", "5", "", false)]
    @test XLSX.getcell(s, 2, 1:2:3) == [XLSX.Cell(XLSX.get_workbook(f), XLSX.CellRef("A2"), "", "", "3", "", false), XLSX.Cell(XLSX.get_workbook(f), XLSX.CellRef("C2"), "", "", "5", "", false)]
    @test XLSX.getcell(s, 2, [1, 3]) == [XLSX.Cell(XLSX.get_workbook(f), XLSX.CellRef("A2"), "", "", "3", "", false), XLSX.Cell(XLSX.get_workbook(f), XLSX.CellRef("C2"), "", "", "5", "", false)]
    @test XLSX.getcell(s, [2], 1) == [XLSX.Cell(XLSX.get_workbook(f), XLSX.CellRef("A2"), "", "", "3", "", false)]
    @test XLSX.getcell(s, [2], [1, 3]) == [XLSX.Cell(XLSX.get_workbook(f), XLSX.CellRef("A2"), "", "", "3", "", false) XLSX.Cell(XLSX.get_workbook(f), XLSX.CellRef("C2"), "", "", "5", "", false)]
    @test_throws XLSX.XLSXError XLSX.getcell(f, "Sheet1!garbage")
    @test_throws XLSX.XLSXError XLSX.getcell(s, "Sheet1!garbage")
    @test_throws XLSX.XLSXError XLSX.getcell(s, "garbage")
    @test_throws XLSX.XLSXError XLSX.getcell(s, "garbage1:garbage2")

    @test XLSX.getcellrange(s, "Sheet1!B:B") == [XLSX.Cell(XLSX.get_workbook(f), XLSX.CellRef("B1"), "", "", "3", "", false); XLSX.Cell(XLSX.get_workbook(f), XLSX.CellRef("B2"), "", "", "4", "", false); XLSX.Cell(XLSX.get_workbook(f), XLSX.CellRef("B3"), "", "", "5", "", false);;]
    @test XLSX.getcellrange(f, "Sheet1!B:B") == [XLSX.Cell(XLSX.get_workbook(f), XLSX.CellRef("B1"), "", "", "3", "", false); XLSX.Cell(XLSX.get_workbook(f), XLSX.CellRef("B2"), "", "", "4", "", false); XLSX.Cell(XLSX.get_workbook(f), XLSX.CellRef("B3"), "", "", "5", "", false);;]
    @test XLSX.getcellrange(s, XLSX.SheetColumnRange("Sheet1!B:B")) == [XLSX.Cell(XLSX.get_workbook(f), XLSX.CellRef("B1"), "", "", "3", "", false); XLSX.Cell(XLSX.get_workbook(f), XLSX.CellRef("B2"), "", "", "4", "", false); XLSX.Cell(XLSX.get_workbook(f), XLSX.CellRef("B3"), "", "", "5", "", false);;]
    @test XLSX.getcellrange(s, "Sheet1!2:2") == [XLSX.Cell(XLSX.get_workbook(f), XLSX.CellRef("A2"), "", "", "3", "", false) XLSX.Cell(XLSX.get_workbook(f), XLSX.CellRef("B2"), "", "", "4", "", false) XLSX.Cell(XLSX.get_workbook(f), XLSX.CellRef("C2"), "", "", "5", "", false)]
    @test XLSX.getcellrange(f, "Sheet1!2:2") == [XLSX.Cell(XLSX.get_workbook(f), XLSX.CellRef("A2"), "", "", "3", "", false) XLSX.Cell(XLSX.get_workbook(f), XLSX.CellRef("B2"), "", "", "4", "", false) XLSX.Cell(XLSX.get_workbook(f), XLSX.CellRef("C2"), "", "", "5", "", false)]
    @test XLSX.getcellrange(s, XLSX.SheetRowRange("Sheet1!2:2")) == [XLSX.Cell(XLSX.get_workbook(f), XLSX.CellRef("A2"), "", "", "3", "", false) XLSX.Cell(XLSX.get_workbook(f), XLSX.CellRef("B2"), "", "", "4", "", false) XLSX.Cell(XLSX.get_workbook(f), XLSX.CellRef("C2"), "", "", "5", "", false)]

    XLSX.addDefinedName(f, "MyName1", "Sheet1!A1")
    XLSX.addDefinedName(s, "MyName2", "Sheet1!A2:A3")
    XLSX.addDefinedName(f, "MyName3", "Sheet1!A2,Sheet1!A3")
    s["MyName1"] = 12.9
    @test s["MyName1"] == 12.9
    s["MyName2"] = 42
    @test s["MyName2"] == [42; 42;;]
    @test XLSX.getcell(s, "MyName1") == XLSX.Cell(XLSX.get_workbook(f), XLSX.CellRef("A1"), "", "", "12.9", "", false)
    @test XLSX.getcell(s, "MyName2") == [XLSX.Cell(XLSX.get_workbook(f), XLSX.CellRef("A2"), "", "", "42", "", false); XLSX.Cell(XLSX.get_workbook(f), XLSX.CellRef("A3"), "", "", "42", "", false);;]
    @test XLSX.getcell(f, "MyName1") == XLSX.Cell(XLSX.get_workbook(f), XLSX.CellRef("A1"), "", "", "12.9", "", false)
    @test XLSX.getcellrange(s, "MyName2") == [XLSX.Cell(XLSX.get_workbook(f), XLSX.CellRef("A2"), "", "", "42", "", false); XLSX.Cell(XLSX.get_workbook(f), XLSX.CellRef("A3"), "", "", "42", "", false);;]
    @test XLSX.getcellrange(s, "MyName3") == [[XLSX.Cell(XLSX.get_workbook(f), XLSX.CellRef("A2"), "", "", "42", "", false);;], [XLSX.Cell(XLSX.get_workbook(f), XLSX.CellRef("A3"), "", "", "42", "", false);;]]
    @test XLSX.getcellrange(f, "MyName3") == [[XLSX.Cell(XLSX.get_workbook(f), XLSX.CellRef("A2"), "", "", "42", "", false);;], [XLSX.Cell(XLSX.get_workbook(f), XLSX.CellRef("A3"), "", "", "42", "", false);;]]

    SAVE_FILES && save_outfile(f)


    @testset "empty cells, hashing and error strings" begin
        e = XLSX.EmptyCell(XLSX.CellRef("A1"))
        @test isempty(e) && !XLSX.iserror(e)
        @test !isempty(XLSX.Cell(XLSX.CellRef("A1"), UInt64(1), UInt32(0), UInt16(0), XLSX.CT_INT, false))
        @test hash(e) == hash(XLSX.EmptyCell(XLSX.CellRef("A1")))
        @test XLSX.get_error_string(nothing) == ""
    end
end

@testset "Time and DateTime" begin
    @test XLSX.excel_value_to_time(0.82291666666666663) == Dates.Time(Dates.Hour(19), Dates.Minute(45))
    @test XLSX.time_to_excel_value(XLSX.excel_value_to_time(0.2)) == 0.2
    @test XLSX.excel_value_to_datetime(43206.805447106482, false) == Date(2018, 4, 16) + Dates.Time(Dates.Hour(19), Dates.Minute(19), Dates.Second(51))
    @test XLSX.excel_value_to_datetime(XLSX.datetime_to_excel_value(Date(2018, 4, 16) + Dates.Time(Dates.Hour(19), Dates.Minute(19), Dates.Second(51)), false), false) == Date(2018, 4, 16) + Dates.Time(Dates.Hour(19), Dates.Minute(19), Dates.Second(51))

    dt = Date(2018, 4, 1)
    @test XLSX.excel_value_to_date(XLSX.date_to_excel_value(dt, false), false) == dt
    @test XLSX.excel_value_to_date(XLSX.date_to_excel_value(dt, true), true) == dt
end

@testset "stream iterator" begin
    f = XLSX.openxlsx(joinpath(data_directory, "general.xlsx"), enable_cache=false)
    s = f["table"]
    for sheetrow in XLSX.eachrow(s)
        for column in 2:4
            cell = XLSX.getcell(sheetrow, column)
            if XLSX.row_number(cell) == 2 && XLSX.column_number(cell) == 2
                @test XLSX.getdata(s, cell) == "Column B"
            end
            if XLSX.row_number(cell) == 12 && XLSX.column_number(cell) == 2
                @test XLSX.getdata(s, cell) == "trash"
            end
        end
    end

# Issue 446
    f = XLSX.openxlsx(joinpath(data_directory, "general.xlsx"), enable_cache=false)

    # The stream iterator moves on from row to row (the failure mode was row 1
    # forever), and yields the rows the file holds, in order.
    seen = [XLSX.row_number(r) for r in XLSX.eachrow(f[1])]
    expected = [XLSX.row_number(r) for r in XLSX.eachrow(XLSX.readxlsx(joinpath(data_directory, "general.xlsx"))[1])]
    @test seen == expected
    @test length(seen) > 1

end

@testset "Base.eachrow, and the cache flag on every holder" begin
    xf = XLSX.newxlsx()
    sh = xf[1]
    sh[1:3, 1:4] = reshape(collect(1:12), 3, 4)

    @test length(collect(Base.eachrow(sh))) == length(collect(XLSX.eachrow(sh)))
    itr = XLSX.eachrow(sh)
    @test XLSX.is_cache_enabled(XLSX.get_workbook(xf)) == XLSX.is_cache_enabled(itr) == XLSX.is_cache_enabled(xf)
end
