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

end