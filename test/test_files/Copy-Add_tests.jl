@testset "add/copy sheet!" begin

    @testset "addsheet!" begin

        new_filename = "template_with_new_sheet.xlsx"
        f = XLSX.open_empty_template()
        s = XLSX.addsheet!(f, "new_sheet")
        s["A1"] = 10
        @test XLSX.sheetnames(f) == ["Sheet1", "new_sheet"]
        XLSX.writexlsx(new_filename, f, overwrite=true)
        SAVE_FILES && save_outfile(new_filename)


        big_sheetname = "aaaaaaaaaabbbbbbbbbbccccccccccd"
        s2 = XLSX.addsheet!(f, big_sheetname)

        XLSX.writexlsx(new_filename, f, overwrite=true)
        SAVE_FILES && save_outfile(new_filename)
        fx = XLSX.opentemplate(new_filename)
        @test XLSX.sheetnames(f) == ["Sheet1", "new_sheet", big_sheetname]

    end

    @testset "invalid sheet names" begin

        f = XLSX.open_empty_template()
        s = XLSX.addsheet!(f, "new_sheet")
        s["A1"] = 10
        invalid_names = [
            "aaaaaaaaaabbbbbbbbbbccccccccccd1",
            "abc:def",
            "abcdef/",
            "\\aaaa",
            "hey?you",
            "[mysheet]",
            "asteri*"
        ]

        for invalid_name in invalid_names
            @test_throws XLSX.XLSXError XLSX.addsheet!(f, invalid_name)
        end

    end

    @testset "copysheet!" begin

        f = XLSX.newxlsx()
        XLSX.renamesheet!(f["Sheet1"], "new_name")
        XLSX.addsheet!(f)
        for x = 1:10, y = 1:10
            f["Sheet1"][x, y] = x + y
            f["new_name"][x, y] = x * y
        end
        XLSX.addDefinedName(f["new_name"], "new_name_range", "A1:B10")
        XLSX.addDefinedName(f["Sheet1"], "Sheet1_range", "C1:D10")
        XLSX.setBorder(f["new_name"], "A1:D10"; allsides=["style" => "thin", "color" => "red"])
        XLSX.setBorder(f["Sheet1"], "A1:D10"; allsides=["style" => "thin", "color" => "red"])
        XLSX.setConditionalFormat(f["new_name"], "A1:D10", :colorScale)

        s3 = XLSX.copysheet!(f["new_name"], "copied_sheet")
        @test s3.name == "copied_sheet"
        @test s3["A1"] == 1
        @test s3[5, 5] == 25
        @test s3[10, 10] == 100
        @test XLSX.get_workbook(s3).worksheet_names == XLSX.get_workbook(f["new_name"]).worksheet_names
        @test XLSX.getConditionalFormats(s3) == XLSX.getConditionalFormats(f["new_name"])
        @test XLSX.getBorder(s3, "C5").border == XLSX.getBorder(f["new_name"], "C5").border

        # Check that the original sheet is unchanged
        s2 = f["new_name"]
        @test s2["A1"] == 1
        @test s2[5, 5] == 25
        @test s2[10, 10] == 100

        s4 = XLSX.copysheet!(s3)
        @test s4.name == "copied_sheet (copy)"
        @test s4["A1"] == 1
        @test s4[5, 5] == 25
        @test s4[10, 10] == 100

        @test XLSX.get_workbook(s4).worksheet_names == XLSX.get_workbook(f["new_name"]).worksheet_names
        XLSX.setBorder(s4, "F1:H10"; allsides=["style" => "thin", "color" => "green"])
        XLSX.setConditionalFormat(s4, "F1:H10", :colorScale; colorscale="redyellowgreen")

        XLSX.writexlsx("copied_sheets.xlsx", f, overwrite=true)
        SAVE_FILES && save_outfile("copied_sheets.xlsx")
        f = XLSX.opentemplate("copied_sheets.xlsx")
        @test XLSX.sheetnames(f) == ["new_name", "Sheet1", "copied_sheet", "copied_sheet (copy)"]
        @test XLSX.get_workbook(f["copied_sheet"]).worksheet_names == XLSX.get_workbook(f["new_name"]).worksheet_names
        @test XLSX.getConditionalFormats(f["copied_sheet (copy)"]) == XLSX.getConditionalFormats(s4)
        @test XLSX.getBorder(f["copied_sheet (copy)"], "C5").border == XLSX.getBorder(f["new_name"], "C5").border
        @test XLSX.getBorder(f["copied_sheet (copy)"], "G5").border == XLSX.getBorder(s4, "G5").border

    end
    isfile("copied_sheets.xlsx") && rm("copied_sheets.xlsx")

    @testset "deletesheet!" begin

        new_filename = "template_with_new_sheet.xlsx"
        big_sheetname = "aaaaaaaaaabbbbbbbbbbccccccccccd"
        fx = XLSX.opentemplate(new_filename)
        XLSX.deletesheet!(fx, big_sheetname)
        @test XLSX.sheetnames(fx) == ["Sheet1", "new_sheet"]
        XLSX.writexlsx(new_filename, fx, overwrite=true)
        SAVE_FILES && save_outfile(new_filename)
        f = XLSX.readxlsx(new_filename)
        @test XLSX.sheetnames(f) == ["Sheet1", "new_sheet"]

        f = XLSX.opentemplate(joinpath(data_directory, "general.xlsx"))
        sc = XLSX.sheetcount(f)
        XLSX.deletesheet!(f, "empty")
        @test XLSX.sheetcount(f) == sc - 1 # Check it's gone.
        @test XLSX.hassheet(f, "empty") == false # Check it's gone.
        @test_throws XLSX.XLSXError XLSX.deletesheet!(f, "empty") # Already deleted.
        @test_throws XLSX.XLSXError XLSX.deletesheet!(f, "nosuchsheet") # Never there.
        s2 = XLSX.addsheet!(f, "this_now")
        @test XLSX.sheetnames(f) == ["general", "table3", "table4", "table", "table2", "table5", "table6", "table7", "lookup", "header_error", "named_ranges_2", "named_ranges", "this_now"]
        XLSX.writexlsx(new_filename, f, overwrite=true)
        SAVE_FILES && save_outfile(new_filename)

        f = XLSX.opentemplate(new_filename)
        @test XLSX.sheetnames(f) == ["general", "table3", "table4", "table", "table2", "table5", "table6", "table7", "lookup", "header_error", "named_ranges_2", "named_ranges", "this_now"]
        XLSX.deletesheet!(f, "named_ranges")
        XLSX.deletesheet!(f["general"])
        @test XLSX.sheetnames(f) == ["table3", "table4", "table", "table2", "table5", "table6", "table7", "lookup", "header_error", "named_ranges_2", "this_now"]
        XLSX.writexlsx(new_filename, f, overwrite=true)
        SAVE_FILES && save_outfile(new_filename)
        dtable = XLSX.readtable(new_filename, "table4", "F:G")
        data, col_names = dtable.data, dtable.column_labels
        @test col_names == [:H2, :H3]
        test_data = Any[Any["C3", missing], Any[missing, "D4"]]
        check_test_data(data, test_data)
        @test XLSX.deletesheet!(f, 1) === f
        @test XLSX.sheetnames(f) == ["table4", "table", "table2", "table5", "table6", "table7", "lookup", "header_error", "named_ranges_2", "this_now"]
        XLSX.writexlsx(new_filename, f, overwrite=true)
        SAVE_FILES && save_outfile(new_filename)
        dtable = XLSX.readtable(new_filename, "table4", "F:G")
        data, col_names = dtable.data, dtable.column_labels
        @test col_names == [:H2, :H3]
        test_data = Any[Any["C3", missing], Any[missing, "D4"]]
        check_test_data(data, test_data)

        f = XLSX.opentemplate(joinpath(data_directory, "Book_1904.xlsx")) # Only one sheet - can't delete
        @test_throws XLSX.XLSXError XLSX.deletesheet!(f, 1)
        s = f[1]
        @test_throws XLSX.XLSXError XLSX.deletesheet!(s)
        @test_throws XLSX.XLSXError XLSX.deletesheet!(f, "Sheet1")

        f = XLSX.openxlsx(joinpath(data_directory, "deletesheet.xlsx"), mode="rw")
        XLSX.deletesheet!(f[1])
        @test XLSX.getcell(f[1], "A1") == XLSX.Cell(XLSX.get_workbook(f), XLSX.CellRef("A1"), "e", "", "#REF!", "", true)
        @test XLSX.getFormula(f[1], "A1") == "=#REF!+#REF!"
        SAVE_FILES && save_outfile(f)

        isfile("template_with_new_sheet.xlsx") && rm("template_with_new_sheet.xlsx")

        # XLSX.jl cannot currently create formula references that refer to other sheets, so we have to manually 
        # insert them into the workbook's formula cache to test. Excel cannot read such a workbook and deletes these 
        # formulas on opening. Tests are for completeness only
        f=XLSX.openxlsx("renamedelete.xlsx", mode="w") 
        s=f[1]
        s[1:10, 1] = collect(1:10)
        XLSX.addsheet!(f)
        s2=f[2]
        s2["A1:B4"]=""
        XLSX.setFormula(s2, "A1", "=Sheet1!A1+10")
        XLSX.getcell(s2, "B1").formula=true
        XLSX.getcell(s2, "B2").formula=true
        XLSX.getcell(s2, "B3").formula=true
        XLSX.getcell(s2, "B4").formula=true
        w=XLSX.get_workbook(s2)
        w.formulas[XLSX.SheetCellRef("Sheet2!B1")] = XLSX.ReferencedFormula("=Sheet1!A1+10", 0, "B1:B4", nothing)
        w.formulas[XLSX.SheetCellRef("Sheet2!B2")] = XLSX.FormulaReference(0, nothing)
        w.formulas[XLSX.SheetCellRef("Sheet2!B3")] = XLSX.FormulaReference(0, nothing)
        w.formulas[XLSX.SheetCellRef("Sheet2!B4")] = XLSX.FormulaReference(0, nothing)

        XLSX.deletesheet!(s)
        @test XLSX.hassheet(f, "Sheet1") == false
        @test XLSX.getcell(s2, "A1") == XLSX.Cell(XLSX.get_workbook(f), XLSX.CellRef("A1"), "e", "", "#REF!", "", true)
        @test XLSX.getcell(s2, "B1") == XLSX.Cell(XLSX.get_workbook(f), XLSX.CellRef("B1"), "e", "", "#REF!", "", true)
        @test XLSX.getcell(s2, "B2") == XLSX.Cell(XLSX.get_workbook(f), XLSX.CellRef("B2"), "e", "", "#REF!", "", true)
        @test XLSX.getcell(s2, "B4") == XLSX.Cell(XLSX.get_workbook(f), XLSX.CellRef("B4"), "e", "", "#REF!", "", true)
        @test XLSX.getFormula(s2, "A1") == "=#REF!+10"
        @test XLSX.getFormula(s2, "B1") == "=#REF!+10"
        @test XLSX.getFormula(s2, "B2") == "=#REF!+10"
        @test XLSX.getFormula(s2, "B4") == "=#REF!+10"
        @test XLSX.get_formula_from_cache(s2, XLSX.CellRef("A1")) == XLSX.Formula("=#REF!+10", nothing, nothing, nothing)
        @test XLSX.get_formula_from_cache(s2, XLSX.CellRef("B1")) == XLSX.ReferencedFormula("=#REF!+10", 0, "B1:B4", nothing)
        @test XLSX.get_formula_from_cache(s2, XLSX.CellRef("B2")) ==  XLSX.FormulaReference(0, nothing)
        @test XLSX.get_formula_from_cache(s2, XLSX.CellRef("B4")) ==  XLSX.FormulaReference(0, nothing)
        SAVE_FILES && save_outfile(f)

    end
    @testset "renamesheet!" begin

        f=XLSX.openxlsx("renamedelete.xlsx", mode="w")
        s=f[1]
        s[1:10, 1] = collect(1:10)
        XLSX.addsheet!(f, "newsheet")
        s2 = f["newsheet"]
        XLSX.setFormula(s2, "A1", "=Sheet1!A1+10")
        XLSX.setFormula(s2, "B1:B3", "=A1+10")

        @test XLSX.getcell(s2, "A1") == XLSX.Cell(XLSX.get_workbook(f), XLSX.CellRef("A1"), "", "", "", "", true)
        @test XLSX.getFormula(s2, "A1") == "=Sheet1!A1+10"
        @test XLSX.get_formula_from_cache(s2, XLSX.CellRef("A1")) == XLSX.Formula("=Sheet1!A1+10", nothing, nothing, nothing)

        w = XLSX.get_workbook(s2)
        w.formulas[XLSX.SheetCellRef("newsheet!B1")] = XLSX.ReferencedFormula("=A1+10", 0, "B1:B3", nothing)
        w.formulas[XLSX.SheetCellRef("newsheet!B2")] = XLSX.FormulaReference(0, nothing)
        w.formulas[XLSX.SheetCellRef("newsheet!B3")] = XLSX.FormulaReference(0, nothing)

        XLSX.renamesheet!(s2, "Sheet4")
        @test s2.name == "Sheet4"
        @test XLSX.getcell(s2, "A1") == XLSX.Cell(XLSX.get_workbook(f), XLSX.CellRef("A1"), "", "", "", "", true)
        @test XLSX.getFormula(s2, "A1") == "=Sheet1!A1+10"
        @test XLSX.get_formula_from_cache(s2, XLSX.CellRef("A1")) == XLSX.Formula("=Sheet1!A1+10", nothing, nothing, nothing)
        @test XLSX.getFormula(s2, "B1") == "=A1+10"
        @test XLSX.getFormula(s2, "B2") == "=A2+10"
        @test XLSX.getFormula(s2, "B3") == "=A3+10"

        XLSX.renamesheet!(s, "Sheet3")
        @test s.name == "Sheet3"
        @test XLSX.getcell(s2, "A1") == XLSX.Cell(XLSX.get_workbook(f), XLSX.CellRef("A1"), "", "", "", "", true)
        @test XLSX.getFormula(s2, "A1") == "=Sheet3!A1+10"
        @test XLSX.get_formula_from_cache(s2, XLSX.CellRef("A1")) == XLSX.Formula("=Sheet3!A1+10", nothing, nothing, nothing)
        SAVE_FILES && save_outfile(f)

    end
end
