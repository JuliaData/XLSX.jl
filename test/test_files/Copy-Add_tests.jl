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

    @testset "copysheet!/deletesheet! part cleanup (issue #427)" begin

        # Shared helper: map a sheet name to the worksheet XML path it actually
        # lives at, by reading workbook.xml + workbook.xml.rels directly rather
        # than assuming file numbering.
        function sheet_xml_path(zipbytes::Vector{UInt8}, sheetname::String)
            r = ZipReader(zipbytes)
            wbxml  = String(zip_readentry(r, "xl/workbook.xml"))
            wbrels = String(zip_readentry(r, "xl/_rels/workbook.xml.rels"))
            sheet_rid = Dict(m.captures[1] => m.captures[2]
                for m in eachmatch(r"<sheet name=\"([^\"]+)\"[^>]*r:id=\"(rId\d+)\"", wbxml))
            rid_target = Dict(m.captures[1] => m.captures[2]
                for m in eachmatch(r"<Relationship Id=\"(rId\d+)\"[^>]*Target=\"([^\"]+)\"", wbrels))
            return "xl/" * rid_target[sheet_rid[sheetname]]
        end

        @testset "deletesheet! on a table-only sheet (no drawing) removes its rels file" begin
            # UTF-16.xlsx's SourceData sheet has a table but no drawing.
            src = joinpath(data_directory, "UTF-16.xlsx")
            p = tempname() * ".xlsx"
            cp(src, p)

            XLSX.openxlsx(p; mode="rw") do xf
                XLSX.deletesheet!(xf, "SourceData")
            end

            r = ZipReader(read(p))
            names = zip_names(r)

            @test !("xl/tables/table1.xml" in names)
            @test !("xl/worksheets/_rels/sheet1.xml.rels" in names)
            @test !occursin("tables/table1.xml", String(zip_readentry(r, "[Content_Types].xml")))

            XLSX.openxlsx(p) do xf2
                @test XLSX.sheetnames(xf2) == ["Sheet1"]
            end
            rm(p; force=true)
        end

        @testset "deletesheet! removes table, comments, VML, drawing, media, and their Overrides together" begin
            src = joinpath(data_directory, "TableCommentsVML.xlsx")
            p = tempname() * ".xlsx"
            cp(src, p)

            XLSX.openxlsx(p; mode="rw") do xf
                XLSX.addsheet!(xf, "Blank")   # deletesheet! refuses to delete the only sheet
                XLSX.deletesheet!(xf, "Sheet1")
            end

            r = ZipReader(read(p))
            names = zip_names(r)

            @test !("xl/tables/table1.xml" in names)
            @test !("xl/comments1.xml" in names)
            @test !("xl/drawings/vmlDrawing1.vml" in names)
            @test !("xl/drawings/drawing1.xml" in names)
            @test !("xl/media/image1.png" in names)
            @test !("xl/worksheets/_rels/sheet1.xml.rels" in names)
            @test !("xl/worksheets/sheet1.xml" in names)

            ctypes = String(zip_readentry(r, "[Content_Types].xml"))
            @test !occursin("tables/table1.xml", ctypes)
            @test !occursin("comments1.xml", ctypes)
            @test !occursin("drawings/drawing1.xml", ctypes)

            XLSX.openxlsx(p) do xf2
                @test XLSX.sheetnames(xf2) == ["Blank"]
            end
            rm(p; force=true)
        end

        @testset "copysheet! strips table/comments/VML but still duplicates the drawing" begin
            src = joinpath(data_directory, "TableCommentsVML.xlsx")
            p = tempname() * ".xlsx"
            cp(src, p)

            XLSX.openxlsx(p; mode="rw") do xf
                XLSX.copysheet!(xf["Sheet1"], "Copy")
            end

            zipbytes = read(p)
            r = ZipReader(zipbytes)
            names = zip_names(r)

            source_path = sheet_xml_path(zipbytes, "Sheet1")
            copy_path   = sheet_xml_path(zipbytes, "Copy")
            source_xml  = String(zip_readentry(r, source_path))
            copy_xml    = String(zip_readentry(r, copy_path))

            # Stripped from the copy...
            @test !occursin("tableParts", copy_xml)
            @test !occursin("legacyDrawing", copy_xml)

            # ...but the original is untouched (catches copynode() aliasing children
            # between the original and the copy rather than deep-copying).
            @test occursin("tableParts", source_xml)
            @test occursin("legacyDrawing", source_xml)

            # The underlying parts still belong to the original sheet.
            @test "xl/tables/table1.xml" in names
            @test "xl/comments1.xml" in names
            @test "xl/drawings/vmlDrawing1.vml" in names

            # The drawing/image IS carried over — duplicated and relinked, not shared.
            m = match(r"<drawing r:id=\"(rId\d+)\"", copy_xml)
            @test m !== nothing

            if m !== nothing
                copy_rels_path = let (d, f) = rsplit(copy_path, "/"; limit=2)
                    "$d/_rels/$f.rels"
                end
                @test copy_rels_path in names

                if copy_rels_path in names
                    copy_rels_xml = String(zip_readentry(r, copy_rels_path))
                    dt = match(Regex("<Relationship Id=\"$(m.captures[1])\"[^>]*Target=\"([^\"]+)\""), copy_rels_xml)
                    @test dt !== nothing
                    if dt !== nothing
                        new_drawing_path = "xl/" * replace(dt.captures[1], "../" => "")
                        @test new_drawing_path in names
                        @test new_drawing_path != "xl/drawings/drawing1.xml"
                    end

                    # The copy's rels must not have picked up comments/VML relationships
                    # it has no corresponding sheet-XML reference for anymore.
                    @test !occursin("relationships/comments", copy_rels_xml)
                    @test !occursin("relationships/vmlDrawing", copy_rels_xml)
                end
            end

            XLSX.openxlsx(p) do xf2
                @test XLSX.sheetnames(xf2) == ["Sheet1", "Copy"]
            end
            rm(p; force=true)
        end

    end
end
