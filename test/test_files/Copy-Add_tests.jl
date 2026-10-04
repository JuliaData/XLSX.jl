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
        SAVE_FILES && save_outfile(fx)

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
        SAVE_FILES && save_outfile(f)

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

        SAVE_FILES && save_outfile(f)
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
        @test XLSX.get_formula_from_cache(s2, XLSX.CellRef("B2")) == XLSX.FormulaReference(0, nothing)
        @test XLSX.get_formula_from_cache(s2, XLSX.CellRef("B4")) == XLSX.FormulaReference(0, nothing)
        SAVE_FILES && save_outfile(f)

    end
    @testset "deletesheet! keeps the active tab on the same sheet" begin
        xf = XLSX.newxlsx("A")
        XLSX.addsheet!(xf, "B")
        XLSX.addsheet!(xf, "C")
        active(x) = parse(Int, XLSX.XML.attributes(first(XLSX.elements_with_tag(first(XLSX.elements_with_tag(
            XLSX.xml_root_element(XLSX.xmlroot(x, "xl/workbook.xml")), "bookViews")), "workbookView")))["activeTab"])
        XLSX.xmlroot(xf, "xl/workbook.xml")                                    # ensure loaded
        wv = first(XLSX.elements_with_tag(first(XLSX.elements_with_tag(
            XLSX.xml_root_element(XLSX.xmlroot(xf, "xl/workbook.xml")), "bookViews")), "workbookView"))
        wv["activeTab"] = "2"                                                  # C is active
        XLSX.deletesheet!(xf, "A")
        @test active(xf) == 1                                                  # still C
        XLSX.deletesheet!(xf, "C")
        @test active(xf) == 0                                                  # C gone: clamped to B
        SAVE_FILES && save_outfile(xf)
    end

    # Chart references to a deleted sheet, invalidated as Excel does it.

    function _chart_elements(node, name)
        out = []
        for ch in XLSX.XML.eachelement(node)
            XLSX.localname(ch) == name && push!(out, ch)
            append!(out, _chart_elements(ch, name))
        end
        return out
    end

    _chart_f_texts(root) = [something(XLSX.XML.is_simple_value(f), "") for f in _chart_elements(root, "f")]

    @testset "deletesheet! invalidates chart references" begin

        @testset "c: chart becomes #REF!, cache kept" begin
            xf = XLSX.openxlsx(joinpath(data_directory, "Chartsheet.xlsx"); mode="rw")
            XLSX.addsheet!(xf, "Keep")
            XLSX.deletesheet!(xf, "Sheet1")

            c = only(XLSX.getCharts(xf))
            @test _chart_f_texts(XLSX.Charts.chart_root(c)) == ["#REF!", "#REF!"]   # as Excel writes it
            @test all(r -> r.values === XLSX.XL_REF, XLSX.getChartRanges(c))
            @test XLSX.getChartData(xf, c.name) isa XLSX.DataTable                  # cache kept

            f = "deletesheet_chart_ref.xlsx"
            XLSX.writexlsx(f, xf; overwrite=true)
            SAVE_FILES && save_outfile(xf)
            yf = XLSX.openxlsx(f; mode="rw")
            d = only(XLSX.getCharts(yf))
            @test all(r -> r.values === XLSX.XL_REF, XLSX.getChartRanges(d))
            a, b = XLSX.getChartData(yf, d.name), XLSX.getChartData(xf, c.name)
            @test a.column_labels == b.column_labels && a.data == b.data
            SAVE_FILES && save_outfile(yf)
            rm(f; force=true)
        end

        @testset "chartEx chart loses the reference" begin
            # chartex_kinds.xlsx with the treemap moved to a chartsheet in Excel,
            # so deleting its data sheet leaves the chart behind.
            xf = XLSX.openxlsx(joinpath(data_directory, "chartex_chartsheet.xlsx"); mode="rw")
            others = filter(c -> XLSX.Charts.getChartType(c) !== :treemap, XLSX.getCharts(xf))
            before = Dict(c.name => string.(XLSX.getChartRanges(c)) for c in others)

            XLSX.deletesheet!(xf, "treemap")

            c = only(filter(c -> XLSX.Charts.getChartType(c) === :treemap, XLSX.getCharts(xf)))
            root = XLSX.Charts.chart_root(c)
            @test isempty(_chart_f_texts(root))
            strdim = only(_chart_elements(root, "strDim"))
            numdim = only(_chart_elements(root, "numDim"))
            @test length(_chart_elements(strdim, "lvl")) == 2                      # one per hierarchy level
            @test length(_chart_elements(numdim, "lvl")) == 1
            @test all(l -> XLSX.get_attr(l, "ptCount") == "0",
                vcat(_chart_elements(strdim, "lvl"), _chart_elements(numdim, "lvl")))
            @test isnothing(XLSX.first_element_with_tag(only(_chart_elements(root, "series")), "tx"))
            @test all(isnothing, XLSX.getChartRanges(c))

            wb = XLSX.get_workbook(xf)
            @test !any(v -> v.value isa XLSX.DefinedNameRangeTypes && v.value.sheet == "treemap",
                values(wb.workbook_names))

            # charts plotting other sheets are untouched
            for d in filter(d -> haskey(before, d.name), XLSX.getCharts(xf))
                @test string.(XLSX.getChartRanges(d)) == before[d.name]
            end

            f = "deletesheet_chartex_ref.xlsx"
            XLSX.writexlsx(f, xf; overwrite=true)
            SAVE_FILES && save_outfile(xf)
            yf = XLSX.openxlsx(f; mode="rw")
            d = only(filter(d -> XLSX.Charts.getChartType(d) === :treemap, XLSX.getCharts(yf)))
            @test all(isnothing, XLSX.getChartRanges(d))
            SAVE_FILES && save_outfile(yf)
            rm(f; force=true)
        end
        @testset "levels in a chartEx reference" begin
            @test XLSX._ref_levels("Data!\$A\$2:\$B\$7", "") == 2         # columns, the default
            @test XLSX._ref_levels("Data!\$A\$2:\$B\$7", "row") == 6      # rows, for dir="row"
            @test XLSX._ref_levels("Data!\$C\$2", "") == 1                # a single cell
        end
    end
    @testset "chartEx chart with direct references (as addChartEx writes)" begin
        xf = XLSX.newxlsx("Data")
        ws = xf["Data"]
        rows = [("Region", "Country", "Sales"),
            ("Europe", "France", 30), ("Europe", "Germany", 45), ("Europe", "Spain", 20),
            ("Asia", "Japan", 50), ("Asia", "India", 35),
            ("Americas", "USA", 60), ("Americas", "Brazil", 25)]
        for (i, r) in enumerate(rows), (j, v) in enumerate(r)
            ws[i, j] = v
        end
        XLSX.Charts.addChartEx(xf, :treemap, "Data!C2:C8"; categories="Data!A2:B8", name_ref="Data!C1")
        XLSX.addsheet!(xf, "Keep")
        XLSX.deletesheet!(xf, "Data")

        c = only(XLSX.getCharts(xf))
        root = XLSX.Charts.chart_root(c)
        @test isempty(_chart_f_texts(root))
        @test length(_chart_elements(only(_chart_elements(root, "strDim")), "lvl")) == 2
        @test length(_chart_elements(only(_chart_elements(root, "numDim")), "lvl")) == 1
        @test isnothing(XLSX.first_element_with_tag(only(_chart_elements(root, "series")), "tx"))
        @test all(isnothing, XLSX.getChartRanges(c))
        SAVE_FILES && save_outfile(xf)
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
            wbxml = String(zip_readentry(r, "xl/workbook.xml"))
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
            SAVE_FILES && save_outfile(p)

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
            SAVE_FILES && save_outfile(p)

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
            SAVE_FILES && save_outfile(p)

            zipbytes = read(p)
            r = ZipReader(zipbytes)
            names = zip_names(r)

            source_path = sheet_xml_path(zipbytes, "Sheet1")
            copy_path = sheet_xml_path(zipbytes, "Copy")
            source_xml = String(zip_readentry(r, source_path))
            copy_xml = String(zip_readentry(r, copy_path))

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
    @testset "deletesheet! removes chart parts (issue #427)" begin
        f = XLSX.opentemplate(joinpath(data_directory, "chart_basic.xlsx"))
        XLSX.addsheet!(f, "Keep")
        XLSX.deletesheet!(f, "Data")

        @test isempty(XLSX.Charts.getCharts(f))
        @test isempty(XLSX.Charts.chart_parts(f))
        @test !haskey(f.data, "xl/charts/chart1.xml")
        @test !haskey(f.data, "xl/drawings/drawing1.xml")
        SAVE_FILES && save_outfile(f)
    end

    @testset "deletesheet! with a chartEx chart" begin
        f = XLSX.opentemplate(joinpath(data_directory, "chart_ex.xlsx"))
        XLSX.addsheet!(f, "Keep")

        XLSX.deletesheet!(f, "Data")
        @test !XLSX.hassheet(f, "Data")
        @test isempty(XLSX.Charts.getCharts(f))

        # The chartEx cluster and the drawing must all go.
        for p in ("xl/charts/chartEx1.xml", "xl/charts/style1.xml",
            "xl/charts/colors1.xml", "xl/drawings/drawing1.xml")
            @test !haskey(f.data, p)
        end
        @test isempty(XLSX.Charts.chartex_parts(f))

        # And the file must still be writable and readable.
        g = XLSX.writexlsx("mytest.xlsx", f, overwrite=true)          # your save_outfile helper
        h = XLSX.readxlsx(g)
        @test isempty(XLSX.Charts.getCharts(h))
        SAVE_FILES && save_outfile(f)
    end

    @testset "copysheet! clones and repoints c: chart parts" begin
        f = XLSX.opentemplate(joinpath(data_directory, "chart_basic.xlsx"))
        XLSX.copysheet!(f["Data"], "Copy")

        charts = XLSX.Charts.getCharts(f)
        @test length(charts) == 2
        @test allunique(c.path for c in charts)
        @test issetequal((c.sheet for c in charts), ["Data", "Copy"])

        cp = only(filter(c -> c.sheet == "Copy", charts))
        orig = only(filter(c -> c.sheet == "Data", charts))

        # The copy plots the copied sheet. getChartRanges returns one named
        # tuple per series for a `c:` chart, so the ranges are in its fields.
        cp_ranges = XLSX.Charts.getChartRanges(cp)
        @test [string(r.categories) for r in cp_ranges] == ["Copy!A2:A5", "Copy!A2:A5"]
        @test [string(r.values) for r in cp_ranges] == ["Copy!B2:B5", "Copy!C2:C5"]

        # Series name refs are repointed too, not just categories and values.
        @test all(s -> startswith(s.name_ref.ref, "Copy!"), XLSX.Charts.getChartSeries(cp))

        # The original is untouched — catches copynode aliasing the two trees.
        orig_ranges = XLSX.Charts.getChartRanges(orig)
        @test [string(r.categories) for r in orig_ranges] == ["Data!A2:A5", "Data!A2:A5"]
        @test [string(r.values) for r in orig_ranges] == ["Data!B2:B5", "Data!C2:C5"]
        @test all(s -> startswith(s.name_ref.ref, "Data!"), XLSX.Charts.getChartSeries(orig))

        # The cache is deliberately left alone, so the copy still reads.
        @test XLSX.Charts.getChartData(cp) isa XLSX.DataTable

        # And it survives a round trip.
        g = XLSX.writexlsx("mytest.xlsx", f; overwrite=true)
        h = XLSX.readxlsx(g)
        cp2 = only(filter(c -> c.sheet == "Copy", XLSX.Charts.getCharts(h)))
        @test [string(r.values) for r in XLSX.Charts.getChartRanges(cp2)] ==
            ["Copy!B2:B5", "Copy!C2:C5"]
        SAVE_FILES && save_outfile(f)
    end
    @testset "copysheet! with a chartEx chart" begin
        f = XLSX.opentemplate(joinpath(data_directory, "chart_ex.xlsx"))
        XLSX.copysheet!(f["Data"], "Copy")

        charts = XLSX.Charts.getCharts(f)
        @test length(charts) == 2
        @test all(c -> c isa XLSX.Charts.ChartEx, charts)
        @test issetequal((c.sheet for c in charts), ["Data", "Copy"])
        @test allunique(c.path for c in charts)
        @test all(c -> XLSX.Charts.getChartType(c) === :waterfall, charts)

        cp = only(filter(c -> c.sheet == "Copy", charts))

        # The copy gets its own hidden defined names, pointing at the copy.
        @test XLSX.Charts._cx_refs(cp) == ["_xlchart.v2.0", "_xlchart.v2.1"]
        @test all(r -> startswith(string(r), "Copy!"), XLSX.Charts.getChartRanges(cp))

        # The original's names and ranges are untouched.
        orig = only(filter(c -> c.sheet == "Data", charts))
        @test XLSX.Charts._cx_refs(orig) == ["_xlchart.v1.0", "_xlchart.v1.1"]
        @test all(r -> startswith(string(r), "Data!"), XLSX.Charts.getChartRanges(orig))

        # Two names added, all still hidden and none user-visible.
        all_names = XLSX.getAllDefinedNames(f; include_system=true)
        @test length(all_names) == 11
        @test all(dn -> dn.hidden, all_names)
        @test isempty(XLSX.getDefinedNames(f))

        g = XLSX.writexlsx("mytest.xlsx", f; overwrite=true)
        h = XLSX.readxlsx(g)
        @test length(XLSX.Charts.getCharts(h)) == 2
        @test allunique(c.path for c in XLSX.Charts.getCharts(h))
        SAVE_FILES && save_outfile(f)
    end

    @testset "copysheet! keeps a shared chartEx name shared" begin
        xf = XLSX.openxlsx(joinpath(data_directory, "chartex_layouts.xlsx"); mode="rw")
        XLSX.copysheet!(xf["pareto"], "pareto2")
        XLSX.writexlsx("mytest.xlsx", xf, overwrite=true)
        SAVE_FILES && save_outfile("mytest.xlsx")

        f = XLSX.readxlsx("mytest.xlsx")
        cxs = filter(x -> x isa XLSX.Charts.ChartEx, XLSX.Charts.getCharts(f))
        orig = only(filter(x -> x.sheet == "pareto", cxs))
        copy = only(filter(x -> x.sheet == "pareto2", cxs))

        ob = XLSX.Charts.getChartDataBlocks(orig)
        cb = XLSX.Charts.getChartDataBlocks(copy)

        # the original is untouched: its two blocks still share one name
        @test ob[1].dimensions[1].formula == ob[2].dimensions[1].formula == "_xlchart.v1.18"

        # the copy's two blocks share one name too, and it is a new one
        shared = cb[1].dimensions[1].formula
        @test cb[2].dimensions[1].formula == shared
        @test shared != "_xlchart.v1.18"

        # every data dimension and series name in the copy points at the copied sheet
        for d in vcat(cb[1].dimensions, cb[2].dimensions)
            @test d.range.sheet == "pareto2"
        end
        @test XLSX.Charts.getSeriesNameRange(copy, 1).sheet == "pareto2"
        @test XLSX.Charts.getSeriesNameRange(copy, 2).sheet == "pareto2"

        # data: cat (shared), val, val; series names: two. Five names, not six.
        @test length(unique(XLSX.Charts._cx_refs(copy))) == 3                      # data dimensions
        @test length(XLSX.Charts._cx_refs(copy)) == 4                              # cat appears twice
        isfile("mytest.xlsx") && rm("mytest.xlsx")
    end

    @testset "copysheet! clones chart parts for both schemas" begin
        f = XLSX.opentemplate(joinpath(data_directory, "chart_mixed.xlsx"))
        XLSX.copysheet!(f["Data"], "Copy")

        charts = XLSX.Charts.getCharts(f)
        @test length(charts) == 4
        @test allunique(c.path for c in charts)
        @test count(c -> c.sheet == "Copy", charts) == 2
        @test issetequal((XLSX.Charts.getChartType(c) for c in charts if c.sheet == "Copy"),
            [:waterfall, :barChart])

        # Stems are handled per family: chart1 -> chart2, chartEx1 -> chartEx2.
        @test issetequal((c.name for c in charts),
            ["chart1", "chart2", "chartEx1", "chartEx2"])

        # Anchors come across with the drawing.
        @test issetequal((c.from for c in charts), ["F1", "F17"])

        # The c: chart is repointed at the copy; the cx: one is not yet.
        bar = only(filter(c -> c.sheet == "Copy" && c isa XLSX.Charts.Chart, charts))
        @test all(r -> startswith(string(r.values), "Copy!"), XLSX.Charts.getChartRanges(bar))

        wf = only(filter(c -> c.sheet == "Copy" && c isa XLSX.Charts.ChartEx, charts))
        @test all(r -> startswith(string(r), "Copy!"), XLSX.Charts.getChartRanges(wf))

        g = XLSX.writexlsx("mytest.xlsx", f; overwrite=true)
        h = XLSX.readxlsx(g)
        @test length(XLSX.Charts.getCharts(h)) == 4
        @test allunique(c.path for c in XLSX.Charts.getCharts(h))
        SAVE_FILES && save_outfile(f)
    end

    isfile("mytest.xlsx") && rm("mytest.xlsx")

    @testset "renaming a sheet repoints chart references" begin
        path = "chart_rename.xlsx"
        cp(joinpath(data_directory, "chart_basic.xlsx"), path; force=true)
        xf = XLSX.openxlsx(path; mode="rw")
        XLSX.renamesheet!(xf["Data"], "Sales")
        c = only(XLSX.Charts.getCharts(xf))
        @test all(r -> startswith(string(r.values), "Sales!"), XLSX.Charts.getChartRanges(c))
        SAVE_FILES && save_outfile(xf)
        isfile(path) && rm(path)
    end

    @testset "chart reference repointing matches whole sheet names" begin
        @test XLSX._repoint_ref_string("Data!\$A\$1:\$A\$5", "Data!", "Sales!") == "Sales!\$A\$1:\$A\$5"
        @test XLSX._repoint_ref_string("MyData!\$A\$1", "Data!", "Sales!") == "MyData!\$A\$1"
        @test XLSX._repoint_ref_string("(Data!\$A\$1,Data!\$C\$1)", "Data!", "Sales!") == "(Sales!\$A\$1,Sales!\$C\$1)"
        @test XLSX._repoint_ref_string("'My Sheet'!\$B\$2", "'My Sheet'!", "Sales!") == "Sales!\$B\$2"
        @test XLSX._repoint_ref_string("[1]Data!\$A\$1", "Data!", "Sales!") == "[1]Data!\$A\$1"
    end
end
