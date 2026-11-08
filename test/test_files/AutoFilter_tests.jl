@testset "AutoFilter" begin
    function fresh()
        xf = XLSX.newxlsx("Data")
        s = xf["Data"]
        s["A1"] = ["Name" "Score" "Group"; "a" 1 "x"; "b" 2 "y"; "c" 3 "x"]
        return xf, s
    end
    sheet_root(xf, n=1) = XLSX.xml_root_element(xf.data["xl/worksheets/sheet$n.xml"])
    sheet_tags(xf, n=1) = [XLSX.localname(c) for c in XLSX.xml_elements(sheet_root(xf, n))]
    af_node(xf, n=1) = XLSX.first_element_with_tag(sheet_root(xf, n), "autoFilter")
    filter_name(xf, s) = get(XLSX.get_workbook(xf).worksheet_names, (s.sheetId, "_xlnm._FilterDatabase"), nothing)
    function roundtrip_bytes(xf)
        io = IOBuffer()
        XLSX.writexlsx(io, xf)
        return take!(io)
    end
    workbook_xml(bytes) = XLSX.ZipArchives.zip_readentry(XLSX.ZipArchives.ZipReader(bytes), "xl/workbook.xml", String)

    @testset "set, get and remove" begin
        xf, s = fresh()
        @test isnothing(XLSX.getAutoFilter(s))
        @test isnothing(XLSX.setAutoFilter(s, "A1:C4"))
        @test XLSX.getAutoFilter(s) == XLSX.CellRange("A1:C4")
        @test af_node(xf)["ref"] == "A1:C4"
        tags = sheet_tags(xf)
        @test findfirst(==("autoFilter"), tags) > findfirst(==("sheetData"), tags)
        @test findfirst(==("autoFilter"), tags) < findfirst(==("pageMargins"), tags)
        dn = filter_name(xf, s)
        @test dn.hidden && dn.value == XLSX.SheetCellRange("Data!A1:C4")

        bytes = roundtrip_bytes(xf)
        @test occursin(r"<definedName name=\"_xlnm._FilterDatabase\" localSheetId=\"0\" hidden=\"1\">Data!\$A\$1:\$C\$4</definedName>", workbook_xml(bytes))
        r = XLSX.readxlsx(IOBuffer(bytes))
        @test XLSX.getAutoFilter(r["Data"]) == XLSX.CellRange("A1:C4")
        @test isempty(XLSX.getDefinedNames(r["Data"]))                  # hidden from the user's names

        @test XLSX.removeAutoFilter(s)
        @test isnothing(XLSX.getAutoFilter(s))
        @test isnothing(filter_name(xf, s))
        @test !XLSX.removeAutoFilter(s)
        @test !occursin("_FilterDatabase", workbook_xml(roundtrip_bytes(xf)))
        SAVE_FILES && (XLSX.setAutoFilter(s, :); save_outfile(xf))
    end

    @testset "range forms" begin
        xf, s = fresh()
        XLSX.setAutoFilter(s, :)
        @test XLSX.getAutoFilter(s) == XLSX.CellRange("A1:C4")
        XLSX.setAutoFilter(s, XLSX.CellRange("A1:B4"))
        @test XLSX.getAutoFilter(s) == XLSX.CellRange("A1:B4")
        @test count(==("autoFilter"), sheet_tags(xf)) == 1                 # replaced, not added
        XLSX.setAutoFilter(xf, "Data!A1:C3")
        @test XLSX.getAutoFilter(s) == XLSX.CellRange("A1:C3")
        XLSX.setAutoFilter(s, XLSX.SheetCellRange("Data!A1:C2"))
        @test XLSX.getAutoFilter(s) == XLSX.CellRange("A1:C2")
        XLSX.addDefinedName(s, "Scores", "A1:C4")
        XLSX.setAutoFilter(s, "Scores")
        @test XLSX.getAutoFilter(s) == XLSX.CellRange("A1:C4")
        @test filter_name(xf, s).value == XLSX.SheetCellRange("Data!A1:C4")

        @test_throws XLSX.XLSXError XLSX.setAutoFilter(s, "A1:B2,D1:E2")
        @test_throws XLSX.XLSXError XLSX.setAutoFilter(s, "B:C")
        @test_throws XLSX.XLSXError XLSX.setAutoFilter(s, "nonsense!")
        @test_throws XLSX.XLSXError XLSX.setAutoFilter(xf, "Nope!A1:B2")
        XLSX.addsheet!(xf, "Other")
        @test_throws XLSX.XLSXError XLSX.setAutoFilter(s, XLSX.SheetCellRange("Other!A1:B2"))
    end

    @testset "criteria are kept for the same range" begin
        xf, s = fresh()
        XLSX.setAutoFilter(s, "A1:C4")
        fc = XML.Element("filterColumn"; colId="2")
        push!(fc, XML.Element("filters", XML.Element("filter"; val="x")))
        push!(af_node(xf), fc)
        XLSX.setAutoFilter(s, "A1:C4")
        @test length(XLSX.xml_elements(af_node(xf))) == 1
        XLSX.setAutoFilter(s, "A1:C3")                                    # a new range starts afresh
        @test isempty(XLSX.xml_elements(af_node(xf)))
    end

    @testset "tables" begin
        xf, s = fresh()
        s["F1"] = ["H1" "H2"; 1 2; 3 4]
        XLSX.addtable!(s, "F1:G3")
        @test_throws XLSX.XLSXError XLSX.setAutoFilter(s, "A1:G4")        # overlaps the table
        XLSX.setAutoFilter(s, "A1:C4")
        @test XLSX.getAutoFilter(s) == XLSX.CellRange("A1:C4")
        r = XLSX.readxlsx(IOBuffer(roundtrip_bytes(xf)))["Data"]
        @test XLSX.getAutoFilter(r) == XLSX.CellRange("A1:C4")
        @test only(XLSX.tables(r)).ref == XLSX.CellRange("F1:G3")

        # ... and the other way: a table may not be placed on, or grow into, the filter
        @test_throws XLSX.XLSXError XLSX.addtable!(s, "B1:C4")
        @test_throws XLSX.XLSXError XLSX.writetable!(s, [[1, 2]], ["N"]; anchor_cell=XLSX.CellRef("C3"), as_table=true)
        @test length(XLSX.tables(s)) == 1
        @test s["C3"] == "y"                                              # refused before writing

        xf, s = fresh()
        s["A10"] = ["H1" "H2"; 1 2]
        XLSX.addtable!(s, "A10:B11"; name="Below")
        XLSX.setAutoFilter(s, "A12:B14")                                   # just below the table
        @test_throws XLSX.XLSXError XLSX.appendtable!(s, "Below", [[3, 4]]; check_empty=false)
        @test_throws XLSX.XLSXError XLSX.settotals!(s, "Below", "H2" => :sum)
        @test XLSX.table(s, "Below").ref == XLSX.CellRange("A10:B11")      # unchanged
        @test ismissing(s["A12"])
        XLSX.removeAutoFilter(s)
        XLSX.appendtable!(s, "Below", [[3, 4]])
        @test XLSX.table(s, "Below").ref == XLSX.CellRange("A10:B12")
    end

    @testset "file modes" begin
        xf, s = fresh()
        XLSX.setAutoFilter(s, "A1:C4")
        bytes = roundtrip_bytes(xf)
        r = XLSX.readxlsx(IOBuffer(bytes))["Data"]
        @test_throws XLSX.XLSXError XLSX.setAutoFilter(r, "A1:B2")
        @test_throws XLSX.XLSXError XLSX.removeAutoFilter(r)
        XLSX.openxlsx(IOBuffer(bytes); enable_cache=false) do f
            @test XLSX.getAutoFilter(f["Data"]) == XLSX.CellRange("A1:C4")
        end
        f = XLSX.openxlsx(IOBuffer(bytes); mode="rw")
        @test XLSX.removeAutoFilter(f["Data"])
        @test !occursin("_FilterDatabase", workbook_xml(roundtrip_bytes(f)))
    end

    @testset "sheet operations" begin
        xf, s = fresh()
        XLSX.setAutoFilter(s, "A1:C4")
        c = XLSX.copysheet!(s, "Copy")
        @test XLSX.getAutoFilter(c) == XLSX.CellRange("A1:C4")
        @test filter_name(xf, c).value == XLSX.SheetCellRange("Copy!A1:C4")
        XLSX.renamesheet!(s, "Renamed")
        @test filter_name(xf, s).value == XLSX.SheetCellRange("Renamed!A1:C4")
        w = workbook_xml(roundtrip_bytes(xf))
        @test occursin("localSheetId=\"0\" hidden=\"1\">Renamed!\$A\$1:\$C\$4<", w)
        @test occursin("localSheetId=\"1\" hidden=\"1\">Copy!\$A\$1:\$C\$4<", w)
        XLSX.deletesheet!(xf, "Renamed")
        @test isnothing(filter_name(xf, s))
        w = workbook_xml(roundtrip_bytes(xf))
        @test occursin("localSheetId=\"0\" hidden=\"1\">Copy!\$A\$1:\$C\$4<", w)
        @test count("_FilterDatabase", w) == 1
    end

    @testset "chartsheets" begin
        xf = XLSX.newxlsx("data")
        xf["data"]["A1:A2"] = [1, 2]
        c = XLSX.Charts.addChart(xf, :line; sheetname="Trend")
        XLSX.Charts.addSeries(c, "data!A1:A2")
        @test isnothing(XLSX.getAutoFilter(xf["Trend"]))
        @test_throws XLSX.XLSXError XLSX.setAutoFilter(xf["Trend"], "A1:A2")
        @test_throws XLSX.XLSXError XLSX.removeAutoFilter(xf["Trend"])
    end
end
