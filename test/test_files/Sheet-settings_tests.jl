@testset "Sheet settings" begin
    function roundtrip_bytes(xf)
        io = IOBuffer()
        XLSX.writexlsx(io, xf)
        return take!(io)
    end
    entry(bytes, name) = XLSX.ZipArchives.zip_readentry(XLSX.ZipArchives.ZipReader(bytes), name, String)
    states(bytes) = Dict(m[1] => m[2] for m in eachmatch(r"<sheet [^>]*name=\"([^\"]+)\"[^>]*state=\"([^\"]+)\"", entry(bytes, "xl/workbook.xml")))
    active_tab(bytes) = (m = match(r"activeTab=\"(\d+)\"", entry(bytes, "xl/workbook.xml")); isnothing(m) ? 0 : parse(Int, m[1]))
    sheet_root(xf, path) = XLSX.xml_root_element(XLSX.get_xml_data(xf, path))
    tab_selected(xf, n) = (v = XLSX.first_element_with_tag(XLSX.first_element_with_tag(sheet_root(xf, "xl/worksheets/sheet$n.xml"), "sheetViews"), "sheetView");
                           get(v, "tabSelected", "0") == "1")
    function three()
        xf = XLSX.newxlsx("A")
        XLSX.addsheet!(xf, "B")
        XLSX.addsheet!(xf, "C")
        return xf
    end

    @testset "visibility" begin
        xf = three()
        @test XLSX.getSheetVisibility(xf["B"]) == :visible
        @test isnothing(XLSX.setSheetVisibility(xf["B"], :hidden))
        XLSX.setSheetVisibility(xf["C"], :veryHidden)
        @test XLSX.getSheetVisibility(xf["B"]) == :hidden
        @test XLSX.getSheetVisibility(xf["C"]) == :veryHidden
        @test occursin("(very hidden)", sprint(show, xf["C"]))
        bytes = roundtrip_bytes(xf)
        @test states(bytes) == Dict("B" => "hidden", "C" => "veryHidden")
        r = XLSX.readxlsx(IOBuffer(bytes))
        @test [XLSX.getSheetVisibility(s) for s in r.workbook.sheets] == [:visible, :hidden, :veryHidden]

        XLSX.setSheetVisibility(xf["C"], :visible)
        @test states(roundtrip_bytes(xf)) == Dict("B" => "hidden")
        @test_throws XLSX.XLSXError XLSX.setSheetVisibility(xf["B"], :shown)
        @test_throws XLSX.XLSXError XLSX.setSheetVisibility(r["A"], :hidden)   # read-only
        SAVE_FILES && save_outfile(xf)
    end

    @testset "the last visible sheet" begin
        xf = three()
        XLSX.setSheetVisibility(xf["A"], :hidden)
        XLSX.setSheetVisibility(xf["B"], :veryHidden)
        @test_throws XLSX.XLSXError XLSX.setSheetVisibility(xf["C"], :hidden)
        @test XLSX.getSheetVisibility(xf["C"]) == :visible
        XLSX.setSheetVisibility(xf["C"], :visible)                            # no change: fine
        XLSX.setSheetVisibility(xf["B"], :hidden)                             # hidden to hidden: fine
        @test XLSX.getSheetVisibility(xf["B"]) == :hidden
    end

    @testset "the active sheet moves off a hidden one" begin
        xf = three()
        @test active_tab(roundtrip_bytes(xf)) == 0
        @test tab_selected(xf, 1)
        XLSX.setSheetVisibility(xf["A"], :hidden)
        @test active_tab(roundtrip_bytes(xf)) == 1
        @test !tab_selected(xf, 1) && tab_selected(xf, 2)
        XLSX.setSheetVisibility(xf["C"], :hidden)                             # not the active sheet
        @test active_tab(roundtrip_bytes(xf)) == 1
        @test tab_selected(xf, 2)
    end

    @testset "is_hidden still works" begin
        xf = three()
        s = xf["B"]
        @test !s.is_hidden
        s.is_hidden = true
        @test s.is_hidden && XLSX.getSheetVisibility(s) == :hidden
        XLSX.setSheetVisibility(s, :veryHidden)
        s.is_hidden = true                                                    # stays very hidden
        @test XLSX.getSheetVisibility(s) == :veryHidden
        s.is_hidden = false
        @test XLSX.getSheetVisibility(s) == :visible
    end

    @testset "tab colour" begin
        xf = three()
        s = xf["A"]
        @test isnothing(XLSX.getTabColor(s))
        @test isnothing(XLSX.setTabColor(s, "FF0000"))
        @test XLSX.getTabColor(s) == "FFFF0000"
        root = sheet_root(xf, "xl/worksheets/sheet1.xml")
        @test XLSX.localname(first(XLSX.xml_elements(root))) == "sheetPr"   # first, as the schema puts it

        XLSX.setTabColor(s, :steelblue)
        @test XLSX.getTabColor(s) == "FF4682B4"
        pr = XLSX.first_element_with_tag(sheet_root(xf, "xl/worksheets/sheet1.xml"), "sheetPr")
        @test length(XLSX.xml_elements(pr)) == 1                              # replaced, not added

        XLSX.setTabColor(xf["B"], :accent2)
        el = only(XLSX.xml_elements(XLSX.first_element_with_tag(sheet_root(xf, "xl/worksheets/sheet2.xml"), "sheetPr")))
        @test el["theme"] == "5"
        @test XLSX.getTabColor(xf["B"]) isa String

        @test_throws XLSX.XLSXError XLSX.setTabColor(s, "not a colour")
        @test XLSX.getTabColor(s) == "FF4682B4"

        bytes = roundtrip_bytes(xf)
        r = XLSX.readxlsx(IOBuffer(bytes))
        @test XLSX.getTabColor(r["A"]) == "FF4682B4"
        @test isnothing(XLSX.getTabColor(r["C"]))
        @test_throws XLSX.XLSXError XLSX.setTabColor(r["C"], "red")
        XLSX.openxlsx(IOBuffer(bytes); enable_cache=false) do f
            @test XLSX.getTabColor(f["A"]) == "FF4682B4"
        end

        @test XLSX.removeTabColor(s)
        @test isnothing(XLSX.getTabColor(s))
        @test isnothing(XLSX.first_element_with_tag(sheet_root(xf, "xl/worksheets/sheet1.xml"), "sheetPr"))   # emptied, so removed
        @test !XLSX.removeTabColor(s)
        SAVE_FILES && save_outfile(xf)
    end

    @testset "tab colour keeps other sheet properties" begin
        xf = three()
        s = xf["A"]
        XLSX.groupRows(s, 2:3)
        root = sheet_root(xf, "xl/worksheets/sheet1.xml")
        pr = XML.Element("sheetPr", XML.Element("outlinePr"; summaryBelow="0"))
        pushfirst!(root.children, pr)
        XLSX.setTabColor(s, "00FF00")
        pr = XLSX.first_element_with_tag(sheet_root(xf, "xl/worksheets/sheet1.xml"), "sheetPr")
        @test [XLSX.localname(c) for c in XLSX.xml_elements(pr)] == ["tabColor", "outlinePr"]
        XLSX.removeTabColor(s)
        pr = XLSX.first_element_with_tag(sheet_root(xf, "xl/worksheets/sheet1.xml"), "sheetPr")
        @test [XLSX.localname(c) for c in XLSX.xml_elements(pr)] == ["outlinePr"]
    end

    @testset "copied sheets" begin
        xf = three()
        XLSX.setTabColor(xf["A"], "red")
        c = XLSX.copysheet!(xf["A"], "Copy")
        @test XLSX.getTabColor(c) == "FFFF0000"
    end

    @testset "chartsheets" begin
        xf = XLSX.newxlsx("data")
        xf["data"]["A1:A2"] = [1, 2]
        c = XLSX.Charts.addChart(xf, :line; sheetname="Trend")
        XLSX.Charts.addSeries(c, "data!A1:A2")
        chs = xf["Trend"]
        XLSX.setTabColor(chs, "0000FF")
        @test XLSX.getTabColor(chs) == "FF0000FF"
        XLSX.setSheetVisibility(chs, :hidden)
        bytes = roundtrip_bytes(xf)
        r = XLSX.readxlsx(IOBuffer(bytes))
        @test XLSX.getTabColor(r["Trend"]) == "FF0000FF"
        @test XLSX.getSheetVisibility(r["Trend"]) == :hidden
        @test XLSX.removeTabColor(chs)
        @test isnothing(XLSX.getTabColor(chs))
    end
end
