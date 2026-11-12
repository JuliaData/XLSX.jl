@testset "sheet settings" begin
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

    sheet_xml(bytes, n=1) = entry(bytes, "xl/worksheets/sheet$n.xml")
    element_atts(xf, name, n=1) = (e = XLSX.first_element_with_tag(sheet_root(xf, "xl/worksheets/sheet$n.xml"), name);
                                   isnothing(e) ? nothing : Dict(XML.attributes(e)))
    sheet_tags(xf, n=1) = [XLSX.localname(c) for c in XLSX.xml_elements(sheet_root(xf, "xl/worksheets/sheet$n.xml"))]

    @testset "protection" begin
        xf = three()
        s = xf["A"]
        @test !XLSX.isSheetProtected(s)
        @test isnothing(XLSX.getSheetProtection(s))
        @test isnothing(XLSX.protectSheet(s))
        @test XLSX.isSheetProtected(s)
        @test element_atts(xf, "sheetProtection") == Dict("sheet" => "1", "objects" => "1", "scenarios" => "1")   # as Excel writes it
        p = XLSX.getSheetProtection(s)
        @test !p.password && p.selectLockedCells && p.selectUnlockedCells
        @test !p.formatCells && !p.sort && !p.objects

        XLSX.protectSheet(s; password="secret", sort=true, autoFilter=true, selectLockedCells=false)
        a = element_atts(xf, "sheetProtection")
        @test a["password"] == "DAA7"                                         # Excel's legacy hash of "secret"
        @test a["sort"] == "0" && a["autoFilter"] == "0" && a["selectLockedCells"] == "1"
        p = XLSX.getSheetProtection(s)
        @test p.password && p.sort && p.autoFilter && !p.selectLockedCells && p.selectUnlockedCells && !p.insertRows
        tags = sheet_tags(xf)
        @test findfirst(==("sheetProtection"), tags) == findfirst(==("sheetData"), tags) + 1

        r = XLSX.readxlsx(IOBuffer(roundtrip_bytes(xf)))
        @test XLSX.getSheetProtection(r["A"]) == p
        @test !XLSX.isSheetProtected(r["B"])
        @test_throws XLSX.XLSXError XLSX.protectSheet(r["B"])

        @test XLSX.unprotectSheet(s)
        @test !XLSX.isSheetProtected(s) && isnothing(element_atts(xf, "sheetProtection"))
        @test !XLSX.unprotectSheet(s)

        @test_throws XLSX.XLSXError XLSX.protectSheet(s; formatcells=true)    # misspelt
        @test_throws XLSX.XLSXError XLSX.protectSheet(s; sort=1)
        @test_throws XLSX.XLSXError XLSX.protectSheet(s; password="")
        @test_throws XLSX.XLSXError XLSX.protectSheet(s; password="pässword")
        @test_throws XLSX.XLSXError XLSX.protectSheet(s; password="x"^16)
        @test !XLSX.isSheetProtected(s)                                       # nothing changed
        SAVE_FILES && (XLSX.protectSheet(s; password="secret"); save_outfile(xf))
    end

    @testset "password hash" begin
        @test XLSX._excel_password_hash("secret") == "DAA7"
        @test length(XLSX._excel_password_hash("x"^15)) == 4
        # openpyxl's form of the hash, the same for ASCII passwords of up to 15 characters
        ref(p) = (h = 0; for (i, c) in enumerate(p); v = Int(c) << i; h ⊻= (v & 0x7fff) | (v >> 15); end;
                  uppercase(string(h ⊻ length(p) ⊻ 0xCE4B; base=16, pad=4)))
        @test all(XLSX._excel_password_hash(p) == ref(p) for p in ("secret", "abcdefghijklmno", "~~~~~~~~~~~~~~~", "Pa55 w0rd!"))
    end

    @testset "page setup" begin
        xf = three()
        s = xf["A"]
        @test XLSX.getPageSetup(s) == (orientation=nothing, paperSize=nothing, scale=nothing, fitToWidth=nothing, fitToHeight=nothing)
        XLSX.setPageSetup(s; orientation=:landscape, paperSize=:A4, scale=80)
        @test XLSX.getPageSetup(s) == (orientation=:landscape, paperSize=9, scale=80, fitToWidth=nothing, fitToHeight=nothing)
        tags = sheet_tags(xf)
        @test findfirst(==("pageSetup"), tags) == findfirst(==("pageMargins"), tags) + 1

        XLSX.setPageSetup(s; fitToWidth=1)                                   # turns scaling off
        @test XLSX.getPageSetup(s) == (orientation=:landscape, paperSize=9, scale=nothing, fitToWidth=1, fitToHeight=0)
        pr = XLSX.first_element_with_tag(sheet_root(xf, "xl/worksheets/sheet1.xml"), "sheetPr")
        @test only(XLSX.xml_elements(pr))["fitToPage"] == "1"
        XLSX.setTabColor(s, "red")                                            # tabColor goes before pageSetUpPr
        pr = XLSX.first_element_with_tag(sheet_root(xf, "xl/worksheets/sheet1.xml"), "sheetPr")
        @test [XLSX.localname(c) for c in XLSX.xml_elements(pr)] == ["tabColor", "pageSetUpPr"]

        XLSX.setPageSetup(s; scale=50)                                        # and back
        @test XLSX.getPageSetup(s).scale == 50 && isnothing(XLSX.getPageSetup(s).fitToWidth)
        pr = XLSX.first_element_with_tag(sheet_root(xf, "xl/worksheets/sheet1.xml"), "sheetPr")
        @test [XLSX.localname(c) for c in XLSX.xml_elements(pr)] == ["tabColor"]
        XLSX.setPageSetup(s; paperSize=70)                                    # any number Excel knows
        @test XLSX.getPageSetup(s).paperSize == 70

        @test_throws XLSX.XLSXError XLSX.setPageSetup(s; orientation=:sideways)
        @test_throws XLSX.XLSXError XLSX.setPageSetup(s; paperSize=:A9)
        @test_throws XLSX.XLSXError XLSX.setPageSetup(s; paperSize=0)
        @test_throws XLSX.XLSXError XLSX.setPageSetup(s; scale=5)
        @test_throws XLSX.XLSXError XLSX.setPageSetup(s; scale=50, fitToWidth=1)
        @test_throws XLSX.XLSXError XLSX.setPageSetup(s; fitToHeight=-1)

        r = XLSX.readxlsx(IOBuffer(roundtrip_bytes(xf)))
        @test XLSX.getPageSetup(r["A"]) == (orientation=:landscape, paperSize=70, scale=50, fitToWidth=nothing, fitToHeight=nothing)
    end

    @testset "page setup keeps the printer settings link" begin
        xf = three()
        s = xf["A"]
        root = sheet_root(xf, "xl/worksheets/sheet1.xml")
        push!(root.children, XML.Element("pageSetup"; paperSize="9", var"r:id"="rId1"))
        XLSX.update_worksheets_xml!(xf)
        XLSX.setPageSetup(s; orientation=:portrait)
        a = element_atts(xf, "pageSetup")
        @test a == Dict("paperSize" => "9", "r:id" => "rId1", "orientation" => "portrait")
    end

    @testset "margins and print options" begin
        xf = three()
        s = xf["A"]
        @test XLSX.getPageMargins(s) == (left=0.7, right=0.7, top=0.75, bottom=0.75, header=0.3, footer=0.3)   # the template's
        XLSX.setPageMargins(s; left=0.25, top=1)
        @test XLSX.getPageMargins(s) == (left=0.25, right=0.7, top=1.0, bottom=0.75, header=0.3, footer=0.3)
        @test_throws XLSX.XLSXError XLSX.setPageMargins(s; left=-1)
        @test_throws XLSX.XLSXError XLSX.setPageMargins(s; gutter=1)

        root = sheet_root(xf, "xl/worksheets/sheet2.xml")             # a sheet without margins
        deleteat!(root.children, findfirst(c -> XLSX.localname(c) == "pageMargins", root.children))
        XLSX.update_worksheets_xml!(xf)
        @test isnothing(XLSX.getPageMargins(xf["B"]))
        XLSX.setPageMargins(xf["B"]; bottom=2)
        @test XLSX.getPageMargins(xf["B"]) == (left=0.7, right=0.7, top=0.75, bottom=2.0, header=0.3, footer=0.3)

        @test XLSX.getPrintOptions(s) == (gridLines=false, headings=false, horizontalCentered=false, verticalCentered=false)
        XLSX.setPrintOptions(s; gridLines=true, horizontalCentered=true)
        @test XLSX.getPrintOptions(s) == (gridLines=true, headings=false, horizontalCentered=true, verticalCentered=false)
        tags = sheet_tags(xf)
        @test findfirst(==("printOptions"), tags) == findfirst(==("pageMargins"), tags) - 1
        XLSX.setPrintOptions(s; gridLines=false)
        @test XLSX.getPrintOptions(s).horizontalCentered && !XLSX.getPrintOptions(s).gridLines
        XLSX.setPrintOptions(s; horizontalCentered=false)
        @test isnothing(element_atts(xf, "printOptions"))                     # emptied, so removed
        @test_throws XLSX.XLSXError XLSX.setPrintOptions(s; gridlines=true)
        @test_throws XLSX.XLSXError XLSX.setPrintOptions(s; gridLines=1)

        XLSX.setPrintOptions(s; headings=true)
        r = XLSX.readxlsx(IOBuffer(roundtrip_bytes(xf)))
        @test XLSX.getPageMargins(r["A"]).left == 0.25
        @test XLSX.getPrintOptions(r["A"]).headings
    end

    @testset "print area and titles" begin
        xf = three()
        s = xf["A"]
        @test isnothing(XLSX.getPrintArea(s)) && isnothing(XLSX.getPrintTitles(s))
        XLSX.setPrintArea(s, "A1:F40")
        @test XLSX.getPrintArea(s) == XLSX.CellRange("A1:F40")
        XLSX.setPrintArea(s, "A1:F40,H1:K10")
        @test XLSX.getPrintArea(s) isa XLSX.NonContiguousRange
        XLSX.setPrintArea(s, XLSX.CellRange("B2:C3"))
        @test XLSX.getPrintArea(s) == XLSX.CellRange("B2:C3")
        @test_throws XLSX.XLSXError XLSX.setPrintArea(s, "B:C")

        XLSX.setPrintTitles(s; rows=1)
        @test XLSX.getPrintTitles(s) == (rows=1:1, cols=nothing)
        XLSX.setPrintTitles(s; rows="1:2", cols="A")
        @test XLSX.getPrintTitles(s) == (rows=1:2, cols=1:1)
        @test_throws XLSX.XLSXError XLSX.setPrintTitles(s)
        @test_throws XLSX.XLSXError XLSX.setPrintTitles(s; rows=[1, 3])

        w = entry(roundtrip_bytes(xf), "xl/workbook.xml")
        @test occursin("<definedName name=\"_xlnm.Print_Area\" localSheetId=\"0\">A!\$B\$2:\$C\$3</definedName>", w)
        @test occursin("<definedName name=\"_xlnm.Print_Titles\" localSheetId=\"0\">A!\$A:\$A,A!\$1:\$2</definedName>", w)

        r = XLSX.readxlsx(IOBuffer(roundtrip_bytes(xf)))
        @test XLSX.getPrintArea(r["A"]) == XLSX.CellRange("B2:C3")
        @test XLSX.getPrintTitles(r["A"]) == (rows=1:2, cols=1:1)

        XLSX.renamesheet!(s, "My Sheet")
        @test XLSX.getPrintTitles(s) == (rows=1:2, cols=1:1)
        w = entry(roundtrip_bytes(xf), "xl/workbook.xml")
        @test occursin(">&apos;My Sheet&apos;!\$B\$2:\$C\$3<", w)
        @test occursin(">&apos;My Sheet&apos;!\$A:\$A,&apos;My Sheet&apos;!\$1:\$2<", w)

        c = XLSX.copysheet!(s, "Copy")
        @test XLSX.getPrintArea(c) == XLSX.CellRange("B2:C3")
        @test XLSX.getPrintTitles(c) == (rows=1:2, cols=1:1)
        @test occursin(">Copy!\$A:\$A,Copy!\$1:\$2<", entry(roundtrip_bytes(xf), "xl/workbook.xml"))

        @test XLSX.removePrintArea(s) && XLSX.removePrintTitles(s)
        @test isnothing(XLSX.getPrintArea(s)) && isnothing(XLSX.getPrintTitles(s))
        @test !XLSX.removePrintArea(s) && !XLSX.removePrintTitles(s)
        @test !isnothing(XLSX.getPrintArea(c))                                # the copy keeps its own
        SAVE_FILES && save_outfile(xf)
    end

    @testset "headers and footers" begin
        xf = three()
        s = xf["A"]
        s[1, 1] = 1
        @test isnothing(XLSX.getHeaderFooter(s))
        XLSX.setHeaderFooter(s; header=(left="Left & text", center="Page &P of &N",
                                        right=XLSX.headerFooterFont(name="Arial", bold=true, italic=true, size=14, color="FF0000") * "Red &D &T"),
                                footer="&F &A &Z")
        hf = XLSX.getHeaderFooter(s)
        @test hf.header == (left="Left & text", center="Page &P of &N", right="&\"Arial,Bold Italic\"&14&KFF0000Red &D &T")
        @test hf.footer == (left=nothing, center="&F &A &Z", right=nothing)            # a plain string is centred
        @test all(isnothing, (hf.evenHeader, hf.evenFooter, hf.firstHeader, hf.firstFooter))
        @test (hf.differentOddEven, hf.differentFirst, hf.scaleWithDoc, hf.alignWithMargins) == (false, false, true, true)
        @test isempty(element_atts(xf, "headerFooter"))                                # defaults left unwritten, as Excel does
        tags = sheet_tags(xf)
        @test findfirst(==("headerFooter"), tags) == findfirst(==("pageMargins"), tags) + 1
        XLSX.setPageSetup(s; orientation=:landscape)                                   # pageSetup goes between them
        tags = sheet_tags(xf)
        @test findfirst(==("headerFooter"), tags) == findfirst(==("pageSetup"), tags) + 1

        XLSX.setHeaderFooter(s; evenHeader="Even &P", firstFooter=(right="First",), scaleWithDoc=false, alignWithMargins=false)
        hf = XLSX.getHeaderFooter(s)
        @test hf.differentOddEven && hf.differentFirst && !hf.scaleWithDoc && !hf.alignWithMargins
        @test hf.evenHeader == (left=nothing, center="Even &P", right=nothing)
        @test hf.firstFooter == (left=nothing, center=nothing, right="First")
        @test hf.header.left == "Left & text"                                          # untouched

        # the same XML as Excel writes for the same settings
        ours = XLSX.first_element_with_tag(sheet_root(xf, "xl/worksheets/sheet1.xml"), "headerFooter")
        ex = XLSX.readxlsx(joinpath(data_directory, "headerfooter_excel.xlsx"))
        theirs = only(XLSX._sheet_top_elements(ex, XLSX._sheet_path(ex[1]), ("headerFooter",)))
        @test XLSX._atts_vector(ours) == XLSX._atts_vector(theirs)
        @test XLSX._header_footer_texts(ours) == XLSX._header_footer_texts(theirs)
        @test [XLSX.localname(c) for c in XLSX.xml_elements(ours)] == [XLSX.localname(c) for c in XLSX.xml_elements(theirs)]
        @test XLSX.getHeaderFooter(ex[1]) == hf                                        # read-only, from Excel's file

        # a header read back can be set again as it is
        c = xf["B"]
        XLSX.setHeaderFooter(c; header=hf.header, footer=hf.footer)
        @test XLSX.getHeaderFooter(c).header == hf.header

        XLSX.setHeaderFooter(s; differentFirst=false, firstHeader="Kept")              # an explicit option wins
        @test !XLSX.getHeaderFooter(s).differentFirst && XLSX.getHeaderFooter(s).firstHeader.center == "Kept"
        XLSX.setHeaderFooter(s; evenHeader="")                                         # "" removes one, leaving even pages blank
        hf = XLSX.getHeaderFooter(s)
        @test isnothing(hf.evenHeader) && hf.differentOddEven
        XLSX.setHeaderFooter(s; differentOddEven=false)
        @test !XLSX.getHeaderFooter(s).differentOddEven
        @test !haskey(element_atts(xf, "headerFooter"), "differentOddEven")
        XLSX.setHeaderFooter(c; firstHeader="")                                         # a blank first page
        @test XLSX.getHeaderFooter(c).differentFirst && isnothing(XLSX.getHeaderFooter(c).firstHeader)

        r = XLSX.readxlsx(IOBuffer(roundtrip_bytes(xf)))
        @test XLSX.getHeaderFooter(r["A"]) == XLSX.getHeaderFooter(s)
        @test XLSX.getHeaderFooter(XLSX.copysheet!(s, "Copy")) == XLSX.getHeaderFooter(s)
        SAVE_FILES && save_outfile(xf)

        @test XLSX.removeHeaderFooter(s)
        @test isnothing(XLSX.getHeaderFooter(s)) && !("headerFooter" in sheet_tags(xf))
        @test !XLSX.removeHeaderFooter(s)
        XLSX.setHeaderFooter(c; header="", footer="", differentFirst=false)             # emptied, so removed
        @test isnothing(XLSX.getHeaderFooter(c))
    end

    @testset "header and footer sections" begin
        @test XLSX._header_footer_sections("&LLeft&CCentre&RRight") == ((left="Left", center="Centre", right="Right"), true)
        @test XLSX._header_footer_sections("Plain") == ((left=nothing, center="Plain", right=nothing), false)
        @test XLSX._header_footer_sections("R&&D &L x") == ((left=" x", center="R&&D ", right=nothing), true)
        @test XLSX._header_footer_sections("&\"Calibri,Regular\"&Rx") == ((left=nothing, center="&\"Calibri,Regular\"", right="x"), true)
        @test XLSX._header_footer_sections("&\"&L,Bold\"y") == ((left=nothing, center="&\"&L,Bold\"y", right=nothing), false)  # quoted
        @test XLSX._header_footer_sections("a &") == ((left=nothing, center="a &", right=nothing), false)
        @test XLSX._header_footer_sections("&Lé&Rü") == ((left="é", center=nothing, right="ü"), true)
        @test XLSX._header_footer_text("&RA", :header) == "&RA"                         # Excel's own form is kept
        @test XLSX._header_footer_text((right="A", left="B"), :header) == "&LB&RA"      # in Excel's order
        @test XLSX._header_footer_text((left="", center=nothing, right="A"), :header) == "&RA"
        @test XLSX._header_footer_text("", :header) == ""
    end

    @testset "header and footer codes" begin
        @test XLSX.headerFooterField(:page) == "&P"
        @test [XLSX.headerFooterField(f) for f in (:pages, :date, :time, :file, :path, :sheet)] == ["&N", "&D", "&T", "&F", "&Z", "&A"]
        @test_throws XLSX.XLSXError XLSX.headerFooterField(:author)
        @test XLSX.headerFooterFont() == ""
        @test XLSX.headerFooterFont(name="Arial") == "&\"Arial,Regular\""
        @test XLSX.headerFooterFont(bold=true) == "&\"-,Bold\""
        @test XLSX.headerFooterFont(italic=true, size=9) == "&\"-,Italic\"&9"
        @test XLSX.headerFooterFont(under="single", strike=true, vertAlign="subscript") == "&U&S&Y"
        @test XLSX.headerFooterFont(under="double", vertAlign="superscript") == "&E&X"
        @test XLSX.headerFooterFont(under="none", vertAlign="baseline") == ""
        @test XLSX.headerFooterFont(color="red") == "&KFF0000"
        @test XLSX.headerFooterFont(color=:accent2) == "&K05+000"
        @test XLSX.headerFooterFont(color=XLSX.SchemeColor(:accent1; lumMod=75)) == "&K04-025"
        @test XLSX.headerFooterFont(color=XLSX.SchemeColor(:accent1; lumMod=60, lumOff=40)) == "&K04+040"
        @test_throws XLSX.XLSXError XLSX.headerFooterFont(name="A,B")
        @test_throws XLSX.XLSXError XLSX.headerFooterFont(name="")
        @test_throws XLSX.XLSXError XLSX.headerFooterFont(size=0)
        @test_throws XLSX.XLSXError XLSX.headerFooterFont(under="triple")
        @test_throws XLSX.XLSXError XLSX.headerFooterFont(vertAlign="up")
        @test_throws XLSX.XLSXError XLSX.headerFooterFont(color="notacolour")
    end

    @testset "header and footer errors" begin
        xf = three()
        s = xf["A"]
        @test_throws XLSX.XLSXError XLSX.setHeaderFooter(s)
        @test_throws XLSX.XLSXError XLSX.setHeaderFooter(s; heading="x")
        @test_throws XLSX.XLSXError XLSX.setHeaderFooter(s; header=1)
        @test_throws XLSX.XLSXError XLSX.setHeaderFooter(s; header=(middle="x",))
        @test_throws XLSX.XLSXError XLSX.setHeaderFooter(s; header=(left=1,))
        @test_throws XLSX.XLSXError XLSX.setHeaderFooter(s; header=(left="a&Rb",))      # a section code inside a section
        @test_throws XLSX.XLSXError XLSX.setHeaderFooter(s; scaleWithDoc=1)
        XLSX.setHeaderFooter(s; header="x"^253)                                         # 255 with its "&C", Excel's limit
        @test_throws XLSX.XLSXError XLSX.setHeaderFooter(s; header="x"^254)
        @test_throws XLSX.XLSXError XLSX.setHeaderFooter(s; header=(left="x"^126, right="y"^126))   # 256, as in Excel
        @test XLSX.getHeaderFooter(s).header.center == "x"^253                          # a failed call changes nothing

        r = XLSX.readxlsx(IOBuffer(roundtrip_bytes(xf)))
        @test_throws XLSX.XLSXError XLSX.setHeaderFooter(r["A"]; header="x")
        @test_throws XLSX.XLSXError XLSX.removeHeaderFooter(r["A"])
    end

    @testset "chartsheets have no page setup or protection" begin
        xf = XLSX.newxlsx("data")
        xf["data"]["A1:A2"] = [1, 2]
        c = XLSX.Charts.addChart(xf, :line; sheetname="Trend")
        XLSX.Charts.addSeries(c, "data!A1:A2")
        chs = xf["Trend"]
        @test !XLSX.isSheetProtected(chs)
        @test_throws XLSX.XLSXError XLSX.protectSheet(chs)
        @test_throws XLSX.XLSXError XLSX.setPageSetup(chs; orientation=:landscape)
        @test_throws XLSX.XLSXError XLSX.setPrintArea(chs, "A1:B2")
        @test_throws XLSX.XLSXError XLSX.setHeaderFooter(chs; header="x")
    end

    @testset "cell protection" begin
        xf = three()
        s = xf["A"]
        s["A1"] = ["Name" "Score"; "a" 1; "b" 2]
        XLSX.setFont(s, "B2"; bold=true)
        xf_children(ref) = [XLSX.localname(c) for c in XLSX.xml_elements(XLSX.styles_cell_xf(XLSX.get_workbook(xf), Int(XLSX.getcell(s, ref).style)))]

        @test XLSX.getCellProtection(s, "A1") == (locked=true, hidden=false)     # Excel's default
        @test XLSX.getCellProtection(s, "Z99") == (locked=true, hidden=false)    # empty

        @test isnothing(XLSX.setCellProtection(s, "B2:B5"; locked=false))       # B4:B5 are empty
        @test all(XLSX.getCellProtection(s, "B$r") == (locked=false, hidden=false) for r in 2:5)
        @test XLSX.getCellProtection(s, "A2").locked
        @test XLSX.getcell(s, "B5") isa XLSX.Cell && ismissing(s["B5"])          # formatted, still no value
        @test haskey(XLSX.getFont(s, "B2").font, "b")                           # other formatting kept
        @test s["B2"] == 1
        @test XLSX.getcell(s, "B4").style == XLSX.getcell(s, "B5").style        # one new style, shared

        XLSX.setCellProtection(s, "B3"; hidden=true)
        @test XLSX.getCellProtection(s, "B3") == (locked=false, hidden=true)     # locked kept
        XLSX.setCellProtection(s, 3, 2; locked=true)
        @test XLSX.getCellProtection(s, "B3") == (locked=true, hidden=true)
        XLSX.setCellProtection(s, "B3")                                          # nothing to change
        @test XLSX.getCellProtection(s, "B3") == (locked=true, hidden=true)

        # alignment and protection both kept, in the schema's order, whichever comes first
        XLSX.setAlignment(s, "B2"; horizontal="center")
        @test xf_children("B2") == ["alignment", "protection"]
        XLSX.setAlignment(s, "A2"; horizontal="center")
        XLSX.setCellProtection(s, "A2"; locked=false)
        @test xf_children("A2") == ["alignment", "protection"]
        @test XLSX.getAlignment(s, "A2").alignment["alignment"]["horizontal"] == "center"

        XLSX.setCellProtection(xf, "A!C1"; hidden=true)
        @test XLSX.getCellProtection(xf, "A!C1").hidden
        XLSX.addDefinedName(s, "Inputs", "D1:D2")
        XLSX.setCellProtection(s, "Inputs,E1"; locked=false)
        @test !XLSX.getCellProtection(s, "D2").locked && !XLSX.getCellProtection(s, "E1").locked

        XLSX.protectSheet(s)
        bytes = roundtrip_bytes(xf)
        r = XLSX.readxlsx(IOBuffer(bytes))["A"]
        @test XLSX.getCellProtection(r, "B5") == (locked=false, hidden=false)
        @test XLSX.getCellProtection(r, "B3") == (locked=true, hidden=true)
        @test XLSX.getCellProtection(r, "A1") == (locked=true, hidden=false)
        @test occursin(r"<c r=\"B5\" s=\"\d+\"/>", sheet_xml(bytes))           # an empty, formatted cell
        @test_throws XLSX.XLSXError XLSX.setCellProtection(r, "A1"; locked=false)   # read-only
        XLSX.openxlsx(IOBuffer(bytes); enable_cache=false) do f
            @test !XLSX.getCellProtection(f["A"], "B5").locked
        end

        @test_throws XLSX.XLSXError XLSX.getCellProtection(s, "A1:B2")
        @test_throws TypeError XLSX.setCellProtection(s, "A1"; locked=1)
        SAVE_FILES && save_outfile(xf)
    end

    @testset "cell protection by rows and columns" begin
        # Whole rows and columns, and `:`, reach as far as the used range, as the
        # format setters do; explicit cells and number ranges reach empty cells too.
        xf = XLSX.newxlsx("A")
        s = xf["A"]
        s["A1:C4"] = 1                                                          # used range A1:C4
        locked(ref) = XLSX.getCellProtection(s, ref).locked

        XLSX.setCellProtection(s, "B:B"; locked=false)
        @test !any(locked("B$r") for r in 1:4)
        @test locked("B5") && locked("A1") && locked("C1")
        @test XLSX.getcell(s, "B5") isa XLSX.EmptyCell                           # no cell made beyond the range

        XLSX.setCellProtection(s, "3:4"; hidden=true)
        @test all(XLSX.getCellProtection(s, "$c$r").hidden for c in "ABC", r in 3:4)
        @test !XLSX.getCellProtection(s, "A2").hidden && !XLSX.getCellProtection(s, "D3").hidden

        XLSX.setCellProtection(s, :, 3; locked=false)                            # column C, rows 1:4
        @test !any(locked("C$r") for r in 1:4) && locked("C5")
        XLSX.setCellProtection(s, 1, :; locked=true)                             # row 1, columns A:C
        @test all(locked("$c" * "1") for c in "ABC")
        XLSX.setCellProtection(s, [2, 4], 1:2; locked=false)
        @test !locked("A2") && !locked("B4") && locked("A3")
        XLSX.setCellProtection(s, 6:7, 1; locked=false)                          # number ranges reach empty cells
        @test !locked("A6") && !locked("A7")

        @test_throws XLSX.XLSXError XLSX.setCellProtection(s, "Z:Z"; locked=false)   # outside the used range
        @test_throws XLSX.XLSXError XLSX.setCellProtection(s, "20:21"; locked=false)
        SAVE_FILES && save_outfile(xf)
    end

    @testset "cell protection on chartsheets" begin
        xf = XLSX.newxlsx("data")
        xf["data"]["A1:A2"] = [1, 2]
        c = XLSX.Charts.addChart(xf, :line; sheetname="Trend")
        XLSX.Charts.addSeries(c, "data!A1:A2")
        @test_throws XLSX.XLSXError XLSX.setCellProtection(xf["Trend"], "A1"; locked=false)
        @test_throws XLSX.XLSXError XLSX.getCellProtection(xf["Trend"], "A1")
    end
end
