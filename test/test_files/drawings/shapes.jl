@testset "text boxes and shapes" begin
    drawing(xf, ws=xf[1]) = XLSX.xml_root_element(XLSX.get_xml_data(xf, XLSX._drawing_path_for_sheet(xf, XLSX._sheet_path(ws))))
    # Each object's XML, without Excel's random creation ids, its language or the
    # positions it works out from the cells' sizes (which differ between Excel's
    # blank workbook and XLSX.jl's).
    norm(s) = replace(replace(s, r">\s+<" => "><"), r"><a:extLst>.*?</a:extLst></xdr:cNvPr>" => "/>", "en-GB" => "en-US",
                      r"<a:off [^>]*/><a:ext [^>]*/>" => "XFRM", "<xdr:rowOff>25400</xdr:rowOff>" => "<xdr:rowOff>0</xdr:rowOff>")
    objects(xf) = [norm(XML.write(a; indentsize=0)) for a in XLSX.xml_elements(drawing(xf))]
    excel(name) = XLSX.openxlsx(joinpath(data_directory, name); mode="rw")
    function roundtrip(xf)
        io = IOBuffer()
        XLSX.writexlsx(io, xf)
        return XLSX.readxlsx(IOBuffer(take!(io)))
    end
    rich() = XLSX.RichTextString(XLSX.RichTextRun("Bold title", Pair{Symbol,Any}[:bold => true, :size => 14]),
                                 XLSX.RichTextRun("\nSecond line", Pair{Symbol,Any}[:color => "FF0000"]))

    @testset "as Excel writes them" begin
        xf = XLSX.newxlsx()
        info = XLSX.addTextBox(xf[1], "B2:D5", "Hello")
        @test info == (sheet="Sheet1", name="TextBox 1", from="B2", to="D5")
        @test objects(xf) == objects(excel("textbox_excel.xlsx"))

        xf = XLSX.newxlsx()
        XLSX.addTextBox(xf[1], "B2:D5", rich(); fill="FFFFCC", line="0000FF", lineWidth=2, lineDash=:dash,
                        horizontal=:center, vertical=:middle)
        @test XLSX.addShape(xf[1], "G2:H5", :rect).name == "Rectangle 2"           # numbered as Excel numbers them
        @test XLSX.addShape(xf[1], "G8:H11", :ellipse; text="Oval").name == "Oval 3"
        @test objects(xf) == objects(excel("shapes_excel.xlsx"))
        SAVE_FILES && save_outfile(xf)
    end

    @testset "reading them" begin
        ex = XLSX.readxlsx(joinpath(data_directory, "shapes_excel.xlsx"))           # read-only
        @test XLSX.getShapes(ex[1]) == [
            (name="TextBox 1", kind=:textbox, from="B2", to="D5", text="Bold title\nSecond line"),
            (name="Rectangle 2", kind=:rect, from="G2", to="H5", text=""),
            (name="Oval 3", kind=:ellipse, from="G8", to="H11", text="Oval")]
        @test isempty(XLSX.getImages(ex[1]))                                        # shapes aren't images
        @test isempty(XLSX.getShapes(XLSX.newxlsx()[1]))
    end

    @testset "options" begin
        xf = XLSX.newxlsx()
        s = xf[1]
        info = XLSX.addTextBox(s, "B2", "Small"; size=(128, 40))                   # snapped to cells, as an image is
        @test (info.from, info.to) == ("B2", "C3")
        @test XLSX.addTextBox(s, XLSX.CellRef("B6"), "Default size").to == "D9"
        XLSX.addTextBox(s, XLSX.CellRange("F2:H4"), "x"; fill=:none, line=:none, name="Plain")
        XLSX.addShape(s, "J2:K4", :rightArrow; fill=:accent2, line=XLSX.SchemeColor(:accent2; lumMod=75),
                      font=(color="white", bold=true, name="Arial", under="single", strike=true, vertAlign="superscript"), text="Go",
                      horizontal=:right, vertical=:bottom)
        x = XML.write(drawing(xf))
        @test occursin("<a:noFill/>", x) && occursin("<a:ln><a:noFill/></a:ln>", replace(x, r">\s+<" => "><"))
        @test occursin("name=\"Plain\"", x) && occursin("prst=\"rightArrow\"", x)
        @test occursin("<a:schemeClr val=\"accent2\"", x)
        @test occursin(r"<a:rPr lang=\"en-US\" sz=\"1100\" b=\"1\" u=\"sng\" strike=\"sngStrike\" baseline=\"30000\"", x)
        @test occursin("typeface=\"Arial\"", x) && occursin("algn=\"r\"", x) && occursin("anchor=\"b\"", x)
        @test [sh.kind for sh in XLSX.getShapes(s)] == [:textbox, :textbox, :textbox, :rightArrow]

        XLSX.addTextBox(s, "B12:E14", "One\r\nTwo\n\nFour")                         # an empty paragraph keeps its place
        @test XLSX.getShapes(s)[end].text == "One\nTwo\n\nFour"
        XLSX.addImage(s, "M2", joinpath(data_directory, "track_start.png"))
        @test length(XLSX.getImages(s)) == 1 && length(XLSX.getShapes(s)) == 5     # each sees its own
        r = roundtrip(xf)
        @test XLSX.getShapes(r[1]) == XLSX.getShapes(s)
    end

    @testset "deleting" begin
        xf = XLSX.newxlsx()
        s = xf[1]
        XLSX.addTextBox(s, "B2", "a")
        XLSX.addShape(s, "B2", :rect)
        XLSX.addShape(s, "F2", :ellipse; name="Target")
        XLSX.addImage(s, "B2", joinpath(data_directory, "track_start.png"))
        @test XLSX.deleteShape(s, "Target") == 1
        @test XLSX.deleteShape(s, "B2") == 2                                        # both, not the image
        @test isempty(XLSX.getShapes(s)) && length(XLSX.getImages(s)) == 1
        @test XLSX.deleteShape(s, XLSX.CellRef("B2")) == 0
        XLSX.deleteImage(s, "B2")
        XLSX.addShape(s, "C3", :rect)
        @test XLSX.deleteShape(s, "C3") == 1
        @test isnothing(XLSX._drawing_path_for_sheet(xf, XLSX._sheet_path(s)))     # the empty drawing goes
    end

    @testset "rows, columns and sheets" begin
        xf = XLSX.newxlsx()
        s = xf[1]
        XLSX.addTextBox(s, "B2:C3", "moves")
        XLSX.addShape(s, "E5:E5", :rect; text="goes")
        XLSX.insertRows(s, 1; count=2)
        XLSX.insertColumns(s, 1)
        @test [(sh.from, sh.to) for sh in XLSX.getShapes(s)] == [("C4", "D5"), ("F7", "F7")]
        XLSX.deleteRows(s, 7)                                                      # all of the shape's rows
        @test [sh.text for sh in XLSX.getShapes(s)] == ["moves"]
        c = XLSX.copysheet!(s, "Copy")
        @test XLSX.getShapes(c) == XLSX.getShapes(s)
        XLSX.addShape(c, "H2", :star5)
        @test length(XLSX.getShapes(c)) == 2 && length(XLSX.getShapes(s)) == 1   # the copy has its own drawing
        @test XLSX.getShapes(roundtrip(xf)["Copy"]) == XLSX.getShapes(c)

        ex = excel("shapes_excel.xlsx")                    # Excel writes no relationships for a drawing of shapes alone
        @test XLSX.getShapes(XLSX.copysheet!(ex[1], "Copy")) == XLSX.getShapes(ex[1])
    end

    @testset "errors" begin
        xf = XLSX.newxlsx()
        s = xf[1]
        @test_throws XLSX.XLSXError XLSX.addShape(s, "B2", :blob)
        @test_throws XLSX.XLSXError XLSX.addTextBox(s, "nonsense", "x")
        @test_throws XLSX.XLSXError XLSX.addTextBox(s, "B2", "x"; size=(0, 10))
        @test_throws XLSX.XLSXError XLSX.addTextBox(s, "B2", "x"; colour="red")
        @test_throws XLSX.XLSXError XLSX.addTextBox(s, "B2", "x"; fill="notacolour")
        @test_throws XLSX.XLSXError XLSX.addTextBox(s, "B2", "x"; line="red", lineDash=:wavy)
        @test_throws XLSX.XLSXError XLSX.addTextBox(s, "B2", "x"; lineWidth=-1)
        @test_throws XLSX.XLSXError XLSX.addTextBox(s, "B2", "x"; horizontal=:middle)
        @test_throws XLSX.XLSXError XLSX.addTextBox(s, "B2", "x"; vertical=:center)
        @test_throws XLSX.XLSXError XLSX.addTextBox(s, "B2", "x"; font=(weight=700,))
        @test_throws XLSX.XLSXError XLSX.addTextBox(s, "B2", "x"; font=(size=0,))
        @test_throws XLSX.XLSXError XLSX.addTextBox(s, "B2", "x"; font=(vertAlign="up",))
        @test_throws XLSX.XLSXError XLSX.addTextBox(s, "B2", "x"; font=14)
        @test_throws XLSX.XLSXError XLSX.addTextBox(s, "B2", "x"; name="")
        @test isempty(XLSX.getShapes(s))                                            # nothing was added
        r = roundtrip(xf)
        @test_throws XLSX.XLSXError XLSX.addTextBox(r[1], "B2", "x")
        @test_throws XLSX.XLSXError XLSX.deleteShape(r[1], "B2")
        XLSX.Charts.addChart(xf, :line; sheetname="Chart")
        @test_throws XLSX.XLSXError XLSX.addTextBox(xf["Chart"], "B2", "x")
    end
end
