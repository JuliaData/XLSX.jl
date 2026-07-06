@testset "Styles" begin

    @testset "Original" begin
        using XLSX: CellValue, id, getcell, setdata!, CellRef
        xfile = XLSX.open_empty_template()
        wb = XLSX.get_workbook(xfile)
        sheet = xfile["Sheet1"]

        datefmt = XLSX.styles_add_numFmt(wb, "yyyymmdd")
        numfmt = XLSX.styles_add_numFmt(wb, "\$* #,##0.00;\$* (#,##0.00);\$* \"-\"??;[Magenta]@")

        #Check format id numbers dont intersect with predefined formats or each other
        @test datefmt == 164
        @test numfmt == 165

        font = XLSX.styles_add_cell_font(wb, Dict("b" => nothing, "sz" => Dict("val" => "24")))
        xroot = XLSX.styles_xmlroot(wb)
        fontnodes = XLSX.find_all_nodes("/$(XLSX.SPREADSHEET_NAMESPACE_XPATH_ARG):styleSheet/$(XLSX.SPREADSHEET_NAMESPACE_XPATH_ARG):fonts/$(XLSX.SPREADSHEET_NAMESPACE_XPATH_ARG):font", xroot)
        fontnode = fontnodes[font+1] # XML is zero indexed so we need to add 1 to get the right node

        # Check the font was written correctly
        @test XML.tag(fontnode) == "font"
        @test length(XML.children(fontnode)) == 2
        @test XML.tag(XML.children(fontnode)[1]) == "b"
        @test XML.tag(XML.children(fontnode)[2]) == "sz"
        @test XML.children(fontnode)[2]["val"] == "24"

        textstyle = XLSX.styles_add_cell_xf(wb, Dict("applyFont" => "true", "fontId" => "$font"))
        datestyle = XLSX.styles_add_cell_xf(wb, Dict("applyNumberFormat" => "1", "numFmtId" => "$datefmt"))
        numstyle = XLSX.styles_add_cell_xf(wb, Dict("applyFont" => "1", "applyNumberFormat" => "1", "fontId" => "$font", "numFmtId" => "$numfmt"))

        xf = XLSX.styles_get_cellXf_with_numFmtId(wb, 1000)
        @test xf == XLSX.EmptyCellDataFormat()
        @test isempty(xf)
        @test id(xf) == UInt64(0)

        @test textstyle isa XLSX.CellDataFormat
        @test !isempty(textstyle)
        @test id(textstyle) == UInt64(1)

        @test XLSX.styles_get_cellXf_with_numFmtId(wb, datefmt) == datestyle
        @test XLSX.styles_numFmt_formatCode(wb, string(datefmt)) == "yyyymmdd"
        @test datestyle isa XLSX.CellDataFormat
        @test !isempty(datestyle)
        @test id(datestyle) == UInt64(2)

        @test XLSX.styles_get_cellXf_with_numFmtId(wb, numfmt) == numstyle
        @test XLSX.styles_numFmt_formatCode(wb, string(numfmt)) == "\$* #,##0.00;\$* (#,##0.00);\$* \"-\"??;[Magenta]@"
        @test numstyle isa XLSX.CellDataFormat
        @test !isempty(numstyle)
        @test id(numstyle) == UInt64(3)

        setdata!(sheet, CellRef("A1"), CellValue(Date(2011, 10, 13), datestyle))
        setdata!(sheet, CellRef("A2"), CellValue(1000, numstyle))
        setdata!(sheet, CellRef("A3"), CellValue(1000.10, numstyle))
        setdata!(sheet, CellRef("A4"), CellValue(-1000.10, numstyle))
        setdata!(sheet, CellRef("A5"), CellValue(0, numstyle))
        setdata!(sheet, CellRef("A6"), CellValue("hello", numstyle))
        setdata!(sheet, CellRef("B1"), CellValue("hello world", textstyle))

        @test sheet["A1"] == Date(2011, 10, 13)
        cell = getcell(sheet, "A1")
        @test cell.style == id(datestyle)
        formatid = XLSX.styles_cell_xf_numFmtId(wb, Int64(cell.style))
        @test formatid == datefmt

        cellstyle = getcell(sheet, "A2").style
        @test cellstyle == id(numstyle)
        formatid = XLSX.styles_cell_xf_numFmtId(wb, Int64(cellstyle))
        @test formatid == numfmt

        @test sheet["A2"] == 1000
        @test sheet["A3"] == 1000.10
        @test XLSX.getcell(sheet, "A3").style == cellstyle
        @test sheet["A4"] == -1000.10
        @test XLSX.getcell(sheet, "A4").style == cellstyle
        @test sheet["A5"] == 0
        @test XLSX.getcell(sheet, "A5").style == cellstyle
        @test sheet["A6"] == "hello"
        @test XLSX.getcell(sheet, "A6").style == cellstyle

        @test sheet["B1"] == "hello world"
        @test XLSX.getcell(sheet, "B1").style == id(textstyle)

        sheet["B2"] = CellValue("hello world", textstyle)
        @test sheet["B2"] == "hello world"
        @test XLSX.getcell(sheet, "B2").style == id(textstyle)

        sheet[3, 1] = CellValue("hello friend", textstyle)
        @test sheet[3, 1] == "hello friend"
        @test XLSX.getcell(sheet, 3, 1).style == id(textstyle)

        # Check CellDataFormat only works with CellValues
        @test_throws MethodError XLSX.CellValue([1, 2, 3, 4], textstyle)
        SAVE_FILES && save_outfile(xfile)

        using XLSX: styles_is_datetime, styles_add_numFmt, styles_add_cell_xf
        # Capitalized and single character numfmts
        xfile = XLSX.open_empty_template()
        wb = XLSX.get_workbook(xfile)
        sheet = xfile["Sheet1"]

        fmt = styles_add_numFmt(wb, "yyyy m d")
        style = styles_add_cell_xf(wb, Dict("numFmtId" => "$fmt"))
        @test styles_is_datetime(wb, style)

        fmt = styles_add_numFmt(wb, "h:m:s")
        style = styles_add_cell_xf(wb, Dict("numFmtId" => "$fmt"))
        @test styles_is_datetime(wb, style)

        fmt = styles_add_numFmt(wb, "0.00")
        style = styles_add_cell_xf(wb, Dict("numFmtId" => "$fmt"))
        @test !styles_is_datetime(wb, style)
        @test XLSX.styles_is_float(wb, style)

        fmt = styles_add_numFmt(wb, "[red]yyyy m d")
        style = styles_add_cell_xf(wb, Dict("numFmtId" => "$fmt"))
        @test styles_is_datetime(wb, style)
        fmt = styles_add_numFmt(wb, "[red] h:m:s")
        style = styles_add_cell_xf(wb, Dict("numFmtId" => "$fmt"))
        @test styles_is_datetime(wb, style)
        fmt = styles_add_numFmt(wb, "[red] 0.00; [magenta] 0.00")
        style = styles_add_cell_xf(wb, Dict("numFmtId" => "$fmt"))
        @test !styles_is_datetime(wb, style)
        @test XLSX.styles_is_float(wb, style)

        fmt = styles_add_numFmt(wb, "YYYY M D")
        style = styles_add_cell_xf(wb, Dict("numFmtId" => "$fmt"))
        @test styles_is_datetime(wb, style)
        fmt = styles_add_numFmt(wb, "H:M:S")
        style = styles_add_cell_xf(wb, Dict("numFmtId" => "$fmt"))
        @test styles_is_datetime(wb, style)

        fmt = styles_add_numFmt(wb, "m")
        style = styles_add_cell_xf(wb, Dict("numFmtId" => "$fmt"))
        @test styles_is_datetime(wb, style)

        fmt = styles_add_numFmt(wb, "M")
        style = styles_add_cell_xf(wb, Dict("numFmtId" => "$fmt"))
        @test styles_is_datetime(wb, style)

        fmt = styles_add_numFmt(wb, "y")
        style = styles_add_cell_xf(wb, Dict("numFmtId" => "$fmt"))
        @test !styles_is_datetime(wb, style)

        fmt = styles_add_numFmt(wb, "[s]")
        style = styles_add_cell_xf(wb, Dict("numFmtId" => "$fmt"))
        @test styles_is_datetime(wb, style)

        fmt = styles_add_numFmt(wb, "am/pm")
        style = styles_add_cell_xf(wb, Dict("numFmtId" => "$fmt"))
        @test styles_is_datetime(wb, style)

        fmt = styles_add_numFmt(wb, "a/p")
        style = styles_add_cell_xf(wb, Dict("numFmtId" => "$fmt"))
        @test styles_is_datetime(wb, style)

        fmt = styles_add_numFmt(wb, "\"Monday\"")
        style = styles_add_cell_xf(wb, Dict("numFmtId" => "$fmt"))
        @test !styles_is_datetime(wb, style)
        @test !XLSX.styles_is_float(wb, style)

        fmt = styles_add_numFmt(wb, "0.00*m")
        style = styles_add_cell_xf(wb, Dict("numFmtId" => "$fmt"))
        @test !styles_is_datetime(wb, style)
        @test XLSX.styles_is_float(wb, style)

        fmt = styles_add_numFmt(wb, "0.00_m")
        style = styles_add_cell_xf(wb, Dict("numFmtId" => "$fmt"))
        @test !styles_is_datetime(wb, style)
        @test XLSX.styles_is_float(wb, style)

        fmt = styles_add_numFmt(wb, "0.00\\d")
        style = styles_add_cell_xf(wb, Dict("numFmtId" => "$fmt"))
        @test !styles_is_datetime(wb, style)
        @test XLSX.styles_is_float(wb, style)

        fmt = styles_add_numFmt(wb, "[red][>1.5]000")
        style = styles_add_cell_xf(wb, Dict("numFmtId" => "$fmt"))
        @test !XLSX.styles_is_float(wb, style)

        fmt = styles_add_numFmt(wb, "0.#")
        style = styles_add_cell_xf(wb, Dict("numFmtId" => "$fmt"))
        @test XLSX.styles_is_float(wb, style)

        fmt = styles_add_numFmt(wb, "\"hello.\" ###")
        style = styles_add_cell_xf(wb, Dict("numFmtId" => "$fmt"))
        @test !XLSX.styles_is_float(wb, style)

        fmt = styles_add_numFmt(wb, ".??")
        style = styles_add_cell_xf(wb, Dict("numFmtId" => "$fmt"))
        @test XLSX.styles_is_float(wb, style)

        fmt = styles_add_numFmt(wb, "#E+00")
        style = styles_add_cell_xf(wb, Dict("numFmtId" => "$fmt"))
        @test XLSX.styles_is_float(wb, style)

        fmt = styles_add_numFmt(wb, "0E+00")
        style = styles_add_cell_xf(wb, Dict("numFmtId" => "$fmt"))
        @test XLSX.styles_is_float(wb, style)
 
        fmt = styles_add_numFmt(wb, "# ??/??")
        style = styles_add_cell_xf(wb, Dict("numFmtId" => "$fmt"))
        @test XLSX.styles_is_float(wb, style)

        fmt = styles_add_numFmt(wb, "*.00")
        style = styles_add_cell_xf(wb, Dict("numFmtId" => "$fmt"))
        @test !XLSX.styles_is_float(wb, style)

        fmt = styles_add_numFmt(wb, "\\.00")
        style = styles_add_cell_xf(wb, Dict("numFmtId" => "$fmt"))
        @test !XLSX.styles_is_float(wb, style)

        fmt = styles_add_numFmt(wb, "00_.")
        style = styles_add_cell_xf(wb, Dict("numFmtId" => "$fmt"))
        @test !XLSX.styles_is_float(wb, style)

        SAVE_FILES && save_outfile(xfile)
    end
    
    @testset "number formats" begin
        XLSX.openxlsx(joinpath(data_directory, "general.xlsx")) do f
            show(IOBuffer(), f)
            sheet = f["general"]
            @test sheet["A1"] == "text"
            @test sheet["B1"] == "regular text"
            @test sheet["A2"] == "integer"
            @test sheet["B2"] == 102
            @test sheet["A3"] == "float"
            @test isapprox(sheet["B3"], 102.2)
            @test sheet["A4"] == "date"
            @test sheet["B4"] == Date(1983, 4, 16)
            @test sheet["A5"] == "hour"
            @test sheet["B5"] == Dates.Time(Dates.Hour(19), Dates.Minute(45))
            @test sheet["A6"] == "datetime"
            @test sheet["B6"] == Date(2018, 4, 16) + Dates.Time(Dates.Hour(19), Dates.Minute(19), Dates.Second(51))
            @test f["general!B7"] == -220.0
            @test f["general!B8"] == -2000
            @test f["general!B9"] == 100000000000000
            @test f["general!B10"] == -100000000000000
        end
    end

    @testset "setFont" begin

        f = XLSX.newxlsx()
        s = f[1]
        s["A1:B2,D1:E2"] = ""

        XLSX.setFont(s, "A1:A2"; bold=true, italic=true, size=24, name="Arial")
        XLSX.setFont(s, "B1:B2"; bold=true, italic=false, size=14, name="Aptos")
        XLSX.setFont(s, "D1:D2"; bold=false, italic=true, size=34, name="Berlin Sans FB Demi")
        XLSX.setFont(s, "E1:E2"; bold=false, italic=false, size=4, name="Times New Roman")
        XLSX.setUniformFont(s, "A1:B2,D1:E2"; color="blue") # `setUniformAttribute()` on a non-contiguous range
        @test XLSX.getFont(s, "A1").font == Dict("b" => nothing, "i" => nothing, "sz" => Dict("val" => "24"), "name" => Dict("val" => "Arial"), "color" => Dict("rgb" => "FF0000FF"))
        @test XLSX.getFont(s, "B2").font == Dict("b" => nothing, "i" => nothing, "sz" => Dict("val" => "24"), "name" => Dict("val" => "Arial"), "color" => Dict("rgb" => "FF0000FF"))
        @test XLSX.getFont(s, "D1").font == Dict("b" => nothing, "i" => nothing, "sz" => Dict("val" => "24"), "name" => Dict("val" => "Arial"), "color" => Dict("rgb" => "FF0000FF"))
        @test XLSX.getFont(s, "E2").font == Dict("b" => nothing, "i" => nothing, "sz" => Dict("val" => "24"), "name" => Dict("val" => "Arial"), "color" => Dict("rgb" => "FF0000FF"))
        SAVE_FILES && save_outfile(f)

        f = XLSX.newxlsx()
        s = f[1]
        s["A1:Z26"] = ""

        XLSX.setFont(s, "Sheet1!A1:A2"; bold=true, italic=true, size=24, name="Arial", color="blue")
        @test XLSX.getFont(s, "A1").font == Dict("b" => nothing, "i" => nothing, "sz" => Dict("val" => "24"), "name" => Dict("val" => "Arial"), "color" => Dict("rgb" => "FF0000FF"))
        XLSX.setFont(s, "Sheet1!Y:Z"; bold=true, italic=false, size=14, name="Aptos", color="blue")
        @test XLSX.getFont(s, "Y20").font == Dict("b" => nothing, "sz" => Dict("val" => "14"), "name" => Dict("val" => "Aptos"), "color" => Dict("rgb" => "FF0000FF"))
        XLSX.setFont(s, "Sheet1!2:3"; bold=false, italic=true, size=34, name="Berlin Sans FB Demi", color="blue")
        @test XLSX.getFont(s, "M3").font == Dict("i" => nothing, "sz" => Dict("val" => "34"), "name" => Dict("val" => "Berlin Sans FB Demi"), "color" => Dict("rgb" => "FF0000FF"))
        XLSX.setFont(f, "Sheet1!A1:A2"; bold=false, italic=false, size=14, name="Aptos", color="green")
        @test XLSX.getFont(s, "A1").font == Dict("sz" => Dict("val" => "14"), "name" => Dict("val" => "Aptos"), "color" => Dict("rgb" => "FF008000"))
        XLSX.setFont(f, "Sheet1!Y:Z"; bold=false, italic=true, size=24, name="Arial", color="green")
        @test XLSX.getFont(s, "Y20").font == Dict("i" => nothing, "sz" => Dict("val" => "24"), "name" => Dict("val" => "Arial"), "color" => Dict("rgb" => "FF008000"))
        XLSX.setFont(f, "Sheet1!2:3"; bold=true, italic=false, size=24, name="Times New Roman", color="green")
        @test XLSX.getFont(s, "M3").font == Dict("b" => nothing, "sz" => Dict("val" => "24"), "name" => Dict("val" => "Times New Roman"), "color" => Dict("rgb" => "FF008000"))
        XLSX.setFont(s, "E1,E2,G2:G4"; bold=false, italic=false, size=4, name="Times New Roman", color="blue")
        @test XLSX.getFont(s, "G3").font == Dict("sz" => Dict("val" => "4"), "name" => Dict("val" => "Times New Roman"), "color" => Dict("rgb" => "FF0000FF"))
        XLSX.setFont(s, :, 15:16; bold=true, italic=false, size=38, name="Wingdings", color="red")
        @test XLSX.getFont(s, "P10").font == Dict("b" => nothing, "sz" => Dict("val" => "38"), "name" => Dict("val" => "Wingdings"), "color" => Dict("rgb" => "FFFF0000"))
        XLSX.setFont(s, 15:16, :; bold=false, italic=true, size=8, name="Wingdings", color="red")
        @test XLSX.getFont(f, "Sheet1!T16").font == Dict("i" => nothing, "sz" => Dict("val" => "8"), "name" => Dict("val" => "Wingdings"), "color" => Dict("rgb" => "FFFF0000"))
        XLSX.setFont(s, [20, 22, 24], :; bold=false, italic=true, size=48, name="Aptos", color="red")
        @test XLSX.getFont(f, "Sheet1!H22").font == Dict("i" => nothing, "sz" => Dict("val" => "48"), "name" => Dict("val" => "Aptos"), "color" => Dict("rgb" => "FFFF0000"))
        XLSX.setUniformFont(s, [15, 16, 20, 22, 24], :; bold=false, italic=true, size=28, name="Aptos", color="red")
        @test XLSX.getFont(f, "Sheet1!H15").font == Dict("i" => nothing, "sz" => Dict("val" => "28"), "name" => Dict("val" => "Aptos"), "color" => Dict("rgb" => "FFFF0000"))
        @test XLSX.getFont(f, "Sheet1!H22").font == Dict("i" => nothing, "sz" => Dict("val" => "28"), "name" => Dict("val" => "Aptos"), "color" => Dict("rgb" => "FFFF0000"))
        SAVE_FILES && save_outfile(f)

        xfile = XLSX.open_empty_template()
        wb = XLSX.get_workbook(xfile)
        sheet = xfile["Sheet1"]

        col_names = ["Integers", "Strings", "Floats", "Booleans", "Dates", "Times", "DateTimes", "AbstractStrings", "Rational", "Irrationals", "MixedStringNothingMissing"]
        data = Vector{Any}(undef, 11)
        data[1] = [1, 2, missing, UInt8(4)]
        data[2] = ["Hey", "You", "Out", "There"]
        data[3] = [101.5, 102.5, missing, 104.5]
        data[4] = [true, false, missing, true]
        data[5] = [Date(2018, 2, 1), Date(2018, 3, 1), Date(2018, 5, 20), Date(2018, 6, 2)]
        data[6] = [Dates.Time(19, 10), Dates.Time(19, 20), Dates.Time(19, 30), Dates.Time(0, 0)]
        data[7] = [Dates.DateTime(2018, 5, 20, 19, 10), Dates.DateTime(2018, 5, 20, 19, 20), Dates.DateTime(2018, 5, 20, 19, 30), Dates.DateTime(2018, 5, 20, 19, 40)]
        data[8] = SubString.(["Hey", "You", "Out", "There"], 1, 2)
        data[9] = [1 // 2, 1 // 3, missing, 22 // 3]
        data[10] = [pi, sqrt(2), missing, sqrt(5)]
        data[11] = [nothing, "middle", missing, nothing]

        XLSX.writetable!(sheet, data, col_names; write_columnnames=true)

        @test isnothing(XLSX.getFont(xfile, "Sheet1!B2")) && isnothing(XLSX.getFont(sheet, "B2"))

        # Default font attributes are present in an empty worksheet until overwritten.
        default_font = XLSX.getDefaultFont(sheet).font
        dname = default_font["name"]["val"]
        dsize = default_font["sz"]["val"]
        dcolorkey = collect(keys(default_font["color"]))[1]
        dcolorval = collect(values(default_font["color"]))[1]

        # Sheet mismatch
        @test_throws XLSX.XLSXError XLSX.setFont(sheet, "S2!A1"; bold=true, size=24, name="Arial")

        XLSX.setFont(sheet, "A1"; bold=true, size=24, name="Arial")
        @test XLSX.getFont(sheet, "A1").font == Dict("b" => nothing, "sz" => Dict("val" => "24"), "name" => Dict("val" => "Arial"), "color" => Dict(dcolorkey => dcolorval))
        XLSX.setFont(sheet, "A1"; size=18)
        @test XLSX.getFont(sheet, "A1").font == Dict("b" => nothing, "sz" => Dict("val" => "18"), "name" => Dict("val" => "Arial"), "color" => Dict(dcolorkey => dcolorval))
        XLSX.setFont(xfile, "Sheet1!A1"; bold=false, size=24, name="Aptos")
        @test XLSX.getFont(xfile, "Sheet1!A1").font == Dict("sz" => Dict("val" => "24"), "name" => Dict("val" => "Aptos"), "color" => Dict(dcolorkey => dcolorval))

        @test XLSX.getFont(xfile, "Sheet1!A1").fontId == XLSX.getFont(sheet, "A1").fontId
        @test XLSX.getFont(xfile, "Sheet1!A1").font == XLSX.getFont(sheet, "A1").font
        @test XLSX.getFont(xfile, "Sheet1!A1").applyFont == XLSX.getFont(sheet, "A1").applyFont

        XLSX.setFont(xfile, "Sheet1!A2"; size=18)
        @test XLSX.getFont(xfile, "Sheet1!A2").font == Dict("sz" => Dict("val" => "18"), "name" => Dict("val" => dname), "color" => Dict(dcolorkey => dcolorval))
        XLSX.setFont(sheet, "A2"; size=24)
        @test XLSX.getFont(sheet, "A2").font == Dict("sz" => Dict("val" => "24"), "name" => Dict("val" => dname), "color" => Dict(dcolorkey => dcolorval))
        XLSX.setFont(sheet, "A2"; size=28, name="Aptos")
        @test XLSX.getFont(sheet, "A2").font == Dict("sz" => Dict("val" => "28"), "name" => Dict("val" => "Aptos"), "color" => Dict(dcolorkey => dcolorval))

        XLSX.setFont(sheet, "A3"; italic=true, name="Berlin Sans FB Demi")
        @test XLSX.getFont(sheet, "A3").font == Dict("i" => nothing, "sz" => Dict("val" => dsize), "name" => Dict("val" => "Berlin Sans FB Demi"), "color" => Dict(dcolorkey => dcolorval))
        XLSX.setFont(xfile, "Sheet1!A3"; size=24)
        @test XLSX.getFont(xfile, "Sheet1!A3").font == Dict("i" => nothing, "sz" => Dict("val" => "24"), "name" => Dict("val" => "Berlin Sans FB Demi"), "color" => Dict(dcolorkey => dcolorval))

        XLSX.setFont(xfile, "Sheet1!A4"; size=28, name="Aptos")
        @test XLSX.getFont(xfile, "Sheet1!A4").font == Dict("sz" => Dict("val" => "28"), "name" => Dict("val" => "Aptos"), "color" => Dict(dcolorkey => dcolorval))

        XLSX.setFont(sheet, "B1"; bold=true, italic=true, size=14, color="FF00FF00")
        @test XLSX.getFont(sheet, "B1").font == Dict("b" => nothing, "i" => nothing, "sz" => Dict("val" => "14"), "name" => Dict("val" => dname), "color" => Dict("rgb" => "FF00FF00"))
        XLSX.setFont(xfile, "Sheet1!B1"; bold=false, italic=false, size=12, name="Berlin Sans FB Demi")
        @test XLSX.getFont(xfile, "Sheet1!B1").font == Dict("sz" => Dict("val" => "12"), "name" => Dict("val" => "Berlin Sans FB Demi"), "color" => Dict("rgb" => "FF00FF00"))

        XLSX.setFont(sheet, "B2"; color="FF000000")
        @test XLSX.getFont(sheet, "B2").font == Dict("sz" => Dict("val" => dsize), "name" => Dict("val" => dname), "color" => Dict("rgb" => "FF000000"))
        XLSX.setFont(xfile, "Sheet1!B2"; bold=true, italic=true, size=14, color="FF00FF00")
        @test XLSX.getFont(xfile, "Sheet1!B2").font == Dict("b" => nothing, "i" => nothing, "sz" => Dict("val" => "14"), "name" => Dict("val" => dname), "color" => Dict("rgb" => "FF00FF00"))

        XLSX.setFont(xfile, "Sheet1!B3"; name="Berlin Sans FB Demi", color="FF000000")
        @test XLSX.getFont(xfile, "Sheet1!B3").font == Dict("sz" => Dict("val" => dsize), "name" => Dict("val" => "Berlin Sans FB Demi"), "color" => Dict("rgb" => "FF000000"))

        XLSX.setFont(sheet, "A1:B2"; size=18, name="Arial")
        @test XLSX.getFont(xfile, "Sheet1!A1").font == Dict("sz" => Dict("val" => "18"), "name" => Dict("val" => "Arial"), "color" => Dict(dcolorkey => dcolorval))
        @test XLSX.getFont(sheet, "A2").font == Dict("sz" => Dict("val" => "18"), "name" => Dict("val" => "Arial"), "color" => Dict(dcolorkey => dcolorval))
        @test XLSX.getFont(xfile, "Sheet1!B1").font == Dict("sz" => Dict("val" => "18"), "name" => Dict("val" => "Arial"), "color" => Dict("rgb" => "FF00FF00"))
        @test XLSX.getFont(xfile, "Sheet1!B2").font == Dict("b" => nothing, "i" => nothing, "sz" => Dict("val" => "18"), "name" => Dict("val" => "Arial"), "color" => Dict("rgb" => "FF00FF00"))


        XLSX.writexlsx("output.xlsx", xfile, overwrite=true)
        SAVE_FILES && save_outfile("output.xlsx")
        @test isfile("output.xlsx")

        XLSX.openxlsx("output.xlsx") do f # Check the updated fonts were written correctly
            s = f["Sheet1"]
            @test XLSX.getFont(f, "Sheet1!A1").font == Dict("sz" => Dict("val" => "18"), "name" => Dict("val" => "Arial"), "color" => Dict(dcolorkey => dcolorval))
            @test XLSX.getFont(s, "A2").font == Dict("sz" => Dict("val" => "18"), "name" => Dict("val" => "Arial"), "color" => Dict(dcolorkey => dcolorval))
            @test XLSX.getFont(f, "Sheet1!A3").font == Dict("i" => nothing, "sz" => Dict("val" => "24"), "name" => Dict("val" => "Berlin Sans FB Demi"), "color" => Dict(dcolorkey => dcolorval))
            @test XLSX.getFont(s, "A4").font == Dict("sz" => Dict("val" => "28"), "name" => Dict("val" => "Aptos"), "color" => Dict(dcolorkey => dcolorval))
            @test XLSX.getFont(f, "Sheet1!B1").font == Dict("sz" => Dict("val" => "18"), "name" => Dict("val" => "Arial"), "color" => Dict("rgb" => "FF00FF00"))
            @test XLSX.getFont(s, "B2").font == Dict("b" => nothing, "i" => nothing, "sz" => Dict("val" => "18"), "name" => Dict("val" => "Arial"), "color" => Dict("rgb" => "FF00FF00"))
            @test XLSX.getFont(f, "Sheet1!B3").font == Dict("sz" => Dict("val" => dsize), "name" => Dict("val" => "Berlin Sans FB Demi"), "color" => Dict("rgb" => "FF000000"))
        end

        # Now try a range
        XLSX.setUniformFont(sheet, "A1:B4"; size=12, name="Times New Roman", color="FF040404")
        @test XLSX.getFont(xfile, "Sheet1!A1").font == Dict("sz" => Dict("val" => "12"), "name" => Dict("val" => "Times New Roman"), "color" => Dict("rgb" => "FF040404"))
        @test XLSX.getFont(xfile, "Sheet1!A4").font == Dict("sz" => Dict("val" => "12"), "name" => Dict("val" => "Times New Roman"), "color" => Dict("rgb" => "FF040404"))
        @test XLSX.getFont(xfile, "Sheet1!B3").font == Dict("sz" => Dict("val" => "12"), "name" => Dict("val" => "Times New Roman"), "color" => Dict("rgb" => "FF040404"))
        @test XLSX.getFont(xfile, "Sheet1!B4").font == Dict("sz" => Dict("val" => "12"), "name" => Dict("val" => "Times New Roman"), "color" => Dict("rgb" => "FF040404"))

        isfile("output.xlsx") && rm("output.xlsx")
        SAVE_FILES && save_outfile(xfile)

        f = XLSX.newxlsx()
        sheet = f[1]
        sheet["A1:E5"] = ""
        XLSX.setFont(sheet, :, [1, 2, 3, 4, 5]; size=18, name="Arial", color="FF040404")
        XLSX.setFont(sheet, 1:3, [1, 3]; size=12, name="Aptos", color="FF040408")
        XLSX.setFont(sheet, [4, 5], [2, 4]; size=6, name="Courier New", color="FF040400")
        @test XLSX.getFont(sheet, "A4").font == Dict("sz" => Dict("val" => "18"), "name" => Dict("val" => "Arial"), "color" => Dict("rgb" => "FF040404"))
        @test XLSX.getFont(sheet, "A3").font == Dict("sz" => Dict("val" => "12"), "name" => Dict("val" => "Aptos"), "color" => Dict("rgb" => "FF040408"))
        @test XLSX.getFont(f, "Sheet1!D5").font == Dict("sz" => Dict("val" => "6"), "name" => Dict("val" => "Courier New"), "color" => Dict("rgb" => "FF040400"))
        @test_throws XLSX.XLSXError XLSX.setFont(sheet, "1:10"; size=18, name="Arial", color="FF040404")
        @test_throws XLSX.XLSXError XLSX.setFont(sheet, "A:K"; size=18, name="Arial", color="FF040404")
        @test_throws XLSX.XLSXError XLSX.setFont(f, "Sheet1!garbage"; size=18, name="Arial", color="FF040404")
        @test_throws XLSX.XLSXError XLSX.setFont(sheet, "Sheet1!garbage"; size=18, name="Arial", color="FF040404")
        @test_throws XLSX.XLSXError XLSX.setFont(sheet, "garbage"; size=18, name="Arial", color="FF040404")
        @test_throws XLSX.XLSXError XLSX.setFont(sheet, "garbage1:garbage2"; size=18, name="Arial", color="FF040404")
        SAVE_FILES && save_outfile(f)

        f = XLSX.newxlsx()
        sheet = f[1]
        sheet["A1:E5"] = ""
        XLSX.setUniformFont(sheet, "Sheet1!A1:E1"; size=18, name="Arial", color="FF040404")
        @test XLSX.getFont(sheet, "D1").font == Dict("sz" => Dict("val" => "18"), "name" => Dict("val" => "Arial"), "color" => Dict("rgb" => "FF040404"))
        XLSX.setUniformFont(sheet, "Sheet1!2:3"; size=18, name="Arial", color="FF040408")
        @test XLSX.getFont(sheet, "A3").font == Dict("sz" => Dict("val" => "18"), "name" => Dict("val" => "Arial"), "color" => Dict("rgb" => "FF040408"))
        XLSX.setUniformFont(sheet, "Sheet1!D:E"; size=18, name="Arial", color="FF040400")
        @test XLSX.getFont(sheet, "E5").font == Dict("sz" => Dict("val" => "18"), "name" => Dict("val" => "Arial"), "color" => Dict("rgb" => "FF040400"))
        XLSX.setUniformFont(sheet, "A1:E1"; size=18, name="Arial", color="FF040304")
        @test XLSX.getFont(sheet, "D1").font == Dict("sz" => Dict("val" => "18"), "name" => Dict("val" => "Arial"), "color" => Dict("rgb" => "FF040304"))
        XLSX.setUniformFont(sheet, "2:3"; size=18, name="Arial", color="FF040308")
        @test XLSX.getFont(sheet, "A3").font == Dict("sz" => Dict("val" => "18"), "name" => Dict("val" => "Arial"), "color" => Dict("rgb" => "FF040308"))
        XLSX.setUniformFont(sheet, "D:E"; size=18, name="Arial", color="FF040300")
        @test XLSX.getFont(sheet, "E5").font == Dict("sz" => Dict("val" => "18"), "name" => Dict("val" => "Arial"), "color" => Dict("rgb" => "FF040300"))
        SAVE_FILES && save_outfile(f)

        f = XLSX.newxlsx()
        sheet = f[1]
        sheet["A1:E5"] = ""
        XLSX.setUniformFont(sheet, :, 1; size=18, name="Arial", color="FF040404")
        @test XLSX.getFont(sheet, "A3").font == Dict("sz" => Dict("val" => "18"), "name" => Dict("val" => "Arial"), "color" => Dict("rgb" => "FF040404"))
        XLSX.setUniformFont(sheet, :, [2, 3]; size=18, name="Arial", color="FF040400")
        @test XLSX.getFont(sheet, "C4").font == Dict("sz" => Dict("val" => "18"), "name" => Dict("val" => "Arial"), "color" => Dict("rgb" => "FF040400"))
        XLSX.setUniformFont(sheet, [1, 3, 4], 5; size=18, name="Arial", color="FF040300")
        @test XLSX.getFont(sheet, "E1").font == Dict("sz" => Dict("val" => "18"), "name" => Dict("val" => "Arial"), "color" => Dict("rgb" => "FF040300"))
        XLSX.setUniformFont(sheet, 5, [3, 4]; size=18, name="Arial", color="FF030300")
        @test XLSX.getFont(sheet, "D5").font == Dict("sz" => Dict("val" => "18"), "name" => Dict("val" => "Arial"), "color" => Dict("rgb" => "FF030300"))
        XLSX.setUniformFont(sheet, [2, 3, 4], [3, 4]; size=18, name="Arial", color="FF030308")
        @test XLSX.getFont(sheet, "C3").font == Dict("sz" => Dict("val" => "18"), "name" => Dict("val" => "Arial"), "color" => Dict("rgb" => "FF030308"))
        XLSX.setUniformFont(sheet, 4:5, 4; size=18, name="Arial", color="FF030408")
        @test XLSX.getFont(sheet, "D4").font == Dict("sz" => Dict("val" => "18"), "name" => Dict("val" => "Arial"), "color" => Dict("rgb" => "FF030408"))
        XLSX.setUniformFont(sheet, :; size=8, name="Arial", color="FF030408")
        @test XLSX.getFont(sheet, "D4").font == Dict("sz" => Dict("val" => "8"), "name" => Dict("val" => "Arial"), "color" => Dict("rgb" => "FF030408"))
        XLSX.setUniformFont(sheet, :, :; size=28, name="Arial", color="FF030408")
        @test XLSX.getFont(sheet, "D4").font == Dict("sz" => Dict("val" => "28"), "name" => Dict("val" => "Arial"), "color" => Dict("rgb" => "FF030408"))
        @test_throws XLSX.XLSXError XLSX.setUniformFont(sheet, :, [1, 3, 10, 15]; size=18, name="Arial", color="FF040404")
        @test_throws XLSX.XLSXError XLSX.setUniformFont(sheet, [1, 3, 10, 15], :; size=18, name="Arial", color="FF040404")
        @test_throws XLSX.XLSXError XLSX.setUniformFont(sheet, 1, [1, 3, 10, 15]; size=18, name="Arial", color="FF040404")
        @test_throws XLSX.XLSXError XLSX.setUniformFont(sheet, [1, 3, 10, 15], 2:3; size=18, name="Arial", color="FF040404")
        @test_throws XLSX.XLSXError XLSX.setUniformFont(f, "Sheet1!garbage"; size=18, name="Arial", color="FF040404")
        @test_throws XLSX.XLSXError XLSX.setUniformFont(sheet, "Sheet1!garbage"; size=18, name="Arial", color="FF040404")
        @test_throws XLSX.XLSXError XLSX.setUniformFont(sheet, "garbage"; size=18, name="Arial", color="FF040404")
        @test_throws XLSX.XLSXError XLSX.setUniformFont(sheet, "garbage1:garbage2"; size=18, name="Arial", color="FF040404")
        SAVE_FILES && save_outfile(f)
    end

    @testset "setBorder" begin
        f = XLSX.open_xlsx_template(joinpath(data_directory, "Borders.xlsx"))
        s = f["Sheet1"]

        @test XLSX.getDefaultBorders(s).border == Dict("left" => nothing, "bottom" => nothing, "right" => nothing, "top" => nothing, "diagonal" => nothing)

        @test isnothing(XLSX.getBorder(s, "A1"))

        @test XLSX.getBorder(s, "B2").border == Dict("left" => Dict("auto" => "1", "style" => "medium"), "bottom" => Dict("auto" => "1", "style" => "medium"), "right" => Dict("auto" => "1", "style" => "medium"), "top" => Dict("auto" => "1", "style" => "medium"), "diagonal" => nothing)
        @test XLSX.getBorder(f, "Sheet1!B4").border == Dict("left" => Dict("rgb" => "FFFF0000", "style" => "hair"), "bottom" => Dict("rgb" => "FFFF0000", "style" => "hair"), "right" => Dict("rgb" => "FFFF0000", "style" => "hair"), "top" => Dict("rgb" => "FFFF0000", "style" => "hair"), "diagonal" => nothing)
        @test XLSX.getBorder(s, "D4").border == Dict("left" => Dict("theme" => "3", "style" => "dashed", "tint" => "0.24994659260841701"), "bottom" => Dict("theme" => "3", "style" => "dashed", "tint" => "0.24994659260841701"), "right" => Dict("theme" => "3", "style" => "dashed", "tint" => "0.24994659260841701"), "top" => Dict("theme" => "3", "style" => "dashed", "tint" => "0.24994659260841701"), "diagonal" => nothing)
        @test XLSX.getBorder(f, "Sheet1!D6").border == Dict("left" => Dict("auto" => "1", "style" => "thick"), "bottom" => Dict("auto" => "1", "style" => "thick"), "right" => Dict("auto" => "1", "style" => "thick"), "top" => Dict("auto" => "1", "style" => "thick"), "diagonal" => nothing)

        XLSX.setBorder(f, "Sheet1!D6"; left=["style" => "dotted", "color" => "FF000FF0"], right=["style" => "medium", "color" => "FF765000"], top=["style" => "thick", "color" => "FF230000"], bottom=["style" => "medium", "color" => "FF0000FF"], diagonal=["style" => "dotted", "color" => "FF00D4D4"])
        @test XLSX.getBorder(s, "D6").border == Dict("left" => Dict("rgb" => "FF000FF0", "style" => "dotted"), "bottom" => Dict("rgb" => "FF0000FF", "style" => "medium"), "right" => Dict("rgb" => "FF765000", "style" => "medium"), "top" => Dict("rgb" => "FF230000", "style" => "thick"), "diagonal" => Dict("rgb" => "FF00D4D4", "style" => "dotted", "direction" => "both"))

        XLSX.setBorder(f, "Sheet1!B2:D4"; left=["style" => "hair"], right=["color" => "FF111111"], top=["style" => "hair"], bottom=["color" => "FF111111"], diagonal=["style" => "hair"])
        @test XLSX.getBorder(s, "B2").border == Dict("left" => Dict("auto" => "1", "style" => "hair"), "bottom" => Dict("rgb" => "FF111111", "style" => "medium"), "right" => Dict("rgb" => "FF111111", "style" => "medium"), "top" => Dict("auto" => "1", "style" => "hair"), "diagonal" => Dict("style" => "hair", "direction" => "both"))
        @test XLSX.getBorder(f, "Sheet1!D4").border == Dict("left" => Dict("theme" => "3", "style" => "hair", "tint" => "0.24994659260841701"), "bottom" => Dict("rgb" => "FF111111", "style" => "dashed"), "right" => Dict("rgb" => "FF111111", "style" => "dashed"), "top" => Dict("theme" => "3", "style" => "hair", "tint" => "0.24994659260841701"), "diagonal" => Dict("style" => "hair", "direction" => "both"))

        XLSX.setBorder(f, "Sheet1!A1:D10"; left=["style" => "hair", "color" => "FF111111"], right=["style" => "hair", "color" => "FF111111"], top=["style" => "hair", "color" => "FF111111"], bottom=["style" => "hair", "color" => "FF111111"], diagonal=["style" => "hair", "color" => "FF111111"])
        @test XLSX.getBorder(s, "B4").border == Dict("left" => Dict("rgb" => "FF111111", "style" => "hair"), "bottom" => Dict("rgb" => "FF111111", "style" => "hair"), "right" => Dict("rgb" => "FF111111", "style" => "hair"), "top" => Dict("rgb" => "FF111111", "style" => "hair"), "diagonal" => Dict("rgb" => "FF111111", "style" => "hair", "direction" => "both"))
        @test XLSX.getBorder(s, "B6").border == Dict("left" => Dict("rgb" => "FF111111", "style" => "hair"), "bottom" => Dict("rgb" => "FF111111", "style" => "hair"), "right" => Dict("rgb" => "FF111111", "style" => "hair"), "top" => Dict("rgb" => "FF111111", "style" => "hair"), "diagonal" => Dict("rgb" => "FF111111", "style" => "hair", "direction" => "both"))
        @test XLSX.getBorder(s, "D4").border == Dict("left" => Dict("rgb" => "FF111111", "style" => "hair"), "bottom" => Dict("rgb" => "FF111111", "style" => "hair"), "right" => Dict("rgb" => "FF111111", "style" => "hair"), "top" => Dict("rgb" => "FF111111", "style" => "hair"), "diagonal" => Dict("rgb" => "FF111111", "style" => "hair", "direction" => "both"))
        @test XLSX.getBorder(s, "D8").border == Dict("left" => Dict("rgb" => "FF111111", "style" => "hair"), "bottom" => Dict("rgb" => "FF111111", "style" => "hair"), "right" => Dict("rgb" => "FF111111", "style" => "hair"), "top" => Dict("rgb" => "FF111111", "style" => "hair"), "diagonal" => Dict("rgb" => "FF111111", "style" => "hair", "direction" => "up"))
        @test XLSX.getBorder(s, "A1").border == Dict("left" => Dict("rgb" => "FF111111", "style" => "hair"), "bottom" => Dict("rgb" => "FF111111", "style" => "hair"), "right" => Dict("rgb" => "FF111111", "style" => "hair"), "top" => Dict("rgb" => "FF111111", "style" => "hair"), "diagonal" => Dict("rgb" => "FF111111", "style" => "hair", "direction" => "both"))
        @test XLSX.getBorder(s, "D10").border == Dict("left" => Dict("rgb" => "FF111111", "style" => "hair"), "bottom" => Dict("rgb" => "FF111111", "style" => "hair"), "right" => Dict("rgb" => "FF111111", "style" => "hair"), "top" => Dict("rgb" => "FF111111", "style" => "hair"), "diagonal" => Dict("rgb" => "FF111111", "style" => "hair", "direction" => "both"))
        @test XLSX.getcell(s, "D11") isa XLSX.EmptyCell
        @test_throws XLSX.XLSXError XLSX.getBorder(s, "D11") # Cannot get a border outside sheet dimension.
        SAVE_FILES && save_outfile(f)

        f = XLSX.newxlsx()
        s = f[1]
        s[1:6, 1:6] = ""
        XLSX.setBorder(s, "B2:E5"; outside=["color" => "FFFF0000", "style" => "thick"])
        @test XLSX.getBorder(s, "B2").border == Dict("left" => Dict("rgb" => "FFFF0000", "style" => "thick"), "bottom" => nothing, "right" => nothing, "top" => Dict("rgb" => "FFFF0000", "style" => "thick"), "diagonal" => nothing)
        @test XLSX.getBorder(s, "B3").border == Dict("left" => Dict("rgb" => "FFFF0000", "style" => "thick"), "bottom" => nothing, "right" => nothing, "top" => nothing, "diagonal" => nothing)
        @test XLSX.getBorder(s, "B4").border == Dict("left" => Dict("rgb" => "FFFF0000", "style" => "thick"), "bottom" => nothing, "right" => nothing, "top" => nothing, "diagonal" => nothing)
        @test XLSX.getBorder(s, "B5").border == Dict("left" => Dict("rgb" => "FFFF0000", "style" => "thick"), "bottom" => Dict("rgb" => "FFFF0000", "style" => "thick"), "right" => nothing, "top" => nothing, "diagonal" => nothing)
        @test XLSX.getBorder(s, "C2").border == Dict("left" => nothing, "bottom" => nothing, "right" => nothing, "top" => Dict("rgb" => "FFFF0000", "style" => "thick"), "diagonal" => nothing)
        @test XLSX.getBorder(s, "C3") === nothing
        @test XLSX.getBorder(s, "C4") === nothing
        @test XLSX.getBorder(s, "C5").border == Dict("left" => nothing, "bottom" => Dict("rgb" => "FFFF0000", "style" => "thick"), "right" => nothing, "top" => nothing, "diagonal" => nothing)
        @test XLSX.getBorder(s, "D2").border == Dict("left" => nothing, "bottom" => nothing, "right" => nothing, "top" => Dict("rgb" => "FFFF0000", "style" => "thick"), "diagonal" => nothing)
        @test XLSX.getBorder(s, "D3") === nothing
        @test XLSX.getBorder(s, "D4") === nothing
        @test XLSX.getBorder(s, "D5").border == Dict("left" => nothing, "bottom" => Dict("rgb" => "FFFF0000", "style" => "thick"), "right" => nothing, "top" => nothing, "diagonal" => nothing)
        @test XLSX.getBorder(s, "E2").border == Dict("left" => nothing, "bottom" => nothing, "right" => Dict("rgb" => "FFFF0000", "style" => "thick"), "top" => Dict("rgb" => "FFFF0000", "style" => "thick"), "diagonal" => nothing)
        @test XLSX.getBorder(s, "E3").border == Dict("left" => nothing, "bottom" => nothing, "right" => Dict("rgb" => "FFFF0000", "style" => "thick"), "top" => nothing, "diagonal" => nothing)
        @test XLSX.getBorder(s, "E4").border == Dict("left" => nothing, "bottom" => nothing, "right" => Dict("rgb" => "FFFF0000", "style" => "thick"), "top" => nothing, "diagonal" => nothing)
        @test XLSX.getBorder(s, "E5").border == Dict("left" => nothing, "bottom" => Dict("rgb" => "FFFF0000", "style" => "thick"), "right" => Dict("rgb" => "FFFF0000", "style" => "thick"), "top" => nothing, "diagonal" => nothing)

        XLSX.setBorder(s, "B2:E5"; outside=["color" => "dodgerblue4"])
        @test XLSX.getBorder(s, "B2").border == Dict("left" => Dict("rgb" => "FF104E8B", "style" => "thick"), "bottom" => nothing, "right" => nothing, "top" => Dict("rgb" => "FF104E8B", "style" => "thick"), "diagonal" => nothing)
        @test XLSX.getBorder(s, "B3").border == Dict("left" => Dict("rgb" => "FF104E8B", "style" => "thick"), "bottom" => nothing, "right" => nothing, "top" => nothing, "diagonal" => nothing)
        @test XLSX.getBorder(s, "B4").border == Dict("left" => Dict("rgb" => "FF104E8B", "style" => "thick"), "bottom" => nothing, "right" => nothing, "top" => nothing, "diagonal" => nothing)
        @test XLSX.getBorder(s, "B5").border == Dict("left" => Dict("rgb" => "FF104E8B", "style" => "thick"), "bottom" => Dict("rgb" => "FF104E8B", "style" => "thick"), "right" => nothing, "top" => nothing, "diagonal" => nothing)
        @test XLSX.getBorder(s, "C2").border == Dict("left" => nothing, "bottom" => nothing, "right" => nothing, "top" => Dict("rgb" => "FF104E8B", "style" => "thick"), "diagonal" => nothing)
        @test XLSX.getBorder(s, "C3") === nothing
        @test XLSX.getBorder(s, "C4") === nothing
        SAVE_FILES && save_outfile(f)

        f = XLSX.newxlsx()
        s = f[1]
        s[1:6, 1:6] = ""
        XLSX.setBorder(s, "Sheet1!A1"; allsides=["color" => "FFFF00FF", "style" => "thick"])
        @test XLSX.getBorder(s, "A1").border == Dict("left" => Dict("rgb" => "FFFF00FF", "style" => "thick"), "bottom" => Dict("rgb" => "FFFF00FF", "style" => "thick"), "right" => Dict("rgb" => "FFFF00FF", "style" => "thick"), "top" => Dict("rgb" => "FFFF00FF", "style" => "thick"), "diagonal" => nothing)
        XLSX.setBorder(s, "Sheet1!A1:E1"; allsides=["color" => "FFFF0000", "style" => "thick"])
        @test XLSX.getBorder(s, "B1").border == Dict("left" => Dict("rgb" => "FFFF0000", "style" => "thick"), "bottom" => Dict("rgb" => "FFFF0000", "style" => "thick"), "right" => Dict("rgb" => "FFFF0000", "style" => "thick"), "top" => Dict("rgb" => "FFFF0000", "style" => "thick"), "diagonal" => nothing)
        XLSX.setBorder(s, "Sheet1!A:E"; left=["color" => "FFFF0001", "style" => "thick"])
        @test XLSX.getBorder(s, "B2").border == Dict("left" => Dict("rgb" => "FFFF0001", "style" => "thick"), "bottom" => nothing, "right" => nothing, "top" => nothing, "diagonal" => nothing)
        XLSX.setBorder(s, "Sheet1!3:4"; left=["color" => "FFFF0002", "style" => "thick"])
        @test XLSX.getBorder(s, "B4").border == Dict("left" => Dict("rgb" => "FFFF0002", "style" => "thick"), "bottom" => nothing, "right" => nothing, "top" => nothing, "diagonal" => nothing)
        XLSX.setBorder(s, "B2,B4"; left=["color" => "FFFF0004", "style" => "thick"])
        @test XLSX.getBorder(s, "B2").border == Dict("left" => Dict("rgb" => "FFFF0004", "style" => "thick"), "bottom" => nothing, "right" => nothing, "top" => nothing, "diagonal" => nothing)
        @test XLSX.getBorder(s, "B4").border == Dict("left" => Dict("rgb" => "FFFF0004", "style" => "thick"), "bottom" => nothing, "right" => nothing, "top" => nothing, "diagonal" => nothing)
        @test XLSX.getBorder(s, "B3").border == Dict("left" => Dict("rgb" => "FFFF0002", "style" => "thick"), "bottom" => nothing, "right" => nothing, "top" => nothing, "diagonal" => nothing)
        SAVE_FILES && save_outfile(f)

        f = XLSX.newxlsx()
        s = f[1]
        s[1:6, 1:6] = ""
        XLSX.setBorder(s, 1, :; left=["color" => "FFFF0001", "style" => "thick"])
        @test XLSX.getBorder(s, "B1").border == Dict("left" => Dict("rgb" => "FFFF0001", "style" => "thick"), "bottom" => nothing, "right" => nothing, "top" => nothing, "diagonal" => nothing)
        XLSX.setBorder(s, [2, 3], :; left=["color" => "FFFF0002", "style" => "thick"])
        @test XLSX.getBorder(s, "D3").border == Dict("left" => Dict("rgb" => "FFFF0002", "style" => "thick"), "bottom" => nothing, "right" => nothing, "top" => nothing, "diagonal" => nothing)
        XLSX.setBorder(s, :, [2, 3]; left=["color" => "FFFF0003", "style" => "thick"])
        @test XLSX.getBorder(s, "C4").border == Dict("left" => Dict("rgb" => "FFFF0003", "style" => "thick"), "bottom" => nothing, "right" => nothing, "top" => nothing, "diagonal" => nothing)
        XLSX.setBorder(s, 4, [2, 3]; left=["color" => "FFFF0004", "style" => "thick"])
        @test XLSX.getBorder(s, "B4").border == Dict("left" => Dict("rgb" => "FFFF0004", "style" => "thick"), "bottom" => nothing, "right" => nothing, "top" => nothing, "diagonal" => nothing)
        XLSX.setBorder(s, 3:2:5, [2, 3]; left=["color" => "FFFF0005", "style" => "thick"])
        @test XLSX.getBorder(s, "C5").border == Dict("left" => Dict("rgb" => "FFFF0005", "style" => "thick"), "bottom" => nothing, "right" => nothing, "top" => nothing, "diagonal" => nothing)
        @test_throws XLSX.XLSXError XLSX.setFont(s, "1:10"; left=["color" => "FFFF0005", "style" => "thick"])
        @test_throws XLSX.XLSXError XLSX.setFont(s, "A:K"; left=["color" => "FFFF0005", "style" => "thick"])
        SAVE_FILES && save_outfile(f)

        f = XLSX.open_xlsx_template(joinpath(data_directory, "Borders.xlsx"))
        s = f["Sheet1"]

        XLSX.setUniformBorder(f, "Sheet1!A1:D4"; left=["style" => "dotted", "color" => "darkseagreen3"],
            right=["style" => "medium", "color" => "FF765000"],
            top=["style" => "thick", "color" => "FF230000"],
            bottom=["style" => "medium", "color" => "FF0000FF"],
            diagonal=["style" => "none"]
        )
        @test XLSX.getBorder(s, "A1").border == Dict("left" => Dict("style" => "dotted", "rgb" => "FF9BCD9B"), "bottom" => Dict("style" => "medium", "rgb" => "FF0000FF"), "right" => Dict("style" => "medium", "rgb" => "FF765000"), "top" => Dict("style" => "thick", "rgb" => "FF230000"), "diagonal" => nothing)
        @test XLSX.getBorder(f, "Sheet1!B2").border == Dict("left" => Dict("style" => "dotted", "rgb" => "FF9BCD9B"), "bottom" => Dict("style" => "medium", "rgb" => "FF0000FF"), "right" => Dict("style" => "medium", "rgb" => "FF765000"), "top" => Dict("style" => "thick", "rgb" => "FF230000"), "diagonal" => nothing)
        @test XLSX.getBorder(f, "Sheet1!D4").border == Dict("left" => Dict("style" => "dotted", "rgb" => "FF9BCD9B"), "bottom" => Dict("style" => "medium", "rgb" => "FF0000FF"), "right" => Dict("style" => "medium", "rgb" => "FF765000"), "top" => Dict("style" => "thick", "rgb" => "FF230000"), "diagonal" => nothing)

        @test XLSX.getcell(s, "C3") isa XLSX.EmptyCell
        @test isnothing(XLSX.getBorder(s, "C3"))
        SAVE_FILES && save_outfile(f)

        f = XLSX.open_xlsx_template(joinpath(data_directory, "customXml.xlsx"))
        s = f["Mock-up"]

        # Sheet mismatch
        @test_throws XLSX.XLSXError XLSX.setUniformBorder(s, "Document History!A1:D4"; left=["style" => "dotted", "color" => "darkseagreen3"],
            right=["style" => "medium", "color" => "FF765000"],
            top=["style" => "thick", "color" => "FF230000"],
            bottom=["style" => "medium", "color" => "FF0000FF"],
            diagonal=["style" => "none"]
        )

        @test XLSX.setUniformBorder(s, "Mock-up!A1:B4,Mock-up!D4:E6"; left=["style" => "dotted", "color" => "darkseagreen3"],
            right=["style" => "medium", "color" => "FF765000"],
            top=["style" => "thick", "color" => "FF230000"],
            bottom=["style" => "medium", "color" => "FF0000FF"],
            diagonal=["style" => "none"]) == 28

        XLSX.setBorder(s, "ID"; left=["style" => "dotted", "color" => "grey36"], bottom=["style" => "medium", "color" => "FF0000FF"], right=["style" => "medium", "color" => "FF765000"], top=["style" => "thick", "color" => "FF230000"], diagonal=nothing)
        @test XLSX.getBorder(s, "ID").border == Dict("left" => Dict("style" => "dotted", "rgb" => "FF5C5C5C"), "bottom" => Dict("style" => "medium", "rgb" => "FF0000FF"), "right" => Dict("style" => "medium", "rgb" => "FF765000"), "top" => Dict("style" => "thick", "rgb" => "FF230000"), "diagonal" => nothing)

        # Location is a non-contiguous range
        XLSX.setBorder(s, "Location"; left=["style" => "hair", "color" => "chocolate4"], right=["style" => "hair", "color" => "chocolate4"], top=["style" => "hair", "color" => "chocolate4"], bottom=["style" => "hair", "color" => "chocolate4"], diagonal=["style" => "hair", "color" => "chocolate4"])
        @test XLSX.getBorder(s, "D18").border == Dict("left" => Dict("rgb" => "FF8B4513", "style" => "hair"), "bottom" => Dict("rgb" => "FF8B4513", "style" => "hair"), "right" => Dict("rgb" => "FF8B4513", "style" => "hair"), "top" => Dict("rgb" => "FF8B4513", "style" => "hair"), "diagonal" => Dict("rgb" => "FF8B4513", "style" => "hair", "direction" => "both"))
        @test XLSX.getBorder(s, "D20").border == Dict("left" => Dict("rgb" => "FF8B4513", "style" => "hair"), "bottom" => Dict("rgb" => "FF8B4513", "style" => "hair"), "right" => Dict("rgb" => "FF8B4513", "style" => "hair"), "top" => Dict("rgb" => "FF8B4513", "style" => "hair"), "diagonal" => Dict("rgb" => "FF8B4513", "style" => "hair", "direction" => "both"))
        @test XLSX.getBorder(s, "J18").border == Dict("left" => Dict("rgb" => "FF8B4513", "style" => "hair"), "bottom" => Dict("rgb" => "FF8B4513", "style" => "hair"), "right" => Dict("rgb" => "FF8B4513", "style" => "hair"), "top" => Dict("rgb" => "FF8B4513", "style" => "hair"), "diagonal" => Dict("rgb" => "FF8B4513", "style" => "hair", "direction" => "both"))
        @test XLSX.getBorder(s, "J20").border == Dict("left" => Dict("rgb" => "FF8B4513", "style" => "hair"), "bottom" => Dict("rgb" => "FF8B4513", "style" => "hair"), "right" => Dict("rgb" => "FF8B4513", "style" => "hair"), "top" => Dict("rgb" => "FF8B4513", "style" => "hair"), "diagonal" => Dict("rgb" => "FF8B4513", "style" => "hair", "direction" => "both"))

        # Can't get attributes on a range.
        @test_throws XLSX.XLSXError XLSX.getBorder(s, "Contiguous")
        SAVE_FILES && save_outfile(f)

        f = XLSX.open_empty_template()
        s = f["Sheet1"]
        s["A1"] = ""
        s["F21"] = ""
        # All these cells are empty.
        @test XLSX.setUniformFont(s, "A2:B4"; size=12, name="Times New Roman", color="chocolate4") == -1
        @test XLSX.setUniformBorder(f, "Sheet1!A2:D4"; left=["style" => "dotted", "color" => "chocolate4"],
            right=["style" => "medium", "color" => "FF765000"],
            top=["style" => "thick", "color" => "FF230000"],
            bottom=["style" => "medium", "color" => "chocolate4"],
            diagonal=["style" => "none"]
        ) == -1
        @test XLSX.setUniformFill(s, "B2:D4"; pattern="gray125", bgColor="FF000000") == -1
        @test XLSX.setFont(s, "A2:F20"; size=18, name="Arial") == -1
        @test XLSX.setBorder(f, "Sheet1!B2:D4"; left=["style" => "hair"], right=["color" => "FF8B4513"], top=["style" => "hair"], bottom=["color" => "chocolate4"], diagonal=["style" => "hair"]) == -1
        @test XLSX.setAlignment(s, "A2:F20"; horizontal="right", wrapText=true) == -1
        @test XLSX.setUniformFill(s, [2, 4], 2:4; pattern="gray125", bgColor="FF000000") == -1
        @test XLSX.setFont(s, [2, 4], 2:4; size=18, name="Arial") == -1
        @test XLSX.setBorder(s, [2, 4], :; left=["style" => "hair"], right=["color" => "FF8B4513"], top=["style" => "hair"], bottom=["color" => "chocolate4"], diagonal=["style" => "hair"]) == -1
        @test XLSX.setAlignment(s, [2, 4], 2:4; horizontal="right", wrapText=true) == -1
        @test XLSX.setUniformFill(s, "B2,C2"; pattern="gray125", bgColor="FF000000") == -1
        @test XLSX.setFont(s, "A2,A4"; size=18, name="Arial") == -1
        @test XLSX.setBorder(f, "Sheet1!B2,Sheet1!C2"; left=["style" => "hair"], right=["color" => "FF8B4513"], top=["style" => "hair"], bottom=["color" => "chocolate4"], diagonal=["style" => "hair"]) == -1
        @test XLSX.setAlignment(s, "A2,B3:C4"; horizontal="right", wrapText=true) == -1
        @test XLSX.setUniformAlignment(s, "B2,D2"; horizontal="right", wrapText=true) == -1
        @test XLSX.setUniformStyle(s, "B2:D2,E3") == -1
        @test_throws XLSX.XLSXError XLSX.setUniformFill(s, "B2,B2"; pattern="gray125", bgColor="FF000000")
        @test_throws XLSX.XLSXError XLSX.setFont(s, "A2,A2"; size=18, name="Arial")
        @test_throws XLSX.XLSXError XLSX.setBorder(f, "Sheet1!B2,Sheet1!B2"; left=["style" => "hair"], right=["color" => "FF8B4513"], top=["style" => "hair"], bottom=["color" => "chocolate4"], diagonal=["style" => "hair"])
        @test_throws XLSX.XLSXError XLSX.setAlignment(s, "A2,A2:A2"; horizontal="right", wrapText=true)
        @test_throws XLSX.XLSXError XLSX.setUniformAlignment(s, "B2,B2"; horizontal="right", wrapText=true)
        @test_throws XLSX.XLSXError XLSX.setUniformStyle(s, "B2:B2,B2")
        SAVE_FILES && save_outfile(f)

        f = XLSX.open_empty_template()
        s = f["Sheet1"]
        # All these cells are outside the sheet dimension.
        @test_throws XLSX.XLSXError XLSX.setUniformFont(s, "A1:B4"; size=12, name="Times New Roman", color="chocolate4")
        @test_throws XLSX.XLSXError XLSX.setUniformBorder(f, "Sheet1!A1:D4"; left=["style" => "dotted", "color" => "chocolate4"],
            right=["style" => "medium", "color" => "FF765000"],
            top=["style" => "thick", "color" => "FF230000"],
            bottom=["style" => "medium", "color" => "chocolate4"],
            diagonal=["style" => "none"]
        )
        @test_throws XLSX.XLSXError XLSX.setUniformFill(s, "B2:D4"; pattern="gray125", bgColor="FF000000")
        @test_throws XLSX.XLSXError XLSX.setFont(s, "A1:F20"; size=18, name="Arial")
        @test_throws XLSX.XLSXError XLSX.setBorder(f, "Sheet1!B2:D4"; left=["style" => "hair"], right=["color" => "FF8B4513"], top=["style" => "hair"], bottom=["color" => "chocolate4"], diagonal=["style" => "hair"])
        @test_throws XLSX.XLSXError XLSX.setAlignment(s, "A1:F20"; horizontal="right", wrapText=true)
        @test_throws XLSX.XLSXError XLSX.setFill(f, "Sheet1!A1"; pattern="none", fgColor="88FF8800")
        @test_throws XLSX.XLSXError XLSX.setFont(s, "A1"; size=18, name="Arial")
        @test_throws XLSX.XLSXError XLSX.setBorder(f, "Sheet1!B2"; left=["style" => "hair"], right=["color" => "FF8B4513"], top=["style" => "hair"], bottom=["color" => "FF8B4513"], diagonal=["style" => "hair"])
        @test_throws XLSX.XLSXError XLSX.setFill(s, "F20"; pattern="none", fgColor="88FF8800")
        SAVE_FILES && save_outfile(f)

        f = XLSX.open_xlsx_template(joinpath(data_directory, "customXml.xlsx"))
        s = f["Mock-up"]

        # Can't set a uniform attribute to a single cell.
        @test_throws MethodError XLSX.setUniformFill(s, "D4"; pattern="gray125", bgColor="FF000000")
        @test_throws MethodError XLSX.setUniformFill(s, "ID"; pattern="darkTrellis", fgColor="FF222222", bgColor="FFDDDDDD")
        @test_throws MethodError XLSX.setUniformFont(s, "B4"; size=12, name="Times New Roman", color="FF040404")
        @test_throws MethodError XLSX.setUniformBorder(f[2], "B4"; left=["style" => "hair"], right=["color" => "FF8B4513"], top=["style" => "hair"], bottom=["color" => "FF8B4513"], diagonal=["style" => "hair"])
        @test_throws MethodError XLSX.setUniformStyle(s, "ID")
        @test_throws MethodError XLSX.setUniformBorder(f, "Mock-up!D4"; left=["style" => "dotted", "color" => "FF000FF0"],
            right=["style" => "medium", "color" => "FF765000"],
            top=["style" => "thick", "color" => "FF230000"],
            bottom=["style" => "medium", "color" => "FF0000FF"],
            diagonal=["style" => "none"]
        )
        SAVE_FILES && save_outfile(f)

        f = XLSX.newxlsx()
        s = f[1]
        s[1:6, 1:6] = ""
        XLSX.setUniformBorder(s, "Sheet1!A:B";
            left=["style" => "dotted", "color" => "darkseagreen3"],
            right=["style" => "medium", "color" => "FF765000"],
            top=["style" => "thick", "color" => "FF230000"],
            bottom=["style" => "medium", "color" => "FF0000FF"],
            diagonal=["style" => "none"]
        )
        @test XLSX.getBorder(s, "A1").border == Dict("left" => Dict("style" => "dotted", "rgb" => "FF9BCD9B"), "bottom" => Dict("style" => "medium", "rgb" => "FF0000FF"), "right" => Dict("style" => "medium", "rgb" => "FF765000"), "top" => Dict("style" => "thick", "rgb" => "FF230000"), "diagonal" => nothing)
        XLSX.setUniformBorder(s, "Sheet1!2:4";
            left=["style" => "dotted", "color" => "FF9BCD9C"],
            right=["style" => "medium", "color" => "FF765000"],
            top=["style" => "thick", "color" => "FF230000"],
            bottom=["style" => "medium", "color" => "FF0000FF"],
            diagonal=["style" => "none"]
        )
        @test XLSX.getBorder(s, "C2").border == Dict("left" => Dict("style" => "dotted", "rgb" => "FF9BCD9C"), "bottom" => Dict("style" => "medium", "rgb" => "FF0000FF"), "right" => Dict("style" => "medium", "rgb" => "FF765000"), "top" => Dict("style" => "thick", "rgb" => "FF230000"), "diagonal" => nothing)
        XLSX.setUniformBorder(s, "A:B";
            left=["style" => "dotted", "color" => "FF9BCD9E"],
            right=["style" => "medium", "color" => "FF765000"],
            top=["style" => "thick", "color" => "FF230000"],
            bottom=["style" => "medium", "color" => "FF0000FF"],
            diagonal=["style" => "none"]
        )
        @test XLSX.getBorder(s, "B3").border == Dict("left" => Dict("style" => "dotted", "rgb" => "FF9BCD9E"), "bottom" => Dict("style" => "medium", "rgb" => "FF0000FF"), "right" => Dict("style" => "medium", "rgb" => "FF765000"), "top" => Dict("style" => "thick", "rgb" => "FF230000"), "diagonal" => nothing)
        XLSX.setUniformBorder(s, "2:4";
            left=["style" => "dotted", "color" => "FF9BCD9D"],
            right=["style" => "medium", "color" => "FF765000"],
            top=["style" => "thick", "color" => "FF230000"],
            bottom=["style" => "medium", "color" => "FF0000FF"],
            diagonal=["style" => "none"]
        )
        @test XLSX.getBorder(s, "C3").border == Dict("left" => Dict("style" => "dotted", "rgb" => "FF9BCD9D"), "bottom" => Dict("style" => "medium", "rgb" => "FF0000FF"), "right" => Dict("style" => "medium", "rgb" => "FF765000"), "top" => Dict("style" => "thick", "rgb" => "FF230000"), "diagonal" => nothing)
        XLSX.setUniformBorder(s, 5, :;
            left=["style" => "dotted", "color" => "FF9BCD8D"],
            right=["style" => "medium", "color" => "FF765000"],
            top=["style" => "thick", "color" => "FF230000"],
            bottom=["style" => "medium", "color" => "FF0000FF"],
            diagonal=["style" => "none"]
        )
        @test XLSX.getBorder(s, "F5").border == Dict("left" => Dict("style" => "dotted", "rgb" => "FF9BCD8D"), "bottom" => Dict("style" => "medium", "rgb" => "FF0000FF"), "right" => Dict("style" => "medium", "rgb" => "FF765000"), "top" => Dict("style" => "thick", "rgb" => "FF230000"), "diagonal" => nothing)
        XLSX.setUniformBorder(s, :, 5;
            left=["style" => "dotted", "color" => "FF9BBD8D"],
            right=["style" => "medium", "color" => "FF765000"],
            top=["style" => "thick", "color" => "FF230000"],
            bottom=["style" => "medium", "color" => "FF0000FF"],
            diagonal=["style" => "none"]
        )
        @test XLSX.getBorder(s, "E2").border == Dict("left" => Dict("style" => "dotted", "rgb" => "FF9BBD8D"), "bottom" => Dict("style" => "medium", "rgb" => "FF0000FF"), "right" => Dict("style" => "medium", "rgb" => "FF765000"), "top" => Dict("style" => "thick", "rgb" => "FF230000"), "diagonal" => nothing)
        XLSX.setUniformBorder(s, :, :;
            left=["style" => "dotted", "color" => "FF9BCD7D"],
            right=["style" => "medium", "color" => "FF765000"],
            top=["style" => "thick", "color" => "FF230000"],
            bottom=["style" => "medium", "color" => "FF0000FF"],
            diagonal=["style" => "none"]
        )
        @test XLSX.getBorder(s, "F5").border == Dict("left" => Dict("style" => "dotted", "rgb" => "FF9BCD7D"), "bottom" => Dict("style" => "medium", "rgb" => "FF0000FF"), "right" => Dict("style" => "medium", "rgb" => "FF765000"), "top" => Dict("style" => "thick", "rgb" => "FF230000"), "diagonal" => nothing)
        XLSX.setUniformBorder(s, :;
            left=["style" => "dotted", "color" => "FF9BCD6D"],
            right=["style" => "medium", "color" => "FF765000"],
            top=["style" => "thick", "color" => "FF230000"],
            bottom=["style" => "medium", "color" => "FF0000FF"],
            diagonal=["style" => "none"]
        )
        @test XLSX.getBorder(s, "D3").border == Dict("left" => Dict("style" => "dotted", "rgb" => "FF9BCD6D"), "bottom" => Dict("style" => "medium", "rgb" => "FF0000FF"), "right" => Dict("style" => "medium", "rgb" => "FF765000"), "top" => Dict("style" => "thick", "rgb" => "FF230000"), "diagonal" => nothing)
        XLSX.setUniformBorder(s, [2, 3], :;
            left=["style" => "dotted", "color" => "FF9BCE6D"],
            right=["style" => "medium", "color" => "FF765000"],
            top=["style" => "thick", "color" => "FF230000"],
            bottom=["style" => "medium", "color" => "FF0000FF"],
            diagonal=["style" => "none"]
        )
        @test XLSX.getBorder(s, "B2").border == Dict("left" => Dict("style" => "dotted", "rgb" => "FF9BCE6D"), "bottom" => Dict("style" => "medium", "rgb" => "FF0000FF"), "right" => Dict("style" => "medium", "rgb" => "FF765000"), "top" => Dict("style" => "thick", "rgb" => "FF230000"), "diagonal" => nothing)
        XLSX.setUniformBorder(s, :, [2, 3];
            left=["style" => "dotted", "color" => "FF9BCB6D"],
            right=["style" => "medium", "color" => "FF765000"],
            top=["style" => "thick", "color" => "FF230000"],
            bottom=["style" => "medium", "color" => "FF0000FF"],
            diagonal=["style" => "none"]
        )
        @test XLSX.getBorder(s, "C6").border == Dict("left" => Dict("style" => "dotted", "rgb" => "FF9BCB6D"), "bottom" => Dict("style" => "medium", "rgb" => "FF0000FF"), "right" => Dict("style" => "medium", "rgb" => "FF765000"), "top" => Dict("style" => "thick", "rgb" => "FF230000"), "diagonal" => nothing)
        XLSX.setUniformBorder(s, 1, [2, 3];
            left=["style" => "dotted", "color" => "FF8BCB6D"],
            right=["style" => "medium", "color" => "FF765000"],
            top=["style" => "thick", "color" => "FF230000"],
            bottom=["style" => "medium", "color" => "FF0000FF"],
            diagonal=["style" => "none"]
        )
        @test XLSX.getBorder(s, "C1").border == Dict("left" => Dict("style" => "dotted", "rgb" => "FF8BCB6D"), "bottom" => Dict("style" => "medium", "rgb" => "FF0000FF"), "right" => Dict("style" => "medium", "rgb" => "FF765000"), "top" => Dict("style" => "thick", "rgb" => "FF230000"), "diagonal" => nothing)
        XLSX.setUniformBorder(s, [1, 2], [4, 5, 6];
            left=["style" => "dotted", "color" => "FF6BCB6D"],
            right=["style" => "medium", "color" => "FF765000"],
            top=["style" => "thick", "color" => "FF230000"],
            bottom=["style" => "medium", "color" => "FF0000FF"],
            diagonal=["style" => "none"]
        )
        @test XLSX.getBorder(s, "E2").border == Dict("left" => Dict("style" => "dotted", "rgb" => "FF6BCB6D"), "bottom" => Dict("style" => "medium", "rgb" => "FF0000FF"), "right" => Dict("style" => "medium", "rgb" => "FF765000"), "top" => Dict("style" => "thick", "rgb" => "FF230000"), "diagonal" => nothing)
        XLSX.setUniformBorder(s, 4, 4;
            left=["style" => "dotted", "color" => "FF7BCB6D"],
            right=["style" => "medium", "color" => "FF765000"],
            top=["style" => "thick", "color" => "FF230000"],
            bottom=["style" => "medium", "color" => "FF0000FF"],
            diagonal=["style" => "none"]
        )
        @test XLSX.getBorder(s, "D4").border == Dict("left" => Dict("style" => "dotted", "rgb" => "FF7BCB6D"), "bottom" => Dict("style" => "medium", "rgb" => "FF0000FF"), "right" => Dict("style" => "medium", "rgb" => "FF765000"), "top" => Dict("style" => "thick", "rgb" => "FF230000"), "diagonal" => nothing)
        SAVE_FILES && save_outfile(f)

        f = XLSX.newxlsx()
        s = f[1]
        s[1:6, 1:6] = ""
        XLSX.setOutsideBorder(s, "Sheet1!A1:A2"; outside=["style" => "dotted", "color" => "FF003FF0"])
        @test XLSX.getBorder(s, "A1").border == Dict("left" => Dict("style" => "dotted", "rgb" => "FF003FF0"), "bottom" => nothing, "right" => Dict("style" => "dotted", "rgb" => "FF003FF0"), "top" => Dict("style" => "dotted", "rgb" => "FF003FF0"), "diagonal" => nothing)
        @test XLSX.getBorder(s, "A2").border == Dict("left" => Dict("style" => "dotted", "rgb" => "FF003FF0"), "bottom" => Dict("style" => "dotted", "rgb" => "FF003FF0"), "right" => Dict("style" => "dotted", "rgb" => "FF003FF0"), "top" => nothing, "diagonal" => nothing)
        XLSX.setOutsideBorder(s, "Sheet1!C:E"; outside=["style" => "dotted", "color" => "FF000FF0"])
        @test XLSX.getBorder(s, "C1").border == Dict("left" => Dict("style" => "dotted", "rgb" => "FF000FF0"), "bottom" => nothing, "right" => nothing, "top" => Dict("style" => "dotted", "rgb" => "FF000FF0"), "diagonal" => nothing)
        @test XLSX.getBorder(s, "E6").border == Dict("left" => nothing, "bottom" => Dict("style" => "dotted", "rgb" => "FF000FF0"), "right" => Dict("style" => "dotted", "rgb" => "FF000FF0"), "top" => nothing, "diagonal" => nothing)
        XLSX.setOutsideBorder(s, "Sheet1!3:5"; outside=["style" => "dotted", "color" => "FF000FFF"])
        @test XLSX.getBorder(s, "A3").border == Dict("left" => Dict("style" => "dotted", "rgb" => "FF000FFF"), "bottom" => nothing, "right" => nothing, "top" => Dict("style" => "dotted", "rgb" => "FF000FFF"), "diagonal" => nothing)
        @test XLSX.getBorder(s, "F5").border == Dict("left" => nothing, "bottom" => Dict("style" => "dotted", "rgb" => "FF000FFF"), "right" => Dict("style" => "dotted", "rgb" => "FF000FFF"), "top" => nothing, "diagonal" => nothing)
        XLSX.setOutsideBorder(s, "C:E"; outside=["style" => "dotted", "color" => "FFFF0FF0"])
        @test XLSX.getBorder(s, "C1").border == Dict("left" => Dict("style" => "dotted", "rgb" => "FFFF0FF0"), "bottom" => nothing, "right" => nothing, "top" => Dict("style" => "dotted", "rgb" => "FFFF0FF0"), "diagonal" => nothing)
        @test XLSX.getBorder(s, "E6").border == Dict("left" => nothing, "bottom" => Dict("style" => "dotted", "rgb" => "FFFF0FF0"), "right" => Dict("style" => "dotted", "rgb" => "FFFF0FF0"), "top" => nothing, "diagonal" => nothing)
        XLSX.setOutsideBorder(s, "3:5"; outside=["style" => "dotted", "color" => "FFF50FFF"])
        @test XLSX.getBorder(s, "A3").border == Dict("left" => Dict("style" => "dotted", "rgb" => "FFF50FFF"), "bottom" => nothing, "right" => nothing, "top" => Dict("style" => "dotted", "rgb" => "FFF50FFF"), "diagonal" => nothing)
        @test XLSX.getBorder(s, "F5").border == Dict("left" => nothing, "bottom" => Dict("style" => "dotted", "rgb" => "FFF50FFF"), "right" => Dict("style" => "dotted", "rgb" => "FFF50FFF"), "top" => nothing, "diagonal" => nothing)
        SAVE_FILES && save_outfile(f)

        f = XLSX.newxlsx()
        s = f[1]
        s[1:6, 1:6] = ""
        XLSX.setOutsideBorder(s, 1, :; outside=["style" => "dotted", "color" => "FF002FF0"])
        @test XLSX.getBorder(s, "A1").border == Dict("left" => Dict("style" => "dotted", "rgb" => "FF002FF0"), "bottom" => Dict("style" => "dotted", "rgb" => "FF002FF0"), "right" => nothing, "top" => Dict("style" => "dotted", "rgb" => "FF002FF0"), "diagonal" => nothing)
        @test XLSX.getBorder(s, "F1").border == Dict("left" => nothing, "bottom" => Dict("style" => "dotted", "rgb" => "FF002FF0"), "right" => Dict("style" => "dotted", "rgb" => "FF002FF0"), "top" => Dict("style" => "dotted", "rgb" => "FF002FF0"), "diagonal" => nothing)
        XLSX.setOutsideBorder(s, :, 1; outside=["style" => "dotted", "color" => "FF003FF0"])
        @test XLSX.getBorder(s, "A1").border == Dict("left" => Dict("rgb" => "FF003FF0", "style" => "dotted"), "bottom" => Dict("rgb" => "FF002FF0", "style" => "dotted"), "right" => Dict("rgb" => "FF003FF0", "style" => "dotted"), "top" => Dict("rgb" => "FF003FF0", "style" => "dotted"), "diagonal" => nothing)
        @test XLSX.getBorder(s, "A6").border == Dict("left" => Dict("style" => "dotted", "rgb" => "FF003FF0"), "bottom" => Dict("style" => "dotted", "rgb" => "FF003FF0"), "right" => Dict("style" => "dotted", "rgb" => "FF003FF0"), "top" => nothing, "diagonal" => nothing)
        XLSX.setOutsideBorder(s, :, :; outside=["style" => "dotted", "color" => "FF000FF0"])
        @test XLSX.getBorder(s, "A1").border == Dict("left" => Dict("rgb" => "FF000FF0", "style" => "dotted"), "bottom" => Dict("rgb" => "FF002FF0", "style" => "dotted"), "right" => Dict("rgb" => "FF003FF0", "style" => "dotted"), "top" => Dict("rgb" => "FF000FF0", "style" => "dotted"), "diagonal" => nothing)
        @test XLSX.getBorder(s, "F6").border == Dict("left" => nothing, "bottom" => Dict("style" => "dotted", "rgb" => "FF000FF0"), "right" => Dict("style" => "dotted", "rgb" => "FF000FF0"), "top" => nothing, "diagonal" => nothing)
        XLSX.setOutsideBorder(s, :; outside=["style" => "dotted", "color" => "FF000FFF"])
        @test XLSX.getBorder(s, "A1").border == Dict("left" => Dict("rgb" => "FF000FFF", "style" => "dotted"), "bottom" => Dict("rgb" => "FF002FF0", "style" => "dotted"), "right" => Dict("rgb" => "FF003FF0", "style" => "dotted"), "top" => Dict("rgb" => "FF000FFF", "style" => "dotted"), "diagonal" => nothing)
        @test XLSX.getBorder(s, "F6").border == Dict("left" => nothing, "bottom" => Dict("style" => "dotted", "rgb" => "FF000FFF"), "right" => Dict("style" => "dotted", "rgb" => "FF000FFF"), "top" => nothing, "diagonal" => nothing)
        XLSX.setOutsideBorder(s, 1:2, 1; outside=["style" => "dotted", "color" => "FFFFFFF0"])
        @test XLSX.getBorder(s, "A1").border == Dict("left" => Dict("rgb" => "FFFFFFF0", "style" => "dotted"), "bottom" => Dict("rgb" => "FF002FF0", "style" => "dotted"), "right" => Dict("rgb" => "FFFFFFF0", "style" => "dotted"), "top" => Dict("rgb" => "FFFFFFF0", "style" => "dotted"), "diagonal" => nothing)
        @test XLSX.getBorder(s, "A2").border == Dict("left" => Dict("rgb" => "FFFFFFF0", "style" => "dotted"), "bottom" => Dict("rgb" => "FFFFFFF0", "style" => "dotted"), "right" => Dict("rgb" => "FFFFFFF0", "style" => "dotted"), "top" => nothing, "diagonal" => nothing)
        SAVE_FILES && save_outfile(f)

    end

    @testset "setFill" begin

        f = XLSX.open_xlsx_template(joinpath(data_directory, "customXml.xlsx"))
        s = f["Mock-up"]

        @test XLSX.getDefaultFill(s).fill == Dict("patternFill" => Dict("patternType" => "none"))

        @test XLSX.getFill(s, "D17").fill == Dict("patternFill" => Dict("bgindexed" => "64", "patternType" => "solid", "fgtint" => "-9.9978637043366805E-2", "fgtheme" => "2"))
        @test XLSX.getFill(f, "Mock-up!D18").fill == Dict("patternFill" => Dict("bgindexed" => "64", "patternType" => "solid", "fgtint" => "-0.499984740745262", "fgtheme" => "2"))

        XLSX.setFill(s, "D17"; pattern="darkTrellis", fgColor="FF222222", bgColor="FFDDDDDD")
        @test XLSX.getFill(s, "D17").fill == Dict("patternFill" => Dict("bgrgb" => "FFDDDDDD", "patternType" => "darkTrellis", "fgrgb" => "FF222222"))

        XLSX.setFill(s, "ID"; pattern="darkTrellis", fgColor="FF222222", bgColor="FFDDDDDD")
        @test XLSX.getFill(s, "ID").fill == Dict("patternFill" => Dict("bgrgb" => "FFDDDDDD", "patternType" => "darkTrellis", "fgrgb" => "FF222222"))
        @test_throws XLSX.XLSXError XLSX.setFill(s, "ID"; pattern="darkTrellis", fgColor="notAcolor", bgColor="FFDDDDDD")
        @test_throws XLSX.XLSXError XLSX.setFill(s, "ID"; pattern="notApattern", fgColor="FF222222", bgColor="FFDDDDDD")
        @test_throws XLSX.XLSXError XLSX.setFill(s, "ID"; pattern="darkTrellis", fgColor="FF222222", bgColor="FFDDDDDDFF")
        @test_throws XLSX.XLSXError XLSX.setFill(s, "ID"; fgColor="FF222222", bgColor="FFDDDDDDFF")

        # Location is a non-contiguous range
        XLSX.setFill(s, "Location"; pattern="lightVertical") # Default colors unchanged
        @test XLSX.getFill(s, "D18").fill == Dict("patternFill" => Dict("bgindexed" => "64", "patternType" => "lightVertical", "fgtint" => "-0.499984740745262", "fgtheme" => "2"))
        @test XLSX.getFill(s, "D20").fill == Dict("patternFill" => Dict("bgindexed" => "64", "patternType" => "lightVertical", "fgtint" => "-0.499984740745262", "fgtheme" => "2"))
        @test XLSX.getFill(s, "J18").fill == Dict("patternFill" => Dict("bgindexed" => "64", "patternType" => "lightVertical", "fgtint" => "-0.499984740745262", "fgtheme" => "2"))
        @test XLSX.getFill(s, "J18").fill == Dict("patternFill" => Dict("bgindexed" => "64", "patternType" => "lightVertical", "fgtint" => "-0.499984740745262", "fgtheme" => "2"))

        XLSX.setFill(s, "Contiguous"; pattern="lightVertical")  # Default colors unchanged
        @test XLSX.getFill(s, "D23").fill == Dict("patternFill" => Dict("patternType" => "lightVertical", "bgindexed" => "64", "fgtheme" => "0"))
        @test XLSX.getFill(s, "D24").fill == Dict("patternFill" => Dict("patternType" => "lightVertical", "bgindexed" => "64", "fgtheme" => "0"))
        @test XLSX.getFill(s, "D25").fill == Dict("patternFill" => Dict("patternType" => "lightVertical", "bgindexed" => "64", "fgtheme" => "0"))
        @test XLSX.getFill(s, "D26").fill == Dict("patternFill" => Dict("patternType" => "lightVertical", "bgindexed" => "64", "fgtheme" => "0"))
        @test XLSX.getFill(s, "D27").fill == Dict("patternFill" => Dict("patternType" => "lightVertical", "bgindexed" => "64", "fgtheme" => "0"))

        # Can't get attributes on a range.
        @test_throws XLSX.XLSXError XLSX.getFill(s, "Contiguous")

        XLSX.setUniformFill(s, "B3:D5"; pattern="lightGrid", fgColor="FF0000FF", bgColor="FF00FF00")
        @test XLSX.getFill(s, "B3").fill == Dict("patternFill" => Dict("bgrgb" => "FF00FF00", "patternType" => "lightGrid", "fgrgb" => "FF0000FF"))
        @test XLSX.getFill(s, "C4").fill == Dict("patternFill" => Dict("bgrgb" => "FF00FF00", "patternType" => "lightGrid", "fgrgb" => "FF0000FF"))
        @test XLSX.getFill(s, "D5").fill == Dict("patternFill" => Dict("bgrgb" => "FF00FF00", "patternType" => "lightGrid", "fgrgb" => "FF0000FF"))

        XLSX.setFill(s, "ID"; pattern="darkTrellis", fgColor="FF222222", bgColor="FFDDDDDD")
        @test XLSX.getFill(s, "ID").fill == Dict("patternFill" => Dict("bgrgb" => "FFDDDDDD", "patternType" => "darkTrellis", "fgrgb" => "FF222222"))

        # Location is a non-contiguous range
        XLSX.setFill(s, "Location"; pattern="lightVertical")
        @test XLSX.getFill(s, "D18").fill == Dict("patternFill" => Dict("bgindexed" => "64", "patternType" => "lightVertical", "fgtint" => "-0.499984740745262", "fgtheme" => "2"))
        @test XLSX.getFill(f, "Mock-up!D20").fill == Dict("patternFill" => Dict("bgindexed" => "64", "patternType" => "lightVertical", "fgtint" => "-0.499984740745262", "fgtheme" => "2"))
        @test XLSX.getFill(s, "J18").fill == Dict("patternFill" => Dict("bgindexed" => "64", "patternType" => "lightVertical", "fgtint" => "-0.499984740745262", "fgtheme" => "2"))
        @test XLSX.getFill(f, "Mock-up!J18").fill == Dict("patternFill" => Dict("bgindexed" => "64", "patternType" => "lightVertical", "fgtint" => "-0.499984740745262", "fgtheme" => "2"))

        XLSX.setFill(s, "Contiguous"; pattern="lightVertical")
        @test XLSX.getFill(s, "D23").fill == Dict("patternFill" => Dict("patternType" => "lightVertical", "bgindexed" => "64", "fgtheme" => "0"))
        @test XLSX.getFill(f, "Mock-up!D24").fill == Dict("patternFill" => Dict("patternType" => "lightVertical", "bgindexed" => "64", "fgtheme" => "0"))
        @test XLSX.getFill(s, "D25").fill == Dict("patternFill" => Dict("patternType" => "lightVertical", "bgindexed" => "64", "fgtheme" => "0"))
        @test XLSX.getFill(f, "Mock-up!D26").fill == Dict("patternFill" => Dict("patternType" => "lightVertical", "bgindexed" => "64", "fgtheme" => "0"))
        @test XLSX.getFill(s, "D27").fill == Dict("patternFill" => Dict("patternType" => "lightVertical", "bgindexed" => "64", "fgtheme" => "0"))

        # Can't get attributes on a range.
        @test_throws XLSX.XLSXError XLSX.getFill(s, "Contiguous")

        XLSX.writexlsx("output.xlsx", f, overwrite=true)
        SAVE_FILES && save_outfile("output.xlsx")
        @test isfile("output.xlsx")

        XLSX.openxlsx("output.xlsx") do f # Check the updated fonts were written correctly
            s = f["Mock-up"]
            @test XLSX.getFill(s, "D23").fill == Dict("patternFill" => Dict("patternType" => "lightVertical", "bgindexed" => "64", "fgtheme" => "0"))
            @test XLSX.getFill(f, "Mock-up!D24").fill == Dict("patternFill" => Dict("patternType" => "lightVertical", "bgindexed" => "64", "fgtheme" => "0"))
            @test XLSX.getFill(s, "D25").fill == Dict("patternFill" => Dict("patternType" => "lightVertical", "bgindexed" => "64", "fgtheme" => "0"))
            @test XLSX.getFill(f, "Mock-up!D26").fill == Dict("patternFill" => Dict("patternType" => "lightVertical", "bgindexed" => "64", "fgtheme" => "0"))
            @test XLSX.getFill(s, "D27").fill == Dict("patternFill" => Dict("patternType" => "lightVertical", "bgindexed" => "64", "fgtheme" => "0"))
        end

        isfile("output.xlsx") && rm("output.xlsx")
        SAVE_FILES && save_outfile(f)

        f = XLSX.newxlsx()
        s = f[1]
        s[1:6, 1:6] = ""
        XLSX.setFill(s, "Sheet1!A1"; pattern="darkTrellis", fgColor="FF222222", bgColor="FFDDDDDD")
        @test XLSX.getFill(s, "A1").fill == Dict("patternFill" => Dict("bgrgb" => "FFDDDDDD", "patternType" => "darkTrellis", "fgrgb" => "FF222222"))
        XLSX.setFill(s, "Sheet1!A2:F2"; pattern="darkTrellis", fgColor="FF222224", bgColor="FFDDDDD4")
        @test XLSX.getFill(s, "A2").fill == Dict("patternFill" => Dict("bgrgb" => "FFDDDDD4", "patternType" => "darkTrellis", "fgrgb" => "FF222224"))
        XLSX.setFill(s, "Sheet1!C:D"; pattern="darkTrellis", fgColor="FF222228", bgColor="FFDDDDD8")
        @test XLSX.getFill(s, "D4").fill == Dict("patternFill" => Dict("bgrgb" => "FFDDDDD8", "patternType" => "darkTrellis", "fgrgb" => "FF222228"))
        XLSX.setFill(s, "Sheet1!5:6"; pattern="darkTrellis", fgColor="FF222220", bgColor="FFDDDDD0")
        @test XLSX.getFill(s, "F5").fill == Dict("patternFill" => Dict("bgrgb" => "FFDDDDD0", "patternType" => "darkTrellis", "fgrgb" => "FF222220"))
        XLSX.setFill(s, "Sheet1!E4:E6,Sheet1!A4"; pattern="darkTrellis", fgColor="FF422220", bgColor="FF4DDDD0")
        @test XLSX.getFill(s, "E4").fill == Dict("patternFill" => Dict("bgrgb" => "FF4DDDD0", "patternType" => "darkTrellis", "fgrgb" => "FF422220"))
        @test XLSX.getFill(s, "E5").fill == Dict("patternFill" => Dict("bgrgb" => "FF4DDDD0", "patternType" => "darkTrellis", "fgrgb" => "FF422220"))
        @test XLSX.getFill(s, "E6").fill == Dict("patternFill" => Dict("bgrgb" => "FF4DDDD0", "patternType" => "darkTrellis", "fgrgb" => "FF422220"))
        @test XLSX.getFill(s, "A4").fill == Dict("patternFill" => Dict("bgrgb" => "FF4DDDD0", "patternType" => "darkTrellis", "fgrgb" => "FF422220"))
        XLSX.setFill(s, :, 2; pattern="darkTrellis", fgColor="FF622220", bgColor="FF6DDDD0")
        @test XLSX.getFill(s, "B4").fill == Dict("patternFill" => Dict("bgrgb" => "FF6DDDD0", "patternType" => "darkTrellis", "fgrgb" => "FF622220"))
        XLSX.setFill(s, [2, 6], :; pattern="darkTrellis", fgColor="FF622222", bgColor="FF6DDDD2")
        @test XLSX.getFill(s, "C2").fill == Dict("patternFill" => Dict("bgrgb" => "FF6DDDD2", "patternType" => "darkTrellis", "fgrgb" => "FF622222"))
        @test XLSX.getFill(s, "F6").fill == Dict("patternFill" => Dict("bgrgb" => "FF6DDDD2", "patternType" => "darkTrellis", "fgrgb" => "FF622222"))
        XLSX.setFill(s, :, [2, 5]; pattern="darkTrellis", fgColor="FF622224", bgColor="FF6DDDD4")
        @test XLSX.getFill(s, "B2").fill == Dict("patternFill" => Dict("bgrgb" => "FF6DDDD4", "patternType" => "darkTrellis", "fgrgb" => "FF622224"))
        @test XLSX.getFill(s, "E3").fill == Dict("patternFill" => Dict("bgrgb" => "FF6DDDD4", "patternType" => "darkTrellis", "fgrgb" => "FF622224"))
        XLSX.setFill(s, 2, [3, 6]; pattern="darkTrellis", fgColor="FF622226", bgColor="FF6DDDD6")
        @test XLSX.getFill(s, "C2").fill == Dict("patternFill" => Dict("bgrgb" => "FF6DDDD6", "patternType" => "darkTrellis", "fgrgb" => "FF622226"))
        XLSX.setFill(s, 2:2:6, [4, 5]; pattern="darkTrellis", fgColor="FF622228", bgColor="FF6DDDD8")
        @test XLSX.getFill(s, "E4").fill == Dict("patternFill" => Dict("bgrgb" => "FF6DDDD8", "patternType" => "darkTrellis", "fgrgb" => "FF622228"))
        SAVE_FILES && save_outfile(f)

        f = XLSX.newxlsx()
        s = f[1]
        s[1:6, 1:6] = ""
        XLSX.setUniformFill(s, "Sheet1!A2:F2"; pattern="darkTrellis", fgColor="FF222224", bgColor="FFDDDDD4")
        @test XLSX.getFill(s, "A2").fill == Dict("patternFill" => Dict("bgrgb" => "FFDDDDD4", "patternType" => "darkTrellis", "fgrgb" => "FF222224"))
        XLSX.setUniformFill(s, "Sheet1!C:D"; pattern="darkTrellis", fgColor="FF222228", bgColor="FFDDDDD8")
        @test XLSX.getFill(s, "D4").fill == Dict("patternFill" => Dict("bgrgb" => "FFDDDDD8", "patternType" => "darkTrellis", "fgrgb" => "FF222228"))
        XLSX.setUniformFill(s, "Sheet1!5:6"; pattern="darkTrellis", fgColor="FF222220", bgColor="FFDDDDD0")
        @test XLSX.getFill(s, "F5").fill == Dict("patternFill" => Dict("bgrgb" => "FFDDDDD0", "patternType" => "darkTrellis", "fgrgb" => "FF222220"))
        XLSX.setUniformFill(s, "A2:F2"; pattern="darkTrellis", fgColor="FF222224", bgColor="FFDDDDD4")
        @test XLSX.getFill(s, "A2").fill == Dict("patternFill" => Dict("bgrgb" => "FFDDDDD4", "patternType" => "darkTrellis", "fgrgb" => "FF222224"))
        XLSX.setUniformFill(s, "C:D"; pattern="darkTrellis", fgColor="FF222228", bgColor="FFDDDDD8")
        @test XLSX.getFill(s, "D4").fill == Dict("patternFill" => Dict("bgrgb" => "FFDDDDD8", "patternType" => "darkTrellis", "fgrgb" => "FF222228"))
        XLSX.setUniformFill(s, "5:6"; pattern="darkTrellis", fgColor="FF222220", bgColor="FFDDDDD0")
        @test XLSX.getFill(s, "F5").fill == Dict("patternFill" => Dict("bgrgb" => "FFDDDDD0", "patternType" => "darkTrellis", "fgrgb" => "FF222220"))
        XLSX.setUniformFill(s, "E4:E6,A4"; pattern="darkTrellis", fgColor="FF422220", bgColor="FF4DDDD0")
        @test XLSX.getFill(s, "E4").fill == Dict("patternFill" => Dict("bgrgb" => "FF4DDDD0", "patternType" => "darkTrellis", "fgrgb" => "FF422220"))
        @test XLSX.getFill(s, "E5").fill == Dict("patternFill" => Dict("bgrgb" => "FF4DDDD0", "patternType" => "darkTrellis", "fgrgb" => "FF422220"))
        @test XLSX.getFill(s, "E6").fill == Dict("patternFill" => Dict("bgrgb" => "FF4DDDD0", "patternType" => "darkTrellis", "fgrgb" => "FF422220"))
        @test XLSX.getFill(s, "A4").fill == Dict("patternFill" => Dict("bgrgb" => "FF4DDDD0", "patternType" => "darkTrellis", "fgrgb" => "FF422220"))
        XLSX.setUniformFill(s, :, 2; pattern="darkTrellis", fgColor="FF622220", bgColor="FF6DDDD0")
        @test XLSX.getFill(s, "B4").fill == Dict("patternFill" => Dict("bgrgb" => "FF6DDDD0", "patternType" => "darkTrellis", "fgrgb" => "FF622220"))
        XLSX.setUniformFill(s, [2, 6], :; pattern="darkTrellis", fgColor="FF622222", bgColor="FF6DDDD2")
        @test XLSX.getFill(s, "C2").fill == Dict("patternFill" => Dict("bgrgb" => "FF6DDDD2", "patternType" => "darkTrellis", "fgrgb" => "FF622222"))
        @test XLSX.getFill(s, "F6").fill == Dict("patternFill" => Dict("bgrgb" => "FF6DDDD2", "patternType" => "darkTrellis", "fgrgb" => "FF622222"))
        XLSX.setUniformFill(s, :, [2, 5]; pattern="darkTrellis", fgColor="FF622224", bgColor="FF6DDDD4")
        @test XLSX.getFill(s, "B2").fill == Dict("patternFill" => Dict("bgrgb" => "FF6DDDD4", "patternType" => "darkTrellis", "fgrgb" => "FF622224"))
        @test XLSX.getFill(s, "E3").fill == Dict("patternFill" => Dict("bgrgb" => "FF6DDDD4", "patternType" => "darkTrellis", "fgrgb" => "FF622224"))
        XLSX.setUniformFill(s, 2, [3, 6]; pattern="darkTrellis", fgColor="FF622226", bgColor="FF6DDDD6")
        @test XLSX.getFill(s, "C2").fill == Dict("patternFill" => Dict("bgrgb" => "FF6DDDD6", "patternType" => "darkTrellis", "fgrgb" => "FF622226"))
        XLSX.setUniformFill(s, [2, 3], 5:6; pattern="darkTrellis", fgColor="FF642226", bgColor="FF64DDD6")
        @test XLSX.getFill(s, "F2").fill == Dict("patternFill" => Dict("bgrgb" => "FF64DDD6", "patternType" => "darkTrellis", "fgrgb" => "FF642226"))
        XLSX.setUniformFill(s, 2:2:6, [4, 5]; pattern="darkTrellis", fgColor="FF622228", bgColor="FF6DDDD8")
        @test XLSX.getFill(s, "E4").fill == Dict("patternFill" => Dict("bgrgb" => "FF6DDDD8", "patternType" => "darkTrellis", "fgrgb" => "FF622228"))
        XLSX.setUniformFill(s, :, :; pattern="darkTrellis", fgColor="FF822228", bgColor="FF8DDDD8")
        @test XLSX.getFill(s, "E4").fill == Dict("patternFill" => Dict("bgrgb" => "FF8DDDD8", "patternType" => "darkTrellis", "fgrgb" => "FF822228"))
        XLSX.setUniformFill(s, :; pattern="darkTrellis", fgColor="FF822288", bgColor="FF8DDD88")
        @test XLSX.getFill(s, "E4").fill == Dict("patternFill" => Dict("bgrgb" => "FF8DDD88", "patternType" => "darkTrellis", "fgrgb" => "FF822288"))
        XLSX.setUniformFill(s, :; pattern="darkTrellis", fgColor="FF822288", bgColor="FF8DDD88")
        @test XLSX.getFill(s, "E4").fill == Dict("patternFill" => Dict("bgrgb" => "FF8DDD88", "patternType" => "darkTrellis", "fgrgb" => "FF822288"))
        XLSX.setUniformFill(s, 1, 1:2; pattern="darkTrellis", fgColor="FF822268", bgColor="FF8DDD68")
        @test XLSX.getFill(s, "B1").fill == Dict("patternFill" => Dict("bgrgb" => "FF8DDD68", "patternType" => "darkTrellis", "fgrgb" => "FF822268"))
        SAVE_FILES && save_outfile(f)

    end

    @testset "setAlignment" begin

        f = XLSX.newxlsx()
        s = f[1]
        s["A1:Z26"] = ""
        XLSX.setAlignment(s, "Sheet1!A1"; horizontal="right", vertical="justify", wrapText=true)
        @test XLSX.getAlignment(s, "A1").alignment == Dict("alignment" => Dict("horizontal" => "right", "vertical" => "justify", "wrapText" => "1"))
        XLSX.setAlignment(s, "Sheet1!A2:C4"; horizontal="right", vertical="justify", wrapText=true)
        @test XLSX.getAlignment(s, "B3").alignment == Dict("alignment" => Dict("horizontal" => "right", "vertical" => "justify", "wrapText" => "1"))
        XLSX.setAlignment(s, "Sheet1!D:E"; horizontal="right", vertical="justify", wrapText=true)
        @test XLSX.getAlignment(s, "D26").alignment == Dict("alignment" => Dict("horizontal" => "right", "vertical" => "justify", "wrapText" => "1"))
        XLSX.setAlignment(s, "Sheet1!25:26"; horizontal="left", vertical="top", wrapText=false)
        @test XLSX.getAlignment(s, "D26").alignment == Dict("alignment" => Dict("horizontal" => "left", "vertical" => "top", "wrapText" => "0"))
        XLSX.setAlignment(s, "G8,H10,J15:M18"; horizontal="left", vertical="bottom", wrapText=true)
        @test XLSX.getAlignment(s, "G8").alignment == Dict("alignment" => Dict("horizontal" => "left", "vertical" => "bottom", "wrapText" => "1"))
        @test XLSX.getAlignment(s, "H10").alignment == Dict("alignment" => Dict("horizontal" => "left", "vertical" => "bottom", "wrapText" => "1"))
        @test XLSX.getAlignment(s, "L16").alignment == Dict("alignment" => Dict("horizontal" => "left", "vertical" => "bottom", "wrapText" => "1"))
        SAVE_FILES && save_outfile(f)

        f = XLSX.newxlsx()
        s = f[1]
        s["A1:Z26"] = ""
        XLSX.setAlignment(s, :, 1:3; horizontal="right", vertical="justify", wrapText=true)
        @test XLSX.getAlignment(s, "B25").alignment == Dict("alignment" => Dict("horizontal" => "right", "vertical" => "justify", "wrapText" => "1"))
        XLSX.setAlignment(s, 8:2:16, :; horizontal="right", vertical="justify", wrapText=true)
        @test XLSX.getAlignment(s, "C12").alignment == Dict("alignment" => Dict("horizontal" => "right", "vertical" => "justify", "wrapText" => "1"))
        XLSX.setAlignment(s, :, [8, 10, 12, 14, 16]; horizontal="right", vertical="top", wrapText=true)
        @test XLSX.getAlignment(s, "L22").alignment == Dict("alignment" => Dict("horizontal" => "right", "vertical" => "top", "wrapText" => "1"))
        XLSX.setAlignment(s, 18, 20:3:26; horizontal="left", vertical="top", wrapText=true)
        @test XLSX.getAlignment(s, "W18").alignment == Dict("alignment" => Dict("horizontal" => "left", "vertical" => "top", "wrapText" => "1"))
        XLSX.setAlignment(s, 18:2:22, 20:3:26; horizontal="left", vertical="bottom", wrapText=true)
        @test XLSX.getAlignment(s, "Z20").alignment == Dict("alignment" => Dict("horizontal" => "left", "vertical" => "bottom", "wrapText" => "1"))
        SAVE_FILES && save_outfile(f)

        f = XLSX.open_xlsx_template(joinpath(data_directory, "customXml.xlsx"))
        s = f["Mock-up"]

        @test XLSX.getAlignment(s, "D51").alignment == Dict("alignment" => Dict("horizontal" => "left", "vertical" => "top", "wrapText" => "1"))
        @test XLSX.getAlignment(s, "D18").alignment == Dict("alignment" => Dict("horizontal" => "center", "vertical" => "top"))

        XLSX.setAlignment(f, "Mock-up!D18"; horizontal="right", wrapText=true)
        @test XLSX.getAlignment(s, "D18").alignment == Dict("alignment" => Dict("horizontal" => "right", "vertical" => "top", "wrapText" => "1"))

        @test XLSX.setAlignment(s, "Location"; horizontal="right", wrapText=true) == -1

        XLSX.setUniformAlignment(s, "B3:D5"; horizontal="right", vertical="justify", wrapText=true)
        @test XLSX.getAlignment(s, "B3").alignment == Dict("alignment" => Dict("horizontal" => "right", "vertical" => "justify", "wrapText" => "1"))
        @test XLSX.getAlignment(s, "C4").alignment == Dict("alignment" => Dict("horizontal" => "right", "vertical" => "justify", "wrapText" => "1"))
        @test XLSX.getAlignment(s, "D5").alignment == Dict("alignment" => Dict("horizontal" => "right", "vertical" => "justify", "wrapText" => "1"))

        XLSX.writexlsx("output.xlsx", f, overwrite=true)
        SAVE_FILES && save_outfile("output.xlsx")
        @test isfile("output.xlsx")

        XLSX.openxlsx("output.xlsx") do f # Check the updated fonts were written correctly
            s = f["Mock-up"]
            @test XLSX.getAlignment(s, "B3").alignment == Dict("alignment" => Dict("horizontal" => "right", "vertical" => "justify", "wrapText" => "1"))
            @test XLSX.getAlignment(s, "C4").alignment == Dict("alignment" => Dict("horizontal" => "right", "vertical" => "justify", "wrapText" => "1"))
            @test XLSX.getAlignment(s, "D5").alignment == Dict("alignment" => Dict("horizontal" => "right", "vertical" => "justify", "wrapText" => "1"))
        end

        isfile("output.xlsx") && rm("output.xlsx")
        SAVE_FILES && save_outfile(f)

        f = XLSX.newxlsx()
        s = f[1]
        s["A1:Z26"] = ""
        XLSX.setUniformAlignment(s, "Sheet1!E5:E6"; horizontal="right", vertical="justify", wrapText=true)
        @test XLSX.getAlignment(s, "E5").alignment == Dict("alignment" => Dict("horizontal" => "right", "vertical" => "justify", "wrapText" => "1"))
        XLSX.setUniformAlignment(s, "Sheet1!A:A"; horizontal="right", vertical="top", wrapText=true)
        @test XLSX.getAlignment(s, "A23").alignment == Dict("alignment" => Dict("horizontal" => "right", "vertical" => "top", "wrapText" => "1"))
        XLSX.setUniformAlignment(s, "Sheet1!15:24"; horizontal="left", vertical="top", wrapText=true)
        @test XLSX.getAlignment(s, "Q15").alignment == Dict("alignment" => Dict("horizontal" => "left", "vertical" => "top", "wrapText" => "1"))
        @test XLSX.getAlignment(s, "A23").alignment == Dict("alignment" => Dict("horizontal" => "left", "vertical" => "top", "wrapText" => "1"))
        XLSX.setUniformAlignment(s, "A:A"; horizontal="right", vertical="bottom", wrapText=true)
        @test XLSX.getAlignment(s, "A15").alignment == Dict("alignment" => Dict("horizontal" => "right", "vertical" => "bottom", "wrapText" => "1"))
        XLSX.setUniformAlignment(s, "10:12"; horizontal="right", vertical="bottom", wrapText=true)
        @test XLSX.getAlignment(s, "Q11").alignment == Dict("alignment" => Dict("horizontal" => "right", "vertical" => "bottom", "wrapText" => "1"))
        SAVE_FILES && save_outfile(f)

        f = XLSX.newxlsx()
        s = f[1]
        s["A1:Z26"] = ""
        XLSX.setUniformAlignment(s, 2, :; horizontal="right", vertical="justify", wrapText=true)
        @test XLSX.getAlignment(s, "E2").alignment == Dict("alignment" => Dict("horizontal" => "right", "vertical" => "justify", "wrapText" => "1"))
        XLSX.setUniformAlignment(s, :, 4:5; horizontal="right", vertical="top", wrapText=true)
        @test XLSX.getAlignment(s, "D23").alignment == Dict("alignment" => Dict("horizontal" => "right", "vertical" => "top", "wrapText" => "1"))
        XLSX.setUniformAlignment(s, :, :; horizontal="left", vertical="top", wrapText=true)
        @test XLSX.getAlignment(s, "Q15").alignment == Dict("alignment" => Dict("horizontal" => "left", "vertical" => "top", "wrapText" => "1"))
        @test XLSX.getAlignment(s, "A23").alignment == Dict("alignment" => Dict("horizontal" => "left", "vertical" => "top", "wrapText" => "1"))
        XLSX.setUniformAlignment(s, :; horizontal="right", vertical="bottom", wrapText=true)
        @test XLSX.getAlignment(s, "A15").alignment == Dict("alignment" => Dict("horizontal" => "right", "vertical" => "bottom", "wrapText" => "1"))
        XLSX.setUniformAlignment(s, :, [8, 12, 14]; horizontal="right", vertical="bottom", wrapText=false)
        @test XLSX.getAlignment(s, "L12").alignment == Dict("alignment" => Dict("horizontal" => "right", "vertical" => "bottom", "wrapText" => "0"))
        XLSX.setUniformAlignment(s, 8:12:20, 3; horizontal="right", vertical="top", wrapText=false)
        @test XLSX.getAlignment(s, "C20").alignment == Dict("alignment" => Dict("horizontal" => "right", "vertical" => "top", "wrapText" => "0"))
        XLSX.setUniformAlignment(s, 8:12:20, [3, 4]; horizontal="justify", vertical="justify", wrapText=false)
        @test XLSX.getAlignment(s, "D8").alignment == Dict("alignment" => Dict("horizontal" => "justify", "vertical" => "justify", "wrapText" => "0"))
        XLSX.setUniformAlignment(s, 8:20, 8; horizontal="justify", vertical="justify", wrapText=true)
        @test XLSX.getAlignment(s, "H15").alignment == Dict("alignment" => Dict("horizontal" => "justify", "vertical" => "justify", "wrapText" => "1"))
        SAVE_FILES && save_outfile(f)

        f = XLSX.newxlsx()
        s = f[1]
        s["A1:Z26"] = ""
        XLSX.setAlignment(s, "A1"; horizontal="right", vertical="justify", wrapText=true)
        XLSX.setUniformAlignment(f, "Sheet1!A1,Sheet1!C3,Sheet1!E5:E6")
        @test XLSX.getAlignment(s, "A1").alignment == Dict("alignment" => Dict("horizontal" => "right", "vertical" => "justify", "wrapText" => "1"))
        @test XLSX.getAlignment(s, "C3").alignment == Dict("alignment" => Dict("horizontal" => "right", "vertical" => "justify", "wrapText" => "1"))
        @test XLSX.getAlignment(s, "E5").alignment == Dict("alignment" => Dict("horizontal" => "right", "vertical" => "justify", "wrapText" => "1"))
        @test XLSX.getAlignment(s, "E6").alignment == Dict("alignment" => Dict("horizontal" => "right", "vertical" => "justify", "wrapText" => "1"))
        @test_throws XLSX.XLSXError XLSX.setAlignment(s, 2, [1, 3, 10, 15, 28]; horizontal="right", vertical="justify", wrapText=true)
        @test_throws XLSX.XLSXError XLSX.setAlignment(s, [1, 3, 10, 15, 28], 2:3; horizontal="right", vertical="justify", wrapText=true)
        @test_throws XLSX.XLSXError XLSX.setAlignment(s, :, [1, 3, 10, 15, 28]; horizontal="right", vertical="justify", wrapText=true)
        @test_throws XLSX.XLSXError XLSX.setAlignment(s, [1, 3, 10, 15, 28], :; horizontal="right", vertical="justify", wrapText=true)
        @test_throws XLSX.XLSXError XLSX.setAlignment(s, "Z100"; horizontal="right", vertical="justify", wrapText=true)
        @test_throws XLSX.XLSXError XLSX.setAlignment(s, "Sheet1!Z100"; horizontal="right", vertical="justify", wrapText=true)
        @test_throws XLSX.XLSXError XLSX.setAlignment(s, "Sheet1!E1:F5,Sheet1!Z100"; horizontal="right", vertical="justify", wrapText=true)
        @test_throws XLSX.XLSXError XLSX.setAlignment(f, "Sheet1!garbage"; horizontal="right", vertical="justify", wrapText=true)
        @test_throws XLSX.XLSXError XLSX.setAlignment(s, "Sheet1!garbage"; horizontal="right", vertical="justify", wrapText=true)
        @test_throws XLSX.XLSXError XLSX.setAlignment(s, "garbage"; horizontal="right", vertical="justify", wrapText=true)
        @test_throws XLSX.XLSXError XLSX.setAlignment(s, "garbage1:garbage2"; horizontal="right", vertical="justify", wrapText=true)
        SAVE_FILES && save_outfile(f)

        f = XLSX.newxlsx()
        s = f[1]
        s["A1:Z26"] = ""
        XLSX.setAlignment(s, "A1"; horizontal="right", vertical="justify", wrapText=true)
        XLSX.setUniformAlignment(s, 1, 1:2:25)
        @test XLSX.getAlignment(s, 1, 1).alignment == Dict("alignment" => Dict("horizontal" => "right", "vertical" => "justify", "wrapText" => "1"))
        @test XLSX.getAlignment(s, 1, 9).alignment == Dict("alignment" => Dict("horizontal" => "right", "vertical" => "justify", "wrapText" => "1"))
        @test XLSX.getAlignment(s, 1, 19).alignment == Dict("alignment" => Dict("horizontal" => "right", "vertical" => "justify", "wrapText" => "1"))
        @test XLSX.getAlignment(s, 1, 25).alignment == Dict("alignment" => Dict("horizontal" => "right", "vertical" => "justify", "wrapText" => "1"))
        @test XLSX.getAlignment(s, 1, 8) === nothing
        @test XLSX.getAlignment(s, 1, 16) === nothing
        @test XLSX.getAlignment(s, 1, 22) === nothing
        @test XLSX.getAlignment(s, 1, 24) === nothing
        SAVE_FILES && save_outfile(f)

        f = XLSX.newxlsx()
        s = f[1]
        s["A1:Z26"] = ""
        XLSX.setAlignment(s, "A2"; horizontal="right", vertical="justify", wrapText=true)
        XLSX.setUniformAlignment(s, 2:2:26, :)
        @test XLSX.getAlignment(s, "A2").alignment == Dict("alignment" => Dict("horizontal" => "right", "vertical" => "justify", "wrapText" => "1"))
        @test XLSX.getAlignment(s, "C4").alignment == Dict("alignment" => Dict("horizontal" => "right", "vertical" => "justify", "wrapText" => "1"))
        @test XLSX.getAlignment(s, "K6").alignment == Dict("alignment" => Dict("horizontal" => "right", "vertical" => "justify", "wrapText" => "1"))
        @test XLSX.getAlignment(s, "Y24").alignment == Dict("alignment" => Dict("horizontal" => "right", "vertical" => "justify", "wrapText" => "1"))
        @test XLSX.getAlignment(s, "A3") === nothing
        @test XLSX.getAlignment(s, "C5") === nothing
        @test XLSX.getAlignment(s, "K7") === nothing
        @test XLSX.getAlignment(s, "Y25") === nothing
        @test_throws XLSX.XLSXError XLSX.setUniformAlignment(s, 2, [1, 3, 10, 15, 28]; horizontal="right", vertical="justify", wrapText=true)
        @test_throws XLSX.XLSXError XLSX.setUniformAlignment(s, [1, 3, 10, 15, 28], 2:3; horizontal="right", vertical="justify", wrapText=true)
        @test_throws XLSX.XLSXError XLSX.setUniformAlignment(s, :, [1, 3, 10, 15, 28]; horizontal="right", vertical="justify", wrapText=true)
        @test_throws XLSX.XLSXError XLSX.setUniformAlignment(s, [1, 3, 10, 15, 28], :; horizontal="right", vertical="justify", wrapText=true)
        @test_throws XLSX.XLSXError XLSX.setUniformAlignment(s, "Z100:Z101"; horizontal="right", vertical="justify", wrapText=true)
        @test_throws XLSX.XLSXError XLSX.setUniformAlignment(s, "Sheet1!Z100:Z101"; horizontal="right", vertical="justify", wrapText=true)
        @test_throws XLSX.XLSXError XLSX.setUniformAlignment(s, "Sheet1!E1:F5,Sheet1!Z100"; horizontal="right", vertical="justify", wrapText=true)
        @test_throws XLSX.XLSXError XLSX.setUniformAlignment(s, "Z100:Z100"; horizontal="right", vertical="justify", wrapText=true)
        @test_throws XLSX.XLSXError XLSX.setUniformAlignment(s, "Sheet1!Z100:Z100"; horizontal="right", vertical="justify", wrapText=true)
        @test_throws XLSX.XLSXError XLSX.setUniformAlignment(s, "Sheet1!Z100,Sheet1!Z100"; horizontal="right", vertical="justify", wrapText=true)
        @test_throws XLSX.XLSXError XLSX.setUniformAlignment(f, "Sheet1!garbage"; horizontal="right", vertical="justify", wrapText=true)
        @test_throws XLSX.XLSXError XLSX.setUniformAlignment(s, "Sheet1!garbage"; horizontal="right", vertical="justify", wrapText=true)
        @test_throws XLSX.XLSXError XLSX.setUniformAlignment(s, "garbage"; horizontal="right", vertical="justify", wrapText=true)
        @test_throws XLSX.XLSXError XLSX.setUniformAlignment(s, "garbage1:garbage2"; horizontal="right", vertical="justify", wrapText=true)
        SAVE_FILES && save_outfile(f)

    end

    @testset "setFormat" begin

        f = XLSX.open_empty_template()
        s = f["Sheet1"]

        s["A1"] = "Hello"
        s["B1"] = "World"
        s["A2"] = 2.367
        s["B2"] = 200450023
        s["C1"] = Dates.Date(2018, 2, 1)
        s["C2"] = 0.45
        s["D1"] = 100.24
        s["D2"] = Dates.Time(0, 19, 30)
        s["E1:F5"] = [Dates.Date(2025, 6, 23) Dates.Date(2025, 6, 23);
            Dates.Date(2025, 6, 23) Dates.Date(2025, 6, 23);
            Dates.Date(2025, 6, 23) Dates.Date(2025, 6, 23);
            Dates.Date(2025, 6, 23) Dates.Date(2025, 6, 23);
            Dates.Date(2025, 6, 23) Dates.Date(2025, 6, 23)
        ]

        @test XLSX.setFormat(s, "B2"; format="Scientific") == 48
        @test XLSX.setFormat(s, "C2"; format="Percentage") == 9
        @test XLSX.setFormat(s, "C1"; format="General") == 0
        @test XLSX.setFormat(s, "D2"; format="Currency") == 7
        @test XLSX.setFormat(s, "C1"; format="LongDate") == 15
        @test XLSX.setFormat(s, "D2"; format="Time") == 21

        @test XLSX.getFormat(s, "A2") === nothing
        @test XLSX.getFormat(s, "D2").format == Dict("numFmt" => Dict("numFmtId" => "21", "formatCode" => "h:mm:ss"))
        @test XLSX.getFormat(s, "E2").format == Dict("numFmt" => Dict("numFmtId" => "14", "formatCode" => "m/d/yyyy"))
        @test XLSX.getFormat(s, "F2").format == Dict("numFmt" => Dict("numFmtId" => "14", "formatCode" => "m/d/yyyy"))


        @test XLSX.setFormat(s, "A2"; format="""_-£* #,##0.00_-;-£* #,##0.00_-;_-£* "-"??_-;_-@_-""") == 164
        @test XLSX.getFormat(s, "A2").format == Dict("numFmt" => Dict("numFmtId" => "164", "formatCode" => "_-£* #,##0.00_-;-£* #,##0.00_-;_-£* \"-\"??_-;_-@_-"))
        @test XLSX.setFormat(s, "A2"; format="_-£* #,##0.00_-;-£* #,##0.00_-;_-£* \"-\"??_-;_-@_-") == 164
        @test XLSX.getFormat(s, "A2").format == Dict("numFmt" => Dict("numFmtId" => "164", "formatCode" => "_-£* #,##0.00_-;-£* #,##0.00_-;_-£* \"-\"??_-;_-@_-"))

        @test XLSX.setFormat(s, "D2"; format="h:mm AM/PM") == 18
        @test XLSX.setFormat(s, "A2"; format="# ??/??") == 13
        @test XLSX.getFormat(s, "A2").format == Dict("numFmt" => Dict("numFmtId" => "13", "formatCode" => "# ??/??"))
        @test XLSX.getFormat(s, "D2").format == Dict("numFmt" => Dict("numFmtId" => "18", "formatCode" => "h:mm AM/PM"))

        @test XLSX.setFormat(s, "E1:E5"; format="General") == -1
        @test XLSX.setFormat(s, "F1:F5"; format="Currency") == -1
        @test XLSX.getFormat(s, "E2").format == Dict("numFmt" => Dict("numFmtId" => "0", "formatCode" => "General"))
        @test XLSX.getFormat(f, "Sheet1!F2").format == Dict("numFmt" => Dict("numFmtId" => "7", "formatCode" => "\$#,##0.00_);(\$#,##0.00)"))


        @test XLSX.setFormat(f, "Sheet1!E1:F5"; format="#,##0.000") == -1
        @test XLSX.setFormat(s, "F1:F5"; format="#,##0.000") == -1
        @test XLSX.getFormat(s, "E2").format == Dict("numFmt" => Dict("numFmtId" => "165", "formatCode" => "#,##0.000"))
        @test XLSX.getFormat(f, "Sheet1!F2").format == Dict("numFmt" => Dict("numFmtId" => "165", "formatCode" => "#,##0.000"))
        @test_throws XLSX.XLSXError XLSX.setFormat(s, "Z100"; format="Currency")
        @test_throws XLSX.XLSXError XLSX.setFormat(s, "Sheet1!Z100"; format="Currency")
        @test_throws XLSX.XLSXError XLSX.setFormat(s, "Sheet1!E1:F5,Sheet1!Z100"; format="Currency")
        @test_throws XLSX.XLSXError XLSX.setFormat(s, "E2"; format="ffzz345")


        XLSX.writexlsx("test.xlsx", f, overwrite=true)
        SAVE_FILES && save_outfile("test.xlsx")

        XLSX.openxlsx("test.xlsx") do f # Check the updated formats were written correctly
            s = f["Sheet1"]
            @test XLSX.getFormat(s, "A2").format == Dict("numFmt" => Dict("numFmtId" => "13", "formatCode" => "# ??/??"))
            @test XLSX.getFormat(s, "D2").format == Dict("numFmt" => Dict("numFmtId" => "18", "formatCode" => "h:mm AM/PM"))
            @test XLSX.getFormat(s, "E2").format == Dict("numFmt" => Dict("numFmtId" => "165", "formatCode" => "#,##0.000"))
            @test XLSX.getFormat(f, "Sheet1!F2").format == Dict("numFmt" => Dict("numFmtId" => "165", "formatCode" => "#,##0.000"))
        end

        isfile("test.xlsx") && rm("test.xlsx")

        f = XLSX.newxlsx()
        s = f[1]
        s["A1:Z26"] = ""
        XLSX.setFormat(s, "Sheet1!E5"; format="Currency")
        @test XLSX.getFormat(f, "Sheet1!E5").format == Dict("numFmt" => Dict("numFmtId" => "7", "formatCode" => "\$#,##0.00_);(\$#,##0.00)"))
        XLSX.setFormat(s, "Sheet1!W5:X8"; format="Currency")
        @test XLSX.getFormat(f, "Sheet1!X7").format == Dict("numFmt" => Dict("numFmtId" => "7", "formatCode" => "\$#,##0.00_);(\$#,##0.00)"))
        XLSX.setFormat(s, "Sheet1!F:G"; format="Currency")
        @test XLSX.getFormat(s, "F3").format == Dict("numFmt" => Dict("numFmtId" => "7", "formatCode" => "\$#,##0.00_);(\$#,##0.00)"))
        XLSX.setFormat(s, "Sheet1!4:8"; format="Currency")
        @test XLSX.getFormat(s, "Q7").format == Dict("numFmt" => Dict("numFmtId" => "7", "formatCode" => "\$#,##0.00_);(\$#,##0.00)"))
        XLSX.setFormat(s, "N4,M8:M15,Z25:Z26"; format="#,##0.000")
        @test XLSX.getFormat(s, "Z26").format == Dict("numFmt" => Dict("numFmtId" => "164", "formatCode" => "#,##0.000"))
        XLSX.setFormat(s, "Sheet1!4:8"; format="39")
        @test XLSX.getFormat(s, "Q7").format == Dict("numFmt" => Dict("numFmtId" => "39", "formatCode" => "#,##0.00_);(#,##0.00)"))
        XLSX.setFormat(s, "Sheet1!4:8"; format="#,##0")
        @test XLSX.getFormat(s, "Q7").format == Dict("numFmt" => Dict("numFmtId" => "3", "formatCode" => "#,##0"))
        SAVE_FILES && save_outfile(f)

        f = XLSX.newxlsx()
        s = f[1]
        s["A1:Z26"] = ""
        XLSX.setFormat(s, :, 2:4; format="Currency")
        @test XLSX.getFormat(f, "Sheet1!B23").format == Dict("numFmt" => Dict("numFmtId" => "7", "formatCode" => "\$#,##0.00_);(\$#,##0.00)"))
        XLSX.setFormat(s, 4:3:10, :; format="Currency")
        @test XLSX.getFormat(s, "Q7").format == Dict("numFmt" => Dict("numFmtId" => "7", "formatCode" => "\$#,##0.00_);(\$#,##0.00)"))
        XLSX.setFormat(s, :, [8, 23, 4]; format="Currency")
        @test XLSX.getFormat(s, "H1").format == Dict("numFmt" => Dict("numFmtId" => "7", "formatCode" => "\$#,##0.00_);(\$#,##0.00)"))
        XLSX.setFormat(s, 25:26, 20:26; format="#,##0.000")
        @test XLSX.getFormat(s, "Z26").format == Dict("numFmt" => Dict("numFmtId" => "164", "formatCode" => "#,##0.000"))
        XLSX.setFormat(s, 25:26, 15; format="#,##0.0000")
        @test XLSX.getFormat(s, "O26").format == Dict("numFmt" => Dict("numFmtId" => "165", "formatCode" => "#,##0.0000"))
        XLSX.setFormat(s, 21:2:25, [15, 16]; format="#,##0.0")
        @test XLSX.getFormat(s, "P25").format == Dict("numFmt" => Dict("numFmtId" => "166", "formatCode" => "#,##0.0"))
        SAVE_FILES && save_outfile(f)

        f = XLSX.newxlsx()
        s = f[1]
        s["A1:Z26"] = ""
        XLSX.setUniformFormat(s, "Sheet1!W5:X8"; format="Currency")
        @test XLSX.getFormat(f, "Sheet1!X7").format == Dict("numFmt" => Dict("numFmtId" => "7", "formatCode" => "\$#,##0.00_);(\$#,##0.00)"))
        XLSX.setUniformFormat(s, "Sheet1!F:G"; format="Currency")
        @test XLSX.getFormat(s, "F3").format == Dict("numFmt" => Dict("numFmtId" => "7", "formatCode" => "\$#,##0.00_);(\$#,##0.00)"))
        XLSX.setUniformFormat(s, "Sheet1!4:8"; format="Currency")
        @test XLSX.getFormat(s, "Q7").format == Dict("numFmt" => Dict("numFmtId" => "7", "formatCode" => "\$#,##0.00_);(\$#,##0.00)"))
        XLSX.setUniformFormat(s, "N4,M8:M15,Z25:Z26"; format="#,##0.000")
        @test XLSX.getFormat(s, "Z26").format == Dict("numFmt" => Dict("numFmtId" => "164", "formatCode" => "#,##0.000"))
        @test_throws XLSX.XLSXError XLSX.setUniformFormat(s, "Z100:Z101"; format="Currency")
        @test_throws XLSX.XLSXError XLSX.setUniformFormat(s, "Sheet1!Z100:Z101"; format="Currency")
        @test_throws XLSX.XLSXError XLSX.setUniformFormat(s, "Sheet1!E1:F5,Sheet1!Z100"; format="Currency")
        SAVE_FILES && save_outfile(f)

        f = XLSX.newxlsx()
        s = f[1]
        s["A1:Z26"] = ""
        XLSX.setUniformFormat(s, :, 2:4; format="Currency")
        @test XLSX.getFormat(f, "Sheet1!B23").format == Dict("numFmt" => Dict("numFmtId" => "7", "formatCode" => "\$#,##0.00_);(\$#,##0.00)"))
        XLSX.setUniformFormat(s, 4:3:10, :; format="Currency")
        @test XLSX.getFormat(s, "Q7").format == Dict("numFmt" => Dict("numFmtId" => "7", "formatCode" => "\$#,##0.00_);(\$#,##0.00)"))
        XLSX.setUniformFormat(s, :, [8, 23, 4]; format="Currency")
        @test XLSX.getFormat(s, "H1").format == Dict("numFmt" => Dict("numFmtId" => "7", "formatCode" => "\$#,##0.00_);(\$#,##0.00)"))
        XLSX.setUniformFormat(s, 25:26, 20:26; format="#,##0.000")
        @test XLSX.getFormat(s, "Z26").format == Dict("numFmt" => Dict("numFmtId" => "164", "formatCode" => "#,##0.000"))
        XLSX.setUniformFormat(s, 25:26, 15; format="#,##0.0000")
        @test XLSX.getFormat(s, "O26").format == Dict("numFmt" => Dict("numFmtId" => "165", "formatCode" => "#,##0.0000"))
        XLSX.setUniformFormat(s, 21:2:25, [15, 16]; format="#,##0.0")
        @test XLSX.getFormat(s, "P25").format == Dict("numFmt" => Dict("numFmtId" => "166", "formatCode" => "#,##0.0"))
        SAVE_FILES && save_outfile(f)

        f = XLSX.newxlsx()
        s = f[1]
        s["A1:Z26"] = ""
        XLSX.setUniformFormat(s, :, :; format="Currency")
        @test XLSX.getFormat(f, "Sheet1!B23").format == Dict("numFmt" => Dict("numFmtId" => "7", "formatCode" => "\$#,##0.00_);(\$#,##0.00)"))
        XLSX.setUniformFormat(s, 4:10, :; format="#,##0.000")
        @test XLSX.getFormat(s, "Q7").format == Dict("numFmt" => Dict("numFmtId" => "164", "formatCode" => "#,##0.000"))
        XLSX.setUniformFormat(s, [8, 23, 4], 8; format="#,##0.0")
        @test XLSX.getFormat(s, "H8").format == Dict("numFmt" => Dict("numFmtId" => "165", "formatCode" => "#,##0.0"))
        SAVE_FILES && save_outfile(f)

    end

    @testset "UniformStyle" begin
        f = XLSX.newxlsx()
        s = f[1]
        s["A1:Z26"] = ""

        XLSX.setFont(s, "A1:F5"; size=18, name="Arial")
        cell_style = XLSX.getcell(s, "A1").style
        @test XLSX.setUniformStyle(s, "A1:F5") == cell_style
        @test XLSX.getcell(s, "F5").style == cell_style

        XLSX.setFont(s, "A6:F10"; size=10, name="Aptos")
        cell_style = XLSX.getcell(s, "E6").style
        @test XLSX.setUniformStyle(s, [6, 7, 8, 9, 10], 5) == cell_style
        @test XLSX.getcell(s, "E8").style == cell_style

        XLSX.setFont(s, "A11:F15"; size=10, name="Times New Roman")
        cell_style = XLSX.getcell(s, "E6").style
        @test XLSX.setUniformStyle(s, [6, 7, 8, 9, 10], :) == cell_style
        @test XLSX.getcell(s, "Z8").style == cell_style

        XLSX.setFont(s, "A16"; size=80, name="Ariel")
        cell_style = XLSX.getcell(s, "A16").style
        @test XLSX.setUniformStyle(s, "A16,A15,D20:E25,F25") == cell_style
        @test XLSX.getcell(s, "A15").style == cell_style
        @test XLSX.getcell(s, "D20").style == cell_style
        @test XLSX.getcell(s, "F25").style == cell_style

        XLSX.setFont(s, "A1"; size=8, name="Aptos")
        cell_style = XLSX.getcell(s, "A1").style
        @test XLSX.setUniformStyle(s, :) == cell_style
        @test XLSX.getcell(s, "A1").style == cell_style
        @test XLSX.getcell(s, "M13").style == cell_style
        @test XLSX.getcell(s, "Z26").style == cell_style
        SAVE_FILES && save_outfile(f)

        f = XLSX.newxlsx()
        s = f[1]
        s["A1:Z26"] = ""
        XLSX.setFont(s, "A1"; size=8, name="Aptos")
        cell_style = XLSX.getcell(s, "A1").style
        @test XLSX.setUniformStyle(s, "Sheet1!A1:A26") == cell_style
        @test XLSX.getcell(s, "A2").style == cell_style
        @test XLSX.getcell(s, "A13").style == cell_style
        @test XLSX.getcell(s, "A26").style == cell_style
        @test XLSX.setUniformStyle(s, "Sheet1!1:2") == cell_style
        @test XLSX.getcell(s, "B1").style == cell_style
        @test XLSX.getcell(s, "M2").style == cell_style
        @test XLSX.getcell(s, "Z1").style == cell_style
        @test XLSX.setUniformStyle(s, "Sheet1!B:C") == cell_style
        @test XLSX.getcell(s, "C3").style == cell_style
        @test XLSX.getcell(s, "B13").style == cell_style
        @test XLSX.getcell(s, "C26").style == cell_style
        XLSX.setFont(s, "A1"; size=8, name="Arial")
        cell_style = XLSX.getcell(s, "A1").style
        @test XLSX.setUniformStyle(s, "A1:A26") == cell_style
        @test XLSX.getcell(s, "A2").style == cell_style
        @test XLSX.getcell(s, "A13").style == cell_style
        @test XLSX.getcell(s, "A26").style == cell_style
        @test XLSX.setUniformStyle(s, "1:2") == cell_style
        @test XLSX.getcell(s, "B1").style == cell_style
        @test XLSX.getcell(s, "M2").style == cell_style
        @test XLSX.getcell(s, "Z1").style == cell_style
        @test XLSX.setUniformStyle(s, "B:C") == cell_style
        @test XLSX.getcell(s, "C3").style == cell_style
        @test XLSX.getcell(s, "B13").style == cell_style
        @test XLSX.getcell(s, "C26").style == cell_style
        SAVE_FILES && save_outfile(f)

        f = XLSX.newxlsx()
        s = f[1]
        s["A1:Z26"] = ""
        XLSX.setFont(s, "A1"; size=8, name="Aptos")
        cell_style = XLSX.getcell(s, "A1").style
        @test XLSX.setUniformStyle(s, 1, :) == cell_style
        @test XLSX.getcell(s, "B1").style == cell_style
        @test XLSX.setUniformStyle(s, :, 2) == cell_style
        @test XLSX.getcell(s, "B13").style == cell_style
        @test XLSX.setUniformStyle(s, :, 5:2:15) == cell_style
        @test XLSX.getcell(s, "E25").style == cell_style
        @test XLSX.setUniformStyle(s, 5:10, [15, 16, 17]) == cell_style
        @test XLSX.getcell(s, "P10").style == cell_style
        @test XLSX.setUniformStyle(s, 5:10, 17:19) == cell_style
        @test XLSX.getcell(s, "S10").style == cell_style
        @test XLSX.setUniformStyle(s, [10, 12, 26], [19, 24, 26]) == cell_style
        @test XLSX.getcell(s, "Z26").style == cell_style
        @test XLSX.setUniformStyle(s, :, :) == cell_style
        @test XLSX.getcell(s, "Y4").style == cell_style
        @test_throws XLSX.XLSXError XLSX.setUniformStyle(s, :, [1, 3, 10, 15, 28])
        @test_throws XLSX.XLSXError XLSX.setUniformStyle(s, [1, 3, 10, 15, 28], :)
        @test_throws XLSX.XLSXError XLSX.setUniformStyle(s, 1, [1, 3, 10, 15, 28])
        @test_throws XLSX.XLSXError XLSX.setUniformStyle(s, [1, 3, 10, 15, 28], 2:3)
        @test_throws XLSX.XLSXError XLSX.setUniformStyle(f, "Sheet1!garbage")
        @test_throws XLSX.XLSXError XLSX.setUniformStyle(s, "Sheet1!garbage")
        @test_throws XLSX.XLSXError XLSX.setUniformStyle(s, "garbage")
        @test_throws XLSX.XLSXError XLSX.setUniformStyle(s, "garbage1:garbage2")
        SAVE_FILES && save_outfile(f)

    end

    @testset "Width and height" begin

        f = XLSX.open_empty_template()
        s = f["Sheet1"]

        s["A1"] = "Hello"
        s["B1"] = "World"
        s["A2"] = 2.367
        s["B2"] = 200450023
        s["C1"] = Dates.Date(2018, 2, 1)
        s["C2"] = 0.45
        s["D1"] = 100.24
        s["D2"] = Dates.Time(0, 19, 30)
        s["E1:F5"] = [Dates.Date(2025, 6, 23) Dates.Date(2025, 6, 23);
            Dates.Date(2025, 6, 23) Dates.Date(2025, 6, 23);
            Dates.Date(2025, 6, 23) Dates.Date(2025, 6, 23);
            Dates.Date(2025, 6, 23) Dates.Date(2025, 6, 23);
            Dates.Date(2025, 6, 23) Dates.Date(2025, 6, 23)
        ]

        XLSX.setColumnWidth(s, "A2"; width=30)
        @test XLSX.getColumnWidth(s, "A2") ≈ 30.7109375

        XLSX.setColumnWidth(s, "B2:C2"; width=10.1)
        @test XLSX.getColumnWidth(s, "B3") ≈ 10.8109375
        @test XLSX.getColumnWidth(s, "C4") ≈ 10.8109375

        XLSX.setRowHeight(s, "A2"; height=30)
        @test XLSX.getRowHeight(s, "A2") ≈ 30.2109375

        XLSX.setRowHeight(s, "B2:C5"; height=10.1)
        @test XLSX.getRowHeight(s, "B3") ≈ 10.3109375
        @test XLSX.getRowHeight(s, "C4") ≈ 10.3109375

        # Make sure setting row height doesn't affect column width
        # and vice versa.
        @test XLSX.getColumnWidth(s, "B3") ≈ 10.8109375
        @test XLSX.getColumnWidth(s, "C4") ≈ 10.8109375
        XLSX.setColumnWidth(s, "B2:C2"; width=30.5)
        @test XLSX.getColumnWidth(s, "B3") ≈ 31.2109375
        @test XLSX.getColumnWidth(f, "Sheet1!C4") ≈ 31.2109375
        @test XLSX.getRowHeight(s, "B3") ≈ 10.3109375
        @test XLSX.getRowHeight(f, "Sheet1!C4") ≈ 10.3109375
        SAVE_FILES && save_outfile(f)

        f = XLSX.open_xlsx_template(joinpath(data_directory, "customXml.xlsx"))
        s = f["Mock-up"]

        XLSX.setColumnWidth(s, "Location"; width=60)
        XLSX.setRowHeight(s, "Location"; height=50)
        @test XLSX.getRowHeight(s, "D18") ≈ 50.2109375
        @test XLSX.getColumnWidth(s, "D18") ≈ 60.7109375
        @test XLSX.getRowHeight(f, "Mock-up!J20") ≈ 50.2109375
        @test XLSX.getColumnWidth(f, "Mock-up!J20") ≈ 60.7109375
        SAVE_FILES && save_outfile(f)

        f = XLSX.newxlsx()
        s = f[1]
        s["A1:Z26"] = ""
        XLSX.setColumnWidth(s, "Sheet1!A1"; width=60)
        @test XLSX.getColumnWidth(s, "A1") ≈ 60.7109375
        XLSX.setColumnWidth(s, "Sheet1!A1:Z1"; width=60)
        @test XLSX.getColumnWidth(s, "R1") ≈ 60.7109375
        XLSX.setColumnWidth(s, "Sheet1!A:B"; width=60)
        @test XLSX.getColumnWidth(s, "B26") ≈ 60.7109375
        XLSX.setColumnWidth(s, "Sheet1!2:3"; width=60)
        @test XLSX.getColumnWidth(s, "R26") ≈ 60.7109375
        XLSX.setColumnWidth(s, "A:B"; width=30.5)
        @test XLSX.getColumnWidth(s, "B26") ≈ 31.2109375
        XLSX.setColumnWidth(s, "2:3"; width=30.5)
        @test XLSX.getColumnWidth(s, "R3") ≈ 31.2109375
        XLSX.setColumnWidth(s, "Sheet1!C5:C7,Sheet1!F5:F7,Sheet1!H7"; width=10.1)
        @test XLSX.getColumnWidth(s, "F26") ≈ 10.8109375
        XLSX.setColumnWidth(s, 5, :; width=10.0)
        @test XLSX.getColumnWidth(s, "Q5") ≈ 10.7109375
        XLSX.setColumnWidth(s, 5:7; width=10.2)
        @test XLSX.getColumnWidth(s, "G22") ≈ 10.9109375
        XLSX.setColumnWidth(s, :, 5:7; width=10.3)
        @test XLSX.getColumnWidth(s, "G22") ≈ 11.0109375
        XLSX.setColumnWidth(s, :, :; width=10.4)
        @test XLSX.getColumnWidth(s, "G22") ≈ 11.1109375
        XLSX.setColumnWidth(s, :; width=10.5)
        @test XLSX.getColumnWidth(s, "G22") ≈ 11.2109375
        XLSX.setColumnWidth(s, 2:3:11, :; width=10.6)
        @test XLSX.getColumnWidth(s, "Z26") ≈ 11.3109375
        XLSX.setColumnWidth(s, 2:3:11; width=10.7)
        @test XLSX.getColumnWidth(s, "E26") ≈ 11.4109375
        XLSX.setColumnWidth(s, :, [2, 3, 11]; width=10.8)
        @test XLSX.getColumnWidth(s, "K15") ≈ 11.5109375
        XLSX.setColumnWidth(s, 3:6, [2, 3, 11]; width=10.9)
        @test XLSX.getColumnWidth(s, "K15") ≈ 11.6109375
        XLSX.setColumnWidth(s, 3:3:6, [2, 3, 11]; width=11.0)
        @test XLSX.getColumnWidth(s, "K15") ≈ 11.7109375
        XLSX.setColumnWidth(s, 11, 7:13; width=11.1)
        @test XLSX.getColumnWidth(s, "K15") ≈ 11.8109375
        SAVE_FILES && save_outfile(f)

        f = XLSX.newxlsx()
        s = f[1]
        s["A1:Z26"] = ""
        XLSX.setRowHeight(s, "Sheet1!A1"; height=10.1)
        @test XLSX.getRowHeight(s, "A1") ≈ 10.3109375
        XLSX.setRowHeight(s, "Sheet1!A1:A26"; height=10.2)
        @test XLSX.getRowHeight(s, "R20") ≈ 10.4109375
        XLSX.setRowHeight(s, "Sheet1!A:B"; height=10.3)
        @test XLSX.getRowHeight(s, "B26") ≈ 10.5109375
        XLSX.setRowHeight(s, "Sheet1!2:3"; height=10.4)
        @test XLSX.getRowHeight(s, "R3") ≈ 10.6109375
        XLSX.setRowHeight(s, "A:B"; height=10.5)
        @test XLSX.getRowHeight(s, "B26") ≈ 10.7109375
        XLSX.setRowHeight(s, "2:3"; height=10.6)
        @test XLSX.getRowHeight(s, "R3") ≈ 10.8109375
        XLSX.setRowHeight(s, "Sheet1!C5:C7,Sheet1!F5:F7,Sheet1!H7"; height=10.7)
        @test XLSX.getRowHeight(s, "F6") ≈ 10.9109375
        XLSX.setRowHeight(s, 5, :; height=10.8)
        @test XLSX.getRowHeight(s, "Q5") ≈ 11.0109375
        XLSX.setRowHeight(s, 5:7; height=10.9)
        @test XLSX.getRowHeight(s, "P6") ≈ 11.1109375
        XLSX.setRowHeight(s, :, 5:7; height=11.0)
        @test XLSX.getRowHeight(s, "G22") ≈ 11.2109375
        XLSX.setRowHeight(s, :, :; height=11.1)
        @test XLSX.getRowHeight(s, "G22") ≈ 11.3109375
        XLSX.setRowHeight(s, :; height=11.2)
        @test XLSX.getRowHeight(s, "G22") ≈ 11.4109375
        XLSX.setRowHeight(s, 2:3:11, :; height=11.3)
        @test XLSX.getRowHeight(s, "J8") ≈ 11.5109375
        XLSX.setRowHeight(s, 2:3:11; height=11.4)
        @test XLSX.getRowHeight(s, "J8") ≈ 11.6109375
        XLSX.setRowHeight(s, :, [2, 3, 11]; height=11.5)
        @test XLSX.getRowHeight(s, "K15") ≈ 11.7109375
        XLSX.setRowHeight(s, 3:6, [2, 3, 11]; height=11.6)
        @test XLSX.getRowHeight(s, "K5") ≈ 11.8109375
        XLSX.setRowHeight(s, 3:3:6, [2, 3, 11]; height=11.7)
        @test XLSX.getRowHeight(s, "K6") ≈ 11.9109375
        XLSX.setRowHeight(s, 11, 7:13; height=11.8)
        @test XLSX.getRowHeight(s, "K11") ≈ 12.0109375
        SAVE_FILES && save_outfile(f)

    end

    @testset "No cache" begin
        XLSX.openxlsx(joinpath(data_directory, "customXml.xlsx"); mode="r", enable_cache=true) do f
            @test_throws XLSX.XLSXError XLSX.getRowHeight(f, "Mock-up!B2") ≈ 23.25 # File not writable
            @test_throws XLSX.XLSXError XLSX.getColumnWidth(f, "Mock-up!B2") # File not writable
        end
        XLSX.openxlsx(joinpath(data_directory, "customXml.xlsx"); mode="r", enable_cache=false) do f
            @test_throws XLSX.XLSXError XLSX.getRowHeight(f, "Mock-up!B2")
            @test_throws XLSX.XLSXError XLSX.getFont(f, "Mock-up!B2")
            @test_throws XLSX.XLSXError XLSX.getFill(f, "Mock-up!B2")
            @test_throws XLSX.XLSXError XLSX.getBorder(f, "Mock-up!B2")
            @test_throws XLSX.XLSXError XLSX.getFormat(f, "Mock-up!B2")
            @test_throws XLSX.XLSXError XLSX.getAlignment(f, "Mock-up!B2")
            @test_throws XLSX.XLSXError XLSX.setRowHeight(f, "Mock-up!B2")
            @test_throws XLSX.XLSXError XLSX.setColumnWidth(f, "Mock-up!B2")
            @test_throws XLSX.XLSXError XLSX.setFont(f, "Mock-up!B2")
            @test_throws XLSX.XLSXError XLSX.setFill(f, "Mock-up!B2")
            @test_throws XLSX.XLSXError XLSX.setBorder(f, "Mock-up!B2")
            @test_throws XLSX.XLSXError XLSX.setFormat(f, "Mock-up!B2")
            @test_throws XLSX.XLSXError XLSX.setAlignment(f, "Mock-up!B2")
            @test_throws XLSX.XLSXError XLSX.setUniformFont(f, "Mock-up!B2:C4")
            @test_throws XLSX.XLSXError XLSX.setUniformFill(f, "Mock-up!B2:C4")
            @test_throws XLSX.XLSXError XLSX.setUniformBorder(f, "Mock-up!B2:C4")
            @test_throws XLSX.XLSXError XLSX.setUniformFormat(f, "Mock-up!B2:C4")
            @test_throws XLSX.XLSXError XLSX.setUniformAlignment(f, "Mock-up!B2:C4")
            @test_throws XLSX.XLSXError XLSX.setOutsideBorder(f, "Mock-up!B2:C4")
        end
    end

    @testset "indexing setAttribute" begin
        f = XLSX.newxlsx() # Empty XLSXFile
        s = f[1] #1×1 XLSX.Worksheet: ["Sheet1"](A1:A1)

        #Can't write to single, empty cells
        @test_throws XLSX.XLSXError XLSX.setFont(s, "A1"; color="grey42")
        @test_throws XLSX.XLSXError XLSX.setFont(s, "A1:A1"; color="grey42")
        @test_throws XLSX.XLSXError XLSX.setFont(s, "A:A"; color="grey42")
        @test_throws XLSX.XLSXError XLSX.setFont(s, "1"; color="grey42")
        @test_throws XLSX.XLSXError XLSX.setFont(s, 1, 1; color="grey42")
        @test_throws XLSX.XLSXError XLSX.setFont(s, [1], 1; color="grey42")
        @test_throws XLSX.XLSXError XLSX.setFont(s, 1, 1:1; color="grey42")
        @test_throws XLSX.XLSXError XLSX.setFont(s, 1, :; color="grey42")
        @test_throws XLSX.XLSXError XLSX.setFont(s, :; color="grey42")
        @test_throws XLSX.XLSXError XLSX.setFont(s, :, :; color="grey42")

        s[2, 1] = ""
        s[3, 3] = ""
        # Skip empty cells silently in ranges 
        @test XLSX.setFont(s, 2:3, 1:3; color="grey42") == -1

        # Outside sheet dimension
        @test_throws XLSX.XLSXError XLSX.getFont(s, 2, 4)
        @test_throws XLSX.XLSXError XLSX.getFont(s, 4, 2)
        @test_throws XLSX.XLSXError XLSX.getFont(s, 4, 4)

        s[1:3, 1:3] = ""
        default_font = XLSX.getDefaultFont(s).font
        dname = default_font["name"]["val"]
        dsize = default_font["sz"]["val"]
        XLSX.setFont(s, "A1"; color="grey42")
        @test XLSX.getFont(s, "A1").font == Dict("name" => Dict("val" => dname), "sz" => Dict("val" => dsize), "color" => Dict("rgb" => "FF6B6B6B"))
        XLSX.setFont(s, 2, 2; color="grey43", name="Ariel")
        @test XLSX.getFont(s, 2, 2).font == Dict("name" => Dict("val" => "Ariel"), "sz" => Dict("val" => dsize), "color" => Dict("rgb" => "FF6E6E6E"))
        XLSX.setFont(s, [2, 3], 1:3; color="grey44", name="Courier New")
        @test XLSX.getFont(s, 3, 1).font == Dict("name" => Dict("val" => "Courier New"), "sz" => Dict("val" => dsize), "color" => Dict("rgb" => "FF707070"))
        @test XLSX.getFont(s, 2, 2).font == Dict("name" => Dict("val" => "Courier New"), "sz" => Dict("val" => dsize), "color" => Dict("rgb" => "FF707070"))
        @test XLSX.getFont(s, 3, 3).font == Dict("name" => Dict("val" => "Courier New"), "sz" => Dict("val" => dsize), "color" => Dict("rgb" => "FF707070"))
        SAVE_FILES && save_outfile(f)

        f = XLSX.newxlsx()
        s = f[1]
        @test_throws XLSX.XLSXError XLSX.setBorder(s, "A1"; allsides=["color" => "grey42", "style" => "thick"])
        @test_throws XLSX.XLSXError XLSX.setBorder(s, "A1:A1"; allsides=["color" => "grey42", "style" => "thick"])
        @test_throws XLSX.XLSXError XLSX.setBorder(s, "A"; allsides=["color" => "grey42", "style" => "thick"])
        @test_throws XLSX.XLSXError XLSX.setBorder(s, "1"; allsides=["color" => "grey42", "style" => "thick"])
        @test_throws XLSX.XLSXError XLSX.setBorder(s, 1, 1; allsides=["color" => "grey42", "style" => "thick"])
        @test_throws XLSX.XLSXError XLSX.setBorder(s, [1], 1; allsides=["color" => "grey42", "style" => "thick"])
        @test_throws XLSX.XLSXError XLSX.setBorder(s, 1, 1:1; allsides=["color" => "grey42", "style" => "thick"])
        @test_throws XLSX.XLSXError XLSX.setBorder(s, :, 1; allsides=["color" => "grey42", "style" => "thick"])
        @test_throws XLSX.XLSXError XLSX.setBorder(s, :; allsides=["color" => "grey42", "style" => "thick"])
        @test_throws XLSX.XLSXError XLSX.setBorder(s, :, :; allsides=["color" => "grey42", "style" => "thick"])
        @test_throws XLSX.XLSXError XLSX.setBorder(s, [2, 3], 1:3; allsides=["color" => "grey42", "style" => "thick"]) == -1
        @test_throws XLSX.XLSXError XLSX.getBorder(s, 2, 1)
        @test_throws XLSX.XLSXError XLSX.getBorder(s, 3, 2)
        @test_throws XLSX.XLSXError XLSX.getBorder(s, 2, 3)
        s[1:3, 1:3] = ""
        XLSX.setBorder(s, "A1"; allsides=["color" => "grey42", "style" => "thick"])
        @test XLSX.getBorder(s, "A1").border == Dict("left" => Dict("rgb" => "FF6B6B6B", "style" => "thick"), "bottom" => Dict("rgb" => "FF6B6B6B", "style" => "thick"), "right" => Dict("rgb" => "FF6B6B6B", "style" => "thick"), "top" => Dict("rgb" => "FF6B6B6B", "style" => "thick"), "diagonal" => nothing)
        XLSX.setBorder(s, 2, 2; allsides=["color" => "grey43", "style" => "thin"])
        @test XLSX.getBorder(s, 2, 2).border == Dict("left" => Dict("rgb" => "FF6E6E6E", "style" => "thin"), "bottom" => Dict("rgb" => "FF6E6E6E", "style" => "thin"), "right" => Dict("rgb" => "FF6E6E6E", "style" => "thin"), "top" => Dict("rgb" => "FF6E6E6E", "style" => "thin"), "diagonal" => nothing)
        XLSX.setBorder(s, [2, 3], 1:3; allsides=["color" => "grey44", "style" => "hair"], diagonal=["color" => "grey44", "style" => "thin", "direction" => "down"])
        @test XLSX.getBorder(s, 3, 1).border == Dict("left" => Dict("rgb" => "FF707070", "style" => "hair"), "bottom" => Dict("rgb" => "FF707070", "style" => "hair"), "right" => Dict("rgb" => "FF707070", "style" => "hair"), "top" => Dict("rgb" => "FF707070", "style" => "hair"), "diagonal" => Dict("rgb" => "FF707070", "style" => "thin", "direction" => "down"))
        @test XLSX.getBorder(s, 2, 2).border == Dict("left" => Dict("rgb" => "FF707070", "style" => "hair"), "bottom" => Dict("rgb" => "FF707070", "style" => "hair"), "right" => Dict("rgb" => "FF707070", "style" => "hair"), "top" => Dict("rgb" => "FF707070", "style" => "hair"), "diagonal" => Dict("rgb" => "FF707070", "style" => "thin", "direction" => "down"))
        @test XLSX.getBorder(s, 3, 3).border == Dict("left" => Dict("rgb" => "FF707070", "style" => "hair"), "bottom" => Dict("rgb" => "FF707070", "style" => "hair"), "right" => Dict("rgb" => "FF707070", "style" => "hair"), "top" => Dict("rgb" => "FF707070", "style" => "hair"), "diagonal" => Dict("rgb" => "FF707070", "style" => "thin", "direction" => "down"))
        SAVE_FILES && save_outfile(f)

        f = XLSX.newxlsx()
        s = f[1]
        @test_throws XLSX.XLSXError XLSX.setFill(s, "A1"; pattern="lightVertical", fgColor="Red", bgColor="blue")
        @test_throws XLSX.XLSXError XLSX.setFill(s, "A1:A1"; pattern="lightVertical", fgColor="Red", bgColor="blue")
        @test_throws XLSX.XLSXError XLSX.setFill(s, "A"; pattern="lightVertical", fgColor="Red", bgColor="blue")
        @test_throws XLSX.XLSXError XLSX.setFill(s, "1"; pattern="lightVertical", fgColor="Red", bgColor="blue")
        @test_throws XLSX.XLSXError XLSX.setFill(s, 1, 1; pattern="lightVertical", fgColor="Red", bgColor="blue")
        @test_throws XLSX.XLSXError XLSX.setFill(s, [1], 1; pattern="lightVertical", fgColor="Red", bgColor="blue")
        @test_throws XLSX.XLSXError XLSX.setFill(s, 1, 1:1; pattern="lightVertical", fgColor="Red", bgColor="blue")
        @test_throws XLSX.XLSXError XLSX.setFill(s, 1, :; pattern="lightVertical", fgColor="Red", bgColor="blue")
        @test_throws XLSX.XLSXError XLSX.setFill(s, :, :; pattern="lightVertical", fgColor="Red", bgColor="blue")
        @test_throws XLSX.XLSXError XLSX.setFill(s, :; pattern="lightVertical", fgColor="Red", bgColor="blue")
        @test_throws XLSX.XLSXError XLSX.setFill(s, [2, 3], 1:3; pattern="lightVertical", fgColor="Red", bgColor="blue") == -1
        @test_throws XLSX.XLSXError XLSX.getFill(s, 2, 1)
        @test_throws XLSX.XLSXError XLSX.getFill(s, 3, 2)
        @test_throws XLSX.XLSXError XLSX.getFill(s, 2, 3)
        s[1:3, 1:3] = ""
        XLSX.setFill(s, "A1"; pattern="lightVertical", fgColor="Red", bgColor="blue")
        @test XLSX.getFill(s, "A1").fill == Dict("patternFill" => Dict("bgrgb" => "FF0000FF", "patternType" => "lightVertical", "fgrgb" => "FFFF0000"))
        XLSX.setFill(s, 2, 2; pattern="lightGrid", fgColor="Red", bgColor="blue")
        @test XLSX.getFill(s, 2, 2).fill == Dict("patternFill" => Dict("bgrgb" => "FF0000FF", "patternType" => "lightGrid", "fgrgb" => "FFFF0000"))
        XLSX.setFill(s, [2, 3], 1:3; pattern="lightGrid", fgColor="Red", bgColor="blue")
        @test XLSX.getFill(s, 3, 1).fill == Dict("patternFill" => Dict("bgrgb" => "FF0000FF", "patternType" => "lightGrid", "fgrgb" => "FFFF0000"))
        @test XLSX.getFill(s, 2, 2).fill == Dict("patternFill" => Dict("bgrgb" => "FF0000FF", "patternType" => "lightGrid", "fgrgb" => "FFFF0000"))
        @test XLSX.getFill(s, 3, 3).fill == Dict("patternFill" => Dict("bgrgb" => "FF0000FF", "patternType" => "lightGrid", "fgrgb" => "FFFF0000"))
        SAVE_FILES && save_outfile(f)

        f = XLSX.newxlsx()
        s = f[1]
        @test_throws XLSX.XLSXError XLSX.setAlignment(s, "A1"; horizontal="right", vertical="justify", wrapText=true)
        @test_throws XLSX.XLSXError XLSX.setAlignment(s, "A1:A1"; horizontal="right", vertical="justify", wrapText=true)
        @test_throws XLSX.XLSXError XLSX.setAlignment(s, "A"; horizontal="right", vertical="justify", wrapText=true)
        @test_throws XLSX.XLSXError XLSX.setAlignment(s, "1"; horizontal="right", vertical="justify", wrapText=true)
        @test_throws XLSX.XLSXError XLSX.setAlignment(s, 1, 1; horizontal="right", vertical="justify", wrapText=true)
        @test_throws XLSX.XLSXError XLSX.setAlignment(s, [1], 1; horizontal="right", vertical="justify", wrapText=true)
        @test_throws XLSX.XLSXError XLSX.setAlignment(s, 1, 1:1; horizontal="right", vertical="justify", wrapText=true)
        @test_throws XLSX.XLSXError XLSX.setAlignment(s, 1, :; horizontal="right", vertical="justify", wrapText=true)
        @test_throws XLSX.XLSXError XLSX.setAlignment(s, :, :; horizontal="right", vertical="justify", wrapText=true)
        @test_throws XLSX.XLSXError XLSX.setAlignment(s, :; horizontal="right", vertical="justify", wrapText=true)
        @test_throws XLSX.XLSXError XLSX.setAlignment(s, [2, 3], 1:3; horizontal="right", vertical="justify", wrapText=true) == -1
        @test_throws XLSX.XLSXError XLSX.getAlignment(s, 2, 1)
        @test_throws XLSX.XLSXError XLSX.getAlignment(s, 3, 2)
        @test_throws XLSX.XLSXError XLSX.getAlignment(s, 2, 3)
        s[1:3, 1:3] = ""
        XLSX.setAlignment(s, "A1"; horizontal="right", vertical="justify", wrapText=true)
        @test XLSX.getAlignment(s, "A1").alignment == Dict("alignment" => Dict("horizontal" => "right", "vertical" => "justify", "wrapText" => "1"))
        XLSX.setAlignment(s, 2, 2; horizontal="right", vertical="justify", wrapText=true, rotation=90)
        @test XLSX.getAlignment(s, 2, 2).alignment == Dict("alignment" => Dict("horizontal" => "right", "vertical" => "justify", "wrapText" => "1", "textRotation" => "90"))
        XLSX.setAlignment(s, [2, 3], 1:3; horizontal="right", vertical="justify", shrink=true, rotation=90)
        @test XLSX.getAlignment(s, 3, 1).alignment == Dict("alignment" => Dict("horizontal" => "right", "vertical" => "justify", "shrinkToFit" => "1", "textRotation" => "90"))
        @test XLSX.getAlignment(s, 2, 2).alignment == Dict("alignment" => Dict("horizontal" => "right", "vertical" => "justify", "wrapText" => "1", "shrinkToFit" => "1", "textRotation" => "90"))
        @test XLSX.getAlignment(s, 3, 3).alignment == Dict("alignment" => Dict("horizontal" => "right", "vertical" => "justify", "shrinkToFit" => "1", "textRotation" => "90"))
        SAVE_FILES && save_outfile(f)

        f = XLSX.newxlsx()
        s = f[1]
        @test_throws XLSX.XLSXError XLSX.setFormat(s, "A1"; format="Percentage")
        @test_throws XLSX.XLSXError XLSX.setFormat(s, "A1:A1"; format="Percentage")
        @test_throws XLSX.XLSXError XLSX.setFormat(s, "A"; format="Percentage")
        @test_throws XLSX.XLSXError XLSX.setFormat(s, "1"; format="Percentage")
        @test_throws XLSX.XLSXError XLSX.setFormat(s, 1, 1; format="Percentage")
        @test_throws XLSX.XLSXError XLSX.setFormat(s, [1], 1; format="Percentage")
        @test_throws XLSX.XLSXError XLSX.setFormat(s, 1, 1:1; format="Percentage")
        @test_throws XLSX.XLSXError XLSX.setFormat(s, 1, :; format="Percentage")
        @test_throws XLSX.XLSXError XLSX.setFormat(s, :, :; format="Percentage")
        @test_throws XLSX.XLSXError XLSX.setFormat(s, :; format="Percentage")
        @test_throws XLSX.XLSXError XLSX.setFormat(s, [2, 3], 1:3; format="Percentage") == -1
        @test_throws XLSX.XLSXError XLSX.getFormat(s, 2, 1)
        @test_throws XLSX.XLSXError XLSX.getFormat(s, 3, 2)
        @test_throws XLSX.XLSXError XLSX.getFormat(s, 2, 3)
        s[1:3, 1:3] = ""
        XLSX.setFormat(s, "A1"; format="#,##0.000;(#,##0.000)")
        @test XLSX.getFormat(s, "A1").format == Dict("numFmt" => Dict("numFmtId" => "164", "formatCode" => "#,##0.000;(#,##0.000)"))
        XLSX.setFormat(s, 2, 2; format="Currency")
        @test XLSX.getFormat(s, 2, 2).format == Dict("numFmt" => Dict("numFmtId" => "7", "formatCode" => "\$#,##0.00_);(\$#,##0.00)"))
        XLSX.setFormat(s, [2, 3], 1:3; format="LongDate")
        @test XLSX.getFormat(s, 3, 1).format == Dict("numFmt" => Dict("numFmtId" => "15", "formatCode" => "d-mmm-yy"))
        @test XLSX.getFormat(s, 2, 2).format == Dict("numFmt" => Dict("numFmtId" => "15", "formatCode" => "d-mmm-yy"))
        @test XLSX.getFormat(s, 3, 3).format == Dict("numFmt" => Dict("numFmtId" => "15", "formatCode" => "d-mmm-yy"))
        SAVE_FILES && save_outfile(f)

        f = XLSX.newxlsx()
        s = f[1]
        @test_throws XLSX.XLSXError XLSX.getColumnWidth(s, "B2") # Cell outside sheet dimension
        s[1:3, 1:3] = ""
        XLSX.setColumnWidth(s, "A1"; width=30)
        @test XLSX.getColumnWidth(s, "A1") ≈ 30.7109375
        XLSX.setColumnWidth(s, 2, 2; width=40)
        @test XLSX.getColumnWidth(s, 2, 2) ≈ 40.7109375
        XLSX.setColumnWidth(s, [2, 3], 1:3; width=50)
        @test XLSX.getColumnWidth(s, 3, 1) ≈ 50.7109375
        @test XLSX.getColumnWidth(s, 2, 2) ≈ 50.7109375
        @test XLSX.getColumnWidth(s, 3, 3) ≈ 50.7109375
        SAVE_FILES && save_outfile(f)

        f = XLSX.newxlsx()
        s = f[1]
        @test_throws XLSX.XLSXError XLSX.getRowHeight(s, "B2") # Cell outside sheet dimension
        s[1:3, 1:3] = ""
        XLSX.setRowHeight(s, "A1"; height=30)
        @test XLSX.getRowHeight(s, "A1") ≈ 30.2109375
        XLSX.setRowHeight(s, 2, 2; height=40)
        @test XLSX.getRowHeight(s, 2, 2) ≈ 40.2109375
        XLSX.setRowHeight(s, [2, 3], 1:3; height=50)
        @test XLSX.getRowHeight(s, 3, 1) ≈ 50.2109375
        @test XLSX.getRowHeight(s, 2, 2) ≈ 50.2109375
        @test XLSX.getRowHeight(s, 3, 3) ≈ 50.2109375
        SAVE_FILES && save_outfile(f)

        f = XLSX.newxlsx()
        s = f[1]
        s[1:30, 1:26] = ""
        XLSX.setUniformFont(s, 1:4, :; size=12, name="Times New Roman", color="FF040404")
        @test XLSX.getFont(f, "Sheet1!A1").font == Dict("sz" => Dict("val" => "12"), "name" => Dict("val" => "Times New Roman"), "color" => Dict("rgb" => "FF040404"))
        @test XLSX.getFont(f, "Sheet1!G2").font == Dict("sz" => Dict("val" => "12"), "name" => Dict("val" => "Times New Roman"), "color" => Dict("rgb" => "FF040404"))
        @test XLSX.getFont(f, "Sheet1!N3").font == Dict("sz" => Dict("val" => "12"), "name" => Dict("val" => "Times New Roman"), "color" => Dict("rgb" => "FF040404"))
        @test XLSX.getFont(f, "Sheet1!Y4").font == Dict("sz" => Dict("val" => "12"), "name" => Dict("val" => "Times New Roman"), "color" => Dict("rgb" => "FF040404"))

        XLSX.setUniformFill(s, :, 2:8; pattern="lightGrid", fgColor="FF0000FF", bgColor="FF00FF00")
        @test XLSX.getFill(s, "B10").fill == Dict("patternFill" => Dict("bgrgb" => "FF00FF00", "patternType" => "lightGrid", "fgrgb" => "FF0000FF"))
        @test XLSX.getFill(s, "D20").fill == Dict("patternFill" => Dict("bgrgb" => "FF00FF00", "patternType" => "lightGrid", "fgrgb" => "FF0000FF"))
        @test XLSX.getFill(s, "F30").fill == Dict("patternFill" => Dict("bgrgb" => "FF00FF00", "patternType" => "lightGrid", "fgrgb" => "FF0000FF"))

        XLSX.setUniformFormat(s, :; format="#,##0.000")
        @test XLSX.getFormat(s, "A1").format == Dict("numFmt" => Dict("numFmtId" => "164", "formatCode" => "#,##0.000"))
        @test XLSX.getFormat(s, "G10").format == Dict("numFmt" => Dict("numFmtId" => "164", "formatCode" => "#,##0.000"))
        @test XLSX.getFormat(s, "M20").format == Dict("numFmt" => Dict("numFmtId" => "164", "formatCode" => "#,##0.000"))
        @test XLSX.getFormat(s, "X30").format == Dict("numFmt" => Dict("numFmtId" => "164", "formatCode" => "#,##0.000"))
        SAVE_FILES && save_outfile(f)

        f = XLSX.open_xlsx_template(joinpath(data_directory, "Borders.xlsx"))
        s = f["Sheet1"]
        XLSX.setUniformBorder(s, [1, 2, 3, 4], 1:4; left=["style" => "dotted", "color" => "darkseagreen3"],
            right=["style" => "medium", "color" => "FF765000"],
            top=["style" => "thick", "color" => "FF230000"],
            bottom=["style" => "medium", "color" => "FF0000FF"],
            diagonal=["style" => "none"]
        )
        @test XLSX.getBorder(s, 1, 1).border == Dict("left" => Dict("style" => "dotted", "rgb" => "FF9BCD9B"), "bottom" => Dict("style" => "medium", "rgb" => "FF0000FF"), "right" => Dict("style" => "medium", "rgb" => "FF765000"), "top" => Dict("style" => "thick", "rgb" => "FF230000"), "diagonal" => nothing)
        @test XLSX.getBorder(s, 2, 2).border == Dict("left" => Dict("style" => "dotted", "rgb" => "FF9BCD9B"), "bottom" => Dict("style" => "medium", "rgb" => "FF0000FF"), "right" => Dict("style" => "medium", "rgb" => "FF765000"), "top" => Dict("style" => "thick", "rgb" => "FF230000"), "diagonal" => nothing)
        @test XLSX.getBorder(s, 4, 4).border == Dict("left" => Dict("style" => "dotted", "rgb" => "FF9BCD9B"), "bottom" => Dict("style" => "medium", "rgb" => "FF0000FF"), "right" => Dict("style" => "medium", "rgb" => "FF765000"), "top" => Dict("style" => "thick", "rgb" => "FF230000"), "diagonal" => nothing)
        SAVE_FILES && save_outfile(f)

    end

    @testset "existing formatting" begin
        f = XLSX.opentemplate(joinpath(data_directory, "customXml.xlsx"))
        s = f[1]
        s["B2"] = pi
        s["D20"] = "Hello World"
        s["J45"] = Dates.Date(2025, 01, 24)
        @test XLSX.getFont(s, "B2").font == Dict("name" => Dict("val" => "Calibri"), "family" => Dict("val" => "2"), "b" => nothing, "sz" => Dict("val" => "18"), "color" => Dict("theme" => "1"), "scheme" => Dict("val" => "minor"))
        @test XLSX.getFill(s, "D20").fill == Dict("patternFill" => Dict("bgindexed" => "64", "patternType" => "solid", "fgtint" => "-0.499984740745262", "fgtheme" => "2"))
        @test XLSX.getBorder(s, "J45").border == Dict("left" => Dict("indexed" => "64", "style" => "thin"), "bottom" => Dict("indexed" => "64", "style" => "thin"), "right" => Dict("indexed" => "64", "style" => "thin"), "top" => Dict("indexed" => "64", "style" => "thin"), "diagonal" => nothing)
        SAVE_FILES && save_outfile(f)
    end
    @testset "Styles caching (issue #426)" begin

        # Builds a workbook with K rows in column 1, each holding a numeric value.
        # `distinct=true` gives every row its own custom numFmt (K distinct cell
        # styles); `distinct=false` gives every row the same custom numFmt (1 style).
        function build_styled_workbook(K::Int; distinct::Bool)
            path = tempname() * ".xlsx"
            f = XLSX.newxlsx()
            sh = f[1]
            for i in 1:K
                sh[i, 1] = 1.5
                fmt = distinct ? "\"v$(i)\"0.00" : "0.00"
                XLSX.setFormat(sh, i, 1; format=fmt)
            end
            XLSX.writexlsx(path, f; overwrite=true)
            return path
        end

        @testset "correctness with many distinct custom formats" begin
            K = 60
            path = build_styled_workbook(K; distinct=true)
            try
                XLSX.openxlsx(path) do xf
                    sh = xf[1]
                    for i in 1:K
                        @test XLSX.getdata(sh, i, 1) isa Float64
                        fmt = XLSX.getFormat(sh, i, 1)
                        @test fmt.format["numFmt"]["formatCode"] == "\"v$(i)\"0.00"
                    end
                end
            finally
                rm(path; force=true)
            end
        end
        @testset "cellXfs/numFmt caches behave like caches" begin
            K = 20
            path = build_styled_workbook(K; distinct=true)
            try
                XLSX.openxlsx(path) do xf
                    wb = XLSX.get_workbook(xf)

                    nodes1 = XLSX.get_cellXfs_nodes(wb)
                    nodes2 = XLSX.get_cellXfs_nodes(wb)
                    @test nodes1 === nodes2                     # same object: not rebuilt on 2nd call
                    @test length(nodes1) >= K                    # at least one xf per distinct style

                    cache1 = XLSX.get_numFmt_cache(wb)
                    cache2 = XLSX.get_numFmt_cache(wb)
                    @test cache1 === cache2
                    @test length(cache1) >= K
                end
            finally
                rm(path; force=true)
            end
        end

        @testset "regression: numFmt cache stays in sync across both write paths" begin
            # styles_add_numFmt creates the first <numFmts> node.
            # styles_add_cell_attribute(wb, ..., "numFmts") is the *separate* path
            # used once <numFmts> already exists (e.g. via conditional formatting).
            # Both must keep wb.numFmt_cache correct — this reproduces the
            # "numFmtId ... not found" bug from testing the original patch.
            path = tempname() * ".xlsx"
            f = XLSX.newxlsx()
            sh = f[1]
            sh["A1"] = 1.5
            sh["A2"] = 2.5

            XLSX.setFormat(sh, 1, 1; format="0.00")               # -> styles_add_numFmt path
            wb = XLSX.get_workbook(sh)
            XLSX.get_numFmt_cache(wb)                              # force-build the cache now,
                                                                    # so the *next* addition must
                                                                    # go through the update path,
                                                                    # not just a fresh build.
            @test XLSX.setConditionalFormat(sh, "A2", :cellIs;
                operator="greaterThan", value="2",
                format=["format" => "0.0"]) == 0

            XLSX.writexlsx(path, f; overwrite=true)
            try
                XLSX.openxlsx(path) do xf
                    @test length(XLSX.getConditionalFormats(xf[1])) == 1
                end
            finally
                rm(path; force=true)
            end
        end

        @testset "reading distinct styles scales ~linearly, not quadratically" begin
            function time_read(path)
                XLSX.openxlsx(x -> XLSX.getdata(x[1]), path)  # warm-up
                times = [@elapsed(XLSX.openxlsx(x -> XLSX.getdata(x[1]), path)) for _ in 1:9]
                sort!(times)
                return times[5]  # median of 9, more robust to one slow outlier than min-of-5
            end

            Ks = (200, 400, 800, 1600)
            times = Float64[]
            for K in Ks
                path = build_styled_workbook(K; distinct=true)
                try
                    push!(times, time_read(path))
                finally
                    rm(path; force=true)
                end
            end

            ratios = [times[i+1] / times[i] for i in 1:length(times)-1]
            @test all(r -> r < 3.5, ratios)
        end
        @testset "cellXfs/numFmt caches are built exactly once regardless of K" begin
            K = 50
            path = build_styled_workbook(K; distinct=true)
            try
                XLSX.openxlsx(path) do xf
                    wb = XLSX.get_workbook(xf)
                    nodes_a = XLSX.get_cellXfs_nodes(wb)
                    for i in 1:K
                        XLSX.getFormat(xf[1], i, 1)   # would previously re-walk cellXfs each time
                    end
                    nodes_b = XLSX.get_cellXfs_nodes(wb)
                    @test nodes_a === nodes_b   # same object: never rebuilt mid-loop
                end
            finally
                rm(path; force=true)
            end
        end
    end
    @testset "Font/Border/Fill caching" begin

        # Builds a workbook with K rows in column 1, each holding a numeric value.
        # `distinct=true` gives every row its own font/border/fill (K distinct
        # style attributes); `distinct=false` gives every row the same attribute.
        function build_font_workbook(K::Int; distinct::Bool)
            path = tempname() * ".xlsx"
            f = XLSX.newxlsx()
            sh = f[1]
            for i in 1:K
                sh[i, 1] = 1.5
                if distinct
                    col = "FF" * uppercase(lpad(string(i % 0x1000000, base=16), 6, '0'))
                    XLSX.setFont(sh, i, 1; size=12, color=col)
                else
                    XLSX.setFont(sh, i, 1; size=12)
                end
            end
            XLSX.writexlsx(path, f; overwrite=true)
            return path
        end

        function build_border_workbook(K::Int; distinct::Bool)
            path = tempname() * ".xlsx"
            f = XLSX.newxlsx()
            sh = f[1]
            for i in 1:K
                sh[i, 1] = 1.5
                col = distinct ? "FF" * uppercase(lpad(string(i % 0x1000000, base=16), 6, '0')) : "FFFF0000"
                XLSX.setBorder(sh, i, 1; allsides=["style"=>"thin", "color"=>col])
            end
            XLSX.writexlsx(path, f; overwrite=true)
            return path
        end

        function build_fill_workbook(K::Int; distinct::Bool)
            path = tempname() * ".xlsx"
            f = XLSX.newxlsx()
            sh = f[1]
            for i in 1:K
                sh[i, 1] = 1.5
                col = distinct ? "FF" * uppercase(lpad(string(i % 0x1000000, base=16), 6, '0')) : "FF00FF00"
                XLSX.setFill(sh, i, 1; pattern="solid", fgColor=col)
            end
            XLSX.writexlsx(path, f; overwrite=true)
            return path
        end

        @testset "correctness with many distinct fonts" begin
            K = 60
            path = build_font_workbook(K; distinct=true)
            try
                XLSX.openxlsx(path) do xf
                    sh = xf[1]
                    for i in 1:K
                        @test XLSX.getdata(sh, i, 1) isa Float64
                        col = "FF" * uppercase(lpad(string(i % 0x1000000, base=16), 6, '0'))
                        font = XLSX.getFont(sh, i, 1)
                        @test font.font["color"]["rgb"] == col
                    end
                end
            finally
                rm(path; force=true)
            end
        end
        @testset "fonts/borders/fills caches behave like caches" begin
            K = 20
            fpath = build_font_workbook(K; distinct=true)
            bpath = build_border_workbook(K; distinct=true)
            gpath = build_fill_workbook(K; distinct=true)
            try
                XLSX.openxlsx(fpath) do xf
                    wb = XLSX.get_workbook(xf)
                    nodes1 = XLSX.get_fonts_nodes(wb)
                    nodes2 = XLSX.get_fonts_nodes(wb)
                    @test nodes1 === nodes2   # same object: not rebuilt on 2nd call
                    @test length(nodes1) >= K
                end
                XLSX.openxlsx(bpath) do xf
                    wb = XLSX.get_workbook(xf)
                    nodes1 = XLSX.get_borders_nodes(wb)
                    nodes2 = XLSX.get_borders_nodes(wb)
                    @test nodes1 === nodes2
                    @test length(nodes1) >= K
                end
                XLSX.openxlsx(gpath) do xf
                    wb = XLSX.get_workbook(xf)
                    nodes1 = XLSX.get_fills_nodes(wb)
                    nodes2 = XLSX.get_fills_nodes(wb)
                    @test nodes1 === nodes2
                    @test length(nodes1) >= K
                end
            finally
                rm(fpath; force=true); rm(bpath; force=true); rm(gpath; force=true)
            end
        end

        @testset "regression: style_table_cache stays in sync across writes" begin
            # Mirrors the numFmt_cache regression test: force-build the cache,
            # then write a *new* distinct font — the addition must go through the
            # push!-sync path in styles_add_cell_attribute, not just a fresh build.
            path = tempname() * ".xlsx"
            f = XLSX.newxlsx()
            sh = f[1]
            sh["A1"] = 1.5
            sh["A2"] = 2.5

            XLSX.setFont(sh, 1, 1; size=14)
            wb = XLSX.get_workbook(sh)
            XLSX.get_fonts_nodes(wb)          # force-build the cache now
            XLSX.setFont(sh, 2, 1; size=16)   # distinct font -> new <font> element

            XLSX.writexlsx(path, f; overwrite=true)
            try
                XLSX.openxlsx(path) do xf
                    sh2 = xf[1]
                    @test parse(Int, XLSX.getFont(sh2, 1, 1).font["sz"]["val"]) == 14
                    @test parse(Int, XLSX.getFont(sh2, 2, 1).font["sz"]["val"]) == 16
                end
            finally
                rm(path; force=true)
            end
        end

        @testset "reading distinct fonts scales ~linearly, not quadratically" begin
            function time_read(path)
                XLSX.openxlsx(x -> XLSX.getdata(x[1]), path)  # warm-up
                times = [@elapsed(XLSX.openxlsx(x -> XLSX.getdata(x[1]), path)) for _ in 1:9]
                sort!(times)
                return times[5]
            end

            Ks = (200, 400, 800, 1600)
            times = Float64[]
            for K in Ks
                path = build_font_workbook(K; distinct=true)
                try
                    push!(times, time_read(path))
                finally
                    rm(path; force=true)
                end
            end

            ratios = [times[i+1] / times[i] for i in 1:length(times)-1]
            @test all(r -> r < 3.5, ratios)
        end

        @testset "fonts/borders/fills caches are built exactly once regardless of K" begin
            K = 50
            path = build_font_workbook(K; distinct=true)
            try
                XLSX.openxlsx(path) do xf
                    wb = XLSX.get_workbook(xf)
                    nodes_a = XLSX.get_fonts_nodes(wb)
                    for i in 1:K
                        XLSX.getFont(xf[1], i, 1)   # would previously re-walk fonts each time
                    end
                    nodes_b = XLSX.get_fonts_nodes(wb)
                    @test nodes_a === nodes_b   # same object: never rebuilt mid-loop
                end
            finally
                rm(path; force=true)
            end
        end
    end
end