@testset "Data validation" begin
    fresh() = (xf = XLSX.newxlsx("Data"); XLSX.addsheet!(xf, "Lists"); (xf, xf["Data"]))
    sheet_root(xf, n=1) = XLSX.xml_root_element(xf.data["xl/worksheets/sheet$n.xml"])
    sheet_tags(xf, n=1) = [XLSX.localname(c) for c in XLSX.xml_elements(sheet_root(xf, n))]
    dv_nodes(xf, n=1) = (b = XLSX.first_element_with_tag(sheet_root(xf, n), "dataValidations");
                         isnothing(b) ? XML.Node[] : XLSX.elements_with_tag(b, "dataValidation"))
    function roundtrip_bytes(xf)
        io = IOBuffer()
        XLSX.writexlsx(io, xf)
        return take!(io)
    end

    @testset "list from values" begin
        xf, s = fresh()
        dv = XLSX.addDataValidation(s, "B2:B100"; values=["Yes", "No", "Say \"maybe\""], prompt="Pick one", promptTitle="Answer")
        @test dv.sheet == "Data"
        @test dv.range == XLSX.CellRange("B2:B100")
        @test dv.type == :list && isnothing(dv.operator)
        @test dv.formula1 == "\"Yes,No,Say \"\"maybe\"\"\""
        @test isnothing(dv.formula2)
        @test dv.allowBlank && dv.dropdown && dv.showInputMessage && dv.showErrorMessage
        @test dv.prompt == "Pick one" && dv.promptTitle == "Answer"
        @test isnothing(dv.error) && isnothing(dv.errorTitle)
        @test dv.errorStyle == :stop
        @test XLSX.getDataValidations(s) == [dv]
        n = only(dv_nodes(xf))
        @test n["type"] == "list"
        @test n["sqref"] == "B2:B100"
        @test !haskey(n, "showDropDown")
        @test only([XML.simple_value(c) for c in XLSX.xml_elements(n)]) == dv.formula1
        SAVE_FILES && save_outfile(xf)
    end

    @testset "list from cells, without the drop-down" begin
        xf, s = fresh()
        dv = XLSX.addDataValidation(s, "C2:C10"; type=:list, formula1="=Lists!\$A\$1:\$A\$12", dropdown=false)
        @test dv.formula1 == "Lists!\$A\$1:\$A\$12"   # stored without the `=`
        @test !dv.dropdown
        @test only(dv_nodes(xf))["showDropDown"] == "1"   # OOXML's inverted flag
        SAVE_FILES && save_outfile(xf)
    end

    @testset "numbers, dates, times and text length" begin
        xf, s = fresh()
        a = XLSX.addDataValidation(s, "A1:A5"; type=:whole, formula1=1, formula2=10)
        @test (a.type, a.operator, a.formula1, a.formula2) == (:whole, :between, "1", "10")
        @test !haskey(dv_nodes(xf)[1], "operator")   # `between` is the default, as Excel writes it
        b = XLSX.addDataValidation(s, "B1"; type=:decimal, operator=:lessThanOrEqual, formula1=2.5)
        @test (b.operator, b.formula1, b.range) == (:lessThanOrEqual, "2.5", XLSX.CellRange("B1:B1"))
        @test dv_nodes(xf)[2]["sqref"] == "B1"
        c = XLSX.addDataValidation(s, "C1"; type=:date, operator=:greaterThan, formula1=Date(2025, 1, 1))
        @test c.formula1 == "45658"
        d = XLSX.addDataValidation(s, "D1"; type=:time, operator=:notBetween, formula1=Time(9), formula2=Time(17, 30))
        @test (d.formula1, d.formula2) == ("0.375", string(17.5 / 24))
        e = XLSX.addDataValidation(s, "E1"; type=:textLength, operator=:lessThan, formula1="=B1")
        @test e.formula1 == "B1"
        f = XLSX.addDataValidation(s, "F1"; type=:date, operator=:equal, formula1=DateTime(2025, 1, 1, 12))
        @test f.formula1 == "45658.5"
        SAVE_FILES && save_outfile(xf)
    end

    @testset "custom and any" begin
        xf, s = fresh()
        c = XLSX.addDataValidation(s, "F2:F50"; type=:custom, formula1="=ISNUMBER(F2)", error="Numbers only",
                                   errorTitle="Oops", errorStyle=:warning)
        @test (c.type, c.formula1, c.error, c.errorTitle, c.errorStyle) == (:custom, "ISNUMBER(F2)", "Numbers only", "Oops", :warning)
        a = XLSX.addDataValidation(s, "G1"; type=:any, prompt="Type your name", showErrorMessage=false, allowBlank=false)
        @test (a.type, a.operator, a.formula1, a.prompt) == (:any, nothing, nothing, "Type your name")
        @test !a.showErrorMessage && !a.allowBlank
        @test !haskey(dv_nodes(xf)[2], "type")
        @test isempty(XLSX.xml_elements(dv_nodes(xf)[2]))
    end

    @testset "range forms" begin
        xf, s = fresh()
        @test XLSX.addDataValidation(s, "D:E"; type=:whole, formula1=0, operator=:greaterThan).range ==
              XLSX.CellRange(XLSX.CellRef(1, 4), XLSX.CellRef(XLSX.EXCEL_MAX_ROWS, 5))
        nc = XLSX.addDataValidation(xf, "Data!G5:G6,Data!J5:J6"; values=["a"]).range
        @test nc isa XLSX.NonContiguousRange
        @test dv_nodes(xf)[2]["sqref"] == "G5:G6 J5:J6"
        @test XLSX.addDataValidation(s, XLSX.CellRef("K1"); values=["b"]).range == XLSX.CellRange("K1:K1")
        @test XLSX.addDataValidation(xf, "Data!L1:L2"; values=["c"]).range == XLSX.CellRange("L1:L2")
        XLSX.addDefinedName(s, "Inputs", "M1:M4")
        @test XLSX.addDataValidation(s, "Inputs"; values=["d"]).range == XLSX.CellRange("M1:M4")
        xf2, s2 = fresh()
        @test XLSX.addDataValidation(s2, "3:3"; type=:any, prompt="p").range ==
              XLSX.CellRange(XLSX.CellRef(3, 1), XLSX.CellRef(3, XLSX.EXCEL_MAX_COLS))
        @test only(dv_nodes(xf2))["sqref"] == "A3:XFD3"
        @test_throws XLSX.XLSXError XLSX.addDataValidation(s, "Lists!A1"; values=["e"])
        @test_throws XLSX.XLSXError XLSX.addDataValidation(s, "notarange"; values=["e"])
    end

    @testset "argument checks" begin
        xf, s = fresh()
        bad = [
            (;),                                                        # no type
            (; type=:number, formula1=1),                               # unknown type
            (; type=:whole, values=["a"]),                              # values on a non-list
            (; values=["a"], formula1="A1"),                            # both
            (; values=String[]),                                        # empty list
            (; values=["a,b"]),                                         # comma in an item
            (; values=[repeat("x", 100), repeat("y", 100), repeat("z", 100)]),   # over 255
            (; type=:list),                                             # no formula1
            (; type=:list, formula1="A1", operator=:equal),             # operator on a list
            (; type=:list, formula1="A1", formula2="A2"),               # formula2 on a list
            (; type=:custom),                                           # no formula1
            (; type=:whole),                                            # no formula1
            (; type=:whole, formula1=1),                                # between needs formula2
            (; type=:whole, operator=:equal, formula1=1, formula2=2),   # formula2 not wanted
            (; type=:whole, operator=:near, formula1=1),                # unknown operator
            (; type=:whole, formula1=true, formula2=2),                 # Bool formula
            (; type=:whole, formula1=[1], formula2=2),                  # unsupported formula type
            (; type=:whole, formula1="=", formula2=2),                  # empty formula
            (; type=:any, formula1=1),                                  # :any takes no formula
            (; type=:whole, operator=:equal, formula1=1, dropdown=false),   # dropdown on a non-list
            (; values=["a"], errorStyle=:fatal),                        # unknown error style
            (; values=["a"], promptTitle=repeat("t", 33)),              # titles ≤ 32
            (; values=["a"], error=repeat("e", 256)),                   # messages ≤ 255
        ]
        for kw in bad
            @test_throws XLSX.XLSXError XLSX.addDataValidation(s, "A1"; kw...)
        end
        @test isempty(XLSX.getDataValidations(s))
        @test !("dataValidations" in sheet_tags(xf))
    end

    @testset "one validation per cell" begin
        xf, s = fresh()
        XLSX.addDataValidation(s, "B2:B10"; values=["a"])
        @test_throws XLSX.XLSXError XLSX.addDataValidation(s, "B10:C12"; values=["b"])
        @test_throws XLSX.XLSXError XLSX.addDataValidation(s, "A1:A2,B5"; values=["b"])
        @test_throws XLSX.XLSXError XLSX.addDataValidation(s, "B:B"; values=["b"])
        XLSX.addDataValidation(s, "B11"; values=["b"])   # adjacent is fine
        @test length(XLSX.getDataValidations(s)) == 2
    end

    @testset "getDataValidation" begin
        xf, s = fresh()
        XLSX.addDataValidation(s, "B2:B10,D2:D10"; values=["a"])
        @test XLSX.getDataValidation(s, "D5").formula1 == "\"a\""
        @test XLSX.getDataValidation(s, 3, 2).formula1 == "\"a\""
        @test isnothing(XLSX.getDataValidation(s, "C5"))
        @test_throws XLSX.XLSXError XLSX.getDataValidation(s, "B2:B3")
    end

    @testset "delete" begin
        xf, s = fresh()
        XLSX.addDataValidation(s, "B2:B10"; values=["a"])
        XLSX.addDataValidation(s, "C2:C10"; values=["b"])
        XLSX.addDataValidation(s, "E1:E5,G1:G5"; values=["c"])
        @test XLSX.deleteDataValidation(s, "B2:B5") == 0      # only part of B2:B10
        @test XLSX.deleteDataValidation(s, "A1:B20") == 1
        @test XLSX.deleteDataValidation(s, "E1:E5") == 0      # G1:G5 is outside
        @test XLSX.deleteDataValidation(s, "E:G") == 1
        @test only([n["sqref"] for n in dv_nodes(xf)]) == "C2:C10"
        @test XLSX.first_element_with_tag(sheet_root(xf), "dataValidations")["count"] == "1"
        @test XLSX.deleteDataValidation(xf, "Data!C2") == 0
        @test XLSX.deleteDataValidation(s, :) == 1
        @test isempty(XLSX.getDataValidations(s))
        @test !("dataValidations" in sheet_tags(xf))
        @test XLSX.deleteDataValidation(s, :) == 0
    end

    @testset "count attribute and schema position" begin
        xf, s = fresh()
        s["A1:A3"] = [1, 2, 3]
        XLSX.setConditionalFormat(s, "A1:A3", :colorScale)
        XLSX.addHyperlink(s, "D1"; url="https://julialang.org")
        XLSX.addDataValidation(s, "B1"; values=["a"])
        XLSX.addDataValidation(s, "C1"; values=["b"])
        @test XLSX.first_element_with_tag(sheet_root(xf), "dataValidations")["count"] == "2"
        tags = sheet_tags(xf)
        @test findfirst(==("conditionalFormatting"), tags) < findfirst(==("dataValidations"), tags) <
              findfirst(==("hyperlinks"), tags) < findfirst(==("pageMargins"), tags)
    end

    @testset "round trip, read-only, with and without the cache" begin
        xf, s = fresh()
        s["A1"] = 7
        XLSX.addDataValidation(s, "B2:B5"; values=["Yes", "No"], prompt="?")
        XLSX.addDataValidation(s, "C:C"; type=:decimal, operator=:greaterThan, formula1=0, errorStyle=:information)
        expected = XLSX.getDataValidations(s)
        bytes = roundtrip_bytes(xf)
        for enable_cache in (true, false)
            r = XLSX.openxlsx(IOBuffer(bytes); enable_cache)
            @test XLSX.getDataValidations(r["Data"]) == expected
            @test XLSX.getDataValidations(r) == expected
            @test r["Data"]["A1"] == 7
        end
        @test_throws XLSX.XLSXError XLSX.addDataValidation(XLSX.readxlsx(IOBuffer(bytes))["Data"], "D1"; values=["x"])
        @test_throws XLSX.XLSXError XLSX.deleteDataValidation(XLSX.readxlsx(IOBuffer(bytes))["Data"], :)
        SAVE_FILES && save_outfile(xf)
    end

    @testset "Excel 2010 (x14) validations are read and deleted" begin
        xf, s = fresh()
        XLSX.addDataValidation(s, "A1"; values=["std"])
        ext = """<extLst><ext uri="{CCE6A557-97BC-4b89-ADB6-D9C93CAAB3DF}" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"><x14:dataValidations count="1" xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main"><x14:dataValidation type="list" allowBlank="1" showInputMessage="1" showErrorMessage="1"><x14:formula1><xm:f>Lists!\$A\$1:\$A\$3</xm:f></x14:formula1><xm:sqref>B2:B5 D2</xm:sqref></x14:dataValidation></x14:dataValidations></ext></extLst>"""
        push!(sheet_root(xf), XLSX.xml_root_element(parse(ext, XML.Node)))

        dvs = XLSX.getDataValidations(s)
        @test length(dvs) == 2
        x = dvs[2]
        @test x.type == :list && x.formula1 == "Lists!\$A\$1:\$A\$3"
        @test x.range isa XLSX.NonContiguousRange
        @test XLSX.getDataValidation(s, "D2") == x

        # Read from a file opened for reading only
        bytes = roundtrip_bytes(xf)
        @test XLSX.getDataValidations(XLSX.readxlsx(IOBuffer(bytes))) == dvs
        @test XLSX.getDataValidations(XLSX.openxlsx(IOBuffer(bytes); enable_cache=false)) == dvs

        # Overlap is checked against them too
        @test_throws XLSX.XLSXError XLSX.addDataValidation(s, "B3"; values=["new"])

        # renamesheet! updates their formulas (handled for every `extLst` formula)
        XLSX.renamesheet!(xf["Lists"], "Choices")
        @test XLSX.getDataValidation(s, "B2").formula1 == "Choices!\$A\$1:\$A\$3"

        # Deleting the last x14 validation removes its <ext> and the emptied <extLst>
        @test XLSX.deleteDataValidation(s, "A1:D5") == 2
        @test isempty(XLSX.getDataValidations(s))
        @test !("extLst" in sheet_tags(xf))
        @test !("dataValidations" in sheet_tags(xf))
    end

    @testset "sheet operations keep formulas right" begin
        xf, s = fresh()
        XLSX.addDataValidation(s, "B2:B5"; type=:list, formula1="Lists!\$A\$1:\$A\$3")
        XLSX.addDataValidation(s, "C2"; type=:custom, formula1="Data!\$A\$1>0")

        c = XLSX.copysheet!(s, "Copy")
        @test [d.formula1 for d in XLSX.getDataValidations(c)] == ["Lists!\$A\$1:\$A\$3", "Copy!\$A\$1>0"]
        @test all(d -> d.sheet == "Copy", XLSX.getDataValidations(c))

        XLSX.renamesheet!(xf["Lists"], "Options")
        @test XLSX.getDataValidation(s, "B3").formula1 == "Options!\$A\$1:\$A\$3"
        @test XLSX.getDataValidation(c, "B3").formula1 == "Options!\$A\$1:\$A\$3"

        XLSX.deletesheet!(xf, "Options")
        @test XLSX.getDataValidation(s, "B3").formula1 == "#REF!"
        @test length(XLSX.getDataValidations(xf)) == 4
    end

    @testset "chartsheets have none" begin
        xf = XLSX.newxlsx("data")
        xf["data"]["A1:A2"] = [1, 2]
        c = XLSX.Charts.addChart(xf, :line; sheetname="Trend")
        XLSX.Charts.addSeries(c, "data!A1:A2")
        @test isempty(XLSX.getDataValidations(xf["Trend"]))
        @test_throws XLSX.XLSXError XLSX.addDataValidation(xf["Trend"], "A1"; values=["a"])
    end
end
