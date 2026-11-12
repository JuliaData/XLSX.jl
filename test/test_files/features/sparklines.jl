@testset "sparklines" begin
    fixture(name; mode="rw") = XLSX.openxlsx(joinpath(data_directory, "sparklines_$name.xlsx"); mode)
    # The sparkline ext as XML text, without Excel's revision ids or layout whitespace.
    ext_xml(ws) = replace(XML.write(XLSX._worksheet_ext(XLSX._sheet_root(ws), XLSX.SPARKLINE_EXT_URI); indentsize=0),
                          r" xr2:uid=\"[^\"]*\"" => "", r">\s+<" => "><")
    groups_xml(ws) = sort([replace(XML.write(g; indentsize=0), r" xr2:uid=\"[^\"]*\"" => "", r">\s+<" => "><")
                           for g in XLSX.xml_elements(only(XLSX.xml_elements(XLSX._worksheet_ext(XLSX._sheet_root(ws), XLSX.SPARKLINE_EXT_URI))))])
    # Each group's type and sparklines, whatever the order Excel keeps them in.
    sparks(ws) = Set((g.type, Set(g.sparklines)) for g in XLSX.getSparklines(ws))
    exts(ws) = [e["uri"] for x in XLSX.xml_elements(XLSX._sheet_root(ws)) if XLSX.localname(x) == "extLst" for e in XLSX.xml_elements(x)]
    function roundtrip(xf)
        io = IOBuffer()
        XLSX.writexlsx(io, xf)
        return XLSX.readxlsx(IOBuffer(take!(io)))
    end
    function data_sheet()
        xf = XLSX.newxlsx()
        s = xf[1]
        for r in 1:3, c in 1:5
            s[r, c] = mod(r * 7 + c * 3, 9) - 3
        end
        return xf, s
    end

    @testset "as Excel writes them" begin
        xf, s = data_sheet()
        @test isempty(XLSX.getSparklines(s))
        XLSX.addSparklines(s, "F1", "A1:E1")
        XLSX.addSparklines(s, "G1:G3", "A1:E3"; type=:column)                     # one per row
        XLSX.addSparklines(s, "H1", "A1:E1"; type=:winloss)
        @test ext_xml(s) == ext_xml(fixture("kinds")[1])                         # the same XML, new groups first

        xf, s = data_sheet()
        XLSX.addSparklines(s, "G1:G3", "A1:E3"; type=:column)
        XLSX.addSparklines(s, "H1", "A1:E1"; type=:winloss)
        XLSX.addSparklines(s, "F1", "A1:E1"; markers=true, high=true, low=true, first=true, last=true, negative=true,
                           seriesColor="FF0000", highColor="00FF00", lineWeight=1.5, emptyCells=:zero, showHidden=true,
                           axis=true, min=-5)
        @test groups_xml(s) == groups_xml(fixture("options")[1])
        SAVE_FILES && save_outfile(xf)
    end

    @testset "reading Excel's" begin
        ex = fixture("options"; mode="r")[1]
        gs = XLSX.getSparklines(ex)
        @test [g.type for g in gs] == [:winloss, :column, :line]
        w, c, l = gs
        @test w.sparklines == ["H1" => "Sheet1!A1:E1"] && w.negative && w.emptyCells == :gap
        @test c.sparklines == ["G1" => "Sheet1!A1:E1", "G2" => "Sheet1!A2:E2", "G3" => "Sheet1!A3:E3"]
        @test (l.markers, l.high, l.low, l.first, l.last, l.negative, l.axis, l.showHidden, l.rightToLeft) ==
              (true, true, true, true, true, true, true, true, false)
        @test (l.emptyCells, l.lineWeight, l.min, l.max, l.dateRange) == (:zero, 1.5, -5.0, :individual, nothing)
        @test l.colors.seriesColor == "FFFF0000" && l.colors.highColor == "FF00FF00" && l.colors.axisColor == "FF000000"
        @test c.lineWeight == 0.75 && c.min == :individual && !c.markers
    end

    @testset "options" begin
        xf, s = data_sheet()
        s["A5:E5"] = [Date(2026, 1, d) for d in (1, 2, 4, 8, 16)]
        XLSX.addSparklines(s, "F1", "A1:E1"; emptyCells=:span, min=:group, max=3, rightToLeft=true, dateRange="A5:E5",
                           axisColor=:accent2, markersColor=XLSX.SchemeColor(:accent1; lumMod=75))
        g = only(XLSX.getSparklines(s))
        @test (g.emptyCells, g.min, g.max, g.rightToLeft, g.dateRange) == (:span, :group, 3.0, true, "Sheet1!A5:E5")
        @test g.colors.axisColor isa String && g.colors.markersColor isa String
        el = only(XLSX.xml_elements(only(XLSX.xml_elements(XLSX._worksheet_ext(XLSX._sheet_root(s), XLSX.SPARKLINE_EXT_URI)))))
        @test [first(p) for p in XLSX._atts_vector(el)] ==                       # the schema's order
              ["manualMax", "dateAxis", "displayEmptyCellsAs", "minAxisType", "maxAxisType", "rightToLeft"]
        @test [XLSX.localname(k) for k in XLSX.xml_elements(el)][end-1:end] == ["f", "sparklines"]   # the date range before the sparklines

        XLSX.addSparklines(s, "B7:F7", "A1:E3")                                  # a row of cells: one per column
        @test only(g for g in XLSX.getSparklines(s) if length(g.sparklines) == 5).sparklines[2] == ("C7" => "Sheet1!B1:B3")
        XLSX.addSparklines(s, "F1", "A2:E2"; type=:column)                       # replaces the sparkline in F1
        @test count(g -> any(p -> first(p) == "F1", g.sparklines), XLSX.getSparklines(s)) == 1
        @test (only(g for g in XLSX.getSparklines(s) if any(p -> first(p) == "F1", g.sparklines))).type == :column
        @test length(XLSX.getSparklines(s)) == 2                                 # the old F1 group, left empty, is gone

        t = XLSX.addsheet!(xf, "My Data")
        XLSX.addSparklines(t, "A1", "Sheet1!A1:E1")
        XLSX.addSparklines(s, XLSX.CellRef("H9"), "'My Data'!A1:A1")
        @test only(XLSX.getSparklines(s)[1].sparklines) == ("H9" => "'My Data'!A1:A1")
        r = roundtrip(xf)
        @test sparks(r[1]) == sparks(s) && sparks(r["My Data"]) == sparks(t)    # read-only
    end

    @testset "in Excel's order among the other extensions" begin
        xf, s = data_sheet()
        XLSX.addSparklines(s, "F1", "A1:E1")
        XLSX.setConditionalFormat(s, "A1:A3", :dataBar)                         # an x14 data bar
        @test exts(s) == [XLSX.X14_CF_EXT_URI, XLSX.SPARKLINE_EXT_URI]
    end

    @testset "deleting" begin
        xf, s = data_sheet()
        XLSX.addSparklines(s, "G1:G3", "A1:E3")
        XLSX.addSparklines(s, "F1", "A1:E1"; type=:column)
        @test XLSX.deleteSparklines(s, "G2")
        @test only(g for g in XLSX.getSparklines(s) if g.type == :line).sparklines == ["G1" => "Sheet1!A1:E1", "G3" => "Sheet1!A3:E3"]
        @test !XLSX.deleteSparklines(s, "G2")
        @test XLSX.deleteSparklines(s, "G3"; group=true)                          # the whole group
        @test [g.type for g in XLSX.getSparklines(s)] == [:column]
        @test XLSX.deleteSparklines(s, "A:F")
        @test isempty(XLSX.getSparklines(s)) && isempty(exts(s))                  # nothing empty left behind
        @test !any(c -> XLSX.localname(c) == "extLst", XLSX.xml_elements(XLSX._sheet_root(s)))

        XLSX.addSparklines(s, "F1", "A1:E1")
        XLSX.setConditionalFormat(s, "A1:A3", :dataBar)
        XLSX.deleteSparklines(s, "F1")
        @test exts(s) == [XLSX.X14_CF_EXT_URI]                                   # other extensions kept
    end

    @testset "inserting and deleting rows and columns, as Excel does" begin
        xf = fixture("kinds")
        s = xf[1]
        XLSX.insertRows(s, 2)                                                     # new rows get sparklines from above
        @test sparks(s) == sparks(fixture("inserted")[1])
        XLSX.deleteRows(s, 1)
        @test sparks(s) == sparks(fixture("deleted")[1])

        xf = fixture("cols_before")
        s = xf[1]
        XLSX.insertColumns(s, 2)                                                  # and new columns from the left
        @test sparks(s) == sparks(fixture("cols_inserted")[1])
        XLSX.insertRows(s, 1)                                                     # nothing above row 1
        @test sparks(s) == sparks(fixture("row_above_first")[1])

        xf, s = data_sheet()
        XLSX.addSparklines(s, "F1:F2", "A1:E2")
        XLSX.addSparklines(s, "G3", "A3:E3")
        XLSX.deleteRows(s, 1:2)                                                   # a group's last sparklines gone
        @test sparks(s) == Set([(:line, Set(["G1" => "Sheet1!A1:E1"]))])
        XLSX.deleteRows(s, 1)
        @test isempty(XLSX.getSparklines(s)) && !(XLSX.SPARKLINE_EXT_URI in exts(s))
        XLSX.addSparklines(s, "F1", "A1:E1")
        XLSX.insertRows(s, 3; count=2)                                            # not below a sparkline: none added
        @test length(only(XLSX.getSparklines(s)).sparklines) == 1

        # data deleted: the sparklines stay, without data, as in Excel
        xf, s = data_sheet()
        XLSX.addSparklines(s, "G1:G2", "A1:E2")
        XLSX.addSparklines(s, "H1", "B1:C1")
        XLSX.deleteColumns(s, "A:E")
        @test sparks(s) == sparks(fixture("data_deleted")[1])
        @test !occursin("#REF!", XML.write(XLSX._sheet_root(s)))
    end

    @testset "sheets copied, renamed and deleted" begin
        xf = fixture("kinds")
        c = XLSX.copysheet!(xf[1], "Sheet1 (2)")                                  # the copy charts its own cells, as in Excel
        @test sparks(c) == sparks(fixture("copied")["Sheet1 (2)"])
        @test sparks(xf[1]) == sparks(fixture("kinds")[1])

        XLSX.renamesheet!(xf[1], "Renamed")
        @test all(startswith(last(p), "Renamed!") for g in XLSX.getSparklines(xf["Renamed"]) for p in g.sparklines)
        t = XLSX.addsheet!(xf, "Chart here")
        XLSX.addSparklines(t, "A1", "Renamed!A1:E1")
        XLSX.deletesheet!(xf, "Renamed")
        @test only(XLSX.getSparklines(t)).sparklines == ["A1" => ""]                 # kept, without data, as Excel keeps it
        @test !occursin("#REF!", XML.write(XLSX._sheet_root(t)))
    end

    @testset "errors" begin
        xf, s = data_sheet()
        @test_throws XLSX.XLSXError XLSX.addSparklines(s, "F1", "A1:E3")          # one cell, several rows
        @test_throws XLSX.XLSXError XLSX.addSparklines(s, "F1:F2", "A1:E3")      # two cells, three rows
        @test_throws XLSX.XLSXError XLSX.addSparklines(s, "F1:G2", "A1:E2")      # not a row or column of cells
        @test_throws XLSX.XLSXError XLSX.addSparklines(s, "F1", "Missing!A1:E1")
        @test_throws XLSX.XLSXError XLSX.addSparklines(s, "F1", "nonsense")
        @test_throws XLSX.XLSXError XLSX.addSparklines(s, "F1", "A1:E1"; type=:pie)
        @test_throws XLSX.XLSXError XLSX.addSparklines(s, "F1", "A1:E1"; emptyCells=:blank)
        @test_throws XLSX.XLSXError XLSX.addSparklines(s, "F1", "A1:E1"; min=:lowest)
        @test_throws XLSX.XLSXError XLSX.addSparklines(s, "F1", "A1:E1"; markers=1)
        @test_throws XLSX.XLSXError XLSX.addSparklines(s, "F1", "A1:E1"; lineWeight=0)
        @test_throws XLSX.XLSXError XLSX.addSparklines(s, "F1", "A1:E1"; colour="red")
        @test_throws XLSX.XLSXError XLSX.addSparklines(s, "F1", "A1:E1"; seriesColor="notacolour")
        @test isempty(XLSX.getSparklines(s))                                     # nothing was added
        r = roundtrip(xf)
        @test_throws XLSX.XLSXError XLSX.addSparklines(r[1], "F1", "A1:E1")
        @test_throws XLSX.XLSXError XLSX.deleteSparklines(r[1], "F1")
        XLSX.Charts.addChart(xf, :line; sheetname="Chart")
        @test_throws XLSX.XLSXError XLSX.addSparklines(xf["Chart"], "F1", "Sheet1!A1:E1")
    end
end
