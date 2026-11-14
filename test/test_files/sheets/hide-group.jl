@testset "hide and group" begin
    function fresh()
        xf = XLSX.newxlsx("Data")
        s = xf["Data"]
        s[1:10, 1:4] = reshape(1:40, 10, 4)
        return xf, s
    end
    sheet_root(xf, n=1) = XLSX.xml_root_element(xf.data["xl/worksheets/sheet$n.xml"])
    col_nodes(xf, n=1) = (b = XLSX.first_element_with_tag(sheet_root(xf, n), "cols");
                          isnothing(b) ? XML.Node[] : collect(XLSX.xml_elements(b)))
    col_spans(xf, n=1) = [(parse(Int, c["min"]), parse(Int, c["max"])) for c in col_nodes(xf, n)]
    format_pr(xf, n=1) = XLSX.first_element_with_tag(sheet_root(xf, n), "sheetFormatPr")
    function roundtrip(xf)
        io = IOBuffer()
        XLSX.writexlsx(io, xf)
        return XLSX.readxlsx(IOBuffer(take!(io)))
    end
    function roundtrip_sheet_xml(xf, n=1)
        io = IOBuffer()
        XLSX.writexlsx(io, xf)
        z = XLSX.ZipArchives.ZipReader(take!(io))
        return String(XLSX.ZipArchives.zip_readentry(z, "xl/worksheets/sheet$n.xml"))
    end

    @testset "hide columns" begin
        xf, s = fresh()
        @test isnothing(XLSX.setColumnHidden(s, "C:D"))
        @test XLSX.isColumnHidden(s, 3) && XLSX.isColumnHidden(s, "D") && XLSX.isColumnHidden(s, "D7")
        @test !XLSX.isColumnHidden(s, 2) && !XLSX.isColumnHidden(s, "E")
        @test col_spans(xf) == [(3, 4)]
        @test only(col_nodes(xf))["width"] == "11.140625"     # the template's default width, not 0
        @test !haskey(only(col_nodes(xf)), "customWidth")

        XLSX.setColumnHidden(s, [2, 6, 7])                    # a vector: separate runs
        @test [XLSX.isColumnHidden(s, c) for c in 1:8] == [false, true, true, true, false, true, true, false]
        @test col_spans(xf) == [(2, 4), (6, 7)]               # equal neighbours are joined

        XLSX.setColumnHidden(s, 3; hidden=false)
        @test [XLSX.isColumnHidden(s, c) for c in 1:5] == [false, true, false, true, false]
        @test col_spans(xf) == [(2, 2), (3, 3), (4, 4), (6, 7)]

        XLSX.setColumnHidden(s, "Z")                          # well beyond the data
        @test XLSX.isColumnHidden(s, 26)

        XLSX.setColumnHidden(xf, "Data!H:I")
        @test XLSX.isColumnHidden(xf, "Data!I")

        r = roundtrip(xf)[1]
        @test [XLSX.isColumnHidden(r, c) for c in 1:9] == [false, true, false, true, false, true, true, true, true]
        @test XLSX.isColumnHidden(r, "Z")
        SAVE_FILES && save_outfile(xf)
    end

    @testset "hiding keeps the width, and widths keep hiding" begin
        xf, s = fresh()
        XLSX.setColumnWidth(s, "B:D"; width=20)
        XLSX.setColumnHidden(s, "C:E")
        @test XLSX.getColumnWidth(s, "C1") ≈ 20.7109375
        @test XLSX.isColumnHidden(s, "C")
        @test col_spans(xf) == [(2, 2), (3, 4), (5, 5)]
        XLSX.setColumnWidth(s, "C"; width=5)
        @test XLSX.isColumnHidden(s, "C")
        @test XLSX.getColumnWidth(s, "C1") ≈ 5.7109375
        @test XLSX.getColumnWidth(s, "D1") ≈ 20.7109375
    end

    @testset "hide rows" begin
        xf, s = fresh()
        @test isnothing(XLSX.setRowHidden(s, 2:3))
        @test XLSX.isRowHidden(s, 2) && XLSX.isRowHidden(s, "3") && XLSX.isRowHidden(s, "B3")
        @test !XLSX.isRowHidden(s, 1) && !XLSX.isRowHidden(s, 4)
        XLSX.setRowHidden(s, "6:7,9")
        @test [XLSX.isRowHidden(s, r) for r in 1:10] == [false, true, true, false, false, true, true, false, true, false]

        XLSX.setRowHidden(s, 20)                              # an empty row
        @test XLSX.isRowHidden(s, 20)
        @test s[20, 1] isa Missing || isnothing(s[20, 1]) || XLSX.getcell(s, 20, 1) isa XLSX.EmptyCell
        xml = roundtrip_sheet_xml(xf)
        @test occursin(r"<row r=\"20\"[^>]*hidden=\"1\"", xml)

        XLSX.setRowHidden(xf, "Data!7"; hidden=false)
        @test !XLSX.isRowHidden(s, 7)
        XLSX.setRowHidden(s, 20; hidden=false)                # the empty row goes again
        @test !haskey(s.cache.cells, 20)
        @test !occursin("r=\"20\"", roundtrip_sheet_xml(xf))

        r = roundtrip(xf)[1]
        @test [XLSX.isRowHidden(r, i) for i in 1:10] == [false, true, true, false, false, true, false, false, true, false]
        @test r[2, 1] == 2                                    # data untouched
        SAVE_FILES && save_outfile(xf)
    end

    @testset "row heights on empty rows" begin
        xf, s = fresh()
        @test XLSX.setRowHeight(s, "A15"; height=30) == 0
        XLSX.setRowHidden(s, 15)
        r = roundtrip(xf)[1]
        @test XLSX.isRowHidden(r, 15)                         # fills the cache
        @test r.cache.row_ht[15] ≈ 30.2109375
    end

    @testset "group rows" begin
        xf, s = fresh()
        @test isnothing(XLSX.groupRows(s, 2:9))
        XLSX.groupRows(s, "4:5")
        @test [XLSX.getRowOutlineLevel(s, r) for r in 1:10] == [0, 1, 1, 2, 2, 1, 1, 1, 1, 0]
        @test format_pr(xf)["outlineLevelRow"] == "2"
        @test !haskey(format_pr(xf), "outlineLevelCol")
        @test format_pr(xf)["defaultRowHeight"] == "16"       # other settings kept
        @test !any(XLSX.isRowHidden(s, r) for r in 1:10)

        XLSX.ungroupRows(s, 4:5)
        @test [XLSX.getRowOutlineLevel(s, r) for r in 1:10] == [0, 1, 1, 1, 1, 1, 1, 1, 1, 0]
        @test format_pr(xf)["outlineLevelRow"] == "1"
        XLSX.ungroupRows(s, 1:10)                             # rows 1 and 10 weren't grouped
        @test all(XLSX.getRowOutlineLevel(s, r) == 0 for r in 1:10)
        @test !haskey(format_pr(xf), "outlineLevelRow")
        @test isnothing(s.unhandled_attributes) || isempty(s.unhandled_attributes)
    end

    @testset "collapsed groups" begin
        xf, s = fresh()
        XLSX.groupRows(s, 3:5; collapsed=true)
        @test all(XLSX.isRowHidden(s, r) for r in 3:5)
        @test XLSX.getRowOutlineLevel(s, 6) == 0
        @test s.unhandled_attributes[6]["collapsed"] == "1"   # the summary row below
        @test !XLSX.isRowHidden(s, 6)

        XLSX.ungroupRows(s, 3:5)
        @test !haskey(something(get(s.unhandled_attributes, 6, nothing), Dict()), "collapsed")
        @test all(XLSX.isRowHidden(s, r) for r in 3:5)        # still hidden, as in Excel
        XLSX.setRowHidden(s, 3:5; hidden=false)
        @test isempty(s.unhandled_attributes)

        XLSX.groupColumns(s, "B:C"; collapsed=true)
        @test XLSX.isColumnHidden(s, "B") && XLSX.isColumnHidden(s, "C")
        @test only(c for c in col_nodes(xf) if c["min"] == "4")["collapsed"] == "1"
        @test format_pr(xf)["outlineLevelCol"] == "1"

        r = roundtrip(xf)[1]
        @test XLSX.getColumnOutlineLevel(r, "B") == 1 && XLSX.isColumnHidden(r, "C")
        SAVE_FILES && save_outfile(xf)
    end

    @testset "summary rows above" begin
        xf, s = fresh()
        root = sheet_root(xf)
        pr = XML.Element("sheetPr", XML.Element("outlinePr"; summaryBelow="0", summaryRight="0"))
        pushfirst!(root.children, pr)
        XLSX.update_worksheets_xml!(xf)
        XLSX.groupRows(s, 4:6; collapsed=true)
        @test s.unhandled_attributes[3]["collapsed"] == "1"
        XLSX.groupColumns(s, "C:D"; collapsed=true)
        @test only(c for c in col_nodes(xf) if c["min"] == "2")["collapsed"] == "1"
        XLSX.groupRows(s, 1:2; collapsed=true)                # no row above row 1: no summary row
        @test XLSX.isRowHidden(s, 1)
    end

    @testset "group columns" begin
        xf, s = fresh()
        XLSX.groupColumns(s, "B:F")
        XLSX.groupColumns(s, [3, 4])
        @test [XLSX.getColumnOutlineLevel(s, c) for c in 1:7] == [0, 1, 2, 2, 1, 1, 0]
        @test col_spans(xf) == [(2, 2), (3, 4), (5, 6)]
        @test format_pr(xf)["outlineLevelCol"] == "2"
        XLSX.groupColumns(xf, "Data!D")
        @test XLSX.getColumnOutlineLevel(xf, "Data!D") == 3
        XLSX.ungroupColumns(s, "A:G")
        @test [XLSX.getColumnOutlineLevel(s, c) for c in 1:7] == [0, 0, 1, 2, 0, 0, 0]
        @test format_pr(xf)["outlineLevelCol"] == "2"
        r = roundtrip(xf)[1]
        @test [XLSX.getColumnOutlineLevel(r, c) for c in 1:7] == [0, 0, 1, 2, 0, 0, 0]
    end

    @testset "outline level limit" begin
        xf, s = fresh()
        for _ in 1:7
            XLSX.groupRows(s, 2:3)
            XLSX.groupColumns(s, 2)
        end
        @test XLSX.getRowOutlineLevel(s, 2) == 7
        @test XLSX.getColumnOutlineLevel(s, 2) == 7
        @test_throws XLSX.XLSXError XLSX.groupRows(s, 1:3)    # row 1 could go deeper, but not rows 2:3
        @test XLSX.getRowOutlineLevel(s, 1) == 0              # nothing changed
        @test_throws XLSX.XLSXError XLSX.groupColumns(s, "B")
        @test format_pr(xf)["outlineLevelRow"] == "7"
    end

    @testset "references" begin
        xf, s = fresh()
        XLSX.addDefinedName(s, "Middle", "B4:C6")
        XLSX.setRowHidden(s, "Middle")
        @test [XLSX.isRowHidden(s, r) for r in 3:7] == [false, true, true, true, false]
        XLSX.groupColumns(s, "Middle")
        @test [XLSX.getColumnOutlineLevel(s, c) for c in 1:4] == [0, 1, 1, 0]
        XLSX.setColumnHidden(s, XLSX.CellRange("D1:D2"))
        @test XLSX.isColumnHidden(s, 4)
        XLSX.setRowHidden(s, XLSX.CellRef("A9"))
        @test XLSX.isRowHidden(s, 9)

        @test_throws XLSX.XLSXError XLSX.setColumnHidden(s, "3:4")       # rows, not columns
        @test_throws XLSX.XLSXError XLSX.setRowHidden(s, "B:C")          # columns, not rows
        @test_throws XLSX.XLSXError XLSX.setColumnHidden(s, 0)
        @test_throws XLSX.XLSXError XLSX.setColumnHidden(s, XLSX.EXCEL_MAX_COLS + 1)
        @test_throws XLSX.XLSXError XLSX.setRowHidden(s, XLSX.EXCEL_MAX_ROWS + 1)
        @test_throws XLSX.XLSXError XLSX.setRowHidden(s, Int[])
        @test_throws XLSX.XLSXError XLSX.setRowHidden(s, "nonsense!")
        @test_throws XLSX.XLSXError XLSX.isRowHidden(s, 2:3)             # one row only
        @test_throws XLSX.XLSXError XLSX.isColumnHidden(s, "B:C")
        @test_throws XLSX.XLSXError XLSX.setRowHidden(xf, "3:4")         # no sheet
        @test_throws XLSX.XLSXError XLSX.setRowHidden(xf, "Nope!3:4")

        XLSX.addsheet!(xf, "Other")
        XLSX.addDefinedName(xf["Other"], "There", "A1:A3")
        @test_throws XLSX.XLSXError XLSX.setRowHidden(s, "There")        # names a cell on another sheet
    end

    @testset "file modes" begin
        xf, s = fresh()
        XLSX.setRowHidden(s, 2)
        XLSX.groupColumns(s, "C")
        io = IOBuffer()
        XLSX.writexlsx(io, xf)
        bytes = take!(io)

        r = XLSX.readxlsx(IOBuffer(bytes))[1]
        @test_throws XLSX.XLSXError XLSX.setRowHidden(r, 3)
        @test_throws XLSX.XLSXError XLSX.groupColumns(r, "D")
        @test XLSX.isRowHidden(r, 2) && XLSX.getColumnOutlineLevel(r, "C") == 1

        XLSX.openxlsx(IOBuffer(bytes); enable_cache=false) do f
            r = f[1]
            @test XLSX.getColumnOutlineLevel(r, "C") == 1       # columns need no cache
            @test_throws XLSX.XLSXError XLSX.isRowHidden(r, 2)  # rows do
        end

        f = XLSX.openxlsx(IOBuffer(bytes); mode="rw")
        XLSX.setRowHidden(f[1], 2; hidden=false)
        @test !XLSX.isRowHidden(f[1], 2)
        @test XLSX.getColumnOutlineLevel(f[1], "C") == 1
        SAVE_FILES && save_outfile(f)
    end

    @testset "sheet operations" begin
        xf, s = fresh()
        XLSX.setRowHidden(s, 3)
        XLSX.groupRows(s, 5:6; collapsed=true)
        XLSX.setColumnHidden(s, "B")
        c = XLSX.copysheet!(s, "Copy")
        @test XLSX.isRowHidden(c, 3) && XLSX.isRowHidden(c, 5)
        @test XLSX.getRowOutlineLevel(c, 6) == 1
        @test XLSX.isColumnHidden(c, "B")
        XLSX.setRowHidden(c, 3; hidden=false)                  # the copy is separate
        @test XLSX.isRowHidden(s, 3) && !XLSX.isRowHidden(c, 3)
        r = roundtrip(xf)
        @test XLSX.isRowHidden(r["Data"], 3) && !XLSX.isRowHidden(r["Copy"], 3)
        @test XLSX.isRowHidden(r["Copy"], 5)
    end

    @testset "chartsheets" begin
        xf = XLSX.newxlsx("data")
        xf["data"]["A1:A2"] = [1, 2]
        c = XLSX.Charts.addChart(xf, :line; sheetname="Trend")
        XLSX.Charts.addSeries(c, "data!A1:A2")
        chs = xf["Trend"]
        @test_throws XLSX.XLSXError XLSX.setRowHidden(chs, 1)
        @test_throws XLSX.XLSXError XLSX.groupColumns(chs, "A")
        @test_throws XLSX.XLSXError XLSX.isColumnHidden(chs, "A")
    end
end
