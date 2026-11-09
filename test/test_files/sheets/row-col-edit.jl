@testset "insert and delete" begin
    # A 10x4 block: A1:D10 holds 1..40, by column.
    function fresh()
        xf = XLSX.newxlsx("Data")
        s = xf["Data"]
        s[1:10, 1:4] = reshape(1:40, 10, 4)
        XLSX.addsheet!(xf, "Other")
        return xf, s
    end
    sheet_root(xf, n=1) = XLSX.xml_root_element(xf.data["xl/worksheets/sheet$n.xml"])
    element(xf, tag, n=1) = XLSX.first_element_with_tag(sheet_root(xf, n), tag)
    function roundtrip(xf)
        io = IOBuffer()
        XLSX.writexlsx(io, xf)
        return XLSX.readxlsx(IOBuffer(take!(io)))
    end
    function roundtrip_bytes(xf)
        io = IOBuffer()
        XLSX.writexlsx(io, xf)
        return take!(io)
    end
    entry(bytes, name) = XLSX.ZipArchives.zip_readentry(XLSX.ZipArchives.ZipReader(bytes), name, String)
    formula(s, ref) = XLSX.getFormula(s, ref)

    @testset "cells move" begin
        xf, s = fresh()
        XLSX.setRowHeight(s, 7; height=30)
        XLSX.setRowHidden(s, 8)
        @test isnothing(XLSX.insertRows(s, 5; count=2))
        @test s[4, 1] == 4 && ismissing(s[5, 1]) && ismissing(s[6, 1]) && s[7, 1] == 5 && s[12, 1] == 10
        @test s[7, 4] == 35
        @test XLSX.getcell(s, 7, 1).ref == XLSX.CellRef("A7")
        @test XLSX.get_dimension(s) == XLSX.CellRange("A1:D12")
        @test s.cache.row_ht[9] ≈ 30.2109375 && XLSX.isRowHidden(s, 10)     # row attributes moved too

        XLSX.deleteRows(s, 5:6)
        @test [s[r, 1] for r in 1:10] == collect(1:10)
        @test XLSX.get_dimension(s) == XLSX.CellRange("A1:D10")

        XLSX.deleteRows(s, "2:3,9")
        @test [s[r, 1] for r in 1:7] == [1, 4, 5, 6, 7, 8, 10]

        XLSX.insertColumns(s, "B")
        @test s[1, 1] == 1 && ismissing(s[1, 2]) && s[1, 3] == 11 && s[1, 5] == 31
        XLSX.deleteColumns(s, [2, 4])
        @test [s[1, c] for c in 1:3] == [1, 11, 31]
        @test XLSX.get_dimension(s) == XLSX.CellRange("A1:C7")

        XLSX.insertRows(xf, "Data!1")
        @test s[2, 1] == 1
        XLSX.deleteColumns(xf, "Data!A")
        @test s[2, 1] == 11

        r = roundtrip(xf)["Data"]
        @test r[2, 1] == 11 && r[8, 2] == 40
        @test XLSX.get_dimension(r) == XLSX.CellRange("A2:B8")
        SAVE_FILES && save_outfile(xf)
    end

    @testset "deleting everything" begin
        xf, s = fresh()
        XLSX.deleteRows(s, 1:10)
        @test isempty(s.cache)
        @test XLSX.get_dimension(s) == XLSX.CellRange("A1:A1")
        @test s.sst_count == 0
        r = roundtrip(xf)["Data"]
        @test ismissing(r[1, 1])
    end

    @testset "shared strings are counted" begin
        xf = XLSX.newxlsx("Data")
        s = xf["Data"]
        s["A1"] = ["a" "b"; "c" "d"]
        @test s.sst_count == 4
        XLSX.deleteRows(s, 1)
        @test s.sst_count == 2
        XLSX.deleteColumns(s, 1)
        @test s.sst_count == 1
        @test roundtrip(xf)["Data"][1, 1] == "d"
    end

    @testset "formulas" begin
        xf, s = fresh()
        XLSX.setFormula(s, "F1", "=SUM(A1:A10)")
        XLSX.setFormula(s, "F2", "=A5*2")
        XLSX.setFormula(s, "F3", "=\$A\$5+A\$4")
        o = xf["Other"]
        XLSX.setFormula(o, "A1", "=Data!A5+A5")
        XLSX.setFormula(o, "A2", "=SUM(Data!B:B)")

        XLSX.insertRows(s, 3)
        @test formula(s, "F1") == "=SUM(A1:A11)"
        @test formula(s, "F2") == "=A6*2"
        @test formula(s, "F4") == "=\$A\$6+A\$5"            # F3 moved down with its row
        @test formula(o, "A1") == "=Data!A6+A5"             # only references to Data move
        @test formula(o, "A2") == "=SUM(Data!B:B)"

        XLSX.deleteRows(s, 6)                                # the old row 5
        @test formula(s, "F2") == "=#REF!*2"
        @test formula(o, "A1") == "=Data!#REF!+A5"
        @test formula(s, "F1") == "=SUM(A1:A10)"

        XLSX.deleteColumns(s, "A")
        @test formula(s, "E1") == "=SUM(#REF!)"              # the formula moved left too
        @test XLSX.getcell(s, "F1") isa XLSX.EmptyCell

        r = roundtrip(xf)
        @test XLSX.getFormula(r["Data"], "E1") == "=SUM(#REF!)"
        @test XLSX.getFormula(r["Other"], "A1") == "=Data!#REF!+A5"
    end

    @testset "formulas on a deleted row go" begin
        xf, s = fresh()
        XLSX.setFormula(s, "F5", "=A1")
        XLSX.deleteRows(s, 5)
        @test !haskey(XLSX.get_workbook(xf).formulas, XLSX.SheetCellRef("Data!F5"))
        @test ismissing(s["F5"]) || !XLSX.getcell(s, "F5").formula
    end

    @testset "shared formulas" begin
        xf, s = fresh()
        XLSX.setFormula(s, "F1:F10", "=A1*2")               # a shared formula
        fs = XLSX.get_workbook(xf).formulas
        @test any(f -> f isa XLSX.ReferencedFormula, values(fs))
        XLSX.insertRows(s, 4)
        @test !any(f -> f isa Union{XLSX.ReferencedFormula,XLSX.FormulaReference}, values(fs))
        @test formula(s, "F1") == "=A1*2"
        @test formula(s, "F3") == "=A3*2"
        @test ismissing(s["F4"])
        @test formula(s, "F5") == "=A5*2"                   # the old F4: =A4*2, moved with A4
        @test formula(s, "F11") == "=A11*2"

        # a block of formulas on another sheet, referring to the edited one (one per
        # cell: see below for a shared one)
        xf, s = fresh()
        o = xf["Other"]
        XLSX.setFormula(o, "A1:A5", "=Data!A1+1")
        XLSX.deleteRows(s, 2)
        @test formula(o, "A1") == "=Data!A1+1"
        @test formula(o, "A2") == "=Data!#REF!+1"
        @test formula(o, "A3") == "=Data!A2+1"

        # one that doesn't refer to the edited sheet stays shared
        xf, s = fresh()
        o = xf["Other"]
        XLSX.setFormula(o, "A1:A5", "=B1+1")
        XLSX.insertRows(s, 2)
        @test any(f -> f isa XLSX.ReferencedFormula, values(XLSX.get_workbook(xf).formulas))
    end

    @testset "shared formulas the edit moves alike stay shared" begin
        shared(xf) = [(k, f) for (k, f) in XLSX.get_workbook(xf).formulas if f isa XLSX.ReferencedFormula]

        # a block below the insert: moves whole, master and all
        xf, s = fresh()
        XLSX.setFormula(s, "F5:F10", "=A5*2")
        XLSX.insertRows(s, 2; count=2)
        k, f = only(shared(xf))
        @test k == XLSX.SheetCellRef("Data!F7") && f.ref == "F7:F12" && f.formula == "=A7*2"
        @test formula(s, "F7") == "=A7*2" && formula(s, "F12") == "=A12*2"
        r = roundtrip(xf)["Data"]
        @test XLSX.getFormula(r, "F12") == "=A12*2"

        # a block above the insert, referring only above it: unchanged
        xf, s = fresh()
        XLSX.setFormula(s, "F1:F3", "=A1+\$B\$2")
        XLSX.insertRows(s, 8)
        k, f = only(shared(xf))
        @test k == XLSX.SheetCellRef("Data!F1") && f.ref == "F1:F3" && f.formula == "=A1+\$B\$2"

        # an absolute reference the insert moves: still the same for every cell
        xf, s = fresh()
        XLSX.setFormula(s, "F1:F3", "=A1+\$B\$9")
        XLSX.insertRows(s, 5)
        @test only(shared(xf))[2].formula == "=A1+\$B\$10"
        @test formula(s, "F3") == "=A3+\$B\$10"

        # references across the edited row: the block moves whole, but its cells don't
        # all come out as the moved master gives them, so it is expanded
        xf, s = fresh()
        XLSX.setFormula(s, "F5:F10", "=A4")
        XLSX.insertRows(s, 5)
        @test isempty(shared(xf))
        @test formula(s, "F6") == "=A4" && formula(s, "F7") == "=A6" && formula(s, "F11") == "=A10"

        # a delete inside the block changes its shape: expanded
        xf, s = fresh()
        XLSX.setFormula(s, "F1:F10", "=A1*2")
        XLSX.deleteRows(s, 4)
        @test isempty(shared(xf))
        @test formula(s, "F4") == "=A4*2" && formula(s, "F9") == "=A9*2"

        # on another sheet, referring to the edited one alike. `setFormula` writes a
        # formula referring to another sheet into each cell, so share it by hand, as
        # Excel writes it.
        xf, s = fresh()
        o = xf["Other"]
        o["A1:A5"] = 0
        fs = XLSX.get_workbook(xf).formulas
        id = XLSX.new_ReferencedFormula_Id(o)
        fs[XLSX.SheetCellRef("Other!A1")] = XLSX.ReferencedFormula("=Data!A10+1", id, "A1:A5", nothing)
        for r in 2:5
            fs[XLSX.SheetCellRef("Other!A$r")] = XLSX.FormulaReference(id, nothing)
        end
        foreach(r -> XLSX.getcell(o, r, 1).formula = true, 1:5)
        @test formula(o, "A3") == "=Data!A12+1"
        XLSX.insertRows(s, 2)
        k, f = only(shared(xf))
        @test k == XLSX.SheetCellRef("Other!A1") && f.ref == "A1:A5" && f.formula == "=Data!A11+1"
        @test formula(o, "A3") == "=Data!A13+1"
    end

    @testset "merged cells" begin
        xf, s = fresh()
        XLSX.mergeCells(s, "B2:C4")
        XLSX.mergeCells(s, "A7:B7")
        XLSX.mergeCells(s, "A9:A10")
        XLSX.insertRows(s, 3)
        @test sort(string.(XLSX.getMergedCells(s))) == ["A10:A11", "A8:B8", "B2:C5"]
        XLSX.deleteRows(s, 8)                               # all of A8:B8
        XLSX.deleteRows(s, 10)                               # half of A9:A10 (was A10:A11), leaving one cell
        @test string.(XLSX.getMergedCells(s)) == ["B2:C5"]
        XLSX.deleteColumns(s, "B:C")
        @test isnothing(XLSX.getMergedCells(s)) || isempty(XLSX.getMergedCells(s))
        @test isnothing(element(xf, "mergeCells"))
    end

    @testset "conditional formats and data validations" begin
        xf, s = fresh()
        XLSX.setConditionalFormat(s, "B2:B9", :cellIs; operator="greaterThan", value="\$A\$5")
        XLSX.addDataValidation(s, "C3:C8"; type=:whole, formula1="\$A\$5", formula2=100)
        XLSX.addDataValidation(s, "D9"; type=:list, values=["a", "b"])
        o = xf["Other"]
        XLSX.addDataValidation(o, "A1"; type=:list, formula1="Data!\$A\$1:\$A\$10")

        XLSX.insertRows(s, 4; count=2)
        cf = only(XLSX.getConditionalFormats(s))
        @test string(first(cf)) == "B2:B11"
        cfr = only(c for c in XLSX.xml_elements(element(xf, "conditionalFormatting")) if XLSX.localname(c) == "cfRule")
        @test XML.simple_value(only(XLSX.xml_elements(cfr))) == "\$A\$7"
        dvs = XLSX.getDataValidations(s)
        @test Set(d.range for d in dvs) == Set([XLSX.CellRange("C3:C10"), XLSX.CellRange("D11:D11")])
        @test only(d for d in dvs if d.range == XLSX.CellRange("C3:C10")).formula1 == "\$A\$7"
        @test only(XLSX.getDataValidations(o)).formula1 == "Data!\$A\$1:\$A\$12"

        XLSX.deleteRows(s, 11)
        @test [d.range for d in XLSX.getDataValidations(s)] == [XLSX.CellRange("C3:C10")]
        XLSX.deleteRows(s, 2:12)
        @test isempty(XLSX.getDataValidations(s))
        @test isempty(XLSX.getConditionalFormats(s))
        @test isnothing(element(xf, "dataValidations"))
        @test only(XLSX.getDataValidations(o)).formula1 == "Data!\$A\$1:\$A\$1"
    end

    @testset "x14 data validations" begin
        xf, s = fresh()
        root = sheet_root(xf)
        ext = XML.parse("""<extLst><ext uri="{CCE6A557-97BC-4b89-ADB6-D9C93CAAB3DF}" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"><x14:dataValidations count="2" xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main"><x14:dataValidation type="list"><x14:formula1><xm:f>Other!\$A\$1:\$A\$3</xm:f></x14:formula1><xm:sqref>B2:B5</xm:sqref></x14:dataValidation><x14:dataValidation type="list"><x14:formula1><xm:f>Data!\$D\$1:\$D\$9</xm:f></x14:formula1><xm:sqref>C7</xm:sqref></x14:dataValidation></x14:dataValidations></ext></extLst>""", XML.Node)
        push!(root.children, XLSX.xml_root_element(ext))
        XLSX.update_worksheets_xml!(xf)
        XLSX.deleteRows(s, 7)
        dvs = XLSX.getDataValidations(s)
        @test [string(d.range) for d in dvs] == ["B2:B5"]
        XLSX.insertRows(s, 1)
        dvs = XLSX.getDataValidations(s)
        @test [string(d.range) for d in dvs] == ["B3:B6"]
        x14 = XLSX.first_element_with_tag(XLSX.first_element_with_tag(element(xf, "extLst"), "ext"), "dataValidations")
        @test x14["count"] == "1"
        XLSX.deleteRows(s, 3:6)
        @test isempty(XLSX.getDataValidations(s))
        @test isnothing(XLSX.first_element_with_tag(XLSX.first_element_with_tag(element(xf, "extLst"), "ext"), "dataValidations"))
    end

    @testset "hyperlinks" begin
        xf, s = fresh()
        XLSX.addHyperlink(s, "B3"; url="https://example.com")
        XLSX.addHyperlink(s, "B6"; location="Other!A1")
        XLSX.insertRows(s, 2)
        @test sort([string(h.ref) for h in XLSX.getHyperlinks(s)]) == ["B4", "B7"]
        XLSX.deleteRows(s, 4)
        h = only(XLSX.getHyperlinks(s))
        @test string(h.ref) == "B6" && h.location == "Other!A1"
        rels = entry(roundtrip_bytes(xf), "xl/worksheets/_rels/sheet1.xml.rels")
        @test !occursin("example.com", rels)                 # the deleted link's relationship went too
    end

    @testset "AutoFilter" begin
        xf, s = fresh()
        XLSX.setAutoFilter(s, "A1:D10")
        af = element(xf, "autoFilter")
        fc = XML.Element("filterColumn"; colId="2")
        push!(af, fc)
        XLSX.update_worksheets_xml!(xf)
        XLSX.insertColumns(s, "B")
        @test XLSX.getAutoFilter(s) == XLSX.CellRange("A1:E10")
        @test only(XLSX.xml_elements(element(xf, "autoFilter")))["colId"] == "3"   # was column C, now D
        @test XLSX.get_workbook(xf).worksheet_names[(s.sheetId, "_xlnm._FilterDatabase")].value == XLSX.SheetCellRange("Data!A1:E10")
        XLSX.deleteColumns(s, "D")
        @test isempty(XLSX.xml_elements(element(xf, "autoFilter")))               # its column went
        XLSX.deleteRows(s, 4)
        @test XLSX.getAutoFilter(s) == XLSX.CellRange("A1:D9")
        XLSX.deleteColumns(s, "A:D")
        @test isnothing(XLSX.getAutoFilter(s))
        @test !haskey(XLSX.get_workbook(xf).worksheet_names, (s.sheetId, "_xlnm._FilterDatabase"))
    end

    @testset "defined names" begin
        xf, s = fresh()
        XLSX.addDefinedName(xf, "Cell", "Data!B5")
        XLSX.addDefinedName(xf, "Block", "Data!B2:C9")
        XLSX.addDefinedName(xf, "Parts", "Data!A1,Data!B5:B6")
        XLSX.addDefinedName(xf, "Calc", "=SUM(Data!A1:A10)*2")
        XLSX.addDefinedName(xf, "Elsewhere", "Other!A5")
        XLSX.addDefinedName(s, "Local", "B5")
        XLSX.setPrintArea(s, "A1:D10")
        XLSX.setPrintTitles(s; rows=1:2)
        names = XLSX.get_workbook(xf).workbook_names
        local_names = XLSX.get_workbook(xf).worksheet_names

        XLSX.insertRows(s, 2)
        @test names["Cell"].value == XLSX.SheetCellRef("Data!B6")
        @test names["Block"].value == XLSX.SheetCellRange("Data!B3:C10")
        @test string(names["Parts"].value) == "Data!A1,Data!B6:B7"
        @test names["Calc"].value == "=SUM(Data!A1:A11)*2"
        @test names["Elsewhere"].value == XLSX.SheetCellRef("Other!A5")
        @test local_names[(s.sheetId, "Local")].value == XLSX.SheetCellRef("Data!B6")
        @test XLSX.getPrintArea(s) == XLSX.CellRange("A1:D11")
        @test XLSX.getPrintTitles(s) == (rows=1:3, cols=nothing)   # the insert was inside the titles

        XLSX.deleteRows(s, 6)
        @test names["Cell"].value == "=#REF!"
        @test string(names["Parts"].value) == "Data!A1,Data!B6"
        XLSX.deleteRows(s, 1)
        @test string(names["Parts"].value) == "Data!B5"
        w = entry(roundtrip_bytes(xf), "xl/workbook.xml")
        @test occursin("<definedName name=\"Cell\">#REF!</definedName>", w)
        @test occursin("<definedName name=\"Block\">Data!\$B\$2:\$C\$8</definedName>", w)
    end

    @testset "panes, selection and column widths" begin
        xf, s = fresh()
        XLSX.freezePanes(s, "C3")                          # two rows and two columns frozen
        XLSX.setColumnWidth(s, "C:D"; width=20)
        XLSX.insertRows(s, 1)
        pane = XLSX.first_element_with_tag(XLSX.first_element_with_tag(element(xf, "sheetViews"), "sheetView"), "pane")
        @test pane["ySplit"] == "3" && pane["topLeftCell"] == "C4"
        XLSX.insertColumns(s, "E")                          # right of the frozen columns: no change to them
        pane = XLSX.first_element_with_tag(XLSX.first_element_with_tag(element(xf, "sheetViews"), "sheetView"), "pane")
        @test pane["xSplit"] == "2"
        spans() = [(c["min"], c["max"]) for c in XLSX.xml_elements(element(xf, "cols"))]
        @test spans() == [("3", "5")]                       # the new column E takes D's width, as in Excel
        XLSX.insertColumns(s, "D")                          # inside the wide span: it grows
        @test spans() == [("3", "6")]
        XLSX.insertColumns(s, "A")                          # before it: it moves; a new A takes nothing
        @test spans() == [("4", "7")]
        XLSX.deleteColumns(s, "D:G")
        @test isnothing(element(xf, "cols"))
    end

    @testset "deleting the frozen rows and columns" begin
        pane(xf) = XLSX.first_element_with_tag(XLSX.first_element_with_tag(element(xf, "sheetViews"), "sheetView"), "pane")
        xf, s = fresh()
        XLSX.freezePanes(s; nrows=2, ncols=1)
        XLSX.deleteRows(s, 1:3)                              # all the frozen rows, and one more
        p = pane(xf)
        @test p["xSplit"] == "1" && !haskey(p, "ySplit")     # only the column stays frozen
        @test p["state"] == "frozen" && p["topLeftCell"] == "B1"
        XLSX.deleteColumns(s, "A")
        @test isnothing(pane(xf))                            # nothing frozen: no pane
        r = roundtrip(xf)["Data"]
        @test r[1, 1] == 14

        xf, s = fresh()
        XLSX.freezePanes(s; nrows=3)
        XLSX.deleteRows(s, 2)
        @test pane(xf)["ySplit"] == "2"
        XLSX.insertRows(s, 1; count=2)
        @test pane(xf)["ySplit"] == "4" && pane(xf)["topLeftCell"] == "A5"
    end

    @testset "page breaks" begin
        xf, s = fresh()
        root = sheet_root(xf)
        brks = XML.Element("rowBreaks"; count="2", manualBreakCount="2")
        push!(brks, XML.Element("brk"; id="3", max="16383", man="1"))
        push!(brks, XML.Element("brk"; id="8", max="16383", man="1"))
        insert!(root.children, XLSX.insert_index(root, "rowBreaks", XLSX.WORKSHEET_ORDER) + 1, brks)
        XLSX.update_worksheets_xml!(xf)
        XLSX.insertRows(s, 5; count=2)
        @test [b["id"] for b in XLSX.xml_elements(element(xf, "rowBreaks"))] == ["3", "10"]
        XLSX.deleteRows(s, 3)
        b = element(xf, "rowBreaks")
        @test [x["id"] for x in XLSX.xml_elements(b)] == ["9"] && b["count"] == "1"
    end

    @testset "charts on other sheets" begin
        xf, s = fresh()
        c = XLSX.Charts.addChart(xf, :line; sheetname="Trend")
        XLSX.Charts.addSeries(c, "Data!A1:A10")
        XLSX.insertRows(s, 3; count=2)
        chart = only(XLSX.parts_with_content_type(xf, XLSX.MIME_CHART))
        fs = String[]
        XLSX._rewrite_chart_formulas!(x -> (push!(fs, x); x), XLSX.xml_root_element(XLSX.get_xml_data(xf, chart)))
        @test any(==("Data!\$A\$1:\$A\$12"), fs)
        @test !any(occursin("A\$10", f) for f in fs)
    end

    # The drawing's anchors, as (from row, from col, to row, to col) 0-based, with offsets.
    function anchors(xf, n=1)
        path = XLSX._drawing_path_for_sheet(xf, "xl/worksheets/sheet$n.xml")
        out = []
        for a in XLSX.xml_elements(XLSX.xml_root_element(XLSX.get_xml_data(xf, path)))
            m(tag) = (el = XLSX.first_element_with_tag(a, tag); isnothing(el) ? nothing :
                      Tuple(parse(Int, XML.simple_value(XLSX.first_element_with_tag(el, t))) for t in ("row", "rowOff", "col", "colOff")))
            push!(out, (XLSX.localname(a), get(a, "editAs", ""), m("from"), m("to")))
        end
        return out
    end
    data_dir = joinpath(dirname(pathof(XLSX)), "..", "test", "data")

    @testset "images and charts" begin
        xf, s = fresh()
        XLSX.addImage(s, "B3:D6", joinpath(data_dir, "track_start.png"))   # two-cell anchor, rows 2:5, cols 1:3 (0-based)
        c = XLSX.Charts.addChart(s, :line; anchor="F2:M18")
        XLSX.Charts.addSeries(c, "Data!A1:A10")
        a0 = anchors(xf)
        @test length(a0) == 2
        img0, chart0 = a0

        XLSX.insertRows(s, 4; count=2)                     # inside both: they stretch
        img, chart = anchors(xf)
        @test img[3][1] == img0[3][1] && img[4][1] == img0[4][1] + 2
        @test chart[3][1] == chart0[3][1] && chart[4][1] == chart0[4][1] + 2

        XLSX.insertRows(s, 1)                              # above both: they move
        img, chart = anchors(xf)
        @test img[3][1] == img0[3][1] + 1 && img[4][1] == img0[4][1] + 3

        XLSX.insertColumns(s, "E")                         # right of the image, inside the chart
        img, chart = anchors(xf)
        @test img[3][3] == img0[3][3] && img[4][3] == img0[4][3]
        @test chart[3][3] == chart0[3][3] + 1 && chart[4][3] == chart0[4][3] + 1

        XLSX.deleteRows(s, 1:3)                            # the image's first row goes
        img, _ = anchors(xf)
        @test img[3][1] == 0 && img[3][2] == 0             # its top is now the top of what was below
        SAVE_FILES && save_outfile(xf)

        # every row of the image deleted: it goes, with its picture, as in Excel
        xf, s = fresh()
        XLSX.addImage(s, "B3:D6", joinpath(data_dir, "track_start.png"))
        XLSX.deleteRows(s, 3:7)
        @test isempty(XLSX.getImages(s))
        @test !any(k -> startswith(k, "xl/media/"), keys(xf.binary_data))
        r = roundtrip(xf)
        @test isempty(XLSX.getImages(r["Data"]))
        SAVE_FILES && save_outfile(xf)
    end

    @testset "how anchors move" begin
        xf, s = fresh()
        XLSX.addImage(s, "B3:D6", joinpath(data_dir, "track_start.png"))
        XLSX.addImage(s, "F3:H6", joinpath(data_dir, "track_start.png"))
        XLSX.addImage(s, "J3:L6", joinpath(data_dir, "track_start.png"))
        path = XLSX._drawing_path_for_sheet(xf, "xl/worksheets/sheet1.xml")
        as = XLSX.xml_elements(XLSX.xml_root_element(XLSX.get_xml_data(xf, path)))
        as[2]["editAs"] = "oneCell"                        # move but don't size
        as[3]["editAs"] = "absolute"                       # don't move or size
        before = anchors(xf)
        XLSX.insertRows(s, 4; count=3)                     # inside all three
        twocell, onecell, absolute = anchors(xf)
        @test twocell[4][1] - twocell[3][1] == before[1][4][1] - before[1][3][1] + 3   # stretched
        @test onecell[3] == before[2][3] && onecell[4] == before[2][4]                 # its top didn't move, so nor did it
        @test absolute == before[3]
        XLSX.insertRows(s, 1)                              # above all three
        twocell, onecell, absolute = anchors(xf)
        @test onecell[3][1] == before[2][3][1] + 1 && onecell[4][1] == before[2][4][1] + 1   # moved, same size
        @test absolute == before[3]
    end

    # A sheet as Excel shows it: each formatted cell's number format, font, fill and
    # border; each row's height, style, outline level and hidden flag; the column
    # definitions; and the merged, conditional-format and validation ranges.
    function formatting(xf, sheet)
        wb = XLSX.get_workbook(xf)
        ws = xf[sheet]
        isnothing(ws.cache) && foreach(_ -> nothing, XLSX.eachrow(ws))
        fonts, fills, borders = XLSX.get_fonts_nodes(wb), XLSX.get_fills_nodes(wb), XLSX.get_borders_nodes(wb)
        node(v, k) = XLSX.XML.write(v[parse(Int, k) + 1])
        style(s) = (x = XLSX.styles_cell_xf(wb, Int(s));
                    (get(x, "numFmtId", "0"), node(fonts, get(x, "fontId", "0")),
                     node(fills, get(x, "fillId", "0")), node(borders, get(x, "borderId", "0"))))
        plain = style(0)
        cells = Dict{String,Any}()
        rows = Dict{Int,Any}()
        for (r, row) in ws.cache.cells
            for cell in values(row)
                st = style(cell.style)
                st == plain || (cells[string(cell.ref)] = st)
            end
            uh = something(get(something(ws.unhandled_attributes, Dict()), r, nothing), Dict())
            a = (get(ws.cache.row_ht, r, nothing), haskey(uh, "s") ? style(parse(Int, uh["s"])) : nothing,
                 get(uh, "outlineLevel", nothing), get(uh, "hidden", nothing))
            all(isnothing, a) || (rows[r] = a)
        end
        root = XLSX._sheet_root(ws)
        cols = [(d["min"], get(d, "max", d["min"]), get(d, "width", nothing), get(d, "hidden", nothing),
                 get(d, "outlineLevel", nothing), haskey(d, "style") ? style(parse(Int, d["style"])) : nothing)
                for c in XLSX.xml_elements(root) if XLSX.localname(c) == "cols" for d in XLSX.xml_elements(c)]
        ranges = sort([(XLSX.localname(n), get(n, "sqref", get(n, "ref", ""))) for c in XLSX.xml_elements(root)
                       for n in (XLSX.localname(c) == "conditionalFormatting" ? [c] :
                                 XLSX.localname(c) in ("mergeCells", "dataValidations") ? XLSX.xml_elements(c) : [])])
        return (; cells, rows, cols, ranges)
    end

    @testset "inserted rows and columns are formatted as Excel formats them" begin
        # Excel's own before and after (made in Excel 365 through its automation): the
        # same inserts, bottom up and right to left.
        for (name, edits) in (
                ("rowcol_inherit", s -> begin
                    foreach(c -> XLSX.insertColumns(s, c), ("U", "Q", "K", "G"))
                    for (r, n) in ((56, 1), (51, 1), (41, 1), (31, 1), (21, 3), (15, 1), (10, 1), (6, 1), (1, 1))
                        XLSX.insertRows(s, r; count = n)
                    end
                end),
                ("rowcol_style", s -> (XLSX.insertColumns(s, "J"); XLSX.insertColumns(s, "E"); XLSX.insertRows(s, 6))))
            xf = XLSX.opentemplate(joinpath(data_directory, "$(name)_before.xlsx"))
            sheet = XLSX.sheetnames(xf)[1]
            edits(xf[sheet])
            ours, excel = formatting(xf, sheet), formatting(XLSX.opentemplate(joinpath(data_directory, "$(name)_after.xlsx")), sheet)
            @test ours.cells == excel.cells
            @test ours.rows == excel.rows
            @test ours.cols == excel.cols
            @test ours.ranges == excel.ranges
            SAVE_FILES && save_outfile(xf)
        end

        # A range ending just before the insert grows over it; whole rows and columns don't.
        E(axis, at, n=1) = XLSX.RowColEdit("Data", axis, at, n)
        @test XLSX._grow_sqref("B2:D9 M50 A:A F1:F3", E(:row, 10, 2)) == "B2:D11 M50 A:A F1:F3"
        @test XLSX._grow_sqref("M50", E(:row, 51)) == "M50:M51"
        @test XLSX._grow_sqref("B2:D9 2:4", E(:column, 5)) == "B2:E9 2:4"
        # At row 1 or column A there is nothing to take formatting from.
        xf, s = fresh()
        XLSX.setFill(s, "A1"; pattern = "solid", fgColor = "FFFF0000")
        XLSX.insertRows(s, 1)
        XLSX.insertColumns(s, "A")
        @test isnothing(XLSX.getFill(s, "A1")) && isnothing(XLSX.getFill(s, "B1"))
        @test !isnothing(XLSX.getFill(s, "B2"))
    end

    @testset "objects whose rows or columns are all deleted" begin
        # Excel's own before and after: a picture of each placement and two charts per
        # sheet; on "Rows" rows 3:18 deleted, on "Cols" columns B:O.
        function find_name(n)
            XLSX.localname(n) == "cNvPr" && return n["name"]
            for k in XLSX.xml_elements(n)
                r = find_name(k)
                isnothing(r) || return r
            end
            return nothing
        end
        function placed(xf, sheet)
            path = XLSX._drawing_path_for_sheet(xf, XLSX._sheet_path(xf[sheet]))
            marker(a, t) = (m = XLSX.first_element_with_tag(a, t);
                            Tuple(XLSX.XML.is_simple_value(XLSX.first_element_with_tag(m, k)) for k in ("col", "colOff", "row", "rowOff")))
            return sort([(find_name(a), marker(a, "from"), marker(a, "to"))
                         for a in XLSX.xml_elements(XLSX.xml_root_element(XLSX.get_xml_data(xf, path)))])
        end
        xf = XLSX.opentemplate(joinpath(data_directory, "rowcol_objects_before.xlsx"))
        XLSX.deleteRows(xf["Rows"], 3:18)
        XLSX.deleteColumns(xf["Cols"], "B:O")
        excel = XLSX.readxlsx(joinpath(data_directory, "rowcol_objects_after.xlsx"))
        for sheet in ("Rows", "Cols")
            @test placed(xf, sheet) == placed(excel, sheet)
            @test [p[1] for p in placed(xf, sheet)] == ["Chart2", "Pic2", "Pic3"]   # move-and-size went
        end
        @test length(XLSX.Charts.getCharts(xf)) == 2                                # with their chart parts
        @test count(k -> startswith(k, "xl/charts/chart"), keys(xf.data)) == 2
        r = roundtrip(xf)
        @test length(XLSX.Charts.getCharts(r)) == 2 && length(XLSX.getImages(r["Rows"])) == 2
        SAVE_FILES && save_outfile(xf)
    end

    @testset "comments" begin
        xf, s = fresh()
        XLSX.addComment(s, "B3", "on B3")
        XLSX.addComment(s, "C8", "on C8")
        XLSX.addComment(s, "D5", "on D5")
        XLSX.insertRows(s, 4; count=2)
        @test Set((string(c.ref), c.text) for c in XLSX.getComments(s)) ==
              Set([("B3", "on B3"), ("C10", "on C8"), ("D7", "on D5")])

        vml_path = XLSX._legacy_vml_path(xf, "xl/worksheets/sheet1.xml")
        vml = XLSX._vml_string(xf, vml_path)
        rows = sort([parse(Int, m[1]) for m in eachmatch(r"<x:Row>(\d+)</x:Row>", vml)])
        @test rows == [2, 6, 9]                           # 0-based: B3, D7, C10

        XLSX.deleteRows(s, 7)                              # the comment on D7 goes, box and all
        @test isnothing(XLSX.getComment(s, "D7"))
        @test length(XLSX.getComments(s)) == 2
        vml = XLSX._vml_string(xf, vml_path)
        @test count(r"<v:shape\b", vml) == 2
        @test XLSX.getComment(s, "C9").text == "on C8"

        XLSX.insertColumns(s, "A")
        @test Set(string(c.ref) for c in XLSX.getComments(s)) == Set(["C3", "D9"])
        vml = XLSX._vml_string(xf, vml_path)
        @test sort([parse(Int, m[1]) for m in eachmatch(r"<x:Column>(\d+)</x:Column>", vml)]) == [2, 3]

        r = roundtrip(xf)["Data"]
        @test XLSX.getComment(r, "D9").text == "on C8"
        SAVE_FILES && save_outfile(xf)                     # the moved comments, before they go

        XLSX.deleteRows(s, 1:20)                           # the last comments go: so do their parts
        @test isempty(XLSX.getComments(s))
        @test isnothing(element(xf, "legacyDrawing"))
        SAVE_FILES && save_outfile(xf)
    end

    @testset "threaded comments" begin
        xf, s = fresh()
        XLSX.addComment(s, "B3", "legacy")
        # A threaded comment as Excel 365 writes it, by hand.
        path = "xl/threadedComments/threadedComment1.xml"
        xf.data[path] = XML.parse("""<ThreadedComments xmlns="http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments"><threadedComment ref="B3" dT="2026-01-01T00:00:00.00" personId="{00000000-0000-0000-0000-000000000001}" id="{00000000-0000-0000-0000-000000000002}"><text>threaded</text></threadedComment></ThreadedComments>""", XML.Node)
        xf.files[path] = true
        XLSX.add_part_rel!(xf, "xl/worksheets/sheet1.xml", path, XLSX.REL_THREADED_COMMENT)
        XLSX.insertRows(s, 1)
        tc = only(XLSX.elements_with_tag(XLSX.xml_root_element(xf.data[path]), "threadedComment"))
        @test tc["ref"] == "B4"
        @test_throws XLSX.XLSXError XLSX.deleteRows(s, 4)  # can't delete a threaded comment
        @test tc["ref"] == "B4" && s[5, 1] == 4           # nothing changed
        XLSX.deleteRows(s, 2)                              # other rows are fine
        @test tc["ref"] == "B3"
    end

    # A table `Sales`, A1:D10, with headers A to D and numbers below.
    function tfresh(; totals=false)
        xf = XLSX.newxlsx("Data")
        s = xf["Data"]
        s["A1"] = ["A" "B" "C" "D"]
        s[2:10, 1:4] = reshape(1:36, 9, 4)
        XLSX.addsheet!(xf, "Other")
        XLSX.addtable!(s, totals ? "A1:D11" : "A1:D10"; name="Sales", has_totals_row=totals)
        return xf, s
    end
    tbl(s) = XLSX.table(s, "Sales")
    table_root(xf, s) = XLSX.xml_root_element(XLSX.get_xml_data(xf, XLSX._table_part_path(s, "Sales")))
    table_filter(xf, s) = XLSX.first_element_with_tag(table_root(xf, s), "autoFilter")

    @testset "tables: rows" begin
        xf, s = tfresh()
        XLSX.insertRows(s, 5; count=2)                     # inside: it grows
        @test tbl(s).ref == XLSX.CellRange("A1:D12")
        @test table_filter(xf, s)["ref"] == "A1:D12"
        XLSX.insertRows(s, 1)                              # at its header: it moves
        @test tbl(s).ref == XLSX.CellRange("A2:D13")
        XLSX.insertRows(s, 14)                             # just after it: unchanged
        @test tbl(s).ref == XLSX.CellRange("A2:D13")
        XLSX.deleteRows(s, 4:5)                            # some data rows: it shrinks
        @test tbl(s).ref == XLSX.CellRange("A2:D11")
        @test_throws XLSX.XLSXError XLSX.deleteRows(s, 2)  # its header row
        @test_throws XLSX.XLSXError XLSX.deleteRows(s, 3:11)   # all its data rows
        @test tbl(s).ref == XLSX.CellRange("A2:D11") && s[2, 1] == "A"   # nothing changed
        r = roundtrip(xf)["Data"]
        @test XLSX.table(r, "Sales").ref == XLSX.CellRange("A2:D11")
        @test isequal(XLSX.gettable(XLSX.table(r, "Sales")).data[1][1:4], [1, missing, missing, 4])   # with the inserted rows
        SAVE_FILES && save_outfile(xf)
    end

    @testset "tables: totals row" begin
        xf, s = tfresh(; totals=true)
        XLSX.settotals!(s, "Sales", "B" => :sum)
        XLSX.insertRows(s, 11)                             # at the totals row: a data row before it
        t = tbl(s)
        @test t.ref == XLSX.CellRange("A1:D12") && t.has_totals_row
        @test table_filter(xf, s)["ref"] == "A1:D11"
        @test_throws XLSX.XLSXError XLSX.deleteRows(s, 12)   # the totals row
        SAVE_FILES && save_outfile(xf)                     # the table with its totals row
        XLSX.removetotals!(s, "Sales")
        XLSX.deleteRows(s, 12)
        @test tbl(s).ref == XLSX.CellRange("A1:D11")
    end

    @testset "tables: deleted whole" begin
        xf, s = tfresh()
        o = xf["Other"]
        XLSX.setFormula(o, "A1", "=SUM(Sales[B])")
        XLSX.setFormula(o, "A2", "=SUM(Data!A1:A10)")
        XLSX.deleteRows(s, 1:10)
        @test isempty(XLSX.tables(s))
        @test formula(o, "A1") == "=SUM(#REF!)"
        @test formula(o, "A2") == "=SUM(Data!#REF!)"
        r = roundtrip(xf)
        @test isempty(XLSX.tables(r["Data"]))

        xf, s = tfresh()
        XLSX.deleteColumns(s, "A:D")
        @test isempty(XLSX.tables(s))
    end

    @testset "tables: columns" begin
        xf, s = tfresh()
        o = xf["Other"]
        XLSX.setFormula(o, "A1", "=SUM(Sales[B])")
        XLSX.setFormula(o, "A2", "=SUM(Sales[C])+Sales[[#Totals],[D]]")
        XLSX.setFormula(o, "A3", "=SUM(Sales[[B]:[C]])")
        XLSX.setFormula(o, "A4", "=\"Sales[B]\"&COUNT(Sales[#Data])")
        root = table_root(xf, s)
        i = findfirst(c -> XLSX.localname(c) == "autoFilter", root.children)
        root.children[i] = XML.Element("autoFilter", XML.Element("filterColumn"; colId="2"),   # on column C
                                       XML.Element("filterColumn"; colId="1"); ref="A1:D10")   # on column B

        XLSX.insertColumns(s, "C"; count=2)                  # inside: new table columns
        t = tbl(s)
        @test t.ref == XLSX.CellRange("A1:F10")
        @test t.columns == ["A", "B", "Column1", "Column2", "C", "D"]
        @test s["C1"] == "Column1" && s["D1"] == "Column2"   # their headers
        @test sort([fc["colId"] for fc in XLSX.xml_elements(table_filter(xf, s))]) == ["1", "4"]
        cols = XLSX.first_element_with_tag(table_root(xf, s), "tableColumns")
        @test cols["count"] == "6"
        @test length(unique(c["id"] for c in XLSX.xml_elements(cols))) == 6

        XLSX.insertColumns(s, "A")                           # at its first column: it moves
        @test tbl(s).ref == XLSX.CellRange("B1:G10")
        XLSX.insertColumns(s, "H")                           # just after it: unchanged
        @test tbl(s).ref == XLSX.CellRange("B1:G10")

        XLSX.deleteColumns(s, "C")                           # table column B
        t = tbl(s)
        @test t.columns == ["A", "Column1", "Column2", "C", "D"]
        @test t.ref == XLSX.CellRange("B1:F10")
        @test [fc["colId"] for fc in XLSX.xml_elements(table_filter(xf, s))] == ["3"]   # B's filter went
        @test formula(o, "A1") == "=SUM(#REF!)"
        @test formula(o, "A2") == "=SUM(Sales[C])+Sales[[#Totals],[D]]"
        @test formula(o, "A3") == "=SUM(#REF!)"
        @test formula(o, "A4") == "=\"Sales[B]\"&COUNT(Sales[#Data])"   # a string, and the whole data
        r = roundtrip(xf)
        @test XLSX.table(r["Data"], "Sales").columns == ["A", "Column1", "Column2", "C", "D"]
        SAVE_FILES && save_outfile(xf)
    end

    @testset "tables: formulas in them" begin
        xf, s = tfresh()
        cols = XLSX.first_element_with_tag(table_root(xf, s), "tableColumns")
        cols.children[end] = XML.Element("tableColumn", XML.Element("calculatedColumnFormula", XML.Text("Other!\$A\$5*2")); id="4", name="D")
        o = xf["Other"]
        XLSX.insertRows(o, 1)                                # a table on another sheet refers to it
        d = last(XLSX.xml_elements(XLSX.first_element_with_tag(table_root(xf, s), "tableColumns")))
        @test XML.simple_value(only(XLSX.xml_elements(d))) == "Other!\$A\$6*2"
    end

    @testset "structured references" begin
        dead = Dict{String,Union{Nothing,Set{String}}}("SALES" => Set(["B"]), "GONE" => nothing)
        k(f) = XLSX._kill_structured_refs(f, dead)
        @test k("=SUM(Sales[B])") == "=SUM(#REF!)"
        @test k("=SUM(sales[b])") == "=SUM(#REF!)"                       # names ignore case
        @test k("=Sales[[#This Row],[B]]+Sales[[#This Row],[C]]") == "=#REF!+Sales[[#This Row],[C]]"
        @test k("=SUM(Sales[[A]:[B]])") == "=SUM(#REF!)"
        @test k("=Sales[@B]") == "=#REF!"
        @test k("=Sales[C]+Gone[X]+Gone[#All]") == "=Sales[C]+#REF!+#REF!"
        @test k("=Sales['[B']]") == "=Sales['[B']]"                      # a column named [B]
        @test k("=Other[B]+\"Sales[B]\"") == "=Other[B]+\"Sales[B]\""
        @test XLSX._structured_ref_columns("[[#Headers],[It''s],['#No]]") == ["It's", "#No"]
    end

    @testset "limits and refusals" begin
        xf, s = fresh()
        s[XLSX.EXCEL_MAX_ROWS - 1, 1] = "near the end"
        @test_throws XLSX.XLSXError XLSX.insertRows(s, 1; count=2)       # would push it off
        XLSX.insertRows(s, 1)                                              # one row still fits
        @test s[XLSX.EXCEL_MAX_ROWS, 1] == "near the end"
        @test_throws XLSX.XLSXError XLSX.insertRows(s, 1; count=0)
        @test_throws XLSX.XLSXError XLSX.insertRows(s, 0)
        @test_throws XLSX.XLSXError XLSX.insertRows(s, "3:4")             # a single row
        @test_throws XLSX.XLSXError XLSX.deleteRows(s, "B:C")
        @test_throws XLSX.XLSXError XLSX.insertColumns(s, XLSX.EXCEL_MAX_COLS + 1)


        xf, s = fresh()
        r = roundtrip(xf)["Data"]
        @test_throws XLSX.XLSXError XLSX.insertRows(r, 1)                 # read-only

        xf = XLSX.newxlsx("data")
        xf["data"]["A1:A2"] = [1, 2]
        c = XLSX.Charts.addChart(xf, :line; sheetname="Trend")
        XLSX.Charts.addSeries(c, "data!A1:A2")
        @test_throws XLSX.XLSXError XLSX.insertRows(xf["Trend"], 1)
    end
end
