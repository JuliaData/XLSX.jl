# One call each for method variants the rest of the suite never reaches: alternative
# argument forms, Base/Tables interface methods and small helpers. Each test checks
# the variant agrees with the form it forwards to, rather than just running it.

@testset "dispatch coverage" begin

    @testset "defined names" begin
        xf = XLSX.newxlsx()
        sh = xf[1]
        wb = XLSX.get_workbook(xf)

        # Integer values that aren't Int64 go through the Integer forwarding methods
        XLSX.addDefName(xf, "wb_int", Int32(5))
        XLSX.addDefName(sh, "ws_int", Int32(6))
        @test XLSX.get_defined_name_value(wb, "wb_int") === Int64(5)
        @test XLSX.get_defined_name_value(sh, "ws_int") === Int64(6)
        @test !isnothing(XLSX.find_workbook_defined_name(xf, "wb_int"))
        @test !isnothing(XLSX.find_worksheet_defined_name(sh, "ws_int"))
        @test XLSX.is_worksheet_defined_name(wb, sh.name, "ws_int")
        @test !XLSX.is_defined_name_value_a_reference(Int32(1))
        @test XLSX.is_defined_name_value_a_constant(Int32(1))

        dnv = XLSX.DefinedNameValue(Int64(1), false)
        @test dnv.hidden == false && dnv.isabs == false

        dns = XLSX.getDefinedNames(xf)
        dn = only(filter(d -> d.name == "wb_int", dns))
        @test hash(dn) == hash(only(filter(d -> d.name == "wb_int", XLSX.getDefinedNames(xf))))
        s = sprint(show, MIME("text/plain"), dn)
        @test occursin("DefinedName: wb_int", s) && occursin("scope: Workbook", s) && occursin("value: 5", s)

        # a non-contiguous defined name follows its sheet when the sheet is renamed
        XLSX.addDefName(sh, "nc", XLSX.NonContiguousRange("Sheet1!A1:A2,Sheet1!C1"))
        XLSX.renamesheet!(sh, "Renamed")
        @test XLSX.get_defined_name_value(sh, "nc").sheet == "Renamed"
        SAVE_FILES && save_outfile(xf)
    end

    @testset "worksheet setters and ranges" begin
        xf = XLSX.newxlsx()
        sh = xf[1]
        sh[1:3, 1:4] = reshape(collect(1:12), 3, 4)

        # setdata! forms
        XLSX.setdata!(sh, XLSX.CellRef("A5"), Int32(7))
        @test sh["A5"] === Int64(7)
        XLSX.setdata!(sh, 6, 1, XLSX.CellValue(sh, "six"))
        @test sh["A6"] == "six"
        XLSX.setdata!(sh, 7, 1:2, [1, 2])
        @test sh["A7"] == 1 && sh["B7"] == 2
        @test XLSX.CellValue(Int32(3), XLSX.EmptyCellDataFormat()).value === Int64(3)

        # formats over vector / step-range columns, and colon rows
        XLSX.setFormat(sh, 1, [1, 3]; format="0.00")
        @test XLSX.getFormat(sh, "A1").format["numFmt"]["formatCode"] == XLSX.getFormat(sh, "C1").format["numFmt"]["formatCode"] == "0.00"
        XLSX.setUniformFormat(sh, 2, 1:2:3; format="0.0")
        @test XLSX.getcell(sh, "A2").style == XLSX.getcell(sh, "C2").style
        @test XLSX.getFormat(sh, "C2").format["numFmt"]["formatCode"] == "0.0"
        XLSX.setUniformFill(sh, 3, :; pattern="solid", fgColor="yellow")
        @test XLSX.getcell(sh, "A3").style == XLSX.getcell(sh, "D3").style
        @test XLSX.getFill(sh, "D3").fill["patternFill"]["fgrgb"] == "FFFFFF00"

        # merged cells: each argument form names exactly the merged range, so unmerges it
        merged() = something(XLSX.getMergedCells(sh), XLSX.CellRange[])
        for (rng, unmerge) in (("A1:A7", () -> XLSX.removeMergedCells(sh, XLSX.ColumnRange("A:A"))),
                               ("A1:A7", () -> XLSX.removeMergedCells(sh, XLSX.SheetColumnRange("Sheet1!A:A"))),
                               ("A1:A7", () -> XLSX.removeMergedCells(sh, :, 1)),
                               ("A1:D1", () -> XLSX.removeMergedCells(sh, XLSX.RowRange("1:1"))),
                               ("A1:D1", () -> XLSX.removeMergedCells(sh, XLSX.SheetRowRange("Sheet1!1:1"))),
                               ("A1:D1", () -> XLSX.removeMergedCells(sh, 1, :)),
                               ("A1:D7", () -> XLSX.removeMergedCells(sh, :, :)))
            XLSX.mergeCells(sh, rng)
            @test length(merged()) == 1
            unmerge()
            @test isempty(merged())
        end
        XLSX.mergeCells(sh, XLSX.NonContiguousRange("Sheet1!A1:B2,Sheet1!C1:D2"))
        @test length(merged()) == 2
        XLSX.removeMergedCells(sh)
        @test isempty(merged())

        # cell ranges and Base methods on references
        @test XLSX.getcellrange(sh, :) == XLSX.getcellrange(sh, XLSX.get_dimension(sh))
        @test isless(XLSX.CellRef("A2"), XLSX.CellRef("B1"))
        @test 2 in XLSX.ColumnRange("A:C") && 2 in XLSX.RowRange("1:3")
        @test convert(XLSX.RowRange, "2:4") == XLSX.RowRange("2:4")
        @test convert(XLSX.RowRange, XLSX.RowRange("2:4")) == XLSX.RowRange("2:4")
        srr = XLSX.SheetRowRange("Sheet1!2:4")
        @test string(srr) == "Sheet1!2:4"
        @test hash(srr) == hash(XLSX.SheetRowRange("Sheet1!2:4"))

        # panes from an anchor cell, Base.eachrow, and the cache flag on every holder
        XLSX.splitFreeze(sh, "B3")
        doc = XLSX.get_worksheet_xml_document(sh)
        i, j = XLSX.get_idces(doc, "worksheet", "sheetViews")
        k, l = XLSX.get_idces(doc[i][j], "sheetView", "pane")
        pane = doc[i][j][k][l]
        @test pane["topLeftCell"] == "B3" && pane["xSplit"] == "1" && pane["ySplit"] == "2"
        @test length(collect(Base.eachrow(sh))) == length(collect(XLSX.eachrow(sh)))
        itr = XLSX.eachrow(sh)
        @test XLSX.is_cache_enabled(XLSX.get_workbook(xf)) == XLSX.is_cache_enabled(itr) == XLSX.is_cache_enabled(xf)

        # the unexported alias and the Workbook forms of the content-type helpers
        XLSX.rename!(sh, "Alias")
        @test sh.name == "Alias"
        wb = XLSX.get_workbook(xf)
        XLSX.add_override!(wb, "/xl/dummy.xml", "application/xml")
        @test occursin("/xl/dummy.xml", XLSX.XML.write(XLSX.xmlroot(xf, "[Content_Types].xml")))
        XLSX.remove_override!(wb, "/xl/dummy.xml")
        @test !occursin("/xl/dummy.xml", XLSX.XML.write(XLSX.xmlroot(xf, "[Content_Types].xml")))
        SAVE_FILES && save_outfile(xf)
    end

    @testset "styles, colors and formulas" begin
        xf = XLSX.newxlsx()
        sh = xf[1]
        wb = XLSX.get_workbook(xf)
        sh["A1"] = 1.5
        XLSX.setFormat(sh, "A1"; format="0.00")
        sid = string(XLSX.getcell(sh, "A1").style)
        @test XLSX.styles_is_float(sh, sid) == XLSX.styles_is_float(wb, parse(Int, sid))
        @test XLSX.styles_is_datetime(sh, sid) == XLSX.styles_is_datetime(wb, parse(Int, sid))
        @test !isempty(XLSX.CellDataFormat(1)) && isempty(XLSX.EmptyCellDataFormat())
        @test XLSX.get_colorant(:red) == XLSX.get_colorant("red")

        @test XLSX.Formula() isa XLSX.EmptyFormula
        f  = XLSX.Formula("SUM(A1:A2)")
        rf = XLSX.ReferencedFormula("A1+1", 0, "B1:B3", Dict("t" => "shared"))
        fr = XLSX.FormulaReference(0, nothing)
        @test !isempty(f) && isempty(XLSX.Formula("")) && !isempty(rf) && !isempty(fr)
        @test hash(f) == hash(XLSX.Formula("SUM(A1:A2)"))
        @test hash(rf) == hash(copy(rf)) && hash(fr) == hash(copy(fr))
        c = copy(rf)
        @test c.unhandled == rf.unhandled && c.unhandled !== rf.unhandled   # Dict copied
        io = IOBuffer()
        XLSX.add_node_formula!(io, XLSX.CellFormula(f, XLSX.EmptyCellDataFormat()), "")
        @test occursin("SUM(A1:A2)", String(take!(io)))
        SAVE_FILES && save_outfile(xf)
    end

    @testset "cells and rich text" begin
        e = XLSX.EmptyCell(XLSX.CellRef("A1"))
        @test isempty(e) && !XLSX.iserror(e)
        @test !isempty(XLSX.Cell(XLSX.CellRef("A1"), UInt64(1), UInt32(0), UInt16(0), XLSX.CT_INT, false))
        @test hash(e) == hash(XLSX.EmptyCell(XLSX.CellRef("A1")))
        @test XLSX.get_error_string(nothing) == ""

        run = XLSX.RichTextRun("ab", nothing)
        @test length(run) == 2
        @test hash(run) == hash(XLSX.RichTextRun("ab"))
        rts = XLSX.RichTextString("ab", [run])
        @test String(rts) == "ab" && ncodeunits(rts) == 2
        @test codeunit(rts) == UInt8 && codeunit(rts, 1) == UInt8('a')
        @test isvalid(rts, 1) && rts[1] == XLSX.RichTextString("a", [XLSX.RichTextRun("a")])
        @test hash(rts) == hash(XLSX.RichTextString("ab", [XLSX.RichTextRun("ab")]))

        xf = XLSX.opentemplate(joinpath(data_directory, "is.xlsx"))
        sh = xf["Sheet1"]
        XLSX.setFont(sh, "A1"; name="Palatino")
        @test XLSX.getRichTextString(sh, 1, 1) == XLSX.getRichTextString(sh, "A1")
        @test XLSX.getRichTextString(xf, "Sheet1!A1") == XLSX.getRichTextString(sh, "A1")
        @test XLSX.get_sst(xf) === XLSX.get_sst(XLSX.get_workbook(xf))
        SAVE_FILES && save_outfile(xf)
    end

    @testset "tables" begin
        xf = XLSX.newxlsx()
        sh = xf[1]
        XLSX.writetable!(sh, [[1, 2], [3, 4]], ["a", "b"])
        XLSX.addtable!(sh, "A1:B3"; name="T")
        t = XLSX.table(sh, "T")
        XLSX.appendtable!(sh, t.id, [[5, 6]])
        @test XLSX.gettable(sh).data[1] == [1, 2, 5]
        XLSX.settotals!(sh, t.id; a=:sum)
        XLSX.settotals!(XLSX.table(sh, "T"); b=:sum)
        @test XLSX.table(sh, "T").has_totals_row

        @test Tables.istable(XLSX.Table) && Tables.rowaccess(XLSX.Table) && Tables.columnaccess(XLSX.Table)
        rows = Tables.rows(XLSX.table(sh, "T"))
        @test Tables.istable(typeof(rows)) && Tables.rowaccess(typeof(rows))
        @test Tables.rowtable(rows) == Tables.rowtable(XLSX.table(sh, "T"))

        dt = XLSX.gettable(sh)
        @test Tables.istable(XLSX.DataTable) && Tables.columnaccess(XLSX.DataTable)
        itr = XLSX.eachtablerow(sh)
        @test Tables.istable(typeof(itr)) && Tables.rowaccess(typeof(itr))
        r = first(itr)
        # by position, with Int and with another integer type (was a MethodError for both)
        @test Tables.getcolumn(r, 1) == Tables.getcolumn(r, Int32(1)) == Tables.getcolumn(r, :a)
        @test XLSX._as_vector((1, 2)) == [1, 2]
        @test XLSX._colname_prefix_string(sh, XLSX.EmptyCell(XLSX.CellRef("A1"))) == "#Empty"
        SAVE_FILES && save_outfile(xf)
    end

    @testset "file array" begin
        p = joinpath(data_directory, "general.xlsx")
        fa = XLSX.FileArray(p)
        @test fa[1] == read(p)[1]
    end

    @testset "charts" begin
        f = XLSX.readxlsx(joinpath(data_directory, "chart_appearance.xlsx"))
        c = XLSX.Charts.getCharts(f)[1]
        @test XLSX.Charts.getChartSchema(c) === :c
        @test XLSX.Charts.chartname(c) == c.name && XLSX.Charts.chartpath(c) == c.path
        @test XLSX.Charts.sheetname(c) == c.sheet
        @test XLSX.Charts.getLegendShapeProps(c) isa Union{Nothing,XLSX.Charts.DrawingShapeProps}
        @test XLSX.Charts.getChartTitleShapeProps(c) isa Union{Nothing,XLSX.Charts.DrawingShapeProps}
        g = first(XLSX.Charts.getChartGroups(c))
        @test isnothing(XLSX.Charts.getGroupSeriesLines(c, g))   # a bar chart without c:serLines

        s1 = XLSX.Charts.getChartSeries(c)[1]
        r = s1.values
        @test XLSX.Charts.iserror(r, 1) == XLSX.Charts.iserror(r)[1]
        @test XLSX.Charts.no_cached_values(r) == false

        @test XLSX.Charts._show_color(XLSX.Charts.SchemeColor(:accent1)) == "accent1"
        pfx = Dict(XLSX.Charts.NS_A => "a")
        @test XLSX.localname(XLSX.Charts._color_node_from(XLSX.Charts.SchemeColor(:accent1), pfx)) == "schemeClr"
        @test XLSX.localname(XLSX.Charts._color_node_from("FF0000", pfx)) == "srgbClr"
        @test XLSX.Charts._check("bar", (:bar, :col), "direction") === :bar

        sh = XLSX.newxlsx()[1]
        @test occursin("\$A\$1:\$A\$2", XLSX.Charts._chart_f(sh, XLSX.NonContiguousRange("Sheet1!A1:A2,Sheet1!C1")))
        @test_throws XLSX.XLSXError XLSX.Charts._chart_f(sh, 42)

        xf = XLSX.opentemplate(joinpath(data_directory, "chart_kinds.xlsx"))
        lc = XLSX.Charts.getCharts(xf["linemarkers"])[1]
        XLSX.Charts.setMarker(lc, 1; symbol=:circle, size=7)
        m = XLSX.Charts.getSeriesMarker(lc, 1)
        @test XLSX.localname(XLSX.Charts._node(lc, m)) == "marker"
        SAVE_FILES && save_outfile(xf)
    end
end
