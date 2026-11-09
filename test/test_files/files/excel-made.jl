# Files Excel 365 wrote, for features first tested against XML written by hand: what
# XLSX.jl reads from each, and its row and column edits compared with the same edits
# made in Excel (through its automation interface), whose results are the *_after,
# *_delA, *_insB, *_rows and *_inside files.
@testset "Excel-made fixtures" begin
    fixture(f) = joinpath(data_directory, f)
    strip_ns(s) = replace(s, r" xmlns(:\w+)?=\"[^\"]*\"" => "", r" xr:uid=\"[^\"]*\"" => "")

    @testset "threaded comments" begin
        s = XLSX.readxlsx(fixture("threaded_comments.xlsx"))[1]
        t = XLSX.getThreadedComments(s)
        @test [(c.ref, c.text, c.resolved) for c in t] == [("A1", "First thread", false), ("B2", "To be resolved", true)]
        @test [r.text for r in t[1].replies] == ["Reply one", "Reply two"] && isempty(t[2].replies)
        @test all(c -> c.author == "Tim Gebbels" && c.date isa Dates.DateTime, t)
        @test [c.ref for c in XLSX.getComments(s)] == ["A1", "B2", "C3"]   # the notes Excel writes beside the threads, and a plain one

        # A row and a column inserted: threads, notes and their boxes move as in Excel.
        # Excel re-sizes a thread's box as it saves, so a box is compared by its top-left cell.
        xf = XLSX.opentemplate(fixture("threaded_comments.xlsx"))
        XLSX.insertRows(xf[1], 1)
        XLSX.insertColumns(xf[1], "A")
        excel = XLSX.readxlsx(fixture("threaded_comments_after.xlsx"))
        threads(x) = [(c.ref, c.text, length(c.replies)) for c in XLSX.getThreadedComments(x[1])]
        @test threads(xf) == threads(excel)
        @test [c.ref for c in XLSX.getComments(xf[1])] == [c.ref for c in XLSX.getComments(excel[1])]
        function boxes(x)
            v = XLSX._vml_string(x, XLSX._legacy_vml_path(x, XLSX._sheet_path(x[1])))
            cells = [(m[1], m[2]) for m in eachmatch(r"<x:Row>(\d+)</x:Row>\s*<x:Column>(\d+)</x:Column>", v)]
            corners = [split(replace(strip(m[1]), " " => ""), ",")[[1, 3]] for m in eachmatch(r"<x:Anchor>([^<]*)</x:Anchor>", v)]
            return cells, corners
        end
        @test boxes(xf) == boxes(excel)
        SAVE_FILES && save_outfile(xf)
    end

    @testset "a data validation listing another sheet's cells" begin
        dvs(x) = [(string(d.range), d.type, d.formula1) for d in XLSX.getDataValidations(x["Sheet1"])]
        @test dvs(XLSX.readxlsx(fixture("dv_crosssheet.xlsx"))) ==
              [("B1:B1", :list, "\"Red, Green\""), ("A1:A10", :list, "Lists!\$A\$1:\$A\$4")]   # the second in x14 form
        xf = XLSX.opentemplate(fixture("dv_crosssheet.xlsx"))
        XLSX.insertRows(xf["Lists"], 1)          # the list moves
        XLSX.insertRows(xf["Sheet1"], 5)         # inside the rule's own range: it grows
        XLSX.deleteRows(xf["Lists"], 3)          # inside the list: it shrinks
        @test dvs(xf) == dvs(XLSX.readxlsx(fixture("dv_crosssheet_after.xlsx")))
        SAVE_FILES && save_outfile(xf)
    end

    @testset "an AutoFilter with criteria" begin
        s = XLSX.readxlsx(fixture("autofilter_range.xlsx"))[1]
        @test XLSX.getAutoFilter(s) == XLSX.CellRange("A1:C9")
        @test [r for r in 1:9 if XLSX.isRowHidden(s, r)] == [2, 5, 8]   # filtered out
        filt(x) = strip_ns(XLSX.XML.write(XLSX.first_element_with_tag(XLSX._sheet_root(x[1]), "autoFilter")))
        hidden(x) = [r for r in 1:12 if XLSX.isRowHidden(x[1], r)]
        # The criteria follow their columns (colId), and the hidden rows their rows.
        for (after, edit) in (("autofilter_range_delA.xlsx", s -> XLSX.deleteColumns(s, "A")),
                              ("autofilter_range_insB.xlsx", s -> XLSX.insertColumns(s, "B")),
                              ("autofilter_range_rows.xlsx", s -> (XLSX.insertRows(s, 4); XLSX.deleteRows(s, 6))))
            xf = XLSX.opentemplate(fixture("autofilter_range.xlsx"))
            edit(xf[1])
            excel = XLSX.opentemplate(fixture(after))
            @test filt(xf) == filt(excel)
            @test hidden(xf) == hidden(excel)
            SAVE_FILES && save_outfile(xf)
        end
    end

    @testset "hidden and grouped rows and columns" begin
        s = XLSX.readxlsx(fixture("hidden_grouped.xlsx"))[1]
        @test [r for r in 1:15 if XLSX.isRowHidden(s, r)] == [5, 9, 10]
        @test [(r, XLSX.getRowOutlineLevel(s, r)) for r in 1:15 if XLSX.getRowOutlineLevel(s, r) > 0] ==
              [(8, 1), (9, 2), (10, 2), (11, 1), (12, 1)]
        @test [c for c in 1:10 if XLSX.isColumnHidden(s, c)] == [3, 5, 6, 7]
        @test [(c, XLSX.getColumnOutlineLevel(s, c)) for c in 1:10 if XLSX.getColumnOutlineLevel(s, c) > 0] ==
              [(5, 1), (6, 1), (7, 1)]

        rows(x) = (w = x[1]; [(r, XLSX.isRowHidden(w, r), XLSX.getRowOutlineLevel(w, r),
                               get(something(get(something(w.unhandled_attributes, Dict()), r, nothing), Dict()), "collapsed", ""))
                              for r in 1:16 if XLSX.isRowHidden(w, r) || XLSX.getRowOutlineLevel(w, r) > 0 ||
                                  haskey(something(get(something(w.unhandled_attributes, Dict()), r, nothing), Dict()), "collapsed")])
        # Column definitions, less a width that isn't custom: Excel writes its default width,
        # which depends on the font's metrics, where XLSX.jl writes none; both mean the default.
        cols(x) = [replace(strip_ns(XLSX.XML.write(d)), r" width=\"[^\"]*\"" => (haskey(d, "customWidth") ? s -> s : ""))
                   for c in XLSX.xml_elements(XLSX._sheet_root(x[1])) if XLSX.localname(c) == "cols" for d in XLSX.xml_elements(c)]
        levels(x) = (n = XLSX.first_element_with_tag(XLSX._sheet_root(x[1]), "sheetFormatPr");
                     (get(n, "outlineLevelRow", nothing), get(n, "outlineLevelCol", nothing)))
        for (after, edit) in (
                # inside the level-2 group; the end of the level-1 group; inside the hidden grouped E:G; the hidden C
                ("hidden_grouped_after.xlsx", s -> (XLSX.insertRows(s, 9); XLSX.deleteRows(s, 12:13);
                                                    XLSX.insertColumns(s, "F"); XLSX.deleteColumns(s, "C"))),
                # between two hidden rows of a collapsed group: it shows, and the group is no longer collapsed
                ("hidden_grouped_inside.xlsx", s -> (XLSX.insertRows(s, 10); XLSX.insertColumns(s, "I"))))
            xf = XLSX.opentemplate(fixture("hidden_grouped.xlsx"))
            edit(xf[1])
            excel = XLSX.opentemplate(fixture(after))
            @test rows(xf) == rows(excel)
            @test cols(xf) == cols(excel)
            @test levels(xf) == levels(excel)
            SAVE_FILES && save_outfile(xf)
        end
    end

    @testset "a protected sheet" begin
        s = XLSX.readxlsx(fixture("protected_sheet.xlsx"))[1]
        @test XLSX.isSheetProtected(s)
        p = XLSX.getSheetProtection(s)
        @test p.password && !p.selectLockedCells && p.selectUnlockedCells && !p.objects && !p.scenarios
        @test [string(XLSX.CellRef(r, c)) for r in 1:6, c in 1:4 if !XLSX.getCellProtection(s, XLSX.CellRef(r, c)).locked] ==
              ["B2", "B3", "B4"]
        @test count(XLSX.getCellProtection(s, XLSX.CellRef(r, c)).hidden for r in 1:6, c in 1:4) == 1
    end

    @testset "formulas referring to another sheet" begin
        # As typed, one formula per cell (crosssheet); and given to whole ranges (range),
        # which Excel shares only where they refer to their own sheet.
        s = XLSX.readxlsx(fixture("shared_formula_crosssheet.xlsx"))["Sheet1"]
        @test [XLSX.getFormula(s, "B$r") for r in 1:10] == ["=Sheet2!A$r*2" for r in 1:10]
        shared = XLSX.readxlsx(fixture("shared_formula_range.xlsx"))
        @test [XLSX.getFormula(shared["Sheet1"], "C$r") for r in 1:3] == ["=B$r+Sheet2!\$A\$1" for r in 1:3]
        raw(n) = XLSX.ZipArchives.zip_readentry(XLSX.ZipArchives.ZipReader(read(fixture("shared_formula_range.xlsx"))), "xl/worksheets/sheet$n.xml", String)
        @test occursin("<f t=\"shared\" ref=\"D1:D10\"", raw(1)) && occursin("<f t=\"shared\" ref=\"C1:C10\"", raw(2))
        @test occursin("<c r=\"B1\"><f>Sheet2!A1*2</f>", raw(1))                                 # the cross-sheet ones aren't

        xf = XLSX.opentemplate(fixture("shared_formula_range.xlsx"))
        XLSX.insertRows(xf["Sheet2"], 5)         # between the rows Sheet1's formulas refer to
        XLSX.deleteRows(xf["Sheet1"], 3)         # inside Sheet1's shared blocks
        XLSX.deleteRows(xf["Sheet2"], 8:9)
        XLSX.insertRows(xf["Sheet1"], 1)
        excel = XLSX.readxlsx(fixture("shared_formula_range_after.xlsx"))
        formulas(x, sh) = [(string(c), XLSX.getFormula(x[sh], c)) for c in (XLSX.CellRef(r, k) for r in 1:12, k in 1:4)
                           if !isempty(something(XLSX.getFormula(x[sh], c), ""))]
        @test formulas(xf, "Sheet1") == formulas(excel, "Sheet1")
        @test formulas(xf, "Sheet2") == formulas(excel, "Sheet2")
        SAVE_FILES && save_outfile(xf)
    end
end
