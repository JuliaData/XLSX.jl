# Filter criteria and sorting with `setAutoFilter`. The expected hidden rows, sorted
# orders and XML are what Excel 365 gave for the same data and the same criteria,
# applied through its automation interface.
@testset "Filter criteria and sort" begin
    sheet_xml(xf, n=1) = replace(XML.write(xf.data["xl/worksheets/sheet$n.xml"]), r">\s+<" => "><")
    af_xml(xf, n=1) = match(r"<autoFilter.*</autoFilter>"s, sheet_xml(xf, n)).match
    hidden(s, rows=2:9) = [r for r in rows if XLSX.isRowHidden(s, r)]
    function roundtrip(xf)
        io = IOBuffer()
        XLSX.writexlsx(io, xf)
        return XLSX.openxlsx(IOBuffer(take!(io)); mode="rw")
    end

    # Excel's test data: 8 rows of a name, a colour (one empty), an amount shown as
    # "#,##0.00" and a date.
    function fresh(n=1)
        xf = XLSX.newxlsx("Sheet1")
        for k in 2:n
            XLSX.addsheet!(xf, "Sheet$k")
        end
        for k in 1:n
            s = xf[k]
            s["A1"] = ["Name" "Colour" "Amount" "When"]
            for (i, (name, col, amt)) in enumerate(zip(["apple", "Banana", "cherry", "date", "Elder", "fig", "grape", "Apricot"],
                                                       ["Red", "Blue", "red", "Green", "", "Blue", "Red", "Green"],
                                                       [5.0, 12.5, 3.0, 20.0, 7.0, 1234.5, 9.0, 15.0]))
                s[i+1, 1] = name
                isempty(col) || (s[i+1, 2] = col)
                s[i+1, 3] = amt
                s[i+1, 4] = XLSX.excel_value_to_date(46300 + 10 * (i - 1), false)
            end
            XLSX.setFormat(s, "C2:C9"; format="#,##0.00")
            XLSX.setFormat(s, "D2:D9"; format="yyyy-mm-dd")
        end
        return xf
    end

    @testset "as Excel filters" begin
        xf = fresh()
        s = xf[1]
        XLSX.setAutoFilter(s, "A1:D9"; filters=["Colour" => ["Red", "Blue", missing], "Amount" => [5, 1234.5, 12.5]])
        @test hidden(s) == [4, 5, 6, 8, 9]   # "red" is shown as Red: any case
        @test af_xml(xf) == """<autoFilter ref="A1:D9"><filterColumn colId="1"><filters blank="1"><filter val="Blue"/><filter val="Red"/></filters></filterColumn><filterColumn colId="2"><filters><filter val="1,234.50"/><filter val="12.50"/><filter val="5.00"/></filters></filterColumn></autoFilter>"""
        @test occursin("""<sheetPr filterMode="1"/>""", sheet_xml(xf))

        XLSX.setAutoFilter(s, "A1:D9"; filters=["Name" => ("==" => "a*", :or, "==" => "*e"), 3 => (">" => 5, :and, "<=" => 20)])
        @test hidden(s) == [2, 3, 4, 6, 7]
        @test af_xml(xf) == """<autoFilter ref="A1:D9"><filterColumn colId="0"><customFilters><customFilter val="a*"/><customFilter val="*e"/></customFilters></filterColumn><filterColumn colId="2"><customFilters and="1"><customFilter operator="greaterThan" val="5"/><customFilter operator="lessThanOrEqual" val="20"/></customFilters></filterColumn></autoFilter>"""

        XLSX.setAutoFilter(s, "A1:D9"; filters=["Amount" => (:top, 3)])
        @test hidden(s) == [2, 3, 4, 6, 8]   # the rows hidden before are shown again
        @test af_xml(xf) == """<autoFilter ref="A1:D9"><filterColumn colId="2"><top10 val="3" filterVal="15"/></filterColumn></autoFilter>"""

        XLSX.setAutoFilter(s, "A1:D9"; filters=["C" => (:bottom, 25, :percent)])
        @test hidden(s) == [3, 5, 6, 7, 8, 9]
        @test af_xml(xf) == """<autoFilter ref="A1:D9"><filterColumn colId="2"><top10 top="0" percent="1" val="25" filterVal="5"/></filterColumn></autoFilter>"""

        XLSX.setAutoFilter(s, "A1:D9"; filters=["Amount" => :aboveAverage])
        @test hidden(s) == [2, 3, 4, 5, 6, 8, 9]
        @test af_xml(xf) == """<autoFilter ref="A1:D9"><filterColumn colId="2"><dynamicFilter type="aboveAverage" val="163.25"/></filterColumn></autoFilter>"""
        XLSX.setAutoFilter(s, "A1:D9"; filters=["Amount" => :belowAverage])
        @test hidden(s) == [7]

        # Read back, with Excel's defined name and the rows' hidden flags
        r = roundtrip(xf)
        @test hidden(r[1]) == [7]
        @test XLSX.getAutoFilter(r[1]) == XLSX.CellRange("A1:D9")
        SAVE_FILES && save_outfile(xf)
    end

    @testset "relative dates" begin
        # The periods Excel gave on 2026-11-13, as serial days, the last one not included
        day = Dates.Date(2026, 11, 13)
        serials(p) = Tuple(XLSX.date_to_excel_value(d, false) for d in XLSX._date_period(p, day))
        @test serials(:today) == (46339, 46340)
        @test serials(:thisWeek) == (46334, 46341)     # from Sunday
        @test serials(:lastMonth) == (46296, 46327)
        @test serials(:thisQuarter) == (46296, 46388)
        @test serials(:thisYear) == (46023, 46388)
        @test serials(:yearToDate) == (46023, 46340)
        @test serials(:yesterday) == (46338, 46339) && serials(:tomorrow) == (46340, 46341)
        @test serials(:nextWeek) == (46341, 46348) && serials(:lastWeek) == (46327, 46334)

        # Days around today, with times: whole days are compared
        xf = XLSX.newxlsx()
        s = xf[1]
        s["A1"] = ["When" "N"]
        t = Dates.today()
        offsets = [-400, -60, -30, -7, -1, 0, 0.5, 1, 7, 30, 60]
        for (i, o) in enumerate(offsets)
            s[i+1, 1] = Dates.DateTime(t) + Dates.Minute(round(Int, o * 24 * 60))
            s[i+1, 2] = i - 1
        end
        rows = 2:12
        XLSX.setAutoFilter(s, "A1:B12"; filters=["When" => :today])
        @test hidden(s, rows) == [2, 3, 4, 5, 6, 9, 10, 11, 12]
        @test occursin("""<dynamicFilter type="today" val="$(XLSX.date_to_excel_value(t, false))" maxVal="$(XLSX.date_to_excel_value(t, false) + 1)"/>""", af_xml(xf))
        XLSX.setAutoFilter(s, "A1:B12"; filters=["N" => :belowAverage])
        @test hidden(s, rows) == [7, 8, 9, 10, 11, 12]   # strictly below

        # Quarters and months of any year, as Excel gave for 2026-11-13 ± offsets
        xf = XLSX.newxlsx()
        s = xf[1]
        s["A1"] = "When"
        for (i, o) in enumerate(offsets)
            s[i+1, 1] = XLSX.excel_value_to_datetime(46339 + o, false)
        end
        XLSX.setAutoFilter(s, "A1:A12"; filters=["When" => :Q1])
        @test hidden(s, rows) == collect(2:11)   # only 2027-01-12
        @test occursin("""<dynamicFilter type="Q1"/>""", af_xml(xf))
        XLSX.setAutoFilter(s, "A1:A12"; filters=[1 => :M3])
        @test hidden(s, rows) == collect(2:12)
        XLSX.setAutoFilter(s, "A1:A12"; filters=[1 => :M11])
        @test hidden(s, rows) == [2, 3, 4, 11, 12]   # November only
    end

    @testset "values and conditions" begin
        xf = XLSX.newxlsx()
        s = xf[1]
        s["A1"] = ["Text" "Num" "Day" "Flag"]
        s["A2"] = ["a*b" 1 Dates.Date(2026, 1, 5) true]
        s["A3"] = ["axb" 2 Dates.DateTime(2026, 1, 5, 18, 0) false]
        s["A4"] = ["A?C" 3 Dates.Date(2026, 1, 6) true]
        s["A5"] = ["abc" 7 Dates.Date(2026, 2, 1) false]
        s[6, 2] = 8
        rng = "A1:D6"
        rows = 2:6

        # Wildcards: `~` before a wildcard means the character itself
        XLSX.setAutoFilter(s, rng; filters=["Text" => ("==" => "a~*b")])
        @test hidden(s, rows) == [3, 4, 5, 6]
        XLSX.setAutoFilter(s, rng; filters=["Text" => ("==" => "a?c")])
        @test hidden(s, rows) == [2, 3, 6]
        XLSX.setAutoFilter(s, rng; filters=["Text" => ("!=" => "a*")])
        @test hidden(s, rows) == [2, 3, 4, 5]   # the empty cell is shown
        XLSX.setAutoFilter(s, rng; filters=["Text" => (">" => "AXA")])   # text compared in any case
        @test hidden(s, rows) == [2, 4, 5, 6]

        # A single value; a date list by the day, written as Excel groups dates
        XLSX.setAutoFilter(s, rng; filters=["Num" => 7])
        @test hidden(s, rows) == [2, 3, 4, 6]
        @test occursin("""<filters><filter val="7"/></filters>""", af_xml(xf))
        XLSX.setAutoFilter(s, rng; filters=["Day" => [Dates.Date(2026, 1, 5)]])
        @test hidden(s, rows) == [4, 5, 6]
        @test occursin("""<dateGroupItem year="2026" month="1" day="5" dateTimeGrouping="day"/>""", af_xml(xf))
        XLSX.setAutoFilter(s, rng; filters=["Day" => (">=" => Dates.Date(2026, 1, 6))])
        @test hidden(s, rows) == [2, 3, 6]
        @test occursin("""val="$(XLSX.date_to_excel_value(Dates.Date(2026, 1, 6), false))\"""", af_xml(xf))
        XLSX.setAutoFilter(s, rng; filters=["Flag" => [true]])
        @test hidden(s, rows) == [3, 5, 6]
        XLSX.setAutoFilter(s, rng; filters=["Flag" => missing])
        @test hidden(s, rows) == [2, 3, 4, 5]
        XLSX.setAutoFilter(s, rng; filters=["Num" => [99]])   # matches nothing: written as given
        @test hidden(s, rows) == collect(rows)
        @test occursin("""<filter val="99"/>""", af_xml(xf))
        XLSX.setAutoFilter(s, rng; filters=["Num" => (:top, 10)])   # more than there are
        @test isempty(hidden(s, rows))
        XLSX.setAutoFilter(s, rng; filters=["Text" => (:top, 2)])   # no numbers
        @test hidden(s, rows) == collect(rows)
        @test occursin("""<top10 val="2"/>""", af_xml(xf))
        @test XLSX._wildcard_match("a.b(c)", "A.B(C)") && !XLSX._wildcard_match("a.b", "axb")
    end

    @testset "columns and errors" begin
        xf = fresh()
        s = xf[1]
        for col in ("Colour", "B", "b", 2)
            XLSX.setAutoFilter(s, "A1:D9"; filters=[col => "Blue"])
            @test hidden(s) == [2, 4, 5, 6, 8, 9]
        end
        XLSX.setAutoFilter(s, "A1:D9"; filters=Dict("Colour" => "Blue"))
        @test hidden(s) == [2, 4, 5, 6, 8, 9]
        XLSX.setAutoFilter(s, "A1:D9"; filters="Colour" => "Blue")
        @test hidden(s) == [2, 4, 5, 6, 8, 9]
        for bad in (["Nope" => "x"], ["E" => "x"], [5 => "x"], [0 => "x"], ["Colour" => "x", 2 => "y"],
                    ["Colour" => :sometimes], ["Colour" => (:top, 0)], ["Colour" => (:top, 101, :percent)],
                    ["Colour" => (:top, 3, :items)], ["Colour" => ("~" => 3)], ["Colour" => (">" => true)],
                    ["Colour" => ("=" => 1, :xor, "=" => 2)], ["Colour" => String[]], ["Colour" => [1 + 2im]],
                    ["Colour"], 3)
            @test_throws XLSX.XLSXError XLSX.setAutoFilter(s, "A1:D9"; filters=bad)
        end
        for bad in ("Nope", "Amount" => :up, String[], fill("Amount", 65))
            @test_throws XLSX.XLSXError XLSX.setAutoFilter(s, "A1:D9"; sort=bad)
        end
        @test hidden(s) == [2, 4, 5, 6, 8, 9]   # a refused call changed nothing
    end

    @testset "keeping and removing" begin
        xf = fresh()
        s = xf[1]
        XLSX.setAutoFilter(s, "A1:D9"; filters=["Colour" => "Blue"])
        XLSX.setAutoFilter(s, "A1:D9"; sort="Amount")   # the criteria are kept
        @test occursin("<filterColumn colId=\"1\">", af_xml(xf)) && occursin("<sortState", af_xml(xf))
        @test [s[r, 1] for r in 2:9] == ["cherry", "apple", "Elder", "grape", "Banana", "Apricot", "date", "fig"]
        @test hidden(s) == [2, 3, 4, 5, 7, 8]   # hidden by the criteria again in the new order
        XLSX.setAutoFilter(s, "A1:D9")   # both kept
        @test occursin("<filterColumn", af_xml(xf)) && occursin("<sortState", af_xml(xf))

        XLSX.setAutoFilter(s, "A1:D9"; filters=[])   # criteria removed, all rows shown; the sort kept
        @test isempty(hidden(s))
        @test !occursin("<filterColumn", af_xml(xf)) && occursin("<sortState", af_xml(xf))
        @test !occursin("filterMode", sheet_xml(xf))

        XLSX.setAutoFilter(s, "A1:D9"; filters=["Colour" => "Blue"])
        XLSX.setAutoFilter(s, "A1:C9")   # another range: neither kept
        @test occursin("""<autoFilter ref="A1:C9"/>""", sheet_xml(xf))
        @test !occursin("sortState", sheet_xml(xf)) && !occursin("filterMode", sheet_xml(xf))

        XLSX.setAutoFilter(s, "A1:D9"; filters=["Colour" => "Blue"], sort="Name" => :descending)
        @test XLSX.removeAutoFilter(s)
        @test !occursin("sortState", sheet_xml(xf)) && !occursin("filterMode", sheet_xml(xf))
        @test !isempty(hidden(s))   # rows stay hidden, as documented

        # A sheet whose <sheetPr> holds more keeps it when the filter goes
        XLSX.setTabColor(s, "FFFF0000")
        XLSX.setAutoFilter(s, "A1:D9"; filters=["Colour" => "Blue"])
        @test occursin("""<sheetPr filterMode="1"><tabColor""", sheet_xml(xf))
        XLSX.removeAutoFilter(s)
        @test occursin("<sheetPr><tabColor", sheet_xml(xf))

        # Excel's file: its criteria kept as the same range is sorted
        xf = XLSX.opentemplate(joinpath(data_directory, "autofilter_range.xlsx"))
        s = xf[1]
        rng = string(XLSX.getAutoFilter(s))
        XLSX.setAutoFilter(s, rng; sort=1 => :descending)
        @test length(collect(eachmatch(r"<filterColumn", af_xml(xf)))) == 2
        @test occursin("xr:uid=", af_xml(xf))   # Excel's other attributes kept
        r = roundtrip(xf)
        @test XLSX.getAutoFilter(r[1]) == XLSX.CellRange(rng)
    end

    @testset "sorting as Excel sorts" begin
        # Numbers, text in any case, logicals, errors, and empty cells last either way
        xf = XLSX.newxlsx()
        s = xf[1]
        s["A1"] = ["K" "Id"]
        s["A2"] = "b"; s["A3"] = 2; s["A4"] = true
        XLSX.setFormula(s, "A5", "=1/0"; value=XLSX.XL_DIV0)
        s["A7"] = "B"; s["A8"] = 10; s["A9"] = false; s["A10"] = "a"
        for r in 2:10
            s[r, 2] = r
        end
        XLSX.setAutoFilter(s, "A1:B10"; sort="K")
        @test [s[r, 2] for r in 2:10] == [3, 8, 10, 2, 7, 9, 4, 5, 6]
        @test XLSX.geterror(s, "A9") == "#DIV/0!" && XLSX.getFormula(s, "A9") == "=1/0"
        XLSX.setAutoFilter(s, "A1:B10"; sort="Id")
        XLSX.setAutoFilter(s, "A1:B10"; sort="K" => :descending)
        @test [s[r, 2] for r in 2:10] == [5, 4, 9, 2, 7, 10, 8, 3, 6]
        @test occursin("""<sortState ref="A2:B10"><sortCondition descending="1" ref="A2:A10"/></sortState>""", sheet_xml(xf))

        # By two columns: the second orders the first's ties
        xf = fresh()
        s = xf[1]
        XLSX.setAutoFilter(s, "A1:D9"; sort=["Colour", "Amount" => :descending])
        @test isequal([(s[r, 2], s[r, 3]) for r in 2:9],
              [("Blue", 1234.5), ("Blue", 12.5), ("Green", 20.0), ("Green", 15.0), ("Red", 9.0), ("Red", 5.0), ("red", 3.0), (missing, 7.0)])
        @test XLSX.getText(s, "C2") == "1,234.50" && XLSX.getText(s, "D2") == "2026-11-24"   # formats move with the cells
        SAVE_FILES && save_outfile(xf)
    end

    @testset "formulas, notes, links" begin
        xf = XLSX.newxlsx()
        s = xf[1]
        s["A1"] = ["K" "F1" "F2" "F3"]
        for (i, v) in enumerate([3, 1, 2])
            r = i + 1
            s[r, 1] = v
            XLSX.setFormula(s, "B$r", "=A$r*10")
            XLSX.setFormula(s, "C$r", "=A\$2+A$(r+1)")
            XLSX.setFormula(s, "D$r", "=SUM(A\$2:A$r)")
        end
        XLSX.setFormula(s, "F2", "=A3")   # outside the range: left alone
        s["E3"] = "beside"                # outside the range's columns: stays
        XLSX.setFill(s, "B3"; pattern="solid", fgColor="FFFF0000")
        XLSX.addComment(s, "A3", "note on key 1")
        XLSX.addHyperlink(s, "B3"; url="https://example.com/1")
        XLSX.addThreadedComment(s, "D4", "on key 2"; author="Tim")
        XLSX.setAutoFilter(s, "A1:D4"; sort="K")

        @test [s[r, 1] for r in 2:4] == [1, 2, 3]
        @test [XLSX.getFormula(s, "B$r") for r in 2:4] == ["=A2*10", "=A3*10", "=A4*10"]
        @test [XLSX.getFormula(s, "C$r") for r in 2:4] == ["=A\$2+A3", "=A\$2+A4", "=A\$2+A5"]
        @test [XLSX.getFormula(s, "D$r") for r in 2:4] == ["=SUM(A\$2:A2)", "=SUM(A\$2:A3)", "=SUM(A\$2:A4)"]
        @test XLSX.getFormula(s, "F2") == "=A3"
        @test s["E3"] == "beside" && ismissing(s["E2"])
        @test XLSX.getFill(s, "B2").fill["patternFill"]["patternType"] == "solid"
        @test isnothing(XLSX.getFill(s, "B3")) || XLSX.getFill(s, "B3").fill["patternFill"]["patternType"] != "solid"
        @test [c.ref for c in XLSX.getComments(s) if !startswith(c.author, "tc=")] == ["A2"]
        @test XLSX.getHyperlink(s, "B2").url == "https://example.com/1" && isnothing(XLSX.getHyperlink(s, "B3"))
        @test only(XLSX.getThreadedComments(s)).ref == "D3"
        vml = XLSX._vml_string(xf, XLSX._legacy_vml_path(xf, XLSX._sheet_path(s)))
        @test occursin(r"<x:Row>1</x:Row>\s*<x:Column>0</x:Column>", vml)   # A2's box
        @test occursin(r"<x:Row>2</x:Row>\s*<x:Column>3</x:Column>", vml)   # D3's box

        r = roundtrip(xf)
        @test XLSX.getFormula(r[1], "C3") == "=A\$2+A4"
        @test XLSX.getComment(r[1], "A2").text == "note on key 1"
        SAVE_FILES && save_outfile(xf)
    end

    @testset "sorting shared formulas" begin
        # Excel's own shared formulas are moved one by one, as plain formulas
        xf = XLSX.newxlsx()
        s = xf[1]
        s["A1"] = ["K" "F"]
        s["A2"] = [3; 1; 2;;]
        XLSX.setFormula(s, "B2:B4", "=A2*2")
        @test XLSX.get_formula_from_cache(s, XLSX.CellRef("B3")) isa XLSX.FormulaReference
        XLSX.setAutoFilter(s, "A1:B4"; sort=1)
        @test [XLSX.getFormula(s, "B$r") for r in 2:4] == ["=A2*2", "=A3*2", "=A4*2"]
        @test all(r -> XLSX.get_formula_from_cache(s, XLSX.CellRef("B$r")) isa XLSX.Formula, 2:4)

        # ...unless the formula is shared with cells outside the rows
        xf = XLSX.newxlsx()
        s = xf[1]
        s["A1"] = ["K" "F"]
        s["A2"] = [3; 1; 2; 5;;]
        XLSX.setFormula(s, "B2:B5", "=A2*2")
        @test_throws XLSX.XLSXError XLSX.setAutoFilter(s, "A1:B4"; sort=1)
        @test [s[r, 1] for r in 2:5] == [3, 1, 2, 5]   # nothing moved
    end

    @testset "what can't be sorted" begin
        xf = fresh()
        s = xf[1]
        XLSX.mergeCells(s, "A3:B3")
        @test_throws XLSX.XLSXError XLSX.setAutoFilter(s, "A1:D9"; sort=1)
        @test XLSX.setAutoFilter(s, "A1:D9"; filters=["Name" => "a*"]) === nothing   # filtering is fine

        xf = fresh()
        s = xf[1]
        XLSX.addHyperlink(s, "B3:B4"; url="https://example.com")
        @test_throws XLSX.XLSXError XLSX.setAutoFilter(s, "A1:D9"; sort=1)
        XLSX.deleteHyperlink(s, "B3")
        XLSX.addHyperlink(s, "D3:E3"; url="https://example.com")   # reaches outside the columns
        @test_throws XLSX.XLSXError XLSX.setAutoFilter(s, "A1:D9"; sort=1)
        @test s["A2"] == "apple"

        # One data row, or rows already in order: nothing moves
        xf = fresh()
        s = xf[1]
        XLSX.setAutoFilter(s, "A1:D2"; sort=1)
        @test s["A2"] == "apple"
        XLSX.setAutoFilter(s, "A1:D9"; sort="When")
        @test [s[r, 1] for r in 2:9] == ["apple", "Banana", "cherry", "date", "Elder", "fig", "grape", "Apricot"]
        @test occursin("""<sortState ref="A2:D9"><sortCondition ref="D2:D9"/></sortState>""", af_xml(xf))
    end
end
