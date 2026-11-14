@testset "ignored errors" begin
    ignored_xml(ws) = only(XLSX._sheet_top_elements(XLSX.get_xlsxfile(ws), XLSX._sheet_path(ws), ("ignoredErrors",)))
    sqrefs(ws) = [(e["sqref"], [first(p) for p in XLSX._atts_vector(e)][2:end]) for e in XLSX.xml_elements(ignored_xml(ws))]
    function roundtrip(xf)
        io = IOBuffer()
        XLSX.writexlsx(io, xf)
        return XLSX.readxlsx(IOBuffer(take!(io)))
    end
    # every cell's ignored warnings, for comparing sheets whatever the order of their ranges
    cells(ws, rng) = Dict(string(c) => XLSX.getIgnoredErrors(ws, c) for c in XLSX.CellRange(rng))

    @testset "as Excel writes them" begin
        ex = XLSX.readxlsx(joinpath(data_directory, "ignored_errors_excel.xlsx"))[1]   # read-only
        @test XLSX.getIgnoredErrors(ex, "A2") == [:evalError, :numberStoredAsText]
        @test [XLSX.getIgnoredErrors(ex, "E$i") for i in 1:9] == [[k] for k in keys(XLSX.IGNORABLE_ERRORS)]
        @test isempty(XLSX.getIgnoredErrors(ex, "B1"))

        # the same steps as were taken in Excel
        xf = XLSX.newxlsx()
        s = xf[1]
        @test isempty(XLSX.getIgnoredErrors(s))
        XLSX.setIgnoredErrors(s, "A1:A3"; numberStoredAsText=true)
        @test sqrefs(s) == [("A1:A3", ["numberStoredAsText"])]
        XLSX.setIgnoredErrors(s, "C5"; numberStoredAsText=true)
        @test sqrefs(s) == [("A1:A3 C5", ["numberStoredAsText"])]                    # one per set of warnings
        XLSX.setIgnoredErrors(s, "A2"; evalError=true)
        @test sqrefs(s) == [("A1 A3 C5", ["numberStoredAsText"]), ("A2", ["evalError", "numberStoredAsText"])]
        for (i, k) in enumerate(keys(XLSX.IGNORABLE_ERRORS))
            XLSX.setIgnoredErrors(s, "E$i"; (k => true,)...)
        end
        @test cells(s, "A1:H12") == cells(ex, "A1:H12")
        @test [last(x) for x in sqrefs(s)] == [last(x) for x in sqrefs(ex)]           # Excel's attribute order
        tags = [XLSX.localname(c) for c in XLSX.xml_elements(XLSX._sheet_root(s))]
        @test findfirst(==("ignoredErrors"), tags) == findfirst(==("pageMargins"), tags) + 1

        @test XLSX.getIgnoredErrors(s)[1] == ("A1,A3,C5,E3" => [:numberStoredAsText])
        r = roundtrip(xf)
        @test XLSX.getIgnoredErrors(r[1]) == XLSX.getIgnoredErrors(s)
        SAVE_FILES && save_outfile(xf)
    end

    @testset "changing them" begin
        xf = XLSX.newxlsx()
        s = xf[1]
        XLSX.setIgnoredErrors(s, "B2:D4"; numberStoredAsText=true, formula=true)
        XLSX.setIgnoredErrors(s, "C3"; formula=false)                                 # a hole in the middle
        @test XLSX.getIgnoredErrors(s, "C3") == [:numberStoredAsText]
        @test XLSX.getIgnoredErrors(s, "B2") == XLSX.getIgnoredErrors(s, "D4") == [:numberStoredAsText, :formula]
        @test length(XLSX.getIgnoredErrors(s)) == 2
        XLSX.setIgnoredErrors(s, "C3"; formula=true)                                  # back again
        @test cells(s, "B2:D4") == Dict(string(c) => [:numberStoredAsText, :formula] for c in XLSX.CellRange("B2:D4"))
        @test length(XLSX.getIgnoredErrors(s)) == 1

        XLSX.setIgnoredErrors(s, "C:C"; evalError=true)                               # whole column
        @test XLSX.getIgnoredErrors(s, "C1048576") == [:evalError]
        @test XLSX.getIgnoredErrors(s, "C3") == [:evalError, :numberStoredAsText, :formula]
        @test any(e -> occursin("C5:C1048576", e["sqref"]), XLSX.xml_elements(ignored_xml(s)))   # to the foot of the sheet
        XLSX.setIgnoredErrors(s, "5:5"; listDataValidation=true)                      # whole row
        @test XLSX.getIgnoredErrors(s, "XFD5") == [:listDataValidation]
        @test XLSX.getIgnoredErrors(s, "C5") == [:evalError, :listDataValidation]

        XLSX.setIgnoredErrors(s, "A1:B2,B2:C3"; twoDigitTextYear=true)               # overlapping parts, counted once
        @test XLSX.getIgnoredErrors(s, "B2") == [:twoDigitTextYear, :numberStoredAsText, :formula]
        @test count(r -> occursin(r"(^|,)B2(,|$)", first(r)), XLSX.getIgnoredErrors(s)) <= 1
        XLSX.setIgnoredErrors(s, XLSX.CellRange("H8:H9"); calculatedColumn=true)
        XLSX.addDefinedName(s, "ids", "J1:J3")
        XLSX.setIgnoredErrors(s, "ids"; numberStoredAsText=true)
        @test XLSX.getIgnoredErrors(s, "H9") == [:calculatedColumn] && XLSX.getIgnoredErrors(s, "J2") == [:numberStoredAsText]

        @test XLSX.removeIgnoredErrors(s, "H8:H9")
        @test !XLSX.removeIgnoredErrors(s, "H8:H9")
        @test isempty(XLSX.getIgnoredErrors(s, "H8"))
        @test XLSX.removeIgnoredErrors(s)
        @test isempty(XLSX.getIgnoredErrors(s)) && !XLSX.removeIgnoredErrors(s)
        @test !("ignoredErrors" in [XLSX.localname(c) for c in XLSX.xml_elements(XLSX._sheet_root(s))])
        XLSX.setIgnoredErrors(s, "A1"; formula=true)
        XLSX.setIgnoredErrors(s, "A1"; formula=false)                                 # emptied, so removed
        @test !("ignoredErrors" in [XLSX.localname(c) for c in XLSX.xml_elements(XLSX._sheet_root(s))])
        XLSX.setIgnoredErrors(s, "A1"; formula=false)                                 # nothing to change
        @test isempty(XLSX.getIgnoredErrors(s))
        SAVE_FILES && save_outfile(xf)
    end

    @testset "inserting and deleting rows and columns" begin
        xf = XLSX.newxlsx()
        s = xf[1]
        s["A1:A6"] = string.(1:6)
        XLSX.setIgnoredErrors(s, "A2:A5"; numberStoredAsText=true)
        XLSX.setIgnoredErrors(s, "A3"; evalError=true)
        XLSX.insertRows(s, 3; count=2)                                                # inside the range
        @test XLSX.getIgnoredErrors(s, "A5") == [:evalError, :numberStoredAsText]     # the old A3
        @test XLSX.getIgnoredErrors(s, "A7") == [:numberStoredAsText]                 # the old A5
        @test isempty(XLSX.getIgnoredErrors(s, "A3"))                                 # between A2 and the old A3, kept apart once A3 differed
        @test isempty(XLSX.getIgnoredErrors(s, "A8"))
        XLSX.setIgnoredErrors(s, "A5"; evalError=false)                               # still editable after the shift
        @test length(XLSX.getIgnoredErrors(s)) == 1
        XLSX.insertColumns(s, 1)
        @test XLSX.getIgnoredErrors(s, "B5") == [:numberStoredAsText] && isempty(XLSX.getIgnoredErrors(s, "A5"))
        XLSX.deleteColumns(s, 2)
        @test isempty(XLSX.getIgnoredErrors(s))
        @test !("ignoredErrors" in [XLSX.localname(c) for c in XLSX.xml_elements(XLSX._sheet_root(s))])
        SAVE_FILES && save_outfile(xf)
    end

    @testset "range arithmetic" begin
        R = XLSX.CellRange
        @test XLSX._overlap(R("A1:C3"), R("B2:D4")) == R("B2:C3")
        @test isnothing(XLSX._overlap(R("A1:A3"), R("B1:B3")))
        @test XLSX._subtract(R("A1:C3"), R("B2:B2")) == [R("A1:C1"), R("A3:C3"), R("A2:A2"), R("C2:C2")]
        @test XLSX._subtract(R("A1:C3"), R("A1:C3")) == R[]
        @test XLSX._subtract(R("A1:A3"), R("C1:C3")) == [R("A1:A3")]
        @test XLSX._minus_all(R("A1:D1"), [R("B1:B1"), R("D1:D1")]) == [R("A1:A1"), R("C1:C1")]
    end

    @testset "errors" begin
        xf = XLSX.newxlsx()
        s = xf[1]
        @test_throws XLSX.XLSXError XLSX.setIgnoredErrors(s, "A1")
        @test_throws XLSX.XLSXError XLSX.setIgnoredErrors(s, "A1"; numberAsText=true)
        @test_throws XLSX.XLSXError XLSX.setIgnoredErrors(s, "A1"; formula=1)
        @test_throws XLSX.XLSXError XLSX.setIgnoredErrors(s, "nonsense"; formula=true)
        @test_throws XLSX.XLSXError XLSX.getIgnoredErrors(s, "A1:B2")
        r = roundtrip(xf)
        @test isempty(XLSX.getIgnoredErrors(r[1]))
        @test_throws XLSX.XLSXError XLSX.setIgnoredErrors(r[1], "A1"; formula=true)
        @test_throws XLSX.XLSXError XLSX.removeIgnoredErrors(r[1])
        c = XLSX.Charts.addChart(xf, :line; sheetname="Chart")
        @test_throws XLSX.XLSXError XLSX.setIgnoredErrors(xf["Chart"], "A1"; formula=true)
    end
end
