@testset "Defined Names" begin # Issue #148
    @test XLSX.is_defined_name_value_a_reference(XLSX.SheetCellRef("Sheet1!A1"))
    @test XLSX.is_defined_name_value_a_reference(XLSX.SheetCellRange("Sheet1!A1:B2"))
    @test !XLSX.is_defined_name_value_a_reference(1)
    @test !XLSX.is_defined_name_value_a_reference(1.2)
    @test !XLSX.is_defined_name_value_a_reference("Hey")
    @test !XLSX.is_defined_name_value_a_reference(missing)

    f = XLSX.opentemplate(joinpath(data_directory, "general.xlsx"))

    result = XLSX.getAllDefinedNames(f)

    @test eltype(result) == XLSX.DefinedName
    @test length(result) == 16
    @test all(r -> r.name isa String, result)
    @test all(r -> isnothing(r.scope) || r.scope isa String, result)

    # sorted by scope then name, workbook scope first
    @test issorted(result, by=r -> (something(r.scope, ""), uppercase(r.name)))
    @test findlast(r -> isnothing(r.scope), result) < findfirst(r -> !isnothing(r.scope), result)

    # workbook scope — values come out typed, not as strings
    @test any(r -> r.name == "CONST_INT" && isnothing(r.scope) && r.value == 100, result)
    @test any(r -> r.name == "CONST_FLOAT" && isnothing(r.scope) && r.value == 10.2, result)
    @test any(r -> r.name == "CONST_DATE" && isnothing(r.scope) && r.value == 43383, result)
    @test any(r -> r.name == "SINGLE_CELL" && isnothing(r.scope) && r.value == XLSX.SheetCellRef("named_ranges!A2"), result)
    @test any(r -> r.name == "RANGE_B4C5" && isnothing(r.scope) && r.value == XLSX.SheetCellRange("named_ranges!B4:C5"), result)
    @test any(r -> r.name == "LOCAL_NAME" && isnothing(r.scope) && r.value == "out there in the cold", result)

    # worksheet scope
    @test any(r -> r.name == "LOCAL_INT" && r.scope == "named_ranges" && r.value == 1000, result)
    @test any(r -> r.name == "LOCAL_NAME" && r.scope == "named_ranges" && r.value == "Hey You", result)
    @test any(r -> r.name == "LOCAL_REF" && r.scope == "named_ranges" && r.value == XLSX.SheetCellRange("named_ranges!A15:B15"), result)
    @test any(r -> r.name == "CONST_LOCAL_INT" && r.scope == "named_ranges" && r.value == 100, result)
    @test any(r -> r.name == "LOCAL_INT" && r.scope == "named_ranges_2" && r.value == 2000, result)
    @test any(r -> r.name == "LOCAL_REF" && r.scope == "named_ranges_2" && r.value == XLSX.SheetCellRange("named_ranges_2!D1:E1"), result)

    # the same name at several scopes
    local_int_entries = filter(r -> r.name == "LOCAL_INT", result)
    @test length(local_int_entries) == 3
    @test any(r -> isnothing(r.scope), local_int_entries)
    @test any(r -> r.scope == "named_ranges", local_int_entries)
    @test any(r -> r.scope == "named_ranges_2", local_int_entries)

    const_local_int_entries = filter(r -> r.name == "CONST_LOCAL_INT", result)
    @test length(const_local_int_entries) == 2
    @test any(r -> isnothing(r.scope), const_local_int_entries)
    @test any(r -> r.scope == "named_ranges", const_local_int_entries)

    # and the scoped accessors partition it
    @test length(XLSX.getDefinedNames(f)) == 9
    @test [r.name for r in XLSX.getAllDefinedNames(f) if isnothing(r.scope)] ==
        [r.name for r in XLSX.getDefinedNames(f)]

    @test f["SINGLE_CELL"] == "single cell A2"
    @test f["RANGE_B4C5"] == Any["range B4:C5" "range B4:C5"; "range B4:C5" "range B4:C5"]
    @test f["CONST_DATE"] == 43383
    @test isapprox(f["CONST_FLOAT"], 10.2)
    @test f["CONST_INT"] == 100
    @test f["LOCAL_INT"] == 2000
    @test f["named_ranges_2"]["LOCAL_INT"] == 2000
    @test f["named_ranges"]["LOCAL_INT"] == 1000
    @test f["named_ranges"]["LOCAL_NAME"] == "Hey You"
    @test f["named_ranges_2"]["LOCAL_NAME"] == "out there in the cold"
    @test f["named_ranges"]["SINGLE_CELL"] == "single cell A2"

    @test_throws XLSX.XLSXError f["header_error"]["LOCAL_REF"]
    @test f["named_ranges"]["LOCAL_REF"][1] == 10
    @test f["named_ranges"]["LOCAL_REF"][2] == 20
    @test f["named_ranges_2"]["LOCAL_REF"][1] == "local"
    @test f["named_ranges_2"]["LOCAL_REF"][2] == "reference"

    XLSX.addDefinedName(f["lookup"], "Life_the_Universe_and_Everything", 42)
    XLSX.addDefinedName(f["lookup"], "FirstName", "Hello World")
    XLSX.addDefinedName(f["lookup"], "single", "C2"; absolute=true)
    XLSX.addDefinedName(f["lookup"], "range", "C3:C5"; absolute=true)
    XLSX.addDefinedName(f["lookup"], "NonContig", "C3:C5,D3:D5"; absolute=true)
    @test f["lookup"]["Life_the_Universe_and_Everything"] == 42
    @test f["lookup"]["FirstName"] == "Hello World"
    @test f["lookup"]["single"] == "NAME"
    @test f["lookup"]["range"] == Any["name1"; "name2"; "name3";;] # A 2D Array, size (3, 1)
    @test f["lookup"]["NonContig"] == [["name1"; "name2"; "name3";;], [100; 200; 300;;]] # NonContiguousRanges return a vector of matrices

    XLSX.addDefinedName(f, "Life_the_Universe_and_Everything", 42)
    XLSX.addDefinedName(f, "FirstName", "Hello World")
    XLSX.addDefinedName(f, "single", "lookup!C2"; absolute=true)
    XLSX.addDefinedName(f, "range", "lookup!C3:C5"; absolute=true)
    XLSX.addDefinedName(f, "NonContig", "lookup!C3:C5,lookup!D3:D5"; absolute=true)
    @test f["Life_the_Universe_and_Everything"] == 42
    @test f["FirstName"] == "Hello World"
    @test f["single"] == "NAME"
    @test f["range"] == Any["name1"; "name2"; "name3";;] # A 2D Array, size (3, 1)
    @test f["NonContig"] == [["name1"; "name2"; "name3";;], [100; 200; 300;;]] # NonContiguousRanges return a vector of matrices

    XLSX.setFont(f["lookup"], "NonContig"; name="Arial", size=12, color="FF0000FF", bold=true, italic=true, under="single", strike=true)
    @test XLSX.getFont(f["lookup"], "C3").font == Dict("i" => nothing, "b" => nothing, "u" => nothing, "strike" => nothing, "sz" => Dict("val" => "12"), "name" => Dict("val" => "Arial"), "color" => Dict("rgb" => "FF0000FF"))
    @test XLSX.getFont(f["lookup"], "C4").font == Dict("i" => nothing, "b" => nothing, "u" => nothing, "strike" => nothing, "sz" => Dict("val" => "12"), "name" => Dict("val" => "Arial"), "color" => Dict("rgb" => "FF0000FF"))
    @test XLSX.getFont(f["lookup"], "C5").font == Dict("i" => nothing, "b" => nothing, "u" => nothing, "strike" => nothing, "sz" => Dict("val" => "12"), "name" => Dict("val" => "Arial"), "color" => Dict("rgb" => "FF0000FF"))
    @test XLSX.getFont(f["lookup"], "D3").font == Dict("i" => nothing, "b" => nothing, "u" => nothing, "strike" => nothing, "sz" => Dict("val" => "12"), "name" => Dict("val" => "Arial"), "color" => Dict("rgb" => "FF0000FF"))
    @test XLSX.getFont(f["lookup"], "D4").font == Dict("i" => nothing, "b" => nothing, "u" => nothing, "strike" => nothing, "sz" => Dict("val" => "12"), "name" => Dict("val" => "Arial"), "color" => Dict("rgb" => "FF0000FF"))
    @test XLSX.getFont(f["lookup"], "D5").font == Dict("i" => nothing, "b" => nothing, "u" => nothing, "strike" => nothing, "sz" => Dict("val" => "12"), "name" => Dict("val" => "Arial"), "color" => Dict("rgb" => "FF0000FF"))
    XLSX.setFont(f, "single"; name="Arial", size=12, color="FF0000FF", bold=true, italic=true, under="double", strike=true)
    @test XLSX.getFont(f["lookup"], "C2").font == Dict("i" => nothing, "b" => nothing, "u" => Dict("val" => "double"), "strike" => nothing, "sz" => Dict("val" => "12"), "name" => Dict("val" => "Arial"), "color" => Dict("rgb" => "FF0000FF"))

    XLSX.writexlsx("mytest.xlsx", f, overwrite=true)
    SAVE_FILES && save_outfile("mytest.xlsx")

    f = XLSX.readxlsx("mytest.xlsx")
    @test f["Life_the_Universe_and_Everything"] == 42
    @test f["FirstName"] == "Hello World"
    @test f["single"] == "NAME"
    @test f["range"] == Any["name1"; "name2"; "name3";;] # A 2D Array, size (3, 1)
    @test f["NonContig"] == [["name1"; "name2"; "name3";;], [100; 200; 300;;]] # NonContiguousRanges return a vector of matrices

    @test XLSX.readdata(joinpath(data_directory, "general.xlsx"), "SINGLE_CELL") == "single cell A2"
    @test XLSX.readdata(joinpath(data_directory, "general.xlsx"), "RANGE_B4C5") == Any["range B4:C5" "range B4:C5"; "range B4:C5" "range B4:C5"]

    f = XLSX.newxlsx()
    s = f[1]
    s["A1:B3"] = "Hello world"
    XLSX.addDefinedName(f, "Life_the_Universe_and_Everything", 42)
    XLSX.addDefinedName(f[1], "FirstName", "Hello World")
    XLSX.addDefinedName(f, "MyCell", "Sheet1!A1")
    XLSX.addDefinedName(f[1], "YourCells", "Sheet1!A2:B3")
    @test_throws XLSX.XLSXError XLSX.addDefinedName(s, "yourcells", "Sheet1!A2:B3") # not unique (case insensitive)
    @test_throws XLSX.XLSXError XLSX.addDefinedName(s, "firstname", "NewText") # not unique (case insensitive)
    @test_throws XLSX.XLSXError s["FirstName"] = 32
    s["MyCell"] = true
    @test s["MyCell"] == true
    s["YourCells"] = false
    @test s["YourCells"] == Any[false false; false false]

    XLSX.writexlsx("mytest.xlsx", f, overwrite=true)
    SAVE_FILES && save_outfile("mytest.xlsx")
    f = XLSX.readxlsx("mytest.xlsx")
    @test f[1]["MyCell"] == true
    @test f[1]["YourCells"] == Any[false false; false false]

    @test_throws XLSX.XLSXError XLSX.addDefinedName(f, "A1", "Sheet1!B1")
    @test_throws XLSX.XLSXError XLSX.addDefinedName(f, "A1:A3", "Sheet1!B2:B3")
    @test_throws XLSX.XLSXError XLSX.addDefinedName(f, "A1,A3", 42)
    @test_throws XLSX.XLSXError XLSX.addDefinedName(s, "Sheet1!A1", "Sheet1!B1")
    @test_throws XLSX.XLSXError XLSX.addDefinedName(s, "Sheet1!A1:A3", "Sheet1!B2:B3")
    @test_throws XLSX.XLSXError XLSX.addDefinedName(s, "Sheet1!A1,Sheet!A3", 42)

    f=XLSX.newxlsx()
    XLSX.addsheet!(f, "Tim's Sheet")
    XLSX.addsheet!(f, "Ano'ther She'et")
    f[2]["A1"] = "tim"
    f[3]["A1"] = "another"
    XLSX.addDefinedName(f, "mine", "Tim's Sheet!A1")
    XLSX.addDefinedName(f, "yours", "Ano'ther She'et!A1")
    @test f["mine"] == "tim"
    @test f["yours"] == "another"
    @test string(XLSX.get_defined_name_value(XLSX.get_workbook(f), "mine")) == "'Tim''s Sheet'!A1"
    @test string(XLSX.get_defined_name_value(XLSX.get_workbook(f), "yours")) == "'Ano''ther She''et'!A1"

    XLSX.writexlsx("mytest.xlsx", f, overwrite=true)
    SAVE_FILES && save_outfile("mytest.xlsx")

    ff = XLSX.openxlsx("mytest.xlsx", mode="rw")

    @test XLSX.hassheet(ff, "Tim's Sheet")
    @test XLSX.hassheet(ff, "Ano'ther She'et")
    @test ff[2]["A1"] == "tim"
    @test ff[3]["A1"] == "another"
    @test ff["mine"] == "tim"
    @test ff["yours"] == "another"
    @test string(XLSX.get_defined_name_value(XLSX.get_workbook(ff), "mine")) == "'Tim''s Sheet'!A1"
    @test string(XLSX.get_defined_name_value(XLSX.get_workbook(ff), "yours")) == "'Ano''ther She''et'!A1"
    SAVE_FILES && save_outfile(ff)

    isfile("mytest.xlsx") && rm("mytest.xlsx")

    @testset "Defined name scope keys" begin
        # `Workbook.worksheet_names` is keyed on `sheetId`, which is stable, not on
        # the sheet's position, which is not. Deleting a sheet is the cheapest way
        # to force the two apart: sheetIds are never reused, so the sheets that
        # survive keep ids that no longer match their ordinals.

        f = XLSX.newxlsx()
        XLSX.addsheet!(f, "two")
        s3 = XLSX.addsheet!(f, "three")
        XLSX.deletesheet!(f, "two")

        wb = XLSX.get_workbook(f)
        @test XLSX.sheetnames(f) == ["Sheet1", "three"]
        @test s3.sheetId == 3
        @test XLSX.ordinal_sheet_number(wb, s3.name) == 2   # id and position now differ

        # --- the name must be findable on the sheet it was added to -------------
        XLSX.addDefinedName(s3, "on_three", "A1")
        @test XLSX.is_worksheet_defined_name(s3, "on_three")
        @test !XLSX.is_worksheet_defined_name(f[1], "on_three")
        @test XLSX.get_defined_name_value(s3, "on_three") == XLSX.SheetCellRef("three!A1")
        @test haskey(wb.worksheet_names, (s3.sheetId, "on_three"))

        # ... which also means adding it twice is caught
        @test_throws XLSX.XLSXError XLSX.addDefinedName(s3, "on_three", "B2")

        # ... and that the scope resolves when listing
        @test any(r -> r.name == "on_three" && r.scope == "three", XLSX.getAllDefinedNames(f))

        # --- and it must survive a round trip on the right sheet ---------------
        io = IOBuffer()
        XLSX.writexlsx(io, f)
        g = XLSX.readxlsx(seekstart(io))

        gs3 = g["three"]
        @test XLSX.is_worksheet_defined_name(gs3, "on_three")
        @test !XLSX.is_worksheet_defined_name(g[1], "on_three")
        @test any(r -> r.name == "on_three" && r.scope == "three", XLSX.getAllDefinedNames(g))

        SAVE_FILES && save_outfile(f)
    end

    @testset "deleteDefinedName" begin

        template = joinpath(data_directory, "general.xlsx")
        dnames(x) = [dn.name for dn in XLSX.getDefinedNames(x)]

        @testset "single name, by scope" begin
            f = XLSX.opentemplate(template)
            ws = f["named_ranges"]

            wb_target = first(dnames(f))
            ws_target = first(dnames(ws))

            @test isnothing(XLSX.deleteDefinedName(f, wb_target))
            @test wb_target ∉ dnames(f)

            XLSX.deleteDefinedName(ws, ws_target)
            @test ws_target ∉ dnames(ws)

            # matched case-insensitively, as in Excel
            next_target = first(dnames(f))
            XLSX.deleteDefinedName(f, lowercase(next_target))
            @test next_target ∉ dnames(f)
            SAVE_FILES && save_outfile(f)
        end

        @testset "scopes are independent" begin
            f = XLSX.opentemplate(template)
            ws = f["named_ranges"]
            ws2 = f["named_ranges_2"]

            # LOCAL_INT exists at workbook scope and on both sheets
            @test "LOCAL_INT" ∈ dnames(f)
            @test "LOCAL_INT" ∈ dnames(ws)
            @test "LOCAL_INT" ∈ dnames(ws2)

            XLSX.deleteDefinedName(ws, "LOCAL_INT")
            @test "LOCAL_INT" ∉ dnames(ws)
            @test "LOCAL_INT" ∈ dnames(f)
            @test "LOCAL_INT" ∈ dnames(ws2)
            SAVE_FILES && save_outfile(f)
        end

        @testset "errors" begin
            f = XLSX.opentemplate(template)
            ws = f["named_ranges"]

            @test_throws XLSX.XLSXError XLSX.deleteDefinedName(f, "NO_SUCH_NAME")
            @test_throws XLSX.XLSXError XLSX.deleteDefinedName(ws, "NO_SUCH_NAME")
            @test_throws XLSX.XLSXError XLSX.deleteDefinedName(f, "")
            @test_throws XLSX.XLSXError XLSX.deleteDefinedName(f, ["", ""])

            # a name defined in the other scope is not defined in this one
            wb_only = first(setdiff(dnames(f), dnames(ws)))
            @test_throws XLSX.XLSXError XLSX.deleteDefinedName(ws, wb_only)
            ws_only = first(setdiff(dnames(ws), dnames(f)))
            @test_throws XLSX.XLSXError XLSX.deleteDefinedName(f, ws_only)
            SAVE_FILES && save_outfile(f)
        end

        @testset "vector of names" begin
            f = XLSX.opentemplate(template)
            targets = dnames(f)[1:2]

            XLSX.deleteDefinedName(f, targets)
            @test all(∉(dnames(f)), targets)

            # an empty collection is a no-op, not an error
            n = length(dnames(f))
            @test isnothing(XLSX.deleteDefinedName(f, String[]))
            @test isnothing(XLSX.deleteDefinedName(f, XLSX.DefinedName[]))
            @test length(dnames(f)) == n
            SAVE_FILES && save_outfile(f)
        end

        @testset "defined-name convenience methods" begin
            xf = XLSX.newxlsx()
            ws = xf[1]
            XLSX.addDefinedName(xf, "wbInt", 5)                       # workbook scope, integer value
            XLSX.addDefinedName(ws, "wsInt", 7)                       # worksheet scope, integer value
            XLSX.addDefinedName(ws, "wsRange", XLSX.SheetCellRange("Sheet1!A1:B2"))
            names_ws = [d.name for d in XLSX.getDefinedNames(ws)]
            @test "wsInt" in names_ws && "wsRange" in names_ws
            @test "wbInt" in [d.name for d in XLSX.getDefinedNames(xf)]
            XLSX.deleteDefinedName(ws, ["wsInt", "wsRange"])          # vector form
            @test isempty(XLSX.getDefinedNames(ws))
            SAVE_FILES && save_outfile(xf)
        end

        @testset "validation precedes deletion" begin
            f = XLSX.opentemplate(template)
            before = dnames(f)
            good = first(before)

            @test_throws XLSX.XLSXError XLSX.deleteDefinedName(f, [good, "NO_SUCH_NAME"])
            @test dnames(f) == before          # the good name survived

            # the same name twice, including via a different casing
            @test_throws XLSX.XLSXError XLSX.deleteDefinedName(f, [good, good])
            @test_throws XLSX.XLSXError XLSX.deleteDefinedName(f, [good, lowercase(good)])
            @test dnames(f) == before
            SAVE_FILES && save_outfile(f)
        end

        @testset "DefinedName vectors compose per scope" begin
            f = XLSX.opentemplate(template)
            ws = f["named_ranges"]

            n_wb = length(dnames(f))
            XLSX.deleteDefinedName(ws, XLSX.getDefinedNames(ws))
            @test isempty(dnames(ws))
            @test length(dnames(f)) == n_wb                    # workbook untouched
            @test !isempty(dnames(f["named_ranges_2"]))        # other sheet untouched

            # scope is the first argument; a row from elsewhere is refused, not routed
            g = XLSX.opentemplate(template)
            gws = g["named_ranges"]
            before = XLSX.getAllDefinedNames(g)
            @test_throws XLSX.XLSXError XLSX.deleteDefinedName(gws, XLSX.getDefinedNames(g))
            @test_throws XLSX.XLSXError XLSX.deleteDefinedName(g, XLSX.getDefinedNames(gws))
            @test_throws XLSX.XLSXError XLSX.deleteDefinedName(gws, XLSX.getAllDefinedNames(g))
            @test XLSX.getAllDefinedNames(g) == before
            SAVE_FILES && save_outfile(g)
            SAVE_FILES && save_outfile(f)
        end

        @testset "clearing one scope" begin
            f = XLSX.opentemplate(template)
            ws = f["named_ranges"]

            XLSX.deleteDefinedName(ws, XLSX.getDefinedNames(ws))
            @test isempty(dnames(ws))
            @test !isempty(dnames(f))
            @test !isempty(dnames(f["named_ranges_2"]))

            XLSX.deleteDefinedName(f, XLSX.getDefinedNames(f))
            @test isempty(dnames(f))
            @test !isempty(dnames(f["named_ranges_2"]))
            SAVE_FILES && save_outfile(f)
        end

        @testset "deleteAllDefinedNames crosses scopes" begin
            f = XLSX.opentemplate(template)
            wb = XLSX.get_workbook(f)

            @test !isempty(XLSX.getAllDefinedNames(f))
            XLSX.deleteAllDefinedNames(f)
            @test isempty(XLSX.getAllDefinedNames(f))
            @test isempty(wb.workbook_names)
            @test isempty(wb.worksheet_names)

            # general.xlsx ships with a <definedNames> block, so this catches the
            # block being written back stale when the last name is removed
            io = IOBuffer()
            XLSX.writexlsx(io, f)
            bytes = take!(io)

            workbook_xml = zip_readentry(ZipReader(bytes), "xl/workbook.xml", String)
            @test !occursin("definedName", workbook_xml)

            g = XLSX.readxlsx(IOBuffer(bytes))
            @test isempty(XLSX.getAllDefinedNames(g))
            SAVE_FILES && save_outfile(f)
        end

        @testset "deletions survive a save" begin
            f = XLSX.opentemplate(template)
            ws = f["named_ranges"]

            gone_wb = first(dnames(f))
            gone_ws = first(dnames(ws))
            XLSX.deleteDefinedName(f, gone_wb)
            XLSX.deleteDefinedName(ws, gone_ws)
            kept = XLSX.getAllDefinedNames(f)

            io = IOBuffer()
            XLSX.writexlsx(io, f)
            g = XLSX.readxlsx(seekstart(io))

            @test XLSX.getAllDefinedNames(g) == kept
            @test gone_wb ∉ dnames(g)
            @test gone_ws ∉ dnames(g["named_ranges"])

            SAVE_FILES && save_outfile(f)
        end
    end

    @testset "Rename sheet updates defined names" begin
        f = XLSX.opentemplate(joinpath(data_directory, "general.xlsx"))
        wb = XLSX.get_workbook(f)
        ws = f["named_ranges"]

        # workbook-scoped range on that sheet, and a worksheet-scoped one
        @test XLSX.get_defined_name_value(wb, "SINGLE_CELL") == XLSX.SheetCellRef("named_ranges!A2")
        @test XLSX.get_defined_name_value(ws, "LOCAL_REF") == XLSX.SheetCellRange("named_ranges!A15:B15")
        # a constant, which must not be touched
        @test XLSX.get_defined_name_value(wb, "CONST_INT") == 100

        XLSX.renamesheet!(ws, "renamed ranges")   # space forces quoting on write

        @test XLSX.get_defined_name_value(wb, "SINGLE_CELL") == XLSX.SheetCellRef("'renamed ranges'!A2")
        @test XLSX.get_defined_name_value(ws, "LOCAL_REF") == XLSX.SheetCellRange("'renamed ranges'!A15:B15")
        @test XLSX.get_defined_name_value(wb, "CONST_INT") == 100

        # values still resolve, and survive a round trip
        @test f["SINGLE_CELL"] == "single cell A2"

        io = IOBuffer()
        XLSX.writexlsx(io, f)
        g = XLSX.readxlsx(seekstart(io))

        @test XLSX.get_defined_name_value(XLSX.get_workbook(g), "SINGLE_CELL") ==
            XLSX.SheetCellRef("'renamed ranges'!A2")
        @test g["SINGLE_CELL"] == "single cell A2"
        @test !any(r -> occursin("named_ranges!", string(r.value)), XLSX.getAllDefinedNames(g))

        SAVE_FILES && save_outfile(f)
    end

    @testset "DefinedName round trip preserves absoluteness" begin
        f = XLSX.opentemplate(joinpath(data_directory, "general.xlsx"))

        XLSX.addDefinedName(f, "REL_REF", "named_ranges!B4"; absolute=false)
        XLSX.addDefinedName(f, "ABS_REF", "named_ranges!B4"; absolute=true)

        rel = only(filter(dn -> dn.name == "REL_REF", XLSX.getDefinedNames(f)))
        abs = only(filter(dn -> dn.name == "ABS_REF", XLSX.getDefinedNames(f)))

        @test rel.absolute == false
        @test abs.absolute == true
        @test rel.value == abs.value            # same target...
        @test rel != abs                        # ...different definitions
        @test sprint(show, rel) != sprint(show, abs)
        @test occursin("\$B\$4", sprint(show, abs))

        # a DefinedName's own fields are enough to recreate it
        XLSX.deleteDefinedName(f, [rel, abs])
        XLSX.addDefinedName(f, rel.name, rel.value; absolute=rel.absolute)
        XLSX.addDefinedName(f, abs.name, abs.value; absolute=abs.absolute)
        @test only(filter(dn -> dn.name == "REL_REF", XLSX.getDefinedNames(f))) == rel
        @test only(filter(dn -> dn.name == "ABS_REF", XLSX.getDefinedNames(f))) == abs

        # and it survives a save
        io = IOBuffer()
        XLSX.writexlsx(io, f)
        g = XLSX.readxlsx(seekstart(io))
        @test only(filter(dn -> dn.name == "REL_REF", XLSX.getDefinedNames(g))) == rel
        @test only(filter(dn -> dn.name == "ABS_REF", XLSX.getDefinedNames(g))) == abs

        # non-contiguous ranges carry a vector of flags, not a scalar
        XLSX.addDefinedName(f, "NC_REF", "named_ranges!A1,named_ranges!C3:D4")
        nc = only(filter(dn -> dn.name == "NC_REF", XLSX.getDefinedNames(f)))
        @test nc.absolute isa Vector{Bool}
        @test length(nc.absolute) == length(nc.value.rng)
        SAVE_FILES && save_outfile(f)
    end

    @testset "DefinedName vector show" begin
        f = XLSX.opentemplate(joinpath(data_directory, "general.xlsx"))
        ws = f["named_ranges"]

        single_scope = sprint(show, MIME"text/plain"(), XLSX.getDefinedNames(ws))
        all_scopes = sprint(show, MIME"text/plain"(), XLSX.getAllDefinedNames(f))

        # header, and the scope column only when the vector spans scopes
        @test occursin("Name", single_scope) && occursin("Value", single_scope)
        @test !occursin("Scope", single_scope)
        @test occursin("Scope", all_scopes)
        @test occursin("[Workbook]", all_scopes)
        @test occursin("[Worksheet: \"named_ranges\"]", all_scopes)
        wb_dn = first(filter(dn -> isnothing(dn.scope), XLSX.getAllDefinedNames(f)))
        ws_dn = first(filter(dn -> dn.scope == "named_ranges", XLSX.getAllDefinedNames(f)))
        @test occursin("[" * XLSX._show_scope(wb_dn) * "]", all_scopes)
        @test occursin("[" * XLSX._show_scope(ws_dn) * "]", all_scopes)

        # one header, one rule, one line per entry
        @test count(==('\n'), all_scopes) == length(XLSX.getAllDefinedNames(f)) + 3

        @test sprint(show, MIME"text/plain"(), XLSX.DefinedName[]) == "0-element Vector{DefinedName}"

        # a name too long for the column is marked as cut, not silently shortened
        long = "A_defined_name_that_is_far_longer_than_the_column"
        XLSX.addDefinedName(f, long, 1)
        out = sprint(show, MIME"text/plain"(), XLSX.getDefinedNames(f))
        @test !occursin(long, out)
        @test occursin("…", out)
        SAVE_FILES && save_outfile(f)
    end

    @testset "parse_defined_name_value" begin
        @testset "parse_defined_name_value - relative sheet cell reference" begin
            v, isabs = XLSX.parse_defined_name_value("Sheet1!A1")
            @test v == XLSX.SheetCellRef("Sheet1!A1")
            @test isabs == false
        end

        @testset "parse_defined_name_value - relative sheet cell range" begin
            v, isabs = XLSX.parse_defined_name_value("Sheet1!A1:B2")
            @test v == XLSX.SheetCellRange("Sheet1!A1:B2")
            @test isabs == false
        end

        @testset "parse_defined_name_value - absolute forms" begin
            v, isabs = XLSX.parse_defined_name_value("Sheet1!\$A\$1")
            @test v == XLSX.SheetCellRef("Sheet1!A1")
            @test isabs == true

            v2, isabs2 = XLSX.parse_defined_name_value("Sheet1!\$A\$1:\$B\$2")
            @test v2 == XLSX.SheetCellRange("Sheet1!A1:B2")
            @test isabs2 == true
        end

        @testset "parse_defined_name_value - empty string" begin
            v, isabs = XLSX.parse_defined_name_value("")
            @test ismissing(v)
            @test isabs == false
        end

        @testset "parse_defined_name_value - quoted, numeric and fallback strings" begin
            @test XLSX.parse_defined_name_value("\"hello\"") == ("hello", false)
            @test ismissing(first(XLSX.parse_defined_name_value("\"\"")))
            @test XLSX.parse_defined_name_value("42") == (42, false)
            @test XLSX.parse_defined_name_value("3.5") == (3.5, false)
            # any other expression is formula text, held with a leading `=`
            @test XLSX.parse_defined_name_value("SomethingElse") == ("=SomethingElse", false)
        end
    end

    @testset "invalid, duplicate and R1C1-style names are rejected" begin
        xf = XLSX.newxlsx()
        sh = xf[1]
        wb = XLSX.get_workbook(xf)
        @test_throws XLSX.XLSXError XLSX.get_defined_name_value(wb, "missing_name")
        @test_throws XLSX.XLSXError XLSX.get_defined_name_value(sh, "missing_name")
        @test !XLSX.is_valid_defined_name("")
        @test !XLSX.is_valid_defined_name("1abc")
        @test !XLSX.is_valid_defined_name("a-b")
        # Excel refuses a file with an R1C1-style defined name, or one that is just R or C,
        # in any case. (A1-style names such as R5 were already rejected.)
        for n in ("R1C1", "RC", "rc", "R12C", "RC3", "R", "C", "r", "c")
            @test !XLSX.is_valid_defined_name(n)
        end
        for n in ("XFE1", "A0", "A1B", "RCX", "Rate", "Col")
            @test XLSX.is_valid_defined_name(n)
        end
        err = try XLSX.addDefinedName(xf, "R1C1", 1); nothing catch e; e end
        @test err isa XLSX.XLSXError && occursin("cell address", err.msg)
        @test_throws XLSX.XLSXError XLSX.addDefinedName(sh, "rc", 1)
        @test_throws XLSX.XLSXError XLSX.addDefName(xf, "", 1)

        XLSX.addDefinedName(xf, "dup", 1)
        @test_throws XLSX.XLSXError XLSX.addDefinedName(xf, "dup", 2)
        @test_throws XLSX.XLSXError XLSX.addDefinedName(xf, "empty", "")
        @test_throws XLSX.XLSXError XLSX.addDefinedName(sh, "empty", "")
        @test_throws XLSX.XLSXError XLSX.addDefinedName(xf, "nosheet", "A1")
        XLSX.addDefinedName(xf, "wb_nc", "Sheet1!A1,Sheet1!B2")
        @test XLSX.get_defined_name_value(wb, "wb_nc") isa XLSX.NonContiguousRange
        SAVE_FILES && save_outfile(xf)
    end

    @testset "Integer values, hashing and show of defined names" begin
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
end
@testset "system defined names" begin

    @testset "hidden survives a round trip" begin
        f = XLSX.opentemplate(joinpath(data_directory, "chart_ex.xlsx"))
        before = XLSX.getAllDefinedNames(f; include_system=true)
        @test length(before) == 9
        @test all(dn -> dn.hidden, before)

        g = XLSX.writexlsx("myfile.xlsx", f, overwrite=true)
        h = XLSX.readxlsx(g)
        after = XLSX.getAllDefinedNames(h; include_system=true)
        isfile("myfile.xlsx") && rm("myfile.xlsx")

        @test length(after) == 9
        @test all(dn -> dn.hidden, after)
        @test Set(dn.name for dn in after) == Set(dn.name for dn in before)
        # values survive too, not just the flag
        @test Dict(dn.name => string(dn.value) for dn in after) ==
            Dict(dn.name => string(dn.value) for dn in before)
        SAVE_FILES && save_outfile(f)
    end

    @testset "hidden from the user-facing accessors" begin
        f = XLSX.readxlsx(joinpath(data_directory, "chart_ex.xlsx"))
        @test isempty(XLSX.getDefinedNames(f))
        @test length(XLSX.getDefinedNames(f; include_system=true)) == 9
        @test isempty(XLSX.getAllDefinedNames(f))
        @test length(XLSX.getAllDefinedNames(f; include_system=true)) == 9
        # sheet scope is unaffected — these are workbook-scoped
        @test isempty(XLSX.getDefinedNames(f["Data"]; include_system=true))
    end

    @testset "internal lookups still see system names" begin
        f = XLSX.readxlsx(joinpath(data_directory, "chart_ex.xlsx"))
        @test XLSX.is_workbook_defined_name(f, "_xlchart.v1.0")
        @test XLSX.get_defined_name_value(XLSX.get_workbook(f), "_xlchart.v1.0") isa XLSX.SheetCellRange
    end

    @testset "system names are protected from deletion" begin
        f = XLSX.opentemplate(joinpath(data_directory, "chart_ex.xlsx"))
        @test_throws XLSX.XLSXError XLSX.deleteDefinedName(f, "_xlchart.v1.0")
        @test XLSX.is_workbook_defined_name(f, "_xlchart.v1.0")   # not deleted

        XLSX.deleteDefinedName(f, "_xlchart.v1.0"; force=true)
        @test !XLSX.is_workbook_defined_name(f, "_xlchart.v1.0")
        SAVE_FILES && save_outfile(f)
    end

    @testset "the delete-everything idiom spares system names" begin
        f = XLSX.opentemplate(joinpath(data_directory, "chart_ex.xlsx"))
        XLSX.addDefinedName(f, "MY_NAME", "Data!A1:A5")

        XLSX.deleteDefinedName(f, XLSX.getDefinedNames(f))
        @test isempty(XLSX.getDefinedNames(f))
        @test length(XLSX.getAllDefinedNames(f; include_system=true)) == 9

        XLSX.deleteAllDefinedNames(f)
        @test length(XLSX.getAllDefinedNames(f; include_system=true)) == 9

        XLSX.deleteAllDefinedNames(f; force=true)
        @test isempty(XLSX.getAllDefinedNames(f; include_system=true))
        SAVE_FILES && save_outfile(f)
    end

    @testset "the guard is all-or-nothing" begin
        f = XLSX.opentemplate(joinpath(data_directory, "chart_ex.xlsx"))
        XLSX.addDefinedName(f, "MY_NAME", "Data!A1:A5")
        @test_throws XLSX.XLSXError XLSX.deleteDefinedName(f, ["MY_NAME", "_xlchart.v1.0"])
        @test XLSX.is_workbook_defined_name(f, "MY_NAME")   # validation precedes deletion
        SAVE_FILES && save_outfile(f)
    end

    @testset "hidden survives a sheet rename" begin
        f = XLSX.opentemplate(joinpath(data_directory, "chart_ex.xlsx"))
        XLSX.renamesheet!(f["Data"], "Renamed")
        dns = XLSX.getAllDefinedNames(f; include_system=true)
        @test length(dns) == 9
        @test all(dn -> dn.hidden, dns)
        @test all(dn -> startswith(string(dn.value), "Renamed!"), dns)
        SAVE_FILES && save_outfile(f)
    end

    @testset "_xlnm built-ins are system names" begin
        f = XLSX.opentemplate(joinpath(data_directory, "customXml.xlsx"))
        ws = f["Mock-up"]

        # Workbook scope: four user names, no built-ins.
        wb_names = XLSX.getDefinedNames(f)
        @test Set(dn.name for dn in wb_names) ==
            Set(["Contiguous", "Location", "Short_Description", "ID"])
        @test length(XLSX.getDefinedNames(f; include_system=true)) == 4

        # Worksheet scope: only the built-in, and only when asked for.
        @test isempty(XLSX.getDefinedNames(ws))
        sheet_names = XLSX.getDefinedNames(ws; include_system=true)
        @test length(sheet_names) == 1
        @test sheet_names[1].name == "_xlnm.Print_Area"
        @test sheet_names[1].hidden == false      # prefix, not the hidden flag

        # Across scopes.
        @test length(XLSX.getAllDefinedNames(f)) == 4
        @test length(XLSX.getAllDefinedNames(f; include_system=true)) == 5

        # Protected from deletion at worksheet scope.
        @test_throws XLSX.XLSXError XLSX.deleteDefinedName(ws, "_xlnm.Print_Area")
        @test XLSX.is_worksheet_defined_name(ws, "_xlnm.Print_Area")

        # User names delete normally, built-in survives the idiom.
        XLSX.deleteDefinedName(f, XLSX.getDefinedNames(f))
        @test isempty(XLSX.getDefinedNames(f))
        @test length(XLSX.getAllDefinedNames(f; include_system=true)) == 1

        XLSX.deleteAllDefinedNames(f)
        @test length(XLSX.getAllDefinedNames(f; include_system=true)) == 1
        XLSX.deleteAllDefinedNames(f; force=true)
        @test isempty(XLSX.getAllDefinedNames(f; include_system=true))

        @test_throws XLSX.XLSXError XLSX.addDefinedName(f["Data"], "_xlnm.Print_Area", "A1:B2")
        err = @test_throws XLSX.XLSXError XLSX.addDefinedName(f, "_xlchart.v1.99", "Data!A1:A5")
        @test occursin("reserved", err.value.msg)
        SAVE_FILES && save_outfile(f)

    end

    @testset "_xlnm survives a round trip without gaining hidden" begin
        f = XLSX.opentemplate(joinpath(data_directory, "customXml.xlsx"))
        g = XLSX.writexlsx("mytest.xlsx", f, overwrite=true)
        h = XLSX.readxlsx(g)
        dns = XLSX.getAllDefinedNames(h; include_system=true)
        @test length(dns) == 5
        pa = only(filter(dn -> dn.name == "_xlnm.Print_Area", dns))
        @test pa.hidden == false
        @test pa.scope == "Mock-up"          # localSheetId round-tripped
        @test string(pa.value) == "'Mock-up'!A1:K116"
        isfile("mytest.xlsx") && rm("mytest.xlsx")
        SAVE_FILES && save_outfile(f)
    end
end

@testset "review fixes: defined names" begin
    @testset "string constants and formulas" begin
        f = XLSX.newxlsx("S")
        XLSX.addDefinedName(f, "greeting", "Say \"hi\"")
        XLSX.addDefinedName(f, "rows", "=COUNTA(S!\$A:\$A)")
        XLSX.update_workbook_xml!(f)
        texts = Dict(n["name"] => XLSX.text_value(n) for n in XLSX.elements_with_tag(XLSX.first_element_with_tag(XLSX.xml_root_element(XLSX.xmlroot(f, "xl/workbook.xml")), "definedNames"), "definedName"))
        @test texts["greeting"] == "\"Say \"\"hi\"\"\""
        @test texts["rows"] == "COUNTA(S!\$A:\$A)"
        XLSX.writexlsx("dn_consts.xlsx", f; overwrite=true)
        f2 = XLSX.readxlsx("dn_consts.xlsx")
        @test f2["greeting"] == "Say \"hi\""
        @test XLSX.get_defined_name_value(f2.workbook, "rows") == "=COUNTA(S!\$A:\$A)"
        @test XLSX.parse_defined_name_value("\"a\"&\"b\"") == ("=\"a\"&\"b\"", false)
        SAVE_FILES && save_outfile("dn_consts.xlsx")
    end
    @testset "names on sheets whose names need quoting" begin
        f = XLSX.newxlsx("2024"); f[1]["A1"] = 42
        a1 = XLSX.addsheet!(f, "A1"); a1["A1"] = 7
        XLSX.addDefinedName(f, "fromDigit", "'2024'!A1")
        XLSX.addDefinedName(f, "fromA1", "'A1'!A1")
        XLSX.update_workbook_xml!(f)
        texts = Dict(n["name"] => XLSX.text_value(n) for n in XLSX.elements_with_tag(XLSX.first_element_with_tag(XLSX.xml_root_element(XLSX.xmlroot(f, "xl/workbook.xml")), "definedNames"), "definedName"))
        @test texts["fromDigit"] == "'2024'!\$A\$1"
        @test texts["fromA1"] == "'A1'!\$A\$1"
        SAVE_FILES && save_outfile(f)
    end
    @testset "names that are cell addresses in any case are refused" begin
        @test !XLSX.is_valid_defined_name("a1")
        @test !XLSX.is_valid_defined_name("xfd100")
    end
    @testset "a non-contiguous name recreated from getDefinedNames" begin
        f = XLSX.newxlsx("S")
        XLSX.addDefinedName(f, "nc", "S!A1,S!B2"; absolute=false)
        dn = only(XLSX.getDefinedNames(f))
        XLSX.deleteDefinedName(f, "nc")
        XLSX.addDefinedName(f, dn.name, dn.value; absolute=dn.absolute)
        @test only(XLSX.getDefinedNames(f)) == dn
        @test_throws XLSX.XLSXError XLSX.addDefinedName(f, "nc2", dn.value; absolute=[true])
        SAVE_FILES && save_outfile(f)
    end
    @testset "copysheet! repoints the copy's names" begin
        f = XLSX.newxlsx("S")
        XLSX.addDefinedName(f[1], "cell", "B2")
        XLSX.addDefinedName(f[1], "nc", "A1,B2")
        t = XLSX.copysheet!(f[1], "T")
        @test Dict(dn.name => dn.value for dn in XLSX.getDefinedNames(t)) ==
              Dict("cell" => XLSX.SheetCellRef("T!B2"), "nc" => XLSX.NonContiguousRange("T!A1,T!B2"))
        SAVE_FILES && save_outfile(f)
    end
    @testset "deletesheet! sets names that refer to it to #REF!" begin
        f = XLSX.newxlsx("S1"); s2 = XLSX.addsheet!(f, "S2")
        XLSX.addDefinedName(s2, "pt", XLSX.SheetCellRef("S1!A1"))
        XLSX.addDefinedName(f, "wb", "S1!A1:B2")
        XLSX.addDefinedName(f, "keep", "S2!A1")
        XLSX.addDefinedName(f[1], "own", "A1")
        # addDefinedName refuses _xlchart names, so add one directly
        wb = XLSX.get_workbook(f)
        wb.workbook_names["_xlchart.v1.0"] = XLSX.DefinedNameValue(XLSX.SheetCellRange("S1!A1:A3"), true, true)
        XLSX.deletesheet!(f, "S1")
        @test only(XLSX.getDefinedNames(s2)).value == "=#REF!"
        @test Dict(dn.name => dn.value for dn in XLSX.getDefinedNames(f)) ==
              Dict("wb" => "=#REF!", "keep" => XLSX.SheetCellRef("S2!A1"))
        @test !haskey(wb.workbook_names, "_xlchart.v1.0")
        @test !any(k -> last(k) == "own", keys(wb.worksheet_names))

        mktempdir() do dir
            p = joinpath(dir, "deleted.xlsx")
            XLSX.writexlsx(p, f)
            g = XLSX.readxlsx(p)
            @test only(XLSX.getDefinedNames(g[1])).value == "=#REF!"
            @test Dict(dn.name => dn.value for dn in XLSX.getDefinedNames(g))["wb"] == "=#REF!"
        end
        SAVE_FILES && save_outfile(f)
    end
    @testset "quoteit" begin
        @test XLSX.quoteit("2024") == "'2024'"
        @test XLSX.quoteit("A1") == "'A1'"
        @test XLSX.quoteit("rc") == "'rc'"
        @test XLSX.quoteit("Sheet1") == "Sheet1"
    end
end
