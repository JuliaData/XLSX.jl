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
    @test issorted(result, by = r -> (something(r.scope, ""), uppercase(r.name)))
    @test findlast(r -> isnothing(r.scope), result) < findfirst(r -> !isnothing(r.scope), result)

    # workbook scope — values come out typed, not as strings
    @test any(r -> r.name == "CONST_INT"   && isnothing(r.scope) && r.value == 100, result)
    @test any(r -> r.name == "CONST_FLOAT" && isnothing(r.scope) && r.value == 10.2, result)
    @test any(r -> r.name == "CONST_DATE"  && isnothing(r.scope) && r.value == 43383, result)
    @test any(r -> r.name == "SINGLE_CELL" && isnothing(r.scope) && r.value == XLSX.SheetCellRef("named_ranges!A2"), result)
    @test any(r -> r.name == "RANGE_B4C5"  && isnothing(r.scope) && r.value == XLSX.SheetCellRange("named_ranges!B4:C5"), result)
    @test any(r -> r.name == "LOCAL_NAME"  && isnothing(r.scope) && r.value == "out there in the cold", result)

    # worksheet scope
    @test any(r -> r.name == "LOCAL_INT"       && r.scope == "named_ranges" && r.value == 1000, result)
    @test any(r -> r.name == "LOCAL_NAME"      && r.scope == "named_ranges" && r.value == "Hey You", result)
    @test any(r -> r.name == "LOCAL_REF"       && r.scope == "named_ranges" && r.value == XLSX.SheetCellRange("named_ranges!A15:B15"), result)
    @test any(r -> r.name == "CONST_LOCAL_INT" && r.scope == "named_ranges" && r.value == 100, result)
    @test any(r -> r.name == "LOCAL_INT"       && r.scope == "named_ranges_2" && r.value == 2000, result)
    @test any(r -> r.name == "LOCAL_REF"       && r.scope == "named_ranges_2" && r.value == XLSX.SheetCellRange("named_ranges_2!D1:E1"), result)

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
    isfile("mytest.xlsx") && rm("mytest.xlsx")

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
    @test s["MyCell"] == true
    @test s["YourCells"] == Any[false false; false false]
    isfile("mytest.xlsx") && rm("mytest.xlsx")

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

    @test XLSX.hassheet(f, "Tim's Sheet")
    @test XLSX.hassheet(f, "Ano'ther She'et")
    @test f[2]["A1"] == "tim"
    @test f[3]["A1"] == "another"
    @test f["mine"] == "tim"
    @test f["yours"] == "another"
    @test string(XLSX.get_defined_name_value(XLSX.get_workbook(ff), "mine")) == "'Tim''s Sheet'!A1"
    @test string(XLSX.get_defined_name_value(XLSX.get_workbook(ff), "yours")) == "'Ano''ther She''et'!A1"

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
    end
end
