@testset "Defined Names" begin # Issue #148 
    @test XLSX.is_defined_name_value_a_reference(XLSX.SheetCellRef("Sheet1!A1"))
    @test XLSX.is_defined_name_value_a_reference(XLSX.SheetCellRange("Sheet1!A1:B2"))
    @test !XLSX.is_defined_name_value_a_reference(1)
    @test !XLSX.is_defined_name_value_a_reference(1.2)
    @test !XLSX.is_defined_name_value_a_reference("Hey")
    @test !XLSX.is_defined_name_value_a_reference(missing)

    f = XLSX.opentemplate(joinpath(data_directory, "general.xlsx"))

    result = XLSX.getDefinedNames(f)

    # Return type
    @test eltype(result) <: NamedTuple
    @test length(result) == 16
    @test all(r -> haskey(r, :name) && haskey(r, :scope) && haskey(r, :value), result)
    @test all(r -> r.name isa String && r.scope isa String && r.value isa String, result)

    # Sorting: sorted by (scope, name) — "Workbook" before sheet names
    @test issorted(result, by = x -> (x.scope, x.name))
    workbook_end = findlast(r -> r.scope == "Workbook", result)
    sheet_start  = findfirst(r -> r.scope != "Workbook", result)
    @test workbook_end < sheet_start

    # Workbook-scoped constants
    @test any(r -> r.name == "CONST_INT"   && r.scope == "Workbook" && r.value == "100",   result)
    @test any(r -> r.name == "CONST_FLOAT" && r.scope == "Workbook" && r.value == "10.2",  result)
    @test any(r -> r.name == "CONST_DATE"  && r.scope == "Workbook" && r.value == "43383", result)

    # Workbook-scoped ranges
    @test any(r -> r.name == "SINGLE_CELL" && r.scope == "Workbook" && r.value == "named_ranges!A2",    result)
    @test any(r -> r.name == "RANGE_B4C5"  && r.scope == "Workbook" && r.value == "named_ranges!B4:C5", result)

    # Workbook-scoped string value
    @test any(r -> r.name == "LOCAL_NAME" && r.scope == "Workbook" && r.value == "out there in the cold", result)

    # Worksheet-scoped: named_ranges sheet
    @test any(r -> r.name == "LOCAL_INT"      && r.scope == "named_ranges" && r.value == "1000",              result)
    @test any(r -> r.name == "LOCAL_NAME"     && r.scope == "named_ranges" && r.value == "Hey You",           result)
    @test any(r -> r.name == "LOCAL_REF"      && r.scope == "named_ranges" && r.value == "named_ranges!A15:B15", result)
    @test any(r -> r.name == "CONST_LOCAL_INT"&& r.scope == "named_ranges" && r.value == "100",               result)

    # Worksheet-scoped: named_ranges_2 sheet
    @test any(r -> r.name == "LOCAL_INT"  && r.scope == "named_ranges_2" && r.value == "2000",                  result)
    @test any(r -> r.name == "LOCAL_NAME" && r.scope == "named_ranges_2" && r.value == "out there in the cold", result)
    @test any(r -> r.name == "LOCAL_REF"  && r.scope == "named_ranges_2" && r.value == "named_ranges_2!D1:E1",  result)

    # Names that exist at both workbook and worksheet scope (shadowing)
    local_int_entries = filter(r -> r.name == "LOCAL_INT", result)
    @test length(local_int_entries) == 3
    @test any(r -> r.scope == "Workbook",       local_int_entries)
    @test any(r -> r.scope == "named_ranges",   local_int_entries)
    @test any(r -> r.scope == "named_ranges_2", local_int_entries)

    const_local_int_entries = filter(r -> r.name == "CONST_LOCAL_INT", result)
    @test length(const_local_int_entries) == 2
    @test any(r -> r.scope == "Workbook",     const_local_int_entries)
    @test any(r -> r.scope == "named_ranges", const_local_int_entries)

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

end
