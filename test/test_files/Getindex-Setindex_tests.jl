@testset "getindex" begin
    f = XLSX.readxlsx(joinpath(data_directory, "Book1.xlsx"))
    show(IOBuffer(), f)
    sheet1 = f["Sheet1"]
    show(IOBuffer(), sheet1)
    @test sheet1["B2"] == "B2"
    @test isapprox(sheet1["C3"], 21.2)
    @test sheet1["B5"] == Date(2018, 3, 21)
    @test sheet1["B8"] == "palavra1"

    @test XLSX.getcell(sheet1, "B2") == XLSX.Cell(XLSX.get_workbook(f), XLSX.CellRef("B2"), "s", "", "0", "", false)
    XLSX.getcell(sheet1, "B:C")
    XLSX.getcell(sheet1, "1:2")
    XLSX.getcell(sheet1, 1:2, 1:2)
    XLSX.getcellrange(sheet1, "B2:C3")
    XLSX.getcellrange(f, "Sheet1!B2:C3")
    XLSX.getcellrange(f, "Sheet1!B:C")
    XLSX.getcellrange(f, "Sheet1!2:3")
    XLSX.getcellrange(f, "Sheet1!B2,Sheet1!C3")
    XLSX.getcellrange(sheet1, 2, 2)
    XLSX.getcellrange(sheet1, 2, :)
    XLSX.getcellrange(sheet1, :, 3)
    XLSX.getcellrange(sheet1, 3, :)
    XLSX.getcellrange(sheet1, "B2:C3")
    XLSX.getcellrange(sheet1, "B2,C3:C4")
    @test_throws XLSX.XLSXError XLSX.getcellrange(f, "B2:C3")

    # a cell can be put in a dict
    c = XLSX.getcell(sheet1, "B2")
    show(IOBuffer(), c)
    dct = Dict("a" => c)
    @test dct["a"] == XLSX.Cell(XLSX.get_workbook(f), XLSX.CellRef("B2"), "s", "", "0", "", false)

    # equality and hash
    @test XLSX.Cell(XLSX.get_workbook(f), XLSX.CellRef("B2"), "s", "", "0", "", false) == XLSX.Cell(XLSX.get_workbook(f), XLSX.CellRef("B2"), "s", "", "0", "", false)
    @test hash(dct["a"]) == hash(XLSX.Cell(XLSX.get_workbook(f), XLSX.CellRef("B2"), "s", "", "0", "", false))

    sheet2 = f[2]
    sheet2_data = [1 2 3; 4 5 6; 7 8 9]
    @test sheet2_data == sheet2["A1:C3"]
    @test sheet2_data == sheet2[:]
    @test sheet2[:] == XLSX.getdata(sheet2)
    @test sheet2[:] == XLSX.getdata(sheet2, :)
    @test XLSX.getdata(sheet2, :, [1, 2]) == sheet2["A1:B3"]

    f = XLSX.readxlsx(joinpath(data_directory, "Book1.xlsx"))
    sheet1 = f["Sheet1"]
    XLSX.getcellrange(sheet1, "B2:C3")
    f = XLSX.openxlsx(joinpath(data_directory, "Book1.xlsx"); enable_cache=false)
    sheet1 = f["Sheet1"]
    XLSX.getcellrange(sheet1, "B2:C3")
end

@testset "setindex" begin
    f = XLSX.newxlsx()
    s = f[1]
    s["A1:A3"] = "Hello world"
    s[2, 1:3] = 42
    s[[1, 3], 2:3] = true
    @test s[1:3, [1, 2, 3]] == Any["Hello world" true true; 42 42 42; "Hello world" true true]
    s[2, :] = 44
    @test s[[1, 2, 3], 1:3] == Any["Hello world" true true; 44 44 44; "Hello world" true true]
    @test s["Sheet1!A1:C3"] == Any["Hello world" true true; 44 44 44; "Hello world" true true]
    @test s["Sheet1!A:C"] == Any["Hello world" true true; 44 44 44; "Hello world" true true]
    @test s["Sheet1!1:3"] == Any["Hello world" true true; 44 44 44; "Hello world" true true]
    s[:, :] = 0
    @test s[:, :] == Any[0 0 0; 0 0 0; 0 0 0]
    s[:] = 1
    @test s[:, 1:3] == Any[1 1 1; 1 1 1; 1 1 1]
    @test s[1:3, :] == Any[1 1 1; 1 1 1; 1 1 1]
    @test s[1:2:3, :] == Any[1 1 1; 1 1 1]
    @test s[1:2:3, 1] == Any[1, 1]
    s["A1,B2,C3"] = "non-contiguous"
    @test s["Sheet1!A1,Sheet1!B2,Sheet1!C3"] == [["non-contiguous";;], ["non-contiguous";;], ["non-contiguous";;]]
    SAVE_FILES && save_outfile(f)

    f = XLSX.newxlsx()
    s = f[1]
    s["A1:A3"] = "Hello world"
    s["A2:C2"] = 42:44
    s[[1, 3], 2:3] = true
    @test s["A1:C3"] == Any["Hello world" true true; 42 43 44; "Hello world" true true]
    s["A2:C2"] = 44:2:48
    @test s[[1, 2, 3], 1:3] == Any["Hello world" true true; 44 46 48; "Hello world" true true]
    s["A2:C2"] = [98, 99, 100]
    s["A3:C3"] = ["goodbye World", "goodbye World", "goodbye World"]
    @test s[[1, 2, 3], 1:3] == Any["Hello world" true true; 98 99 100; "goodbye World" "goodbye World" "goodbye World"]
    SAVE_FILES && save_outfile(f)

    f = XLSX.newxlsx()
    s = f[1]
    s[[1, 2, 3], :] = "Hello world"
    s[:, [1, 2, 3, 4]] = 42
    s[:, 1:3] = true
    @test s["Sheet1!1:3"] == Any[true true true 42; true true true 42; true true true 42]
    s["Sheet1!A1"] = "Goodbye world"
    @test s["Sheet1!A1"] == "Goodbye world"
    s["Sheet1!A1:A3"] = "Goodbye cruel world"
    @test s["Sheet1!A1:A3"] == ["Goodbye cruel world"; "Goodbye cruel world"; "Goodbye cruel world";;]
    s["Sheet1!1:2"] = "Bright Lights"
    @test s["A1,B2,C3"] == [["Bright Lights";;], ["Bright Lights";;], [true;;]]
    s["Sheet1!C:D"] = "Beat my Retreat"
    @test s["B1,C2,D3"] == [["Bright Lights";;], ["Beat my Retreat";;], ["Beat my Retreat";;]]
    s["Sheet1!B1,Sheet1!C2,Sheet1!D3"] = "Night Comes In"
    @test s["Sheet1!B1,Sheet1!C2,Sheet1!D3"] == [["Night Comes In";;], ["Night Comes In";;], ["Night Comes In";;]]
    SAVE_FILES && save_outfile(f)

    f = XLSX.newxlsx()
    s = f[1]
    s[[1, 2, 3], :] = "Hello world"
    s[:, [1, 2, 3, 4]] = 42
    s[:, 1:3] = true
    @test f["Sheet1!1:3"] == Any[true true true 42; true true true 42; true true true 42]
    s["Sheet1!A1"] = "Goodbye world"
    @test f["Sheet1!A1"] == "Goodbye world"
    s["Sheet1!A1:A3"] = "Goodbye cruel world"
    @test s["Sheet1!A1:A3"] == ["Goodbye cruel world"; "Goodbye cruel world"; "Goodbye cruel world";;]
    s["Sheet1!1:2"] = "Bright Lights"
    @test s["A1,B2,C3"] == [["Bright Lights";;], ["Bright Lights";;], [true;;]]
    s["Sheet1!C:D"] = "Beat my Retreat"
    @test s["B1,C2,D3"] == [["Bright Lights";;], ["Beat my Retreat";;], ["Beat my Retreat";;]]
    s["Sheet1!B1,Sheet1!C2,Sheet1!D3"] = "Night Comes In"
    @test s["Sheet1!B1,Sheet1!C2,Sheet1!D3"] == [["Night Comes In";;], ["Night Comes In";;], ["Night Comes In";;]]
    @test_throws XLSX.XLSXError s["Sheet1!garbage"] = 1
    @test_throws XLSX.XLSXError s["garbage"] = 1
    @test_throws XLSX.XLSXError s["garbage1:garbage2"] = 1
    SAVE_FILES && save_outfile(f)


    f = XLSX.newxlsx()
    s = f[1]
    for i in 1:5
        for j in 1:5
            s[i, j] = i + j
        end
    end
    @test s[1:5, 1:5] == [2 3 4 5 6; 3 4 5 6 7; 4 5 6 7 8; 5 6 7 8 9; 6 7 8 9 10]
    s[1:3, 1:2:5] = 99
    @test s[1:5, 1:5] == [99 3 99 5 99; 99 4 99 6 99; 99 5 99 7 99; 5 6 7 8 9; 6 7 8 9 10]
    s[1:2:5, 4:5] = -99
    @test s[1:5, 1:5] == [99 3 99 -99 -99; 99 4 99 6 99; 99 5 99 -99 -99; 5 6 7 8 9; 6 7 8 -99 -99]
    s[[2, 4], [3, 5]] = 0
    @test s[1:5, 1:5] == [99 3 99 -99 -99; 99 4 0 6 0; 99 5 99 -99 -99; 5 6 0 8 0; 6 7 8 -99 -99]
    @test s[[2, 4], [3, 5]] == [0 0; 0 0]
    SAVE_FILES && save_outfile(f)

end