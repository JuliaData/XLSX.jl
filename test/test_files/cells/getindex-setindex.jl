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
    refs(cells) = map(c -> string(c.ref), cells)
    @test refs(XLSX.getcell(sheet1, "B:C")) == [string(c, r) for r in 2:8, c in "BC"]
    @test refs(XLSX.getcell(sheet1, "1:2")) == ["B1" "C1"; "B2" "C2"]
    @test refs(XLSX.getcell(sheet1, 1:2, 1:2)) == ["A1" "B1"; "A2" "B2"]
    @test refs(XLSX.getcellrange(sheet1, "B2:C3")) == ["B2" "C2"; "B3" "C3"]
    @test refs(XLSX.getcellrange(f, "Sheet1!B2:C3")) == ["B2" "C2"; "B3" "C3"]
    @test size(XLSX.getcellrange(f, "Sheet1!B:C")) == (7, 2)
    @test refs(XLSX.getcellrange(f, "Sheet1!2:3")) == ["B2" "C2"; "B3" "C3"]
    @test map(refs, XLSX.getcellrange(f, "Sheet1!B2,Sheet1!C3")) == [["B2";;], ["C3";;]]
    @test refs(XLSX.getcellrange(sheet1, 2, 2)) == ["B2";;]
    @test refs(XLSX.getcellrange(sheet1, 2, :)) == ["B2" "C2"]
    @test refs(XLSX.getcellrange(sheet1, :, 3)) == [string(c, r) for r in 2:8, c in "C"]
    @test refs(XLSX.getcellrange(sheet1, 3, :)) == ["B3" "C3"]
    @test map(refs, XLSX.getcellrange(sheet1, "B2,C3:C4")) == [["B2";;], ["C3"; "C4";;]]
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

    # getcellrange gives the same cells with and without the cache
    f = XLSX.readxlsx(joinpath(data_directory, "Book1.xlsx"))
    sheet1 = f["Sheet1"]
    @test [XLSX.getdata(sheet1, c) for c in XLSX.getcellrange(sheet1, "B2:C3")] == Any["B2" "C2"; 10.5 21.2]
    f = XLSX.openxlsx(joinpath(data_directory, "Book1.xlsx"); enable_cache=false)
    sheet1 = f["Sheet1"]
    @test [XLSX.getdata(sheet1, c) for c in XLSX.getcellrange(sheet1, "B2:C3")] == Any["B2" "C2"; 10.5 21.2]

    @testset "issue 460: attributes on <v>" begin
        xf = XLSX.readxlsx(joinpath(data_directory, "issue460_whitespace.xlsx"))
        sh = xf[1]
        @test sh["B4"] == " 6000056"          # t="str", <v xml:space="preserve">
        @test sh["B6"] == 6000056             # numeric, <v xml:space="preserve">
        @test all(!ismissing, sh["B2:B7"])
    end
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


    @testset "write column" begin
        filename = "setindex_forms.xlsx"
        col_data = collect(1:50)

        XLSX.openxlsx(filename, mode="w") do xf
            sheet = xf[1]
            sheet[:, 2] = col_data
            sheet[51:100, 3] = col_data
            sheet[2, 4, dim=1] = col_data
        end
        SAVE_FILES && save_outfile(filename)

        XLSX.openxlsx(filename) do xf
            sheet = xf[1]

            for (row, val) in enumerate(col_data)
                @test sheet[row, 2] == val
                @test sheet[50+row, 3] == val
                @test sheet[row+1, 4] == val
            end
        end
    end

    @testset "write matrix with anchor cell" begin
        filename = "setindex_forms.xlsx"
        test_data = [1 2 3; 4 5 6; 7 8 9]
        XLSX.openxlsx(filename, mode="w") do xf
            sheet = xf[1]
            sheet["A7"] = test_data
        end
        SAVE_FILES && save_outfile(filename)

        XLSX.openxlsx(filename) do xf
            sheet = xf[1]
            rows, cols = size(test_data)
            for c in 1:cols, r in 1:rows
                @test sheet[r+6, c] == test_data[r, c]
            end
        end

        xf = XLSX.newxlsx()
        sheet = xf[1]
        sheet[7, 1] = test_data
        XLSX.writexlsx(filename, xf, overwrite=true)
        SAVE_FILES && save_outfile(filename)

        XLSX.openxlsx(filename) do xf
            sheet = xf[1]
            rows, cols = size(test_data)
            for c in 1:cols, r in 1:rows
                @test sheet[r+6, c] == test_data[r, c]
            end
        end
    end

    @testset "write matrix with range" begin
        filename = "setindex_forms.xlsx"
        test_data = [1 2 3; 4 5 6; 7 8 9]
        XLSX.openxlsx(filename, mode="w") do xf
            sheet = xf[1]
            sheet["A7:C9"] = test_data
        end
        SAVE_FILES && save_outfile(filename)

        XLSX.openxlsx(filename) do xf
            sheet = xf[1]
            rows, cols = size(test_data)
            for c in 1:cols, r in 1:rows
                @test sheet[r+6, c] == test_data[r, c]
            end
        end

        xf = XLSX.newxlsx()
        sheet = xf[1]
        sheet[7:9, 1:3] = test_data
        XLSX.writexlsx(filename, xf, overwrite=true)
        SAVE_FILES && save_outfile(filename)

        XLSX.openxlsx(filename) do xf
            sheet = xf[1]
            rows, cols = size(test_data)
            for c in 1:cols, r in 1:rows
                @test sheet[r+6, c] == test_data[r, c]
            end
        end
    end

    @testset "write matrix with range mismatch" begin
        filename = "setindex_forms.xlsx"
        test_data = [1 2 3; 4 5 6; 7 8 9]
        XLSX.openxlsx(filename, mode="w") do xf
            sheet = xf[1]
            @test_throws XLSX.XLSXError sheet["A7:C10"] = test_data
        end
        SAVE_FILES && save_outfile(filename)
    end

    @testset "write matrix with heterogeneous data types" begin
        filename = "setindex_forms.xlsx"
        # issue #97
        test_data = ["A" "B"; 1 2; "a" "b"; 0 "x"]
        XLSX.openxlsx(filename, mode="w") do xf
            sheet = xf[1]
            sheet["B2"] = test_data
        end
        SAVE_FILES && save_outfile(filename)

        XLSX.openxlsx(filename, mode="r") do xf
            sheet = xf[1]
            @test sheet["B2"] == "A"
            @test sheet["C2"] == "B"
            @test sheet["B3"] == 1
            @test sheet["C3"] == 2
            @test sheet["B4"] == "a"
            @test sheet["C4"] == "b"
            @test sheet["B5"] == 0
            @test sheet["C5"] == "x"
        end
    end

    @testset "setdata! - references" begin
        xf = XLSX.newxlsx()
        sh = xf[1]
        sh[1:3, 1:3] = reshape(collect(1:9), 3, 3)

        XLSX.addDefName(sh, "ws_const", 5)
        XLSX.addDefName(xf, "wb_const", 6)
        @test_throws XLSX.XLSXError XLSX.setdata!(sh, "ws_const", 1)
        @test_throws XLSX.XLSXError XLSX.setdata!(sh, "wb_const", 1)
        @test_throws XLSX.XLSXError XLSX.setdata!(sh, "not a ref", 1)

        XLSX.setdata!(sh, "B:B", 0)
        @test all(==(0), sh["B1:B3"])
        XLSX.setdata!(sh, "3:3", 7)
        @test all(==(7), sh["A3:C3"])
        XLSX.setdata!(sh, "A1,C1", 9)
        @test sh["A1"] == 9 && sh["C1"] == 9

        @test_throws XLSX.XLSXError XLSX.setdata!(sh, "not a ref", [1 2; 3 4])
        @test_throws XLSX.XLSXError XLSX.setdata!(sh, "not a ref", [1, 2], 1)   # was an UndefVarError
        @test_throws XLSX.XLSXError XLSX.target_cell_ref_from_offset(1, 1, 1, 3)
        SAVE_FILES && save_outfile(xf)
    end

    @testset "setdata! - a new value restyles an existing cell" begin
        xf = XLSX.newxlsx()
        sh = xf[1]
        sh["A1"] = 1.5
        XLSX.setFormat(sh, "A1"; format="0.00")
        sh["A1"] = Dates.Time(12, 30)
        @test sh["A1"] == Dates.Time(12, 30)
        sh["A2"] = 1.5
        XLSX.setFormat(sh, "A2"; format="0.00")
        sh["A2"] = Dates.DateTime(2024, 1, 2, 3, 4)
        @test sh["A2"] == Dates.DateTime(2024, 1, 2, 3, 4)
        sh["A3"] = Dates.Date(2024, 1, 2)
        sh["A3"] = true
        @test sh["A3"] === true

        XLSX.setdata!(sh, XLSX.CellRef("B1"), XLSX.Formula("A2+1"))      # new cell
        @test XLSX.getcell(sh, "B1").formula
        sh["B2"] = 1
        XLSX.setFormat(sh, "B2"; format="0.00")
        XLSX.setdata!(sh, XLSX.CellRef("B2"), XLSX.Formula("A2+2"))      # styled cell
        @test XLSX.getcell(sh, "B2").formula
        SAVE_FILES && save_outfile(xf)
    end

    @testset "setdata! argument forms" begin
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
        SAVE_FILES && save_outfile(xf)
    end

    # A vector is written along one row or column, or across from one cell; anywhere
    # else it was an "ambiguous" MethodError, and is now an XLSXError saying so.
    @testset "vector values in ws[r, c] = v" begin
        xf = XLSX.newxlsx()
        sh = xf[1]
        sh["A1:D4"] = ""
        sh[2, 2] = [1, 2, 3]                       # across from B2, as sh["B2"] = v
        @test [sh[2, c] for c in 2:4] == [1, 2, 3]
        sh[3, 1:2] = [5, 6]
        @test [sh[3, 1], sh[3, 2]] == [5, 6]
        sh[:, 4] = [7, 8]
        @test [sh[1, 4], sh[2, 4]] == [7, 8]
        for (r, c) in ((1:2, 1:2), (1, [1, 3]), ([1, 3], 2), (1:2, :), (:, :), (1:2:3, 1:2:3))
            @test_throws XLSX.XLSXError (sh[r, c] = [1, 2])
        end
        sh[1, [1, 3]] = 9                          # any other value still fills each cell
        @test sh[1, 1] == 9 && sh[1, 3] == 9
    end

    # The cache keeps written rows unsorted until something needs their order. Reading a
    # range must sort them first: it once binary-searched them unsorted, returning wrong
    # values or writing out of bounds (and crashing Julia's GC later).
    @testset "getdata over a range after rows written out of order" begin
        xf = XLSX.newxlsx()
        sh = xf[1]
        for r in (5, 1, 9, 2, 3)
            sh[r, 1] = r
        end
        @test vec(XLSX.getdata(sh, "A2:A3")) == [2, 3]

        rng = Random.Xoshiro(1)
        for _ in 1:200
            xf = XLSX.newxlsx()
            sh = xf[1]
            rows = Random.shuffle(rng, 1:12)[1:rand(rng, 3:10)]
            for r in rows
                sh[r, 1] = r
            end
            top = rand(rng, 1:10)
            bottom = rand(rng, top:12)
            got = vec(XLSX.getdata(sh, XLSX.CellRange(top, 1, bottom, 1)))
            @test isequal(got, [r in rows ? r : missing for r in top:bottom])
        end
    end
end
