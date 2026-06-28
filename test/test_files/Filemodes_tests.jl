@testset "filemodes" begin

    sheetname = "New Sheet"
    filename = "test_file.xlsx"
    if isfile(filename)
        rm(filename)
    end

    data = [
        1 "a" Date(2018, 1, 1);
        2 missing Date(2018, 1, 2);
        missing "c" Date(2018, 1, 3)
    ]

    # can't read or edit a file that does not exist
    @test_throws XLSX.XLSXError XLSX.openxlsx(filename, mode="r") do xf
        error("This should fail.")
    end

    @test_throws XLSX.XLSXError XLSX.openxlsx(filename, mode="rw") do xf
        error("This should fail.")
    end

    # test create new file
    XLSX.openxlsx(filename, mode="w") do xf
        sheet = xf[1]
        XLSX.renamesheet!(sheet, sheetname)

        sheet["A1"] = data[1, :]
        sheet[2, :] = data[2, :]
        sheet[2, 1] = "test overwrite"
        sheet[3, 2:3] = data[3, 2:3]
    end
    SAVE_FILES && save_outfile(filename)

    @test isfile(filename)
    XLSX.openxlsx(filename) do xf
        sheet = xf[sheetname]
        read_data = sheet[:]

        @test isequal(read_data[1, :], data[1, :])
        @test isequal(read_data[2, :], vcat(["test overwrite"], data[2, 2:end]))
        @test isequal(read_data[3, :], data[3, :])
    end

    # test overwrite file
    @test isfile(filename)
    new_data = [1 2 3;]
    XLSX.openxlsx(filename, mode="w") do xf
        sheet = xf[1]
        sheet[1, :] = new_data[1, :]
    end
    SAVE_FILES && save_outfile(filename)

    XLSX.openxlsx(filename) do xf
        sheet = xf[1]
        read_data = sheet[:]

        @test isequal(read_data, new_data)
    end

    # test edit file
    XLSX.openxlsx(filename, mode="rw") do xf
        sheet = xf[1]
        sheet[1, 2] = "hello"
        sheet["B6"] = 5
    end
    SAVE_FILES && save_outfile(filename)

    XLSX.openxlsx(filename) do xf
        sheet = xf[1]
        read_data = sheet[:]

        @test read_data[1, 1] == new_data[1, 1]
        @test read_data[1, 2] == "hello"
        @test read_data[1, 3] == new_data[1, 3]
        @test read_data[6, 2] == 5
    end

    # test writing throws error if flag not set
    XLSX.openxlsx(filename) do xf
        sheet = xf[1]
        @test_throws XLSX.XLSXError sheet[1, 1] = "failure"
    end

    @test_throws XLSX.XLSXError f = XLSX.openxlsx(filename; mode="rw", enable_cache=false) # Cache must be enabled to open in `write` mode.

    @testset "write column" begin
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
        test_data = [1 2 3; 4 5 6; 7 8 9]
        XLSX.openxlsx(filename, mode="w") do xf
            sheet = xf[1]
            @test_throws XLSX.XLSXError sheet["A7:C10"] = test_data
        end
        SAVE_FILES && save_outfile(filename)
    end

    @testset "write matrix with heterogeneous data types" begin
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

    @testset "doctest for writetable!" begin
        columns = Vector()
        push!(columns, [1, 2, 3])
        push!(columns, ["a", "b", "c"])

        labels = ["column_1", "column_2"]

        XLSX.openxlsx(filename, mode="w") do xf
            sheet = xf[1]
            XLSX.writetable!(sheet, columns, labels, anchor_cell=XLSX.CellRef("B2"))
        end
        SAVE_FILES && save_outfile(filename)

        # read data back
        XLSX.openxlsx(filename) do xf
            sheet = xf[1]
            @test sheet["B2"] == "column_1"
            @test sheet["C2"] == "column_2"
            @test sheet["B3"] == 1
            @test sheet["B4"] == 2
            @test sheet["B5"] == 3
            @test sheet["C3"] == "a"
            @test sheet["C4"] == "b"
            @test sheet["C5"] == "c"
        end
    end

    @testset "openxlsx without do-syntax" begin
        let
            xf = XLSX.openxlsx(filename)
            sheet = xf[1]
            @test sheet["B2"] == "column_1"
        end

        let
            xf = XLSX.openxlsx(filename, mode="w")
            sheet = xf[1]
            sheet["A1"] = "openxlsx without do-syntax"
            XLSX.writexlsx(filename, xf, overwrite=true)
            SAVE_FILES && save_outfile(filename)
        end

        let
            xf = XLSX.openxlsx(filename)
            sheet = xf[1]
            @test sheet["A1"] == "openxlsx without do-syntax"
        end
    end

    isfile(filename) && rm(filename)
end
