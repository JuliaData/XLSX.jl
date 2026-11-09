@testset "filemodes" begin

    sheetname = "New Sheet"
    filename = "test_file.xlsx"
    isfile(filename) && rm(filename)

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

    @testset "openxlsx / parse_file_mode - argument errors" begin
        missing_file = "definitely_not_a_file_12345.xlsx"

        @test_throws XLSX.XLSXError XLSX.openxlsx(missing_file)
        @test_throws XLSX.XLSXError XLSX.openxlsx(identity, missing_file)
        @test_throws XLSX.XLSXError XLSX.openxlsx(missing_file; mode="q")
        @test XLSX.parse_file_mode("wr") == (true, true)
        @test XLSX.parse_file_mode("RW") == (true, true)
    end

    @testset "missing file or bad `mode`" begin
        @test_throws XLSX.XLSXError XLSX.openxlsx("noSuchFile.xlsx")
        @test_throws XLSX.XLSXError XLSX.openxlsx(joinpath(data_directory, "Book1.xlsx"); mode="tg")
    end

    @testset "write-only mode" begin
        XLSX.openxlsx("mytest.xlsx", mode="w") do f
            f[1]["A1"] = 1
            @test f.source == "mytest.xlsx"
        end
        SAVE_FILES && save_outfile("mytest.xlsx")
        ef = XLSX.readxlsx("mytest.xlsx")
        @test ef["Sheet1"]["A1"] == 1
        f = XLSX.openxlsx("mytest2.xlsx", mode="w")
        @test f.source == "mytest2.xlsx"
        f[1]["A1"] = 1
        XLSX.writexlsx("mytest3.xlsx", f, overwrite=true)
        SAVE_FILES && save_outfile("mytest3.xlsx")
        ef = XLSX.readxlsx("mytest3.xlsx")
        @test ef["Sheet1"]["A1"] == 1
        f = XLSX.newxlsx()
        @test f.source == "blank.xlsx"
        f[1]["A1"] = 1
        XLSX.writexlsx("mytest.xlsx", f, overwrite=true)
        SAVE_FILES && save_outfile("mytest.xlsx")
        ef = XLSX.readxlsx("mytest.xlsx")
        @test ef["Sheet1"]["A1"] == 1
        for f in ["mytest.xlsx", "mytest2.xlsx", "mytest3.xlsx"]
            isfile(f) && rm(f)
        end
    end

    @testset "read-only and no-cache files reject formula and merge calls" begin
        general = joinpath(data_directory, "general.xlsx")
        
        ro = XLSX.readxlsx(general)
        s = ro[1]
        @test_throws XLSX.XLSXError XLSX.setFormula(s, "A1", "=1+1")
        @test_throws XLSX.XLSXError XLSX.setFormula(s, "A1:A2", "=1+1")
        @test_throws XLSX.XLSXError XLSX.isMergedCell(s, "A1")
        @test_throws XLSX.XLSXError XLSX.getMergedBaseCell(s, "A1")
        @test_throws XLSX.XLSXError XLSX.mergeCells(s, "A1:B2")

        XLSX.openxlsx(general; enable_cache=false) do nc
            s = nc[1]
            @test_throws XLSX.XLSXError XLSX.getMergedCells(s)
        end
    end
end
