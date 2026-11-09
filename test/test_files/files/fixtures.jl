@testset "reading fixtures" begin
    ef_blank_ptbr_1904 = XLSX.readxlsx(joinpath(data_directory, "blank_ptbr_1904.xlsx"))
    ef_Book1 = XLSX.readxlsx(joinpath(data_directory, "Book1.xlsx"))
    ef_Book_1904 = XLSX.readxlsx(joinpath(data_directory, "Book_1904.xlsx"))
    ef_book_1904_ptbr = XLSX.readxlsx(joinpath(data_directory, "book_1904_ptbr.xlsx"))
    ef_book_sparse = XLSX.readxlsx(joinpath(data_directory, "book_sparse.xlsx"))
    ef_book_sparse_2 = XLSX.readxlsx(joinpath(data_directory, "book_sparse_2.xlsx"))

    @test XLSX.readxlsx(joinpath(data_directory, "missing_numFmtId.xlsx"))["Koldioxid (CO2)"][7, 5] ≈ 208634.871365389

    @test open(joinpath(data_directory, "blank_ptbr_1904.xlsx")) do io
        XLSX.readxlsx(io)
    end isa XLSX.XLSXFile

    @test ef_Book1.source == joinpath(data_directory, "Book1.xlsx")
    @test length(keys(ef_Book1.data)) > 0

    @test ef_Book_1904.source == joinpath(data_directory, "Book_1904.xlsx")
    @test length(keys(ef_Book_1904.data)) > 0

    @test !XLSX.isdate1904(ef_Book1)
    @test XLSX.isdate1904(ef_Book_1904)
    @test XLSX.isdate1904(ef_blank_ptbr_1904)
    @test XLSX.isdate1904(ef_book_1904_ptbr)

    @test XLSX.sheetnames(ef_Book1) == ["Sheet1", "Sheet2"]
    @test XLSX.sheetcount(ef_Book1) == 2
    @test ef_Book1["Sheet1"].name == "Sheet1"
    @test ef_Book1[1].name == "Sheet1"

    @test XLSX.sst_unformatted_string(ef_Book1.workbook, Int64(0)) == "B2" # index is 0-based
    @test XLSX.sst_unformatted_string(ef_Book1, Int64(0)) == "B2"
    @test XLSX.sst_unformatted_string(ef_Book1, "0") == "B2"

    @test !XLSX.has_relationship_by_type(ef_Book1.workbook, "invalid_type")

    @test XLSX.get_dimension(ef_Book1["Sheet1"]) == XLSX.range"B2:C8"
    @test XLSX.isdate1904(ef_Book1["Sheet1"]) == false

    @testset "show(Workbook)" begin
        xf = XLSX.readxlsx(joinpath(data_directory, "Book1.xlsx"))
        wb = XLSX.get_workbook(xf)
        s = sprint(show, wb)

        @test startswith(s, "Workbook(")
        @test occursin("Book1.xlsx", s)
        @test occursin("Worksheet", s)
        for sh in XLSX.sheetnames(xf)
            @test occursin(sh, s)
        end

        @test occursin(" containing 1 Worksheet\n", sprint(show, XLSX.get_workbook(XLSX.newxlsx())))
        
        io = IOBuffer(read(joinpath(data_directory, "Book1.xlsx")))
        @test occursin("Workbook(IOBuffer)", sprint(show, XLSX.get_workbook(XLSX.readxlsx(io))))

    end
    @testset "Read XLS file error" begin
        @test_throws XLSX.XLSXError XLSX.readxlsx(joinpath(data_directory, "old.xls"))
        try
            XLSX.readxlsx(joinpath(data_directory, "old.xls"))
            @test false # didn't throw exception
        catch e
            @test occursin("This package does not support XLS file format", "$e")
        end

    end

    # Issue #293
    @testset "read .xltx file" begin
        # saving a template writes the .xlsx beside it, so work on a copy
        cp(joinpath(data_directory, "Template File.xltx"), "Template File.xltx"; force=true)
        xf = XLSX.openxlsx("Template File.xltx"; mode="RW")
        s=xf[1]
        @test s["P5"] == 5
        @test XLSX.getFormula(s, "B5") == "=RANDBETWEEN(0,100)"
        @test xf.template_type == XLSX.XLTXTemplate

        XLSX.savexlsx(xf)
        SAVE_FILES && save_outfile("Template File.xlsx")
        @test isfile("Template File.xlsx")
        xf = XLSX.readxlsx("Template File.xlsx")
        s=xf[1]
        @test s["P5"] == 5
        @test XLSX.getFormula(s, "B5") == "=RANDBETWEEN(0,100)"
        @test xf.template_type == XLSX.NotATemplate
        isfile("Template File.xlsx") && rm("Template File.xlsx")

        XLSX.openxlsx("Template File.xltx"; mode="RW") do xf
            s=xf[1]
            @test s["P5"] == 5
            @test XLSX.getFormula(s, "B5") == "=RANDBETWEEN(0,100)"
            @test xf.template_type == XLSX.XLTXTemplate
        end
        @test isfile("Template File.xlsx")
        SAVE_FILES && save_outfile("Template File.xlsx")
        isfile("Template File.xlsx") && rm("Template File.xlsx")
    end

    # Issue #401
    @testset "macro enabled files" begin
        mf = XLSX.openxlsx(joinpath(data_directory, "macro-enabled.xlsm"); mode="Rw")
        @test mf[1]["A1"] == "hello"
        XLSX.writexlsx("mytest.xlsm", mf; overwrite=true)
        SAVE_FILES && save_outfile("mytest.xlsm")
        mf = XLSX.openxlsx("mytest.xlsm"; mode="rw")
        @test mf[1]["A1"] == "hello"
        isfile("mytest.xlsm") && rm("mytest.xlsm")

        SAVE_FILES && save_outfile(mf)
        cp(joinpath(data_directory, "macro-enabled2.xltm"), "macro-enabled2.xltm"; force=true)
        mf = XLSX.openxlsx("macro-enabled2.xltm"; mode="rW")
        @test mf[1]["A1"] == "hello"
        @test mf.template_type == XLSX.XLTMTemplate
        XLSX.savexlsx(mf)
        SAVE_FILES && save_outfile("macro-enabled2.xlsm")
        @test isfile("macro-enabled2.xlsm")
        mf = XLSX.readxlsx("macro-enabled2.xlsm")
        s=mf[1]
        @test mf[1]["A1"] == "hello"
        @test mf.template_type == XLSX.NotATemplate
        isfile("macro-enabled2.xlsm") && rm("macro-enabled2.xlsm")
        
        XLSX.openxlsx("macro-enabled2.xltm"; mode="wr") do mf
            @test mf[1]["A1"] == "hello"
            @test mf.template_type == XLSX.XLTMTemplate
        end
        @test isfile("macro-enabled2.xlsm")
        SAVE_FILES && save_outfile("macro-enabled2.xlsm")
        mf = XLSX.openxlsx("macro-enabled2.xlsm"; mode="WR")
        @test mf[1]["A1"] == "hello"
        @test mf.template_type == XLSX.NotATemplate
        isfile("macro-enabled2.xlsm") && rm("macro-enabled2.xlsm")

    end

    # Issue #403
    @testset "UTF-16 customXml" begin
        xf1 = XLSX.openxlsx(joinpath(data_directory, "UTF-16.xlsx"); mode="rw")
        XLSX.writexlsx("UTF-16_test.xlsx", xf1; overwrite=true)
        SAVE_FILES && save_outfile("UTF-16_test.xlsx")
        xf2 = XLSX.openxlsx("UTF-16_test.xlsx"; mode="rw")
        @test xf1[1]["E3"] == xf2[1]["E3"]
        @test xf1[1]["R99"] == xf2[1]["R99"]
        SAVE_FILES && save_outfile(xf2)

        isfile("UTF-16_test.xlsx") && rm("UTF-16_test.xlsx")

    end
    @testset "Read password protected file error" begin
        @test_throws XLSX.XLSXError XLSX.readxlsx(joinpath(data_directory, "password.xlsx")) # password for this file is simply "password"
        try
            XLSX.readxlsx(joinpath(data_directory, "password.xlsx"))
            @test false # didn't throw exception
        catch e
            @test occursin("This package does not support password protected files", "$e")
        end

    end

    @testset "Read invalid XLSX error" begin
        @test_throws XLSX.XLSXError XLSX.readxlsx(joinpath(src_data_directory, "sheet_template.xml"))
        try
            XLSX.readxlsx(joinpath(src_data_directory, "sheet_template.xml"))
            @test false # didn't throw exception
        catch e
            @test occursin("is not a valid XLSX file", "$e")
        end
    end

    @testset "Book1.xlsx" begin
        f = XLSX.readxlsx(joinpath(data_directory, "Book1.xlsx"))
        sheet = f["Sheet1"]
        @test ismissing(sheet["A1"])
        @test sheet["B2"] == "B2"
        @test sheet["C2"] == "C2"
        @test isapprox(sheet["B3"], 10.5)
        @test isapprox(sheet["C3"], 21.2)
        @test sheet["B4"] == Date(2018, 3, 21)
        @test sheet["C4"] == Date(2018, 3, 22)
        @test sheet["B5"] == Date(2018, 3, 21)
        @test sheet["C5"] == Date(2018, 3, 22)
        @test sheet["B6"] == true
        @test sheet["C6"] == false
        @test sheet["B7"] == 1
        @test sheet["C7"] == 2
        @test sheet["B8"] == "palavra1"
        @test sheet["C8"] == "palavra2"
        @test XLSX.get_dimension(sheet) == XLSX.CellRange("B2:C8")

        sheet2 = f["Sheet2"]
        @test XLSX.get_dimension(sheet2) == XLSX.CellRange("A1:C3")
        @test axes(sheet2, 1) == 1:3
        @test axes(sheet2, 2) == 1:3
        @test_throws ArgumentError axes(sheet2, 3)
        @test sheet2[1, :] == Any[1 2 3]
        @test sheet2[1:2, :] == Any[1 2 3; 4 5 6]
        @test sheet2[:, 2] == permutedims(Any[2 5 8])
        @test sheet2[:, 2:3] == Any[2 3; 5 6; 8 9]
        @test sheet2[1:2, 2:3] == Any[2 3; 5 6]

        @test XLSX.getdata(f, XLSX.SheetCellRef("Sheet1!B2")) == "B2"
        @test XLSX.getdata(f, XLSX.SheetCellRange("Sheet1!B2:B3"))[1] == "B2"
        @test XLSX.getdata(f, XLSX.SheetCellRange("Sheet1!B2:B3"))[2] == 10.5
        @test f["Sheet1!B2"] == "B2"
        @test f["Sheet1!B2:B3"][1] == "B2"
        @test f["Sheet1!B2:B3"][2] == 10.5
        @test string(XLSX.SheetCellRange("Sheet1!B2:B3")) == "Sheet1!B2:B3"
    end

    @testset "book_1904_ptbr.xlsx" begin
        f = XLSX.readxlsx(joinpath(data_directory, "book_1904_ptbr.xlsx"))

        @test f["Plan1"][:] == Any["Coluna A" "Coluna B" "Coluna C" "Coluna D";
            10 10.5 Date(2018, 3, 22) "linha 2";
            20 20.5 Date(2017, 12, 31) "linha 3";
            30 30.5 Date(2018, 1, 1) "linha 4"]

        @test f["Plan2"]["A1"] == "Merge de A1:D1"
        @test ismissing(f["Plan2"]["B1"])
        @test f["Plan2"]["C2"] == "C2"
        @test f["Plan2"]["D3"] == "D3"
        @test f["NEGOCIAÇÕES Descrição"]["A1"] == "Negociações"
        @test f["NEGOCIAÇÕES Descrição"]["B1"] == 10
        @test f["NEGOCIAÇÕES Descrição!A1"] == "Negociações"
        @test f["NEGOCIAÇÕES Descrição!B1"] == 10
    end

    @testset "numbers.xlsx" begin
        f = XLSX.readxlsx(joinpath(data_directory, "numbers.xlsx"))
        floats = f["float"][:]
        for n in floats
            if !ismissing(n)
                @test isa(n, Float64)
            end
        end

        ints = f["int"][:]
        for n in ints
            if !ismissing(n)
                @test isa(n, Int64)
            end
        end

        error_sheet = f["error"]
        @test error_sheet["A1"] == "errors"
        @test !XLSX.iserror(XLSX.getcell(error_sheet, "A1"))
        @test XLSX.iserror(XLSX.getcell(error_sheet, "A2"))
        @test XLSX.iserror(XLSX.getcell(f, "error!A2"))
        @test ismissing(error_sheet["A2"])
        @test ismissing(error_sheet["A3"])
        @test ismissing(error_sheet["A4"])
        emptycell = XLSX.getcell(error_sheet, "B1")
        @test !XLSX.iserror(emptycell)
        @test ismissing(XLSX.getdata(error_sheet, emptycell))
        @test XLSX.row_number(emptycell) == 1
        @test XLSX.column_number(emptycell) == 2
    end


    @testset "typed values in general.xlsx" begin
        XLSX.openxlsx(joinpath(data_directory, "general.xlsx")) do f
            show(IOBuffer(), f)
            sheet = f["general"]
            @test sheet["A1"] == "text"
            @test sheet["B1"] == "regular text"
            @test sheet["A2"] == "integer"
            @test sheet["B2"] == 102
            @test sheet["A3"] == "float"
            @test isapprox(sheet["B3"], 102.2)
            @test sheet["A4"] == "date"
            @test sheet["B4"] == Date(1983, 4, 16)
            @test sheet["A5"] == "hour"
            @test sheet["B5"] == Dates.Time(Dates.Hour(19), Dates.Minute(45))
            @test sheet["A6"] == "datetime"
            @test sheet["B6"] == Date(2018, 4, 16) + Dates.Time(Dates.Hour(19), Dates.Minute(19), Dates.Second(51))
            @test f["general!B7"] == -220.0
            @test f["general!B8"] == -2000
            @test f["general!B9"] == 100000000000000
            @test f["general!B10"] == -100000000000000
        end
    end
end

@testset "customXml parts" begin
    # issue #210
    template = XLSX.open_xlsx_template(joinpath(data_directory, "customXml.xlsx"))
    filename_copy = "customXml_copy.xlsx"
    for sn in XLSX.sheetnames(template)
        sheet = template[sn]
        sheet["Q1"] = "Can't"
        sheet["Q2"] = "write"
        sheet["Q3"] = "this"
        sheet["Q4"] = "template"
    end
    @test XLSX.writexlsx(filename_copy, template, overwrite=true) == abspath(filename_copy)
    SAVE_FILES && save_outfile(filename_copy)
    @test isfile(filename_copy)
    f_copy = XLSX.readxlsx(filename_copy)
    test_Xmlread = [["Can't", "write", "this", "template"]]
    for sn in XLSX.sheetnames(f_copy)
        sheet = f_copy[sn]
        data = [[sheet["Q1"], sheet["Q2"], sheet["Q3"], sheet["Q4"]]]
        check_test_data(data, test_Xmlread)
    end
    isfile(filename_copy) && rm(filename_copy)
end

@testset "show xlsx" begin
    @testset "single sheet" begin
        xf = XLSX.readxlsx(joinpath(src_data_directory, "blank.xlsx"))
        io = IOBuffer()
        show(io, xf)
        lines = split(String(take!(io)), '\n')

        expected_tail = "            sheetname size          range        \n-------------------------------------------------\n               Sheet1 1x1           A1:A1        \n"

        @test join(lines[2:end], '\n') == expected_tail
    end

    @testset "multiple sheets" begin
        xf = XLSX.readxlsx(joinpath(data_directory, "Book1.xlsx"))
        io = IOBuffer()
        show(io, xf)
        lines = split(String(take!(io)), '\n')

        expected_tail = "            sheetname size          range        \n-------------------------------------------------\n               Sheet1 7x2           B2:C8        \n               Sheet2 3x3           A1:C3        \n"

        @test join(lines[2:end], '\n') == expected_tail
    end

    @testset "IOBuffer source" begin
        source = IOBuffer()
        write(source, read(joinpath(data_directory, "Book1.xlsx")))
        seekstart(source)
        xf = XLSX.readxlsx(source)
        io=IOBuffer()
        show(io, xf)
        result = String(take!(io))
        expected = "XLSXFile(IOBuffer) containing 2 Worksheets\n            sheetname size          range        \n-------------------------------------------------\n               Sheet1 7x2           B2:C8        \n               Sheet2 3x3           A1:C3        \n"

        @test result == expected
    end
end

# issues #62, #75
@testset "relative paths" begin
    let
        xf = XLSX.readxlsx(joinpath(data_directory, "openpyxl.xlsx"))
        @test XLSX.sheetnames(xf) == ["Sheet", "Test1"]
        @test xf["Test1"]["A1"] == "One"
        @test xf["Test1"]["A2"] == 1
        show(IOBuffer(), xf)
        show(IOBuffer(), xf["Sheet"])
        show(IOBuffer(), xf["Test1"])
    end

    let
        dtable = XLSX.readtable(joinpath(data_directory, "openpyxl.xlsx"), "Test1")
        data, col_names = dtable.data, dtable.column_labels
        @test data == [[1, 3], [2, 4]]
        @test col_names == [:One, :Two]
    end

    # openpyxl writes directory entries ("xl/", ...) into the zip: a rw round trip
    # must skip them rather than fail writing them back
    let
        xf = XLSX.openxlsx(joinpath(data_directory, "openpyxl.xlsx"); mode="rw")
        @test !any(endswith("/"), keys(xf.binary_data))
        io = IOBuffer()
        XLSX.writexlsx(io, xf)
        xf2 = XLSX.readxlsx(IOBuffer(take!(io)))
        @test XLSX.sheetnames(xf2) == ["Sheet", "Test1"]
        @test xf2["Test1"]["A1"] == "One"
        @test xf2["Test1"]["A2"] == 1
    end
end

# issues #62, #71
@testset "windows compatibility" begin
    xf = XLSX.open_xlsx_template(joinpath(data_directory, "issue62_71.xlsx"))
    @test xf["Sheet1"]["A1"] == "One"
    @test xf["Sheet1"]["A2"] == 1

    @test collect(keys(xf.binary_data)) == ["xl/printerSettings/printerSettings1.bin"]
end
