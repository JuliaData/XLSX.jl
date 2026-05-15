
import XLSX
import Tables
using Test, Dates, XML
using OrderedCollections: OrderedDict
import DataFrames, Random
import Distributions as Dist
import CSV
using StyledStrings


data_directory = joinpath(dirname(pathof(XLSX)), "..", "test", "data")

@assert isdir(data_directory)

src_data_directory = joinpath(dirname(pathof(XLSX)), "data")

@assert isdir(src_data_directory)

@testset "read test files" begin
    ef_blank_ptbr_1904 = XLSX.readxlsx(joinpath(data_directory, "blank_ptbr_1904.xlsx"))
    ef_Book1 = XLSX.readxlsx(joinpath(data_directory, "Book1.xlsx"))
    ef_Book_1904 = XLSX.readxlsx(joinpath(data_directory, "Book_1904.xlsx"))
    ef_book_1904_ptbr = XLSX.readxlsx(joinpath(data_directory, "book_1904_ptbr.xlsx"))
    ef_book_sparse = XLSX.readxlsx(joinpath(data_directory, "book_sparse.xlsx"))
    ef_book_sparse_2 = XLSX.readxlsx(joinpath(data_directory, "book_sparse_2.xlsx"))

    XLSX.readxlsx(joinpath(data_directory, "missing_numFmtId.xlsx"))["Koldioxid (CO2)"][7, 5]

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

    @testset "Read XLS file error" begin
        @test_throws XLSX.XLSXError XLSX.readxlsx(joinpath(data_directory, "old.xls"))
        try
            XLSX.readxlsx(joinpath(data_directory, "old.xls"))
            @test false # didn't throw exception
        catch e
            @test occursin("This package does not support XLS file format", "$e")
        end

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
#        @test_throws XLSX.XLSXError XLSX.readxlsx(joinpath(data_directory, "Template File.xltx"))
#        try
#            XLSX.readxlsx(joinpath(data_directory, "Template File.xltx"))
#            @test false # didn't throw exception
#        catch e
#            @test occursin("does not support Excel template files", "$e")
#        end
    end

    @testset "read .xltx file" begin
        xf = XLSX.readxlsx(joinpath(data_directory, "Template File.xltx"))
        s=xf[1]
        @test s["P5"] == 5
        @test XLSX.getFormula(s, "B5") == "=RANDBETWEEN(0,100)"
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
        ef = XLSX.readxlsx("mytest.xlsx")
        @test ef["Sheet1"]["A1"] == 1
        f = XLSX.openxlsx("mytest2.xlsx", mode="w")
        @test f.source == "mytest2.xlsx"
        f[1]["A1"] = 1
        XLSX.writexlsx("mytest3.xlsx", f, overwrite=true)
        ef = XLSX.readxlsx("mytest3.xlsx")
        @test ef["Sheet1"]["A1"] == 1
        f = XLSX.newxlsx()
        @test f.source == "blank.xlsx"
        f[1]["A1"] = 1
        XLSX.writexlsx("mytest.xlsx", f, overwrite=true)
        ef = XLSX.readxlsx("mytest.xlsx")
        @test ef["Sheet1"]["A1"] == 1
        for f in ["mytest.xlsx", "mytest2.xlsx", "mytest3.xlsx"]
            isfile(f) && rm(f)
        end
    end

    @testset "Fix timestamp" begin
        t = Dates.now(Dates.UTC) - Dates.Second(1)
        xf = XLSX.newxlsx()
        f = "docProps/core.xml"
        date_format = Dates.dateformat"yyyy-mm-ddTHH:MM:SSZ"
        i, j = XLSX.get_idces(xf.data[f], "cp:coreProperties", "dcterms:created") # i, j should always be found unless the `blank.xlsx` file is updated
        @test t < DateTime(XML.value(xf.data[f][i][j][1]), date_format)
        i, j = XLSX.get_idces(xf.data[f], "cp:coreProperties", "dcterms:modified")
        @test t < DateTime(XML.value(xf.data[f][i][j][1]), date_format)

        xf = XLSX.newxlsx(; update_timestamp=false) # do not update timestamp
        i, j = XLSX.get_idces(xf.data[f], "cp:coreProperties", "dcterms:created") # i, j should always be found unless the `blank.xlsx` file is updated
        @test XML.value(xf.data[f][i][j][1]) == "2018-05-22T02:41:32Z"
        i, j = XLSX.get_idces(xf.data[f], "cp:coreProperties", "dcterms:modified")
        @test XML.value(xf.data[f][i][j][1]) == "2018-05-22T02:42:04Z"
    end

end

@testset "Cell names" begin
    @test !XLSX.is_valid_cellname("A0")
    @test XLSX.is_valid_cellname("A1")
    @test !XLSX.is_valid_cellname("A")
    @test !XLSX.is_valid_cellname("1")
    @test XLSX.is_valid_cellname("XFD1048576")
    @test !XLSX.is_valid_cellname("XFD1048577")
    @test XLSX.is_valid_cellname("XFD1")
    @test !XLSX.is_valid_cellname("ZFD1")
    @test XLSX.is_valid_column_name("A")
    @test XLSX.is_valid_column_name("AZ")
    @test XLSX.is_valid_column_name("AAZ")
    @test !XLSX.is_valid_column_name("AAAZ")
    @test !XLSX.is_valid_column_name(":")
    @test !XLSX.is_valid_column_name("É")
    @test XLSX.is_valid_row_name("1")
    @test XLSX.is_valid_row_name("12")
    @test XLSX.is_valid_row_name("123")
    @test !XLSX.is_valid_row_name("012")
    @test !XLSX.is_valid_row_name(":")
    @test !XLSX.is_valid_row_name("A")

    @test XLSX.is_valid_sheet_cellname("Sheet1!A2")
    @test !XLSX.is_valid_sheet_cellname("Sheet1!A2:B3")
    @test !XLSX.is_valid_sheet_cellname("Sheet1!A0")
    @test !XLSX.is_valid_sheet_cellname("A1")
    @test !XLSX.is_valid_sheet_cellname("Sheet1!")
    @test !XLSX.is_valid_sheet_cellname("Sheet1")
    @test !XLSX.is_valid_sheet_cellname("Sheet1!A")
    @test !XLSX.is_valid_sheet_cellname("Sheet1!1")
    @test XLSX.is_valid_sheet_cellname("NEGOCIAÇÕES Descrição!A1")

    @test XLSX.is_valid_sheet_cellrange("Sheet1!A1:B4")
    @test !XLSX.is_valid_sheet_cellrange("Sheet1!:B4")
    @test !XLSX.is_valid_sheet_cellrange("Sheet1!A1:")
    @test !XLSX.is_valid_sheet_cellrange("Sheet1!:")
    @test !XLSX.is_valid_sheet_cellrange("A1:B4")
    @test !XLSX.is_valid_sheet_cellrange("Sheet1!")
    @test !XLSX.is_valid_sheet_cellrange("Sheet1")
    @test !XLSX.is_valid_sheet_cellrange("mysheet!A1")

    @test XLSX.is_valid_sheet_column_range("Sheet1!A:B")
    @test XLSX.is_valid_sheet_column_range("Sheet1!AB:BC")
    @test !XLSX.is_valid_sheet_column_range("A:B")
    @test !XLSX.is_valid_sheet_column_range("Sheet1!")
    @test !XLSX.is_valid_sheet_column_range("Sheet1")
    @test XLSX.is_valid_sheet_row_range("Sheet1!1:2")
    @test XLSX.is_valid_sheet_row_range("Sheet1!12:23")
    @test !XLSX.is_valid_sheet_row_range("1:2")
    @test !XLSX.is_valid_sheet_row_range("Sheet1!")
    @test !XLSX.is_valid_sheet_row_range("Sheet1")

    @test XLSX.is_valid_non_contiguous_range("Sheet1!B1,Sheet1!B3")
    @test XLSX.is_valid_non_contiguous_range("Sheet1!B1,Sheet1!GZ75:HB127")
    @test XLSX.is_valid_non_contiguous_range("B2,B5")
    @test XLSX.is_valid_non_contiguous_range("C3:C5,D6,G7:G8")
    @test !XLSX.is_valid_non_contiguous_range("Sheet1!C3,Sheet2!C3")
    @test !XLSX.is_valid_non_contiguous_range("Sheet1!B3")
    @test !XLSX.is_valid_non_contiguous_range("Sheet1!B3:C6")

    @test in(XLSX.SheetCellRef("Sheet1!A1"), XLSX.NonContiguousRange("Sheet1!A1,Sheet1!B2")) == true
    @test in(XLSX.SheetCellRef("Sheet1!B2"), XLSX.NonContiguousRange("Sheet1!A1,Sheet1!B2")) == true
    @test in(XLSX.SheetCellRef("Sheet1!A2"), XLSX.NonContiguousRange("Sheet1!A1,Sheet1!B2")) == false

    cn = XLSX.CellRef("A1")
    @test string(cn) == "A1"
    @test XLSX.column_name(cn) == "A"
    @test XLSX.row_number(cn) == 1
    @test XLSX.column_number(cn) == 1

    cn = XLSX.CellRef("AB1")
    @test string(cn) == "AB1"
    @test XLSX.column_name(cn) == "AB"
    @test XLSX.row_number(cn) == 1
    @test XLSX.column_number(cn) == 28

    cn = XLSX.CellRef("AMI1")
    @test string(cn) == "AMI1"
    @test XLSX.column_name(cn) == "AMI"
    @test XLSX.row_number(cn) == 1
    @test XLSX.column_number(cn) == 1023

    cn = XLSX.CellRef("XFD1048576")
    @test string(cn) == "XFD1048576"
    @test XLSX.column_name(cn) == "XFD"
    @test XLSX.row_number(cn) == XLSX.EXCEL_MAX_ROWS
    @test XLSX.column_number(cn) == XLSX.EXCEL_MAX_COLS

    v_column_numbers = [1, 15, 22, 23, 24, 25, 26, 27, 28, 29, 30, 38, 39, 40, 41, 42, 43, 44, 45, 46, 47, 48, 49, 50, 51, 52, 53, 54, 55, 56, 57, 58, 59, 60, 61, 74, 75, 76, 77, 78, 79, 80, 81, 82, 83, 84, 85, 86, 284, 285, 286, 287, 288, 289, 296, 297, 299, 300, 301, 700, 701, 702, 703, 704, 705, 706, 727, 728, 729, 730, 731, 1008, 1013, 1014, 1015, 1016, 1017, 1018, 1023, 1024, 1376, 1377, 1378, 1379, 1380, 1381, 3379, 3380, 3381, 3382, 3383, 3403, 3404, 3405, 3406, 3407, 16250, 16251, 16354, 16355, 16384]

    v_column_names = ["A", "O", "V", "W", "X", "Y", "Z", "AA", "AB", "AC", "AD", "AL", "AM", "AN", "AO", "AP", "AQ", "AR", "AS", "AT", "AU", "AV", "AW", "AX", "AY", "AZ", "BA", "BB", "BC", "BD", "BE", "BF", "BG", "BH", "BI", "BV", "BW", "BX", "BY", "BZ", "CA", "CB", "CC", "CD", "CE", "CF", "CG", "CH", "JX", "JY", "JZ", "KA", "KB", "KC", "KJ", "KK", "KM", "KN", "KO", "ZX", "ZY", "ZZ", "AAA", "AAB", "AAC", "AAD", "AAY", "AAZ", "ABA", "ABB", "ABC", "ALT", "ALY", "ALZ", "AMA", "AMB", "AMC", "AMD", "AMI", "AMJ", "AZX", "AZY", "AZZ", "BAA", "BAB", "BAC", "DYY", "DYZ", "DZA", "DZB", "DZC", "DZW", "DZX", "DZY", "DZZ", "EAA", "WZZ", "XAA", "XDZ", "XEA", "XFD"]

    @assert length(v_column_names) == length(v_column_numbers) "Test script is wrong."

    for i in axes(v_column_names, 1)
        @test XLSX.encode_column_number(v_column_numbers[i]) == v_column_names[i]
        @test XLSX.decode_column_number(v_column_names[i]) == v_column_numbers[i]
    end

    @testset "CellRef" begin
        ref = XLSX.CellRef(12, 2)
        @test XLSX.cellname(ref) == "B12"
        show(IOBuffer(), ref)
    end

    cr = XLSX.range"A1:C4"
    @test string(cr) == "A1:C4"
    @test XLSX.row_number(cr.start) == 1
    @test XLSX.column_number(cr.start) == 1
    @test XLSX.row_number(cr.stop) == 4
    @test XLSX.column_number(cr.stop) == 3
    @test size(cr) == (4, 3)
    show(IOBuffer(), cr)

    cr = XLSX.range"B2:C8"
    @test XLSX.ref"B2" ∈ cr
    @test XLSX.ref"B3" ∈ cr
    @test XLSX.ref"C2" ∈ cr
    @test XLSX.ref"C3" ∈ cr
    @test XLSX.ref"C8" ∈ cr
    @test XLSX.ref"A1" ∉ cr
    @test XLSX.ref"C9" ∉ cr
    @test XLSX.ref"D4" ∉ cr
    @test size(cr) == (7, 2)

    fullrng = XLSX.range"B2:E5"
    @test fullrng ⊆ fullrng
    @test XLSX.range"B3:D4" ⊆ fullrng
    @test !issubset(XLSX.range"A1:E5", fullrng)

    @test XLSX.is_valid_cellrange("B2:C8")
    @test !XLSX.is_valid_cellrange("A:B")
    @test_throws XLSX.XLSXError XLSX.CellRange("Z10:A1")
    @test_throws XLSX.XLSXError XLSX.CellRange("Z1:A1")

    # hashing and equality
    @test XLSX.CellRef("AMI1") == XLSX.CellRef("AMI1")
    @test hash(XLSX.CellRef("AMI1")) == hash(XLSX.CellRef("AMI1"))
    @test XLSX.CellRange("A1:C4") == XLSX.CellRange("A1:C4")
    @test hash(XLSX.CellRange("A1:C4")) == hash(XLSX.CellRange("A1:C4"))

    # relative cell position
    rng = XLSX.range"B2:D4"
    @test XLSX.relative_cell_position(XLSX.ref"C3", rng) == (2, 2)
    @test XLSX.relative_cell_position(XLSX.ref"B2", rng) == (1, 1)
    @test XLSX.relative_cell_position(XLSX.ref"C4", rng) == (3, 2)
    @test XLSX.relative_cell_position(XLSX.ref"D4", rng) == (3, 3)
    @test XLSX.relative_cell_position(XLSX.EmptyCell(XLSX.ref"D4"), rng) == (3, 3)

    # SheetCellRef, SheetCellRange, SheetColumnRange
    ref = XLSX.SheetCellRef("Sheet1!A2")
    @test string(ref) == "Sheet1!A2"
    @test ref.sheet == "Sheet1"
    @test ref.cellref == XLSX.CellRef("A2")
    @test XLSX.SheetCellRef("Sheet1!A2") == XLSX.SheetCellRef("Sheet1!A2")
    @test hash(XLSX.SheetCellRef("Sheet1!A2")) == hash(XLSX.SheetCellRef("Sheet1!A2"))
    show(IOBuffer(), ref)

    ref = XLSX.SheetCellRange("Sheet1!A1:B4")
    @test ref.sheet == "Sheet1"
    @test ref.rng == XLSX.CellRange("A1:B4")
    @test_throws XLSX.XLSXError XLSX.SheetCellRange("Sheet1!B4:A1")
    @test XLSX.SheetCellRange("Sheet1!A1:B4") == XLSX.SheetCellRange("Sheet1!A1:B4")
    @test hash(XLSX.SheetCellRange("Sheet1!A1:B4")) == hash(XLSX.SheetCellRange("Sheet1!A1:B4"))
    show(IOBuffer(), ref)

    ref = XLSX.SheetColumnRange("Sheet1!A:B")
    @test string(ref) == "Sheet1!A:B"
    @test ref.sheet == "Sheet1"
    @test ref.colrng == XLSX.ColumnRange("A:B")
    @test XLSX.SheetColumnRange("Sheet1!A:B") == XLSX.SheetColumnRange("Sheet1!A:B")
    @test hash(XLSX.SheetColumnRange("Sheet1!A:B")) == hash(XLSX.SheetColumnRange("Sheet1!A:B"))
    show(IOBuffer(), ref)

    ref = XLSX.SheetColumnRange("Sheet1!A:AA")
    @test string(ref) == "Sheet1!A:AA"
    @test ref.sheet == "Sheet1"
    @test ref.colrng == XLSX.ColumnRange("A:AA")
    @test XLSX.SheetColumnRange("Sheet1!A:AA") == XLSX.SheetColumnRange("Sheet1!A:AA")
    @test hash(XLSX.SheetColumnRange("Sheet1!A:AA")) == hash(XLSX.SheetColumnRange("Sheet1!A:AA"))
    show(IOBuffer(), ref)

    ref = XLSX.SheetColumnRange("Sheet1!AA:AA")
    @test string(ref) == "Sheet1!AA:AA"
    @test ref.sheet == "Sheet1"
    @test ref.colrng == XLSX.ColumnRange("AA:AA")
    @test XLSX.SheetColumnRange("Sheet1!AA:AA") == XLSX.SheetColumnRange("Sheet1!AA:AA")
    @test hash(XLSX.SheetColumnRange("Sheet1!AA:AA")) == hash(XLSX.SheetColumnRange("Sheet1!AA:AA"))
    show(IOBuffer(), ref)

    @test XLSX.is_valid_fixed_sheet_cellname("named_ranges!\$A\$2")
    @test XLSX.is_valid_fixed_sheet_cellrange("named_ranges!\$B\$4:\$C\$5")
    @test !XLSX.is_valid_fixed_sheet_cellname("named_ranges!A2")
    @test !XLSX.is_valid_fixed_sheet_cellrange("named_ranges!B4:C5")
    @test XLSX.SheetCellRef("named_ranges!\$A\$2") == XLSX.SheetCellRef("named_ranges!A2")
    @test XLSX.SheetCellRange("named_ranges!\$B\$4:\$C\$5") == XLSX.SheetCellRange("named_ranges!B4:C5")
end

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

end

@testset "formulas" begin

    @testset "simple formulas" begin
        f = XLSX.newxlsx("mySheet")
        for i = 1:10
            f[1][i, 1] = i
        end
        f[1]["B1:B10"] = 10
        XLSX.setFormula(f[1], "C1:C10", "=A1+B1")
        @test XLSX.get_formula_from_cache(f[1], XLSX.CellRef("C1")) == XLSX.ReferencedFormula("=A1+B1", 0, "C1:C10", nothing)
        @test XLSX.get_formula_from_cache(f[1], XLSX.CellRef("C2")) == XLSX.FormulaReference(0, nothing)
        @test XLSX.get_formula_from_cache(f[1], XLSX.CellRef("C10")) == XLSX.FormulaReference(0, nothing)
        XLSX.setFormula(f[1], 11, 1:3, "sum(A1:A10)")
        @test XLSX.get_formula_from_cache(f[1], XLSX.CellRef("C11")) == XLSX.FormulaReference(1, nothing)
        XLSX.setFormula(f[1], 11, 4, "=sum(A1:C10)")
        @test XLSX.get_formula_from_cache(f[1], XLSX.CellRef("D11")) == XLSX.Formula("=sum(A1:C10)", nothing, nothing, nothing)
        XLSX.setFormula(f, "mySheet!A11:C11", "=A10/\$D11")
        XLSX.writexlsx("formulas.xlsx", f, overwrite=true)

        f = XLSX.openxlsx("formulas.xlsx", mode="rw")
        @test XLSX.get_formula_from_cache(f[1], XLSX.CellRef("C1")) == XLSX.ReferencedFormula("=A1+B1", 0, "C1:C10", nothing)
        @test XLSX.get_formula_from_cache(f[1], XLSX.CellRef("C2")) == XLSX.FormulaReference(0, nothing)
        @test XLSX.get_formula_from_cache(f[1], XLSX.CellRef("C10")) == XLSX.FormulaReference(0, nothing)
        @test XLSX.get_formula_from_cache(f[1], XLSX.CellRef("A11")) == XLSX.ReferencedFormula("=A10/\$D11", 1, "A11:C11", nothing)
        @test XLSX.get_formula_from_cache(f[1], XLSX.CellRef("C11")) == XLSX.FormulaReference(1, nothing)
        @test XLSX.get_formula_from_cache(f[1], XLSX.CellRef("D11")) == XLSX.Formula("=sum(A1:C10)", nothing, nothing, nothing)
        isfile("formulas.xlsx") && rm("formulas.xlsx")

        f = XLSX.newxlsx("mySheet")
        s = f["mySheet"]
        s[1:12, 1] = [x for x in 1:12]
        s["B1:L12"] = ""
        XLSX.setFormula(s, "B:L", "=\$A1+A1")
        XLSX.addsheet!(f, "moreFormulas")
        XLSX.addsheet!(f, "empty")
        s1 = f["moreFormulas"]
        s2 = f["empty"]
        s1[1:12, 1:12] = ""
        s2[1:12, 1:12] = ""
        XLSX.setFormula(s1, 1, :, "=max(mySheet!A1:A12)")
        XLSX.setFormula(s1, "2", "=min(\$A1:A1)")
        XLSX.setFormula(s2, :, :, "=mySheet!A\$1+mySheet!A1")
        @test XLSX.get_formula_from_cache(f[1], XLSX.CellRef("B1")) == XLSX.ReferencedFormula("=\$A1+A1", 0, "B1:L12", nothing)
        @test XLSX.get_formula_from_cache(f[1], XLSX.CellRef("B2")) == XLSX.FormulaReference(0, nothing)
        @test XLSX.get_formula_from_cache(f[1], XLSX.CellRef("L10")) == XLSX.FormulaReference(0, nothing)
        @test XLSX.get_formula_from_cache(f[2], XLSX.CellRef("A1")) == XLSX.Formula("=max(mySheet!A1:A12)", nothing, nothing, nothing)
        @test XLSX.get_formula_from_cache(f[2], XLSX.CellRef("G1")) == XLSX.Formula("=max(mySheet!G1:G12)", nothing, nothing, nothing)
        @test XLSX.get_formula_from_cache(f[2], XLSX.CellRef("L1")) == XLSX.Formula("=max(mySheet!L1:L12)", nothing, nothing, nothing)
        @test XLSX.get_formula_from_cache(f[2], XLSX.CellRef("A2")) == XLSX.ReferencedFormula("=min(\$A1:A1)", 0, "A2:L2", nothing)
        @test XLSX.get_formula_from_cache(f[2], XLSX.CellRef("L2")) == XLSX.FormulaReference(0, nothing)
        @test XLSX.get_formula_from_cache(f[3], XLSX.CellRef("A1")) == XLSX.Formula("=mySheet!A\$1+mySheet!A1", nothing, nothing, nothing)
        @test XLSX.get_formula_from_cache(f[3], XLSX.CellRef("F6")) == XLSX.Formula("=mySheet!F\$1+mySheet!F6", nothing, nothing, nothing)
        @test XLSX.get_formula_from_cache(f[3], XLSX.CellRef("L12")) == XLSX.Formula("=mySheet!L\$1+mySheet!L12", nothing, nothing, nothing)
        XLSX.writexlsx("formulas.xlsx", f, overwrite=true)

        f = XLSX.openxlsx("formulas.xlsx", mode="rw")
        @test XLSX.get_formula_from_cache(f[1], XLSX.CellRef("B1")) == XLSX.ReferencedFormula("=\$A1+A1", 0, "B1:L12", nothing)
        @test XLSX.get_formula_from_cache(f[1], XLSX.CellRef("B2")) == XLSX.FormulaReference(0, nothing)
        @test XLSX.get_formula_from_cache(f[1], XLSX.CellRef("L10")) == XLSX.FormulaReference(0, nothing)
        @test XLSX.get_formula_from_cache(f[2], XLSX.CellRef("A1")) == XLSX.Formula("=max(mySheet!A1:A12)", nothing, nothing, nothing)
        @test XLSX.get_formula_from_cache(f[2], XLSX.CellRef("G1")) == XLSX.Formula("=max(mySheet!G1:G12)", nothing, nothing, nothing)
        @test XLSX.get_formula_from_cache(f[2], XLSX.CellRef("L1")) == XLSX.Formula("=max(mySheet!L1:L12)", nothing, nothing, nothing)
        @test XLSX.get_formula_from_cache(f[2], XLSX.CellRef("A2")) == XLSX.ReferencedFormula("=min(\$A1:A1)", 0, "A2:L2", nothing)
        @test XLSX.get_formula_from_cache(f[2], XLSX.CellRef("L2")) == XLSX.FormulaReference(0, nothing)
        @test XLSX.get_formula_from_cache(f[3], XLSX.CellRef("A1")) == XLSX.Formula("=mySheet!A\$1+mySheet!A1", nothing, nothing, nothing)
        @test XLSX.get_formula_from_cache(f[3], XLSX.CellRef("F6")) == XLSX.Formula("=mySheet!F\$1+mySheet!F6", nothing, nothing, nothing)
        @test XLSX.get_formula_from_cache(f[3], XLSX.CellRef("L12")) == XLSX.Formula("=mySheet!L\$1+mySheet!L12", nothing, nothing, nothing)
        isfile("formulas.xlsx") && rm("formulas.xlsx")

        f = XLSX.newxlsx("mySheet")
        s = f[1]
        s[1:10, 1:10] = rand(10, 10)
        XLSX.setFont(s, :; color="red")
        XLSX.setFormula(s, "mySheet!B2:J2", "=\$A\$2+A2")
        XLSX.setFormula(s, "mySheet!3:4", "=\$A1+B1")
        XLSX.setFormula(s, "mySheet!F:G", "=E\$1+E1")
        XLSX.setFormula(s, "J10", "=A1+10")
        @test XLSX.get_formula_from_cache(s, XLSX.CellRef("D2")) == XLSX.FormulaReference(0, nothing)
        @test XLSX.get_formula_from_cache(s, XLSX.CellRef("D4")) == XLSX.FormulaReference(1, nothing)
        @test XLSX.get_formula_from_cache(s, XLSX.CellRef("G4")) == XLSX.FormulaReference(2, nothing)
        @test XLSX.get_formula_from_cache(s, XLSX.CellRef("J10")) == XLSX.Formula("=A1+10", nothing, nothing, nothing)
        XLSX.setFormula(s, :, 10, "=\$A\$1+5")
        @test XLSX.get_formula_from_cache(f[1], XLSX.CellRef("J1")) == XLSX.ReferencedFormula("=\$A\$1+5", 3, "J1:J10", nothing)
        @test XLSX.get_formula_from_cache(f[1], XLSX.CellRef("J2")) == XLSX.FormulaReference(3, nothing)
        @test XLSX.get_formula_from_cache(f[1], XLSX.CellRef("J8")) == XLSX.FormulaReference(3, nothing)
        XLSX.setFormula(s, 10, :, "=\$A\$1+6")
        @test XLSX.get_formula_from_cache(f[1], XLSX.CellRef("A10")) == XLSX.ReferencedFormula("=\$A\$1+6", 4, "A10:J10", nothing)
        @test XLSX.get_formula_from_cache(f[1], XLSX.CellRef("B10")) == XLSX.FormulaReference(4, nothing)
        @test XLSX.get_formula_from_cache(f[1], XLSX.CellRef("G10")) == XLSX.FormulaReference(4, nothing)
    end
    @testset "dynamic array" begin
        f = XLSX.openxlsx(joinpath(data_directory, "Unique.xlsx"), mode="rw")
        @test XLSX.getcell(f[1], "C1") == XLSX.Cell(XLSX.get_workbook(f), XLSX.CellRef("C1"), "", "", "1", "1", true)
        @test XLSX.get_formula_from_cache(f[1], XLSX.CellRef("C1")) == XLSX.Formula("_xlfn.UNIQUE(A1:A9)", "array", "C1:C3", nothing)
        s = f[1]
        XLSX.setFormula(s, "B1", "=A1")
        @test XLSX.get_formula_from_cache(f[1], XLSX.CellRef("B1")) == XLSX.Formula("=A1", nothing, nothing, nothing)
        XLSX.setFormula(s, "B2:B10", "=A2+B1")
        @test XLSX.get_formula_from_cache(f[1], XLSX.CellRef("B2")) == XLSX.ReferencedFormula("=A2+B1", 0, "B2:B10", nothing)
        @test XLSX.get_formula_from_cache(f[1], XLSX.CellRef("B5")) == XLSX.FormulaReference(0, nothing)
        XLSX.setFormula(s, "D1", "=sort(B1:B10,,-1)")
        @test XLSX.getcell(f[1], "D1") == XLSX.Cell(XLSX.get_workbook(f), XLSX.CellRef("D1"), "", "", "", "1", true)
        @test XLSX.get_formula_from_cache(f[1], XLSX.CellRef("D1")) == XLSX.Formula("=_xlfn.SORT(B1:B10,,-1)", "array", "D1:D1", nothing)
        XLSX.writexlsx("formulas.xlsx", f, overwrite=true)

        f = XLSX.opentemplate("formulas.xlsx")
        @test XLSX.getcell(f[1], "C1") == XLSX.Cell(XLSX.get_workbook(f), XLSX.CellRef("C1"), "", "", "1", "1", true)
        @test XLSX.get_formula_from_cache(f[1], XLSX.CellRef("C1")) == XLSX.Formula("_xlfn.UNIQUE(A1:A9)", "array", "C1:C3", nothing)
        s = f[1]
        @test XLSX.get_formula_from_cache(f[1], XLSX.CellRef("B1")) == XLSX.Formula("=A1", nothing, nothing, nothing)
        @test XLSX.get_formula_from_cache(f[1], XLSX.CellRef("B2")) == XLSX.ReferencedFormula("=A2+B1", 0, "B2:B10", nothing)
        @test XLSX.get_formula_from_cache(f[1], XLSX.CellRef("B5")) == XLSX.FormulaReference(0, nothing)
        @test XLSX.getcell(f[1], "D1") == XLSX.Cell(XLSX.get_workbook(f), XLSX.CellRef("D1"), "", "", "", "1", true)
        @test XLSX.get_formula_from_cache(f[1], XLSX.CellRef("D1")) == XLSX.Formula("=_xlfn.SORT(B1:B10,,-1)", "array", "D1:D1", nothing)
        isfile("formulas.xlsx") && rm("formulas.xlsx")

        f = XLSX.newxlsx("mySheet")
        s = f[1]
        s[1:5, 1] = [x for x in 3:3:15]
        s[1:5, 2] = ""
        XLSX.setFormula(s, "mySheet!B1", "=sort(A1:A5, , -1)")
        @test XLSX.getcell(f[1], "B1") == XLSX.Cell(XLSX.get_workbook(f), XLSX.CellRef("B1"), "", "", "", "1", true)
        @test XLSX.get_formula_from_cache(f[1], XLSX.CellRef("B1")) == XLSX.Formula("=_xlfn.SORT(A1:A5, , -1)", "array", "B1:B1", nothing)
        XLSX.setFormula(s, "C1", "=if(A1:A5>30, \"High\", \"Low\")")
        XLSX.setFormula(s, "D1", "=OFFSET(A2:A5, -1, 0)")
        XLSX.setFormula(s, "E1", "=CHOOSE(1,A1:A2,A3:A4)")
        s["F1:F2"] = [2; 4]
        XLSX.setFormula(s, "G1", "=INDEX(A1:A5,F1:F2)")
        @test XLSX.getFormula(s, "C1") == "=if(A1:A5>30, \"High\", \"Low\")"
        @test XLSX.getFormula(s, "D1") == "=OFFSET(A2:A5, -1, 0)"
        @test XLSX.getFormula(s, "E1") == "=CHOOSE(1,A1:A2,A3:A4)"
        @test XLSX.getFormula(s, "G1") == "=INDEX(A1:A5,F1:F2)"
    end
    @testset "getFormula" begin
        f = XLSX.newxlsx()
        s = f[1]
        s[1:10, 1] = 10:10:100
        s[1:6, 7] = 1:6
        XLSX.setFormula(s, "B1", "=A1")
        XLSX.setFormula(s, "B2:B10", "=A2+B1")
        XLSX.setFormula(f, "Sheet1!C1", "=B1:B10")
        XLSX.setFormula(s, "D1", "=INT(B1:B10/100)+1")
        XLSX.setFormula(s, "E1", "=groupby(D1#,B1:B10,sum)")
        XLSX.setFormula(s, "H1", "=frequency(D1#,G1:G6)")
        @test XLSX.get_referenced_formula(s, XLSX.CellRef("B4")) == "=A4+B3"
        @test XLSX.getFormula(s, "B1") == "=A1"
        @test XLSX.getFormula(s, 2, 2) == "=A2+B1"
        @test XLSX.getFormula(f, "Sheet1!B10") == "=A10+B9"
        @test XLSX.getFormula(s, "D1") == "=INT(B1:B10/100)+1"
        @test XLSX.getFormula(s, "E1") == "=_xlfn.GROUPBY(_xlfn.ANCHORARRAY(D1),B1:B10,_xleta.sum)"
        @test XLSX.getFormula(s, "H1") == "=frequency(_xlfn.ANCHORARRAY(D1),G1:G6)"
    end

    @testset "spillranges" begin
        SPILL_REF_TESTS = [
            # Basic spill references
            ("=G1#", "=ANCHORARRAY(G1)", "basic_cell"),
            ("=Sheet1!A1#", "=ANCHORARRAY(Sheet1!A1)", "basic_sheet"),
            ("='My Sheet'!B2#", "=ANCHORARRAY('My Sheet'!B2)", "quoted_sheet"),
            ("=Table1[Column]#", "=ANCHORARRAY(Table1[Column])", "structured_ref"),

            # Spill references inside functions
            ("=VLOOKUP(M1,G1#,3,FALSE)", "=VLOOKUP(M1,ANCHORARRAY(G1),3,FALSE)", "vlookup"),
            ("=SUM(Table1[Column]#)", "=SUM(ANCHORARRAY(Table1[Column]))", "sum_structured"),
            ("=IF(G1#>0, G1#, 0)", "=IF(ANCHORARRAY(G1)>0, ANCHORARRAY(G1), 0)", "if_branch"),
            ("=INDEX(Table1[Column]#, 2)", "=INDEX(ANCHORARRAY(Table1[Column]), 2)", "index_structured"),
            ("=FILTER(G1#, G1#>0)", "=FILTER(ANCHORARRAY(G1), ANCHORARRAY(G1)>0)", "filter"),

            # Multiple spill references
            ("=G1#+H1#", "=ANCHORARRAY(G1)+ANCHORARRAY(H1)", "add_two"),
            ("=SUM(G1#,H1#)", "=SUM(ANCHORARRAY(G1),ANCHORARRAY(H1))", "sum_two"),
            ("=IF(G1#>H1#, G1#, H1#)", "=IF(ANCHORARRAY(G1)>ANCHORARRAY(H1), ANCHORARRAY(G1), ANCHORARRAY(H1))", "if_compare"),
            ("=VLOOKUP(M1,G1#,3,FALSE)+VLOOKUP(M2,H1#,2,FALSE)", "=VLOOKUP(M1,ANCHORARRAY(G1),3,FALSE)+VLOOKUP(M2,ANCHORARRAY(H1),2,FALSE)", "vlookup_double"),

            # Sheet and workbook prefixes
            ("='Sales Data'!Table1[Revenue]#", "=ANCHORARRAY('Sales Data'!Table1[Revenue])", "sheet_structured"),
            ("='[Book1.xlsx]Sheet1'!A1#", "=ANCHORARRAY('[Book1.xlsx]Sheet1'!A1)", "book_sheet"),
            ("=Sheet1!\$B\$2#", "=ANCHORARRAY(Sheet1!\$B\$2)", "absolute_sheet"),
            ("=Table1[[#All],[Column]]#", "=ANCHORARRAY(Table1[[#All],[Column]])", "multi_structured"),
            ("='Table 1'[[#All],[Column]]#", "=ANCHORARRAY('Table 1'[[#All],[Column]])", "quoted_structured"),
            ("='My Table'[[#Headers],[Column1],[Column2]]#", "=ANCHORARRAY('My Table'[[#Headers],[Column1],[Column2]])", "quoted_multi_structured"),

            # Non-spill references (should remain unchanged)
            ("=G1", "=G1", "plain_cell"),
            ("=Sheet1!A1", "=Sheet1!A1", "plain_sheet"),
            ("=Table1[Column]", "=Table1[Column]", "plain_structured"),
            ("=VLOOKUP(M1,G1,3,FALSE)", "=VLOOKUP(M1,G1,3,FALSE)", "vlookup_plain"),
            ("=SUM(A1:A10)", "=SUM(A1:A10)", "sum_range")
        ]
        for (input, expected, label) in SPILL_REF_TESTS
            output = XLSX.anchor_spill_refs(input)
            @test output == expected
        end
    end
    @testset "external references" begin
        f = XLSX.openxlsx(joinpath(data_directory, "linked-1.xlsx"), mode="rw")
        s = f[1]
        @test XLSX.getFormula(s, "A1") == "=[1]Sheet1!\$A\$1"
        @test occursin("linked-2.xlsx]", XLSX.getFormula(s, "A1"; get_external_refs=true))
        f = XLSX.openxlsx(joinpath(data_directory, "linked-2.xlsx"), mode="rw")
        s = f[1]
        @test XLSX.getFormula(s, "B1") == "=[1]Sheet1!\$B\$1"
        @test occursin("linked-1.xlsx]", XLSX.getFormula(s, "B1"; get_external_refs=true))
    end

    @testset "ReferencedFormulae" begin

        f = XLSX.openxlsx(joinpath(data_directory, "reftest.xlsx"), mode="rw")

        s = f[1]
        wb = XLSX.get_workbook(s)
        @test XLSX.getcell(s, "A2") == XLSX.Cell(XLSX.get_workbook(f), XLSX.CellRef("A2"), "", "", "20", "", true)
        @test XLSX.get_formula_from_cache(s, XLSX.CellRef("A2")) == XLSX.ReferencedFormula("SUM(O2:S2)", 0, "A2:A10", nothing)
        @test XLSX.getcell(s, "A3") == XLSX.Cell(XLSX.get_workbook(f), XLSX.CellRef("A3"), "", "", "25", "", true)
        @test XLSX.get_formula_from_cache(s, XLSX.CellRef("A3")) == XLSX.FormulaReference(0, nothing)
        s["A2"] = 3
        @test XLSX.getcell(s, "A2") == XLSX.Cell(XLSX.get_workbook(f), XLSX.CellRef("A2"), "", "", "3", "", false)
        @test XLSX.getcell(s, "A3") == XLSX.Cell(XLSX.get_workbook(f), XLSX.CellRef("A3"), "", "", "25", "", true)
        @test XLSX.get_formula_from_cache(s, XLSX.CellRef("A3")) == XLSX.ReferencedFormula("SUM(O3:S3)", 0, "A3:A10", nothing)
        @test XLSX.get_formula_from_cache(s, XLSX.CellRef("A4")) == XLSX.FormulaReference(0, nothing)
        
        s2 = f[2]
        @test XLSX.getcell(s2, "A1") == XLSX.Cell(XLSX.get_workbook(f), XLSX.CellRef("A1"), "", "", "54", "", true)
        @test XLSX.get_formula_from_cache(s2, XLSX.CellRef("A1")) == XLSX.Formula("SECOND(NOW())", nothing, nothing, Dict("ca" => "1"))
        @test XLSX.getcell(s2, "A2") == XLSX.Cell(XLSX.get_workbook(f), XLSX.CellRef("A2"), "", "", "54", "", true)
        @test XLSX.get_formula_from_cache(s2, XLSX.CellRef("A2")) == XLSX.ReferencedFormula("SECOND(NOW())", 1, "A2:A5", Dict("ca" => "1"))
        s2["A2"] = 3
        @test XLSX.getcell(s2, "A2") == XLSX.Cell(XLSX.get_workbook(f), XLSX.CellRef("A2"), "", "", "3", "", false)
        @test XLSX.get_formula_from_cache(s2, XLSX.CellRef("A3")).id == 1
        @test XLSX.get_formula_from_cache(s2, XLSX.CellRef("A3")).unhandled == Dict("ca" => "1")
        @test XLSX.getcell(s2, "A3").formula == true
        @test XLSX.getcell(s2, "A3") == XLSX.Cell(XLSX.get_workbook(f), XLSX.CellRef("A3"), "", "", "54", "", true)
        @test XLSX.get_formula_from_cache(s2, XLSX.CellRef("A3")) == XLSX.ReferencedFormula("SECOND(NOW())", 1, "A3:A5", Dict("ca" => "1"))
        @test XLSX.getcell(s2, "B1") == XLSX.Cell(XLSX.get_workbook(f), XLSX.CellRef("B1"), "", "", "54", "", true)
        @test XLSX.get_formula_from_cache(s2, XLSX.CellRef("B1")) == XLSX.ReferencedFormula("SECOND(NOW())", 0, "B1:C5", Dict("ca" => "1"))
        s2["B1"] = 3
        @test XLSX.getcell(s2, "B1") == XLSX.Cell(XLSX.get_workbook(f), XLSX.CellRef("B1"), "", "", "3", "", false)
        @test XLSX.getcell(s2, "B2") == XLSX.Cell(XLSX.get_workbook(f), XLSX.CellRef("B2"), "", "", "54", "", true)
        @test XLSX.get_formula_from_cache(s2, XLSX.CellRef("B2")) == XLSX.ReferencedFormula("SECOND(NOW())", 0, "B2:C5", Dict("ca" => "1"))
        @test XLSX.getcell(s2, "C1") == XLSX.Cell(XLSX.get_workbook(f), XLSX.CellRef("C1"), "", "", "54", "", true)
        @test XLSX.get_formula_from_cache(s2, XLSX.CellRef("C1")) == XLSX.Formula("SECOND(NOW())", nothing, "C1", Dict("ca" => "1"))

        XLSX.writexlsx("mytest.xlsx", f, overwrite=true)
        f2 = XLSX.openxlsx("mytest.xlsx", mode="rw")

        s = f2[1]
        @test XLSX.getcell(s, "A2") == XLSX.Cell(XLSX.get_workbook(f), XLSX.CellRef("A2"), "", "", "3", "", false)
        @test XLSX.getcell(s, "A3") == XLSX.Cell(XLSX.get_workbook(f), XLSX.CellRef("A3"), "", "", "25", "", true)
        @test XLSX.get_formula_from_cache(s, XLSX.CellRef("A3")) == XLSX.ReferencedFormula("SUM(O3:S3)", 0, "A3:A10", nothing)

        s2 = f[2]
        @test XLSX.getcell(s2, "A2") == XLSX.Cell(XLSX.get_workbook(f), XLSX.CellRef("A2"), "", "", "3", "", false)
        @test XLSX.get_formula_from_cache(s2, XLSX.CellRef("A3")).id == 1
        @test XLSX.get_formula_from_cache(s2, XLSX.CellRef("A3")).unhandled == Dict("ca" => "1")
        @test XLSX.getcell(s2, "A3") == XLSX.Cell(XLSX.get_workbook(f), XLSX.CellRef("A3"), "", "", "54", "", true)
        @test XLSX.get_formula_from_cache(s2, XLSX.CellRef("A3")) == XLSX.ReferencedFormula("SECOND(NOW())", 1, "A3:A5", Dict("ca" => "1"))
        @test XLSX.getcell(s2, "B1") == XLSX.Cell(XLSX.get_workbook(f), XLSX.CellRef("B1"), "", "", "3", "", false)
        @test XLSX.getcell(s2, "B2") == XLSX.Cell(XLSX.get_workbook(f), XLSX.CellRef("B2"), "", "", "54", "", true)
        @test XLSX.get_formula_from_cache(s2, XLSX.CellRef("B2")) == XLSX.ReferencedFormula("SECOND(NOW())", 0, "B2:C5", Dict("ca" => "1"))
        @test XLSX.getcell(s2, "C1") == XLSX.Cell(XLSX.get_workbook(f), XLSX.CellRef("C1"), "", "", "54", "", true)
        @test XLSX.get_formula_from_cache(s2, XLSX.CellRef("C1")) ==XLSX.Formula("SECOND(NOW())", nothing, "C1", Dict("ca" => "1"))

    end
end

@testset "getcell" begin
    f = XLSX.newxlsx()
    s = f[1]
    for i in 1:3
        for j in 1:3
            s[i, j] = i + j
        end
    end
    @test XLSX.getcell(s, "A1") == XLSX.Cell(XLSX.get_workbook(f), XLSX.CellRef("A1"), "", "", "2", "", false)
    @test XLSX.getcell(s, "Sheet1!A1") == XLSX.Cell(XLSX.get_workbook(f), XLSX.CellRef("A1"), "", "", "2", "", false)
    @test XLSX.getcell(f, "Sheet1!A1") == XLSX.Cell(XLSX.get_workbook(f), XLSX.CellRef("A1"), "", "", "2", "", false)
    @test XLSX.getcell(s, XLSX.SheetCellRef("Sheet1!A1")) == XLSX.Cell(XLSX.get_workbook(f), XLSX.CellRef("A1"), "", "", "2", "", false)
    @test XLSX.getcell(f, XLSX.SheetCellRef("Sheet1!A1")) == XLSX.Cell(XLSX.get_workbook(f), XLSX.CellRef("A1"), "", "", "2", "", false)
    @test XLSX.getcell(s, "B1:B3") == [XLSX.Cell(XLSX.get_workbook(f), XLSX.CellRef("B1"), "", "", "3", "", false); XLSX.Cell(XLSX.get_workbook(f), XLSX.CellRef("B2"), "", "", "4", "", false); XLSX.Cell(XLSX.get_workbook(f), XLSX.CellRef("B3"), "", "", "5", "", false);;]
    @test XLSX.getcell(s, "Sheet1!B1:B3") == [XLSX.Cell(XLSX.get_workbook(f), XLSX.CellRef("B1"), "", "", "3", "", false); XLSX.Cell(XLSX.get_workbook(f), XLSX.CellRef("B2"), "", "", "4", "", false); XLSX.Cell(XLSX.get_workbook(f), XLSX.CellRef("B3"), "", "", "5", "", false);;]
    @test XLSX.getcell(f, "Sheet1!B1:B3") == [XLSX.Cell(XLSX.get_workbook(f), XLSX.CellRef("B1"), "", "", "3", "", false); XLSX.Cell(XLSX.get_workbook(f), XLSX.CellRef("B2"), "", "", "4", "", false); XLSX.Cell(XLSX.get_workbook(f), XLSX.CellRef("B3"), "", "", "5", "", false);;]
    @test XLSX.getcell(s, XLSX.SheetCellRange("Sheet1!B1:B3")) == [XLSX.Cell(XLSX.get_workbook(f), XLSX.CellRef("B1"), "", "", "3", "", false); XLSX.Cell(XLSX.get_workbook(f), XLSX.CellRef("B2"), "", "", "4", "", false); XLSX.Cell(XLSX.get_workbook(f), XLSX.CellRef("B3"), "", "", "5", "", false);;]
    @test XLSX.getcell(s, "B1,B3") == [[XLSX.Cell(XLSX.get_workbook(f), XLSX.CellRef("B1"), "", "", "3", "", false);;], [XLSX.Cell(XLSX.get_workbook(f), XLSX.CellRef("B3"), "", "", "5", "", false);;]]
    @test XLSX.getcell(s, "Sheet1!B1,Sheet1!B3") == [[XLSX.Cell(XLSX.get_workbook(f), XLSX.CellRef("B1"), "", "", "3", "", false);;], [XLSX.Cell(XLSX.get_workbook(f), XLSX.CellRef("B3"), "", "", "5", "", false);;]]
    @test XLSX.getcell(f, "Sheet1!B1,Sheet1!B3") == [[XLSX.Cell(XLSX.get_workbook(f), XLSX.CellRef("B1"), "", "", "3", "", false);;], [XLSX.Cell(XLSX.get_workbook(f), XLSX.CellRef("B3"), "", "", "5", "", false);;]]
    @test XLSX.getcell(s, "B:B") == [XLSX.Cell(XLSX.get_workbook(f), XLSX.CellRef("B1"), "", "", "3", "", false); XLSX.Cell(XLSX.get_workbook(f), XLSX.CellRef("B2"), "", "", "4", "", false); XLSX.Cell(XLSX.get_workbook(f), XLSX.CellRef("B3"), "", "", "5", "", false);;]
    @test XLSX.getcell(s, "Sheet1!B:B") == [XLSX.Cell(XLSX.get_workbook(f), XLSX.CellRef("B1"), "", "", "3", "", false); XLSX.Cell(XLSX.get_workbook(f), XLSX.CellRef("B2"), "", "", "4", "", false); XLSX.Cell(XLSX.get_workbook(f), XLSX.CellRef("B3"), "", "", "5", "", false);;]
    @test XLSX.getcell(f, "Sheet1!B:B") == [XLSX.Cell(XLSX.get_workbook(f), XLSX.CellRef("B1"), "", "", "3", "", false); XLSX.Cell(XLSX.get_workbook(f), XLSX.CellRef("B2"), "", "", "4", "", false); XLSX.Cell(XLSX.get_workbook(f), XLSX.CellRef("B3"), "", "", "5", "", false);;]
    @test XLSX.getcell(s, XLSX.SheetColumnRange("Sheet1!B:B")) == [XLSX.Cell(XLSX.get_workbook(f), XLSX.CellRef("B1"), "", "", "3", "", false); XLSX.Cell(XLSX.get_workbook(f), XLSX.CellRef("B2"), "", "", "4", "", false); XLSX.Cell(XLSX.get_workbook(f), XLSX.CellRef("B3"), "", "", "5", "", false);;]
    @test XLSX.getcell(s, "Sheet1!2:2") == [XLSX.Cell(XLSX.get_workbook(f), XLSX.CellRef("A2"), "", "", "3", "", false) XLSX.Cell(XLSX.get_workbook(f), XLSX.CellRef("B2"), "", "", "4", "", false) XLSX.Cell(XLSX.get_workbook(f), XLSX.CellRef("C2"), "", "", "5", "", false)]
    @test XLSX.getcell(f, "Sheet1!2:2") == [XLSX.Cell(XLSX.get_workbook(f), XLSX.CellRef("A2"), "", "", "3", "", false) XLSX.Cell(XLSX.get_workbook(f), XLSX.CellRef("B2"), "", "", "4", "", false) XLSX.Cell(XLSX.get_workbook(f), XLSX.CellRef("C2"), "", "", "5", "", false)]
    @test XLSX.getcell(s, XLSX.SheetRowRange("Sheet1!2:2")) == [XLSX.Cell(XLSX.get_workbook(f), XLSX.CellRef("A2"), "", "", "3", "", false) XLSX.Cell(XLSX.get_workbook(f), XLSX.CellRef("B2"), "", "", "4", "", false) XLSX.Cell(XLSX.get_workbook(f), XLSX.CellRef("C2"), "", "", "5", "", false)]
    @test XLSX.getcell(s, "2:2") == [XLSX.Cell(XLSX.get_workbook(f), XLSX.CellRef("A2"), "", "", "3", "", false) XLSX.Cell(XLSX.get_workbook(f), XLSX.CellRef("B2"), "", "", "4", "", false) XLSX.Cell(XLSX.get_workbook(f), XLSX.CellRef("C2"), "", "", "5", "", false)]
    @test XLSX.getcell(s, :, 2) == [XLSX.Cell(XLSX.get_workbook(f), XLSX.CellRef("B1"), "", "", "3", "", false); XLSX.Cell(XLSX.get_workbook(f), XLSX.CellRef("B2"), "", "", "4", "", false); XLSX.Cell(XLSX.get_workbook(f), XLSX.CellRef("B3"), "", "", "5", "", false);;]
    @test XLSX.getcell(s, 2, :) == [XLSX.Cell(XLSX.get_workbook(f), XLSX.CellRef("A2"), "", "", "3", "", false) XLSX.Cell(XLSX.get_workbook(f), XLSX.CellRef("B2"), "", "", "4", "", false) XLSX.Cell(XLSX.get_workbook(f), XLSX.CellRef("C2"), "", "", "5", "", false)]
    @test XLSX.getcell(s, 2, 1:2:3) == [XLSX.Cell(XLSX.get_workbook(f), XLSX.CellRef("A2"), "", "", "3", "", false), XLSX.Cell(XLSX.get_workbook(f), XLSX.CellRef("C2"), "", "", "5", "", false)]
    @test XLSX.getcell(s, 2, [1, 3]) == [XLSX.Cell(XLSX.get_workbook(f), XLSX.CellRef("A2"), "", "", "3", "", false), XLSX.Cell(XLSX.get_workbook(f), XLSX.CellRef("C2"), "", "", "5", "", false)]
    @test XLSX.getcell(s, [2], 1) == [XLSX.Cell(XLSX.get_workbook(f), XLSX.CellRef("A2"), "", "", "3", "", false)]
    @test XLSX.getcell(s, [2], [1, 3]) == [XLSX.Cell(XLSX.get_workbook(f), XLSX.CellRef("A2"), "", "", "3", "", false) XLSX.Cell(XLSX.get_workbook(f), XLSX.CellRef("C2"), "", "", "5", "", false)]
    @test_throws XLSX.XLSXError XLSX.getcell(f, "Sheet1!garbage")
    @test_throws XLSX.XLSXError XLSX.getcell(s, "Sheet1!garbage")
    @test_throws XLSX.XLSXError XLSX.getcell(s, "garbage")
    @test_throws XLSX.XLSXError XLSX.getcell(s, "garbage1:garbage2")

    @test XLSX.getcellrange(s, "Sheet1!B:B") == [XLSX.Cell(XLSX.get_workbook(f), XLSX.CellRef("B1"), "", "", "3", "", false); XLSX.Cell(XLSX.get_workbook(f), XLSX.CellRef("B2"), "", "", "4", "", false); XLSX.Cell(XLSX.get_workbook(f), XLSX.CellRef("B3"), "", "", "5", "", false);;]
    @test XLSX.getcellrange(f, "Sheet1!B:B") == [XLSX.Cell(XLSX.get_workbook(f), XLSX.CellRef("B1"), "", "", "3", "", false); XLSX.Cell(XLSX.get_workbook(f), XLSX.CellRef("B2"), "", "", "4", "", false); XLSX.Cell(XLSX.get_workbook(f), XLSX.CellRef("B3"), "", "", "5", "", false);;]
    @test XLSX.getcellrange(s, XLSX.SheetColumnRange("Sheet1!B:B")) == [XLSX.Cell(XLSX.get_workbook(f), XLSX.CellRef("B1"), "", "", "3", "", false); XLSX.Cell(XLSX.get_workbook(f), XLSX.CellRef("B2"), "", "", "4", "", false); XLSX.Cell(XLSX.get_workbook(f), XLSX.CellRef("B3"), "", "", "5", "", false);;]
    @test XLSX.getcellrange(s, "Sheet1!2:2") == [XLSX.Cell(XLSX.get_workbook(f), XLSX.CellRef("A2"), "", "", "3", "", false) XLSX.Cell(XLSX.get_workbook(f), XLSX.CellRef("B2"), "", "", "4", "", false) XLSX.Cell(XLSX.get_workbook(f), XLSX.CellRef("C2"), "", "", "5", "", false)]
    @test XLSX.getcellrange(f, "Sheet1!2:2") == [XLSX.Cell(XLSX.get_workbook(f), XLSX.CellRef("A2"), "", "", "3", "", false) XLSX.Cell(XLSX.get_workbook(f), XLSX.CellRef("B2"), "", "", "4", "", false) XLSX.Cell(XLSX.get_workbook(f), XLSX.CellRef("C2"), "", "", "5", "", false)]
    @test XLSX.getcellrange(s, XLSX.SheetRowRange("Sheet1!2:2")) == [XLSX.Cell(XLSX.get_workbook(f), XLSX.CellRef("A2"), "", "", "3", "", false) XLSX.Cell(XLSX.get_workbook(f), XLSX.CellRef("B2"), "", "", "4", "", false) XLSX.Cell(XLSX.get_workbook(f), XLSX.CellRef("C2"), "", "", "5", "", false)]

    XLSX.addDefinedName(f, "MyName1", "Sheet1!A1")
    XLSX.addDefinedName(s, "MyName2", "Sheet1!A2:A3")
    XLSX.addDefinedName(f, "MyName3", "Sheet1!A2,Sheet1!A3")
    s["MyName1"] = 12.9
    @test s["MyName1"] == 12.9
    s["MyName2"] = 42
    @test s["MyName2"] == [42; 42;;]
    @test XLSX.getcell(s, "MyName1") == XLSX.Cell(XLSX.get_workbook(f), XLSX.CellRef("A1"), "", "", "12.9", "", false)
    @test XLSX.getcell(s, "MyName2") == [XLSX.Cell(XLSX.get_workbook(f), XLSX.CellRef("A2"), "", "", "42", "", false); XLSX.Cell(XLSX.get_workbook(f), XLSX.CellRef("A3"), "", "", "42", "", false);;]
    @test XLSX.getcell(f, "MyName1") == XLSX.Cell(XLSX.get_workbook(f), XLSX.CellRef("A1"), "", "", "12.9", "", false)
    @test XLSX.getcellrange(s, "MyName2") == [XLSX.Cell(XLSX.get_workbook(f), XLSX.CellRef("A2"), "", "", "42", "", false); XLSX.Cell(XLSX.get_workbook(f), XLSX.CellRef("A3"), "", "", "42", "", false);;]
    @test XLSX.getcellrange(s, "MyName3") == [[XLSX.Cell(XLSX.get_workbook(f), XLSX.CellRef("A2"), "", "", "42", "", false);;], [XLSX.Cell(XLSX.get_workbook(f), XLSX.CellRef("A3"), "", "", "42", "", false);;]]
    @test XLSX.getcellrange(f, "MyName3") == [[XLSX.Cell(XLSX.get_workbook(f), XLSX.CellRef("A2"), "", "", "42", "", false);;], [XLSX.Cell(XLSX.get_workbook(f), XLSX.CellRef("A3"), "", "", "42", "", false);;]]

end

@testset "Time and DateTime" begin
    @test XLSX.excel_value_to_time(0.82291666666666663) == Dates.Time(Dates.Hour(19), Dates.Minute(45))
    @test XLSX.time_to_excel_value(XLSX.excel_value_to_time(0.2)) == 0.2
    @test XLSX.excel_value_to_datetime(43206.805447106482, false) == Date(2018, 4, 16) + Dates.Time(Dates.Hour(19), Dates.Minute(19), Dates.Second(51))
    @test XLSX.excel_value_to_datetime(XLSX.datetime_to_excel_value(Date(2018, 4, 16) + Dates.Time(Dates.Hour(19), Dates.Minute(19), Dates.Second(51)), false), false) == Date(2018, 4, 16) + Dates.Time(Dates.Hour(19), Dates.Minute(19), Dates.Second(51))

    dt = Date(2018, 4, 1)
    @test XLSX.excel_value_to_date(XLSX.date_to_excel_value(dt, false), false) == dt
    @test XLSX.excel_value_to_date(XLSX.date_to_excel_value(dt, true), true) == dt
end

@testset "number formats" begin
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

@testset "Error Values" begin
    f = XLSX.openxlsx(joinpath(data_directory, "Errors.xlsx"), mode="rw")
    s=f[1]
    @test Int64(XLSX.getcell(s, "A1").value) == 1
    @test Int64(XLSX.getcell(s, "B1").value) == 2
    @test Int64(XLSX.getcell(s, "C1").value) == 3
    @test Int64(XLSX.getcell(s, "D1").value) == 4
    @test Int64(XLSX.getcell(s, "E1").value) == 5
    @test Int64(XLSX.getcell(s, "F1").value) == 6
    @test Int64(XLSX.getcell(s, "G1").value) == 7
    @test Int64(XLSX.getcell(s, "H1").value) == 3

    XLSX.writexlsx("mytest.xlsx", f, overwrite=true)

    f2 = XLSX.openxlsx("mytest.xlsx", mode="rw")
    s=f[1]
    @test Int64(XLSX.getcell(s, "A1").value) == 1
    @test Int64(XLSX.getcell(s, "B1").value) == 2
    @test Int64(XLSX.getcell(s, "C1").value) == 3
    @test Int64(XLSX.getcell(s, "D1").value) == 4
    @test Int64(XLSX.getcell(s, "E1").value) == 5
    @test Int64(XLSX.getcell(s, "F1").value) == 6
    @test Int64(XLSX.getcell(s, "G1").value) == 7
    @test Int64(XLSX.getcell(s, "H1").value) == 3

    isfile("mytest.xlsx") && rm("mytest.xlsx")

    f = XLSX.readxlsx(joinpath(data_directory, "Errors.xlsx"))
    sh = f[1]
    @test XLSX.iserror(sh, "A1") == true
    @test XLSX.iserror(sh, 1, 1) == true
    @test XLSX.iserror(sh, "I1") == false
    @test XLSX.iserror(sh, 1, 9) == false
    @test XLSX.iserror(sh, "A1:I1") == [true true true true true true true true false]
    @test XLSX.iserror(sh, 1, 1:9) == [true true true true true true true true false]
    @test XLSX.iserror(sh, "A1:B1,D1:E1") == [[true true], [true true]]
    @test XLSX.iserror(sh, 1, [1, 2, 4, 5]) == [true, true, true, true]
    @test XLSX.iserror(sh, :) == Bool[
                                        1 1 1 1 1 1 1 1
                                        0 0 0 0 0 0 0 0
                                        0 0 0 0 0 0 0 0
                                        0 0 0 0 0 0 0 0
                                        0 0 0 0 0 0 0 0
                                     ]
    @test XLSX.geterror(sh, "A1") == "#NULL!"
    @test XLSX.geterror(sh, 1, 1) == "#NULL!"
    @test XLSX.geterror(sh, "I1") == ""
    @test XLSX.geterror(sh, 1, 9) == ""
    @test XLSX.geterror(s, "A1:I1") == ["#NULL!"  "#DIV/0!"  "#VALUE!"  "#REF!"  "#NAME?"  "#NUM!"  "#N/A"  "#VALUE!"  ""]
    @test XLSX.geterror(s, 1, 1:9) == ["#NULL!"  "#DIV/0!"  "#VALUE!"  "#REF!"  "#NAME?"  "#NUM!"  "#N/A"  "#VALUE!"  ""]
    @test XLSX.geterror(sh, "A1:B1,D1:E1") == [["#NULL!" "#DIV/0!"], ["#REF!" "#NAME?"]]
    @test XLSX.geterror(sh, 1, [1, 2, 4, 5]) == ["#NULL!", "#DIV/0!", "#REF!", "#NAME?"]
    @test XLSX.geterror(sh, :) == [ "#NULL!"  "#DIV/0!"  "#VALUE!"  "#REF!"  "#NAME?"  "#NUM!"  "#N/A"  "#VALUE!"
                                    ""        ""         ""         ""       ""        ""       ""      ""
                                    ""        ""         ""         ""       ""        ""       ""      ""
                                    ""        ""         ""         ""       ""        ""       ""      ""
                                    ""        ""         ""         ""       ""        ""       ""      ""
                                 ]

end

@testset "No Dimension" begin
    noDim = XLSX.openxlsx(joinpath(data_directory, "NoDim.xlsx"), mode="rw") # This file is the same as customXml but has the dimension nodes removed.
    Dim = XLSX.readxlsx(joinpath(data_directory, "customXml.xlsx"))
    @test noDim[1].dimension == Dim[1].dimension
    @test noDim[2].dimension == Dim[2].dimension

    f = XLSX.newxlsx()
    s = f[1]
    for i = 10:20, j = 10:20
        s[i, j] = i + j
    end
    XLSX.set_dimension!(s, XLSX.CellRange(XLSX.CellRef("J10"), XLSX.CellRef("T20")))
    @test XLSX.get_dimension(s) == XLSX.CellRange(XLSX.CellRef("J10"), XLSX.CellRef("T20"))
    s["A1"] = 2
    @test XLSX.get_dimension(s) == XLSX.CellRange(XLSX.CellRef("A1"), XLSX.CellRef("T20"))

end

@testset "Range intersect" begin
    @test XLSX.intersects(XLSX.CellRange(XLSX.CellRef("A3"), XLSX.CellRef("E4")), XLSX.CellRange(XLSX.CellRef("C1"), XLSX.CellRef("D6")))
    @test XLSX.intersects(XLSX.CellRange(XLSX.CellRef("C1"), XLSX.CellRef("D6")), XLSX.CellRange(XLSX.CellRef("A3"), XLSX.CellRef("E4")))
    @test XLSX.intersects(XLSX.CellRange(XLSX.CellRef("A3"), XLSX.CellRef("D4")), XLSX.CellRange(XLSX.CellRef("C1"), XLSX.CellRef("E6")))
    @test XLSX.intersects(XLSX.CellRange(XLSX.CellRef("C1"), XLSX.CellRef("E6")), XLSX.CellRange(XLSX.CellRef("A3"), XLSX.CellRef("D4")))
    @test XLSX.intersects(XLSX.CellRange(XLSX.CellRef("C3"), XLSX.CellRef("E4")), XLSX.CellRange(XLSX.CellRef("C1"), XLSX.CellRef("G6")))
    @test XLSX.intersects(XLSX.CellRange(XLSX.CellRef("C1"), XLSX.CellRef("G6")), XLSX.CellRange(XLSX.CellRef("C3"), XLSX.CellRef("E4")))
    @test XLSX.intersects(XLSX.CellRange(XLSX.CellRef("A1"), XLSX.CellRef("E4")), XLSX.CellRange(XLSX.CellRef("C3"), XLSX.CellRef("G6")))
    @test XLSX.intersects(XLSX.CellRange(XLSX.CellRef("C3"), XLSX.CellRef("G6")), XLSX.CellRange(XLSX.CellRef("A1"), XLSX.CellRef("E4")))
    @test XLSX.intersects(XLSX.CellRange(XLSX.CellRef("C3"), XLSX.CellRef("D4")), XLSX.CellRange(XLSX.CellRef("A1"), XLSX.CellRef("E6")))
end

@testset "Column Range" begin

    @testset "Single Column" begin
        c = XLSX.ColumnRange("C")
        @test c.start == 3
        @test c.stop == 3
        show(IOBuffer(), c)

        c = XLSX.ColumnRange("AA")
        @test c.start == 27
        @test c.stop == 27
    end

    @testset "Multiple Columns" begin
        c = XLSX.ColumnRange("A:Z")
        @test c.start == 1
        @test c.stop == 26

        c = XLSX.ColumnRange("A:AA")
        @test c.start == 1
        @test c.stop == 27

        cr = XLSX.ColumnRange("B:D")
        @test string(cr) == "B:D"
        @test cr.start == 2
        @test cr.stop == 4
        @test length(cr) == 3

        cr = XLSX.ColumnRange("B")
        @test string(cr) == "B:B"
        @test cr.start == 2
        @test cr.stop == 2
        @test length(cr) == 1

        @test XLSX.ColumnRange("B1:D3") == XLSX.ColumnRange("B:D")

        @test_throws XLSX.XLSXError XLSX.ColumnRange("D:A")
        @test collect(XLSX.ColumnRange("B:D")) == ["B", "C", "D"]
        @test XLSX.ColumnRange("B:D") == XLSX.ColumnRange("B:D")
        @test hash(XLSX.ColumnRange("B:D")) == hash(XLSX.ColumnRange("B:D"))
    end
end

@testset "Row Range" begin # Issue #150
    cr = XLSX.RowRange("2:5")
    @test string(cr) == "2:5"
    @test cr.start == 2
    @test cr.stop == 5
    @test length(cr) == 4
    @test collect(cr) == ["2", "3", "4", "5"]

    cr = XLSX.RowRange("2")
    @test string(cr) == "2:2"
    @test cr.start == 2
    @test cr.stop == 2
    @test length(cr) == 1
    @test collect(cr) == ["2"]

    @test XLSX.RowRange("B1:D3") == XLSX.RowRange("1:3")
    @test_throws XLSX.XLSXError XLSX.RowRange("5:2")
    @test XLSX.RowRange("2:5") == XLSX.RowRange("2:5")
    @test hash(XLSX.RowRange("2:5")) == hash(XLSX.RowRange("2:5"))
end

@testset "Non-contiguous Range" begin
    cr = XLSX.NonContiguousRange("Sheet1!D1:D3,Sheet1!B1:B3")
    @test string(cr) == "Sheet1!D1:D3,Sheet1!B1:B3"
    @test cr.sheet == "Sheet1"
    @test cr.rng == [XLSX.CellRange("D1:D3"), XLSX.CellRange("B1:B3")]
    @test length(cr) == 6
    @test length(XLSX.NonContiguousRange("Sheet1!B1:B1,Sheet1!B1")) == 1
    @test collect(cr.rng) == [XLSX.CellRange("D1:D3"), XLSX.CellRange("B1:B3")]
    @test XLSX.NonContiguousRange("Sheet1!D1:D3,Sheet1!B1:B3") == XLSX.NonContiguousRange("Sheet1!D1:D3,Sheet1!B1:B3")
    @test hash(XLSX.NonContiguousRange("Sheet1!D1:D3,Sheet1!B1:B3")) == hash(XLSX.NonContiguousRange("Sheet1!D1:D3,Sheet1!B1:B3"))

    f = XLSX.newxlsx("Sheet 1")
    s = f["Sheet 1"]
    for cell in XLSX.CellRange("A1:D6")
        s[cell] = ""
    end
    cr = XLSX.NonContiguousRange(s, "D1:D3,A2,B1:B3")
    @test string(cr) == "'Sheet 1'!D1:D3,'Sheet 1'!A2,'Sheet 1'!B1:B3"
    @test cr.sheet == "Sheet 1"
    @test cr.rng == [XLSX.CellRange("D1:D3"), XLSX.CellRef("A2"), XLSX.CellRange("B1:B3")]
    @test length(cr) == 7
    @test length(XLSX.NonContiguousRange(s, "B1:B1,B1")) == 1
    @test collect(cr.rng) == [XLSX.CellRange("D1:D3"), XLSX.CellRef("A2"), XLSX.CellRange("B1:B3")]
    @test XLSX.NonContiguousRange(s, "D1:D3,A2,B1:B3") == XLSX.NonContiguousRange(s, "D1:D3,A2,B1:B3")
    @test hash(XLSX.NonContiguousRange(s, "D1:D3,A2,B1:B3")) == hash(XLSX.NonContiguousRange(s, "D1:D3,A2,B1:B3"))

    @test_throws XLSX.XLSXError XLSX.NonContiguousRange("Sheet1!D1:D3,B1:B3")
    @test_throws XLSX.XLSXError XLSX.NonContiguousRange("Sheet1!D1:D3,Sheet2!B1:B3")
    @test_throws XLSX.XLSXError XLSX.NonContiguousRange("B1:D3")
    @test_throws XLSX.XLSXError XLSX.NonContiguousRange("2:5")
end

@testset "CellRange iterator" begin
    rng = XLSX.CellRange("A2:C4")
    @test collect(rng) == [XLSX.CellRef("A2"), XLSX.CellRef("B2"), XLSX.CellRef("C2"), XLSX.CellRef("A3"), XLSX.CellRef("B3"), XLSX.CellRef("C3"), XLSX.CellRef("A4"), XLSX.CellRef("B4"), XLSX.CellRef("C4")]
end

# Checks whether `data` equals `test_data`
function check_test_data(data::Vector{S}, test_data::Vector{T}) where {S,T}

    @test length(data) == length(test_data)

    function size_of_data(d::Vector{T}) where {T}
        isempty(d) && return (0, 0)
        return length(d[1]), length(d)
    end

    rows, cols = size_of_data(test_data)

    for col in 1:cols
        @test length(data[col]) == length(test_data[col])
    end

    for row in 1:rows, col in 1:cols
        test_value = test_data[col][row]
        value = data[col][row]

        if test_value === nothing
            @test ismissing(value)
        elseif ismissing(test_value) || (isa(test_value, AbstractString) && isempty(test_value))
            @test ismissing(value) || (isa(value, AbstractString) && isempty(value))
        else
            if isa(test_value, Integer) || isa(value, Integer)
                @test isa(test_value, Integer)
                @test isa(value, Integer)
            end

            if isa(test_value, Real) && !isa(test_value, Integer)
                @test isapprox(value, test_value)
            else
                @test value == test_value
            end
        end
    end

    nothing
end

@testset "Strict formats" begin
    @testset "Simple" begin
        XLSX.openxlsx(joinpath(data_directory, "strict.xlsx")) do f
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

        XLSX.openxlsx(joinpath(data_directory, "strict.xlsx")) do xf
            empty_sheet = XLSX.getsheet(xf, "empty")
            @test_throws XLSX.XLSXError XLSX.gettable(empty_sheet)
            itr = XLSX.eachrow(empty_sheet)
            @test_throws XLSX.XLSXError XLSX.find_row(itr, 1)
            @test_throws XLSX.XLSXError XLSX.getsheet(xf, "invalid_sheet")
        end

        f = XLSX.readxlsx(joinpath(data_directory, "strict.xlsx"))
        tb5 = f["table5"]
        test_data = Vector{Any}(undef, 1)
        test_data[1] = [1, 2, 3, 4, 5]
        dtable = XLSX.gettable(tb5)
        data, col_names = dtable.data, dtable.column_labels
        @test col_names == [:HEADER]
        check_test_data(data, test_data)
        tb6 = f["table6"]
        dtable = XLSX.gettable(tb6, first_row=3)
        data, col_names = dtable.data, dtable.column_labels
        @test col_names == [:HEADER]
        check_test_data(data, test_data)
        tb7 = f["table7"]
        dtable = XLSX.gettable(tb7, first_row=3)
        data, col_names = dtable.data, dtable.column_labels
        @test col_names == [:HEADER]
        check_test_data(data, test_data)

        sheet_lookup = f["lookup"]
        test_data = Vector{Any}(undef, 3)
        test_data[1] = [10, 20, 30]
        test_data[2] = ["name1", "name2", "name3"]
        test_data[3] = [100, 200, 300]
        dtable = XLSX.gettable(sheet_lookup)
        data, col_names = dtable.data, dtable.column_labels
        @test col_names == [:ID, :NAME, :VALUE]
        check_test_data(data, test_data)

        header_error_sheet = f["header_error"]
        dtable = XLSX.gettable(header_error_sheet, "B:E")
        data, col_names = dtable.data, dtable.column_labels
        @test col_names == [:COLUMN_A, :COLUMN_B, Symbol("COLUMN_A_2"), Symbol("#Empty")]

        XLSX.writexlsx("mytest.xlsx", XLSX.openxlsx(joinpath(data_directory, "strict.xlsx"); mode="rw"), overwrite=true)

        XLSX.openxlsx("mytest.xlsx") do f
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

        XLSX.openxlsx("mytest.xlsx") do xf
            empty_sheet = XLSX.getsheet(xf, "empty")
            @test_throws XLSX.XLSXError XLSX.gettable(empty_sheet)
            itr = XLSX.eachrow(empty_sheet)
            @test_throws XLSX.XLSXError XLSX.find_row(itr, 1)
            @test_throws XLSX.XLSXError XLSX.getsheet(xf, "invalid_sheet")
        end

        f = XLSX.readxlsx("mytest.xlsx")
        tb5 = f["table5"]
        test_data = Vector{Any}(undef, 1)
        test_data[1] = [1, 2, 3, 4, 5]
        dtable = XLSX.gettable(tb5)
        data, col_names = dtable.data, dtable.column_labels
        @test col_names == [:HEADER]
        check_test_data(data, test_data)
        tb6 = f["table6"]
        dtable = XLSX.gettable(tb6, first_row=3)
        data, col_names = dtable.data, dtable.column_labels
        @test col_names == [:HEADER]
        check_test_data(data, test_data)
        tb7 = f["table7"]
        dtable = XLSX.gettable(tb7, first_row=3)
        data, col_names = dtable.data, dtable.column_labels
        @test col_names == [:HEADER]
        check_test_data(data, test_data)

        sheet_lookup = f["lookup"]
        test_data = Vector{Any}(undef, 3)
        test_data[1] = [10, 20, 30]
        test_data[2] = ["name1", "name2", "name3"]
        test_data[3] = [100, 200, 300]
        dtable = XLSX.gettable(sheet_lookup)
        data, col_names = dtable.data, dtable.column_labels
        @test col_names == [:ID, :NAME, :VALUE]
        check_test_data(data, test_data)

        header_error_sheet = f["header_error"]
        dtable = XLSX.gettable(header_error_sheet, "B:E")
        data, col_names = dtable.data, dtable.column_labels
        @test col_names == [:COLUMN_A, :COLUMN_B, Symbol("COLUMN_A_2"), Symbol("#Empty")]

        isfile("mytest.xlsx") && rm("mytest.xlsx")
    end

    @testset "With chartsheet" begin # From Issue #233
        f = XLSX.openxlsx(joinpath(data_directory, "Strict-foo.xlsx"), mode="rw")
        Expected = """            sheetname size          range        
            -------------------------------------------------
                         Tabelle1 13x4          A2:D14       
                        Diagramm1 Chartsheet\n"""
        result = sprint(show, f)
        idx = findfirst(==('\n'), result)
        after = result[idx+1:end]
        @test after == Expected
        @test sprint(show, f[1]) == "13×4 XLSX.Worksheet: [\"Tabelle1\"](A2:D14) "
        @test sprint(show, f[2]) == "Chartsheet: [\"Diagramm1\"] "
        @test_throws XLSX.XLSXError XLSX.copysheet!(f["Diagramm1"], "Diagramm1_copy")
        @test_throws XLSX.XLSXError XLSX.deletesheet!(f["Diagramm1"])
        @test_throws XLSX.XLSXError XLSX.gettable(f["Diagramm1"])
        @test_throws XLSX.XLSXError XLSX.gettable(f["Diagramm1"], "A:B")
        XLSX.writexlsx("mytest.xlsx", f, overwrite=true)

        XLSX.openxlsx("mytest.xlsx") do f
            result = sprint(show, f)
            idx = findfirst(==('\n'), result)
            after = result[idx+1:end]
            @test after == Expected
            @test sprint(show, f[1]) == "13×4 XLSX.Worksheet: [\"Tabelle1\"](A2:D14) "
            @test sprint(show, f[2]) == "Chartsheet: [\"Diagramm1\"] "
        end
        isfile("mytest.xlsx") && rm("mytest.xlsx")
    end


end

@testset "Table" begin

    @test Tables.istable(XLSX.DataTable)

    @testset "Index" begin
        index = XLSX.Index("A:B", ["First", "Second"])
        @test index.column_labels == [:First, :Second]
        @test index.lookup[:First] == 1
        @test index.lookup[:Second] == 2
    end

    @testset "Bounds" begin
        f = XLSX.readxlsx(joinpath(data_directory, "book_sparse.xlsx"))
        s = f["Sheet1"]

        report = Vector{String}()
        for r in XLSX.eachrow(s)
            if !isempty(r)
                push!(report, string(XLSX.row_number(r), " - ", XLSX.column_bounds(r)))

                if XLSX.row_number(r) == 2
                    @test XLSX.last_column_index(r, 2) == 2
                elseif XLSX.row_number(r) == 3
                    @test XLSX.last_column_index(r, 3) == 4
                elseif XLSX.row_number(r) == 6
                    @test XLSX.last_column_index(r, 1) == 4
                    @test XLSX.last_column_index(r, 2) == 4
                    @test XLSX.last_column_index(r, 3) == 4
                    @test XLSX.last_column_index(r, 4) == 4
                    @test_throws XLSX.XLSXError XLSX.last_column_index(r, 5)
                elseif XLSX.row_number(r) == 9
                    @test XLSX.last_column_index(r, 2) == 3
                    @test XLSX.last_column_index(r, 3) == 3
                    @test XLSX.last_column_index(r, 5) == 5
                end
            end
        end

        @test report == ["2 - (2, 2)", "3 - (3, 4)", "6 - (1, 4)", "9 - (2, 5)"]
    end

    XLSX.openxlsx(joinpath(data_directory, "general.xlsx")) do f
        f["general"][:]
    end

    f = XLSX.readxlsx(joinpath(data_directory, "general.xlsx"))
    s = f["table"]
    s[:]
    dtable = XLSX.gettable(s)
    data, col_names = dtable.data, dtable.column_labels
    @test col_names == [Symbol("Column B"), Symbol("Column C"), Symbol("Column D"), Symbol("Column E"), Symbol("Column F"), Symbol("Column G")]

    test_data = Vector{Any}(undef, 6)
    test_data[1] = collect(1:8)
    test_data[2] = ["Str1", missing, "Str1", "Str1", "Str2", "Str2", "Str2", "Str2"]
    test_data[3] = [Date(2018, 4, 21) + Dates.Day(i) for i in 0:7]
    test_data[4] = [missing, missing, missing, missing, missing, "a", "b", missing]
    test_data[5] = [0.2001132319, 0.2793987377, 0.0950591677, 0.0744023067, 0.8242278091, 0.6205883578, 0.9174151018, 0.6749604883]
    test_data[6] = [missing for i in 1:8]

    check_test_data(data, test_data)

    @test XLSX.infer_eltype(data[1]) == Int64
    @test XLSX.infer_eltype(data[2]) == Union{Missing,String}
    @test XLSX.infer_eltype(data[3]) == Date
    @test XLSX.infer_eltype(data[4]) == Union{Missing,String}
    @test XLSX.infer_eltype(data[5]) == Float64
    @test XLSX.infer_eltype(data[6]) == Any
    @test XLSX.infer_eltype(Vector{Int}()) == Int
    @test XLSX.infer_eltype(Vector{Float64}()) == Float64
    @test XLSX.infer_eltype(Vector{Any}()) == Any
    @test XLSX.infer_eltype([1, "1", 10.2]) == Any
    @test XLSX.infer_eltype([1, 10]) == Int
    @test XLSX.infer_eltype([1.0, 10.0]) == Float64
    @test XLSX.infer_eltype([1, 10.2]) == Float64 # Promote mixed int/float columns to float (#192)

    dtable_inferred = XLSX.gettable(s, infer_eltypes=true)
    data_inferred, col_names = dtable_inferred.data, dtable_inferred.column_labels
    @test eltype(data_inferred[1]) == Int64
    @test eltype(data_inferred[2]) == Union{Missing,String}
    @test eltype(data_inferred[3]) == Date
    @test eltype(data_inferred[4]) == Union{Missing,String}
    @test eltype(data_inferred[5]) == Float64
    @test eltype(data_inferred[6]) == Any

    function stop_function(r::XLSX.TableRow)
        v = r[Symbol("Column C")]
        return !ismissing(v) && v == "Str2"
    end

    function never_reaches_stop(r::XLSX.TableRow)
        v = r[Symbol("Column C")]
        return !ismissing(v) && v == "never was found"
    end

    dtable = XLSX.gettable(s, stop_in_row_function=never_reaches_stop)
    data, col_names = dtable.data, dtable.column_labels
    check_test_data(data, test_data)

    dtable = XLSX.gettable(s, stop_in_row_function=stop_function)
    data, col_names = dtable.data, dtable.column_labels
    @test col_names == [Symbol("Column B"), Symbol("Column C"), Symbol("Column D"), Symbol("Column E"), Symbol("Column F"), Symbol("Column G")]

    test_data = Vector{Any}(undef, 6)
    test_data[1] = collect(1:4)
    test_data[2] = ["Str1", missing, "Str1", "Str1"]
    test_data[3] = [Date(2018, 4, 21) + Dates.Day(i) for i in 0:3]
    test_data[4] = [missing, missing, missing, missing]
    test_data[5] = [0.2001132319, 0.2793987377, 0.0950591677, 0.0744023067]
    test_data[6] = [missing for i in 1:4]

    check_test_data(data, test_data)

    dtable = XLSX.gettable(s, stop_in_empty_row=false)
    data, col_names = dtable.data, dtable.column_labels
    @test col_names == [Symbol("Column B"), Symbol("Column C"), Symbol("Column D"), Symbol("Column E"), Symbol("Column F"), Symbol("Column G")]

    # test keep_empty_rows
    for (stop_in_empty_row, keep_empty_rows, n_rows) in [
        (false, false, 9),
        (false, true, 11),
        (true, false, 8),
        (true, true, 8)
    ]
        dtable = XLSX.gettable(s; stop_in_empty_row=stop_in_empty_row, keep_empty_rows=keep_empty_rows)
        @test all(col_name -> length(Tables.getcolumn(dtable, col_name)) == n_rows, Tables.columnnames(dtable))
    end

    test_data = Vector{Any}(undef, 6)
    test_data[1] = [1, 2, 3, 4, 5, 6, 7, 8, "trash"]
    test_data[2] = ["Str1", missing, "Str1", "Str1", "Str2", "Str2", "Str2", "Str2", missing]
    test_data[3] = Any[Date(2018, 4, 21) + Dates.Day(i) for i in 0:7]
    push!(test_data[3], "trash")
    test_data[4] = [missing, missing, missing, missing, missing, "a", "b", missing, missing]
    test_data[5] = [0.2001132319, 0.2793987377, 0.0950591677, 0.0744023067, 0.8242278091, 0.6205883578, 0.9174151018, 0.6749604883, "trash"]
    test_data[6] = Any[missing for i in 1:8]
    push!(test_data[6], "trash")

    check_test_data(data, test_data)

    # queries based on ColumnRange
    x = XLSX.getcellrange(s, XLSX.ColumnRange("B:D"))
    @test size(x) == (12, 3)
    y = XLSX.getcellrange(s, "B:D")
    @test size(y) == (12, 3)
    @test x == y
    @test_throws XLSX.XLSXError XLSX.getcellrange(s, "D:B")
    @test_throws XLSX.XLSXError XLSX.getcellrange(s, "A:C1")

    d = XLSX.getdata(s, "B:D")
    @test size(d) == (12, 3)
    @test_throws XLSX.XLSXError XLSX.getdata(s, "A:C1")
    @test d[1, 1] == "Column B"
    @test d[1, 2] == "Column C"
    @test d[1, 3] == "Column D"
    @test d[9, 1] == 8
    @test d[9, 2] == "Str2"
    @test d[9, 3] == Date(2018, 4, 28)
    @test d[11, 1] == "trash"
    @test ismissing(d[11, 2])
    @test d[11, 3] == "trash"
    @test ismissing(d[12, 1])
    @test ismissing(d[12, 2])
    @test ismissing(d[12, 3])

    d1 = XLSX.getdata(s, "2:3")
    @test size(d1) == (2, 8)
    @test d1[1, 2] == "Column B"
    @test d1[1, 4] == "Column D"
    @test d1[2, 2] == 1
    @test d1[2, 4] == Date(2018, 4, 21)

    d2 = f["table!B:D"]
    @test size(d) == size(d2)
    @test all(d .=== d2)

    @test_throws XLSX.XLSXError f["table!B1:D"]
    @test_throws XLSX.XLSXError f["table!D:B"]

    s = f["table2"]
    test_data = Vector{Any}(undef, 3)
    test_data[1] = ["A1", "A2", "A3", missing]
    test_data[2] = ["B1", "B2", missing, "B4"]
    test_data[3] = [missing, missing, missing, missing]

    dtable = XLSX.gettable(s)
    data, col_names = dtable.data, dtable.column_labels

    @test col_names == [:HA, :HB, :HC]
    check_test_data(data, test_data)

    for (ri, rowdata) in enumerate(XLSX.eachtablerow(s))
        if ismissing(test_data[1][ri])
            @test ismissing(rowdata[:HA])
        else
            @test rowdata[:HA] == test_data[1][ri]
        end

        @test XLSX.table_columns_count(rowdata) == 3
        @test XLSX.row_number(rowdata) == ri
        @test XLSX.get_column_labels(rowdata) == col_names
        @test XLSX.get_column_label(rowdata, 1) == :HA
        @test XLSX.get_column_label(rowdata, 2) == :HB
        @test XLSX.get_column_label(rowdata, 3) == :HC

        @test_throws XLSX.XLSXError XLSX.getdata(rowdata, :INVALID_COLUMN)
    end

    override_col_names_strs = ["ColumnA", "ColumnB", "ColumnC"]
    override_col_names = [Symbol(i) for i in override_col_names_strs]

    dtable = XLSX.gettable(s, column_labels=override_col_names_strs)
    data, col_names = dtable.data, dtable.column_labels

    @test col_names == override_col_names
    check_test_data(data, test_data)

    dtable = XLSX.gettable(s, "A:B", first_row=1)
    data, col_names = dtable.data, dtable.column_labels
    test_data_AB_cols = Vector{Any}(undef, 2)
    test_data_AB_cols[1] = test_data[1]
    test_data_AB_cols[2] = test_data[2]
    @test col_names == [:HA, :HB]
    check_test_data(data, test_data_AB_cols)

    dtable = XLSX.gettable(s, "A:B")
    data, col_names = dtable.data, dtable.column_labels
    test_data_AB_cols = Vector{Any}(undef, 2)
    test_data_AB_cols[1] = test_data[1]
    test_data_AB_cols[2] = test_data[2]
    @test col_names == [:HA, :HB]
    check_test_data(data, test_data_AB_cols)

    dtable = XLSX.gettable(s, "B:B", first_row=2)
    data, col_names = dtable.data, dtable.column_labels
    @test col_names == [:B1]
    @test length(data) == 1
    @test length(data[1]) == 1
    @test data[1][1] == "B2"

    dtable = XLSX.gettable(s, "B:C")
    data, col_names = dtable.data, dtable.column_labels
    @test col_names == [:HB, :HC]
    test_data_BC_cols = Vector{Any}(undef, 2)
    test_data_BC_cols[1] = ["B1", "B2"]
    test_data_BC_cols[2] = [missing, missing]
    check_test_data(data, test_data_BC_cols)

    dtable = XLSX.gettable(s, "B:C", first_row=2, header=false)
    data, col_names = dtable.data, dtable.column_labels
    @test col_names == [:B, :C]
    check_test_data(data, test_data_BC_cols)

    s = f["table3"]
    test_data = Vector{Any}(undef, 3)
    test_data[1] = [missing, missing, "B5"]
    test_data[2] = ["C3", missing, missing]
    test_data[3] = [missing, "D4", missing]
    dtable = XLSX.gettable(s)
    data, col_names = dtable.data, dtable.column_labels
    @test col_names == [:H1, :H2, :H3]
    check_test_data(data, test_data)
    @test_throws XLSX.XLSXError XLSX.find_row(XLSX.eachrow(s), 20)

    for r in XLSX.eachrow(s)
        @test isempty(XLSX.getcell(r, "A"))
        @test XLSX.getdata(s, XLSX.getcell(r, "B")) == "H1"
        @test r[2] == "H1"
        @test r["B"] == "H1"
        break
    end

    @test XLSX._find_first_row_with_data(s, 5) == 5
    @test_throws XLSX.XLSXError XLSX._find_first_row_with_data(s, 7)

    s = f["table4"]
    dtable = XLSX.gettable(s)
    data, col_names = dtable.data, dtable.column_labels
    @test col_names == [:H1, :H2, :H3]
    check_test_data(data, test_data)

    @testset "empty/invalid" begin
        XLSX.openxlsx(joinpath(data_directory, "general.xlsx")) do xf
            empty_sheet = XLSX.getsheet(xf, "empty")
            @test_throws XLSX.XLSXError XLSX.gettable(empty_sheet)
            itr = XLSX.eachrow(empty_sheet)
            @test_throws XLSX.XLSXError XLSX.find_row(itr, 1)
            @test_throws XLSX.XLSXError XLSX.getsheet(xf, "invalid_sheet")
        end
    end

    @testset "sheets 6/7/lookup/header_error" begin
        f = XLSX.readxlsx(joinpath(data_directory, "general.xlsx"))
        tb5 = f["table5"]
        test_data = Vector{Any}(undef, 1)
        test_data[1] = [1, 2, 3, 4, 5]
        dtable = XLSX.gettable(tb5)
        data, col_names = dtable.data, dtable.column_labels
        @test col_names == [:HEADER]
        check_test_data(data, test_data)
        tb6 = f["table6"]
        dtable = XLSX.gettable(tb6, first_row=3)
        data, col_names = dtable.data, dtable.column_labels
        @test col_names == [:HEADER]
        check_test_data(data, test_data)
        tb7 = f["table7"]
        dtable = XLSX.gettable(tb7, first_row=3)
        data, col_names = dtable.data, dtable.column_labels
        @test col_names == [:HEADER]
        check_test_data(data, test_data)

        sheet_lookup = f["lookup"]
        test_data = Vector{Any}(undef, 3)
        test_data[1] = [10, 20, 30]
        test_data[2] = ["name1", "name2", "name3"]
        test_data[3] = [100, 200, 300]
        dtable = XLSX.gettable(sheet_lookup)
        data, col_names = dtable.data, dtable.column_labels
        @test col_names == [:ID, :NAME, :VALUE]
        check_test_data(data, test_data)

        header_error_sheet = f["header_error"]
        dtable = XLSX.gettable(header_error_sheet, "B:E")
        data, col_names = dtable.data, dtable.column_labels
        @test col_names == [:COLUMN_A, :COLUMN_B, Symbol("COLUMN_A_2"), Symbol("#Empty")]
    end

    @testset "Consecutive passes" begin
        # Consecutive passes should yield the same results
        XLSX.openxlsx(joinpath(data_directory, "general.xlsx")) do f
            sl = f["lookup"]
            dtable = XLSX.gettable(sl)
            data, col_names = dtable.data, dtable.column_labels
            @test col_names == [:ID, :NAME, :VALUE]
            check_test_data(data, test_data)

            dtable = XLSX.gettable(sl)
            data, col_names = dtable.data, dtable.column_labels
            @test col_names == [:ID, :NAME, :VALUE]
            check_test_data(data, test_data)
        end
    end

    @testset "Transposed Tables" begin
        test_data = Vector{Any}(undef, 6)
        test_data[1] = [1940, 1950, 1960, 1970, 1980, 1990, 2000, 2010, 2020]
        test_data[2] = [1, 2, 3, 4, 5, 6, 7, 8, 9]
        test_data[3] = [10, 20, 30, 40, 50, 60, 70, 80, 90]
        test_data[4] = [100, 200, 300, 400, 500, 600, 700, 800, 900]
        test_data[5] = [0.1, 0.2, 0.3, 0.4, 0.5, 0.6, 0.7, 0.8, 0.9]
        test_data[6] = Any["Hello", Date(2025, 12, 19), 3, 3.33, "Hello", Date(2025, 12, 19), 3, 3.33, true]

        f = XLSX.openxlsx(joinpath(data_directory, "HTable.xlsx"), mode="rw")
        dtable = XLSX.gettransposedtable(f[1])
        data, colnames = dtable.data, dtable.column_labels
        @test colnames == Symbol.(["Year", "Col A", "Col B", "Col C", "Col D", "Col E"])
        dtable = XLSX.gettransposedtable(f[1]; normalizenames=true)
        data, colnames = dtable.data, dtable.column_labels
        @test colnames == Symbol.(["Year", "Col_A", "Col_B", "Col_C", "Col_D", "Col_E"])
        check_test_data(data, test_data)

        dtable = XLSX.readtransposedtable(joinpath(data_directory, "HTable.xlsx"))
        data, colnames = dtable.data, dtable.column_labels
        @test colnames == Symbol.(["Year", "Col A", "Col B", "Col C", "Col D", "Col E"])
        dtable = XLSX.readtransposedtable(joinpath(data_directory, "HTable.xlsx"); normalizenames=true)
        data, colnames = dtable.data, dtable.column_labels
        @test colnames == Symbol.(["Year", "Col_A", "Col_B", "Col_C", "Col_D", "Col_E"])
        check_test_data(data, test_data)

        dtable = XLSX.readtransposedtable(joinpath(data_directory, "HTable.xlsx"), "Offset", "2:7")
        data, colnames = dtable.data, dtable.column_labels
        @test colnames == Symbol.(["Year", "Col A", "Col B", "Col C", "Col D", "Col E"])
        check_test_data(data, test_data)

        dtable = XLSX.readtransposedtable(joinpath(data_directory, "HTable.xlsx"), "Multiple", "2:7")
        data, colnames = dtable.data, dtable.column_labels
        @test colnames == Symbol.(["Year", "Col A", "Col B", "Col C", "Col D", "Col E"])
        check_test_data(data, test_data)

        @test_throws XLSX.XLSXError XLSX.readtransposedtable(joinpath(data_directory, "HTable.xlsx"), "Multiple", "2:9") # Beyond worksheet dimension
        @test_throws XLSX.XLSXError XLSX.readtransposedtable(joinpath(data_directory, "HTable.xlsx"), "Multiple", "b:d") # Invalid row range
        test_data[1] = [1840, 1841, 1842, 1843, 1844, 1845, 1846, 1847, 1848]
        test_data[2] = [12.4, 12.6, 12.8, 13.0, 13.2, 13.4, 13.6, 13.8, 14.0]
        test_data[3] = [0.045, 0.046, 0.047, 0.048, 0.049, 0.050, 0.051, 0.052, 0.053]
        test_data[4] = [true, true, false, true, false, true, true, true, false]
        test_data[5] = [Time(10, 22), Time(10, 23), Time(10, 24), Time(10, 25), Time(10, 26), Time(10, 27), Time(10, 28), Time(10, 29), Time(10, 30)]
        test_data[6] = Any["Hello", Date(2025, 12, 19), 3, 3.33, "Hello", Date(2025, 12, 19), 3, 3.33, true]
        dtable = XLSX.readtransposedtable(joinpath(data_directory, "HTable.xlsx"), "Multiple", "2:7"; first_column="M")
        data, colnames = dtable.data, dtable.column_labels
        @test colnames == Symbol.(["date", "name1", "name2", "name3", "name4", "name5"])
        check_test_data(data, test_data)

        @test_throws XLSX.XLSXError XLSX.readtransposedtable(joinpath(data_directory, "HTable.xlsx"), "Multiple", "2:7"; first_column=4.2) # Invalid first_column
        @test_throws XLSX.XLSXError XLSX.readtransposedtable(joinpath(data_directory, "HTable.xlsx"), "Multiple", "2:7"; first_column=200) # first_column beyond worksheet dimension

        dtable = XLSX.readtransposedtable(joinpath(data_directory, "HTable.xlsx"), "Multiple", "2:7"; first_column=13, column_labels=["year", "Col_2", "Col_3", "Col_4", "Col_5", "Col_6"])
        data, colnames = dtable.data, dtable.column_labels
        @test colnames == Symbol.(["year", "Col_2", "Col_3", "Col_4", "Col_5", "Col_6"])
        check_test_data(data, test_data)

        @test_throws XLSX.XLSXError XLSX.readtransposedtable(joinpath(data_directory, "HTable.xlsx"), "Multiple", "2:7"; first_column=13, column_labels=["Col_2", "Col_3", "Col_4", "Col_5", "Col_6"]) # Missing one label

        test_data[1] = Any["date", 1840, 1841, 1842, 1843, 1844, 1845, 1846, 1847, 1848]
        test_data[2] = Any["name1", 12.4, 12.6, 12.8, 13.0, 13.2, 13.4, 13.6, 13.8, 14.0]
        test_data[3] = Any["name2", 0.045, 0.046, 0.047, 0.048, 0.049, 0.050, 0.051, 0.052, 0.053]
        test_data[4] = Any["name3", true, true, false, true, false, true, true, true, false]
        test_data[5] = Any["name4", Time(10, 22), Time(10, 23), Time(10, 24), Time(10, 25), Time(10, 26), Time(10, 27), Time(10, 28), Time(10, 29), Time(10, 30)]
        test_data[6] = Any["name5", "Hello", Date(2025, 12, 19), 3, 3.33, "Hello", Date(2025, 12, 19), 3, 3.33, true]
        dtable = XLSX.readtransposedtable(joinpath(data_directory, "HTable.xlsx"), "Multiple", "2:7"; first_column=13, header=false, column_labels=["year", "Col2", "Col3", "Col4", "Col5", "Col6"])
        data, colnames = dtable.data, dtable.column_labels
        @test colnames == Symbol.(["year", "Col2", "Col3", "Col4", "Col5", "Col6"])

        dtable = XLSX.readtransposedtable(joinpath(data_directory, "HTable.xlsx"), "Multiple", "2:7"; first_column=13, header=false)
        data, colnames = dtable.data, dtable.column_labels
        @test colnames == Symbol.(["Col_1", "Col_2", "Col_3", "Col_4", "Col_5", "Col_6"])

        test_data = Vector{Any}(undef, 4)
        test_data[1] = ["A", "B", "C", "D"]
        test_data[2] = [10, 20, 30, 40]
        test_data[3] = [15, 25, 35, 40]
        test_data[4] = [20, 30, 40, 50]
        dtable = XLSX.readtransposedtable(joinpath(data_directory, "HTable.xlsx"), "Example")
        data, colnames = dtable.data, dtable.column_labels
        @test colnames == Symbol.(["Category", "Variable 1", "Variable 2", "Variable 3"])
        check_test_data(data, test_data)

        XLSX.addsheet!(f, "ExpandedDim")
        s = f["ExpandedDim"]
        s["B2"] = "Category"
        s["B3"] = "Variable 1"
        s["B4"] = "Variable 2"
        s["B5"] = "Variable 3"
        s["C2"] = test_data[1]
        s["C3"] = test_data[2]
        s["C4"] = test_data[3]
        s["C5"] = test_data[4]
        s["M10"] = "ExtraData"
        @test XLSX.get_dimension(s) == XLSX.CellRange("A1:M10")
        dtable = XLSX.gettransposedtable(s)
        data, colnames = dtable.data, dtable.column_labels
        @test colnames == Symbol.(["Category", "Variable 1", "Variable 2", "Variable 3"])
        check_test_data(data, test_data)

    end
end

@testset "Helper functions" begin

    test_data = Vector{Any}(undef, 3)
    test_data[1] = [missing, missing, "B5"]
    test_data[2] = ["C3", missing, missing]
    test_data[3] = [missing, "D4", missing]

    dtable = XLSX.readtable(joinpath(data_directory, "general.xlsx"), "table4")
    data, col_names = dtable.data, dtable.column_labels
    @test col_names == [:H1, :H2, :H3]

    dtable = XLSX.readtable(joinpath(data_directory, "general.xlsx"), "table4"; enable_cache=false)
    data, col_names = dtable.data, dtable.column_labels
    @test col_names == [:H1, :H2, :H3]

    @testset "Tables.jl DataTable interface" begin
        df = DataFrames.DataFrame(dtable)
        @test DataFrames.names(df) == ["H1", "H2", "H3"]
        @test size(df) == (3, 3)
        @test df[1, :H2] == "C3"
        @test df[3, :H1] == "B5"
        @test df[2, :H3] == "D4"
        @test ismissing(df[1, 1])
    end

    check_test_data(data, test_data)

    @test XLSX.readdata(joinpath(data_directory, "general.xlsx"), "table4", "E12") == "H1"
    test_data = Array{Any,2}(undef, 2, 1)
    test_data[1, 1] = "H2"
    test_data[2, 1] = "C3"

    @test XLSX.readdata(joinpath(data_directory, "general.xlsx"), "table4", "F12:F13") == test_data

    @testset "readtable select single column" begin
        dtable = XLSX.readtable(joinpath(data_directory, "general.xlsx"), "table4", "F")
        data, col_names = dtable.data, dtable.column_labels
        @test col_names == [:H2]
        @test data == Any[Any["C3"]]
    end

    @testset "readtable select column range" begin

        dtable = XLSX.readtable(joinpath(data_directory, "general.xlsx"), "table4", "F:G")
        data, col_names = dtable.data, dtable.column_labels
        @test col_names == [:H2, :H3]
        test_data = Any[Any["C3", missing], Any[missing, "D4"]]
        check_test_data(data, test_data)

        @test_throws XLSX.XLSXError XLSX.readtable(joinpath(data_directory, "general.xlsx"), "table4", "F:G"; header=false, column_labels=["a", "b", "c"])

    end

    @testset "readtable empty rows" begin

        t = XLSX.readtable(joinpath(data_directory, "EmptyTableRows.xlsx"), "EmptyRow", "B:C"; first_row=2, stop_in_empty_row=false, keep_empty_rows=true)
        test_data = Vector{Any}(undef, 2)
        test_data[1] = [1, 2, missing, missing, 3, 4, 5]
        test_data[2] = ["a", "b", missing, missing, "c", "d", "e"]
        check_test_data(t.data, test_data)
        @test t.column_labels == [Symbol("Col A"), Symbol("Col B")]
        @test t.column_label_index == Dict(Symbol("Col A") => 1, Symbol("Col B") => 2)

        t = XLSX.readtable(joinpath(data_directory, "EmptyTableRows.xlsx"), "EmptyCols", "B:C"; first_row=2, stop_in_empty_row=false, keep_empty_rows=true)
        test_data = Vector{Any}(undef, 2)
        test_data[1] = [1, 2, missing, missing, 3, 4, 5]
        test_data[2] = ["a", "b", missing, missing, "c", "d", "e"]
        check_test_data(t.data, test_data)
        @test t.column_labels == [Symbol("Col A"), Symbol("Col B")]
        @test t.column_label_index == Dict(Symbol("Col A") => 1, Symbol("Col B") => 2)

        t = XLSX.readtable(joinpath(data_directory, "EmptyTableRows.xlsx"), "MixedEmpty", "B:C"; first_row=2, stop_in_empty_row=false, keep_empty_rows=true)
        test_data = Vector{Any}(undef, 2)
        test_data[1] = [1, 2, missing, missing, 3, 4, 5, missing, missing, missing, missing, missing, 6, 7, 8, missing, missing, missing, missing, missing, missing, missing, 9, 10, 11, 12, 13, 14, 15, 16]
        test_data[2] = ["a", "b", missing, missing, "c", "d", "e", missing, missing, missing, missing, missing, "c", "d", "e", missing, missing, missing, missing, missing, missing, missing, "c", "d", "e", "c", "d", "e", "c", "d"]
        check_test_data(t.data, test_data)
        @test t.column_labels == [Symbol("Col A"), Symbol("Col B")]
        @test t.column_label_index == Dict(Symbol("Col A") => 1, Symbol("Col B") => 2)

        t = XLSX.readtable(joinpath(data_directory, "EmptyTableRows.xlsx"), "EmptyRow", "B:C"; first_row=2, stop_in_empty_row=false, keep_empty_rows=false)
        test_data = Vector{Any}(undef, 2)
        test_data[1] = [1, 2, 3, 4, 5]
        test_data[2] = ["a", "b", "c", "d", "e"]
        check_test_data(t.data, test_data)
        @test t.column_labels == [Symbol("Col A"), Symbol("Col B")]
        @test t.column_label_index == Dict(Symbol("Col A") => 1, Symbol("Col B") => 2)

        t = XLSX.readtable(joinpath(data_directory, "EmptyTableRows.xlsx"), "EmptyCols", "B:C"; first_row=2, stop_in_empty_row=false, keep_empty_rows=false)
        test_data = Vector{Any}(undef, 2)
        test_data[1] = [1, 2, 3, 4, 5]
        test_data[2] = ["a", "b", "c", "d", "e"]
        check_test_data(t.data, test_data)
        @test t.column_labels == [Symbol("Col A"), Symbol("Col B")]
        @test t.column_label_index == Dict(Symbol("Col A") => 1, Symbol("Col B") => 2)

        t = XLSX.readtable(joinpath(data_directory, "EmptyTableRows.xlsx"), "MixedEmpty", "B:C"; first_row=2, stop_in_empty_row=false, keep_empty_rows=false)
        test_data = Vector{Any}(undef, 2)
        test_data[1] = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16]
        test_data[2] = ["a", "b", "c", "d", "e", "c", "d", "e", "c", "d", "e", "c", "d", "e", "c", "d"]
        check_test_data(t.data, test_data)
        @test t.column_labels == [Symbol("Col A"), Symbol("Col B")]
        @test t.column_label_index == Dict(Symbol("Col A") => 1, Symbol("Col B") => 2)

        t = XLSX.readtable(joinpath(data_directory, "EmptyTableRows.xlsx"), "LeadingEmpty", "B:C"; first_row=2, stop_in_empty_row=false, keep_empty_rows=true)
        test_data = Vector{Any}(undef, 2)
        test_data[1] = [missing, missing, missing, 1, 2, missing, missing, 3, 4, 5, missing, missing, missing, missing, missing, 6, 7, 8, missing, missing, missing, missing, missing, missing, missing, 9, 10, 11, 12, 13, 14, 15, 16]
        test_data[2] = [missing, missing, missing, "a", "b", missing, missing, "c", "d", "e", missing, missing, missing, missing, missing, "c", "d", "e", missing, missing, missing, missing, missing, missing, missing, "c", "d", "e", "c", "d", "e", "c", "d"]
        check_test_data(t.data, test_data)
        @test t.column_labels == [Symbol("col A"), Symbol("col B")]
        @test t.column_label_index == Dict(Symbol("col A") => 1, Symbol("col B") => 2)

        t = XLSX.readtable(joinpath(data_directory, "EmptyTableRows.xlsx"), "LeadingEmpty", "B:C"; first_row=2, stop_in_empty_row=false, keep_empty_rows=false)
        test_data = Vector{Any}(undef, 2)
        test_data[1] = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16]
        test_data[2] = ["a", "b", "c", "d", "e", "c", "d", "e", "c", "d", "e", "c", "d", "e", "c", "d"]
        check_test_data(t.data, test_data)
        @test t.column_labels == [Symbol("col A"), Symbol("col B")]
        @test t.column_label_index == Dict(Symbol("col A") => 1, Symbol("col B") => 2)

        t = XLSX.readtable(joinpath(data_directory, "EmptyTableRows.xlsx"), "LeadingMixed", "B:C"; first_row=2, stop_in_empty_row=false, keep_empty_rows=true)
        test_data = Vector{Any}(undef, 2)
        test_data[1] = [missing, missing, missing, 1, 2, missing, missing, 3, 4, 5, missing, missing, missing, missing, missing, 6, 7, 8, missing, missing, missing, missing, missing, missing, missing, 9, 10, 11, 12, 13, 14, 15, 16]
        test_data[2] = [missing, missing, missing, "a", "b", missing, missing, "c", "d", "e", missing, missing, missing, missing, missing, "c", "d", "e", missing, missing, missing, missing, missing, missing, missing, "c", "d", "e", "c", "d", "e", "c", "d"]
        check_test_data(t.data, test_data)
        @test t.column_labels == [Symbol("col A"), Symbol("col B")]
        @test t.column_label_index == Dict(Symbol("col A") => 1, Symbol("col B") => 2)

        t = XLSX.readtable(joinpath(data_directory, "EmptyTableRows.xlsx"), "LeadingMixed", "B:C"; first_row=2, stop_in_empty_row=false, keep_empty_rows=false)
        test_data = Vector{Any}(undef, 2)
        test_data[1] = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16]
        test_data[2] = ["a", "b", "c", "d", "e", "c", "d", "e", "c", "d", "e", "c", "d", "e", "c", "d"]
        check_test_data(t.data, test_data)
        @test t.column_labels == [Symbol("col A"), Symbol("col B")]
        @test t.column_label_index == Dict(Symbol("col A") => 1, Symbol("col B") => 2)
    end

    @testset "stop function" begin

        t = XLSX.readtable(joinpath(data_directory, "EmptyTableRows.xlsx"), "LeadingEmpty", "B:C"; first_row=2, stop_in_empty_row=false, keep_empty_rows=true, stop_in_row_function=x -> !ismissing(x.cell_values[1]) && x.cell_values[1] == 2)
        test_data = Vector{Any}(undef, 2)
        test_data[1] = [missing, missing, missing, 1]
        test_data[2] = [missing, missing, missing, "a"]
        check_test_data(t.data, test_data)
        @test t.column_labels == [Symbol("col A"), Symbol("col B")]
        @test t.column_label_index == Dict(Symbol("col A") => 1, Symbol("col B") => 2)

        t = XLSX.readtable(joinpath(data_directory, "EmptyTableRows.xlsx"), "LeadingEmpty", "B:C"; first_row=2, stop_in_empty_row=false, keep_empty_rows=false, stop_in_row_function=x -> !ismissing(x.cell_values[1]) && x.cell_values[1] == 14)
        test_data = Vector{Any}(undef, 2)
        test_data[1] = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13]
        test_data[2] = ["a", "b", "c", "d", "e", "c", "d", "e", "c", "d", "e", "c", "d"]
        check_test_data(t.data, test_data)
        @test t.column_labels == [Symbol("col A"), Symbol("col B")]
        @test t.column_label_index == Dict(Symbol("col A") => 1, Symbol("col B") => 2)

        t = XLSX.readtable(joinpath(data_directory, "EmptyTableRows.xlsx"), "LeadingEmpty", "B:C"; first_row=2, stop_in_empty_row=false, keep_empty_rows=true, stop_in_row_function=x -> ismissing(x.cell_values[1]))
        @test isempty(t.data[1])
        @test isempty(t.data[2])
        @test t.column_labels == [Symbol("col A"), Symbol("col B")]
        @test t.column_label_index == Dict(Symbol("col A") => 1, Symbol("col B") => 2)

        t = XLSX.readtable(joinpath(data_directory, "EmptyTableRows.xlsx"), "LeadingEmpty", "B:C"; first_row=2, stop_in_empty_row=false, keep_empty_rows=false, stop_in_row_function=x -> !ismissing(x.cell_values[1]) && x.cell_values[1] == 1)
        @test isempty(t.data[1])
        @test isempty(t.data[2])
        @test t.column_labels == [Symbol("col A"), Symbol("col B")]
        @test t.column_label_index == Dict(Symbol("col A") => 1, Symbol("col B") => 2)

        t = XLSX.readtable(joinpath(data_directory, "EmptyTableRows.xlsx"), "LeadingMixed", "B:C"; first_row=2, stop_in_empty_row=false, keep_empty_rows=true, stop_in_row_function=x -> ismissing(x.cell_values[1]))
        @test isempty(t.data[1])
        @test isempty(t.data[2])
        @test t.column_labels == [Symbol("col A"), Symbol("col B")]
        @test t.column_label_index == Dict(Symbol("col A") => 1, Symbol("col B") => 2)

        t = XLSX.readtable(joinpath(data_directory, "EmptyTableRows.xlsx"), "LeadingMixed", "B:C"; first_row=2, stop_in_empty_row=true)
        @test isempty(t.data[1])
        @test isempty(t.data[2])
        @test t.column_labels == [Symbol("col A"), Symbol("col B")]
        @test t.column_label_index == Dict(Symbol("col A") => 1, Symbol("col B") => 2)

    end

    @testset "normalizenames" begin
        test_data = ["hello", "Hello 1", "123", Symbol("name")]
        @test XLSX.normalizename.(test_data) == [:hello, :Hello_1, :_123, :name]

        data = Vector{Any}()
        push!(data, [:sym1, :sym2, :sym3])
        push!(data, [1.0, 2.0, 3.0])
        push!(data, ["abc", "DeF", "gHi"])
        push!(data, [true, true, false])
        cols = ["1 col", "col \$2", "local", "col:4"]

        XLSX.writetable("mytest.xlsx", data, cols; overwrite=true)
        df = DataFrames.DataFrame(XLSX.readtable("mytest.xlsx", "Sheet1", normalizenames=true))
        @test DataFrames.names(df) == Any["_1_col", "col_2", "_local", "col_4"]

    end

    @testset "missing_strings" begin # issue #90
        t = XLSX.readtable(joinpath(data_directory, "missing_strings.xlsx"); missing_strings="N/A", stop_in_empty_row=false)
        test_data = Vector{Any}(undef, 2)
        test_data[1] = [1, 2, 3, missing, 4, 5]
        test_data[2] = [4, 5, 6, "Null", 7, 8]
        check_test_data(t.data, test_data)

        t = XLSX.readtable(joinpath(data_directory, "missing_strings.xlsx"); missing_strings="Null", stop_in_empty_row=false)
        test_data = Vector{Any}(undef, 2)
        test_data[1] = [1, 2, 3, "N/A", 4, 5]
        test_data[2] = [4, 5, 6, missing, 7, 8]
        check_test_data(t.data, test_data)

        t = XLSX.readtable(joinpath(data_directory, "missing_strings.xlsx"); missing_strings=["Null", "N/A"], stop_in_empty_row=false)
        test_data = Vector{Any}(undef, 2)
        test_data[1] = [1, 2, 3, missing, 4, 5]
        test_data[2] = [4, 5, 6, missing, 7, 8]
        check_test_data(t.data, test_data)

        t = XLSX.readtable(joinpath(data_directory, "missing_strings.xlsx"); missing_strings=["Null", "N/A"], stop_in_empty_row=false, keep_empty_rows=true)
        test_data = Vector{Any}(undef, 2)
        test_data[1] = [missing, 1, 2, 3, missing, missing, 4, 5]
        test_data[2] = [missing, 4, 5, 6, missing, missing, 7, 8]
        check_test_data(t.data, test_data)

    end

    @testset "Read DataFrame" begin

        df = XLSX.readto(joinpath(data_directory, "general.xlsx"), "table4", "F:G", DataFrames.DataFrame)
        @test names(df) == ["H2", "H3"]
        @test size(df) == (2, 2)
        @test df[1, :H2] == "C3"
        @test df[2, :H3] == "D4"
        @test ismissing(df[1, 2])
        @test ismissing(df[2, 1])

        df = XLSX.readto(joinpath(data_directory, "general.xlsx"), "table4", DataFrames.DataFrame)
        @test names(df) == ["H1", "H2", "H3"]
        @test size(df) == (3, 3)
        @test df[1, :H2] == "C3"
        @test df[2, :H3] == "D4"
        @test ismissing(df[1, :H1])
        @test ismissing(df[2, :H2])

        df = XLSX.readto(joinpath(data_directory, "general.xlsx"), DataFrames.DataFrame)
        @test names(df) == ["text", "regular text"]
        @test size(df) == (9, 2)
        @test df[1, "text"] == "integer"
        @test df[2, "regular text"] == 102.2
        @test df[3, 2] == Dates.Date(1983, 04, 16)
        @test df[5, 2] == Dates.DateTime(2018, 04, 16, 19, 19, 51)

        @test_throws XLSX.XLSXError XLSX.readto(joinpath(data_directory, "general.xlsx"))           # No sink
        @test_throws XLSX.XLSXError df = XLSX.readto(joinpath(data_directory, "general.xlsx"), 3)        # No sink
        @test_throws XLSX.XLSXError df = XLSX.readto(joinpath(data_directory, "general.xlsx"), 3, "F:G") # No sink

    end

end

@testset "Write" begin
    f = XLSX.open_xlsx_template(joinpath(data_directory, "general.xlsx"))
    filename_copy = "general_copy.xlsx"

    XLSX.writexlsx(filename_copy, f)
    @test isfile(filename_copy)

    f_copy = XLSX.readxlsx(filename_copy)

    s = f_copy["table"]
    s[:]
    dtable = XLSX.gettable(s)
    data, col_names = dtable.data, dtable.column_labels
    @test col_names == [Symbol("Column B"), Symbol("Column C"), Symbol("Column D"), Symbol("Column E"), Symbol("Column F"), Symbol("Column G")]

    test_data = Vector{Any}(undef, 6)
    test_data[1] = collect(1:8)
    test_data[2] = ["Str1", missing, "Str1", "Str1", "Str2", "Str2", "Str2", "Str2"]
    test_data[3] = [Date(2018, 4, 21) + Dates.Day(i) for i in 0:7]
    test_data[4] = [missing, missing, missing, missing, missing, "a", "b", missing]
    test_data[5] = [0.2001132319, 0.2793987377, 0.0950591677, 0.0744023067, 0.8242278091, 0.6205883578, 0.9174151018, 0.6749604883]
    test_data[6] = [missing for i in 1:8]
    check_test_data(data, test_data)
    isfile(filename_copy) && rm(filename_copy)
end

@testset "Save" begin
    f = XLSX.openxlsx("saveable.xlsx", mode="w")
    XLSX.renamesheet!(f["Sheet1"], "new_name")
    s = f[1]
    s[1:10, 1:10] = "hello world"
    @test XLSX.savexlsx(f) == abspath("saveable.xlsx")
    f2 = XLSX.openxlsx("saveable.xlsx", mode="rw")
    @test XLSX.hassheet(f2, "new_name")
    @test f2["new_name"][1, 1] == "hello world"
    @test f2["new_name"][10, 10] == "hello world"
    f2["new_name"][1:5, 1:5] = "goodbye world"
    XLSX.savexlsx(f2)
    f3 = XLSX.openxlsx("saveable.xlsx", mode="r")
    #f3["new_name"][:]
    @test f3["new_name"][1, 1] == "goodbye world"
    @test f3["new_name"][5, 5] == "goodbye world"
    @test f3["new_name"][10, 10] == "hello world"
    isfile("saveable.xlsx") && rm("saveable.xlsx")
end

@testset "CustomXml" begin
    # issue #210
    # None of the example .xlsx files in the test suite include custoimXml internal files
    # but customXml.xlsx does
    template = XLSX.open_xlsx_template(joinpath(data_directory, "customXml.xlsx"))
    filename_copy = "customXml_copy.xlsx"
    for sn in XLSX.sheetnames(template)
        sheet = template[sn]
        sheet["Q1"] = "Cant" # using an apostrophe here causes a test failure reading the copy - "Evaluated: "Can&apos;t" == "Can't""
        sheet["Q2"] = "write"
        sheet["Q3"] = "this"
        sheet["Q4"] = "template"
    end
    @test XLSX.writexlsx(filename_copy, template, overwrite=true) == abspath(filename_copy) # This is where the bug will throw if customXml internal files present.
    @test isfile(filename_copy)
    f_copy = XLSX.readxlsx(filename_copy) # Don't really think this second part is necessary.
    test_Xmlread = [["Cant", "write", "this", "template"]]
    for sn in XLSX.sheetnames(f_copy)
        sheet = template[sn]
        data = [[sheet["Q1"], sheet["Q2"], sheet["Q3"], sheet["Q4"]]]
        check_test_data(data, test_Xmlread)
    end
    isfile(filename_copy) && rm(filename_copy)
end

@testset "Edit Template" begin
    new_filename = "new_file_from_empty_template.xlsx"
    isfile(new_filename) && rm(new_filename)
    f = XLSX.open_empty_template()
    f["Sheet1"]["A1"] = "Hello"
    f["Sheet1"]["A2"] = 10
    XLSX.writexlsx(new_filename, f, overwrite=true)

    f = XLSX.readxlsx(new_filename)
    @test f["Sheet1"]["A1"] == "Hello"
    @test f["Sheet1"]["A2"] == 10

    rm(new_filename)
end

@testset "add/copy sheet!" begin

    @testset "addsheet!" begin

        new_filename = "template_with_new_sheet.xlsx"
        f = XLSX.open_empty_template()
        s = XLSX.addsheet!(f, "new_sheet")
        s["A1"] = 10
        @test XLSX.sheetnames(f) == ["Sheet1", "new_sheet"]
        XLSX.writexlsx(new_filename, f, overwrite=true)


        big_sheetname = "aaaaaaaaaabbbbbbbbbbccccccccccd"
        s2 = XLSX.addsheet!(f, big_sheetname)

        XLSX.writexlsx(new_filename, f, overwrite=true)
        fx = XLSX.opentemplate(new_filename)
        @test XLSX.sheetnames(f) == ["Sheet1", "new_sheet", big_sheetname]

    end

    @testset "invalid sheet names" begin

        f = XLSX.open_empty_template()
        s = XLSX.addsheet!(f, "new_sheet")
        s["A1"] = 10
        invalid_names = [
            "aaaaaaaaaabbbbbbbbbbccccccccccd1",
            "abc:def",
            "abcdef/",
            "\\aaaa",
            "hey?you",
            "[mysheet]",
            "asteri*"
        ]

        for invalid_name in invalid_names
            @test_throws XLSX.XLSXError XLSX.addsheet!(f, invalid_name)
        end

    end

    @testset "copysheet!" begin

        f = XLSX.newxlsx()
        XLSX.renamesheet!(f["Sheet1"], "new_name")
        XLSX.addsheet!(f)
        for x = 1:10, y = 1:10
            f["Sheet1"][x, y] = x + y
            f["new_name"][x, y] = x * y
        end
        XLSX.addDefinedName(f["new_name"], "new_name_range", "A1:B10")
        XLSX.addDefinedName(f["Sheet1"], "Sheet1_range", "C1:D10")
        XLSX.setBorder(f["new_name"], "A1:D10"; allsides=["style" => "thin", "color" => "red"])
        XLSX.setBorder(f["Sheet1"], "A1:D10"; allsides=["style" => "thin", "color" => "red"])
        XLSX.setConditionalFormat(f["new_name"], "A1:D10", :colorScale)

        s3 = XLSX.copysheet!(f["new_name"], "copied_sheet")
        @test s3.name == "copied_sheet"
        @test s3["A1"] == 1
        @test s3[5, 5] == 25
        @test s3[10, 10] == 100
        @test XLSX.get_workbook(s3).worksheet_names == XLSX.get_workbook(f["new_name"]).worksheet_names
        @test XLSX.getConditionalFormats(s3) == XLSX.getConditionalFormats(f["new_name"])
        @test XLSX.getBorder(s3, "C5").border == XLSX.getBorder(f["new_name"], "C5").border

        # Check that the original sheet is unchanged
        s2 = f["new_name"]
        @test s2["A1"] == 1
        @test s2[5, 5] == 25
        @test s2[10, 10] == 100

        s4 = XLSX.copysheet!(s3)
        @test s4.name == "copied_sheet (copy)"
        @test s4["A1"] == 1
        @test s4[5, 5] == 25
        @test s4[10, 10] == 100

        @test XLSX.get_workbook(s4).worksheet_names == XLSX.get_workbook(f["new_name"]).worksheet_names
        XLSX.setBorder(s4, "F1:H10"; allsides=["style" => "thin", "color" => "green"])
        XLSX.setConditionalFormat(s4, "F1:H10", :colorScale; colorscale="redyellowgreen")

        XLSX.writexlsx("copied_sheets.xlsx", f, overwrite=true)
        f = XLSX.opentemplate("copied_sheets.xlsx")
        @test XLSX.sheetnames(f) == ["new_name", "Sheet1", "copied_sheet", "copied_sheet (copy)"]
        @test XLSX.get_workbook(f["copied_sheet"]).worksheet_names == XLSX.get_workbook(f["new_name"]).worksheet_names
        @test XLSX.getConditionalFormats(f["copied_sheet (copy)"]) == XLSX.getConditionalFormats(s4)
        @test XLSX.getBorder(f["copied_sheet (copy)"], "C5").border == XLSX.getBorder(f["new_name"], "C5").border
        @test XLSX.getBorder(f["copied_sheet (copy)"], "G5").border == XLSX.getBorder(s4, "G5").border

    end
    isfile("copied_sheets.xlsx") && rm("copied_sheets.xlsx")

    @testset "deletesheet!" begin

        new_filename = "template_with_new_sheet.xlsx"
        big_sheetname = "aaaaaaaaaabbbbbbbbbbccccccccccd"
        fx = XLSX.opentemplate(new_filename)
        XLSX.deletesheet!(fx, big_sheetname)
        @test XLSX.sheetnames(fx) == ["Sheet1", "new_sheet"]
        XLSX.writexlsx(new_filename, fx, overwrite=true)
        f = XLSX.readxlsx(new_filename)
        @test XLSX.sheetnames(f) == ["Sheet1", "new_sheet"]

        f = XLSX.opentemplate(joinpath(data_directory, "general.xlsx"))
        sc = XLSX.sheetcount(f)
        XLSX.deletesheet!(f, "empty")
        @test XLSX.sheetcount(f) == sc - 1 # Check it's gone.
        @test XLSX.hassheet(f, "empty") == false # Check it's gone.
        @test_throws XLSX.XLSXError XLSX.deletesheet!(f, "empty") # Already deleted.
        @test_throws XLSX.XLSXError XLSX.deletesheet!(f, "nosuchsheet") # Never there.
        s2 = XLSX.addsheet!(f, "this_now")
        @test XLSX.sheetnames(f) == ["general", "table3", "table4", "table", "table2", "table5", "table6", "table7", "lookup", "header_error", "named_ranges_2", "named_ranges", "this_now"]
        XLSX.writexlsx(new_filename, f, overwrite=true)

        f = XLSX.opentemplate(new_filename)
        @test XLSX.sheetnames(f) == ["general", "table3", "table4", "table", "table2", "table5", "table6", "table7", "lookup", "header_error", "named_ranges_2", "named_ranges", "this_now"]
        XLSX.deletesheet!(f, "named_ranges")
        XLSX.deletesheet!(f["general"])
        @test XLSX.sheetnames(f) == ["table3", "table4", "table", "table2", "table5", "table6", "table7", "lookup", "header_error", "named_ranges_2", "this_now"]
        XLSX.writexlsx(new_filename, f, overwrite=true)
        dtable = XLSX.readtable(new_filename, "table4", "F:G")
        data, col_names = dtable.data, dtable.column_labels
        @test col_names == [:H2, :H3]
        test_data = Any[Any["C3", missing], Any[missing, "D4"]]
        check_test_data(data, test_data)
        @test XLSX.deletesheet!(f, 1) === f
        @test XLSX.sheetnames(f) == ["table4", "table", "table2", "table5", "table6", "table7", "lookup", "header_error", "named_ranges_2", "this_now"]
        XLSX.writexlsx(new_filename, f, overwrite=true)
        dtable = XLSX.readtable(new_filename, "table4", "F:G")
        data, col_names = dtable.data, dtable.column_labels
        @test col_names == [:H2, :H3]
        test_data = Any[Any["C3", missing], Any[missing, "D4"]]
        check_test_data(data, test_data)

        f = XLSX.opentemplate(joinpath(data_directory, "Book_1904.xlsx")) # Only one sheet - can't delete
        @test_throws XLSX.XLSXError XLSX.deletesheet!(f, 1)
        s = f[1]
        @test_throws XLSX.XLSXError XLSX.deletesheet!(s)
        @test_throws XLSX.XLSXError XLSX.deletesheet!(f, "Sheet1")

        f = XLSX.openxlsx(joinpath(data_directory, "deletesheet.xlsx"), mode="rw")
        XLSX.deletesheet!(f[1])
        @test XLSX.getcell(f[1], "A1") == XLSX.Cell(XLSX.get_workbook(f), XLSX.CellRef("A1"), "e", "", "#REF!", "", true)
        @test XLSX.getFormula(f[1], "A1") == "=#REF!+#REF!"

        isfile("template_with_new_sheet.xlsx") && rm("template_with_new_sheet.xlsx")

        # XLSX.jl cannot currently create formula references that refer to other sheets, so we have to manually 
        # insert them into the workbook's formula cache to test. Excel cannot read such a workbook and deletes these 
        # formulas on opening. Tests are for completeness only
        f=XLSX.openxlsx("renamedelete.xlsx", mode="w") 
        s=f[1]
        s[1:10, 1] = collect(1:10)
        XLSX.addsheet!(f)
        s2=f[2]
        s2["A1:B4"]=""
        XLSX.setFormula(s2, "A1", "=Sheet1!A1+10")
        XLSX.getcell(s2, "B1").formula=true
        XLSX.getcell(s2, "B2").formula=true
        XLSX.getcell(s2, "B3").formula=true
        XLSX.getcell(s2, "B4").formula=true
        w=XLSX.get_workbook(s2)
        w.formulas[XLSX.SheetCellRef("Sheet2!B1")] = XLSX.ReferencedFormula("=Sheet1!A1+10", 0, "B1:B4", nothing)
        w.formulas[XLSX.SheetCellRef("Sheet2!B2")] = XLSX.FormulaReference(0, nothing)
        w.formulas[XLSX.SheetCellRef("Sheet2!B3")] = XLSX.FormulaReference(0, nothing)
        w.formulas[XLSX.SheetCellRef("Sheet2!B4")] = XLSX.FormulaReference(0, nothing)

        XLSX.deletesheet!(s)
        @test XLSX.hassheet(f, "Sheet1") == false
        @test XLSX.getcell(s2, "A1") == XLSX.Cell(XLSX.get_workbook(f), XLSX.CellRef("A1"), "e", "", "#REF!", "", true)
        @test XLSX.getcell(s2, "B1") == XLSX.Cell(XLSX.get_workbook(f), XLSX.CellRef("B1"), "e", "", "#REF!", "", true)
        @test XLSX.getcell(s2, "B2") == XLSX.Cell(XLSX.get_workbook(f), XLSX.CellRef("B2"), "e", "", "#REF!", "", true)
        @test XLSX.getcell(s2, "B4") == XLSX.Cell(XLSX.get_workbook(f), XLSX.CellRef("B4"), "e", "", "#REF!", "", true)
        @test XLSX.getFormula(s2, "A1") == "=#REF!+10"
        @test XLSX.getFormula(s2, "B1") == "=#REF!+10"
        @test XLSX.getFormula(s2, "B2") == "=#REF!+10"
        @test XLSX.getFormula(s2, "B4") == "=#REF!+10"
        @test XLSX.get_formula_from_cache(s2, XLSX.CellRef("A1")) == XLSX.Formula("=#REF!+10", nothing, nothing, nothing)
        @test XLSX.get_formula_from_cache(s2, XLSX.CellRef("B1")) == XLSX.ReferencedFormula("=#REF!+10", 0, "B1:B4", nothing)
        @test XLSX.get_formula_from_cache(s2, XLSX.CellRef("B2")) ==  XLSX.FormulaReference(0, nothing)
        @test XLSX.get_formula_from_cache(s2, XLSX.CellRef("B4")) ==  XLSX.FormulaReference(0, nothing)

    end
    @testset "renamesheet!" begin

        # XLSX.jl cannot currently create formula references that refer to other sheets, so we have to manually 
        # insert them into the workbook's formula cache to test. Excel cannot read such a workbook and deletes these 
        # formulas on opening. Tests are for completeness only
        f=XLSX.openxlsx("renamedelete.xlsx", mode="w")
        s=f[1]
        s[1:10, 1] = collect(1:10)
        XLSX.addsheet!(f)
        s2=f[2]
        s2["A1:B4"]=""
        XLSX.setFormula(s2, "A1", "=Sheet1!A1+10")
        XLSX.getcell(s2, "B1").formula=true
        XLSX.getcell(s2, "B2").formula=true
        XLSX.getcell(s2, "B3").formula=true
        XLSX.getcell(s2, "B4").formula=true
        w=XLSX.get_workbook(s2)
        w.formulas[XLSX.SheetCellRef("Sheet2!B1")] = XLSX.ReferencedFormula("=Sheet1!A1+10", 0, "B1:B4", nothing)
        w.formulas[XLSX.SheetCellRef("Sheet2!B2")] = XLSX.FormulaReference(0, nothing)
        w.formulas[XLSX.SheetCellRef("Sheet2!B3")] = XLSX.FormulaReference(0, nothing)
        w.formulas[XLSX.SheetCellRef("Sheet2!B4")] = XLSX.FormulaReference(0, nothing)

        XLSX.renamesheet!(s2, "Sheet4")
        @test s2.name == "Sheet4"
        @test XLSX.getcell(s2, "A1") == XLSX.Cell(XLSX.get_workbook(f), XLSX.CellRef("A1"), "", "", "", "", true)
        @test XLSX.getcell(s2, "B1") == XLSX.Cell(XLSX.get_workbook(f), XLSX.CellRef("B1"), "", "", "", "", true)
        @test XLSX.getcell(s2, "B2") == XLSX.Cell(XLSX.get_workbook(f), XLSX.CellRef("B2"), "", "", "", "", true)
        @test XLSX.getcell(s2, "B4") == XLSX.Cell(XLSX.get_workbook(f), XLSX.CellRef("B4"), "", "", "", "", true)
        @test XLSX.getFormula(s2, "A1") == "=Sheet1!A1+10"
        @test XLSX.getFormula(s2, "B1") == "=Sheet1!A1+10"
        @test XLSX.getFormula(s2, "B2") == "=Sheet1!A2+10"
        @test XLSX.getFormula(s2, "B4") == "=Sheet1!A4+10"
        @test XLSX.get_formula_from_cache(s2, XLSX.CellRef("A1")) == XLSX.Formula("=Sheet1!A1+10", nothing, nothing, nothing)
        @test XLSX.get_formula_from_cache(s2, XLSX.CellRef("B1")) == XLSX.ReferencedFormula("=Sheet1!A1+10", 0, "B1:B4", nothing)
        @test XLSX.get_formula_from_cache(s2, XLSX.CellRef("B2")) ==  XLSX.FormulaReference(0, nothing)
        @test XLSX.get_formula_from_cache(s2, XLSX.CellRef("B4")) ==  XLSX.FormulaReference(0, nothing)

        XLSX.renamesheet!(s, "Sheet3")
        @test s.name == "Sheet3"
        @test XLSX.getcell(s2, "A1") == XLSX.Cell(XLSX.get_workbook(f), XLSX.CellRef("A1"), "", "", "", "", true)
        @test XLSX.getcell(s2, "B1") == XLSX.Cell(XLSX.get_workbook(f), XLSX.CellRef("B1"), "", "", "", "", true)
        @test XLSX.getcell(s2, "B2") == XLSX.Cell(XLSX.get_workbook(f), XLSX.CellRef("B2"), "", "", "", "", true)
        @test XLSX.getcell(s2, "B4") == XLSX.Cell(XLSX.get_workbook(f), XLSX.CellRef("B4"), "", "", "", "", true)
        @test XLSX.getFormula(s2, "A1") == "=Sheet3!A1+10"
        @test XLSX.getFormula(s2, "B2") == "=Sheet3!A2+10"
        @test XLSX.getFormula(s2, "B3") == "=Sheet3!A3+10"
        @test XLSX.getFormula(s2, "B4") == "=Sheet3!A4+10"
        @test XLSX.get_formula_from_cache(s2, XLSX.CellRef("A1")) == XLSX.Formula("=Sheet3!A1+10", nothing, nothing, nothing)
        @test XLSX.get_formula_from_cache(s2, XLSX.CellRef("B1")) == XLSX.ReferencedFormula("=Sheet3!A1+10", 0, "B1:B4", nothing)
        @test XLSX.get_formula_from_cache(s2, XLSX.CellRef("B2")) ==  XLSX.FormulaReference(0, nothing)
        @test XLSX.get_formula_from_cache(s2, XLSX.CellRef("B4")) ==  XLSX.FormulaReference(0, nothing)

    end
end

@testset "Edit" begin
    f = XLSX.open_xlsx_template(joinpath(data_directory, "general.xlsx"))
    s = f["general"]
    @test_throws XLSX.XLSXError s["A1"] = :sym
    XLSX.renamesheet!(s, "general") # no-op
    @test_throws XLSX.XLSXError XLSX.renamesheet!(s, "table") # name is taken
    XLSX.renamesheet!(s, "renamed_sheet") # retain old function to avoid breaking change
    @test s.name == "renamed_sheet"
    s["A1"] = "Hey You!"
    s["B1"] = "Out there in the cold..."
    s["A2"] = "Getting lonely getting old..."
    s["B2"] = "Can you feel me?"
    s["A3"] = 1000
    s["B3"] = 99.99

    # create a new sheet
    s = XLSX.addsheet!(f, "my_new_sheet_1")
    s = XLSX.addsheet!(f, "my_new_sheet_2")
    s["B1"] = "This is a new sheet"
    s["B2"] = "This is a new sheet"
    s = XLSX.addsheet!(f)
    s["B1"] = "unnamed sheet"

    XLSX.writexlsx("general_copy_2.xlsx", f, overwrite=true)
    @test isfile("general_copy_2.xlsx")

    XLSX.openxlsx("general_copy_2.xlsx") do f
        s = f["renamed_sheet"]
        @test s["A1"] == "Hey You!"
        @test s["B1"] == "Out there in the cold..."
        @test s["A2"] == "Getting lonely getting old..."
        @test s["B2"] == "Can you feel me?"
        @test s["A3"] == 1000
        @test s["B3"] == 99.99
        f["my_new_sheet_1"]
        @test f["my_new_sheet_2"]["B1"] == "This is a new sheet"
        @test f["my_new_sheet_2"]["B2"] == "This is a new sheet"
        @test f["Sheet1"]["B1"] == "unnamed sheet"
    end

    isfile("general_copy_2.xlsx") && rm("general_copy_2.xlsx")
end

@testset "writetable" begin

    @testset "single" begin
        col_names = ["Integers", "Strings", "Floats", "Booleans", "Dates", "Times", "DateTimes", "AbstractStrings", "Rational", "Irrationals", "MixedStringNothingMissing"]
        data = Vector{Any}(undef, 11)
        data[1] = [1, 2, missing, UInt8(4)]
        data[2] = ["Hey", "You", "Out", "There"]
        data[3] = [101.5, 102.5, missing, 104.5]
        data[4] = [true, false, missing, true]
        data[5] = [Date(2018, 2, 1), Date(2018, 3, 1), Date(2018, 5, 20), Date(2018, 6, 2)]
        data[6] = [Dates.Time(19, 10), Dates.Time(19, 20), Dates.Time(19, 30), Dates.Time(0, 0)]
        data[7] = [Dates.DateTime(2018, 5, 20, 19, 10), Dates.DateTime(2018, 5, 20, 19, 20), Dates.DateTime(2018, 5, 20, 19, 30), Dates.DateTime(2018, 5, 20, 19, 40)]
        data[8] = SubString.(["Hey", "You", "Out", "There"], 1, 2)
        data[9] = [1 // 2, 1 // 3, missing, 22 // 3]
        data[10] = [pi, sqrt(2), missing, sqrt(5)]
        data[11] = [nothing, "middle", missing, nothing]

        XLSX.writetable("output_table.xlsx", data, col_names, overwrite=true, sheetname="report", anchor_cell="B2")
        @test isfile("output_table.xlsx")

        dtable = XLSX.readtable("output_table.xlsx", "report")
        read_data, read_column_names = dtable.data, dtable.column_labels
        @test length(read_column_names) == length(col_names)
        for c in axes(col_names, 1)
            @test Symbol(col_names[c]) == read_column_names[c]
        end
        check_test_data(read_data, data)
    end

    @testset "multiple" begin
        report_1_column_names = ["HEADER_A", "HEADER_B"]
        report_1_data = Vector{Any}(undef, 2)
        report_1_data[1] = [1, 2, 3]
        report_1_data[2] = ["A", "B", ""]

        report_2_column_names = ["COLUMN_A", "COLUMN_B"]
        report_2_data = Vector{Any}(undef, 2)
        report_2_data[1] = [Date(2017, 2, 1), Date(2018, 2, 1)]
        report_2_data[2] = [10.2, 10.3]

        XLSX.writetable("output_tables.xlsx", overwrite=true, REPORT_A=(report_1_data, report_1_column_names), REPORT_B=(report_2_data, report_2_column_names))
        @test isfile("output_tables.xlsx")

        dtable = XLSX.readtable("output_tables.xlsx", "REPORT_A")
        data, labels = dtable.data, dtable.column_labels
        @test labels[1] == :HEADER_A
        @test labels[2] == :HEADER_B
        check_test_data(data, report_1_data)

        dtable = XLSX.readtable("output_tables.xlsx", "REPORT_B")
        data, labels = dtable.data, dtable.column_labels
        @test labels[1] == :COLUMN_A
        @test labels[2] == :COLUMN_B
        check_test_data(data, report_2_data)
        XLSX.writetable("output_tables.xlsx", [("REPORT_A", report_1_data, report_1_column_names), ("REPORT_B", report_2_data, report_2_column_names)], overwrite=true)

        dtable = XLSX.readtable("output_tables.xlsx", "REPORT_A")
        data, labels = dtable.data, dtable.column_labels
        @test labels[1] == :HEADER_A
        @test labels[2] == :HEADER_B
        check_test_data(data, report_1_data)

        dtable = XLSX.readtable("output_tables.xlsx", "REPORT_B")
        data, labels = dtable.data, dtable.column_labels
        @test labels[1] == :COLUMN_A
        @test labels[2] == :COLUMN_B
        check_test_data(data, report_2_data)

        report_1_column_names = [:HEADER_A, :HEADER_B]
        report_2_column_names = [:COLUMN_A, :COLUMN_B]
        XLSX.writetable("output_tables.xlsx", [("REPORT_A", report_1_data, report_1_column_names), ("REPORT_B", report_2_data, report_2_column_names)], overwrite=true)

        dtable = XLSX.readtable("output_tables.xlsx", "REPORT_A")
        data, labels = dtable.data, dtable.column_labels
        @test labels[1] == :HEADER_A
        @test labels[2] == :HEADER_B
        check_test_data(data, report_1_data)

        dtable = XLSX.readtable("output_tables.xlsx", "REPORT_B")
        data, labels = dtable.data, dtable.column_labels
        @test labels[1] == :COLUMN_A
        @test labels[2] == :COLUMN_B
        check_test_data(data, report_2_data)

        report_1_column_names = ["HEADER_A", "HEADER_B"]
        report_1_data = [["1", "2", "3"], ["A", "B", ""]]

        report_2_column_names = ["COLUMN_A", "COLUMN_B"]
        report_2_data = Vector{Any}(undef, 2)
        report_2_data[1] = [Date(2017, 2, 1), Date(2018, 2, 1)]
        report_2_data[2] = [10.2, 10.3]
        XLSX.writetable("output_tables.xlsx", [("REPORT_A", report_1_data, report_1_column_names), ("REPORT_B", report_2_data, report_2_column_names)], overwrite=true)

        dtable = XLSX.readtable("output_tables.xlsx", "REPORT_A")
        data, labels = dtable.data, dtable.column_labels
        @test labels[1] == :HEADER_A
        @test labels[2] == :HEADER_B
        check_test_data(data, report_1_data)

        dtable = XLSX.readtable("output_tables.xlsx", "REPORT_B")
        data, labels = dtable.data, dtable.column_labels
        @test labels[1] == :COLUMN_A
        @test labels[2] == :COLUMN_B
        check_test_data(data, report_2_data)
    end

    @testset "writetable to IO" begin
        dt = XLSX.DataTable(Any[Any[1, 2, 3], Any[4, 5, 6]], [:a, :b])
        io = IOBuffer()
        XLSX.writetable(io, "Test" => dt)
        seek(io, 0)
        dt_read = XLSX.readtable(io, "Test")
        @test dt_read.data == dt.data
        @test dt_read.column_labels == dt.column_labels
        @test dt_read.column_label_index == dt.column_label_index
    end

    @testset "extended types" begin # Issue #239
        @enum enums begin
            enum1
            enum2
            enum3
        end

        data = Vector{Any}()
        push!(data, [:sym1, :sym2, :sym3])
        push!(data, [1.0, 2.0, 3.0])
        push!(data, ["abc", "DeF", "gHi"])
        push!(data, [true, true, false])
        push!(data, [XLSX.CellRef("A1"), XLSX.CellRef("B2"), XLSX.CellRef("CCC34000")])
        push!(data, collect(instances(enums)))
        cols = [string(eltype(x)) for x in data]

        XLSX.writetable("mytest.xlsx", data, cols; overwrite=true)

        f = XLSX.readxlsx("mytest.xlsx")
        @test f[1]["A1"] == "Symbol"
        @test f[1]["A1:A4"] == Any["Symbol"; "sym1"; "sym2"; "sym3";;] # A 2D Array, size (4, 1)
        @test f[1]["A1"] == "Symbol"
        @test f[1]["E1:E4"] == Any["XLSX.CellRef"; "A1"; "B2"; "CCC34000";;]
    end

    @testset "NaN and Inf" begin # Issues #342 and #179
        col_names = ["Integers", "Strings", "Floats", "Booleans", "Dates", "Times", "DateTimes"]
        data = Vector{Any}(undef, 7)
        data[1] = [1, 2, missing, 4]
        data[2] = ["Hey", "You", "Out", "There"]
        data[3] = [-Inf, Inf, missing, NaN]
        data[4] = [true, false, missing, true]
        data[5] = [Date(2018, 2, 1), Date(2018, 3, 1), Date(2018, 5, 20), Date(2018, 6, 2)]
        data[6] = [Dates.Time(19, 10), Dates.Time(19, 20), Dates.Time(19, 30), Dates.Time(19, 40)]
        data[7] = [Dates.DateTime(2018, 5, 20, 19, 10), Dates.DateTime(2018, 5, 20, 19, 20), Dates.DateTime(2018, 5, 20, 19, 30), Dates.DateTime(2018, 5, 20, 19, 40)]
        table = NamedTuple{Tuple(Symbol(x) for x in col_names)}(Tuple(data))

        f=XLSX.newxlsx()
        s=f[1]
        XLSX.writetable!(s, table)
        @test s["C2"] == "-Inf"
        @test s["C3"] == "Inf"
        @test ismissing(s["C4"])
        @test s["C5"] == "NaN"

        XLSX.writetable("mytest.xlsx", data, col_names; overwrite=true)

        f2=XLSX.readxlsx("mytest.xlsx")
        s2=f2[1]
        @test s2["C2"] == "-Inf"
        @test s2["C3"] == "Inf"
        @test ismissing(s2["C4"])
        @test s2["C5"] == "NaN"

        df=XLSX.readto("mytest.xlsx", DataFrames.DataFrame)
        @test all(isequal.(df.Floats, ["-Inf", "Inf", missing, "NaN"]))

    end

    # delete files created by this testset
    delete_files = ["output_table.xlsx", "output_tables.xlsx", "mytest.xlsx"]
    for f in delete_files
        isfile(f) && rm(f)
    end
end

@testset "Styles" begin

    @testset "Original" begin
        using XLSX: CellValue, id, getcell, setdata!, CellRef
        xfile = XLSX.open_empty_template()
        wb = XLSX.get_workbook(xfile)
        sheet = xfile["Sheet1"]

        datefmt = XLSX.styles_add_numFmt(wb, "yyyymmdd")
        numfmt = XLSX.styles_add_numFmt(wb, "\$* #,##0.00;\$* (#,##0.00);\$* \"-\"??;[Magenta]@")

        #Check format id numbers dont intersect with predefined formats or each other
        @test datefmt == 164
        @test numfmt == 165

        font = XLSX.styles_add_cell_font(wb, Dict("b" => nothing, "sz" => Dict("val" => "24")))
        xroot = XLSX.styles_xmlroot(wb)
        fontnodes = XLSX.find_all_nodes("/$(XLSX.SPREADSHEET_NAMESPACE_XPATH_ARG):styleSheet/$(XLSX.SPREADSHEET_NAMESPACE_XPATH_ARG):fonts/$(XLSX.SPREADSHEET_NAMESPACE_XPATH_ARG):font", xroot)
        fontnode = fontnodes[font+1] # XML is zero indexed so we need to add 1 to get the right node

        # Check the font was written correctly
        @test XML.tag(fontnode) == "font"
        @test length(XML.children(fontnode)) == 2
        @test XML.tag(XML.children(fontnode)[1]) == "b"
        @test XML.tag(XML.children(fontnode)[2]) == "sz"
        @test XML.children(fontnode)[2]["val"] == "24"

        textstyle = XLSX.styles_add_cell_xf(wb, Dict("applyFont" => "true", "fontId" => "$font"))
        datestyle = XLSX.styles_add_cell_xf(wb, Dict("applyNumberFormat" => "1", "numFmtId" => "$datefmt"))
        numstyle = XLSX.styles_add_cell_xf(wb, Dict("applyFont" => "1", "applyNumberFormat" => "1", "fontId" => "$font", "numFmtId" => "$numfmt"))

        xf = XLSX.styles_get_cellXf_with_numFmtId(wb, 1000)
        @test xf == XLSX.EmptyCellDataFormat()
        @test isempty(xf)
        @test id(xf) == UInt64(0)

        @test textstyle isa XLSX.CellDataFormat
        @test !isempty(textstyle)
        @test id(textstyle) == UInt64(1)

        @test XLSX.styles_get_cellXf_with_numFmtId(wb, datefmt) == datestyle
        @test XLSX.styles_numFmt_formatCode(wb, string(datefmt)) == "yyyymmdd"
        @test datestyle isa XLSX.CellDataFormat
        @test !isempty(datestyle)
        @test id(datestyle) == UInt64(2)

        @test XLSX.styles_get_cellXf_with_numFmtId(wb, numfmt) == numstyle
        @test XLSX.styles_numFmt_formatCode(wb, string(numfmt)) == "\$* #,##0.00;\$* (#,##0.00);\$* &quot;-&quot;??;[Magenta]@"
        @test numstyle isa XLSX.CellDataFormat
        @test !isempty(numstyle)
        @test id(numstyle) == UInt64(3)

        setdata!(sheet, CellRef("A1"), CellValue(Date(2011, 10, 13), datestyle))
        setdata!(sheet, CellRef("A2"), CellValue(1000, numstyle))
        setdata!(sheet, CellRef("A3"), CellValue(1000.10, numstyle))
        setdata!(sheet, CellRef("A4"), CellValue(-1000.10, numstyle))
        setdata!(sheet, CellRef("A5"), CellValue(0, numstyle))
        setdata!(sheet, CellRef("A6"), CellValue("hello", numstyle))
        setdata!(sheet, CellRef("B1"), CellValue("hello world", textstyle))

        @test sheet["A1"] == Date(2011, 10, 13)
        cell = getcell(sheet, "A1")
        @test cell.style == id(datestyle)
        formatid = XLSX.styles_cell_xf_numFmtId(wb, Int64(cell.style))
        @test formatid == datefmt

        cellstyle = getcell(sheet, "A2").style
        @test cellstyle == id(numstyle)
        formatid = XLSX.styles_cell_xf_numFmtId(wb, Int64(cellstyle))
        @test formatid == numfmt

        @test sheet["A2"] == 1000
        @test sheet["A3"] == 1000.10
        @test XLSX.getcell(sheet, "A3").style == cellstyle
        @test sheet["A4"] == -1000.10
        @test XLSX.getcell(sheet, "A4").style == cellstyle
        @test sheet["A5"] == 0
        @test XLSX.getcell(sheet, "A5").style == cellstyle
        @test sheet["A6"] == "hello"
        @test XLSX.getcell(sheet, "A6").style == cellstyle

        @test sheet["B1"] == "hello world"
        @test XLSX.getcell(sheet, "B1").style == id(textstyle)

        sheet["B2"] = CellValue("hello world", textstyle)
        @test sheet["B2"] == "hello world"
        @test XLSX.getcell(sheet, "B2").style == id(textstyle)

        sheet[3, 1] = CellValue("hello friend", textstyle)
        @test sheet[3, 1] == "hello friend"
        @test XLSX.getcell(sheet, 3, 1).style == id(textstyle)

        # Check CellDataFormat only works with CellValues
        @test_throws MethodError XLSX.CellValue([1, 2, 3, 4], textstyle)

        using XLSX: styles_is_datetime, styles_add_numFmt, styles_add_cell_xf
        # Capitalized and single character numfmts
        xfile = XLSX.open_empty_template()
        wb = XLSX.get_workbook(xfile)
        sheet = xfile["Sheet1"]

        fmt = styles_add_numFmt(wb, "yyyy m d")
        style = styles_add_cell_xf(wb, Dict("numFmtId" => "$fmt"))
        @test styles_is_datetime(wb, style)

        fmt = styles_add_numFmt(wb, "h:m:s")
        style = styles_add_cell_xf(wb, Dict("numFmtId" => "$fmt"))
        @test styles_is_datetime(wb, style)

        fmt = styles_add_numFmt(wb, "0.00")
        style = styles_add_cell_xf(wb, Dict("numFmtId" => "$fmt"))
        @test !styles_is_datetime(wb, style)
        @test XLSX.styles_is_float(wb, style)

        fmt = styles_add_numFmt(wb, "[red]yyyy m d")
        style = styles_add_cell_xf(wb, Dict("numFmtId" => "$fmt"))
        @test styles_is_datetime(wb, style)
        fmt = styles_add_numFmt(wb, "[red] h:m:s")
        style = styles_add_cell_xf(wb, Dict("numFmtId" => "$fmt"))
        @test styles_is_datetime(wb, style)
        fmt = styles_add_numFmt(wb, "[red] 0.00; [magenta] 0.00")
        style = styles_add_cell_xf(wb, Dict("numFmtId" => "$fmt"))
        @test !styles_is_datetime(wb, style)
        @test XLSX.styles_is_float(wb, style)

        fmt = styles_add_numFmt(wb, "YYYY M D")
        style = styles_add_cell_xf(wb, Dict("numFmtId" => "$fmt"))
        @test styles_is_datetime(wb, style)
        fmt = styles_add_numFmt(wb, "H:M:S")
        style = styles_add_cell_xf(wb, Dict("numFmtId" => "$fmt"))
        @test styles_is_datetime(wb, style)

        fmt = styles_add_numFmt(wb, "m")
        style = styles_add_cell_xf(wb, Dict("numFmtId" => "$fmt"))
        @test styles_is_datetime(wb, style)

        fmt = styles_add_numFmt(wb, "M")
        style = styles_add_cell_xf(wb, Dict("numFmtId" => "$fmt"))
        @test styles_is_datetime(wb, style)

        fmt = styles_add_numFmt(wb, "y")
        style = styles_add_cell_xf(wb, Dict("numFmtId" => "$fmt"))
        @test !styles_is_datetime(wb, style)

        fmt = styles_add_numFmt(wb, "[s]")
        style = styles_add_cell_xf(wb, Dict("numFmtId" => "$fmt"))
        @test styles_is_datetime(wb, style)

        fmt = styles_add_numFmt(wb, "am/pm")
        style = styles_add_cell_xf(wb, Dict("numFmtId" => "$fmt"))
        @test styles_is_datetime(wb, style)

        fmt = styles_add_numFmt(wb, "a/p")
        style = styles_add_cell_xf(wb, Dict("numFmtId" => "$fmt"))
        @test styles_is_datetime(wb, style)

        fmt = styles_add_numFmt(wb, "\"Monday\"")
        style = styles_add_cell_xf(wb, Dict("numFmtId" => "$fmt"))
        @test !styles_is_datetime(wb, style)
        @test !XLSX.styles_is_float(wb, style)

        fmt = styles_add_numFmt(wb, "0.00*m")
        style = styles_add_cell_xf(wb, Dict("numFmtId" => "$fmt"))
        @test !styles_is_datetime(wb, style)
        @test XLSX.styles_is_float(wb, style)

        fmt = styles_add_numFmt(wb, "0.00_m")
        style = styles_add_cell_xf(wb, Dict("numFmtId" => "$fmt"))
        @test !styles_is_datetime(wb, style)
        @test XLSX.styles_is_float(wb, style)

        fmt = styles_add_numFmt(wb, "0.00\\d")
        style = styles_add_cell_xf(wb, Dict("numFmtId" => "$fmt"))
        @test !styles_is_datetime(wb, style)
        @test XLSX.styles_is_float(wb, style)

        fmt = styles_add_numFmt(wb, "[red][>1.5]000")
        style = styles_add_cell_xf(wb, Dict("numFmtId" => "$fmt"))
        @test !XLSX.styles_is_float(wb, style)

        fmt = styles_add_numFmt(wb, "0.#")
        style = styles_add_cell_xf(wb, Dict("numFmtId" => "$fmt"))
        @test XLSX.styles_is_float(wb, style)

        fmt = styles_add_numFmt(wb, "\"hello.\" ###")
        style = styles_add_cell_xf(wb, Dict("numFmtId" => "$fmt"))
        @test !XLSX.styles_is_float(wb, style)

        fmt = styles_add_numFmt(wb, ".??")
        style = styles_add_cell_xf(wb, Dict("numFmtId" => "$fmt"))
        @test XLSX.styles_is_float(wb, style)

        fmt = styles_add_numFmt(wb, "#e+00")
        style = styles_add_cell_xf(wb, Dict("numFmtId" => "$fmt"))
        @test XLSX.styles_is_float(wb, style)

        fmt = styles_add_numFmt(wb, "0e00")
        style = styles_add_cell_xf(wb, Dict("numFmtId" => "$fmt"))
        @test XLSX.styles_is_float(wb, style)

        fmt = styles_add_numFmt(wb, "# ??/??")
        style = styles_add_cell_xf(wb, Dict("numFmtId" => "$fmt"))
        @test XLSX.styles_is_float(wb, style)

        fmt = styles_add_numFmt(wb, "*.00")
        style = styles_add_cell_xf(wb, Dict("numFmtId" => "$fmt"))
        @test !XLSX.styles_is_float(wb, style)

        fmt = styles_add_numFmt(wb, "\\.00")
        style = styles_add_cell_xf(wb, Dict("numFmtId" => "$fmt"))
        @test !XLSX.styles_is_float(wb, style)

        fmt = styles_add_numFmt(wb, "00_.")
        style = styles_add_cell_xf(wb, Dict("numFmtId" => "$fmt"))
        @test !XLSX.styles_is_float(wb, style)

    end

    @testset "setFont" begin

        f = XLSX.newxlsx()
        s = f[1]
        s["A1:B2,D1:E2"] = ""

        XLSX.setFont(s, "A1:A2"; bold=true, italic=true, size=24, name="Arial")
        XLSX.setFont(s, "B1:B2"; bold=true, italic=false, size=14, name="Aptos")
        XLSX.setFont(s, "D1:D2"; bold=false, italic=true, size=34, name="Berlin Sans FB Demi")
        XLSX.setFont(s, "E1:E2"; bold=false, italic=false, size=4, name="Times New Roman")
        XLSX.setUniformFont(s, "A1:B2,D1:E2"; color="blue") # `setUniformAttribute()` on a non-contiguous range
        @test XLSX.getFont(s, "A1").font == Dict("b" => nothing, "i" => nothing, "sz" => Dict("val" => "24"), "name" => Dict("val" => "Arial"), "color" => Dict("rgb" => "FF0000FF"))
        @test XLSX.getFont(s, "B2").font == Dict("b" => nothing, "i" => nothing, "sz" => Dict("val" => "24"), "name" => Dict("val" => "Arial"), "color" => Dict("rgb" => "FF0000FF"))
        @test XLSX.getFont(s, "D1").font == Dict("b" => nothing, "i" => nothing, "sz" => Dict("val" => "24"), "name" => Dict("val" => "Arial"), "color" => Dict("rgb" => "FF0000FF"))
        @test XLSX.getFont(s, "E2").font == Dict("b" => nothing, "i" => nothing, "sz" => Dict("val" => "24"), "name" => Dict("val" => "Arial"), "color" => Dict("rgb" => "FF0000FF"))

        f = XLSX.newxlsx()
        s = f[1]
        s["A1:Z26"] = ""

        XLSX.setFont(s, "Sheet1!A1:A2"; bold=true, italic=true, size=24, name="Arial", color="blue")
        @test XLSX.getFont(s, "A1").font == Dict("b" => nothing, "i" => nothing, "sz" => Dict("val" => "24"), "name" => Dict("val" => "Arial"), "color" => Dict("rgb" => "FF0000FF"))
        XLSX.setFont(s, "Sheet1!Y:Z"; bold=true, italic=false, size=14, name="Aptos", color="blue")
        @test XLSX.getFont(s, "Y20").font == Dict("b" => nothing, "sz" => Dict("val" => "14"), "name" => Dict("val" => "Aptos"), "color" => Dict("rgb" => "FF0000FF"))
        XLSX.setFont(s, "Sheet1!2:3"; bold=false, italic=true, size=34, name="Berlin Sans FB Demi", color="blue")
        @test XLSX.getFont(s, "M3").font == Dict("i" => nothing, "sz" => Dict("val" => "34"), "name" => Dict("val" => "Berlin Sans FB Demi"), "color" => Dict("rgb" => "FF0000FF"))
        XLSX.setFont(f, "Sheet1!A1:A2"; bold=false, italic=false, size=14, name="Aptos", color="green")
        @test XLSX.getFont(s, "A1").font == Dict("sz" => Dict("val" => "14"), "name" => Dict("val" => "Aptos"), "color" => Dict("rgb" => "FF008000"))
        XLSX.setFont(f, "Sheet1!Y:Z"; bold=false, italic=true, size=24, name="Arial", color="green")
        @test XLSX.getFont(s, "Y20").font == Dict("i" => nothing, "sz" => Dict("val" => "24"), "name" => Dict("val" => "Arial"), "color" => Dict("rgb" => "FF008000"))
        XLSX.setFont(f, "Sheet1!2:3"; bold=true, italic=false, size=24, name="Times New Roman", color="green")
        @test XLSX.getFont(s, "M3").font == Dict("b" => nothing, "sz" => Dict("val" => "24"), "name" => Dict("val" => "Times New Roman"), "color" => Dict("rgb" => "FF008000"))
        XLSX.setFont(s, "E1,E2,G2:G4"; bold=false, italic=false, size=4, name="Times New Roman", color="blue")
        @test XLSX.getFont(s, "G3").font == Dict("sz" => Dict("val" => "4"), "name" => Dict("val" => "Times New Roman"), "color" => Dict("rgb" => "FF0000FF"))
        XLSX.setFont(s, :, 15:16; bold=true, italic=false, size=38, name="Wingdings", color="red")
        @test XLSX.getFont(s, "P10").font == Dict("b" => nothing, "sz" => Dict("val" => "38"), "name" => Dict("val" => "Wingdings"), "color" => Dict("rgb" => "FFFF0000"))
        XLSX.setFont(s, 15:16, :; bold=false, italic=true, size=8, name="Wingdings", color="red")
        @test XLSX.getFont(f, "Sheet1!T16").font == Dict("i" => nothing, "sz" => Dict("val" => "8"), "name" => Dict("val" => "Wingdings"), "color" => Dict("rgb" => "FFFF0000"))
        XLSX.setFont(s, [20, 22, 24], :; bold=false, italic=true, size=48, name="Aptos", color="red")
        @test XLSX.getFont(f, "Sheet1!H22").font == Dict("i" => nothing, "sz" => Dict("val" => "48"), "name" => Dict("val" => "Aptos"), "color" => Dict("rgb" => "FFFF0000"))
        XLSX.setUniformFont(s, [15, 16, 20, 22, 24], :; bold=false, italic=true, size=28, name="Aptos", color="red")
        @test XLSX.getFont(f, "Sheet1!H15").font == Dict("i" => nothing, "sz" => Dict("val" => "28"), "name" => Dict("val" => "Aptos"), "color" => Dict("rgb" => "FFFF0000"))
        @test XLSX.getFont(f, "Sheet1!H22").font == Dict("i" => nothing, "sz" => Dict("val" => "28"), "name" => Dict("val" => "Aptos"), "color" => Dict("rgb" => "FFFF0000"))

        xfile = XLSX.open_empty_template()
        wb = XLSX.get_workbook(xfile)
        sheet = xfile["Sheet1"]

        col_names = ["Integers", "Strings", "Floats", "Booleans", "Dates", "Times", "DateTimes", "AbstractStrings", "Rational", "Irrationals", "MixedStringNothingMissing"]
        data = Vector{Any}(undef, 11)
        data[1] = [1, 2, missing, UInt8(4)]
        data[2] = ["Hey", "You", "Out", "There"]
        data[3] = [101.5, 102.5, missing, 104.5]
        data[4] = [true, false, missing, true]
        data[5] = [Date(2018, 2, 1), Date(2018, 3, 1), Date(2018, 5, 20), Date(2018, 6, 2)]
        data[6] = [Dates.Time(19, 10), Dates.Time(19, 20), Dates.Time(19, 30), Dates.Time(0, 0)]
        data[7] = [Dates.DateTime(2018, 5, 20, 19, 10), Dates.DateTime(2018, 5, 20, 19, 20), Dates.DateTime(2018, 5, 20, 19, 30), Dates.DateTime(2018, 5, 20, 19, 40)]
        data[8] = SubString.(["Hey", "You", "Out", "There"], 1, 2)
        data[9] = [1 // 2, 1 // 3, missing, 22 // 3]
        data[10] = [pi, sqrt(2), missing, sqrt(5)]
        data[11] = [nothing, "middle", missing, nothing]

        XLSX.writetable!(sheet, data, col_names; write_columnnames=true)

        @test isnothing(XLSX.getFont(xfile, "Sheet1!B2")) && isnothing(XLSX.getFont(sheet, "B2"))

        # Default font attributes are present in an empty worksheet until overwritten.
        default_font = XLSX.getDefaultFont(sheet).font
        dname = default_font["name"]["val"]
        dsize = default_font["sz"]["val"]
        dcolorkey = collect(keys(default_font["color"]))[1]
        dcolorval = collect(values(default_font["color"]))[1]

        # Sheet mismatch
        @test_throws XLSX.XLSXError XLSX.setFont(sheet, "S2!A1"; bold=true, size=24, name="Arial")

        XLSX.setFont(sheet, "A1"; bold=true, size=24, name="Arial")
        @test XLSX.getFont(sheet, "A1").font == Dict("b" => nothing, "sz" => Dict("val" => "24"), "name" => Dict("val" => "Arial"), "color" => Dict(dcolorkey => dcolorval))
        XLSX.setFont(sheet, "A1"; size=18)
        @test XLSX.getFont(sheet, "A1").font == Dict("b" => nothing, "sz" => Dict("val" => "18"), "name" => Dict("val" => "Arial"), "color" => Dict(dcolorkey => dcolorval))
        XLSX.setFont(xfile, "Sheet1!A1"; bold=false, size=24, name="Aptos")
        @test XLSX.getFont(xfile, "Sheet1!A1").font == Dict("sz" => Dict("val" => "24"), "name" => Dict("val" => "Aptos"), "color" => Dict(dcolorkey => dcolorval))

        @test XLSX.getFont(xfile, "Sheet1!A1").fontId == XLSX.getFont(sheet, "A1").fontId
        @test XLSX.getFont(xfile, "Sheet1!A1").font == XLSX.getFont(sheet, "A1").font
        @test XLSX.getFont(xfile, "Sheet1!A1").applyFont == XLSX.getFont(sheet, "A1").applyFont

        XLSX.setFont(xfile, "Sheet1!A2"; size=18)
        @test XLSX.getFont(xfile, "Sheet1!A2").font == Dict("sz" => Dict("val" => "18"), "name" => Dict("val" => dname), "color" => Dict(dcolorkey => dcolorval))
        XLSX.setFont(sheet, "A2"; size=24)
        @test XLSX.getFont(sheet, "A2").font == Dict("sz" => Dict("val" => "24"), "name" => Dict("val" => dname), "color" => Dict(dcolorkey => dcolorval))
        XLSX.setFont(sheet, "A2"; size=28, name="Aptos")
        @test XLSX.getFont(sheet, "A2").font == Dict("sz" => Dict("val" => "28"), "name" => Dict("val" => "Aptos"), "color" => Dict(dcolorkey => dcolorval))

        XLSX.setFont(sheet, "A3"; italic=true, name="Berlin Sans FB Demi")
        @test XLSX.getFont(sheet, "A3").font == Dict("i" => nothing, "sz" => Dict("val" => dsize), "name" => Dict("val" => "Berlin Sans FB Demi"), "color" => Dict(dcolorkey => dcolorval))
        XLSX.setFont(xfile, "Sheet1!A3"; size=24)
        @test XLSX.getFont(xfile, "Sheet1!A3").font == Dict("i" => nothing, "sz" => Dict("val" => "24"), "name" => Dict("val" => "Berlin Sans FB Demi"), "color" => Dict(dcolorkey => dcolorval))

        XLSX.setFont(xfile, "Sheet1!A4"; size=28, name="Aptos")
        @test XLSX.getFont(xfile, "Sheet1!A4").font == Dict("sz" => Dict("val" => "28"), "name" => Dict("val" => "Aptos"), "color" => Dict(dcolorkey => dcolorval))

        XLSX.setFont(sheet, "B1"; bold=true, italic=true, size=14, color="FF00FF00")
        @test XLSX.getFont(sheet, "B1").font == Dict("b" => nothing, "i" => nothing, "sz" => Dict("val" => "14"), "name" => Dict("val" => dname), "color" => Dict("rgb" => "FF00FF00"))
        XLSX.setFont(xfile, "Sheet1!B1"; bold=false, italic=false, size=12, name="Berlin Sans FB Demi")
        @test XLSX.getFont(xfile, "Sheet1!B1").font == Dict("sz" => Dict("val" => "12"), "name" => Dict("val" => "Berlin Sans FB Demi"), "color" => Dict("rgb" => "FF00FF00"))

        XLSX.setFont(sheet, "B2"; color="FF000000")
        @test XLSX.getFont(sheet, "B2").font == Dict("sz" => Dict("val" => dsize), "name" => Dict("val" => dname), "color" => Dict("rgb" => "FF000000"))
        XLSX.setFont(xfile, "Sheet1!B2"; bold=true, italic=true, size=14, color="FF00FF00")
        @test XLSX.getFont(xfile, "Sheet1!B2").font == Dict("b" => nothing, "i" => nothing, "sz" => Dict("val" => "14"), "name" => Dict("val" => dname), "color" => Dict("rgb" => "FF00FF00"))

        XLSX.setFont(xfile, "Sheet1!B3"; name="Berlin Sans FB Demi", color="FF000000")
        @test XLSX.getFont(xfile, "Sheet1!B3").font == Dict("sz" => Dict("val" => dsize), "name" => Dict("val" => "Berlin Sans FB Demi"), "color" => Dict("rgb" => "FF000000"))

        XLSX.setFont(sheet, "A1:B2"; size=18, name="Arial")
        @test XLSX.getFont(xfile, "Sheet1!A1").font == Dict("sz" => Dict("val" => "18"), "name" => Dict("val" => "Arial"), "color" => Dict(dcolorkey => dcolorval))
        @test XLSX.getFont(sheet, "A2").font == Dict("sz" => Dict("val" => "18"), "name" => Dict("val" => "Arial"), "color" => Dict(dcolorkey => dcolorval))
        @test XLSX.getFont(xfile, "Sheet1!B1").font == Dict("sz" => Dict("val" => "18"), "name" => Dict("val" => "Arial"), "color" => Dict("rgb" => "FF00FF00"))
        @test XLSX.getFont(xfile, "Sheet1!B2").font == Dict("b" => nothing, "i" => nothing, "sz" => Dict("val" => "18"), "name" => Dict("val" => "Arial"), "color" => Dict("rgb" => "FF00FF00"))


        XLSX.writexlsx("output.xlsx", xfile, overwrite=true)
        @test isfile("output.xlsx")

        XLSX.openxlsx("output.xlsx") do f # Check the updated fonts were written correctly
            s = f["Sheet1"]
            @test XLSX.getFont(f, "Sheet1!A1").font == Dict("sz" => Dict("val" => "18"), "name" => Dict("val" => "Arial"), "color" => Dict(dcolorkey => dcolorval))
            @test XLSX.getFont(s, "A2").font == Dict("sz" => Dict("val" => "18"), "name" => Dict("val" => "Arial"), "color" => Dict(dcolorkey => dcolorval))
            @test XLSX.getFont(f, "Sheet1!A3").font == Dict("i" => nothing, "sz" => Dict("val" => "24"), "name" => Dict("val" => "Berlin Sans FB Demi"), "color" => Dict(dcolorkey => dcolorval))
            @test XLSX.getFont(s, "A4").font == Dict("sz" => Dict("val" => "28"), "name" => Dict("val" => "Aptos"), "color" => Dict(dcolorkey => dcolorval))
            @test XLSX.getFont(f, "Sheet1!B1").font == Dict("sz" => Dict("val" => "18"), "name" => Dict("val" => "Arial"), "color" => Dict("rgb" => "FF00FF00"))
            @test XLSX.getFont(s, "B2").font == Dict("b" => nothing, "i" => nothing, "sz" => Dict("val" => "18"), "name" => Dict("val" => "Arial"), "color" => Dict("rgb" => "FF00FF00"))
            @test XLSX.getFont(f, "Sheet1!B3").font == Dict("sz" => Dict("val" => dsize), "name" => Dict("val" => "Berlin Sans FB Demi"), "color" => Dict("rgb" => "FF000000"))
        end

        # Now try a range
        XLSX.setUniformFont(sheet, "A1:B4"; size=12, name="Times New Roman", color="FF040404")
        @test XLSX.getFont(xfile, "Sheet1!A1").font == Dict("sz" => Dict("val" => "12"), "name" => Dict("val" => "Times New Roman"), "color" => Dict("rgb" => "FF040404"))
        @test XLSX.getFont(xfile, "Sheet1!A4").font == Dict("sz" => Dict("val" => "12"), "name" => Dict("val" => "Times New Roman"), "color" => Dict("rgb" => "FF040404"))
        @test XLSX.getFont(xfile, "Sheet1!B3").font == Dict("sz" => Dict("val" => "12"), "name" => Dict("val" => "Times New Roman"), "color" => Dict("rgb" => "FF040404"))
        @test XLSX.getFont(xfile, "Sheet1!B4").font == Dict("sz" => Dict("val" => "12"), "name" => Dict("val" => "Times New Roman"), "color" => Dict("rgb" => "FF040404"))

        isfile("output.xlsx") && rm("output.xlsx")

        f = XLSX.newxlsx()
        sheet = f[1]
        sheet["A1:E5"] = ""
        XLSX.setFont(sheet, :, [1, 2, 3, 4, 5]; size=18, name="Arial", color="FF040404")
        XLSX.setFont(sheet, 1:3, [1, 3]; size=12, name="Aptos", color="FF040408")
        XLSX.setFont(sheet, [4, 5], [2, 4]; size=6, name="Courier New", color="FF040400")
        @test XLSX.getFont(sheet, "A4").font == Dict("sz" => Dict("val" => "18"), "name" => Dict("val" => "Arial"), "color" => Dict("rgb" => "FF040404"))
        @test XLSX.getFont(sheet, "A3").font == Dict("sz" => Dict("val" => "12"), "name" => Dict("val" => "Aptos"), "color" => Dict("rgb" => "FF040408"))
        @test XLSX.getFont(f, "Sheet1!D5").font == Dict("sz" => Dict("val" => "6"), "name" => Dict("val" => "Courier New"), "color" => Dict("rgb" => "FF040400"))
        @test_throws XLSX.XLSXError XLSX.setFont(sheet, "1:10"; size=18, name="Arial", color="FF040404")
        @test_throws XLSX.XLSXError XLSX.setFont(sheet, "A:K"; size=18, name="Arial", color="FF040404")
        @test_throws XLSX.XLSXError XLSX.setFont(f, "Sheet1!garbage"; size=18, name="Arial", color="FF040404")
        @test_throws XLSX.XLSXError XLSX.setFont(sheet, "Sheet1!garbage"; size=18, name="Arial", color="FF040404")
        @test_throws XLSX.XLSXError XLSX.setFont(sheet, "garbage"; size=18, name="Arial", color="FF040404")
        @test_throws XLSX.XLSXError XLSX.setFont(sheet, "garbage1:garbage2"; size=18, name="Arial", color="FF040404")

        f = XLSX.newxlsx()
        sheet = f[1]
        sheet["A1:E5"] = ""
        XLSX.setUniformFont(sheet, "Sheet1!A1:E1"; size=18, name="Arial", color="FF040404")
        @test XLSX.getFont(sheet, "D1").font == Dict("sz" => Dict("val" => "18"), "name" => Dict("val" => "Arial"), "color" => Dict("rgb" => "FF040404"))
        XLSX.setUniformFont(sheet, "Sheet1!2:3"; size=18, name="Arial", color="FF040408")
        @test XLSX.getFont(sheet, "A3").font == Dict("sz" => Dict("val" => "18"), "name" => Dict("val" => "Arial"), "color" => Dict("rgb" => "FF040408"))
        XLSX.setUniformFont(sheet, "Sheet1!D:E"; size=18, name="Arial", color="FF040400")
        @test XLSX.getFont(sheet, "E5").font == Dict("sz" => Dict("val" => "18"), "name" => Dict("val" => "Arial"), "color" => Dict("rgb" => "FF040400"))
        XLSX.setUniformFont(sheet, "A1:E1"; size=18, name="Arial", color="FF040304")
        @test XLSX.getFont(sheet, "D1").font == Dict("sz" => Dict("val" => "18"), "name" => Dict("val" => "Arial"), "color" => Dict("rgb" => "FF040304"))
        XLSX.setUniformFont(sheet, "2:3"; size=18, name="Arial", color="FF040308")
        @test XLSX.getFont(sheet, "A3").font == Dict("sz" => Dict("val" => "18"), "name" => Dict("val" => "Arial"), "color" => Dict("rgb" => "FF040308"))
        XLSX.setUniformFont(sheet, "D:E"; size=18, name="Arial", color="FF040300")
        @test XLSX.getFont(sheet, "E5").font == Dict("sz" => Dict("val" => "18"), "name" => Dict("val" => "Arial"), "color" => Dict("rgb" => "FF040300"))

        f = XLSX.newxlsx()
        sheet = f[1]
        sheet["A1:E5"] = ""
        XLSX.setUniformFont(sheet, :, 1; size=18, name="Arial", color="FF040404")
        @test XLSX.getFont(sheet, "A3").font == Dict("sz" => Dict("val" => "18"), "name" => Dict("val" => "Arial"), "color" => Dict("rgb" => "FF040404"))
        XLSX.setUniformFont(sheet, :, [2, 3]; size=18, name="Arial", color="FF040400")
        @test XLSX.getFont(sheet, "C4").font == Dict("sz" => Dict("val" => "18"), "name" => Dict("val" => "Arial"), "color" => Dict("rgb" => "FF040400"))
        XLSX.setUniformFont(sheet, [1, 3, 4], 5; size=18, name="Arial", color="FF040300")
        @test XLSX.getFont(sheet, "E1").font == Dict("sz" => Dict("val" => "18"), "name" => Dict("val" => "Arial"), "color" => Dict("rgb" => "FF040300"))
        XLSX.setUniformFont(sheet, 5, [3, 4]; size=18, name="Arial", color="FF030300")
        @test XLSX.getFont(sheet, "D5").font == Dict("sz" => Dict("val" => "18"), "name" => Dict("val" => "Arial"), "color" => Dict("rgb" => "FF030300"))
        XLSX.setUniformFont(sheet, [2, 3, 4], [3, 4]; size=18, name="Arial", color="FF030308")
        @test XLSX.getFont(sheet, "C3").font == Dict("sz" => Dict("val" => "18"), "name" => Dict("val" => "Arial"), "color" => Dict("rgb" => "FF030308"))
        XLSX.setUniformFont(sheet, 4:5, 4; size=18, name="Arial", color="FF030408")
        @test XLSX.getFont(sheet, "D4").font == Dict("sz" => Dict("val" => "18"), "name" => Dict("val" => "Arial"), "color" => Dict("rgb" => "FF030408"))
        XLSX.setUniformFont(sheet, :; size=8, name="Arial", color="FF030408")
        @test XLSX.getFont(sheet, "D4").font == Dict("sz" => Dict("val" => "8"), "name" => Dict("val" => "Arial"), "color" => Dict("rgb" => "FF030408"))
        XLSX.setUniformFont(sheet, :, :; size=28, name="Arial", color="FF030408")
        @test XLSX.getFont(sheet, "D4").font == Dict("sz" => Dict("val" => "28"), "name" => Dict("val" => "Arial"), "color" => Dict("rgb" => "FF030408"))
        @test_throws XLSX.XLSXError XLSX.setUniformFont(sheet, :, [1, 3, 10, 15]; size=18, name="Arial", color="FF040404")
        @test_throws XLSX.XLSXError XLSX.setUniformFont(sheet, [1, 3, 10, 15], :; size=18, name="Arial", color="FF040404")
        @test_throws XLSX.XLSXError XLSX.setUniformFont(sheet, 1, [1, 3, 10, 15]; size=18, name="Arial", color="FF040404")
        @test_throws XLSX.XLSXError XLSX.setUniformFont(sheet, [1, 3, 10, 15], 2:3; size=18, name="Arial", color="FF040404")
        @test_throws XLSX.XLSXError XLSX.setUniformFont(f, "Sheet1!garbage"; size=18, name="Arial", color="FF040404")
        @test_throws XLSX.XLSXError XLSX.setUniformFont(sheet, "Sheet1!garbage"; size=18, name="Arial", color="FF040404")
        @test_throws XLSX.XLSXError XLSX.setUniformFont(sheet, "garbage"; size=18, name="Arial", color="FF040404")
        @test_throws XLSX.XLSXError XLSX.setUniformFont(sheet, "garbage1:garbage2"; size=18, name="Arial", color="FF040404")
    end

    @testset "setBorder" begin
        f = XLSX.open_xlsx_template(joinpath(data_directory, "Borders.xlsx"))
        s = f["Sheet1"]

        @test XLSX.getDefaultBorders(s).border == Dict("left" => nothing, "bottom" => nothing, "right" => nothing, "top" => nothing, "diagonal" => nothing)

        @test isnothing(XLSX.getBorder(s, "A1"))

        @test XLSX.getBorder(s, "B2").border == Dict("left" => Dict("auto" => "1", "style" => "medium"), "bottom" => Dict("auto" => "1", "style" => "medium"), "right" => Dict("auto" => "1", "style" => "medium"), "top" => Dict("auto" => "1", "style" => "medium"), "diagonal" => nothing)
        @test XLSX.getBorder(f, "Sheet1!B4").border == Dict("left" => Dict("rgb" => "FFFF0000", "style" => "hair"), "bottom" => Dict("rgb" => "FFFF0000", "style" => "hair"), "right" => Dict("rgb" => "FFFF0000", "style" => "hair"), "top" => Dict("rgb" => "FFFF0000", "style" => "hair"), "diagonal" => nothing)
        @test XLSX.getBorder(s, "D4").border == Dict("left" => Dict("theme" => "3", "style" => "dashed", "tint" => "0.24994659260841701"), "bottom" => Dict("theme" => "3", "style" => "dashed", "tint" => "0.24994659260841701"), "right" => Dict("theme" => "3", "style" => "dashed", "tint" => "0.24994659260841701"), "top" => Dict("theme" => "3", "style" => "dashed", "tint" => "0.24994659260841701"), "diagonal" => nothing)
        @test XLSX.getBorder(f, "Sheet1!D6").border == Dict("left" => Dict("auto" => "1", "style" => "thick"), "bottom" => Dict("auto" => "1", "style" => "thick"), "right" => Dict("auto" => "1", "style" => "thick"), "top" => Dict("auto" => "1", "style" => "thick"), "diagonal" => nothing)

        XLSX.setBorder(f, "Sheet1!D6"; left=["style" => "dotted", "color" => "FF000FF0"], right=["style" => "medium", "color" => "FF765000"], top=["style" => "thick", "color" => "FF230000"], bottom=["style" => "medium", "color" => "FF0000FF"], diagonal=["style" => "dotted", "color" => "FF00D4D4"])
        @test XLSX.getBorder(s, "D6").border == Dict("left" => Dict("rgb" => "FF000FF0", "style" => "dotted"), "bottom" => Dict("rgb" => "FF0000FF", "style" => "medium"), "right" => Dict("rgb" => "FF765000", "style" => "medium"), "top" => Dict("rgb" => "FF230000", "style" => "thick"), "diagonal" => Dict("rgb" => "FF00D4D4", "style" => "dotted", "direction" => "both"))

        XLSX.setBorder(f, "Sheet1!B2:D4"; left=["style" => "hair"], right=["color" => "FF111111"], top=["style" => "hair"], bottom=["color" => "FF111111"], diagonal=["style" => "hair"])
        @test XLSX.getBorder(s, "B2").border == Dict("left" => Dict("auto" => "1", "style" => "hair"), "bottom" => Dict("rgb" => "FF111111", "style" => "medium"), "right" => Dict("rgb" => "FF111111", "style" => "medium"), "top" => Dict("auto" => "1", "style" => "hair"), "diagonal" => Dict("style" => "hair", "direction" => "both"))
        @test XLSX.getBorder(f, "Sheet1!D4").border == Dict("left" => Dict("theme" => "3", "style" => "hair", "tint" => "0.24994659260841701"), "bottom" => Dict("rgb" => "FF111111", "style" => "dashed"), "right" => Dict("rgb" => "FF111111", "style" => "dashed"), "top" => Dict("theme" => "3", "style" => "hair", "tint" => "0.24994659260841701"), "diagonal" => Dict("style" => "hair", "direction" => "both"))

        XLSX.setBorder(f, "Sheet1!A1:D10"; left=["style" => "hair", "color" => "FF111111"], right=["style" => "hair", "color" => "FF111111"], top=["style" => "hair", "color" => "FF111111"], bottom=["style" => "hair", "color" => "FF111111"], diagonal=["style" => "hair", "color" => "FF111111"])
        @test XLSX.getBorder(s, "B4").border == Dict("left" => Dict("rgb" => "FF111111", "style" => "hair"), "bottom" => Dict("rgb" => "FF111111", "style" => "hair"), "right" => Dict("rgb" => "FF111111", "style" => "hair"), "top" => Dict("rgb" => "FF111111", "style" => "hair"), "diagonal" => Dict("rgb" => "FF111111", "style" => "hair", "direction" => "both"))
        @test XLSX.getBorder(s, "B6").border == Dict("left" => Dict("rgb" => "FF111111", "style" => "hair"), "bottom" => Dict("rgb" => "FF111111", "style" => "hair"), "right" => Dict("rgb" => "FF111111", "style" => "hair"), "top" => Dict("rgb" => "FF111111", "style" => "hair"), "diagonal" => Dict("rgb" => "FF111111", "style" => "hair", "direction" => "both"))
        @test XLSX.getBorder(s, "D4").border == Dict("left" => Dict("rgb" => "FF111111", "style" => "hair"), "bottom" => Dict("rgb" => "FF111111", "style" => "hair"), "right" => Dict("rgb" => "FF111111", "style" => "hair"), "top" => Dict("rgb" => "FF111111", "style" => "hair"), "diagonal" => Dict("rgb" => "FF111111", "style" => "hair", "direction" => "both"))
        @test XLSX.getBorder(s, "D8").border == Dict("left" => Dict("rgb" => "FF111111", "style" => "hair"), "bottom" => Dict("rgb" => "FF111111", "style" => "hair"), "right" => Dict("rgb" => "FF111111", "style" => "hair"), "top" => Dict("rgb" => "FF111111", "style" => "hair"), "diagonal" => Dict("rgb" => "FF111111", "style" => "hair", "direction" => "up"))
        @test XLSX.getBorder(s, "A1").border == Dict("left" => Dict("rgb" => "FF111111", "style" => "hair"), "bottom" => Dict("rgb" => "FF111111", "style" => "hair"), "right" => Dict("rgb" => "FF111111", "style" => "hair"), "top" => Dict("rgb" => "FF111111", "style" => "hair"), "diagonal" => Dict("rgb" => "FF111111", "style" => "hair", "direction" => "both"))
        @test XLSX.getBorder(s, "D10").border == Dict("left" => Dict("rgb" => "FF111111", "style" => "hair"), "bottom" => Dict("rgb" => "FF111111", "style" => "hair"), "right" => Dict("rgb" => "FF111111", "style" => "hair"), "top" => Dict("rgb" => "FF111111", "style" => "hair"), "diagonal" => Dict("rgb" => "FF111111", "style" => "hair", "direction" => "both"))
        @test XLSX.getcell(s, "D11") isa XLSX.EmptyCell
        @test_throws XLSX.XLSXError XLSX.getBorder(s, "D11") # Cannot get a border outside sheet dimension.

        f = XLSX.newxlsx()
        s = f[1]
        s[1:6, 1:6] = ""
        XLSX.setBorder(s, "B2:E5"; outside=["color" => "FFFF0000", "style" => "thick"])
        @test XLSX.getBorder(s, "B2").border == Dict("left" => Dict("rgb" => "FFFF0000", "style" => "thick"), "bottom" => nothing, "right" => nothing, "top" => Dict("rgb" => "FFFF0000", "style" => "thick"), "diagonal" => nothing)
        @test XLSX.getBorder(s, "B3").border == Dict("left" => Dict("rgb" => "FFFF0000", "style" => "thick"), "bottom" => nothing, "right" => nothing, "top" => nothing, "diagonal" => nothing)
        @test XLSX.getBorder(s, "B4").border == Dict("left" => Dict("rgb" => "FFFF0000", "style" => "thick"), "bottom" => nothing, "right" => nothing, "top" => nothing, "diagonal" => nothing)
        @test XLSX.getBorder(s, "B5").border == Dict("left" => Dict("rgb" => "FFFF0000", "style" => "thick"), "bottom" => Dict("rgb" => "FFFF0000", "style" => "thick"), "right" => nothing, "top" => nothing, "diagonal" => nothing)
        @test XLSX.getBorder(s, "C2").border == Dict("left" => nothing, "bottom" => nothing, "right" => nothing, "top" => Dict("rgb" => "FFFF0000", "style" => "thick"), "diagonal" => nothing)
        @test XLSX.getBorder(s, "C3") === nothing
        @test XLSX.getBorder(s, "C4") === nothing
        @test XLSX.getBorder(s, "C5").border == Dict("left" => nothing, "bottom" => Dict("rgb" => "FFFF0000", "style" => "thick"), "right" => nothing, "top" => nothing, "diagonal" => nothing)
        @test XLSX.getBorder(s, "D2").border == Dict("left" => nothing, "bottom" => nothing, "right" => nothing, "top" => Dict("rgb" => "FFFF0000", "style" => "thick"), "diagonal" => nothing)
        @test XLSX.getBorder(s, "D3") === nothing
        @test XLSX.getBorder(s, "D4") === nothing
        @test XLSX.getBorder(s, "D5").border == Dict("left" => nothing, "bottom" => Dict("rgb" => "FFFF0000", "style" => "thick"), "right" => nothing, "top" => nothing, "diagonal" => nothing)
        @test XLSX.getBorder(s, "E2").border == Dict("left" => nothing, "bottom" => nothing, "right" => Dict("rgb" => "FFFF0000", "style" => "thick"), "top" => Dict("rgb" => "FFFF0000", "style" => "thick"), "diagonal" => nothing)
        @test XLSX.getBorder(s, "E3").border == Dict("left" => nothing, "bottom" => nothing, "right" => Dict("rgb" => "FFFF0000", "style" => "thick"), "top" => nothing, "diagonal" => nothing)
        @test XLSX.getBorder(s, "E4").border == Dict("left" => nothing, "bottom" => nothing, "right" => Dict("rgb" => "FFFF0000", "style" => "thick"), "top" => nothing, "diagonal" => nothing)
        @test XLSX.getBorder(s, "E5").border == Dict("left" => nothing, "bottom" => Dict("rgb" => "FFFF0000", "style" => "thick"), "right" => Dict("rgb" => "FFFF0000", "style" => "thick"), "top" => nothing, "diagonal" => nothing)

        XLSX.setBorder(s, "B2:E5"; outside=["color" => "dodgerblue4"])
        @test XLSX.getBorder(s, "B2").border == Dict("left" => Dict("rgb" => "FF104E8B", "style" => "thick"), "bottom" => nothing, "right" => nothing, "top" => Dict("rgb" => "FF104E8B", "style" => "thick"), "diagonal" => nothing)
        @test XLSX.getBorder(s, "B3").border == Dict("left" => Dict("rgb" => "FF104E8B", "style" => "thick"), "bottom" => nothing, "right" => nothing, "top" => nothing, "diagonal" => nothing)
        @test XLSX.getBorder(s, "B4").border == Dict("left" => Dict("rgb" => "FF104E8B", "style" => "thick"), "bottom" => nothing, "right" => nothing, "top" => nothing, "diagonal" => nothing)
        @test XLSX.getBorder(s, "B5").border == Dict("left" => Dict("rgb" => "FF104E8B", "style" => "thick"), "bottom" => Dict("rgb" => "FF104E8B", "style" => "thick"), "right" => nothing, "top" => nothing, "diagonal" => nothing)
        @test XLSX.getBorder(s, "C2").border == Dict("left" => nothing, "bottom" => nothing, "right" => nothing, "top" => Dict("rgb" => "FF104E8B", "style" => "thick"), "diagonal" => nothing)
        @test XLSX.getBorder(s, "C3") === nothing
        @test XLSX.getBorder(s, "C4") === nothing

        f = XLSX.newxlsx()
        s = f[1]
        s[1:6, 1:6] = ""
        XLSX.setBorder(s, "Sheet1!A1"; allsides=["color" => "FFFF00FF", "style" => "thick"])
        @test XLSX.getBorder(s, "A1").border == Dict("left" => Dict("rgb" => "FFFF00FF", "style" => "thick"), "bottom" => Dict("rgb" => "FFFF00FF", "style" => "thick"), "right" => Dict("rgb" => "FFFF00FF", "style" => "thick"), "top" => Dict("rgb" => "FFFF00FF", "style" => "thick"), "diagonal" => nothing)
        XLSX.setBorder(s, "Sheet1!A1:E1"; allsides=["color" => "FFFF0000", "style" => "thick"])
        @test XLSX.getBorder(s, "B1").border == Dict("left" => Dict("rgb" => "FFFF0000", "style" => "thick"), "bottom" => Dict("rgb" => "FFFF0000", "style" => "thick"), "right" => Dict("rgb" => "FFFF0000", "style" => "thick"), "top" => Dict("rgb" => "FFFF0000", "style" => "thick"), "diagonal" => nothing)
        XLSX.setBorder(s, "Sheet1!A:E"; left=["color" => "FFFF0001", "style" => "thick"])
        @test XLSX.getBorder(s, "B2").border == Dict("left" => Dict("rgb" => "FFFF0001", "style" => "thick"), "bottom" => nothing, "right" => nothing, "top" => nothing, "diagonal" => nothing)
        XLSX.setBorder(s, "Sheet1!3:4"; left=["color" => "FFFF0002", "style" => "thick"])
        @test XLSX.getBorder(s, "B4").border == Dict("left" => Dict("rgb" => "FFFF0002", "style" => "thick"), "bottom" => nothing, "right" => nothing, "top" => nothing, "diagonal" => nothing)
        XLSX.setBorder(s, "B2,B4"; left=["color" => "FFFF0004", "style" => "thick"])
        @test XLSX.getBorder(s, "B2").border == Dict("left" => Dict("rgb" => "FFFF0004", "style" => "thick"), "bottom" => nothing, "right" => nothing, "top" => nothing, "diagonal" => nothing)
        @test XLSX.getBorder(s, "B4").border == Dict("left" => Dict("rgb" => "FFFF0004", "style" => "thick"), "bottom" => nothing, "right" => nothing, "top" => nothing, "diagonal" => nothing)
        @test XLSX.getBorder(s, "B3").border == Dict("left" => Dict("rgb" => "FFFF0002", "style" => "thick"), "bottom" => nothing, "right" => nothing, "top" => nothing, "diagonal" => nothing)

        f = XLSX.newxlsx()
        s = f[1]
        s[1:6, 1:6] = ""
        XLSX.setBorder(s, 1, :; left=["color" => "FFFF0001", "style" => "thick"])
        @test XLSX.getBorder(s, "B1").border == Dict("left" => Dict("rgb" => "FFFF0001", "style" => "thick"), "bottom" => nothing, "right" => nothing, "top" => nothing, "diagonal" => nothing)
        XLSX.setBorder(s, [2, 3], :; left=["color" => "FFFF0002", "style" => "thick"])
        @test XLSX.getBorder(s, "D3").border == Dict("left" => Dict("rgb" => "FFFF0002", "style" => "thick"), "bottom" => nothing, "right" => nothing, "top" => nothing, "diagonal" => nothing)
        XLSX.setBorder(s, :, [2, 3]; left=["color" => "FFFF0003", "style" => "thick"])
        @test XLSX.getBorder(s, "C4").border == Dict("left" => Dict("rgb" => "FFFF0003", "style" => "thick"), "bottom" => nothing, "right" => nothing, "top" => nothing, "diagonal" => nothing)
        XLSX.setBorder(s, 4, [2, 3]; left=["color" => "FFFF0004", "style" => "thick"])
        @test XLSX.getBorder(s, "B4").border == Dict("left" => Dict("rgb" => "FFFF0004", "style" => "thick"), "bottom" => nothing, "right" => nothing, "top" => nothing, "diagonal" => nothing)
        XLSX.setBorder(s, 3:2:5, [2, 3]; left=["color" => "FFFF0005", "style" => "thick"])
        @test XLSX.getBorder(s, "C5").border == Dict("left" => Dict("rgb" => "FFFF0005", "style" => "thick"), "bottom" => nothing, "right" => nothing, "top" => nothing, "diagonal" => nothing)
        @test_throws XLSX.XLSXError XLSX.setFont(s, "1:10"; left=["color" => "FFFF0005", "style" => "thick"])
        @test_throws XLSX.XLSXError XLSX.setFont(s, "A:K"; left=["color" => "FFFF0005", "style" => "thick"])

        f = XLSX.open_xlsx_template(joinpath(data_directory, "Borders.xlsx"))
        s = f["Sheet1"]

        XLSX.setUniformBorder(f, "Sheet1!A1:D4"; left=["style" => "dotted", "color" => "darkseagreen3"],
            right=["style" => "medium", "color" => "FF765000"],
            top=["style" => "thick", "color" => "FF230000"],
            bottom=["style" => "medium", "color" => "FF0000FF"],
            diagonal=["style" => "none"]
        )
        @test XLSX.getBorder(s, "A1").border == Dict("left" => Dict("style" => "dotted", "rgb" => "FF9BCD9B"), "bottom" => Dict("style" => "medium", "rgb" => "FF0000FF"), "right" => Dict("style" => "medium", "rgb" => "FF765000"), "top" => Dict("style" => "thick", "rgb" => "FF230000"), "diagonal" => nothing)
        @test XLSX.getBorder(f, "Sheet1!B2").border == Dict("left" => Dict("style" => "dotted", "rgb" => "FF9BCD9B"), "bottom" => Dict("style" => "medium", "rgb" => "FF0000FF"), "right" => Dict("style" => "medium", "rgb" => "FF765000"), "top" => Dict("style" => "thick", "rgb" => "FF230000"), "diagonal" => nothing)
        @test XLSX.getBorder(f, "Sheet1!D4").border == Dict("left" => Dict("style" => "dotted", "rgb" => "FF9BCD9B"), "bottom" => Dict("style" => "medium", "rgb" => "FF0000FF"), "right" => Dict("style" => "medium", "rgb" => "FF765000"), "top" => Dict("style" => "thick", "rgb" => "FF230000"), "diagonal" => nothing)

        @test XLSX.getcell(s, "C3") isa XLSX.EmptyCell
        @test isnothing(XLSX.getBorder(s, "C3"))

        f = XLSX.open_xlsx_template(joinpath(data_directory, "customXml.xlsx"))
        s = f["Mock-up"]

        # Sheet mismatch
        @test_throws XLSX.XLSXError XLSX.setUniformBorder(s, "Document History!A1:D4"; left=["style" => "dotted", "color" => "darkseagreen3"],
            right=["style" => "medium", "color" => "FF765000"],
            top=["style" => "thick", "color" => "FF230000"],
            bottom=["style" => "medium", "color" => "FF0000FF"],
            diagonal=["style" => "none"]
        )

        @test XLSX.setUniformBorder(s, "Mock-up!A1:B4,Mock-up!D4:E6"; left=["style" => "dotted", "color" => "darkseagreen3"],
            right=["style" => "medium", "color" => "FF765000"],
            top=["style" => "thick", "color" => "FF230000"],
            bottom=["style" => "medium", "color" => "FF0000FF"],
            diagonal=["style" => "none"]) == 28

        XLSX.setBorder(s, "ID"; left=["style" => "dotted", "color" => "grey36"], bottom=["style" => "medium", "color" => "FF0000FF"], right=["style" => "medium", "color" => "FF765000"], top=["style" => "thick", "color" => "FF230000"], diagonal=nothing)
        @test XLSX.getBorder(s, "ID").border == Dict("left" => Dict("style" => "dotted", "rgb" => "FF5C5C5C"), "bottom" => Dict("style" => "medium", "rgb" => "FF0000FF"), "right" => Dict("style" => "medium", "rgb" => "FF765000"), "top" => Dict("style" => "thick", "rgb" => "FF230000"), "diagonal" => nothing)

        # Location is a non-contiguous range
        XLSX.setBorder(s, "Location"; left=["style" => "hair", "color" => "chocolate4"], right=["style" => "hair", "color" => "chocolate4"], top=["style" => "hair", "color" => "chocolate4"], bottom=["style" => "hair", "color" => "chocolate4"], diagonal=["style" => "hair", "color" => "chocolate4"])
        @test XLSX.getBorder(s, "D18").border == Dict("left" => Dict("rgb" => "FF8B4513", "style" => "hair"), "bottom" => Dict("rgb" => "FF8B4513", "style" => "hair"), "right" => Dict("rgb" => "FF8B4513", "style" => "hair"), "top" => Dict("rgb" => "FF8B4513", "style" => "hair"), "diagonal" => Dict("rgb" => "FF8B4513", "style" => "hair", "direction" => "both"))
        @test XLSX.getBorder(s, "D20").border == Dict("left" => Dict("rgb" => "FF8B4513", "style" => "hair"), "bottom" => Dict("rgb" => "FF8B4513", "style" => "hair"), "right" => Dict("rgb" => "FF8B4513", "style" => "hair"), "top" => Dict("rgb" => "FF8B4513", "style" => "hair"), "diagonal" => Dict("rgb" => "FF8B4513", "style" => "hair", "direction" => "both"))
        @test XLSX.getBorder(s, "J18").border == Dict("left" => Dict("rgb" => "FF8B4513", "style" => "hair"), "bottom" => Dict("rgb" => "FF8B4513", "style" => "hair"), "right" => Dict("rgb" => "FF8B4513", "style" => "hair"), "top" => Dict("rgb" => "FF8B4513", "style" => "hair"), "diagonal" => Dict("rgb" => "FF8B4513", "style" => "hair", "direction" => "both"))
        @test XLSX.getBorder(s, "J20").border == Dict("left" => Dict("rgb" => "FF8B4513", "style" => "hair"), "bottom" => Dict("rgb" => "FF8B4513", "style" => "hair"), "right" => Dict("rgb" => "FF8B4513", "style" => "hair"), "top" => Dict("rgb" => "FF8B4513", "style" => "hair"), "diagonal" => Dict("rgb" => "FF8B4513", "style" => "hair", "direction" => "both"))

        # Can't get attributes on a range.
        @test_throws XLSX.XLSXError XLSX.getBorder(s, "Contiguous")

        f = XLSX.open_empty_template()
        s = f["Sheet1"]
        s["A1"] = ""
        s["F21"] = ""
        # All these cells are empty.
        @test XLSX.setUniformFont(s, "A2:B4"; size=12, name="Times New Roman", color="chocolate4") == -1
        @test XLSX.setUniformBorder(f, "Sheet1!A2:D4"; left=["style" => "dotted", "color" => "chocolate4"],
            right=["style" => "medium", "color" => "FF765000"],
            top=["style" => "thick", "color" => "FF230000"],
            bottom=["style" => "medium", "color" => "chocolate4"],
            diagonal=["style" => "none"]
        ) == -1
        @test XLSX.setUniformFill(s, "B2:D4"; pattern="gray125", bgColor="FF000000") == -1
        @test XLSX.setFont(s, "A2:F20"; size=18, name="Arial") == -1
        @test XLSX.setBorder(f, "Sheet1!B2:D4"; left=["style" => "hair"], right=["color" => "FF8B4513"], top=["style" => "hair"], bottom=["color" => "chocolate4"], diagonal=["style" => "hair"]) == -1
        @test XLSX.setAlignment(s, "A2:F20"; horizontal="right", wrapText=true) == -1
        @test XLSX.setUniformFill(s, [2, 4], 2:4; pattern="gray125", bgColor="FF000000") == -1
        @test XLSX.setFont(s, [2, 4], 2:4; size=18, name="Arial") == -1
        @test XLSX.setBorder(s, [2, 4], :; left=["style" => "hair"], right=["color" => "FF8B4513"], top=["style" => "hair"], bottom=["color" => "chocolate4"], diagonal=["style" => "hair"]) == -1
        @test XLSX.setAlignment(s, [2, 4], 2:4; horizontal="right", wrapText=true) == -1
        @test XLSX.setUniformFill(s, "B2,C2"; pattern="gray125", bgColor="FF000000") == -1
        @test XLSX.setFont(s, "A2,A4"; size=18, name="Arial") == -1
        @test XLSX.setBorder(f, "Sheet1!B2,Sheet1!C2"; left=["style" => "hair"], right=["color" => "FF8B4513"], top=["style" => "hair"], bottom=["color" => "chocolate4"], diagonal=["style" => "hair"]) == -1
        @test XLSX.setAlignment(s, "A2,B3:C4"; horizontal="right", wrapText=true) == -1
        @test XLSX.setUniformAlignment(s, "B2,D2"; horizontal="right", wrapText=true) == -1
        @test XLSX.setUniformStyle(s, "B2:D2,E3") == -1
        @test_throws XLSX.XLSXError XLSX.setUniformFill(s, "B2,B2"; pattern="gray125", bgColor="FF000000")
        @test_throws XLSX.XLSXError XLSX.setFont(s, "A2,A2"; size=18, name="Arial")
        @test_throws XLSX.XLSXError XLSX.setBorder(f, "Sheet1!B2,Sheet1!B2"; left=["style" => "hair"], right=["color" => "FF8B4513"], top=["style" => "hair"], bottom=["color" => "chocolate4"], diagonal=["style" => "hair"])
        @test_throws XLSX.XLSXError XLSX.setAlignment(s, "A2,A2:A2"; horizontal="right", wrapText=true)
        @test_throws XLSX.XLSXError XLSX.setUniformAlignment(s, "B2,B2"; horizontal="right", wrapText=true)
        @test_throws XLSX.XLSXError XLSX.setUniformStyle(s, "B2:B2,B2")

        f = XLSX.open_empty_template()
        s = f["Sheet1"]
        # All these cells are outside the sheet dimension.
        @test_throws XLSX.XLSXError XLSX.setUniformFont(s, "A1:B4"; size=12, name="Times New Roman", color="chocolate4")
        @test_throws XLSX.XLSXError XLSX.setUniformBorder(f, "Sheet1!A1:D4"; left=["style" => "dotted", "color" => "chocolate4"],
            right=["style" => "medium", "color" => "FF765000"],
            top=["style" => "thick", "color" => "FF230000"],
            bottom=["style" => "medium", "color" => "chocolate4"],
            diagonal=["style" => "none"]
        )
        @test_throws XLSX.XLSXError XLSX.setUniformFill(s, "B2:D4"; pattern="gray125", bgColor="FF000000")
        @test_throws XLSX.XLSXError XLSX.setFont(s, "A1:F20"; size=18, name="Arial")
        @test_throws XLSX.XLSXError XLSX.setBorder(f, "Sheet1!B2:D4"; left=["style" => "hair"], right=["color" => "FF8B4513"], top=["style" => "hair"], bottom=["color" => "chocolate4"], diagonal=["style" => "hair"])
        @test_throws XLSX.XLSXError XLSX.setAlignment(s, "A1:F20"; horizontal="right", wrapText=true)
        @test_throws XLSX.XLSXError XLSX.setFill(f, "Sheet1!A1"; pattern="none", fgColor="88FF8800")
        @test_throws XLSX.XLSXError XLSX.setFont(s, "A1"; size=18, name="Arial")
        @test_throws XLSX.XLSXError XLSX.setBorder(f, "Sheet1!B2"; left=["style" => "hair"], right=["color" => "FF8B4513"], top=["style" => "hair"], bottom=["color" => "FF8B4513"], diagonal=["style" => "hair"])
        @test_throws XLSX.XLSXError XLSX.setFill(s, "F20"; pattern="none", fgColor="88FF8800")

        f = XLSX.open_xlsx_template(joinpath(data_directory, "customXml.xlsx"))
        s = f["Mock-up"]

        # Can't set a uniform attribute to a single cell.
        @test_throws MethodError XLSX.setUniformFill(s, "D4"; pattern="gray125", bgColor="FF000000")
        @test_throws MethodError XLSX.setUniformFill(s, "ID"; pattern="darkTrellis", fgColor="FF222222", bgColor="FFDDDDDD")
        @test_throws MethodError XLSX.setUniformFont(s, "B4"; size=12, name="Times New Roman", color="FF040404")
        @test_throws MethodError XLSX.setUniformBorder(f[2], "B4"; left=["style" => "hair"], right=["color" => "FF8B4513"], top=["style" => "hair"], bottom=["color" => "FF8B4513"], diagonal=["style" => "hair"])
        @test_throws MethodError XLSX.setUniformStyle(s, "ID")
        @test_throws MethodError XLSX.setUniformBorder(f, "Mock-up!D4"; left=["style" => "dotted", "color" => "FF000FF0"],
            right=["style" => "medium", "color" => "FF765000"],
            top=["style" => "thick", "color" => "FF230000"],
            bottom=["style" => "medium", "color" => "FF0000FF"],
            diagonal=["style" => "none"]
        )

        f = XLSX.newxlsx()
        s = f[1]
        s[1:6, 1:6] = ""
        XLSX.setUniformBorder(s, "Sheet1!A:B";
            left=["style" => "dotted", "color" => "darkseagreen3"],
            right=["style" => "medium", "color" => "FF765000"],
            top=["style" => "thick", "color" => "FF230000"],
            bottom=["style" => "medium", "color" => "FF0000FF"],
            diagonal=["style" => "none"]
        )
        @test XLSX.getBorder(s, "A1").border == Dict("left" => Dict("style" => "dotted", "rgb" => "FF9BCD9B"), "bottom" => Dict("style" => "medium", "rgb" => "FF0000FF"), "right" => Dict("style" => "medium", "rgb" => "FF765000"), "top" => Dict("style" => "thick", "rgb" => "FF230000"), "diagonal" => nothing)
        XLSX.setUniformBorder(s, "Sheet1!2:4";
            left=["style" => "dotted", "color" => "FF9BCD9C"],
            right=["style" => "medium", "color" => "FF765000"],
            top=["style" => "thick", "color" => "FF230000"],
            bottom=["style" => "medium", "color" => "FF0000FF"],
            diagonal=["style" => "none"]
        )
        @test XLSX.getBorder(s, "C2").border == Dict("left" => Dict("style" => "dotted", "rgb" => "FF9BCD9C"), "bottom" => Dict("style" => "medium", "rgb" => "FF0000FF"), "right" => Dict("style" => "medium", "rgb" => "FF765000"), "top" => Dict("style" => "thick", "rgb" => "FF230000"), "diagonal" => nothing)
        XLSX.setUniformBorder(s, "A:B";
            left=["style" => "dotted", "color" => "FF9BCD9E"],
            right=["style" => "medium", "color" => "FF765000"],
            top=["style" => "thick", "color" => "FF230000"],
            bottom=["style" => "medium", "color" => "FF0000FF"],
            diagonal=["style" => "none"]
        )
        @test XLSX.getBorder(s, "B3").border == Dict("left" => Dict("style" => "dotted", "rgb" => "FF9BCD9E"), "bottom" => Dict("style" => "medium", "rgb" => "FF0000FF"), "right" => Dict("style" => "medium", "rgb" => "FF765000"), "top" => Dict("style" => "thick", "rgb" => "FF230000"), "diagonal" => nothing)
        XLSX.setUniformBorder(s, "2:4";
            left=["style" => "dotted", "color" => "FF9BCD9D"],
            right=["style" => "medium", "color" => "FF765000"],
            top=["style" => "thick", "color" => "FF230000"],
            bottom=["style" => "medium", "color" => "FF0000FF"],
            diagonal=["style" => "none"]
        )
        @test XLSX.getBorder(s, "C3").border == Dict("left" => Dict("style" => "dotted", "rgb" => "FF9BCD9D"), "bottom" => Dict("style" => "medium", "rgb" => "FF0000FF"), "right" => Dict("style" => "medium", "rgb" => "FF765000"), "top" => Dict("style" => "thick", "rgb" => "FF230000"), "diagonal" => nothing)
        XLSX.setUniformBorder(s, 5, :;
            left=["style" => "dotted", "color" => "FF9BCD8D"],
            right=["style" => "medium", "color" => "FF765000"],
            top=["style" => "thick", "color" => "FF230000"],
            bottom=["style" => "medium", "color" => "FF0000FF"],
            diagonal=["style" => "none"]
        )
        @test XLSX.getBorder(s, "F5").border == Dict("left" => Dict("style" => "dotted", "rgb" => "FF9BCD8D"), "bottom" => Dict("style" => "medium", "rgb" => "FF0000FF"), "right" => Dict("style" => "medium", "rgb" => "FF765000"), "top" => Dict("style" => "thick", "rgb" => "FF230000"), "diagonal" => nothing)
        XLSX.setUniformBorder(s, :, 5;
            left=["style" => "dotted", "color" => "FF9BBD8D"],
            right=["style" => "medium", "color" => "FF765000"],
            top=["style" => "thick", "color" => "FF230000"],
            bottom=["style" => "medium", "color" => "FF0000FF"],
            diagonal=["style" => "none"]
        )
        @test XLSX.getBorder(s, "E2").border == Dict("left" => Dict("style" => "dotted", "rgb" => "FF9BBD8D"), "bottom" => Dict("style" => "medium", "rgb" => "FF0000FF"), "right" => Dict("style" => "medium", "rgb" => "FF765000"), "top" => Dict("style" => "thick", "rgb" => "FF230000"), "diagonal" => nothing)
        XLSX.setUniformBorder(s, :, :;
            left=["style" => "dotted", "color" => "FF9BCD7D"],
            right=["style" => "medium", "color" => "FF765000"],
            top=["style" => "thick", "color" => "FF230000"],
            bottom=["style" => "medium", "color" => "FF0000FF"],
            diagonal=["style" => "none"]
        )
        @test XLSX.getBorder(s, "F5").border == Dict("left" => Dict("style" => "dotted", "rgb" => "FF9BCD7D"), "bottom" => Dict("style" => "medium", "rgb" => "FF0000FF"), "right" => Dict("style" => "medium", "rgb" => "FF765000"), "top" => Dict("style" => "thick", "rgb" => "FF230000"), "diagonal" => nothing)
        XLSX.setUniformBorder(s, :;
            left=["style" => "dotted", "color" => "FF9BCD6D"],
            right=["style" => "medium", "color" => "FF765000"],
            top=["style" => "thick", "color" => "FF230000"],
            bottom=["style" => "medium", "color" => "FF0000FF"],
            diagonal=["style" => "none"]
        )
        @test XLSX.getBorder(s, "D3").border == Dict("left" => Dict("style" => "dotted", "rgb" => "FF9BCD6D"), "bottom" => Dict("style" => "medium", "rgb" => "FF0000FF"), "right" => Dict("style" => "medium", "rgb" => "FF765000"), "top" => Dict("style" => "thick", "rgb" => "FF230000"), "diagonal" => nothing)
        XLSX.setUniformBorder(s, [2, 3], :;
            left=["style" => "dotted", "color" => "FF9BCE6D"],
            right=["style" => "medium", "color" => "FF765000"],
            top=["style" => "thick", "color" => "FF230000"],
            bottom=["style" => "medium", "color" => "FF0000FF"],
            diagonal=["style" => "none"]
        )
        @test XLSX.getBorder(s, "B2").border == Dict("left" => Dict("style" => "dotted", "rgb" => "FF9BCE6D"), "bottom" => Dict("style" => "medium", "rgb" => "FF0000FF"), "right" => Dict("style" => "medium", "rgb" => "FF765000"), "top" => Dict("style" => "thick", "rgb" => "FF230000"), "diagonal" => nothing)
        XLSX.setUniformBorder(s, :, [2, 3];
            left=["style" => "dotted", "color" => "FF9BCB6D"],
            right=["style" => "medium", "color" => "FF765000"],
            top=["style" => "thick", "color" => "FF230000"],
            bottom=["style" => "medium", "color" => "FF0000FF"],
            diagonal=["style" => "none"]
        )
        @test XLSX.getBorder(s, "C6").border == Dict("left" => Dict("style" => "dotted", "rgb" => "FF9BCB6D"), "bottom" => Dict("style" => "medium", "rgb" => "FF0000FF"), "right" => Dict("style" => "medium", "rgb" => "FF765000"), "top" => Dict("style" => "thick", "rgb" => "FF230000"), "diagonal" => nothing)
        XLSX.setUniformBorder(s, 1, [2, 3];
            left=["style" => "dotted", "color" => "FF8BCB6D"],
            right=["style" => "medium", "color" => "FF765000"],
            top=["style" => "thick", "color" => "FF230000"],
            bottom=["style" => "medium", "color" => "FF0000FF"],
            diagonal=["style" => "none"]
        )
        @test XLSX.getBorder(s, "C1").border == Dict("left" => Dict("style" => "dotted", "rgb" => "FF8BCB6D"), "bottom" => Dict("style" => "medium", "rgb" => "FF0000FF"), "right" => Dict("style" => "medium", "rgb" => "FF765000"), "top" => Dict("style" => "thick", "rgb" => "FF230000"), "diagonal" => nothing)
        XLSX.setUniformBorder(s, [1, 2], [4, 5, 6];
            left=["style" => "dotted", "color" => "FF6BCB6D"],
            right=["style" => "medium", "color" => "FF765000"],
            top=["style" => "thick", "color" => "FF230000"],
            bottom=["style" => "medium", "color" => "FF0000FF"],
            diagonal=["style" => "none"]
        )
        @test XLSX.getBorder(s, "E2").border == Dict("left" => Dict("style" => "dotted", "rgb" => "FF6BCB6D"), "bottom" => Dict("style" => "medium", "rgb" => "FF0000FF"), "right" => Dict("style" => "medium", "rgb" => "FF765000"), "top" => Dict("style" => "thick", "rgb" => "FF230000"), "diagonal" => nothing)
        XLSX.setUniformBorder(s, 4, 4;
            left=["style" => "dotted", "color" => "FF7BCB6D"],
            right=["style" => "medium", "color" => "FF765000"],
            top=["style" => "thick", "color" => "FF230000"],
            bottom=["style" => "medium", "color" => "FF0000FF"],
            diagonal=["style" => "none"]
        )
        @test XLSX.getBorder(s, "D4").border == Dict("left" => Dict("style" => "dotted", "rgb" => "FF7BCB6D"), "bottom" => Dict("style" => "medium", "rgb" => "FF0000FF"), "right" => Dict("style" => "medium", "rgb" => "FF765000"), "top" => Dict("style" => "thick", "rgb" => "FF230000"), "diagonal" => nothing)

        f = XLSX.newxlsx()
        s = f[1]
        s[1:6, 1:6] = ""
        XLSX.setOutsideBorder(s, "Sheet1!A1:A2"; outside=["style" => "dotted", "color" => "FF003FF0"])
        @test XLSX.getBorder(s, "A1").border == Dict("left" => Dict("style" => "dotted", "rgb" => "FF003FF0"), "bottom" => nothing, "right" => Dict("style" => "dotted", "rgb" => "FF003FF0"), "top" => Dict("style" => "dotted", "rgb" => "FF003FF0"), "diagonal" => nothing)
        @test XLSX.getBorder(s, "A2").border == Dict("left" => Dict("style" => "dotted", "rgb" => "FF003FF0"), "bottom" => Dict("style" => "dotted", "rgb" => "FF003FF0"), "right" => Dict("style" => "dotted", "rgb" => "FF003FF0"), "top" => nothing, "diagonal" => nothing)
        XLSX.setOutsideBorder(s, "Sheet1!C:E"; outside=["style" => "dotted", "color" => "FF000FF0"])
        @test XLSX.getBorder(s, "C1").border == Dict("left" => Dict("style" => "dotted", "rgb" => "FF000FF0"), "bottom" => nothing, "right" => nothing, "top" => Dict("style" => "dotted", "rgb" => "FF000FF0"), "diagonal" => nothing)
        @test XLSX.getBorder(s, "E6").border == Dict("left" => nothing, "bottom" => Dict("style" => "dotted", "rgb" => "FF000FF0"), "right" => Dict("style" => "dotted", "rgb" => "FF000FF0"), "top" => nothing, "diagonal" => nothing)
        XLSX.setOutsideBorder(s, "Sheet1!3:5"; outside=["style" => "dotted", "color" => "FF000FFF"])
        @test XLSX.getBorder(s, "A3").border == Dict("left" => Dict("style" => "dotted", "rgb" => "FF000FFF"), "bottom" => nothing, "right" => nothing, "top" => Dict("style" => "dotted", "rgb" => "FF000FFF"), "diagonal" => nothing)
        @test XLSX.getBorder(s, "F5").border == Dict("left" => nothing, "bottom" => Dict("style" => "dotted", "rgb" => "FF000FFF"), "right" => Dict("style" => "dotted", "rgb" => "FF000FFF"), "top" => nothing, "diagonal" => nothing)
        XLSX.setOutsideBorder(s, "C:E"; outside=["style" => "dotted", "color" => "FFFF0FF0"])
        @test XLSX.getBorder(s, "C1").border == Dict("left" => Dict("style" => "dotted", "rgb" => "FFFF0FF0"), "bottom" => nothing, "right" => nothing, "top" => Dict("style" => "dotted", "rgb" => "FFFF0FF0"), "diagonal" => nothing)
        @test XLSX.getBorder(s, "E6").border == Dict("left" => nothing, "bottom" => Dict("style" => "dotted", "rgb" => "FFFF0FF0"), "right" => Dict("style" => "dotted", "rgb" => "FFFF0FF0"), "top" => nothing, "diagonal" => nothing)
        XLSX.setOutsideBorder(s, "3:5"; outside=["style" => "dotted", "color" => "FFF50FFF"])
        @test XLSX.getBorder(s, "A3").border == Dict("left" => Dict("style" => "dotted", "rgb" => "FFF50FFF"), "bottom" => nothing, "right" => nothing, "top" => Dict("style" => "dotted", "rgb" => "FFF50FFF"), "diagonal" => nothing)
        @test XLSX.getBorder(s, "F5").border == Dict("left" => nothing, "bottom" => Dict("style" => "dotted", "rgb" => "FFF50FFF"), "right" => Dict("style" => "dotted", "rgb" => "FFF50FFF"), "top" => nothing, "diagonal" => nothing)

        f = XLSX.newxlsx()
        s = f[1]
        s[1:6, 1:6] = ""
        XLSX.setOutsideBorder(s, 1, :; outside=["style" => "dotted", "color" => "FF002FF0"])
        @test XLSX.getBorder(s, "A1").border == Dict("left" => Dict("style" => "dotted", "rgb" => "FF002FF0"), "bottom" => Dict("style" => "dotted", "rgb" => "FF002FF0"), "right" => nothing, "top" => Dict("style" => "dotted", "rgb" => "FF002FF0"), "diagonal" => nothing)
        @test XLSX.getBorder(s, "F1").border == Dict("left" => nothing, "bottom" => Dict("style" => "dotted", "rgb" => "FF002FF0"), "right" => Dict("style" => "dotted", "rgb" => "FF002FF0"), "top" => Dict("style" => "dotted", "rgb" => "FF002FF0"), "diagonal" => nothing)
        XLSX.setOutsideBorder(s, :, 1; outside=["style" => "dotted", "color" => "FF003FF0"])
        @test XLSX.getBorder(s, "A1").border == Dict("left" => Dict("rgb" => "FF003FF0", "style" => "dotted"), "bottom" => Dict("rgb" => "FF002FF0", "style" => "dotted"), "right" => Dict("rgb" => "FF003FF0", "style" => "dotted"), "top" => Dict("rgb" => "FF003FF0", "style" => "dotted"), "diagonal" => nothing)
        @test XLSX.getBorder(s, "A6").border == Dict("left" => Dict("style" => "dotted", "rgb" => "FF003FF0"), "bottom" => Dict("style" => "dotted", "rgb" => "FF003FF0"), "right" => Dict("style" => "dotted", "rgb" => "FF003FF0"), "top" => nothing, "diagonal" => nothing)
        XLSX.setOutsideBorder(s, :, :; outside=["style" => "dotted", "color" => "FF000FF0"])
        @test XLSX.getBorder(s, "A1").border == Dict("left" => Dict("rgb" => "FF000FF0", "style" => "dotted"), "bottom" => Dict("rgb" => "FF002FF0", "style" => "dotted"), "right" => Dict("rgb" => "FF003FF0", "style" => "dotted"), "top" => Dict("rgb" => "FF000FF0", "style" => "dotted"), "diagonal" => nothing)
        @test XLSX.getBorder(s, "F6").border == Dict("left" => nothing, "bottom" => Dict("style" => "dotted", "rgb" => "FF000FF0"), "right" => Dict("style" => "dotted", "rgb" => "FF000FF0"), "top" => nothing, "diagonal" => nothing)
        XLSX.setOutsideBorder(s, :; outside=["style" => "dotted", "color" => "FF000FFF"])
        @test XLSX.getBorder(s, "A1").border == Dict("left" => Dict("rgb" => "FF000FFF", "style" => "dotted"), "bottom" => Dict("rgb" => "FF002FF0", "style" => "dotted"), "right" => Dict("rgb" => "FF003FF0", "style" => "dotted"), "top" => Dict("rgb" => "FF000FFF", "style" => "dotted"), "diagonal" => nothing)
        @test XLSX.getBorder(s, "F6").border == Dict("left" => nothing, "bottom" => Dict("style" => "dotted", "rgb" => "FF000FFF"), "right" => Dict("style" => "dotted", "rgb" => "FF000FFF"), "top" => nothing, "diagonal" => nothing)
        XLSX.setOutsideBorder(s, 1:2, 1; outside=["style" => "dotted", "color" => "FFFFFFF0"])
        @test XLSX.getBorder(s, "A1").border == Dict("left" => Dict("rgb" => "FFFFFFF0", "style" => "dotted"), "bottom" => Dict("rgb" => "FF002FF0", "style" => "dotted"), "right" => Dict("rgb" => "FFFFFFF0", "style" => "dotted"), "top" => Dict("rgb" => "FFFFFFF0", "style" => "dotted"), "diagonal" => nothing)
        @test XLSX.getBorder(s, "A2").border == Dict("left" => Dict("rgb" => "FFFFFFF0", "style" => "dotted"), "bottom" => Dict("rgb" => "FFFFFFF0", "style" => "dotted"), "right" => Dict("rgb" => "FFFFFFF0", "style" => "dotted"), "top" => nothing, "diagonal" => nothing)

    end

    @testset "setFill" begin

        f = XLSX.open_xlsx_template(joinpath(data_directory, "customXml.xlsx"))
        s = f["Mock-up"]

        @test XLSX.getDefaultFill(s).fill == Dict("patternFill" => Dict("patternType" => "none"))

        @test XLSX.getFill(s, "D17").fill == Dict("patternFill" => Dict("bgindexed" => "64", "patternType" => "solid", "fgtint" => "-9.9978637043366805E-2", "fgtheme" => "2"))
        @test XLSX.getFill(f, "Mock-up!D18").fill == Dict("patternFill" => Dict("bgindexed" => "64", "patternType" => "solid", "fgtint" => "-0.499984740745262", "fgtheme" => "2"))

        XLSX.setFill(s, "D17"; pattern="darkTrellis", fgColor="FF222222", bgColor="FFDDDDDD")
        @test XLSX.getFill(s, "D17").fill == Dict("patternFill" => Dict("bgrgb" => "FFDDDDDD", "patternType" => "darkTrellis", "fgrgb" => "FF222222"))

        XLSX.setFill(s, "ID"; pattern="darkTrellis", fgColor="FF222222", bgColor="FFDDDDDD")
        @test XLSX.getFill(s, "ID").fill == Dict("patternFill" => Dict("bgrgb" => "FFDDDDDD", "patternType" => "darkTrellis", "fgrgb" => "FF222222"))
        @test_throws XLSX.XLSXError XLSX.setFill(s, "ID"; pattern="darkTrellis", fgColor="notAcolor", bgColor="FFDDDDDD")
        @test_throws XLSX.XLSXError XLSX.setFill(s, "ID"; pattern="notApattern", fgColor="FF222222", bgColor="FFDDDDDD")
        @test_throws XLSX.XLSXError XLSX.setFill(s, "ID"; pattern="darkTrellis", fgColor="FF222222", bgColor="FFDDDDDDFF")
        @test_throws XLSX.XLSXError XLSX.setFill(s, "ID"; fgColor="FF222222", bgColor="FFDDDDDDFF")

        # Location is a non-contiguous range
        XLSX.setFill(s, "Location"; pattern="lightVertical") # Default colors unchanged
        @test XLSX.getFill(s, "D18").fill == Dict("patternFill" => Dict("bgindexed" => "64", "patternType" => "lightVertical", "fgtint" => "-0.499984740745262", "fgtheme" => "2"))
        @test XLSX.getFill(s, "D20").fill == Dict("patternFill" => Dict("bgindexed" => "64", "patternType" => "lightVertical", "fgtint" => "-0.499984740745262", "fgtheme" => "2"))
        @test XLSX.getFill(s, "J18").fill == Dict("patternFill" => Dict("bgindexed" => "64", "patternType" => "lightVertical", "fgtint" => "-0.499984740745262", "fgtheme" => "2"))
        @test XLSX.getFill(s, "J18").fill == Dict("patternFill" => Dict("bgindexed" => "64", "patternType" => "lightVertical", "fgtint" => "-0.499984740745262", "fgtheme" => "2"))

        XLSX.setFill(s, "Contiguous"; pattern="lightVertical")  # Default colors unchanged
        @test XLSX.getFill(s, "D23").fill == Dict("patternFill" => Dict("patternType" => "lightVertical", "bgindexed" => "64", "fgtheme" => "0"))
        @test XLSX.getFill(s, "D24").fill == Dict("patternFill" => Dict("patternType" => "lightVertical", "bgindexed" => "64", "fgtheme" => "0"))
        @test XLSX.getFill(s, "D25").fill == Dict("patternFill" => Dict("patternType" => "lightVertical", "bgindexed" => "64", "fgtheme" => "0"))
        @test XLSX.getFill(s, "D26").fill == Dict("patternFill" => Dict("patternType" => "lightVertical", "bgindexed" => "64", "fgtheme" => "0"))
        @test XLSX.getFill(s, "D27").fill == Dict("patternFill" => Dict("patternType" => "lightVertical", "bgindexed" => "64", "fgtheme" => "0"))

        # Can't get attributes on a range.
        @test_throws XLSX.XLSXError XLSX.getFill(s, "Contiguous")

        XLSX.setUniformFill(s, "B3:D5"; pattern="lightGrid", fgColor="FF0000FF", bgColor="FF00FF00")
        @test XLSX.getFill(s, "B3").fill == Dict("patternFill" => Dict("bgrgb" => "FF00FF00", "patternType" => "lightGrid", "fgrgb" => "FF0000FF"))
        @test XLSX.getFill(s, "C4").fill == Dict("patternFill" => Dict("bgrgb" => "FF00FF00", "patternType" => "lightGrid", "fgrgb" => "FF0000FF"))
        @test XLSX.getFill(s, "D5").fill == Dict("patternFill" => Dict("bgrgb" => "FF00FF00", "patternType" => "lightGrid", "fgrgb" => "FF0000FF"))

        XLSX.setFill(s, "ID"; pattern="darkTrellis", fgColor="FF222222", bgColor="FFDDDDDD")
        @test XLSX.getFill(s, "ID").fill == Dict("patternFill" => Dict("bgrgb" => "FFDDDDDD", "patternType" => "darkTrellis", "fgrgb" => "FF222222"))

        # Location is a non-contiguous range
        XLSX.setFill(s, "Location"; pattern="lightVertical")
        @test XLSX.getFill(s, "D18").fill == Dict("patternFill" => Dict("bgindexed" => "64", "patternType" => "lightVertical", "fgtint" => "-0.499984740745262", "fgtheme" => "2"))
        @test XLSX.getFill(f, "Mock-up!D20").fill == Dict("patternFill" => Dict("bgindexed" => "64", "patternType" => "lightVertical", "fgtint" => "-0.499984740745262", "fgtheme" => "2"))
        @test XLSX.getFill(s, "J18").fill == Dict("patternFill" => Dict("bgindexed" => "64", "patternType" => "lightVertical", "fgtint" => "-0.499984740745262", "fgtheme" => "2"))
        @test XLSX.getFill(f, "Mock-up!J18").fill == Dict("patternFill" => Dict("bgindexed" => "64", "patternType" => "lightVertical", "fgtint" => "-0.499984740745262", "fgtheme" => "2"))

        XLSX.setFill(s, "Contiguous"; pattern="lightVertical")
        @test XLSX.getFill(s, "D23").fill == Dict("patternFill" => Dict("patternType" => "lightVertical", "bgindexed" => "64", "fgtheme" => "0"))
        @test XLSX.getFill(f, "Mock-up!D24").fill == Dict("patternFill" => Dict("patternType" => "lightVertical", "bgindexed" => "64", "fgtheme" => "0"))
        @test XLSX.getFill(s, "D25").fill == Dict("patternFill" => Dict("patternType" => "lightVertical", "bgindexed" => "64", "fgtheme" => "0"))
        @test XLSX.getFill(f, "Mock-up!D26").fill == Dict("patternFill" => Dict("patternType" => "lightVertical", "bgindexed" => "64", "fgtheme" => "0"))
        @test XLSX.getFill(s, "D27").fill == Dict("patternFill" => Dict("patternType" => "lightVertical", "bgindexed" => "64", "fgtheme" => "0"))

        # Can't get attributes on a range.
        @test_throws XLSX.XLSXError XLSX.getFill(s, "Contiguous")

        XLSX.writexlsx("output.xlsx", f, overwrite=true)
        @test isfile("output.xlsx")

        XLSX.openxlsx("output.xlsx") do f # Check the updated fonts were written correctly
            s = f["Mock-up"]
            @test XLSX.getFill(s, "D23").fill == Dict("patternFill" => Dict("patternType" => "lightVertical", "bgindexed" => "64", "fgtheme" => "0"))
            @test XLSX.getFill(f, "Mock-up!D24").fill == Dict("patternFill" => Dict("patternType" => "lightVertical", "bgindexed" => "64", "fgtheme" => "0"))
            @test XLSX.getFill(s, "D25").fill == Dict("patternFill" => Dict("patternType" => "lightVertical", "bgindexed" => "64", "fgtheme" => "0"))
            @test XLSX.getFill(f, "Mock-up!D26").fill == Dict("patternFill" => Dict("patternType" => "lightVertical", "bgindexed" => "64", "fgtheme" => "0"))
            @test XLSX.getFill(s, "D27").fill == Dict("patternFill" => Dict("patternType" => "lightVertical", "bgindexed" => "64", "fgtheme" => "0"))
        end

        isfile("output.xlsx") && rm("output.xlsx")

        f = XLSX.newxlsx()
        s = f[1]
        s[1:6, 1:6] = ""
        XLSX.setFill(s, "Sheet1!A1"; pattern="darkTrellis", fgColor="FF222222", bgColor="FFDDDDDD")
        @test XLSX.getFill(s, "A1").fill == Dict("patternFill" => Dict("bgrgb" => "FFDDDDDD", "patternType" => "darkTrellis", "fgrgb" => "FF222222"))
        XLSX.setFill(s, "Sheet1!A2:F2"; pattern="darkTrellis", fgColor="FF222224", bgColor="FFDDDDD4")
        @test XLSX.getFill(s, "A2").fill == Dict("patternFill" => Dict("bgrgb" => "FFDDDDD4", "patternType" => "darkTrellis", "fgrgb" => "FF222224"))
        XLSX.setFill(s, "Sheet1!C:D"; pattern="darkTrellis", fgColor="FF222228", bgColor="FFDDDDD8")
        @test XLSX.getFill(s, "D4").fill == Dict("patternFill" => Dict("bgrgb" => "FFDDDDD8", "patternType" => "darkTrellis", "fgrgb" => "FF222228"))
        XLSX.setFill(s, "Sheet1!5:6"; pattern="darkTrellis", fgColor="FF222220", bgColor="FFDDDDD0")
        @test XLSX.getFill(s, "F5").fill == Dict("patternFill" => Dict("bgrgb" => "FFDDDDD0", "patternType" => "darkTrellis", "fgrgb" => "FF222220"))
        XLSX.setFill(s, "Sheet1!E4:E6,Sheet1!A4"; pattern="darkTrellis", fgColor="FF422220", bgColor="FF4DDDD0")
        @test XLSX.getFill(s, "E4").fill == Dict("patternFill" => Dict("bgrgb" => "FF4DDDD0", "patternType" => "darkTrellis", "fgrgb" => "FF422220"))
        @test XLSX.getFill(s, "E5").fill == Dict("patternFill" => Dict("bgrgb" => "FF4DDDD0", "patternType" => "darkTrellis", "fgrgb" => "FF422220"))
        @test XLSX.getFill(s, "E6").fill == Dict("patternFill" => Dict("bgrgb" => "FF4DDDD0", "patternType" => "darkTrellis", "fgrgb" => "FF422220"))
        @test XLSX.getFill(s, "A4").fill == Dict("patternFill" => Dict("bgrgb" => "FF4DDDD0", "patternType" => "darkTrellis", "fgrgb" => "FF422220"))
        XLSX.setFill(s, :, 2; pattern="darkTrellis", fgColor="FF622220", bgColor="FF6DDDD0")
        @test XLSX.getFill(s, "B4").fill == Dict("patternFill" => Dict("bgrgb" => "FF6DDDD0", "patternType" => "darkTrellis", "fgrgb" => "FF622220"))
        XLSX.setFill(s, [2, 6], :; pattern="darkTrellis", fgColor="FF622222", bgColor="FF6DDDD2")
        @test XLSX.getFill(s, "C2").fill == Dict("patternFill" => Dict("bgrgb" => "FF6DDDD2", "patternType" => "darkTrellis", "fgrgb" => "FF622222"))
        @test XLSX.getFill(s, "F6").fill == Dict("patternFill" => Dict("bgrgb" => "FF6DDDD2", "patternType" => "darkTrellis", "fgrgb" => "FF622222"))
        XLSX.setFill(s, :, [2, 5]; pattern="darkTrellis", fgColor="FF622224", bgColor="FF6DDDD4")
        @test XLSX.getFill(s, "B2").fill == Dict("patternFill" => Dict("bgrgb" => "FF6DDDD4", "patternType" => "darkTrellis", "fgrgb" => "FF622224"))
        @test XLSX.getFill(s, "E3").fill == Dict("patternFill" => Dict("bgrgb" => "FF6DDDD4", "patternType" => "darkTrellis", "fgrgb" => "FF622224"))
        XLSX.setFill(s, 2, [3, 6]; pattern="darkTrellis", fgColor="FF622226", bgColor="FF6DDDD6")
        @test XLSX.getFill(s, "C2").fill == Dict("patternFill" => Dict("bgrgb" => "FF6DDDD6", "patternType" => "darkTrellis", "fgrgb" => "FF622226"))
        XLSX.setFill(s, 2:2:6, [4, 5]; pattern="darkTrellis", fgColor="FF622228", bgColor="FF6DDDD8")
        @test XLSX.getFill(s, "E4").fill == Dict("patternFill" => Dict("bgrgb" => "FF6DDDD8", "patternType" => "darkTrellis", "fgrgb" => "FF622228"))

        f = XLSX.newxlsx()
        s = f[1]
        s[1:6, 1:6] = ""
        XLSX.setUniformFill(s, "Sheet1!A2:F2"; pattern="darkTrellis", fgColor="FF222224", bgColor="FFDDDDD4")
        @test XLSX.getFill(s, "A2").fill == Dict("patternFill" => Dict("bgrgb" => "FFDDDDD4", "patternType" => "darkTrellis", "fgrgb" => "FF222224"))
        XLSX.setUniformFill(s, "Sheet1!C:D"; pattern="darkTrellis", fgColor="FF222228", bgColor="FFDDDDD8")
        @test XLSX.getFill(s, "D4").fill == Dict("patternFill" => Dict("bgrgb" => "FFDDDDD8", "patternType" => "darkTrellis", "fgrgb" => "FF222228"))
        XLSX.setUniformFill(s, "Sheet1!5:6"; pattern="darkTrellis", fgColor="FF222220", bgColor="FFDDDDD0")
        @test XLSX.getFill(s, "F5").fill == Dict("patternFill" => Dict("bgrgb" => "FFDDDDD0", "patternType" => "darkTrellis", "fgrgb" => "FF222220"))
        XLSX.setUniformFill(s, "A2:F2"; pattern="darkTrellis", fgColor="FF222224", bgColor="FFDDDDD4")
        @test XLSX.getFill(s, "A2").fill == Dict("patternFill" => Dict("bgrgb" => "FFDDDDD4", "patternType" => "darkTrellis", "fgrgb" => "FF222224"))
        XLSX.setUniformFill(s, "C:D"; pattern="darkTrellis", fgColor="FF222228", bgColor="FFDDDDD8")
        @test XLSX.getFill(s, "D4").fill == Dict("patternFill" => Dict("bgrgb" => "FFDDDDD8", "patternType" => "darkTrellis", "fgrgb" => "FF222228"))
        XLSX.setUniformFill(s, "5:6"; pattern="darkTrellis", fgColor="FF222220", bgColor="FFDDDDD0")
        @test XLSX.getFill(s, "F5").fill == Dict("patternFill" => Dict("bgrgb" => "FFDDDDD0", "patternType" => "darkTrellis", "fgrgb" => "FF222220"))
        XLSX.setUniformFill(s, "E4:E6,A4"; pattern="darkTrellis", fgColor="FF422220", bgColor="FF4DDDD0")
        @test XLSX.getFill(s, "E4").fill == Dict("patternFill" => Dict("bgrgb" => "FF4DDDD0", "patternType" => "darkTrellis", "fgrgb" => "FF422220"))
        @test XLSX.getFill(s, "E5").fill == Dict("patternFill" => Dict("bgrgb" => "FF4DDDD0", "patternType" => "darkTrellis", "fgrgb" => "FF422220"))
        @test XLSX.getFill(s, "E6").fill == Dict("patternFill" => Dict("bgrgb" => "FF4DDDD0", "patternType" => "darkTrellis", "fgrgb" => "FF422220"))
        @test XLSX.getFill(s, "A4").fill == Dict("patternFill" => Dict("bgrgb" => "FF4DDDD0", "patternType" => "darkTrellis", "fgrgb" => "FF422220"))
        XLSX.setUniformFill(s, :, 2; pattern="darkTrellis", fgColor="FF622220", bgColor="FF6DDDD0")
        @test XLSX.getFill(s, "B4").fill == Dict("patternFill" => Dict("bgrgb" => "FF6DDDD0", "patternType" => "darkTrellis", "fgrgb" => "FF622220"))
        XLSX.setUniformFill(s, [2, 6], :; pattern="darkTrellis", fgColor="FF622222", bgColor="FF6DDDD2")
        @test XLSX.getFill(s, "C2").fill == Dict("patternFill" => Dict("bgrgb" => "FF6DDDD2", "patternType" => "darkTrellis", "fgrgb" => "FF622222"))
        @test XLSX.getFill(s, "F6").fill == Dict("patternFill" => Dict("bgrgb" => "FF6DDDD2", "patternType" => "darkTrellis", "fgrgb" => "FF622222"))
        XLSX.setUniformFill(s, :, [2, 5]; pattern="darkTrellis", fgColor="FF622224", bgColor="FF6DDDD4")
        @test XLSX.getFill(s, "B2").fill == Dict("patternFill" => Dict("bgrgb" => "FF6DDDD4", "patternType" => "darkTrellis", "fgrgb" => "FF622224"))
        @test XLSX.getFill(s, "E3").fill == Dict("patternFill" => Dict("bgrgb" => "FF6DDDD4", "patternType" => "darkTrellis", "fgrgb" => "FF622224"))
        XLSX.setUniformFill(s, 2, [3, 6]; pattern="darkTrellis", fgColor="FF622226", bgColor="FF6DDDD6")
        @test XLSX.getFill(s, "C2").fill == Dict("patternFill" => Dict("bgrgb" => "FF6DDDD6", "patternType" => "darkTrellis", "fgrgb" => "FF622226"))
        XLSX.setUniformFill(s, [2, 3], 5:6; pattern="darkTrellis", fgColor="FF642226", bgColor="FF64DDD6")
        @test XLSX.getFill(s, "F2").fill == Dict("patternFill" => Dict("bgrgb" => "FF64DDD6", "patternType" => "darkTrellis", "fgrgb" => "FF642226"))
        XLSX.setUniformFill(s, 2:2:6, [4, 5]; pattern="darkTrellis", fgColor="FF622228", bgColor="FF6DDDD8")
        @test XLSX.getFill(s, "E4").fill == Dict("patternFill" => Dict("bgrgb" => "FF6DDDD8", "patternType" => "darkTrellis", "fgrgb" => "FF622228"))
        XLSX.setUniformFill(s, :, :; pattern="darkTrellis", fgColor="FF822228", bgColor="FF8DDDD8")
        @test XLSX.getFill(s, "E4").fill == Dict("patternFill" => Dict("bgrgb" => "FF8DDDD8", "patternType" => "darkTrellis", "fgrgb" => "FF822228"))
        XLSX.setUniformFill(s, :; pattern="darkTrellis", fgColor="FF822288", bgColor="FF8DDD88")
        @test XLSX.getFill(s, "E4").fill == Dict("patternFill" => Dict("bgrgb" => "FF8DDD88", "patternType" => "darkTrellis", "fgrgb" => "FF822288"))
        XLSX.setUniformFill(s, :; pattern="darkTrellis", fgColor="FF822288", bgColor="FF8DDD88")
        @test XLSX.getFill(s, "E4").fill == Dict("patternFill" => Dict("bgrgb" => "FF8DDD88", "patternType" => "darkTrellis", "fgrgb" => "FF822288"))
        XLSX.setUniformFill(s, 1, 1:2; pattern="darkTrellis", fgColor="FF822268", bgColor="FF8DDD68")
        @test XLSX.getFill(s, "B1").fill == Dict("patternFill" => Dict("bgrgb" => "FF8DDD68", "patternType" => "darkTrellis", "fgrgb" => "FF822268"))

    end

    @testset "setAlignment" begin

        f = XLSX.newxlsx()
        s = f[1]
        s["A1:Z26"] = ""
        XLSX.setAlignment(s, "Sheet1!A1"; horizontal="right", vertical="justify", wrapText=true)
        @test XLSX.getAlignment(s, "A1").alignment == Dict("alignment" => Dict("horizontal" => "right", "vertical" => "justify", "wrapText" => "1"))
        XLSX.setAlignment(s, "Sheet1!A2:C4"; horizontal="right", vertical="justify", wrapText=true)
        @test XLSX.getAlignment(s, "B3").alignment == Dict("alignment" => Dict("horizontal" => "right", "vertical" => "justify", "wrapText" => "1"))
        XLSX.setAlignment(s, "Sheet1!D:E"; horizontal="right", vertical="justify", wrapText=true)
        @test XLSX.getAlignment(s, "D26").alignment == Dict("alignment" => Dict("horizontal" => "right", "vertical" => "justify", "wrapText" => "1"))
        XLSX.setAlignment(s, "Sheet1!25:26"; horizontal="left", vertical="top", wrapText=false)
        @test XLSX.getAlignment(s, "D26").alignment == Dict("alignment" => Dict("horizontal" => "left", "vertical" => "top", "wrapText" => "0"))
        XLSX.setAlignment(s, "G8,H10,J15:M18"; horizontal="left", vertical="bottom", wrapText=true)
        @test XLSX.getAlignment(s, "G8").alignment == Dict("alignment" => Dict("horizontal" => "left", "vertical" => "bottom", "wrapText" => "1"))
        @test XLSX.getAlignment(s, "H10").alignment == Dict("alignment" => Dict("horizontal" => "left", "vertical" => "bottom", "wrapText" => "1"))
        @test XLSX.getAlignment(s, "L16").alignment == Dict("alignment" => Dict("horizontal" => "left", "vertical" => "bottom", "wrapText" => "1"))

        f = XLSX.newxlsx()
        s = f[1]
        s["A1:Z26"] = ""
        XLSX.setAlignment(s, :, 1:3; horizontal="right", vertical="justify", wrapText=true)
        @test XLSX.getAlignment(s, "B25").alignment == Dict("alignment" => Dict("horizontal" => "right", "vertical" => "justify", "wrapText" => "1"))
        XLSX.setAlignment(s, 8:2:16, :; horizontal="right", vertical="justify", wrapText=true)
        @test XLSX.getAlignment(s, "C12").alignment == Dict("alignment" => Dict("horizontal" => "right", "vertical" => "justify", "wrapText" => "1"))
        XLSX.setAlignment(s, :, [8, 10, 12, 14, 16]; horizontal="right", vertical="top", wrapText=true)
        @test XLSX.getAlignment(s, "L22").alignment == Dict("alignment" => Dict("horizontal" => "right", "vertical" => "top", "wrapText" => "1"))
        XLSX.setAlignment(s, 18, 20:3:26; horizontal="left", vertical="top", wrapText=true)
        @test XLSX.getAlignment(s, "W18").alignment == Dict("alignment" => Dict("horizontal" => "left", "vertical" => "top", "wrapText" => "1"))
        XLSX.setAlignment(s, 18:2:22, 20:3:26; horizontal="left", vertical="bottom", wrapText=true)
        @test XLSX.getAlignment(s, "Z20").alignment == Dict("alignment" => Dict("horizontal" => "left", "vertical" => "bottom", "wrapText" => "1"))

        f = XLSX.open_xlsx_template(joinpath(data_directory, "customXml.xlsx"))
        s = f["Mock-up"]

        @test XLSX.getAlignment(s, "D51").alignment == Dict("alignment" => Dict("horizontal" => "left", "vertical" => "top", "wrapText" => "1"))
        @test XLSX.getAlignment(s, "D18").alignment == Dict("alignment" => Dict("horizontal" => "center", "vertical" => "top"))

        XLSX.setAlignment(f, "Mock-up!D18"; horizontal="right", wrapText=true)
        @test XLSX.getAlignment(s, "D18").alignment == Dict("alignment" => Dict("horizontal" => "right", "vertical" => "top", "wrapText" => "1"))

        @test XLSX.setAlignment(s, "Location"; horizontal="right", wrapText=true) == -1

        XLSX.setUniformAlignment(s, "B3:D5"; horizontal="right", vertical="justify", wrapText=true)
        @test XLSX.getAlignment(s, "B3").alignment == Dict("alignment" => Dict("horizontal" => "right", "vertical" => "justify", "wrapText" => "1"))
        @test XLSX.getAlignment(s, "C4").alignment == Dict("alignment" => Dict("horizontal" => "right", "vertical" => "justify", "wrapText" => "1"))
        @test XLSX.getAlignment(s, "D5").alignment == Dict("alignment" => Dict("horizontal" => "right", "vertical" => "justify", "wrapText" => "1"))

        XLSX.writexlsx("output.xlsx", f, overwrite=true)
        @test isfile("output.xlsx")

        XLSX.openxlsx("output.xlsx") do f # Check the updated fonts were written correctly
            s = f["Mock-up"]
            @test XLSX.getAlignment(s, "B3").alignment == Dict("alignment" => Dict("horizontal" => "right", "vertical" => "justify", "wrapText" => "1"))
            @test XLSX.getAlignment(s, "C4").alignment == Dict("alignment" => Dict("horizontal" => "right", "vertical" => "justify", "wrapText" => "1"))
            @test XLSX.getAlignment(s, "D5").alignment == Dict("alignment" => Dict("horizontal" => "right", "vertical" => "justify", "wrapText" => "1"))
        end

        isfile("output.xlsx") && rm("output.xlsx")

        f = XLSX.newxlsx()
        s = f[1]
        s["A1:Z26"] = ""
        XLSX.setUniformAlignment(s, "Sheet1!E5:E6"; horizontal="right", vertical="justify", wrapText=true)
        @test XLSX.getAlignment(s, "E5").alignment == Dict("alignment" => Dict("horizontal" => "right", "vertical" => "justify", "wrapText" => "1"))
        XLSX.setUniformAlignment(s, "Sheet1!A:A"; horizontal="right", vertical="top", wrapText=true)
        @test XLSX.getAlignment(s, "A23").alignment == Dict("alignment" => Dict("horizontal" => "right", "vertical" => "top", "wrapText" => "1"))
        XLSX.setUniformAlignment(s, "Sheet1!15:24"; horizontal="left", vertical="top", wrapText=true)
        @test XLSX.getAlignment(s, "Q15").alignment == Dict("alignment" => Dict("horizontal" => "left", "vertical" => "top", "wrapText" => "1"))
        @test XLSX.getAlignment(s, "A23").alignment == Dict("alignment" => Dict("horizontal" => "left", "vertical" => "top", "wrapText" => "1"))
        XLSX.setUniformAlignment(s, "A:A"; horizontal="right", vertical="bottom", wrapText=true)
        @test XLSX.getAlignment(s, "A15").alignment == Dict("alignment" => Dict("horizontal" => "right", "vertical" => "bottom", "wrapText" => "1"))
        XLSX.setUniformAlignment(s, "10:12"; horizontal="right", vertical="bottom", wrapText=true)
        @test XLSX.getAlignment(s, "Q11").alignment == Dict("alignment" => Dict("horizontal" => "right", "vertical" => "bottom", "wrapText" => "1"))

        f = XLSX.newxlsx()
        s = f[1]
        s["A1:Z26"] = ""
        XLSX.setUniformAlignment(s, 2, :; horizontal="right", vertical="justify", wrapText=true)
        @test XLSX.getAlignment(s, "E2").alignment == Dict("alignment" => Dict("horizontal" => "right", "vertical" => "justify", "wrapText" => "1"))
        XLSX.setUniformAlignment(s, :, 4:5; horizontal="right", vertical="top", wrapText=true)
        @test XLSX.getAlignment(s, "D23").alignment == Dict("alignment" => Dict("horizontal" => "right", "vertical" => "top", "wrapText" => "1"))
        XLSX.setUniformAlignment(s, :, :; horizontal="left", vertical="top", wrapText=true)
        @test XLSX.getAlignment(s, "Q15").alignment == Dict("alignment" => Dict("horizontal" => "left", "vertical" => "top", "wrapText" => "1"))
        @test XLSX.getAlignment(s, "A23").alignment == Dict("alignment" => Dict("horizontal" => "left", "vertical" => "top", "wrapText" => "1"))
        XLSX.setUniformAlignment(s, :; horizontal="right", vertical="bottom", wrapText=true)
        @test XLSX.getAlignment(s, "A15").alignment == Dict("alignment" => Dict("horizontal" => "right", "vertical" => "bottom", "wrapText" => "1"))
        XLSX.setUniformAlignment(s, :, [8, 12, 14]; horizontal="right", vertical="bottom", wrapText=false)
        @test XLSX.getAlignment(s, "L12").alignment == Dict("alignment" => Dict("horizontal" => "right", "vertical" => "bottom", "wrapText" => "0"))
        XLSX.setUniformAlignment(s, 8:12:20, 3; horizontal="right", vertical="top", wrapText=false)
        @test XLSX.getAlignment(s, "C20").alignment == Dict("alignment" => Dict("horizontal" => "right", "vertical" => "top", "wrapText" => "0"))
        XLSX.setUniformAlignment(s, 8:12:20, [3, 4]; horizontal="justify", vertical="justify", wrapText=false)
        @test XLSX.getAlignment(s, "D8").alignment == Dict("alignment" => Dict("horizontal" => "justify", "vertical" => "justify", "wrapText" => "0"))
        XLSX.setUniformAlignment(s, 8:20, 8; horizontal="justify", vertical="justify", wrapText=true)
        @test XLSX.getAlignment(s, "H15").alignment == Dict("alignment" => Dict("horizontal" => "justify", "vertical" => "justify", "wrapText" => "1"))

        f = XLSX.newxlsx()
        s = f[1]
        s["A1:Z26"] = ""
        XLSX.setAlignment(s, "A1"; horizontal="right", vertical="justify", wrapText=true)
        XLSX.setUniformAlignment(f, "Sheet1!A1,Sheet1!C3,Sheet1!E5:E6")
        @test XLSX.getAlignment(s, "A1").alignment == Dict("alignment" => Dict("horizontal" => "right", "vertical" => "justify", "wrapText" => "1"))
        @test XLSX.getAlignment(s, "C3").alignment == Dict("alignment" => Dict("horizontal" => "right", "vertical" => "justify", "wrapText" => "1"))
        @test XLSX.getAlignment(s, "E5").alignment == Dict("alignment" => Dict("horizontal" => "right", "vertical" => "justify", "wrapText" => "1"))
        @test XLSX.getAlignment(s, "E6").alignment == Dict("alignment" => Dict("horizontal" => "right", "vertical" => "justify", "wrapText" => "1"))
        @test_throws XLSX.XLSXError XLSX.setAlignment(s, 2, [1, 3, 10, 15, 28]; horizontal="right", vertical="justify", wrapText=true)
        @test_throws XLSX.XLSXError XLSX.setAlignment(s, [1, 3, 10, 15, 28], 2:3; horizontal="right", vertical="justify", wrapText=true)
        @test_throws XLSX.XLSXError XLSX.setAlignment(s, :, [1, 3, 10, 15, 28]; horizontal="right", vertical="justify", wrapText=true)
        @test_throws XLSX.XLSXError XLSX.setAlignment(s, [1, 3, 10, 15, 28], :; horizontal="right", vertical="justify", wrapText=true)
        @test_throws XLSX.XLSXError XLSX.setAlignment(s, "Z100"; horizontal="right", vertical="justify", wrapText=true)
        @test_throws XLSX.XLSXError XLSX.setAlignment(s, "Sheet1!Z100"; horizontal="right", vertical="justify", wrapText=true)
        @test_throws XLSX.XLSXError XLSX.setAlignment(s, "Sheet1!E1:F5,Sheet1!Z100"; horizontal="right", vertical="justify", wrapText=true)
        @test_throws XLSX.XLSXError XLSX.setAlignment(f, "Sheet1!garbage"; horizontal="right", vertical="justify", wrapText=true)
        @test_throws XLSX.XLSXError XLSX.setAlignment(s, "Sheet1!garbage"; horizontal="right", vertical="justify", wrapText=true)
        @test_throws XLSX.XLSXError XLSX.setAlignment(s, "garbage"; horizontal="right", vertical="justify", wrapText=true)
        @test_throws XLSX.XLSXError XLSX.setAlignment(s, "garbage1:garbage2"; horizontal="right", vertical="justify", wrapText=true)

        f = XLSX.newxlsx()
        s = f[1]
        s["A1:Z26"] = ""
        XLSX.setAlignment(s, "A1"; horizontal="right", vertical="justify", wrapText=true)
        XLSX.setUniformAlignment(s, 1, 1:2:25)
        @test XLSX.getAlignment(s, 1, 1).alignment == Dict("alignment" => Dict("horizontal" => "right", "vertical" => "justify", "wrapText" => "1"))
        @test XLSX.getAlignment(s, 1, 9).alignment == Dict("alignment" => Dict("horizontal" => "right", "vertical" => "justify", "wrapText" => "1"))
        @test XLSX.getAlignment(s, 1, 19).alignment == Dict("alignment" => Dict("horizontal" => "right", "vertical" => "justify", "wrapText" => "1"))
        @test XLSX.getAlignment(s, 1, 25).alignment == Dict("alignment" => Dict("horizontal" => "right", "vertical" => "justify", "wrapText" => "1"))
        @test XLSX.getAlignment(s, 1, 8) === nothing
        @test XLSX.getAlignment(s, 1, 16) === nothing
        @test XLSX.getAlignment(s, 1, 22) === nothing
        @test XLSX.getAlignment(s, 1, 24) === nothing

        f = XLSX.newxlsx()
        s = f[1]
        s["A1:Z26"] = ""
        XLSX.setAlignment(s, "A2"; horizontal="right", vertical="justify", wrapText=true)
        XLSX.setUniformAlignment(s, 2:2:26, :)
        @test XLSX.getAlignment(s, "A2").alignment == Dict("alignment" => Dict("horizontal" => "right", "vertical" => "justify", "wrapText" => "1"))
        @test XLSX.getAlignment(s, "C4").alignment == Dict("alignment" => Dict("horizontal" => "right", "vertical" => "justify", "wrapText" => "1"))
        @test XLSX.getAlignment(s, "K6").alignment == Dict("alignment" => Dict("horizontal" => "right", "vertical" => "justify", "wrapText" => "1"))
        @test XLSX.getAlignment(s, "Y24").alignment == Dict("alignment" => Dict("horizontal" => "right", "vertical" => "justify", "wrapText" => "1"))
        @test XLSX.getAlignment(s, "A3") === nothing
        @test XLSX.getAlignment(s, "C5") === nothing
        @test XLSX.getAlignment(s, "K7") === nothing
        @test XLSX.getAlignment(s, "Y25") === nothing
        @test_throws XLSX.XLSXError XLSX.setUniformAlignment(s, 2, [1, 3, 10, 15, 28]; horizontal="right", vertical="justify", wrapText=true)
        @test_throws XLSX.XLSXError XLSX.setUniformAlignment(s, [1, 3, 10, 15, 28], 2:3; horizontal="right", vertical="justify", wrapText=true)
        @test_throws XLSX.XLSXError XLSX.setUniformAlignment(s, :, [1, 3, 10, 15, 28]; horizontal="right", vertical="justify", wrapText=true)
        @test_throws XLSX.XLSXError XLSX.setUniformAlignment(s, [1, 3, 10, 15, 28], :; horizontal="right", vertical="justify", wrapText=true)
        @test_throws XLSX.XLSXError XLSX.setUniformAlignment(s, "Z100:Z101"; horizontal="right", vertical="justify", wrapText=true)
        @test_throws XLSX.XLSXError XLSX.setUniformAlignment(s, "Sheet1!Z100:Z101"; horizontal="right", vertical="justify", wrapText=true)
        @test_throws XLSX.XLSXError XLSX.setUniformAlignment(s, "Sheet1!E1:F5,Sheet1!Z100"; horizontal="right", vertical="justify", wrapText=true)
        @test_throws XLSX.XLSXError XLSX.setUniformAlignment(s, "Z100:Z100"; horizontal="right", vertical="justify", wrapText=true)
        @test_throws XLSX.XLSXError XLSX.setUniformAlignment(s, "Sheet1!Z100:Z100"; horizontal="right", vertical="justify", wrapText=true)
        @test_throws XLSX.XLSXError XLSX.setUniformAlignment(s, "Sheet1!Z100,Sheet1!Z100"; horizontal="right", vertical="justify", wrapText=true)
        @test_throws XLSX.XLSXError XLSX.setUniformAlignment(f, "Sheet1!garbage"; horizontal="right", vertical="justify", wrapText=true)
        @test_throws XLSX.XLSXError XLSX.setUniformAlignment(s, "Sheet1!garbage"; horizontal="right", vertical="justify", wrapText=true)
        @test_throws XLSX.XLSXError XLSX.setUniformAlignment(s, "garbage"; horizontal="right", vertical="justify", wrapText=true)
        @test_throws XLSX.XLSXError XLSX.setUniformAlignment(s, "garbage1:garbage2"; horizontal="right", vertical="justify", wrapText=true)

    end

    @testset "setFormat" begin

        f = XLSX.open_empty_template()
        s = f["Sheet1"]

        s["A1"] = "Hello"
        s["B1"] = "World"
        s["A2"] = 2.367
        s["B2"] = 200450023
        s["C1"] = Dates.Date(2018, 2, 1)
        s["C2"] = 0.45
        s["D1"] = 100.24
        s["D2"] = Dates.Time(0, 19, 30)
        s["E1:F5"] = [Dates.Date(2025, 6, 23) Dates.Date(2025, 6, 23);
            Dates.Date(2025, 6, 23) Dates.Date(2025, 6, 23);
            Dates.Date(2025, 6, 23) Dates.Date(2025, 6, 23);
            Dates.Date(2025, 6, 23) Dates.Date(2025, 6, 23);
            Dates.Date(2025, 6, 23) Dates.Date(2025, 6, 23)
        ]

        @test XLSX.setFormat(s, "B2"; format="Scientific") == 48
        @test XLSX.setFormat(s, "C2"; format="Percentage") == 9
        @test XLSX.setFormat(s, "C1"; format="General") == 0
        @test XLSX.setFormat(s, "D2"; format="Currency") == 7
        @test XLSX.setFormat(s, "C1"; format="LongDate") == 15
        @test XLSX.setFormat(s, "D2"; format="Time") == 21

        @test XLSX.getFormat(s, "A2") === nothing
        @test XLSX.getFormat(s, "D2").format == Dict("numFmt" => Dict("numFmtId" => "21", "formatCode" => "h:mm:ss"))
        @test XLSX.getFormat(s, "E2").format == Dict("numFmt" => Dict("numFmtId" => "14", "formatCode" => "m/d/yyyy"))
        @test XLSX.getFormat(s, "F2").format == Dict("numFmt" => Dict("numFmtId" => "14", "formatCode" => "m/d/yyyy"))


        @test XLSX.setFormat(s, "A2"; format="""_-£* #,##0.00_-;-£* #,##0.00_-;_-£* "-"??_-;_-@_-""") == 164
        @test XLSX.getFormat(s, "A2").format == Dict("numFmt" => Dict("numFmtId" => "164", "formatCode" => "_-£* #,##0.00_-;-£* #,##0.00_-;_-£* \"-\"??_-;_-@_-"))
        @test XLSX.setFormat(s, "A2"; format="_-£* #,##0.00_-;-£* #,##0.00_-;_-£* \"-\"??_-;_-@_-") == 164
        @test XLSX.getFormat(s, "A2").format == Dict("numFmt" => Dict("numFmtId" => "164", "formatCode" => "_-£* #,##0.00_-;-£* #,##0.00_-;_-£* \"-\"??_-;_-@_-"))

        @test XLSX.setFormat(s, "D2"; format="h:mm AM/PM") == 18
        @test XLSX.setFormat(s, "A2"; format="# ??/??") == 13
        @test XLSX.getFormat(s, "A2").format == Dict("numFmt" => Dict("numFmtId" => "13", "formatCode" => "# ??/??"))
        @test XLSX.getFormat(s, "D2").format == Dict("numFmt" => Dict("numFmtId" => "18", "formatCode" => "h:mm AM/PM"))

        @test XLSX.setFormat(s, "E1:E5"; format="General") == -1
        @test XLSX.setFormat(s, "F1:F5"; format="Currency") == -1
        @test XLSX.getFormat(s, "E2").format == Dict("numFmt" => Dict("numFmtId" => "0", "formatCode" => "General"))
        @test XLSX.getFormat(f, "Sheet1!F2").format == Dict("numFmt" => Dict("numFmtId" => "7", "formatCode" => "\$#,##0.00_);(\$#,##0.00)"))


        @test XLSX.setFormat(f, "Sheet1!E1:F5"; format="#,##0.000") == -1
        @test XLSX.setFormat(s, "F1:F5"; format="#,##0.000") == -1
        @test XLSX.getFormat(s, "E2").format == Dict("numFmt" => Dict("numFmtId" => "165", "formatCode" => "#,##0.000"))
        @test XLSX.getFormat(f, "Sheet1!F2").format == Dict("numFmt" => Dict("numFmtId" => "165", "formatCode" => "#,##0.000"))
        @test_throws XLSX.XLSXError XLSX.setFormat(s, "Z100"; format="Currency")
        @test_throws XLSX.XLSXError XLSX.setFormat(s, "Sheet1!Z100"; format="Currency")
        @test_throws XLSX.XLSXError XLSX.setFormat(s, "Sheet1!E1:F5,Sheet1!Z100"; format="Currency")
        @test_throws XLSX.XLSXError XLSX.setFormat(s, "E2"; format="ffzz345")


        XLSX.writexlsx("test.xlsx", f, overwrite=true)

        XLSX.openxlsx("test.xlsx") do f # Check the updated formats were written correctly
            s = f["Sheet1"]
            @test XLSX.getFormat(s, "A2").format == Dict("numFmt" => Dict("numFmtId" => "13", "formatCode" => "# ??/??"))
            @test XLSX.getFormat(s, "D2").format == Dict("numFmt" => Dict("numFmtId" => "18", "formatCode" => "h:mm AM/PM"))
            @test XLSX.getFormat(s, "E2").format == Dict("numFmt" => Dict("numFmtId" => "165", "formatCode" => "#,##0.000"))
            @test XLSX.getFormat(f, "Sheet1!F2").format == Dict("numFmt" => Dict("numFmtId" => "165", "formatCode" => "#,##0.000"))
        end

        isfile("test.xlsx") && rm("test.xlsx")

        f = XLSX.newxlsx()
        s = f[1]
        s["A1:Z26"] = ""
        XLSX.setFormat(s, "Sheet1!E5"; format="Currency")
        @test XLSX.getFormat(f, "Sheet1!E5").format == Dict("numFmt" => Dict("numFmtId" => "7", "formatCode" => "\$#,##0.00_);(\$#,##0.00)"))
        XLSX.setFormat(s, "Sheet1!W5:X8"; format="Currency")
        @test XLSX.getFormat(f, "Sheet1!X7").format == Dict("numFmt" => Dict("numFmtId" => "7", "formatCode" => "\$#,##0.00_);(\$#,##0.00)"))
        XLSX.setFormat(s, "Sheet1!F:G"; format="Currency")
        @test XLSX.getFormat(s, "F3").format == Dict("numFmt" => Dict("numFmtId" => "7", "formatCode" => "\$#,##0.00_);(\$#,##0.00)"))
        XLSX.setFormat(s, "Sheet1!4:8"; format="Currency")
        @test XLSX.getFormat(s, "Q7").format == Dict("numFmt" => Dict("numFmtId" => "7", "formatCode" => "\$#,##0.00_);(\$#,##0.00)"))
        XLSX.setFormat(s, "N4,M8:M15,Z25:Z26"; format="#,##0.000")
        @test XLSX.getFormat(s, "Z26").format == Dict("numFmt" => Dict("numFmtId" => "164", "formatCode" => "#,##0.000"))
        XLSX.setFormat(s, "Sheet1!4:8"; format="39")
        @test XLSX.getFormat(s, "Q7").format == Dict("numFmt" => Dict("numFmtId" => "39", "formatCode" => "#,##0.00_);(#,##0.00)"))
        XLSX.setFormat(s, "Sheet1!4:8"; format="#,##0")
        @test XLSX.getFormat(s, "Q7").format == Dict("numFmt" => Dict("numFmtId" => "3", "formatCode" => "#,##0"))

        f = XLSX.newxlsx()
        s = f[1]
        s["A1:Z26"] = ""
        XLSX.setFormat(s, :, 2:4; format="Currency")
        @test XLSX.getFormat(f, "Sheet1!B23").format == Dict("numFmt" => Dict("numFmtId" => "7", "formatCode" => "\$#,##0.00_);(\$#,##0.00)"))
        XLSX.setFormat(s, 4:3:10, :; format="Currency")
        @test XLSX.getFormat(s, "Q7").format == Dict("numFmt" => Dict("numFmtId" => "7", "formatCode" => "\$#,##0.00_);(\$#,##0.00)"))
        XLSX.setFormat(s, :, [8, 23, 4]; format="Currency")
        @test XLSX.getFormat(s, "H1").format == Dict("numFmt" => Dict("numFmtId" => "7", "formatCode" => "\$#,##0.00_);(\$#,##0.00)"))
        XLSX.setFormat(s, 25:26, 20:26; format="#,##0.000")
        @test XLSX.getFormat(s, "Z26").format == Dict("numFmt" => Dict("numFmtId" => "164", "formatCode" => "#,##0.000"))
        XLSX.setFormat(s, 25:26, 15; format="#,##0.0000")
        @test XLSX.getFormat(s, "O26").format == Dict("numFmt" => Dict("numFmtId" => "165", "formatCode" => "#,##0.0000"))
        XLSX.setFormat(s, 21:2:25, [15, 16]; format="#,##0.0")
        @test XLSX.getFormat(s, "P25").format == Dict("numFmt" => Dict("numFmtId" => "166", "formatCode" => "#,##0.0"))

        f = XLSX.newxlsx()
        s = f[1]
        s["A1:Z26"] = ""
        XLSX.setUniformFormat(s, "Sheet1!W5:X8"; format="Currency")
        @test XLSX.getFormat(f, "Sheet1!X7").format == Dict("numFmt" => Dict("numFmtId" => "7", "formatCode" => "\$#,##0.00_);(\$#,##0.00)"))
        XLSX.setUniformFormat(s, "Sheet1!F:G"; format="Currency")
        @test XLSX.getFormat(s, "F3").format == Dict("numFmt" => Dict("numFmtId" => "7", "formatCode" => "\$#,##0.00_);(\$#,##0.00)"))
        XLSX.setUniformFormat(s, "Sheet1!4:8"; format="Currency")
        @test XLSX.getFormat(s, "Q7").format == Dict("numFmt" => Dict("numFmtId" => "7", "formatCode" => "\$#,##0.00_);(\$#,##0.00)"))
        XLSX.setUniformFormat(s, "N4,M8:M15,Z25:Z26"; format="#,##0.000")
        @test XLSX.getFormat(s, "Z26").format == Dict("numFmt" => Dict("numFmtId" => "164", "formatCode" => "#,##0.000"))
        @test_throws XLSX.XLSXError XLSX.setUniformFormat(s, "Z100:Z101"; format="Currency")
        @test_throws XLSX.XLSXError XLSX.setUniformFormat(s, "Sheet1!Z100:Z101"; format="Currency")
        @test_throws XLSX.XLSXError XLSX.setUniformFormat(s, "Sheet1!E1:F5,Sheet1!Z100"; format="Currency")

        f = XLSX.newxlsx()
        s = f[1]
        s["A1:Z26"] = ""
        XLSX.setUniformFormat(s, :, 2:4; format="Currency")
        @test XLSX.getFormat(f, "Sheet1!B23").format == Dict("numFmt" => Dict("numFmtId" => "7", "formatCode" => "\$#,##0.00_);(\$#,##0.00)"))
        XLSX.setUniformFormat(s, 4:3:10, :; format="Currency")
        @test XLSX.getFormat(s, "Q7").format == Dict("numFmt" => Dict("numFmtId" => "7", "formatCode" => "\$#,##0.00_);(\$#,##0.00)"))
        XLSX.setUniformFormat(s, :, [8, 23, 4]; format="Currency")
        @test XLSX.getFormat(s, "H1").format == Dict("numFmt" => Dict("numFmtId" => "7", "formatCode" => "\$#,##0.00_);(\$#,##0.00)"))
        XLSX.setUniformFormat(s, 25:26, 20:26; format="#,##0.000")
        @test XLSX.getFormat(s, "Z26").format == Dict("numFmt" => Dict("numFmtId" => "164", "formatCode" => "#,##0.000"))
        XLSX.setUniformFormat(s, 25:26, 15; format="#,##0.0000")
        @test XLSX.getFormat(s, "O26").format == Dict("numFmt" => Dict("numFmtId" => "165", "formatCode" => "#,##0.0000"))
        XLSX.setUniformFormat(s, 21:2:25, [15, 16]; format="#,##0.0")
        @test XLSX.getFormat(s, "P25").format == Dict("numFmt" => Dict("numFmtId" => "166", "formatCode" => "#,##0.0"))

        f = XLSX.newxlsx()
        s = f[1]
        s["A1:Z26"] = ""
        XLSX.setUniformFormat(s, :, :; format="Currency")
        @test XLSX.getFormat(f, "Sheet1!B23").format == Dict("numFmt" => Dict("numFmtId" => "7", "formatCode" => "\$#,##0.00_);(\$#,##0.00)"))
        XLSX.setUniformFormat(s, 4:10, :; format="#,##0.000")
        @test XLSX.getFormat(s, "Q7").format == Dict("numFmt" => Dict("numFmtId" => "164", "formatCode" => "#,##0.000"))
        XLSX.setUniformFormat(s, [8, 23, 4], 8; format="#,##0.0")
        @test XLSX.getFormat(s, "H8").format == Dict("numFmt" => Dict("numFmtId" => "165", "formatCode" => "#,##0.0"))

    end

    @testset "UniformStyle" begin
        f = XLSX.newxlsx()
        s = f[1]
        s["A1:Z26"] = ""

        XLSX.setFont(s, "A1:F5"; size=18, name="Arial")
        cell_style = XLSX.getcell(s, "A1").style
        @test XLSX.setUniformStyle(s, "A1:F5") == cell_style
        @test XLSX.getcell(s, "F5").style == cell_style

        XLSX.setFont(s, "A6:F10"; size=10, name="Aptos")
        cell_style = XLSX.getcell(s, "E6").style
        @test XLSX.setUniformStyle(s, [6, 7, 8, 9, 10], 5) == cell_style
        @test XLSX.getcell(s, "E8").style == cell_style

        XLSX.setFont(s, "A11:F15"; size=10, name="Times New Roman")
        cell_style = XLSX.getcell(s, "E6").style
        @test XLSX.setUniformStyle(s, [6, 7, 8, 9, 10], :) == cell_style
        @test XLSX.getcell(s, "Z8").style == cell_style

        XLSX.setFont(s, "A16"; size=80, name="Ariel")
        cell_style = XLSX.getcell(s, "A16").style
        @test XLSX.setUniformStyle(s, "A16,A15,D20:E25,F25") == cell_style
        @test XLSX.getcell(s, "A15").style == cell_style
        @test XLSX.getcell(s, "D20").style == cell_style
        @test XLSX.getcell(s, "F25").style == cell_style

        XLSX.setFont(s, "A1"; size=8, name="Aptos")
        cell_style = XLSX.getcell(s, "A1").style
        @test XLSX.setUniformStyle(s, :) == cell_style
        @test XLSX.getcell(s, "A1").style == cell_style
        @test XLSX.getcell(s, "M13").style == cell_style
        @test XLSX.getcell(s, "Z26").style == cell_style
        f = XLSX.newxlsx()
        s = f[1]
        s["A1:Z26"] = ""
        XLSX.setFont(s, "A1"; size=8, name="Aptos")
        cell_style = XLSX.getcell(s, "A1").style
        @test XLSX.setUniformStyle(s, "Sheet1!A1:A26") == cell_style
        @test XLSX.getcell(s, "A2").style == cell_style
        @test XLSX.getcell(s, "A13").style == cell_style
        @test XLSX.getcell(s, "A26").style == cell_style
        @test XLSX.setUniformStyle(s, "Sheet1!1:2") == cell_style
        @test XLSX.getcell(s, "B1").style == cell_style
        @test XLSX.getcell(s, "M2").style == cell_style
        @test XLSX.getcell(s, "Z1").style == cell_style
        @test XLSX.setUniformStyle(s, "Sheet1!B:C") == cell_style
        @test XLSX.getcell(s, "C3").style == cell_style
        @test XLSX.getcell(s, "B13").style == cell_style
        @test XLSX.getcell(s, "C26").style == cell_style
        XLSX.setFont(s, "A1"; size=8, name="Arial")
        cell_style = XLSX.getcell(s, "A1").style
        @test XLSX.setUniformStyle(s, "A1:A26") == cell_style
        @test XLSX.getcell(s, "A2").style == cell_style
        @test XLSX.getcell(s, "A13").style == cell_style
        @test XLSX.getcell(s, "A26").style == cell_style
        @test XLSX.setUniformStyle(s, "1:2") == cell_style
        @test XLSX.getcell(s, "B1").style == cell_style
        @test XLSX.getcell(s, "M2").style == cell_style
        @test XLSX.getcell(s, "Z1").style == cell_style
        @test XLSX.setUniformStyle(s, "B:C") == cell_style
        @test XLSX.getcell(s, "C3").style == cell_style
        @test XLSX.getcell(s, "B13").style == cell_style
        @test XLSX.getcell(s, "C26").style == cell_style
        f = XLSX.newxlsx()
        s = f[1]
        s["A1:Z26"] = ""
        XLSX.setFont(s, "A1"; size=8, name="Aptos")
        cell_style = XLSX.getcell(s, "A1").style
        @test XLSX.setUniformStyle(s, 1, :) == cell_style
        @test XLSX.getcell(s, "B1").style == cell_style
        @test XLSX.setUniformStyle(s, :, 2) == cell_style
        @test XLSX.getcell(s, "B13").style == cell_style
        @test XLSX.setUniformStyle(s, :, 5:2:15) == cell_style
        @test XLSX.getcell(s, "E25").style == cell_style
        @test XLSX.setUniformStyle(s, 5:10, [15, 16, 17]) == cell_style
        @test XLSX.getcell(s, "P10").style == cell_style
        @test XLSX.setUniformStyle(s, 5:10, 17:19) == cell_style
        @test XLSX.getcell(s, "S10").style == cell_style
        @test XLSX.setUniformStyle(s, [10, 12, 26], [19, 24, 26]) == cell_style
        @test XLSX.getcell(s, "Z26").style == cell_style
        @test XLSX.setUniformStyle(s, :, :) == cell_style
        @test XLSX.getcell(s, "Y4").style == cell_style
        @test_throws XLSX.XLSXError XLSX.setUniformStyle(s, :, [1, 3, 10, 15, 28])
        @test_throws XLSX.XLSXError XLSX.setUniformStyle(s, [1, 3, 10, 15, 28], :)
        @test_throws XLSX.XLSXError XLSX.setUniformStyle(s, 1, [1, 3, 10, 15, 28])
        @test_throws XLSX.XLSXError XLSX.setUniformStyle(s, [1, 3, 10, 15, 28], 2:3)
        @test_throws XLSX.XLSXError XLSX.setUniformStyle(f, "Sheet1!garbage")
        @test_throws XLSX.XLSXError XLSX.setUniformStyle(s, "Sheet1!garbage")
        @test_throws XLSX.XLSXError XLSX.setUniformStyle(s, "garbage")
        @test_throws XLSX.XLSXError XLSX.setUniformStyle(s, "garbage1:garbage2")

    end

    @testset "Width and height" begin

        f = XLSX.open_empty_template()
        s = f["Sheet1"]

        s["A1"] = "Hello"
        s["B1"] = "World"
        s["A2"] = 2.367
        s["B2"] = 200450023
        s["C1"] = Dates.Date(2018, 2, 1)
        s["C2"] = 0.45
        s["D1"] = 100.24
        s["D2"] = Dates.Time(0, 19, 30)
        s["E1:F5"] = [Dates.Date(2025, 6, 23) Dates.Date(2025, 6, 23);
            Dates.Date(2025, 6, 23) Dates.Date(2025, 6, 23);
            Dates.Date(2025, 6, 23) Dates.Date(2025, 6, 23);
            Dates.Date(2025, 6, 23) Dates.Date(2025, 6, 23);
            Dates.Date(2025, 6, 23) Dates.Date(2025, 6, 23)
        ]

        XLSX.setColumnWidth(s, "A2"; width=30)
        @test XLSX.getColumnWidth(s, "A2") ≈ 30.7109375

        XLSX.setColumnWidth(s, "B2:C2"; width=10.1)
        @test XLSX.getColumnWidth(s, "B3") ≈ 10.8109375
        @test XLSX.getColumnWidth(s, "C4") ≈ 10.8109375

        XLSX.setRowHeight(s, "A2"; height=30)
        @test XLSX.getRowHeight(s, "A2") ≈ 30.2109375

        XLSX.setRowHeight(s, "B2:C5"; height=10.1)
        @test XLSX.getRowHeight(s, "B3") ≈ 10.3109375
        @test XLSX.getRowHeight(s, "C4") ≈ 10.3109375

        # Make sure setting row height doesn't affect column width
        # and vice versa.
        @test XLSX.getColumnWidth(s, "B3") ≈ 10.8109375
        @test XLSX.getColumnWidth(s, "C4") ≈ 10.8109375
        XLSX.setColumnWidth(s, "B2:C2"; width=30.5)
        @test XLSX.getColumnWidth(s, "B3") ≈ 31.2109375
        @test XLSX.getColumnWidth(f, "Sheet1!C4") ≈ 31.2109375
        @test XLSX.getRowHeight(s, "B3") ≈ 10.3109375
        @test XLSX.getRowHeight(f, "Sheet1!C4") ≈ 10.3109375

        f = XLSX.open_xlsx_template(joinpath(data_directory, "customXml.xlsx"))
        s = f["Mock-up"]

        XLSX.setColumnWidth(s, "Location"; width=60)
        XLSX.setRowHeight(s, "Location"; height=50)
        @test XLSX.getRowHeight(s, "D18") ≈ 50.2109375
        @test XLSX.getColumnWidth(s, "D18") ≈ 60.7109375
        @test XLSX.getRowHeight(f, "Mock-up!J20") ≈ 50.2109375
        @test XLSX.getColumnWidth(f, "Mock-up!J20") ≈ 60.7109375

        f = XLSX.newxlsx()
        s = f[1]
        s["A1:Z26"] = ""
        XLSX.setColumnWidth(s, "Sheet1!A1"; width=60)
        @test XLSX.getColumnWidth(s, "A1") ≈ 60.7109375
        XLSX.setColumnWidth(s, "Sheet1!A1:Z1"; width=60)
        @test XLSX.getColumnWidth(s, "R1") ≈ 60.7109375
        XLSX.setColumnWidth(s, "Sheet1!A:B"; width=60)
        @test XLSX.getColumnWidth(s, "B26") ≈ 60.7109375
        XLSX.setColumnWidth(s, "Sheet1!2:3"; width=60)
        @test XLSX.getColumnWidth(s, "R26") ≈ 60.7109375
        XLSX.setColumnWidth(s, "A:B"; width=30.5)
        @test XLSX.getColumnWidth(s, "B26") ≈ 31.2109375
        XLSX.setColumnWidth(s, "2:3"; width=30.5)
        @test XLSX.getColumnWidth(s, "R3") ≈ 31.2109375
        XLSX.setColumnWidth(s, "Sheet1!C5:C7,Sheet1!F5:F7,Sheet1!H7"; width=10.1)
        @test XLSX.getColumnWidth(s, "F26") ≈ 10.8109375
        XLSX.setColumnWidth(s, 5, :; width=10.0)
        @test XLSX.getColumnWidth(s, "Q5") ≈ 10.7109375
        XLSX.setColumnWidth(s, 5:7; width=10.2)
        @test XLSX.getColumnWidth(s, "G22") ≈ 10.9109375
        XLSX.setColumnWidth(s, :, 5:7; width=10.3)
        @test XLSX.getColumnWidth(s, "G22") ≈ 11.0109375
        XLSX.setColumnWidth(s, :, :; width=10.4)
        @test XLSX.getColumnWidth(s, "G22") ≈ 11.1109375
        XLSX.setColumnWidth(s, :; width=10.5)
        @test XLSX.getColumnWidth(s, "G22") ≈ 11.2109375
        XLSX.setColumnWidth(s, 2:3:11, :; width=10.6)
        @test XLSX.getColumnWidth(s, "Z26") ≈ 11.3109375
        XLSX.setColumnWidth(s, 2:3:11; width=10.7)
        @test XLSX.getColumnWidth(s, "E26") ≈ 11.4109375
        XLSX.setColumnWidth(s, :, [2, 3, 11]; width=10.8)
        @test XLSX.getColumnWidth(s, "K15") ≈ 11.5109375
        XLSX.setColumnWidth(s, 3:6, [2, 3, 11]; width=10.9)
        @test XLSX.getColumnWidth(s, "K15") ≈ 11.6109375
        XLSX.setColumnWidth(s, 3:3:6, [2, 3, 11]; width=11.0)
        @test XLSX.getColumnWidth(s, "K15") ≈ 11.7109375
        XLSX.setColumnWidth(s, 11, 7:13; width=11.1)
        @test XLSX.getColumnWidth(s, "K15") ≈ 11.8109375

        f = XLSX.newxlsx()
        s = f[1]
        s["A1:Z26"] = ""
        XLSX.setRowHeight(s, "Sheet1!A1"; height=10.1)
        @test XLSX.getRowHeight(s, "A1") ≈ 10.3109375
        XLSX.setRowHeight(s, "Sheet1!A1:A26"; height=10.2)
        @test XLSX.getRowHeight(s, "R20") ≈ 10.4109375
        XLSX.setRowHeight(s, "Sheet1!A:B"; height=10.3)
        @test XLSX.getRowHeight(s, "B26") ≈ 10.5109375
        XLSX.setRowHeight(s, "Sheet1!2:3"; height=10.4)
        @test XLSX.getRowHeight(s, "R3") ≈ 10.6109375
        XLSX.setRowHeight(s, "A:B"; height=10.5)
        @test XLSX.getRowHeight(s, "B26") ≈ 10.7109375
        XLSX.setRowHeight(s, "2:3"; height=10.6)
        @test XLSX.getRowHeight(s, "R3") ≈ 10.8109375
        XLSX.setRowHeight(s, "Sheet1!C5:C7,Sheet1!F5:F7,Sheet1!H7"; height=10.7)
        @test XLSX.getRowHeight(s, "F6") ≈ 10.9109375
        XLSX.setRowHeight(s, 5, :; height=10.8)
        @test XLSX.getRowHeight(s, "Q5") ≈ 11.0109375
        XLSX.setRowHeight(s, 5:7; height=10.9)
        @test XLSX.getRowHeight(s, "P6") ≈ 11.1109375
        XLSX.setRowHeight(s, :, 5:7; height=11.0)
        @test XLSX.getRowHeight(s, "G22") ≈ 11.2109375
        XLSX.setRowHeight(s, :, :; height=11.1)
        @test XLSX.getRowHeight(s, "G22") ≈ 11.3109375
        XLSX.setRowHeight(s, :; height=11.2)
        @test XLSX.getRowHeight(s, "G22") ≈ 11.4109375
        XLSX.setRowHeight(s, 2:3:11, :; height=11.3)
        @test XLSX.getRowHeight(s, "J8") ≈ 11.5109375
        XLSX.setRowHeight(s, 2:3:11; height=11.4)
        @test XLSX.getRowHeight(s, "J8") ≈ 11.6109375
        XLSX.setRowHeight(s, :, [2, 3, 11]; height=11.5)
        @test XLSX.getRowHeight(s, "K15") ≈ 11.7109375
        XLSX.setRowHeight(s, 3:6, [2, 3, 11]; height=11.6)
        @test XLSX.getRowHeight(s, "K5") ≈ 11.8109375
        XLSX.setRowHeight(s, 3:3:6, [2, 3, 11]; height=11.7)
        @test XLSX.getRowHeight(s, "K6") ≈ 11.9109375
        XLSX.setRowHeight(s, 11, 7:13; height=11.8)
        @test XLSX.getRowHeight(s, "K11") ≈ 12.0109375

    end

    @testset "No cache" begin
        XLSX.openxlsx(joinpath(data_directory, "customXml.xlsx"); mode="r", enable_cache=true) do f
            @test_throws XLSX.XLSXError XLSX.getRowHeight(f, "Mock-up!B2") ≈ 23.25 # File not writable
            @test_throws XLSX.XLSXError XLSX.getColumnWidth(f, "Mock-up!B2") # File not writable
        end
        XLSX.openxlsx(joinpath(data_directory, "customXml.xlsx"); mode="r", enable_cache=false) do f
            @test_throws XLSX.XLSXError XLSX.getRowHeight(f, "Mock-up!B2")
            @test_throws XLSX.XLSXError XLSX.getFont(f, "Mock-up!B2")
            @test_throws XLSX.XLSXError XLSX.getFill(f, "Mock-up!B2")
            @test_throws XLSX.XLSXError XLSX.getBorder(f, "Mock-up!B2")
            @test_throws XLSX.XLSXError XLSX.getFormat(f, "Mock-up!B2")
            @test_throws XLSX.XLSXError XLSX.getAlignment(f, "Mock-up!B2")
            @test_throws XLSX.XLSXError XLSX.setRowHeight(f, "Mock-up!B2")
            @test_throws XLSX.XLSXError XLSX.setColumnWidth(f, "Mock-up!B2")
            @test_throws XLSX.XLSXError XLSX.setFont(f, "Mock-up!B2")
            @test_throws XLSX.XLSXError XLSX.setFill(f, "Mock-up!B2")
            @test_throws XLSX.XLSXError XLSX.setBorder(f, "Mock-up!B2")
            @test_throws XLSX.XLSXError XLSX.setFormat(f, "Mock-up!B2")
            @test_throws XLSX.XLSXError XLSX.setAlignment(f, "Mock-up!B2")
            @test_throws XLSX.XLSXError XLSX.setUniformFont(f, "Mock-up!B2:C4")
            @test_throws XLSX.XLSXError XLSX.setUniformFill(f, "Mock-up!B2:C4")
            @test_throws XLSX.XLSXError XLSX.setUniformBorder(f, "Mock-up!B2:C4")
            @test_throws XLSX.XLSXError XLSX.setUniformFormat(f, "Mock-up!B2:C4")
            @test_throws XLSX.XLSXError XLSX.setUniformAlignment(f, "Mock-up!B2:C4")
            @test_throws XLSX.XLSXError XLSX.setOutsideBorder(f, "Mock-up!B2:C4")
        end
    end

    @testset "indexing setAttribute" begin
        f = XLSX.newxlsx() # Empty XLSXFile
        s = f[1] #1×1 XLSX.Worksheet: ["Sheet1"](A1:A1)

        #Can't write to single, empty cells
        @test_throws XLSX.XLSXError XLSX.setFont(s, "A1"; color="grey42")
        @test_throws XLSX.XLSXError XLSX.setFont(s, "A1:A1"; color="grey42")
        @test_throws XLSX.XLSXError XLSX.setFont(s, "A:A"; color="grey42")
        @test_throws XLSX.XLSXError XLSX.setFont(s, "1"; color="grey42")
        @test_throws XLSX.XLSXError XLSX.setFont(s, 1, 1; color="grey42")
        @test_throws XLSX.XLSXError XLSX.setFont(s, [1], 1; color="grey42")
        @test_throws XLSX.XLSXError XLSX.setFont(s, 1, 1:1; color="grey42")
        @test_throws XLSX.XLSXError XLSX.setFont(s, 1, :; color="grey42")
        @test_throws XLSX.XLSXError XLSX.setFont(s, :; color="grey42")
        @test_throws XLSX.XLSXError XLSX.setFont(s, :, :; color="grey42")

        s[2, 1] = ""
        s[3, 3] = ""
        # Skip empty cells silently in ranges 
        @test XLSX.setFont(s, 2:3, 1:3; color="grey42") == -1

        # Outside sheet dimension
        @test_throws XLSX.XLSXError XLSX.getFont(s, 2, 4)
        @test_throws XLSX.XLSXError XLSX.getFont(s, 4, 2)
        @test_throws XLSX.XLSXError XLSX.getFont(s, 4, 4)

        s[1:3, 1:3] = ""
        default_font = XLSX.getDefaultFont(s).font
        dname = default_font["name"]["val"]
        dsize = default_font["sz"]["val"]
        XLSX.setFont(s, "A1"; color="grey42")
        @test XLSX.getFont(s, "A1").font == Dict("name" => Dict("val" => dname), "sz" => Dict("val" => dsize), "color" => Dict("rgb" => "FF6B6B6B"))
        XLSX.setFont(s, 2, 2; color="grey43", name="Ariel")
        @test XLSX.getFont(s, 2, 2).font == Dict("name" => Dict("val" => "Ariel"), "sz" => Dict("val" => dsize), "color" => Dict("rgb" => "FF6E6E6E"))
        XLSX.setFont(s, [2, 3], 1:3; color="grey44", name="Courier New")
        @test XLSX.getFont(s, 3, 1).font == Dict("name" => Dict("val" => "Courier New"), "sz" => Dict("val" => dsize), "color" => Dict("rgb" => "FF707070"))
        @test XLSX.getFont(s, 2, 2).font == Dict("name" => Dict("val" => "Courier New"), "sz" => Dict("val" => dsize), "color" => Dict("rgb" => "FF707070"))
        @test XLSX.getFont(s, 3, 3).font == Dict("name" => Dict("val" => "Courier New"), "sz" => Dict("val" => dsize), "color" => Dict("rgb" => "FF707070"))

        f = XLSX.newxlsx()
        s = f[1]
        @test_throws XLSX.XLSXError XLSX.setBorder(s, "A1"; allsides=["color" => "grey42", "style" => "thick"])
        @test_throws XLSX.XLSXError XLSX.setBorder(s, "A1:A1"; allsides=["color" => "grey42", "style" => "thick"])
        @test_throws XLSX.XLSXError XLSX.setBorder(s, "A"; allsides=["color" => "grey42", "style" => "thick"])
        @test_throws XLSX.XLSXError XLSX.setBorder(s, "1"; allsides=["color" => "grey42", "style" => "thick"])
        @test_throws XLSX.XLSXError XLSX.setBorder(s, 1, 1; allsides=["color" => "grey42", "style" => "thick"])
        @test_throws XLSX.XLSXError XLSX.setBorder(s, [1], 1; allsides=["color" => "grey42", "style" => "thick"])
        @test_throws XLSX.XLSXError XLSX.setBorder(s, 1, 1:1; allsides=["color" => "grey42", "style" => "thick"])
        @test_throws XLSX.XLSXError XLSX.setBorder(s, :, 1; allsides=["color" => "grey42", "style" => "thick"])
        @test_throws XLSX.XLSXError XLSX.setBorder(s, :; allsides=["color" => "grey42", "style" => "thick"])
        @test_throws XLSX.XLSXError XLSX.setBorder(s, :, :; allsides=["color" => "grey42", "style" => "thick"])
        @test_throws XLSX.XLSXError XLSX.setBorder(s, [2, 3], 1:3; allsides=["color" => "grey42", "style" => "thick"]) == -1
        @test_throws XLSX.XLSXError XLSX.getBorder(s, 2, 1)
        @test_throws XLSX.XLSXError XLSX.getBorder(s, 3, 2)
        @test_throws XLSX.XLSXError XLSX.getBorder(s, 2, 3)
        s[1:3, 1:3] = ""
        XLSX.setBorder(s, "A1"; allsides=["color" => "grey42", "style" => "thick"])
        @test XLSX.getBorder(s, "A1").border == Dict("left" => Dict("rgb" => "FF6B6B6B", "style" => "thick"), "bottom" => Dict("rgb" => "FF6B6B6B", "style" => "thick"), "right" => Dict("rgb" => "FF6B6B6B", "style" => "thick"), "top" => Dict("rgb" => "FF6B6B6B", "style" => "thick"), "diagonal" => nothing)
        XLSX.setBorder(s, 2, 2; allsides=["color" => "grey43", "style" => "thin"])
        @test XLSX.getBorder(s, 2, 2).border == Dict("left" => Dict("rgb" => "FF6E6E6E", "style" => "thin"), "bottom" => Dict("rgb" => "FF6E6E6E", "style" => "thin"), "right" => Dict("rgb" => "FF6E6E6E", "style" => "thin"), "top" => Dict("rgb" => "FF6E6E6E", "style" => "thin"), "diagonal" => nothing)
        XLSX.setBorder(s, [2, 3], 1:3; allsides=["color" => "grey44", "style" => "hair"], diagonal=["color" => "grey44", "style" => "thin", "direction" => "down"])
        @test XLSX.getBorder(s, 3, 1).border == Dict("left" => Dict("rgb" => "FF707070", "style" => "hair"), "bottom" => Dict("rgb" => "FF707070", "style" => "hair"), "right" => Dict("rgb" => "FF707070", "style" => "hair"), "top" => Dict("rgb" => "FF707070", "style" => "hair"), "diagonal" => Dict("rgb" => "FF707070", "style" => "thin", "direction" => "down"))
        @test XLSX.getBorder(s, 2, 2).border == Dict("left" => Dict("rgb" => "FF707070", "style" => "hair"), "bottom" => Dict("rgb" => "FF707070", "style" => "hair"), "right" => Dict("rgb" => "FF707070", "style" => "hair"), "top" => Dict("rgb" => "FF707070", "style" => "hair"), "diagonal" => Dict("rgb" => "FF707070", "style" => "thin", "direction" => "down"))
        @test XLSX.getBorder(s, 3, 3).border == Dict("left" => Dict("rgb" => "FF707070", "style" => "hair"), "bottom" => Dict("rgb" => "FF707070", "style" => "hair"), "right" => Dict("rgb" => "FF707070", "style" => "hair"), "top" => Dict("rgb" => "FF707070", "style" => "hair"), "diagonal" => Dict("rgb" => "FF707070", "style" => "thin", "direction" => "down"))

        f = XLSX.newxlsx()
        s = f[1]
        @test_throws XLSX.XLSXError XLSX.setFill(s, "A1"; pattern="lightVertical", fgColor="Red", bgColor="blue")
        @test_throws XLSX.XLSXError XLSX.setFill(s, "A1:A1"; pattern="lightVertical", fgColor="Red", bgColor="blue")
        @test_throws XLSX.XLSXError XLSX.setFill(s, "A"; pattern="lightVertical", fgColor="Red", bgColor="blue")
        @test_throws XLSX.XLSXError XLSX.setFill(s, "1"; pattern="lightVertical", fgColor="Red", bgColor="blue")
        @test_throws XLSX.XLSXError XLSX.setFill(s, 1, 1; pattern="lightVertical", fgColor="Red", bgColor="blue")
        @test_throws XLSX.XLSXError XLSX.setFill(s, [1], 1; pattern="lightVertical", fgColor="Red", bgColor="blue")
        @test_throws XLSX.XLSXError XLSX.setFill(s, 1, 1:1; pattern="lightVertical", fgColor="Red", bgColor="blue")
        @test_throws XLSX.XLSXError XLSX.setFill(s, 1, :; pattern="lightVertical", fgColor="Red", bgColor="blue")
        @test_throws XLSX.XLSXError XLSX.setFill(s, :, :; pattern="lightVertical", fgColor="Red", bgColor="blue")
        @test_throws XLSX.XLSXError XLSX.setFill(s, :; pattern="lightVertical", fgColor="Red", bgColor="blue")
        @test_throws XLSX.XLSXError XLSX.setFill(s, [2, 3], 1:3; pattern="lightVertical", fgColor="Red", bgColor="blue") == -1
        @test_throws XLSX.XLSXError XLSX.getFill(s, 2, 1)
        @test_throws XLSX.XLSXError XLSX.getFill(s, 3, 2)
        @test_throws XLSX.XLSXError XLSX.getFill(s, 2, 3)
        s[1:3, 1:3] = ""
        XLSX.setFill(s, "A1"; pattern="lightVertical", fgColor="Red", bgColor="blue")
        @test XLSX.getFill(s, "A1").fill == Dict("patternFill" => Dict("bgrgb" => "FF0000FF", "patternType" => "lightVertical", "fgrgb" => "FFFF0000"))
        XLSX.setFill(s, 2, 2; pattern="lightGrid", fgColor="Red", bgColor="blue")
        @test XLSX.getFill(s, 2, 2).fill == Dict("patternFill" => Dict("bgrgb" => "FF0000FF", "patternType" => "lightGrid", "fgrgb" => "FFFF0000"))
        XLSX.setFill(s, [2, 3], 1:3; pattern="lightGrid", fgColor="Red", bgColor="blue")
        @test XLSX.getFill(s, 3, 1).fill == Dict("patternFill" => Dict("bgrgb" => "FF0000FF", "patternType" => "lightGrid", "fgrgb" => "FFFF0000"))
        @test XLSX.getFill(s, 2, 2).fill == Dict("patternFill" => Dict("bgrgb" => "FF0000FF", "patternType" => "lightGrid", "fgrgb" => "FFFF0000"))
        @test XLSX.getFill(s, 3, 3).fill == Dict("patternFill" => Dict("bgrgb" => "FF0000FF", "patternType" => "lightGrid", "fgrgb" => "FFFF0000"))

        f = XLSX.newxlsx()
        s = f[1]
        @test_throws XLSX.XLSXError XLSX.setAlignment(s, "A1"; horizontal="right", vertical="justify", wrapText=true)
        @test_throws XLSX.XLSXError XLSX.setAlignment(s, "A1:A1"; horizontal="right", vertical="justify", wrapText=true)
        @test_throws XLSX.XLSXError XLSX.setAlignment(s, "A"; horizontal="right", vertical="justify", wrapText=true)
        @test_throws XLSX.XLSXError XLSX.setAlignment(s, "1"; horizontal="right", vertical="justify", wrapText=true)
        @test_throws XLSX.XLSXError XLSX.setAlignment(s, 1, 1; horizontal="right", vertical="justify", wrapText=true)
        @test_throws XLSX.XLSXError XLSX.setAlignment(s, [1], 1; horizontal="right", vertical="justify", wrapText=true)
        @test_throws XLSX.XLSXError XLSX.setAlignment(s, 1, 1:1; horizontal="right", vertical="justify", wrapText=true)
        @test_throws XLSX.XLSXError XLSX.setAlignment(s, 1, :; horizontal="right", vertical="justify", wrapText=true)
        @test_throws XLSX.XLSXError XLSX.setAlignment(s, :, :; horizontal="right", vertical="justify", wrapText=true)
        @test_throws XLSX.XLSXError XLSX.setAlignment(s, :; horizontal="right", vertical="justify", wrapText=true)
        @test_throws XLSX.XLSXError XLSX.setAlignment(s, [2, 3], 1:3; horizontal="right", vertical="justify", wrapText=true) == -1
        @test_throws XLSX.XLSXError XLSX.getAlignment(s, 2, 1)
        @test_throws XLSX.XLSXError XLSX.getAlignment(s, 3, 2)
        @test_throws XLSX.XLSXError XLSX.getAlignment(s, 2, 3)
        s[1:3, 1:3] = ""
        XLSX.setAlignment(s, "A1"; horizontal="right", vertical="justify", wrapText=true)
        @test XLSX.getAlignment(s, "A1").alignment == Dict("alignment" => Dict("horizontal" => "right", "vertical" => "justify", "wrapText" => "1"))
        XLSX.setAlignment(s, 2, 2; horizontal="right", vertical="justify", wrapText=true, rotation=90)
        @test XLSX.getAlignment(s, 2, 2).alignment == Dict("alignment" => Dict("horizontal" => "right", "vertical" => "justify", "wrapText" => "1", "textRotation" => "90"))
        XLSX.setAlignment(s, [2, 3], 1:3; horizontal="right", vertical="justify", shrink=true, rotation=90)
        @test XLSX.getAlignment(s, 3, 1).alignment == Dict("alignment" => Dict("horizontal" => "right", "vertical" => "justify", "shrinkToFit" => "1", "textRotation" => "90"))
        @test XLSX.getAlignment(s, 2, 2).alignment == Dict("alignment" => Dict("horizontal" => "right", "vertical" => "justify", "wrapText" => "1", "shrinkToFit" => "1", "textRotation" => "90"))
        @test XLSX.getAlignment(s, 3, 3).alignment == Dict("alignment" => Dict("horizontal" => "right", "vertical" => "justify", "shrinkToFit" => "1", "textRotation" => "90"))

        f = XLSX.newxlsx()
        s = f[1]
        @test_throws XLSX.XLSXError XLSX.setFormat(s, "A1"; format="Percentage")
        @test_throws XLSX.XLSXError XLSX.setFormat(s, "A1:A1"; format="Percentage")
        @test_throws XLSX.XLSXError XLSX.setFormat(s, "A"; format="Percentage")
        @test_throws XLSX.XLSXError XLSX.setFormat(s, "1"; format="Percentage")
        @test_throws XLSX.XLSXError XLSX.setFormat(s, 1, 1; format="Percentage")
        @test_throws XLSX.XLSXError XLSX.setFormat(s, [1], 1; format="Percentage")
        @test_throws XLSX.XLSXError XLSX.setFormat(s, 1, 1:1; format="Percentage")
        @test_throws XLSX.XLSXError XLSX.setFormat(s, 1, :; format="Percentage")
        @test_throws XLSX.XLSXError XLSX.setFormat(s, :, :; format="Percentage")
        @test_throws XLSX.XLSXError XLSX.setFormat(s, :; format="Percentage")
        @test_throws XLSX.XLSXError XLSX.setFormat(s, [2, 3], 1:3; format="Percentage") == -1
        @test_throws XLSX.XLSXError XLSX.getFormat(s, 2, 1)
        @test_throws XLSX.XLSXError XLSX.getFormat(s, 3, 2)
        @test_throws XLSX.XLSXError XLSX.getFormat(s, 2, 3)
        s[1:3, 1:3] = ""
        XLSX.setFormat(s, "A1"; format="#,##0.000;(#,##0.000)")
        @test XLSX.getFormat(s, "A1").format == Dict("numFmt" => Dict("numFmtId" => "164", "formatCode" => "#,##0.000;(#,##0.000)"))
        XLSX.setFormat(s, 2, 2; format="Currency")
        @test XLSX.getFormat(s, 2, 2).format == Dict("numFmt" => Dict("numFmtId" => "7", "formatCode" => "\$#,##0.00_);(\$#,##0.00)"))
        XLSX.setFormat(s, [2, 3], 1:3; format="LongDate")
        @test XLSX.getFormat(s, 3, 1).format == Dict("numFmt" => Dict("numFmtId" => "15", "formatCode" => "d-mmm-yy"))
        @test XLSX.getFormat(s, 2, 2).format == Dict("numFmt" => Dict("numFmtId" => "15", "formatCode" => "d-mmm-yy"))
        @test XLSX.getFormat(s, 3, 3).format == Dict("numFmt" => Dict("numFmtId" => "15", "formatCode" => "d-mmm-yy"))

        f = XLSX.newxlsx()
        s = f[1]
        @test_throws XLSX.XLSXError XLSX.getColumnWidth(s, "B2") # Cell outside sheet dimension
        s[1:3, 1:3] = ""
        XLSX.setColumnWidth(s, "A1"; width=30)
        @test XLSX.getColumnWidth(s, "A1") ≈ 30.7109375
        XLSX.setColumnWidth(s, 2, 2; width=40)
        @test XLSX.getColumnWidth(s, 2, 2) ≈ 40.7109375
        XLSX.setColumnWidth(s, [2, 3], 1:3; width=50)
        @test XLSX.getColumnWidth(s, 3, 1) ≈ 50.7109375
        @test XLSX.getColumnWidth(s, 2, 2) ≈ 50.7109375
        @test XLSX.getColumnWidth(s, 3, 3) ≈ 50.7109375

        f = XLSX.newxlsx()
        s = f[1]
        @test_throws XLSX.XLSXError XLSX.getRowHeight(s, "B2") # Cell outside sheet dimension
        s[1:3, 1:3] = ""
        XLSX.setRowHeight(s, "A1"; height=30)
        @test XLSX.getRowHeight(s, "A1") ≈ 30.2109375
        XLSX.setRowHeight(s, 2, 2; height=40)
        @test XLSX.getRowHeight(s, 2, 2) ≈ 40.2109375
        XLSX.setRowHeight(s, [2, 3], 1:3; height=50)
        @test XLSX.getRowHeight(s, 3, 1) ≈ 50.2109375
        @test XLSX.getRowHeight(s, 2, 2) ≈ 50.2109375
        @test XLSX.getRowHeight(s, 3, 3) ≈ 50.2109375

        f = XLSX.newxlsx()
        s = f[1]
        s[1:30, 1:26] = ""
        XLSX.setUniformFont(s, 1:4, :; size=12, name="Times New Roman", color="FF040404")
        @test XLSX.getFont(f, "Sheet1!A1").font == Dict("sz" => Dict("val" => "12"), "name" => Dict("val" => "Times New Roman"), "color" => Dict("rgb" => "FF040404"))
        @test XLSX.getFont(f, "Sheet1!G2").font == Dict("sz" => Dict("val" => "12"), "name" => Dict("val" => "Times New Roman"), "color" => Dict("rgb" => "FF040404"))
        @test XLSX.getFont(f, "Sheet1!N3").font == Dict("sz" => Dict("val" => "12"), "name" => Dict("val" => "Times New Roman"), "color" => Dict("rgb" => "FF040404"))
        @test XLSX.getFont(f, "Sheet1!Y4").font == Dict("sz" => Dict("val" => "12"), "name" => Dict("val" => "Times New Roman"), "color" => Dict("rgb" => "FF040404"))

        XLSX.setUniformFill(s, :, 2:8; pattern="lightGrid", fgColor="FF0000FF", bgColor="FF00FF00")
        @test XLSX.getFill(s, "B10").fill == Dict("patternFill" => Dict("bgrgb" => "FF00FF00", "patternType" => "lightGrid", "fgrgb" => "FF0000FF"))
        @test XLSX.getFill(s, "D20").fill == Dict("patternFill" => Dict("bgrgb" => "FF00FF00", "patternType" => "lightGrid", "fgrgb" => "FF0000FF"))
        @test XLSX.getFill(s, "F30").fill == Dict("patternFill" => Dict("bgrgb" => "FF00FF00", "patternType" => "lightGrid", "fgrgb" => "FF0000FF"))

        XLSX.setUniformFormat(s, :; format="#,##0.000")
        @test XLSX.getFormat(s, "A1").format == Dict("numFmt" => Dict("numFmtId" => "164", "formatCode" => "#,##0.000"))
        @test XLSX.getFormat(s, "G10").format == Dict("numFmt" => Dict("numFmtId" => "164", "formatCode" => "#,##0.000"))
        @test XLSX.getFormat(s, "M20").format == Dict("numFmt" => Dict("numFmtId" => "164", "formatCode" => "#,##0.000"))
        @test XLSX.getFormat(s, "X30").format == Dict("numFmt" => Dict("numFmtId" => "164", "formatCode" => "#,##0.000"))

        f = XLSX.open_xlsx_template(joinpath(data_directory, "Borders.xlsx"))
        s = f["Sheet1"]
        XLSX.setUniformBorder(s, [1, 2, 3, 4], 1:4; left=["style" => "dotted", "color" => "darkseagreen3"],
            right=["style" => "medium", "color" => "FF765000"],
            top=["style" => "thick", "color" => "FF230000"],
            bottom=["style" => "medium", "color" => "FF0000FF"],
            diagonal=["style" => "none"]
        )
        @test XLSX.getBorder(s, 1, 1).border == Dict("left" => Dict("style" => "dotted", "rgb" => "FF9BCD9B"), "bottom" => Dict("style" => "medium", "rgb" => "FF0000FF"), "right" => Dict("style" => "medium", "rgb" => "FF765000"), "top" => Dict("style" => "thick", "rgb" => "FF230000"), "diagonal" => nothing)
        @test XLSX.getBorder(s, 2, 2).border == Dict("left" => Dict("style" => "dotted", "rgb" => "FF9BCD9B"), "bottom" => Dict("style" => "medium", "rgb" => "FF0000FF"), "right" => Dict("style" => "medium", "rgb" => "FF765000"), "top" => Dict("style" => "thick", "rgb" => "FF230000"), "diagonal" => nothing)
        @test XLSX.getBorder(s, 4, 4).border == Dict("left" => Dict("style" => "dotted", "rgb" => "FF9BCD9B"), "bottom" => Dict("style" => "medium", "rgb" => "FF0000FF"), "right" => Dict("style" => "medium", "rgb" => "FF765000"), "top" => Dict("style" => "thick", "rgb" => "FF230000"), "diagonal" => nothing)

    end

    @testset "existing formatting" begin
        f = XLSX.opentemplate(joinpath(data_directory, "customXml.xlsx"))
        s = f[1]
        s["B2"] = pi
        s["D20"] = "Hello World"
        s["J45"] = Dates.Date(2025, 01, 24)
        @test XLSX.getFont(s, "B2").font == Dict("name" => Dict("val" => "Calibri"), "family" => Dict("val" => "2"), "b" => nothing, "sz" => Dict("val" => "18"), "color" => Dict("theme" => "1"), "scheme" => Dict("val" => "minor"))
        @test XLSX.getFill(s, "D20").fill == Dict("patternFill" => Dict("bgindexed" => "64", "patternType" => "solid", "fgtint" => "-0.499984740745262", "fgtheme" => "2"))
        @test XLSX.getBorder(s, "J45").border == Dict("left" => Dict("indexed" => "64", "style" => "thin"), "bottom" => Dict("indexed" => "64", "style" => "thin"), "right" => Dict("indexed" => "64", "style" => "thin"), "top" => Dict("indexed" => "64", "style" => "thin"), "diagonal" => nothing)
    end
end

@testset "Conditional Formats" begin

    @testset "DataBar" begin

        f = XLSX.newxlsx()
        s = f[1]
        for i in 1:5, j in 1:5
            s[i, j] = i + j
        end

        @test_throws MethodError XLSX.setConditionalFormat(s, "A1,A3", :dataBar) # Non-contiguous ranges not allowed
        @test_throws MethodError XLSX.setConditionalFormat(s, [1], 1, :dataBar) # Vectors may be non-contiguous
        @test_throws MethodError XLSX.setConditionalFormat(s, 1, 1:3:7, :dataBar) # StepRange is non-contiguous
        @test XLSX.setConditionalFormat(s, "1:1", :dataBar) == 0
        @test XLSX.setConditionalFormat(s, 2, :, :dataBar; databar="greengrad") == 0
        @test XLSX.setConditionalFormat(s, 3, 1:5, :dataBar;
            min_type="least",
            min_val="green", #should be ignored because type=least
            max_type="percentile",
            max_val="50",
        ) == 0
        @test XLSX.setConditionalFormat(s, "Sheet1!A4:E4", :dataBar;
            min_type="automatic",
            max_type="automatic",
        ) == 0
        @test XLSX.setConditionalFormat(f, "Sheet1!A5:E5", :dataBar;
            min_type="num",
            min_val="\$A\$1",
            max_type="formula",
            max_val="\$A\$2"
        ) == 0
        @test sort!(XLSX.getConditionalFormats(s), by = x -> x.second.priority, rev=true) == [
            XLSX.CellRange("A5:E5") => (type="dataBar", priority=5), 
            XLSX.CellRange("A4:E4") => (type="dataBar", priority=4), 
            XLSX.CellRange("A3:E3") => (type="dataBar", priority=3), 
            XLSX.CellRange("A2:E2") => (type="dataBar", priority=2), 
            XLSX.CellRange("A1:E1") => (type="dataBar", priority=1),
        ]
        @test XLSX.setConditionalFormat(s, "A1", :dataBar) == 0
        @test XLSX.setConditionalFormat(s, "A1:C3", :dataBar) == 0
        @test XLSX.setConditionalFormat(s, "Sheet1!A1", :dataBar) == 0
        @test XLSX.setConditionalFormat(s, "Sheet1!A1:A2", :dataBar) == 0
        @test XLSX.setConditionalFormat(s, "Sheet1!1:2", :dataBar) == 0
        @test XLSX.setConditionalFormat(s, "2:4", :dataBar) == 0
        @test XLSX.setConditionalFormat(s, "A:C", :dataBar) == 0
        @test XLSX.setConditionalFormat(s, "Sheet1!A:C", :dataBar) == 0
        @test XLSX.setConditionalFormat(f, "Sheet1!1:2", :dataBar) == 0
        @test XLSX.setConditionalFormat(f, "Sheet1!A:C", :dataBar) == 0
        @test XLSX.setConditionalFormat(s, :, 1:3, :dataBar) == 0
        @test XLSX.setConditionalFormat(s, 1:3, :, :dataBar) == 0
        @test XLSX.setConditionalFormat(s, "2:4", :dataBar) == 0
        @test XLSX.setConditionalFormat(s, "A:C", :dataBar) == 0
        @test XLSX.setConditionalFormat(s, "Sheet1!A:C", :dataBar) == 0
        @test XLSX.setConditionalFormat(s, :, :dataBar) == 0
        @test XLSX.setConditionalFormat(s, :, :, :dataBar) == 0
        @test length(XLSX.getConditionalFormats(s)) == 22
        @test sort!(XLSX.getConditionalFormats(s), by = x -> x.second.priority, rev=true) == [
            XLSX.CellRange("A1:E5") => (type="dataBar", priority=22),
            XLSX.CellRange("A1:E5") => (type="dataBar", priority=21),
            XLSX.CellRange("A1:C5") => (type="dataBar", priority=20),
            XLSX.CellRange("A1:C5") => (type="dataBar", priority=19),
            XLSX.CellRange("A2:E4") => (type="dataBar", priority=18),
            XLSX.CellRange("A1:E3") => (type="dataBar", priority=17),
            XLSX.CellRange("A1:C5") => (type="dataBar", priority=16),
            XLSX.CellRange("A1:C5") => (type="dataBar", priority=15),
            XLSX.CellRange("A1:E2") => (type="dataBar", priority=14),
            XLSX.CellRange("A1:C5") => (type="dataBar", priority=13),
            XLSX.CellRange("A1:C5") => (type="dataBar", priority=12),
            XLSX.CellRange("A2:E4") => (type="dataBar", priority=11),
            XLSX.CellRange("A1:E2") => (type="dataBar", priority=10),
            XLSX.CellRange("A1:A2") => (type="dataBar", priority=9),
            XLSX.CellRange("A1:A1") => (type="dataBar", priority=8),
            XLSX.CellRange("A1:C3") => (type="dataBar", priority=7),
            XLSX.CellRange("A1:A1") => (type="dataBar", priority=6),
            XLSX.CellRange("A5:E5") => (type="dataBar", priority=5),
            XLSX.CellRange("A4:E4") => (type="dataBar", priority=4),
            XLSX.CellRange("A3:E3") => (type="dataBar", priority=3),
            XLSX.CellRange("A2:E2") => (type="dataBar", priority=2),
            XLSX.CellRange("A1:E1") => (type="dataBar", priority=1)
        ]

        f = XLSX.newxlsx()
        s = f[1]
        for i in 1:5, j in 1:5
            s[i, j] = i + j
        end
        @test XLSX.setConditionalFormat(s, "A1:A5", :dataBar) == 0
        @test XLSX.setConditionalFormat(s, :, 2, :dataBar; databar="orange") == 0
        @test XLSX.setConditionalFormat(s, "Sheet1!E:E", :dataBar; databar="purplegrad") == 0
        @test XLSX.setConditionalFormat(s, 1:5, 3:4, :dataBar;
            borders="false",
            min_type="percentile",
            min_val="25",
            max_type="percentile",
            max_val="75"
        ) == 0
        @test sort!(XLSX.getConditionalFormats(s), by = x -> x.second.priority, rev=true) == [
            XLSX.CellRange("C1:D5") => (type="dataBar", priority=4), 
            XLSX.CellRange("E1:E5") => (type="dataBar", priority=3), 
            XLSX.CellRange("B1:B5") => (type="dataBar", priority=2), 
            XLSX.CellRange("A1:A5") => (type="dataBar", priority=1)
        ]

        f = XLSX.newxlsx()
        s = f[1]
        for i in 1:5, j in 1:5
            s[i, j] = i + j
        end

        @test XLSX.setConditionalFormat(s, :, 1:4, :dataBar;
            databar="red",
            borders="true",
            fill_col="blue",
            border_col="yellow",
            neg_fill_col="magenta",
            neg_border_col="green",
            axis_col="cyan"
        ) == 0

        f = XLSX.newxlsx()
        s = f[1]
        for i in 1:5, j in 1:5
            s[i, j] = i + j
        end
        XLSX.addDefinedName(s, "myRange", "A1:B5")
        @test XLSX.setConditionalFormat(s, "myRange", :dataBar;
            showVal="false",
            direction="leftToRight",
            borders="true",
            sameNegBorders="false"
        ) == 0
        XLSX.addDefinedName(s, "myNCRange", "C1:C5,D1:D5")
        @test_throws MethodError XLSX.setConditionalFormat(s, "myNCRange", :dataBar; # Non-contiguous ranges not allowed
            showVal="false",
            direction="leftToRight",
            borders="true",
            sameNegBorders="false"
        )
        @test_throws XLSX.XLSXError XLSX.setConditionalFormat(s, "A1:A2", :dataBar;
            databar="rainbow"
        )

        f = XLSX.newxlsx()
        s = f[1]
        for i in 1:5, j in 1:12
            s[i, j] = i + j
        end
        s[1, 13] = 5

        @test XLSX.setConditionalFormat(s, :, 1, :dataBar;
            databar="orange",
            sameNegFill="true",
            sameNegBorders="true"
        ) == 0
        @test XLSX.setConditionalFormat(s, :, 2, :dataBar;
            databar="orange",
            axis_pos="none"
        ) == 0
        @test XLSX.setConditionalFormat(s, :, 3, :dataBar;
            databar="orange",
            axis_pos="middle"
        ) == 0
        @test XLSX.setConditionalFormat(s, :, 4, :dataBar;
            databar="orange",
            min_type="num",
            min_val="Sheet1!\$M\$1"
        ) == 0
        @test XLSX.setConditionalFormat(s, :, 5, :dataBar;
            databar="orange",
            showVal="false",
            direction="rightToLeft",
            borders="true",
            sameNegBorders="false",
            sameNegFill="false"
        ) == 0

        @test_throws XLSX.XLSXError XLSX.setConditionalFormat(s, :, 4, :dataBar;
            axis_pos="nonsense",
            databar="orange",
            min_type="num",
            min_val="Sheet1!\$M\$1"
        )
        @test_throws XLSX.XLSXError XLSX.setConditionalFormat(s, :, 4, :dataBar;
            borders="nonsense",
            databar="orange",
            min_type="num",
            min_val="Sheet1!\$M\$1"
        )
        @test_throws XLSX.XLSXError XLSX.setConditionalFormat(s, :, 4, :dataBar;
            fill_col="nonsense",
            databar="orange",
            min_type="num",
            min_val="Sheet1!\$M\$1"
        )
        @test_throws XLSX.XLSXError XLSX.setConditionalFormat(s, :, 4, :dataBar;
            sameNegFill="nonsense",
            databar="orange",
            min_type="num",
            min_val="Sheet1!\$M\$1"
        )
        @test_throws XLSX.XLSXError XLSX.setConditionalFormat(s, :, 4, :dataBar;
            databar="orange",
            min_type="num",
            min_val="Sheet2!\$M\$1"
        )

        f = XLSX.newxlsx()
        s = f[1]
        for i in 1:5, j in 1:12
            s[i, j] = i + j
        end
        for (j, k) in enumerate(keys(XLSX.databars))
            @test XLSX.setConditionalFormat(s, :, j, :dataBar; databar=k) == 0
        end
    end

    @testset "colorScale" begin

        f = XLSX.newxlsx()
        s = f[1]
        for i in 1:5, j in 1:5
            s[i, j] = i + j
        end

        @test_throws XLSX.XLSXError XLSX.setConditionalFormat(s, "A1,A3", :wrongOne)
        @test_throws XLSX.XLSXError XLSX.setConditionalFormat(s, 1, 2, :wrongOne)

        @test_throws MethodError XLSX.setConditionalFormat(s, "A1,A3", :colorScale) # Non-contiguous ranges not allowed
        @test_throws MethodError XLSX.setConditionalFormat(s, [1], 1, :colorScale) # Vectors may be non-contiguous
        @test_throws MethodError XLSX.setConditionalFormat(s, 1, 1:3:7, :colorScale) # StepRange is non-contiguous
        @test XLSX.setConditionalFormat(s, "1:1", :colorScale) == 0
        @test XLSX.setConditionalFormat(s, 2, :, :colorScale; colorscale="redwhiteblue") == 0
        @test XLSX.setConditionalFormat(s, 3, 1:5, :colorScale;
            min_type="min",
            min_col="green",
            mid_type="percentile",
            mid_val="50",
            mid_col="red",
            max_type="max",
            max_col="blue"
        ) == 0
        @test XLSX.setConditionalFormat(s, "Sheet1!A4:E4", :colorScale;
            min_type="min",
            min_col="tomato",
            max_type="max",
            max_col="gold4"
        ) == 0
        @test XLSX.setConditionalFormat(f, "Sheet1!A5:E5", :colorScale;
            min_type="min",
            min_col="yellow",
            max_type="max",
            max_col="darkgreen"
        ) == 0
        @test sort!(XLSX.getConditionalFormats(s), by = x -> x.second.priority, rev=true) == [
            XLSX.CellRange("A5:E5") => (type="colorScale", priority=5), 
            XLSX.CellRange("A4:E4") => (type="colorScale", priority=4), 
            XLSX.CellRange("A3:E3") => (type="colorScale", priority=3), 
            XLSX.CellRange("A2:E2") => (type="colorScale", priority=2), 
            XLSX.CellRange("A1:E1") => (type="colorScale", priority=1)
        ]
        @test XLSX.setConditionalFormat(s, "A1", :colorScale) == 0
        @test XLSX.setConditionalFormat(s, "A1:C3", :colorScale) == 0
        @test XLSX.setConditionalFormat(s, "Sheet1!A1", :colorScale) == 0
        @test XLSX.setConditionalFormat(s, "Sheet1!A1:A2", :colorScale) == 0
        @test XLSX.setConditionalFormat(s, "Sheet1!1:2", :colorScale) == 0
        @test XLSX.setConditionalFormat(s, "2:4", :colorScale) == 0
        @test XLSX.setConditionalFormat(s, "A:C", :colorScale) == 0
        @test XLSX.setConditionalFormat(s, "Sheet1!A:C", :colorScale) == 0
        @test XLSX.setConditionalFormat(f, "Sheet1!1:2", :colorScale) == 0
        @test XLSX.setConditionalFormat(f, "Sheet1!A:C", :colorScale) == 0
        @test XLSX.setConditionalFormat(s, :, 1:3, :colorScale) == 0
        @test XLSX.setConditionalFormat(s, 1:3, :, :colorScale) == 0
        @test XLSX.setConditionalFormat(s, "2:4", :colorScale) == 0
        @test XLSX.setConditionalFormat(s, "A:C", :colorScale) == 0
        @test XLSX.setConditionalFormat(s, "Sheet1!A:C", :colorScale) == 0
        @test XLSX.setConditionalFormat(s, :, :colorScale) == 0
        @test XLSX.setConditionalFormat(s, :, :, :colorScale) == 0
        @test length(XLSX.getConditionalFormats(s)) == 22
        @test sort!(XLSX.getConditionalFormats(s), by = x -> x.second.priority, rev=true) == [
            XLSX.CellRange("A1:E5") => (type="colorScale", priority=22),
            XLSX.CellRange("A1:E5") => (type="colorScale", priority=21),
            XLSX.CellRange("A1:C5") => (type="colorScale", priority=20),
            XLSX.CellRange("A1:C5") => (type="colorScale", priority=19),
            XLSX.CellRange("A2:E4") => (type="colorScale", priority=18),
            XLSX.CellRange("A1:E3") => (type="colorScale", priority=17),
            XLSX.CellRange("A1:C5") => (type="colorScale", priority=16),
            XLSX.CellRange("A1:C5") => (type="colorScale", priority=15),
            XLSX.CellRange("A1:E2") => (type="colorScale", priority=14),
            XLSX.CellRange("A1:C5") => (type="colorScale", priority=13),
            XLSX.CellRange("A1:C5") => (type="colorScale", priority=12),
            XLSX.CellRange("A2:E4") => (type="colorScale", priority=11),
            XLSX.CellRange("A1:E2") => (type="colorScale", priority=10),
            XLSX.CellRange("A1:A2") => (type="colorScale", priority=9),
            XLSX.CellRange("A1:A1") => (type="colorScale", priority=8),
            XLSX.CellRange("A1:C3") => (type="colorScale", priority=7),
            XLSX.CellRange("A1:A1") => (type="colorScale", priority=6),
            XLSX.CellRange("A5:E5") => (type="colorScale", priority=5),
            XLSX.CellRange("A4:E4") => (type="colorScale", priority=4),
            XLSX.CellRange("A3:E3") => (type="colorScale", priority=3),
            XLSX.CellRange("A2:E2") => (type="colorScale", priority=2),
            XLSX.CellRange("A1:E1") => (type="colorScale", priority=1)
        ]

        f = XLSX.newxlsx()
        s = f[1]
        for i in 1:5, j in 1:5
            s[i, j] = i + j
        end
        @test XLSX.setConditionalFormat(s, "A1:A5", :colorScale) == 0
        @test XLSX.setConditionalFormat(s, :, 2, :colorScale; colorscale="redwhiteblue") == 0
        @test XLSX.setConditionalFormat(s, "Sheet1!E:E", :colorScale; colorscale="greenwhitered") == 0
        @test XLSX.setConditionalFormat(s, 1:5, 3:4, :colorScale;
            min_type="min",
            min_col="green",
            mid_type="percentile",
            mid_val="50",
            mid_col="red",
            max_type="max",
            max_col="blue"
        ) == 0
        @test sort!(XLSX.getConditionalFormats(s), by = x -> x.second.priority, rev=true) == [
            XLSX.CellRange("C1:D5") => (type="colorScale", priority=4), 
            XLSX.CellRange("E1:E5") => (type="colorScale", priority=3), 
            XLSX.CellRange("B1:B5") => (type="colorScale", priority=2), 
            XLSX.CellRange("A1:A5") => (type="colorScale", priority=1)
        ]

        f = XLSX.newxlsx()
        s = f[1]
        for i in 1:5, j in 1:5
            s[i, j] = i + j
        end

        @test XLSX.setConditionalFormat(s, :, 1:4, :colorScale;
            min_type="min",
            min_col="green",
            mid_type="percentile",
            mid_val="\$E\$4",
            mid_col="red",
            max_type="max",
            max_col="blue"
        ) == 0
        @test XLSX.setConditionalFormat(s, :, 5, :colorScale;
            min_type="min",
            min_col="green",
            mid_type="percentile",
            mid_val="Sheet1!\$E\$4",
            mid_col="red",
            max_type="max",
            max_col="blue"
        ) == 0
        @test_throws XLSX.XLSXError XLSX.setConditionalFormat(s, :, 5, :colorScale;
            min_type="min",
            min_col="green",
            mid_type="percentile",
            mid_val="Sheet2!\$E\$4",
            mid_col="red",
            max_type="max",
            max_col="blue"
        )

        f = XLSX.newxlsx()
        s = f[1]
        for i in 1:5, j in 1:5
            s[i, j] = i + j
        end
        XLSX.addDefinedName(s, "myRange", "A1:B5")
        @test XLSX.setConditionalFormat(s, "myRange", :colorScale;
            min_type="min",
            min_col="green",
            mid_type="percentile",
            mid_val="50",
            mid_col="red",
            max_type="max",
            max_col="blue"
        ) == 0
        XLSX.addDefinedName(s, "myNCRange", "C1:C5,D1:D5")
        @test_throws MethodError XLSX.setConditionalFormat(s, "myNCRange", :colorScale; # Non-contiguous ranges not allowed
            min_type="min",
            min_col="green",
            mid_type="percentile",
            mid_val="50",
            mid_col="red",
            max_type="max",
            max_col="blue"
        )
        @test_throws XLSX.XLSXError XLSX.setConditionalFormat(s, "A1:A2", :colorScale;
            colorscale="rainbow"
        )

        f = XLSX.newxlsx()
        s = f[1]
        for i in 1:5, j in 1:12
            s[i, j] = i + j
        end
        for (j, k) in enumerate(keys(XLSX.colorscales))
            @test XLSX.setConditionalFormat(s, :, j, :colorScale; colorscale=k) == 0
        end
    end

    @testset "iconSet" begin
        f = XLSX.newxlsx()
        s = f[1]
        for i in 1:5, j in 1:5
            s[i, j] = i + j
        end
        @test_throws MethodError XLSX.setConditionalFormat(s, "A1,A3", :iconSet) # Non-contiguous ranges not allowed
        @test_throws MethodError XLSX.setConditionalFormat(s, [1], 1, :iconSet) # Vectors may be non-contiguous
        @test_throws MethodError XLSX.setConditionalFormat(s, 1, 1:3:7, :iconSet) # StepRange is non-contiguous
        @test XLSX.setConditionalFormat(s, "1:1", :iconSet) == 0
        @test XLSX.setConditionalFormat(s, 2, :, :iconSet; iconset="3Arrows") == 0
        @test XLSX.setConditionalFormat(s, 3, 1:5, :iconSet;
            min_type="percent",
            min_val="20",
            max_type="num",
            max_val="4"
        ) == 0
        @test XLSX.setConditionalFormat(s, "Sheet1!A4:E4", :iconSet;
            min_type="percentile",
            min_val="10",
            max_type="num",
            max_val="\$C\$4"
        ) == 0
        @test XLSX.setConditionalFormat(f, "Sheet1!A5:E5", :iconSet;
            min_type="percentile",
            min_val="\$D\$5",
            max_type="percent",
            max_val="95"
        ) == 0
        @test sort!(XLSX.getConditionalFormats(s), by = x -> x.second.priority, rev=true) == [
            XLSX.CellRange("A5:E5") => (type="iconSet", priority=5), 
            XLSX.CellRange("A4:E4") => (type="iconSet", priority=4), 
            XLSX.CellRange("A3:E3") => (type="iconSet", priority=3), 
            XLSX.CellRange("A2:E2") => (type="iconSet", priority=2), 
            XLSX.CellRange("A1:E1") => (type="iconSet", priority=1)
        ]
        @test XLSX.setConditionalFormat(s, "A1", :iconSet) == 0
        @test XLSX.setConditionalFormat(s, "A1:C3", :iconSet) == 0
        @test XLSX.setConditionalFormat(s, "Sheet1!A1", :iconSet) == 0
        @test XLSX.setConditionalFormat(s, "Sheet1!A1:A2", :iconSet) == 0
        @test XLSX.setConditionalFormat(s, "Sheet1!1:2", :iconSet) == 0
        @test XLSX.setConditionalFormat(s, "2:4", :iconSet) == 0
        @test XLSX.setConditionalFormat(s, "A:C", :iconSet) == 0
        @test XLSX.setConditionalFormat(s, "Sheet1!A:C", :iconSet) == 0
        @test XLSX.setConditionalFormat(f, "Sheet1!1:2", :iconSet) == 0
        @test XLSX.setConditionalFormat(f, "Sheet1!A:C", :iconSet) == 0
        @test XLSX.setConditionalFormat(s, :, 1:3, :iconSet) == 0
        @test XLSX.setConditionalFormat(s, 1:3, :, :iconSet) == 0
        @test XLSX.setConditionalFormat(s, "2:4", :iconSet) == 0
        @test XLSX.setConditionalFormat(s, "A:C", :iconSet) == 0
        @test XLSX.setConditionalFormat(s, "Sheet1!A:C", :iconSet) == 0
        @test XLSX.setConditionalFormat(s, :, :iconSet) == 0
        @test XLSX.setConditionalFormat(s, :, :, :iconSet) == 0
        @test length(XLSX.getConditionalFormats(s)) == 22

        @test sort!(XLSX.getConditionalFormats(s), by = x -> x.second.priority, rev=true) == [
            XLSX.CellRange("A1:E5") => (type="iconSet", priority=22),
            XLSX.CellRange("A1:E5") => (type="iconSet", priority=21),
            XLSX.CellRange("A1:C5") => (type="iconSet", priority=20),
            XLSX.CellRange("A1:C5") => (type="iconSet", priority=19),
            XLSX.CellRange("A2:E4") => (type="iconSet", priority=18),
            XLSX.CellRange("A1:E3") => (type="iconSet", priority=17),
            XLSX.CellRange("A1:C5") => (type="iconSet", priority=16),
            XLSX.CellRange("A1:C5") => (type="iconSet", priority=15),
            XLSX.CellRange("A1:E2") => (type="iconSet", priority=14),
            XLSX.CellRange("A1:C5") => (type="iconSet", priority=13),
            XLSX.CellRange("A1:C5") => (type="iconSet", priority=12),
            XLSX.CellRange("A2:E4") => (type="iconSet", priority=11),
            XLSX.CellRange("A1:E2") => (type="iconSet", priority=10),
            XLSX.CellRange("A1:A2") => (type="iconSet", priority=9),
            XLSX.CellRange("A1:A1") => (type="iconSet", priority=8),
            XLSX.CellRange("A1:C3") => (type="iconSet", priority=7),
            XLSX.CellRange("A1:A1") => (type="iconSet", priority=6),
            XLSX.CellRange("A5:E5") => (type="iconSet", priority=5),
            XLSX.CellRange("A4:E4") => (type="iconSet", priority=4),
            XLSX.CellRange("A3:E3") => (type="iconSet", priority=3),
            XLSX.CellRange("A2:E2") => (type="iconSet", priority=2),
            XLSX.CellRange("A1:E1") => (type="iconSet", priority=1)
        ]

        f = XLSX.newxlsx()
        s = f[1]

        XLSX.writetable!(s, [collect(1:10), collect(1:10), collect(1:10), collect(1:10), collect(1:10), collect(1:10)],
            ["normal", "showVal=\"false\"", "reverse=\"true\"", "min_gte=\"false\"", "extra1", "extra2"])
        s["G1"] = 3
        s["G4"] = "y"

        @test XLSX.setConditionalFormat(s, "A2:A11", :iconSet;
            min_type="num", max_type="formula",
            min_val="3", max_val="if(\$G\$4=\"y\", \$G\$1+5, 10)") == 0

        @test XLSX.setConditionalFormat(s, "A2:A11", :iconSet;
            min_type="num", max_type="num",
            min_val="3", max_val="8") == 0

        @test XLSX.setConditionalFormat(s, "B2:B11", :iconSet; iconset="4TrafficLights",
            min_type="num", mid_type="percent", max_type="num",
            min_val="3", mid_val="50", max_val="8",
            showVal="false") == 0

        @test XLSX.setConditionalFormat(s, "C2:C11", :iconSet; iconset="3Symbols2",
            min_type="num", mid_type="percentile", max_type="num",
            min_val="3", mid_val="50", max_val="8",
            reverse="true") == 0

        @test XLSX.setConditionalFormat(s, "D2:D11", :iconSet; iconset="5Arrows",
            min_type="num", mid_type="percentile", mid2_type="percentile", max_type="num",
            min_val="3", mid_val="45", mid2_val="65", max_val="8",
            min_gte="false", max_gte="false") == 0

        @test XLSX.setConditionalFormat(s, "E2:E11", :iconSet; iconset="3Stars",
            reverse="true",
            showVal="false",
            min_type="num", mid_type="percentile", mid2_type="percentile", max_type="num",
            min_val="3", mid_val="45", mid2_val="65", max_val="8",
            min_gte="false", max_gte="false") == 0

        @test XLSX.setConditionalFormat(s, "F2:F11", :iconSet; iconset="5Boxes",
            reverse="true",
            showVal="false",
            min_type="num", mid_type="percentile", mid2_type="percentile", max_type="num",
            min_val="3", mid_val="45", mid2_val="65", max_val="8",
            min_gte="false", mid_gte="false", mid2_gte="false", max_gte="false") == 0

        @test sort!(XLSX.getConditionalFormats(s), by = x -> x.second.priority, rev=true) == [
            XLSX.CellRange("F2:F11") => (type="iconSet", priority=7),
            XLSX.CellRange("E2:E11") => (type="iconSet", priority=6),
            XLSX.CellRange("D2:D11") => (type="iconSet", priority=5),
            XLSX.CellRange("C2:C11") => (type="iconSet", priority=4),
            XLSX.CellRange("B2:B11") => (type="iconSet", priority=3),
            XLSX.CellRange("A2:A11") => (type="iconSet", priority=2),
            XLSX.CellRange("A2:A11") => (type="iconSet", priority=1),
        ]

        f = XLSX.newxlsx()
        s = f[1]
        for i = 0:3
            for j = 1:13
                s[i+1, j] = i * 13 + j
            end
        end
        for j = 1:13
            @test XLSX.setConditionalFormat(s, 1:4, j, :iconSet; # Create a custom 4-icon set in each column.
                iconset="Custom",
                icon_list=[j, 13 + j, 26 + j, 39 + j],
                min_type="percent", mid_type="percent", max_type="percent",
                min_val="25", mid_val="50", max_val="75"
            ) == 0
        end

        @test XLSX.setConditionalFormat(s, 1:4, 1, :iconSet;
            iconset="Custom",
            icon_list=[1, 2, 3, 4, 5],
            min_type="percent", max_type="percent",
            min_val="25", max_val="75",
            min_gte="false", max_gte="false"
        ) == 0
        @test XLSX.setConditionalFormat(s, 1:4, 1, :iconSet;
            iconset="Custom",
            showVal="false",
            icon_list=[1, 2, 3, 4, 5],
            min_type="percent", mid_type="percent", max_type="percent",
            min_val="25", mid_val="50", max_val="75"
        ) == 0
        @test XLSX.setConditionalFormat(s, 1:4, 1, :iconSet;
            iconset="Custom",
            reverse="true",
            icon_list=[1, 2, 3, 4, 5],
            min_type="percent", mid_type="percent", mid2_type="percentile", max_type="percent",
            min_val="25", mid_val="50", mid2_val="60", max_val="75"
        ) == 0

        @test XLSX.setConditionalFormat(s, "A2:M2", :iconSet;
            iconset="Custom",
            icon_list=[31, 24, 11],
            min_type="num", max_type="formula",
            min_val="3", max_val="if(\$G\$4=\"y\", \$G\$1+5, 10)") == 0

        @test_throws XLSX.XLSXError XLSX.setConditionalFormat(s, 1:4, 1, :iconSet;
            iconset="Custom",
            icon_list=[1, 2, 3, 4, 5],
            min_type="percent", mid_type="madeUp", mid2_type="percentile", max_type="num",
            min_val="25", mid_val="50", mid2_val="60", max_val="75"
        )
        @test_throws XLSX.XLSXError XLSX.setConditionalFormat(s, 1:4, 1, :iconSet;
            iconset="Custom",
            icon_list=[99, 2, 3, 4, 5],
            min_type="percent", mid_type="percent", mid2_type="percentile", max_type="num",
            min_val="25", mid_val="50", mid2_val="60", max_val="75"
        )
        @test_throws XLSX.XLSXError XLSX.setConditionalFormat(s, 1:4, 1, :iconSet;
            iconset="Custom",
            min_type="percent", mid_type="percent", max_type="percent",
            min_val="25", mid_val="50", max_val="75"
        )
        @test_throws XLSX.XLSXError XLSX.setConditionalFormat(s, 1:4, 1, :iconSet;
            iconset="Custom",
            icon_list=[],
            min_type="percent", mid_type="percent", max_type="percent",
            min_val="25", mid_val="50", max_val="75"
        )
        @test_throws XLSX.XLSXError XLSX.setConditionalFormat(s, 1:4, 1, :iconSet;
            iconset="Custom",
            icon_list=[1, 13, 26],
            min_type="percent", mid_type="percent", max_type="percent",
            min_val="25", mid_val="50", max_val="75"
        )
        @test_throws XLSX.XLSXError XLSX.setConditionalFormat(s, 1:4, 1, :iconSet;
            iconset="Custom",
            icon_list=[1, 13, 26, 39],
            min_type="percent", max_type="percent",
            min_val="25"
        ) == 0
        @test_throws XLSX.XLSXError XLSX.setConditionalFormat(s, 1:4, 1, :iconSet;
            iconset="Custom",
            icon_list=[1, 13, 26, 39],
            min_type="percent",
            min_val="25", max_val="75"
        ) == 0
        @test_throws XLSX.XLSXError XLSX.setConditionalFormat(s, 1:4, 1, :iconSet;
            iconset="Custom",
            icon_list=[1, 13, 26, 39]
        ) == 0
        @test_throws XLSX.XLSXError XLSX.setConditionalFormat(s, 1:4, 1, :iconSet;
            iconset="10ThousandManiacs",
        ) == 0


        @test sort!(XLSX.getConditionalFormats(s), by = x -> x.second.priority) == [
            XLSX.CellRange("A1:A4") => (type="iconSet", priority=1),
            XLSX.CellRange("B1:B4") => (type="iconSet", priority=2),
            XLSX.CellRange("C1:C4") => (type="iconSet", priority=3),
            XLSX.CellRange("D1:D4") => (type="iconSet", priority=4),
            XLSX.CellRange("E1:E4") => (type="iconSet", priority=5),
            XLSX.CellRange("F1:F4") => (type="iconSet", priority=6),
            XLSX.CellRange("G1:G4") => (type="iconSet", priority=7),
            XLSX.CellRange("H1:H4") => (type="iconSet", priority=8),
            XLSX.CellRange("I1:I4") => (type="iconSet", priority=9),
            XLSX.CellRange("J1:J4") => (type="iconSet", priority=10),
            XLSX.CellRange("K1:K4") => (type="iconSet", priority=11),
            XLSX.CellRange("L1:L4") => (type="iconSet", priority=12),
            XLSX.CellRange("M1:M4") => (type="iconSet", priority=13),
            XLSX.CellRange("A1:A4") => (type="iconSet", priority=14),
            XLSX.CellRange("A1:A4") => (type="iconSet", priority=15),
            XLSX.CellRange("A1:A4") => (type="iconSet", priority=16),
            XLSX.CellRange("A2:M2") => (type="iconSet", priority=17)
        ]

        XLSX.addDefinedName(s, "myRange", "A1:B2")
        XLSX.addDefinedName(s, "myNCRange", "C1:C5,D1:D5")
        @test XLSX.setConditionalFormat(s, "myRange", :iconSet) == 0
        @test_throws MethodError XLSX.setConditionalFormat(s, "myNCRange", :iconSet)

        f = XLSX.newxlsx()
        s = f[1]
        for i in 1:5, j in 1:21
            s[i, j] = i + j
        end
        for (j, k) in enumerate(keys(XLSX.iconsets))
            if k == "Custom"
                @test XLSX.setConditionalFormat(s, :, j, :iconSet;
                    iconset=k,
                    icon_list=[1, 2, 3, 4, 5],
                    min_type="num", mid_type="num", mid2_type="num", max_type="num",
                    min_val="8", mid_val="12", mid2_val="15", max_val="18",
                ) == 0
            else
                @test XLSX.setConditionalFormat(s, :, j, :iconSet; iconset=k) == 0
            end
        end
        f = XLSX.newxlsx()
        s = f[1]
        for i in 1:3, j in 1:21
            s[i, j] = i + j
        end
        @test XLSX.setConditionalFormat(s, "Sheet1!A1:E1", :iconSet;
            min_type="percentile",
            min_val="10",
            max_type="num",
            max_val="Sheet1!\$C\$4"
        ) == 0
        @test XLSX.setConditionalFormat(f, "Sheet1!A2:E2", :iconSet;
            min_type="percentile",
            min_val="Sheet1!\$D\$5",
            max_type="percent",
            max_val="95"
        ) == 0
        @test_throws XLSX.XLSXError XLSX.setConditionalFormat(s, "Sheet1!A1:E1", :iconSet;
            min_type="percentile",
            min_val="10",
            max_type="num",
            max_val="Sheet2!\$C\$4"
        )
        @test_throws XLSX.XLSXError XLSX.setConditionalFormat(f, "Sheet1!A5:E5", :iconSet;
            min_type="percentile",
            min_val="Sheet2!\$D\$5",
            max_type="percent",
            max_val="95"
        )
        @test XML.tag(XLSX.get_x14_icon("3Triangles")) == "x14:cfRule"
        @test XML.attributes(XLSX.get_x14_icon("3Stars")) == OrderedDict("type" => "iconSet", "priority" => "1", "id" => "XXXX-xxxx-XXXX")
        @test length(XML.children(XLSX.get_x14_icon("5Boxes"))) == 1
        @test typeof(XLSX.get_x14_icon("Custom")) == XML.Node
    end

    @testset "cellIs" begin
        f = XLSX.newxlsx()
        s = f[1]
        for i in 1:5, j in 1:5
            s[i, j] = i + j
        end
        @test_throws MethodError XLSX.setConditionalFormat(s, "A1,A3", :cellIs) # Non-contiguous ranges not allowed
        @test_throws MethodError XLSX.setConditionalFormat(s, [1], 1, :cellIs) # Vectors may be non-contiguous
        @test_throws MethodError XLSX.setConditionalFormat(s, 1, 1:3:7, :cellIs) # StepRange is non-contiguous
        @test_throws XLSX.XLSXError XLSX.setConditionalFormat(s, "A1:A3", :cellIs; dxStyle="madeUp") # dxStyle invalid
        @test XLSX.setConditionalFormat(s, "1:1", :cellIs) == 0
        @test XLSX.setConditionalFormat(s, 2, :, :cellIs; dxStyle="greenfilltext") == 0
        @test XLSX.setConditionalFormat(s, 3, 1:5, :cellIs;
            operator="between",
            value="2",
            value2="3",
            stopIfTrue="true",
            fill=["pattern" => "none", "bgColor" => "FFFFC7CE"],
            format=["format" => "0.00%"],
            font=["color" => "blue", "bold" => "true"]
        ) == 0

        @test XLSX.setConditionalFormat(s, "Sheet1!A4:E4", :cellIs;
            operator="greaterThan",
            value="4",
            fill=["pattern" => "none", "bgColor" => "green"],
            format=["format" => "0.0"],
            font=["color" => "red", "italic" => "true"]
        ) == 0
        @test XLSX.setConditionalFormat(f, "Sheet1!A5:E5", :cellIs;
            operator="lessThan",
            value="2",
            fill=["pattern" => "none", "bgColor" => "yellow"],
            format=["format" => "0.0"],
            font=["color" => "green"],
            border=["style" => "thick", "color" => "coral"]
        ) == 0
        @test sort!(XLSX.getConditionalFormats(s), by = x -> x.second.priority, rev=true) == [
            XLSX.CellRange("A5:E5") => (type="cellIs", priority=5), 
            XLSX.CellRange("A4:E4") => (type="cellIs", priority=4), 
            XLSX.CellRange("A3:E3") => (type="cellIs", priority=3), 
            XLSX.CellRange("A2:E2") => (type="cellIs", priority=2), 
            XLSX.CellRange("A1:E1") => (type="cellIs", priority=1)
        ]
        @test XLSX.setConditionalFormat(s, "A1", :cellIs) == 0
        @test XLSX.setConditionalFormat(s, "A1:C3", :cellIs) == 0
        @test XLSX.setConditionalFormat(s, "Sheet1!A1", :cellIs) == 0
        @test XLSX.setConditionalFormat(s, "Sheet1!A1:A2", :cellIs) == 0
        @test XLSX.setConditionalFormat(s, "Sheet1!1:2", :cellIs) == 0
        @test XLSX.setConditionalFormat(s, "2:4", :cellIs) == 0
        @test XLSX.setConditionalFormat(s, "A:C", :cellIs) == 0
        @test XLSX.setConditionalFormat(s, "Sheet1!A:C", :cellIs) == 0
        @test XLSX.setConditionalFormat(f, "Sheet1!1:2", :cellIs) == 0
        @test XLSX.setConditionalFormat(f, "Sheet1!A:C", :cellIs) == 0
        @test XLSX.setConditionalFormat(s, :, 1:3, :cellIs) == 0
        @test XLSX.setConditionalFormat(s, 1:3, :, :cellIs) == 0
        @test XLSX.setConditionalFormat(s, "2:4", :cellIs) == 0
        @test XLSX.setConditionalFormat(s, "A:C", :cellIs) == 0
        @test XLSX.setConditionalFormat(s, "Sheet1!A:C", :cellIs) == 0
        @test XLSX.setConditionalFormat(s, :, :cellIs) == 0
        @test XLSX.setConditionalFormat(s, :, :, :cellIs) == 0
        @test length(XLSX.getConditionalFormats(s)) == 22
        @test sort!(XLSX.getConditionalFormats(s), by = x -> x.second.priority, rev=true) == [
            XLSX.CellRange("A1:E5") => (type="cellIs", priority=22),
            XLSX.CellRange("A1:E5") => (type="cellIs", priority=21),
            XLSX.CellRange("A1:C5") => (type="cellIs", priority=20),
            XLSX.CellRange("A1:C5") => (type="cellIs", priority=19),
            XLSX.CellRange("A2:E4") => (type="cellIs", priority=18),
            XLSX.CellRange("A1:E3") => (type="cellIs", priority=17),
            XLSX.CellRange("A1:C5") => (type="cellIs", priority=16),
            XLSX.CellRange("A1:C5") => (type="cellIs", priority=15),
            XLSX.CellRange("A1:E2") => (type="cellIs", priority=14),
            XLSX.CellRange("A1:C5") => (type="cellIs", priority=13),
            XLSX.CellRange("A1:C5") => (type="cellIs", priority=12),
            XLSX.CellRange("A2:E4") => (type="cellIs", priority=11),
            XLSX.CellRange("A1:E2") => (type="cellIs", priority=10),
            XLSX.CellRange("A1:A2") => (type="cellIs", priority=9),
            XLSX.CellRange("A1:A1") => (type="cellIs", priority=8),
            XLSX.CellRange("A1:C3") => (type="cellIs", priority=7),
            XLSX.CellRange("A1:A1") => (type="cellIs", priority=6),
            XLSX.CellRange("A5:E5") => (type="cellIs", priority=5),
            XLSX.CellRange("A4:E4") => (type="cellIs", priority=4),
            XLSX.CellRange("A3:E3") => (type="cellIs", priority=3),
            XLSX.CellRange("A2:E2") => (type="cellIs", priority=2),
            XLSX.CellRange("A1:E1") => (type="cellIs", priority=1)
        ]
        @test_throws XLSX.XLSXError XLSX.setConditionalFormat(s, "Sheet1!A4:E4", :cellIs;
            operator="madeUp",
            value="4",
            fill=["pattern" => "none", "bgColor" => "green"],
            format=["format" => "0.0"],
            font=["color" => "red", "italic" => "true"]
        )

        f = XLSX.newxlsx()
        s = f[1]
        for i in 1:5, j in 1:5
            s[i, j] = i + j
        end
        XLSX.setConditionalFormat(s, "A1:A5", :cellIs)
        XLSX.setConditionalFormat(s, :, 2, :cellIs; dxStyle="redborder")
        XLSX.setConditionalFormat(s, "Sheet1!E:E", :cellIs; dxStyle="redfilltext")
        XLSX.setConditionalFormat(s, 1:5, 3:4, :cellIs;
            operator="between",
            value="2",
            value2="4",
            fill=["pattern" => "none", "bgColor" => "yellow"],
            format=["format" => "0.0"],
            font=["color" => "green"],
            border=["style" => "thick", "color" => "coral"]
        ) == 0

        @test sort!(XLSX.getConditionalFormats(s), by = x -> x.second.priority, rev=true) == [
            XLSX.CellRange("C1:D5") => (type="cellIs", priority=4), 
            XLSX.CellRange("E1:E5") => (type="cellIs", priority=3), 
            XLSX.CellRange("B1:B5") => (type="cellIs", priority=2), 
            XLSX.CellRange("A1:A5") => (type="cellIs", priority=1)
        ]

        f = XLSX.newxlsx()
        s = f[1]
        for i in 1:5, j in 1:5
            s[i, j] = i + j
        end

        @test XLSX.setConditionalFormat(s, :, 1:4, :cellIs;
            operator="lessThan",
            value="\$E\$4",
            fill=["pattern" => "none", "bgColor" => "yellow"],
            format=["format" => "0.0"],
            font=["color" => "green", "under" => "double"],
            border=["style" => "thin", "color" => "coral"]
        ) == 0

        f = XLSX.newxlsx()
        s = f[1]
        for i in 1:5, j in 1:5
            s[i, j] = i + j
        end
        XLSX.addDefinedName(s, "myRange", "A1:B5")
        @test XLSX.setConditionalFormat(s, "myRange", :cellIs;
            operator="lessThan",
            value="2",
            fill=["pattern" => "none", "bgColor" => "yellow"],
            format=["format" => "0.0"],
            font=["color" => "green"],
            border=["style" => "hair", "color" => "cyan"]
        ) == 0
        XLSX.addDefinedName(s, "myNCRange", "C1:C5,D1:D5")
        @test_throws MethodError XLSX.setConditionalFormat(s, "myNCRange", :cellIs; # Non-contiguous ranges not allowed
            operator="lessThan",
            value="2",
            fill=["pattern" => "none", "bgColor" => "yellow"],
            format=["format" => "0.0"],
            font=["color" => "green"],
            border=["style" => "hair", "color" => "cyan"]
        )

        f = XLSX.newxlsx()
        s = f[1]
        for i in 1:5, j in 1:6
            s[i, j] = i + j
        end
        for (j, k) in enumerate(keys(XLSX.highlights))
            @test XLSX.setConditionalFormat(s, :, j, :cellIs; dxStyle=k) == 0
        end
    end


    @testset "containsText" begin
        f = XLSX.newxlsx()
        s = f[1]
        s["A1:E1"] = "Hello World"
        s["A2:E2"] = "Life the universe and everything"
        s["A3:E3"] = "Once upon a time"
        s["A4:E4"] = "In America"
        s["A5:E5"] = "a"
        @test_throws MethodError XLSX.setConditionalFormat(s, "A1,A3", :containsText; value="a") # Non-contiguous ranges not allowed
        @test_throws MethodError XLSX.setConditionalFormat(s, [1], 1, :containsText; value="a") # Vectors may be non-contiguous
        @test_throws MethodError XLSX.setConditionalFormat(s, 1, 1:3:7, :containsText; value="a") # StepRange is non-contiguous
        @test_throws XLSX.XLSXError XLSX.setConditionalFormat(s, "1:1", :containsText) # value must be defined
        @test XLSX.setConditionalFormat(s, "1:1", :containsText; value="a") == 0
        @test XLSX.setConditionalFormat(s, 2, :, :containsText; value="a", dxStyle="greenfilltext") == 0
        @test XLSX.setConditionalFormat(s, 3, 1:5, :containsText;
            operator="notContainsText",
            value="a",
            fill=["pattern" => "none", "bgColor" => "FFFFC7CE"],
            format=["format" => "0.00%"],
            font=["color" => "blue", "bold" => "true"]
        ) == 0

        @test XLSX.setConditionalFormat(s, "Sheet1!A4:E4", :containsText;
            operator="notContainsText",
            value="a",
            fill=["pattern" => "none", "bgColor" => "green"],
            format=["format" => "0.0"],
            font=["color" => "red", "italic" => "true"]
        ) == 0
        @test XLSX.setConditionalFormat(f, "Sheet1!A5:E5", :containsText;
            operator="beginsWith",
            value="a",
            fill=["pattern" => "none", "bgColor" => "yellow"],
            format=["format" => "0.0"],
            font=["color" => "green"],
            border=["style" => "thick", "color" => "coral"]
        ) == 0
        @test sort!(XLSX.getConditionalFormats(s), by = x -> x.second.priority, rev=true) == [
            XLSX.CellRange("A5:E5") => (type="beginsWith", priority=5), 
            XLSX.CellRange("A4:E4") => (type="notContainsText", priority=4), 
            XLSX.CellRange("A3:E3") => (type="notContainsText", priority=3), 
            XLSX.CellRange("A2:E2") => (type="containsText", priority=2), 
            XLSX.CellRange("A1:E1") => (type="containsText", priority=1)
        ]
        #        @test XLSX.getConditionalFormats(s) == [XLSX.CellRange("A5:E5") => (type = "containsText", priority = 5), XLSX.CellRange("A4:E4") => (type = "containsText", priority = 4), XLSX.CellRange("A3:E3") => (type = "containsText", priority = 3), XLSX.CellRange("A2:E2") => (type = "containsText", priority = 2), XLSX.CellRange("A1:E1") => (type = "containsText", priority = 1)]
        @test XLSX.setConditionalFormat(s, "A1", :containsText; value="a") == 0
        @test XLSX.setConditionalFormat(s, "A1:C3", :containsText; value="a") == 0
        @test XLSX.setConditionalFormat(s, "Sheet1!A1", :containsText; value="a") == 0
        @test XLSX.setConditionalFormat(s, "Sheet1!A1:A2", :containsText; value="a") == 0
        @test XLSX.setConditionalFormat(s, "Sheet1!1:2", :containsText; value="a") == 0
        @test XLSX.setConditionalFormat(s, "2:4", :containsText; value="a") == 0
        @test XLSX.setConditionalFormat(s, "A:C", :containsText; value="a") == 0
        @test XLSX.setConditionalFormat(s, "Sheet1!A:C", :containsText; value="a") == 0
        @test XLSX.setConditionalFormat(f, "Sheet1!1:2", :containsText; value="a") == 0
        @test XLSX.setConditionalFormat(f, "Sheet1!A:C", :containsText; value="a") == 0
        @test XLSX.setConditionalFormat(s, :, 1:3, :containsText; value="a") == 0
        @test XLSX.setConditionalFormat(s, 1:3, :, :containsText; value="a") == 0
        @test XLSX.setConditionalFormat(s, "2:4", :containsText; value="a") == 0
        @test XLSX.setConditionalFormat(s, "A:C", :containsText; value="a") == 0
        @test XLSX.setConditionalFormat(s, "Sheet1!A:C", :containsText; value="a") == 0
        @test XLSX.setConditionalFormat(s, :, :containsText; value="a") == 0
        @test XLSX.setConditionalFormat(s, :, :, :containsText; value="a") == 0
        @test length(XLSX.getConditionalFormats(s)) == 22
        @test sort!(XLSX.getConditionalFormats(s), by = x -> x.second.priority, rev=true) == [
            XLSX.CellRange("A1:E5") => (type="containsText", priority=22),
            XLSX.CellRange("A1:E5") => (type="containsText", priority=21),
            XLSX.CellRange("A1:C5") => (type="containsText", priority=20),
            XLSX.CellRange("A1:C5") => (type="containsText", priority=19),
            XLSX.CellRange("A2:E4") => (type="containsText", priority=18),
            XLSX.CellRange("A1:E3") => (type="containsText", priority=17),
            XLSX.CellRange("A1:C5") => (type="containsText", priority=16),
            XLSX.CellRange("A1:C5") => (type="containsText", priority=15),
            XLSX.CellRange("A1:E2") => (type="containsText", priority=14),
            XLSX.CellRange("A1:C5") => (type="containsText", priority=13),
            XLSX.CellRange("A1:C5") => (type="containsText", priority=12),
            XLSX.CellRange("A2:E4") => (type="containsText", priority=11),
            XLSX.CellRange("A1:E2") => (type="containsText", priority=10),
            XLSX.CellRange("A1:A2") => (type="containsText", priority=9),
            XLSX.CellRange("A1:A1") => (type="containsText", priority=8),
            XLSX.CellRange("A1:C3") => (type="containsText", priority=7),
            XLSX.CellRange("A1:A1") => (type="containsText", priority=6),
            XLSX.CellRange("A5:E5") => (type="beginsWith", priority=5),
            XLSX.CellRange("A4:E4") => (type="notContainsText", priority=4),
            XLSX.CellRange("A3:E3") => (type="notContainsText", priority=3),
            XLSX.CellRange("A2:E2") => (type="containsText", priority=2),
            XLSX.CellRange("A1:E1") => (type="containsText", priority=1)
        ]

        f = XLSX.newxlsx()
        s = f[1]
        s["A1:E1"] = "Hello World"
        s["A2:E2"] = "Life the universe and everything"
        s["A3:E3"] = "Once upon a time"
        s["A4:E4"] = "In America"
        s["A5:E5"] = "a"
        XLSX.setConditionalFormat(s, "A1:A5", :containsText; value="a")
        XLSX.setConditionalFormat(s, :, 2, :containsText; value="a", dxStyle="redborder")
        XLSX.setConditionalFormat(s, "Sheet1!E:E", :containsText; value="a", dxStyle="redfilltext")
        XLSX.setConditionalFormat(s, 1:5, 3:4, :containsText;
            operator="endsWith",
            value="a",
            fill=["pattern" => "none", "bgColor" => "yellow"],
            format=["format" => "0.0"],
            font=["color" => "green"],
            border=["style" => "thick", "color" => "coral"]
        ) == 0

        @test sort!(XLSX.getConditionalFormats(s), by = x -> x.second.priority, rev=true) == [
            XLSX.CellRange("C1:D5") => (type="endsWith", priority=4), 
            XLSX.CellRange("E1:E5") => (type="containsText", priority=3), 
            XLSX.CellRange("B1:B5") => (type="containsText", priority=2), 
            XLSX.CellRange("A1:A5") => (type="containsText", priority=1)
        ]

        f = XLSX.newxlsx()
        s = f[1]
        s["A1:E1"] = "Hello World"
        s["A2:E2"] = "Life the universe and everything"
        s["A3:E3"] = "Once upon a time"
        s["A4:E4"] = "In America"
        s["A5:E5"] = "a"

        @test XLSX.setConditionalFormat(s, :, 1:4, :containsText;
            operator="containsText",
            value="Sheet1!\$E\$5",
            stopIfTrue="true",
            fill=["pattern" => "none", "bgColor" => "yellow"],
            format=["format" => "0.0"],
            font=["color" => "green", "under" => "double"],
            border=["style" => "thin", "color" => "coral"]
        ) == 0

        f = XLSX.newxlsx()
        s = f[1]
        s["A1:E1"] = "Hello World"
        s["A2:E2"] = "Life the universe and everything"
        s["A3:E3"] = "Once upon a time"
        s["A4:E4"] = "In America"
        s["A5:E5"] = "a"
        XLSX.addDefinedName(s, "myRange", "A1:B5")
        @test XLSX.setConditionalFormat(s, "myRange", :containsText;
            operator="notContainsText",
            value="a",
            fill=["pattern" => "none", "bgColor" => "yellow"],
            format=["format" => "0.0"],
            font=["color" => "green"],
            border=["style" => "hair", "color" => "cyan"]
        ) == 0
        @test_throws XLSX.XLSXError XLSX.setConditionalFormat(s, "myRange", :containsText;
            operator="madeUp",
            value="a",
            fill=["pattern" => "none", "bgColor" => "yellow"],
            format=["format" => "0.0"],
            font=["color" => "green"],
            border=["style" => "hair", "color" => "cyan"]
        ) == 0
        XLSX.addDefinedName(s, "myNCRange", "C1:C5,D1:D5")
        @test_throws MethodError XLSX.setConditionalFormat(s, "myNCRange", :containsText; # Non-contiguous ranges not allowed
            operator="beginsWith",
            value="a",
            fill=["pattern" => "none", "bgColor" => "yellow"],
            format=["format" => "0.0"],
            font=["color" => "green"],
            border=["style" => "hair", "color" => "cyan"]
        )

    end

    @testset "top10" begin
        f = XLSX.newxlsx()
        s = f[1]
        for i = 1:10
            for j = 1:10
                s[i, j] = i * j
            end
        end
        @test_throws MethodError XLSX.setConditionalFormat(s, "A1,A3", :top10) # Non-contiguous ranges not allowed
        @test_throws MethodError XLSX.setConditionalFormat(s, [1], 1, :top10) # Vectors may be non-contiguous
        @test_throws MethodError XLSX.setConditionalFormat(s, 1, 1:3:7, :top10) # StepRange is non-contiguous
        @test XLSX.setConditionalFormat(s, "1:1", :top10) == 0
        @test XLSX.setConditionalFormat(s, 2, :, :top10; dxStyle="greenfilltext") == 0
        @test XLSX.setConditionalFormat(s, 1:10, 1:10, :top10;
            operator="topN",
            value="5",
            stopIfTrue="true",
            fill=["pattern" => "none", "bgColor" => "green"],
            border=["style" => "thick", "color" => "coral"],
            font=["color" => "blue", "bold" => "true", "strike" => "true"]
        ) == 0
        @test XLSX.setConditionalFormat(s, 1:10, 1:10, :top10;
            operator="bottomN",
            value="5",
            stopIfTrue="true",
            fill=["pattern" => "lightVertical", "fgColor" => "grey", "bgColor" => "FFFFC7CE"],
            border=["style" => "thick", "color" => "coral"],
            font=["color" => "blue", "bold" => "true", "strike" => "true"]
        ) == 0
        @test XLSX.setConditionalFormat(s, 1:10, 1:10, :top10;
            operator="topN%",
            value="20",
            fill=["pattern" => "none", "bgColor" => "FFFFC7CE"],
            border=["style" => "thick", "color" => "coral"],
            font=["color" => "blue", "bold" => "true", "strike" => "true"]
        ) == 0
        @test XLSX.setConditionalFormat(s, 1:10, 1:10, :top10;
            operator="bottomN%",
            value="30",
            fill=["pattern" => "none", "bgColor" => "pink"],
            border=["style" => "thick", "color" => "coral"],
            font=["color" => "blue", "bold" => "true", "italic" => "true"]
        ) == 0
        @test sort!(XLSX.getConditionalFormats(s), by = x -> x.second.priority) == [
            XLSX.CellRange("A1:J1") => (type="top10", priority=1),
            XLSX.CellRange("A2:J2") => (type="top10", priority=2),
            XLSX.CellRange("A1:J10") => (type="top10", priority=3), 
            XLSX.CellRange("A1:J10") => (type="top10", priority=4), 
            XLSX.CellRange("A1:J10") => (type="top10", priority=5), 
            XLSX.CellRange("A1:J10") => (type="top10", priority=6), 
        ]

        @test XLSX.setConditionalFormat(s, "A1", :top10) == 0
        @test XLSX.setConditionalFormat(s, "A1:C3", :top10) == 0
        @test XLSX.setConditionalFormat(s, "Sheet1!A1", :top10) == 0
        @test XLSX.setConditionalFormat(s, "Sheet1!A1:A2", :top10) == 0
        @test XLSX.setConditionalFormat(s, "Sheet1!1:2", :top10) == 0
        @test XLSX.setConditionalFormat(s, "2:4", :top10) == 0
        @test XLSX.setConditionalFormat(s, "A:C", :top10) == 0
        @test XLSX.setConditionalFormat(s, "Sheet1!A:C", :top10) == 0
        @test XLSX.setConditionalFormat(f, "Sheet1!1:2", :top10) == 0
        @test XLSX.setConditionalFormat(f, "Sheet1!A:C", :top10) == 0
        @test XLSX.setConditionalFormat(s, :, 1:3, :top10) == 0
        @test XLSX.setConditionalFormat(s, 1:3, :, :top10) == 0
        @test XLSX.setConditionalFormat(s, "2:4", :top10) == 0
        @test XLSX.setConditionalFormat(s, "A:C", :top10) == 0
        @test XLSX.setConditionalFormat(s, "Sheet1!A:C", :top10) == 0
        @test XLSX.setConditionalFormat(s, :, :top10) == 0
        @test XLSX.setConditionalFormat(s, :, :, :top10) == 0
        @test length(XLSX.getConditionalFormats(s)) == 23
        @test sort!(XLSX.getConditionalFormats(s), by = x -> x.second.priority) == [
            XLSX.CellRange("A1:J1") => (type="top10", priority=1),
            XLSX.CellRange("A2:J2") => (type="top10", priority=2),
            XLSX.CellRange("A1:J10") => (type="top10", priority=3),
            XLSX.CellRange("A1:J10") => (type="top10", priority=4),
            XLSX.CellRange("A1:J10") => (type="top10", priority=5),
            XLSX.CellRange("A1:J10") => (type="top10", priority=6),
            XLSX.CellRange("A1:A1") => (type="top10", priority=7),
            XLSX.CellRange("A1:C3") => (type="top10", priority=8),
            XLSX.CellRange("A1:A1") => (type="top10", priority=9),
            XLSX.CellRange("A1:A2") => (type="top10", priority=10),
            XLSX.CellRange("A1:J2") => (type="top10", priority=11),
            XLSX.CellRange("A2:J4") => (type="top10", priority=12),
            XLSX.CellRange("A1:C10") => (type="top10", priority=13),
            XLSX.CellRange("A1:C10") => (type="top10", priority=14),
            XLSX.CellRange("A1:J2") => (type="top10", priority=15),
            XLSX.CellRange("A1:C10") => (type="top10", priority=16),
            XLSX.CellRange("A1:C10") => (type="top10", priority=17),
            XLSX.CellRange("A1:J3") => (type="top10", priority=18),
            XLSX.CellRange("A2:J4") => (type="top10", priority=19),
            XLSX.CellRange("A1:C10") => (type="top10", priority=20),
            XLSX.CellRange("A1:C10") => (type="top10", priority=21),
            XLSX.CellRange("A1:J10") => (type="top10", priority=22),
            XLSX.CellRange("A1:J10") => (type="top10", priority=23),
        ]

        f = XLSX.newxlsx()
        s = f[1]
        for i = 1:10
            for j = 1:10
                s[i, j] = i * j
            end
        end
        XLSX.setConditionalFormat(s, "A1:A5", :top10)
        XLSX.setConditionalFormat(s, :, 2, :top10; dxStyle="redborder")
        XLSX.setConditionalFormat(s, "Sheet1!E:E", :top10; dxStyle="redfilltext")
        XLSX.setConditionalFormat(s, 1:5, 3:4, :top10;
            operator="topN%",
            value="20",
            fill=["pattern" => "none", "bgColor" => "yellow"],
            format=["format" => "0.0"],
            font=["color" => "green"],
            border=["style" => "thick", "color" => "coral"]
        ) == 0

        @test sort!(XLSX.getConditionalFormats(s), by = x -> x.second.priority, rev=true) == [
            XLSX.CellRange("C1:D5") => (type="top10", priority=4), 
            XLSX.CellRange("E1:E10") => (type="top10", priority=3), 
            XLSX.CellRange("B1:B10") => (type="top10", priority=2), 
            XLSX.CellRange("A1:A5") => (type="top10", priority=1)]

        f = XLSX.newxlsx()
        s = f[1]
        for i = 1:10
            for j = 1:10
                s[i, j] = i * j
            end
        end

        @test XLSX.setConditionalFormat(s, :, 1:4, :top10;
            operator="bottomN",
            value="\$E\$4",
            fill=["pattern" => "none", "bgColor" => "yellow"],
            format=["format" => "0.0"],
            font=["color" => "green", "under" => "double"],
            border=["style" => "thin", "color" => "coral"]
        ) == 0

        f = XLSX.newxlsx()
        s = f[1]
        for i = 1:10
            for j = 1:10
                s[i, j] = i * j
            end
        end
        XLSX.addDefinedName(s, "myRange", "A1:E5")
        @test XLSX.setConditionalFormat(s, "myRange", :top10;
            operator="topN%",
            value="2",
            fill=["pattern" => "none", "bgColor" => "yellow"],
            format=["format" => "0.0"],
            font=["color" => "green"],
            border=["style" => "medium", "color" => "cyan"]
        ) == 0
        XLSX.addDefinedName(s, "myNCRange", "C1:C5,D1:D5")
        @test_throws MethodError XLSX.setConditionalFormat(s, "myNCRange", :top10; # Non-contiguous ranges not allowed
            operator="bottomN%",
            value="2",
            fill=["pattern" => "none", "bgColor" => "yellow"],
            format=["format" => "0.0"],
            font=["color" => "green"],
            border=["style" => "hair", "color" => "cyan"]
        )
        @test_throws XLSX.XLSXError XLSX.setConditionalFormat(s, "myRange", :top10;
            operator="madeUp",
            value="2",
            fill=["pattern" => "none", "bgColor" => "yellow"],
            format=["format" => "0.0"],
            font=["color" => "green"],
            border=["style" => "hair", "color" => "cyan"]
        )

    end

    @testset "aboveAverage" begin
        f = XLSX.newxlsx()
        s = f[1]
        d = Dist.Normal()
        columns = [rand(d, 1000), rand(d, 1000), rand(d, 1000)]
        XLSX.writetable!(s, columns, ["normal1", "normal2", "normal3"])
        @test_throws MethodError XLSX.setConditionalFormat(s, "A2:A1001,C1:C1000", :aboveAverage) # Non-contiguous ranges not allowed
        @test_throws MethodError XLSX.setConditionalFormat(s, [2, 3, 19], 1:3, :aboveAverage) # Vectors may be non-contiguous
        @test_throws MethodError XLSX.setConditionalFormat(s, 2, 1:3:7, :aboveAverage) # StepRange is non-contiguous
        @test XLSX.setConditionalFormat(s, "2:2", :aboveAverage) == 0
        @test XLSX.setConditionalFormat(s, 2, :, :aboveAverage; dxStyle="greenfilltext") == 0
        @test XLSX.setConditionalFormat(s, 2:10, 1:3, :aboveAverage;
            operator="plus3StdDev",
            stopIfTrue="true",
            fill=["pattern" => "none", "bgColor" => "FFFFC7CE"],
            border=["style" => "thick", "color" => "coral"],
            font=["color" => "blue", "bold" => "true", "strike" => "true"]
        ) == 0
        @test XLSX.setConditionalFormat(s, 2:1001, 1:3, :aboveAverage;
            operator="minus3StdDev",
            stopIfTrue="true",
            fill=["pattern" => "lightVertical", "fgColor" => "grey", "bgColor" => "FFFFC7CE"],
            border=["style" => "thick", "color" => "coral"],
            font=["color" => "blue", "bold" => "true", "strike" => "true"]
        ) == 0
        @test XLSX.setConditionalFormat(s, 2:1001, 1:3, :aboveAverage;
            operator="plus2StdDev",
            stopIfTrue="true",
            fill=["pattern" => "none", "bgColor" => "FFFFC7CE"],
            border=["style" => "thick", "color" => "coral"],
            font=["color" => "blue", "bold" => "true", "strike" => "true"]
        ) == 0
        @test XLSX.setConditionalFormat(s, 2:1001, 1:3, :aboveAverage;
            operator="minus2StdDev",
            stopIfTrue="true",
            fill=["pattern" => "none", "bgColor" => "pink"],
            border=["style" => "thick", "color" => "coral"],
            font=["color" => "blue", "bold" => "true", "italic" => "true"]
        ) == 0
        @test XLSX.setConditionalFormat(s, 2:1001, 1:3, :aboveAverage;
            operator="plus1StdDev",
            stopIfTrue="true",
            fill=["pattern" => "none", "bgColor" => "FFFFCFCE"],
            border=["style" => "thick", "color" => "coral"],
            font=["color" => "yellow", "bold" => "true", "strike" => "true"]
        ) == 0
        @test XLSX.setConditionalFormat(s, 1:1001, 1:3, :aboveAverage;
            operator="minus1StdDev",
            stopIfTrue="true",
            fill=["pattern" => "none", "bgColor" => "yellow"],
            border=["style" => "thick", "color" => "coral"],
            font=["color" => "green", "bold" => "true", "italic" => "true"]
        ) == 0
        @test XLSX.setConditionalFormat(s, 1:1001, 1:3, :aboveAverage;
            operator="aboveAverage",
            fill=["pattern" => "none", "bgColor" => "FFFFCFCE"],
            border=["style" => "thick", "color" => "gray"],
            font=["color" => "yellow", "bold" => "true", "strike" => "true"]
        ) == 0
        @test XLSX.setConditionalFormat(s, 1:1001, 1:3, :aboveAverage;
            operator="belowAverage",
            fill=["pattern" => "none", "bgColor" => "yellow"],
            border=["style" => "thick", "color" => "green"],
            font=["color" => "green", "bold" => "true", "italic" => "true"]
        ) == 0
        @test sort!(XLSX.getConditionalFormats(s), by = x -> x.second.priority) == [
            XLSX.CellRange("A2:C2") => (type="aboveAverage", priority=1),
            XLSX.CellRange("A2:C2") => (type="aboveAverage", priority=2),
            XLSX.CellRange("A2:C10") => (type="aboveAverage", priority=3),
            XLSX.CellRange("A2:C1001") => (type="aboveAverage", priority=4),
            XLSX.CellRange("A2:C1001") => (type="aboveAverage", priority=5),
            XLSX.CellRange("A2:C1001") => (type="aboveAverage", priority=6),
            XLSX.CellRange("A2:C1001") => (type="aboveAverage", priority=7),
            XLSX.CellRange("A1:C1001") => (type="aboveAverage", priority=8),
            XLSX.CellRange("A1:C1001") => (type="aboveAverage", priority=9),
            XLSX.CellRange("A1:C1001") => (type="aboveAverage", priority=10)
         ]

        @test XLSX.setConditionalFormat(s, "A1", :aboveAverage) == 0
        @test XLSX.setConditionalFormat(s, "A1:C3", :aboveAverage) == 0
        @test XLSX.setConditionalFormat(s, "Sheet1!A1", :aboveAverage) == 0
        @test XLSX.setConditionalFormat(s, "Sheet1!A1:A2", :aboveAverage) == 0
        @test XLSX.setConditionalFormat(s, "Sheet1!1:2", :aboveAverage) == 0
        @test XLSX.setConditionalFormat(s, "2:4", :aboveAverage) == 0
        @test XLSX.setConditionalFormat(s, "A:C", :aboveAverage) == 0
        @test XLSX.setConditionalFormat(s, "Sheet1!A:C", :aboveAverage) == 0
        @test XLSX.setConditionalFormat(f, "Sheet1!1:2", :aboveAverage) == 0
        @test XLSX.setConditionalFormat(f, "Sheet1!A:C", :aboveAverage) == 0
        @test XLSX.setConditionalFormat(s, :, 1:3, :aboveAverage) == 0
        @test XLSX.setConditionalFormat(s, 1:3, :, :aboveAverage) == 0
        @test XLSX.setConditionalFormat(s, "2:4", :aboveAverage) == 0
        @test XLSX.setConditionalFormat(s, "A:C", :aboveAverage) == 0
        @test XLSX.setConditionalFormat(s, "Sheet1!A:C", :aboveAverage) == 0
        @test XLSX.setConditionalFormat(s, :, :aboveAverage) == 0
        @test XLSX.setConditionalFormat(s, :, :, :aboveAverage) == 0
        @test length(XLSX.getConditionalFormats(s)) == 27
        @test sort!(XLSX.getConditionalFormats(s), by = x -> x.second.priority) == [
            XLSX.CellRange("A2:C2") => (type="aboveAverage", priority=1),
            XLSX.CellRange("A2:C2") => (type="aboveAverage", priority=2),
            XLSX.CellRange("A2:C10") => (type="aboveAverage", priority=3),
            XLSX.CellRange("A2:C1001") => (type="aboveAverage", priority=4),
            XLSX.CellRange("A2:C1001") => (type="aboveAverage", priority=5),
            XLSX.CellRange("A2:C1001") => (type="aboveAverage", priority=6),
            XLSX.CellRange("A2:C1001") => (type="aboveAverage", priority=7),
            XLSX.CellRange("A1:C1001") => (type="aboveAverage", priority=8),
            XLSX.CellRange("A1:C1001") => (type="aboveAverage", priority=9),
            XLSX.CellRange("A1:C1001") => (type="aboveAverage", priority=10),
            XLSX.CellRange("A1:A1") => (type="aboveAverage", priority=11),
            XLSX.CellRange("A1:C3") => (type="aboveAverage", priority=12),
            XLSX.CellRange("A1:A1") => (type="aboveAverage", priority=13),
            XLSX.CellRange("A1:A2") => (type="aboveAverage", priority=14),
            XLSX.CellRange("A1:C2") => (type="aboveAverage", priority=15),
            XLSX.CellRange("A2:C4") => (type="aboveAverage", priority=16),
            XLSX.CellRange("A1:C1001") => (type="aboveAverage", priority=17),
            XLSX.CellRange("A1:C1001") => (type="aboveAverage", priority=18),
            XLSX.CellRange("A1:C2") => (type="aboveAverage", priority=19),
            XLSX.CellRange("A1:C1001") => (type="aboveAverage", priority=20),
            XLSX.CellRange("A1:C1001") => (type="aboveAverage", priority=21),
            XLSX.CellRange("A1:C3") => (type="aboveAverage", priority=22),
            XLSX.CellRange("A2:C4") => (type="aboveAverage", priority=23),
            XLSX.CellRange("A1:C1001") => (type="aboveAverage", priority=24),
            XLSX.CellRange("A1:C1001") => (type="aboveAverage", priority=25),
            XLSX.CellRange("A1:C1001") => (type="aboveAverage", priority=26),
            XLSX.CellRange("A1:C1001") => (type="aboveAverage", priority=27),
        ]

        f = XLSX.newxlsx()
        s = f[1]
        for i = 1:10
            for j = 1:10
                s[i, j] = i * j
            end
        end
        @test XLSX.setConditionalFormat(s, "A1:A5", :aboveAverage) == 0
        @test XLSX.setConditionalFormat(s, :, 2, :aboveAverage; dxStyle="redborder") == 0
        @test XLSX.setConditionalFormat(s, "Sheet1!E:E", :aboveAverage; dxStyle="redfilltext") == 0
        @test XLSX.setConditionalFormat(s, 1:5, 3:4, :aboveAverage;
            operator="aboveEqAverage",
            fill=["pattern" => "none", "bgColor" => "yellow"],
            format=["format" => "0.0"],
            font=["color" => "green"],
            border=["style" => "thick", "color" => "coral"]
        ) == 0
        @test_throws XLSX.XLSXError XLSX.setConditionalFormat(s, 1:5, 3:4, :aboveAverage;
            operator="madeup",
            fill=["pattern" => "none", "bgColor" => "yellow"],
            format=["format" => "0.0"],
            font=["color" => "green"],
            border=["style" => "thick", "color" => "coral"]
        ) == 0

        @test sort!(XLSX.getConditionalFormats(s), by = x -> x.second.priority, rev=true) == [
            XLSX.CellRange("C1:D5") => (type="aboveAverage", priority=4), 
            XLSX.CellRange("E1:E10") => (type="aboveAverage", priority=3), 
            XLSX.CellRange("B1:B10") => (type="aboveAverage", priority=2), 
            XLSX.CellRange("A1:A5") => (type="aboveAverage", priority=1)
        ]

        f = XLSX.newxlsx()
        s = f[1]
        for i = 1:10
            for j = 1:10
                s[i, j] = i * j
            end
        end

        @test XLSX.setConditionalFormat(s, :, 1:4, :aboveAverage;
            operator="belowEqAverage",
            fill=["pattern" => "none", "bgColor" => "yellow"],
            format=["format" => "0.0"],
            font=["color" => "green", "under" => "double"],
            border=["style" => "thin", "color" => "coral"]
        ) == 0

        f = XLSX.newxlsx()
        s = f[1]
        for i = 1:10
            for j = 1:10
                s[i, j] = i * j
            end
        end
        XLSX.addDefinedName(s, "myRange", "A1:E5")
        @test XLSX.setConditionalFormat(s, "myRange", :aboveAverage;
            operator="aboveEqAverage",
            fill=["pattern" => "none", "bgColor" => "yellow"],
            format=["format" => "0.0"],
            font=["color" => "green"],
            border=["style" => "medium", "color" => "cyan"]
        ) == 0
        XLSX.addDefinedName(s, "myNCRange", "C1:C5,D1:D5")
        @test_throws MethodError XLSX.setConditionalFormat(s, "myNCRange", :aboveAverage; # Non-contiguous ranges not allowed
            operator="belowEqAverage",
            fill=["pattern" => "none", "bgColor" => "yellow"],
            format=["format" => "0.0"],
            font=["color" => "green"],
            border=["style" => "hair", "color" => "cyan"]
        )

    end

    @testset "timePeriod" begin
        f = XLSX.newxlsx()
        s = f[1]
        todaynow = Dates.today()
        s[1, 1:10] = todaynow - Dates.Year(1)
        s[2, 1:10] = todaynow - Dates.Month(1)
        s[3, 1:10] = todaynow - Dates.Day(14)
        s[4, 1:10] = todaynow - Dates.Day(5)
        s[5, 1:10] = todaynow - Dates.Day(1)
        s[6, 1:10] = todaynow
        s[7, 1:10] = todaynow + Dates.Day(1)
        s[8, 1:10] = todaynow + Dates.Day(14)
        s[9, 1:10] = todaynow + Dates.Month(1)
        s[10, 1:10] = todaynow + Dates.Year(1)

        @test_throws MethodError XLSX.setConditionalFormat(s, "A1:A5,C1:C5", :timePeriod) # Non-contiguous ranges not allowed
        @test_throws MethodError XLSX.setConditionalFormat(s, [2, 3, 8], 1:3, :timePeriod) # Vectors may be non-contiguous
        @test_throws MethodError XLSX.setConditionalFormat(s, 2, 1:3:7, :timePeriod) # StepRange is non-contiguous
        @test XLSX.setConditionalFormat(s, "2:2", :timePeriod) == 0
        @test XLSX.setConditionalFormat(s, 2, :, :timePeriod; dxStyle="greenfilltext") == 0
        @test_throws XLSX.XLSXError XLSX.setConditionalFormat(s, 1:10, 1:3, :timePeriod;
            operator="madeUp",
            stopIfTrue="true",
            fill=["pattern" => "none", "bgColor" => "FFFFC7CE"],
            border=["style" => "thick", "color" => "coral"],
            font=["color" => "blue", "bold" => "true", "strike" => "true"]
        )
        @test XLSX.setConditionalFormat(s, 1:10, 1:3, :timePeriod;
            operator="today",
            stopIfTrue="true",
            fill=["pattern" => "none", "bgColor" => "FFFFC7CE"],
            border=["style" => "thick", "color" => "coral"],
            font=["color" => "blue", "bold" => "true", "strike" => "true"]
        ) == 0
        @test XLSX.setConditionalFormat(s, 1:10, 1:3, :timePeriod;
            operator="yesterday",
            stopIfTrue="true",
            fill=["pattern" => "lightVertical", "fgColor" => "grey", "bgColor" => "FFFFC7CE"],
            border=["style" => "thick", "color" => "coral"],
            font=["color" => "blue", "bold" => "true", "strike" => "true"]
        ) == 0
        @test XLSX.setConditionalFormat(s, 1:10, 1:3, :timePeriod;
            operator="tomorrow",
            stopIfTrue="true",
            fill=["pattern" => "none", "bgColor" => "FFFFC7CE"],
            border=["style" => "thick", "color" => "coral"],
            font=["color" => "blue", "bold" => "true", "strike" => "true"]
        ) == 0
        @test XLSX.setConditionalFormat(s, 1:10, 1:3, :timePeriod;
            operator="lastMonth",
            stopIfTrue="true",
            fill=["pattern" => "none", "bgColor" => "pink"],
            border=["style" => "thick", "color" => "coral"],
            font=["color" => "blue", "bold" => "true", "italic" => "true"]
        ) == 0
        @test XLSX.setConditionalFormat(s, 1:10, 1:3, :timePeriod;
            operator="thisMonth",
            stopIfTrue="true",
            fill=["pattern" => "none", "bgColor" => "FFCC4411"],
            border=["style" => "thick", "color" => "coral"],
            font=["color" => "yellow", "bold" => "true", "strike" => "true"]
        ) == 0
        @test XLSX.setConditionalFormat(s, 1:10, 1:3, :timePeriod;
            operator="nextMonth",
            stopIfTrue="true",
            fill=["pattern" => "none", "bgColor" => "FFFFCFCE"],
            border=["style" => "thick", "color" => "coral"],
            font=["color" => "yellow", "bold" => "true", "strike" => "true"]
        ) == 0
        @test XLSX.setConditionalFormat(s, 1:10, 1:3, :timePeriod;
            operator="last7Days",
            stopIfTrue="true",
            fill=["pattern" => "none", "bgColor" => "yellow"],
            border=["style" => "thick", "color" => "coral"],
            font=["color" => "green", "bold" => "true", "italic" => "true"]
        ) == 0
        @test sort!(XLSX.getConditionalFormats(s), by = x -> x.second.priority) == [
            XLSX.CellRange("A2:J2") => (type="timePeriod", priority=1),
            XLSX.CellRange("A2:J2") => (type="timePeriod", priority=2),
            XLSX.CellRange("A1:C10") => (type="timePeriod", priority=3),
            XLSX.CellRange("A1:C10") => (type="timePeriod", priority=4),
            XLSX.CellRange("A1:C10") => (type="timePeriod", priority=5),
            XLSX.CellRange("A1:C10") => (type="timePeriod", priority=6),
            XLSX.CellRange("A1:C10") => (type="timePeriod", priority=7),
            XLSX.CellRange("A1:C10") => (type="timePeriod", priority=8),
            XLSX.CellRange("A1:C10") => (type="timePeriod", priority=9)
        ]

        @test XLSX.setConditionalFormat(s, "A1", :timePeriod) == 0
        @test XLSX.setConditionalFormat(s, "A1:C3", :timePeriod) == 0
        @test XLSX.setConditionalFormat(s, "Sheet1!A1", :timePeriod) == 0
        @test XLSX.setConditionalFormat(s, "Sheet1!A1:A2", :timePeriod) == 0
        @test XLSX.setConditionalFormat(s, "Sheet1!1:2", :timePeriod) == 0
        @test XLSX.setConditionalFormat(s, "2:4", :timePeriod) == 0
        @test XLSX.setConditionalFormat(s, "A:C", :timePeriod) == 0
        @test XLSX.setConditionalFormat(s, "Sheet1!A:C", :timePeriod) == 0
        @test XLSX.setConditionalFormat(f, "Sheet1!1:2", :timePeriod) == 0
        @test XLSX.setConditionalFormat(f, "Sheet1!A:C", :timePeriod) == 0
        @test XLSX.setConditionalFormat(s, :, 1:3, :timePeriod) == 0
        @test XLSX.setConditionalFormat(s, 1:3, :, :timePeriod) == 0
        @test XLSX.setConditionalFormat(s, "2:4", :timePeriod) == 0
        @test XLSX.setConditionalFormat(s, "A:C", :timePeriod) == 0
        @test XLSX.setConditionalFormat(s, "Sheet1!A:C", :timePeriod) == 0
        @test XLSX.setConditionalFormat(s, :, :timePeriod) == 0
        @test XLSX.setConditionalFormat(s, :, :, :timePeriod) == 0
        @test length(XLSX.getConditionalFormats(s)) == 26

        @test sort!(XLSX.getConditionalFormats(s), by = x -> x.second.priority) == [
            XLSX.CellRange("A2:J2") => (type="timePeriod", priority=1),
            XLSX.CellRange("A2:J2") => (type="timePeriod", priority=2),
            XLSX.CellRange("A1:C10") => (type="timePeriod", priority=3),
            XLSX.CellRange("A1:C10") => (type="timePeriod", priority=4),
            XLSX.CellRange("A1:C10") => (type="timePeriod", priority=5),
            XLSX.CellRange("A1:C10") => (type="timePeriod", priority=6),
            XLSX.CellRange("A1:C10") => (type="timePeriod", priority=7),
            XLSX.CellRange("A1:C10") => (type="timePeriod", priority=8),
            XLSX.CellRange("A1:C10") => (type="timePeriod", priority=9),
            XLSX.CellRange("A1:A1") => (type="timePeriod", priority=10),
            XLSX.CellRange("A1:C3") => (type="timePeriod", priority=11),
            XLSX.CellRange("A1:A1") => (type="timePeriod", priority=12),
            XLSX.CellRange("A1:A2") => (type="timePeriod", priority=13),
            XLSX.CellRange("A1:J2") => (type="timePeriod", priority=14),
            XLSX.CellRange("A2:J4") => (type="timePeriod", priority=15),
            XLSX.CellRange("A1:C10") => (type="timePeriod", priority=16),
            XLSX.CellRange("A1:C10") => (type="timePeriod", priority=17),
            XLSX.CellRange("A1:J2") => (type="timePeriod", priority=18),
            XLSX.CellRange("A1:C10") => (type="timePeriod", priority=19),
            XLSX.CellRange("A1:C10") => (type="timePeriod", priority=20),
            XLSX.CellRange("A1:J3") => (type="timePeriod", priority=21),
            XLSX.CellRange("A2:J4") => (type="timePeriod", priority=22),
            XLSX.CellRange("A1:C10") => (type="timePeriod", priority=23),
            XLSX.CellRange("A1:C10") => (type="timePeriod", priority=24),
            XLSX.CellRange("A1:J10") => (type="timePeriod", priority=25),
            XLSX.CellRange("A1:J10") => (type="timePeriod", priority=26)
        ]

        f = XLSX.newxlsx()
        s = f[1]
        s[1, 1:10] = todaynow - Dates.Year(1)
        s[2, 1:10] = todaynow - Dates.Month(1)
        s[3, 1:10] = todaynow - Dates.Day(14)
        s[4, 1:10] = todaynow - Dates.Day(5)
        s[5, 1:10] = todaynow - Dates.Day(1)
        s[6, 1:10] = todaynow
        s[7, 1:10] = todaynow + Dates.Day(1)
        s[8, 1:10] = todaynow + Dates.Day(14)
        s[9, 1:10] = todaynow + Dates.Month(1)
        s[10, 1:10] = todaynow + Dates.Year(1)
        XLSX.setConditionalFormat(s, "A1:A5", :timePeriod)
        XLSX.setConditionalFormat(s, :, 2, :timePeriod; dxStyle="redborder")
        XLSX.setConditionalFormat(s, "Sheet1!E:E", :timePeriod; dxStyle="redfilltext")
        XLSX.setConditionalFormat(s, 1:5, 3:4, :timePeriod;
            operator="lastWeek",
            fill=["pattern" => "none", "bgColor" => "yellow"],
            format=["format" => "0.0"],
            font=["color" => "green"],
            border=["style" => "thick", "color" => "coral"]
        ) == 0

        @test sort!(XLSX.getConditionalFormats(s), by = x -> x.second.priority, rev=true) == [
            XLSX.CellRange("C1:D5") => (type="timePeriod", priority=4), 
            XLSX.CellRange("E1:E10") => (type="timePeriod", priority=3), 
            XLSX.CellRange("B1:B10") => (type="timePeriod", priority=2), 
            XLSX.CellRange("A1:A5") => (type="timePeriod", priority=1)
        ]

        f = XLSX.newxlsx()
        s = f[1]
        s[1, 1:10] = todaynow - Dates.Year(1)
        s[2, 1:10] = todaynow - Dates.Month(1)
        s[3, 1:10] = todaynow - Dates.Day(14)
        s[4, 1:10] = todaynow - Dates.Day(5)
        s[5, 1:10] = todaynow - Dates.Day(1)
        s[6, 1:10] = todaynow
        s[7, 1:10] = todaynow + Dates.Day(1)
        s[8, 1:10] = todaynow + Dates.Day(14)
        s[9, 1:10] = todaynow + Dates.Month(1)
        s[10, 1:10] = todaynow + Dates.Year(1)

        @test XLSX.setConditionalFormat(s, :, 1:4, :timePeriod;
            operator="thisWeek",
            fill=["pattern" => "none", "bgColor" => "yellow"],
            format=["format" => "0.0"],
            font=["color" => "green", "under" => "double"],
            border=["style" => "thin", "color" => "coral"]
        ) == 0

        f = XLSX.newxlsx()
        s = f[1]
        s[1, 1:10] = todaynow - Dates.Year(1)
        s[2, 1:10] = todaynow - Dates.Month(1)
        s[3, 1:10] = todaynow - Dates.Day(14)
        s[4, 1:10] = todaynow - Dates.Day(5)
        s[5, 1:10] = todaynow - Dates.Day(1)
        s[6, 1:10] = todaynow
        s[7, 1:10] = todaynow + Dates.Day(1)
        s[8, 1:10] = todaynow + Dates.Day(14)
        s[9, 1:10] = todaynow + Dates.Month(1)
        s[10, 1:10] = todaynow + Dates.Year(1)
        XLSX.addDefinedName(s, "myRange", "A1:E5")
        @test XLSX.setConditionalFormat(s, "myRange", :timePeriod;
            operator="nextWeek",
            fill=["pattern" => "none", "bgColor" => "yellow"],
            format=["format" => "0.0"],
            font=["color" => "green"],
            border=["style" => "medium", "color" => "cyan"]
        ) == 0
        XLSX.addDefinedName(s, "myNCRange", "C1:C5,D1:D5")
        @test_throws MethodError XLSX.setConditionalFormat(s, "myNCRange", :timePeriod; # Non-contiguous ranges not allowed
            operator="lastWeek",
            fill=["pattern" => "none", "bgColor" => "yellow"],
            format=["format" => "0.0"],
            font=["color" => "green"],
            border=["style" => "hair", "color" => "cyan"]
        )

    end

    @testset "expression" begin
        f = XLSX.newxlsx()
        s = f[1]
        for i = 1:10
            for j = 1:10
                s[i, j] = i * j
            end
        end

        @test_throws MethodError XLSX.setConditionalFormat(s, "A1:A5,C1:C5", :expression; formula="A1>3") # Non-contiguous ranges not allowed
        @test_throws MethodError XLSX.setConditionalFormat(s, [2, 3, 8], 1:3, :expression; formula="A1 > 11") # Vectors may be non-contiguous
        @test_throws MethodError XLSX.setConditionalFormat(s, 2, 1:3:7, :expression; formula="A1 < 7") # StepRange is non-contiguous
        @test XLSX.setConditionalFormat(s, "2:2", :expression; formula="A1 = 16") == 0
        @test XLSX.setConditionalFormat(s, 2, :, :expression; formula="A1 < 16", dxStyle="greenfilltext") == 0
        @test_throws XLSX.XLSXError XLSX.setConditionalFormat(s, 1:10, 1:3, :expression;
            stopIfTrue="true",
            fill=["pattern" => "none", "bgColor" => "FFFFC7CE"],
            border=["style" => "thick", "color" => "coral"],
            font=["color" => "blue", "bold" => "true", "strike" => "true"]
        ) == 0
        @test XLSX.setConditionalFormat(s, 1:10, 1:3, :expression;
            formula="A1 > 15",
            stopIfTrue="true",
            fill=["pattern" => "none", "bgColor" => "FFFFC7CE"],
            border=["style" => "thick", "color" => "coral"],
            font=["color" => "blue", "bold" => "true", "strike" => "true"]
        ) == 0
        @test XLSX.setConditionalFormat(s, 1:10, 1:3, :expression;
            formula="iseven(A1)",
            stopIfTrue="true",
            fill=["pattern" => "lightVertical", "fgColor" => "grey", "bgColor" => "FFFFC7CE"],
            border=["style" => "thick", "color" => "coral"],
            font=["color" => "blue", "bold" => "true", "strike" => "true"]
        ) == 0
        @test XLSX.setConditionalFormat(s, 1:10, 1:3, :expression;
            formula="A1 < 10",
            stopIfTrue="true",
            fill=["pattern" => "none", "bgColor" => "FFFFC7CE"],
            border=["style" => "thick", "color" => "coral"],
            font=["color" => "blue", "bold" => "true", "strike" => "true"]
        ) == 0
        @test XLSX.setConditionalFormat(s, 1:10, 1:3, :expression;
            formula="A1 < 5",
            stopIfTrue="true",
            fill=["pattern" => "none", "bgColor" => "pink"],
            border=["style" => "thick", "color" => "coral"],
            font=["color" => "blue", "bold" => "true", "italic" => "true"]
        ) == 0
        @test sort!(XLSX.getConditionalFormats(s), by = x -> x.second.priority) == [
            XLSX.CellRange("A2:J2") => (type="expression", priority=1),
            XLSX.CellRange("A2:J2") => (type="expression", priority=2),
            XLSX.CellRange("A1:C10") => (type="expression", priority=3),
            XLSX.CellRange("A1:C10") => (type="expression", priority=4),
            XLSX.CellRange("A1:C10") => (type="expression", priority=5),
            XLSX.CellRange("A1:C10") => (type="expression", priority=6),
        ]

        @test XLSX.setConditionalFormat(s, "A1", :expression; formula="iseven(A1)") == 0
        @test XLSX.setConditionalFormat(s, "A1:C3", :expression; formula="iseven(A1)") == 0
        @test XLSX.setConditionalFormat(s, "Sheet1!A1", :expression; formula="iseven(A1)") == 0
        @test XLSX.setConditionalFormat(s, "Sheet1!A1:A2", :expression; formula="iseven(A1)") == 0
        @test XLSX.setConditionalFormat(s, "Sheet1!1:2", :expression; formula="iseven(A1)") == 0
        @test XLSX.setConditionalFormat(s, "2:4", :expression; formula="iseven(A1)") == 0
        @test XLSX.setConditionalFormat(s, "A:C", :expression; formula="iseven(A1)") == 0
        @test XLSX.setConditionalFormat(s, "Sheet1!A:C", :expression; formula="iseven(A1)") == 0
        @test XLSX.setConditionalFormat(f, "Sheet1!1:2", :expression; formula="iseven(A1)") == 0
        @test XLSX.setConditionalFormat(f, "Sheet1!A:C", :expression; formula="iseven(A1)") == 0
        @test XLSX.setConditionalFormat(s, :, 1:3, :expression; formula="iseven(A1)") == 0
        @test XLSX.setConditionalFormat(s, 1:3, :, :expression; formula="iseven(A1)") == 0
        @test XLSX.setConditionalFormat(s, "2:4", :expression; formula="iseven(A1)") == 0
        @test XLSX.setConditionalFormat(s, "A:C", :expression; formula="iseven(A1)") == 0
        @test XLSX.setConditionalFormat(s, "Sheet1!A:C", :expression; formula="iseven(A1)") == 0
        @test XLSX.setConditionalFormat(s, :, :expression; formula="iseven(A1)") == 0
        @test XLSX.setConditionalFormat(s, :, :, :expression; formula="iseven(A1)") == 0
        @test length(XLSX.getConditionalFormats(s)) == 23
        @test sort!(XLSX.getConditionalFormats(s), by = x -> x.second.priority, rev=true) == [
            XLSX.CellRange("A1:J10") => (type="expression", priority=23),
            XLSX.CellRange("A1:J10") => (type="expression", priority=22),
            XLSX.CellRange("A1:C10") => (type="expression", priority=21),
            XLSX.CellRange("A1:C10") => (type="expression", priority=20),
            XLSX.CellRange("A2:J4") => (type="expression", priority=19),
            XLSX.CellRange("A1:J3") => (type="expression", priority=18),
            XLSX.CellRange("A1:C10") => (type="expression", priority=17),
            XLSX.CellRange("A1:C10") => (type="expression", priority=16),
            XLSX.CellRange("A1:J2") => (type="expression", priority=15),
            XLSX.CellRange("A1:C10") => (type="expression", priority=14),
            XLSX.CellRange("A1:C10") => (type="expression", priority=13),
            XLSX.CellRange("A2:J4") => (type="expression", priority=12),
            XLSX.CellRange("A1:J2") => (type="expression", priority=11),
            XLSX.CellRange("A1:A2") => (type="expression", priority=10),
            XLSX.CellRange("A1:A1") => (type="expression", priority=9),
            XLSX.CellRange("A1:C3") => (type="expression", priority=8),
            XLSX.CellRange("A1:A1") => (type="expression", priority=7),
            XLSX.CellRange("A1:C10") => (type="expression", priority=6),
            XLSX.CellRange("A1:C10") => (type="expression", priority=5),
            XLSX.CellRange("A1:C10") => (type="expression", priority=4),
            XLSX.CellRange("A1:C10") => (type="expression", priority=3),
            XLSX.CellRange("A2:J2") => (type="expression", priority=2),
            XLSX.CellRange("A2:J2") => (type="expression", priority=1),
        ]

        f = XLSX.newxlsx()
        s = f[1]
        for i = 1:10
            for j = 1:10
                s[i, j] = i * j
            end
        end
        XLSX.setConditionalFormat(s, "A1:A5", :expression; formula="A1=1")
        XLSX.setConditionalFormat(s, :, 2, :expression; formula="A1=1", dxStyle="redborder")
        XLSX.setConditionalFormat(s, "Sheet1!E:E", :expression; formula="A1=1", dxStyle="redfilltext")
        XLSX.setConditionalFormat(s, 1:5, 3:4, :expression;
            formula="A1=1",
            fill=["pattern" => "none", "bgColor" => "yellow"],
            format=["format" => "0.0"],
            font=["color" => "green"],
            border=["style" => "thick", "color" => "coral"]
        ) == 0

        @test sort!(XLSX.getConditionalFormats(s), by = x -> x.second.priority, rev=true) == [
            XLSX.CellRange("C1:D5") => (type="expression", priority=4), 
            XLSX.CellRange("E1:E10") => (type="expression", priority=3), 
            XLSX.CellRange("B1:B10") => (type="expression", priority=2), 
            XLSX.CellRange("A1:A5") => (type="expression", priority=1)
        ]

        f = XLSX.newxlsx()
        s = f[1]
        for i = 1:10
            for j = 1:10
                s[i, j] = i * j
            end
        end
        @test XLSX.setConditionalFormat(s, :, 1:4, :expression;
            formula="A1 > \$E\$3",
            fill=["pattern" => "none", "bgColor" => "yellow"],
            format=["format" => "0.0"],
            font=["color" => "green", "under" => "double"],
            border=["style" => "thin", "color" => "coral"]
        ) == 0

        f = XLSX.newxlsx()
        s = f[1]
        for i = 1:10
            for j = 1:10
                s[i, j] = i * j
            end
        end
        XLSX.addDefinedName(f, "myTest", "Sheet1!L11")
        s["L11"] = 70
        XLSX.addDefinedName(s, "myRange", "F6:J10")

        @test XLSX.setConditionalFormat(s, "myRange", :expression;
            formula="E5 > myTest",
            fill=["pattern" => "none", "bgColor" => "yellow"],
            format=["format" => "0.0"],
            font=["color" => "green"],
            border=["style" => "medium", "color" => "cyan"]
        ) == 0
        XLSX.addDefinedName(s, "myNCRange", "C1:C5,D1:D5")
        @test_throws MethodError XLSX.setConditionalFormat(s, "myNCRange", :expression; # Non-contiguous ranges not allowed
            formula="C4 < myTest",
            fill=["pattern" => "none", "bgColor" => "yellow"],
            format=["format" => "0.0"],
            font=["color" => "green"],
            border=["style" => "hair", "color" => "cyan"]
        )

    end

    @testset "containsErrors" begin
        f = XLSX.newxlsx()
        s = f[1]
        for i = 1:10
            for j = 1:10
                s[i, j] = i * j
            end
        end
        @test_throws MethodError XLSX.setConditionalFormat(s, "A1:A5,C1:C5", :containsErrors) # Non-contiguous ranges not allowed
        @test_throws MethodError XLSX.setConditionalFormat(s, [2, 3, 8], 1:3, :containsErrors) # Vectors may be non-contiguous
        @test_throws MethodError XLSX.setConditionalFormat(s, 2, 1:3:7, :containsErrors) # StepRange is non-contiguous
        @test XLSX.setConditionalFormat(s, "2:2", :containsErrors) == 0
        @test XLSX.setConditionalFormat(s, 2, :, :containsErrors; dxStyle="greenfilltext") == 0
        @test XLSX.setConditionalFormat(s, 1:10, 1:3, :containsErrors;
            stopIfTrue="true",
            fill=["pattern" => "none", "bgColor" => "FFFFC7CE"],
            border=["style" => "thick", "color" => "coral"],
            font=["color" => "blue", "bold" => "true", "strike" => "true"]
        ) == 0
        @test XLSX.setConditionalFormat(s, 1:10, 1:3, :notContainsErrors;
            stopIfTrue="true",
            fill=["pattern" => "lightVertical", "fgColor" => "grey", "bgColor" => "FFFFC7CE"],
            border=["style" => "thick", "color" => "coral"],
            font=["color" => "blue", "bold" => "true", "strike" => "true"]
        ) == 0
        @test XLSX.setConditionalFormat(s, 1:10, 1:3, :containsBlanks;
            stopIfTrue="true",
            fill=["pattern" => "none", "bgColor" => "FFFFC7CE"],
            border=["style" => "thick", "color" => "coral"],
            font=["color" => "blue", "bold" => "true", "strike" => "true"]
        ) == 0
        @test XLSX.setConditionalFormat(s, 1:10, 1:3, :notContainsBlanks;
            stopIfTrue="true",
            fill=["pattern" => "none", "bgColor" => "pink"],
            border=["style" => "thick", "color" => "coral"],
            font=["color" => "blue", "bold" => "true", "italic" => "true"]
        ) == 0
        @test XLSX.setConditionalFormat(s, 1:10, 1:3, :uniqueValues;
            stopIfTrue="true",
            fill=["pattern" => "none", "bgColor" => "FFFFCFCE"],
            border=["style" => "thick", "color" => "coral"],
            font=["color" => "yellow", "bold" => "true", "strike" => "true"]
        ) == 0
        @test XLSX.setConditionalFormat(s, 1:10, 1:3, :duplicateValues;
            stopIfTrue="true",
            fill=["pattern" => "none", "bgColor" => "yellow"],
            border=["style" => "thick", "color" => "coral"],
            font=["color" => "green", "bold" => "true", "italic" => "true"]
        ) == 0
        @test sort!(XLSX.getConditionalFormats(s), by = x -> x.second.priority) == [
            XLSX.CellRange("A2:J2") => (type="containsErrors", priority=1),
            XLSX.CellRange("A2:J2") => (type="containsErrors", priority=2),
            XLSX.CellRange("A1:C10") => (type="containsErrors", priority=3),
            XLSX.CellRange("A1:C10") => (type="notContainsErrors", priority=4),
            XLSX.CellRange("A1:C10") => (type="containsBlanks", priority=5),
            XLSX.CellRange("A1:C10") => (type="notContainsBlanks", priority=6),
            XLSX.CellRange("A1:C10") => (type="uniqueValues", priority=7),
            XLSX.CellRange("A1:C10") => (type="duplicateValues", priority=8),
        ]

        @test XLSX.setConditionalFormat(s, "A1", :containsErrors) == 0
        @test XLSX.setConditionalFormat(s, "A1:C3", :notContainsErrors) == 0
        @test XLSX.setConditionalFormat(s, "Sheet1!A1", :containsBlanks) == 0
        @test XLSX.setConditionalFormat(s, "Sheet1!A1:A2", :notContainsBlanks) == 0
        @test XLSX.setConditionalFormat(s, "Sheet1!1:2", :uniqueValues) == 0
        @test XLSX.setConditionalFormat(s, "2:4", :duplicateValues) == 0
        @test XLSX.setConditionalFormat(s, "A:C", :containsErrors) == 0
        @test XLSX.setConditionalFormat(s, "Sheet1!A:C", :containsErrors) == 0
        @test XLSX.setConditionalFormat(f, "Sheet1!1:2", :containsErrors) == 0
        @test XLSX.setConditionalFormat(f, "Sheet1!A:C", :containsErrors) == 0
        @test XLSX.setConditionalFormat(s, :, 1:3, :containsErrors) == 0
        @test XLSX.setConditionalFormat(s, 1:3, :, :notContainsErrors) == 0
        @test XLSX.setConditionalFormat(s, "2:4", :containsBlanks) == 0
        @test XLSX.setConditionalFormat(s, "A:C", :notContainsBlanks) == 0
        @test XLSX.setConditionalFormat(s, "Sheet1!A:C", :containsErrors) == 0
        @test XLSX.setConditionalFormat(s, :, :uniqueValues) == 0
        @test XLSX.setConditionalFormat(s, :, :, :duplicateValues) == 0
        @test length(XLSX.getConditionalFormats(s)) == 25
        @test sort!(XLSX.getConditionalFormats(s), by = x -> x.second.priority, rev=true) == [
            XLSX.CellRange("A1:J10") => (type="duplicateValues", priority=25),
            XLSX.CellRange("A1:J10") => (type="uniqueValues", priority=24),
            XLSX.CellRange("A1:C10") => (type="containsErrors", priority=23),
            XLSX.CellRange("A1:C10") => (type="notContainsBlanks", priority=22),
            XLSX.CellRange("A2:J4") => (type="containsBlanks", priority=21),
            XLSX.CellRange("A1:J3") => (type="notContainsErrors", priority=20),
            XLSX.CellRange("A1:C10") => (type="containsErrors", priority=19),
            XLSX.CellRange("A1:C10") => (type="containsErrors", priority=18),
            XLSX.CellRange("A1:J2") => (type="containsErrors", priority=17),
            XLSX.CellRange("A1:C10") => (type="containsErrors", priority=16),
            XLSX.CellRange("A1:C10") => (type="containsErrors", priority=15),
            XLSX.CellRange("A2:J4") => (type="duplicateValues", priority=14),
            XLSX.CellRange("A1:J2") => (type="uniqueValues", priority=13),
            XLSX.CellRange("A1:A2") => (type="notContainsBlanks", priority=12),
            XLSX.CellRange("A1:A1") => (type="containsBlanks", priority=11),
            XLSX.CellRange("A1:C3") => (type="notContainsErrors", priority=10),
            XLSX.CellRange("A1:A1") => (type="containsErrors", priority=9),
            XLSX.CellRange("A1:C10") => (type="duplicateValues", priority=8),
            XLSX.CellRange("A1:C10") => (type="uniqueValues", priority=7),
            XLSX.CellRange("A1:C10") => (type="notContainsBlanks", priority=6),
            XLSX.CellRange("A1:C10") => (type="containsBlanks", priority=5),
            XLSX.CellRange("A1:C10") => (type="notContainsErrors", priority=4),
            XLSX.CellRange("A1:C10") => (type="containsErrors", priority=3),
            XLSX.CellRange("A2:J2") => (type="containsErrors", priority=2),
            XLSX.CellRange("A2:J2") => (type="containsErrors", priority=1),
        ]

        f = XLSX.newxlsx()
        s = f[1]
        for i = 1:10
            for j = 1:10
                s[i, j] = i * j
            end
        end
        XLSX.setConditionalFormat(s, "A1:A5", :containsErrors)
        XLSX.setConditionalFormat(s, :, 2, :notContainsErrors; dxStyle="redborder")
        XLSX.setConditionalFormat(s, "Sheet1!E:E", :containsBlanks; dxStyle="redfilltext")
        XLSX.setConditionalFormat(s, 1:5, 3:4, :uniqueValues;
            fill=["pattern" => "none", "bgColor" => "yellow"],
            format=["format" => "0.0"],
            font=["color" => "green"],
            border=["style" => "thick", "color" => "coral"]
        ) == 0

        @test sort!(XLSX.getConditionalFormats(s), by = x -> x.second.priority, rev=true) == [
            XLSX.CellRange("C1:D5") => (type="uniqueValues", priority=4), 
            XLSX.CellRange("E1:E10") => (type="containsBlanks", priority=3), 
            XLSX.CellRange("B1:B10") => (type="notContainsErrors", priority=2), 
            XLSX.CellRange("A1:A5") => (type="containsErrors", priority=1)
        ]

        f = XLSX.newxlsx()
        s = f[1]
        for i = 1:10
            for j = 1:10
                s[i, j] = i * j
            end
        end

        @test XLSX.setConditionalFormat(s, :, 1:4, :containsErrors;
            fill=["pattern" => "none", "bgColor" => "yellow"],
            format=["format" => "0.0"],
            font=["color" => "green", "under" => "double"],
            border=["style" => "thin", "color" => "coral"]
        ) == 0

        f = XLSX.newxlsx()
        s = f[1]
        for i = 1:10
            for j = 1:10
                s[i, j] = i * j
            end
        end
        XLSX.addDefinedName(s, "myRange", "A1:E5")
        @test XLSX.setConditionalFormat(s, "myRange", :containsErrors;
            fill=["pattern" => "none", "bgColor" => "yellow"],
            format=["format" => "0.0"],
            font=["color" => "green"],
            border=["style" => "medium", "color" => "cyan"]
        ) == 0
        XLSX.addDefinedName(s, "myNCRange", "C1:C5,D1:D5")
        @test_throws MethodError XLSX.setConditionalFormat(s, "myNCRange", :containsErrors; # Non-contiguous ranges not allowed
            fill=["pattern" => "none", "bgColor" => "yellow"],
            format=["format" => "0.0"],
            font=["color" => "green"],
            border=["style" => "hair", "color" => "cyan"]
        )

    end

end

@testset "merged cells" begin
    XLSX.openxlsx(joinpath(data_directory, "customXml.xlsx")) do f
        @test_throws XLSX.XLSXError XLSX.getMergedCells(f["Mock-up"]) # File isn't writeable
    end
    f = XLSX.openxlsx(joinpath(data_directory, "customXml.xlsx"); mode="rw")
    mc = sort(XLSX.getMergedCells(f["Mock-up"]))
    @test length(mc) == 25
    @test mc == sort(XLSX.CellRange[XLSX.CellRange("D49:H49"), XLSX.CellRange("D72:J72"), XLSX.CellRange("F94:J94"), XLSX.CellRange("F96:J96"), XLSX.CellRange("F84:J84"), XLSX.CellRange("F86:J86"), XLSX.CellRange("D62:J63"), XLSX.CellRange("D51:J53"), XLSX.CellRange("D55:J60"), XLSX.CellRange("D92:J92"), XLSX.CellRange("D82:J82"), XLSX.CellRange("D74:J74"), XLSX.CellRange("D67:J68"), XLSX.CellRange("D47:H47"), XLSX.CellRange("D9:H9"), XLSX.CellRange("D11:G11"), XLSX.CellRange("D12:G12"), XLSX.CellRange("D14:E14"), XLSX.CellRange("D16:E16"), XLSX.CellRange("D32:F32"), XLSX.CellRange("D38:J38"), XLSX.CellRange("D34:J34"), XLSX.CellRange("D18:E18"), XLSX.CellRange("D20:E20"), XLSX.CellRange("D13:G13")])
    s = f["Mock-up"]
    @test XLSX.isMergedCell(f, "Mock-up!D47")
    @test XLSX.isMergedCell(f, "Mock-up!D49"; mergedCells=mc)
    @test XLSX.isMergedCell(s, "H84")
    @test XLSX.isMergedCell(s, "G84"; mergedCells=mc)
    @test XLSX.isMergedCell(s, "Short_Description")
    @test !XLSX.isMergedCell(f, "Mock-up!B2")
    @test !XLSX.isMergedCell(s, "H40"; mergedCells=mc)
    @test !XLSX.isMergedCell(s, "ID"; mergedCells=mc)
    @test_throws XLSX.XLSXError XLSX.isMergedCell(s, "Contiguous"; mergedCells=mc) # Can't test a range
    @test_throws XLSX.XLSXError XLSX.getMergedBaseCell(s, "Location")

    @test XLSX.getMergedBaseCell(f[1], "F72") == (baseCell=XLSX.CellRef("D72"), baseValue=Dates.Date("2025-03-24"))
    @test XLSX.getMergedBaseCell(f, "Mock-up!G72") == (baseCell=XLSX.CellRef("D72"), baseValue=Dates.Date("2025-03-24"))
    @test XLSX.getMergedBaseCell(s, "H53") == (baseCell=XLSX.CellRef("D51"), baseValue="Hello World")
    @test XLSX.getMergedBaseCell(s, "G52") == (baseCell=XLSX.CellRef("D51"), baseValue="Hello World")
    @test XLSX.getMergedBaseCell(s, 53, 8) == (baseCell=XLSX.CellRef("D51"), baseValue="Hello World")
    @test XLSX.getMergedBaseCell(s, "Short_Description") == (baseCell=XLSX.CellRef("D51"), baseValue="Hello World")
    @test isnothing(XLSX.getMergedBaseCell(s, "F73"))
    @test isnothing(XLSX.getMergedBaseCell(f, "Mock-up!H73"))
    @test_throws XLSX.XLSXError XLSX.getMergedBaseCell(s, "Location") # Can't get base cell for a range

    @test isnothing(XLSX.getMergedCells(f["Document History"]))
    s = f["Document History"]
    @test !XLSX.isMergedCell(f, "Document History!B2")
    @test !XLSX.isMergedCell(s, "C5"; mergedCells=XLSX.getMergedCells(f["Document History"]))

    f = XLSX.opentemplate(joinpath(data_directory, "testmerge.xlsx"))
    @test XLSX.mergeCells(f, "Sheet1!A1:B2") == 0
    @test f[1]["A1"] == "Tables"
    @test ismissing(f[1]["B2"])
    @test f[1]["C3"] == 4
    @test XLSX.mergeCells(f[1], 4:6, 4:6) == 0
    @test f[1][4, 4] == 9
    @test ismissing(f[1][5, 5])
    @test f[1][7, 7] == 36
    @test XLSX.mergeCells(f[1], "J") == 0
    @test f[1]["J1"] == 9
    @test ismissing(f[1]["J2"])
    @test ismissing(f[1]["J12"])
    @test XLSX.isMergedCell(f[1], "J8")
    mc = XLSX.getMergedCells(f["Sheet1"])
    @test XLSX.isMergedCell(f[1], "J9"; mergedCells=mc)
    @test XLSX.getMergedBaseCell(f[1], "J12") == (baseCell=XLSX.CellRef("J1"), baseValue=9)

    @test_throws XLSX.XLSXError XLSX.mergeCells(f[1], "Sheet1!M13:M13")       # Single cell
    @test_throws XLSX.XLSXError XLSX.mergeCells(f[1], 1, :)                   # Overlapping
    @test_throws XLSX.XLSXError XLSX.mergeCells(f[1], 10, :)                  # Overlapping
    @test_throws XLSX.XLSXError XLSX.mergeCells(f["Sheet1"], "M1:P15")        # Outside dimension
    @test_throws XLSX.XLSXError XLSX.mergeCells(f["Sheet1"], "Sheet2!L1:M2")  # Sheets don't match

    XLSX.writexlsx("outfile.xlsx", f, overwrite=true)

    XLSX.openxlsx("outfile.xlsx"; mode="rw") do f
        mc = sort(XLSX.getMergedCells(f["Sheet1"]))
        @test length(mc) == 3
        @test mc == sort(XLSX.CellRange[XLSX.CellRange("A1:B2"), XLSX.CellRange("D4:F6"), XLSX.CellRange("J1:J13")])
        @test XLSX.isMergedCell(f[1], "B2")
        @test XLSX.isMergedCell(f[1], 6, 6; mergedCells=mc)
        @test XLSX.getMergedBaseCell(f[1], "F6") == (baseCell=XLSX.CellRef("D4"), baseValue=9)
        @test f[1]["A1"] == "Tables"
        @test ismissing(f[1]["B2"])
        @test f[1]["C3"] == 4
        @test f[1][4, 4] == 9
        @test ismissing(f[1][5, 5])
        @test f[1][7, 7] == 36
        @test f[1]["J1"] == 9
        @test ismissing(f[1]["J2"])
        @test ismissing(f[1]["J12"])
        @test XLSX.isMergedCell(f[1], "J8")
        @test XLSX.isMergedCell(f[1], "J9"; mergedCells=XLSX.getMergedCells(f["Sheet1"]))
        @test XLSX.getMergedBaseCell(f[1], "J12") == (baseCell=XLSX.CellRef("J1"), baseValue=9)
    end
    isfile("outfile.xlsx") && rm("outfile.xlsx")

    f = XLSX.newxlsx()
    s = f[1]
    for i in 1:3, j in 1:3
        s[i, j] = i + j
    end
    XLSX.mergeCells(s, "Sheet1!A:B")
    @test XLSX.getMergedBaseCell(f, "Sheet1!B2") == (baseCell=XLSX.CellRef("A1"), baseValue=2)

    f = XLSX.newxlsx()
    s = f[1]
    for i in 1:4, j in 1:4
        s[i, j] = i + j
    end
    XLSX.mergeCells(s, "Sheet1!2:3")
    @test XLSX.getMergedBaseCell(f, "Sheet1!C3") == (baseCell=XLSX.CellRef("A2"), baseValue=3)
    XLSX.mergeCells(s, "Sheet1!4:4")
    @test XLSX.getMergedBaseCell(f, "Sheet1!C4") == (baseCell=XLSX.CellRef("A4"), baseValue=5)

    f = XLSX.newxlsx()
    s = f[1]
    for i in 1:3, j in 1:3
        s[i, j] = i + j
    end
    XLSX.mergeCells(s, :, 2:3)
    @test XLSX.getMergedBaseCell(f, "Sheet1!C3") == (baseCell=XLSX.CellRef("B1"), baseValue=3)

    f = XLSX.newxlsx()
    s = f[1]
    for i in 1:3, j in 1:3
        s[i, j] = i + j
    end
    XLSX.mergeCells(s, :, :)
    @test XLSX.getMergedBaseCell(f, "Sheet1!B2") == (baseCell=XLSX.CellRef("A1"), baseValue=2)

    f = XLSX.newxlsx()
    s = f[1]
    for i in 1:3, j in 1:3
        s[i, j] = i + j
    end
    XLSX.mergeCells(s, :)
    @test XLSX.getMergedBaseCell(f, "Sheet1!B2") == (baseCell=XLSX.CellRef("A1"), baseValue=2)



end

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
    end

    @testset "write matrix with heterogeneous data types" begin
        # issue #97
        test_data = ["A" "B"; 1 2; "a" "b"; 0 "x"]
        XLSX.openxlsx(filename, mode="w") do xf
            sheet = xf[1]
            sheet["B2"] = test_data
        end

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
        end

        let
            xf = XLSX.openxlsx(filename)
            sheet = xf[1]
            @test sheet["A1"] == "openxlsx without do-syntax"
        end
    end

    isfile(filename) && rm(filename)
end

@testset "escape" begin

    @test XLSX.escape("hello&world<'") == "hello&amp;world&lt;&apos;"
    @test XLSX.unescape("hello&amp;world&lt;&apos;") == "hello&world<'"

    esc_filename = "output_table_escape_test.xlsx"
    isfile(esc_filename) && rm(esc_filename)

    esc_col_names = ["&' & \" < > '", "I❤Julia", "\"<'&O-O&'>\"", "<&>"]
    esc_sheetname = "& & \" > < "
    esc_data = Vector{Any}(undef, 4)
    esc_data[1] = ["11&&", "12\"&", "13<&", "14>&", "15'&"]
    esc_data[2] = ["21&&&&", "22&\"&&", "23&<&&", "24&>&&", "25&'&&"]
    esc_data[3] = ["31&&&&&&", "32&&\"&&&", "33&&<&&&", "34&&>&&&", "35&&'&&&"]
    esc_data[4] = ["41& &; &&", "42\" \"; \"\"", "43< <; <<", "44> >; >>", "45' '; ''"]
    XLSX.writetable(esc_filename, esc_data, esc_col_names, overwrite=true, sheetname=esc_sheetname)

    dtable = XLSX.readtable(esc_filename, esc_sheetname)
    r1_data, r1_col_names = dtable.data, dtable.column_labels
    check_test_data(r1_data, esc_data)
    @test r1_col_names[4] == Symbol(esc_col_names[4])
    @test r1_col_names[3] == Symbol(esc_col_names[3])
    @test r1_col_names[2] == Symbol(esc_col_names[2])
    @test r1_col_names[1] == Symbol(esc_col_names[1])
    isfile(esc_filename) && rm(esc_filename)

    # compare to the backup version: escape.xlsx
    dtable = XLSX.readtable(joinpath(data_directory, "escape.xlsx"), esc_sheetname)
    r2_data, r2_col_names = dtable.data, dtable.column_labels
    check_test_data(r2_data, esc_data)
    check_test_data(r2_data, r1_data)
    @test string(r2_col_names[4]) == esc_col_names[4]
    @test string(r2_col_names[3]) == esc_col_names[3]
    @test string(r2_col_names[2]) == esc_col_names[2]
    @test string(r2_col_names[1]) == esc_col_names[1]

    esc_col_names = ["&; &amp; &quot; &lt; &gt; &apos; ", "I❤Julia", "\"<'&O-O&'>\"", "<&>"]
    esc_sheetname = string( esc_col_names[1],esc_col_names[2],esc_col_names[3],esc_col_names[4])[1:30] # There is a hard limit in Excel
    esc_data = Vector{Any}(undef, 4)
    esc_data[1] = ["11&amp;&",    "12&quot;&",    "13&lt;&",    "14&gt;&",    "15&apos;&"    ]
    esc_data[2] = ["21&&amp;&&",  "22&&quot;&&",  "23&&lt;&&",  "24&&gt;&&",  "25&&apos;&&"  ]
    esc_data[3] = ["31&&&amp;&&&","32&&&quot;&&&","33&&&lt;&&&","34&&&gt;&&&","35&&&apos;&&&"]
    esc_data[4] = ["41& &; &&",   "42\" \"; \"\"","43< <; <<",  "44> >; >>",  "45' '; ''"    ]
    XLSX.writetable(esc_filename, esc_data, esc_col_names, overwrite=true, sheetname=esc_sheetname)

    dtable = XLSX.readtable(esc_filename, esc_sheetname)
    r3_data, r3_col_names = dtable.data, dtable.column_labels
    check_test_data(r3_data, esc_data)
    @test r3_col_names[4] == Symbol( esc_col_names[4] )
    @test r3_col_names[3] == Symbol( esc_col_names[3] )
    @test r3_col_names[2] == Symbol( esc_col_names[2] )
    @test r3_col_names[1] == Symbol( esc_col_names[1] )
    isfile(esc_filename) && rm(esc_filename)

    # compare to the backup version: escape2.xlsx
    dtable = XLSX.readtable(joinpath(data_directory, "escape2.xlsx"), esc_sheetname)
    r4_data, r4_col_names = dtable.data, dtable.column_labels
    check_test_data(r4_data, esc_data)
    check_test_data(r4_data, r3_data)
    @test r4_col_names[4] == Symbol( esc_col_names[4] )
    @test r4_col_names[3] == Symbol( esc_col_names[3] )
    @test r4_col_names[2] == Symbol( esc_col_names[2] )
    @test r4_col_names[1] == Symbol( esc_col_names[1] )

end

# issue #67
@testset "row_index" begin
    filename = "test_pr67.xlsx"
    XLSX.openxlsx(filename, mode="w") do xf
        xf[1]["A2"] = 5
        xf[1]["A1"] = 7
    end
    @test isfile(filename)
    isfile(filename) && rm(filename)
end

@testset "show xlsx" begin
    @testset "single sheet" begin
        xf = XLSX.readxlsx(joinpath(src_data_directory, "blank.xlsx"))
        show(IOBuffer(), xf)
    end

    @testset "multiple sheets" begin
        xf = XLSX.readxlsx(joinpath(data_directory, "Book1.xlsx"))
        show(IOBuffer(), xf)
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
end

# issues #380, #362, #267, #170
@testset "no default namespace" begin

    f = XLSX.openxlsx(joinpath(data_directory, "No-Default_NameSpace.xlsx"), mode="rw")
    @test XLSX.get_dimension(f[2])==XLSX.CellRange("A1:N15")
    XLSX.addDefinedName(f, "xfile", "XLSX-Export!B2")
    XLSX.addDefinedName(f[2], "wsheet", 200.2)
    XLSX.mergeCells(f[2], "A11:N11")
    XLSX.setRowHeight(f[2], 11; height=25)
    XLSX.setColumnWidth(f[2], "B"; width=50)
    @test XLSX.setFont(f[2], [2, 4], 2:4; size=18, name="Arial") == -1
    @test XLSX.setBorder(f[2], [2, 4], :; left=["style" => "hair"], right=["color" => "FF8B4513"], top=["style" => "hair"], bottom=["color" => "chocolate4"], diagonal=["style" => "hair"]) == -1
    @test XLSX.setFill(f[2], "G5"; pattern="darkTrellis", fgColor="FF222222", bgColor="FFDDDDDD") == 2
    @test XLSX.setConditionalFormat(f[2], "J2:J10", :colorScale) == 0
    XLSX.mergeCells(f[2], "A4:C5")
    @test XLSX.setConditionalFormat(f[2], "H2:H15", :dataBar) == 0
    XLSX.setFormula(f[2], "A16:N16", "=sum(A2:A15)")
    XLSX.addImage(f[2], "B17", joinpath(data_directory, "track_start.jpg"))
    XLSX.copysheet!(f[2], "newSheet")

    XLSX.writexlsx("mytest.xlsx", f, overwrite=true)

    f2 = XLSX.openxlsx("mytest.xlsx", mode="rw")
    @test XLSX.get_dimension(f2[2])==XLSX.CellRange("A1:N16")
    @test f2["xfile"] == f2["XLSX-Export"]["B2"]
    @test f2["XLSX-Export"]["wsheet"] == 200.2
    @test XLSX.getMergedCells(f2[2]) == XLSX.CellRange[XLSX.CellRange("A11:N11"), XLSX.CellRange("A4:C5")]
    @test XLSX.isMergedCell(f2[2], "D11")
    @test XLSX.getMergedBaseCell(f2[2], "D11") == (baseCell=XLSX.CellRef("A11"), baseValue=13953)
    @test ismissing(f2[2]["D11"])
    @test XLSX.getRowHeight(f2[2], "B11") ≈ 25.2109375
    @test XLSX.getColumnWidth(f2[2], 11, 2) ≈ 50.7109375
    @test XLSX.getFont(f2[2], 2, 2).font == Dict("name" => Dict("val" => "Arial"), "sz" => Dict("val" => "18"), "color" => Dict("theme" => "1"))
    @test XLSX.getBorder(f2[2], 4, 6).border == Dict("left" => Dict("style" => "hair"), "bottom" => nothing, "right" => nothing, "top" => Dict("style" => "hair"), "diagonal" => Dict("style" => "hair", "direction" => "both"))
    @test XLSX.getFill(f2[2], "G5").fill == Dict("patternFill" => Dict("bgrgb" => "FFDDDDDD", "patternType" => "darkTrellis", "fgrgb" => "FF222222"))
    @test sort!(XLSX.getConditionalFormats(f2[2]), by = x -> x.second.priority, rev=true) == [
        XLSX.CellRange("H2:H15") => (type = "dataBar", priority = 2),
        XLSX.CellRange("J2:J10") => (type = "colorScale", priority = 1)
    ]
    @test XLSX.getFormula(f2[2], "B16") == "=sum(B2:B15)"
    @test XLSX.getFormula(f2[2], "M16") == "=sum(M2:M15)"

    @test XLSX.hassheet(f, "newSheet")
    @test XLSX.getImages(f["newSheet"]) == [(sheet = "newSheet", media_name = "image1.jpg", from = "B17", to = "E27")]

    for row = 1:16
        for col = 1:14
            @test (ismissing(f[1][row, col]) && ismissing(f2[1][row, col])) || (f2[1][row, col] == f[1][row, col])
            @test (ismissing(f[2][row, col]) && ismissing(f[3][row, col])) || (f[2][row, col] == f[3][row, col])
            @test (ismissing(f[3][row, col]) && ismissing(f2[3][row, col])) || (f2[3][row, col] == f[3][row, col])
        end
    end

    df = XLSX.readto(joinpath(data_directory, "No-Default_NameSpace.xlsx"), 2, DataFrames.DataFrame)
    @test DataFrames.names(df) == String[
        "NR ",
        "LC ",
        "LC-title ",
        "X [m]",
        "Xi ",
        "MNR ",
        "SIGU [MPa]",
        "DL [m]",
        "node1 ",
        "X [m]_2",
        "Y [m]",
        "Z [m]",
        "node2 ",
        "NREF "
    ]
    @test DataFrames.nrow(df) == 14

    XLSX.writetable("mytest.xlsx", "Sheet1" => df; overwrite=true)
    df2 = XLSX.readto("mytest.xlsx", DataFrames.DataFrame)

    @test DataFrames.names(df) == DataFrames.names(df2)
    @test DataFrames.nrow(df) == DataFrames.nrow(df2)
    @test isequal(df, df2)


    f = XLSX.openxlsx(joinpath(data_directory, "No-Default_NameSpace2.xlsx"), mode="rw")
    @test XLSX.get_dimension(f[1])==XLSX.CellRange("A1:R1001")

    XLSX.writexlsx("mytest.xlsx", f, overwrite=true)
    f2 = XLSX.readxlsx("mytest.xlsx")
    @test XLSX.get_dimension(f2[1])==XLSX.CellRange("A1:R1001")

    for row =1:100:1001
        for col = 1:2:17
            @test (ismissing(f[1][row, col]) && ismissing(f2[1][row, col])) || (f2[1][row, col] == f[1][row, col])
        end
    end

    df = XLSX.readto(joinpath(data_directory, "No-Default_NameSpace2.xlsx"), DataFrames.DataFrame; first_row=3)
    @test DataFrames.names(df) == String[
        "Stock Code",
        "Name of Securities",
        "Category",
        "Sub-Category",
        "Board Lot",
        "ISIN",
        "Expiry Date",
        "Subject to Stamp Duty",
        "Shortsell Eligible",
        "CAS Eligible",
        "VCM Eligible",
        "Admitted to CCASS",
        "Debt Securities Board Lot (Nominal)",
        "Debt Securities Investor Type",
        "POS Eligible",
        "Spread Table\r\n1 = Part A\r\n3 = Part B\r\n5 = Part D\r\n4 & 6 = Part E",
        "Trading Currency",
        "RMB Counter"
        ]
    @test DataFrames.nrow(df) == 998

    XLSX.writetable("mytest.xlsx", "Sheet1" => df; overwrite=true)
    df2 = XLSX.readto("mytest.xlsx", DataFrames.DataFrame)

    @test DataFrames.names(df) == DataFrames.names(df2)
    @test DataFrames.nrow(df) == DataFrames.nrow(df2)

    for row in 1:5:DataFrames.nrow(df)
        for col in 1:2:DataFrames.ncol(df)
            if !(ismissing(df2[row, col]))
                @test df[row, col] == df2[row, col]
            end
        end
    end
    isfile("mytest.xlsx") && rm("mytest.xlsx")
end

# issues #62, #71
@testset "windows compatibility" begin
    xf = XLSX.open_xlsx_template(joinpath(data_directory, "issue62_71.xlsx"))
    @test xf["Sheet1"]["A1"] == "One"
    @test xf["Sheet1"]["A2"] == 1

    @test collect(keys(xf.binary_data)) == ["xl/printerSettings/printerSettings1.bin"]
end

# issue #117
@testset "whitespace nodes" begin
    xf = XLSX.readxlsx(joinpath(data_directory, "noutput_first_second_third.xlsx"))
    @test XLSX.sheetnames(xf) == ["NOTES", "DATA"]
    @test xf["NOTES"]["A1"] == "Nominal GNP/GDP"
    @test xf["NOTES"]["A9"] == "Last updated on: August 29, 2019"
    @test xf["DATA"]["A5"] == "Date"
    @test xf["DATA"]["A6"] == "1965:Q3"
    @test xf["DATA"]["B6"] ≈ 6.7731
    @test xf["DATA"]["E5"] == "Most_Recent"
    @test xf["DATA"]["E7"] ≈ 12.6215
end

# issue #303
@testset "xml:space" begin
    f = XLSX.openxlsx(joinpath(data_directory, "sstTest.xlsx"), mode="rw")
    s = f[1]
    @test XLSX.getdata(s, :) == ["  hello" "    "; "  hello  " "    "; " hello\">" "    "; "hello\">" "    "; "  hello" "    "]
    s["C1"] = " "
    s["C2"] = " hello"
    s["C3"] = "hello "
    s["C4"] = " hello "
    s["C5"] = " \"hello\" "
    @test XLSX.getdata(s, "C1:C5") == Any[" "; " hello"; "hello "; " hello "; " \"hello\" ";;]
    XLSX.writexlsx("mydata.xlsx", f, overwrite=true)
    @test XLSX.readdata("mydata.xlsx", 1, :) == ["  hello" "    " " "; "  hello  " "    " " hello"; " hello\">" "    " "hello "; "hello\">" "    " " hello "; "  hello" "    " " \"hello\" "]
    XLSX.writetable("mydata.xlsx", [["  hello", "  hello  ", " hello\">", "hello\">", "  hello"], ["    ", "    ", "    ", "    ", "    "], [" ", " hello", "hello ", " hello ", " \"hello\" "]], ["Col_A", "Col_B", "Col_C"]; overwrite=true)
    @test XLSX.readdata("mydata.xlsx", 1, :) == ["Col_A" "Col_B" "Col_C"; "  hello" "    " " "; "  hello  " "    " " hello"; " hello\">" "    " "hello "; "hello\">" "    " " hello "; "  hello" "    " " \"hello\" "]
    isfile("mydata.xlsx") && rm("mydata.xlsx")
end

# issue #243
@testset "xml bom" begin
    xf = XLSX.readxlsx(joinpath(data_directory, "Bom - issue243.xlsx"))
    @test XLSX.sheetnames(xf) == ["QMJ Factors", "Definition", "Data Sources", "--> Additional Global Factors", "Disclosures"]
    @test XLSX.sheetcount(xf) == 5
    @test XLSX.hassheet(xf, "QMJ Factors") == true
    @test xf["QMJ Factors"]["H833"] ≈ -0.0686846616503713
end

@testset "inlineStr" begin
    xf = XLSX.readxlsx(joinpath(data_directory, "inlinestr.xlsx"))
    sheet = xf["Requirements"]
    @test sheet["A1"] == "NN"
    @test sheet["A2"] == 1
    @test sheet["B1"] == "Hierarchy"
    @test sheet["B2"] == "+"
    @test ismissing(sheet["C1"])
    @test ismissing(sheet["C2"])
    @test sheet["D1"] == "Outline Number"
    @test sheet["D2"] == "1."
    @test sheet["E1"] == "ID"
    @test sheet["E2"] == "RQ11610"
    @test sheet["F1"] == "Name"
    @test sheet["F2"] == "requirement"
    @test sheet["G1"] == "Type"
    @test sheet["G2"] == "Textual Requirement"
    @test sheet["H1"] == "Description"
    @test sheet["H2"] == "test"
    @test ismissing(sheet["I1"])
    @test ismissing(sheet["I2"])
    @test ismissing(sheet["J1"])
    @test ismissing(sheet["J2"])

    xf = XLSX.openxlsx(joinpath(data_directory, "inlinestr.xlsx"), mode="rw")
    XLSX.writexlsx("mydata.xlsx", xf, overwrite=true)
    xf = XLSX.readxlsx("mydata.xlsx")
    sheet = xf["Requirements"]
    @test sheet["A1"] == "NN"
    @test sheet["A2"] == 1
    @test sheet["B1"] == "Hierarchy"
    @test sheet["B2"] == "+"
    @test ismissing(sheet["C1"])
    @test ismissing(sheet["C2"])
    @test sheet["D1"] == "Outline Number"
    @test sheet["D2"] == "1."
    @test sheet["E1"] == "ID"
    @test sheet["E2"] == "RQ11610"
    @test sheet["F1"] == "Name"
    @test sheet["F2"] == "requirement"
    @test sheet["G1"] == "Type"
    @test sheet["G2"] == "Textual Requirement"
    @test sheet["H1"] == "Description"
    @test sheet["H2"] == "test"
    @test ismissing(sheet["I1"])
    @test ismissing(sheet["I2"])
    @test ismissing(sheet["J1"])
    @test ismissing(sheet["J2"])
    isfile("mydata.xlsx") && rm("mydata.xlsx")
end

@testset "rich text formats" begin
    @testset "setAttributes" begin
        xf = XLSX.readxlsx(joinpath(data_directory, "is.xlsx"))
        sh = xf["Sheet1"]
        @test_throws XLSX.XLSXError XLSX.setFont(sh, "A1"; name="Palatino") # Must be in write mode for rich text formats

        xf = XLSX.opentemplate(joinpath(data_directory, "is.xlsx"))
        sh = xf["Sheet1"]
        @test XLSX.getFont(sh, "A1") === nothing
        XLSX.setFont(sh, "A1"; name="Palatino")
        r = XLSX.getRichTextString(sh, "A1")
        @test r.runs[4].text == "ki"
        @test r.runs[4].atts == Dict(:bold => true, :color => "FFFFFFFF", :size => 11)
        @test r.runs[6].text == "ty"
        @test r.runs[6].atts == Dict(:color => "FFFF0000", :size => 11)
        @test XLSX.getFont(sh, "A1").font == Dict("name" => Dict("val" => "Palatino"), "sz" => Dict("val" => "11"), "color" => Dict("theme" => "1"))
        XLSX.setFont(sh, "A1:F2"; name="Palatino")
        r = XLSX.getRichTextString(sh, "B2")
        @test r.runs[4].text == "k"
        @test r.runs[4].atts == Dict(:bold => true, :color => "FFFFFFFF", :size => 11, :under => true)
        r = XLSX.getRichTextString(sh, "F2")
        @test r.runs[3].text == "lo "
        @test r.runs[3].atts == Dict(:color => "FFFFFFFF", :size => 11)
        @test XLSX.getFont(sh, "B2").font["name"] == Dict("val" => "Palatino")
        XLSX.setFont(sh, "B"; under="none")
        @test haskey(XLSX.getFont(sh, "B2").font, "u") == false
        XLSX.setFont(sh, "C1,D2:E2", color="orange")
        r = XLSX.getRichTextString(sh, "C1")
        @test r.runs[6].text == "ty "
        @test r.runs[6].atts == Dict(:size => 11)
        r = XLSX.getRichTextString(sh, "E2")
        @test r.runs[9].text == "t"
        @test r.runs[9].atts == Dict(:size => 24)
        @test XLSX.getFont(sh, "C1").font["color"] == Dict("rgb" => XLSX.get_color("orange"))
        @test XLSX.getFont(sh, "D2").font["color"] == Dict("rgb" => XLSX.get_color("orange"))
        @test XLSX.getFont(sh, "D1").font["color"] == Dict("theme" => "1")

        xf = XLSX.opentemplate(joinpath(data_directory, "is.xlsx"))
        sh = xf["Sheet1"]
        XLSX.setUniformFont(sh, [1, 2], 1:2; under="double")
        @test XLSX.getFont(sh, "A1").font["u"] == Dict("val" => "double")
        @test XLSX.getFont(sh, "B2").font["u"] == Dict("val" => "double")
        @test XLSX.getFont(sh, "A1").font["sz"] == Dict("val" => "11")
        @test XLSX.getFont(sh, "B2").font["sz"] == Dict("val" => "11")
        @test XLSX.getFont(sh, "A1").font["color"] == Dict("theme" => "1")
        @test XLSX.getFont(sh, "B2").font["color"] == Dict("theme" => "1")
        @test XLSX.getFont(sh, "C2") === nothing
        r = XLSX.getRichTextString(sh, "C1")
        @test r.runs[3].text == "lo "
        @test r.runs[3].atts == Dict(:color => "FFFFFFFF", :name => "Aptos Narrow", :size => 11)
        @test r.runs[4].text == "ki"
        @test r.runs[4].atts == Dict(:bold => true, :color => "FFFFFFFF", :italic => true, :name => "Aptos Narrow", :size => 11)

        XLSX.setUniformFont(sh, :, 3:2:5; color="orange")
        @test XLSX.getFont(sh, "C1").font["sz"] == Dict("val" => "11")
        @test XLSX.getFont(sh, "E2").font["sz"] == Dict("val" => "11")
        @test XLSX.getFont(sh, "C1").font["color"] == Dict("rgb" => XLSX.get_color("orange"))
        @test XLSX.getFont(sh, "E2").font["color"] == Dict("rgb" => XLSX.get_color("orange"))
        @test XLSX.getFont(sh, "D2") === nothing
        r = XLSX.getRichTextString(sh, "E2")
        @test r.runs[9].text == "t"
        @test r.runs[9].atts == Dict(:name => "Aptos Narrow", :size => 24)
        @test r.runs[10].text == "y "
        @test r.runs[10].atts == Dict(:name => "Aptos Narrow", :size => 11)

        xf = XLSX.opentemplate(joinpath(data_directory, "is.xlsx"))
        sh = xf["Sheet1"]
        XLSX.setUniformFont(sh, "A1,C2,D1:F1"; under="double")
        @test XLSX.getFont(sh, "A1").font["u"] == Dict("val" => "double")
        @test XLSX.getFont(sh, "C2").font["u"] == Dict("val" => "double")
        @test XLSX.getFont(sh, "D1").font["sz"] == Dict("val" => "11")
        @test XLSX.getFont(sh, "E1").font["sz"] == Dict("val" => "11")
        @test XLSX.getFont(sh, "F1").font["color"] == Dict("theme" => "1")
        @test XLSX.getFont(sh, "C2").font["color"] == Dict("theme" => "1")
        @test XLSX.getFont(sh, "B2") === nothing
        r = XLSX.getRichTextString(sh, "F1")
        @test r.runs[3].text == "lo "
        @test r.runs[3].atts == Dict(:color => "FFFFFFFF", :name => "Calibri", :size => 11)
        @test r.runs[8].text == " "
        @test r.runs[8].atts == Dict(:color => "FFFFFFFF", :name => "Aptos Narrow", :size => 11)
        xf = XLSX.opentemplate(joinpath(data_directory, "is.xlsx"))
        sh = xf["Sheet1"]
        XLSX.setFont(sh, "A1"; under="double", color="orange", name="Palatino", size=20, bold=true, italic=true, strike=true)
        XLSX.setUniformStyle(sh, "A1:F2")
        @test XLSX.getFont(sh, "A1").font["u"] == Dict("val" => "double")
        @test XLSX.getFont(sh, "C2").font["u"] == Dict("val" => "double")
        @test XLSX.getFont(sh, "D1").font["sz"] == Dict("val" => "11")
        @test XLSX.getFont(sh, "E1").font["sz"] == Dict("val" => "20")
        @test XLSX.getFont(sh, "F1").font["color"] == Dict("rgb" => XLSX.get_color("orange"))
        @test XLSX.getFont(sh, "C2").font["color"] == Dict("rgb" => XLSX.get_color("orange"))
        @test XLSX.getFont(sh, "A1").font["name"] == Dict("val" => "Palatino")
        @test XLSX.getFont(sh, "F2").font["name"] == Dict("val" => "Palatino")
        @test haskey(XLSX.getFont(sh, "B2").font, "b") == true
        @test haskey(XLSX.getFont(sh, "B2").font, "i") == true
        @test haskey(XLSX.getFont(sh, "B2").font, "strike") == true
    end

    @static if VERSION >= v"1.11-"
        @testset "styledStrings" begin
            f=XLSX.newxlsx()
            s=f[1]
            s["A1"] = styled"{yellow:hello} {blue:there}"
            s["A2"] = styled"The {bold:{italic:quick {(foreground=#cd853f):brown} fox} jumps over the {(foreground=#FFC000):lazy} dog}"
            s["A3"] = styled"In terms of color, we have named faces for the 16 standard terminal colors:\n {black:■} {red:■} {green:■} {yellow:■} {blue:■} {magenta:■} {cyan:■} {white:■}\n {bright_black:■} {bright_red:■} {bright_green:■} {bright_yellow:■} {bright_blue:■} {bright_magenta:■} {bright_cyan:■} {bright_white:■}"
            s["A4"] = styled"\nThe basic font-style attributes are {bold:bold}, {light:light}, {italic:italic},\n {underline:underline}, and {strikethrough:strikethrough}"
            s["A5"] = styled"{red:{(height=2.0):{bold:Hello}} {(height=120):Tim!}}"
            StyledStrings.addface!(:MyFace => StyledStrings.Face(foreground = 0xFF7700))
            s["A6"] = styled"{MyFace:this is orange text}"
            s["A7"] = styled"{(weight=bold):{(underline=true, slant=italic):deleted {(strikethrough=true):and more {(height=2.5):too}}, aswell}}. And now unbolded!"
            s["A8"] = styled"Hello {(font=Consolas):computer - do {bright_green:you have {(height=2.4):any} {(font=Palatino, foreground=#cd853f):rhymes}} for} me"
            @test s["A1"] == "hello there"
            @test s["A2"] == "The quick brown fox jumps over the lazy dog"
            @test s["A3"] == "In terms of color, we have named faces for the 16 standard terminal colors:\n ■ ■ ■ ■ ■ ■ ■ ■\n ■ ■ ■ ■ ■ ■ ■ ■"
            @test s["A4"] == "\nThe basic font-style attributes are bold, light, italic,\n underline, and strikethrough" # in Excel, light is simply plain text (not bold).
            @test s["A5"] == "Hello Tim!"
            @test s["A6"] == "this is orange text"
            @test s["A7"] == "deleted and more too, aswell. And now unbolded!"
            @test s["A8"] == "Hello computer - do you have any rhymes for me"
            @test XLSX.get_workbook(s).sst.shared_strings[Int(XLSX.getcell(s, "A1").value)+1] == "<si><r><rPr><color rgb=\"FFE5A509\"/><sz val=\"12\"/></rPr><t>hello</t></r><r><rPr><sz val=\"12\"/></rPr><t xml:space=\"preserve\"> </t></r><r><rPr><color rgb=\"FF195EB3\"/><sz val=\"12\"/></rPr><t>there</t></r></si>"
            @test XLSX.get_workbook(s).sst.shared_strings[Int(XLSX.getcell(s, "A2").value)+1] == "<si><r><rPr><sz val=\"12\"/></rPr><t xml:space=\"preserve\">The </t></r><r><rPr><b/><i/><sz val=\"12\"/></rPr><t xml:space=\"preserve\">quick </t></r><r><rPr><b/><i/><color rgb=\"FFCD853F\"/><sz val=\"12\"/></rPr><t>brown</t></r><r><rPr><b/><i/><sz val=\"12\"/></rPr><t xml:space=\"preserve\"> fox</t></r><r><rPr><b/><sz val=\"12\"/></rPr><t xml:space=\"preserve\"> jumps over the </t></r><r><rPr><b/><color rgb=\"FFFFC000\"/><sz val=\"12\"/></rPr><t>lazy</t></r><r><rPr><b/><sz val=\"12\"/></rPr><t xml:space=\"preserve\"> dog</t></r></si>"
            @test XLSX.get_workbook(s).sst.shared_strings[Int(XLSX.getcell(s, "A3").value)+1] == "<si><r><rPr><sz val=\"12\"/></rPr><t xml:space=\"preserve\">In terms of color, we have named faces for the 16 standard terminal colors:\n </t></r><r><rPr><color rgb=\"FF1C1A23\"/><sz val=\"12\"/></rPr><t>■</t></r><r><rPr><sz val=\"12\"/></rPr><t xml:space=\"preserve\"> </t></r><r><rPr><color rgb=\"FFA51C2C\"/><sz val=\"12\"/></rPr><t>■</t></r><r><rPr><sz val=\"12\"/></rPr><t xml:space=\"preserve\"> </t></r><r><rPr><color rgb=\"FF25A268\"/><sz val=\"12\"/></rPr><t>■</t></r><r><rPr><sz val=\"12\"/></rPr><t xml:space=\"preserve\"> </t></r><r><rPr><color rgb=\"FFE5A509\"/><sz val=\"12\"/></rPr><t>■</t></r><r><rPr><sz val=\"12\"/></rPr><t xml:space=\"preserve\"> </t></r><r><rPr><color rgb=\"FF195EB3\"/><sz val=\"12\"/></rPr><t>■</t></r><r><rPr><sz val=\"12\"/></rPr><t xml:space=\"preserve\"> </t></r><r><rPr><color rgb=\"FF803D9B\"/><sz val=\"12\"/></rPr><t>■</t></r><r><rPr><sz val=\"12\"/></rPr><t xml:space=\"preserve\"> </t></r><r><rPr><color rgb=\"FF0097A7\"/><sz val=\"12\"/></rPr><t>■</t></r><r><rPr><sz val=\"12\"/></rPr><t xml:space=\"preserve\"> </t></r><r><rPr><color rgb=\"FFDDDCD9\"/><sz val=\"12\"/></rPr><t>■</t></r><r><rPr><sz val=\"12\"/></rPr><t xml:space=\"preserve\">\n </t></r><r><rPr><color rgb=\"FF76757A\"/><sz val=\"12\"/></rPr><t>■</t></r><r><rPr><sz val=\"12\"/></rPr><t xml:space=\"preserve\"> </t></r><r><rPr><color rgb=\"FFED333B\"/><sz val=\"12\"/></rPr><t>■</t></r><r><rPr><sz val=\"12\"/></rPr><t xml:space=\"preserve\"> </t></r><r><rPr><color rgb=\"FF33D079\"/><sz val=\"12\"/></rPr><t>■</t></r><r><rPr><sz val=\"12\"/></rPr><t xml:space=\"preserve\"> </t></r><r><rPr><color rgb=\"FFF6D22C\"/><sz val=\"12\"/></rPr><t>■</t></r><r><rPr><sz val=\"12\"/></rPr><t xml:space=\"preserve\"> </t></r><r><rPr><color rgb=\"FF3583E4\"/><sz val=\"12\"/></rPr><t>■</t></r><r><rPr><sz val=\"12\"/></rPr><t xml:space=\"preserve\"> </t></r><r><rPr><color rgb=\"FFBF60CA\"/><sz val=\"12\"/></rPr><t>■</t></r><r><rPr><sz val=\"12\"/></rPr><t xml:space=\"preserve\"> </t></r><r><rPr><color rgb=\"FF26C6DA\"/><sz val=\"12\"/></rPr><t>■</t></r><r><rPr><sz val=\"12\"/></rPr><t xml:space=\"preserve\"> </t></r><r><rPr><color rgb=\"FFF6F5F4\"/><sz val=\"12\"/></rPr><t>■</t></r></si>"
            @test XLSX.get_workbook(s).sst.shared_strings[Int(XLSX.getcell(s, "A4").value)+1] == "<si><r><rPr><sz val=\"12\"/></rPr><t xml:space=\"preserve\">\nThe basic font-style attributes are </t></r><r><rPr><b/><sz val=\"12\"/></rPr><t>bold</t></r><r><rPr><sz val=\"12\"/></rPr><t xml:space=\"preserve\">, </t></r><r><rPr><sz val=\"12\"/></rPr><t>light</t></r><r><rPr><sz val=\"12\"/></rPr><t xml:space=\"preserve\">, </t></r><r><rPr><i/><sz val=\"12\"/></rPr><t>italic</t></r><r><rPr><sz val=\"12\"/></rPr><t xml:space=\"preserve\">,\n </t></r><r><rPr><sz val=\"12\"/><u/></rPr><t>underline</t></r><r><rPr><sz val=\"12\"/></rPr><t xml:space=\"preserve\">, and </t></r><r><rPr><strike/><sz val=\"12\"/></rPr><t>strikethrough</t></r></si>"
            @test XLSX.get_workbook(s).sst.shared_strings[Int(XLSX.getcell(s, "A5").value)+1] == "<si><r><rPr><b/><color rgb=\"FFA51C2C\"/><sz val=\"24\"/></rPr><t>Hello</t></r><r><rPr><color rgb=\"FFA51C2C\"/><sz val=\"12\"/></rPr><t xml:space=\"preserve\"> </t></r><r><rPr><color rgb=\"FFA51C2C\"/><sz val=\"12\"/></rPr><t>Tim!</t></r></si>"
            @test XLSX.get_workbook(s).sst.shared_strings[Int(XLSX.getcell(s, "A6").value)+1] == "<si>\n  <t>this is orange text</t>\n</si>"
            @test XLSX.getFont(s, "A6").font == Dict("name" => Dict("val" => "Calibri"), "sz" => Dict("val" => "12"), "color" => Dict("rgb" => "FFFF7700"))
            @test XLSX.get_workbook(s).sst.shared_strings[Int(XLSX.getcell(s, "A7").value)+1] == "<si><r><rPr><b/><i/><sz val=\"12\"/><u/></rPr><t xml:space=\"preserve\">deleted </t></r><r><rPr><b/><i/><strike/><sz val=\"12\"/><u/></rPr><t xml:space=\"preserve\">and more </t></r><r><rPr><b/><i/><strike/><sz val=\"30\"/><u/></rPr><t>too</t></r><r><rPr><b/><i/><sz val=\"12\"/><u/></rPr><t>, aswell</t></r><r><rPr><sz val=\"12\"/></rPr><t>. And now unbolded!</t></r></si>"
        end

    end

    @testset "RichTextString" begin
        f=XLSX.newxlsx()
        s=f[1]
        rtf1=XLSX.RichTextRun("Hello", (:vertAlign => "subscript"))
        rtf2=XLSX.RichTextRun(" Kitty ", (:color => "green", :size => 14, :bold => true, :under => true))
        rtf3=XLSX.RichTextRun("Hello", [:color => "green", :size => 14, :under => true])
        s["A1"] = XLSX.RichTextString(rtf1, rtf2, rtf3)
        @test XLSX.getRichTextString(s, "A1").runs[1].atts == Dict(:vertAlign => "subscript")
        @test XLSX.getRichTextString(s, "A1") == XLSX.RichTextString(rtf1, rtf2, rtf3)

        rtf4=XLSX.RichTextRun("Hell", (color = "red", size = 18, name = "Times New Roman"))
        rtf5=XLSX.RichTextRun("o", [:color => "green", :size => 24, :vertAlign => "superscript", :name => "Arial"])
        rtf6=XLSX.RichTextRun(" Kitt", [:color => "blue", :size => 12, :name => "Consolas"])
        rtf7=XLSX.RichTextRun("y", [:color => "green", :size => 14, :vertAlign => "subscript"])
        s["A2"] = XLSX.RichTextString(rtf4, rtf5, rtf6, rtf7)
        @test XLSX.getRichTextString(s, "A2").runs[1].atts == Dict(:color => XLSX.get_color("red"), :size => 18, :name => "Times New Roman")
        @test XLSX.getRichTextString(s, "A2") == XLSX.RichTextString(rtf4, rtf5, rtf6, rtf7)

        @test XLSX.getRichTextString(s, "A1")*XLSX.getRichTextString(s, "A2") == XLSX.RichTextString(rtf1, rtf2, rtf3, rtf4, rtf5, rtf6, rtf7)
        @test string((XLSX.getRichTextString(s, "A1")*XLSX.getRichTextString(s, "A2"))[8:24]) == "itty HelloHello K"
        @test length((XLSX.getRichTextString(s, "A1")*XLSX.getRichTextString(s, "A2"))[8:24]) == 17
        rtf1=XLSX.RichTextRun("single run", [:color => "green", :size => 14, :vertAlign => "subscript"])
        s["A3"] = XLSX.RichTextString(rtf1)

        @test s["A1"] == "Hello Kitty Hello"
        @test s["A2"] == "Hello Kitty"
        @test s["A3"] == "single run"
        @test XLSX.get_workbook(s).sst.shared_strings[Int(XLSX.getcell(s, "A1").value)+1] == "<si><r><rPr><vertAlign val=\"subscript\"/></rPr><t>Hello</t></r><r><rPr><b/><color rgb=\"FF008000\"/><sz val=\"14\"/><u/></rPr><t xml:space=\"preserve\"> Kitty </t></r><r><rPr><color rgb=\"FF008000\"/><sz val=\"14\"/><u/></rPr><t>Hello</t></r></si>"
        @test XLSX.get_workbook(s).sst.shared_strings[Int(XLSX.getcell(s, "A2").value)+1] == "<si><r><rPr><rFont val=\"Times New Roman\"/><color rgb=\"FFFF0000\"/><sz val=\"18\"/></rPr><t>Hell</t></r><r><rPr><rFont val=\"Arial\"/><color rgb=\"FF008000\"/><sz val=\"24\"/><vertAlign val=\"superscript\"/></rPr><t>o</t></r><r><rPr><rFont val=\"Consolas\"/><color rgb=\"FF0000FF\"/><sz val=\"12\"/></rPr><t xml:space=\"preserve\"> Kitt</t></r><r><rPr><color rgb=\"FF008000\"/><sz val=\"14\"/><vertAlign val=\"subscript\"/></rPr><t>y</t></r></si>"
        @test XLSX.get_workbook(s).sst.shared_strings[Int(XLSX.getcell(s, "A3").value)+1] == "<si>\n  <t>single run</t>\n</si>"
        @test XLSX.getFont(s, "A3").font == Dict("name" => Dict("val" => "Calibri"), "sz" => Dict("val" => "14"), "color" => Dict("rgb" => "FF008000"), "vertAlign" => Dict("val" => "subscript"))

        rts_expected_result = "RichTextString: \"Hello Kitty\" \n" *
                              " containing 4 runs:\n" *
                              " Run text                 Run attributes\n" *
                              " -------------------------------------------------------------------------------------------\n" *
                              " \"Hell\"                   [:color => \"FFFF0000\", :name => \"Times New Roman\", :size => 18]   \n" *
                              " \"o\"                      [:color => \"FF008000\", :name => \"Arial\", :size => 24, :vertAlign…]\n" *
                              " \" Kitt\"                  [:color => \"FF0000FF\", :name => \"Consolas\", :size => 12]          \n" *
                              " \"y\"                      [:color => \"FF008000\", :size => 14, :vertAlign => \"subscript\"]    \n"
        rtr_expected_result = "RichTextRun (\"Hell\"  [:color => \"FFFF0000\", :name => \"Times New Roman\", :size => 18])"
        io1 = IOBuffer()
        io2 = IOBuffer()
        show(io1, XLSX.getRichTextString(s, "A2"))
        @test String(take!(io1)) == rts_expected_result
        show(io1, XLSX.getRichTextString(s, "A2").runs[1])
        @test String(take!(io1)) == rtr_expected_result

        rt1 = XLSX.RichTextRun("Water is H")
        rt2 = XLSX.RichTextRun("2", :vertAlign => "subscript")
        rt3 = XLSX.RichTextRun("O!")
        rts_expected_result = "RichTextString: \"Water is H2O!\" \n" *
                               " containing 3 runs:\n" *
                               " Run text                 Run attributes\n" *
                               " -------------------------------------------------------------------------------------------\n" *
                               " \"Water is H\"             [ ]                                                               \n" *
                               " \"2\"                      [:vertAlign => \"subscript\"]                                       \n" *
                               " \"O!\"                     [ ]                                                               \n"
        io1 = IOBuffer()
        show(io1, XLSX.RichTextString(rt1, rt2, rt3))
        @test String(take!(io1)) == rts_expected_result

    end
end

# issue #299 & 301
@testset "empty_v" begin
    xf = XLSX.openxlsx(joinpath(data_directory, "empty_v.xlsx"), mode="rw")
    sheet1 = xf["Sheet1"]
    @test XLSX.getcell(sheet1, "A1") == XLSX.Cell(XLSX.get_workbook(xf), XLSX.CellRef("A1"), "str", "", "", "", true)
    @test XLSX.get_formula_from_cache(sheet1, XLSX.CellRef("A1")) == XLSX.Formula("\"\"")
    XLSX.writexlsx("mytest.xlsx", xf, overwrite=true)
    xf2 = XLSX.readxlsx("mytest.xlsx")
    sheet1 = xf2["Sheet1"]
    @test XLSX.getcell(xf2[1], "A1") == XLSX.Cell(XLSX.get_workbook(xf2), XLSX.CellRef("A1"), "str", "", "", "", true)
    @test XLSX.get_formula_from_cache(sheet1, XLSX.CellRef("A1")) == XLSX.Formula("\"\"")
    isfile("mytest.xlsx") && rm("mytest.xlsx")
end

@testset "Tables.jl integration" begin
    f = XLSX.readxlsx(joinpath(data_directory, "general.xlsx"))
    s = f["table"]
    ct = XLSX.eachtablerow(s) |> Tables.columntable
    @test isequal(ct, NamedTuple{(Symbol("Column B"), Symbol("Column C"), Symbol("Column D"), Symbol("Column E"), Symbol("Column F"), Symbol("Column G"))}(([1, 2, 3, 4, 5, 6, 7, 8], Union{Missing,String}["Str1", missing, "Str1", "Str1", "Str2", "Str2", "Str2", "Str2"], Date[Date(2018, 04, 21), Date(2018, 04, 22), Date(2018, 04, 23), Date(2018, 04, 24), Date(2018, 04, 25), Date(2018, 04, 26), Date(2018, 04, 27), Date(2018, 04, 28)], Union{Missing,String}[missing, missing, missing, missing, missing, "a", "b", missing], [0.2001132319106511, 0.27939873773400004, 0.09505916768351352, 0.07440230673248627, 0.82422780912015, 0.620588357787471, 0.9174151017732964, 0.6749604882690108], Missing[missing, missing, missing, missing, missing, missing, missing, missing])))
    rt = XLSX.eachtablerow(s) |> Tables.rowtable
    rt2 = [
        NamedTuple{(Symbol("Column B"), Symbol("Column C"), Symbol("Column D"), Symbol("Column E"), Symbol("Column F"), Symbol("Column G"))}((1, "Str1", Date(2018, 04, 21), missing, 0.2001132319106511, missing)),
        NamedTuple{(Symbol("Column B"), Symbol("Column C"), Symbol("Column D"), Symbol("Column E"), Symbol("Column F"), Symbol("Column G"))}((2, missing, Date(2018, 04, 22), missing, 0.27939873773400004, missing)),
        NamedTuple{(Symbol("Column B"), Symbol("Column C"), Symbol("Column D"), Symbol("Column E"), Symbol("Column F"), Symbol("Column G"))}((3, "Str1", Date(2018, 04, 23), missing, 0.09505916768351352, missing)),
        NamedTuple{(Symbol("Column B"), Symbol("Column C"), Symbol("Column D"), Symbol("Column E"), Symbol("Column F"), Symbol("Column G"))}((4, "Str1", Date(2018, 04, 24), missing, 0.07440230673248627, missing)),
        NamedTuple{(Symbol("Column B"), Symbol("Column C"), Symbol("Column D"), Symbol("Column E"), Symbol("Column F"), Symbol("Column G"))}((5, "Str2", Date(2018, 04, 25), missing, 0.82422780912015, missing)),
        NamedTuple{(Symbol("Column B"), Symbol("Column C"), Symbol("Column D"), Symbol("Column E"), Symbol("Column F"), Symbol("Column G"))}((6, "Str2", Date(2018, 04, 26), "a", 0.620588357787471, missing)),
        NamedTuple{(Symbol("Column B"), Symbol("Column C"), Symbol("Column D"), Symbol("Column E"), Symbol("Column F"), Symbol("Column G"))}((7, "Str2", Date(2018, 04, 27), "b", 0.9174151017732964, missing)),
        NamedTuple{(Symbol("Column B"), Symbol("Column C"), Symbol("Column D"), Symbol("Column E"), Symbol("Column F"), Symbol("Column G"))}((8, "Str2", Date(2018, 04, 28), missing, 0.6749604882690108, missing))
    ]
    @test isequal(rt, rt2)
    ct = XLSX.eachtablerow(f["table2"]) |> Tables.columntable
    @test length(ct) == 3
    @test length(ct[1]) == 4
    ct = XLSX.eachtablerow(f["general"]) |> Tables.columntable
    @test length(ct) == 2
    @test length(ct[1]) == 9
    ct = XLSX.eachtablerow(f["table3"]) |> Tables.columntable
    @test length(ct) == 3
    @test length(ct[1]) == 3
    ct = XLSX.eachtablerow(f["table4"]) |> Tables.columntable
    @test length(ct) == 3
    @test length(ct[1]) == 3
    ct = XLSX.eachtablerow(f["table5"]) |> Tables.columntable
    @test length(ct) == 1
    @test length(ct[1]) == 5
    ct = XLSX.eachtablerow(f["table6"]) |> Tables.columntable
    @test length(ct) == 1
    @test isempty(ct.hey)
    #    @test ct.hey == Any[missing]
    ct = XLSX.eachtablerow(f["table7"]; header=false) |> Tables.columntable
    @test length(ct) == 1
    @test length(ct[1]) == 1

    # write
    col_names = ["Integers", "Strings", "Floats", "Booleans", "Dates", "Times", "DateTimes"]
    data = Vector{Any}(undef, 7)
    data[1] = [1, 2, missing, 4]
    data[2] = ["Hey", "You", "Out", "There"]
    data[3] = [101.5, 102.5, missing, 104.5]
    data[4] = [true, false, missing, true]
    data[5] = [Date(2018, 2, 1), Date(2018, 3, 1), Date(2018, 5, 20), Date(2018, 6, 2)]
    data[6] = [Dates.Time(19, 10), Dates.Time(19, 20), Dates.Time(19, 30), Dates.Time(19, 40)]
    data[7] = [Dates.DateTime(2018, 5, 20, 19, 10), Dates.DateTime(2018, 5, 20, 19, 20), Dates.DateTime(2018, 5, 20, 19, 30), Dates.DateTime(2018, 5, 20, 19, 40)]
    table = NamedTuple{Tuple(Symbol(x) for x in col_names)}(Tuple(data))

    XLSX.writetable("output_table.xlsx", table, overwrite=true, sheetname="report", anchor_cell="B2")
    @test isfile("output_table.xlsx")

    XLSX.openxlsx("output_table2.xlsx", mode="w") do xf
        sheet = XLSX.getsheet(xf, 1)
        XLSX.renamesheet!(sheet, "report")
        XLSX.writetable!(sheet, table)
    end

    for file in ["output_table.xlsx", "output_table2.xlsx"]
        try
            f = XLSX.readxlsx(file)
            s = f["report"]
            table2 = XLSX.eachtablerow(s) |> Tables.columntable
            @test isequal(table, table2)
        finally
            isfile(file) && rm(file)
        end
    end

    # multiple tables in same file
    table2 = (a=[1, 2, 3, 4], b=["a", "b", "c", "d"])
    XLSX.writetable("output_table3.xlsx", "report1" => table, "report2" => table2)
    XLSX.writetable("output_table4.xlsx", ["report1" => table, "report2" => table2])
    for file in ["output_table4.xlsx", "output_table3.xlsx"]
        try
            f = XLSX.readxlsx(file)
            result1 = Tables.columntable(XLSX.eachtablerow(f["report1"]))
            result2 = Tables.columntable(XLSX.eachtablerow(f["report2"]))
            @test isequal(table, result1)
            @test isequal(table2, result2)
        finally
            isfile(file) && rm(file)
        end
    end

    @testset "Tables.jl with DataFrames" begin
        f = XLSX.readxlsx(joinpath(data_directory, "general.xlsx"))
        s = f["table"]
        df = XLSX.eachtablerow(s) |> DataFrames.DataFrame
        @test size(df) == (8, 6)
        @test df[!, "Column B"] == collect(1:8)
        @test df[!, "Column D"] == collect(Date(2018, 4, 21):Dates.Day(1):Date(2018, 4, 28))
        @test all(ismissing.(df[!, "Column G"]))

        file = joinpath(@__DIR__, "test_report.xlsx")

        try
            df1 = DataFrames.DataFrame(COL1=[10, 20, 30], COL2=["Fist", "Sec", "Third"])
            df2 = DataFrames.DataFrame(AA=["aa", "bb"], AB=[10.1, 10.2])
            XLSX.writetable(file, "REPORT_A" => df1, "REPORT_B" => df2, overwrite=true)
        finally
            isfile(file) && rm(file)
        end
    end

    @testset "Tables.jl as sink" begin
        f = CSV.read(joinpath(data_directory, "iris.csv"), XLSX.XLSXFile)
        @test XLSX.hassheet(f, "Sheet1") == true
        sheet = f["Sheet1"]
        @test sheet["A1"] == "sepal_length"
        @test sheet["B1"] == "sepal_width"
        @test sheet["C1"] == "petal_length"
        @test sheet["D1"] == "petal_width"
        @test sheet["E1"] == "species"
        @test sheet["A150"] ≈ 6.2
        @test sheet["B150"] ≈ 3.4
        @test sheet["C150"] ≈ 5.4
        @test sheet["D150"] ≈ 2.3
        @test sheet["E150"] == "virginica"
        XLSX.setFormula(f[1], "G1", "=GROUPBY(E1:E151,A1:D151,AVERAGE,3,1)")
        @test XLSX.get_formula_from_cache(sheet, XLSX.CellRef("G1")) == XLSX.Formula("_xlfn.GROUPBY(E1:E151,A1:D151,_xleta.AVERAGE,3,1)", "array", "G1:G1", nothing)
        f[1]["M1"] = "versicolor"
        XLSX.setFormula(f[1], "M2", "=VLOOKUP(M1,G1#,3,FALSE)")
        @test XLSX.get_formula_from_cache(sheet, XLSX.CellRef("M2")) == XLSX.Formula("=VLOOKUP(M1,_xlfn.ANCHORARRAY(G1),3,FALSE)", "array", "M2:M2", nothing)
        XLSX.setFormula(f[1], "G1", "=GROUPBY(E1:E151,A1:D151,STDEV.P,3,1)")
        @test XLSX.get_formula_from_cache(sheet, XLSX.CellRef("G1")) == XLSX.Formula("_xlfn.GROUPBY(E1:E151,A1:D151,_xleta.STDEV.P,3,1)", "array", "G1:G1", nothing)
        XLSX.setFormula(f[1], "G10", "_xlfn.GROUPBY(E1:E151,A1:D151,_xlfn.LAMBDA(_xlpm.x,AVERAGE(_xlpm.x)),3,1)"; raw=true)
        @test XLSX.get_formula_from_cache(sheet, XLSX.CellRef("G10")) == XLSX.Formula("_xlfn.GROUPBY(E1:E151,A1:D151,_xlfn.LAMBDA(_xlpm.x,AVERAGE(_xlpm.x)),3,1)", "array", "G10:G10", nothing)
    end

    @testset "type inference in `eachtablerow`" begin
        f = XLSX.readxlsx(joinpath(data_directory, "general.xlsx"))
        df = XLSX.eachtablerow(f["lookup"], "B:J") |> DataFrames.DataFrame
        @test eltype.(eachcol(df)) == [Int64, String, Int64, Any, Int64, String, Any, Int64, Int64]
    end
end

@testset "stream iterator" begin
    f = XLSX.openxlsx(joinpath(data_directory, "general.xlsx"), enable_cache=false)
    s = f["table"]
    for sheetrow in XLSX.eachrow(s)
        for column in 2:4
            cell = XLSX.getcell(sheetrow, column)
            if XLSX.row_number(cell) == 2 && XLSX.column_number(cell) == 2
                @test XLSX.getdata(s, cell) == "Column B"
            end
            if XLSX.row_number(cell) == 12 && XLSX.column_number(cell) == 2
                @test XLSX.getdata(s, cell) == "trash"
            end
        end
    end
end


@testset "Add Images" begin
    REL_IMAGE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
    jpeg  = joinpath(data_directory, "track_start.jpg")
    png   = joinpath(data_directory, "track_start.png")
    bytes = read(jpeg)

    # Helper so each testset gets a fresh workbook
    fresh() = (xf = XLSX.newxlsx(); (xf, xf["Sheet1"]))

    @testset "cell addressing variants" begin
        cases = [
            ((1, 1),                      "A1", nothing),
            (("B2",),                     "B2", nothing),
            (("C3:D5",),                  "C3", "D5"),
            ((CellRef("E7"),),            "E7", nothing),
            ((XLSX.CellRange("F4:H9"),), "F4", "H9"),
        ]
        for (args, exp_from, exp_to) in cases
            xf, s = fresh()
            info = XLSX.addImage(s, args..., png)
            @test info.from == exp_from
            exp_to === nothing || @test info.to == exp_to
            @test haskey(xf.binary_data, "xl/media/" * info.media_name)
        end
    end
    @testset "IOBuffer input" begin
        xf, s = fresh()
        info = XLSX.addImage(s, 3, 4, IOBuffer(copy(bytes)))
        @test startswith(info.media_name, "image")
        @test xf.binary_data["xl/media/" * info.media_name] == bytes
        @test haskey(xf.data, "xl/drawings/drawing1.xml")
    end

    @testset "drawing XML and relationships" begin
        xf, s = fresh()
        info = XLSX.addImage(s, "B2", jpeg; size=(128, 128))

        # Relationships
        rels_root = xf.data["xl/drawings/_rels/drawing1.xml.rels"][end]
        rel_nodes = [n for n in XML.children(rels_root) if XML.tag(n) == "Relationship"]
        @test !isempty(rel_nodes)
        @test any(get(XML.attributes(n), "Type", "") == REL_IMAGE for n in rel_nodes)

        # Anchor geometry
        drawing_root = xf.data["xl/drawings/drawing1.xml"][end]
        anchors = [n for n in XML.children(drawing_root) if XML.tag(n) == "xdr:twoCellAnchor"]
        @test length(anchors) == 1
        @test XLSX._parse_cell_marker(anchors[1], "from"; is_to=false) == info.from
        @test XLSX._parse_cell_marker(anchors[1], "to";   is_to=true)  == info.to
    end

    @testset "multiple images" begin
        xf, s = fresh()
        info1 = XLSX.addImage(s, 1, 1, jpeg)
        info2 = XLSX.addImage(s, 5, 5, jpeg)

        @test info1.media_name != info2.media_name
        @test haskey(xf.binary_data, "xl/media/" * info1.media_name)
        @test haskey(xf.binary_data, "xl/media/" * info2.media_name)

        rels_root = xf.data["xl/drawings/_rels/drawing1.xml.rels"][end]
        rel_nodes = [n for n in XML.children(rels_root) if XML.tag(n) == "Relationship"]
        @test length(rel_nodes) == 2
        @test all(get(XML.attributes(n), "Type", "") == REL_IMAGE for n in rel_nodes)

        drawing_root = xf.data["xl/drawings/drawing1.xml"][end]
        anchors = [n for n in XML.children(drawing_root) if XML.tag(n) == "xdr:twoCellAnchor"]
        @test length(anchors) == 2
        @test XLSX._parse_cell_marker(anchors[1], "from"; is_to=false) == info1.from
        @test XLSX._parse_cell_marker(anchors[2], "from"; is_to=false) == info2.from
    end

    @testset "round-trip (file and IOBuffer)" begin
        for (label, src) in [("file path", jpeg), ("IOBuffer", IOBuffer(copy(bytes)))]
            xf, s = fresh()
            XLSX.addImage(s, 1, 1, src)
            tmp = tempname() * ".xlsx"
            XLSX.writexlsx(tmp, xf)
            @test isfile(tmp) && filesize(tmp) > 0

            xf2  = XLSX.readxlsx(tmp)
            imgs = XLSX.getImages(xf2)
            @test length(imgs) == 1
            @test imgs[1].sheet == "Sheet1"
            @test startswith(imgs[1].media_name, "image")
        end
    end

    @testset "invalid cell reference" begin
        xf, s = fresh()
        @test_throws ArgumentError XLSX.addImage(s, "ZZZ9999", jpeg)
    end

    @testset "image cleaned up when sheet deleted" begin
        xf = XLSX.newxlsx()
        s1 = xf["Sheet1"]
        XLSX.addImage(s1, 1, 1, png)
        info = XLSX.getImages(s1)[1]

        wb = XLSX.get_workbook(xf)
        XLSX.addsheet!(wb, "Sheet2")  # need a second sheet to allow deletion
        XLSX.deletesheet!(wb, "Sheet1")

        # Media removed
        @test !haskey(xf.binary_data, "xl/media/" * info.media_name)

        # Drawing XML and rels removed
        @test !haskey(xf.data, "xl/drawings/drawing1.xml")
        @test !haskey(xf.data, "xl/drawings/_rels/drawing1.xml.rels")

        # No images reported
        @test isempty(XLSX.getImages(xf))
    end

    @testset "shared media preserved when only one sheet deleted" begin
        xf = XLSX.newxlsx()
        s1 = xf["Sheet1"]
        wb = XLSX.get_workbook(xf)

        XLSX.addImage(s1, 1, 1, jpeg)

        XLSX.copysheet!(s1, "Sheet2")
        s2 = xf["Sheet2"]

        info1 = XLSX.getImages(s1)[1]
        info2 = XLSX.getImages(s2)[1]

        # Both sheets reference the same media file
        @test info1.media_name == info2.media_name
        media_key = "xl/media/" * info1.media_name

        XLSX.deletesheet!(wb, "Sheet1")

        # Media still present — Sheet2 still references it
        @test haskey(xf.binary_data, media_key)

        # Sheet2 image still retrievable
        imgs = XLSX.getImages(xf)
        @test length(imgs) == 1
        @test imgs[1].sheet == "Sheet2"
    end
end