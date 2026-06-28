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
