@testset "Shifting references" begin
    E(axis, at, count; sheet="Data") = XLSX.RowColEdit(sheet, axis, at, count)
    ins_rows(at, n=1) = E(:row, at, n)
    del_rows(at, n=1) = E(:row, at, -n)
    ins_cols(at, n=1) = E(:column, at, n)
    del_cols(at, n=1) = E(:column, at, -n)
    F(f, e; own="Data") = XLSX._shift_formula(f, e; own_sheet=own)
    R(s, e) = (r = XLSX._shift_cellrange(XLSX.CellRange(s), e); isnothing(r) ? nothing : string(r))
    C(s, e) = (r = XLSX._shift_cellref(XLSX.CellRef(s), e); isnothing(r) ? nothing : string(r))
    MAXR, MAXC = XLSX.EXCEL_MAX_ROWS, XLSX.EXCEL_MAX_COLS

    @testset "edits" begin
        @test_throws XLSX.XLSXError E(:diagonal, 1, 1)
        @test_throws XLSX.XLSXError E(:row, 0, 1)
        @test_throws XLSX.XLSXError E(:row, MAXR + 1, 1)
        @test_throws XLSX.XLSXError E(:column, MAXC + 1, 1)
        @test_throws XLSX.XLSXError E(:row, 1, 0)
        @test_throws XLSX.XLSXError E(:row, MAXR, -2)                 # past the last row
        @test E(:row, MAXR, -1).count == -1
    end

    @testset "cells" begin
        @test C("B4", ins_rows(5)) == "B4"                           # before the insert: unmoved
        @test C("B5", ins_rows(5)) == "B6"                           # at it: moves down
        @test C("B5", ins_rows(5, 3)) == "B8"
        @test C("B5", ins_cols(2)) == "C5"
        @test C("B5", ins_cols(3)) == "B5"
        @test C("B5", del_rows(5)) === nothing                       # deleted
        @test C("B9", del_rows(5, 3)) == "B6"
        @test C("B4", del_rows(5, 3)) == "B4"
        @test C("D5", del_cols(2, 2)) == "B5"
        @test C("C5", del_cols(2, 2)) === nothing
        @test C("A$MAXR", ins_rows(1)) === nothing                   # pushed off the sheet
        @test C("XFD1", ins_cols(1)) === nothing
    end

    @testset "ranges" begin
        @test R("B2:D9", ins_rows(5, 2)) == "B2:D11"                  # an insert inside grows it
        @test R("B2:D9", ins_rows(2)) == "B3:D10"                     # at its first row: moves
        @test R("B2:D9", ins_rows(10)) == "B2:D9"                     # just after it: unchanged
        @test R("B2:D9", ins_rows(9)) == "B2:D10"                     # at its last row: grows
        @test R("B2:D9", del_rows(4, 2)) == "B2:D7"                   # a delete inside shrinks it
        @test R("B2:D9", del_rows(1, 3)) == "B1:D6"                   # cut at the top
        @test R("B2:D9", del_rows(8, 5)) == "B2:D7"                   # cut at the bottom
        @test R("B2:D9", del_rows(2, 8)) === nothing                  # all of it
        @test R("B2:D9", del_rows(1, 20)) === nothing
        @test R("B2:D9", ins_cols(3)) == "B2:E9"
        @test R("B2:D9", del_cols(1)) == "A2:C9"
        @test R("B2:D9", del_cols(2, 3)) === nothing
        @test R("A1:A$MAXR", ins_rows(5)) == "A1:A$MAXR"            # the end stops at the last row
        @test R("A$(MAXR-1):A$MAXR", ins_rows(1, 5)) === nothing     # all pushed off
    end

    @testset "sqref" begin
        @test XLSX._shift_sqref("A1 B3:C9 E12", ins_rows(3)) == "A1 B4:C10 E13"
        @test XLSX._shift_sqref("A1 B3:C9 E12", del_rows(12)) == "A1 B3:C9"
        @test XLSX._shift_sqref("B3", del_rows(3)) === nothing
        @test XLSX._shift_sqref("A:A 3:3", ins_rows(1)) == "A:A 4:4"   # whole column unmoved, whole row moved
        @test XLSX._shift_sqref("A:A 3:3", ins_cols(1)) == "B:B 3:3"
    end

    @testset "formulas" begin
        @test F("=SUM(A1:A10)", ins_rows(5)) == "=SUM(A1:A11)"
        @test F("SUM(A1:A10)", ins_rows(5)) == "SUM(A1:A11)"             # stored without `=`
        @test F("=\$B\$5*B5+\$B5+B\$5", ins_rows(2)) == "=\$B\$6*B6+\$B6+B\$6"   # absolute or not
        @test F("=B5*2", ins_cols(1)) == "=C5*2"
        @test F("=B5+C7", del_rows(5)) == "=#REF!+C6"
        @test F("=SUM(B5:B5)", del_rows(5)) == "=SUM(#REF!)"
        @test F("=SUM(A:A)+SUM(3:5)", ins_rows(1)) == "=SUM(A:A)+SUM(4:6)"
        @test F("=SUM(A:C)", ins_cols(2)) == "=SUM(A:D)"
        @test F("=SUM(\$A:\$C)", del_cols(1)) == "=SUM(\$A:\$B)"
        @test F("=SUM(B:B)", del_cols(2)) == "=SUM(#REF!)"
        @test F("=A10:A1", ins_rows(5)) == "=A11:A1"                     # reversed ends kept in order

        # which sheet
        @test F("=Data!A5+A5", ins_rows(1); own="Other") == "=Data!A6+A5"
        @test F("=DATA!A5", ins_rows(1); own="Other") == "=DATA!A6"       # sheet names ignore case
        @test F("=Other!A5+A5", ins_rows(1)) == "=Other!A5+A6"
        @test F("=Data!A5", ins_rows(1); own=nothing) == "=Data!A6"       # a defined name
        @test F("=A5", ins_rows(1); own=nothing) == "=A5"
        @test F("='My Data'!A5+'My Data'!B2:B9", E(:row, 3, 1; sheet="My Data"); own="X") ==
              "='My Data'!A6+'My Data'!B2:B10"
        @test F("='It''s'!A5", E(:row, 1, 1; sheet="It's"); own="X") == "='It''s'!A6"
        @test F("=Data!B5", del_rows(5); own="X") == "=Data!#REF!"

        # left alone
        @test F("=\"A5\"&A5", ins_rows(1)) == "=\"A5\"&A6"               # in a string
        @test F("=\"say \"\"A5\"\"\"&A5", ins_rows(1)) == "=\"say \"\"A5\"\"\"&A6"
        @test F("=[1]Data!A5+A5", ins_rows(1)) == "=[1]Data!A5+A6"       # another workbook
        @test F("=SUM(Data:Other!A5)", ins_rows(1)) == "=SUM(Data:Other!A5)"   # 3-D
        @test F("=Table1[Col1]+A5", ins_rows(1)) == "=Table1[Col1]+A6"
        @test F("=Table1[[#Headers],[A5]]", ins_rows(1)) == "=Table1[[#Headers],[A5]]"
        @test F("=LOG10(A5)+ATAN2(1,2)", ins_rows(1)) == "=LOG10(A6)+ATAN2(1,2)"   # function names
        @test F("=INDIRECT(\"B5\")", ins_rows(1)) == "=INDIRECT(\"B5\")"
        @test F("=Sales*2", ins_rows(1)) == "=Sales*2"                    # a defined name
        @test F("=Data2024!A5", ins_rows(1)) == "=Data2024!A5"            # another sheet, whose name ends in digits
        @test F("=A5", ins_rows(1); own="data") == "=A6"
        @test F("=1+2", ins_rows(1)) == "=1+2"
        @test F("", ins_rows(1)) == ""
        @test F("=A5 — é", ins_rows(1)) == "=A6 — é"                      # non-ASCII text around
    end
end
