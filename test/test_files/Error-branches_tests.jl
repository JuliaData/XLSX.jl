# Error and rarely taken branches: bad input, read-only files, files opened without
# the cache, and the style updates `setdata!` makes when a value replaces another
# of a different kind.

@testset "error branches" begin

    general = joinpath(data_directory, "general.xlsx")

    @testset "saving and writing" begin
        @test_throws XLSX.XLSXError XLSX.savexlsx(XLSX.newxlsx())   # blank: no file name

        outfile = "errbranch_exists.xlsx"
        XLSX.writetable(outfile, [[1, 2]], ["a"]; overwrite=true)
        @test_throws XLSX.XLSXError XLSX.writetable(outfile; Sheet1=([[1]], ["a"]))
        @test_throws XLSX.XLSXError XLSX.writetable(outfile, [("Sheet1", [[1]], ["a"])])
        SAVE_FILES && save_outfile(outfile)
        isfile(outfile) && rm(outfile)
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

    @testset "read-only and no-cache files" begin
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

    @testset "merged cells and row heights" begin
        xf = XLSX.newxlsx()
        sh = xf[1]
        sh[1:2, 1:2] = [1 2; 3 4]
        @test_throws XLSX.XLSXError XLSX.isMergedCell(sh, "Z99")
        @test_throws XLSX.XLSXError XLSX.getMergedBaseCell(sh, "Z99")
        @test isnothing(XLSX.getMergedBaseCell(sh, "A1"))            # no merged cells at all

        @test_throws XLSX.XLSXError XLSX.setRowHeight(sh, "A1"; height=-1)
        @test XLSX.setRowHeight(sh, "A1") == 0                        # no height: no-op
        @test XLSX.setRowHeight(sh, "A10"; height=20) == -1           # row has no cells
        SAVE_FILES && save_outfile(xf)
    end

    @testset "cell references and ranges" begin
        # out-of-range column and row numbers are rejected
        @test_throws XLSX.XLSXError XLSX.encode_column_number(0)
        @test_throws XLSX.XLSXError XLSX.encode_column_number(XLSX.EXCEL_MAX_COLS + 1)
        @test XLSX.encode_column_number(XLSX.EXCEL_MAX_COLS) == "XFD"
        @test_throws XLSX.XLSXError XLSX.is_valid_row_range(string(XLSX.EXCEL_MAX_ROWS + 1))
        @test XLSX.is_valid_row_range(string(XLSX.EXCEL_MAX_ROWS))
        @test_throws XLSX.XLSXError XLSX.split_cellname("A")

        @test !XLSX.is_valid_row_name("0")
        @test !XLSX.is_valid_cellrange("A0:B2") && !XLSX.is_valid_cellrange("A1:B0")
        @test !XLSX.is_valid_column_range("A:ZZZZ")
        @test !XLSX.is_valid_row_range("0:2")
        @test !XLSX.is_valid_sheet_cellrange("Sheet1!A0:B2")
        @test !XLSX.is_valid_sheet_column_range("Sheet1!A:ZZZZ")
        @test !XLSX.is_valid_sheet_row_range("Sheet1!0:2")

        @test sprint(show, XLSX.SheetRowRange("Sheet1!2:4")) == "Sheet1!2:4"
        @test_throws XLSX.XLSXError XLSX.NonContiguousRange("Sheet1!A1,Sheet2!B2")
        @test_throws XLSX.XLSXError XLSX.nCR("Sheet1", ["not a ref"])
        nc = XLSX.nCR("Sheet1", ["\$A\$1", "\$B\$1:\$C\$2"])
        @test nc.rng == [XLSX.CellRef("A1"), XLSX.CellRange("B1:C2")]
    end

    @testset "defined names" begin
        xf = XLSX.newxlsx()
        sh = xf[1]
        wb = XLSX.get_workbook(xf)
        @test_throws XLSX.XLSXError XLSX.get_defined_name_value(wb, "missing_name")
        @test_throws XLSX.XLSXError XLSX.get_defined_name_value(sh, "missing_name")
        @test !XLSX.is_valid_defined_name("")
        @test !XLSX.is_valid_defined_name("1abc")
        @test !XLSX.is_valid_defined_name("a-b")
        @test_throws XLSX.XLSXError XLSX.addDefName(xf, "", 1)

        XLSX.addDefinedName(xf, "dup", 1)
        @test_throws XLSX.XLSXError XLSX.addDefinedName(xf, "dup", 2)
        @test_throws XLSX.XLSXError XLSX.addDefinedName(xf, "empty", "")
        @test_throws XLSX.XLSXError XLSX.addDefinedName(sh, "empty", "")
        @test_throws XLSX.XLSXError XLSX.addDefinedName(xf, "nosheet", "A1")
        XLSX.addDefinedName(xf, "wb_nc", "Sheet1!A1,Sheet1!B2")
        @test XLSX.get_defined_name_value(wb, "wb_nc") isa XLSX.NonContiguousRange
        SAVE_FILES && save_outfile(xf)
    end

    @testset "cells, tables and worksheets" begin
        @test_throws XLSX.XLSXError XLSX.get_error_type("#BOGUS!")
        @test_throws XLSX.XLSXError XLSX.get_error_string(UInt64(999))

        dt = XLSX.DataTable(Any[[1, 2]], [:a])
        @test_throws XLSX.XLSXError Tables.getcolumn(dt, :nope)

        sh = XLSX.newxlsx()[1]                     # nothing written: no dimension
        @test occursin("Sheet1", sprint(show, sh))
        # was an UndefVarError (`ws` in the message)
        @test_throws XLSX.XLSXError XLSX.ordinal_sheet_number(XLSX.get_workbook(sh), "nope")
    end

    @testset "DrawingML colour resolution" begin
        wb = XLSX.get_workbook(XLSX.newxlsx())
        el(tag; kw...) = XLSX.XML.Element("a:" * tag; kw...)
        resolve(n) = XLSX.Charts.resolve_color_base(wb, n)
        @test resolve(el("sysClr"; val="windowText", lastClr="112233")) == "112233"
        @test resolve(el("sysClr"; val="window")) == "FFFFFF"
        @test resolve(el("sysClr"; val="windowText")) == "000000"
        @test resolve(el("prstClr"; val="dkBlue")) == "00008B"       # darkblue
        @test resolve(el("prstClr"; val="ltGray")) == "D3D3D3"       # lightgray
        @test resolve(el("prstClr"; val="medPurple")) == "9370DB"    # mediumpurple
        @test resolve(el("prstClr"; val="notAColour")) == "000000"
        @test resolve(el("schemeClr"; val="phClr")) == "000000"      # placeholder
        @test resolve(el("unknownClr"; val="x")) == "000000"
    end

    @testset "formatting through defined names and odd references" begin
        xf = XLSX.newxlsx()
        sh = xf[1]
        sh[1:3, 1:3] = reshape(collect(1:9), 3, 3)
        XLSX.addDefName(sh, "ws_const", 5)
        XLSX.addDefName(xf, "wb_const", 6)

        # a constant name is not a cell, whichever way it is reached
        @test_throws XLSX.XLSXError XLSX.setFont(xf, "wb_const"; bold=true)
        @test_throws XLSX.XLSXError XLSX.setFont(sh, "ws_const"; bold=true)
        @test_throws XLSX.XLSXError XLSX.setFont(sh, "wb_const"; bold=true)
        @test_throws XLSX.XLSXError XLSX.getFont(sh, "wb_const")
        @test_throws XLSX.XLSXError XLSX.getFont(sh, "A1:B2")       # a getter takes one cell
        @test_throws XLSX.XLSXError XLSX.getFont(xf, "Sheet1!Z99")  # outside the dimension

        # non-contiguous ranges, with and without the sheet name
        XLSX.setFont(sh, "A1,C1"; bold=true)
        @test XLSX.getFont(sh, "A1").font["b"] === nothing && XLSX.getFont(sh, "C1").font["b"] === nothing
        XLSX.setFont(sh, "Sheet1!A2,Sheet1!C2"; italic=true)
        @test haskey(XLSX.getFont(sh, "C2").font, "i")
        @test_throws XLSX.XLSXError XLSX.setFont(sh, XLSX.NonContiguousRange("Sheet1!A1,Sheet1!Z99"); bold=true)
        SAVE_FILES && save_outfile(xf)
    end

    @testset "borders through the XLSXFile" begin
        xf = XLSX.newxlsx()
        sh = xf[1]
        sh[1:2, 1:2] = [1 2; 3 4]
        XLSX.setBorder(xf, "Sheet1!A1:B2"; outside=["style" => "thin", "color" => "black"])
        @test XLSX.getBorder(sh, "A1").border["top"]["style"] == "thin"
        SAVE_FILES && save_outfile(xf)
    end
end
