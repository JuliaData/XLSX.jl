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
        SAVE_FILES && save_outfile("formulas.xlsx")

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
        XLSX.setFormula(s2,:,:,"=mySheet!A\$1+mySheet!A1")
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
        SAVE_FILES && save_outfile("formulas.xlsx")

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
        SAVE_FILES && save_outfile("formulas.xlsx")

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
        @test XLSX.get_formula_from_cache(s, XLSX.CellRef("A3")) == XLSX.ReferencedFormula("SUM(O3:S3)", 4, "A3:A10", nothing)
        @test XLSX.get_formula_from_cache(s, XLSX.CellRef("A4")) == XLSX.FormulaReference(4, nothing)

        s2 = f[2]
        @test XLSX.getcell(s2, "A1") == XLSX.Cell(XLSX.get_workbook(f), XLSX.CellRef("A1"), "", "", "54", "", true)
        @test XLSX.get_formula_from_cache(s2, XLSX.CellRef("A1")) == XLSX.Formula("SECOND(NOW())", nothing, nothing, Dict("ca" => "1"))
        @test XLSX.getcell(s2, "A2") == XLSX.Cell(XLSX.get_workbook(f), XLSX.CellRef("A2"), "", "", "54", "", true)
        @test XLSX.get_formula_from_cache(s2, XLSX.CellRef("A2")) == XLSX.ReferencedFormula("SECOND(NOW())", 1, "A2:A5", Dict("ca" => "1"))
        s2["A2"] = 3
        @test XLSX.getcell(s2, "A2") == XLSX.Cell(XLSX.get_workbook(f), XLSX.CellRef("A2"), "", "", "3", "", false)
        @test (XLSX.get_formula_from_cache(s2, XLSX.CellRef("A3"))).id == 2
        @test XLSX.get_formula_from_cache(s2, XLSX.CellRef("A3")).unhandled == Dict("ca" => "1")
        @test XLSX.getcell(s2, "A3").formula == true
        @test XLSX.getcell(s2, "A3") == XLSX.Cell(XLSX.get_workbook(f), XLSX.CellRef("A3"), "", "", "54", "", true)
        @test XLSX.get_formula_from_cache(s2, XLSX.CellRef("A3")) == XLSX.ReferencedFormula("SECOND(NOW())", 2, "A3:A5", Dict("ca" => "1"))
        @test XLSX.getcell(s2, "B1") == XLSX.Cell(XLSX.get_workbook(f), XLSX.CellRef("B1"), "", "", "54", "", true)
        @test XLSX.get_formula_from_cache(s2, XLSX.CellRef("B1")) == XLSX.ReferencedFormula("SECOND(NOW())", 0, "B1:C5", Dict("ca" => "1"))
        s2["B1"] = 3
        @test XLSX.getcell(s2, "B1") == XLSX.Cell(XLSX.get_workbook(f), XLSX.CellRef("B1"), "", "", "3", "", false)
        @test XLSX.getcell(s2, "B2") == XLSX.Cell(XLSX.get_workbook(f), XLSX.CellRef("B2"), "", "", "54", "", true)
        @test XLSX.get_formula_from_cache(s2, XLSX.CellRef("B2")) == XLSX.ReferencedFormula("SECOND(NOW())", 3, "B2:C5", Dict("ca" => "1"))
        @test XLSX.getcell(s2, "C1") == XLSX.Cell(XLSX.get_workbook(f), XLSX.CellRef("C1"), "", "", "54", "", true)
        @test XLSX.get_formula_from_cache(s2, XLSX.CellRef("C1")) == XLSX.Formula("SECOND(NOW())", nothing, nothing, Dict("ca" => "1"))

        XLSX.writexlsx("mytest.xlsx", f, overwrite=true)
        SAVE_FILES && save_outfile("mytest.xlsx")
        f2 = XLSX.openxlsx("mytest.xlsx", mode="rw")

        s = f2[1]
        @test XLSX.getcell(s, "A2") == XLSX.Cell(XLSX.get_workbook(f), XLSX.CellRef("A2"), "", "", "3", "", false)
        @test XLSX.getcell(s, "A3") == XLSX.Cell(XLSX.get_workbook(f), XLSX.CellRef("A3"), "", "", "25", "", true)
        @test XLSX.get_formula_from_cache(s, XLSX.CellRef("A3")) == XLSX.ReferencedFormula("SUM(O3:S3)", 4, "A3:A10", nothing)

        s2 = f[2]
        @test XLSX.getcell(s2, "A2") == XLSX.Cell(XLSX.get_workbook(f), XLSX.CellRef("A2"), "", "", "3", "", false)
        @test (XLSX.get_formula_from_cache(s2, XLSX.CellRef("A3"))).id == 2
        @test XLSX.get_formula_from_cache(s2, XLSX.CellRef("A3")).unhandled == Dict("ca" => "1")
        @test XLSX.getcell(s2, "A3") == XLSX.Cell(XLSX.get_workbook(f), XLSX.CellRef("A3"), "", "", "54", "", true)
        @test XLSX.get_formula_from_cache(s2, XLSX.CellRef("A3")) == XLSX.ReferencedFormula("SECOND(NOW())", 2, "A3:A5", Dict("ca" => "1"))
        @test XLSX.getcell(s2, "B1") == XLSX.Cell(XLSX.get_workbook(f), XLSX.CellRef("B1"), "", "", "3", "", false)
        @test XLSX.getcell(s2, "B2") == XLSX.Cell(XLSX.get_workbook(f), XLSX.CellRef("B2"), "", "", "54", "", true)
        @test XLSX.get_formula_from_cache(s2, XLSX.CellRef("B2")) == XLSX.ReferencedFormula("SECOND(NOW())", 3, "B2:C5", Dict("ca" => "1"))
        @test XLSX.getcell(s2, "C1") == XLSX.Cell(XLSX.get_workbook(f), XLSX.CellRef("C1"), "", "", "54", "", true)
        @test XLSX.get_formula_from_cache(s2, XLSX.CellRef("C1")) == XLSX.Formula("SECOND(NOW())", nothing, nothing, Dict("ca" => "1"))

    end
    isfile("mytest.xlsx") && rm("mytest.xlsx")

    @testset "ReferencedFormulae - single-cell remainder shifts references" begin
        f = XLSX.openxlsx("shift_test.xlsx", mode="w")
        s = f[1]

        # Build a 2-column ReferencedFormula block B1:C1 with a formula
        # containing a real relative reference, so a shift is observable.
        XLSX.setFormula(s, "B1:C1"; val="A1+1")

        # Confirm the block was created as a single ReferencedFormula group
        # before we overwrite the master.
        rf_before = XLSX.get_formula_from_cache(s, XLSX.CellRef("B1"))
        @test rf_before isa XLSX.ReferencedFormula
        @test rf_before.ref == "B1:C1"

        # Overwriting the master (B1) should trigger rereference_formulae,
        # leaving C1 as the single-cell remainder.
        s["B1"] = 99

        c1_formula = XLSX.get_formula_from_cache(s, XLSX.CellRef("C1"))
        @test c1_formula isa XLSX.Formula
        @test c1_formula.ref === nothing
        @test c1_formula.type === nothing
        # The reference should be shifted one column right: A1 -> B1
        @test c1_formula.formula == "B1+1"

#        XLSX.writexlsx("shift_test_out.xlsx", f, overwrite=true)
        SAVE_FILES && save_outfile(f)
    end
end


@testset "parallel formulas" begin

    f = XLSX.newxlsx()
    s = f[1]
    for row in 1:1000
        s[row, 1] = row
        XLSX.setFormula(s, "B$row", "= A$row * 2 + 5")
    end
    XLSX.copysheet!(s, "Sheet2")
    XLSX.copysheet!(s, "Sheet3")

    io = IOBuffer()
    XLSX.writexlsx(io, f)
    SAVE_FILES && save_outfile(io)
    seekstart(io)
    f2 = XLSX.openxlsx(io)

    @testset "sheet count and names preserved" begin
        @test XLSX.sheetcount(XLSX.get_workbook(f2)) == 3
        @test XLSX.sheetnames(XLSX.get_workbook(f2)) == ["Sheet1", "Sheet2", "Sheet3"]
    end

    @testset "numeric values consistent across all sheets" begin
        for sheet_no in 1:3
            for row in [1, 100, 500, 999, 1000]
                @test f2[sheet_no]["A$row"] == row
            end
        end
    end

    @testset "formula strings consistent across all sheets" begin
        for sheet_no in 1:3
            for row in [1, 50, 99, 500, 1000]
                @test XLSX.getFormula(f2[sheet_no], "B$row") == XLSX.getFormula(f[1], "B$row")
            end
        end
    end

    @testset "cache fully populated for all sheets" begin
        for sheet_no in 1:3
            @test !isnothing(f2[sheet_no].cache)
            @test f2[sheet_no].cache.is_full
        end
    end

    @testset "repeated round-trips produce consistent results" begin
        for trial in 1:5
            io2 = IOBuffer()
            XLSX.writexlsx(io2, f)
            SAVE_FILES && save_outfile(io2)
            seekstart(io2)
            f3 = XLSX.openxlsx(io2)
            for sheet_no in 1:3
                @test f3[sheet_no]["A500"] == 500
                @test XLSX.getFormula(f3[sheet_no], "B500") == XLSX.getFormula(f[1], "B500")
            end
        end
    end

    # Issue #395
    @testset "Multi-threaded read" begin
        N_FORMULAS = 5000 # Should be a multiple of ROW_CHUNKSIZE
        N_ITER = 5

        xf = XLSX.newxlsx()
        sheet = xf[1]
        sheet["A1"] = "n";
        sheet["B1"] = "double n";
        sheet["C1"] = "formula"
        for i in 1:N_FORMULAS
            sheet["A$(i+1)"] = i
            XLSX.setFormula(sheet, "B$(i+1)", "A$(i+1)*2")
            sheet["C$(i+1)"] = "= A$(i+1) * 2"
        end
        io = IOBuffer()
        XLSX.writexlsx(io, xf)
        SAVE_FILES && save_outfile(io)

        for iter in 1:N_ITER
            seekstart(io)
            try
                df = XLSX.readtable(io, 1)
            catch e
                println("Error in iteration $iter: $e")
                @test false
            end
        end
    end

    # issue #299 & 301
    @testset "empty_v" begin
        xf = XLSX.openxlsx(joinpath(data_directory, "empty_v.xlsx"), mode="rw")
        sheet1 = xf["Sheet1"]
        @test XLSX.getcell(sheet1, "A1") == XLSX.Cell(XLSX.get_workbook(xf), XLSX.CellRef("A1"), "str", "", "", "", true)
        @test XLSX.get_formula_from_cache(sheet1, XLSX.CellRef("A1")) == XLSX.Formula("\"\"")
        XLSX.writexlsx("mytest.xlsx", xf, overwrite=true)
        SAVE_FILES && save_outfile("mytest.xlsx")
        xf2 = XLSX.readxlsx("mytest.xlsx")
        sheet1 = xf2["Sheet1"]
        @test XLSX.getcell(xf2[1], "A1") == XLSX.Cell(XLSX.get_workbook(xf2), XLSX.CellRef("A1"), "str", "", "", "", true)
        @test XLSX.get_formula_from_cache(sheet1, XLSX.CellRef("A1")) == XLSX.Formula("\"\"")
        isfile("mytest.xlsx") && rm("mytest.xlsx")
    end

end