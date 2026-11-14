
#
# string-types.jl - every public function that takes a name, a reference or a text
# keyword accepts any AbstractString, not only a String (issue #498). Each case runs
# with a SubString (from `strip`, `split`, `first`...), a LazyString, a slice of a
# RichTextString and, on Julia 1.11+, a StyledStrings `styled"…"` string. A styled
# string never `==` a String, so it also catches a name compared before converting.
#

# Each maker returns a string of the given kind whose plain text is `s`.
const STRING_KINDS = Pair{String,Function}[
    "SubString"      => s -> SubString(s * "  ", 1, length(s)),
    "LazyString"     => s -> lazy"$(s)",
    "RichTextString" => s -> SubString(XLSX.RichTextString(XLSX.RichTextRun(s, [:bold => true]),
                                                           XLSX.RichTextRun(" ")), 1, length(s)),
]
@static if VERSION >= v"1.11-"
    push!(STRING_KINDS, "AnnotatedString" => s -> styled"{bold:$(s)}")
end

@testset "$kind" for (kind, mk) in STRING_KINDS

    @testset "sheet names" begin
        xf = XLSX.newxlsx(mk("First"))
        @test XLSX.sheetnames(xf) == ["First"]
        ws = XLSX.addsheet!(xf, mk("Second"))
        @test ws.name == "Second" && ws.name isa String
        XLSX.renamesheet!(ws, mk("Third"))
        @test ws.name == "Third" && ws.name isa String
        @test xf[mk("Third")] === ws
        @test XLSX.hassheet(xf, mk("third"))
        @test XLSX.copysheet!(ws, mk("Copy")).name == "Copy"
        cs = XLSX.Charts.addChart(xf, :line; sheetname=mk("Trend"))
        @test "Trend" in XLSX.sheetnames(xf)
        # the renamed name reaches workbook.xml, not just the Worksheet
        io = IOBuffer(); XLSX.writexlsx(io, xf); seekstart(io)
        @test XLSX.sheetnames(XLSX.readxlsx(io)) == ["First", "Third", "Copy", "Trend"]
        SAVE_FILES && save_outfile(xf)
    end

    @testset "references" begin
        xf = XLSX.newxlsx("Data")
        ws = xf[1]
        ws["A1"] = 1; ws["A2"] = 2; ws["B1"] = "x"; ws["D2"] = 0
        XLSX.addDefinedName(xf, mk("cells"), mk("Data!A1:A2"))
        XLSX.addDefinedName(xf, mk("greeting"), mk("Hello"))
        @test only(filter(d -> d.name == "greeting", XLSX.getDefinedNames(xf))).value == "Hello"

        @test XLSX.setFont(xf, mk("Data!A1"); bold=true) isa Int
        @test haskey(XLSX.getFont(xf, mk("Data!A1")).font, "b")
        @test haskey(XLSX.getFont(ws, mk("A1")).font, "b")
        XLSX.setFont(xf, mk("cells"); italic=true)
        @test haskey(XLSX.getFont(ws, "A2").font, "i")
        XLSX.setUniformFill(xf, mk("Data!A1:A2"); pattern="solid", fgColor="red")
        @test XLSX.getFill(xf, mk("Data!A2")).fill["patternFill"]["patternType"] == "solid"
        XLSX.setBorder(xf, mk("Data!B1"); left=["style" => "thin"])
        @test !isnothing(XLSX.getBorder(ws, mk("B1")))
        XLSX.mergeCells(xf, mk("Data!C1:D1"))
        @test XLSX.isMergedCell(xf, mk("Data!D1"))
        @test XLSX.getMergedBaseCell(ws, mk("D1")).baseCell == XLSX.CellRef("C1")
        XLSX.setColumnWidth(xf, mk("Data!A1"); width=20)
        @test XLSX.getColumnWidth(xf, mk("Data!A1")) ≈ 20 atol=1
        XLSX.setRowHeight(xf, mk("Data!A1"); height=30)
        @test XLSX.getRowHeight(ws, mk("A1")) ≈ 30 atol=1
        @test XLSX.getText(xf, mk("Data!B1")) == "x"
        XLSX.setConditionalFormat(xf, mk("Data!A1:A2"), :cellIs; operator="greaterThan", value="1")
        XLSX.addDataValidation(xf, mk("Data!B2"); type=:whole, operator=:between, formula1="1", formula2="9")
        XLSX.setFormula(xf, mk("Data!E1"), "=A1+A2")
        @test XLSX.getFormula(xf, mk("Data!E1")) == "=A1+A2"
        @test XLSX.getFormula(ws, mk("E1")) == "=A1+A2"
        @test isnothing(XLSX.getRichTextString(xf, mk("Data!B1")))
        SAVE_FILES && save_outfile(xf)
    end

    @testset "text values" begin
        xf = XLSX.newxlsx()
        ws = xf[1]
        ws["A1"] = 1; ws["A2"] = 2
        XLSX.setFont(ws, "A1"; name=mk("Arial"), under=mk("double"), vertAlign=mk("superscript"))
        @test XLSX.getFont(ws, "A1").font["name"]["val"] == "Arial"
        XLSX.setFill(ws, "A1"; pattern=mk("darkGray"), fgColor="blue")
        XLSX.setAlignment(ws, "A1"; horizontal=mk("center"), vertical=mk("top"))
        @test XLSX.getAlignment(ws, "A1").alignment["alignment"]["horizontal"] == "center"
        XLSX.setFormat(ws, "A2"; format=mk("0.000"))
        @test XLSX.getFormat(ws, "A2").format["numFmt"]["formatCode"] == "0.000"
        XLSX.setConditionalFormat(ws, "A1:A2", :cellIs; operator="greaterThan", value="1",
                                  format=["format" => mk("0.0")])
        XLSX.setColoredDataBars(ws, "A1:A2"; bands=1, direction=mk("leftToRight"))
        XLSX.setFormula(ws, "B1", mk("=SORT(A1:A2)"))
        @test XLSX.getFormula(ws, "B1") == "=_xlfn.SORT(A1:A2)"
        run = XLSX.RichTextRun(mk("run"), [:italic => true])
        @test run.text == "run" && run.text isa String
        @test XLSX.RichTextRun(mk("plain")).text == "plain"
        SAVE_FILES && save_outfile(xf)
    end

    @testset "reading and writing tables" begin
        f = "string-types-$(kind).xlsx"
        XLSX.writetable(f, [[1, 2], ["a", "b"]], ["n", "s"];
                        overwrite=true, sheetname=mk("Data"), anchor_cell=mk("B2"))
        @test XLSX.readdata(f, mk("Data"), "B3") == 1
        t = XLSX.readtable(f, mk("Data"); first_row=3, header=false,
                           column_labels=[mk("num"), mk("str")])
        @test t.column_labels == [:num, :str]
        t = XLSX.readtransposedtable(f, mk("Data"); first_column=mk("B"), header=false,
                                     column_labels=[mk("x"), mk("y"), mk("z")])
        @test t.column_labels == [:x, :y, :z]
        SAVE_FILES && save_outfile(f)
    end
end
