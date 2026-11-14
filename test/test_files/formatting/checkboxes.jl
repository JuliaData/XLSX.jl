@testset "checkboxes" begin
    excel_file = joinpath(data_directory, "checkboxes_excel.xlsx")   # made in Excel 365: see below
    # B2:B5 checkboxes, B3 ticked; C2 a checkbox on a bold, filled, bordered cell; D2 ("x")
    # and D3 (5) given checkboxes, which Excel shows as their values.
    bags(xf) = XML.write(XLSX.xml_root_element(XLSX.get_xml_data(xf, XLSX._workbook_part(xf, XLSX.REL_FEATURE_BAG))); indentsize=0)
    xf_node(xf, ref) = XLSX.styles_cell_xf(XLSX.get_workbook(xf), Int(XLSX.getcell(xf[1], ref).style))
    ext(xf, ref) = XML.write(XLSX.first_element_with_tag(xf_node(xf, ref), "extLst"); indentsize=0)
    function roundtrip(xf)
        io = IOBuffer()
        XLSX.writexlsx(io, xf)
        return XLSX.openxlsx(IOBuffer(take!(io)); mode="rw")
    end
    boxes(cells...) = [(cell=c, checked=t) for (c, t) in cells]

    @testset "reading Excel's" begin
        ex = XLSX.readxlsx(excel_file)                                              # read-only
        s = ex[1]
        @test XLSX.getCheckboxes(s) == boxes(("B2", false), ("C2", false), ("D2", false), ("B3", true),
                                             ("D3", false), ("B4", false), ("B5", false))
        @test XLSX.hasCheckbox(s, "B3") && XLSX.hasCheckbox(s, 2, 3) && XLSX.hasCheckbox(ex, "Sheet1!D2")
        @test !XLSX.hasCheckbox(s, "A1") && !XLSX.hasCheckbox(s, "E2")
        @test s["B3"] === true && s["D2"] == "x"                                    # the values are the cells' own
        @test isempty(XLSX.getCheckboxes(XLSX.newxlsx()[1])) && !XLSX.hasCheckbox(XLSX.newxlsx()[1], "A1")
        st = XLSX.openxlsx(excel_file; enable_cache=false)                         # streamed rows
        @test XLSX.getCheckboxes(st[1]) == XLSX.getCheckboxes(s)
    end

    @testset "as Excel writes them" begin
        ex = XLSX.readxlsx(excel_file)
        xf = XLSX.newxlsx()
        s = xf[1]
        s["D2"] = "x"
        s["D3"] = 5
        s["C2"] = false
        XLSX.setFont(s, "C2"; bold=true)
        XLSX.setCheckbox(s, "B2:B5")
        s["B3"] = true                                                              # ticked
        XLSX.setCheckbox(s, "C2")
        XLSX.setCheckbox(s, "D2:D3")
        @test bags(xf) == bags(ex)
        @test ext(xf, "B2") == ext(ex, "B2")
        @test XLSX.getCheckboxes(s) == XLSX.getCheckboxes(ex[1])
        @test s["B2"] === false && s["D2"] == "x" && s["D3"] == 5                  # empty cells become FALSE
        @test haskey(XLSX.getFont(s, "C2").font, "b")                # other formatting kept
        @test XLSX.getCheckboxes(roundtrip(xf)[1]) == XLSX.getCheckboxes(s)
        SAVE_FILES && save_outfile(xf)
    end

    @testset "with other formatting" begin
        xf = XLSX.newxlsx()
        s = xf[1]
        s["A1"] = 1
        s["C2"] = 1
        XLSX.setCheckbox(s, "B2")
        XLSX.setFill(s, "B2"; pattern="solid", fgColor="yellow")                   # a later format keeps it
        XLSX.setBorder(s, "B2"; allsides=["style" => "thin"])
        @test XLSX.hasCheckbox(s, "B2")
        style = XLSX.getcell(s, "B2").style
        XLSX.setCheckbox(s, "B2")                                                  # already has one
        @test XLSX.getcell(s, "B2").style == style
        XLSX.setCheckbox(s, "C:C")                                                 # whole columns, within the used range
        @test XLSX.getCheckboxes(s) == boxes(("C1", false), ("B2", false), ("C2", false))   # by row
        @test s["C2"] == 1                                                         # kept: Excel shows it, not a box
        XLSX.setCheckbox(xf, "Sheet1!D1")
        @test XLSX.hasCheckbox(s, "D1")
        @test length(XLSX.elements_with_tag(XLSX.xml_root_element(XLSX.get_xml_data(xf, XLSX._workbook_part(xf, XLSX.REL_FEATURE_BAG))), "bag")) == 4
        SAVE_FILES && save_outfile(xf)
    end

    @testset "removing" begin
        xf = XLSX.newxlsx()
        s = xf[1]
        XLSX.removeCheckbox(s, "A1")                                               # none yet: nothing to do
        @test isnothing(XLSX._workbook_part(xf, XLSX.REL_FEATURE_BAG))
        s["B2:B3"] = false
        XLSX.setFont(s, "B2:B3"; italic=true)
        italic = XLSX.getcell(s, "B2").style
        XLSX.setCheckbox(s, "B2:B3")
        s["B3"] = true
        XLSX.removeCheckbox(s, "B2:B3")
        @test isempty(XLSX.getCheckboxes(s))
        @test s["B2"] === false && s["B3"] === true                               # the values stay
        @test XLSX.getcell(s, "B2").style == italic                               # back to the format it had
        @test isnothing(XLSX.first_element_with_tag(xf_node(xf, "B2"), "extLst"))
        XLSX.setCheckbox(s, 4, 2)
        XLSX.removeCheckbox(xf, "Sheet1!B4")
        @test !XLSX.hasCheckbox(s, "B4")
        SAVE_FILES && save_outfile(xf)
    end

    @testset "in Excel's file" begin
        ex = XLSX.openxlsx(excel_file; mode="rw")
        before = bags(ex)
        XLSX.setCheckbox(ex[1], "E2")
        @test bags(ex) == before                                                   # its bags are reused
        @test XLSX.getcell(ex[1], "E2").style == XLSX.getcell(ex[1], "B2").style
        XLSX.removeCheckbox(ex[1], "C2")
        @test !XLSX.hasCheckbox(ex[1], "C2") && haskey(XLSX.getFont(ex[1], "C2").font, "b")
        ro = XLSX.readxlsx(excel_file)
        @test_throws XLSX.XLSXError XLSX.setCheckbox(ro[1], "A1")
        @test_throws XLSX.XLSXError XLSX.removeCheckbox(ro[1], "A1")
        SAVE_FILES && save_outfile(ex)
    end
end
