@testset "sheet view" begin
    sheet_views(xf, ws) = only(XLSX._sheet_top_elements(xf, XLSX._sheet_path(ws), ("sheetViews",)))
    sheet_view(xf, ws) = first(XLSX.xml_elements(sheet_views(xf, ws)))
    view_xml(xf, ws) = replace(XML.write(sheet_views(xf, ws)), " tabSelected=\"0\"" => "")   # as addsheet! writes it; Excel leaves out the default
    function roundtrip(xf)
        io = IOBuffer()
        XLSX.writexlsx(io, xf)
        return XLSX.readxlsx(IOBuffer(take!(io)))
    end
    defaults = (zoom=100, view=:normal, showGridLines=true, showRowColHeaders=true, showZeros=true,
                      rightToLeft=false, topLeftCell=XLSX.CellRef("A1"), activeCell=XLSX.CellRef("A1"), selection="A1")

    @testset "settings" begin
        xf = XLSX.newxlsx()
        s = xf[1]
        @test XLSX.getSheetView(s) == defaults
        XLSX.setSheetView(s; zoom=150, showGridLines=false, showRowColHeaders=false, showZeros=false, rightToLeft=true)
        v = XLSX.getSheetView(s)
        @test (v.zoom, v.showGridLines, v.showRowColHeaders, v.showZeros, v.rightToLeft) == (150, false, false, false, true)
        sv = sheet_view(xf, s)
        @test [first(p) for p in XLSX._atts_vector(sv)] ==            # in the schema's order, as Excel writes them
              ["showGridLines", "showRowColHeaders", "showZeros", "rightToLeft", "tabSelected", "zoomScale", "zoomScaleNormal", "workbookViewId"]
        XLSX.setSheetView(s; zoom=100, showGridLines=true, showRowColHeaders=true, showZeros=true, rightToLeft=false)
        @test XLSX.getSheetView(s) == defaults
        @test Dict(XLSX._atts_vector(sheet_view(xf, s))) == Dict("tabSelected" => "1", "zoomScaleNormal" => "100", "workbookViewId" => "0")

        XLSX.setSheetView(s; topLeftCell="C10")
        @test XLSX.getSheetView(s).topLeftCell == XLSX.CellRef("C10")
        XLSX.setSheetView(s; topLeftCell=XLSX.CellRef("A1"))
        @test !haskey(sheet_view(xf, s), "topLeftCell")

        XLSX.setSheetView(s; view=:pageLayout, zoom=80)
        @test (XLSX.getSheetView(s).view, XLSX.getSheetView(s).zoom) == (:pageLayout, 80)
        XLSX.setSheetView(s; view=:pageBreakPreview)                     # Excel's first zoom for it
        @test (XLSX.getSheetView(s).view, XLSX.getSheetView(s).zoom) == (:pageBreakPreview, 60)
        XLSX.setSheetView(s; view=:pageLayout)                           # each view keeps its zoom
        @test XLSX.getSheetView(s).zoom == 80
        XLSX.setSheetView(s; view=:normal)
        @test XLSX.getSheetView(s).zoom == 100 && !haskey(sheet_view(xf, s), "view")

        @test XLSX.getSheetView(roundtrip(xf)[1]) == XLSX.getSheetView(s)
        SAVE_FILES && save_outfile(xf)
    end

    @testset "selection" begin
        xf = XLSX.newxlsx()
        s = xf[1]
        XLSX.setSheetView(s; selection="B2:D5,F7", activeCell="C3")
        v = XLSX.getSheetView(s)
        @test (v.selection, v.activeCell) == ("B2:D5,F7", XLSX.CellRef("C3"))
        @test occursin("<selection activeCell=\"C3\" sqref=\"B2:D5 F7\"/>", view_xml(xf, s))
        XLSX.setSheetView(s; selection="B2:D5")                          # its first cell is active
        @test XLSX.getSheetView(s).activeCell == XLSX.CellRef("B2")
        XLSX.setSheetView(s; activeCell="H8")                            # alone, selected on its own
        @test XLSX.getSheetView(s).selection == "H8"
        @test count("<selection", view_xml(xf, s)) == 1
        XLSX.setSheetView(s; selection=XLSX.CellRange("A1:B2"))
        @test XLSX.getSheetView(s).selection == "A1:B2"
        XLSX.addDefinedName(s, "picked", "D4:E5")
        XLSX.setSheetView(s; selection="picked")
        @test XLSX.getSheetView(s).selection == "D4:E5"
        @test XLSX.getSheetView(roundtrip(xf)[1]).selection == "D4:E5"
        SAVE_FILES && save_outfile(xf)
    end

    @testset "with panes" begin
        xf = XLSX.newxlsx()
        s = xf[1]
        XLSX.freezePanes(s, "B3")
        XLSX.setSheetView(s; topLeftCell="B15", activeCell="E20")
        v = XLSX.getSheetView(s)
        @test (v.topLeftCell, v.activeCell) == (XLSX.CellRef("B15"), XLSX.CellRef("E20"))
        sv = sheet_view(xf, s)
        @test !haskey(sv, "topLeftCell")                                  # the scroll goes on the pane
        @test [XLSX._selection_pane(c) for c in XLSX.xml_elements(sv) if XLSX.localname(c) == "selection"] ==
              ["topRight", "bottomLeft", "bottomRight"]
        XLSX.setSheetView(s; selection="A1:C8")                           # on the active pane, as in Excel
        @test occursin("<selection pane=\"bottomRight\" activeCell=\"A1\" sqref=\"A1:C8\"/>", view_xml(xf, s))
        @test_throws XLSX.XLSXError XLSX.setSheetView(s; topLeftCell="A15")   # in the frozen column
        @test_throws XLSX.XLSXError XLSX.setSheetView(s; topLeftCell="C2")    # in the frozen rows
        XLSX.setSheetView(s; zoom=120)                                    # the panes are kept
        @test count("<selection", view_xml(xf, s)) == 3 && occursin("state=\"frozen\"", view_xml(xf, s))

        XLSX.removePanes(s)
        XLSX.splitPanes(s; nrows=2)                                       # a split pane scrolls anywhere
        XLSX.setSheetView(s; topLeftCell="A1")
        @test XLSX.getSheetView(s).topLeftCell == XLSX.CellRef("A1")

        # freezing counts from row 1, so it brings the view back there, as Excel freezes what is in view
        t = XLSX.addsheet!(xf, "scrolled")
        XLSX.setSheetView(t; topLeftCell="A10", zoom=90)
        XLSX.freezePanes(t; nrows=1)
        @test !haskey(sheet_view(xf, t), "topLeftCell") && XLSX.getSheetView(t).zoom == 90
        @test XLSX.getSheetView(t).topLeftCell == XLSX.CellRef("A2")
        SAVE_FILES && save_outfile(xf)
    end

    @testset "as Excel writes it" begin
        ex = XLSX.readxlsx(joinpath(data_directory, "sheetview_excel.xlsx"))
        xf = XLSX.newxlsx("zoom")
        for n in ("flags", "sel", "scroll", "pagebreak", "frozen")
            XLSX.addsheet!(xf, n)
        end
        XLSX.setSheetView(xf["zoom"]; zoom=150)
        XLSX.setSheetView(xf["flags"]; showGridLines=false, showRowColHeaders=false, showZeros=false, rightToLeft=true)
        XLSX.setSheetView(xf["sel"]; selection="B2:D5,F7", activeCell="C3")
        XLSX.setSheetView(xf["scroll"]; topLeftCell="C10")
        XLSX.setSheetView(xf["pagebreak"]; view=:pageBreakPreview)
        XLSX.freezePanes(xf["frozen"], "B3")
        XLSX.setSheetView(xf["frozen"]; topLeftCell="B15", activeCell="E20")
        for n in ("zoom", "flags", "sel", "scroll", "pagebreak", "frozen")
            @test XLSX.getSheetView(ex[n]) == XLSX.getSheetView(xf[n])
            n == "pagebreak" && continue   # Excel also writes the normal view's zoom, unchanged at 100
            @test view_xml(ex, ex[n]) == view_xml(xf, xf[n])
        end
        @test XLSX.getSheetView(ex["pagebreak"]).zoom == 60
        SAVE_FILES && save_outfile(xf)
    end

    @testset "errors" begin
        xf = XLSX.newxlsx()
        s = xf[1]
        @test_throws XLSX.XLSXError XLSX.setSheetView(s)
        @test_throws XLSX.XLSXError XLSX.setSheetView(s; gridLines=false)
        @test_throws XLSX.XLSXError XLSX.setSheetView(s; showGridLines=0)
        @test_throws XLSX.XLSXError XLSX.setSheetView(s; zoom=9)
        @test_throws XLSX.XLSXError XLSX.setSheetView(s; zoom=401)
        @test_throws XLSX.XLSXError XLSX.setSheetView(s; zoom=150.5)
        @test_throws XLSX.XLSXError XLSX.setSheetView(s; view=:outline)
        @test_throws XLSX.XLSXError XLSX.setSheetView(s; view="normal")
        @test_throws XLSX.XLSXError XLSX.setSheetView(s; topLeftCell="A1:B2")
        @test_throws XLSX.XLSXError XLSX.setSheetView(s; activeCell="Z99", selection="A1:B2")
        @test_throws XLSX.XLSXError XLSX.setSheetView(s; selection="B:C")
        @test_throws XLSX.XLSXError XLSX.setSheetView(s; selection=1)
        @test_throws XLSX.XLSXError XLSX.setSheetView(s; zoom=150, view=:sideways)  # nothing changes
        @test XLSX.getSheetView(s).zoom == 100

        r = roundtrip(xf)
        @test XLSX.getSheetView(r[1]) == defaults                  # read-only works
        @test_throws XLSX.XLSXError XLSX.setSheetView(r[1]; zoom=150)

        c = XLSX.Charts.addChart(xf, :line; sheetname="Chart")
        @test_throws XLSX.XLSXError XLSX.setSheetView(xf["Chart"]; zoom=150)
        SAVE_FILES && save_outfile(xf)
    end
end
