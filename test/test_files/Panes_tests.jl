@testset "freeze/split panes" begin

    # White-box helper: walk down to the <pane> element (and any <selection>
    # elements) of a worksheet's first <sheetView>, or `nothing` if absent.
    function _pane_and_selections(ws::XLSX.Worksheet)
        doc = XLSX.get_worksheet_xml_document(ws)

        i, j = XLSX.get_idces(doc, "worksheet", "sheetViews")
        j === nothing && return nothing, nothing, XML.Node[]
        sheetViews = doc[i][j]

        k, l = XLSX.get_idces(sheetViews, "sheetView", "pane")
        k === nothing && return nothing, nothing, XML.Node[]
        sheetView = sheetViews[k]
        pane = l === nothing ? nothing : sheetView[l]

        children = something(sheetView.children, XML.Node{String}[])
        selections = [c for c in children if XLSX.localname(c) == "selection"]
        return sheetView, pane, selections
    end

    @testset "freezePanes basic (nrows/ncols)" begin
        f = XLSX.open_empty_template()
        s = f["Sheet1"]
        s["A1:E10"] = ""

        XLSX.freezePanes(s; nrows=2, ncols=1)
        sheetView, pane, selections = _pane_and_selections(s)

        @test pane !== nothing
        @test pane["topLeftCell"] == "B3"
        @test pane["xSplit"] == "1"
        @test pane["ySplit"] == "2"
        @test pane["state"] == "frozen"
        @test pane["activePane"] == "bottomRight"

        @test length(selections) == 3
        bytag(pn) = only(filter(sel -> sel["pane"] == pn, selections))
        @test bytag("topRight")["activeCell"] == "B1"
        @test bytag("bottomLeft")["activeCell"] == "A3"
        @test bytag("bottomRight")["activeCell"] == "B3"

        SAVE_FILES && save_outfile(f)
    end

    @testset "freezePanes defaults" begin
        f = XLSX.open_empty_template()
        s = f["Sheet1"]
        s["A1:E10"] = ""

        XLSX.freezePanes(s)   # nrows=1, ncols=0 by default
        _, pane, selections = _pane_and_selections(s)

        @test pane["topLeftCell"] == "A2"
        @test pane["ySplit"] == "1"
        @test !haskey(pane, "xSplit")   # no column freeze requested
        @test pane["activePane"] == "bottomLeft"
        @test length(selections) == 1
        @test only(selections)["pane"] == "bottomLeft"

        SAVE_FILES && save_outfile(f)
    end

    @testset "freezePanes anchor_cell form matches nrows/ncols form" begin
        f = XLSX.open_empty_template()
        s = f["Sheet1"]
        s["A1:E10"] = ""

        XLSX.freezePanes(s; nrows=1, ncols=1)
        doc1 = XML.write(XLSX.get_worksheet_xml_document(s))

        f2 = XLSX.open_empty_template()
        s2 = f2["Sheet1"]
        s2["A1:E10"] = ""
        XLSX.freezePanes(s2, "B2")
        doc2 = XML.write(XLSX.get_worksheet_xml_document(s2))

        @test doc1 == doc2
    end

    @testset "idempotency" begin
        f = XLSX.open_empty_template()
        s = f["Sheet1"]
        s["A1:E10"] = ""

        XLSX.freezePanes(s; nrows=2, ncols=1)
        first_write = XML.write(XLSX.get_worksheet_xml_document(s))
        XLSX.freezePanes(s; nrows=2, ncols=1)
        second_write = XML.write(XLSX.get_worksheet_xml_document(s))

        @test first_write == second_write
    end

    @testset "splitFreeze uses frozenSplit state" begin
        f = XLSX.open_empty_template()
        s = f["Sheet1"]
        s["A1:E10"] = ""

        XLSX.splitFreeze(s; nrows=3, ncols=2)
        _, pane, _ = _pane_and_selections(s)

        @test pane["state"] == "frozenSplit"
        @test pane["topLeftCell"] == "C4"
        @test pane["xSplit"] == "2"
        @test pane["ySplit"] == "3"

        SAVE_FILES && save_outfile(f)
    end

    @testset "splitPanes has no state and a selectable corner pane" begin
        f = XLSX.open_empty_template()
        s = f["Sheet1"]
        s["A1:E10"] = ""

        XLSX.splitPanes(s; nrows=1, ncols=0)
        _, pane, selections = _pane_and_selections(s)

        @test !haskey(pane, "state")
        @test !haskey(pane, "xSplit")     # ncols == 0
        @test haskey(pane, "ySplit")
        @test pane["activePane"] == "bottomLeft"

        # Unlike frozen/frozenSplit, a plain split's corner pane is
        # selectable and does get its own <selection>.
        @test length(selections) == 2
        panenames = Set(sel["pane"] for sel in selections)
        @test panenames == Set(["topLeft", "bottomLeft"])

        SAVE_FILES && save_outfile(f)
    end

    @testset "removePanes clears pane/selection but keeps the view" begin
        f = XLSX.open_empty_template()
        s = f["Sheet1"]
        s["A1:E10"] = ""

        XLSX.freezePanes(s; nrows=2, ncols=1)
        XLSX.removePanes(s)
        sheetView, pane, selections = _pane_and_selections(s)

        @test sheetView !== nothing
        @test pane === nothing
        @test isempty(selections)

        SAVE_FILES && save_outfile(f)
    end

    @testset "preserves existing sheetView attributes" begin
        f = XLSX.open_empty_template()
        s = f["Sheet1"]
        s["A1:E10"] = ""

        # A fresh template's default <sheetView> already carries tabSelected;
        # confirm freezePanes doesn't clobber it or workbookViewId.
        sheetView_before, _, _ = _pane_and_selections(s)
        @test sheetView_before["tabSelected"] == "1"

        XLSX.freezePanes(s; nrows=1, ncols=1)
        sheetView_after, _, _ = _pane_and_selections(s)

        @test sheetView_after["tabSelected"] == "1"
        @test sheetView_after["workbookViewId"] == "0"
    end

    @testset "creates sheetViews from scratch when entirely absent" begin
        f = XLSX.open_empty_template()
        s = f["Sheet1"]
        s["A1:E10"] = ""

        # Force the no-sheetViews-at-all case, which a fresh template doesn't
        # naturally produce (it always starts with a default sheetView).
        doc = XLSX.get_worksheet_xml_document(s)
        i, j = XLSX.get_idces(doc, "worksheet", "sheetViews")
        @test j !== nothing
        deleteat!(XML.children(doc[i]), j)
        XLSX.set_worksheet_xml_document!(s, doc)
        _, j2 = XLSX.get_idces(XLSX.get_worksheet_xml_document(s), "worksheet", "sheetViews")
        @test j2 === nothing

        XLSX.freezePanes(s; nrows=1, ncols=1)
        sheetView, pane, selections = _pane_and_selections(s)

        @test sheetView !== nothing
        @test pane["topLeftCell"] == "B2"
        @test length(selections) == 3

        # sheetViews must land after dimension and before sheetFormatPr/sheetData.
        doc2 = XLSX.get_worksheet_xml_document(s)
        i2, _ = XLSX.get_idces(doc2, "worksheet", "sheetViews")
        tags = [XLSX.localname(c) for c in XML.children(doc2[i2])]
        @test findfirst(==("sheetViews"), tags) < findfirst(==("sheetData"), tags)

        SAVE_FILES && save_outfile(f)
    end

    @testset "overwrites an existing pane of a different kind" begin
        f = XLSX.open_empty_template()
        s = f["Sheet1"]
        s["A1:E10"] = ""

        XLSX.splitFreeze(s; nrows=3, ncols=2)
        _, pane1, _ = _pane_and_selections(s)
        @test pane1["state"] == "frozenSplit"

        XLSX.freezePanes(s; nrows=1, ncols=1)
        sheetView, pane2, selections = _pane_and_selections(s)

        @test pane2["state"] == "frozen"
        @test pane2["topLeftCell"] == "B2"
        @test length(selections) == 3   # not 6 -- old selections were replaced, not appended

        SAVE_FILES && save_outfile(f)
    end

    @testset "invalid nrows/ncols" begin
        f = XLSX.open_empty_template()
        s = f["Sheet1"]
        s["A1:E10"] = ""

        @test_throws XLSX.XLSXError XLSX.freezePanes(s; nrows=-1)
        @test_throws XLSX.XLSXError XLSX.freezePanes(s; ncols=-1)
    end

    @testset "save/reopen round trip (namespaced fixture, no do-block)" begin
        # No-Default_NameSpace.xlsx has two sheets that exercise different
        # starting shapes, both under an explicit x: prefix (no bare default
        # namespace):
        #   "Sheet1"      -> sheet.xml  -- no <sheetViews> at all
        #   "XLSX-Export" -> sheet5.xml -- existing self-closed <x:sheetView/>
        src = joinpath(data_directory, "No-Default_NameSpace.xlsx")
        @test isfile(src)

        tmp = joinpath(data_directory, "mytest.xlsx")

        try
            # -- Round 1: open the pristine fixture, apply panes to both
            # tabs, explicitly save AS tmp (never overwriting the source
            # fixture; no do-block means nothing is saved until this call). --
            xf = XLSX.openxlsx(src, mode="rw")
            sh1 = xf["Sheet1"]
            sh2 = xf["XLSX-Export"]

            @test XLSX.get_prefix(sh1) == "x"
            @test XLSX.get_prefix(sh2) == "x"

            XLSX.freezePanes(sh1; nrows=1, ncols=1)
            XLSX.splitFreeze(sh2; nrows=2, ncols=1)

            XLSX.writexlsx(tmp, xf; overwrite=true)

            # -- Round 2: reopen tmp, confirm round 1's panes survived the
            # round trip, then apply different operations and explicitly
            # resave -- again via writexlsx (given the same name again),
            # never savexlsx. --
            xf2 = XLSX.openxlsx(tmp, mode="rw")
            sh1b = xf2["Sheet1"]
            sh2b = xf2["XLSX-Export"]

            _, pane1b, _ = _pane_and_selections(sh1b)
            @test pane1b["state"] == "frozen"
            @test pane1b["topLeftCell"] == "B2"

            _, pane2b, _ = _pane_and_selections(sh2b)
            @test pane2b["state"] == "frozenSplit"
            @test pane2b["topLeftCell"] == "B3"

            XLSX.splitPanes(sh1b; nrows=1, ncols=0)
            XLSX.removePanes(sh2b)

            XLSX.writexlsx(tmp, xf2; overwrite=true)

            # -- Round 3: reopen tmp again, confirm round 2's changes stuck. --
            xf3 = XLSX.openxlsx(tmp, mode="rw")
            sh1c = xf3["Sheet1"]
            sh2c = xf3["XLSX-Export"]

            _, pane1c, selections1c = _pane_and_selections(sh1c)
            @test !haskey(pane1c, "state")   # splitPanes never sets state
            @test length(selections1c) == 2  # topLeft + bottomLeft (corner selectable)

            sheetView2c, pane2c, selections2c = _pane_and_selections(sh2c)
            @test sheetView2c !== nothing     # view itself survives
            @test pane2c === nothing          # removePanes cleared it
            @test isempty(selections2c)

            SAVE_FILES && save_outfile(xf3)
        finally
            isfile(tmp) && rm(tmp; force=true)
        end
    end

end
