@testset "hyperlinks" begin
    REL_HYPERLINK = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"
    URL = "https://julialang.org"

    fresh() = (xf = XLSX.newxlsx("Data"); (xf, xf["Data"]))
    rels(xf, path) = [XML.attributes(n) for n in XML.children(XLSX.xml_root_element(xf.data[path]))
                      if XML.nodetype(n) == XML.Element]
    sheet_tags(xf, path) = [XLSX.localname(n) for n in XML.children(XLSX.xml_root_element(xf.data[path]))
                            if XML.nodetype(n) == XML.Element]
    function roundtrip(xf)
        io = IOBuffer()
        XLSX.writexlsx(io, xf)
        return XLSX.readxlsx(IOBuffer(take!(io)))
    end

    @testset "cell addressing variants" begin
        for (args, ref) in (((1, 1), "A1"), (("B2",), "B2"), ((XLSX.CellRef("E7"),), "E7"))
            xf, s = fresh()
            info = XLSX.addHyperlink(s, args...; url=URL)
            @test info == (sheet="Data", ref=ref, url=URL, location=nothing, display=nothing, tooltip=nothing)
            @test XLSX.getHyperlink(s, args...) == info
            SAVE_FILES && save_outfile(xf)
        end
        xf, s = fresh()
        @test_throws XLSX.XLSXError XLSX.addHyperlink(s, "notacell"; url=URL)
        @test_throws XLSX.XLSXError XLSX.getHyperlink(s, "A1:B2")
        @test isnothing(XLSX.getHyperlink(s, "A1"))
        @test isempty(XLSX.getHyperlinks(s))
    end

    @testset "argument checks" begin
        xf, s = fresh()
        @test_throws XLSX.XLSXError XLSX.addHyperlink(s, "A1")
        @test_throws XLSX.XLSXError XLSX.addHyperlink(s, "A1"; url=URL, location="Data!A1")
        @test_throws XLSX.XLSXError XLSX.addHyperlink(s, "A1"; url="")
        @test_throws XLSX.XLSXError XLSX.addHyperlink(s, "A1"; location="")
        @test_throws XLSX.XLSXError XLSX.addHyperlink(s, "A1"; url="https://example.com/" * "x"^2100)
        @test isempty(XLSX.getHyperlinks(s))
        @test !haskey(xf.data, "xl/worksheets/_rels/sheet1.xml.rels")
    end

    @testset "external and internal links" begin
        xf, s = fresh()
        XLSX.addHyperlink(s, "A1"; url=URL, display="Julia", tooltip="The Julia site")
        XLSX.addHyperlink(s, "A2"; location="Data!C5")
        XLSX.addHyperlink(s, "A3"; url="mailto:team@example.com")

        rs = rels(xf, "xl/worksheets/_rels/sheet1.xml.rels")
        @test length(rs) == 2
        @test all(r -> r["Type"] == REL_HYPERLINK && r["TargetMode"] == "External", rs)
        @test Set(r["Target"] for r in rs) == Set([URL, "mailto:team@example.com"])

        tags = sheet_tags(xf, "xl/worksheets/sheet1.xml")
        @test "hyperlinks" in tags
        @test findfirst(==("hyperlinks"), tags) < findfirst(==("pageMargins"), tags)

        @test XLSX.getHyperlinks(s) == [
            (sheet="Data", ref="A1", url=URL, location=nothing, display="Julia", tooltip="The Julia site"),
            (sheet="Data", ref="A2", url=nothing, location="Data!C5", display=nothing, tooltip=nothing),
            (sheet="Data", ref="A3", url="mailto:team@example.com", location=nothing, display=nothing, tooltip=nothing),
        ]
        SAVE_FILES && save_outfile(xf)
    end

    @testset "the same URL shares one relationship" begin
        xf, s = fresh()
        XLSX.addHyperlink(s, "A1"; url=URL)
        XLSX.addHyperlink(s, "A2"; url=URL)
        @test length(rels(xf, "xl/worksheets/_rels/sheet1.xml.rels")) == 1
        @test XLSX.deleteHyperlink(s, "A1")
        @test length(rels(xf, "xl/worksheets/_rels/sheet1.xml.rels")) == 1   # still used by A2
        @test XLSX.getHyperlink(s, "A2").url == URL
        @test XLSX.deleteHyperlink(s, "A2")
        @test isempty(rels(xf, "xl/worksheets/_rels/sheet1.xml.rels"))
    end

    @testset "cell value and font" begin
        xf, s = fresh()
        s["B1"] = "kept"
        XLSX.addHyperlink(s, "A1"; url=URL, display="Julia")
        XLSX.addHyperlink(s, "A2"; url=URL)
        XLSX.addHyperlink(s, "A3"; location="Data!A1")
        XLSX.addHyperlink(s, "B1"; url=URL, display="ignored for the cell")
        @test s["A1"] == "Julia"
        @test s["A2"] == URL
        @test s["A3"] == "Data!A1"
        @test s["B1"] == "kept"
        f = XLSX.getFont(s, "A1").font
        @test f["color"] == Dict("rgb" => "FF0563C1")
        @test haskey(f, "u")
        @test isnothing(XLSX.getFont(s, "B1")) || !haskey(XLSX.getFont(s, "B1").font, "u")

        # Deleting the link leaves the cell alone
        @test XLSX.deleteHyperlink(s, "A1")
        @test s["A1"] == "Julia"
        @test haskey(XLSX.getFont(s, "A1").font, "u")
    end

    @testset "replacing a link" begin
        xf, s = fresh()
        XLSX.addHyperlink(s, "A1"; url=URL)
        XLSX.addHyperlink(s, "A1"; location="Data!B2", tooltip="moved")
        @test XLSX.getHyperlinks(s) == [(sheet="Data", ref="A1", url=nothing, location="Data!B2", display=nothing, tooltip="moved")]
        @test isempty(rels(xf, "xl/worksheets/_rels/sheet1.xml.rels"))
        @test s["A1"] == URL   # the cell was filled by the first link; the second finds it full
    end

    @testset "delete" begin
        xf, s = fresh()
        @test !XLSX.deleteHyperlink(s, "A1")
        XLSX.addHyperlink(s, "A1"; url=URL)
        XLSX.addHyperlink(s, 2, 1; location="Data!A1")
        @test !XLSX.deleteHyperlink(s, "C3")
        @test XLSX.deleteHyperlink(s, 1, 1)
        @test !XLSX.deleteHyperlink(s, "A1")
        @test XLSX.deleteHyperlink(s, XLSX.CellRef("A2"))
        @test isempty(XLSX.getHyperlinks(s))
        @test !("hyperlinks" in sheet_tags(xf, "xl/worksheets/sheet1.xml"))
        SAVE_FILES && save_outfile(xf)
    end

    @testset "a link on a range of cells" begin
        xf, s = fresh()
        s["D2"] = 7
        info = XLSX.addHyperlink(s, "C1:D2"; url=URL, display="Julia", tooltip="tip")
        @test info.ref == "C1:D2"
        block = XLSX.first_element_with_tag(XLSX.xml_root_element(xf.data["xl/worksheets/sheet1.xml"]), "hyperlinks")
        @test only(XLSX.elements_with_tag(block, "hyperlink"))["ref"] == "C1:D2"   # one link, as Excel writes it
        @test length(rels(xf, "xl/worksheets/_rels/sheet1.xml.rels")) == 1
        @test XLSX.getHyperlink(s, "D2") == info
        @test isnothing(XLSX.getHyperlink(s, "E2"))

        # As Excel does: the text in the top-left cell, the link font on the empty cells
        @test s["C1"] == "Julia"
        @test ismissing(s["D1"]) && ismissing(s["C2"])
        @test s["D2"] == 7
        for ref in ("C1", "D1", "C2")
            @test haskey(XLSX.getFont(s, ref).font, "u")
        end
        @test XLSX.getcell(s, "D1").style == XLSX.getcell(s, "C2").style
        @test isnothing(XLSX.getFont(s, "D2")) || !haskey(XLSX.getFont(s, "D2").font, "u")   # a value keeps its format

        # Overlapping links are refused; the same range is replaced
        @test_throws XLSX.XLSXError XLSX.addHyperlink(s, "D1"; url=URL)
        @test_throws XLSX.XLSXError XLSX.addHyperlink(s, "B2:C3"; url=URL)
        XLSX.addHyperlink(s, "C1:D2"; location="Data!A1")
        @test XLSX.getHyperlinks(s) == [(sheet="Data", ref="C1:D2", url=nothing, location="Data!A1", display=nothing, tooltip=nothing)]
        @test isempty(rels(xf, "xl/worksheets/_rels/sheet1.xml.rels"))
        @test s["C1"] == "Julia"   # no longer empty

        # By CellRange, and a one-cell range is written as the cell
        @test XLSX.addHyperlink(s, XLSX.CellRange("F1:F3"); url=URL).ref == "F1:F3"
        @test XLSX.addHyperlink(s, "H1:H1"; url=URL).ref == "H1"
        @test XLSX.getHyperlink(s, "F2").ref == "F1:F3"

        # Deleting by a cell removes the whole range's link; by a range, every link it overlaps
        @test XLSX.deleteHyperlink(s, "D2")
        @test [h.ref for h in XLSX.getHyperlinks(s)] == ["F1:F3", "H1"]
        @test !XLSX.deleteHyperlink(s, "A5:B9")
        @test XLSX.deleteHyperlink(s, "F1:H3")
        @test isempty(XLSX.getHyperlinks(s))

        XLSX.addHyperlink(s, "A10:B11"; url=URL)
        @test XLSX.getHyperlink(roundtrip(xf)["Data"], "B11").ref == "A10:B11"
        SAVE_FILES && save_outfile(xf)
    end

    @testset "round trip, read-only, with and without the cache" begin
        xf, s = fresh()
        s["C3"] = 1
        XLSX.addHyperlink(s, "A1"; url=URL, display="Julia", tooltip="tip")
        XLSX.addHyperlink(s, "A2"; location="Data!C3")
        expected = XLSX.getHyperlinks(s)
        io = IOBuffer()
        XLSX.writexlsx(io, xf)
        bytes = take!(io)
        for enable_cache in (true, false)
            r = XLSX.openxlsx(IOBuffer(bytes); enable_cache)
            @test XLSX.getHyperlinks(r) == expected
            @test XLSX.getHyperlink(r["Data"], "A2").location == "Data!C3"
            @test r["Data"]["A1"] == "Julia"   # reading the links leaves the cells readable
            @test r["Data"]["C3"] == 1
        end
        @test_throws XLSX.XLSXError XLSX.addHyperlink(XLSX.readxlsx(IOBuffer(bytes))["Data"], "B1"; url=URL)
        @test_throws XLSX.XLSXError XLSX.deleteHyperlink(XLSX.readxlsx(IOBuffer(bytes))["Data"], "A1")

        # Edit after reopening for writing
        xf2 = XLSX.openxlsx(IOBuffer(bytes); mode="rw")
        @test XLSX.deleteHyperlink(xf2["Data"], "A1")
        XLSX.addHyperlink(xf2["Data"], "B5"; url="https://example.com")
        r = roundtrip(xf2)
        @test [h.ref for h in XLSX.getHyperlinks(r)] == ["A2", "B5"]
        @test XLSX.getHyperlink(r["Data"], "B5").url == "https://example.com"
    end

    @testset "copysheet! gives the copy its own relationships" begin
        xf, s = fresh()
        XLSX.addHyperlink(s, "A1"; url=URL)
        XLSX.addHyperlink(s, "A2"; location="Data!B1")
        c = XLSX.copysheet!(s, "Copy")
        copy_path = XLSX._sheet_path(c)
        crels = rels(xf, XLSX._sheet_rels_path(copy_path))
        @test length(crels) == 1 && only(crels)["Target"] == URL
        @test [(h.ref, h.url, h.location) for h in XLSX.getHyperlinks(c)] == [("A1", URL, nothing), ("A2", nothing, "Data!B1")]

        # Deleting the original's link leaves the copy's
        @test XLSX.deleteHyperlink(s, "A1")
        @test XLSX.getHyperlink(c, "A1").url == URL

        r = roundtrip(xf)
        @test XLSX.getHyperlink(r["Copy"], "A1").url == URL
        SAVE_FILES && save_outfile(xf)
    end

    @testset "deletesheet! with links" begin
        xf, s = fresh()
        o = XLSX.addsheet!(xf, "Other")
        XLSX.addHyperlink(o, "A1"; url=URL)
        XLSX.addHyperlink(s, "A1"; url=URL)
        XLSX.deletesheet!(xf, "Other")
        r = roundtrip(xf)
        @test XLSX.getHyperlinks(r) == [(sheet="Data", ref="A1", url=URL, location=nothing, display=nothing, tooltip=nothing)]
    end

    @testset "chartsheets have none" begin
        xf, s = fresh()
        s["A1"] = 1; s["A2"] = 2
        c = XLSX.Charts.addChart(xf, :column; title="t")
        XLSX.Charts.addSeries(c, "Data!A1:A2")
        cs = xf[XLSX.sheetnames(xf)[end]]
        @test isempty(XLSX.getHyperlinks(cs))
        @test_throws XLSX.XLSXError XLSX.addHyperlink(cs, "A1"; url=URL)
        @test length(XLSX.getHyperlinks(xf)) == 0
    end
end
