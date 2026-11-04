@testset "Comments" begin
    REL_COMMENTS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments"
    REL_VML      = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing"
    fixture      = joinpath(data_directory, "TableCommentsVML.xlsx")

    fresh() = (xf = XLSX.newxlsx(); (xf, xf["Sheet1"]))
    rel_types(xf, path) = [XML.attributes(n)["Type"] for n in XML.children(XLSX.xml_root_element(xf.data[path]))
                           if XML.nodetype(n) == XML.Element]
    sheet_tags(xf, path) = [XLSX.localname(n) for n in XML.children(XLSX.xml_root_element(xf.data[path]))
                            if XML.nodetype(n) == XML.Element]
    vml(xf, path) = String(copy(xf.binary_data[path]))

    @testset "cell addressing variants" begin
        for (args, ref) in (((1, 1), "A1"), (("B2",), "B2"), ((XLSX.CellRef("E7"),), "E7"))
            xf, s = fresh()
            info = XLSX.addComment(s, args..., "note")
            @test info == (sheet="Sheet1", ref=ref, author="", text="note")
            @test XLSX.getComment(s, args...) == info
            SAVE_FILES && save_outfile(xf)
        end
        xf, s = fresh()
        @test_throws XLSX.XLSXError XLSX.addComment(s, "A1:B2", "range")
        @test_throws XLSX.XLSXError XLSX.addComment(s, "notacell", "bad")
        @test_throws XLSX.XLSXError XLSX.addComment(s, "A1", "bad"; width=0)
        @test isnothing(XLSX.getComment(s, "A1"))
        @test isempty(XLSX.getComments(s))
    end

    @testset "parts, relationships and content types" begin
        xf, s = fresh()
        XLSX.addComment(s, "B3", "first"; author="Tim")
        @test haskey(xf.data, "xl/comments1.xml")
        @test haskey(xf.binary_data, "xl/drawings/vmlDrawing1.vml")
        types = rel_types(xf, "xl/worksheets/_rels/sheet1.xml.rels")
        @test REL_COMMENTS in types && REL_VML in types
        ct = XML.write(xf.data["[Content_Types].xml"])
        @test occursin("/xl/comments1.xml", ct)
        @test occursin("Extension=\"vml\"", ct)

        tags = sheet_tags(xf, "xl/worksheets/sheet1.xml")
        @test count(==("legacyDrawing"), tags) == 1
        @test findfirst(==("legacyDrawing"), tags) > findfirst(==("pageMargins"), tags)

        v = vml(xf, "xl/drawings/vmlDrawing1.vml")
        @test occursin("_x0000_s1025", v)
        @test occursin("<x:Row>2</x:Row>", v) && occursin("<x:Column>1</x:Column>", v)

        # A second note shares the parts and takes the next shape id
        XLSX.addComment(s, "D1", "second"; author="Tim", visible=true)
        @test count(==("legacyDrawing"), sheet_tags(xf, "xl/worksheets/sheet1.xml")) == 1
        v = vml(xf, "xl/drawings/vmlDrawing1.vml")
        @test occursin("_x0000_s1026", v)
        @test occursin("<x:Visible/>", v)
        authors = XLSX.first_element_with_tag(XLSX.xml_root_element(xf.data["xl/comments1.xml"]), "authors")
        @test length(XLSX.elements_with_tag(authors, "author")) == 1
        SAVE_FILES && save_outfile(xf)
    end

    @testset "replace" begin
        xf, s = fresh()
        XLSX.addComment(s, "B3", "old"; author="Tim")
        XLSX.addComment(s, "B3", "new"; author="Ann")
        @test XLSX.getComments(s) == [(sheet="Sheet1", ref="B3", author="Ann", text="new")]
        @test count("ObjectType=\"Note\"", vml(xf, "xl/drawings/vmlDrawing1.vml")) == 1
        SAVE_FILES && save_outfile(xf)
    end

    @testset "write and read back" begin
        xf, s = fresh()
        text = "Escaped & <quoted> \"text\"\nSecond line   with spaces "
        XLSX.addComment(s, "C4", text; author="Tim")
        XLSX.addsheet!(xf, "Two")
        XLSX.addComment(xf["Two"], "A2", "on sheet two")
        XLSX.writexlsx("comments.xlsx", xf; overwrite=true)

        r = XLSX.readxlsx("comments.xlsx")
        @test XLSX.getComment(r["Sheet1"], "C4") == (sheet="Sheet1", ref="C4", author="Tim", text=text)
        @test XLSX.getComments(r) == [
            (sheet="Sheet1", ref="C4", author="Tim", text=text),
            (sheet="Two",    ref="A2", author="",    text="on sheet two"),
        ]

        # Each sheet's VML claims its own idmap block, so shape ids never collide
        z  = ZipArchives.ZipReader(read("comments.xlsx"))
        v1 = String(ZipArchives.zip_readentry(z, "xl/drawings/vmlDrawing1.vml"))
        v2 = String(ZipArchives.zip_readentry(z, "xl/drawings/vmlDrawing2.vml"))
        @test occursin("data=\"1\"", v1) && occursin("_x0000_s1025", v1)
        @test occursin("data=\"2\"", v2) && occursin("_x0000_s2049", v2)
        SAVE_FILES && save_outfile(xf)
    end

    @testset "delete" begin
        xf, s = fresh()
        XLSX.addComment(s, "A1", "one")
        XLSX.addComment(s, "B2", "two")
        @test XLSX.deleteComment(s, "A1")
        @test !XLSX.deleteComment(s, "A1")
        @test !XLSX.deleteComment(s, "Z9")
        @test [c.ref for c in XLSX.getComments(s)] == ["B2"]
        @test !occursin("<x:Row>0</x:Row>", vml(xf, "xl/drawings/vmlDrawing1.vml"))

        # The last one gone: no comments part, VML, relationships or <legacyDrawing>
        @test XLSX.deleteComment(s, 2, 2)
        @test !haskey(xf.data, "xl/comments1.xml")
        @test !haskey(xf.binary_data, "xl/drawings/vmlDrawing1.vml")
        types = rel_types(xf, "xl/worksheets/_rels/sheet1.xml.rels")
        @test REL_COMMENTS ∉ types && REL_VML ∉ types
        @test "legacyDrawing" ∉ sheet_tags(xf, "xl/worksheets/sheet1.xml")
        @test !occursin("/xl/comments1.xml", XML.write(xf.data["[Content_Types].xml"]))

        # ... and comments can be added again afterwards
        XLSX.addComment(s, "C3", "again")
        @test [c.ref for c in XLSX.getComments(s)] == ["C3"]
        SAVE_FILES && save_outfile(xf)
    end

    @testset "existing comments in a template" begin
        xf = XLSX.opentemplate(fixture)
        s  = xf[1]
        @test XLSX.getComments(s) ==
            [(sheet=s.name, ref="A1", author="Tim", text="Test comment for copysheet!/deletesheet! fixture.")]

        XLSX.addComment(s, "D4", "added"; author="Tim")
        v = vml(xf, "xl/drawings/vmlDrawing1.vml")
        @test occursin("_x0000_s1025", v) && occursin("_x0000_s1026", v)
        @test count("_x0000_t202\" coordsize", v) == 1  # shapetype not duplicated
        @test length(XLSX.elements_with_tag(
            XLSX.first_element_with_tag(XLSX.xml_root_element(xf.data["xl/comments1.xml"]), "authors"), "author")) == 1

        XLSX.writexlsx("template-comments.xlsx", xf; overwrite=true)
        r = XLSX.readxlsx("template-comments.xlsx")
        @test [c.ref for c in XLSX.getComments(r)] == ["A1", "D4"]
        @test length(XLSX.getImages(r)) == 1   # the sheet's image is untouched

        XLSX.deleteComment(s, "A1")
        XLSX.deleteComment(s, "D4")
        @test isempty(XLSX.getComments(s))
        @test !haskey(xf.binary_data, "xl/drawings/vmlDrawing1.vml")
        @test "drawing" in sheet_tags(xf, "xl/worksheets/sheet1.xml")
        SAVE_FILES && save_outfile(xf)
    end

    @testset "deletesheet! removes comment parts" begin
        xf, s = fresh()
        XLSX.addsheet!(xf, "Two")
        XLSX.addComment(xf["Two"], "A1", "gone with the sheet")
        XLSX.deletesheet!(xf, "Two")
        @test isempty(XLSX.getComments(xf))
        @test !any(endswith(".vml"), keys(xf.binary_data))
        @test !any(startswith("xl/comments"), keys(xf.data))
        SAVE_FILES && save_outfile(xf)
    end

    @testset "read-only and threaded comments" begin
        xf, s = fresh()
        XLSX.addComment(s, "A1", "note")
        XLSX.writexlsx("ro-comments.xlsx", xf; overwrite=true)
        SAVE_FILES && save_outfile("ro-comments.xlsx")
        r = XLSX.readxlsx("ro-comments.xlsx")
        @test_throws XLSX.XLSXError XLSX.addComment(r[1], "B1", "nope")
        @test_throws XLSX.XLSXError XLSX.deleteComment(r[1], "A1")

        # A cell with a threaded comment is refused rather than desynchronised
        xf, s = fresh()
        XLSX.addComment(s, "A1", "[Threaded comment] fallback")
        path = "xl/threadedComments/threadedComment1.xml"
        xf.data[path] = parse("""<?xml version="1.0" encoding="UTF-8"?>
            <ThreadedComments xmlns="http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments">
            <threadedComment ref="A1" dT="2026-01-01T00:00:00.00" personId="{0}" id="{1}"><text>hi</text></threadedComment>
            </ThreadedComments>""", XML.Node)
        xf.files[path] = true
        XLSX.add_part_rel!(xf, "xl/worksheets/sheet1.xml", path,
                           "http://schemas.microsoft.com/office/2017/10/relationships/threadedComment")
        @test_throws XLSX.XLSXError XLSX.addComment(s, "A1", "replace")
        @test_throws XLSX.XLSXError XLSX.deleteComment(s, "A1")
        @test XLSX.addComment(s, "B1", "a different cell is fine").ref == "B1"
    end
end
