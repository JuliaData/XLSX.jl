@testset "Add Images" begin
    REL_IMAGE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
    jpeg  = joinpath(data_directory, "track_start.jpg")
    png   = joinpath(data_directory, "track_start.png")
    bytes = read(jpeg)

    # Helper so each testset gets a fresh workbook
    fresh() = (xf = XLSX.newxlsx(); (xf, xf["Sheet1"]))

    @testset "cell addressing variants" begin
        cases = [
            ((1, 1),                      "A1", nothing),
            (("B2",),                     "B2", nothing),
            (("C3:D5",),                  "C3", "D5"),
            ((XLSX.CellRef("E7"),),       "E7", nothing),
            ((XLSX.CellRange("F4:H9"),),  "F4", "H9"),
        ]
        for (args, exp_from, exp_to) in cases
            xf, s = fresh()
            info = XLSX.addImage(s, args..., png)
            @test info.from == exp_from
            exp_to === nothing || @test info.to == exp_to
            @test haskey(xf.binary_data, "xl/media/" * info.media_name)
            SAVE_FILES && save_outfile(xf)
        end
    end
    @testset "IOBuffer input" begin
        xf, s = fresh()
        info = XLSX.addImage(s, 3, 4, IOBuffer(copy(bytes)))
        @test startswith(info.media_name, "image")
        @test xf.binary_data["xl/media/" * info.media_name] == bytes
        @test haskey(xf.data, "xl/drawings/drawing1.xml")
        SAVE_FILES && save_outfile(xf)
    end

    @testset "drawing XML and relationships" begin
        xf, s = fresh()
        info = XLSX.addImage(s, "B2", jpeg; size=(128, 128))

        # Relationships
        rels_root = xf.data["xl/drawings/_rels/drawing1.xml.rels"][end]
        rel_nodes = [n for n in XML.children(rels_root) if XML.tag(n) == "Relationship"]
        @test !isempty(rel_nodes)
        @test any(get(XML.attributes(n), "Type", "") == REL_IMAGE for n in rel_nodes)

        # Anchor geometry
        drawing_root = xf.data["xl/drawings/drawing1.xml"][end]
        anchors = [n for n in XML.children(drawing_root) if XML.tag(n) == "xdr:twoCellAnchor"]
        @test length(anchors) == 1
        @test XLSX._parse_cell_marker(anchors[1], "from"; is_to=false) == info.from
        @test XLSX._parse_cell_marker(anchors[1], "to";   is_to=true)  == info.to
        SAVE_FILES && save_outfile(xf)
    end

    @testset "multiple images" begin
        xf, s = fresh()
        info1 = XLSX.addImage(s, 1, 1, jpeg)
        info2 = XLSX.addImage(s, 5, 5, jpeg)

        @test info1.media_name != info2.media_name
        @test haskey(xf.binary_data, "xl/media/" * info1.media_name)
        @test haskey(xf.binary_data, "xl/media/" * info2.media_name)

        rels_root = xf.data["xl/drawings/_rels/drawing1.xml.rels"][end]
        rel_nodes = [n for n in XML.children(rels_root) if XML.tag(n) == "Relationship"]
        @test length(rel_nodes) == 2
        @test all(get(XML.attributes(n), "Type", "") == REL_IMAGE for n in rel_nodes)

        drawing_root = xf.data["xl/drawings/drawing1.xml"][end]
        anchors = [n for n in XML.children(drawing_root) if XML.tag(n) == "xdr:twoCellAnchor"]
        @test length(anchors) == 2
        @test XLSX._parse_cell_marker(anchors[1], "from"; is_to=false) == info1.from
        @test XLSX._parse_cell_marker(anchors[2], "from"; is_to=false) == info2.from
        SAVE_FILES && save_outfile(xf)
    end

    @testset "round-trip (file and IOBuffer)" begin
        for (label, src) in [("file path", jpeg), ("IOBuffer", IOBuffer(copy(bytes)))]
            xf, s = fresh()
            XLSX.addImage(s, 1, 1, src)
            tmp = tempname() * ".xlsx"
            XLSX.writexlsx(tmp, xf)
            SAVE_FILES && save_outfile(tmp)
            @test isfile(tmp) && filesize(tmp) > 0

            xf2  = XLSX.readxlsx(tmp)
            imgs = XLSX.getImages(xf2)
            @test length(imgs) == 1
            @test imgs[1].sheet == "Sheet1"
            @test startswith(imgs[1].media_name, "image")
        end
    end

    @testset "invalid cell reference" begin
        xf, s = fresh()
        @test_throws ArgumentError XLSX.addImage(s, "ZZZ9999", jpeg)
        SAVE_FILES && save_outfile(xf)
    end

    @testset "image cleaned up when sheet deleted" begin
        xf = XLSX.newxlsx()
        s1 = xf["Sheet1"]
        XLSX.addImage(s1, 1, 1, png)
        info = XLSX.getImages(s1)[1]

        wb = XLSX.get_workbook(xf)
        XLSX.addsheet!(wb, "Sheet2")  # need a second sheet to allow deletion
        XLSX.deletesheet!(wb, "Sheet1")

        # Media removed
        @test !haskey(xf.binary_data, "xl/media/" * info.media_name)

        # Drawing XML and rels removed
        @test !haskey(xf.data, "xl/drawings/drawing1.xml")
        @test !haskey(xf.data, "xl/drawings/_rels/drawing1.xml.rels")

        # No images reported
        @test isempty(XLSX.getImages(xf))
        SAVE_FILES && save_outfile(xf)
    end

    @testset "shared media preserved when only one sheet deleted" begin
        xf = XLSX.newxlsx()
        s1 = xf["Sheet1"]
        wb = XLSX.get_workbook(xf)

        XLSX.addImage(s1, 1, 1, jpeg)

        XLSX.copysheet!(s1, "Sheet2")
        s2 = xf["Sheet2"]

        info1 = XLSX.getImages(s1)[1]
        info2 = XLSX.getImages(s2)[1]

        # Both sheets reference the same media file
        @test info1.media_name == info2.media_name
        media_key = "xl/media/" * info1.media_name

        XLSX.deletesheet!(wb, "Sheet1")

        # Media still present — Sheet2 still references it
        @test haskey(xf.binary_data, media_key)

        # Sheet2 image still retrievable
        imgs = XLSX.getImages(xf)
        @test length(imgs) == 1
        @test imgs[1].sheet == "Sheet2"
        SAVE_FILES && save_outfile(xf)
    end
end