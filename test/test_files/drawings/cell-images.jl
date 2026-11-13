@testset "pictures in cells" begin
    excel_file = joinpath(data_directory, "cellimages_excel.xlsx")   # made in Excel 365: see below
    # B2: a PNG placed in the cell, alt text "Logo\n"; B3: a JPEG; B4: B2's PNG again;
    # C4: a floating picture moved into the cell; B5: =IMAGE(url); D2: =SEQUENCE(3).
    part(xf, rel) = XLSX.xml_root_element(XLSX.get_xml_data(xf, XLSX._workbook_part(xf, rel)))
    els(root, tag) = [replace(XML.write(n), r">\s+<" => "><") for n in XLSX.elements_with_tag(root, tag)]
    saved(xf) = (io = IOBuffer(); XLSX.writexlsx(io, xf); take!(io))
    sheet_xml(xf) = replace(zip_readentry(ZipReader(saved(xf)), "xl/worksheets/sheet1.xml", String), r">\s+<" => "><")
    roundtrip(xf) = XLSX.openxlsx(IOBuffer(saved(xf)); mode="rw")
    png = read(joinpath(data_directory, "track_start.png"))
    url = "https://upload.wikimedia.org/wikipedia/commons/4/47/PNG_transparency_demonstration_1.png"

    @testset "reading Excel's" begin
        ex = XLSX.readxlsx(excel_file)                                              # read-only
        s = ex[1]
        b2 = XLSX.getCellImage(s, "B2")
        @test (b2.format, b2.alt, b2.url) == (:png, "Logo\n", nothing)
        @test XLSX.getCellImage(s, "B3").format == :jpeg
        @test XLSX.getCellImage(s, "B4") == XLSX.CellImage(b2.bytes, :png, "", nothing)   # the same picture, shared
        @test XLSX.getCellImage(s, 4, 3).format == :png                             # C4, moved into the cell
        b5 = XLSX.getCellImage(ex, "Sheet1!B5")                                    # an IMAGE formula's picture
        @test b5.url == url && b5.format == :png && !isempty(b5.bytes)
        @test isnothing(XLSX.getCellImage(s, "D2")) && isnothing(XLSX.getCellImage(s, "A1"))
        @test [c.cell for c in XLSX.getCellImages(s)] == ["B2", "B3", "B4", "C4", "B5"]
        @test ismissing(s["B2"]) && XLSX.iserror(s, "B2") && XLSX.geterror(s, "B2") == "#VALUE!"   # as before: Excel stores #VALUE!
        @test XLSX.getText(s, "B2") == ""
        t = XLSX.readtable(excel_file, 1, "B:D"; first_row=2, header=false).data
        @test all(ismissing, t[1]) && t[3] == [1, 2, 3]
        st = XLSX.openxlsx(excel_file; enable_cache=false)                         # streamed rows
        @test all(ismissing(r[2]) for r in XLSX.eachrow(st[1]))
        @test [c.cell for c in XLSX.getCellImages(st[1])] == ["B2", "B3", "B4", "C4", "B5"]
        @test sprint(show, b2) == "CellImage(:png, $(length(b2.bytes)) bytes, alt=\"Logo\\n\")"
        @test occursin("url=", sprint(show, b5))
    end

    @testset "kept when saved" begin
        ex = XLSX.openxlsx(excel_file; mode="rw")
        before = XLSX.getCellImages(ex[1])
        r = roundtrip(ex)
        @test XLSX.getCellImages(r[1]) == before
        for (c, vm) in (("B2", 1), ("B3", 2), ("B4", 3), ("C4", 4), ("B5", 5))
            @test occursin("<c r=\"$c\" t=\"e\" vm=\"$vm\">", sheet_xml(r))
        end
        @test occursin(r"<c r=\"B5\" t=\"e\" vm=\"5\"><f>_xlfn\.IMAGE", sheet_xml(r))   # an IMAGE formula's too
        @test occursin("<c r=\"D2\" t=\"n\" cm=\"1\">", sheet_xml(r))
    end

    @testset "as Excel writes them" begin
        ex = XLSX.readxlsx(excel_file)
        pics = Dict(c.cell => c.image for c in XLSX.getCellImages(ex[1]))
        xf = XLSX.newxlsx()
        s = xf[1]
        XLSX.setFormula(s, "D2", "=SEQUENCE(3)")
        s["B2"] = XLSX.CellImage(pics["B2"].bytes; alt="Logo\n")
        s["B3"] = XLSX.CellImage(pics["B3"].bytes)
        s["B4"] = XLSX.CellImage(pics["B2"].bytes)
        s["C4"] = XLSX.CellImage(pics["C4"].bytes)
        # The same rich values, structures and metadata as Excel's, apart from its web picture
        @test els(part(xf, XLSX.REL_RD_RICHVALUE), "rv") == els(part(ex, XLSX.REL_RD_RICHVALUE), "rv")[1:4]
        @test els(part(xf, XLSX.REL_RD_STRUCTURE), "s") == els(part(ex, XLSX.REL_RD_STRUCTURE), "s")[1:2]
        @test length(XLSX.elements_with_tag(part(xf, XLSX.REL_RICHVALUEREL), "rel")) == 3   # B2's picture shared
        meta, exmeta = part(xf, XLSX.REL_SHEET_METADATA), part(ex, XLSX.REL_SHEET_METADATA)
        for tag in ("metadataTypes", "cellMetadata")
            @test els(meta, tag) == els(exmeta, tag)
        end
        @test els(XLSX.first_element_with_tag(meta, "valueMetadata"), "bk") ==
              els(XLSX.first_element_with_tag(exmeta, "valueMetadata"), "bk")[1:4]
        @test XML.write(part(xf, XLSX.REL_RD_TYPES)) == XML.write(part(ex, XLSX.REL_RD_TYPES))
        @test occursin("<c r=\"B2\" t=\"e\" vm=\"1\"><v>#VALUE!</v></c>", sheet_xml(xf))
        r = roundtrip(xf)
        @test XLSX.getCellImages(r[1]) == [(cell=c, image=XLSX.CellImage(pics[c].bytes, pics[c].format, c == "B2" ? "Logo\n" : "", nothing))
                                           for c in ("B2", "B3", "B4", "C4")]
        SAVE_FILES && save_outfile(xf)
    end

    @testset "dynamic arrays share the metadata" begin
        # A dynamic-array formula after a picture: its `cm` must point at its own record
        xf = XLSX.newxlsx()
        s = xf[1]
        s["B2"] = XLSX.CellImage(png)
        XLSX.setFormula(s, "D2", "=SEQUENCE(3)")
        XLSX.setFormula(s, "E2", "=SEQUENCE(2)")                                   # one record for all
        meta = part(xf, XLSX.REL_SHEET_METADATA)
        @test [XLSX.get_attr(t, "name") for t in XLSX.xml_elements(XLSX.first_element_with_tag(meta, "metadataTypes"))] == ["XLRICHVALUE", "XLDAPR"]
        @test els(XLSX.first_element_with_tag(meta, "cellMetadata"), "bk") == ["<bk><rc t=\"2\" v=\"0\"/></bk>"]
        @test XLSX.getcell(s, "D2").meta == 1 && XLSX.getcell(s, "E2").meta == 1
        @test length(XLSX.elements_with_tag(meta, "futureMetadata")) == 2
        SAVE_FILES && save_outfile(xf)
    end

    @testset "writing them" begin
        xf = XLSX.newxlsx()
        s = xf[1]
        info = XLSX.addImage(s, "B2", joinpath(data_directory, "track_start.png"); inCell=true, alt="Start")
        @test info == (sheet="Sheet1", media_name="image1.png", from="B2", to="B2")
        @test XLSX.getCellImage(s, "B2") == XLSX.CellImage(png, :png, "Start", nothing)
        @test XLSX.addImage(s, 3, 2, IOBuffer(png); inCell=true).media_name == "image1.png"   # shared
        @test isempty(XLSX.getImages(s))                                            # not floating
        s["D2:E3"] = XLSX.CellImage(png)                                           # every cell of a range
        @test length(XLSX.getCellImages(s)) == 6
        @test count(k -> startswith(k, "xl/media/"), keys(xf.binary_data)) == 1
        alt = XLSX.CellImage(IOBuffer(png); alt=" spaced ")
        s["F2"] = alt
        @test occursin("<v xml:space=\"preserve\"> spaced </v>", XML.write(part(xf, XLSX.REL_RD_RICHVALUE)))
        @test XLSX.getCellImage(roundtrip(xf)[1], "F2").alt == " spaced "

        XLSX.setFont(s, "B2"; bold=true)
        s["B2"] = XLSX.CellImage(png)                                              # keeps the cell's format
        @test haskey(XLSX.getFont(s, "B2").font, "b")
        s["B2"] = 42                                                               # a value replaces it
        @test isnothing(XLSX.getCellImage(s, "B2")) && s["B2"] == 42

        @test_throws XLSX.XLSXError XLSX.addImage(s, "B2:C3", IOBuffer(png); inCell=true)
        @test_throws XLSX.XLSXError XLSX.addImage(s, "B2", IOBuffer(png); inCell=true, size=(10, 10))
        @test_throws XLSX.XLSXError XLSX.addImage(s, "B2", IOBuffer(png); alt="x")      # alt is for inCell
        @test_throws XLSX.XLSXError s["B2"] = XLSX.CellImage(UInt8[], :png, "", url)    # a web picture
        @test_throws XLSX.XLSXError XLSX.CellImage(UInt8[1, 2, 3])                      # not a picture
        ro = XLSX.readxlsx(excel_file)
        @test_throws XLSX.XLSXError ro[1]["A1"] = XLSX.CellImage(png)
        SAVE_FILES && save_outfile(xf)
    end

    @testset "deleting" begin
        xf = XLSX.newxlsx()
        s = xf[1]
        XLSX.addImage(s, "B2", IOBuffer(png); inCell=true)
        XLSX.addImage(s, "B2", IOBuffer(png))                                      # floating, from the same cell
        XLSX.setFont(s, "B2"; italic=true)
        @test XLSX.deleteImage(s, "B2") == 2
        @test isempty(XLSX.getCellImages(s)) && isempty(XLSX.getImages(s))
        @test haskey(XLSX.getFont(s, "B2").font, "i")                              # its format is kept
        @test XLSX.deleteImage(s, "B2") == 0
        s["C2"] = XLSX.CellImage(png)
        s["C2"] = missing
        @test isempty(XLSX.getCellImages(roundtrip(xf)[1]))   # an unused rich value: Excel drops it when it saves
    end

    @testset "tables and sheet edits" begin
        xf = XLSX.newxlsx()
        s = xf[1]
        XLSX.writetable!(s, [["a", "b"], [XLSX.CellImage(png), missing]], ["name", "picture"])
        @test [c.cell for c in XLSX.getCellImages(s)] == ["B2"]
        @test isequal(XLSX.gettable(roundtrip(xf)[1]).data[2], [missing, missing])    # pictures aren't data
        XLSX.copysheet!(s, "Copy")
        XLSX.insertRows(xf["Copy"], 1; count=2)
        @test [c.cell for c in XLSX.getCellImages(xf["Copy"])] == ["B4"]
        @test XLSX.getcell(xf["Copy"], "B4").value == XLSX.getcell(s, "B2").value    # one rich value, as Excel shares it
        XLSX.deleteRows(s, 2)
        @test isempty(XLSX.getCellImages(s))
        @test_throws XLSX.XLSXError XLSX.writetable(IOBuffer(), [[XLSX.CellImage(png)]], ["p"]; streaming=true)
    end
end
