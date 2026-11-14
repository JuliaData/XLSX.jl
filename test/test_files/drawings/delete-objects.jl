@testset "deleteImage, deleteChart, threads" begin
    data_dir = joinpath(dirname(pathof(XLSX)), "..", "test", "data")
    png = joinpath(data_dir, "track_start.png")
    jpg = joinpath(data_dir, "track_start.jpg")
    function fresh()
        xf = XLSX.newxlsx("Data")
        s = xf["Data"]
        s[1:10, 1:2] = reshape(1:20, 10, 2)
        return xf, s
    end
    sheet_root(xf, n=1) = XLSX.xml_root_element(xf.data["xl/worksheets/sheet$n.xml"])
    element(xf, tag, n=1) = XLSX.first_element_with_tag(sheet_root(xf, n), tag)
    function roundtrip_bytes(xf)
        io = IOBuffer()
        XLSX.writexlsx(io, xf)
        return take!(io)
    end
    roundtrip(xf) = XLSX.readxlsx(IOBuffer(roundtrip_bytes(xf)))
    zip_names(bytes) = XLSX.ZipArchives.zip_names(XLSX.ZipArchives.ZipReader(bytes))
    media(xf) = sort([k for k in keys(xf.binary_data) if startswith(k, "xl/media/")])

    @testset "deleteImage" begin
        xf, s = fresh()
        XLSX.addImage(s, "D2:F6", png)
        XLSX.addImage(s, "H2", jpg)
        @test length(XLSX.getImages(s)) == 2
        @test length(media(xf)) == 2

        @test XLSX.deleteImage(s, "D2") == 1
        imgs = XLSX.getImages(s)
        @test length(imgs) == 1 && imgs[1].from == "H2"
        @test length(media(xf)) == 1 && !any(endswith(".png"), media(xf))   # its picture went too
        @test XLSX.deleteImage(s, "D2") == 0                                 # nothing there now
        @test XLSX.deleteImage(s, "A1") == 0

        r = roundtrip(xf)
        @test length(XLSX.getImages(r["Data"])) == 1

        @test XLSX.deleteImage(s, only(XLSX.getImages(s))) == 1               # the last one
        @test isempty(XLSX.getImages(s))
        @test isempty(media(xf))
        @test isnothing(element(xf, "drawing"))                               # the drawing went with it
        @test isnothing(XLSX._drawing_path_for_sheet(xf, "xl/worksheets/sheet1.xml"))
        bytes = roundtrip_bytes(xf)
        @test !any(startswith("xl/drawings/"), zip_names(bytes))
        @test !any(startswith("xl/media/"), zip_names(bytes))
        @test XLSX.readxlsx(IOBuffer(bytes))["Data"][1, 1] == 1
        SAVE_FILES && save_outfile(xf)
    end

    @testset "deleteImage: several in one cell, and arguments" begin
        xf, s = fresh()
        XLSX.addImage(s, "D2", png)
        XLSX.addImage(s, "D2", jpg)
        @test XLSX.deleteImage(s, XLSX.CellRef("D2")) == 2
        @test isempty(XLSX.getImages(s))
        @test_throws XLSX.XLSXError XLSX.deleteImage(s, "D2:E3")
        @test_throws XLSX.XLSXError XLSX.deleteImage(s, "nonsense")

        xf, s = fresh()
        XLSX.addsheet!(xf, "Other")
        XLSX.addImage(xf["Other"], "B2", png)
        @test_throws XLSX.XLSXError XLSX.deleteImage(s, only(XLSX.getImages(xf["Other"])))   # on another sheet

        r = roundtrip(xf)
        @test_throws XLSX.XLSXError XLSX.deleteImage(r["Other"], "B2")       # read-only
    end

    @testset "deleteImage: a picture another sheet still shows" begin
        xf, s = fresh()
        XLSX.addImage(s, "D2", png)
        c = XLSX.copysheet!(s, "Copy")                                       # shares the picture
        @test length(media(xf)) == 1
        @test XLSX.deleteImage(c, "D2") == 1
        @test length(media(xf)) == 1                                          # still shown on Data
        r = roundtrip(xf)
        @test length(XLSX.getImages(r["Data"])) == 1 && isempty(XLSX.getImages(r["Copy"]))
    end

    @testset "deleteChart" begin
        xf, s = fresh()
        c1 = XLSX.Charts.addChart(s, :line; anchor="D2:J12")
        XLSX.Charts.addSeries(c1, "Data!A1:A10")
        c2 = XLSX.Charts.addChart(s, :column; anchor="D14:J24")
        XLSX.Charts.addSeries(c2, "Data!B1:B10")
        XLSX.addImage(s, "L2", png)
        @test length(XLSX.Charts.getCharts(s)) == 2

        XLSX.Charts.deleteChart(only(c for c in XLSX.Charts.getCharts(s) if c.from == "D2"))
        cs = XLSX.Charts.getCharts(s)
        @test length(cs) == 1 && only(cs).from == "D14"
        @test length(XLSX.parts_with_content_type(xf, XLSX.MIME_CHART)) == 1   # its part went
        @test length(XLSX.getImages(s)) == 1                                   # the image stays
        @test length(XLSX.Charts.getChartSeries(only(cs))) == 1

        r = roundtrip(xf)
        @test length(XLSX.Charts.getCharts(r["Data"])) == 1

        XLSX.Charts.deleteChart(only(XLSX.Charts.getCharts(s)))
        XLSX.deleteImage(s, "L2")
        @test isempty(XLSX.Charts.getCharts(s))
        @test isnothing(element(xf, "drawing"))
        bytes = roundtrip_bytes(xf)
        @test !any(startswith("xl/charts/"), zip_names(bytes))
        @test isempty(XLSX.Charts.getCharts(XLSX.readxlsx(IOBuffer(bytes))))
        SAVE_FILES && save_outfile(xf)
    end

    @testset "deleteChart: chartEx" begin
        xf = XLSX.newxlsx("Data")
        s = xf["Data"]
        s["A1"] = ["Step" "Amount"; "Start" 10; "Up" 5; "Down" -3]
        c = XLSX.Charts.addChartEx(s, :waterfall, "B2:B4"; anchor="D2:K16", categories="A2:A4")
        names() = [k for k in keys(XLSX.get_workbook(xf).workbook_names) if startswith(k, "_xlchart.")]
        XLSX.Charts.deleteChart(only(XLSX.Charts.getCharts(s)))
        @test isempty(XLSX.Charts.getCharts(s))
        @test isempty(XLSX.parts_with_content_type(xf, XLSX.MIME_CHARTEX))
        @test isempty(names())                                                # its hidden names went
        @test isnothing(element(xf, "drawing"))
        @test isempty(XLSX.Charts.getCharts(roundtrip(xf)))
        SAVE_FILES && save_outfile(xf)
    end

    @testset "deleteChart: refusals" begin
        xf, s = fresh()
        c = XLSX.Charts.addChart(xf, :line; sheetname="Trend")
        XLSX.Charts.addSeries(c, "Data!A1:A10")
        @test_throws XLSX.XLSXError XLSX.Charts.deleteChart(only(XLSX.Charts.getCharts(xf)))   # a chartsheet
        c = XLSX.Charts.addChart(s, :line; anchor="D2:J12")
        XLSX.Charts.addSeries(c, "Data!A1:A10")
        r = roundtrip(xf)
        @test_throws XLSX.XLSXError XLSX.Charts.deleteChart(only(XLSX.Charts.getCharts(r["Data"])))   # read-only
    end

    @testset "threaded comments" begin
        xf, s = fresh()
        @test isempty(XLSX.getThreadedComments(s))
        # As Excel 365 writes them: the threads, and the people in the workbook.
        thread = "xl/threadedComments/threadedComment1.xml"
        xf.data[thread] = XML.parse("""
            <ThreadedComments xmlns="http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments">
            <threadedComment ref="B3" dT="2026-03-01T09:30:00.25" personId="{P1}" id="{T1}"><text>Is this right?</text></threadedComment>
            <threadedComment ref="B3" dT="2026-03-01T10:00:00.00" personId="{P2}" id="{T2}" parentId="{T1}"><text>Yes, checked.</text></threadedComment>
            <threadedComment ref="B3" dT="2026-03-02T08:00:00.00" personId="{P1}" id="{T3}" parentId="{T1}"><text>Thanks</text></threadedComment>
            <threadedComment ref="D7" dT="2026-03-05T12:00:00" personId="{P2}" id="{T4}" done="1"><text>Done &amp; dusted</text></threadedComment>
            </ThreadedComments>""", XML.Node)
        xf.files[thread] = true
        XLSX.add_part_rel!(xf, "xl/worksheets/sheet1.xml", thread, XLSX.REL_THREADED_COMMENT)
        people = "xl/persons/person.xml"
        xf.data[people] = XML.parse("""
            <personList xmlns="http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments">
            <person displayName="Ana" id="{P1}" userId="ana@example.com" providerId="AD"/>
            <person displayName="Ben" id="{P2}" userId="ben@example.com" providerId="AD"/>
            </personList>""", XML.Node)
        xf.files[people] = true
        XLSX.add_relationship!(XLSX.get_workbook(xf), "persons/person.xml", XLSX.REL_PERSON)

        ts = XLSX.getThreadedComments(s)
        @test length(ts) == 2
        t = XLSX.getThreadedComment(s, "B3")
        @test t.sheet == "Data" && t.author == "Ana" && t.text == "Is this right?"
        @test t.date == XLSX.Dates.DateTime(2026, 3, 1, 9, 30, 0, 250)
        @test !t.resolved
        @test [(r.author, r.text) for r in t.replies] == [("Ben", "Yes, checked."), ("Ana", "Thanks")]
        @test t.replies[1].date == XLSX.Dates.DateTime(2026, 3, 1, 10)
        d = XLSX.getThreadedComment(s, 7, 4)
        @test d.text == "Done & dusted" && d.resolved && isempty(d.replies) && d.author == "Ben"
        @test d.date == XLSX.Dates.DateTime(2026, 3, 5, 12)
        @test isnothing(XLSX.getThreadedComment(s, "A1"))
        @test XLSX.getThreadedComments(xf) == ts

        r = roundtrip(xf)                                                     # read-only
        @test XLSX.getThreadedComments(r) == ts
        XLSX.openxlsx(IOBuffer(roundtrip_bytes(xf)); enable_cache=false) do f
            @test XLSX.getThreadedComment(f["Data"], "D7").text == "Done & dusted"
        end
    end

    @testset "threaded comments: no people part" begin
        xf, s = fresh()
        thread = "xl/threadedComments/threadedComment1.xml"
        xf.data[thread] = XML.parse("""<ThreadedComments xmlns="http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments"><threadedComment ref="A1" personId="{P1}" id="{T1}"><text>hi</text></threadedComment></ThreadedComments>""", XML.Node)
        xf.files[thread] = true
        XLSX.add_part_rel!(xf, "xl/worksheets/sheet1.xml", thread, XLSX.REL_THREADED_COMMENT)
        t = only(XLSX.getThreadedComments(s))
        @test t.author == "" && isnothing(t.date) && t.text == "hi"
    end
end
