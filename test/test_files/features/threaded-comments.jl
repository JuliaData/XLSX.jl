@testset "Threaded comments" begin
    REL_THREADED = "http://schemas.microsoft.com/office/2017/10/relationships/threadedComment"
    REL_PERSON   = "http://schemas.microsoft.com/office/2017/10/relationships/person"
    PATH         = "xl/threadedComments/threadedComment1.xml"
    fixture      = joinpath(data_directory, "threaded_comments.xlsx")

    fresh() = (xf = XLSX.newxlsx("Data"); (xf, xf["Data"]))
    root(xf, path) = XLSX.xml_root_element(xf.data[path])
    els(xf, path, tag) = XLSX.elements_with_tag(root(xf, path), tag)
    rel_types(xf, path) = [XML.attributes(n)["Type"] for n in XML.children(root(xf, path)) if XML.nodetype(n) == XML.Element]
    function roundtrip(xf; kw...)
        io = IOBuffer()
        XLSX.writexlsx(io, xf)
        return XLSX.openxlsx(IOBuffer(take!(io)); kw...)
    end
    when = Dates.DateTime(2026, 11, 13, 10, 15, 30, 250)

    @testset "write and read back" begin
        xf, s = fresh()
        t = XLSX.addThreadedComment(s, "B3", "Is this final?"; author="Tim", date=when)
        @test t == (sheet="Data", ref="B3", author="Tim", text="Is this final?", date=when, resolved=false,
                    replies=NamedTuple{(:author, :text, :date), Tuple{String, String, Union{Nothing,Dates.DateTime}}}[])
        t = XLSX.addThreadedReply(s, 3, 2, "Yes."; author="Ann", date=when + Dates.Hour(1))
        @test [(r.author, r.text, r.date) for r in t.replies] == [("Ann", "Yes.", when + Dates.Hour(1))]
        XLSX.addThreadedReply(s, XLSX.CellRef("B3"), "Thanks"; author="Tim", date=when)
        XLSX.addThreadedComment(s, 1, 1, "Rename?"; author="Ann")
        @test isnothing(XLSX.setThreadedCommentResolved(s, "A1"))
        @test XLSX.getThreadedComment(s, "A1").resolved
        XLSX.setThreadedCommentResolved(s, "A1", false)
        @test !XLSX.getThreadedComment(s, "A1").resolved
        XLSX.setThreadedCommentResolved(s, 1, 1, true)
        @test XLSX.getThreadedComment(s, "A1").date isa Dates.DateTime   # now, by default

        for enable_cache in (true, false)
            r = roundtrip(xf; enable_cache)
            ts = XLSX.getThreadedComments(r)
            @test [(c.ref, c.author, c.text, c.resolved, [x.text for x in c.replies]) for c in ts] ==
                  [("A1", "Ann", "Rename?", true, String[]), ("B3", "Tim", "Is this final?", false, ["Yes.", "Thanks"])]
            @test ts[2].date == when
        end
        SAVE_FILES && save_outfile(xf)
    end

    @testset "as Excel writes them" begin
        xf, s = fresh()
        XLSX.addThreadedComment(s, "B3", "Second"; author="Tim", date=when)
        XLSX.addThreadedComment(s, "A1", "First"; author="Ann", date=when)
        XLSX.addThreadedReply(s, "B3", "A reply"; author="Ann", date=when)

        # The parts, their relationships and content types
        @test REL_THREADED in rel_types(xf, "xl/worksheets/_rels/sheet1.xml.rels")
        @test REL_PERSON in rel_types(xf, "xl/_rels/workbook.xml.rels")
        ct = XML.write(xf.data["[Content_Types].xml"])
        @test occursin("/$PATH\" ContentType=\"application/vnd.ms-excel.threadedcomments+xml\"", ct)
        @test occursin("/xl/persons/person.xml\" ContentType=\"application/vnd.ms-excel.person+xml\"", ct)

        # Threads in cell order, replies after their thread, people by name
        nodes = els(xf, PATH, "threadedComment")
        @test [n["ref"] for n in nodes] == ["A1", "B3", "B3"]
        @test nodes[3]["parentId"] == nodes[2]["id"] && !haskey(XML.attributes(nodes[2]), "parentId")
        @test collect(keys(XML.attributes(nodes[3]))) == ["ref", "dT", "personId", "id", "parentId"]
        @test nodes[2]["dT"] == "2026-11-13T10:15:30.25"
        people = els(xf, "xl/persons/person.xml", "person")
        @test [p["displayName"] for p in people] == ["Ann", "Tim"]
        @test all(p -> p["providerId"] == "None" && p["userId"] == p["displayName"], people)
        @test nodes[1]["personId"] == nodes[3]["personId"] == people[1]["id"]   # one person per name
        @test length(unique(n["id"] for n in nodes)) == 3

        # Each thread has a note for older versions of Excel, tied to it by the thread's id
        notes = XLSX.getComments(s)
        @test [c.ref for c in notes] == ["B3", "A1"]
        @test notes[1].author == "tc=" * nodes[2]["id"]
        note = only(n for n in XLSX.elements_with_tag(XLSX.first_element_with_tag(root(xf, "xl/comments1.xml"), "commentList"), "comment")
                    if n["ref"] == "B3")
        @test note["xr:uid"] == nodes[2]["id"]
        @test root(xf, "xl/comments1.xml")["mc:Ignorable"] == "xr"
        @test occursin("<x:Row>2</x:Row>", XLSX._vml_string(xf, "xl/drawings/vmlDrawing1.vml"))   # and its box

        # The note's text is Excel's, as in its own file
        excel = XLSX.readxlsx(fixture)[1]
        xf2, s2 = fresh()
        XLSX.addThreadedComment(s2, "A1", "First thread"; author="Tim Gebbels")
        XLSX.addThreadedReply(s2, "A1", "Reply one"; author="Tim Gebbels")
        XLSX.addThreadedReply(s2, "A1", "Reply two"; author="Tim Gebbels")
        @test XLSX.getComment(s2, "A1").text == XLSX.getComment(excel, "A1").text
        SAVE_FILES && save_outfile(xf)
    end

    @testset "refused" begin
        xf, s = fresh()
        XLSX.addComment(s, "C1", "a note")
        XLSX.addThreadedComment(s, "A1", "hi"; author="Tim")
        @test_throws XLSX.XLSXError XLSX.addThreadedComment(s, "A1", "again"; author="Tim")   # has a thread
        @test_throws XLSX.XLSXError XLSX.addThreadedComment(s, "C1", "hi"; author="Tim")      # has a note
        @test_throws XLSX.XLSXError XLSX.addThreadedComment(s, "B1", ""; author="Tim")
        @test_throws XLSX.XLSXError XLSX.addThreadedComment(s, "B1", "hi"; author="")
        @test_throws UndefKeywordError XLSX.addThreadedComment(s, "B1", "hi")
        @test_throws XLSX.XLSXError XLSX.addThreadedComment(s, "B1:B2", "hi"; author="Tim")
        @test_throws XLSX.XLSXError XLSX.addThreadedReply(s, "B1", "no thread"; author="Tim")
        @test_throws XLSX.XLSXError XLSX.setThreadedCommentResolved(s, "B1")
        @test_throws XLSX.XLSXError XLSX.addComment(s, "A1", "over a thread")
        @test_throws XLSX.XLSXError XLSX.deleteComment(s, "A1")

        r = roundtrip(xf)
        @test_throws XLSX.XLSXError XLSX.addThreadedComment(r[1], "B1", "hi"; author="Tim")
        @test_throws XLSX.XLSXError XLSX.addThreadedReply(r[1], "A1", "hi"; author="Tim")
        @test_throws XLSX.XLSXError XLSX.setThreadedCommentResolved(r[1], "A1")
        @test_throws XLSX.XLSXError XLSX.deleteThreadedComment(r[1], "A1")

        s["A5"] = 1
        c = XLSX.Charts.addChart(xf, :column; title="t")
        XLSX.Charts.addSeries(c, "Data!A5:A5")
        cs = xf[XLSX.sheetnames(xf)[end]]
        @test_throws XLSX.XLSXError XLSX.addThreadedComment(cs, "A1", "hi"; author="Tim")
        @test isempty(XLSX.getThreadedComments(cs))
    end

    @testset "delete" begin
        xf, s = fresh()
        XLSX.addComment(s, "C1", "a note")
        XLSX.addThreadedComment(s, "A1", "one"; author="Tim")
        XLSX.addThreadedReply(s, "A1", "reply"; author="Ann")
        XLSX.addThreadedComment(s, "B2", "two"; author="Tim")
        @test !XLSX.deleteThreadedComment(s, "D4")
        @test XLSX.deleteThreadedComment(s, 1, 1)
        @test [c.ref for c in XLSX.getThreadedComments(s)] == ["B2"]
        @test [c.ref for c in XLSX.getComments(s)] == ["C1", "B2"]   # its note gone, the others kept
        @test !occursin(r"<x:Row>0</x:Row>\s*<x:Column>0</x:Column>", XLSX._vml_string(xf, "xl/drawings/vmlDrawing1.vml"))

        # The last thread gone: its part too; a cell can then have a note again
        @test XLSX.deleteThreadedComment(s, "B2")
        @test !(REL_THREADED in rel_types(xf, "xl/worksheets/_rels/sheet1.xml.rels"))
        @test !haskey(xf.data, PATH)
        @test !occursin(PATH, XML.write(xf.data["[Content_Types].xml"]))
        @test XLSX.addComment(s, "A1", "now a note").ref == "A1"
        @test XLSX.deleteComment(s, "C1") && XLSX.deleteComment(s, "A1")
        @test isempty(XLSX.getComments(s))

        # Starting again makes the part again
        XLSX.addThreadedComment(s, "A1", "again"; author="Tim")
        @test XLSX.getThreadedComment(roundtrip(xf)[1], "A1").text == "again"
    end

    @testset "in Excel's file" begin
        xf = XLSX.opentemplate(fixture)
        s = xf[1]
        people = length(els(xf, "xl/persons/person.xml", "person"))
        XLSX.addThreadedReply(s, "A1", "Reply three"; author="Tim Gebbels")
        @test length(els(xf, "xl/persons/person.xml", "person")) == people   # the same person
        @test XLSX.deleteThreadedComment(s, "B2")
        XLSX.addThreadedComment(s, "D4", "New"; author="Someone Else")
        XLSX.setThreadedCommentResolved(s, "A1")

        r = roundtrip(xf)
        ts = XLSX.getThreadedComments(r)
        @test [(c.ref, c.author, c.resolved, [x.text for x in c.replies]) for c in ts] ==
              [("A1", "Tim Gebbels", true, ["Reply one", "Reply two", "Reply three"]), ("D4", "Someone Else", false, String[])]
        @test [c.ref for c in XLSX.getComments(r)] == ["A1", "C3", "D4"]   # Excel's plain note on C3 kept
        @test endswith(XLSX.getComment(r[1], "A1").text, "Reply:\n    Reply two\nReply:\n    Reply three")
        SAVE_FILES && save_outfile(xf)
    end
end
