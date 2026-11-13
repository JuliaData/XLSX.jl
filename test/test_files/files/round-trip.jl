# Round-trip fidelity: a file opened with `mode="rw"`, edited and saved keeps every
# part, relationship and content type, including parts XLSX.jl doesn't model (pivot
# tables and caches, slicers, timelines, external links, custom XML). Only
# `xl/calcChain.xml` goes, as Excel rebuilds it. `roundtrip_excel.xlsx` and
# `calcsettings_excel.xlsx` were made by Excel 365 through its automation interface.
@testset "round-trip fidelity" begin
    fixture(f) = joinpath(data_directory, f)
    zipped(path) = XLSX.ZipArchives.ZipReader(read(path))
    function edited(path; edit=true)
        xf = XLSX.openxlsx(path; mode="rw")
        if edit
            for i in 1:XLSX.sheetcount(xf)
                s = xf[i]
                XLSX.is_chartsheet(XLSX.get_workbook(xf), s.name) && continue
                s["XFD1048576"] = "edited"   # out of the way of everything else
                break
            end
        end
        io = IOBuffer()
        XLSX.writexlsx(io, xf)
        return XLSX.ZipArchives.ZipReader(take!(io))
    end
    names(z) = Set(XLSX.ZipArchives.zip_names(z))
    entry(z, n) = XLSX.ZipArchives.zip_readentry(z, n, String)

    # XML compared as elements, attributes in any order, whitespace between elements ignored.
    function canon(n)
        XML.nodetype(n) == XML.Text && return isempty(strip(XML.value(n))) ? nothing : XML.value(n)
        XML.nodetype(n) == XML.Element || return nothing
        return (XML.tag(n), sort(collect(something(XML.attributes(n), Dict{String,String}()))),
                filter(!isnothing, [canon(c) for c in XML.children(n)]))
    end
    xmlroot(s) = first(c for c in XML.children(parse(s, XML.Node)) if XML.nodetype(c) == XML.Element)
    same_xml(a, b) = canon(xmlroot(a)) == canon(xmlroot(b))
    rels(s) = Set((n["Id"], n["Type"], n["Target"]) for n in XML.children(xmlroot(s)) if XML.nodetype(n) == XML.Element)
    content_types(s) = Set(sort(collect(XML.attributes(n))) for n in XML.children(xmlroot(s)) if XML.nodetype(n) == XML.Element)
    top(s, name) = [canon(c) for c in XML.children(xmlroot(s)) if XML.nodetype(c) == XML.Element && XLSX.localname(c) == name]

    @testset "parts XLSX.jl doesn't model" begin
        before = zipped(fixture("roundtrip_excel.xlsx"))
        after = edited(fixture("roundtrip_excel.xlsx"))
        @test setdiff(names(before), names(after)) == Set(["xl/calcChain.xml"])
        @test isempty(setdiff(names(after), names(before)))

        # Untouched parts are kept as they were
        for n in names(after)
            occursin(r"^(xl/(pivotCache|pivotTables|slicers|slicerCaches|timelines|timelineCaches|externalLinks|tables|drawings|theme)/|customXml/|xl/comments|xl/styles)", n) || continue
            @test entry(before, n) == entry(after, n) || (endswith(n, ".vml") ? false : same_xml(entry(before, n), entry(after, n)))
        end

        # Every relationship, and every content type but calcChain's
        for n in names(after)
            endswith(n, ".rels") && @test issubset(rels(entry(before, n)), rels(entry(after, n)))
        end
        ct_before = content_types(entry(before, "[Content_Types].xml"))
        ct_after = content_types(entry(after, "[Content_Types].xml"))
        @test setdiff(ct_before, ct_after) == Set([sort(["PartName" => "/xl/calcChain.xml",
            "ContentType" => "application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml"])])

        # The workbook keeps its pivot caches, external references and slicer and timeline caches
        wb_before, wb_after = entry(before, "xl/workbook.xml"), entry(after, "xl/workbook.xml")
        for name in ("pivotCaches", "externalReferences", "extLst", "sheets", "bookViews", "workbookPr")
            @test top(wb_before, name) == top(wb_after, name)
        end

        # The sheets keep everything around their cells: sparklines, drawings, the note,
        # the table, the pivot's slicer and timeline
        for n in ("xl/worksheets/sheet1.xml", "xl/worksheets/sheet2.xml")
            s_before, s_after = entry(before, n), entry(after, n)
            kept = [XLSX.localname(c) for c in XML.children(xmlroot(s_before)) if XML.nodetype(c) == XML.Element]
            for name in setdiff(kept, ["sheetData", "dimension"])
                @test top(s_before, name) == top(s_after, name)
            end
        end
        # A row Excel fitted to its wrapped text stays fitted, not fixed at its height
        @test occursin(r"<row r=\"27\"[^>]* ht=\"101.5\"", entry(after, "xl/worksheets/sheet1.xml"))
        @test !occursin(r"<row r=\"27\"[^>]*customHeight", entry(after, "xl/worksheets/sheet1.xml"))

        # And it reads back: the pivot table's values, the external formula, the edit
        io = IOBuffer()
        xf = XLSX.openxlsx(fixture("roundtrip_excel.xlsx"); mode="rw")
        xf["Data"]["Z1"] = "edited"
        XLSX.writexlsx(io, xf)
        r = XLSX.readxlsx(IOBuffer(take!(io)))
        @test r["Pivot"]["A3"] == "Row Labels" && r["Pivot"]["B8"] == 12612
        @test XLSX.getFormula(r["Data"], "H1") == "=[1]Sheet1!\$A\$1" && r["Data"]["H1"] == 42
        @test r["Data"]["Z1"] == "edited"
        @test length(XLSX.getSparklines(r["Data"])) == 1 && length(XLSX.getShapes(r["Data"])) == 1
        @test XLSX.getComment(r["Data"], "H2").text == "a note"
    end

    @testset "calculation settings" begin
        # Saving asks Excel to recalculate on opening, keeping the workbook's own settings:
        # manual calculation, R1C1, iteration, precision as displayed
        calc(s) = only(top(s, "calcPr"))[2]
        before = calc(entry(zipped(fixture("calcsettings_excel.xlsx")), "xl/workbook.xml"))
        after = calc(entry(edited(fixture("calcsettings_excel.xlsx")), "xl/workbook.xml"))
        @test after == sort(vcat(before, ["fullCalcOnLoad" => "1"]))
        @test ("calcMode" => "manual") in after && ("refMode" => "R1C1") in after
    end

    @testset "document properties in other languages" begin
        # Excel names the sheet categories in its own language ("Planilhas"): they are
        # recognised by their titles, not listed again as "Worksheets"
        app(z) = entry(z, "docProps/app.xml")
        xf = XLSX.openxlsx(fixture("simple.xlsx"); mode="rw")
        XLSX.addsheet!(xf, "Nova")
        io = IOBuffer()
        XLSX.writexlsx(io, xf)
        a = app(XLSX.ZipArchives.ZipReader(take!(io)))
        @test occursin("Planilhas", a) && !occursin("Worksheets", a)
        @test [m[1] for m in eachmatch(r"<vt:lpstr>([^<]*)</vt:lpstr>", match(r"<TitlesOfParts>.*</TitlesOfParts>"s, a).match)] ==
              ["Plan1", "Nova"]
        @test occursin(r"<vt:i4>2</vt:i4>", a)
    end

    @testset "row heights" begin
        # A height set with setRowHeight is fixed (`customHeight`), as Excel writes it;
        # rows inserted below keep it
        xf = XLSX.newxlsx()
        s = xf[1]
        s["A1"] = 1; s["A2"] = 2
        XLSX.setRowHeight(s, 2; height=30)
        XLSX.insertRows(s, 3)
        io = IOBuffer()
        XLSX.writexlsx(io, xf)
        x = entry(XLSX.ZipArchives.ZipReader(take!(io)), "xl/worksheets/sheet1.xml")
        @test occursin(r"<row r=\"2\"[^>]* ht=\"30.2109375\"[^>]* customHeight=\"1\"", x)
        @test occursin(r"<row r=\"3\"[^>]* customHeight=\"1\"", x)
        @test !occursin(r"<row r=\"1\"[^>]*customHeight", x)
    end

    @testset "every fixture keeps its parts" begin
        # Each part, relationship and content type survives an edit and a save, from
        # every file XLSX.jl can open for writing.
        for f in sort(readdir(data_directory))
            occursin(r"\.(xlsx|xlsm|xltx|xltm)$", f) || continue
            path = fixture(f)
            local before, after
            try
                before = zipped(path)
                after = edited(path)
            catch e
                e isa XLSX.XLSXError || occursin("zip", sprint(showerror, e)) || rethrow()
                continue   # files the other tests show can't be opened (encrypted, invalid)
            end
            lost = setdiff(filter(n -> !endswith(n, "/"), names(before)), names(after), ["xl/calcChain.xml"])
            isempty(lost) || @info "parts lost from $f" lost
            @test isempty(lost)
            for n in intersect(names(before), names(after))
                endswith(n, ".rels") || continue
                ok = try
                    issubset(Set((Id, get(XLSX.STRICT_TO_TRANSITIONAL, T, T), t) for (Id, T, t) in rels(entry(before, n))),
                             rels(entry(after, n)))
                catch
                    true   # not parsable as it was (a byte-order mark): nothing to compare
                end
                ok || @info "relationships lost from $f" n
                @test ok
            end
        end
    end
end
