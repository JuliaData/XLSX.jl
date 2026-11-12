@testset "document properties" begin
    function roundtrip(xf)
        io = IOBuffer()
        XLSX.writexlsx(io, xf)
        return XLSX.readxlsx(IOBuffer(take!(io)))
    end
    root(xf, path) = XLSX.xml_root_element(xf.data[path])
    tags(xf, path) = [XML.tag(c) for c in XLSX.xml_elements(root(xf, path))]
    excel_file() = joinpath(data_directory, "docprops_excel.xlsx")

    @testset "reading Excel's" begin
        ex = XLSX.readxlsx(excel_file())                      # read-only
        p = XLSX.getDocumentProperties(ex)
        @test (p.title, p.subject, p.creator, p.keywords, p.description) == ("Sales & <report>", "Q3", "Ann Author", "sales; q3", "Line one")
        @test (p.category, p.contentStatus, p.company, p.manager) == ("Finance", "Draft", "Acme Ltd", "Mo Manager")
        @test p.created == DateTime(2024, 1, 2, 3, 4, 5)
        @test p.modified isa DateTime && p.lastModifiedBy isa String
        c = XLSX.getCustomProperties(ex)
        @test collect(c) == ["Project" => "Apollo", "Count" => 42, "Ratio" => 0.25, "Approved" => true, "Due" => DateTime(2025, 6, 29, 23)]
        @test XLSX.getCustomProperty(ex, "count") == 42                # names ignore case, as in Excel
        @test isnothing(XLSX.getCustomProperty(ex, "Missing"))
        @test XLSX.getCustomProperties(XLSX.openxlsx(excel_file(); enable_cache=false)) == c
    end

    @testset "writing as Excel writes" begin
        ex = XLSX.readxlsx(excel_file())
        xf = XLSX.newxlsx()
        @test isempty(XLSX.getCustomProperties(xf))
        XLSX.setDocumentProperties(xf; title="Sales & <report>", subject="Q3", creator="Ann Author", keywords="sales; q3",
                                   description="Line one", category="Finance", contentStatus="Draft",
                                   created=DateTime(2024, 1, 2, 3, 4, 5), company="Acme Ltd", manager="Mo Manager")
        for (n, v) in ("Project" => "Apollo", "Count" => 42, "Ratio" => 0.25, "Approved" => true, "Due" => DateTime(2025, 6, 29, 23))
            XLSX.setCustomProperty(xf, n, v)
        end
        @test tags(xf, "docProps/core.xml") == tags(ex, "docProps/core.xml")    # Excel's order
        texts(x, path) = [(XML.tag(c), XLSX._element_text(c), XLSX._atts_vector(c)) for c in XLSX.xml_elements(root(x, path))
                          if XLSX.localname(c) ∉ ("lastModifiedBy", "modified")]
        @test texts(xf, "docProps/core.xml") == texts(ex, "docProps/core.xml")
        app_tags(x) = filter(t -> t ∉ ("TotalTime",), tags(x, "docProps/app.xml"))  # Excel adds the time spent editing
        @test app_tags(xf) == app_tags(ex)
        @test XML.write(root(xf, "docProps/custom.xml")) == XML.write(root(ex, "docProps/custom.xml"))
        rel = only(r for r in xf.relationships if r.Type == XLSX._CUSTOM_PART.rel)
        @test rel.Target == "docProps/custom.xml"
        @test "docProps/custom.xml" in XLSX.parts_with_content_type(xf, XLSX._CUSTOM_PART.content_type)

        r = roundtrip(xf)
        @test XLSX.getCustomProperties(r) == XLSX.getCustomProperties(ex)
        p, q = XLSX.getDocumentProperties(r), XLSX.getDocumentProperties(ex)
        @test all(getfield(p, k) == getfield(q, k) for k in keys(p) if k ∉ (:lastModifiedBy, :modified))
        SAVE_FILES && save_outfile(xf)
    end

    @testset "changing a file's" begin
        xf = XLSX.openxlsx(excel_file(); mode="rw")
        XLSX.setDocumentProperties(xf; subject="", title="New title", modified=Date(2026, 3, 1))
        p = XLSX.getDocumentProperties(xf)
        @test isnothing(p.subject) && p.title == "New title" && p.modified == DateTime(2026, 3, 1)
        @test !any(t -> endswith(t, ":subject"), tags(xf, "docProps/core.xml"))
        XLSX.setDocumentProperties(xf; subject="Back")                 # where Excel puts it
        @test tags(xf, "docProps/core.xml")[1:3] == ["dc:title", "dc:subject", "dc:creator"]

        XLSX.setCustomProperty(xf, "COUNT", 7)                          # replaces, keeping its name and pid
        @test XLSX.getCustomProperties(xf)["Count"] == 7 && !haskey(XLSX.getCustomProperties(xf), "COUNT")
        @test XLSX.xml_elements(root(xf, "docProps/custom.xml"))[2]["pid"] == "3"
        XLSX.setCustomProperty(xf, "Big", 3_000_000_000)                # beyond Excel's whole numbers
        @test XLSX.getCustomProperty(xf, "Big") === 3.0e9
        @test last(XLSX.xml_elements(root(xf, "docProps/custom.xml")))["pid"] == "7"
        XLSX.setCustomProperty(xf, "Day", Date(2026, 1, 15))
        @test XLSX.getCustomProperty(xf, "Day") == DateTime(2026, 1, 15)
        @test XLSX.deleteCustomProperty(xf, "ratio")
        @test !XLSX.deleteCustomProperty(xf, "ratio")
        @test collect(keys(XLSX.getCustomProperties(xf))) == ["Project", "Count", "Approved", "Due", "Big", "Day"]

        r = roundtrip(xf)
        @test XLSX.getCustomProperties(r) == XLSX.getCustomProperties(xf)
        @test XLSX.getDocumentProperties(r) == XLSX.getDocumentProperties(xf)
    end

    @testset "other programs' properties are kept" begin
        xf = XLSX.openxlsx(excel_file(); mode="rw")
        rt = root(xf, "docProps/custom.xml")
        label = """<property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="20" name="MSIP_Label_0f1e_Enabled"><vt:lpwstr>true</vt:lpwstr></property>"""
        listp = """<property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="21" name="Reviewers"><vt:vector size="2" baseType="lpwstr"><vt:lpwstr>A</vt:lpwstr><vt:lpwstr>B</vt:lpwstr></vt:vector></property>"""
        for s in (label, listp)
            el = XLSX.xml_root_element(parse("<Properties xmlns=\"$(XLSX._NS_CUSTOM)\" xmlns:vt=\"$(XLSX.DOCPROPS_VTYPES_NS)\">$s</Properties>", XML.Node))[1]
            push!(rt, el)
        end
        before = [XML.write(c) for c in XLSX.xml_elements(root(xf, "docProps/custom.xml")) if c["name"] in ("MSIP_Label_0f1e_Enabled", "Reviewers")]
        XLSX.setCustomProperty(xf, "Project", "Gemini")
        XLSX.setCustomProperty(xf, "New", 1)
        XLSX.deleteCustomProperty(xf, "Due")
        r = roundtrip(xf)
        after = [XML.write(c) for c in XLSX.xml_elements(root(r, "docProps/custom.xml")) if c["name"] in ("MSIP_Label_0f1e_Enabled", "Reviewers")]
        @test after == before
        @test XLSX.getCustomProperty(r, "Reviewers") == ["A", "B"]      # a list is read as a vector
        @test XLSX.getCustomProperty(r, "New") == 1
        @test last(XLSX.xml_elements(root(r, "docProps/custom.xml")))["pid"] == "22"        # after the highest pid
    end

    @testset "a file without the parts" begin
        xf = XLSX.newxlsx()
        for part in (XLSX._CORE_PART, XLSX._APP_PART)                   # some writers leave them out
            rels = XLSX.get_package_relationship_root(xf)
            deleteat!(rels.children, findfirst(c -> XML.nodetype(c) == XML.Element && c["Type"] == part.rel, rels.children))
            filter!(r -> r.Type != part.rel, xf.relationships)
            delete!(xf.data, part.path)
            delete!(xf.files, part.path)
        end
        @test all(isnothing, XLSX.getDocumentProperties(xf))
        XLSX.setDocumentProperties(xf; title="T", company="C")
        r = roundtrip(xf)
        @test XLSX.getDocumentProperties(r).title == "T" && XLSX.getDocumentProperties(r).company == "C"
        @test count(rr -> rr.Type in (XLSX._CORE_PART.rel, XLSX._APP_PART.rel), r.relationships) == 2
    end

    @testset "dates" begin
        @test XLSX._parse_w3cdtf("2024-01-02T03:04:05Z") == DateTime(2024, 1, 2, 3, 4, 5)
        @test XLSX._parse_w3cdtf("2024-01-02T03:04:05.25Z") == DateTime(2024, 1, 2, 3, 4, 5, 250)
        @test XLSX._parse_w3cdtf("2024-01-02T03:04:05+01:30") == DateTime(2024, 1, 2, 1, 34, 5)
        @test XLSX._parse_w3cdtf("2024-01-02T03:04-02:00") == DateTime(2024, 1, 2, 5, 4)
        @test XLSX._parse_w3cdtf("2024-01-02") == DateTime(2024, 1, 2)
        @test isnothing(XLSX._parse_w3cdtf("2024-13-02")) && isnothing(XLSX._parse_w3cdtf("yesterday"))
        @test XLSX._w3cdtf(DateTime(2024, 1, 2, 3, 4, 5, 999)) == "2024-01-02T03:04:05Z"
        xf = XLSX.newxlsx()
        XLSX._set_property_elements!(xf, "docProps/core.xml", [XLSX._CORE_PROPERTIES[:created] => ("not a date", true)], XLSX._CORE_ORDER)
        @test XLSX.getDocumentProperties(xf).created == "not a date"   # kept as written
    end

    @testset "errors" begin
        xf = XLSX.newxlsx()
        @test_throws XLSX.XLSXError XLSX.setDocumentProperties(xf)
        @test_throws XLSX.XLSXError XLSX.setDocumentProperties(xf; author="A")
        @test_throws XLSX.XLSXError XLSX.setDocumentProperties(xf; title=1)
        @test_throws XLSX.XLSXError XLSX.setDocumentProperties(xf; created="2024-01-01")
        @test_throws XLSX.XLSXError XLSX.setCustomProperty(xf, "", 1)
        @test_throws XLSX.XLSXError XLSX.setCustomProperty(xf, "n"^256, 1)
        @test_throws XLSX.XLSXError XLSX.setCustomProperty(xf, "x", NaN)
        @test_throws XLSX.XLSXError XLSX.setCustomProperty(xf, "x", [1, 2])
        @test isempty(XLSX.getCustomProperties(xf))                     # nothing was added
        XLSX.setCustomProperty(xf, "n"^255, 1)
        r = roundtrip(xf)
        @test XLSX.getCustomProperty(r, "n"^255) == 1
        @test_throws XLSX.XLSXError XLSX.setDocumentProperties(r; title="T")
        @test_throws XLSX.XLSXError XLSX.setCustomProperty(r, "x", 1)
        @test_throws XLSX.XLSXError XLSX.deleteCustomProperty(r, "x")
    end
end
