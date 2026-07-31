# Additional testsets for read.jl coverage.
#
# Most of these drive internals directly rather than through a fixture file:
# the uncovered branches are mostly strict-OOXML and malformed-package paths
# that would otherwise need a purpose-built .xlsx each.

const _STRICT_MAIN = "http://purl.oclc.org/ooxml/spreadsheetml/main"
const _TRANS_MAIN  = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
const _CT_SHEET    = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"
const _CT_TEMPLATE = "application/vnd.openxmlformats-officedocument.spreadsheetml.template.main+xml"

_root_of(s::AbstractString) = XLSX.xml_root_element(parse(s, XLSX.XML.Node))

# Locate a child of `root` by localname and an attribute value.
function _find_child(root, tag::AbstractString, attr::AbstractString, val::AbstractString)
    for (i, c) in enumerate(XLSX.XML.children(root))
        XLSX.localname(c) == tag || continue
        XLSX.get_attr(c, attr, "") == val && return (i, c)
    end
    return (nothing, nothing)
end

@testset "read.jl coverage" begin


# ===========================================================================
# Namespace resolution
# ===========================================================================

    @testset "get_default_namespace - single prefixed namespace" begin
        r = _root_of("""<x:workbook xmlns:x="$_TRANS_MAIN"/>""")
        @test XLSX.get_default_namespace(r) == _TRANS_MAIN
    end

    @testset "get_default_namespace - unprefixed default preferred" begin
        r = _root_of("""<workbook xmlns="$_TRANS_MAIN" xmlns:r="http://example.com/r"/>""")
        @test XLSX.get_default_namespace(r) == _TRANS_MAIN
    end

    @testset "get_default_namespace - prefixed spreadsheet ns as fallback" begin
        # No unprefixed default (issues #380/#362/#267/#170)
        r = _root_of("""<x:workbook xmlns:x="$_TRANS_MAIN" xmlns:r="http://example.com/r"/>""")
        @test XLSX.get_default_namespace(r) == _TRANS_MAIN
    end

    @testset "get_default_namespace - none found errors" begin
        r = _root_of("""<thing xmlns:a="http://example.com/a" xmlns:b="http://example.com/b"/>""")
        @test_throws XLSX.XLSXError XLSX.get_default_namespace(r)
    end

    @testset "get_default_namespace_prefix - no namespaces at all" begin
        r = _root_of("<thing/>")
        @test isnothing(XLSX.get_default_namespace_prefix(r))

        r2 = _root_of("""<thing xmlns:a="http://example.com/a" xmlns:b="http://example.com/b"/>""")
        @test isnothing(XLSX.get_default_namespace_prefix(r2))
    end

    @testset "build_ns_dict! - part still held as a raw String" begin
        # Covers the `val isa String` branch: a read part whose data hasn't been
        # parsed yet has its prefix sniffed from the raw text instead.
        xf = XLSX.newxlsx()
        f = "xl/styles.xml"
        @test haskey(xf.files, f)

        xf.data[f] = """<x:styleSheet xmlns:x="$_TRANS_MAIN"/>"""
        delete!(xf.namespace, f)

        XLSX.build_ns_dict!(xf)
        @test xf.namespace[f] == "x"
        @test xf.data[f] isa String   # left unparsed by build_ns_dict!
    end

    @testset "_get_ns_prefix_from_string - default vs prefixed vs absent" begin
        @test isnothing(XLSX._get_ns_prefix_from_string(nothing))
        @test isnothing(XLSX._get_ns_prefix_from_string("""<worksheet xmlns="$_TRANS_MAIN"/>"""))
        @test XLSX._get_ns_prefix_from_string("""<x:worksheet xmlns:x="$_TRANS_MAIN"/>""") == "x"
        @test isnothing(XLSX._get_ns_prefix_from_string("""<worksheet xmlns="http://example.com/other"/>"""))
    end

    @testset "get_sst_prefix - prefixed shared strings namespace" begin
        xf = XLSX.newxlsx()
        sh = xf[1]

        xf.namespace["xl/sharedStrings.xml"] = "x"
        @test XLSX.get_sst_prefix(sh) == "x:"

        xf.namespace["xl/sharedStrings.xml"] = nothing
        @test XLSX.get_sst_prefix(sh) == ""
    end


    # ===========================================================================
    # Strict OOXML detection and conversion
    # ===========================================================================

    @testset "is_strict_ooxml - strict namespace without conformance attribute" begin
        xf = XLSX.newxlsx()
        xf.data["xl/workbook.xml"] = parse(
            """<workbook xmlns="$_STRICT_MAIN"/>""", XLSX.XML.Node)
        @test XLSX.is_strict_ooxml(xf) == true
    end

    @testset "is_strict_ooxml - conformance attribute" begin
        xf = XLSX.newxlsx()
        xf.data["xl/workbook.xml"] = parse(
            """<workbook xmlns="$_TRANS_MAIN" conformance="strict"/>""", XLSX.XML.Node)
        @test XLSX.is_strict_ooxml(xf) == true
    end

    @testset "is_strict_ooxml - falls back to relationship types in _rels/.rels" begin
        xf = XLSX.newxlsx()
        # workbook root gives no hint at all
        xf.data["xl/workbook.xml"] = parse("""<workbook xmlns="$_TRANS_MAIN"/>""", XLSX.XML.Node)
        xf.data["_rels/.rels"] = parse(
            """<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
            <Relationship Id="rId1" Target="xl/workbook.xml"
                Type="http://purl.oclc.org/ooxml/officeDocument/relationships/officeDocument"/>
            </Relationships>""", XLSX.XML.Node)
        @test XLSX.is_strict_ooxml(xf) == true
    end

    @testset "is_strict_ooxml - ordinary transitional file" begin
        xf = XLSX.newxlsx()
        @test XLSX.is_strict_ooxml(xf) == false
    end

    @testset "_strict_to_transitional_node! - remaps and drops conformance" begin
        node = XLSX.XML.Element("workbook"; xmlns=_STRICT_MAIN, conformance="strict")
        XLSX._strict_to_transitional_node!(node, "xl/workbook.xml")

        @test XLSX.get_attr(node, "xmlns") == _TRANS_MAIN
        @test XLSX.get_attr(node, "conformance", "") == ""
    end

    @testset "_strict_to_transitional_node! - node with no attributes is a no-op" begin
        node = XLSX.XML.Element("workbook")
        @test isnothing(XLSX._strict_to_transitional_node!(node, "xl/workbook.xml"))
    end

    @testset "_strict_to_transitional_node! - unknown purl namespace errors" begin
        node = XLSX.XML.Element("thing"; xmlns="http://purl.oclc.org/ooxml/notAThing/main")
        @test_throws XLSX.XLSXError XLSX._strict_to_transitional_node!(node, "xl/thing.xml")
    end

    @testset "convert_strict_to_transitional! - part held as a raw String is parsed" begin
        xf = XLSX.newxlsx()
        f = "xl/styles.xml"
        xf.data[f] = """<styleSheet xmlns="$_STRICT_MAIN"/>"""

        XLSX.convert_strict_to_transitional!(xf, 1)

        @test xf.data[f] isa XLSX.XML.Node
        @test XLSX.get_attr(XLSX.xml_root_element(xf.data[f]), "xmlns") == _TRANS_MAIN
    end

    @testset "convert_strict_to_transitional! - worksheet String substitution (pass 3)" begin
        xf = XLSX.newxlsx()
        f = XLSX.get_relationship_target_by_id("xl", XLSX.get_workbook(xf), xf[1].relationship_id)
        xf.data[f] = """<worksheet xmlns="$_STRICT_MAIN" conformance="strict"><sheetData/></worksheet>"""

        XLSX.convert_strict_to_transitional!(xf, 3)

        data = xf.data[f]
        @test data isa String
        @test occursin(_TRANS_MAIN, data)
        @test !occursin("purl.oclc.org", data)
        @test !occursin("conformance", data)
    end


    # ===========================================================================
    # Content types / workbook parsing
    # ===========================================================================

    @testset "ensure_workbook_is_xlsx! - template type carried on the Default entry" begin
        # Covers the `isnothing(workbook_override)` branch: no Override for
        # /xl/workbook.xml, so the type comes from (and is rewritten on) Default.
        xf = XLSX.newxlsx()
        root = XLSX.xml_root_element(xf.data["[Content_Types].xml"])

        i, _ = _find_child(root, "Override", "PartName", "/xl/workbook.xml")
        isnothing(i) || deleteat!(root.children, i)

        _, def = _find_child(root, "Default", "Extension", "xml")
        @test !isnothing(def)
        def["ContentType"] = _CT_TEMPLATE

        XLSX.ensure_workbook_is_xlsx!(xf)

        @test xf.template_type == XLSX.XLTXTemplate
        _, def2 = _find_child(root, "Default", "Extension", "xml")
        @test XLSX.get_attr(def2, "ContentType") == _CT_SHEET
    end

    @testset "ensure_workbook_is_xlsx! - unknown workbook content type errors" begin
        xf = XLSX.newxlsx()
        root = XLSX.xml_root_element(xf.data["[Content_Types].xml"])
        _, ovr = _find_child(root, "Override", "PartName", "/xl/workbook.xml")
        @test !isnothing(ovr)
        ovr["ContentType"] = "application/vnd.example.not-a-workbook+xml"

        @test_throws XLSX.XLSXError XLSX.ensure_workbook_is_xlsx!(xf)
    end

    @testset "ensure_workbook_is_xlsx! - no content type at all errors" begin
        xf = XLSX.newxlsx()
        root = XLSX.xml_root_element(xf.data["[Content_Types].xml"])

        i, _ = _find_child(root, "Override", "PartName", "/xl/workbook.xml")
        isnothing(i) || deleteat!(root.children, i)
        j, _ = _find_child(root, "Default", "Extension", "xml")
        isnothing(j) || deleteat!(root.children, j)

        @test_throws XLSX.XLSXError XLSX.ensure_workbook_is_xlsx!(xf)
    end

    @testset "check_minimum_requirements - missing mandatory part errors" begin
        xf = XLSX.newxlsx()
        delete!(xf.files, "xl/_rels/workbook.xml.rels")
        @test_throws XLSX.XLSXError XLSX.check_minimum_requirements(xf)
    end

    @testset "parse_workbook! - root element is not <workbook>" begin
        xf = XLSX.newxlsx()
        xf.data["xl/workbook.xml"] = parse("""<notAWorkbook xmlns="$_TRANS_MAIN"/>""", XLSX.XML.Node)
        @test_throws XLSX.XLSXError XLSX.parse_workbook!(xf)
    end

    @testset "parse_workbook! - unsupported node inside <sheets>" begin
        xf = XLSX.newxlsx()
        xroot = XLSX.xml_root_element(XLSX.xmlroot(xf, "xl/workbook.xml"))
        sheets = only(filter(c -> XLSX.localname(c) == "sheets", XLSX.xml_elements(xroot)))
        push!(sheets, XLSX.XML.Element("notASheet"))

        @test_throws XLSX.XLSXError XLSX.parse_workbook!(xf)
    end

    # Find-or-create <workbookPr> and set date1904, so the branch under test is
    # the first workbookPr `parse_workbook!` encounters.
    function _set_date1904!(xf::XLSX.XLSXFile, v::AbstractString)
        xroot = XLSX.xml_root_element(XLSX.xmlroot(xf, "xl/workbook.xml"))
        idx = findfirst(c -> XLSX.localname(c) == "workbookPr", XLSX.XML.children(xroot))
        if isnothing(idx)
            pr = XLSX.XML.Element("workbookPr")
            pushfirst!(xroot.children, pr)
        else
            pr = XLSX.XML.children(xroot)[idx]
        end
        pr["date1904"] = v
        return nothing
    end

    @testset "parse_workbook! - date1904 false forms" begin
        for v in ("0", "false")
            xf = XLSX.newxlsx()
            _set_date1904!(xf, v)
            XLSX.parse_workbook!(xf)
            @test XLSX.get_workbook(xf).date1904 == false
        end
    end

    @testset "parse_workbook! - date1904 true forms" begin
        for v in ("1", "true")
            xf = XLSX.newxlsx()
            _set_date1904!(xf, v)
            XLSX.parse_workbook!(xf)
            @test XLSX.get_workbook(xf).date1904 == true
        end
    end

    @testset "parse_workbook! - unparseable date1904 errors" begin
        xf = XLSX.newxlsx()
        _set_date1904!(xf, "maybe")
        @test_throws XLSX.XLSXError XLSX.parse_workbook!(xf)
    end


    # ===========================================================================
    # parse_defined_name_value
    # ===========================================================================

    @testset "parse_defined_name_value - relative sheet cell reference" begin
        v, isabs = XLSX.parse_defined_name_value("Sheet1!A1")
        @test v == XLSX.SheetCellRef("Sheet1!A1")
        @test isabs == false
    end

    @testset "parse_defined_name_value - relative sheet cell range" begin
        v, isabs = XLSX.parse_defined_name_value("Sheet1!A1:B2")
        @test v == XLSX.SheetCellRange("Sheet1!A1:B2")
        @test isabs == false
    end

    @testset "parse_defined_name_value - absolute forms" begin
        v, isabs = XLSX.parse_defined_name_value("Sheet1!\$A\$1")
        @test v == XLSX.SheetCellRef("Sheet1!A1")
        @test isabs == true

        v2, isabs2 = XLSX.parse_defined_name_value("Sheet1!\$A\$1:\$B\$2")
        @test v2 == XLSX.SheetCellRange("Sheet1!A1:B2")
        @test isabs2 == true
    end

    @testset "parse_defined_name_value - empty string" begin
        v, isabs = XLSX.parse_defined_name_value("")
        @test ismissing(v)
        @test isabs == false
    end

    @testset "parse_defined_name_value - quoted, numeric and fallback strings" begin
        @test XLSX.parse_defined_name_value("\"hello\"") == ("hello", false)
        @test ismissing(first(XLSX.parse_defined_name_value("\"\"")))
        @test XLSX.parse_defined_name_value("42") == (42, false)
        @test XLSX.parse_defined_name_value("3.5") == (3.5, false)
        @test XLSX.parse_defined_name_value("SomethingElse") == ("SomethingElse", false)
    end


    # ===========================================================================
    # process_file failure path
    # ===========================================================================

    @testset "process_file - unreadable zip entry throws XLSXError" begin
        io = IOBuffer()
        XLSX.ZipArchives.ZipWriter(io) do w
            XLSX.ZipArchives.zip_newfile(w, "present.xml")
            write(w, """<thing/>""")
        end
        reader = XLSX.ZipArchives.ZipReader(take!(io))

        @test XLSX.process_file(reader, "present.xml").name == "present.xml"
        @test_throws XLSX.XLSXError XLSX.process_file(reader, "absent.xml")
    end


    # ===========================================================================
    # Source-not-found and bad-argument paths on the public read entry points
    # ===========================================================================

    @testset "readtable - file not found, every arity" begin
        missing_file = "definitely_not_a_file_12345.xlsx"
        @test !isfile(missing_file)

        @test_throws XLSX.XLSXError XLSX.readtable(missing_file)
        @test_throws XLSX.XLSXError XLSX.readtable(missing_file, "Sheet1")
        @test_throws XLSX.XLSXError XLSX.readtable(missing_file, "Sheet1", XLSX.ColumnRange("A:B"))
    end

    @testset "readtable - columns argument is not a valid column range" begin
        outfile = "read_badrange.xlsx"
        isfile(outfile) && rm(outfile)
        XLSX.writetable(outfile, [[1, 2], [3, 4]], ["a", "b"])

        @test_throws XLSX.XLSXError XLSX.readtable(outfile, 1, "not a range")
        @test_throws XLSX.XLSXError XLSX.readtable(outfile, 1, "A1:B2")  # cell range, not columns

        isfile(outfile) && rm(outfile)
    end

    @testset "readtransposedtable - file not found, every arity" begin
        missing_file = "definitely_not_a_file_12345.xlsx"

        @test_throws XLSX.XLSXError XLSX.readtransposedtable(missing_file)
        @test_throws XLSX.XLSXError XLSX.readtransposedtable(missing_file, "Sheet1")
        @test_throws XLSX.XLSXError XLSX.readtransposedtable(missing_file, "Sheet1", "1:3")
    end

    @testset "openxlsx / parse_file_mode - argument errors" begin
        missing_file = "definitely_not_a_file_12345.xlsx"

        @test_throws XLSX.XLSXError XLSX.openxlsx(missing_file)
        @test_throws XLSX.XLSXError XLSX.openxlsx(identity, missing_file)
        @test_throws XLSX.XLSXError XLSX.openxlsx(missing_file; mode="q")
        @test XLSX.parse_file_mode("wr") == (true, true)
        @test XLSX.parse_file_mode("RW") == (true, true)
    end
end