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
        SAVE_FILES && save_outfile(xf)
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
        SAVE_FILES && save_outfile(xf)
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
        SAVE_FILES && save_outfile(xf)
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
        SAVE_FILES && save_outfile(xf)
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
        SAVE_FILES && save_outfile(xf)
    end

    @testset "ensure_workbook_is_xlsx! - unknown workbook content type errors" begin
        xf = XLSX.newxlsx()
        root = XLSX.xml_root_element(xf.data["[Content_Types].xml"])
        _, ovr = _find_child(root, "Override", "PartName", "/xl/workbook.xml")
        @test !isnothing(ovr)
        ovr["ContentType"] = "application/vnd.example.not-a-workbook+xml"

        @test_throws XLSX.XLSXError XLSX.ensure_workbook_is_xlsx!(xf)
        SAVE_FILES && save_outfile(xf)
    end

    @testset "ensure_workbook_is_xlsx! - no content type at all errors" begin
        xf = XLSX.newxlsx()
        root = XLSX.xml_root_element(xf.data["[Content_Types].xml"])

        i, _ = _find_child(root, "Override", "PartName", "/xl/workbook.xml")
        isnothing(i) || deleteat!(root.children, i)
        j, _ = _find_child(root, "Default", "Extension", "xml")
        isnothing(j) || deleteat!(root.children, j)

        @test_throws XLSX.XLSXError XLSX.ensure_workbook_is_xlsx!(xf)
        SAVE_FILES && save_outfile(xf)
    end

    @testset "check_minimum_requirements - missing mandatory part errors" begin
        xf = XLSX.newxlsx()
        delete!(xf.files, "xl/_rels/workbook.xml.rels")
        @test_throws XLSX.XLSXError XLSX.check_minimum_requirements(xf)
        SAVE_FILES && save_outfile(xf)
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
        SAVE_FILES && save_outfile(xf)
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
        XLSX.writetable(outfile, [[1, 2], [3, 4]], ["a", "b"]; overwrite=true)
        SAVE_FILES && save_outfile(outfile)

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

# Read-only, cache-on: the `<sheetData>` stub made at open is kept and swapped in by
# the first `eachrow`, instead of splitting the sheet XML a second time (#462).
@testset "sheet stub reuse" begin

    @testset "stub kept at open and swapped in after the cache fill" begin
        nsheets = 0
        for file in sort(readdir(data_directory))
            any(ext -> endswith(lowercase(file), ext), (".xlsx", ".xlsm", ".xltx", ".xltm")) || continue
            xf = try XLSX.readxlsx(joinpath(data_directory, file)) catch; continue end
            wb = XLSX.get_workbook(xf)
            for ws in wb.sheets
                XLSX.is_chartsheet(wb, ws.name) && continue
                target = XLSX.get_relationship_target_by_id("xl", wb, ws.relationship_id)
                raw = xf.data[target]
                @test raw isa String
                expected = XLSX.splitNode(raw, "sheetData")[1]
                @test get(xf.sheet_stubs, target, nothing) == expected
                XLSX.eachrow(ws)
                @test xf.data[target] == expected
                @test !haskey(xf.sheet_stubs, target)
                XLSX.eachrow(ws)                       # already filled: nothing changes
                @test xf.data[target] == expected
                nsheets += 1
            end
            @test isempty(xf.sheet_stubs)
        end
        @test nsheets > 100
    end

    @testset "stub of a Strict OOXML sheet is converted with the sheet" begin
        for file in ("strict.xlsx", "Strict-foo.xlsx", "chart_strict.xlsx")
            xf = XLSX.readxlsx(joinpath(data_directory, file))
            wb = XLSX.get_workbook(xf)
            ws = first(s for s in wb.sheets if !XLSX.is_chartsheet(wb, s.name))
            target = XLSX.get_relationship_target_by_id("xl", wb, ws.relationship_id)
            stub = xf.sheet_stubs[target]
            @test !occursin("purl.oclc.org/ooxml", stub)
            XLSX.eachrow(ws)
            doc = XLSX.get_xml_data(xf, target)
            @test !occursin("purl.oclc.org/ooxml", XML.write(doc))
            @test XLSX.readtable(joinpath(data_directory, file), ws.name).column_labels ==
                  XLSX.gettable(ws).column_labels
        end
    end

    @testset "no stubs where the cache isn't used lazily" begin
        file = joinpath(data_directory, "general.xlsx")
        XLSX.openxlsx(file; enable_cache = false) do xf
            @test isempty(xf.sheet_stubs)
            @test !isempty(collect(XLSX.eachrow(xf[1])))
        end
        # "rw" writes back on close, so use a copy
        copy_path = "general_copy.xlsx"
        cp(file, copy_path; force=true)
        XLSX.openxlsx(copy_path; mode = "rw") do xf
            @test isempty(xf.sheet_stubs)
        end
        SAVE_FILES && save_outfile(copy_path)
        isfile(copy_path) && rm(copy_path)
    end

    @testset "reading the stub-swapped sheet afterwards" begin
        # Code that parses the worksheet XML after the cache fill must see the same
        # document whether the stub came from open or from the fallback split.
        for file in ("testmerge.xlsx", "Book1.xlsx", "strict.xlsx")
            path = joinpath(data_directory, file)
            xf = XLSX.readxlsx(path)
            wb = XLSX.get_workbook(xf)
            target = XLSX.get_relationship_target_by_id("xl", wb, xf[1].relationship_id)
            XLSX.eachrow(xf[1])
            doc = XLSX.get_xml_data(xf, target)
            xf2 = XLSX.readxlsx(path)
            empty!(xf2.sheet_stubs)                  # force the fallback split
            XLSX.eachrow(xf2[1])
            doc2 = XLSX.get_xml_data(xf2, target)
            @test XML.write(doc) == XML.write(doc2)
        end
    end
end


# The cache fill reads each `<c>` in the same cursor walk as its children (#462). It
# must build exactly the cells `Cell(::LazyNode, …)` builds from the same XML.

# Every cell of a sheet, built one at a time with `Cell(::LazyNode, …)`, and the
# formulas they record.
function _cells_by_lazynode(xf, ws)
    wb = XLSX.get_workbook(xf)
    target = XLSX.get_relationship_target_by_id("xl", wb, ws.relationship_id)
    raw = xf.data[target]
    sst_pfx = XLSX.get_sst_prefix(ws)
    formulas = Dict{XLSX.SheetCellRef, XLSX.AbstractFormula}()
    cells = Dict{Tuple{Int,Int}, XLSX.Cell}()
    c = XML.Cursor(raw)
    in_data = false
    while XML.next!(c) !== nothing
        XML.nodetype(c) == XML.Element || continue
        name = XLSX.localname(c)
        d = XML.depth(c)
        if d == 2
            in_data = name == "sheetData"
            in_data || XML.skip_element!(c)
        elseif in_data && d == 4 && name == "c"
            cell = XLSX.Cell(XML.LazyNode(c), ws, sst_pfx, formulas, xf.load_formulas)
            XML.skip_element!(c)
            cells[(XLSX.row_number(cell), XLSX.column_number(cell))] = cell
        elseif in_data && d == 4
            XML.skip_element!(c)
        end
    end
    return cells, formulas
end

_same_fields(a, b) = typeof(a) == typeof(b) &&
    all(isequal(getfield(a, f), getfield(b, f)) for f in fieldnames(typeof(a)))

@testset "cache fill reads cells as Cell(::LazyNode) does" begin
    dirs = [data_directory]
    fixtures = get(ENV, "XLSX_DIFF_FIXTURES", "")
    isempty(fixtures) || push!(dirs, fixtures)
    nsheets = 0
    for dir in dirs, file in sort(readdir(dir))
        any(ext -> endswith(lowercase(file), ext), (".xlsx", ".xlsm", ".xltx", ".xltm")) || continue
        path = joinpath(dir, file)
        xf = try XLSX.readxlsx(path) catch; continue end
        ref_xf = XLSX.readxlsx(path)                  # a second copy for the reference
        wb, ref_wb = XLSX.get_workbook(xf), XLSX.get_workbook(ref_xf)
        for (ws, ref_ws) in zip(wb.sheets, ref_wb.sheets)
            XLSX.is_chartsheet(wb, ws.name) && continue
            expected, expected_formulas = _cells_by_lazynode(ref_xf, ref_ws)
            XLSX.eachrow(ws)                         # fill the cache
            got = Dict((r, col) => cell for (r, row) in ws.cache.cells for (col, cell) in row)
            same = keys(got) == keys(expected) &&
                   all(_same_fields(got[k], expected[k]) for k in keys(expected))
            same || println("cache fill differs from Cell(::LazyNode): $file, sheet $(ws.name)")
            @test same
            got_formulas = filter(p -> first(p).sheet == ws.name, wb.formulas)
            @test keys(got_formulas) == keys(expected_formulas)
            @test all(_same_fields(got_formulas[k], expected_formulas[k]) for k in keys(expected_formulas))
            @test ws.sst_count == count(c -> c.datatype == XLSX.CT_STRING, values(expected))
            nsheets += 1
        end
    end
    @test nsheets > 100
end

@testset "cache fill: cell forms" begin
    # Each cell form on its own, through the cache fill and through Cell(::LazyNode).
    sst = "<si><t>a</t></si><si><t>b</t></si>"
    forms = [
        "number"            => "<c r=\"A1\"><v>1.5</v></c>",
        "styled number"     => "<c r=\"A1\" s=\"1\"><v>45000</v></c>",
        "shared string"     => "<c r=\"A1\" t=\"s\"><v>1</v></c>",
        "inline string"     => "<c r=\"A1\" t=\"inlineStr\"><is><t>hi</t></is></c>",
        "inline rich text"  => "<c r=\"A1\" t=\"inlineStr\"><is><r><rPr><b/></rPr><t>b</t></r><r><t>x</t></r></is></c>",
        "inline, empty"     => "<c r=\"A1\" t=\"inlineStr\"><is><t/></is></c>",
        "inline, two is"    => "<c r=\"A1\" t=\"inlineStr\"><is><t>1</t></is><is><t>2</t></is></c>",
        "inline with v, f"  => "<c r=\"A1\" t=\"inlineStr\"><f>X()</f><v>9</v><is><t>z</t></is></c>",
        "formula + value"   => "<c r=\"A1\"><f>1+1</f><v>2</v></c>",
        "formula, no value" => "<c r=\"A1\"><f>1+1</f></c>",
        "shared formula"    => "<c r=\"A1\"><f t=\"shared\" ref=\"A1:A3\" si=\"0\">B1*2</f><v>4</v></c>",
        "shared reference"  => "<c r=\"A1\"><f t=\"shared\" si=\"0\"/><v>4</v></c>",
        "array formula"     => "<c r=\"A1\"><f t=\"array\" ref=\"A1:A2\">SEQUENCE(2)</f><v>1</v></c>",
        "cm metadata"       => "<c r=\"A1\" cm=\"1\"><f t=\"array\" ref=\"A1\">X()</f><v>1</v></c>",
        "empty v"           => "<c r=\"A1\"><v/></c>",
        "v, empty text"     => "<c r=\"A1\"><v></v></c>",
        "v holding CDATA"   => "<c r=\"A1\" t=\"str\"><v><![CDATA[a<b]]></v></c>",
        "v with entity"     => "<c r=\"A1\" t=\"str\"><v>a&amp;b</v></c>",
        "styled empty"      => "<c r=\"A1\" s=\"2\"/>",
        "error"             => "<c r=\"A1\" t=\"e\"><v>#N/A</v></c>",
        "boolean"           => "<c r=\"A1\" t=\"b\"><v>1</v></c>",
        "formula string"    => "<c r=\"A1\" t=\"str\"><f>\"x\"</f><v>x</v></c>",
        "unknown child"     => "<c r=\"A1\"><extLst><ext/></extLst><v>3</v></c>",
        "comment inside"    => "<c r=\"A1\"><!-- note --><v>3</v></c>",
        "whitespace inside" => "<c r=\"A1\">\n  <v>3</v>\n</c>",
        "two v"             => "<c r=\"A1\"><v>1</v><v>2</v></c>",
    ]
    for (name, cxml) in forms
        rows = "<row r=\"1\">$cxml<c r=\"B1\"><v>7</v></c></row><row r=\"2\"><c r=\"A2\"><v>8</v></c></row>"
        bytes = _diff_build_xlsx(rows, ["a", "b"])
        xf, ref_xf = XLSX.readxlsx(IOBuffer(bytes)), XLSX.readxlsx(IOBuffer(bytes))
        expected, expected_formulas = _cells_by_lazynode(ref_xf, ref_xf[1])
        ws = xf[1]
        XLSX.eachrow(ws)
        got = Dict((r, col) => cell for (r, row) in ws.cache.cells for (col, cell) in row)
        same = keys(got) == keys(expected) && all(_same_fields(got[k], expected[k]) for k in keys(expected))
        same || println("cache fill differs from Cell(::LazyNode): $name")
        @test same
        wbf = XLSX.get_workbook(xf).formulas
        @test keys(wbf) == keys(expected_formulas)
        @test all(_same_fields(wbf[k], expected_formulas[k]) for k in keys(expected_formulas))
        @test XLSX.getdata(ws, "B1") == 7 && XLSX.getdata(ws, "A2") == 8
    end
end


# `_parse_cell_float` must give exactly what `parse(Float64, s)` gives, bit for bit,
# and throw exactly where it throws (#462: cell floats are parsed with Parsers.jl).
@testset "cell value float parsing" begin
    same_as_base(s) = begin
        b = try parse(Float64, s) catch e; typeof(e) end
        f = try XLSX._parse_cell_float(s) catch e; typeof(e) end
        b isa Float64 && f isa Float64 ? reinterpret(UInt64, b) == reinterpret(UInt64, f) : b == f
    end

    @testset "edge cases" begin
        for s in ["0", "-0", "0.0", "-0.0", "00012.500", "1", "-1", ".5", "-.5", "1.", "5E0",
                  "1E+22", "1E22", "1E23", "1E-22", "1E-23", "1e5", "2.5E-3", "2.5e-03", "1E+05",
                  "123456789012345", "1234567890123456", "9007199254740993", "999999999999999",
                  "0.000000000000000000001", "1.50000000000000000000", "0.30000000000000004",
                  "1.7976931348623157E+308", "2.2250738585072014E-308", "4.9E-324", "1E+400", "1E-400",
                  "45000.5", "-123.456", "12345678901234.5", "1234567890123.45",
                  # 16-25 significant digits, subnormals, overflow and halfway cases
                  "12345.678901234567", "0.1000000000000000055511151231257827", "1234567890123456789012345",
                  "2.2250738585072011E-308", "2.4703282292062327E-324", "2.4703282292062328E-324",
                  "1.7976931348623158E+308", "1.7976931348623159E+308", "1E+309", "9007199254740992.5",
                  "9007199254740994.9999999999999999", "8.98846567431158E+307", "1E+2147483648",
                  # not plain decimals: both must agree (Base parses some, rejects others)
                  "", "-", ".", "-.", "1e", "1E+", "1e-", "E5", "abc", "1.2.3", " 1.5", "1.5 ",
                  "+1.5", "1_0", "0x10", "0x1p3", "Inf", "-Inf", "NaN", "nan", "Infinity", "1,5", "1d5",
                  "--1", "1e5.5", "	2
", "１"]
            ok = same_as_base(s)
            ok || println("float parse differs from Base for ", repr(s))
            @test ok
        end
        @test XLSX._parse_cell_float("-0") === -0.0
        @test XLSX._parse_cell_float(SubString("<v>12345.678901234567</v>", 4, 21)) === 12345.678901234567
    end

    @testset "random values in the forms Excel writes" begin
        rng = Random.MersenneTwister(462)
        nbad = 0
        for _ in 1:300_000
            x = (2rand(rng) - 1) * 10.0^rand(rng, -30:30)
            k = rand(rng, 1:17)
            for s in (string(x), string(round(x; sigdigits = k)),
                      uppercase(string(round(x; sigdigits = k))),
                      string(round(x; digits = rand(rng, 0:6))),
                      string(rand(rng, -10^9:10^9)))
                same_as_base(s) || (nbad += 1; nbad <= 5 && println("float parse differs from Base for ", repr(s)))
            end
        end
        @test nbad == 0
        # extremes of the exponent range
        @test all(same_as_base(string((2rand(rng) - 1) * 10.0^rand(rng, -307:308))) for _ in 1:20_000)
    end

    @testset "every <v> in test/data" begin
        nvals = 0
        nbad = 0
        for file in sort(readdir(data_directory))
            any(ext -> endswith(lowercase(file), ext), (".xlsx", ".xlsm", ".xltx", ".xltm")) || continue
            zip = try ZipArchives.ZipReader(read(joinpath(data_directory, file))) catch; continue end
            for name in ZipArchives.zip_names(zip)
                occursin(r"xl/worksheets/sheet\d*\.xml", name) || continue
                xml = String(ZipArchives.zip_readentry(zip, name))
                for m in eachmatch(r"<(?:\w+:)?v>([^<]*)</(?:\w+:)?v>", xml)
                    s = m.captures[1]
                    tryparse(Float64, s) === nothing && continue
                    nvals += 1
                    same_as_base(s) && continue
                    nbad += 1
                    nbad <= 5 && println("float parse differs from Base for ", repr(s), " in ", file)
                end
            end
        end
        @test nbad == 0
        @test nvals > 1000
    end
end

# General-format cells and SST indices use `Parsers.tryparse(Int64, v)`, which must give
# exactly what `tryparse(Int64, v)` gives.
@testset "cell value integer parsing" begin
    for s in ["0", "-0", "+5", "00012", "42", "-42", "9223372036854775807", "9223372036854775808",
              "-9223372036854775808", "-9223372036854775809", "99999999999999999999", "1.0", "1e3",
              "1E3", "12345.678901234567", "", "-", "+", " 12 ", "	7
", "_1", "1_0", "0x1F", "0b101",
              "0o17", "1,0", "١٢", "abc"]
        ok = isequal(XLSX.Parsers.tryparse(Int64, s), tryparse(Int64, s))
        ok || println("integer parse differs from Base for ", repr(s))
        @test ok
        w = "<v>" * s * "</v>"
        @test isequal(XLSX.Parsers.tryparse(Int64, SubString(w, 4, prevind(w, ncodeunits(w) - 3))), tryparse(Int64, s))
    end
    rng = Random.MersenneTwister(462)
    nbad = 0
    for _ in 1:200_000
        s = string(rand(rng, (rand(rng, -10^6:10^6), rand(rng, Int64), rand(rng, Int128), (2rand(rng) - 1) * 10.0^rand(rng, -5:20))))
        isequal(XLSX.Parsers.tryparse(Int64, s), tryparse(Int64, s)) || (nbad += 1)
    end
    @test nbad == 0
end
