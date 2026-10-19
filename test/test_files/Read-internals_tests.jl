# Reading internals: workbook parsing, the worksheet cache fill and its sheet stubs,
# cell value parsing, and the sheet dimension.
#
# Most of these drive internals directly rather than through a fixture file:
# the branches they reach are mostly malformed-package paths that would otherwise
# need a purpose-built .xlsx each.

const _CT_SHEET    = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"
const _CT_TEMPLATE = "application/vnd.openxmlformats-officedocument.spreadsheetml.template.main+xml"

# Locate a child of `root` by localname and an attribute value.
function _find_child(root, tag::AbstractString, attr::AbstractString, val::AbstractString)
    for (i, c) in enumerate(XLSX.XML.children(root))
        XLSX.localname(c) == tag || continue
        XLSX.get_attr(c, attr, "") == val && return (i, c)
    end
    return (nothing, nothing)
end

@testset "workbook parsing and file arrays" begin

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
        SAVE_FILES && save_outfile(xf)
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

    @testset "FileArray" begin
        p = joinpath(data_directory, "general.xlsx")
        fa = XLSX.FileArray(p)
        @test fa[1] == read(p)[1]
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

# General-format cells and SST indices use `_parse_cell_int(v)` (Parsers.jl), which must give
# exactly what `tryparse(Int64, v)` gives.
@testset "cell value integer parsing" begin
    for s in ["0", "-0", "+5", "00012", "42", "-42", "9223372036854775807", "9223372036854775808",
              "-9223372036854775808", "-9223372036854775809", "99999999999999999999", "1.0", "1e3",
              "1E3", "12345.678901234567", "", "-", "+", " 12 ", "	7
", "_1", "1_0", "0x1F", "0b101",
              "0o17", "1,0", "١٢", "abc"]
        ok = isequal(XLSX._parse_cell_int(s), tryparse(Int64, s))
        ok || println("integer parse differs from Base for ", repr(s))
        @test ok
        w = "<v>" * s * "</v>"
        @test isequal(XLSX._parse_cell_int(SubString(w, 4, prevind(w, ncodeunits(w) - 3))), tryparse(Int64, s))
    end
    rng = Random.MersenneTwister(462)
    nbad = 0
    for _ in 1:200_000
        s = string(rand(rng, (rand(rng, -10^6:10^6), rand(rng, Int64), rand(rng, Int128), (2rand(rng) - 1) * 10.0^rand(rng, -5:20))))
        isequal(XLSX._parse_cell_int(s), tryparse(Int64, s)) || (nbad += 1)
    end
    @test nbad == 0
end

@testset "sheet dimensions and ranges across cache modes" begin
    @testset "Read-only range discovery" begin
        bytes = read(joinpath(data_directory, "NoDim.xlsx"))
        expected = XLSX.openxlsx(IOBuffer(bytes); mode = "rw")[1]["A1:K116"]
        factories = (XLSX.readxlsx,
                     io -> XLSX.openxlsx(io; enable_cache = true),
                     io -> XLSX.openxlsx(io; enable_cache = false))
        cellvalues(ws, cells) = map(c -> XLSX.getdata(ws, c), cells)
        ranges = (
            (ws -> ws[:], expected),
            (ws -> ws[:, :], expected),
            (ws -> XLSX.getdata(ws), expected),
            (ws -> ws[1:2, :], expected[1:2, :]),
            (ws -> ws[:, 1:2], expected[:, 1:2]),
            (ws -> ws[[1, 2], :], expected[1:2, :]),
            (ws -> ws[:, [1, 2]], expected[:, 1:2]),
            (ws -> ws["1:2"], expected[1:2, :]),
            (ws -> ws["A:B"], expected[:, 1:2]),
            (ws -> cellvalues(ws, XLSX.getcellrange(ws, :)), expected),
            (ws -> cellvalues(ws, XLSX.getcell(ws, :)), expected),
            (ws -> cellvalues(ws, XLSX.getcell(ws, 1:2, :)), expected[1:2, :]),
            (ws -> cellvalues(ws, XLSX.getcell(ws, :, 1:2)), expected[:, 1:2]),
            (ws -> cellvalues(ws, XLSX.getcellrange(ws, "1:2")), expected[1:2, :]),
            (ws -> cellvalues(ws, XLSX.getcellrange(ws, "A:B")), expected[:, 1:2]),
        )
        for openfile in factories
            for (readrange, reference) in ranges
                ws = openfile(IOBuffer(bytes))[1]
                @test isequal(readrange(ws), reference)
                @test ws.dimension == XLSX.CellRange("A1:K116") # recorded by the first read
                @test isequal(readrange(ws), reference)
            end
            ws = openfile(IOBuffer(bytes))[1]
            @test axes(ws, 1) == 1:116
            @test axes(ws, 2) == 1:11
            @test XLSX.get_dimension(ws) == XLSX.CellRange("A1:K116")
            !XLSX.is_cache_enabled(ws) && @test ws.cache === nothing

            ws = openfile(IOBuffer(bytes))[1]
            sprint(show, ws)
            @test ws.cache === nothing
            @test ws.dimension === nothing

            ws = openfile(IOBuffer(read(joinpath(data_directory, "simple.xlsx"))))[1]
            @test XLSX.get_dimension(ws) == XLSX.CellRange("A1:B3")
            @test ws.cache === nothing
        end

        # An uncached row iterator records the dimension only when it runs to the end.
        ws = XLSX.openxlsx(IOBuffer(bytes); enable_cache = false)[1]
        for _ in XLSX.eachrow(ws) end
        @test ws.dimension == XLSX.CellRange("A1:K116")
        ws = XLSX.openxlsx(IOBuffer(bytes); enable_cache = false)[1]
        for _ in XLSX.eachrow(ws)
            break
        end
        @test ws.dimension === nothing

        # A recorded dimension is never recomputed.
        for enable_cache in (true, false)
            ws = XLSX.openxlsx(joinpath(data_directory, "simple.xlsx"); enable_cache)[1]
            dimension = ws.dimension
            ws[:]; ws[1:2, :]; ws[:, 2]; axes(ws, 2)
            for _ in XLSX.eachrow(ws) end
            @test ws.dimension === dimension
        end

        sparse = Matrix{Any}(missing, 5, 3)
        sparse[1, 1] = 11
        sparse[5, 3] = 22
        styled = Matrix{Any}(missing, 3, 3)
        styled[1, 1] = 11
        inline = Matrix{Any}(missing, 2, 2)
        inline[1, 1] = "hi"
        inline[2, 2] = 5
        emptyrow = Matrix{Any}(missing, 4, 1)
        emptyrow[1, 1] = 1
        sheets = (
            ("<sheetData/>", "A1:A1", fill(missing, 1, 1)),
            ("<sheetData><row r=\"7\" ht=\"18\" customHeight=\"1\"/></sheetData>", "A1:A1", fill(missing, 1, 1)),
            ("<sheetData><row r=\"3\"><c r=\"C3\"><v>11</v></c></row><row r=\"7\"><c r=\"E7\"><v>22</v></c></row></sheetData>", "C3:E7", sparse),
            # a formatting-only cell extends the dimension
            ("<sheetData><row r=\"2\"><c r=\"B2\"><v>11</v></c></row><row r=\"4\"><c r=\"D4\" s=\"0\"/></row></sheetData>", "B2:D4", styled),
            ("<sheetData><row r=\"1\"><c r=\"A1\" t=\"inlineStr\"><is><t>hi</t></is></c></row><row r=\"2\"><c r=\"B2\"><v>5</v></c></row></sheetData>", "A1:B2", inline),
            # a row without cells still extends the rows
            ("<sheetData><row r=\"2\"><c r=\"B2\"><v>1</v></c></row><row r=\"5\" ht=\"18\" customHeight=\"1\"/></sheetData>", "B2:B5", emptyrow),
        )
        for (sheetdata, dimension, reference) in sheets
            path = patch_xlsx_entry(joinpath(src_data_directory, "blank.xlsx"), "xl/worksheets/sheet1.xml",
                _ -> "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">" * sheetdata * "</worksheet>")
            try
                rng = XLSX.CellRange(dimension)
                top, left = rng.start.row_number, rng.start.column_number
                for openfile in factories
                    ws = openfile(IOBuffer(read(path)))[1]
                    @test isequal(ws[:], reference)
                    @test isequal(ws[:], reference)
                    @test XLSX.get_dimension(ws) == rng
                    !XLSX.is_cache_enabled(ws) && @test ws.cache === nothing

                    # partial reads take the other extent from the sheet
                    ws = openfile(IOBuffer(read(path)))[1]
                    @test isequal(ws[top:top, :], reference[1:1, :])
                    ws = openfile(IOBuffer(read(path)))[1]
                    @test isequal(ws[:, left], reference[:, 1:1])
                    ws = openfile(IOBuffer(read(path)))[1]
                    @test axes(ws, 1) == rng.start.row_number:rng.stop.row_number
                    @test axes(ws, 2) == rng.start.column_number:rng.stop.column_number
                end
            finally
                rm(path)
            end
        end
    end

    noDim = XLSX.openxlsx(joinpath(data_directory, "NoDim.xlsx"), mode="rw")
    Dim = XLSX.readxlsx(joinpath(data_directory, "customXml.xlsx"))
    @test noDim[1].dimension == Dim[1].dimension
    @test noDim[2].dimension == Dim[2].dimension

    f = XLSX.newxlsx()
    s = f[1]
    for i = 10:20, j = 10:20
        s[i, j] = i + j
    end
    XLSX.set_dimension!(s, XLSX.CellRange(XLSX.CellRef("J10"), XLSX.CellRef("T20")))
    @test XLSX.get_dimension(s) == XLSX.CellRange(XLSX.CellRef("J10"), XLSX.CellRef("T20"))
    s["A1"] = 2
    @test XLSX.get_dimension(s) == XLSX.CellRange(XLSX.CellRef("A1"), XLSX.CellRef("T20"))
    SAVE_FILES && save_outfile(f)

    @testset "Worksheet ranges across cache modes" begin
        for filename in ("simple.xlsx", "strict.xlsx", "NoDim.xlsx")
            bytes = read(joinpath(data_directory, filename))
            writable = XLSX.openxlsx(IOBuffer(bytes); mode = "rw")
            expected = writable[1][:]
            range = "A1:" * XLSX.encode_column_number(size(expected, 2)) * string(size(expected, 1))
            for openfile in (XLSX.readxlsx,
                             io -> XLSX.openxlsx(io; enable_cache = true),
                             io -> XLSX.openxlsx(io; enable_cache = false))
                xf = openfile(IOBuffer(bytes))
                @test XLSX.sheetnames(xf) == XLSX.sheetnames(writable)
                @test isequal(xf[1][range], expected)
                @test isequal(xf[1][range], expected)
            end
        end
    end
end

@testset "review fixes: reading" begin
    @testset "a worksheet part not named sheetN.xml" begin
        XLSX.writetable("base_part.xlsx", [[1, 2, 3]], ["a"]; overwrite=true)
        r = ZipReader(read("base_part.xlsx"))
        open("renamed_part.xlsx", "w") do io
            ZipWriter(io) do w
                for n in zip_names(r)
                    bytes = zip_readentry(r, n)
                    if n in ("xl/_rels/workbook.xml.rels", "[Content_Types].xml")
                        bytes = Vector{UInt8}(replace(String(bytes), "worksheets/sheet1.xml" => "worksheets/Data.xml"))
                    end
                    zip_newfile(w, n == "xl/worksheets/sheet1.xml" ? "xl/worksheets/Data.xml" : n)
                    write(w, bytes)
                end
            end
        end
        @test XLSX.readxlsx("renamed_part.xlsx")[1]["A3"] == 2
        @test XLSX.readtable("renamed_part.xlsx", 1).data == Any[[1, 2, 3]]
        XLSX.openxlsx("renamed_part.xlsx"; mode="rw") do xf
            xf[1]["B1"] = "new"
        end
        @test XLSX.readxlsx("renamed_part.xlsx")[1]["B1"] == "new"
        SAVE_FILES && save_outfile("renamed_part.xlsx")
    end
    @testset "negative number in a date-formatted cell" begin
        f = XLSX.newxlsx(); s = f[1]
        st = Int(XLSX.get_num_style_index(s, 14).id)
        @test XLSX.process_tv(f.workbook, "", "-1", st) == (XLSX.CT_INT, reinterpret(UInt64, Int64(-1)))
        @test XLSX.process_tv(f.workbook, "", "-1.5", st)[1] == XLSX.CT_FLOAT
    end
    @testset "a datetime that rounds up to midnight" begin
        @test XLSX.excel_value_to_datetime(45000.999999999, false) == DateTime(2023, 3, 16)
    end
    @testset "too-short file and missing sheet index give XLSXError" begin
        touch("empty.xlsx")
        @test_throws XLSX.XLSXError XLSX.readxlsx("empty.xlsx")
        XLSX.writetable("one_sheet.xlsx", [[1, 2]], ["a"]; overwrite=true)
        @test_throws XLSX.XLSXError XLSX.readtable("one_sheet.xlsx", 5)
    end
end
