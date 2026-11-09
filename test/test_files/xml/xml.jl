# issue #117
@testset "whitespace nodes" begin
    xf = XLSX.readxlsx(joinpath(data_directory, "noutput_first_second_third.xlsx"))
    @test XLSX.sheetnames(xf) == ["NOTES", "DATA"]
    @test xf["NOTES"]["A1"] == "Nominal GNP/GDP"
    @test xf["NOTES"]["A9"] == "Last updated on: August 29, 2019"
    @test xf["DATA"]["A5"] == "Date"
    @test xf["DATA"]["A6"] == "1965:Q3"
    @test xf["DATA"]["B6"] ≈ 6.7731
    @test xf["DATA"]["E5"] == "Most_Recent"
    @test xf["DATA"]["E7"] ≈ 12.6215
end

# issue #303
@testset "xml:space" begin
    f = XLSX.openxlsx(joinpath(data_directory, "sstTest.xlsx"), mode="rw")
    s = f[1]
    @test XLSX.getdata(s, :) == ["  hello" "    "; "  hello  " "    "; " hello\">" "    "; "hello\">" "    "; "  hello" "    "]
    s["C1"] = " "
    s["C2"] = " hello"
    s["C3"] = "hello "
    s["C4"] = " hello "
    s["C5"] = " \"hello\" "
    @test XLSX.getdata(s, "C1:C5") == Any[" "; " hello"; "hello "; " hello "; " \"hello\" ";;]
    XLSX.writexlsx("mydata.xlsx", f, overwrite=true)
    SAVE_FILES && save_outfile("mydata.xlsx")
    @test XLSX.readdata("mydata.xlsx", 1, :) == ["  hello" "    " " "; "  hello  " "    " " hello"; " hello\">" "    " "hello "; "hello\">" "    " " hello "; "  hello" "    " " \"hello\" "]
    XLSX.writetable("mydata.xlsx", [["  hello", "  hello  ", " hello\">", "hello\">", "  hello"], ["    ", "    ", "    ", "    ", "    "], [" ", " hello", "hello ", " hello ", " \"hello\" "]], ["Col_A", "Col_B", "Col_C"]; overwrite=true)
    SAVE_FILES && save_outfile("mydata.xlsx")
    @test XLSX.readdata("mydata.xlsx", 1, :) == ["Col_A" "Col_B" "Col_C"; "  hello" "    " " "; "  hello  " "    " " hello"; " hello\">" "    " "hello "; "hello\">" "    " " hello "; "  hello" "    " " \"hello\" "]
    isfile("mydata.xlsx") && rm("mydata.xlsx")
end

# issue #243
@testset "xml bom" begin
    xf = XLSX.readxlsx(joinpath(data_directory, "Bom - issue243.xlsx"))
    @test XLSX.sheetnames(xf) == ["QMJ Factors", "Definition", "Data Sources", "--> Additional Global Factors", "Disclosures"]
    @test XLSX.sheetcount(xf) == 5
    @test XLSX.hassheet(xf, "QMJ Factors") == true
    @test xf["QMJ Factors"]["H833"] ≈ -0.0686846616503713
end

# XML.jl reads a String in place whatever its line ends (XML.jl#194). XLSX.jl relies on it
# not to copy, at each read, every part saved by Excel, whose XML declaration ends with
# CR LF (#482).
@testset "no copy for line ends" begin
    decl = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
    body = "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">" *
           "<dimension ref=\"A1:A100000\"/><sheetData>" *
           repeat("<row r=\"1\"><c r=\"A1\"><v>1</v></c></row>\n", 100_000) * "</sheetData></worksheet>"
    lazy(x) = XML.parse(x, XML.LazyNode)
    nsprefix(x) = XLSX._get_ns_prefix_from_string(x)
    for doc in ("$decl\r\n$body", replace("$decl\n$body", "\n" => "\r\n"))
        lazy(doc); nsprefix(doc)
        @test (@allocated lazy(doc)) < sizeof(doc) ÷ 100
        @test (@allocated nsprefix(doc)) < sizeof(doc) ÷ 100
        # the stub is a new String, but it is the only copy
        XLSX.splitNode(doc, "sheetData")
        @test (@allocated XLSX.splitNode(doc, "sheetData")) < sizeof(doc) ÷ 100
    end

    # A streamed read drops a leading BOM and line feed, as the other reads do (#243)
    path = joinpath(data_directory, "Bom - issue243.xlsx")
    rows(cache) = XLSX.openxlsx(path; enable_cache = cache) do xf
        s = xf["QMJ Factors"]
        (count(_ -> true, XLSX.eachrow(s)), s["H833"])
    end
    @test rows(false) == rows(true)
end

@testset "escape" begin

    @test XML.escape("hello&world<'") == "hello&amp;world&lt;&apos;"
    @test XML.unescape("hello&amp;world&lt;&apos;") == "hello&world<'"
    # XLSX.jl relies on these to write CR, LF and tab so that they read back unchanged (#482)
    @test XML.escape("a\r\nb") == "a&#13;\nb"
    @test XML.escape("x\ty\n"; attribute = true) == "x&#9;y&#10;"

    esc_filename = "output_table_escape_test.xlsx"

    esc_col_names = ["&' & \" < > '", "I❤Julia", "\"<'&O-O&'>\"", "<&>"]
    esc_sheetname = "& & \" > < "
    esc_data = Vector{Any}(undef, 4)
    esc_data[1] = ["11&&", "12\"&", "13<&", "14>&", "15'&"]
    esc_data[2] = ["21&&&&", "22&\"&&", "23&<&&", "24&>&&", "25&'&&"]
    esc_data[3] = ["31&&&&&&", "32&&\"&&&", "33&&<&&&", "34&&>&&&", "35&&'&&&"]
    esc_data[4] = ["41& &; &&", "42\" \"; \"\"", "43< <; <<", "44> >; >>", "45' '; ''"]
    XLSX.writetable(esc_filename, esc_data, esc_col_names, overwrite=true, sheetname=esc_sheetname)
    SAVE_FILES && save_outfile(esc_filename)

    dtable = XLSX.readtable(esc_filename, esc_sheetname)
    r1_data, r1_col_names = dtable.data, dtable.column_labels
    check_test_data(r1_data, esc_data)
    @test r1_col_names[4] == Symbol(esc_col_names[4])
    @test r1_col_names[3] == Symbol(esc_col_names[3])
    @test r1_col_names[2] == Symbol(esc_col_names[2])
    @test r1_col_names[1] == Symbol(esc_col_names[1])

    # compare to the backup version: escape.xlsx
    dtable = XLSX.readtable(joinpath(data_directory, "escape.xlsx"), esc_sheetname)
    r2_data, r2_col_names = dtable.data, dtable.column_labels
    check_test_data(r2_data, esc_data)
    check_test_data(r2_data, r1_data)
    @test string(r2_col_names[4]) == esc_col_names[4]
    @test string(r2_col_names[3]) == esc_col_names[3]
    @test string(r2_col_names[2]) == esc_col_names[2]
    @test string(r2_col_names[1]) == esc_col_names[1]

    esc_col_names = ["&; &amp; &quot; &lt; &gt; &apos; ", "I❤Julia", "\"<'&O-O&'>\"", "<&>"]
    esc_sheetname = string( esc_col_names[1],esc_col_names[2],esc_col_names[3],esc_col_names[4])[1:30] # There is a hard limit in Excel
    esc_data = Vector{Any}(undef, 4)
    esc_data[1] = ["11&amp;&",    "12&quot;&",    "13&lt;&",    "14&gt;&",    "15&apos;&"    ]
    esc_data[2] = ["21&&amp;&&",  "22&&quot;&&",  "23&&lt;&&",  "24&&gt;&&",  "25&&apos;&&"  ]
    esc_data[3] = ["31&&&amp;&&&","32&&&quot;&&&","33&&&lt;&&&","34&&&gt;&&&","35&&&apos;&&&"]
    esc_data[4] = ["41& &; &&",   "42\" \"; \"\"","43< <; <<",  "44> >; >>",  "45' '; ''"    ]
    XLSX.writetable(esc_filename, esc_data, esc_col_names, overwrite=true, sheetname=esc_sheetname)
    SAVE_FILES && save_outfile(esc_filename)

    dtable = XLSX.readtable(esc_filename, esc_sheetname)
    r3_data, r3_col_names = dtable.data, dtable.column_labels
    check_test_data(r3_data, esc_data)
    @test r3_col_names[4] == Symbol( esc_col_names[4] )
    @test r3_col_names[3] == Symbol( esc_col_names[3] )
    @test r3_col_names[2] == Symbol( esc_col_names[2] )
    @test r3_col_names[1] == Symbol( esc_col_names[1] )
    isfile(esc_filename) && rm(esc_filename)

    # compare to the backup version: escape2.xlsx
    dtable = XLSX.readtable(joinpath(data_directory, "escape2.xlsx"), esc_sheetname)
    r4_data, r4_col_names = dtable.data, dtable.column_labels
    check_test_data(r4_data, esc_data)
    check_test_data(r4_data, r3_data)
    @test r4_col_names[4] == Symbol( esc_col_names[4] )
    @test r4_col_names[3] == Symbol( esc_col_names[3] )
    @test r4_col_names[2] == Symbol( esc_col_names[2] )
    @test r4_col_names[1] == Symbol( esc_col_names[1] )


end

# The `splicetext`-based `splitNode` that preceded the raw-scanner version (#462), kept
# here as the reference the real one must match exactly. Both splice the bytes they are
# given, CRs included (#482).
function _splitnode_reference(xml_str::String, skipnode::String)
    c = XML.Cursor(xml_str)
    XML.next!(c)
    while !XML.eof(c) && XML.nodetype(c) != XML.Element
        XML.next!(c)
    end
    XML.eof(c) && return xml_str, ""
    target_lazy = nothing
    while XML.next!(c) !== nothing
        XML.depth(c) == 0 && break
        XML.depth(c) != 2 && (XML.skip_element!(c); continue)
        XML.nodetype(c) == XML.Element || continue
        if XLSX.localname(c) == skipnode
            target_lazy = XML.LazyNode(c)
            XML.skip_element!(c)
            break
        end
        XML.skip_element!(c)
    end
    isnothing(target_lazy) && return xml_str, ""
    target_tag = XML.tag(target_lazy)
    attrs = XML.attributes(target_lazy)
    replacement = if isnothing(attrs) || isempty(attrs)
        "<$(target_tag)/>"
    else
        "<$(target_tag) $(join(("$(k)=\"$(XML.escape(v; attribute = true))\"" for (k,v) in attrs), " "))/>"
    end
    return XML.splicetext(target_lazy, replacement), ""
end

# Both must return the same result, or throw the same type of exception.
function _splitnode_agree(xml_str::String)
    expected = try _splitnode_reference(xml_str, "sheetData") catch e; e end
    actual   = try XLSX.splitNode(xml_str, "sheetData") catch e; e end
    expected isa Exception && return actual isa Exception && typeof(actual) == typeof(expected)
    actual isa Exception && return false
    return actual[1] isa String && actual == expected
end

@testset "splitNode" begin

    @testset "matches the reference on every worksheet" begin
        dirs = [data_directory]
        fixtures = get(ENV, "XLSX_DIFF_FIXTURES", "")
        isempty(fixtures) || push!(dirs, fixtures)
        nsheets = 0
        for dir in dirs, file in sort(readdir(dir))
            any(ext -> endswith(lowercase(file), ext), (".xlsx", ".xlsm", ".xltx", ".xltm")) || continue
            zip = try ZipArchives.ZipReader(read(joinpath(dir, file))) catch; continue end
            for name in ZipArchives.zip_names(zip)
                occursin(r"xl/worksheets/sheet\d*\.xml", name) || continue
                nsheets += 1
                ok = _splitnode_agree(XLSX.strip_bom_and_lf(XLSX._readentry_string(zip, name)))
                ok || println("splitNode differs from the reference: $file $name")
                @test ok
            end
        end
        @test nsheets > 100
    end

    ws(body; ns = "") = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n" *
        "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"$ns>$body</worksheet>"
    rows = "<row r=\"1\"><c r=\"A1\"><v>1</v></c></row><row r=\"2\"><c r=\"A2\" t=\"inlineStr\"><is><t>x</t></is></c></row>"

    @testset "edge cases" begin
        cases = [
            "plain"               => ws("<dimension ref=\"A1:A2\"/><sheetData>$rows</sheetData><pageMargins left=\"0.7\"/>"),
            "attributes"          => ws("<sheetData foo=\"1\" bar=\"a&amp;b\">$rows</sheetData><pageMargins left=\"0.7\"/>"),
            "self-closing"        => ws("<dimension ref=\"A1\"/><sheetData/><pageMargins left=\"0.7\"/>"),
            "self-closing, attrs" => ws("<sheetData foo=\"1\"/><pageMargins left=\"0.7\"/>"),
            "empty element"       => ws("<sheetData></sheetData><pageMargins left=\"0.7\"/>"),
            "prefixed"            => "<x:worksheet xmlns:x=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><x:sheetData><x:row r=\"1\"><x:c r=\"A1\"><x:v>1</x:v></x:c></x:row></x:sheetData><x:pageMargins left=\"0.7\"/></x:worksheet>",
            "comment with close"  => ws("<sheetData><!-- </sheetData> -->$rows</sheetData><pageMargins left=\"0.7\"/>"),
            "CDATA with close"    => ws("<sheetData><row r=\"1\"><c r=\"A1\" t=\"str\"><v><![CDATA[</sheetData>]]></v></c></row></sheetData><pageMargins left=\"0.7\"/>"),
            "PI inside"           => ws("<sheetData><?pi </sheetData> ?>$rows</sheetData><pageMargins left=\"0.7\"/>"),
            "multi-byte around"   => ws("<sheetPr codeName=\"日本語é\"/><sheetData><row r=\"1\"><c r=\"A1\" t=\"inlineStr\"><is><t>😀ü</t></is></c></row></sheetData><headerFooter><oddHeader>Ωmega ✓</oddHeader></headerFooter>"),
            "multi-byte adjacent" => ws("<sheetPr codeName=\"é\"/><sheetData>$rows</sheetData><!--ü-->"),
            "no sheetData"        => ws("<dimension ref=\"A1\"/><pageMargins left=\"0.7\"/>"),
            "nested, not depth 2" => ws("<extLst><sheetData>$rows</sheetData></extLst>"),
            "second sheetData"    => ws("<sheetData>$rows</sheetData><sheetData><row r=\"9\"/></sheetData>"),
            "root not closed"     => "<worksheet><sheetData>$rows</sheetData>",
            "sheetData not closed" => "<worksheet><sheetData>$rows",
            "only the root"       => "<worksheet/>",
            "no elements"         => "<?xml version=\"1.0\"?>",
        ]
        # Each case again with CR LF, then a lone CR, between every pair of tags
        for (name, xml) in copy(cases)
            push!(cases, "$name, CR LF" => replace(xml, "><" => ">\r\n<"),
                         "$name, CR"    => replace(xml, "><" => ">\r<"))
        end
        for (name, xml) in cases
            ok = _splitnode_agree(xml)
            ok || println("splitNode differs from the reference: $name")
            @test ok
        end
        # The result itself, for the cases where it's easy to state
        @test XLSX.splitNode(ws("<sheetData foo=\"1\">$rows</sheetData><pageMargins left=\"0.7\"/>"), "sheetData")[1] ==
              ws("<sheetData foo=\"1\"/><pageMargins left=\"0.7\"/>")
        @test XLSX.splitNode(ws("<sheetPr codeName=\"é\"/><sheetData>$rows</sheetData><!--ü-->"), "sheetData")[1] ==
              ws("<sheetPr codeName=\"é\"/><sheetData/><!--ü-->")
        @test XLSX.splitNode(ws("<dimension ref=\"A1\"/>"), "sheetData")[1] == ws("<dimension ref=\"A1\"/>")
        # An attribute value is written back escaped, so the stub stays well-formed
        stub = XLSX.splitNode(ws("<sheetData bar=\"a&amp;b&#10;c\">$rows</sheetData>"), "sheetData")[1]
        @test stub == ws("<sheetData bar=\"a&amp;b&#10;c\"/>")
        sd = XLSX.xml_root_element(XML.parse(stub, XML.Node))[1]
        @test XML.attributes(sd)["bar"] == "a&b\nc"
    end

    # The stub keeps the bytes it was given, CRs included (#482)
    @testset "CR line endings" begin
        cases = [
            "plain"      => ("<sheetFormatPr defaultRowHeight=\"15\"/>\n<sheetData>\n$rows\n</sheetData>\n<pageMargins left=\"0.7\"/>",
                             "<sheetFormatPr defaultRowHeight=\"15\"/>\n<sheetData/>\n<pageMargins left=\"0.7\"/>"),
            "attributes" => ("<sheetFormatPr defaultRowHeight=\"15\"/>\n<sheetData foo=\"1\">\n$rows</sheetData>\n<pageMargins left=\"0.7\"/>",
                             "<sheetFormatPr defaultRowHeight=\"15\"/>\n<sheetData foo=\"1\"/>\n<pageMargins left=\"0.7\"/>"),
        ]
        for eol in ("\r\n", "\r"), (name, (body, expected)) in cases
            xml = replace(ws(body), "\n" => eol)
            @test _splitnode_agree(xml)
            out = XLSX.splitNode(xml, "sheetData")[1]
            out == replace(ws(expected), "\n" => eol) || println("splitNode CR case failed: $name, $(repr(eol))")
            @test out == replace(ws(expected), "\n" => eol)
            @test XML.parse(out, XML.Node) isa XML.Node
        end
        # The #482 shape: one CR LF, after the declaration, in a large sheet
        big = replace(ws("<sheetData>" * repeat(rows, 20_000) * "</sheetData><pageMargins left=\"0.7\"/>"), "?>\n" => "?>\r\n")
        @test XLSX.splitNode(big, "sheetData")[1] == replace(ws("<sheetData/><pageMargins left=\"0.7\"/>"), "?>\n" => "?>\r\n")
    end

    @testset "uses no XML.jl internals" begin
        # XML.jl v0.5 removes `_skip_element_raw` and may change any internal field;
        # `sourceview`, briefly on its `main`, was removed before any release (XML.jl#194)
        for (root, _, files) in walkdir(dirname(pathof(XLSX))), f in files
            endswith(f, ".jl") || continue
            text = read(joinpath(root, f), String)
            for pat in (r"XMLTokenizer", r"\.token\.offset", r"\.st\.state", r"sourceview")
                hit = occursin(pat, text)
                hit && println("XML.jl internal $(pat.pattern) used in $f")
                @test !hit
            end
        end
    end
end

# A newline or comment after the root element's end tag must not be mistaken for the root.
@testset "trailing nodes after root element" begin
    src = "trailing_base.xlsx"
    xf = XLSX.newxlsx()
    xf[1]["A1:B3"] = 1
    XLSX.writexlsx(src, xf; overwrite=true)
    for tail in ("\n", "<!-- x -->")
        f = patch_xlsx_entry(src, "xl/worksheets/sheet1.xml", s -> s * tail)
        xf = XLSX.openxlsx(f; mode="rw")
        s = xf[1]
        XLSX.setConditionalFormat(s, "A1:B3", :cellIs)
        @test length(XLSX.getConditionalFormats(s)) == 1
        XLSX.writexlsx("trailing_out.xlsx", xf; overwrite=true)
        SAVE_FILES && save_outfile("trailing_out.xlsx")
        xf = XLSX.openxlsx("trailing_out.xlsx"; mode="rw")
        @test length(XLSX.getConditionalFormats(xf[1])) == 1
        XLSX.clearConditionalFormats(xf[1], "A1:B3")
        @test isempty(XLSX.getConditionalFormats(xf[1]))
        xf = XLSX.openxlsx(f; mode="rw")   # cleared on a sheet whose XML still has the tail
        XLSX.setConditionalFormat(xf[1], "A1:B3", :cellIs)
        XLSX.clearConditionalFormats(xf[1], "A1:B3")
        @test isempty(XLSX.getConditionalFormats(xf[1]))
        rm(f)

        f = patch_xlsx_entry(src, "xl/workbook.xml", s -> s * tail)
        xf = XLSX.openxlsx(f; mode="rw")
        XLSX.addDefinedName(xf, "trailingName", "Sheet1!A1:B2")
        XLSX.writexlsx("trailing_out.xlsx", xf; overwrite=true)
        SAVE_FILES && save_outfile("trailing_out.xlsx")
        @test haskey(XLSX.get_workbook(XLSX.readxlsx("trailing_out.xlsx")).workbook_names, "trailingName")
        rm(f)
    end
    rm(src)
    rm("trailing_out.xlsx")
end

@testset "insert_index" begin
    # Nothing in the order list precedes the target: insert at the top, not after the first child.
    root = XML.Element("worksheet", XML.Element("pageMargins"))
    @test XLSX.insert_index(root, "sheetPr", XLSX.WORKSHEET_ORDER) == 0
    @test XLSX.insert_index(XML.Element("worksheet"), "sheetPr", XLSX.WORKSHEET_ORDER) == 0
end

@testset "remove_attr! - node with no attributes, and absent key" begin
    bare = XLSX.XML.Element("x")
    @test isnothing(XLSX.remove_attr!(bare, "anything"))

    node = XLSX.XML.Element("x"; a="1", b="2")
    @test isnothing(XLSX.remove_attr!(node, "notThere"))
    @test XLSX.get_attr(node, "a") == "1"
    @test XLSX.get_attr(node, "b") == "2"

    XLSX.remove_attr!(node, "a")
    @test XLSX.get_attr(node, "a", "") == ""
    @test !haskey(XML.attributes(node), "a")
    @test XLSX.get_attr(node, "b") == "2"
end
