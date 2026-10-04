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

@testset "escape" begin

    @test XML.escape("hello&world<'") == "hello&amp;world&lt;&apos;"
    @test XML.unescape("hello&amp;world&lt;&apos;") == "hello&world<'"

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

# The `splicetext`-based `splitNode` that preceded the raw-scanner version (#462),
# kept here as the reference the faster one must match exactly.
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
        "<$(target_tag) $(join(("$(k)=\"$(v)\"" for (k,v) in attrs), " "))/>"
    end
    return XML.splicetext(target_lazy, replacement), ""
end

# Both must return the same result, or throw the same type of exception.
function _splitnode_agree(xml_str::String)
    expected = try _splitnode_reference(xml_str, "sheetData") catch e; e end
    actual   = try XLSX.splitNode(xml_str, "sheetData") catch e; e end
    expected isa Exception && return actual isa Exception && typeof(actual) == typeof(expected)
    return actual == expected
end

@testset "splitNode" begin

    @testset "matches the splicetext reference on every worksheet" begin
        dirs = [data_directory]
        fixtures = get(ENV, "XLSX_DIFF_FIXTURES", "")
        isempty(fixtures) || push!(dirs, fixtures)
        nsheets = 0
        for dir in dirs, file in sort(readdir(dir))
            any(ext -> endswith(lowercase(file), ext), (".xlsx", ".xlsm", ".xltx", ".xltm")) || continue
            zip = try ZipArchives.ZipReader(read(joinpath(dir, file))) catch; continue end
            for name in ZipArchives.zip_names(zip)
                occursin(r"xl/worksheets/sheet\d*\.xml", name) || continue
                bytes = ZipArchives.zip_readentry(zip, name)
                XLSX.strip_bom_and_lf!(bytes)
                nsheets += 1
                ok = _splitnode_agree(String(bytes))
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
    end

    @testset "_element_span" begin
        xml = ws("<sheetPr codeName=\"é\"/><sheetData>$rows</sheetData><pageMargins left=\"0.7\"/>")
        c = XML.Cursor(xml)
        while XML.next!(c) !== nothing
            XML.nodetype(c) == XML.Element && XLSX.localname(c) == "sheetData" && break
        end
        n = XML.LazyNode(c)
        start, stop = XLSX._element_span(n)
        @test SubString(xml, start, prevind(xml, stop)) == XML.sourcetext(n)
        @test xml[start:start+10] == "<sheetData>"
        @test xml[prevind(xml, stop, 12):prevind(xml, stop)] == "</sheetData>"
    end
end
