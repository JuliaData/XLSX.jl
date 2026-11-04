# `xl/sharedStrings.xml` must be well-formed and must decode back to the original text.
function check_sst(path::AbstractString)
    sst = String(zip_readentry(ZipReader(read(path)), "xl/sharedStrings.xml"))
    XML.parse(sst, XML.Node)   # throws on the corrupt output produced before the fix
    return sst
end

# Replace A1's inlineStr content outright. Errors rather than silently no-opping if
# the fixture changes shape — a patch that quietly does nothing produces tests that
# pass while covering the wrong thing.
function set_a1_inline(s::AbstractString, inner::AbstractString)
    m = match(r"<c\b[^>]*\br=\"A1\"[^>]*>.*?</c>"s, s)
    m === nothing && error("no A1 cell found in sheet XML")
    occursin("inlineStr", m.match) || error("A1 is not an inlineStr cell: $(m.match)")
    stop = m.offset + ncodeunits(m.match) - 1
    return s[1:prevind(s, m.offset)] *
           "<c r=\"A1\" t=\"inlineStr\"><is>$inner</is></c>" *
           s[nextind(s, stop):end]
end
@testset "inlineStr" begin
    xf = XLSX.readxlsx(joinpath(data_directory, "inlinestr.xlsx"))
    sheet = xf["Requirements"]
    @test sheet["A1"] == "NN"
    @test sheet["A2"] == 1
    @test sheet["B1"] == "Hierarchy"
    @test sheet["B2"] == "+"
    @test ismissing(sheet["C1"])
    @test ismissing(sheet["C2"])
    @test sheet["D1"] == "Outline Number"
    @test sheet["D2"] == "1."
    @test sheet["E1"] == "ID"
    @test sheet["E2"] == "RQ11610"
    @test sheet["F1"] == "Name"
    @test sheet["F2"] == "requirement"
    @test sheet["G1"] == "Type"
    @test sheet["G2"] == "Textual Requirement"
    @test sheet["H1"] == "Description"
    @test sheet["H2"] == "test"
    @test ismissing(sheet["I1"])
    @test ismissing(sheet["I2"])
    @test ismissing(sheet["J1"])
    @test ismissing(sheet["J2"])

    xf = XLSX.openxlsx(joinpath(data_directory, "inlinestr.xlsx"), mode="rw")
    XLSX.writexlsx("mydata.xlsx", xf, overwrite=true)
    SAVE_FILES && save_outfile("mydata.xlsx")
    xf = XLSX.readxlsx("mydata.xlsx")
    sheet = xf["Requirements"]
    @test sheet["A1"] == "NN"
    @test sheet["A2"] == 1
    @test sheet["B1"] == "Hierarchy"
    @test sheet["B2"] == "+"
    @test ismissing(sheet["C1"])
    @test ismissing(sheet["C2"])
    @test sheet["D1"] == "Outline Number"
    @test sheet["D2"] == "1."
    @test sheet["E1"] == "ID"
    @test sheet["E2"] == "RQ11610"
    @test sheet["F1"] == "Name"
    @test sheet["F2"] == "requirement"
    @test sheet["G1"] == "Type"
    @test sheet["G2"] == "Textual Requirement"
    @test sheet["H1"] == "Description"
    @test sheet["H2"] == "test"
    @test ismissing(sheet["I1"])
    @test ismissing(sheet["I2"])
    @test ismissing(sheet["J1"])
    @test ismissing(sheet["J2"])
    isfile("mydata.xlsx") && rm("mydata.xlsx")

    @testset "markup in inline strings is escaped on write (issue #447)" begin
        src = joinpath(data_directory, "one_and_two_start.xlsx")
        sheetxml = "xl/worksheets/sheet1.xml"

        # (text written into <is>, expected value after round-trip)
        cases = [
            ("<t>one &amp; two</t>",        "one & two"),
            ("<t>a &lt;b&gt; c</t>",        "a <b> c"),
            ("<t><![CDATA[x & y]]></t>",    "x & y"),
            ("<t>quote \" and &apos;</t>",  "quote \" and '"),
            # Multiple runs: exercises the recursive branch of `_rewrite_node` and the
            # attribute path (`rFont val`), which the single-<t> cases never reach.
            ("""<r><rPr><b/><rFont val="A &amp; B"/></rPr><t>one &amp; two</t></r>""" *
             """<r><t xml:space="preserve"> &lt;then&gt; </t></r>""" *
             """<r><rPr><i/></rPr><t>three</t></r>""",
             "one & two <then> three"),
            # A CR written as a reference must be written back as one, or a reader
            # normalises it to LF (#482); likewise a LF in an attribute value.
            ("<t>a&#13;&#10;b</t>",          "a\r\nb"),
            ("""<r><rPr><rFont val="A&#10;B"/></rPr><t>x&#13;y</t></r>""", "x\ry"),
        ]

        for (inner, expected) in cases
            f = patch_xlsx_entry(src, sheetxml, s -> set_a1_inline(s, inner))
            xf = XLSX.openxlsx(f, mode="rw")
            @test xf[1]["A1"] == expected          # read decodes correctly (passes pre-fix)

            out = "strings_inline_escape_out.xlsx"
            XLSX.writexlsx(out, xf, overwrite=true)

            sst = check_sst(out)                   # throws pre-fix
            @test !occursin(r"&(?!amp;|lt;|gt;|apos;|quot;|#)", sst)
            @test XLSX.readxlsx(out)[1]["A1"] == expected

            SAVE_FILES && save_outfile(xf)
            isfile(f) && rm(f)
            isfile(out) && rm(out)
        end
    end
end

@testset "rich text formats" begin
    @testset "setFont, setUniformFont and setUniformStyle on rich text" begin
        xf = XLSX.readxlsx(joinpath(data_directory, "is.xlsx"))
        sh = xf["Sheet1"]
        @test_throws XLSX.XLSXError XLSX.setFont(sh, "A1"; name="Palatino") # Must be in write mode for rich text formats

        xf = XLSX.opentemplate(joinpath(data_directory, "is.xlsx"))
        sh = xf["Sheet1"]
        @test XLSX.getFont(sh, "A1") === nothing
        XLSX.setFont(sh, "A1"; name="Palatino")
        r = XLSX.getRichTextString(sh, "A1")
        @test r.runs[4].text == "ki"
        @test r.runs[4].atts == Dict(:bold => true, :color => "FF000000", :size => 11)
        @test r.runs[6].text == "ty"
        @test r.runs[6].atts == Dict(:color => "FFFF0000", :size => 11)
        # An empty cell holds no rich text, so nothing rather than an error.
        @test isnothing(XLSX.getRichTextString(sh, "Z99"))
        @test isnothing(XLSX.getRichTextString(sh, 99, 26))
        @test isnothing(XLSX.getRichTextString(xf, "Sheet1!Z99"))     # outside the dimension too
        @test XLSX.getFont(sh, "A1").font ==Dict("name" => Dict("val" => "Palatino"), "sz" => Dict("val" => "11"), "color" => Dict("theme" => "1"))
        XLSX.setFont(sh, "A1:F2"; name="Palatino")
        r = XLSX.getRichTextString(sh, "B2")
        @test r.runs[4].text == "k"
        @test r.runs[4].atts == Dict(:bold => true, :color => "FF000000", :size => 11, :under => true)
        r = XLSX.getRichTextString(sh, "F2")
        @test r.runs[3].text == "lo "
        @test r.runs[3].atts == Dict(:color => "FF000000", :size => 11)
        @test XLSX.getFont(sh, "B2").font["name"] == Dict("val" => "Palatino")
        XLSX.setFont(sh, "B"; under="none")
        @test haskey(XLSX.getFont(sh, "B2").font, "u") == false
        XLSX.setFont(sh, "C1,D2:E2", color="orange")
        r = XLSX.getRichTextString(sh, "C1")
        @test r.runs[6].text == "ty "
        @test r.runs[6].atts == Dict(:size => 11)
        r = XLSX.getRichTextString(sh, "E2")
        @test r.runs[9].text == "t"
        @test r.runs[9].atts == Dict(:size => 24)
        @test XLSX.getFont(sh, "C1").font["color"] == Dict("rgb" => XLSX.get_color("orange"))
        @test XLSX.getFont(sh, "D2").font["color"] == Dict("rgb" => XLSX.get_color("orange"))
        @test XLSX.getFont(sh, "D1").font["color"] == Dict("theme" => "1")
        SAVE_FILES && save_outfile(xf)

        xf = XLSX.opentemplate(joinpath(data_directory, "is.xlsx"))
        sh = xf["Sheet1"]
        XLSX.setUniformFont(sh, [1, 2], 1:2; under="double")
        @test XLSX.getFont(sh, "A1").font["u"] == Dict("val" => "double")
        @test XLSX.getFont(sh, "B2").font["u"] == Dict("val" => "double")
        @test XLSX.getFont(sh, "A1").font["sz"] == Dict("val" => "11")
        @test XLSX.getFont(sh, "B2").font["sz"] == Dict("val" => "11")
        @test XLSX.getFont(sh, "A1").font["color"] == Dict("theme" => "1")
        @test XLSX.getFont(sh, "B2").font["color"] == Dict("theme" => "1")
        @test XLSX.getFont(sh, "C2") === nothing
        r = XLSX.getRichTextString(sh, "C1")
        @test r.runs[3].text == "lo "
        @test r.runs[3].atts == Dict(:color => "FF000000", :name => "Aptos Narrow", :size => 11)
        @test r.runs[4].text == "ki"
        @test r.runs[4].atts == Dict(:bold => true, :color => "FF000000", :italic => true, :name => "Aptos Narrow", :size => 11)

        XLSX.setUniformFont(sh, :, 3:2:5; color="orange")
        @test XLSX.getFont(sh, "C1").font["sz"] == Dict("val" => "11")
        @test XLSX.getFont(sh, "E2").font["sz"] == Dict("val" => "11")
        @test XLSX.getFont(sh, "C1").font["color"] == Dict("rgb" => XLSX.get_color("orange"))
        @test XLSX.getFont(sh, "E2").font["color"] == Dict("rgb" => XLSX.get_color("orange"))
        @test XLSX.getFont(sh, "D2") === nothing
        r = XLSX.getRichTextString(sh, "E2")
        @test r.runs[9].text == "t"
        @test r.runs[9].atts == Dict(:name => "Aptos Narrow", :size => 24)
        @test r.runs[10].text == "y "
        @test r.runs[10].atts == Dict(:name => "Aptos Narrow", :size => 11)
        SAVE_FILES && save_outfile(xf)

        xf = XLSX.opentemplate(joinpath(data_directory, "is.xlsx"))
        sh = xf["Sheet1"]
        XLSX.setUniformFont(sh, "A1,C2,D1:F1"; under="double")
        @test XLSX.getFont(sh, "A1").font["u"] == Dict("val" => "double")
        @test XLSX.getFont(sh, "C2").font["u"] == Dict("val" => "double")
        @test XLSX.getFont(sh, "D1").font["sz"] == Dict("val" => "11")
        @test XLSX.getFont(sh, "E1").font["sz"] == Dict("val" => "11")
        @test XLSX.getFont(sh, "F1").font["color"] == Dict("theme" => "1")
        @test XLSX.getFont(sh, "C2").font["color"] == Dict("theme" => "1")
        @test XLSX.getFont(sh, "B2") === nothing
        r = XLSX.getRichTextString(sh, "F1")
        @test r.runs[3].text == "lo "
        @test r.runs[3].atts == Dict(:color => "FF000000", :name => "Calibri", :size => 11)
        @test r.runs[8].text == " "
        @test r.runs[8].atts == Dict(:color => "FF000000", :name => "Aptos Narrow", :size => 11)
        SAVE_FILES && save_outfile(xf)

        xf = XLSX.opentemplate(joinpath(data_directory, "is.xlsx"))
        sh = xf["Sheet1"]
        XLSX.setFont(sh, "A1"; under="double", color="orange", name="Palatino", size=20, bold=true, italic=true, strike=true)
        XLSX.setUniformStyle(sh, "A1:F2")
        @test XLSX.getFont(sh, "A1").font["u"] == Dict("val" => "double")
        @test XLSX.getFont(sh, "C2").font["u"] == Dict("val" => "double")
        @test XLSX.getFont(sh, "D1").font["sz"] == Dict("val" => "11")
        @test XLSX.getFont(sh, "E1").font["sz"] == Dict("val" => "20")
        @test XLSX.getFont(sh, "F1").font["color"] == Dict("rgb" => XLSX.get_color("orange"))
        @test XLSX.getFont(sh, "C2").font["color"] == Dict("rgb" => XLSX.get_color("orange"))
        @test XLSX.getFont(sh, "A1").font["name"] == Dict("val" => "Palatino")
        @test XLSX.getFont(sh, "F2").font["name"] == Dict("val" => "Palatino")
        @test haskey(XLSX.getFont(sh, "B2").font, "b") == true
        @test haskey(XLSX.getFont(sh, "B2").font, "i") == true
        @test haskey(XLSX.getFont(sh, "B2").font, "strike") == true
        SAVE_FILES && save_outfile(xf)
    end

    # A run's own properties override the cell font, so a font change removes the ones
    # it sets from every run. `vertAlign` was missed, and a uniform style threw on it.
    @testset "every font property reaches the runs of rich text" begin
        function rich_sheet()
            xf = XLSX.newxlsx()
            sh = xf[1]
            sh["A1"] = XLSX.RichTextString([XLSX.RichTextRun("ab", [:bold => true]),
                                            XLSX.RichTextRun("cd", [:italic => true, :vertAlign => "subscript"])])
            sh["A2"] = XLSX.RichTextString([XLSX.RichTextRun("ef", [:bold => true]),
                                            XLSX.RichTextRun("gh", [:vertAlign => "subscript"])])
            return xf, sh
        end
        run_atts(sh, ref) = [r.atts for r in XLSX.getRichTextString(sh, ref).runs]

        xf, sh = rich_sheet()
        XLSX.setFont(sh, "A1"; vertAlign="superscript")
        @test run_atts(sh, "A1") == [Dict(:bold => true), Dict(:italic => true)]
        @test XLSX.getFont(sh, "A1").font["vertAlign"] == Dict("val" => "superscript")

        xf, sh = rich_sheet()
        XLSX.setUniformFont(sh, "A1:A2"; vertAlign="superscript")
        @test run_atts(sh, "A1") == [Dict(:bold => true), Dict(:italic => true)]
        @test run_atts(sh, "A2")[2] === nothing || isempty(run_atts(sh, "A2")[2])

        xf, sh = rich_sheet()
        XLSX.setFont(sh, "A1"; vertAlign="superscript")
        XLSX.setUniformStyle(sh, "A1:A2")
        @test XLSX.getFont(sh, "A2").font["vertAlign"] == Dict("val" => "superscript")
        @test all(a -> isnothing(a) || !haskey(a, :vertAlign), run_atts(sh, "A2"))

        # A font size Excel wrote as a fraction is a tag like any other.
        xf, sh = rich_sheet()
        cell = XLSX.getcell(sh, "A1")
        font = XLSX.CellFont(0, Dict{String,Union{Dict{String,String},Nothing}}("sz" => Dict("val" => "10.5")), "1")
        @test XLSX.update_sharedString_font(sh, cell, font) isa Integer

        # Run properties the new font doesn't set are kept, including ones XLSX.jl
        # doesn't otherwise read, such as `charset` and `outline`.
        xf, sh = rich_sheet()
        wb = XLSX.get_workbook(sh)
        si = "<si><r><rPr><b/><charset val=\"0\"/></rPr><t>ab</t></r><r><rPr><i/><outline/></rPr><t>cd</t></r></si>"
        cell = XLSX.getcell(sh, "A1")
        cell.value = reinterpret(UInt64, Int64(XLSX.add_formatted_string!(wb, si)))
        XLSX.setFont(sh, "A1"; bold=true)
        s = XLSX.get_sst(wb).shared_strings[reinterpret(Int64, XLSX.getcell(sh, "A1").value) + 1]
        @test occursin("charset", s) && occursin("outline", s) && occursin("<i", s)
        @test !occursin("<b/>", s)
        SAVE_FILES && save_outfile(xf)
    end

    @static if VERSION >= v"1.11-"
        @testset "styledStrings" begin
            f=XLSX.newxlsx()
            s=f[1]
            s["A1"] = styled"{yellow:hello} {blue:there}"
            s["A2"] = styled"The {bold:{italic:quick {(foreground=#cd853f):brown} fox} jumps over the {(foreground=#FFC000):lazy} dog}"
            s["A3"] = styled"In terms of color, we have named faces for the 16 standard terminal colors:\n {black:■} {red:■} {green:■} {yellow:■} {blue:■} {magenta:■} {cyan:■} {white:■}\n {bright_black:■} {bright_red:■} {bright_green:■} {bright_yellow:■} {bright_blue:■} {bright_magenta:■} {bright_cyan:■} {bright_white:■}"
            s["A4"] = styled"\nThe basic font-style attributes are {bold:bold}, {light:light}, {italic:italic},\n {underline:underline}, and {strikethrough:strikethrough}"
            s["A5"] = styled"{red:{(height=2.0):{bold:Hello}} {(height=120):Tim!}}"
            StyledStrings.addface!(:MyFace => StyledStrings.Face(foreground = 0xFF7700))
            s["A6"] = styled"{MyFace:this is orange text}"
            s["A7"] = styled"{(weight=bold):{(underline=true, slant=italic):deleted {(strikethrough=true):and more {(height=2.5):too}}, aswell}}. And now unbolded!"
            s["A8"] = styled"Hello {(font=Consolas):computer - do {bright_green:you have {(height=2.4):any} {(font=Palatino, foreground=#cd853f):rhymes}} for} me"
            @test s["A1"] == "hello there"
            @test s["A2"] == "The quick brown fox jumps over the lazy dog"
            @test s["A3"] == "In terms of color, we have named faces for the 16 standard terminal colors:\n ■ ■ ■ ■ ■ ■ ■ ■\n ■ ■ ■ ■ ■ ■ ■ ■"
            @test s["A4"] == "\nThe basic font-style attributes are bold, light, italic,\n underline, and strikethrough" # in Excel, light is simply plain text (not bold).
            @test s["A5"] == "Hello Tim!"
            @test s["A6"] == "this is orange text"
            @test s["A7"] == "deleted and more too, aswell. And now unbolded!"
            @test s["A8"] == "Hello computer - do you have any rhymes for me"
            @test XLSX.get_workbook(s).sst.shared_strings[Int(XLSX.getcell(s, "A1").value)+1] == "<si><r><rPr><color rgb=\"FFE5A509\"/></rPr><t>hello</t></r><r><t xml:space=\"preserve\"> </t></r><r><rPr><color rgb=\"FF195EB3\"/></rPr><t>there</t></r></si>"
            @test XLSX.get_workbook(s).sst.shared_strings[Int(XLSX.getcell(s, "A2").value)+1] == "<si><r><t xml:space=\"preserve\">The </t></r><r><rPr><b/><i/></rPr><t xml:space=\"preserve\">quick </t></r><r><rPr><b/><i/><color rgb=\"FFCD853F\"/></rPr><t>brown</t></r><r><rPr><b/><i/></rPr><t xml:space=\"preserve\"> fox</t></r><r><rPr><b/></rPr><t xml:space=\"preserve\"> jumps over the </t></r><r><rPr><b/><color rgb=\"FFFFC000\"/></rPr><t>lazy</t></r><r><rPr><b/></rPr><t xml:space=\"preserve\"> dog</t></r></si>"
            @test XLSX.get_workbook(s).sst.shared_strings[Int(XLSX.getcell(s, "A3").value)+1] == "<si><r><t xml:space=\"preserve\">In terms of color, we have named faces for the 16 standard terminal colors:\n </t></r><r><rPr><color rgb=\"FF1C1A23\"/></rPr><t>■</t></r><r><t xml:space=\"preserve\"> </t></r><r><rPr><color rgb=\"FFA51C2C\"/></rPr><t>■</t></r><r><t xml:space=\"preserve\"> </t></r><r><rPr><color rgb=\"FF25A268\"/></rPr><t>■</t></r><r><t xml:space=\"preserve\"> </t></r><r><rPr><color rgb=\"FFE5A509\"/></rPr><t>■</t></r><r><t xml:space=\"preserve\"> </t></r><r><rPr><color rgb=\"FF195EB3\"/></rPr><t>■</t></r><r><t xml:space=\"preserve\"> </t></r><r><rPr><color rgb=\"FF803D9B\"/></rPr><t>■</t></r><r><t xml:space=\"preserve\"> </t></r><r><rPr><color rgb=\"FF0097A7\"/></rPr><t>■</t></r><r><t xml:space=\"preserve\"> </t></r><r><rPr><color rgb=\"FFDDDCD9\"/></rPr><t>■</t></r><r><t xml:space=\"preserve\">\n </t></r><r><rPr><color rgb=\"FF76757A\"/></rPr><t>■</t></r><r><t xml:space=\"preserve\"> </t></r><r><rPr><color rgb=\"FFED333B\"/></rPr><t>■</t></r><r><t xml:space=\"preserve\"> </t></r><r><rPr><color rgb=\"FF33D079\"/></rPr><t>■</t></r><r><t xml:space=\"preserve\"> </t></r><r><rPr><color rgb=\"FFF6D22C\"/></rPr><t>■</t></r><r><t xml:space=\"preserve\"> </t></r><r><rPr><color rgb=\"FF3583E4\"/></rPr><t>■</t></r><r><t xml:space=\"preserve\"> </t></r><r><rPr><color rgb=\"FFBF60CA\"/></rPr><t>■</t></r><r><t xml:space=\"preserve\"> </t></r><r><rPr><color rgb=\"FF26C6DA\"/></rPr><t>■</t></r><r><t xml:space=\"preserve\"> </t></r><r><rPr><color rgb=\"FFF6F5F4\"/></rPr><t>■</t></r></si>"
            @test XLSX.get_workbook(s).sst.shared_strings[Int(XLSX.getcell(s, "A4").value)+1] == "<si><r><t xml:space=\"preserve\">\nThe basic font-style attributes are </t></r><r><rPr><b/></rPr><t>bold</t></r><r><t xml:space=\"preserve\">, </t></r><r><t>light</t></r><r><t xml:space=\"preserve\">, </t></r><r><rPr><i/></rPr><t>italic</t></r><r><t xml:space=\"preserve\">,\n </t></r><r><rPr><u/></rPr><t>underline</t></r><r><t xml:space=\"preserve\">, and </t></r><r><rPr><strike/></rPr><t>strikethrough</t></r></si>"
            @test XLSX.get_workbook(s).sst.shared_strings[Int(XLSX.getcell(s, "A5").value)+1] == "<si><r><rPr><b/><color rgb=\"FFA51C2C\"/><sz val=\"24\"/></rPr><t>Hello</t></r><r><rPr><color rgb=\"FFA51C2C\"/></rPr><t xml:space=\"preserve\"> </t></r><r><rPr><color rgb=\"FFA51C2C\"/></rPr><t>Tim!</t></r></si>"
            @test XLSX.get_workbook(s).sst.shared_strings[Int(XLSX.getcell(s, "A6").value)+1] == "<si>\n  <t>this is orange text</t>\n</si>"
            @test XLSX.getFont(s, "A6").font == Dict("name" => Dict("val" => "Calibri"), "sz" => Dict("val" => "12"), "color" => Dict("rgb" => "FFFF7700"))
            @test XLSX.get_workbook(s).sst.shared_strings[Int(XLSX.getcell(s, "A7").value)+1] == "<si><r><rPr><b/><i/><u/></rPr><t xml:space=\"preserve\">deleted </t></r><r><rPr><b/><i/><strike/><u/></rPr><t xml:space=\"preserve\">and more </t></r><r><rPr><b/><i/><strike/><sz val=\"30\"/><u/></rPr><t>too</t></r><r><rPr><b/><i/><u/></rPr><t>, aswell</t></r><r><t>. And now unbolded!</t></r></si>"
            SAVE_FILES && save_outfile(f)
        end
        @testset "styledStrings - defaults and edge cases" begin

            # Runs with no attributes have `atts === nothing`, not an empty Dict.
            _hasatt(r, k) = !isnothing(r.atts) && haskey(r.atts, k)
            _getatt(r, k, d=false) = isnothing(r.atts) ? d : get(r.atts, k, d)

            # Empty annotated strings must not throw (they yield one empty
            # region on some Julia versions and none on others).
            f = XLSX.newxlsx()
            s = f[1]
            s["A1"] = styled""
            s["A2"] = styled"{red:}"
            @test ismissing(s["A1"])
            @test ismissing(s["A2"])

            # Substrings of annotated strings keep their styling, including
            # annotations clipped by the substring boundary.
            ss = styled"The {bold:quick brown} fox"
            s["B1"] = ss[5:15]            # "quick brown", fully inside the bold run
            s["B2"] = ss[9:19]            # "k brown fox", straddling the closing edge
            @test s["B1"] == "quick brown"
            @test s["B2"] == "k brown fox"
            b2 = XLSX.getRichTextString(s, "B2")
            @test length(b2.runs) == 2
            @test b2.runs[1].text == "k brown"
            @test _getatt(b2.runs[1], :bold) == true
            @test !_getatt(b2.runs[2], :bold)

            # The default height is inherited from the cell font, not baked
            # into every run at 12pt.
            s["C1"] = styled"plain {bold:bold} plain"
            c1 = XLSX.getRichTextString(s, "C1")
            @test all(r -> !_hasatt(r, :size), c1.runs)
            @test !occursin("<sz", XLSX.get_workbook(s).sst.shared_strings[
                Int(XLSX.getcell(s, "C1").value) + 1])

            # A non-default height is still emitted. The trailing unstyled text
            # keeps this a rich text string rather than collapsing to a cell font.
            s["C2"] = styled"{(height=180):big} normal"
            c2 = XLSX.getRichTextString(s, "C2")
            @test c2.runs[1].atts[:size] == 18
            @test !_hasatt(c2.runs[2], :size)

            # A cell font applied over a styled string survives it.
            s["C3"] = styled"plain {bold:bold} {(height=180):big}"
            XLSX.setFont(s, "C3"; name = "Aptos", size = 14)
            c3 = XLSX.getRichTextString(s, "C3")
            @test XLSX.getFont(s, "C3").font["sz"]["val"] == "14"
            @test all(r -> !_hasatt(r, :size), c3.runs)   # explicit 18 stripped by setFont
            @test _getatt(c3.runs[2], :bold) == true      # untouched by setFont

            # The default font name is still dropped, now via the default
            # face rather than a hardcoded "monospace".
            s["C4"] = styled"{(font=Consolas):code} normal"
            c4 = XLSX.getRichTextString(s, "C4")
            @test c4.runs[1].atts[:name] == "Consolas"
            @test !_hasatt(c4.runs[2], :name)

            # `setFont` strips the attributes it sets from the runs, so a cell-wide
            # font name overrides an explicit run-level one.
            s["C5"] = styled"{(font=Consolas):code} normal"
            XLSX.setFont(s, "C5"; name = "Aptos")
            @test s["C5"] == "code normal"
            @test isnothing(XLSX.getRichTextString(s, "C5"))
            @test XLSX.getFont(s, "C5").font["name"]["val"] == "Consolas"


            # Two identically-formatted runs collapse to a cell-level font.
            s["D1"] = XLSX.RichTextString(
                XLSX.RichTextRun("aa", [:bold => true, :size => 20]),
                XLSX.RichTextRun("bb", [:bold => true, :size => 20]),
            )
            XLSX.setFont(s, "D1"; size = 14)
            @test s["D1"] == "aabb"
            @test isnothing(XLSX.getRichTextString(s, "D1"))
            fnt = XLSX.getFont(s, "D1").font
            @test fnt["sz"]["val"] == "20"      # run wins over setFont — see C5
            @test haskey(fnt, "b")

            # Entities survive the strip-and-rewrite round trip.
            s["D2"] = styled"a & b {bold:c & d} e & f"
            @test s["D2"] == "a & b c & d e & f"
            XLSX.setFont(s, "D2"; italic = true)
            @test s["D2"] == "a & b c & d e & f"
            raw = XLSX.get_workbook(s).sst.shared_strings[Int(XLSX.getcell(s, "D2").value) + 1]
            @test !occursin("&amp;amp;", raw)          # not double-escaped
            @test count("&amp;", raw) == 3             # all three ampersands escaped
            @test !occursin(" & ", raw)                # none left bare

            s["E1"] = styled"{bold:first}"
            XLSX.setFont(s, "E1"; size = 14)
            ft = XLSX.getFont(s, "E1").font
            s["E1"] = styled"plain {italic:second}"        # mixed → stays rich text
            @test s["E1"] == "plain second"
            @test XLSX.getFont(s, "E1").font == ft

            s["E2"] = "text"
            XLSX.setFont(s, "E2"; size = 20)
            ft = XLSX.getFont(s, "E2").font
            s["E2"] = styled""
            @test XLSX.getFont(s, "E2").font == ft


            dflt = XLSX.getDefaultFont(s).font
            s["E3"] = styled"{italic:all italic}"
            @test XLSX.getFont(s, "E3").font["sz"] == dflt["sz"]
            @test haskey(XLSX.getFont(s, "E3").font, "i")

            s["A1"] = "x"; XLSX.setFont(s, "A1"; name = "Arial", size = 14, italic = true)
            s["A1"] = styled"{bold:all bold}"                       # one run: merged into the cell font
            ft = XLSX.getFont(s, "A1").font
            @test ft["name"]["val"] == "Arial" && ft["sz"]["val"] == "14"
            @test haskey(ft, "i") && haskey(ft, "b")

            # A colour naming a face without its own foreground resolves through
            # inheritance, or to no colour at all, rather than throwing.
            _sst(ref) = XLSX.get_workbook(s).sst.shared_strings[Int(XLSX.getcell(s, ref).value) + 1]
            s["F1"] = styled"{(foreground=highlight):x} y"          # highlight inherits emphasis (blue)
            @test s["F1"] == "x y"
            @test occursin("<color rgb=\"FF195EB3\"/>", _sst("F1"))
            s["F2"] = styled"{(foreground=bold):x} y"               # bold has no colour
            @test s["F2"] == "x y"
            @test !occursin("<color", _sst("F2"))
            StyledStrings.addface!(:XLSXInhOnly => StyledStrings.Face(inherit = [:red]))
            s["F3"] = styled"{(foreground=XLSXInhOnly):x} y"
            @test s["F3"] == "x y"
            @test occursin("<color rgb=\"FFA51C2C\"/>", _sst("F3"))

            # Faces whose foregrounds name each other give no colour rather than a
            # stack overflow, while a long but finite chain still resolves.
            StyledStrings.addface!(:XLSXCycA => StyledStrings.Face(foreground = :XLSXCycB))
            StyledStrings.addface!(:XLSXCycB => StyledStrings.Face(foreground = :XLSXCycA))
            s["F4"] = styled"{XLSXCycA:x} y"
            @test s["F4"] == "x y"
            @test !occursin("<color", _sst("F4"))
            s["F5"] = styled"{(foreground=XLSXCycA):x} y"
            @test s["F5"] == "x y"
            @test !occursin("<color", _sst("F5"))
            for i in 1:4
                StyledStrings.addface!(Symbol("XLSXChain$i") => StyledStrings.Face(foreground = Symbol("XLSXChain$(i+1)")))
            end
            StyledStrings.addface!(:XLSXChain5 => StyledStrings.Face(foreground = :red))
            s["F6"] = styled"{XLSXChain1:x} y"
            @test occursin("<color rgb=\"FFA51C2C\"/>", _sst("F6"))
            SAVE_FILES && save_outfile(f)
        end
    end
    @testset "RichTextString" begin
        f=XLSX.newxlsx()
        s=f[1]
        rtf1=XLSX.RichTextRun("Hello", (:vertAlign => "subscript"))
        rtf2=XLSX.RichTextRun(" Kitty ", (:color => "green", :size => 14, :bold => true, :under => true))
        rtf3=XLSX.RichTextRun("Hello", [:color => "green", :size => 14, :under => true])
        s["A1"] = XLSX.RichTextString(rtf1, rtf2, rtf3)
        @test XLSX.getRichTextString(s, "A1").runs[1].atts == Dict(:vertAlign => "subscript")
        @test XLSX.getRichTextString(s, "A1") == XLSX.RichTextString(rtf1, rtf2, rtf3)

        rtf4=XLSX.RichTextRun("Hell", (color = "red", size = 18, name = "Times New Roman"))
        rtf5=XLSX.RichTextRun("o", [:color => "green", :size => 24, :vertAlign => "superscript", :name => "Arial"])
        rtf6=XLSX.RichTextRun(" Kitt", [:color => "blue", :size => 12, :name => "Consolas"])
        rtf7=XLSX.RichTextRun("y", [:color => "green", :size => 14, :vertAlign => "subscript"])
        s["A2"] = XLSX.RichTextString(rtf4, rtf5, rtf6, rtf7)
        @test XLSX.getRichTextString(s, "A2").runs[1].atts == Dict(:color => XLSX.get_color("red"), :size => 18, :name => "Times New Roman")
        @test XLSX.getRichTextString(s, "A2") == XLSX.RichTextString(rtf4, rtf5, rtf6, rtf7)

        @test XLSX.getRichTextString(s, "A1")*XLSX.getRichTextString(s, "A2") == XLSX.RichTextString(rtf1, rtf2, rtf3, rtf4, rtf5, rtf6, rtf7)
        @test string((XLSX.getRichTextString(s, "A1")*XLSX.getRichTextString(s, "A2"))[8:24]) == "itty HelloHello K"
        @test length((XLSX.getRichTextString(s, "A1")*XLSX.getRichTextString(s, "A2"))[8:24]) == 17
        rtf1=XLSX.RichTextRun("single run", [:color => "green", :size => 14, :vertAlign => "subscript"])
        s["A3"] = XLSX.RichTextString(rtf1)

        @test s["A1"] == "Hello Kitty Hello"
        @test s["A2"] == "Hello Kitty"
        @test s["A3"] == "single run"
        @test XLSX.get_workbook(s).sst.shared_strings[Int(XLSX.getcell(s, "A1").value)+1] == "<si><r><rPr><vertAlign val=\"subscript\"/></rPr><t>Hello</t></r><r><rPr><b/><color rgb=\"FF008000\"/><sz val=\"14\"/><u/></rPr><t xml:space=\"preserve\"> Kitty </t></r><r><rPr><color rgb=\"FF008000\"/><sz val=\"14\"/><u/></rPr><t>Hello</t></r></si>"
        @test XLSX.get_workbook(s).sst.shared_strings[Int(XLSX.getcell(s, "A2").value)+1] == "<si><r><rPr><rFont val=\"Times New Roman\"/><color rgb=\"FFFF0000\"/><sz val=\"18\"/></rPr><t>Hell</t></r><r><rPr><rFont val=\"Arial\"/><color rgb=\"FF008000\"/><sz val=\"24\"/><vertAlign val=\"superscript\"/></rPr><t>o</t></r><r><rPr><rFont val=\"Consolas\"/><color rgb=\"FF0000FF\"/><sz val=\"12\"/></rPr><t xml:space=\"preserve\"> Kitt</t></r><r><rPr><color rgb=\"FF008000\"/><sz val=\"14\"/><vertAlign val=\"subscript\"/></rPr><t>y</t></r></si>"
        @test XLSX.get_workbook(s).sst.shared_strings[Int(XLSX.getcell(s, "A3").value)+1] == "<si>\n  <t>single run</t>\n</si>"
        @test XLSX.getFont(s, "A3").font == Dict("name" => Dict("val" => "Calibri"), "sz" => Dict("val" => "14"), "color" => Dict("rgb" => "FF008000"), "vertAlign" => Dict("val" => "subscript"))

        rts_expected_result = "RichTextString: \"Hello Kitty\" \n" *
                              " containing 4 runs:\n" *
                              " Run text                 Run attributes\n" *
                              " -------------------------------------------------------------------------------------------\n" *
                              " \"Hell\"                   [:color => \"FFFF0000\", :name => \"Times New Roman\", :size => 18]   \n" *
                              " \"o\"                      [:color => \"FF008000\", :name => \"Arial\", :size => 24, :vertAlign…]\n" *
                              " \" Kitt\"                  [:color => \"FF0000FF\", :name => \"Consolas\", :size => 12]          \n" *
                              " \"y\"                      [:color => \"FF008000\", :size => 14, :vertAlign => \"subscript\"]    \n"
        rtr_expected_result = "RichTextRun (\"Hell\"  [:color => \"FFFF0000\", :name => \"Times New Roman\", :size => 18])"
        io1 = IOBuffer()
        io2 = IOBuffer()
        show(io1, XLSX.getRichTextString(s, "A2"))
        @test String(take!(io1)) == rts_expected_result
        show(io1, XLSX.getRichTextString(s, "A2").runs[1])
        @test String(take!(io1)) == rtr_expected_result

        rt1 = XLSX.RichTextRun("Water is H")
        rt2 = XLSX.RichTextRun("2", :vertAlign => "subscript")
        rt3 = XLSX.RichTextRun("O!")
        rts_expected_result = "RichTextString: \"Water is H2O!\" \n" *
                               " containing 3 runs:\n" *
                               " Run text                 Run attributes\n" *
                               " -------------------------------------------------------------------------------------------\n" *
                               " \"Water is H\"             [ ]                                                               \n" *
                               " \"2\"                      [:vertAlign => \"subscript\"]                                       \n" *
                               " \"O!\"                     [ ]                                                               \n"
        io1 = IOBuffer()
        show(io1, XLSX.RichTextString(rt1, rt2, rt3))
        @test String(take!(io1)) == rts_expected_result
        SAVE_FILES && save_outfile(f)

    end

    @testset "RichTextRun and RichTextString as strings" begin
        run = XLSX.RichTextRun("ab", nothing)
        @test length(run) == 2
        @test hash(run) == hash(XLSX.RichTextRun("ab"))
        rts = XLSX.RichTextString("ab", [run])
        @test String(rts) == "ab" && ncodeunits(rts) == 2
        @test codeunit(rts) == UInt8 && codeunit(rts, 1) == UInt8('a')
        @test isvalid(rts, 1) && rts[1] == XLSX.RichTextString("a", [XLSX.RichTextRun("a")])
        @test hash(rts) == hash(XLSX.RichTextString("ab", [XLSX.RichTextRun("ab")]))

        xf = XLSX.opentemplate(joinpath(data_directory, "is.xlsx"))
        sh = xf["Sheet1"]
        XLSX.setFont(sh, "A1"; name="Palatino")
        @test XLSX.getRichTextString(sh, 1, 1) == XLSX.getRichTextString(sh, "A1")
        @test XLSX.getRichTextString(xf, "Sheet1!A1") == XLSX.getRichTextString(sh, "A1")
        @test XLSX.get_sst(xf) === XLSX.get_sst(XLSX.get_workbook(xf))
        SAVE_FILES && save_outfile(xf)
    end

    @testset "rich text: every color form" begin
        xf = XLSX.newxlsx()
        s = xf[1]
        sst(ref) = XLSX.get_workbook(s).sst.shared_strings[Int(XLSX.getcell(s, ref).value)+1]

        runs = [XLSX.RichTextRun("a", [:color => :coral]),
                XLSX.RichTextRun("b", [:color => colorant"seagreen"]),
                XLSX.RichTextRun("c", [:color => XLSX.SchemeColor(:accent1; lumMod = 60, lumOff = 40)]),
                XLSX.RichTextRun("d", [:color => "accent2"]),
                XLSX.RichTextRun("e", [:color => :tx1])]
        s["A1"] = XLSX.RichTextString(runs)
        @test sst("A1") == "<si><r><rPr><color rgb=\"FFFF7F50\"/></rPr><t>a</t></r>" *
                           "<r><rPr><color rgb=\"FF2E8B57\"/></rPr><t>b</t></r>" *
                           "<r><rPr><color theme=\"4\" tint=\"0.4\"/></rPr><t>c</t></r>" *
                           "<r><rPr><color theme=\"5\"/></rPr><t>d</t></r>" *
                           "<r><rPr><color theme=\"1\"/></rPr><t>e</t></r></si>"

        # Reading resolves theme colors to RGB, as before
        rts = XLSX.getRichTextString(s, "A1")
        @test rts.runs[1].atts[:color] == "FFFF7F50"
        @test rts.runs[3].atts[:color] isa String

        # Equality and hashing treat every spelling of one color alike
        r(c) = XLSX.RichTextRun("x", [:color => c])
        @test r("red") == r(:red) == r(colorant"red") == r("FFFF0000")
        @test hash(r("red")) == hash(r(colorant"red"))
        @test r(:accent1) == r(XLSX.SchemeColor(:accent1)) == r("accent1")
        @test hash(r(:accent1)) == hash(r(XLSX.SchemeColor(:accent1)))
        @test r(:tx1) == r(:dk1)
        @test r(:accent1) != r("red")

        # A chart-only theme color (shade) is a valid run color: comparing it does not throw,
        # but a cell cannot hold it
        shaded = r(XLSX.SchemeColor(:accent1; shade = 50))
        @test shaded == r(XLSX.SchemeColor(:accent1; shade = 50))
        @test shaded != r(:accent1)
        @test_throws XLSX.XLSXError (s["A2"] = XLSX.RichTextString([shaded, XLSX.RichTextRun("y")]))

        # An invalid color is caught when the run is built
        @test_throws XLSX.XLSXError XLSX.RichTextRun("x", [:color => "notacolor"])
        @test_throws XLSX.XLSXError XLSX.RichTextRun("x", [:color => 42])

        # A single run is applied to the whole cell with setFont
        s["A3"] = XLSX.RichTextString([XLSX.RichTextRun("one", [:color => XLSX.SchemeColor(:accent3; lumMod = 75)])])
        @test XLSX.getFont(s, "A3").font["color"] == Dict("theme" => "6", "tint" => "-0.25")

        # In chart text, theme colors stay theme colors
        @test XLSX.Charts._rich_run_props(Dict(:color => :accent1)).fill.fgcolor == XLSX.SchemeColor(:accent1)
        @test XLSX.Charts._rich_run_props(Dict(:color => XLSX.SchemeColor(:accent1; shade = 50))).fill.fgcolor ==
              XLSX.SchemeColor(:accent1; shade = 50)
        @test XLSX.Charts._rich_run_props(Dict(:color => colorant"red")).fill.fgcolor == "FFFF0000"
        SAVE_FILES && save_outfile(xf)
    end
end

@testset "review fixes: rich text" begin
    f = XLSX.newxlsx(); s = f[1]
    s["A1"] = XLSX.RichTextString(XLSX.RichTextRun("plain"))
    s["A2"] = XLSX.RichTextString(XLSX.RichTextRun("u", [:under => true, :size => 12.5]))
    @test s["A1"] == "plain" && s["A2"] == "u"
    @test XLSX.getFont(s, "A2").font["u"] === nothing
    @test XLSX.getFont(s, "A2").font["sz"]["val"] == "12.5"
    SAVE_FILES && save_outfile(f)

    rt = XLSX.getRichTextString(f.workbook, "<si><r><rPr><sz val=\"10.5\"/><b val=\"0\"/><u val=\"none\"/></rPr><t>ab</t></r><r><rPr><b/></rPr><t>cd</t></r></si>")
    @test rt.runs[1].atts == Dict{Symbol,Any}(:size => 10.5)
    @test rt.runs[2].atts == Dict{Symbol,Any}(:bold => true)

    r = XLSX.RichTextString(XLSX.RichTextRun("é"), XLSX.RichTextRun("ab", [:bold => true]))
    @test r[3:4] == XLSX.RichTextString(XLSX.RichTextRun("ab", [:bold => true]))
    @test r[1:3].runs == [XLSX.RichTextRun("é"), XLSX.RichTextRun("a", [:bold => true])]

    a = XLSX.RichTextRun("x", [:bold => true, :size => 12, :italic => true, :name => "Arial", :color => "red", :under => true, :strike => true, :vertAlign => "subscript"])
    b = XLSX.RichTextRun("x", [:vertAlign => "subscript", :strike => true, :under => true, :color => "red", :name => "Arial", :italic => true, :size => 12, :bold => true])
    @test a == b && hash(a) == hash(b)
end

# A reader normalises a CR in text to LF, and CR, LF and tab in an attribute value to a
# space, unless they are written as character references (#482).
@testset "CR, LF and tab survive a round trip" begin
    vals = ["a\r\nb", "a\rb", "a\nb", "\r", "x\ty"]
    f = XLSX.newxlsx(); s = f[1]
    for (i, v) in enumerate(vals)
        s[i, 1] = v
        s[i, 2] = XLSX.RichTextString(XLSX.RichTextRun(v, [:bold => true]), XLSX.RichTextRun("!"))
    end
    s["D1"] = "a\r\nb"; s["E1"] = "c\td"; s["D2"] = 1; s["E2"] = 2
    XLSX.addtable!(s, "D1:E2"; name = "T")
    XLSX.setFormula(s, "C1", "=1+\r\n2")
    out = "cr_round_trip.xlsx"
    XLSX.writexlsx(out, f; overwrite = true)

    sst = check_sst(out)
    @test occursin("a&#13;\nb", sst) && !occursin("\r", sst)
    zip = ZipReader(read(out))
    tbl = String(zip_readentry(zip, only(filter(n -> occursin(r"^xl/tables/", n), zip_names(zip)))))
    @test occursin("name=\"a&#13;&#10;b\"", tbl) && occursin("name=\"c&#9;d\"", tbl)
    # Excel refuses these two, so they are written differently:
    # a run of only "\r" keeps xml:space, and a formula's line breaks become LF
    @test occursin("<t xml:space=\"preserve\">&#13;</t>", sst)
    sheet = String(zip_readentry(zip, "xl/worksheets/sheet1.xml"))
    @test occursin("1+\n2", sheet) && !occursin("&#13;", sheet)
    @test XLSX.getFormula(s, "C1") == "=1+\n2"

    for xf in (XLSX.readxlsx(out), XLSX.openxlsx(out; mode = "rw"))
        s2 = xf[1]
        @test [s2[i, 1] for i in eachindex(vals)] == vals
        @test [s2[i, 2] for i in eachindex(vals)] == vals .* "!"
        @test XLSX.getRichTextString(s2, "B1").runs[1] == XLSX.RichTextRun("a\r\nb", [:bold => true])
        @test XLSX.table(s2, "T").columns == ["a\r\nb", "c\td"]
        xf.is_writable && @test XLSX.getFormula(s2, "C1") == "=1+\n2"
    end
    SAVE_FILES && save_outfile(f)
    isfile(out) && rm(out)
end

# A shared string table with CR line ends reads as the same table with LF ones (#482).
@testset "shared strings with CR line ends" begin
    f = XLSX.newxlsx(); s = f[1]
    for i in 1:4; s[i, 1] = "s$i"; end
    src = "sst_eol_src.xlsx"
    XLSX.writexlsx(src, f; overwrite = true)
    sst(eol) = replace("""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="4" uniqueCount="4">
        <si><t>plain</t></si>
        <si><t xml:space="preserve">line1
        line2</t></si>
        <si><r><rPr><b/></rPr><t>bold</t></r><r><t xml:space="preserve"> rest</t></r></si>
        <si><t>漢字</t><rPh sb="0" eb="2"><t>かんじ</t></rPh></si>
        </sst>""", "\n" => eol)
    expected = ["plain", "line1\nline2", "bold rest", "漢字"]
    for eol in ("\n", "\r\n", "\r")
        p = patch_xlsx_entry(src, "xl/sharedStrings.xml", _ -> sst(eol))
        @test [XLSX.readxlsx(p)[1][i, 1] for i in 1:4] == expected
        xf = XLSX.openxlsx(p; mode = "rw")
        @test [xf[1][i, 1] for i in 1:4] == expected
        raw = XLSX.get_sst(xf.workbook).shared_strings
        # the raw XML is the part's own bytes, its line ends included (XML.jl#194)
        @test raw[2] == "<si><t xml:space=\"preserve\">line1$(eol)line2</t></si>"
        @test raw[4] == "<si><t>漢字</t><rPh sb=\"0\" eb=\"2\"><t>かんじ</t></rPh></si>"
        # written back unchanged, and read the same again
        out = "sst_eol_out.xlsx"
        XLSX.writexlsx(out, xf; overwrite = true)
        @test [XLSX.readxlsx(out)[1][i, 1] for i in 1:4] == expected
        SAVE_FILES && save_outfile(out)
        rm(p); rm(out)
    end
    rm(src)
end
