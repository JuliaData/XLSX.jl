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
end

@testset "rich text formats" begin
    @testset "setAttributes" begin
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
        @test XLSX.getFont(sh, "A1").font == Dict("name" => Dict("val" => "Palatino"), "sz" => Dict("val" => "11"), "color" => Dict("theme" => "1"))
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
            @test XLSX.get_workbook(s).sst.shared_strings[Int(XLSX.getcell(s, "A1").value)+1] == "<si><r><rPr><color rgb=\"FFE5A509\"/><sz val=\"12\"/></rPr><t>hello</t></r><r><rPr><sz val=\"12\"/></rPr><t xml:space=\"preserve\"> </t></r><r><rPr><color rgb=\"FF195EB3\"/><sz val=\"12\"/></rPr><t>there</t></r></si>"
            @test XLSX.get_workbook(s).sst.shared_strings[Int(XLSX.getcell(s, "A2").value)+1] == "<si><r><rPr><sz val=\"12\"/></rPr><t xml:space=\"preserve\">The </t></r><r><rPr><b/><i/><sz val=\"12\"/></rPr><t xml:space=\"preserve\">quick </t></r><r><rPr><b/><i/><color rgb=\"FFCD853F\"/><sz val=\"12\"/></rPr><t>brown</t></r><r><rPr><b/><i/><sz val=\"12\"/></rPr><t xml:space=\"preserve\"> fox</t></r><r><rPr><b/><sz val=\"12\"/></rPr><t xml:space=\"preserve\"> jumps over the </t></r><r><rPr><b/><color rgb=\"FFFFC000\"/><sz val=\"12\"/></rPr><t>lazy</t></r><r><rPr><b/><sz val=\"12\"/></rPr><t xml:space=\"preserve\"> dog</t></r></si>"
            @test XLSX.get_workbook(s).sst.shared_strings[Int(XLSX.getcell(s, "A3").value)+1] == "<si><r><rPr><sz val=\"12\"/></rPr><t xml:space=\"preserve\">In terms of color, we have named faces for the 16 standard terminal colors:\n </t></r><r><rPr><color rgb=\"FF1C1A23\"/><sz val=\"12\"/></rPr><t>■</t></r><r><rPr><sz val=\"12\"/></rPr><t xml:space=\"preserve\"> </t></r><r><rPr><color rgb=\"FFA51C2C\"/><sz val=\"12\"/></rPr><t>■</t></r><r><rPr><sz val=\"12\"/></rPr><t xml:space=\"preserve\"> </t></r><r><rPr><color rgb=\"FF25A268\"/><sz val=\"12\"/></rPr><t>■</t></r><r><rPr><sz val=\"12\"/></rPr><t xml:space=\"preserve\"> </t></r><r><rPr><color rgb=\"FFE5A509\"/><sz val=\"12\"/></rPr><t>■</t></r><r><rPr><sz val=\"12\"/></rPr><t xml:space=\"preserve\"> </t></r><r><rPr><color rgb=\"FF195EB3\"/><sz val=\"12\"/></rPr><t>■</t></r><r><rPr><sz val=\"12\"/></rPr><t xml:space=\"preserve\"> </t></r><r><rPr><color rgb=\"FF803D9B\"/><sz val=\"12\"/></rPr><t>■</t></r><r><rPr><sz val=\"12\"/></rPr><t xml:space=\"preserve\"> </t></r><r><rPr><color rgb=\"FF0097A7\"/><sz val=\"12\"/></rPr><t>■</t></r><r><rPr><sz val=\"12\"/></rPr><t xml:space=\"preserve\"> </t></r><r><rPr><color rgb=\"FFDDDCD9\"/><sz val=\"12\"/></rPr><t>■</t></r><r><rPr><sz val=\"12\"/></rPr><t xml:space=\"preserve\">\n </t></r><r><rPr><color rgb=\"FF76757A\"/><sz val=\"12\"/></rPr><t>■</t></r><r><rPr><sz val=\"12\"/></rPr><t xml:space=\"preserve\"> </t></r><r><rPr><color rgb=\"FFED333B\"/><sz val=\"12\"/></rPr><t>■</t></r><r><rPr><sz val=\"12\"/></rPr><t xml:space=\"preserve\"> </t></r><r><rPr><color rgb=\"FF33D079\"/><sz val=\"12\"/></rPr><t>■</t></r><r><rPr><sz val=\"12\"/></rPr><t xml:space=\"preserve\"> </t></r><r><rPr><color rgb=\"FFF6D22C\"/><sz val=\"12\"/></rPr><t>■</t></r><r><rPr><sz val=\"12\"/></rPr><t xml:space=\"preserve\"> </t></r><r><rPr><color rgb=\"FF3583E4\"/><sz val=\"12\"/></rPr><t>■</t></r><r><rPr><sz val=\"12\"/></rPr><t xml:space=\"preserve\"> </t></r><r><rPr><color rgb=\"FFBF60CA\"/><sz val=\"12\"/></rPr><t>■</t></r><r><rPr><sz val=\"12\"/></rPr><t xml:space=\"preserve\"> </t></r><r><rPr><color rgb=\"FF26C6DA\"/><sz val=\"12\"/></rPr><t>■</t></r><r><rPr><sz val=\"12\"/></rPr><t xml:space=\"preserve\"> </t></r><r><rPr><color rgb=\"FFF6F5F4\"/><sz val=\"12\"/></rPr><t>■</t></r></si>"
            @test XLSX.get_workbook(s).sst.shared_strings[Int(XLSX.getcell(s, "A4").value)+1] == "<si><r><rPr><sz val=\"12\"/></rPr><t xml:space=\"preserve\">\nThe basic font-style attributes are </t></r><r><rPr><b/><sz val=\"12\"/></rPr><t>bold</t></r><r><rPr><sz val=\"12\"/></rPr><t xml:space=\"preserve\">, </t></r><r><rPr><sz val=\"12\"/></rPr><t>light</t></r><r><rPr><sz val=\"12\"/></rPr><t xml:space=\"preserve\">, </t></r><r><rPr><i/><sz val=\"12\"/></rPr><t>italic</t></r><r><rPr><sz val=\"12\"/></rPr><t xml:space=\"preserve\">,\n </t></r><r><rPr><sz val=\"12\"/><u/></rPr><t>underline</t></r><r><rPr><sz val=\"12\"/></rPr><t xml:space=\"preserve\">, and </t></r><r><rPr><strike/><sz val=\"12\"/></rPr><t>strikethrough</t></r></si>"
            @test XLSX.get_workbook(s).sst.shared_strings[Int(XLSX.getcell(s, "A5").value)+1] == "<si><r><rPr><b/><color rgb=\"FFA51C2C\"/><sz val=\"24\"/></rPr><t>Hello</t></r><r><rPr><color rgb=\"FFA51C2C\"/><sz val=\"12\"/></rPr><t xml:space=\"preserve\"> </t></r><r><rPr><color rgb=\"FFA51C2C\"/><sz val=\"12\"/></rPr><t>Tim!</t></r></si>"
            @test XLSX.get_workbook(s).sst.shared_strings[Int(XLSX.getcell(s, "A6").value)+1] == "<si>\n  <t>this is orange text</t>\n</si>"
            @test XLSX.getFont(s, "A6").font == Dict("name" => Dict("val" => "Calibri"), "sz" => Dict("val" => "12"), "color" => Dict("rgb" => "FFFF7700"))
            @test XLSX.get_workbook(s).sst.shared_strings[Int(XLSX.getcell(s, "A7").value)+1] == "<si><r><rPr><b/><i/><sz val=\"12\"/><u/></rPr><t xml:space=\"preserve\">deleted </t></r><r><rPr><b/><i/><strike/><sz val=\"12\"/><u/></rPr><t xml:space=\"preserve\">and more </t></r><r><rPr><b/><i/><strike/><sz val=\"30\"/><u/></rPr><t>too</t></r><r><rPr><b/><i/><sz val=\"12\"/><u/></rPr><t>, aswell</t></r><r><rPr><sz val=\"12\"/></rPr><t>. And now unbolded!</t></r></si>"
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
end