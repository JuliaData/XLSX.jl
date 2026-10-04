@testset "DrawingML colours" begin

    # A workbook is needed only for schemeClr resolution; any file with a
    # standard theme will do.
    f = XLSX.opentemplate(joinpath(data_directory, "chart_basic.xlsx"))
    wb = XLSX.get_workbook(f)

    # Parse a colour from an XML fragment, as it appears inside a fill.
    function color_of(xml::String)
        doc = parse(xml, XML.Node)
        XLSX.Charts.parse_drawing_color(wb, XLSX.xml_root_element(doc))
    end

    @testset "theme colour map" begin
        m = XLSX.get_theme_color_map(wb)
        for name in ("dk1", "lt1", "dk2", "lt2", "accent1", "accent2", "accent3",
            "accent4", "accent5", "accent6", "hlink", "folHlink")
            @test haskey(m, name)
            @test occursin(r"^[0-9A-Fa-f]{6}$", m[name])
        end

        # DrawingML aliases, not index-swapped like the spreadsheet ordering.
        @test m["tx1"] == m["dk1"]
        @test m["bg1"] == m["lt1"]
        @test m["tx2"] == m["dk2"]
        @test m["bg2"] == m["lt2"]

        # Cached: the same dict object comes back.
        @test XLSX.get_theme_color_map(wb) === m
    end

    @testset "srgbClr" begin
        c = color_of("""<a:srgbClr xmlns:a="$(XLSX.NS_A)" val="FF0000"/>""")
        @test c.kind === :srgb
        @test c.val == "FF0000"
        @test c.rgb == "FF0000"
        @test c.alpha == 1.0
        @test isempty(c.transforms)
    end

    @testset "schemeClr resolves through the theme" begin
        c = color_of("""<a:schemeClr xmlns:a="$(XLSX.NS_A)" val="accent1"/>""")
        @test c.kind === :scheme
        @test c.val == "accent1"                       # the reference is kept as authored
        @test c.rgb == uppercase(XLSX.get_theme_color_map(wb)["accent1"])
        @test isempty(c.transforms)
    end

    @testset "lumMod and lumOff apply in document order" begin
        # "Accent 1, Lighter 40%" as Excel writes it.
        c = color_of("""
            <a:schemeClr xmlns:a="$(XLSX.NS_A)" val="accent1">
                <a:lumMod val="60000"/><a:lumOff val="40000"/>
            </a:schemeClr>""")
        @test c.transforms == [:lumMod => 60000, :lumOff => 40000]
        @test c.rgb != uppercase(XLSX.get_theme_color_map(wb)["accent1"])  # it moved
        @test occursin(r"^[0-9A-F]{6}$", c.rgb)

        # Lightening raises luminance; darkening lowers it.
        base = color_of("""<a:schemeClr xmlns:a="$(XLSX.NS_A)" val="accent1"/>""")
        dark = color_of("""
           <a:schemeClr xmlns:a="$(XLSX.NS_A)" val="accent1">
               <a:lumMod val="75000"/>
           </a:schemeClr>""")
        lum(h) = convert(Colors.HSL, parse(Colors.Colorant, "#" * h)).l
        @test lum(c.rgb) > lum(base.rgb)
        @test lum(dark.rgb) < lum(base.rgb)
    end

    @testset "alpha sets alpha, not rgb" begin
        c = color_of("""
            <a:srgbClr xmlns:a="$(XLSX.NS_A)" val="00FF00">
                <a:alpha val="50000"/>
            </a:srgbClr>""")
        @test c.rgb == "00FF00"
        @test c.alpha ≈ 0.5
    end

    @testset "shade and tint go through the gamma transfer" begin
        # A 50% shade of mid grey is not 50% of its sRGB value: applying the
        # transform in gamma-encoded space would give 3F3F3F.
        c = color_of("""
            <a:srgbClr xmlns:a="$(XLSX.NS_A)" val="7F7F7F">
                <a:shade val="50000"/>
            </a:srgbClr>""")
        @test c.rgb != "3F3F3F"
        @test parse(Int, c.rgb[1:2]; base=16) > 0x3F      # lighter than the naive result

        t = color_of("""
            <a:srgbClr xmlns:a="$(XLSX.NS_A)" val="000000">
                <a:tint val="50000"/>
            </a:srgbClr>""")
        @test t.rgb != "000000"                            # tint moves black toward white
    end

    @testset "sysClr prefers lastClr" begin
        c = color_of("""<a:sysClr xmlns:a="$(XLSX.NS_A)" val="windowText" lastClr="112233"/>""")
        @test c.kind === :sys
        @test c.val == "windowText"
        @test c.rgb == "112233"
    end

    @testset "hslClr" begin
        # hue in 60000ths of a degree; sat and lum in thousandths of a percent.
        c = color_of("""<a:hslClr xmlns:a="$(XLSX.NS_A)" hue="0" sat="100000" lum="50000"/>""")
        @test c.kind === :hsl
        @test c.rgb == "FF0000"
        @test c.alpha == 1.0
        @test color_of("""<a:hslClr xmlns:a="$(XLSX.NS_A)" hue="14400000" sat="100000" lum="50000"/>""").rgb == "0000FF"
    end

    @testset "scrgbClr" begin
        # Linear RGB, thousandths of a percent per channel.
        c = color_of("""<a:scrgbClr xmlns:a="$(XLSX.NS_A)" r="100000" g="0" b="0"/>""")
        @test c.kind === :scrgb
        @test c.rgb == "FF0000"
        @test c.alpha == 1.0
        @test color_of("""<a:scrgbClr xmlns:a="$(XLSX.NS_A)" r="50000" g="50000" b="50000"/>""").rgb == "BCBCBC"   # linear 0.5 is sRGB 188
    end

    @testset "unknown transforms are no-ops" begin
        plain = color_of("""<a:srgbClr xmlns:a="$(XLSX.NS_A)" val="336699"/>""")
        c = color_of("""
            <a:srgbClr xmlns:a="$(XLSX.NS_A)" val="336699">
                <a:hueMod val="50000"/>
            </a:srgbClr>""")
        @test c.rgb == plain.rgb                # not applied...
        @test c.transforms == [:hueMod => 50000]  # ...but recorded
    end

    @testset "no colour child" begin
        doc = parse("""<a:noFill xmlns:a="$(XLSX.NS_A)"/>""", XML.Node)
        @test isnothing(XLSX.Charts.parse_drawing_color(wb, XLSX.xml_root_element(doc)))
        @test isnothing(XLSX.Charts.parse_drawing_color(wb, nothing))
    end

    @testset "parent element is searched for its colour child" begin
        doc = parse("""
            <a:solidFill xmlns:a="$(XLSX.NS_A)"><a:srgbClr val="ABCDEF"/></a:solidFill>""", XML.Node)
        c = XLSX.Charts.parse_drawing_color(wb, XLSX.xml_root_element(doc))
        @test c.rgb == "ABCDEF"
    end

    fill_of(xml) = XLSX.Charts.parse_drawing_fill(wb, XLSX.xml_root_element(XML.parse(xml, XML.Node)))

    @testset "solid fill" begin
        fl = fill_of("""
            <a:solidFill xmlns:a="$(XLSX.NS_A)"><a:srgbClr val="FF0000"/></a:solidFill>""")
        @test fl.kind === :solid
        @test fl.fgcolor.rgb == "FF0000"
        @test isnothing(fl.bgcolor)
        @test isnothing(fl.preset)
    end

    @testset "no fill" begin
        fl = fill_of("""<a:noFill xmlns:a="$(XLSX.NS_A)"/>""")
        @test fl.kind === :none
        @test isnothing(fl.fgcolor)
    end

    @testset "pattern fill exposes both colours" begin
        fl = fill_of("""
            <a:pattFill xmlns:a="$(XLSX.NS_A)" prst="pct25">
                <a:fgClr><a:srgbClr val="112233"/></a:fgClr>
                <a:bgClr><a:schemeClr val="bg1"/></a:bgClr>
            </a:pattFill>""")
        @test fl.kind === :pattern
        @test fl.preset == "pct25"
        @test fl.fgcolor.rgb == "112233"
        @test fl.bgcolor.val == "bg1"
        @test fl.bgcolor.rgb == uppercase(XLSX.get_theme_color_map(wb)["bg1"])
    end

    @testset "gradient is identified but not modelled" begin
        fl = fill_of("""
            <a:gradFill xmlns:a="$(XLSX.NS_A)">
                <a:gsLst>
                    <a:gs pos="0"><a:srgbClr val="FFFFFF"/></a:gs>
                    <a:gs pos="100000"><a:srgbClr val="000000"/></a:gs>
                </a:gsLst>
            </a:gradFill>""")
        @test fl.kind === :gradient
        @test isnothing(fl.fgcolor)          # not the first stop, which would mislead
        @test XLSX.localname(fl.raw) == "gradFill"
    end

    @testset "parent element is searched for its fill child" begin
        fl = fill_of("""
            <c:spPr xmlns:c="$(XLSX.Charts.NS_C)" xmlns:a="$(XLSX.NS_A)">
                <a:solidFill><a:srgbClr val="ABCDEF"/></a:solidFill>
            </c:spPr>""")
        @test fl.fgcolor.rgb == "ABCDEF"
    end

    @testset "no fill child" begin
        doc = XML.parse("""<c:spPr xmlns:c="$(XLSX.Charts.NS_C)"/>""", XML.Node)
        @test isnothing(XLSX.Charts.parse_drawing_fill(wb, XLSX.xml_root_element(doc)))
        @test isnothing(XLSX.Charts.parse_drawing_fill(wb, nothing))
    end

    @testset "parent element is searched for its fill child" begin
        fl = fill_of("""
            <a:spPr xmlns:a="$(XLSX.NS_A)">
                <a:solidFill><a:srgbClr val="ABCDEF"/></a:solidFill>
            </a:spPr>""")
        @test fl.fgcolor.rgb == "ABCDEF"
    end

    @testset "no fill child" begin
        doc = XML.parse("""<a:spPr xmlns:a="$(XLSX.NS_A)"/>""", XML.Node)
        @test isnothing(XLSX.Charts.parse_drawing_fill(wb, XLSX.xml_root_element(doc)))
        @test isnothing(XLSX.Charts.parse_drawing_fill(wb, nothing))
    end

    line_of(xml) = XLSX.Charts.parse_drawing_line(wb, XLSX.xml_root_element(XML.parse(xml, XML.Node)))

    @testset "solid line with width and dash" begin
        ln = line_of("""
            <a:ln xmlns:a="$(XLSX.NS_A)" w="19050" cap="rnd" cmpd="sng">
                <a:solidFill><a:srgbClr val="FF0000"/></a:solidFill>
                <a:prstDash val="dash"/>
            </a:ln>""")
        @test ln.fill.kind === :solid
        @test ln.fill.fgcolor.rgb == "FF0000"
        @test ln.width == 1.5
        @test ln.dash == "dash"          # an element, not an attribute
        @test ln.cap == "rnd"
        @test ln.compound == "sng"
    end

    @testset "line with no outline" begin
        ln = line_of("""
            <a:ln xmlns:a="$(XLSX.NS_A)"><a:noFill/></a:ln>""")
        @test ln.fill.kind === :none
        @test isnothing(ln.fill.fgcolor)
        @test isnothing(ln.width)
        @test isnothing(ln.dash)
    end

    @testset "line that sets only a width" begin
        ln = line_of("""<a:ln xmlns:a="$(XLSX.NS_A)" w="9525"/>""")
        @test isnothing(ln.fill)         # says nothing about the stroke...
        @test ln.width == 0.75
        @test isnothing(ln.cap)
        @test isnothing(ln.compound)
    end

    @testset "line colour goes through the theme" begin
        ln = line_of("""
            <a:ln xmlns:a="$(XLSX.NS_A)">
                <a:solidFill>
                    <a:schemeClr val="tx1"><a:lumMod val="15000"/><a:lumOff val="85000"/></a:schemeClr>
                </a:solidFill>
            </a:ln>""")
        @test ln.fill.fgcolor.val == "tx1"
        @test ln.fill.fgcolor.rgb == "D9D9D9"    # the standard Office gridline grey
    end

    @testset "parent element is searched for its ln child" begin
        ln = line_of("""
            <c:spPr xmlns:c="$(XLSX.Charts.NS_C)" xmlns:a="$(XLSX.NS_A)">
                <a:solidFill><a:srgbClr val="ABCDEF"/></a:solidFill>
                <a:ln w="12700"><a:solidFill><a:srgbClr val="123456"/></a:solidFill></a:ln>
            </c:spPr>""")
        # The line's own fill, not the shape's.
        @test ln.fill.fgcolor.rgb == "123456"
        @test ln.width == 1.0
    end

    @testset "no ln child" begin
        doc = XML.parse("""
            <c:spPr xmlns:c="$(XLSX.Charts.NS_C)" xmlns:a="$(XLSX.NS_A)">
                <a:solidFill><a:srgbClr val="ABCDEF"/></a:solidFill>
            </c:spPr>""", XML.Node)
        @test isnothing(XLSX.Charts.parse_drawing_line(wb, XLSX.xml_root_element(doc)))
        @test isnothing(XLSX.Charts.parse_drawing_line(wb, nothing))
        ln_bare = line_of("""<a:ln xmlns:a="$(XLSX.NS_A)"/>""")
        @test ln_bare !== nothing
        @test isnothing(ln_bare.width)
    end
    SAVE_FILES && save_outfile(f)
end

# =============================================================================
# DrawingML Text
## =============================================================================









# Parse a fragment into its root element. XML.parse returns a document node;
# the element is its last child (a declaration may precede it).
function _frag(s::String)
    doc = XML.parse(XML.Node, s)
    els = filter(c -> XML.nodetype(c) === XML.Element, XML.children(doc))
    return els[end]
end

const _NSDECL = "xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\""

@testset "DrawingML text" begin

    # -----------------------------------------------------------------------
    @testset "attribute readers" begin
        el = _frag("""<a:rPr $_NSDECL sz="1197" b="0" i="1" baseline="30000"
                      kern="1200" spc="-50" lvl="2" rot="-2700000" marL="228600"/>""")

        # units
        @test XLSX.Charts._attr_pt(el, "sz") == 11.97           # 1/100 pt
        @test XLSX.Charts._attr_pt(el, "kern") == 12.0
        @test XLSX.Charts._attr_pt(el, "spc") == -0.5
        @test XLSX.Charts._attr_pct_opt(el, "baseline") == 0.3  # thousandths of a percent -> fraction
        @test XLSX.Charts._attr_deg(el, "rot") == -45.0         # 1/60000 deg
        @test XLSX.Charts._attr_emu(el, "marL") == 18.0         # EMU -> pt
        @test XLSX.Charts._attr_int(el, "lvl") == 2

        # absent stays distinct from explicit
        @test XLSX.Charts._attr_bool(el, "b") === false         # written as 0
        @test XLSX.Charts._attr_bool(el, "i") === true
        @test XLSX.Charts._attr_bool(el, "u") === nothing       # not written at all
        @test XLSX._attr(el, "missing") === nothing
        @test XLSX.Charts._attr_pt(el, "missing") === nothing
        @test XLSX.Charts._attr_pct_opt(el, "missing") === nothing

        # tolerates a missing child element
        @test XLSX._attr(nothing, "typeface") === nothing
        @test XLSX.first_element_with_tag(nothing, "latin") === nothing

        # bool spellings
        b2 = _frag("""<a:rPr $_NSDECL b="true" i="on" u="none"/>""")
        @test XLSX.Charts._attr_bool(b2, "b") === true
        @test XLSX.Charts._attr_bool(b2, "i") === true
        @test XLSX._attr(b2, "u") == "none"              # enum stays a string

        # unparseable reads as absent, not as zero
        bad = _frag("""<a:rPr $_NSDECL sz="large"/>""")
        @test XLSX.Charts._attr_pt(bad, "sz") === nothing
    end

    # -----------------------------------------------------------------------
    @testset "spacing" begin
        pPr = _frag("""<a:pPr $_NSDECL>
                         <a:lnSpc><a:spcPct val="150000"/></a:lnSpc>
                         <a:spcBef><a:spcPts val="1200"/></a:spcBef>
                       </a:pPr>""")
        @test XLSX.Charts._attr_spacing(pPr, "lnSpc") == (:frac, 1.5)
        @test XLSX.Charts._attr_spacing(pPr, "spcBef") == (:pts, 12.0)
        @test XLSX.Charts._attr_spacing(pPr, "spcAft") === nothing
    end

    # -----------------------------------------------------------------------
    @testset "hand-built text bodies" begin
        # A workbook is needed only for theme colour resolution; any will do.
        XLSX.openxlsx(joinpath(data_directory, "chart_basic.xlsx")) do xf
            wb = XLSX.get_workbook(xf)

            @testset "formatting-only txPr (the common shape)" begin
                node = _frag("""<c:txPr xmlns:c="$(XLSX.Charts.NS_C)" $_NSDECL>
                      <a:bodyPr rot="-60000000" vert="horz" wrap="square" anchor="ctr"/>
                      <a:lstStyle/>
                      <a:p>
                        <a:pPr algn="ctr">
                          <a:defRPr sz="900" b="1" i="0" u="none" strike="noStrike"/>
                        </a:pPr>
                        <a:endParaRPr lang="en-GB"/>
                      </a:p>
                    </c:txPr>""")

                t = XLSX.Charts.parse_drawing_text(wb, node)
                @test t isa XLSX.Charts.DrawingText
                @test length(t.paragraphs) == 1
                @test isempty(t.paragraphs[1].runs)
                @test XLSX.Charts.text_content(t) == ""
                @test t.liststyle !== nothing              # preserved, even empty

                @test t.body.rotation == -1000.0
                @test t.body.anchor == "ctr"
                @test t.body.autofit === nothing           # absent, not :none

                # the font lives in defRPr, and default_run_props finds it
                p = XLSX.Charts.default_run_props(t)
                @test p !== nothing
                @test p.size == 9.0
                @test p.bold === true
                @test p.italic === false
                @test p.under == "none"
                @test t.paragraphs[1].props.align == "ctr"

                @test XLSX.Charts.is_uniform(t)                        # no runs: uniform
            end

            @testset "uniform multi-run rich text" begin
                node = _frag("""<c:rich xmlns:c="$(XLSX.Charts.NS_C)" $_NSDECL>
                      <a:bodyPr/><a:lstStyle/>
                      <a:p>
                        <a:r><a:rPr lang="en-GB" sz="1400" b="1"/><a:t>Quarterly </a:t></a:r>
                        <a:r><a:rPr lang="en-US" sz="1400" b="1"/><a:t>Revenue</a:t></a:r>
                      </a:p>
                    </c:rich>""")

                t = XLSX.Charts.parse_drawing_text(wb, node; tag="rich")
                @test length(XLSX.Charts.text_runs(t)) == 2
                @test XLSX.Charts.text_content(t) == "Quarterly Revenue"

                # differing only by lang — Excel splits runs on spellcheck
                # boundaries, and that must not read as mixed formatting
                @test XLSX.Charts.is_uniform(t)
                @test XLSX.Charts.default_run_props(t) !== nothing
                @test XLSX.Charts.default_run_props(t).size == 14.0
            end

            @testset "mixed formatting" begin
                node = _frag("""<c:rich xmlns:c="$(XLSX.Charts.NS_C)" $_NSDECL>
                      <a:bodyPr/><a:lstStyle/>
                      <a:p>
                        <a:r><a:rPr sz="1400" b="0"/><a:t>Revenue </a:t></a:r>
                        <a:r><a:rPr sz="1400" b="1"/><a:t>2024</a:t></a:r>
                      </a:p>
                    </c:rich>""")

                t = XLSX.Charts.parse_drawing_text(wb, node; tag="rich")
                @test XLSX.Charts.text_content(t) == "Revenue 2024"
                @test !XLSX.Charts.is_uniform(t)

                # no single answer, so no answer
                @test XLSX.Charts.default_run_props(t) === nothing
                # ...but the first fragment is still reachable
                @test XLSX.Charts.first_run_props(t).bold === false
                @test XLSX.Charts.text_runs(t)[2].props.bold === true
            end

            @testset "breaks and paragraphs" begin
                node = _frag("""<c:rich xmlns:c="$(XLSX.Charts.NS_C)" $_NSDECL>
                      <a:bodyPr/><a:lstStyle/>
                      <a:p><a:r><a:t>One</a:t></a:r><a:br/><a:r><a:t>Two</a:t></a:r></a:p>
                      <a:p><a:r><a:t>Three</a:t></a:r></a:p>
                    </c:rich>""")

                t = XLSX.Charts.parse_drawing_text(wb, node; tag="rich")
                @test length(t.paragraphs) == 2
                @test XLSX.Charts.text_content(t) == "One\nTwo\nThree"
                @test [r.kind for r in t.paragraphs[1].runs] == [:run, :br, :run]
                @test XLSX.Charts.DrawingText("Revenue\nby region") ==
                    XLSX.Charts.DrawingText("Revenue", "by region")
                @test length(XLSX.Charts.DrawingText("a\nb\nc").paragraphs) == 3
                @test XLSX.Charts.text_content(XLSX.Charts.DrawingText("a\nb")) == "a\nb"
            end

            @testset "DrawingText show" begin
                t = XLSX.Charts.DrawingText("Revenue\nby region")
                @test sprint(show, t) == "DrawingText(\"Revenue\\nby region\", 2 paragraph(s), 2 run(s))"
                @test occursin("formatting only", sprint(show, XLSX.Charts.DrawingText("")))
            end

            @testset "a newline in a DrawingParagraph is a line break" begin
                P = XLSX.Charts.DrawingParagraph
                R = XLSX.Charts.DrawingRun
                kinds(p) = [r.kind for r in p.runs]
                p = P("a\nb")
                @test kinds(p) == [:run, :br, :run]
                @test [r.text for r in p.runs] == ["a", "\n", "b"]
                # the same paragraph the explicit form builds, and the reader produces
                @test p == P("a", R("\n"; kind=:br), "b")
                @test kinds(P("a\n\nb")) == [:run, :br, :br, :run]
                @test kinds(P("a\n")) == [:run, :br]
                @test kinds(P("a", R("b"))) == [:run, :run]               # no newline: unchanged
                @test length(P("").runs) == 1
                # Windows and lone-CR line endings
                @test kinds(P("a\r\nb")) == [:run, :br, :run]
                @test [r.text for r in P("a\r\nb").runs] == ["a", "\n", "b"]   # no stray \r
                @test kinds(P("a\rb")) == [:run, :br, :run]
                # DrawingText still splits into paragraphs first, on every line ending
                @test length(XLSX.Charts.DrawingText("a\r\nb").paragraphs) == 2
                @test only(first(XLSX.Charts.DrawingText("a\r\nb").paragraphs).runs).text == "a"
            end

            @testset "indented XML (nodetype guard)" begin
                # Excel writes chart parts unindented; a formatted or
                # hand-edited part has whitespace text nodes between runs.
                node = _frag("""<c:rich xmlns:c="$(XLSX.Charts.NS_C)" $_NSDECL>
                      <a:bodyPr/>
                      <a:lstStyle/>
                      <a:p>
                        <a:r><a:rPr sz="1000"/><a:t>Spaced</a:t></a:r>
                        <a:r><a:rPr sz="1000"/><a:t> out</a:t></a:r>
                      </a:p>
                    </c:rich>""")

                t = XLSX.Charts.parse_drawing_text(wb, node; tag="rich")
                @test length(XLSX.Charts.text_runs(t)) == 2          # not 2 + whitespace nodes
                @test XLSX.Charts.text_content(t) == "Spaced out"
            end

            @testset "parse from parent, and absent txPr" begin
                parent = _frag("""<c:valAx xmlns:c="$(XLSX.Charts.NS_C)" $_NSDECL>
                      <c:delete val="0"/>
                      <c:txPr><a:bodyPr/><a:lstStyle/>
                        <a:p><a:pPr><a:defRPr sz="1000"/></a:pPr></a:p>
                      </c:txPr>
                    </c:valAx>""")
                t = XLSX.Charts.parse_drawing_text(wb, parent)          # searches for txPr
                @test t !== nothing
                @test XLSX.Charts.default_run_props(t).size == 10.0

                bare = _frag("""<c:valAx xmlns:c="$(XLSX.Charts.NS_C)"><c:delete val="0"/></c:valAx>""")
                @test XLSX.Charts.parse_drawing_text(wb, bare) === nothing
            end

            @testset "solidFill inside defRPr" begin
                node = _frag("""<c:txPr xmlns:c="$(XLSX.Charts.NS_C)" $_NSDECL>
                      <a:bodyPr/><a:lstStyle/>
                      <a:p><a:pPr><a:defRPr sz="900">
                        <a:solidFill><a:schemeClr val="tx1">
                          <a:lumMod val="65000"/><a:lumOff val="35000"/>
                        </a:schemeClr></a:solidFill>
                        <a:latin typeface="+mn-lt"/>
                      </a:defRPr></a:pPr></a:p>
                    </c:txPr>""")

                t = XLSX.Charts.parse_drawing_text(wb, node)
                p = XLSX.Charts.default_run_props(t)
                @test p.fill !== nothing
                @test p.fill.kind == :solid
                # same transform the colour tests already pin
                @test p.fill.fgcolor.rgb == "595959"
                @test p.latin == "+mn-lt"           # kept as written
                @test p.ea === nothing
                @test p.line === nothing
            end
        end
    end

    # -----------------------------------------------------------------------
    # -----------------------------------------------------------------------
    @testset "fixtures" begin

        # NOTE: `Chart` is metadata only — it does not retain the parsed node,
        # so this reaches into the package by path. That is a stopgap: the
        # surgical-write design needs charts to hold their XML, and when stage
        # 3 decides how (a `raw` field, or a lookup keyed on `path`), these
        # three lines become one accessor. Don't copy this pattern elsewhere.
        _chart_root(xf, c) = XLSX.xml_root_element(XLSX.get_xml_data(xf, c.path))

        XLSX.openxlsx(joinpath(data_directory, "chart_basic.xlsx")) do xf
            wb = XLSX.get_workbook(xf)
            c = first(XLSX.Charts.getCharts(xf["Data"]))

            root = _chart_root(xf, c)                          # c:chartSpace
            chart = XLSX.first_element_with_tag(root, "chart")
            @test chart !== nothing

            @testset "axis txPr" begin
                plotarea = XLSX.first_element_with_tag(chart, "plotArea")
                @test plotarea !== nothing

                ax = XLSX.first_element_with_tag(plotarea, "valAx")
                @test ax !== nothing

                t = XLSX.Charts.parse_drawing_text(wb, ax)
                @test t !== nothing
                @test XLSX.Charts.text_content(t) == ""          # formatting only
                @test isempty(XLSX.Charts.text_runs(t))
                @test XLSX.Charts.is_uniform(t)
                @test t.raw !== nothing              # kept for write-back

                # <a:bodyPr rot="-60000000" spcFirstLastPara="1"
                #  vertOverflow="ellipsis" vert="horz" wrap="square"
                #  anchor="ctr" anchorCtr="1"/>
                @test t.body.rotation == -1000.0
                @test t.body.vertical == "horz"
                @test t.body.wrap == "square"
                @test t.body.anchor == "ctr"
                @test t.body.anchorctr === true
                @test t.body.spcfirstlastpara === true
                @test t.body.vertoverflow == "ellipsis"
                @test t.body.horzoverflow === nothing
                @test t.body.autofit === nothing     # absent, not :none
                @test t.body.insetleft === nothing

                # The font is on a:pPr/a:defRPr — this txPr has no runs.
                @test t.paragraphs[1].props.defprops !== nothing
                p = XLSX.Charts.default_run_props(t)
                @test p === t.paragraphs[1].props.defprops

                # <a:defRPr sz="900" b="0" i="0" u="none" strike="noStrike"
                #  kern="1200" baseline="0"/>
                @test p.size == 9.0
                @test p.kern == 12.0
                @test p.under == "none"
                @test p.strike == "noStrike"

                # Written as 0, so explicitly false — not absent.
                @test p.bold === false
                @test p.italic === false
                @test p.baseline == 0.0

                # Not written at all.
                @test p.caps === nothing
                @test p.spacing === nothing
                @test p.lang === nothing
                @test p.line === nothing

                # tx1 + lumMod 65000 / lumOff 35000
                @test p.fill !== nothing
                @test p.fill.kind == :solid
                @test p.fill.fgcolor.rgb == "595959"

                # Theme references, kept as written.
                @test p.latin == "+mn-lt"
                @test p.ea == "+mn-ea"
                @test p.cs == "+mn-cs"
                @test XLSX.resolve_theme_font(wb, p.latin) !== nothing
            end
            @testset "title rich text" begin
                title = XLSX.first_element_with_tag(chart, "title")
                if title !== nothing
                    tx = XLSX.first_element_with_tag(title, "tx")
                    if tx !== nothing
                        t = XLSX.Charts.parse_drawing_text(wb, tx; tag="rich")
                        if t !== nothing
                            # `getChartTitle` is parsed independently, so this
                            # cross-checks the two paths agree.
                            @test XLSX.Charts.text_content(t) == XLSX.Charts.getChartTitle(c)
                            @test !isempty(XLSX.Charts.text_runs(t))
                        end
                    end
                end
            end
        end
        @testset "text_content reads a raw body as it reads a parsed one" begin
            xml = """<c:rich xmlns:c="$(XLSX.Charts.NS_C)" xmlns:a="$(XLSX.NS_A)"><a:bodyPr/><a:lstStyle/>""" *
                """<a:p><a:r><a:t>One</a:t></a:r></a:p>""" *
                """<a:p><a:r><a:t>Two</a:t></a:r><a:fld id="{0}" type="x"><a:t>!</a:t></a:fld></a:p></c:rich>"""
            node = XLSX.xml_root_element(parse(xml, XLSX.XML.Node))
            @test XLSX.Charts.text_content(node) == "One\nTwo!"
        end
        @testset "theme fonts" begin
            XLSX.openxlsx(joinpath(data_directory, "chart_theme_colors.xlsx")) do xf
                wb = XLSX.get_workbook(xf)

                fonts = XLSX.get_theme_fonts(wb)
                @test fonts isa Dict{String,String}
                @test haskey(fonts, "+mn-lt")
                @test haskey(fonts, "+mj-lt")

                @test XLSX.resolve_theme_font(wb, "+mn-lt") == fonts["+mn-lt"]
                @test XLSX.resolve_theme_font(wb, "Calibri") == "Calibri"
                @test XLSX.resolve_theme_font(wb, nothing) === nothing
                @test XLSX.resolve_theme_font(wb, "+xx-lt") === nothing   # undefined ref
            end
        end
    end

    @testset "run properties overlay defRPr" begin
        f = XLSX.readxlsx(joinpath(data_directory, "chartex_formatted.xlsx"))
        c = only(filter(x -> x isa XLSX.Charts.ChartEx, XLSX.Charts.getCharts(f)))
        wb = XLSX.Charts._wb(c)

        # Empty defRPr, formatting on the run: the run answers.
        title = XLSX.Charts.parse_drawing_text(wb, first(XLSX.elements_with_tag(XLSX.Charts._cx_chart(c), "title")))
        rp = XLSX.Charts.default_run_props(title)
        @test rp.size ≈ 18.0
        @test rp.fill.fgcolor.val == "8C8026"
        @test rp.latin == "Comic Sans MS"

        # No run: defRPr answers as written; endParaRPr does not contribute.
        labels = XLSX.Charts.parse_drawing_text(wb, XLSX.Charts._cx_datalabels(c, 1))
        lp = XLSX.Charts.default_run_props(labels)
        @test lp.size ≈ 11.0
        @test lp.bold === true
        @test lp.fill.fgcolor.val == "7030A0"
        @test isnothing(lp.latin)              # font is only on endParaRPr

        # defRPr and run disagree: the run wins field by field, defRPr fills the rest.
        xml = """<cx:title xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex"
                           xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                   <cx:txPr><a:bodyPr/><a:lstStyle/><a:p>
                     <a:pPr><a:defRPr sz="1400" b="1"/></a:pPr>
                     <a:r><a:rPr sz="1800"/><a:t>x</a:t></a:r>
                   </a:p></cx:txPr></cx:title>"""
        node = XLSX.xml_root_element(parse(xml, XLSX.XML.Node))
        mp = XLSX.Charts.default_run_props(XLSX.Charts.parse_drawing_text(wb, node))
        @test mp.size ≈ 18.0                   # from the run
        @test mp.bold === true                 # from defRPr
    end
    @testset "setting a run property clears it from the runs" begin
        tmp = "runclear.xlsx"
        cp(joinpath(data_directory, "chartex_layouts.xlsx"), tmp; force=true)
        xf = XLSX.openxlsx(tmp; mode="rw")
        h = only(filter(x -> x isa XLSX.Charts.ChartEx && x.sheet == "histogram", XLSX.Charts.getCharts(xf)))

        # Excel wrote this title with sz on both a:defRPr and the run's a:rPr
        XLSX.Charts.setChartTitleTextProp(h, :size, 20)
        tx = XLSX.first_element_with_tag(
            XLSX.first_element_with_tag(XLSX.Charts._cx_chart(h), "title"), "txPr")
        p = XLSX.first_element_with_tag(tx, "p")
        def = XLSX.first_element_with_tag(XLSX.first_element_with_tag(p, "pPr"), "defRPr")
        rpr = XLSX.first_element_with_tag(XLSX.first_element_with_tag(p, "r"), "rPr")

        @test XLSX.get_attr(def, "sz") == "2000"
        @test isempty(XLSX.get_attr(rpr, "sz"))          # cleared, so the default renders
        @test XLSX.get_attr(rpr, "lang") == "en-GB"      # the run keeps everything else

        # what the reader resolves now matches what Excel will render
        @test XLSX.Charts.default_run_props(XLSX.Charts.getChartTitleTextProps(h)).size ≈ 20.0

        # a composite field: the run's own solidFill must go too
        XLSX.Charts.setChartTitleTextProp(h, :fill, "FFC00000")
        rpr = XLSX.first_element_with_tag(XLSX.first_element_with_tag(
            XLSX.first_element_with_tag(
                XLSX.first_element_with_tag(XLSX.Charts._cx_chart(h), "title"), "txPr"), "p"), "r")
        @test isnothing(XLSX.first_element_with_tag(
            XLSX.first_element_with_tag(rpr, "rPr"), "solidFill"))
        @test XLSX.Charts.default_run_props(XLSX.Charts.getChartTitleTextProps(h)).fill.fgcolor.rgb == "C00000"
        SAVE_FILES && save_outfile(xf)
        isfile(tmp) && rm(tmp)
    end
    @testset "formatting cascade" begin
        f = XLSX.readxlsx(joinpath(data_directory, "chartex_formatted.xlsx"))
        c = only(filter(x -> x isa XLSX.Charts.ChartEx, XLSX.Charts.getCharts(f)))

        # Series fill: theme colour with a transform.
        e = XLSX.Charts.getSeriesFill(c, 1)
        @test e.site.level == :series
        @test e.value.fgcolor.val == "accent2"

        # Point 5 (idx 4) has its own fill: explicit colour with transforms.
        e5 = XLSX.Charts.getSeriesFill(c, 1; point=5)
        @test e5.site.level == :point
        @test [s.level for s in e5.chain] == [:point, :series]
        @test e5.value.fgcolor.kind == :srgb
        @test e5.value.fgcolor.val == "4EA72E"                                  # as written
        @test e5.value.fgcolor.transforms == [:lumMod => 60000, :lumOff => 40000]
        @test e5.value.fgcolor.rgb == "8ED973"                                  # transforms applied

        # Point 1 has no dataPt: no point rung, resolves at the series.
        e1 = XLSX.Charts.getSeriesFill(c, 1; point=1)
        @test e1.site.level == :series
        @test [s.level for s in e1.chain] == [:series]
        @test_throws XLSX.XLSXError XLSX.Charts.getSeriesFill(c, 1; point=0)

        # Label text: set on dataLabels/txPr defRPr.
        s = XLSX.Charts.getLabelTextProp(c, 1, :size)
        @test s.value ≈ 11.0
        @test s.site.level == :series
        @test [x.level for x in s.chain] == [:series, :chartspace]
        @test XLSX.Charts.getLabelTextProp(c, 1, :bold).value === true
        @test XLSX.Charts.getLabelTextProp(c, 1, :fill).value.fgcolor.val == "7030A0"
        @test_throws XLSX.XLSXError XLSX.Charts.getLabelTextProp(c, 1, :nonsense)

        # Nothing written anywhere: no value, no site, but a full chain.
        g = XLSX.readxlsx(joinpath(data_directory, "chartex_layouts.xlsx"))
        w = only(filter(x -> x isa XLSX.Charts.ChartEx && x.sheet == "waterfall", XLSX.Charts.getCharts(g)))
        u = XLSX.Charts.getSeriesFill(w, 1)
        @test isnothing(u.value) && isnothing(u.site)
        @test length(u.chain) == 1
        @test isnothing(only(u.chain).props)
        @test isnothing(XLSX.Charts.getLabelTextProp(w, 1, :size).value)
    end
    @testset "DrawingML shape properties" begin

        XLSX.openxlsx(joinpath(data_directory, "chart_basic.xlsx")) do xf
            wb = XLSX.get_workbook(xf)

            @testset "solid fill and line" begin
                node = _frag("""<c:spPr xmlns:c="$(XLSX.Charts.NS_C)" $_NSDECL>
                      <a:solidFill><a:srgbClr val="4472C4"/></a:solidFill>
                      <a:ln w="19050"><a:solidFill><a:srgbClr val="203864"/></a:solidFill>
                        <a:prstDash val="dash"/></a:ln>
                    </c:spPr>""")

                sp = XLSX.Charts.parse_drawing_shape_props(wb, node)
                @test sp isa XLSX.Charts.DrawingShapeProps
                @test sp.fill.kind == :solid
                @test sp.fill.fgcolor.rgb == "4472C4"
                @test sp.line.width == 1.5              # points, converted at parse
                @test sp.line.dash == "dash"
                @test sp.line.fill.fgcolor.rgb == "203864"
                @test XLSX.Charts.hasFill(sp)
                @test XLSX.Charts.hasLine(sp)
                @test sp.effects === nothing
                @test sp.bwmode === nothing
                @test sp.raw !== nothing
            end

            @testset "noFill is not absence" begin
                # Chart-area shape: transparent background, no border. Both are
                # deliberate, and neither is the same as omitting the element.
                node = _frag("""<c:spPr xmlns:c="$(XLSX.Charts.NS_C)" $_NSDECL>
                      <a:noFill/>
                      <a:ln><a:noFill/></a:ln>
                    </c:spPr>""")

                sp = XLSX.Charts.parse_drawing_shape_props(wb, node)
                @test sp.fill !== nothing               # present...
                @test sp.fill.kind == :none             # ...and explicitly invisible
                @test !XLSX.Charts.hasFill(sp)

                @test sp.line !== nothing
                @test sp.line.fill.kind == :none
                @test !XLSX.Charts.hasLine(sp)
            end

            @testset "absence is inheritance" begin
                node = _frag("""<c:spPr xmlns:c="$(XLSX.Charts.NS_C)" $_NSDECL/>""")

                sp = XLSX.Charts.parse_drawing_shape_props(wb, node)
                @test sp !== nothing                    # the element exists
                @test sp.fill === nothing               # but sets nothing
                @test sp.line === nothing
                @test !XLSX.Charts.hasFill(sp)
                @test !XLSX.Charts.hasLine(sp)
            end

            @testset "parse from parent, and no spPr at all" begin
                ser = _frag("""<c:ser xmlns:c="$(XLSX.Charts.NS_C)" $_NSDECL>
                      <c:idx val="0"/>
                      <c:spPr><a:solidFill><a:schemeClr val="accent1"/></a:solidFill></c:spPr>
                    </c:ser>""")
                sp = XLSX.Charts.parse_drawing_shape_props(wb, ser)
                @test sp !== nothing
                @test sp.fill.fgcolor.rgb == "156082"   # accent1, as pinned elsewhere

                bare = _frag("""<c:ser xmlns:c="$(XLSX.Charts.NS_C)"><c:idx val="0"/></c:ser>""")
                @test XLSX.Charts.parse_drawing_shape_props(wb, bare) === nothing

                # chains without a guard
                @test XLSX.Charts.parse_drawing_shape_props(wb, nothing) === nothing
            end

            @testset "effects preserved, not modelled" begin
                node = _frag("""<c:spPr xmlns:c="$(XLSX.Charts.NS_C)" $_NSDECL>
                      <a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill>
                      <a:effectLst><a:outerShdw blurRad="50800" dist="38100"/></a:effectLst>
                    </c:spPr>""")

                sp = XLSX.Charts.parse_drawing_shape_props(wb, node)
                @test sp.effects !== nothing
                @test XLSX.localname(sp.effects) == "effectLst"
            end

            @testset "pattern fill keeps DrawingML vocabulary" begin
                node = _frag("""<c:spPr xmlns:c="$(XLSX.Charts.NS_C)" $_NSDECL>
                      <a:pattFill prst="ltUpDiag">
                        <a:fgClr><a:srgbClr val="000000"/></a:fgClr>
                        <a:bgClr><a:srgbClr val="FFFFFF"/></a:bgClr>
                      </a:pattFill>
                    </c:spPr>""")

                sp = XLSX.Charts.parse_drawing_shape_props(wb, node)
                @test sp.fill.kind == :pattern
                @test sp.fill.preset == "ltUpDiag"      # not "lightUp"
                @test XLSX.Charts.hasFill(sp)
            end
        end

        @testset "fixture: series spPr" begin
            XLSX.openxlsx(joinpath(data_directory, "chart_basic.xlsx")) do xf
                wb = XLSX.get_workbook(xf)
                c = first(XLSX.Charts.getCharts(xf["Data"]))

                root = XLSX.xml_root_element(XLSX.get_xml_data(xf, c.path))
                chart = XLSX.first_element_with_tag(root, "chart")
                plotarea = XLSX.first_element_with_tag(chart, "plotArea")
                barchart = XLSX.first_element_with_tag(plotarea, "barChart")
                @test barchart !== nothing

                ser = XLSX.first_element_with_tag(barchart, "ser")
                @test ser !== nothing

                sp = XLSX.Charts.parse_drawing_shape_props(wb, ser)
                @test sp !== nothing
                @test sp.raw !== nothing

                @test XLSX.Charts.hasFill(sp)
                @test sp.fill.kind == :solid
                @test sp.fill.fgcolor.rgb == "156082"     # accent1

                # Excel writes an explicit "no border", which is not the same as
                # omitting a:ln. hasLine must be false while sp.line is not nothing.
                @test sp.line !== nothing
                @test sp.line.fill.kind == :none
                @test !XLSX.Charts.hasLine(sp)

                @test sp.effects !== nothing
                @test XLSX.localname(sp.effects) == "effectLst"
                if sp !== nothing
                    @test sp.raw !== nothing
                end
            end
        end
    end

    @testset "SchemeColor" begin
        @testset "construction" begin
            @test XLSX.Charts.SchemeColor(:accent1).token === :accent1
            @test isempty(XLSX.Charts.SchemeColor(:accent1).transforms)
            @test XLSX.Charts.SchemeColor(:accent1; lumMod=75).transforms == [:lumMod => 75.0]
            @test XLSX.Charts.SchemeColor(:tx1; lumMod=65, lumOff=35).transforms ==
                [:lumMod => 65.0, :lumOff => 35.0]
            # integers convert
            @test XLSX.Charts.SchemeColor(:accent1; alpha=80).transforms == [:alpha => 80.0]
            # vector form takes any order
            @test XLSX.Charts.SchemeColor(:accent2, [:shade => 50.0, :alpha => 80.0]).transforms ==
                [:shade => 50.0, :alpha => 80.0]
        end

        @testset "validation" begin
            @test_throws XLSX.XLSXError XLSX.Charts.SchemeColor(:accent7)
            @test_throws XLSX.XLSXError XLSX.Charts.SchemeColor(:acccent1)          # typo, not silently accepted
            @test_throws XLSX.XLSXError XLSX.Charts.SchemeColor(:accent1, [:lumMud => 75.0])
            @test_throws XLSX.XLSXError XLSX.Charts.SchemeColor(:theme1)            # spreadsheet vocabulary, not DrawingML
        end

        @testset "aliases" begin
            # preserved as written
            @test XLSX.Charts.SchemeColor(:lt1).token === :lt1
            @test XLSX.Charts.SchemeColor(:dk2).token === :dk2
            # but equal and equally hashed
            @test XLSX.Charts.SchemeColor(:lt1) == XLSX.Charts.SchemeColor(:bg1)
            @test XLSX.Charts.SchemeColor(:dk1) == XLSX.Charts.SchemeColor(:tx1)
            @test XLSX.Charts.SchemeColor(:lt2) == XLSX.Charts.SchemeColor(:bg2)
            @test XLSX.Charts.SchemeColor(:dk2) == XLSX.Charts.SchemeColor(:tx2)
            @test hash(XLSX.Charts.SchemeColor(:lt1)) == hash(XLSX.Charts.SchemeColor(:bg1))
            @test isequal(XLSX.Charts.SchemeColor(:lt1), XLSX.Charts.SchemeColor(:bg1))
            # so one Dict slot, not two
            d = Dict(XLSX.Charts.SchemeColor(:lt1) => 1)
            d[XLSX.Charts.SchemeColor(:bg1)] = 2
            @test length(d) == 1
            # aliases with matching transforms are still equal
            @test XLSX.Charts.SchemeColor(:lt1; lumMod=50) == XLSX.Charts.SchemeColor(:bg1; lumMod=50)
            # different slots are not
            @test XLSX.Charts.SchemeColor(:accent1) != XLSX.Charts.SchemeColor(:accent2)
        end

        @testset "transform order is significant" begin
            a = XLSX.Charts.SchemeColor(:accent1; lumMod=75, lumOff=25)
            b = XLSX.Charts.SchemeColor(:accent1, [:lumOff => 25.0, :lumMod => 75.0])
            @test a.transforms != b.transforms
            @test a != b
            @test XLSX.Charts.SchemeColor(:accent1; lumMod=75) != XLSX.Charts.SchemeColor(:accent1; lumMod=50)
            @test XLSX.Charts.SchemeColor(:accent1) != XLSX.Charts.SchemeColor(:accent1; lumMod=75)
        end

        xf = XLSX.openxlsx(joinpath(data_directory, "chart_appearance.xlsx"))
        wb = XLSX.get_workbook(xf[1])
        @testset "SchemeColor round trip" begin

            # resolve_color_base does the theme lookup only; the transforms are applied
            # separately, so composing the two is what a reader effectively does.
            resolved(wb, n) = first(XLSX.apply_drawingml_transforms(
                XLSX.Charts.resolve_color_base(wb, n),
                XLSX.Charts.parse_drawing_color(wb, n).transforms))

            # accent1, untransformed
            n = XLSX.Charts._scheme_color_node(XLSX.Charts.SchemeColor(:accent1))
            @test XML.write(n) == """<a:schemeClr val="accent1"/>"""
            dc = XLSX.Charts.parse_drawing_color(wb, n)
            @test dc.kind === :scheme && dc.val == "accent1"

            # accent1 + lumMod 75% — 104862 against the Office theme, verified in stage 2
            n = XLSX.Charts._scheme_color_node(XLSX.Charts.SchemeColor(:accent1; lumMod=75))
            @test occursin("""val="75000\"""", XML.write(n))
            dc = XLSX.Charts.parse_drawing_color(wb, n)
            @test length(dc.transforms) == 1
            @test XLSX.Charts.resolve_color_base(wb, n) == "156082"

            # tx1 + lumMod 65 / lumOff 35 — 595959, and order is preserved through the XML
            n = XLSX.Charts._scheme_color_node(XLSX.Charts.SchemeColor(:tx1; lumMod=65, lumOff=35))
            @test XLSX.localname.(collect(XML.eachelement(n))) == ["lumMod", "lumOff"]
            @test XLSX.Charts.resolve_color_base(wb, n) == "000000"

            # what we build parses back to what we built
            for sc in (XLSX.Charts.SchemeColor(:accent1), XLSX.Charts.SchemeColor(:accent1; lumMod=75),
                XLSX.Charts.SchemeColor(:tx1; lumMod=65, lumOff=35),
                XLSX.Charts.SchemeColor(:accent2, [:shade => 50.0, :alpha => 80.0]))
                dc = XLSX.Charts.parse_drawing_color(wb, XLSX.Charts._scheme_color_node(sc))
                @test dc.val == String(sc.token)
                @test length(dc.transforms) == length(sc.transforms)
            end

            @test XLSX.Charts.resolve_color_base(wb, XLSX.Charts._scheme_color_node(XLSX.Charts.SchemeColor(:accent1))) == "156082"
            @test resolved(wb, XLSX.Charts._scheme_color_node(XLSX.Charts.SchemeColor(:accent1; lumMod=75))) == "104862"
            @test resolved(wb, XLSX.Charts._scheme_color_node(XLSX.Charts.SchemeColor(:tx1; lumMod=65, lumOff=35))) == "595959"

            # solidFill wrapper
            f = XLSX.Charts.parse_drawing_fill(wb, XLSX.Charts._solid_fill_node(XLSX.Charts._scheme_color_node(XLSX.Charts.SchemeColor(:accent1))))
            @test f.kind === :solid && f.fgcolor.val == "accent1"

            # alpha rides alongside the hex rather than altering it
            n = XLSX.Charts._scheme_color_node(XLSX.Charts.SchemeColor(:accent1; alpha=50))
            hex, alpha = XLSX.apply_drawingml_transforms(
                XLSX.Charts.resolve_color_base(wb, n), XLSX.Charts.parse_drawing_color(wb, n).transforms)
            @test hex == "156082"
            @test alpha ≈ 0.5
        end

        @testset "srgb color nodes" begin
            # 8-digit hex passes through, split into six digits plus alpha
            @test XML.write(XLSX.Charts._srgb_color_node("FFFF0000")) == """<a:srgbClr val="FF0000"/>"""
            @test XML.write(XLSX.Charts._srgb_color_node("ffff0000")) == """<a:srgbClr val="FF0000"/>"""   # case

            # Colors.jl names
            @test XML.write(XLSX.Charts._srgb_color_node("red")) == """<a:srgbClr val="FF0000"/>"""
            @test XML.write(XLSX.Charts._srgb_color_node(:red)) == """<a:srgbClr val="FF0000"/>"""
            @test XML.write(XLSX.Charts._srgb_color_node("grey")) == XML.write(XLSX.Charts._srgb_color_node("gray"))

            # opaque emits no alpha child
            @test isnothing(XLSX.Charts._srgb_color_node("FF0000FF").children) ||
                isempty(XLSX.Charts._srgb_color_node("FF0000FF").children)

            # partial alpha becomes a transform
            n = XLSX.Charts._srgb_color_node("800000FF")
            @test XLSX.localname.(collect(XML.eachelement(n))) == ["alpha"]
            @test XLSX.get_attr(first(XML.eachelement(n)), "val") == "50196"   # 128/255

            # colorants, opaque and transparent
            @test XML.write(XLSX.Charts._srgb_color_node(Colors.RGB(1, 0, 0))) == """<a:srgbClr val="FF0000"/>"""
            n = XLSX.Charts._srgb_color_node(Colors.ARGB(1, 0, 0, 0.5))
            @test XLSX.get_attr(n, "val") == "FF0000"
            @test !isempty(XML.children(n))

            # invalid names throw get_color's message
            @test_throws XLSX.XLSXError XLSX.Charts._srgb_color_node("notacolor")

            # round trip through the parser
            dc = XLSX.Charts.parse_drawing_color(wb, XLSX.Charts._srgb_color_node("red"))
            @test dc.kind === :srgb && dc.rgb == "FF0000"
        end

        @testset "_ln_with_*" begin
            pfx = Dict(XLSX.NS_A => "a")
            ln() = XML.Element("a:ln")

            @test XLSX.get_attr(XLSX.Charts._ln_with_width(ln(), 2), "w") == "25400"
            @test_throws XLSX.XLSXError XLSX.Charts._ln_with_width(ln(), 2000)
            @test_throws XLSX.XLSXError XLSX.Charts._ln_with_width(ln(), -1)
            @test XLSX.get_attr(XLSX.Charts._ln_with_width(XLSX.Charts._ln_with_width(ln(), 2), :inherit), "w") == ""

            # Excel names and DrawingML names both work, and resolve to the same thing
            @test XML.write(XLSX.Charts._ln_with_dash(ln(), :roundDot, pfx)) ==
                XML.write(XLSX.Charts._ln_with_dash(ln(), :sysDot, pfx))
            @test_throws XLSX.XLSXError XLSX.Charts._ln_with_dash(ln(), :dotted, pfx)

            @test XLSX.get_attr(XLSX.Charts._ln_with_cap(ln(), :round), "cap") == "rnd"
            @test XLSX.get_attr(XLSX.Charts._ln_with_cap(ln(), :rnd), "cap") == "rnd"
            @test XLSX.get_attr(XLSX.Charts._ln_with_compound(ln(), :double), "cmpd") == "dbl"

            # join, and the limit that depends on it
            l = XLSX.Charts._ln_with_join(ln(), :miter, pfx)
            @test XLSX.localname.(collect(XML.eachelement(l))) == ["miter"]
            @test XLSX.get_attr(first(XML.eachelement(XLSX.Charts._ln_with_miter_limit(l, 8))), "lim") == "800000"
            @test_throws XLSX.XLSXError XLSX.Charts._ln_with_miter_limit(ln(), 8)
            @test_throws XLSX.XLSXError XLSX.Charts._ln_with_miter_limit(XLSX.Charts._ln_with_join(ln(), :bevel, pfx), 8)

            # schema order: fill, dash, join
            full = XLSX.Charts._ln_with_join(XLSX.Charts._ln_with_dash(XLSX.Charts._ln_with_color(ln(), "red", pfx), :dash, pfx), :round, pfx)
            @test XLSX.localname.(collect(XML.eachelement(full))) == ["solidFill", "prstDash", "round"]

            # each choice group replaces rather than accumulates
            @test XLSX.localname.(collect(XML.eachelement(
                XLSX.Charts._ln_with_color(XLSX.Charts._ln_with_color(ln(), "red", pfx), :none, pfx)))) == ["noFill"]
        end

        @testset "text body round trip" begin
            pfx = Dict(XLSX.NS_A => "a", XLSX.Charts.NS_C => "c")

            rt(t) = XLSX.Charts.parse_drawing_text(wb, XLSX.Charts._text_from(t, "txPr", pfx))

            # plain text
            t = XLSX.Charts.DrawingText("Revenue by Region")
            g = rt(t)
            @test length(g.paragraphs) == 1
            @test XLSX.Charts.text_content(g) == "Revenue by Region"

            # run properties survive, in both directions
            t = XLSX.Charts.DrawingText(XLSX.Charts.DrawingParagraph(
                XLSX.Charts.DrawingRun("Revenue", props=XLSX.Charts.DrawingRunProps(size=14.0, bold=true,
                    latin="Calibri"))))
            g = rt(t)
            rp = XLSX.Charts.first_run_props(g)
            @test rp.size ≈ 14.0 && rp.bold === true && rp.latin == "Calibri"
            @test isnothing(rp.italic)          # absent stays absent

            # paragraph defaults
            t = XLSX.Charts.DrawingText(XLSX.Charts.DrawingParagraph("x",
                props=XLSX.Charts.DrawingParaProps(align="ctr",
                    defprops=XLSX.Charts.DrawingRunProps(size=10.5))))
            g = rt(t)
            @test g.paragraphs[1].props.align == "ctr"
            @test g.paragraphs[1].props.defprops.size ≈ 10.5

            # a solid fill on a run, scheme and srgb
            for color in (XLSX.Charts.DrawingColor(:srgb, "FF0000", Pair{Symbol,Int}[], "FF0000", 1.0),
                XLSX.Charts.DrawingColor(:scheme, "accent1", [:lumMod => 75000], "104862", 1.0))
                t = XLSX.Charts.DrawingText(XLSX.Charts.DrawingParagraph(XLSX.Charts.DrawingRun("x",
                    props=XLSX.Charts.DrawingRunProps(fill=XLSX.Charts.DrawingFill(:solid; fgcolor=color)))))
                f = XLSX.Charts.first_run_props(rt(t)).fill
                @test f.kind === :solid
                @test f.fgcolor.val == color.val
                @test f.fgcolor.transforms == color.transforms
            end

            # a text outline
            t = XLSX.Charts.DrawingText(XLSX.Charts.DrawingParagraph(XLSX.Charts.DrawingRun("x",
                props=XLSX.Charts.DrawingRunProps(line=XLSX.Charts.DrawingLine(width=1.5, dash="sysDot")))))
            l = XLSX.Charts.first_run_props(rt(t)).line
            @test l.width ≈ 1.5 && l.dash == "sysDot"

            # body properties, including the three-way autofit
            for (af, extra) in ((:none, ()), (:shape, ()),
                (:normal, (fontscale=0.9, linespacereduction=0.1)))
                t = XLSX.Charts.DrawingText("x"; body=XLSX.Charts.DrawingBodyProps(; rotation=-45.0, anchor="ctr",
                    autofit=af, extra...))
                b = rt(t).body
                @test b.rotation ≈ -45.0 && b.anchor == "ctr" && b.autofit === af
            end
            b = rt(XLSX.Charts.DrawingText("x"; body=XLSX.Charts.DrawingBodyProps(autofit=:normal, fontscale=0.9))).body
            @test b.fontscale ≈ 0.9

            # several runs keep document order
            t = XLSX.Charts.DrawingText(XLSX.Charts.DrawingParagraph("one", XLSX.Charts.DrawingRun("\n", kind=:br), "two"))
            g = rt(t)
            @test [r.kind for r in g.paragraphs[1].runs] == [:run, :br, :run]
            @test XLSX.Charts.text_content(g) == "one\ntwo"

            # schema order inside a:p and a:defRPr
            n = XLSX.Charts._text_from(XLSX.Charts.DrawingText(XLSX.Charts.DrawingParagraph("x",
                props=XLSX.Charts.DrawingParaProps(defprops=XLSX.Charts.DrawingRunProps(size=10.0)))), "txPr", pfx)
            p = XLSX.first_element_with_tag(n, "p")
            @test XLSX.localname.(collect(XML.eachelement(p))) == ["pPr", "r"]

            # what cannot be written says so
            @test_throws XLSX.XLSXError XLSX.Charts._fill_node_from(XLSX.Charts.DrawingFill(:gradient), pfx)
            @test_throws XLSX.XLSXError XLSX.Charts._color_node_from(
                XLSX.Charts.DrawingColor(:scrgb, "", Pair{Symbol,Int}[], "000000", 1.0), pfx)
        end

    end

    @testset "text_content reads a raw body as it reads a parsed one" begin
        xml = """<c:tx xmlns:c="$(XLSX.Charts.NS_C)" xmlns:a="$(XLSX.NS_A)"><c:rich><a:bodyPr/><a:lstStyle/>""" *
            """<a:p><a:r><a:t>One</a:t></a:r><a:br/><a:r><a:t>Two</a:t></a:r></a:p>""" *
            """<a:p><a:r><a:t>Three</a:t></a:r><a:fld id="{0}" type="x"><a:t>!</a:t></a:fld></a:p>""" *
            """</c:rich></c:tx>"""
        tx = XLSX.xml_root_element(parse(xml, XLSX.XML.Node))
        rich = XLSX.first_element_with_tag(tx, "rich")
        wb = XLSX.get_workbook(XLSX.newxlsx())
        @test XLSX.Charts.text_content(rich) == "One\nTwo\nThree!"
        @test XLSX.Charts.text_content(rich) == XLSX.Charts.text_content(XLSX.Charts.parse_drawing_text(wb, tx; tag="rich"))
    end

    @testset "show methods: DrawingLine and ChartMarker" begin
        # Both types previously fell back to Julia's default struct printing, which
        # dumps `raw`. These pin the summary and the multi-line forms.
        plain(x) = sprint(show, MIME("text/plain"), x)

        @testset "DrawingLine" begin
            # Built in code, so no `raw` and nothing inherited from a file.
            ln = XLSX.Charts.DrawingLine(; width=2.25, dash="sysDot", cap="rnd")
            @test repr(ln) == "XLSX.Charts.DrawingLine(2.25pt, sysDot)"
            s = plain(ln)
            @test occursin("stroke: inherited", s)
            @test occursin("width: 2.25 pt", s)
            @test occursin("dash: sysDot", s)
            @test occursin("cap: rnd", s)
            @test !occursin("compound", s)          # absent fields are not printed
            @test !occursin("miter", s)

            @test repr(XLSX.Charts.DrawingLine()) == "XLSX.Charts.DrawingLine(inherited)"

            none = XLSX.Charts.DrawingLine(; fill=XLSX.Charts.DrawingFill(:none))
            @test repr(none) == "XLSX.Charts.DrawingLine(no stroke)"
            @test occursin("stroke: none", plain(none))

            red = XLSX.Charts.DrawingLine(;
                fill=XLSX.Charts.DrawingFill(:solid; fgcolor="FF0000"),
                width=1.5)
            @test occursin("1.5pt", repr(red))
            @test occursin("stroke: ", plain(red))

            # Read from a file: the line on chart_appearance's first series is an
            # explicit <a:noFill/>, not an absent element.
            xf = XLSX.readxlsx(joinpath(data_directory, "chart_appearance.xlsx"))
            c = XLSX.Charts.getCharts(xf)[1]
            fl = XLSX.Charts.getSeriesShapeProps(c, 1).line
            @test repr(fl) == "XLSX.Charts.DrawingLine(no stroke)"
            @test !occursin("Element", plain(fl))   # `raw` stays out of the output
        end


    end

end