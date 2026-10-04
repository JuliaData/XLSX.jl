@testset "Colors" begin
    xf = XLSX.readxlsx(joinpath(data_directory, "is.xlsx"))
    sh = xf["Sheet1"]
    wb = XLSX.get_workbook(xf)

    @testset "get_color accepts every hex form" begin
        @test XLSX.get_color("FF0000")    == "FFFF0000"
        @test XLSX.get_color("ff0000")    == "FFFF0000"
        @test XLSX.get_color("#FF0000")   == "FFFF0000"
        @test XLSX.get_color("80FF0000")  == "80FF0000"
        @test XLSX.get_color("#FFFF0000") == "FFFF0000"      # ARGB red, not CSS yellow
        @test XLSX.get_color("red")       == "FFFF0000"
        @test XLSX.get_color(:red)        == "FFFF0000"
        @test XLSX.get_color("grey")      == "FF808080"
        @test_throws XLSX.XLSXError XLSX.get_color("notacolour")
    end
    
    @testset "theme_xmlroot" begin
        # theme1.xml parses cleanly and has the correct root element
        theme_root = XLSX.theme_xmlroot(wb)
        @test theme_root !== nothing

        # second call returns the cached value, not a fresh parse
        @test XLSX.theme_xmlroot(wb) === theme_root

        # the 12 theme colors are loaded in OOXML index order (lt1, dk1, lt2, dk2, ...)
        colors = XLSX.get_theme_colors(wb)
        @test length(colors) == 12
        @test colors[1]  == "FFFFFF"   # theme=0: lt1
        @test colors[2]  == "000000"   # theme=1: dk1
        @test colors[3]  == "E8E8E8"   # theme=2: lt2
        @test colors[4]  == "0E2841"   # theme=3: dk2
        @test colors[5]  == "156082"   # theme=4: accent1
        @test colors[10] == "4EA72E"   # theme=9: accent6
        @test colors[11] == "467886"   # theme=10: hlink
        @test colors[12] == "96607D"   # theme=11: folHlink
    end

    @testset "resolveColor" begin
        # rgb: returned unchanged
        @test XLSX.resolveColor(wb, Dict("rgb" => "FFAABBCC")) == "FFAABBCC"

        # auto: resolves to black
        @test XLSX.resolveColor(wb, Dict("auto" => "1")) == "FF000000"

        # no recognised key: falls back to black
        @test XLSX.resolveColor(wb, Dict{String,String}()) == "FF000000"

        # indexed: prepends FF to INDEXED_PALETTE entry
        @test XLSX.resolveColor(wb, Dict("indexed" => "0")) == "FF000000"
        @test XLSX.resolveColor(wb, Dict("indexed" => "2")) == "FFFF0000"

        # theme, no tint
        @test XLSX.resolveColor(wb, Dict("theme" => "0")) == "FFFFFFFF"   # lt1: white
        @test XLSX.resolveColor(wb, Dict("theme" => "1")) == "FF000000"   # dk1: black
        @test XLSX.resolveColor(wb, Dict("theme" => "9")) == "FF4EA72E"   # accent6

        # theme with positive tint (lightens toward white)
        @test XLSX.resolveColor(wb, Dict("theme" => "3", "tint" => "0.24994659260841701")) == "FF4A5E70"

        # theme with negative tint (darkens toward black)
        @test XLSX.resolveColor(wb, Dict("theme" => "4", "tint" => "-0.5")) == "FF0A3041"

        # Worksheet and XLSXFile dispatch methods reach the same result
        @test XLSX.resolveColor(sh, Dict("theme" => "9")) == "FF4EA72E"
        @test XLSX.resolveColor(xf, Dict("theme" => "9")) == "FF4EA72E"

        # prefix kwarg — same resolution, just different key names
        @test XLSX.resolveColor(wb, Dict("fgrgb" => "FFAABBCC"); prefix="fg") == "FFAABBCC"
        @test XLSX.resolveColor(wb, Dict("fgtheme" => "1"); prefix="fg") == "FF000000"

        # theme colors are consistent with what getRichTextString resolved in the actual file
        # is.xlsx has theme="1" runs (dk1, black) and theme="9" runs (accent6, green)
        rts = XLSX.getRichTextString(sh, "B2")
        dk1_run = findfirst(r -> r.atts !== nothing && get(r.atts, :color, nothing) == "FF000000", rts.runs)
        @test dk1_run !== nothing
    end

    @testset "resolveColor (Borders.xlsx)" begin
        f = XLSX.readxlsx(joinpath(data_directory, "Borders.xlsx"))
        s = f["Sheet1"]
        wb = XLSX.get_workbook(f)

        # theme_xmlroot works on a different fixture - same theme, independent cache
        theme_root = XLSX.theme_xmlroot(wb)
        @test theme_root !== nothing
        colors = XLSX.get_theme_colors(wb)
        @test colors[4] == "0E2841"   # dk2, index 3

        # D4: theme="3" tint="0.24994659260841701" - the motivating example
        d4_top = XLSX.getBorder(s, "D4").border["top"]
        @test XLSX.resolveColor(s, d4_top) == "FF4A5E70"

        # all four sides of D4 have the same color
        for side in ("left", "right", "top", "bottom")
            @test XLSX.resolveColor(wb, XLSX.getBorder(s, "D4").border[side]) == "FF4A5E70"
        end

        # B4: explicit rgb - resolveColor passes it through unchanged
        b4_top = XLSX.getBorder(s, "B4").border["top"]
        @test XLSX.resolveColor(s, b4_top) == "FFFF0000"

        # B2: auto color on border
        b2_top = XLSX.getBorder(s, "B2").border["top"]
        @test XLSX.resolveColor(s, b2_top) == "FF000000"
    end

    @testset "rgb normalization" begin
        @test XLSX.resolveColor(wb, Dict("rgb" => "aabbcc")) == "FFAABBCC"
        @test XLSX.resolveColor(wb, Dict("rgb" => "  aabbcc  ")) == "FFAABBCC"
        @test XLSX.resolveColor(wb, Dict("rgb" => "00ff00")) == "FF00FF00"
        @test XLSX.resolveColor(wb, Dict("rgb" => "80AABBCC")) == "80AABBCC"
        @test XLSX.resolveColor(wb, Dict("rgb" => "80aabbcc")) == "80AABBCC"
        @test_throws XLSX.XLSXError XLSX.resolveColor(wb, Dict("rgb" => "ABC"))   # if you choose to throw on bad format
    end

    @testset "indexed bounds" begin
        maxidx = length(XLSX.INDEXED_PALETTE) - 1
        @test XLSX.resolveColor(wb, Dict("indexed" => "0")) == "FF" * XLSX.INDEXED_PALETTE[1]
        @test XLSX.resolveColor(wb, Dict("indexed" => string(maxidx))) == "FF" * XLSX.INDEXED_PALETTE[maxidx+1]
        # If you choose to throw on out-of-range:
        @test_throws XLSX.XLSXError XLSX.resolveColor(wb, Dict("indexed" => string(maxidx+1)))
        @test_throws XLSX.XLSXError XLSX.resolveColor(wb, Dict("indexed" => "-1"))
    end

    @testset "extremes and invalid values" begin
        @test XLSX.resolveColor(wb, Dict("theme" => "3", "tint" => "0")) == XLSX.resolveColor(wb, Dict("theme" => "3"))
        @test XLSX.resolveColor(wb, Dict("theme" => "3", "tint" => "1.0"))  == "FFFFFFFF"
        @test XLSX.resolveColor(wb, Dict("theme" => "3", "tint" => "-1.0")) == "FF000000"
        @test_throws XLSX.XLSXError XLSX.resolveColor(wb, Dict("theme" => "3", "tint" => "nan"))
        @test_throws XLSX.XLSXError XLSX.resolveColor(wb, Dict("theme" => "3", "tint" => "inf"))
        @test_throws XLSX.XLSXError XLSX.resolveColor(wb, Dict("theme" => "3", "tint" => "notanumber"))
        @test_throws XLSX.XLSXError XLSX.resolveColor(wb, Dict("indexed" => "64"))
        @test_throws XLSX.XLSXError XLSX.resolveColor(wb, Dict("theme" => "notanint"))
        @test_throws XLSX.XLSXError XLSX.resolveColor(wb, Dict("indexed" => "foo"))
        @test_throws XLSX.XLSXError XLSX.resolveColor(wb, Dict("theme" => "-1"))
        @test_throws XLSX.XLSXError XLSX.resolveColor(wb, Dict("theme" => "12"))
        @test_throws XLSX.XLSXError XLSX.resolveColor(wb, Dict("theme" => "notanumber"))
    end

    @testset "theme colour transforms match Excel" begin
        f = XLSX.opentemplate(joinpath(data_directory, "chart_theme_colors.xlsx"))
        wb = XLSX.get_workbook(f)
        ch = XLSX.xml_root_element(f.data["xl/charts/chart1.xml"])

        cols = XLSX.Charts.DrawingColor[]
        walk(n) = for c in XML.eachelement(n)
            XLSX.localname(c) == "solidFill" ?
                (col = XLSX.Charts.parse_drawing_color(wb, c); isnothing(col) || push!(cols, col)) :
                walk(c)
        end
        walk(ch)

        # Values confirmed against what Excel reports in More Fill Colors.
        function byval(v, t)
            hits = filter(c -> c.val == v && c.transforms == t, cols)
            @test !isempty(hits)
            @test allequal(c.rgb for c in hits)     # same input, same output everywhere
            return first(hits).rgb
        end

        @test byval("accent1", Pair{Symbol,Int}[])                              == "156082"
        @test byval("accent1", [:lumMod => 60000, :lumOff => 40000])            == "46B1E1"
        @test byval("accent1", [:lumMod => 75000])                              == "104862"
        @test byval("tx1",     [:lumMod => 65000, :lumOff => 35000])            == "595959"
        @test byval("tx1",     [:lumMod => 15000, :lumOff => 85000])            == "D9D9D9"
        SAVE_FILES && save_outfile(f)
    end

    @testset "every colour argument takes every colour form" begin
        # The same colour in each accepted form. Coral rather than red, so a
        # named colour is resolved through Colors.jl rather than a short table.
        forms = ("FFFF7F50", "#FF7F50", :coral, "coral")
        # Chart colours read back as RRGGBB, cell colours as AARRGGBB.
        iscoral(rgb) = !isnothing(rgb) && endswith(uppercase(String(rgb)), "FF7F50")

        function first_descendant(n, tag)
            for k in XML.eachelement(n)
                XLSX.localname(k) == tag && return k
                d = first_descendant(k, tag)
                isnothing(d) || return d
            end
            return nothing
        end
        # The colour in an element's own c:txPr
        txfill(n) = XLSX.Charts._attr(
            first_descendant(XLSX.first_element_with_tag(n, "txPr"), "srgbClr"), "val")

        tmp = "colour_forms.xlsx"
        cp(joinpath(data_directory, "chart_kinds.xlsx"), tmp; force = true)
        try
            xf = XLSX.openxlsx(tmp; mode = "rw")
            c  = first(XLSX.Charts.getCharts(xf["stackedbar"]))
            lc = first(XLSX.Charts.getCharts(xf["linemarkers"]))
            catax = only(XLSX.Charts.getChartAxes(c, :category))
            c = XLSX.Charts.setAxisTitleText(c, catax, "Questions")
            ws = XLSX.addsheet!(xf, "colours")

            for (i, col) in enumerate(forms)
                @testset "$(repr(col))" begin
                    # series fill and outline
                    c = XLSX.Charts.setSeriesFill(c, 1, col)
                    @test iscoral(XLSX.Charts.getSeriesFill(c, 1).value.fgcolor.rgb)
                    c = XLSX.Charts.setSeriesLine(c, 1; color = col)
                    @test iscoral(XLSX.Charts.getSeriesLine(c, 1).value.fill.fgcolor.rgb)
                    c = XLSX.Charts.setSeriesLineColor(c, 2, col)
                    @test iscoral(XLSX.Charts.getSeriesLine(c, 2).value.fill.fgcolor.rgb)

                    # addSeries
                    c = XLSX.Charts.addSeries(c, "stackedbar!B2:B5"; color = col)
                    n = length(XLSX.Charts.getChartSeries(c))
                    @test iscoral(XLSX.Charts.getSeriesFill(c, n).value.fgcolor.rgb)

                    # axis line and plot area border
                    c = XLSX.Charts.setAxisLine(c, catax; color = col)
                    @test iscoral(XLSX.Charts.getAxisShapeProps(c, catax).line.fill.fgcolor.rgb)
                    c = XLSX.Charts.setPlotAreaLine(c; color = col)
                    @test iscoral(XLSX.Charts.getPlotAreaShapeProps(c).line.fill.fgcolor.rgb)

                    # text colour, at every text setter
                    c = XLSX.Charts.setLabelTextProp(c, 1, :fill, col)
                    @test iscoral(XLSX.Charts.getLabelTextProp(c, 1, :fill).value.fgcolor.rgb)
                    c = XLSX.Charts.setChartTitleTextProp(c, :fill, col)
                    @test iscoral(txfill(XLSX.Charts.getChartTitleNode(c)))
                    c = XLSX.Charts.setLegendTextProp(c, :fill, col)
                    @test iscoral(txfill(XLSX.Charts.getChartLegend(c)))
                    c = XLSX.Charts.setAxisTextProp(c, catax, :fill, col)
                    @test iscoral(txfill(XLSX.Charts._axnode(c, catax)))
                    c = XLSX.Charts.setAxisTitleTextProp(c, catax, :fill, col)
                    @test iscoral(txfill(XLSX.first_element_with_tag(XLSX.Charts._axnode(c, catax), "title")))
                    c = XLSX.Charts.setChartSpaceTextProp(c, :fill, col)
                    @test iscoral(txfill(XLSX.Charts.chart_root(c)))

                    # markers
                    lc = XLSX.Charts.setMarkerFill(lc, 1, col)
                    @test iscoral(XLSX.Charts.getSeriesMarker(lc, 1).shape.fill.fgcolor.rgb)
                    lc = XLSX.Charts.setMarkerLineColor(lc, 1, col)
                    @test iscoral(XLSX.Charts.getSeriesMarker(lc, 1).shape.line.fill.fgcolor.rgb)

                    # cells: Symbol colours not yet accepted (borders need work)
                    ws["A$i"] = "x"
                    if col isa Symbol
                        @test_broken (XLSX.setFont(ws, "A$i"; color = col); true) # formatting setters can't take a symbol for a colour yet.
                    else
                        XLSX.setFont(ws, "A$i"; color = col)
                        @test iscoral(XLSX.getFont(ws, "A$i").font["color"]["rgb"])
                    end
                end
            end

            # The last form's colours survive a write and reopen
            SAVE_FILES && save_outfile(xf)
            XLSX.writexlsx(tmp, xf; overwrite = true)
            xf2 = XLSX.openxlsx(tmp; mode = "rw")
            SAVE_FILES && save_outfile(xf2)
            c2 = first(XLSX.Charts.getCharts(xf2["stackedbar"]))
            @test iscoral(XLSX.Charts.getSeriesFill(c2, 1).value.fgcolor.rgb)
            @test iscoral(XLSX.Charts.getPlotAreaShapeProps(c2).line.fill.fgcolor.rgb)
        finally
            isfile(tmp) && rm(tmp)
        end
    end
end