@testset "Colors" begin
    xf = XLSX.readxlsx(joinpath(data_directory, "is.xlsx"))
    sh = xf["Sheet1"]
    wb = XLSX.get_workbook(xf)

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
end