# Formatting a range whose cells start with different styles. `setFont` and the
# other per-cell setters merge their keywords into each cell's own format, the
# `setUniform*` setters give every cell the first cell's, and the styles written
# don't depend on how the range is given, or on whether it is given at once.

const MIXED_ROWS  = 40
const MIXED_BANDS = 8      # bands of 5 rows, each starting with a font of its own

_band_color(k) = "FF" * string((k + 1) * 0x1f3d5b % 0xffffff; base=16, pad=6) |> uppercase
_band_rows(k) = 5k+1:5k+5

# A sheet of 40 x 6 numbers whose bands of rows start with distinct fonts.
function _banded_sheet()
    xf = XLSX.newxlsx()
    s = xf[1]
    s[1:MIXED_ROWS, 1:6] = [r * 10 + c for r in 1:MIXED_ROWS, c in 1:6]
    for k in 0:MIXED_BANDS-1
        XLSX.setFont(s, "A$(first(_band_rows(k))):F$(last(_band_rows(k)))"; size=8 + k, color=_band_color(k))
    end
    return xf, s
end

# Every cell of every band keeps its band's colour and size, and passes `extra`.
function _bands_keep_font(s, extra)
    ok = true
    for k in 0:MIXED_BANDS-1, r in _band_rows(k), c in 1:6
        f = XLSX.getFont(s, r, c).font
        ok &= f["color"]["rgb"] == _band_color(k) && f["sz"]["val"] == string(8 + k) && extra(f)
    end
    return ok
end

_atts(run) = isnothing(run.atts) ? Dict{Symbol,Any}() : Dict(run.atts)
_part(xf, name) = (io = IOBuffer(); XLSX.writexlsx(io, xf); zip_readentry(ZipReader(take!(io)), name, String))

# --- workload (also run, before the change to memoised range formatting, to make
# --- the files in test/data/format_identity)
function _format_identity_workload()
    xf = XLSX.newxlsx()
    s = xf[1]
    s[1:40, 1:6] = [r * 10 + c for r in 1:40, c in 1:6]
    rich = XLSX.RichTextString([XLSX.RichTextRun("ab", [:bold => true, :color => "FF0000"]),
                                XLSX.RichTextRun("cd", [:italic => true])])
    s["A41"] = rich
    s["B41"] = rich
    s["C41"] = "plain"
    s[41, 4:6] = [1.5 2.5 3.5]
    for k in 0:7
        XLSX.setFont(s, "A$(5k+1):F$(5k+5)"; size=8 + k, color="FF" * lpad(string(k * 0x202020; base=16), 6, '0'))
    end
    XLSX.setFill(s, "A1:C40"; pattern="solid", fgColor="FFFFFF00")
    XLSX.setFill(s, "D1:F20"; pattern="darkGrid", fgColor="red")
    XLSX.setBorder(s, "B2:E30"; left=["style" => "thin", "color" => "blue"])
    XLSX.setBorder(s, "A1:F40"; allsides=["style" => "medium"])
    XLSX.setAlignment(s, "A1:F20"; vertical="top")
    XLSX.setAlignment(s, "A1:F40"; horizontal="center")
    XLSX.setFormat(s, "A1:A40"; format="0.000")
    XLSX.setFormat(s, "B1:C40"; format="Percentage")
    XLSX.setFont(s, "A1:F41"; bold=true)
    XLSX.setFont(s, "A:B"; italic=true)
    XLSX.setFont(s, "3:4"; under="double")
    XLSX.setFill(s, "A1:A3,C5:D9,F12"; pattern="lightGrid", bgColor="green")
    XLSX.setUniformFont(s, "A30:F35"; strike=true)
    XLSX.setUniformFill(s, "E1:F10"; pattern="gray125")
    XLSX.setUniformBorder(s, "A36:C40"; top=["style" => "thick"])
    XLSX.setUniformAlignment(s, "D36:F40"; wrapText=true)
    XLSX.setUniformFormat(s, "A36:F38"; format="#,##0.0")
    XLSX.setOutsideBorder(s, "B10:D15"; outside=["style" => "dotted"])
    XLSX.setFont(s, 1:40, [2, 4]; color="orange")
    XLSX.setFont(s, [1, 3, 5], :; size=20)
    XLSX.setFont(s, "A1:F41"; name="Arial", vertAlign="superscript")
    return xf
end
# --- end of workload

@testset "mixed-style ranges" begin

    @testset "setFont keeps fonts" begin
        forms = Any[
            s -> XLSX.setFont(s, "A1:F40"; bold=true),
            s -> XLSX.setFont(s, "A:F"; bold=true),
            s -> XLSX.setFont(s, "1:40"; bold=true),
            s -> XLSX.setFont(s, "A1:C40,D1:F20,D21:F40"; bold=true),
            s -> XLSX.setFont(s, 1:40, [1, 2, 3, 4, 5, 6]; bold=true),
            s -> XLSX.setFont(s, :, :; bold=true),
            s -> XLSX.setFont(s, 1:40, 1:6; bold=true),
        ]
        for f in forms
            xf, s = _banded_sheet()
            @test f(s) == -1
            @test _bands_keep_font(s, f -> haskey(f, "b"))
        end
        xf, s = _banded_sheet()
        XLSX.addDefinedName(xf, "allcells", "Sheet1!A1:F40")
        XLSX.setFont(xf, "allcells"; bold=true)
        @test _bands_keep_font(s, f -> haskey(f, "b"))
        XLSX.setFont(s, "A1:F1"; italic=true)          # a second call on the same styles
        @test _bands_keep_font(s, f -> haskey(f, "b"))
        @test all(haskey(XLSX.getFont(s, 1, c).font, "i") for c in 1:6)
        @test !haskey(XLSX.getFont(s, 2, 1).font, "i")
        @test XLSX.setFont(s, "C7:C7"; size=30) == XLSX.getFont(s, 7, 3).fontId   # a one-cell range returns its id
    end

    @testset "setFill keeps colours" begin
        xf, s = _banded_sheet()
        for k in 0:MIXED_BANDS-1
            XLSX.setFill(s, "A$(first(_band_rows(k))):F$(last(_band_rows(k)))"; pattern="solid", fgColor=_band_color(k))
        end
        XLSX.setFill(s, "A1:F40"; pattern="darkGrid")
        @test all(XLSX.getFill(s, r, c).fill["patternFill"]["fgrgb"] == _band_color(k) &&
                  XLSX.getFill(s, r, c).fill["patternFill"]["patternType"] == "darkGrid"
                  for k in 0:MIXED_BANDS-1 for r in _band_rows(k) for c in 1:6)
        @test _bands_keep_font(s, f -> true)
    end

    @testset "setBorder keeps colours" begin
        xf, s = _banded_sheet()
        for k in 0:MIXED_BANDS-1
            XLSX.setBorder(s, "A$(first(_band_rows(k))):F$(last(_band_rows(k)))"; left=["style" => "dashed", "color" => _band_color(k)])
        end
        XLSX.setBorder(s, "A1:F40"; allsides=["style" => "thin"])
        ok = true
        for k in 0:MIXED_BANDS-1, r in _band_rows(k), c in 1:6
            b = XLSX.getBorder(s, r, c).border
            ok &= b["left"]["rgb"] == _band_color(k) && all(b[side]["style"] == "thin" for side in ("left", "right", "top", "bottom"))
        end
        @test ok
        @test _bands_keep_font(s, f -> true)
    end

    @testset "setAlignment merges" begin
        xf, s = _banded_sheet()
        XLSX.setAlignment(s, "A1:F20"; vertical="top")
        XLSX.setAlignment(s, "A21:F40"; vertical="bottom", wrapText=true)
        XLSX.setAlignment(s, "A1:F40"; horizontal="center")
        a(r, c) = XLSX.getAlignment(s, r, c).alignment["alignment"]
        @test all(a(r, c) == Dict("horizontal" => "center", "vertical" => "top") for r in 1:20, c in 1:6)
        @test all(a(r, c) == Dict("horizontal" => "center", "vertical" => "bottom", "wrapText" => "1") for r in 21:40, c in 1:6)
        @test _bands_keep_font(s, f -> true)
    end

    @testset "setFormat keeps fonts" begin
        xf, s = _banded_sheet()
        id = XLSX.setFormat(s, "B2:B2"; format="0.0000")
        XLSX.setFormat(s, "A1:F40"; format="0.0000")
        @test all(XLSX.getFormat(s, r, c).numFmtId == id for r in 1:40, c in 1:6)
        @test _bands_keep_font(s, f -> true)
    end

    @testset "setUniform makes alike" begin
        xf, s = _banded_sheet()
        id = XLSX.setUniformFont(s, "A1:F40"; bold=true)
        @test all(XLSX.getFont(s, r, c).fontId == id for r in 1:40, c in 1:6)
        @test XLSX.getFont(s, 40, 6).font["color"]["rgb"] == _band_color(0)

        xf, s = _banded_sheet()
        XLSX.setFill(s, "A21:F40"; pattern="solid", fgColor="red")
        id = XLSX.setUniformFill(s, "A1:F40"; pattern="lightGrid")
        @test all(XLSX.getFill(s, r, c).fillId == id for r in 1:40, c in 1:6)
        @test _bands_keep_font(s, f -> true)                # only the fill is made uniform

        xf, s = _banded_sheet()
        XLSX.setBorder(s, "A21:F40"; left=["style" => "thick"])
        id = XLSX.setUniformBorder(s, "A1:F40"; top=["style" => "thin"])
        @test all(XLSX.getBorder(s, r, c).borderId == id for r in 1:40, c in 1:6)

        xf, s = _banded_sheet()
        XLSX.setAlignment(s, "A21:F40"; vertical="bottom")
        XLSX.setUniformAlignment(s, "A1:F40"; horizontal="right")
        @test all(XLSX.getAlignment(s, r, c).alignment["alignment"] == Dict("horizontal" => "right") for r in 1:40, c in 1:6)
        @test _bands_keep_font(s, f -> true)

        xf, s = _banded_sheet()
        XLSX.setFormat(s, "A21:F40"; format="0.0")
        id = XLSX.setUniformFormat(s, "A1:F40"; format="0.00%")
        @test all(XLSX.getFormat(s, r, c).numFmtId == id for r in 1:40, c in 1:6)
        @test _bands_keep_font(s, f -> true)

        xf, s = _banded_sheet()                             # vector and non-contiguous forms
        id = XLSX.setUniformFont(s, 1:40, [1, 3, 5]; italic=true)
        @test all(XLSX.getFont(s, r, c).fontId == id for r in 1:40, c in (1, 3, 5))
        @test !haskey(XLSX.getFont(s, 1, 2).font, "i")
        id = XLSX.setUniformFill(s, "A1:A40,C1:C40"; pattern="solid", fgColor="blue")
        @test all(XLSX.getFill(s, r, c).fillId == id for r in 1:40, c in (1, 3))
    end

    @testset "rich text" begin
        xf, s = _banded_sheet()
        rich = XLSX.RichTextString([XLSX.RichTextRun("ab", [:bold => true, :color => "FF0000"]),
                                    XLSX.RichTextRun("cd", [:italic => true])])
        s["A1"] = rich
        s["B1"] = rich
        s["A12"] = rich
        s["C1"] = "plain"
        XLSX.setFont(s, "A1:F40"; bold=true, italic=false)
        for ref in ("A1", "B1", "A12")
            runs = XLSX.getRichTextString(s, ref).runs
            @test [r.text for r in runs] == ["ab", "cd"]
            @test !any(haskey(_atts(r), k) for r in runs for k in (:bold, :italic))
            @test _atts(runs[1])[:color] == "FFFF0000"
            @test haskey(XLSX.getFont(s, ref).font, "b")
        end
        @test s["C1"] == "plain" && haskey(XLSX.getFont(s, "C1").font, "b")
        @test _bands_keep_font(s, f -> haskey(f, "b"))

        xf, s = _banded_sheet()
        s["A1"] = rich
        s["B1"] = rich
        XLSX.setUniformFont(s, "A1:F40"; size=9)
        @test !any(haskey(_atts(r), :size) for ref in ("A1", "B1") for r in XLSX.getRichTextString(s, ref).runs)
    end

    @testset "range same as by cell" begin
        edits = [
            (XLSX.setFont,      (; bold=true, color="red")),
            (XLSX.setFill,      (; pattern="solid", fgColor="FF00FF00")),
            (XLSX.setBorder,    (; left=["style" => "thick"], diagonal=["style" => "thin", "direction" => "up"])),
            (XLSX.setBorder,    (; allsides=["style" => "hair", "color" => "blue"])),
            (XLSX.setAlignment, (; horizontal="left", wrapText=true)),
            (XLSX.setFormat,    (; format="0.0;[Red]-0.0")),
        ]
        for (f, kw) in edits
            xa, sa = _banded_sheet()
            xb, sb = _banded_sheet()
            f(sa, "A1:F40"; kw...)
            for r in 1:40, c in 1:6
                f(sb, r, c; kw...)
            end
            @test _part(xa, "xl/styles.xml") == _part(xb, "xl/styles.xml")
            @test _part(xa, "xl/worksheets/sheet1.xml") == _part(xb, "xl/worksheets/sheet1.xml")
        end
    end

    @testset "output unchanged" begin
        # Files written by the workload above, on Julia 1.13, before range formatting
        # was memoised. `setAlignment` takes an `<alignment>`'s attributes from a Dict,
        # whose order differs between Julia versions, so styles.xml is compared as XML
        # (attributes in any order, elements, and so ids, in the same order).
        dir = joinpath(data_directory, "format_identity")
        xf = _format_identity_workload()
        @test parse(_part(xf, "xl/styles.xml"), XML.Node) == parse(read(joinpath(dir, "styles.xml"), String), XML.Node)
        for (part, file) in (("xl/worksheets/sheet1.xml", "sheet1.xml"), ("xl/sharedStrings.xml", "sharedStrings.xml"))
            @test _part(xf, part) == read(joinpath(dir, file), String)
        end
    end
end
