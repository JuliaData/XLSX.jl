#=
The text Excel shows for a cell (`getText`), and which number format codes are valid
(`setFormat`'s check). Both are compared with Excel 365's own results, read by COM into
test/data/number_formats_excel.tsv and test/data/number_format_validity_excel.tsv.
=#

function display_fixture_value(kind, v)
    kind == "number" && return parse(Float64, v)
    kind == "date" && return Float64(XLSX.date_to_excel_value(Date(v), false))
    kind == "datetime" && return XLSX.datetime_to_excel_value(DateTime(v), false)
    kind == "time" && return XLSX.time_to_excel_value(Time(v))
    kind == "bool" && return v == "true"
    return String(v)
end

@testset "rendering as Excel does" begin
    # Excel's results that getText doesn't reproduce: a fill repeats to the column's
    # width, which getText has no width for; and two times exactly half-way between
    # hundredths of a second, which Excel's floating-point arithmetic rounds up.
    known = Dict(("* 0", "42") => "42",
                 ("ss.00", "2026-03-07T14:05:09.125") => "09.12",
                 ("ss.00", "2026-03-07T14:05:09.135") => "09.13")
    n = 0
    for line in readlines(joinpath(data_directory, "number_formats_excel.tsv"))
        startswith(line, '#') && continue
        kind, v, fmt, text = split(line, '\t')
        excel = chop(text; head=1, tail=1)
        expected = get(known, (fmt, v), excel)
        got = XLSX._display_text(display_fixture_value(kind, v), fmt)
        got == expected || @info "differs from Excel" kind v fmt excel got
        @test got == expected
        n += 1
    end
    @test n == 201
end

@testset "format codes Excel accepts" begin
    for line in readlines(joinpath(data_directory, "number_format_validity_excel.tsv"))
        startswith(line, '#') && continue
        verdict, code = split(line, '\t'; limit=2)
        @test XLSX.is_valid_format(code) == (verdict == "accepts")
    end
    # setFormat takes the codes Excel accepts, and refuses the others
    f = XLSX.newxlsx()
    s = f[1]
    s["A1:A4"] = 1234.5
    @test XLSX.setFormat(s, "A1"; format="0.00E-00") >= XLSX.PREDEFINED_NUMFMT_COUNT
    @test XLSX.setFormat(s, "A2"; format="# ?/4") >= XLSX.PREDEFINED_NUMFMT_COUNT
    @test XLSX.getText(s, "A1") == "1.23E03"
    @test XLSX.getText(s, "A2") == "1234 2/4"
    @test_throws XLSX.XLSXError XLSX.setFormat(s, "A3"; format="[Purple]0")
    @test_throws XLSX.XLSXError XLSX.setFormat(s, "A3"; format="0 h")
    @test_throws XLSX.XLSXError XLSX.setFormat(s, "A3"; format="0;0;0;0;0")
    @test_throws XLSX.XLSXError XLSX.setFormat(s, "A3"; format="\"open")
    SAVE_FILES && save_outfile(f)
end

@testset "built-in formats" begin
    # Excel's codes for the built-ins (6, 8, 38 and 40 colour negatives red; 47 has its colon)
    @test XLSX.builtinFormats["6"] == "\$#,##0_);[Red](\$#,##0)"
    @test XLSX.builtinFormats["38"] == "#,##0_);[Red](#,##0)"
    @test XLSX.builtinFormats["47"] == "mm:ss.0"
    f = XLSX.newxlsx()
    s = f[1]
    s["A1:A3"] = -1234.5
    @test XLSX.setFormat(s, "A1"; format="\$#,##0_);[Red](\$#,##0)") == 6   # recognised as the built-in
    @test XLSX.getFormat(s, "A1").format["numFmt"]["formatCode"] == "\$#,##0_);[Red](\$#,##0)"
    @test XLSX.getText(s, "A1") == "(\$1,235)"
    XLSX.setFormat(s, "A2"; format="14")
    @test all(==('#'), XLSX.getText(s, "A2"))   # a negative date: Excel shows ####
    s["A3"] = 0.001                              # the value first: a number loses a date format
    XLSX.setFormat(s, "A3"; format="47")
    @test XLSX.getText(s, "A3") == "01:26.4"
    SAVE_FILES && save_outfile(f)
end

@testset "getText" begin
    f = XLSX.newxlsx("S")
    s = f[1]
    s["A1"] = 1234.5
    XLSX.setFormat(s, "A1"; format="#,##0.00")
    s["A2"] = Date(2026, 3, 7)                     # the default date format, 14
    s["A3"] = DateTime(2026, 3, 7, 14, 5)          # 22
    s["A4"] = Time(9, 5)                           # 20
    s["A5"] = "text"
    s["A6"] = true
    s["A7"] = 1 / 3
    s["A8"] = 42
    XLSX.setFormula(s, "A9", "=NA()"; value=XLSX.XL_NA)
    XLSX.setFormula(s, "A10", "=A8*2"; value=84)
    XLSX.setFormula(s, "A11", "=A8*3")             # no cached value
    s["A12"] = XLSX.RichTextString([XLSX.RichTextRun("bold", [:bold => true]), XLSX.RichTextRun(" plain", Pair{Symbol,Any}[])])
    s["A13"] = 0.256
    XLSX.setFormat(s, "A13"; format="0.0%")
    XLSX.addDefinedName(f, "amount", "S!A1")
    expected = ["1,234.50", "3/7/2026", "3/7/2026 14:05", "9:05", "text", "TRUE", "0.333333333",
                "42", "#N/A", "84", "", "bold plain", "25.6%"]
    @test [XLSX.getText(s, "A$r") for r in 1:13] == expected
    @test XLSX.getText(s, "B1") == ""                # empty
    @test XLSX.getText(s, "Z99") == ""               # outside the sheet
    @test XLSX.getText(s, 13, 1) == "25.6%"
    @test XLSX.getText(s, XLSX.CellRef("A1")) == "1,234.50"
    @test XLSX.getText(f, "S!A2") == "3/7/2026"
    @test XLSX.getText(s, "amount") == "1,234.50"
    @test_throws XLSX.XLSXError XLSX.getText(s, "not a cell")

    # read only, with and without the cache
    io = IOBuffer()
    XLSX.writexlsx(io, f)
    SAVE_FILES && save_outfile(f)
    bytes = take!(io)
    for enable_cache in (true, false)
        XLSX.openxlsx(IOBuffer(bytes); enable_cache) do r
            @test [XLSX.getText(r[1], "A$i") for i in 1:13] == expected
        end
    end

    # the 1904 date system, and the days Excel's 1900 system gets wrong
    @test XLSX._display_text(0, "yyyy-mm-dd"; date1904=true) == "1904-01-01"
    @test XLSX._display_text(1, "dddd"; date1904=true) == "Saturday"
    @test XLSX._display_text(60, "yyyy-mm-dd") == "1900-02-29"
    @test XLSX._display_text(1, "dddd") == "Sunday"             # as Excel has it

    # padding measured for width: `_x` as `x`, `?` as a digit
    @test XLSX._display_text(1.5, "_(0.0?_)") == " 1.5  "
    @test XLSX._display_text(1.5, "_(0.0?_)"; for_width=true) == "(1.50)"
end
