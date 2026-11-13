#=
autofitColumns, compared with Excel 365's own AutoFit of the same cells:
test/data/autofit_excel.xlsx is the workbook as Excel saved it, and
test/data/autofit_excel_widths.tsv the widths Excel gave each column.
=#

@testset "widths as Excel's AutoFit sets them" begin
    xf = XLSX.openxlsx(joinpath(data_directory, "autofit_excel.xlsx"); mode="rw")
    s = xf[1]
    got = XLSX.autofitColumns(s)
    errors = Float64[]
    for line in readlines(joinpath(data_directory, "autofit_excel_widths.tsv"))
        startswith(line, '#') && continue
        col, width, text = split(line, '\t')
        c = parse(Int, col)
        c == 29 && continue   # format 39, which Excel showed in the UK locale as -1,234.50
        push!(errors, got[c] - parse(Float64, width))
    end
    # approximate: the median within 0.25 characters, each within 1.5 (the widest
    # miss is Arial at 12 points, from font hinting at another size than measured)
    @test sort(abs.(errors))[cld(length(errors), 2)] <= 0.25
    @test maximum(abs.(errors)) <= 1.5
    @test XLSX.getColumnWidth(s, 1, 37) ≈ got[37] + XLSX.EXCEL_COLUMN_WIDTH_PADDING
    SAVE_FILES && save_outfile(xf)
end

@testset "autofitColumns" begin
    f = XLSX.newxlsx()
    s = f[1]
    s["A1"] = "short"
    s["A2"] = "a much longer piece of text"
    s["B1"] = 1234567.891
    XLSX.setFormat(s, "B1"; format="#,##0.00")
    s["C1"] = Date(2026, 3, 7)
    XLSX.setFormat(s, "C1"; format="dddd, mmmm d, yyyy")
    s["D1"] = "two\nlines, the second longer"
    s["G1"] = "bold"
    XLSX.setFont(s, "G1"; bold=true, size=20)
    s["E1"] = "merged across E1:F1, so left out"
    XLSX.mergeCells(s, "E1:F1")
    s["E2"] = "e"
    s["H1"] = ""                                   # shows nothing
    XLSX.setColumnWidth(s, 8; width=30)

    widths = XLSX.autofitColumns(s)
    @test sort(collect(keys(widths))) == [1, 2, 3, 4, 5, 7]    # H shows nothing: kept
    @test widths[1] > widths[5]                                 # the longer text
    @test 9.5 < widths[2] < 12      # "1,234,567.89", as formatted (Excel gives 10.67)
    @test widths[5] < 3                                         # only "e": E1 is merged
    @test widths[4] < XLSX.autofitColumns(s, "A")[1]            # the longest line of D1 is shorter than A2
    @test widths[7] > 5                                         # 20 pt bold
    @test XLSX.getColumnWidth(s, 1, 8) ≈ 30 + XLSX.EXCEL_COLUMN_WIDTH_PADDING
    @test XLSX.getColumnWidth(s, 1, 1) ≈ widths[1] + XLSX.EXCEL_COLUMN_WIDTH_PADDING

    # columns by letter, number, range or vector, and the width limits
    @test collect(keys(XLSX.autofitColumns(s, "B:C"))) |> sort == [2, 3]
    @test collect(keys(XLSX.autofitColumns(s, "A,C:D"))) |> sort == [1, 3, 4]
    @test collect(keys(XLSX.autofitColumns(s, 2))) == [2]
    @test collect(keys(XLSX.autofitColumns(s, [1, 7]))) |> sort == [1, 7]
    @test collect(keys(XLSX.autofitColumns(s, 1:2))) |> sort == [1, 2]
    @test XLSX.autofitColumns(s, "A"; max_width=10)[1] == 10
    @test XLSX.autofitColumns(s, "E"; min_width=8)[5] == 8
    SAVE_FILES && save_outfile(f)

    @test_throws XLSX.XLSXError XLSX.autofitColumns(s; max_width=300)
    @test_throws XLSX.XLSXError XLSX.autofitColumns(s; min_width=20, max_width=10)
    @test_throws XLSX.XLSXError XLSX.autofitColumns(s, "B2")
    @test_throws XLSX.XLSXError XLSX.autofitColumns(s, 1.5)
    io = IOBuffer()
    XLSX.writexlsx(io, f)
    @test_throws XLSX.XLSXError XLSX.autofitColumns(XLSX.readxlsx(seekstart(io))[1])   # read only

    # fonts: a family not in the table is measured as its nearest name, or as Calibri
    @test XLSX._font_widths("Calibri Light", false) === XLSX._font_widths("calibri", false)
    @test XLSX._font_widths("Aptos Narrow Display", false) === XLSX._font_widths("aptos narrow", false)
    @test XLSX._font_widths("Comic Sans MS", true) === XLSX._font_widths("Calibri", true)
    @test XLSX._char_em(XLSX._font_widths("Calibri", false), '中') ≈ 2 * XLSX._char_em(XLSX._font_widths("Calibri", false), '0')
end
