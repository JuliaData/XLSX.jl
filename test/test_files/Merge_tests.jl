@testset "merged cells" begin
    XLSX.openxlsx(joinpath(data_directory, "customXml.xlsx")) do f
        @test_throws XLSX.XLSXError XLSX.getMergedCells(f["Mock-up"]) # File isn't writeable
    end
    f = XLSX.openxlsx(joinpath(data_directory, "customXml.xlsx"); mode="rw")
    mc = sort(XLSX.getMergedCells(f["Mock-up"]))
    @test length(mc) == 25
    @test mc == sort(XLSX.CellRange[XLSX.CellRange("D49:H49"), XLSX.CellRange("D72:J72"), XLSX.CellRange("F94:J94"), XLSX.CellRange("F96:J96"), XLSX.CellRange("F84:J84"), XLSX.CellRange("F86:J86"), XLSX.CellRange("D62:J63"), XLSX.CellRange("D51:J53"), XLSX.CellRange("D55:J60"), XLSX.CellRange("D92:J92"), XLSX.CellRange("D82:J82"), XLSX.CellRange("D74:J74"), XLSX.CellRange("D67:J68"), XLSX.CellRange("D47:H47"), XLSX.CellRange("D9:H9"), XLSX.CellRange("D11:G11"), XLSX.CellRange("D12:G12"), XLSX.CellRange("D14:E14"), XLSX.CellRange("D16:E16"), XLSX.CellRange("D32:F32"), XLSX.CellRange("D38:J38"), XLSX.CellRange("D34:J34"), XLSX.CellRange("D18:E18"), XLSX.CellRange("D20:E20"), XLSX.CellRange("D13:G13")])
    s = f["Mock-up"]
    @test XLSX.isMergedCell(f, "Mock-up!D47")
    @test XLSX.isMergedCell(f, "Mock-up!D49"; mergedCells=mc)
    @test XLSX.isMergedCell(s, "H84")
    @test XLSX.isMergedCell(s, "G84"; mergedCells=mc)
    @test XLSX.isMergedCell(s, "Short_Description")
    @test !XLSX.isMergedCell(f, "Mock-up!B2")
    @test !XLSX.isMergedCell(s, "H40"; mergedCells=mc)
    @test !XLSX.isMergedCell(s, "ID"; mergedCells=mc)
    @test_throws XLSX.XLSXError XLSX.isMergedCell(s, "Contiguous"; mergedCells=mc) # Can't test a range
    @test_throws XLSX.XLSXError XLSX.getMergedBaseCell(s, "Location")

    @test XLSX.getMergedBaseCell(f[1], "F72") == (baseCell=XLSX.CellRef("D72"), baseValue=Dates.Date("2025-03-24"))
    @test XLSX.getMergedBaseCell(f, "Mock-up!G72") == (baseCell=XLSX.CellRef("D72"), baseValue=Dates.Date("2025-03-24"))
    @test XLSX.getMergedBaseCell(s, "H53") == (baseCell=XLSX.CellRef("D51"), baseValue="Hello World")
    @test XLSX.getMergedBaseCell(s, "G52") == (baseCell=XLSX.CellRef("D51"), baseValue="Hello World")
    @test XLSX.getMergedBaseCell(s, 53, 8) == (baseCell=XLSX.CellRef("D51"), baseValue="Hello World")
    @test XLSX.getMergedBaseCell(s, "Short_Description") == (baseCell=XLSX.CellRef("D51"), baseValue="Hello World")
    @test isnothing(XLSX.getMergedBaseCell(s, "F73"))
    @test isnothing(XLSX.getMergedBaseCell(f, "Mock-up!H73"))
    @test_throws XLSX.XLSXError XLSX.getMergedBaseCell(s, "Location") # Can't get base cell for a range

    @test isnothing(XLSX.getMergedCells(f["Document History"]))
    s = f["Document History"]
    @test !XLSX.isMergedCell(f, "Document History!B2")
    @test !XLSX.isMergedCell(s, "C5"; mergedCells=XLSX.getMergedCells(f["Document History"]))

    f = XLSX.opentemplate(joinpath(data_directory, "testmerge.xlsx"))
    @test XLSX.mergeCells(f, "Sheet1!A1:B2") == 0
    @test f[1]["A1"] == "Tables"
    @test ismissing(f[1]["B2"])
    @test f[1]["C3"] == 4
    @test XLSX.mergeCells(f[1], 4:6, 4:6) == 0
    @test f[1][4, 4] == 9
    @test ismissing(f[1][5, 5])
    @test f[1][7, 7] == 36
    @test XLSX.mergeCells(f[1], "J") == 0
    @test f[1]["J1"] == 9
    @test ismissing(f[1]["J2"])
    @test ismissing(f[1]["J12"])
    @test XLSX.isMergedCell(f[1], "J8")
    mc = XLSX.getMergedCells(f["Sheet1"])
    @test XLSX.isMergedCell(f[1], "J9"; mergedCells=mc)
    @test XLSX.getMergedBaseCell(f[1], "J12") == (baseCell=XLSX.CellRef("J1"), baseValue=9)

    @test_throws XLSX.XLSXError XLSX.mergeCells(f[1], "Sheet1!M13:M13")       # Single cell
    @test_throws XLSX.XLSXError XLSX.mergeCells(f[1], 1, :)                   # Overlapping
    @test_throws XLSX.XLSXError XLSX.mergeCells(f[1], 10, :)                  # Overlapping
    @test_throws XLSX.XLSXError XLSX.mergeCells(f["Sheet1"], "M1:P15")        # Outside dimension
    @test_throws XLSX.XLSXError XLSX.mergeCells(f["Sheet1"], "Sheet2!L1:M2")  # Sheets don't match

    XLSX.writexlsx("outfile.xlsx", f, overwrite=true)
    SAVE_FILES && save_outfile("outfile.xlsx")

    XLSX.openxlsx("outfile.xlsx"; mode="rw") do f
        mc = sort(XLSX.getMergedCells(f["Sheet1"]))
        @test length(mc) == 3
        @test mc == sort(XLSX.CellRange[XLSX.CellRange("A1:B2"), XLSX.CellRange("D4:F6"), XLSX.CellRange("J1:J13")])
        @test XLSX.isMergedCell(f[1], "B2")
        @test XLSX.isMergedCell(f[1], 6, 6; mergedCells=mc)
        @test XLSX.getMergedBaseCell(f[1], "F6") == (baseCell=XLSX.CellRef("D4"), baseValue=9)
        @test f[1]["A1"] == "Tables"
        @test ismissing(f[1]["B2"])
        @test f[1]["C3"] == 4
        @test f[1][4, 4] == 9
        @test ismissing(f[1][5, 5])
        @test f[1][7, 7] == 36
        @test f[1]["J1"] == 9
        @test ismissing(f[1]["J2"])
        @test ismissing(f[1]["J12"])
        @test XLSX.isMergedCell(f[1], "J8")
        @test XLSX.isMergedCell(f[1], "J9"; mergedCells=XLSX.getMergedCells(f["Sheet1"]))
        @test XLSX.getMergedBaseCell(f[1], "J12") == (baseCell=XLSX.CellRef("J1"), baseValue=9)
    end
    isfile("outfile.xlsx") && rm("outfile.xlsx")

    f = XLSX.newxlsx()
    s = f[1]
    for i in 1:3, j in 1:3
        s[i, j] = i + j
    end
    XLSX.mergeCells(s, "Sheet1!A:B")
    @test XLSX.getMergedBaseCell(f, "Sheet1!B2") == (baseCell=XLSX.CellRef("A1"), baseValue=2)
    SAVE_FILES && save_outfile(f)

    f = XLSX.newxlsx()
    s = f[1]
    for i in 1:4, j in 1:4
        s[i, j] = i + j
    end
    XLSX.mergeCells(s, "Sheet1!2:3")
    @test XLSX.getMergedBaseCell(f, "Sheet1!C3") == (baseCell=XLSX.CellRef("A2"), baseValue=3)
    XLSX.mergeCells(s, "Sheet1!4:4")
    @test XLSX.getMergedBaseCell(f, "Sheet1!C4") == (baseCell=XLSX.CellRef("A4"), baseValue=5)
    SAVE_FILES && save_outfile(f)

    f = XLSX.newxlsx()
    s = f[1]
    for i in 1:3, j in 1:3
        s[i, j] = i + j
    end
    XLSX.mergeCells(s, :, 2:3)
    @test XLSX.getMergedBaseCell(f, "Sheet1!C3") == (baseCell=XLSX.CellRef("B1"), baseValue=3)
    SAVE_FILES && save_outfile(f)

    f = XLSX.newxlsx()
    s = f[1]
    for i in 1:3, j in 1:3
        s[i, j] = i + j
    end
    XLSX.mergeCells(s, :, :)
    @test XLSX.getMergedBaseCell(f, "Sheet1!B2") == (baseCell=XLSX.CellRef("A1"), baseValue=2)
    SAVE_FILES && save_outfile(f)

    f = XLSX.newxlsx()
    s = f[1]
    for i in 1:3, j in 1:3
        s[i, j] = i + j
    end
    XLSX.mergeCells(s, :)
    @test XLSX.getMergedBaseCell(f, "Sheet1!B2") == (baseCell=XLSX.CellRef("A1"), baseValue=2)
    SAVE_FILES && save_outfile(f)
end
@testset "unmerge cells" begin

    XLSX.openxlsx(joinpath(data_directory, "customXml.xlsx")) do f
        @test_throws XLSX.XLSXError XLSX.removeMergedCells(f["Mock-up"]) # File isn't writeable
    end

    f = XLSX.opentemplate(joinpath(data_directory, "customXml.xlsx"))
    s = f["Mock-up"]

    @test length(XLSX.getMergedCells(s)) == 25

    # Naming any cell of a merged range unmerges the whole range.
    @test XLSX.removeMergedCells(s, "F72") == 0
    @test !XLSX.isMergedCell(s, "D72")
    @test !XLSX.isMergedCell(s, "J72")
    @test length(XLSX.getMergedCells(s)) == 24

    # Values are untouched by unmerging.
    @test s["D72"] == Dates.Date("2025-03-24")

    # An enclosed range works too, and the sheetcell form reaches the same range.
    @test XLSX.removeMergedCells(f, "Mock-up!E11:F11") == 0
    @test isnothing(XLSX.getMergedBaseCell(s, "G11"))
    @test length(XLSX.getMergedCells(s)) == 23

    # Each merged range must enclose the argument in its own right, so three
    # single-row merges need three calls.
    @test XLSX.removeMergedCells(s, "D12") == 0
    @test XLSX.removeMergedCells(s, "D13") == 0
    @test isnothing(XLSX.getMergedBaseCell(s, "G12"))
    @test isnothing(XLSX.getMergedBaseCell(s, "G13"))
    @test length(XLSX.getMergedCells(s)) == 21

    # A range exactly equal to the merge is enclosed by it.
    @test XLSX.removeMergedCells(s, "D14:E14,D16:E16") == -1 # process_ncranges return
    @test !XLSX.isMergedCell(s, "E14")
    @test !XLSX.isMergedCell(s, "E16")
    @test length(XLSX.getMergedCells(s)) == 19

    # Index forms.
    @test XLSX.removeMergedCells(s, 18, 4) == 0             # cell D18, inside D18:E18
    @test !XLSX.isMergedCell(s, "E18")
    @test XLSX.removeMergedCells(s, 20, 4:5) == 0           # D20:E20 exactly
    @test !XLSX.isMergedCell(s, "E20")
    @test length(XLSX.getMergedCells(s)) == 17

    # A large merged range encloses a smaller range argument.
    @test XLSX.isMergedCell(s, "H57")
    @test XLSX.removeMergedCells(s, "E56:F58") == 0         # inside D55:J60
    @test !XLSX.isMergedCell(s, "H57")

    # A range wider than the merge encloses nothing: no-op, not an error.
    n = length(XLSX.getMergedCells(s))
    @test XLSX.isMergedCell(s, "F47")                       # D47:H47
    @test XLSX.removeMergedCells(s, "D47:J47") == 0
    @test XLSX.isMergedCell(s, "F47")                       # still merged
    @test length(XLSX.getMergedCells(s)) == n

    # Unmerging a cell that isn't merged is also a no-op.
    @test XLSX.removeMergedCells(s, "B2") == 0
    @test length(XLSX.getMergedCells(s)) == n

    # Out of dimension / mismatched sheet still throw.
    @test_throws XLSX.XLSXError XLSX.removeMergedCells(s, "ZZ1:ZZ2")
    @test_throws XLSX.XLSXError XLSX.removeMergedCells(s, "Document History!D9:H9")

    # Remove all: the <mergeCells> block is dropped entirely.
    @test XLSX.removeMergedCells(s) == 0
    @test isnothing(XLSX.getMergedCells(s))
    @test !XLSX.isMergedCell(s, "D9")

    # ...and removing all again is a no-op.
    @test XLSX.removeMergedCells(s) == 0
    @test XLSX.removeMergedCells(s, :) == 0

    # A sheet that never had merged cells.
    @test isnothing(XLSX.getMergedCells(f["Document History"]))
    @test XLSX.removeMergedCells(f["Document History"]) == 0

    # Round trip: re-merge after unmerging, and check the file still opens.
    @test XLSX.mergeCells(s, "D9:H9") == 0
    @test XLSX.isMergedCell(s, "F9")

    XLSX.writexlsx("unmerge_outfile.xlsx", f, overwrite=true)
    @test isfile("unmerge_outfile.xlsx")
    g = XLSX.opentemplate("unmerge_outfile.xlsx")
    @test XLSX.getMergedCells(g["Mock-up"]) == XLSX.CellRange[XLSX.CellRange("D9:H9")]
    isfile("unmerge_outfile.xlsx") && rm("unmerge_outfile.xlsx")
    
    # Values written into a merged range (invisible in Excel) survive unmerging
    # and become visible again.
    f2 = XLSX.newxlsx()
    s2 = f2[1]
    s2["A1:A3"] = 5
    @test XLSX.mergeCells(s2, "A1:A3") == 0
    s2["A2"] = "text here now"
    @test s2["A2"] == "text here now"
    @test ismissing(s2["A3"])
    @test XLSX.removeMergedCells(s2, "A2") == 0
    @test isnothing(XLSX.getMergedCells(s2))
    @test s2["A1"] == 5
    @test s2["A2"] == "text here now"   # no longer hidden
    @test ismissing(s2["A3"])           # blanked by the merge, not restored
end
