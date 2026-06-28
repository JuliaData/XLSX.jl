@testset "no default namespace" begin
    # issues #380, #362, #267, #170
    
    f = XLSX.openxlsx(joinpath(data_directory, "No-Default_NameSpace.xlsx"), mode="rw")
    @test XLSX.get_dimension(f[2])==XLSX.CellRange("A1:N15")
    XLSX.addDefinedName(f, "xfile", "XLSX-Export!B2")
    XLSX.addDefinedName(f[2], "wsheet", 200.2)
    XLSX.mergeCells(f[2], "A11:N11")
    XLSX.setRowHeight(f[2], 11; height=25)
    XLSX.setColumnWidth(f[2], "B"; width=50)
    @test XLSX.setFont(f[2], [2, 4], 2:4; size=18, name="Arial") == -1
    @test XLSX.setBorder(f[2], [2, 4], :; left=["style" => "hair"], right=["color" => "FF8B4513"], top=["style" => "hair"], bottom=["color" => "chocolate4"], diagonal=["style" => "hair"]) == -1
    @test XLSX.setFill(f[2], "G5"; pattern="darkTrellis", fgColor="FF222222", bgColor="FFDDDDDD") == 2
    @test XLSX.setConditionalFormat(f[2], "J2:J10", :colorScale) == 0
    XLSX.mergeCells(f[2], "A4:C5")
    @test XLSX.setConditionalFormat(f[2], "H2:H15", :dataBar) == 0
    XLSX.setFormula(f[2], "A16:N16", "=sum(A2:A15)")
    XLSX.addImage(f[2], "B17", joinpath(data_directory, "track_start.jpg"))
    XLSX.copysheet!(f[2], "newSheet")

    XLSX.writexlsx("mytest.xlsx", f, overwrite=true)
    SAVE_FILES && save_outfile("mytest.xlsx")

    f2 = XLSX.openxlsx("mytest.xlsx", mode="rw")
    @test XLSX.get_dimension(f2[2])==XLSX.CellRange("A1:N16")
    @test f2["xfile"] == f2["XLSX-Export"]["B2"]
    @test f2["XLSX-Export"]["wsheet"] == 200.2
    @test XLSX.getMergedCells(f2[2]) == XLSX.CellRange[XLSX.CellRange("A11:N11"), XLSX.CellRange("A4:C5")]
    @test XLSX.isMergedCell(f2[2], "D11")
    @test XLSX.getMergedBaseCell(f2[2], "D11") == (baseCell=XLSX.CellRef("A11"), baseValue=13953)
    @test ismissing(f2[2]["D11"])
    @test XLSX.getRowHeight(f2[2], "B11") ≈ 25.2109375
    @test XLSX.getColumnWidth(f2[2], 11, 2) ≈ 50.7109375
    @test XLSX.getFont(f2[2], 2, 2).font == Dict("name" => Dict("val" => "Arial"), "sz" => Dict("val" => "18"), "color" => Dict("theme" => "1"))
    @test XLSX.getBorder(f2[2], 4, 6).border == Dict("left" => Dict("style" => "hair"), "bottom" => nothing, "right" => nothing, "top" => Dict("style" => "hair"), "diagonal" => Dict("style" => "hair", "direction" => "both"))
    @test XLSX.getFill(f2[2], "G5").fill == Dict("patternFill" => Dict("bgrgb" => "FFDDDDDD", "patternType" => "darkTrellis", "fgrgb" => "FF222222"))
    @test sort!(XLSX.getConditionalFormats(f2[2]), by = x -> x.second.priority, rev=true) == [
        XLSX.CellRange("H2:H15") => (type = "dataBar", priority = 2),
        XLSX.CellRange("J2:J10") => (type = "colorScale", priority = 1)
    ]
    @test XLSX.getFormula(f2[2], "B16") == "=sum(B2:B15)"
    @test XLSX.getFormula(f2[2], "M16") == "=sum(M2:M15)"

    @test XLSX.hassheet(f, "newSheet")
    @test XLSX.getImages(f["newSheet"]) == [(sheet = "newSheet", media_name = "image1.jpg", from = "B17", to = "E27")]

    for row = 1:16
        for col = 1:14
            @test (ismissing(f[1][row, col]) && ismissing(f2[1][row, col])) || (f2[1][row, col] == f[1][row, col])
            @test (ismissing(f[2][row, col]) && ismissing(f[3][row, col])) || (f[2][row, col] == f[3][row, col])
            @test (ismissing(f[3][row, col]) && ismissing(f2[3][row, col])) || (f2[3][row, col] == f[3][row, col])
        end
    end

    df = XLSX.readto(joinpath(data_directory, "No-Default_NameSpace.xlsx"), 2, DataFrames.DataFrame)
    @test DataFrames.names(df) == String[
        "NR ",
        "LC ",
        "LC-title ",
        "X [m]",
        "Xi ",
        "MNR ",
        "SIGU [MPa]",
        "DL [m]",
        "node1 ",
        "X [m]_2",
        "Y [m]",
        "Z [m]",
        "node2 ",
        "NREF "
    ]
    @test DataFrames.nrow(df) == 14

    XLSX.writetable("mytest.xlsx", "Sheet1" => df; overwrite=true)
    SAVE_FILES && save_outfile("mytest.xlsx")
    df2 = XLSX.readto("mytest.xlsx", DataFrames.DataFrame)

    @test DataFrames.names(df) == DataFrames.names(df2)
    @test DataFrames.nrow(df) == DataFrames.nrow(df2)
    @test isequal(df, df2)


    f = XLSX.openxlsx(joinpath(data_directory, "No-Default_NameSpace2.xlsx"), mode="rw")
    @test XLSX.get_dimension(f[1])==XLSX.CellRange("A1:R1001")

    XLSX.writexlsx("mytest.xlsx", f, overwrite=true)
    SAVE_FILES && save_outfile("mytest.xlsx")
    f2 = XLSX.readxlsx("mytest.xlsx")
    @test XLSX.get_dimension(f2[1])==XLSX.CellRange("A1:R1001")

    for row =1:100:1001
        for col = 1:2:17
            @test (ismissing(f[1][row, col]) && ismissing(f2[1][row, col])) || (f2[1][row, col] == f[1][row, col])
        end
    end

    df = XLSX.readto(joinpath(data_directory, "No-Default_NameSpace2.xlsx"), DataFrames.DataFrame; first_row=3)
    @test DataFrames.names(df) == String[
        "Stock Code",
        "Name of Securities",
        "Category",
        "Sub-Category",
        "Board Lot",
        "ISIN",
        "Expiry Date",
        "Subject to Stamp Duty",
        "Shortsell Eligible",
        "CAS Eligible",
        "VCM Eligible",
        "Admitted to CCASS",
        "Debt Securities Board Lot (Nominal)",
        "Debt Securities Investor Type",
        "POS Eligible",
        "Spread Table\r\n1 = Part A\r\n3 = Part B\r\n5 = Part D\r\n4 & 6 = Part E",
        "Trading Currency",
        "RMB Counter"
        ]
    @test DataFrames.nrow(df) == 998

    XLSX.writetable("mytest.xlsx", "Sheet1" => df; overwrite=true)
    SAVE_FILES && save_outfile("mytest.xlsx")
    df2 = XLSX.readto("mytest.xlsx", DataFrames.DataFrame)

    @test DataFrames.names(df) == DataFrames.names(df2)
    @test DataFrames.nrow(df) == DataFrames.nrow(df2)

    for row in 1:5:DataFrames.nrow(df)
        for col in 1:2:DataFrames.ncol(df)
            if !(ismissing(df2[row, col]))
                @test df[row, col] == df2[row, col]
            end
        end
    end
    isfile("mytest.xlsx") && rm("mytest.xlsx")
end