@testset "Strict formats" begin
    @testset "Simple" begin
        XLSX.openxlsx(joinpath(data_directory, "strict.xlsx")) do f
            show(IOBuffer(), f)
            sheet = f["general"]
            @test sheet["A1"] == "text"
            @test sheet["B1"] == "regular text"
            @test sheet["A2"] == "integer"
            @test sheet["B2"] == 102
            @test sheet["A3"] == "float"
            @test isapprox(sheet["B3"], 102.2)
            @test sheet["A4"] == "date"
            @test sheet["B4"] == Date(1983, 4, 16)
            @test sheet["A5"] == "hour"
            @test sheet["B5"] == Dates.Time(Dates.Hour(19), Dates.Minute(45))
            @test sheet["A6"] == "datetime"
            @test sheet["B6"] == Date(2018, 4, 16) + Dates.Time(Dates.Hour(19), Dates.Minute(19), Dates.Second(51))
            @test f["general!B7"] == -220.0
            @test f["general!B8"] == -2000
            @test f["general!B9"] == 100000000000000
            @test f["general!B10"] == -100000000000000
        end

        XLSX.openxlsx(joinpath(data_directory, "strict.xlsx")) do xf
            empty_sheet = XLSX.getsheet(xf, "empty")
            @test_throws XLSX.XLSXError XLSX.gettable(empty_sheet)
            itr = XLSX.eachrow(empty_sheet)
            @test_throws XLSX.XLSXError XLSX.find_row(itr, 1)
            @test_throws XLSX.XLSXError XLSX.getsheet(xf, "invalid_sheet")
        end

        f = XLSX.readxlsx(joinpath(data_directory, "strict.xlsx"))
        tb5 = f["table5"]
        test_data = Vector{Any}(undef, 1)
        test_data[1] = [1, 2, 3, 4, 5]
        dtable = XLSX.gettable(tb5)
        data, col_names = dtable.data, dtable.column_labels
        @test col_names == [:HEADER]
        check_test_data(data, test_data)
        tb6 = f["table6"]
        dtable = XLSX.gettable(tb6, first_row=3)
        data, col_names = dtable.data, dtable.column_labels
        @test col_names == [:HEADER]
        check_test_data(data, test_data)
        tb7 = f["table7"]
        dtable = XLSX.gettable(tb7, first_row=3)
        data, col_names = dtable.data, dtable.column_labels
        @test col_names == [:HEADER]
        check_test_data(data, test_data)

        sheet_lookup = f["lookup"]
        test_data = Vector{Any}(undef, 3)
        test_data[1] = [10, 20, 30]
        test_data[2] = ["name1", "name2", "name3"]
        test_data[3] = [100, 200, 300]
        dtable = XLSX.gettable(sheet_lookup)
        data, col_names = dtable.data, dtable.column_labels
        @test col_names == [:ID, :NAME, :VALUE]
        check_test_data(data, test_data)

        header_error_sheet = f["header_error"]
        dtable = XLSX.gettable(header_error_sheet, "B:E")
        data, col_names = dtable.data, dtable.column_labels
        @test col_names == [:COLUMN_A, :COLUMN_B, Symbol("COLUMN_A_2"), Symbol("#Empty")]

        XLSX.writexlsx("mytest.xlsx", XLSX.openxlsx(joinpath(data_directory, "strict.xlsx"); mode="rw"), overwrite=true)
        SAVE_FILES && save_outfile("mytest.xlsx")

        XLSX.openxlsx("mytest.xlsx") do f
            show(IOBuffer(), f)
            sheet = f["general"]
            @test sheet["A1"] == "text"
            @test sheet["B1"] == "regular text"
            @test sheet["A2"] == "integer"
            @test sheet["B2"] == 102
            @test sheet["A3"] == "float"
            @test isapprox(sheet["B3"], 102.2)
            @test sheet["A4"] == "date"
            @test sheet["B4"] == Date(1983, 4, 16)
            @test sheet["A5"] == "hour"
            @test sheet["B5"] == Dates.Time(Dates.Hour(19), Dates.Minute(45))
            @test sheet["A6"] == "datetime"
            @test sheet["B6"] == Date(2018, 4, 16) + Dates.Time(Dates.Hour(19), Dates.Minute(19), Dates.Second(51))
            @test f["general!B7"] == -220.0
            @test f["general!B8"] == -2000
            @test f["general!B9"] == 100000000000000
            @test f["general!B10"] == -100000000000000
        end

        XLSX.openxlsx("mytest.xlsx") do xf
            empty_sheet = XLSX.getsheet(xf, "empty")
            @test_throws XLSX.XLSXError XLSX.gettable(empty_sheet)
            itr = XLSX.eachrow(empty_sheet)
            @test_throws XLSX.XLSXError XLSX.find_row(itr, 1)
            @test_throws XLSX.XLSXError XLSX.getsheet(xf, "invalid_sheet")
        end

        f = XLSX.readxlsx("mytest.xlsx")
        tb5 = f["table5"]
        test_data = Vector{Any}(undef, 1)
        test_data[1] = [1, 2, 3, 4, 5]
        dtable = XLSX.gettable(tb5)
        data, col_names = dtable.data, dtable.column_labels
        @test col_names == [:HEADER]
        check_test_data(data, test_data)
        tb6 = f["table6"]
        dtable = XLSX.gettable(tb6, first_row=3)
        data, col_names = dtable.data, dtable.column_labels
        @test col_names == [:HEADER]
        check_test_data(data, test_data)
        tb7 = f["table7"]
        dtable = XLSX.gettable(tb7, first_row=3)
        data, col_names = dtable.data, dtable.column_labels
        @test col_names == [:HEADER]
        check_test_data(data, test_data)

        sheet_lookup = f["lookup"]
        test_data = Vector{Any}(undef, 3)
        test_data[1] = [10, 20, 30]
        test_data[2] = ["name1", "name2", "name3"]
        test_data[3] = [100, 200, 300]
        dtable = XLSX.gettable(sheet_lookup)
        data, col_names = dtable.data, dtable.column_labels
        @test col_names == [:ID, :NAME, :VALUE]
        check_test_data(data, test_data)

        header_error_sheet = f["header_error"]
        dtable = XLSX.gettable(header_error_sheet, "B:E")
        data, col_names = dtable.data, dtable.column_labels
        @test col_names == [:COLUMN_A, :COLUMN_B, Symbol("COLUMN_A_2"), Symbol("#Empty")]

        isfile("mytest.xlsx") && rm("mytest.xlsx")
    end

    @testset "With chartsheet" begin # From Issue #233
        f = XLSX.openxlsx(joinpath(data_directory, "Strict-foo.xlsx"), mode="rw")
        Expected = """            sheetname size          range        
            -------------------------------------------------
                         Tabelle1 13x4          A2:D14       
                        Diagramm1 Chartsheet\n"""
        result = sprint(show, f)
        idx = findfirst(==('\n'), result)
        after = result[idx+1:end]
        @test after == Expected
        @test sprint(show, f[1]) == "13×4 XLSX.Worksheet: [\"Tabelle1\"](A2:D14) "
        @test sprint(show, f[2]) == "Chartsheet: [\"Diagramm1\"] "
        @test_throws XLSX.XLSXError XLSX.copysheet!(f["Diagramm1"], "Diagramm1_copy")
        @test_throws XLSX.XLSXError XLSX.deletesheet!(f["Diagramm1"])
        @test_throws XLSX.XLSXError XLSX.gettable(f["Diagramm1"])
        @test_throws XLSX.XLSXError XLSX.gettable(f["Diagramm1"], "A:B")
        XLSX.writexlsx("mytest.xlsx", f, overwrite=true)
        SAVE_FILES && save_outfile("mytest.xlsx")

        XLSX.openxlsx("mytest.xlsx") do f
            result = sprint(show, f)
            idx = findfirst(==('\n'), result)
            after = result[idx+1:end]
            @test after == Expected
            @test sprint(show, f[1]) == "13×4 XLSX.Worksheet: [\"Tabelle1\"](A2:D14) "
            @test sprint(show, f[2]) == "Chartsheet: [\"Diagramm1\"] "
        end
        isfile("mytest.xlsx") && rm("mytest.xlsx")
    end

end