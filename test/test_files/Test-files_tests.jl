@testset "read test files" begin
    ef_blank_ptbr_1904 = XLSX.readxlsx(joinpath(data_directory, "blank_ptbr_1904.xlsx"))
    ef_Book1 = XLSX.readxlsx(joinpath(data_directory, "Book1.xlsx"))
    ef_Book_1904 = XLSX.readxlsx(joinpath(data_directory, "Book_1904.xlsx"))
    ef_book_1904_ptbr = XLSX.readxlsx(joinpath(data_directory, "book_1904_ptbr.xlsx"))
    ef_book_sparse = XLSX.readxlsx(joinpath(data_directory, "book_sparse.xlsx"))
    ef_book_sparse_2 = XLSX.readxlsx(joinpath(data_directory, "book_sparse_2.xlsx"))

    XLSX.readxlsx(joinpath(data_directory, "missing_numFmtId.xlsx"))["Koldioxid (CO2)"][7, 5]

    @test open(joinpath(data_directory, "blank_ptbr_1904.xlsx")) do io
        XLSX.readxlsx(io)
    end isa XLSX.XLSXFile

    @test ef_Book1.source == joinpath(data_directory, "Book1.xlsx")
    @test length(keys(ef_Book1.data)) > 0

    @test ef_Book_1904.source == joinpath(data_directory, "Book_1904.xlsx")
    @test length(keys(ef_Book_1904.data)) > 0

    @test !XLSX.isdate1904(ef_Book1)
    @test XLSX.isdate1904(ef_Book_1904)
    @test XLSX.isdate1904(ef_blank_ptbr_1904)
    @test XLSX.isdate1904(ef_book_1904_ptbr)

    @test XLSX.sheetnames(ef_Book1) == ["Sheet1", "Sheet2"]
    @test XLSX.sheetcount(ef_Book1) == 2
    @test ef_Book1["Sheet1"].name == "Sheet1"
    @test ef_Book1[1].name == "Sheet1"

    @test XLSX.sst_unformatted_string(ef_Book1.workbook, Int64(0)) == "B2" # index is 0-based
    @test XLSX.sst_unformatted_string(ef_Book1, Int64(0)) == "B2"
    @test XLSX.sst_unformatted_string(ef_Book1, "0") == "B2"

    @test !XLSX.has_relationship_by_type(ef_Book1.workbook, "invalid_type")

    @test XLSX.get_dimension(ef_Book1["Sheet1"]) == XLSX.range"B2:C8"
    @test XLSX.isdate1904(ef_Book1["Sheet1"]) == false

    @testset "Read XLS file error" begin
        @test_throws XLSX.XLSXError XLSX.readxlsx(joinpath(data_directory, "old.xls"))
        try
            XLSX.readxlsx(joinpath(data_directory, "old.xls"))
            @test false # didn't throw exception
        catch e
            @test occursin("This package does not support XLS file format", "$e")
        end

    end

    # Issue #293
    @testset "read .xltx file" begin
        xf = XLSX.openxlsx(joinpath(data_directory, "Template File.xltx"); mode="RW")
        s=xf[1]
        @test s["P5"] == 5
        @test XLSX.getFormula(s, "B5") == "=RANDBETWEEN(0,100)"
        @test xf.template_type == XLSX.XLTXTemplate

        XLSX.savexlsx(xf)
        SAVE_FILES && save_outfile(joinpath(data_directory, "Template File.xlsx"))
        @test isfile(joinpath(data_directory, "Template File.xlsx"))
        xf = XLSX.readxlsx(joinpath(data_directory, "Template File.xlsx"))
        s=xf[1]
        @test s["P5"] == 5
        @test XLSX.getFormula(s, "B5") == "=RANDBETWEEN(0,100)"
        @test xf.template_type == XLSX.NotATemplate
        isfile(joinpath(data_directory, "Template File.xlsx")) && rm(joinpath(data_directory, "Template File.xlsx"))

        XLSX.openxlsx(joinpath(data_directory, "Template File.xltx"); mode="RW") do xf
            s=xf[1]
            @test s["P5"] == 5
            @test XLSX.getFormula(s, "B5") == "=RANDBETWEEN(0,100)"
            @test xf.template_type == XLSX.XLTXTemplate
        end
        @test isfile(joinpath(data_directory, "Template File.xlsx"))
        SAVE_FILES && save_outfile(joinpath(data_directory, "Template File.xlsx"))
        isfile(joinpath(data_directory, "Template File.xlsx")) && rm(joinpath(data_directory, "Template File.xlsx"))
    end

    # Issue #401
    @testset "macro enabled files" begin
        mf = XLSX.openxlsx(joinpath(data_directory, "macro-enabled.xlsm"); mode="Rw")
        @test mf[1]["A1"] == "hello"
        XLSX.writexlsx("mytest.xlsm", mf; overwrite=true)
        SAVE_FILES && save_outfile("mytest.xlsm")
        mf = XLSX.openxlsx("mytest.xlsm"; mode="rw")
        @test mf[1]["A1"] == "hello"
        isfile("mytest.xlsm") && rm("mytest.xlsm")

        mf = XLSX.openxlsx(joinpath(data_directory, "macro-enabled2.xltm"); mode="rW")
        @test mf[1]["A1"] == "hello"
        @test mf.template_type == XLSX.XLTMTemplate
        XLSX.savexlsx(mf)
        SAVE_FILES && save_outfile(joinpath(data_directory, "macro-enabled2.xlsm"))
        @test isfile(joinpath(data_directory, "macro-enabled2.xlsm"))
        mf = XLSX.readxlsx(joinpath(data_directory, "macro-enabled2.xlsm"))
        s=mf[1]
        @test mf[1]["A1"] == "hello"
        @test mf.template_type == XLSX.NotATemplate
        isfile(joinpath(data_directory, "macro-enabled2.xlsm")) && rm(joinpath(data_directory, "macro-enabled2.xlsm"))
        
        XLSX.openxlsx(joinpath(data_directory, "macro-enabled2.xltm"); mode="wr") do mf
            @test mf[1]["A1"] == "hello"
            @test mf.template_type == XLSX.XLTMTemplate
        end
        @test isfile(joinpath(data_directory, "macro-enabled2.xlsm"))
        SAVE_FILES && save_outfile(joinpath(data_directory, "macro-enabled2.xlsm"))
        mf = XLSX.openxlsx(joinpath(data_directory, "macro-enabled2.xlsm"); mode="WR")
        @test mf[1]["A1"] == "hello"
        @test mf.template_type == XLSX.NotATemplate
        isfile(joinpath(data_directory, "macro-enabled2.xlsm")) && rm(joinpath(data_directory, "macro-enabled2.xlsm"))

    end

    # Issue #403
    @testset "UTF-16 customXml" begin
        try
            xf1 = XLSX.openxlsx(joinpath(data_directory, "UTF-16.xlsx"); mode="rw")
            XLSX.writexlsx("UTF-16_test.xlsx", xf1; overwrite=true)
            SAVE_FILES && save_outfile("UTF-16_test.xlsx")
            xf2 = XLSX.openxlsx("UTF-16_test.xlsx"; mode="rw")
            @test xf1[1]["E3"] == xf2[1]["E3"]
            @test xf1[1]["R99"] == xf2[1]["R99"]
        catch e
            @test false
        end

        isfile("UTF-16_test.xlsx") && rm("UTF-16_test.xlsx")

    end
    @testset "Read password protected file error" begin
        @test_throws XLSX.XLSXError XLSX.readxlsx(joinpath(data_directory, "password.xlsx")) # password for this file is simply "password"
        try
            XLSX.readxlsx(joinpath(data_directory, "password.xlsx"))
            @test false # didn't throw exception
        catch e
            @test occursin("This package does not support password protected files", "$e")
        end

    end

    @testset "Read invalid XLSX error" begin
        @test_throws XLSX.XLSXError XLSX.readxlsx(joinpath(src_data_directory, "sheet_template.xml"))
        try
            XLSX.readxlsx(joinpath(src_data_directory, "sheet_template.xml"))
            @test false # didn't throw exception
        catch e
            @test occursin("is not a valid XLSX file", "$e")
        end
    end

    @testset "missing file or bad `mode`" begin
        @test_throws XLSX.XLSXError XLSX.openxlsx("noSuchFile.xlsx")
        @test_throws XLSX.XLSXError XLSX.openxlsx(joinpath(data_directory, "Book1.xlsx"); mode="tg")
    end

    @testset "write-only mode" begin
        XLSX.openxlsx("mytest.xlsx", mode="w") do f
            f[1]["A1"] = 1
            @test f.source == "mytest.xlsx"
        end
        SAVE_FILES && save_outfile("mytest.xlsx")
        ef = XLSX.readxlsx("mytest.xlsx")
        @test ef["Sheet1"]["A1"] == 1
        f = XLSX.openxlsx("mytest2.xlsx", mode="w")
        @test f.source == "mytest2.xlsx"
        f[1]["A1"] = 1
        XLSX.writexlsx("mytest3.xlsx", f, overwrite=true)
        SAVE_FILES && save_outfile("mytest3.xlsx")
        ef = XLSX.readxlsx("mytest3.xlsx")
        @test ef["Sheet1"]["A1"] == 1
        f = XLSX.newxlsx()
        @test f.source == "blank.xlsx"
        f[1]["A1"] = 1
        XLSX.writexlsx("mytest.xlsx", f, overwrite=true)
        SAVE_FILES && save_outfile("mytest.xlsx")
        ef = XLSX.readxlsx("mytest.xlsx")
        @test ef["Sheet1"]["A1"] == 1
        for f in ["mytest.xlsx", "mytest2.xlsx", "mytest3.xlsx"]
            isfile(f) && rm(f)
        end
    end

    @testset "Fix timestamp" begin
        t = Dates.now(Dates.UTC) - Dates.Second(1)
        xf = XLSX.newxlsx()
        f = "docProps/core.xml"
        date_format = Dates.dateformat"yyyy-mm-ddTHH:MM:SSZ"
        i, j = XLSX.get_idces(xf.data[f], "cp:coreProperties", "dcterms:created") # i, j should always be found unless the `blank.xlsx` file is updated
        @test t < DateTime(XML.value(xf.data[f][i][j][1]), date_format)
        i, j = XLSX.get_idces(xf.data[f], "cp:coreProperties", "dcterms:modified")
        @test t < DateTime(XML.value(xf.data[f][i][j][1]), date_format)
        SAVE_FILES && save_outfile(xf)

        xf = XLSX.newxlsx(; update_timestamp=false) # do not update timestamp
        i, j = XLSX.get_idces(xf.data[f], "cp:coreProperties", "dcterms:created") # i, j should always be found unless the `blank.xlsx` file is updated
        @test XML.value(xf.data[f][i][j][1]) == "2018-05-22T02:41:32Z"
        i, j = XLSX.get_idces(xf.data[f], "cp:coreProperties", "dcterms:modified")
        @test XML.value(xf.data[f][i][j][1]) == "2018-05-22T02:42:04Z"
        SAVE_FILES && save_outfile(xf)
    end

    @testset "Book1.xlsx" begin
        f = XLSX.readxlsx(joinpath(data_directory, "Book1.xlsx"))
        sheet = f["Sheet1"]
        @test ismissing(sheet["A1"])
        @test sheet["B2"] == "B2"
        @test sheet["C2"] == "C2"
        @test isapprox(sheet["B3"], 10.5)
        @test isapprox(sheet["C3"], 21.2)
        @test sheet["B4"] == Date(2018, 3, 21)
        @test sheet["C4"] == Date(2018, 3, 22)
        @test sheet["B5"] == Date(2018, 3, 21)
        @test sheet["C5"] == Date(2018, 3, 22)
        @test sheet["B6"] == true
        @test sheet["C6"] == false
        @test sheet["B7"] == 1
        @test sheet["C7"] == 2
        @test sheet["B8"] == "palavra1"
        @test sheet["C8"] == "palavra2"
        @test XLSX.get_dimension(sheet) == XLSX.CellRange("B2:C8")

        sheet2 = f["Sheet2"]
        @test XLSX.get_dimension(sheet2) == XLSX.CellRange("A1:C3")
        @test axes(sheet2, 1) == 1:3
        @test axes(sheet2, 2) == 1:3
        @test_throws ArgumentError axes(sheet2, 3)
        @test sheet2[1, :] == Any[1 2 3]
        @test sheet2[1:2, :] == Any[1 2 3; 4 5 6]
        @test sheet2[:, 2] == permutedims(Any[2 5 8])
        @test sheet2[:, 2:3] == Any[2 3; 5 6; 8 9]
        @test sheet2[1:2, 2:3] == Any[2 3; 5 6]


        @test XLSX.getdata(f, XLSX.SheetCellRef("Sheet1!B2")) == "B2"
        @test XLSX.getdata(f, XLSX.SheetCellRange("Sheet1!B2:B3"))[1] == "B2"
        @test XLSX.getdata(f, XLSX.SheetCellRange("Sheet1!B2:B3"))[2] == 10.5
        @test f["Sheet1!B2"] == "B2"
        @test f["Sheet1!B2:B3"][1] == "B2"
        @test f["Sheet1!B2:B3"][2] == 10.5
        @test string(XLSX.SheetCellRange("Sheet1!B2:B3")) == "Sheet1!B2:B3"
    end

    @testset "book_1904_ptbr.xlsx" begin
        f = XLSX.readxlsx(joinpath(data_directory, "book_1904_ptbr.xlsx"))

        @test f["Plan1"][:] == Any["Coluna A" "Coluna B" "Coluna C" "Coluna D";
            10 10.5 Date(2018, 3, 22) "linha 2";
            20 20.5 Date(2017, 12, 31) "linha 3";
            30 30.5 Date(2018, 1, 1) "linha 4"]

        @test f["Plan2"]["A1"] == "Merge de A1:D1"
        @test ismissing(f["Plan2"]["B1"])
        @test f["Plan2"]["C2"] == "C2"
        @test f["Plan2"]["D3"] == "D3"
        @test f["NEGOCIAÇÕES Descrição"]["A1"] == "Negociações"
        @test f["NEGOCIAÇÕES Descrição"]["B1"] == 10
        @test f["NEGOCIAÇÕES Descrição!A1"] == "Negociações"
        @test f["NEGOCIAÇÕES Descrição!B1"] == 10
    end

    @testset "numbers.xlsx" begin
        f = XLSX.readxlsx(joinpath(data_directory, "numbers.xlsx"))
        floats = f["float"][:]
        for n in floats
            if !ismissing(n)
                @test isa(n, Float64)
            end
        end

        ints = f["int"][:]
        for n in ints
            if !ismissing(n)
                @test isa(n, Int64)
            end
        end

        error_sheet = f["error"]
        @test error_sheet["A1"] == "errors"
        @test !XLSX.iserror(XLSX.getcell(error_sheet, "A1"))
        @test XLSX.iserror(XLSX.getcell(error_sheet, "A2"))
        @test XLSX.iserror(XLSX.getcell(f, "error!A2"))
        @test ismissing(error_sheet["A2"])
        @test ismissing(error_sheet["A3"])
        @test ismissing(error_sheet["A4"])
        emptycell = XLSX.getcell(error_sheet, "B1")
        @test !XLSX.iserror(emptycell)
        @test ismissing(XLSX.getdata(error_sheet, emptycell))
        @test XLSX.row_number(emptycell) == 1
        @test XLSX.column_number(emptycell) == 2
    end

end

@testset "CustomXml" begin
    # issue #210
    template = XLSX.open_xlsx_template(joinpath(data_directory, "customXml.xlsx"))
    filename_copy = "customXml_copy.xlsx"
    for sn in XLSX.sheetnames(template)
        sheet = template[sn]
        sheet["Q1"] = "Can't"
        sheet["Q2"] = "write"
        sheet["Q3"] = "this"
        sheet["Q4"] = "template"
    end
    @test XLSX.writexlsx(filename_copy, template, overwrite=true) == abspath(filename_copy)
    SAVE_FILES && save_outfile(filename_copy)
    @test isfile(filename_copy)
    f_copy = XLSX.readxlsx(filename_copy)
    test_Xmlread = [["Can't", "write", "this", "template"]]
    for sn in XLSX.sheetnames(f_copy)
        sheet = template[sn]
        data = [[sheet["Q1"], sheet["Q2"], sheet["Q3"], sheet["Q4"]]]
        check_test_data(data, test_Xmlread)
    end
    isfile(filename_copy) && rm(filename_copy)
end

@testset "docProps/app.xml" begin # issue #428

    function parse_heading_pairs(app_xml::String)
        m = match(r"<HeadingPairs>(.*?)</HeadingPairs>"s, app_xml)
        m === nothing && return Tuple{String,Int}[]
        inner = m.captures[1]
        labels = [c.captures[1] for c in eachmatch(r"<vt:lpstr>([^<]*)</vt:lpstr>", inner)]
        counts = [parse(Int, c.captures[1]) for c in eachmatch(r"<vt:i4>([^<]*)</vt:i4>", inner)]
        return collect(zip(labels, counts))
    end

    function parse_titles_of_parts(app_xml::String)
        m = match(r"<TitlesOfParts>(.*?)</TitlesOfParts>"s, app_xml)
        m === nothing && return String[]
        inner = m.captures[1]
        return [c.captures[1] for c in eachmatch(r"<vt:lpstr>([^<]*)</vt:lpstr>", inner)]
    end

    @testset "plain worksheets: rename + add updates HeadingPairs and TitlesOfParts" begin
        src = joinpath(data_directory, "Book1.xlsx")
        p = tempname() * ".xlsx"
        cp(src, p)
 
        XLSX.openxlsx(p; mode="rw") do xf
            XLSX.renamesheet!(xf["Sheet1"], "Renamed")
            XLSX.addsheet!(xf, "Added")
        end

        app = String(zip_readentry(ZipReader(read(p)), "docProps/app.xml"))

        @test !occursin(">Sheet1<", app)   # stale title gone

        titles = parse_titles_of_parts(app)
        @test "Renamed" in titles
        @test "Added" in titles
        @test length(titles) == 3

        pairs = parse_heading_pairs(app)
        ws_idx = findfirst(pr -> pr[1] == "Worksheets", pairs)
        @test ws_idx !== nothing
        @test pairs[ws_idx][2] == 3

        rm(p; force=true)
    end

    @testset "chartsheet workbook: Worksheets/Charts categories stay separate and correctly counted" begin
        src = joinpath(data_directory, "Chartsheet.xlsx")
        p = tempname() * ".xlsx"
        cp(src, p)

        XLSX.openxlsx(p; mode="rw") do xf
            XLSX.renamesheet!(xf["Sheet1"], "RenamedData")
            XLSX.addsheet!(xf, "ExtraData")
        end

        app = String(zip_readentry(ZipReader(read(p)), "docProps/app.xml"))
        pairs  = parse_heading_pairs(app)
        titles = parse_titles_of_parts(app)

        ws_idx = findfirst(pr -> pr[1] == "Worksheets", pairs)
        ch_idx = findfirst(pr -> pr[1] == "Charts", pairs)
        @test ws_idx !== nothing
        @test ch_idx !== nothing
        @test pairs[ws_idx][2] == 2   # RenamedData, ExtraData
        @test pairs[ch_idx][2] == 1   # Chart1 untouched by a worksheet-only operation

        @test "RenamedData" in titles
        @test "ExtraData" in titles
        @test "Chart1" in titles
        @test !occursin(">Sheet1<", app)
        @test length(titles) == 3

        # Worksheets titles are grouped before Charts titles regardless of tab
        # order — with only one chart, it must be the last title.
        @test titles[end] == "Chart1"

        XLSX.openxlsx(p) do xf2
            @test issetequal(XLSX.sheetnames(xf2), ["RenamedData", "ExtraData", "Chart1"])
        end

        rm(p; force=true)
    end

end
