const _STRICT_MAIN = "http://purl.oclc.org/ooxml/spreadsheetml/main"

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
        @test_throws XLSX.XLSXError XLSX.deletesheet!(f["Tabelle1"])     # the only worksheet
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
        f= XLSX.openxlsx("mytest.xlsx"; mode="rw")
        XLSX.deletesheet!(f["Diagramm1"])
        @test XLSX.sheetnames(f) == ["Tabelle1"]
        @test isempty(XLSX.Charts.getCharts(f))
        isfile("mytest.xlsx") && rm("mytest.xlsx")
        SAVE_FILES && save_outfile(f)
    end
    @testset "readtable on Strict OOXML with target_sheet" begin
        f = joinpath(data_directory, "strict.xlsx")
        # readtable loads only the target sheet; the others must not be touched
        for s in ("general", "table", "lookup")
            @test XLSX.readtable(f, s) isa XLSX.DataTable
        end
        @test XLSX.readtable(f, 1) isa XLSX.DataTable
    end
    @testset "Normalisation reaches every element" begin
        # Strict namespaces can be declared below the root, e.g. xmlns:r on
        # cx:chart inside mc:AlternateContent in a drawing part. A write must
        # leave no Strict URI anywhere in the package.
        for fixture in ("strict.xlsx", "Strict-foo.xlsx", "chart_ex.xlsx")
            XLSX.writexlsx("mytest.xlsx", XLSX.openxlsx(joinpath(data_directory, fixture); mode="rw"), overwrite=true)
            SAVE_FILES && save_outfile("mytest.xlsx")

            r = ZipArchives.ZipReader(read("mytest.xlsx"))
            for name in ZipArchives.zip_names(r)
                (endswith(name, ".xml") || endswith(name, ".rels")) || continue
                s = ZipArchives.zip_readentry(r, name, String)
                @test !occursin("purl.oclc.org", s)
                occursin("purl.oclc.org", s) && @info "Strict URI survived" fixture name
            end
            isfile("mytest.xlsx") && rm("mytest.xlsx")
        end
    end
end

@testset "detecting and converting Strict OOXML" begin
    @testset "is_strict_ooxml - strict namespace without conformance attribute" begin
        xf = XLSX.newxlsx()
        xf.data["xl/workbook.xml"] = parse(
            """<workbook xmlns="$_STRICT_MAIN"/>""", XLSX.XML.Node)
        @test XLSX.is_strict_ooxml(xf) == true
    end

    @testset "is_strict_ooxml - conformance attribute" begin
        xf = XLSX.newxlsx()
        xf.data["xl/workbook.xml"] = parse(
            """<workbook xmlns="$_TRANS_MAIN" conformance="strict"/>""", XLSX.XML.Node)
        @test XLSX.is_strict_ooxml(xf) == true
    end

    @testset "is_strict_ooxml - falls back to relationship types in _rels/.rels" begin
        xf = XLSX.newxlsx()
        # workbook root gives no hint at all
        xf.data["xl/workbook.xml"] = parse("""<workbook xmlns="$_TRANS_MAIN"/>""", XLSX.XML.Node)
        xf.data["_rels/.rels"] = parse(
            """<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
            <Relationship Id="rId1" Target="xl/workbook.xml"
                Type="http://purl.oclc.org/ooxml/officeDocument/relationships/officeDocument"/>
            </Relationships>""", XLSX.XML.Node)
        @test XLSX.is_strict_ooxml(xf) == true
    end

    @testset "is_strict_ooxml - ordinary transitional file" begin
        xf = XLSX.newxlsx()
        @test XLSX.is_strict_ooxml(xf) == false
        SAVE_FILES && save_outfile(xf)
    end

    @testset "_strict_to_transitional_node! - remaps and drops conformance" begin
        node = XLSX.XML.Element("workbook"; xmlns=_STRICT_MAIN, conformance="strict")
        XLSX._strict_to_transitional_node!(node, "xl/workbook.xml")

        @test XLSX.get_attr(node, "xmlns") == _TRANS_MAIN
        @test XLSX.get_attr(node, "conformance", "") == ""
    end

    @testset "_strict_to_transitional_node! - node with no attributes is a no-op" begin
        node = XLSX.XML.Element("workbook")
        @test isnothing(XLSX._strict_to_transitional_node!(node, "xl/workbook.xml"))
    end

    @testset "_strict_to_transitional_node! - unknown purl namespace errors" begin
        node = XLSX.XML.Element("thing"; xmlns="http://purl.oclc.org/ooxml/notAThing/main")
        @test_throws XLSX.XLSXError XLSX._strict_to_transitional_node!(node, "xl/thing.xml")
    end

    @testset "convert_strict_to_transitional! - part held as a raw String is parsed" begin
        xf = XLSX.newxlsx()
        f = "xl/styles.xml"
        xf.data[f] = """<styleSheet xmlns="$_STRICT_MAIN"/>"""

        XLSX.convert_strict_to_transitional!(xf, 1)

        @test xf.data[f] isa XLSX.XML.Node
        @test XLSX.get_attr(XLSX.xml_root_element(xf.data[f]), "xmlns") == _TRANS_MAIN
        SAVE_FILES && save_outfile(xf)
    end

    @testset "convert_strict_to_transitional! - worksheet String substitution (pass 3)" begin
        xf = XLSX.newxlsx()
        f = XLSX.get_relationship_target_by_id("xl", XLSX.get_workbook(xf), xf[1].relationship_id)
        xf.data[f] = """<worksheet xmlns="$_STRICT_MAIN" conformance="strict"><sheetData/></worksheet>"""

        XLSX.convert_strict_to_transitional!(xf, 3)

        data = xf.data[f]
        @test data isa String
        @test occursin(_TRANS_MAIN, data)
        @test !occursin("purl.oclc.org", data)
        @test !occursin("conformance", data)
        SAVE_FILES && save_outfile(xf)
    end
end
