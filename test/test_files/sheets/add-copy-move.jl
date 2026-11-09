@testset "add/copy sheet!" begin
    # The text of every `name` element (e.g. "formula", "formula1") in a worksheet's XML.
    function _sheet_xml_texts(ws, name)
        out = String[]
        walk(node) = for c in XLSX.XML.eachelement(node)
            XLSX.localname(c) == name ? push!(out, XLSX.XML.simple_value(c)) : walk(c)
        end
        walk(XLSX.xml_root_element(XLSX.get_worksheet_xml_document(ws)))
        return out
    end
    # Add a list data validation on `sqref` with `formula1`; there is no API for one.
    function _add_list_validation!(ws, sqref, formula1)
        root = XLSX.xml_root_element(XLSX.get_worksheet_xml_document(ws))
        dv = XLSX.XML.Element("dataValidations"; count="1")
        push!(dv, XLSX.XML.Element("dataValidation", XLSX.XML.Element("formula1", XLSX.XML.Text(formula1));
            type="list", sqref))
        insert!(root.children, XLSX.insert_index(root, "dataValidations", XLSX.WORKSHEET_ORDER) + 1, dv)
    end

    # Chart references to a deleted sheet, invalidated as Excel does it.

    function _chart_elements(node, name)
        out = []
        for ch in XLSX.XML.eachelement(node)
            XLSX.localname(ch) == name && push!(out, ch)
            append!(out, _chart_elements(ch, name))
        end
        return out
    end

    _chart_f_texts(root) = [something(XLSX.XML.is_simple_value(f), "") for f in _chart_elements(root, "f")]

    @testset "addsheet!" begin
        @testset "adding named sheets" begin

            new_filename = "template_with_new_sheet.xlsx"
            f = XLSX.open_empty_template()
            s = XLSX.addsheet!(f, "new_sheet")
            s["A1"] = 10
            @test XLSX.sheetnames(f) == ["Sheet1", "new_sheet"]
            XLSX.writexlsx(new_filename, f, overwrite=true)
            SAVE_FILES && save_outfile(new_filename)


            big_sheetname = "aaaaaaaaaabbbbbbbbbbccccccccccd"
            s2 = XLSX.addsheet!(f, big_sheetname)

            XLSX.writexlsx(new_filename, f, overwrite=true)
            SAVE_FILES && save_outfile(new_filename)
            fx = XLSX.opentemplate(new_filename)
            @test XLSX.sheetnames(fx) == ["Sheet1", "new_sheet", big_sheetname]
            SAVE_FILES && save_outfile(fx)

        end

        @testset "invalid sheet names" begin

            f = XLSX.open_empty_template()
            s = XLSX.addsheet!(f, "new_sheet")
            s["A1"] = 10
            invalid_names = [
                "aaaaaaaaaabbbbbbbbbbccccccccccd1",
                "abc:def",
                "abcdef/",
                "\\aaaa",
                "hey?you",
                "[mysheet]",
                "asteri*"
            ]

            for invalid_name in invalid_names
                @test_throws XLSX.XLSXError XLSX.addsheet!(f, invalid_name)
            end

        end

        @testset "addsheet! after reopening gives a fresh xr:uid" begin
            # `uuid_rng` has a fixed seed, so a reopened file replays the ids it was first written with.
            uid(f, e) = match(r"xr:uid=\"(\{[^\"]+\})\"", String(XLSX.ZipArchives.zip_readentry(XLSX.ZipArchives.ZipReader(read(f)), e)))[1]
            xf = XLSX.newxlsx()
            XLSX.addsheet!(xf, "S2")
            XLSX.writexlsx("uid.xlsx", xf; overwrite=true)
            for out in ("uid2.xlsx", "uid3.xlsx")
                xf = XLSX.openxlsx("uid.xlsx"; mode="rw")
                XLSX.addsheet!(xf, "S3")
                XLSX.writexlsx(out, xf; overwrite=true)
            end
            SAVE_FILES && save_outfile("uid2.xlsx")
            uids = [uid("uid2.xlsx", "xl/worksheets/sheet$k.xml") for k in 1:3]
            @test allunique(uids)
            @test uids[3] == uid("uid3.xlsx", "xl/worksheets/sheet3.xml")   # still deterministic
            @test XLSX._unique_guid(XLSX.Random.Xoshiro(1), Set{String}()) != XLSX._unique_guid(XLSX.Random.Xoshiro(1), Set([XLSX._unique_guid(XLSX.Random.Xoshiro(1), Set{String}())]))
            foreach(rm, ("uid.xlsx", "uid2.xlsx", "uid3.xlsx"))
        end
    end

    @testset "copysheet!" begin
        @testset "the copy's references to the original sheet are repointed at the copy" begin
            xf = XLSX.newxlsx()
            s = xf[1]
            s[1:3, 1:4] = reshape(collect(1:12), 3, 4)
            XLSX.addsheet!(xf, "Other")
            XLSX.setFormula(s, "E1", "=Sheet1!A1+Other!A1")
            XLSX.setConditionalFormat(s, "B1:B3", :expression; formula="AND(Sheet1!\$M\$2>0, Other!\$A\$1>0, C1=\"abc\")")
            _add_list_validation!(s, "D1", "Sheet1!\$A\$1:\$A\$3")
            XLSX.addDefinedName(s, "tot", "=SUM(Sheet1!\$A:\$A)")

            c = XLSX.copysheet!(s, "Copy")
            wb = XLSX.get_workbook(xf)
            # string literals are kept as written; the rest of a CF formula is uppercased
            @test _sheet_xml_texts(s, "formula") == ["(AND(SHEET1!\$M\$2>0, OTHER!\$A\$1>0, C1=\"abc\"))"]
            for (ws, sh) in ((s, "Sheet1"), (c, "Copy"))
                @test XLSX.getFormula(ws, "E1") == "=$(sh)!A1+Other!A1"
                @test _sheet_xml_texts(ws, "formula1") == ["$(sh)!\$A\$1:\$A\$3"]
                @test XLSX.get_defined_name_value(ws, "tot") == "=SUM($(sh)!\$A:\$A)"
            end
            @test _sheet_xml_texts(c, "formula") == ["(AND(Copy!\$M\$2>0, OTHER!\$A\$1>0, C1=\"abc\"))"]

            path = "copy_repoint.xlsx"
            XLSX.writexlsx(path, xf; overwrite=true)
            SAVE_FILES && save_outfile(path)
            xf2 = XLSX.openxlsx(path; mode="rw")
            @test XLSX.getFormula(xf2["Copy"], "E1") == "=Copy!A1+Other!A1"
            @test _sheet_xml_texts(xf2["Copy"], "formula") == ["(AND(Copy!\$M\$2>0, OTHER!\$A\$1>0, C1=\"abc\"))"]
            @test _sheet_xml_texts(xf2["Copy"], "formula1") == ["Copy!\$A\$1:\$A\$3"]
            isfile(path) && rm(path)
        end

        @testset "values, defined names, borders and conditional formats are copied" begin

            f = XLSX.newxlsx()
            XLSX.renamesheet!(f["Sheet1"], "new_name")
            XLSX.addsheet!(f)
            for x = 1:10, y = 1:10
                f["Sheet1"][x, y] = x + y
                f["new_name"][x, y] = x * y
            end
            XLSX.addDefinedName(f["new_name"], "new_name_range", "A1:B10")
            XLSX.addDefinedName(f["Sheet1"], "Sheet1_range", "C1:D10")
            XLSX.setBorder(f["new_name"], "A1:D10"; allsides=["style" => "thin", "color" => "red"])
            XLSX.setBorder(f["Sheet1"], "A1:D10"; allsides=["style" => "thin", "color" => "red"])
            XLSX.setConditionalFormat(f["new_name"], "A1:D10", :colorScale)

            s3 = XLSX.copysheet!(f["new_name"], "copied_sheet")
            @test s3.name == "copied_sheet"
            @test s3["A1"] == 1
            @test s3[5, 5] == 25
            @test s3[10, 10] == 100
            @test XLSX.get_workbook(s3).worksheet_names == XLSX.get_workbook(f["new_name"]).worksheet_names
            @test XLSX.getConditionalFormats(s3) == XLSX.getConditionalFormats(f["new_name"])
            @test XLSX.getBorder(s3, "C5").border == XLSX.getBorder(f["new_name"], "C5").border

            # Check that the original sheet is unchanged
            s2 = f["new_name"]
            @test s2["A1"] == 1
            @test s2[5, 5] == 25
            @test s2[10, 10] == 100

            s4 = XLSX.copysheet!(s3)
            @test s4.name == "copied_sheet (copy)"
            @test s4["A1"] == 1
            @test s4[5, 5] == 25
            @test s4[10, 10] == 100

            @test XLSX.get_workbook(s4).worksheet_names == XLSX.get_workbook(f["new_name"]).worksheet_names
            XLSX.setBorder(s4, "F1:H10"; allsides=["style" => "thin", "color" => "green"])
            XLSX.setConditionalFormat(s4, "F1:H10", :colorScale; colorscale="redyellowgreen")

            XLSX.writexlsx("copied_sheets.xlsx", f, overwrite=true)
            SAVE_FILES && save_outfile("copied_sheets.xlsx")
            f = XLSX.opentemplate("copied_sheets.xlsx")
            @test XLSX.sheetnames(f) == ["new_name", "Sheet1", "copied_sheet", "copied_sheet (copy)"]
            @test XLSX.get_workbook(f["copied_sheet"]).worksheet_names == XLSX.get_workbook(f["new_name"]).worksheet_names
            @test XLSX.getConditionalFormats(f["copied_sheet (copy)"]) == XLSX.getConditionalFormats(s4)
            @test XLSX.getBorder(f["copied_sheet (copy)"], "C5").border == XLSX.getBorder(f["new_name"], "C5").border
            @test XLSX.getBorder(f["copied_sheet (copy)"], "G5").border == XLSX.getBorder(s4, "G5").border
            SAVE_FILES && save_outfile(f)

        end

        @testset "copysheet! clones and repoints c: chart parts" begin
            f = XLSX.opentemplate(joinpath(data_directory, "chart_basic.xlsx"))
            XLSX.copysheet!(f["Data"], "Copy")

            charts = XLSX.Charts.getCharts(f)
            @test length(charts) == 2
            @test allunique(c.path for c in charts)
            @test issetequal((c.sheet for c in charts), ["Data", "Copy"])

            cp = only(filter(c -> c.sheet == "Copy", charts))
            orig = only(filter(c -> c.sheet == "Data", charts))

            # The copy plots the copied sheet. getChartRanges returns one named
            # tuple per series for a `c:` chart, so the ranges are in its fields.
            cp_ranges = XLSX.Charts.getChartRanges(cp)
            @test [string(r.categories) for r in cp_ranges] == ["Copy!A2:A5", "Copy!A2:A5"]
            @test [string(r.values) for r in cp_ranges] == ["Copy!B2:B5", "Copy!C2:C5"]

            # Series name refs are repointed too, not just categories and values.
            @test all(s -> startswith(s.name_ref.ref, "Copy!"), XLSX.Charts.getChartSeries(cp))

            # The original is untouched — catches copynode aliasing the two trees.
            orig_ranges = XLSX.Charts.getChartRanges(orig)
            @test [string(r.categories) for r in orig_ranges] == ["Data!A2:A5", "Data!A2:A5"]
            @test [string(r.values) for r in orig_ranges] == ["Data!B2:B5", "Data!C2:C5"]
            @test all(s -> startswith(s.name_ref.ref, "Data!"), XLSX.Charts.getChartSeries(orig))

            # The cache is deliberately left alone, so the copy still reads.
            @test XLSX.Charts.getChartData(cp) isa XLSX.DataTable

            # And it survives a round trip.
            g = XLSX.writexlsx("mytest.xlsx", f; overwrite=true)
            h = XLSX.readxlsx(g)
            cp2 = only(filter(c -> c.sheet == "Copy", XLSX.Charts.getCharts(h)))
            @test [string(r.values) for r in XLSX.Charts.getChartRanges(cp2)] ==
                ["Copy!B2:B5", "Copy!C2:C5"]
            SAVE_FILES && save_outfile(f)
        end

        @testset "copysheet! with a chartEx chart" begin
            f = XLSX.opentemplate(joinpath(data_directory, "chart_ex.xlsx"))
            XLSX.copysheet!(f["Data"], "Copy")

            charts = XLSX.Charts.getCharts(f)
            @test length(charts) == 2
            @test all(c -> c isa XLSX.Charts.ChartEx, charts)
            @test issetequal((c.sheet for c in charts), ["Data", "Copy"])
            @test allunique(c.path for c in charts)
            @test all(c -> XLSX.Charts.getChartType(c) === :waterfall, charts)

            cp = only(filter(c -> c.sheet == "Copy", charts))

            # The copy gets its own hidden defined names, pointing at the copy.
            @test XLSX.Charts._cx_refs(cp) == ["_xlchart.v2.0", "_xlchart.v2.1"]
            @test all(r -> startswith(string(r), "Copy!"), XLSX.Charts.getChartRanges(cp))

            # The original's names and ranges are untouched.
            orig = only(filter(c -> c.sheet == "Data", charts))
            @test XLSX.Charts._cx_refs(orig) == ["_xlchart.v1.0", "_xlchart.v1.1"]
            @test all(r -> startswith(string(r), "Data!"), XLSX.Charts.getChartRanges(orig))

            # Two names added, all still hidden and none user-visible.
            all_names = XLSX.getAllDefinedNames(f; include_system=true)
            @test length(all_names) == 11
            @test all(dn -> dn.hidden, all_names)
            @test isempty(XLSX.getDefinedNames(f))

            g = XLSX.writexlsx("mytest.xlsx", f; overwrite=true)
            h = XLSX.readxlsx(g)
            @test length(XLSX.Charts.getCharts(h)) == 2
            @test allunique(c.path for c in XLSX.Charts.getCharts(h))
            SAVE_FILES && save_outfile(f)
        end

        @testset "copysheet! keeps a shared chartEx name shared" begin
            xf = XLSX.openxlsx(joinpath(data_directory, "chartex_layouts.xlsx"); mode="rw")
            XLSX.copysheet!(xf["pareto"], "pareto2")
            XLSX.writexlsx("mytest.xlsx", xf, overwrite=true)
            SAVE_FILES && save_outfile("mytest.xlsx")

            f = XLSX.readxlsx("mytest.xlsx")
            cxs = filter(x -> x isa XLSX.Charts.ChartEx, XLSX.Charts.getCharts(f))
            orig = only(filter(x -> x.sheet == "pareto", cxs))
            copy = only(filter(x -> x.sheet == "pareto2", cxs))

            ob = XLSX.Charts.getChartDataBlocks(orig)
            cb = XLSX.Charts.getChartDataBlocks(copy)

            # the original is untouched: its two blocks still share one name
            @test ob[1].dimensions[1].formula == ob[2].dimensions[1].formula == "_xlchart.v1.18"

            # the copy's two blocks share one name too, and it is a new one
            shared = cb[1].dimensions[1].formula
            @test cb[2].dimensions[1].formula == shared
            @test shared != "_xlchart.v1.18"

            # every data dimension and series name in the copy points at the copied sheet
            for d in vcat(cb[1].dimensions, cb[2].dimensions)
                @test d.range.sheet == "pareto2"
            end
            @test XLSX.Charts.getSeriesNameRange(copy, 1).sheet == "pareto2"
            @test XLSX.Charts.getSeriesNameRange(copy, 2).sheet == "pareto2"

            # data: cat (shared), val, val; series names: two. Five names, not six.
            @test length(unique(XLSX.Charts._cx_refs(copy))) == 3                      # data dimensions
            @test length(XLSX.Charts._cx_refs(copy)) == 4                              # cat appears twice
            isfile("mytest.xlsx") && rm("mytest.xlsx")
        end

        @testset "copysheet! clones chart parts for both schemas" begin
            f = XLSX.opentemplate(joinpath(data_directory, "chart_mixed.xlsx"))
            XLSX.copysheet!(f["Data"], "Copy")

            charts = XLSX.Charts.getCharts(f)
            @test length(charts) == 4
            @test allunique(c.path for c in charts)
            @test count(c -> c.sheet == "Copy", charts) == 2
            @test issetequal((XLSX.Charts.getChartType(c) for c in charts if c.sheet == "Copy"),
                [:waterfall, :barChart])

            # Stems are handled per family: chart1 -> chart2, chartEx1 -> chartEx2.
            @test issetequal((c.name for c in charts),
                ["chart1", "chart2", "chartEx1", "chartEx2"])

            # Anchors come across with the drawing.
            @test issetequal((c.from for c in charts), ["F1", "F17"])

            # Both the c: and the cx: chart are repointed at the copy.
            bar = only(filter(c -> c.sheet == "Copy" && c isa XLSX.Charts.Chart, charts))
            @test all(r -> startswith(string(r.values), "Copy!"), XLSX.Charts.getChartRanges(bar))

            wf = only(filter(c -> c.sheet == "Copy" && c isa XLSX.Charts.ChartEx, charts))
            @test all(r -> startswith(string(r), "Copy!"), XLSX.Charts.getChartRanges(wf))

            g = XLSX.writexlsx("mytest.xlsx", f; overwrite=true)
            h = XLSX.readxlsx(g)
            @test length(XLSX.Charts.getCharts(h)) == 4
            @test allunique(c.path for c in XLSX.Charts.getCharts(h))
            SAVE_FILES && save_outfile(f)
        end

        @testset "copysheet! clones chart defined names" begin
            f = XLSX.opentemplate(joinpath(data_directory, "chart_ex.xlsx"))
            before = XLSX.getAllDefinedNames(f; include_system=true)
            XLSX.copysheet!(f["Data"], "Copy")
            after = XLSX.getAllDefinedNames(f; include_system=true)

            # A chartEx chart's sources are hidden defined names, so the copy needs
            # its own set pointing at the copied sheet.
            @test length(before) == 9
            @test length(after) == 11
            @test all(dn -> isnothing(dn.scope), after)   # all workbook-scoped
            @test all(dn -> dn.hidden, after)             # and all still hidden
            @test isempty(XLSX.getDefinedNames(f))        # none user-visible

            # The originals are unchanged; the new ones point at the copy.
            added = setdiff(Set(dn.name for dn in after), Set(dn.name for dn in before))
            @test added == Set(["_xlchart.v2.0", "_xlchart.v2.1"])
            @test all(dn -> startswith(string(dn.value), "Copy!"),
                filter(dn -> dn.name in added, after))
            @test all(dn -> startswith(string(dn.value), "Data!"),
                filter(dn -> dn.name ∉ added, after))
            SAVE_FILES && save_outfile(f)
        end

        # The copy once shared the original's cache, so a change to either sheet
        # showed up in both.
        @testset "copysheet! gives the copy its own cells" begin
            f = XLSX.newxlsx()
            s = f[1]
            s["A1"] = 1
            s["A2"] = 2
            c = XLSX.copysheet!(s)

            c["A1"] = 100
            c["A5"] = 5
            XLSX.setFont(c, "A2"; bold=true)
            XLSX.setRowHeight(c, 2; height=30)
            XLSX.setFormula(c, "C1", "=A1+1")
            @test c["A1"] == 100 && c["A5"] == 5
            @test s["A1"] == 1
            @test ismissing(s["A5"])
            @test isnothing(XLSX.getFont(s, "A2"))
            @test isnothing(XLSX.getRowHeight(s, "A2"))
            @test isnothing(XLSX.getFormula(s, "C1"))

            # and the other way round
            s["B1"] = "original"
            @test ismissing(c["B1"])

            io = IOBuffer()
            XLSX.writexlsx(io, f)
            r = XLSX.readxlsx(seekstart(io))
            @test r[1]["A1"] == 1 && ismissing(r[1]["A5"]) && r[1]["B1"] == "original"
            @test r[2]["A1"] == 100 && r[2]["A5"] == 5 && ismissing(r[2]["B1"])
            SAVE_FILES && save_outfile(f)
        end
    end

    @testset "deletesheet!" begin
        @testset "conditional format, data validation and defined name formulas lose references to the deleted sheet" begin
            xf = XLSX.newxlsx()
            s = xf[1]
            s[1:3, 1:4] = reshape(collect(1:12), 3, 4)
            XLSX.addsheet!(xf, "Gone")
            XLSX.addsheet!(xf, "MyGone")
            XLSX.setConditionalFormat(s, "B1:B3", :expression; formula="AND(Gone!\$A\$1>0, MyGone!A1>0)")
            _add_list_validation!(s, "D1", "Gone!\$A\$1:\$A\$3")
            XLSX.addDefinedName(xf, "twice", "=Gone!\$A\$1*2")
            XLSX.addDefinedName(s, "local", "=SUM(Gone!A1:A3, Sheet1!A1)")

            XLSX.deletesheet!(xf, "Gone")
            @test _sheet_xml_texts(s, "formula") == ["(AND(#REF!>0, MYGONE!A1>0))"]
            @test _sheet_xml_texts(s, "formula1") == ["#REF!"]
            @test XLSX.get_defined_name_value(XLSX.get_workbook(xf), "twice") == "=#REF!*2"
            @test XLSX.get_defined_name_value(s, "local") == "=SUM(#REF!, Sheet1!A1)"
            SAVE_FILES && save_outfile(xf)

            # a read-only file can still delete a sheet, and read the rest lazily afterwards
            ro = XLSX.readxlsx(joinpath(data_directory, "general.xlsx"))
            XLSX.deletesheet!(ro, "table2")
            @test size(ro["table"][:]) == (12, 8)
        end

        @testset "deleting sheets, and formulas that referred to them" begin

            new_filename = "template_with_new_sheet.xlsx"
            big_sheetname = "aaaaaaaaaabbbbbbbbbbccccccccccd"
            f0 = XLSX.open_empty_template()
            XLSX.addsheet!(f0, "new_sheet")
            XLSX.addsheet!(f0, big_sheetname)
            XLSX.writexlsx(new_filename, f0, overwrite=true)
            fx = XLSX.opentemplate(new_filename)
            XLSX.deletesheet!(fx, big_sheetname)
            @test XLSX.sheetnames(fx) == ["Sheet1", "new_sheet"]
            XLSX.writexlsx(new_filename, fx, overwrite=true)
            SAVE_FILES && save_outfile(new_filename)
            f = XLSX.readxlsx(new_filename)
            @test XLSX.sheetnames(f) == ["Sheet1", "new_sheet"]

            f = XLSX.opentemplate(joinpath(data_directory, "general.xlsx"))
            sc = XLSX.sheetcount(f)
            XLSX.deletesheet!(f, "empty")
            @test XLSX.sheetcount(f) == sc - 1 # Check it's gone.
            @test XLSX.hassheet(f, "empty") == false # Check it's gone.
            @test_throws XLSX.XLSXError XLSX.deletesheet!(f, "empty") # Already deleted.
            @test_throws XLSX.XLSXError XLSX.deletesheet!(f, "nosuchsheet") # Never there.
            s2 = XLSX.addsheet!(f, "this_now")
            @test XLSX.sheetnames(f) == ["general", "table3", "table4", "table", "table2", "table5", "table6", "table7", "lookup", "header_error", "named_ranges_2", "named_ranges", "this_now"]
            XLSX.writexlsx(new_filename, f, overwrite=true)
            SAVE_FILES && save_outfile(new_filename)

            f = XLSX.opentemplate(new_filename)
            @test XLSX.sheetnames(f) == ["general", "table3", "table4", "table", "table2", "table5", "table6", "table7", "lookup", "header_error", "named_ranges_2", "named_ranges", "this_now"]
            XLSX.deletesheet!(f, "named_ranges")
            XLSX.deletesheet!(f["general"])
            @test XLSX.sheetnames(f) == ["table3", "table4", "table", "table2", "table5", "table6", "table7", "lookup", "header_error", "named_ranges_2", "this_now"]
            XLSX.writexlsx(new_filename, f, overwrite=true)
            SAVE_FILES && save_outfile(new_filename)
            dtable = XLSX.readtable(new_filename, "table4", "F:G")
            data, col_names = dtable.data, dtable.column_labels
            @test col_names == [:H2, :H3]
            test_data = Any[Any["C3", missing], Any[missing, "D4"]]
            check_test_data(data, test_data)
            @test XLSX.deletesheet!(f, 1) === f
            @test XLSX.sheetnames(f) == ["table4", "table", "table2", "table5", "table6", "table7", "lookup", "header_error", "named_ranges_2", "this_now"]
            XLSX.writexlsx(new_filename, f, overwrite=true)
            SAVE_FILES && save_outfile(new_filename)
            dtable = XLSX.readtable(new_filename, "table4", "F:G")
            data, col_names = dtable.data, dtable.column_labels
            @test col_names == [:H2, :H3]
            test_data = Any[Any["C3", missing], Any[missing, "D4"]]
            check_test_data(data, test_data)

            f = XLSX.opentemplate(joinpath(data_directory, "Book_1904.xlsx")) # Only one sheet - can't delete
            @test_throws XLSX.XLSXError XLSX.deletesheet!(f, 1)
            s = f[1]
            @test_throws XLSX.XLSXError XLSX.deletesheet!(s)
            @test_throws XLSX.XLSXError XLSX.deletesheet!(f, "Sheet1")

            SAVE_FILES && save_outfile(f)
            f = XLSX.openxlsx(joinpath(data_directory, "deletesheet.xlsx"), mode="rw")
            XLSX.deletesheet!(f[1])
            @test XLSX.getcell(f[1], "A1") == XLSX.Cell(XLSX.get_workbook(f), XLSX.CellRef("A1"), "e", "", "#REF!", "", true)
            @test XLSX.getFormula(f[1], "A1") == "=#REF!+#REF!"
            SAVE_FILES && save_outfile(f)

            isfile("template_with_new_sheet.xlsx") && rm("template_with_new_sheet.xlsx")

            # XLSX.jl cannot currently create formula references that refer to other sheets, so we have to manually 
            # insert them into the workbook's formula cache to test. Excel cannot read such a workbook and deletes these 
            # formulas on opening. Tests are for completeness only
            f=XLSX.openxlsx("renamedelete.xlsx", mode="w")
            s=f[1]
            s[1:10, 1] = collect(1:10)
            XLSX.addsheet!(f)
            s2=f[2]
            s2["A1:B4"]=""
            XLSX.setFormula(s2, "A1", "=Sheet1!A1+10")
            XLSX.getcell(s2, "B1").formula=true
            XLSX.getcell(s2, "B2").formula=true
            XLSX.getcell(s2, "B3").formula=true
            XLSX.getcell(s2, "B4").formula=true
            w=XLSX.get_workbook(s2)
            w.formulas[XLSX.SheetCellRef("Sheet2!B1")] = XLSX.ReferencedFormula("=Sheet1!A1+10", 0, "B1:B4", nothing)
            w.formulas[XLSX.SheetCellRef("Sheet2!B2")] = XLSX.FormulaReference(0, nothing)
            w.formulas[XLSX.SheetCellRef("Sheet2!B3")] = XLSX.FormulaReference(0, nothing)
            w.formulas[XLSX.SheetCellRef("Sheet2!B4")] = XLSX.FormulaReference(0, nothing)

            XLSX.deletesheet!(s)
            @test XLSX.hassheet(f, "Sheet1") == false
            @test XLSX.getcell(s2, "A1") == XLSX.Cell(XLSX.get_workbook(f), XLSX.CellRef("A1"), "e", "", "#REF!", "", true)
            @test XLSX.getcell(s2, "B1") == XLSX.Cell(XLSX.get_workbook(f), XLSX.CellRef("B1"), "e", "", "#REF!", "", true)
            @test XLSX.getcell(s2, "B2") == XLSX.Cell(XLSX.get_workbook(f), XLSX.CellRef("B2"), "e", "", "#REF!", "", true)
            @test XLSX.getcell(s2, "B4") == XLSX.Cell(XLSX.get_workbook(f), XLSX.CellRef("B4"), "e", "", "#REF!", "", true)
            @test XLSX.getFormula(s2, "A1") == "=#REF!+10"
            @test XLSX.getFormula(s2, "B1") == "=#REF!+10"
            @test XLSX.getFormula(s2, "B2") == "=#REF!+10"
            @test XLSX.getFormula(s2, "B4") == "=#REF!+10"
            @test XLSX.get_formula_from_cache(s2, XLSX.CellRef("A1")) == XLSX.Formula("=#REF!+10", nothing, nothing, nothing)
            @test XLSX.get_formula_from_cache(s2, XLSX.CellRef("B1")) == XLSX.ReferencedFormula("=#REF!+10", 0, "B1:B4", nothing)
            @test XLSX.get_formula_from_cache(s2, XLSX.CellRef("B2")) == XLSX.FormulaReference(0, nothing)
            @test XLSX.get_formula_from_cache(s2, XLSX.CellRef("B4")) == XLSX.FormulaReference(0, nothing)
            SAVE_FILES && save_outfile(f)

        end

        @testset "deletesheet! keeps the active tab on the same sheet" begin
            xf = XLSX.newxlsx("A")
            XLSX.addsheet!(xf, "B")
            XLSX.addsheet!(xf, "C")
            active(x) = parse(Int, XLSX.XML.attributes(first(XLSX.elements_with_tag(first(XLSX.elements_with_tag(
                XLSX.xml_root_element(XLSX.xmlroot(x, "xl/workbook.xml")), "bookViews")), "workbookView")))["activeTab"])
            XLSX.xmlroot(xf, "xl/workbook.xml")                                    # ensure loaded
            wv = first(XLSX.elements_with_tag(first(XLSX.elements_with_tag(
                XLSX.xml_root_element(XLSX.xmlroot(xf, "xl/workbook.xml")), "bookViews")), "workbookView"))
            wv["activeTab"] = "2"                                                  # C is active
            XLSX.deletesheet!(xf, "A")
            @test active(xf) == 1                                                  # still C
            XLSX.deletesheet!(xf, "C")
            @test active(xf) == 0                                                  # C gone: clamped to B
            SAVE_FILES && save_outfile(xf)
        end

        @testset "deletesheet! invalidates chart references" begin

            @testset "c: chart becomes #REF!, cache kept" begin
                xf = XLSX.openxlsx(joinpath(data_directory, "Chartsheet.xlsx"); mode="rw")
                XLSX.addsheet!(xf, "Keep")
                XLSX.deletesheet!(xf, "Sheet1")

                c = only(XLSX.getCharts(xf))
                @test _chart_f_texts(XLSX.Charts.chart_root(c)) == ["#REF!", "#REF!"]   # as Excel writes it
                @test all(r -> r.values === XLSX.XL_REF, XLSX.getChartRanges(c))
                @test XLSX.getChartData(xf, c.name) isa XLSX.DataTable                  # cache kept

                f = "deletesheet_chart_ref.xlsx"
                XLSX.writexlsx(f, xf; overwrite=true)
                SAVE_FILES && save_outfile(xf)
                yf = XLSX.openxlsx(f; mode="rw")
                d = only(XLSX.getCharts(yf))
                @test all(r -> r.values === XLSX.XL_REF, XLSX.getChartRanges(d))
                a, b = XLSX.getChartData(yf, d.name), XLSX.getChartData(xf, c.name)
                @test a.column_labels == b.column_labels && a.data == b.data
                SAVE_FILES && save_outfile(yf)
                isfile(f) && rm(f)
            end

            @testset "chartEx chart loses the reference" begin
                # chartex_kinds.xlsx with the treemap moved to a chartsheet in Excel,
                # so deleting its data sheet leaves the chart behind.
                xf = XLSX.openxlsx(joinpath(data_directory, "chartex_chartsheet.xlsx"); mode="rw")
                others = filter(c -> XLSX.Charts.getChartType(c) !== :treemap, XLSX.getCharts(xf))
                before = Dict(c.name => string.(XLSX.getChartRanges(c)) for c in others)

                XLSX.deletesheet!(xf, "treemap")

                c = only(filter(c -> XLSX.Charts.getChartType(c) === :treemap, XLSX.getCharts(xf)))
                root = XLSX.Charts.chart_root(c)
                @test isempty(_chart_f_texts(root))
                strdim = only(_chart_elements(root, "strDim"))
                numdim = only(_chart_elements(root, "numDim"))
                @test length(_chart_elements(strdim, "lvl")) == 2                      # one per hierarchy level
                @test length(_chart_elements(numdim, "lvl")) == 1
                @test all(l -> XLSX.get_attr(l, "ptCount") == "0",
                    vcat(_chart_elements(strdim, "lvl"), _chart_elements(numdim, "lvl")))
                @test isnothing(XLSX.first_element_with_tag(only(_chart_elements(root, "series")), "tx"))
                @test all(isnothing, XLSX.getChartRanges(c))

                wb = XLSX.get_workbook(xf)
                @test !any(v -> v.value isa XLSX.DefinedNameRangeTypes && v.value.sheet == "treemap",
                    values(wb.workbook_names))

                # charts plotting other sheets are untouched
                for d in filter(d -> haskey(before, d.name), XLSX.getCharts(xf))
                    @test string.(XLSX.getChartRanges(d)) == before[d.name]
                end

                f = "deletesheet_chartex_ref.xlsx"
                XLSX.writexlsx(f, xf; overwrite=true)
                SAVE_FILES && save_outfile(xf)
                yf = XLSX.openxlsx(f; mode="rw")
                d = only(filter(d -> XLSX.Charts.getChartType(d) === :treemap, XLSX.getCharts(yf)))
                @test all(isnothing, XLSX.getChartRanges(d))
                SAVE_FILES && save_outfile(yf)
                isfile(f) && rm(f)
            end
            @testset "levels in a chartEx reference" begin
                @test XLSX._ref_levels("Data!\$A\$2:\$B\$7", "") == 2         # columns, the default
                @test XLSX._ref_levels("Data!\$A\$2:\$B\$7", "row") == 6      # rows, for dir="row"
                @test XLSX._ref_levels("Data!\$C\$2", "") == 1                # a single cell
            end
        end

        @testset "deletesheet! invalidates direct chartEx references (as addChartEx writes)" begin
            xf = XLSX.newxlsx("Data")
            ws = xf["Data"]
            rows = [("Region", "Country", "Sales"),
                ("Europe", "France", 30), ("Europe", "Germany", 45), ("Europe", "Spain", 20),
                ("Asia", "Japan", 50), ("Asia", "India", 35),
                ("Americas", "USA", 60), ("Americas", "Brazil", 25)]
            for (i, r) in enumerate(rows), (j, v) in enumerate(r)
                ws[i, j] = v
            end
            XLSX.Charts.addChartEx(xf, :treemap, "Data!C2:C8"; categories="Data!A2:B8", name_ref="Data!C1")
            XLSX.addsheet!(xf, "Keep")
            XLSX.deletesheet!(xf, "Data")

            c = only(XLSX.getCharts(xf))
            root = XLSX.Charts.chart_root(c)
            @test isempty(_chart_f_texts(root))
            @test length(_chart_elements(only(_chart_elements(root, "strDim")), "lvl")) == 2
            @test length(_chart_elements(only(_chart_elements(root, "numDim")), "lvl")) == 1
            @test isnothing(XLSX.first_element_with_tag(only(_chart_elements(root, "series")), "tx"))
            @test all(isnothing, XLSX.getChartRanges(c))
            SAVE_FILES && save_outfile(xf)
        end

        @testset "deletesheet! removes chart parts (issue #427)" begin
            f = XLSX.opentemplate(joinpath(data_directory, "chart_basic.xlsx"))
            XLSX.addsheet!(f, "Keep")
            XLSX.deletesheet!(f, "Data")

            @test isempty(XLSX.Charts.getCharts(f))
            @test isempty(XLSX.Charts.chart_parts(f))
            @test !haskey(f.data, "xl/charts/chart1.xml")
            @test !haskey(f.data, "xl/drawings/drawing1.xml")
            SAVE_FILES && save_outfile(f)
        end

        @testset "deletesheet! with a chartEx chart" begin
            f = XLSX.opentemplate(joinpath(data_directory, "chart_ex.xlsx"))
            XLSX.addsheet!(f, "Keep")

            XLSX.deletesheet!(f, "Data")
            @test !XLSX.hassheet(f, "Data")
            @test isempty(XLSX.Charts.getCharts(f))

            # The chartEx cluster and the drawing must all go.
            for p in ("xl/charts/chartEx1.xml", "xl/charts/style1.xml",
                "xl/charts/colors1.xml", "xl/drawings/drawing1.xml")
                @test !haskey(f.data, p)
            end
            @test isempty(XLSX.Charts.chartex_parts(f))

            # And the file must still be writable and readable.
            g = XLSX.writexlsx("mytest.xlsx", f, overwrite=true)
            h = XLSX.readxlsx(g)
            @test isempty(XLSX.Charts.getCharts(h))
            SAVE_FILES && save_outfile(f)
        end
    end

    @testset "renamesheet!" begin
        @testset "renaming sheets" begin

            f=XLSX.openxlsx("renamedelete.xlsx", mode="w")
            s=f[1]
            s[1:10, 1] = collect(1:10)
            XLSX.addsheet!(f, "newsheet")
            s2 = f["newsheet"]
            XLSX.setFormula(s2, "A1", "=Sheet1!A1+10")
            XLSX.setFormula(s2, "B1:B3", "=A1+10")

            @test XLSX.getcell(s2, "A1") == XLSX.Cell(XLSX.get_workbook(f), XLSX.CellRef("A1"), "", "", "", "", true)
            @test XLSX.getFormula(s2, "A1") == "=Sheet1!A1+10"
            @test XLSX.get_formula_from_cache(s2, XLSX.CellRef("A1")) == XLSX.Formula("=Sheet1!A1+10", nothing, nothing, nothing)

            w = XLSX.get_workbook(s2)
            w.formulas[XLSX.SheetCellRef("newsheet!B1")] = XLSX.ReferencedFormula("=A1+10", 0, "B1:B3", nothing)
            w.formulas[XLSX.SheetCellRef("newsheet!B2")] = XLSX.FormulaReference(0, nothing)
            w.formulas[XLSX.SheetCellRef("newsheet!B3")] = XLSX.FormulaReference(0, nothing)

            XLSX.renamesheet!(s2, "Sheet4")
            @test s2.name == "Sheet4"
            @test XLSX.getcell(s2, "A1") == XLSX.Cell(XLSX.get_workbook(f), XLSX.CellRef("A1"), "", "", "", "", true)
            @test XLSX.getFormula(s2, "A1") == "=Sheet1!A1+10"
            @test XLSX.get_formula_from_cache(s2, XLSX.CellRef("A1")) == XLSX.Formula("=Sheet1!A1+10", nothing, nothing, nothing)
            @test XLSX.getFormula(s2, "B1") == "=A1+10"
            @test XLSX.getFormula(s2, "B2") == "=A2+10"
            @test XLSX.getFormula(s2, "B3") == "=A3+10"

            XLSX.renamesheet!(s, "Sheet3")
            @test s.name == "Sheet3"
            @test XLSX.getcell(s2, "A1") == XLSX.Cell(XLSX.get_workbook(f), XLSX.CellRef("A1"), "", "", "", "", true)
            @test XLSX.getFormula(s2, "A1") == "=Sheet3!A1+10"
            @test XLSX.get_formula_from_cache(s2, XLSX.CellRef("A1")) == XLSX.Formula("=Sheet3!A1+10", nothing, nothing, nothing)
            SAVE_FILES && save_outfile(f)

        end

        @testset "renaming a sheet repoints chart references" begin
            path = "chart_rename.xlsx"
            cp(joinpath(data_directory, "chart_basic.xlsx"), path; force=true)
            xf = XLSX.openxlsx(path; mode="rw")
            XLSX.renamesheet!(xf["Data"], "Sales")
            c = only(XLSX.Charts.getCharts(xf))
            @test all(r -> startswith(string(r.values), "Sales!"), XLSX.Charts.getChartRanges(c))
            SAVE_FILES && save_outfile(xf)
            isfile(path) && rm(path)
        end

        @testset "renaming a sheet repoints conditional format, data validation and defined name formulas" begin
            xf = XLSX.newxlsx()
            s = xf[1]
            s[1:3, 1:4] = reshape(collect(1:12), 3, 4)
            XLSX.setConditionalFormat(s, "B1:B3", :expression; formula="Sheet1!\$M\$2>0")
            XLSX.addsheet!(xf, "Other")
            other = xf["Other"]
            other[1:3, 1] = [1, 2, 3]
            XLSX.setConditionalFormat(other, "A1:A3", :expression; formula="AND(Sheet1!\$M\$3>0, MySheet1!A1>0)")
            XLSX.addDefinedName(xf, "total", "=SUM(Sheet1!\$A:\$A)")

            _add_list_validation!(other, "B1", "Sheet1!\$A\$1:\$A\$3")

            XLSX.renamesheet!(s, "My Data")

            formulas(ws) = _sheet_xml_texts(ws, "formula")
            formula1(ws) = _sheet_xml_texts(ws, "formula1")
            @test formulas(s) == ["('My Data'!\$M\$2>0)"]
            @test formulas(other) == ["(AND('My Data'!\$M\$3>0, MYSHEET1!A1>0))"]   # `MySheet1!` is another sheet
            @test formula1(other) == ["'My Data'!\$A\$1:\$A\$3"]
            @test XLSX.get_defined_name_value(XLSX.get_workbook(xf), "total") == "=SUM('My Data'!\$A:\$A)"

            path = "cf_rename.xlsx"
            XLSX.writexlsx(path, xf; overwrite=true)
            SAVE_FILES && save_outfile(path)
            xf2 = XLSX.readxlsx(path)
            @test formulas(xf2["Other"]) == ["(AND('My Data'!\$M\$3>0, MYSHEET1!A1>0))"]
            @test formula1(xf2["Other"]) == ["'My Data'!\$A\$1:\$A\$3"]
            isfile(path) && rm(path)
        end

        @testset "chart reference repointing matches whole sheet names" begin
            @test XLSX._repoint_ref_string("Data!\$A\$1:\$A\$5", "Data!", "Sales!") == "Sales!\$A\$1:\$A\$5"
            @test XLSX._repoint_ref_string("MyData!\$A\$1", "Data!", "Sales!") == "MyData!\$A\$1"
            @test XLSX._repoint_ref_string("(Data!\$A\$1,Data!\$C\$1)", "Data!", "Sales!") == "(Sales!\$A\$1,Sales!\$C\$1)"
            @test XLSX._repoint_ref_string("'My Sheet'!\$B\$2", "'My Sheet'!", "Sales!") == "Sales!\$B\$2"
            @test XLSX._repoint_ref_string("[1]Data!\$A\$1", "Data!", "Sales!") == "[1]Data!\$A\$1"
        end

        @testset "rename! alias, and the Workbook forms of the content-type helpers" begin
            xf = XLSX.newxlsx()
            sh = xf[1]
            sh[1:3, 1:4] = reshape(collect(1:12), 3, 4)

            # the unexported alias and the Workbook forms of the content-type helpers
            XLSX.rename!(sh, "Alias")
            @test sh.name == "Alias"
            wb = XLSX.get_workbook(xf)
            XLSX.add_override!(wb, "/xl/dummy.xml", "application/xml")
            @test occursin("/xl/dummy.xml", XLSX.XML.write(XLSX.xmlroot(xf, "[Content_Types].xml")))
            XLSX.remove_override!(wb, "/xl/dummy.xml")
            @test !occursin("/xl/dummy.xml", XLSX.XML.write(XLSX.xmlroot(xf, "[Content_Types].xml")))
            SAVE_FILES && save_outfile(xf)
        end
    end

    @testset "copysheet!/deletesheet! part cleanup (issue #427)" begin

        # Shared helper: map a sheet name to the worksheet XML path it actually
        # lives at, by reading workbook.xml + workbook.xml.rels directly rather
        # than assuming file numbering.
        function sheet_xml_path(zipbytes::Vector{UInt8}, sheetname::String)
            r = ZipReader(zipbytes)
            wbxml = String(zip_readentry(r, "xl/workbook.xml"))
            wbrels = String(zip_readentry(r, "xl/_rels/workbook.xml.rels"))
            sheet_rid = Dict(m.captures[1] => m.captures[2]
                             for m in eachmatch(r"<sheet name=\"([^\"]+)\"[^>]*r:id=\"(rId\d+)\"", wbxml))
            rid_target = Dict(m.captures[1] => m.captures[2]
                              for m in eachmatch(r"<Relationship Id=\"(rId\d+)\"[^>]*Target=\"([^\"]+)\"", wbrels))
            return "xl/" * rid_target[sheet_rid[sheetname]]
        end

        @testset "deletesheet! on a table-only sheet (no drawing) removes its rels file" begin
            # UTF-16.xlsx's SourceData sheet has a table but no drawing.
            src = joinpath(data_directory, "UTF-16.xlsx")
            p = "copyadd_deletesheet_on_a_table_only.xlsx"
            cp(src, p; force=true)

            XLSX.openxlsx(p; mode="rw") do xf
                XLSX.deletesheet!(xf, "SourceData")
            end
            SAVE_FILES && save_outfile(p)

            r = ZipReader(read(p))
            names = zip_names(r)

            @test !("xl/tables/table1.xml" in names)
            @test !("xl/worksheets/_rels/sheet1.xml.rels" in names)
            @test !occursin("tables/table1.xml", String(zip_readentry(r, "[Content_Types].xml")))

            XLSX.openxlsx(p) do xf2
                @test XLSX.sheetnames(xf2) == ["Sheet1"]
            end
            isfile(p) && rm(p)
        end

        @testset "deletesheet! removes table, comments, VML, drawing, media, and their Overrides together" begin
            src = joinpath(data_directory, "TableCommentsVML.xlsx")
            p = "copyadd_deletesheet_removes_table.xlsx"
            cp(src, p; force=true)

            XLSX.openxlsx(p; mode="rw") do xf
                XLSX.addsheet!(xf, "Blank")   # deletesheet! refuses to delete the only sheet
                XLSX.deletesheet!(xf, "Sheet1")
            end
            SAVE_FILES && save_outfile(p)

            r = ZipReader(read(p))
            names = zip_names(r)

            @test !("xl/tables/table1.xml" in names)
            @test !("xl/comments1.xml" in names)
            @test !("xl/drawings/vmlDrawing1.vml" in names)
            @test !("xl/drawings/drawing1.xml" in names)
            @test !("xl/media/image1.png" in names)
            @test !("xl/worksheets/_rels/sheet1.xml.rels" in names)
            @test !("xl/worksheets/sheet1.xml" in names)

            ctypes = String(zip_readentry(r, "[Content_Types].xml"))
            @test !occursin("tables/table1.xml", ctypes)
            @test !occursin("comments1.xml", ctypes)
            @test !occursin("drawings/drawing1.xml", ctypes)

            XLSX.openxlsx(p) do xf2
                @test XLSX.sheetnames(xf2) == ["Blank"]
            end
            isfile(p) && rm(p)
        end

        @testset "copysheet! strips table/comments/VML but still duplicates the drawing" begin
            src = joinpath(data_directory, "TableCommentsVML.xlsx")
            p = "copyadd_copysheet_strips_table_comments.xlsx"
            cp(src, p; force=true)

            XLSX.openxlsx(p; mode="rw") do xf
                XLSX.copysheet!(xf["Sheet1"], "Copy")
            end
            SAVE_FILES && save_outfile(p)

            zipbytes = read(p)
            r = ZipReader(zipbytes)
            names = zip_names(r)

            source_path = sheet_xml_path(zipbytes, "Sheet1")
            copy_path = sheet_xml_path(zipbytes, "Copy")
            source_xml = String(zip_readentry(r, source_path))
            copy_xml = String(zip_readentry(r, copy_path))

            # Stripped from the copy...
            @test !occursin("tableParts", copy_xml)
            @test !occursin("legacyDrawing", copy_xml)

            # ...but the original is untouched (catches copynode() aliasing children
            # between the original and the copy rather than deep-copying).
            @test occursin("tableParts", source_xml)
            @test occursin("legacyDrawing", source_xml)

            # The underlying parts still belong to the original sheet.
            @test "xl/tables/table1.xml" in names
            @test "xl/comments1.xml" in names
            @test "xl/drawings/vmlDrawing1.vml" in names

            # The drawing/image IS carried over — duplicated and relinked, not shared.
            m = match(r"<drawing r:id=\"(rId\d+)\"", copy_xml)
            @test m !== nothing

            if m !== nothing
                copy_rels_path = let (d, f) = rsplit(copy_path, "/"; limit=2)
                    "$d/_rels/$f.rels"
                end
                @test copy_rels_path in names

                if copy_rels_path in names
                    copy_rels_xml = String(zip_readentry(r, copy_rels_path))
                    dt = match(Regex("<Relationship Id=\"$(m.captures[1])\"[^>]*Target=\"([^\"]+)\""), copy_rels_xml)
                    @test dt !== nothing
                    if dt !== nothing
                        new_drawing_path = "xl/" * replace(dt.captures[1], "../" => "")
                        @test new_drawing_path in names
                        @test new_drawing_path != "xl/drawings/drawing1.xml"
                    end

                    # The copy's rels must not have picked up comments/VML relationships
                    # it has no corresponding sheet-XML reference for anymore.
                    @test !occursin("relationships/comments", copy_rels_xml)
                    @test !occursin("relationships/vmlDrawing", copy_rels_xml)
                end
            end

            XLSX.openxlsx(p) do xf2
                @test XLSX.sheetnames(xf2) == ["Sheet1", "Copy"]
            end
            isfile(p) && rm(p)
        end

    end

    @testset "editing a template: renaming, editing cells and adding sheets" begin
        f = XLSX.open_xlsx_template(joinpath(data_directory, "general.xlsx"))
        s = f["general"]
        @test_throws XLSX.XLSXError s["A1"] = :sym
        XLSX.renamesheet!(s, "general") # no-op
        @test_throws XLSX.XLSXError XLSX.renamesheet!(s, "table") # name is taken
        XLSX.renamesheet!(s, "renamed_sheet") # retain old function to avoid breaking change
        @test s.name == "renamed_sheet"
        s["A1"] = "Hey You!"
        s["B1"] = "Out there in the cold..."
        s["A2"] = "Getting lonely getting old..."
        s["B2"] = "Can you feel me?"
        s["A3"] = 1000
        s["B3"] = 99.99

        # create a new sheet
        s = XLSX.addsheet!(f, "my_new_sheet_1")
        s = XLSX.addsheet!(f, "my_new_sheet_2")
        s["B1"] = "This is a new sheet"
        s["B2"] = "This is a new sheet"
        s = XLSX.addsheet!(f)
        s["B1"] = "unnamed sheet"

        XLSX.writexlsx("general_copy_2.xlsx", f, overwrite=true)
        SAVE_FILES && save_outfile("general_copy_2.xlsx")
        @test isfile("general_copy_2.xlsx")

        XLSX.openxlsx("general_copy_2.xlsx") do f
            s = f["renamed_sheet"]
            @test s["A1"] == "Hey You!"
            @test s["B1"] == "Out there in the cold..."
            @test s["A2"] == "Getting lonely getting old..."
            @test s["B2"] == "Can you feel me?"
            @test s["A3"] == 1000
            @test s["B3"] == 99.99
            f["my_new_sheet_1"]
            @test f["my_new_sheet_2"]["B1"] == "This is a new sheet"
            @test f["my_new_sheet_2"]["B2"] == "This is a new sheet"
            @test f["Sheet1"]["B1"] == "unnamed sheet"
        end

        isfile("general_copy_2.xlsx") && rm("general_copy_2.xlsx")
    end

    @testset "renaming and adding sheets updates docProps/app.xml (issue #428)" begin # issue #428

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
            p = "testfiles_plain_worksheets_rename_add.xlsx"
            cp(src, p; force=true)

            XLSX.openxlsx(p; mode="rw") do xf
                XLSX.renamesheet!(xf["Sheet1"], "Renamed")
                XLSX.addsheet!(xf, "Added")
            end
            SAVE_FILES && save_outfile(p)

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

            isfile(p) && rm(p)
        end

        @testset "chartsheet workbook: Worksheets/Charts categories stay separate and correctly counted" begin
            src = joinpath(data_directory, "Chartsheet.xlsx")
            p = "testfiles_chartsheet_workbook_worksheets.xlsx"
            cp(src, p; force=true)

            XLSX.openxlsx(p; mode="rw") do xf
                XLSX.renamesheet!(xf["Sheet1"], "RenamedData")
                XLSX.addsheet!(xf, "ExtraData")
            end
            SAVE_FILES && save_outfile(p)

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

            isfile(p) && rm(p)
        end

    end
end

@testset "review fixes: copysheet! shared formula ids" begin
    f = XLSX.newxlsx("S"); s = f[1]
    s["A1:A4"] = 1
    XLSX.setFormula(s, "B1:B4", "=A1+1")
    c = XLSX.copysheet!(s, "C")
    XLSX.setFormula(c, "D1:D4", "=A1*3")
    ids = [f.id for (k, f) in f.workbook.formulas if k.sheet == "C" && f isa XLSX.ReferencedFormula]
    @test allunique(ids)
    SAVE_FILES && save_outfile(f)
end

@testset "movesheet!" begin
    function abc()
        xf = XLSX.newxlsx("A")
        XLSX.addsheet!(xf, "B")
        XLSX.addsheet!(xf, "C")
        for n in ("A", "B", "C")
            xf[n]["A1"] = n
        end
        return xf
    end
    book_view(x, key) = parse(Int, XLSX.XML.attributes(first(XLSX.elements_with_tag(first(XLSX.elements_with_tag(
        XLSX.xml_root_element(XLSX.xmlroot(x, "xl/workbook.xml")), "bookViews")), "workbookView")))[key])

    @testset "before and after, by name, position and Worksheet" begin
        xf = abc()
        @test XLSX.movesheet!(xf, "C"; before="A") === xf
        @test XLSX.sheetnames(xf) == ["C", "A", "B"]
        XLSX.movesheet!(xf, 1; after=3)
        @test XLSX.sheetnames(xf) == ["A", "B", "C"]
        XLSX.movesheet!(xf["A"]; after=xf["B"])
        @test XLSX.sheetnames(xf) == ["B", "A", "C"]
        XLSX.movesheet!(XLSX.get_workbook(xf), "C"; before=1)
        @test XLSX.sheetnames(xf) == ["C", "B", "A"]
        XLSX.movesheet!(xf, "C"; after=XLSX.sheetcount(xf))
        @test XLSX.sheetnames(xf) == ["B", "A", "C"]
        SAVE_FILES && save_outfile(xf)
    end

    @testset "no-op moves" begin
        xf = abc()
        XLSX.movesheet!(xf, "B"; before="B")
        XLSX.movesheet!(xf, "B"; before="C")
        XLSX.movesheet!(xf, "B"; after="A")
        XLSX.movesheet!(xf, "A"; before=1)
        @test XLSX.sheetnames(xf) == ["A", "B", "C"]
    end

    @testset "errors" begin
        xf = abc()
        @test_throws XLSX.XLSXError XLSX.movesheet!(xf, "A")
        @test_throws XLSX.XLSXError XLSX.movesheet!(xf, "A"; before="B", after="C")
        @test_throws XLSX.XLSXError XLSX.movesheet!(xf, "nosuchsheet"; before="A")
        @test_throws XLSX.XLSXError XLSX.movesheet!(xf, "A"; after="nosuchsheet")
        @test_throws XLSX.XLSXError XLSX.movesheet!(xf, "A"; after=4)
        @test_throws XLSX.XLSXError XLSX.movesheet!(xf, "A"; after=0)
        @test_throws XLSX.XLSXError XLSX.movesheet!(xf, "A"; after=big(2)^70)
        @test_throws XLSX.XLSXError XLSX.movesheet!(xf, "A"; before=true)
        @test_throws XLSX.XLSXError XLSX.movesheet!(xf, true; before="C")
        @test_throws XLSX.XLSXError XLSX.movesheet!(xf, "A"; after=abc()["B"])  # another workbook's sheet
        @test XLSX.sheetnames(xf) == ["A", "B", "C"]

        io = IOBuffer()
        XLSX.writexlsx(io, xf)
        ro = XLSX.readxlsx(seekstart(io))
        @test_throws XLSX.XLSXError XLSX.movesheet!(ro, "C"; before="A")
    end

    @testset "round trip keeps cells and sheet-scoped names with their sheets" begin
        xf = abc()
        XLSX.addDefinedName(xf["A"], "onA", "A1")
        XLSX.addDefinedName(xf["C"], "onC", "A1")
        XLSX.movesheet!(xf, "A"; after="C")
        @test XLSX.sheetnames(xf) == ["B", "C", "A"]

        io = IOBuffer()
        XLSX.writexlsx(io, xf)
        g = XLSX.readxlsx(seekstart(io))
        @test XLSX.sheetnames(g) == ["B", "C", "A"]
        for n in ("A", "B", "C")
            @test g[n]["A1"] == n
        end
        @test XLSX.is_worksheet_defined_name(g["A"], "onA")
        @test XLSX.is_worksheet_defined_name(g["C"], "onC")
        @test !XLSX.is_worksheet_defined_name(g["B"], "onA")
        @test !XLSX.is_worksheet_defined_name(g["C"], "onA")
        @test XLSX.get_defined_name_value(g["A"], "onA") == XLSX.SheetCellRef("A!A1")
        SAVE_FILES && save_outfile(xf)
    end

    @testset "active tab and first sheet stay on the same sheet" begin
        xf = abc()
        XLSX.xmlroot(xf, "xl/workbook.xml")
        wv = first(XLSX.elements_with_tag(first(XLSX.elements_with_tag(
            XLSX.xml_root_element(XLSX.xmlroot(xf, "xl/workbook.xml")), "bookViews")), "workbookView"))
        wv["activeTab"] = "1"                                                  # B is active
        wv["firstSheet"] = "2"                                                 # C is first shown
        XLSX.movesheet!(xf, "B"; before="A")                                   # B, A, C
        @test book_view(xf, "activeTab") == 0
        @test book_view(xf, "firstSheet") == 2
        XLSX.movesheet!(xf, "C"; before=1)                                     # C, B, A
        @test book_view(xf, "activeTab") == 1
        @test book_view(xf, "firstSheet") == 0

        io = IOBuffer()
        XLSX.writexlsx(io, xf)
        g = XLSX.opentemplate(seekstart(io))
        @test book_view(g, "activeTab") == 1
        SAVE_FILES && save_outfile(io)
    end

    @testset "chartsheet workbook" begin
        p = "testfiles_movesheet_chartsheet.xlsx"
        cp(joinpath(data_directory, "Chartsheet.xlsx"), p; force=true)
        names = XLSX.openxlsx(p; mode="rw") do xf
            ch = only(filter(n -> XLSX.is_chartsheet(XLSX.get_workbook(xf), n), XLSX.sheetnames(xf)))
            XLSX.movesheet!(xf, ch; before=1)
            XLSX.sheetnames(xf)
        end
        XLSX.openxlsx(p) do xf2
            @test XLSX.sheetnames(xf2) == names
            @test XLSX.is_chartsheet(XLSX.get_workbook(xf2), names[1])
        end
        SAVE_FILES && save_outfile(p)
        isfile(p) && rm(p)
    end
end
