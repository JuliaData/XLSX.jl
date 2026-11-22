_root_of(s::AbstractString) = XLSX.xml_root_element(parse(s, XLSX.XML.Node))

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

    @test XLSX.hassheet(f2, "newSheet")
    @test XLSX.getImages(f2["newSheet"]) == [(sheet = "newSheet", media_name = "image1.jpg", from = "B17", to = "E27")]

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
    @test replace.(DataFrames.names(df), "\r\n" => "\n") == String[
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
        "Spread Table\n1 = Part A\n3 = Part B\n5 = Part D\n4 & 6 = Part E",
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

@testset "namespace resolution" begin
    @testset "get_default_namespace - single prefixed namespace" begin
        r = _root_of("""<x:workbook xmlns:x="$_TRANS_MAIN"/>""")
        @test XLSX.get_default_namespace(r) == _TRANS_MAIN
    end

    @testset "get_default_namespace - unprefixed default preferred" begin
        r = _root_of("""<workbook xmlns="$_TRANS_MAIN" xmlns:r="http://example.com/r"/>""")
        @test XLSX.get_default_namespace(r) == _TRANS_MAIN
    end

    @testset "get_default_namespace - prefixed spreadsheet ns as fallback" begin
        # No unprefixed default (issues #380/#362/#267/#170)
        r = _root_of("""<x:workbook xmlns:x="$_TRANS_MAIN" xmlns:r="http://example.com/r"/>""")
        @test XLSX.get_default_namespace(r) == _TRANS_MAIN
    end

    @testset "get_default_namespace - none found errors" begin
        r = _root_of("""<thing xmlns:a="http://example.com/a" xmlns:b="http://example.com/b"/>""")
        @test_throws XLSX.XLSXError XLSX.get_default_namespace(r)
    end

    @testset "get_default_namespace_prefix - no namespaces at all" begin
        r = _root_of("<thing/>")
        @test isnothing(XLSX.get_default_namespace_prefix(r))

        r2 = _root_of("""<thing xmlns:a="http://example.com/a" xmlns:b="http://example.com/b"/>""")
        @test isnothing(XLSX.get_default_namespace_prefix(r2))
    end

    @testset "build_ns_dict! - part still held as a raw String" begin
        # Covers the `val isa String` branch: a read part whose data hasn't been
        # parsed yet has its prefix sniffed from the raw text instead.
        xf = XLSX.newxlsx()
        f = "xl/styles.xml"
        @test haskey(xf.files, f)

        xf.data[f] = """<x:styleSheet xmlns:x="$_TRANS_MAIN"/>"""
        delete!(xf.namespace, f)

        XLSX.build_ns_dict!(xf)
        @test xf.namespace[f] == "x"
        @test xf.data[f] isa String   # left unparsed by build_ns_dict!
    end

    @testset "_get_ns_prefix_from_string - default vs prefixed vs absent" begin
        @test isnothing(XLSX._get_ns_prefix_from_string(nothing))
        @test isnothing(XLSX._get_ns_prefix_from_string("""<worksheet xmlns="$_TRANS_MAIN"/>"""))
        @test XLSX._get_ns_prefix_from_string("""<x:worksheet xmlns:x="$_TRANS_MAIN"/>""") == "x"
        @test isnothing(XLSX._get_ns_prefix_from_string("""<worksheet xmlns="http://example.com/other"/>"""))
    end

    @testset "get_sst_prefix - prefixed shared strings namespace" begin
        xf = XLSX.newxlsx()
        sh = xf[1]

        xf.namespace["xl/sharedStrings.xml"] = "x"
        @test XLSX.get_sst_prefix(sh) == "x:"

        xf.namespace["xl/sharedStrings.xml"] = nothing
        @test XLSX.get_sst_prefix(sh) == ""
        SAVE_FILES && save_outfile(xf)
    end

    @testset "prefixed attribute lookup" begin
        blip = parse("""<a:blip xmlns:rel="$(XLSX.NS_R)" rel:embed="rId7" cstate="print"/>""", XML.Node)[end]
        @test XLSX.get_prefixed_attr(blip, "embed") == "rId7"
        @test XLSX.get_prefixed_attr(blip, "cstate") === nothing   # unprefixed: skipped by design
        @test XLSX.get_prefixed_attr(blip, "id") === nothing
        # The attribute collection itself, as the cursor paths have it
        @test XLSX.get_prefixed_attr(XML.attributes(blip), "embed") == "rId7"
        @test XLSX.get_prefixed_attr(Dict("relationships:id" => "rId3"), "id") == "rId3"
        @test XLSX.get_prefixed_attr(nothing, "id") === nothing
        # A namespace declaration is not an attribute: `xmlns:id` declares a prefix `id`
        @test XLSX.get_prefixed_attr(Dict("xmlns:id" => XLSX.NS_R), "id") === nothing
    end
end

# The prefixes used in `xml` (on elements and attributes) that no enclosing element binds.
function _unbound_prefixes(xml::AbstractString)
    out = Set{String}()
    function walk(n, scope)
        XML.nodetype(n) == XML.Element || return
        atts = something(XML.attributes(n), Dict{String,String}())
        scope = union(scope, Set(k[7:end] for k in keys(atts) if startswith(k, "xmlns:")))
        for name in Iterators.flatten(((XML.tag(n),), (k for k in keys(atts) if !startswith(k, "xmlns"))))
            i = findfirst(':', name)
            isnothing(i) && continue
            p = name[1:prevind(name, i)]
            p in scope || p == "xml" || push!(out, p)
        end
        foreach(c -> walk(c, scope), something(XML.children(n), []))
    end
    walk(XLSX.xml_root_element(parse(String(xml), XML.Node)), Set{String}())
    return out
end
_zip_part(file, part) = String(zip_readentry(ZipReader(read(file)), part))

# Each copy gets its own name, so that patch_xlsx_entry's output doesn't clash.
function _patched(src, entry, f; as)
    cp(src, as; force=true)
    out = patch_xlsx_entry(as, entry, f)
    rm(as)
    return out
end

@testset "relationship prefix other than r (#507)" begin
    NS_R = XLSX.NS_R
    XLSX.writetable("plain507.xlsx", Any[[1, 2, 3], ["x", "y", "z"]], ["a", "b"]; overwrite=true)
    relabel(s) = replace(s, " r:id=\"" => " xmlns:relationships=\"$NS_R\" relationships:id=\"")

    @testset "workbook <sheet>: $(variant)" for variant in ("root keeps xmlns:r", "only relationships:")
        drop_root = variant == "only relationships:"
        f = _patched("plain507.xlsx", "xl/workbook.xml",
                     s -> drop_root ? replace(relabel(s), " xmlns:r=\"$NS_R\"" => "") : relabel(s);
                     as = drop_root ? "wb507b.xlsx" : "wb507a.xlsx")
        wb_xml = _zip_part(f, "xl/workbook.xml")
        @test occursin("relationships:id=", wb_xml)
        @test occursin("xmlns:r=", wb_xml) == !drop_root

        xf = XLSX.readxlsx(f)
        @test XLSX.sheetnames(xf) == ["Sheet1"]
        @test xf[1]["A2"] == 1
        @test !XLSX.is_chartsheet(XLSX.get_workbook(xf), "Sheet1")
        dt = XLSX.readtable(f, 1)
        @test dt.column_labels == [:a, :b]
        @test dt.data[2] == ["x", "y", "z"]

        # Writing it back: the new <sheet> elements use `r:id`, so `r` must be bound.
        xf = XLSX.openxlsx(f, mode="rw")
        xf[1]["C1"] = "new"
        XLSX.writexlsx("out507.xlsx", xf, overwrite=true)
        @test isempty(_unbound_prefixes(_zip_part("out507.xlsx", "xl/workbook.xml")))
        xf2 = XLSX.readxlsx("out507.xlsx")
        @test XLSX.sheetnames(xf2) == ["Sheet1"]
        @test xf2[1]["C1"] == "new"
        @test xf2[1]["B4"] == "z"
    end

    @testset "<tablePart> under another prefix" begin
        f = _patched(joinpath(data_directory, "two_tables.xlsx"), "xl/worksheets/sheet1.xml",
                     s -> replace(s, "<tablePart r:id=" => "<tablePart xmlns:rel=\"$NS_R\" rel:id=");
                     as = "tables507.xlsx")
        @test occursin("rel:id=", _zip_part(f, "xl/worksheets/sheet1.xml"))
        XLSX.openxlsx(f) do xf   # read-only: the cursor path
            @test sort([t.name for t in XLSX.tables(xf["Sheet1"])]) == ["Age_height", "IO_Table"]
            @test sort(XLSX.get_worksheet_table_rids(xf, xf["Sheet1"])) == ["rId1", "rId2"]
        end
        xf = XLSX.openxlsx(f, mode="rw")   # parsed tree
        @test XLSX.table(xf["Sheet1"], "IO_Table").ref == XLSX.CellRange("A1:C8")
        XLSX.deletetable!(xf["Sheet1"], "Age_height")
        XLSX.writexlsx("tables507_out.xlsx", xf, overwrite=true)
        XLSX.openxlsx("tables507_out.xlsx") do xf2
            @test [t.name for t in XLSX.tables(xf2["Sheet1"])] == ["IO_Table"]
        end
    end

    @testset "external link under other prefixes" begin
        # The external book's `r:id`s and the `xxl21:` alternateUrls, renamed.
        f1 = _patched(joinpath(data_directory, "linked-1.xlsx"), "xl/externalLinks/externalLink1.xml",
                      s -> replace(s, "xmlns:r=" => "xmlns:rel=", " r:id=" => " rel:id=", "xxl21" => "x21");
                      as = "linked507.xlsx")
        f = _patched(f1, "xl/workbook.xml",
                     s -> replace(s, "<externalReference r:id=" => "<externalReference xmlns:rel=\"$NS_R\" rel:id=");
                     as = "linked507b.xlsx")
        @test occursin("rel:id=", _zip_part(f, "xl/workbook.xml"))
        @test occursin("<x21:alternateUrls>", _zip_part(f, "xl/externalLinks/externalLink1.xml"))
        xf = XLSX.openxlsx(f, mode="rw")
        @test XLSX.getFormula(xf[1], "A1") == "=[1]Sheet1!\$A\$1"
        @test occursin("linked-2.xlsx]", XLSX.getFormula(xf[1], "A1"; get_external_refs=true))
    end
end

@testset "writing into parts that bind namespaces differently" begin
    NS_R = XLSX.NS_R
    XLSX.writetable("plain_ns.xlsx", Any[[1, 2, 3], ["x", "y", "z"]], ["a", "b"]; overwrite=true)

    @testset "addtable! to a sheet without xmlns:r" begin
        f = _patched("plain_ns.xlsx", "xl/worksheets/sheet1.xml", s -> replace(s, " xmlns:r=\"$NS_R\"" => ""); as = "nor.xlsx")
        @test !occursin("xmlns:r=", _zip_part(f, "xl/worksheets/sheet1.xml"))
        xf = XLSX.openxlsx(f, mode="rw")
        XLSX.addtable!(xf[1], XLSX.CellRange("A1:B4"); name="Table_507")
        XLSX.writexlsx("nor_out.xlsx", xf, overwrite=true)
        sheet_xml = _zip_part("nor_out.xlsx", "xl/worksheets/sheet1.xml")
        @test occursin("<tablePart r:id=", sheet_xml)
        @test isempty(_unbound_prefixes(sheet_xml))
        XLSX.openxlsx("nor_out.xlsx") do xf2
            @test [t.name for t in XLSX.tables(xf2[1])] == ["Table_507"]
        end
    end

    @testset "addImage and addTextBox into a drawing with its own prefixes" begin
        xf = XLSX.openxlsx("plain_ns.xlsx", mode="rw")
        XLSX.addImage(xf[1], "D2", joinpath(data_directory, "track_start.jpg"))
        XLSX.writexlsx("img_ns.xlsx", xf, overwrite=true)
        # As another writer might: `d:` and `x:` for xdr and a, and `rel:` declared on the
        # root for the relationship namespace, with no `r` anywhere.
        f = _patched("img_ns.xlsx", "xl/drawings/drawing1.xml", s -> begin
            s = replace(s, " xmlns:r=\"$NS_R\"" => "")
            s = replace(s, "xdr:" => "d:", "xmlns:xdr=" => "xmlns:d=", "<a:" => "<x:", "</a:" => "</x:",
                           "xmlns:a=" => "xmlns:x=", "r:embed=" => "rel:embed=")
            replace(s, "<d:wsDr " => "<d:wsDr xmlns:rel=\"$NS_R\" "; count=1)
        end; as = "img_ns2.xlsx")
        d_xml = _zip_part(f, "xl/drawings/drawing1.xml")
        @test !occursin("xdr:", d_xml) && !occursin("xmlns:r=", d_xml)
        @test isempty(_unbound_prefixes(d_xml))

        xf = XLSX.openxlsx(f, mode="rw")
        XLSX.addImage(xf[1], "J2", joinpath(data_directory, "track_start.jpg"))
        XLSX.addTextBox(xf[1], "B8:E12", "note")
        XLSX.writexlsx("img_ns_out.xlsx", xf, overwrite=true)
        d_out = _zip_part("img_ns_out.xlsx", "xl/drawings/drawing1.xml")
        @test isempty(_unbound_prefixes(d_out))
        @test !occursin("xdr:", d_out)   # the new anchors took the drawing's own prefixes
        @test length(XLSX.getImages(XLSX.readxlsx("img_ns_out.xlsx")[1])) == 2
    end

    @testset "copysheet! keeps one uid under the sheet's own prefix" begin
        NS_XR = "http://schemas.microsoft.com/office/spreadsheetml/2014/revision"
        f = _patched("plain_ns.xlsx", "xl/worksheets/sheet1.xml",
                     s -> replace(s, "xmlns:xr=\"$NS_XR\"" => "xmlns:rev=\"$NS_XR\"", " xr:uid=" => " rev:uid=",
                                     "mc:Ignorable=\"x14ac xr " => "mc:Ignorable=\"x14ac rev ");
                     as = "uid.xlsx")
        xf = XLSX.openxlsx(f, mode="rw")
        XLSX.copysheet!(xf[1], "Copy")
        XLSX.writexlsx("uid_out.xlsx", xf, overwrite=true)
        sheets = filter(n -> startswith(n, "xl/worksheets/sheet") && endswith(n, ".xml"), zip_names(ZipReader(read("uid_out.xlsx"))))
        @test length(sheets) == 2
        uids = String[]
        for p in sheets
            root = XLSX.xml_root_element(parse(_zip_part("uid_out.xlsx", p), XML.Node))
            ks = [k for k in keys(XML.attributes(root)) if occursin(':', k) && !startswith(k, "xmlns") && endswith(k, ":uid")]
            @test length(ks) == 1
            push!(uids, XML.attributes(root)[only(ks)])
        end
        @test allunique(uids)
        @test XLSX.sheetnames(XLSX.readxlsx("uid_out.xlsx")) == ["Sheet1", "Copy"]
    end
    for f in ("plain507.xlsx", "out507.xlsx", "tables507_out.xlsx", "plain_ns.xlsx", "nor_out.xlsx",
              "img_ns.xlsx", "img_ns_out.xlsx", "uid_out.xlsx")
        isfile(f) && rm(f)
    end
    foreach(f -> startswith(f, "patched_") && rm(f), readdir())
end
