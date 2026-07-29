# Helper: compare two Table structs field-by-field (incl. nested TableStyleInfo).
function tables_equal(a::XLSX.Table, b::XLSX.Table)
    a.id == b.id || return false
    a.name == b.name || return false
    a.display_name == b.display_name || return false
    a.ref == b.ref || return false
    a.columns == b.columns || return false
    a.has_totals_row == b.has_totals_row || return false

    (a.style === nothing) != (b.style === nothing) && return false
    if a.style !== nothing
        a.style.name == b.style.name || return false
        a.style.show_first_column == b.style.show_first_column || return false
        a.style.show_last_column == b.style.show_last_column || return false
        a.style.show_row_stripes == b.style.show_row_stripes || return false
        a.style.show_column_stripes == b.style.show_column_stripes || return false
    end
    return true
end

# Internal helper (test-only): read back the `totalsRowFunction`/`totalsRowLabel`
# attributes for a given column directly from the table's own XML part, since
# `XLSX.Table` (by design) only exposes `has_totals_row::Bool`, not per-column
# totals settings. Reuses the same internal accessors `settotals!` itself uses.
function _totals_col_attrs(sheet::XLSX.Worksheet, table_name::AbstractString, col_name::AbstractString)
    xf = XLSX.get_xlsxfile(sheet)
    t = XLSX.table(sheet, table_name)

    # Deliberately avoid touching the worksheet's own XML part (`sheet_path`)
    # here. In a read-only reopened file, `xf.data[sheet_path]` is kept as a
    # raw, unparsed string so the lazy per-sheet cache fill (triggered by
    # `eachrow`, on first cell access) can parse it on demand. Calling
    # `get_xml_data` on it directly — as an earlier version of this helper
    # did — permanently converts it to a parsed Node and breaks that lazy
    # fill for the rest of the session. `get_worksheet_table_rids` is the
    # existing, read-only-safe accessor Phase 1 built for exactly this: it
    # never mutates `xf.data[sheet_path]`.
    table_path = nothing
    for r_id in XLSX.get_worksheet_table_rids(xf, sheet)
        path = XLSX.get_worksheet_relationship_target(xf, sheet, r_id)
        if XLSX.get_attr(XLSX.root_element(XLSX.get_xml_data(xf, path)), "name") == table_name
            table_path = path
            break
        end
    end
    isnothing(table_path) && error("Internal test helper error: table part not found for `$table_name`.")

    # The table part itself is always fully parsed regardless of cache mode
    # (tables aren't touched by the streaming/eachrow lazy-fill machinery),
    # so `get_xml_data` here is safe.
    table_doc = XLSX.get_xml_data(xf, table_path)
    i, j = XLSX.get_idces(table_doc, "table", "tableColumns")
    col_idx = findfirst(==(col_name), t.columns)
    col_node = collect(XLSX.xml_elements(table_doc[i][j]))[col_idx]

    func  = XLSX.get_attr(col_node, "totalsRowFunction", "")
    label = XLSX.get_attr(col_node, "totalsRowLabel", "")
    return (func == "" ? nothing : func), (label == "" ? nothing : label)

end

@testset "Excel Tables" begin

    @testset "tables(sheet)" begin
        XLSX.openxlsx("data/two_tables.xlsx") do xf
            sh = xf["Sheet1"]
            tbls = XLSX.tables(sh)

            @test length(tbls) == 2
            @test all(t -> t isa XLSX.Table, tbls)

            names = [t.name for t in tbls]
            @test "IO_Table" in names
            @test "Age_height" in names
        end
    end

    @testset "table(sheet, name) - IO_Table" begin
        XLSX.openxlsx("data/two_tables.xlsx") do xf
            sh = xf["Sheet1"]
            t = XLSX.table(sh, "IO_Table")

            @test t.id == 1
            @test t.name == "IO_Table"
            @test t.display_name == "IO_Table"
            @test t.ref == XLSX.CellRange("A1:C8")
            @test t.columns == ["id", "input", "output"]
            @test t.has_totals_row == false

            @test !isnothing(t.style)
            @test t.style.name == "TableStyleMedium2"
            @test t.style.show_row_stripes == true
            @test t.style.show_column_stripes == false
            @test t.style.show_first_column == false
            @test t.style.show_last_column == false
        end
    end

    @testset "table(sheet, name) - Age_height" begin
        XLSX.openxlsx("data/two_tables.xlsx") do xf
            sh = xf["Sheet1"]
            t = XLSX.table(sh, "Age_height")

            @test t.id == 2
            @test t.name == "Age_height"
            @test t.display_name == "Age_height"
            @test t.ref == XLSX.CellRange("E1:G5")
            @test t.columns == ["name", "age", "height"]
            @test t.has_totals_row == false
        end
    end

    @testset "table(sheet, id)" begin
        XLSX.openxlsx("data/two_tables.xlsx") do xf
            sh = xf["Sheet1"]

            t1 = XLSX.table(sh, 1)
            @test t1.name == "IO_Table"

            t2 = XLSX.table(sh, 2)
            @test t2.name == "Age_height"
        end
    end

    @testset "table(sheet, name/id) - not found" begin
        XLSX.openxlsx("data/two_tables.xlsx") do xf
            sh = xf["Sheet1"]

            @test_throws KeyError XLSX.table(sh, "NoSuchTable")
            @test_throws KeyError XLSX.table(sh, 999)
        end
    end

    @testset "table(sheet, name) - totals row (Sheet2, with_total)" begin
        XLSX.openxlsx("data/two_tables.xlsx") do xf
            sh = xf["Sheet2"]
            t = XLSX.table(sh, "with_total")

            @test t.id == 3
            @test t.name == "with_total"
            @test t.ref == XLSX.CellRange("A3:C11")
            @test t.columns == ["start", "stop", "sin"]

            # This table signals its totals row via `totalsRowCount="1"`,
            # not `totalsRowShown` — regression check for that path.
            @test t.has_totals_row == true
        end
    end

    @testset "tables(sheet) - sheet with no tables" begin
        # blank.xlsx (used internally by XLSX.newxlsx) has a single sheet with no tables
        xf = XLSX.newxlsx()
        sh = xf[1]
        @test XLSX.tables(sh) == XLSX.Table[]
    end

    @testset "consistent under enable_cache=false" begin
        XLSX.openxlsx("data/two_tables.xlsx", enable_cache=false) do xf
            sh = xf["Sheet1"]
            tbls = XLSX.tables(sh)

            @test length(tbls) == 2
            @test XLSX.table(sh, "IO_Table").ref == XLSX.CellRange("A1:C8")
            @test XLSX.table(sh, "Age_height").ref == XLSX.CellRange("E1:G5")
        end
    end

    @testset "consistent under mode=rw" begin
        XLSX.openxlsx("data/two_tables.xlsx", mode="r", enable_cache=true) do xf
            sh = xf["Sheet1"]
            @test length(XLSX.tables(sh)) == 2
        end
    end


    # NOTE: writing always requires enable_cache=true (rw mode's own contract:
    # "Cache must be enabled for files in `write` mode"). `opentemplate` already
    # opens in rw/cache=true, so every write below satisfies that automatically.
    # `enable_cache=false` is only ever used for the *reopen-to-verify* step,
    # never for the write itself.

    @testset "no edits - tables survive save" begin
        original_tables = Dict{String,Vector{XLSX.Table}}()
        XLSX.openxlsx("data/two_tables.xlsx") do xf
            for sh_name in ("Sheet1", "Sheet2")
                original_tables[sh_name] = XLSX.tables(xf[sh_name])
            end
        end

        f = XLSX.opentemplate("data/two_tables.xlsx")
        SAVE_FILES && save_outfile(f)

        outfile = "two_tables_roundtrip_noedits.xlsx"
        isfile(outfile) && rm(outfile)
        XLSX.writexlsx(outfile, f, overwrite=true)
        SAVE_FILES && save_outfile(outfile)

        XLSX.openxlsx(outfile) do xf2
            for sh_name in ("Sheet1", "Sheet2")
                before = original_tables[sh_name]
                after  = XLSX.tables(xf2[sh_name])

                @test length(after) == length(before)
                for t_before in before
                    idx = findfirst(t -> t.name == t_before.name, after)
                    @test idx !== nothing
                    @test tables_equal(t_before, after[idx])
                end
            end
        end

        isfile(outfile) && rm(outfile)
    end

    @testset "unrelated cell edit elsewhere - tables still intact" begin
        f = XLSX.opentemplate("data/two_tables.xlsx")

        # Edit a cell well outside any table's range on Sheet1
        # (IO_Table is A1:C8, Age_height is E1:G5).
        f["Sheet1"]["Z100"] = "unrelated edit"
        SAVE_FILES && save_outfile(f)

        outfile = "two_tables_roundtrip_celledit.xlsx"
        isfile(outfile) && rm(outfile)
        XLSX.writexlsx(outfile, f, overwrite=true)
        SAVE_FILES && save_outfile(outfile)

        XLSX.openxlsx(outfile) do xf2
            @test xf2["Sheet1"]["Z100"] == "unrelated edit"

            sh1_tables = XLSX.tables(xf2["Sheet1"])
            @test length(sh1_tables) == 2
            io_table = XLSX.table(xf2["Sheet1"], "IO_Table")
            @test io_table.ref == XLSX.CellRange("A1:C8")
            @test io_table.columns == ["id", "input", "output"]

            sh2_tables = XLSX.tables(xf2["Sheet2"])
            @test length(sh2_tables) == 1
            wt = XLSX.table(xf2["Sheet2"], "with_total")
            @test wt.has_totals_row == true
            @test wt.ref == XLSX.CellRange("A3:C11")
        end

        isfile(outfile) && rm(outfile)
    end

    @testset "reopen after round trip under enable_cache=false" begin
        f = XLSX.opentemplate("data/two_tables.xlsx")
        SAVE_FILES && save_outfile(f)

        outfile = "two_tables_roundtrip_nocache.xlsx"
        isfile(outfile) && rm(outfile)
        XLSX.writexlsx(outfile, f, overwrite=true)
        SAVE_FILES && save_outfile(outfile)

        # Only the *reopen* is enable_cache=false here — the write above
        # still went through cache=true, as it must.
        XLSX.openxlsx(outfile, enable_cache=false) do xf2
            @test length(XLSX.tables(xf2["Sheet1"])) == 2
            @test length(XLSX.tables(xf2["Sheet2"])) == 1
            @test XLSX.table(xf2["Sheet2"], "with_total").has_totals_row == true
        end

        isfile(outfile) && rm(outfile)
    end

    @testset "addtable! - basic creation" begin
        f = XLSX.newxlsx()
        sh = f[1]

        sh["A1"] = "id"; sh["B1"] = "name"; sh["C1"] = "score"
        sh["A2"] = 1; sh["B2"] = "alice"; sh["C2"] = 10.5
        sh["A3"] = 2; sh["B3"] = "bob";   sh["C3"] = 20.0

        t = XLSX.addtable!(sh, "A1:C3"; name="MyTable")

        @test t isa XLSX.Table
        @test t.name == "MyTable"
        @test t.display_name == "MyTable"
        @test t.ref == XLSX.CellRange("A1:C3")
        @test t.columns == ["id", "name", "score"]
        @test t.has_totals_row == false
        @test isnothing(t.style)

        # confirm it's visible via the read API too
        @test length(XLSX.tables(sh)) == 1
        @test XLSX.table(sh, "MyTable").ref == XLSX.CellRange("A1:C3")
        @test XLSX.table(sh, 1).name == "MyTable"
    end

    @testset "addtable! - auto-generated name" begin
        f = XLSX.newxlsx()
        sh = f[1]
        sh["A1"] = "x"; sh["B1"] = "y"
        sh["A2"] = 1;   sh["B2"] = 2

        t = XLSX.addtable!(sh, "A1:B2")
        @test t.name == "Table1"

        sh["D1"] = "p"; sh["E1"] = "q"
        sh["D2"] = 1;   sh["E2"] = 2
        t2 = XLSX.addtable!(sh, "D1:E2")
        @test t2.name == "Table2"
    end

    @testset "addtable! - with style and totals row (empty row, no warning)" begin
        f = XLSX.newxlsx()
        sh = f[1]
        sh["A1"] = "qty"; sh["B1"] = "price"
        sh["A2"] = 3;     sh["B2"] = 9.99
        sh["A3"] = 5;     sh["B3"] = 4.50
        sh["A4"] = missing; sh["B4"] = missing  # totals row placeholder, genuinely empty

        t = @test_logs XLSX.addtable!(sh, "A1:B4"; name="Sales", style="TableStyleMedium9", has_totals_row=true)

        @test t.has_totals_row == true
        @test !isnothing(t.style)
        @test t.style.name == "TableStyleMedium9"
    end

    @testset "addtable! - has_totals_row with nonempty last row warns, still creates table" begin
        f = XLSX.newxlsx()
        sh = f[1]
        sh["A1"] = "qty";   sh["B1"] = "price"
        sh["A2"] = 3;       sh["B2"] = 9.99
        sh["A3"] = 5;       sh["B3"] = 4.50
        # last row already has real content — e.g. left over from writetable!,
        # or intentionally pre-authored totals content. Either way, `addtable!`
        # must NOT throw and must NOT alter these cells; it only warns.
        sh["A4"] = "Total"; sh["B4"] = 14.49

        t = @test_logs (:warn, r"already has content") match_mode=:any XLSX.addtable!(
            sh, "A1:B4"; name="Sales", has_totals_row=true)

        @test t.has_totals_row == true
        @test t.ref == XLSX.CellRange("A1:B4")

        # cell contents must be completely untouched by addtable!
        @test sh["A4"] == "Total"
        @test sh["B4"] == 14.49
    end

    @testset "addtable! - has_totals_row=false always treats last row as data, even if blank" begin
        f = XLSX.newxlsx()
        sh = f[1]
        sh["A1"] = "qty"; sh["B1"] = "price"
        sh["A2"] = 3;     sh["B2"] = 9.99
        sh["A3"] = missing; sh["B3"] = missing  # blank last row, but has_totals_row=false

        t = @test_logs XLSX.addtable!(sh, "A1:B3"; name="NoTotals", has_totals_row=false)

        @test t.has_totals_row == false
        @test t.ref == XLSX.CellRange("A1:B3")
    end

    @testset "addtable! - name collisions" begin
        f = XLSX.newxlsx()
        sh = f[1]
        sh["A1"] = "a"; sh["B1"] = "b"
        sh["A2"] = 1;   sh["B2"] = 2
        XLSX.addtable!(sh, "A1:B2"; name="Dup")

        sh["D1"] = "c"; sh["E1"] = "d"
        sh["D2"] = 1;   sh["E2"] = 2
        @test_throws XLSX.XLSXError XLSX.addtable!(sh, "D1:E2"; name="Dup")
    end

    @testset "addtable! - invalid header (empty cell)" begin
        f = XLSX.newxlsx()
        sh = f[1]
        sh["A1"] = "a"; sh["B1"] = missing
        sh["A2"] = 1;   sh["B2"] = 2
        @test_throws XLSX.XLSXError XLSX.addtable!(sh, "A1:B2")
    end

    @testset "addtable! - invalid header (duplicate names)" begin
        f = XLSX.newxlsx()
        sh = f[1]
        sh["A1"] = "same"; sh["B1"] = "same"
        sh["A2"] = 1;       sh["B2"] = 2
        @test_throws XLSX.XLSXError XLSX.addtable!(sh, "A1:B2")
    end

    @testset "addtable! - displayName validation" begin
        f = XLSX.newxlsx()
        sh = f[1]
        sh["A1"] = "a"; sh["B1"] = "b"
        sh["A2"] = 1;   sh["B2"] = 2
        @test_throws XLSX.XLSXError XLSX.addtable!(sh, "A1:B2"; name="has space")
    end

    @testset "addtable! - round trip (create -> save -> reopen)" begin
        f = XLSX.newxlsx()
        sh = f[1]
        sh["A1"] = "id"; sh["B1"] = "value"
        sh["A2"] = 1;    sh["B2"] = 100
        sh["A3"] = 2;    sh["B3"] = 200
        XLSX.addtable!(sh, "A1:B3"; name="RoundTripTable", style="TableStyleLight1")

        SAVE_FILES && save_outfile(f)

        outfile = "newtable_roundtrip.xlsx"
        isfile(outfile) && rm(outfile)
        XLSX.writexlsx(outfile, f, overwrite=true)
        SAVE_FILES && save_outfile(outfile)

        XLSX.openxlsx(outfile) do xf2
            sh2 = xf2[1]
            @test length(XLSX.tables(sh2)) == 1
            t = XLSX.table(sh2, "RoundTripTable")
            @test t.ref == XLSX.CellRange("A1:B3")
            @test t.columns == ["id", "value"]
            @test t.style.name == "TableStyleLight1"
        end

        XLSX.openxlsx(outfile, enable_cache=false) do xf2
            @test length(XLSX.tables(xf2[1])) == 1
            @test XLSX.table(xf2[1], "RoundTripTable").columns == ["id", "value"]
        end

        isfile(outfile) && rm(outfile)
    end

    @testset "deletetable! - by name, single table" begin
        f = XLSX.newxlsx()
        sh = f[1]
        sh["A1"] = "a"; sh["B1"] = "b"
        sh["A2"] = 1;   sh["B2"] = 2
        XLSX.addtable!(sh, "A1:B2"; name="ToDelete")

        @test length(XLSX.tables(sh)) == 1
        XLSX.deletetable!(sh, "ToDelete")
        @test length(XLSX.tables(sh)) == 0
        @test_throws KeyError XLSX.table(sh, "ToDelete")
    end

    @testset "deletetable! - by id" begin
        f = XLSX.newxlsx()
        sh = f[1]
        sh["A1"] = "a"; sh["B1"] = "b"
        sh["A2"] = 1;   sh["B2"] = 2
        t = XLSX.addtable!(sh, "A1:B2"; name="ById")

        XLSX.deletetable!(sh, t.id)
        @test length(XLSX.tables(sh)) == 0
    end

    @testset "deletetable! - one of several tables survives" begin
        f = XLSX.newxlsx()
        sh = f[1]

        sh["A1"] = "a"; sh["B1"] = "b"
        sh["A2"] = 1;   sh["B2"] = 2
        t1 = XLSX.addtable!(sh, "A1:B2"; name="Keep")

        sh["D1"] = "c"; sh["E1"] = "d"
        sh["D2"] = 3;   sh["E2"] = 4
        t2 = XLSX.addtable!(sh, "D1:E2"; name="Remove")

        @test length(XLSX.tables(sh)) == 2

        XLSX.deletetable!(sh, "Remove")

        remaining = XLSX.tables(sh)
        @test length(remaining) == 1
        @test remaining[1].name == "Keep"
        @test XLSX.table(sh, "Keep").ref == XLSX.CellRange("A1:B2")
        @test_throws KeyError XLSX.table(sh, "Remove")
    end

    @testset "deletetable! - survives round trip, remaining table intact" begin
        f = XLSX.newxlsx()
        sh = f[1]

        sh["A1"] = "a"; sh["B1"] = "b"
        sh["A2"] = 1;   sh["B2"] = 2
        XLSX.addtable!(sh, "A1:B2"; name="Keep")

        sh["D1"] = "c"; sh["E1"] = "d"
        sh["D2"] = 3;   sh["E2"] = 4
        XLSX.addtable!(sh, "D1:E2"; name="Remove")

        XLSX.deletetable!(sh, "Remove")

        SAVE_FILES && save_outfile(f)

        outfile = "deletetable_roundtrip.xlsx"
        isfile(outfile) && rm(outfile)
        XLSX.writexlsx(outfile, f, overwrite=true)
        SAVE_FILES && save_outfile(outfile)

        XLSX.openxlsx(outfile) do xf2
            sh2 = xf2[1]
            @test length(XLSX.tables(sh2)) == 1
            @test XLSX.table(sh2, "Keep").ref == XLSX.CellRange("A1:B2")
            @test_throws KeyError XLSX.table(sh2, "Remove")
        end

        isfile(outfile) && rm(outfile)
    end

    @testset "deletetable! - on existing fixture, other sheet's tables untouched" begin
        f = XLSX.opentemplate("data/two_tables.xlsx")
        sh1 = f["Sheet1"]
        sh2 = f["Sheet2"]

        @test length(XLSX.tables(sh1)) == 2
        @test length(XLSX.tables(sh2)) == 1

        XLSX.deletetable!(sh1, "IO_Table")

        @test length(XLSX.tables(sh1)) == 1
        @test XLSX.table(sh1, "Age_height").ref == XLSX.CellRange("E1:G5")
        @test_throws KeyError XLSX.table(sh1, "IO_Table")

        # Sheet2's table must be completely unaffected
        @test length(XLSX.tables(sh2)) == 1
        @test XLSX.table(sh2, "with_total").has_totals_row == true

        SAVE_FILES && save_outfile(f)

        outfile = "deletetable_fixture_roundtrip.xlsx"
        isfile(outfile) && rm(outfile)
        XLSX.writexlsx(outfile, f, overwrite=true)
        SAVE_FILES && save_outfile(outfile)

        XLSX.openxlsx(outfile) do xf2
            @test length(XLSX.tables(xf2["Sheet1"])) == 1
            @test XLSX.table(xf2["Sheet1"], "Age_height").columns == ["name", "age", "height"]
            @test_throws KeyError XLSX.table(xf2["Sheet1"], "IO_Table")

            @test length(XLSX.tables(xf2["Sheet2"])) == 1
            @test XLSX.table(xf2["Sheet2"], "with_total").ref == XLSX.CellRange("A3:C11")
        end

        isfile(outfile) && rm(outfile)
    end

    @testset "settotals! - adds a totals row where none existed" begin
        f = XLSX.newxlsx()
        sh = f[1]
        sh["A1"] = "qty";   sh["B1"] = "price"
        sh["A2"] = 3;       sh["B2"] = 9.99
        sh["A3"] = 5;       sh["B3"] = 4.50

        XLSX.addtable!(sh, "A1:B3"; name="Sales")
        @test XLSX.table(sh, "Sales").has_totals_row == false
        @test XLSX.table(sh, "Sales").ref == XLSX.CellRange("A1:B3")

        t = XLSX.settotals!(sh, "Sales", "price" => :sum)

        @test t.has_totals_row == true
        @test t.ref == XLSX.CellRange("A1:B4")  # extended by one row

        func, label = _totals_col_attrs(sh, "Sales", "price")
        @test func == "sum"
        @test isnothing(label)

        # formula cell has no cached value (per the no-calc-engine convention)
        @test ismissing(sh["B4"])
    end

    @testset "settotals! - errors if extension row already has content" begin
        f = XLSX.newxlsx()
        sh = f[1]
        sh["A1"] = "qty"; sh["B1"] = "price"
        sh["A2"] = 3;     sh["B2"] = 9.99
        sh["A3"] = "oops"; sh["B3"] = 1  # occupies what would become the totals row

        XLSX.addtable!(sh, "A1:B2"; name="Sales")  # table itself only spans A1:B2
        @test_throws XLSX.XLSXError XLSX.settotals!(sh, "Sales", "price" => :sum)
    end

    @testset "settotals! - label column" begin
        f = XLSX.newxlsx()
        sh = f[1]
        sh["A1"] = "name"; sh["B1"] = "score"
        sh["A2"] = "alice"; sh["B2"] = 10
        sh["A3"] = "bob";   sh["B3"] = 20

        XLSX.addtable!(sh, "A1:B3"; name="Scores")
        t = XLSX.settotals!(sh, "Scores", "name" => "Grand Total", "score" => :sum)

        @test t.has_totals_row == true

        func, label = _totals_col_attrs(sh, "Scores", "name")
        @test isnothing(func)
        @test label == "Grand Total"
        @test sh["A4"] == "Grand Total"  # literal value written, not just the attribute

        func2, label2 = _totals_col_attrs(sh, "Scores", "score")
        @test func2 == "sum"
        @test isnothing(label2)
    end

    @testset "settotals! - custom formula (:custom tuple)" begin
        f = XLSX.newxlsx()
        sh = f[1]
        sh["A1"] = "revenue"; sh["B1"] = "cost"; sh["C1"] = "margin"
        sh["A2"] = 100;       sh["B2"] = 60;     sh["C2"] = 40
        sh["A3"] = 200;       sh["B3"] = 90;     sh["C3"] = 110

        XLSX.addtable!(sh, "A1:C3"; name="PnL")
        t = XLSX.settotals!(sh, "PnL",
            "revenue" => :sum,
            "cost"    => :sum,
            "margin"  => (:custom, "PnL[revenue]-PnL[cost]"),
        )

        @test t.has_totals_row == true

        func, label = _totals_col_attrs(sh, "PnL", "margin")
        @test func == "custom"
        @test isnothing(label)
        @test ismissing(sh["C4"])  # formula cell, no cached value
    end

    @testset "settotals! - rejects malformed custom tuple" begin
        f = XLSX.newxlsx()
        sh = f[1]
        sh["A1"] = "a"; sh["B1"] = "b"
        sh["A2"] = 1;   sh["B2"] = 2
        XLSX.addtable!(sh, "A1:B2"; name="T")

        @test_throws XLSX.XLSXError XLSX.settotals!(sh, "T", "b" => (:notcustom, "SUM(1,2)"))
    end

    @testset "settotals! - unknown column name errors" begin
        f = XLSX.newxlsx()
        sh = f[1]
        sh["A1"] = "a"; sh["B1"] = "b"
        sh["A2"] = 1;   sh["B2"] = 2
        XLSX.addtable!(sh, "A1:B2"; name="T")

        @test_throws XLSX.XLSXError XLSX.settotals!(sh, "T", "nope" => :sum)
    end

    @testset "settotals! - unknown built-in function symbol errors" begin
        f = XLSX.newxlsx()
        sh = f[1]
        sh["A1"] = "a"; sh["B1"] = "b"
        sh["A2"] = 1;   sh["B2"] = 2
        XLSX.addtable!(sh, "A1:B2"; name="T")

        @test_throws XLSX.XLSXError XLSX.settotals!(sh, "T", "b" => :median)
    end

    @testset "settotals! - invalid value type errors" begin
        f = XLSX.newxlsx()
        sh = f[1]
        sh["A1"] = "a"; sh["B1"] = "b"
        sh["A2"] = 1;   sh["B2"] = 2
        XLSX.addtable!(sh, "A1:B2"; name="T")

        @test_throws XLSX.XLSXError XLSX.settotals!(sh, "T", "b" => 3.14)
    end

    @testset "settotals! - only mentioned columns change; others untouched" begin
        f = XLSX.newxlsx()
        sh = f[1]
        sh["A1"] = "a"; sh["B1"] = "b"; sh["C1"] = "c"
        sh["A2"] = 1;   sh["B2"] = 2;   sh["C2"] = 3

        XLSX.addtable!(sh, "A1:C2"; name="T")
        XLSX.settotals!(sh, "T", "a" => :sum, "b" => :average)

        func_a, _ = _totals_col_attrs(sh, "T", "a")
        func_b, _ = _totals_col_attrs(sh, "T", "b")
        func_c, label_c = _totals_col_attrs(sh, "T", "c")

        @test func_a == "sum"
        @test func_b == "average"
        @test isnothing(func_c)
        @test isnothing(label_c)

        # calling again on only column "a" must not disturb "b"
        XLSX.settotals!(sh, "T", "a" => :max)
        func_a2, _ = _totals_col_attrs(sh, "T", "a")
        func_b2, _ = _totals_col_attrs(sh, "T", "b")
        @test func_a2 == "max"
        @test func_b2 == "average"  # unchanged
    end

    @testset "settotals! - overwriting a column's totals replaces stale content" begin
        f = XLSX.newxlsx()
        sh = f[1]
        sh["A1"] = "a"; sh["B1"] = "b"
        sh["A2"] = 1;   sh["B2"] = 2

        XLSX.addtable!(sh, "A1:B2"; name="T")
        XLSX.settotals!(sh, "T", "b" => "Label first")
        @test sh["B3"] == "Label first"
        func1, label1 = _totals_col_attrs(sh, "T", "b")
        @test isnothing(func1)
        @test label1 == "Label first"

        # now switch that same column to a function — old label attribute
        # and cell value must be fully replaced, not left lingering
        XLSX.settotals!(sh, "T", "b" => :sum)
        func2, label2 = _totals_col_attrs(sh, "T", "b")
        @test func2 == "sum"
        @test isnothing(label2)  # old totalsRowLabel attribute must be gone
        @test ismissing(sh["B3"])  # old literal "Label first" value must be gone
    end

    @testset "settotals! - kwarg form" begin
        f = XLSX.newxlsx()
        sh = f[1]
        sh["A1"] = "revenue"; sh["B1"] = "margin"
        sh["A2"] = 100;       sh["B2"] = 40

        XLSX.addtable!(sh, "A1:B2"; name="Rpt")
        t = XLSX.settotals!(sh, "Rpt"; revenue=:sum, margin=:average)

        @test t.has_totals_row == true
        func_r, _ = _totals_col_attrs(sh, "Rpt", "revenue")
        func_m, _ = _totals_col_attrs(sh, "Rpt", "margin")
        @test func_r == "sum"
        @test func_m == "average"
    end

    @testset "settotals! - by id" begin
        f = XLSX.newxlsx()
        sh = f[1]
        sh["A1"] = "a"; sh["B1"] = "b"
        sh["A2"] = 1;   sh["B2"] = 2
        tbl = XLSX.addtable!(sh, "A1:B2"; name="ById")

        t = XLSX.settotals!(sh, tbl.id, "b" => :sum)
        @test t.has_totals_row == true
        func, _ = _totals_col_attrs(sh, "ById", "b")
        @test func == "sum"
    end

    @testset "settotals! - updating a table that already has a totals row (real fixture)" begin
        f = XLSX.opentemplate("data/two_tables.xlsx")
        sh2 = f["Sheet2"]

        t_before = XLSX.table(sh2, "with_total")
        @test t_before.has_totals_row == true
        ref_before = t_before.ref

        func_before, _ = _totals_col_attrs(sh2, "with_total", "sin")
        @test func_before == "sum"

        # change the existing function without adding a new row
        t_after = XLSX.settotals!(sh2, "with_total", "sin" => :average)

        @test t_after.ref == ref_before  # unchanged — table already had a totals row
        func_after, _ = _totals_col_attrs(sh2, "with_total", "sin")
        @test func_after == "average"
    end

    @testset "settotals! - round trip (create -> save -> reopen)" begin
        f = XLSX.newxlsx()
        sh = f[1]
        sh["A1"] = "id"; sh["B1"] = "amount"; sh["C1"] = "note"
        sh["A2"] = 1;    sh["B2"] = 50;       sh["C2"] = "x"
        sh["A3"] = 2;    sh["B3"] = 75;       sh["C3"] = "y"

        XLSX.addtable!(sh, "A1:C3"; name="RT")
        XLSX.settotals!(sh, "RT", "amount" => :sum, "note" => "Total")

        SAVE_FILES && save_outfile(f)

        outfile = "settotals_roundtrip.xlsx"
        isfile(outfile) && rm(outfile)
        XLSX.writexlsx(outfile, f, overwrite=true)
        SAVE_FILES && save_outfile(outfile)

        XLSX.openxlsx(outfile) do xf2
            sh2 = xf2[1]
            t = XLSX.table(sh2, "RT")
            @test t.has_totals_row == true
            @test t.ref == XLSX.CellRange("A1:C4")

            func, _ = _totals_col_attrs(sh2, "RT", "amount")
            @test func == "sum"
            _, label = _totals_col_attrs(sh2, "RT", "note")
            @test label == "Total"
            @test sh2["C4"] == "Total"
        end

        XLSX.openxlsx(outfile, enable_cache=false) do xf2
            t = XLSX.table(xf2[1], "RT")
            @test t.has_totals_row == true
            @test t.ref == XLSX.CellRange("A1:C4")
        end

        isfile(outfile) && rm(outfile)
    end

    @testset "show(Table) - compact form, no style, no totals" begin
        f = XLSX.newxlsx()
        sh = f[1]
        sh["A1"] = "id"; sh["B1"] = "name"; sh["C1"] = "score"
        sh["A2"] = 1; sh["B2"] = "alice"; sh["C2"] = 10.5

        t = XLSX.addtable!(sh, "A1:C2"; name="Plain")
        s = sprint(show, t)

        @test occursin("Plain", s)
        @test occursin(string(t.ref), s)
        @test occursin("3 cols", s)
        @test !occursin("totals", s)  # no totals row -> no "+totals" marker
    end

    @testset "show(Table) - compact form, with totals" begin
        f = XLSX.newxlsx()
        sh = f[1]
        sh["A1"] = "qty"; sh["B1"] = "price"
        sh["A2"] = 3;     sh["B2"] = 9.99

        XLSX.addtable!(sh, "A1:B2"; name="Sales")
        t = XLSX.settotals!(sh, "Sales", "price" => :sum)
        s = sprint(show, t)

        @test occursin("Sales", s)
        @test occursin("2 cols", s)
        @test occursin("totals", s)
    end

    @testset "show(Table) - compact form used in Vector display" begin
        f = XLSX.newxlsx()
        sh = f[1]
        sh["A1"] = "a"; sh["B1"] = "b"
        sh["A2"] = 1;   sh["B2"] = 2
        XLSX.addtable!(sh, "A1:B2"; name="T1")

        sh["D1"] = "c"; sh["E1"] = "d"
        sh["D2"] = 3;   sh["E2"] = 4
        XLSX.addtable!(sh, "D1:E2"; name="T2")

        tbls = XLSX.tables(sh)

        # plain `show` — no array summary header, just each element's compact form
        s = sprint(show, tbls)
        @test occursin("T1", s)
        @test occursin("T2", s)

        # MIME"text/plain" — this is the form that adds the "2-element" summary
        s_repl = sprint(show, MIME("text/plain"), tbls)
        @test occursin("T1", s_repl)
        @test occursin("T2", s_repl)
        @test occursin("2-element", s_repl)
    end

    @testset "show(text/plain, Table) - no style, no totals" begin
        f = XLSX.newxlsx()
        sh = f[1]
        sh["A1"] = "id"; sh["B1"] = "name"
        sh["A2"] = 1;    sh["B2"] = "alice"

        t = XLSX.addtable!(sh, "A1:B2"; name="Basic")
        s = sprint(show, MIME("text/plain"), t)

        @test occursin("Basic", s)
        @test occursin(string(t.id), s)
        @test occursin(string(t.ref), s)
        @test occursin("id, name", s)
        @test occursin("style   : none", s)
        @test occursin("totals  : no", s)
        @test !occursin("displayName", s)  # name == display_name, so no extra note
    end

    @testset "show(text/plain, Table) - with style and totals row" begin
        f = XLSX.newxlsx()
        sh = f[1]
        sh["A1"] = "region"; sh["B1"] = "revenue"
        sh["A2"] = "North";  sh["B2"] = 1000

        XLSX.addtable!(sh, "A1:B2"; name="Sales", style="TableStyleMedium9")
        t = XLSX.settotals!(sh, "Sales", "region" => "Total", "revenue" => :sum)
        s = sprint(show, MIME("text/plain"), t)

        @test occursin("Sales", s)
        @test occursin("region, revenue", s)
        @test occursin("TableStyleMedium9", s)
        @test occursin("row stripes", s)  # default style_info sets show_row_stripes=true
        @test occursin("totals  : yes", s)
    end

    @testset "show(TableStyleInfo) - compact form" begin
        style = XLSX.TableStyleInfo("TableStyleLight1", false, false, true, false)
        s = sprint(show, style)

        @test occursin("TableStyleLight1", s)
        @test occursin("rows", s)
        @test !occursin("first", s)
        @test !occursin("last", s)
        @test !occursin("cols", s)
    end

    @testset "show(TableStyleInfo) - no name, no flags" begin
        style = XLSX.TableStyleInfo(nothing, false, false, false, false)
        s = sprint(show, style)

        @test occursin("none", s)
        @test !occursin(",", s)  # no flags appended
    end

    @testset "show(TableStyleInfo) - all flags set" begin
        style = XLSX.TableStyleInfo("Custom", true, true, true, true)
        s = sprint(show, style)

        @test occursin("Custom", s)
        @test occursin("first", s)
        @test occursin("last", s)
        @test occursin("rows", s)
        @test occursin("cols", s)
    end

    @testset "show output does not error on real fixture tables" begin
        XLSX.openxlsx("data/two_tables.xlsx") do xf
            for t in XLSX.tables(xf["Sheet1"])
                @test !isempty(sprint(show, t))
                @test !isempty(sprint(show, MIME("text/plain"), t))
            end
            for t in XLSX.tables(xf["Sheet2"])
                @test !isempty(sprint(show, t))
                @test !isempty(sprint(show, MIME("text/plain"), t))
            end
        end
    end

    @testset "writetable! - as_table=true basic" begin
        f = XLSX.newxlsx()
        sh = f[1]
        data = [[1, 2, 3], ["alice", "bob", "carol"]]
        cols = ["id", "name"]

        XLSX.writetable!(sh, data, cols; as_table=true, table_name="People")

        t = XLSX.table(sh, "People")
        @test t.ref == XLSX.CellRange("A1:B4")
        @test t.columns == ["id", "name"]
        @test t.has_totals_row == false
    end

    @testset "writetable! - as_table=false writes no table (default)" begin
        f = XLSX.newxlsx()
        sh = f[1]
        data = [[1, 2], [3, 4]]
        cols = ["a", "b"]

        XLSX.writetable!(sh, data, cols)  # as_table defaults to false

        @test isempty(XLSX.tables(sh))
    end

    @testset "writetable! - as_table=true requires write_columnnames=true" begin
        f = XLSX.newxlsx()
        sh = f[1]
        data = [[1, 2], [3, 4]]
        cols = ["a", "b"]

        @test_throws XLSX.XLSXError XLSX.writetable!(sh, data, cols;
            as_table=true, write_columnnames=false)
    end

    @testset "writetable! - as_table=true requires at least one data row" begin
        f = XLSX.newxlsx()
        sh = f[1]
        data = [Int[], String[]]
        cols = ["a", "b"]

        @test_throws XLSX.XLSXError XLSX.writetable!(sh, data, cols; as_table=true)
    end

    @testset "writetable! - as_table=true with style and anchor offset" begin
        f = XLSX.newxlsx()
        sh = f[1]
        data = [[10, 20], [1.5, 2.5]]
        cols = ["qty", "price"]

        XLSX.writetable!(sh, data, cols;
            anchor_cell=XLSX.CellRef("C3"), as_table=true,
            table_name="Prices", table_style="TableStyleMedium9")

        t = XLSX.table(sh, "Prices")
        @test t.ref == XLSX.CellRange("C3:D5")
        @test t.style.name == "TableStyleMedium9"
    end

    @testset "writetable (single-sheet, new file) - as_table=true" begin
        outfile = "writetable_astable_single.xlsx"
        isfile(outfile) && rm(outfile)

        columns = [[1, 2, 3], ["x", "y", "z"]]
        colnames = ["num", "letter"]
        XLSX.writetable(outfile, columns, colnames; as_table=true, table_name="Data")
        SAVE_FILES && save_outfile(outfile)

        XLSX.openxlsx(outfile) do xf
            t = XLSX.table(xf[1], "Data")
            @test t.ref == XLSX.CellRange("A1:B4")
            @test t.columns == ["num", "letter"]
        end

        isfile(outfile) && rm(outfile)
    end

    @testset "writetable (kwarg multi-sheet) - as_table=true, table names from sheet names" begin
        outfile = "writetable_astable_multi_kw.xlsx"
        isfile(outfile) && rm(outfile)

        colsA = [[1, 2], [3, 4]]
        namesA = ["a", "b"]
        colsB = [[5, 6], [7, 8]]
        namesB = ["c", "d"]

        XLSX.writetable(outfile, as_table=true, table_style="TableStyleLight1",
            REPORT_A=(colsA, namesA), REPORT_B=(colsB, namesB))
        SAVE_FILES && save_outfile(outfile)

        XLSX.openxlsx(outfile) do xf
            tA = XLSX.table(xf["REPORT_A"], "REPORT_A")
            @test tA.ref == XLSX.CellRange("A1:B3")
            @test tA.style.name == "TableStyleLight1"

            tB = XLSX.table(xf["REPORT_B"], "REPORT_B")
            @test tB.ref == XLSX.CellRange("A1:B3")
            @test tB.style.name == "TableStyleLight1"
        end

        isfile(outfile) && rm(outfile)
    end

    @testset "writetable (Vector{Tuple} multi-sheet) - as_table=true, table names from sheet names" begin
        outfile = "writetable_astable_multi_vec.xlsx"
        isfile(outfile) && rm(outfile)

        colsA = [[1, 2], [3, 4]]
        namesA = ["a", "b"]
        colsB = [[5, 6], [7, 8]]
        namesB = ["c", "d"]

        XLSX.writetable(outfile, [
            ("First", colsA, namesA),
            ("Second", colsB, namesB),
        ]; as_table=true)
        SAVE_FILES && save_outfile(outfile)

        XLSX.openxlsx(outfile) do xf
            @test XLSX.table(xf["First"], "First").ref == XLSX.CellRange("A1:B3")
            @test XLSX.table(xf["Second"], "Second").ref == XLSX.CellRange("A1:B3")
        end

        isfile(outfile) && rm(outfile)
    end

    @testset "writetable - table name normalized from an invalid sheet name (spaces)" begin
        outfile = "writetable_astable_normalize.xlsx"
        isfile(outfile) && rm(outfile)

        cols = [[1, 2], [3, 4]]
        names = ["x", "y"]

        # sheet names may contain spaces; table names may not — expect a
        # warning and a normalized fallback name ("Report A" -> "Report_A")
        @test_logs (:warn, r"valid Excel Table name"i) match_mode=:any begin
            XLSX.writetable(outfile, [("Report A", cols, names)]; as_table=true)
        end
        SAVE_FILES && save_outfile(outfile)

        XLSX.openxlsx(outfile) do xf
            sh = xf["Report A"]
            tbls = XLSX.tables(sh)
            @test length(tbls) == 1
            @test tbls[1].name == "Report_A"
        end

        isfile(outfile) && rm(outfile)
    end

    @testset "writetable - normalized name falls back to auto-generated on collision" begin
        outfile = "writetable_astable_normalize_collision.xlsx"
        isfile(outfile) && rm(outfile)

        cols = [[1, 2], [3, 4]]
        names = ["x", "y"]

        # "Report A" normalizes to "Report_A"; second sheet is literally
        # named "Report_A" already — its table name candidate collides with
        # the first table, so it must fall back to an auto-generated name.
        # Two warnings are expected here (normalization + collision) but
        # aren't the focus of this test, so they're suppressed rather than
        # left to print to the console.
        with_logger(NullLogger()) do
            XLSX.writetable(outfile, [
                ("Report A", cols, names),
                ("Report_A", cols, names),
            ]; as_table=true)
        end
        SAVE_FILES && save_outfile(outfile)

        XLSX.openxlsx(outfile) do xf
            t1 = only(XLSX.tables(xf["Report A"]))
            t2 = only(XLSX.tables(xf["Report_A"]))

            @test t1.name == "Report_A"
            @test t2.name != "Report_A"   # forced to fall back
            @test t2.name != t1.name
            @test startswith(t2.name, "Table")  # addtable!'s auto-naming pattern
        end

        isfile(outfile) && rm(outfile)
    end

    @testset "writetable - reserved Julia keyword as sheet name still normalizes" begin
        outfile = "writetable_astable_reserved.xlsx"
        isfile(outfile) && rm(outfile)

        cols = [[1, 2], [3, 4]]
        names = ["x", "y"]

        # "for" is a valid Excel sheet name (no restricted characters) but a
        # reserved Julia keyword; normalizename prepends "_" for these. The
        # resulting warning isn't the focus of this test, so it's suppressed.
        with_logger(NullLogger()) do
            XLSX.writetable(outfile, [("for", cols, names)]; as_table=true)
        end
        SAVE_FILES && save_outfile(outfile)

        XLSX.openxlsx(outfile) do xf
            t = only(XLSX.tables(xf["for"]))
            @test t.name == "_for"
        end

        isfile(outfile) && rm(outfile)
    end

    @testset "writetable - as_table=false leaves no tables in multi-sheet write" begin
        outfile = "writetable_no_table_multi.xlsx"
        isfile(outfile) && rm(outfile)

        cols = [[1, 2], [3, 4]]
        names = ["x", "y"]
        XLSX.writetable(outfile, [("Sheet1", cols, names), ("Sheet2", cols, names)])
        SAVE_FILES && save_outfile(outfile)

        XLSX.openxlsx(outfile) do xf
            @test isempty(XLSX.tables(xf["Sheet1"]))
            @test isempty(XLSX.tables(xf["Sheet2"]))
        end

        isfile(outfile) && rm(outfile)
    end

   @testset "Tables.istable / rowaccess / columnaccess" begin
        @test Tables.istable(XLSX.Table)
        @test Tables.istable(XLSX.XLSXTableRowIterator)
        @test Tables.rowaccess(XLSX.Table)
        @test Tables.rowaccess(XLSX.XLSXTableRowIterator)
        @test Tables.columnaccess(XLSX.Table)
    end

    @testset "Tables.schema - column names, types unknown (nothing)" begin
        f = XLSX.newxlsx()
        sh = f[1]
        sh["A1"] = "id"; sh["B1"] = "name"; sh["C1"] = "score"
        sh["A2"] = 1; sh["B2"] = "alice"; sh["C2"] = 10.5

        t = XLSX.addtable!(sh, "A1:C2"; name="T")
        sch = Tables.schema(t)

        # `nothing` (not a declared `fill(Any, n)`) is deliberate: declaring
        # Any explicitly would make DataFrames (and other Tables.jl sinks)
        # trust that declaration literally and permanently box every column
        # as Any, rather than inferring real types from the data — exactly
        # the issue #225 regression.
        @test sch.names == (:id, :name, :score)
        @test isnothing(sch.types)
    end

    @testset "Tables.columns / DataFrame infer concrete column types (issue #225 regression)" begin
        f = XLSX.newxlsx()
        sh = f[1]
        sh["A1"] = "score"; sh["B1"] = "label"
        sh["A2"] = 10.5;    sh["B2"] = "alice"
        sh["A3"] = 20.0;    sh["B3"] = "bob"
        sh["A4"] = 15.25;   sh["B4"] = "carol"

        t = XLSX.addtable!(sh, "A1:B4"; name="T")

        # Direct Tables.columns(t) — the underlying mechanism
        cols = Tables.columns(t)
        @test eltype(cols.score) != Any
        @test eltype(cols.score) <: Union{Missing,Float64}
        @test eltype(cols.label) <: Union{Missing,String}

        # Same, via the row iterator returned by eachtablerow — this is the
        # exact path that previously regressed to Any/Any (issue #225)
        it = XLSX.eachtablerow(t)
        cols_via_iter = Tables.columns(it)
        @test eltype(cols_via_iter.score) != Any
        @test eltype(cols_via_iter.score) <: Union{Missing,Float64}

        # And the full DataFrame(...) round trip
        df = DataFrames.DataFrame(it)
        @test eltype(df.score) != Any
        @test eltype(df.score) <: Union{Missing,Float64}
        @test eltype(df.label) <: Union{Missing,String}
    end

    @testset "Tables.schema - matches between Table and its row iterator" begin
        f = XLSX.newxlsx()
        sh = f[1]
        sh["A1"] = "a"; sh["B1"] = "b"
        sh["A2"] = 1;   sh["B2"] = 2

        t = XLSX.addtable!(sh, "A1:B2"; name="T")
        it = Tables.rows(t)

        @test it isa XLSX.XLSXTableRowIterator
        @test Tables.rows(it) === it  # identity, matching TableRowIterator convention
        @test Tables.schema(it).names == Tables.schema(t).names
        @test Tables.schema(it).types == Tables.schema(t).types
    end

    @testset "Tables.columnnames matches table columns" begin
        f = XLSX.newxlsx()
        sh = f[1]
        sh["A1"] = "x"; sh["B1"] = "y"
        sh["A2"] = 1;   sh["B2"] = 2

        t = XLSX.addtable!(sh, "A1:B2"; name="T")
        row = first(XLSX.eachtablerow(t))
        @test Tables.columnnames(row) == [:x, :y]
        @test Tables.columnnames(t) == [:x, :y]
    end

    @testset "Tables.columns - values match what was written" begin
        f = XLSX.newxlsx()
        sh = f[1]
        sh["A1"] = "id"; sh["B1"] = "name"; sh["C1"] = "score"
        sh["A2"] = 1; sh["B2"] = "alice"; sh["C2"] = 10.5
        sh["A3"] = 2; sh["B3"] = "bob";   sh["C3"] = 20.0
        sh["A4"] = 3; sh["B4"] = "carol"; sh["C4"] = 15.0

        t = XLSX.addtable!(sh, "A1:C4"; name="People")
        cols = Tables.columns(t)

        @test cols.id == [1, 2, 3]
        @test cols.name == ["alice", "bob", "carol"]
        @test cols.score == [10.5, 20.0, 15.0]
    end

    @testset "eachtablerow - length excludes header row" begin
        f = XLSX.newxlsx()
        sh = f[1]
        sh["A1"] = "a"; sh["B1"] = "b"
        sh["A2"] = 1;   sh["B2"] = 2
        sh["A3"] = 3;   sh["B3"] = 4
        sh["A4"] = 5;   sh["B4"] = 6

        t = XLSX.addtable!(sh, "A1:B4"; name="T")
        rows = collect(XLSX.eachtablerow(t))

        @test length(rows) == 3
        @test length(XLSX.eachtablerow(t)) == 3  # Base.length via iterator, not just collect
    end

    @testset "eachtablerow - minimum valid table (one data row)" begin
        f = XLSX.newxlsx()
        sh = f[1]
        sh["A1"] = "a"; sh["B1"] = "b"
        sh["A2"] = 1;   sh["B2"] = 2

        t = XLSX.addtable!(sh, "A1:B2"; name="T")
        rows = collect(XLSX.eachtablerow(t))

        @test length(rows) == 1
        @test Tables.getcolumn(rows[1], :a) == 1
        @test Tables.getcolumn(rows[1], :b) == 2
        @test Tables.getcolumn(rows[1], 1) == 1
        @test Tables.getcolumn(rows[1], 2) == 2
    end

    @testset "eachtablerow - values match by name and by index" begin
        f = XLSX.newxlsx()
        sh = f[1]
        sh["A1"] = "id"; sh["B1"] = "name"
        sh["A2"] = 1;    sh["B2"] = "alice"
        sh["A3"] = 2;    sh["B3"] = "bob"

        t = XLSX.addtable!(sh, "A1:B3"; name="People")
        rows = collect(XLSX.eachtablerow(t))

        @test Tables.getcolumn(rows[1], :id) == 1
        @test Tables.getcolumn(rows[1], :name) == "alice"
        @test Tables.getcolumn(rows[2], 1) == 2
        @test Tables.getcolumn(rows[2], 2) == "bob"
    end

    @testset "eachtablerow - totals row excluded (created via settotals!)" begin
        f = XLSX.newxlsx()
        sh = f[1]
        sh["A1"] = "item"; sh["B1"] = "amount"
        sh["A2"] = "Apples"; sh["B2"] = 12
        sh["A3"] = "Pears";  sh["B3"] = 8

        XLSX.addtable!(sh, "A1:B3"; name="Bulk")
        t = XLSX.settotals!(sh, "Bulk", "item" => "Total", "amount" => :sum)

        rows = collect(XLSX.eachtablerow(t))
        @test length(rows) == 2  # totals row (row 4) must NOT be included
        @test Tables.getcolumn(rows[1], :item) == "Apples"
        @test Tables.getcolumn(rows[2], :item) == "Pears"

        cols = Tables.columns(t)
        @test cols.item == ["Apples", "Pears"]
        @test cols.amount == [12, 8]
    end

    @testset "eachtablerow - missing values pass through" begin
        f = XLSX.newxlsx()
        sh = f[1]
        sh["A1"] = "a"; sh["B1"] = "b"
        sh["A2"] = 1;   sh["B2"] = missing
        sh["A3"] = missing; sh["B3"] = 4

        t = XLSX.addtable!(sh, "A1:B3"; name="T")
        rows = collect(XLSX.eachtablerow(t))

        @test ismissing(Tables.getcolumn(rows[1], :b))
        @test ismissing(Tables.getcolumn(rows[2], :a))
    end

    @testset "eachtablerow - a fully blank row within ref is preserved, not skipped" begin
        # Unlike gettable/TableRowIterator (which infers table bounds from
        # content and offers stop_in_empty_row/keep_empty_rows/
        # stop_in_row_function to resolve that ambiguity), a Table's `ref` is
        # already authoritative — every row between header and totals (if
        # any) is unconditionally part of the table, whether blank or not.
        # eachtablerow must never drop a row just because it happens to be
        # entirely empty.
        f = XLSX.newxlsx()
        sh = f[1]
        sh["A1"] = "a"; sh["B1"] = "b"
        sh["A2"] = 1;   sh["B2"] = 10
        # row 3 deliberately left entirely blank
        sh["A4"] = 3;   sh["B4"] = 30

        t = XLSX.addtable!(sh, "A1:B4"; name="T")
        rows = collect(XLSX.eachtablerow(t))

        @test length(rows) == 3  # rows 2, 3, 4 — the blank row 3 must still be present
        @test Tables.getcolumn(rows[1], :a) == 1
        @test ismissing(Tables.getcolumn(rows[2], :a))
        @test ismissing(Tables.getcolumn(rows[2], :b))
        @test Tables.getcolumn(rows[3], :a) == 3

        cols = Tables.columns(t)
        @test length(cols.a) == 3
        @test ismissing(cols.a[2])
    end

    @testset "Tables.rowtable - generic Tables.jl round trip" begin
        f = XLSX.newxlsx()
        sh = f[1]
        sh["A1"] = "id"; sh["B1"] = "score"
        sh["A2"] = 1;    sh["B2"] = 10.5
        sh["A3"] = 2;    sh["B3"] = 20.0

        t = XLSX.addtable!(sh, "A1:B3"; name="T")

        nt_rows = Tables.rowtable(t)
        @test length(nt_rows) == 2
        @test nt_rows[1].id == 1
        @test nt_rows[1].score == 10.5
        @test nt_rows[2].id == 2
        @test nt_rows[2].score == 20.0
    end

    @testset "real fixture (two_tables.xlsx) - with_total (Sheet2), has a totals row" begin
        XLSX.openxlsx("data/two_tables.xlsx") do xf
            sh2 = xf["Sheet2"]
            t = XLSX.table(sh2, "with_total")
            @test t.has_totals_row == true
            @test t.columns == ["start", "stop", "sin"]

            rows = collect(XLSX.eachtablerow(t))
            expected_data_rows = (t.ref.stop.row_number - 1) - (t.ref.start.row_number + 1) + 1
            @test length(rows) == expected_data_rows

            # the last row iterated must be genuine data — not the totals
            # row. The totals row's "sin" column holds a formula (no cached
            # value), so getdata would return `missing`; confirm the last
            # data row's "sin" value is real, present data instead.
            last_row = rows[end]
            @test !ismissing(Tables.getcolumn(last_row, :sin))
        end
    end

    @testset "real fixture (two_tables.xlsx) - Age_height (Sheet1), no totals row" begin
        XLSX.openxlsx("data/two_tables.xlsx") do xf
            sh = xf["Sheet1"]
            t = XLSX.table(sh, "Age_height")
            @test t.has_totals_row == false

            rows = collect(XLSX.eachtablerow(t))
            @test length(rows) == t.ref.stop.row_number - t.ref.start.row_number  # all rows below header
        end
    end

    @testset "real fixture (two_tables.xlsx) - IO_Table, no totals row" begin
        XLSX.openxlsx("data/two_tables.xlsx") do xf
            sh = xf["Sheet1"]
            t = XLSX.table(sh, "IO_Table")
            @test t.has_totals_row == false

            rows = collect(XLSX.eachtablerow(t))
            @test length(rows) == t.ref.stop.row_number - t.ref.start.row_number  # all rows below header
            @test Tables.columnnames(rows[1]) == Symbol.(t.columns)
        end
    end

    @testset "t.sheet identity is preserved" begin
        f = XLSX.newxlsx()
        sh = f[1]
        sh["A1"] = "a"; sh["B1"] = "b"
        sh["A2"] = 1;   sh["B2"] = 2

        XLSX.addtable!(sh, "A1:B2"; name="T")
        t = XLSX.table(sh, "T")
        @test t.sheet === sh
    end

    @testset "row-access path (Tables.rows) resolves schema/columnnames without erroring" begin
        # Regression check: PrettyTables (and other row-access consumers)
        # call Tables.schema/Tables.columnnames on Tables.rows(t) — the
        # XLSXTableRowIterator — not on `t` directly.
        f = XLSX.newxlsx()
        sh = f[1]
        sh["A1"] = "id"; sh["B1"] = "name"
        sh["A2"] = 1;    sh["B2"] = "alice"
        sh["A3"] = 2;    sh["B3"] = "bob"

        t = XLSX.addtable!(sh, "A1:B3"; name="T")
        it = Tables.rows(t)

        @test !isnothing(Tables.schema(it))
        @test Tables.schema(it).names == (:id, :name)
        @test !isnothing(Tables.columnnames(first(it)))
        @test collect(it) isa Vector{XLSX.XLSXTableRow}
    end

    @testset "basic matrix shape and values" begin
        f = XLSX.newxlsx()
        sh = f[1]
        sh["A1"] = "region"; sh["B1"] = "revenue"
        sh["A2"] = "North";  sh["B2"] = 1000
        sh["A3"] = "South";  sh["B3"] = 1500
        sh["A4"] = "East";   sh["B4"] = 900

        t = XLSX.addtable!(sh, "A1:B4"; name="Sales")
        m = XLSX.getdata(t)

        @test m isa Matrix{Any}
        @test size(m) == (3, 2)
        @test m[1, 1] == "North"; @test m[1, 2] == 1000
        @test m[2, 1] == "South"; @test m[2, 2] == 1500
        @test m[3, 1] == "East";  @test m[3, 2] == 900
    end

    @testset "excludes totals row" begin
        f = XLSX.newxlsx()
        sh = f[1]
        sh["A1"] = "item";   sh["B1"] = "amount"
        sh["A2"] = "Apples"; sh["B2"] = 12
        sh["A3"] = "Pears";  sh["B3"] = 8

        XLSX.addtable!(sh, "A1:B3"; name="Bulk")
        t = XLSX.settotals!(sh, "Bulk", "item" => "Total", "amount" => :sum)

        m = XLSX.getdata(t)
        @test size(m) == (2, 2)  # totals row (row 4) must not appear
        @test m[end, 1] == "Pears"
        @test m[end, 2] == 8
    end

    @testset "blank row within ref is preserved, not skipped" begin
        f = XLSX.newxlsx()
        sh = f[1]
        sh["A1"] = "a"; sh["B1"] = "b"
        sh["A2"] = 1;   sh["B2"] = 10
        # row 3 deliberately left entirely blank
        sh["A4"] = 3;   sh["B4"] = 30

        t = XLSX.addtable!(sh, "A1:B4"; name="T")
        m = XLSX.getdata(t)

        @test size(m) == (3, 2)
        @test m[1, 1] == 1
        @test ismissing(m[2, 1])
        @test ismissing(m[2, 2])
        @test m[3, 1] == 3
    end

    @testset "minimum valid table (one data row)" begin
        f = XLSX.newxlsx()
        sh = f[1]
        sh["A1"] = "a"; sh["B1"] = "b"
        sh["A2"] = 1;   sh["B2"] = 2

        t = XLSX.addtable!(sh, "A1:B2"; name="T")
        m = XLSX.getdata(t)

        @test size(m) == (1, 2)
        @test m[1, 1] == 1
        @test m[1, 2] == 2
    end

    @testset "real fixture (two_tables.xlsx) - IO_Table" begin
        XLSX.openxlsx("data/two_tables.xlsx") do xf
            sh = xf["Sheet1"]
            t = XLSX.table(sh, "IO_Table")
            m = XLSX.getdata(t)

            @test size(m) == (t.ref.stop.row_number - t.ref.start.row_number, length(t.columns))
        end
    end

end
