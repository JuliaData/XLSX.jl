# `NoSchemaCols` exposes column names but deliberately returns `nothing` from
# `Tables.schema`, forcing the `Tables.columnnames(Tables.columns(...))`
# fallback. `NoNamesSource` returns `nothing` from schema *and* errors on
# column access, forcing the positional fallback.
struct NoSchemaCols
    nt::NamedTuple
end
Tables.istable(::Type{NoSchemaCols}) = true
Tables.columnaccess(::Type{NoSchemaCols}) = true
Tables.columns(x::NoSchemaCols) = x
Tables.schema(::NoSchemaCols) = nothing
Tables.columnnames(x::NoSchemaCols) = collect(keys(getfield(x, :nt)))
Tables.getcolumn(x::NoSchemaCols, nm::Symbol) = getfield(x, :nt)[nm]
Tables.getcolumn(x::NoSchemaCols, i::Int) = getfield(x, :nt)[i]
Tables.rows(x::NoSchemaCols) = Tables.rows(getfield(x, :nt))

struct PosRow <: Tables.AbstractRow
    v::Vector{Any}
end
Tables.getcolumn(r::PosRow, i::Int) = getfield(r, :v)[i]
Tables.getcolumn(r::PosRow, ::Symbol) = error("this source has no column names")
Tables.columnnames(::PosRow) = error("this source has no column names")

struct NoNamesSource
    rows::Vector{Vector{Any}}
end
Tables.istable(::Type{NoNamesSource}) = true
Tables.rowaccess(::Type{NoNamesSource}) = true
Tables.schema(::NoNamesSource) = nothing
Tables.rows(x::NoNamesSource) = [PosRow(r) for r in getfield(x, :rows)]
Tables.columns(::NoNamesSource) = error("this source has no column access")

# Parse a table part from a literal XML string, for driving the parser
# error paths directly rather than through a real file.
_tbl_doc(s::AbstractString) = parse(s, XLSX.XML.Node)

const _TBL_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"

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

function fresh_abc()
    f = XLSX.newxlsx()
    sh = f[1]
    sh["A1"] = "a"; sh["B1"] = "b"; sh["C1"] = "c"
    sh["A2"] = 1;   sh["B2"] = 2;   sh["C2"] = 3
    XLSX.addtable!(sh, "A1:C2"; name="T")
    return sh
end
function fresh()
    f = XLSX.newxlsx()
    sh = f[1]
    sh["A1"] = "x"; sh["B1"] = "y"
    sh["A2"] = 1;   sh["B2"] = 10
    XLSX.addtable!(sh, "A1:B2"; name="T")
    return sh
end

function make_readto_file(outfile)
    isfile(outfile) && rm(outfile)
    f = XLSX.newxlsx("First")

    sh1 = f["First"]
    sh1["A1"] = "a"; sh1["B1"] = "b"
    sh1["A2"] = 1;   sh1["B2"] = 10
    sh1["A3"] = 2;   sh1["B3"] = 20
    XLSX.addtable!(sh1, "A1:B3"; name="TableOne")

    sh2 = XLSX.addsheet!(f, "Second")
    sh2["A1"] = "id"; sh2["B1"] = "name"; sh2["C1"] = "score"
    sh2["A2"] = 1;    sh2["B2"] = "alice"; sh2["C2"] = 10.5
    sh2["A3"] = 2;    sh2["B3"] = "bob";   sh2["C3"] = 20.0
    sh2["A4"] = 3;    sh2["B4"] = "carol"; sh2["C4"] = 15.25
    XLSX.addtable!(sh2, "A1:C4"; name="TableTwo")
    # data outside the table, to confirm it isn't picked up
    sh2["E1"] = "outside"; sh2["E2"] = "not in table"

    XLSX.writexlsx(outfile, f, overwrite=true)
    return outfile
end


# Shared fixture: 3 sheets, one table each, plus extra non-table data
# on the middle sheet to confirm reads are scoped to the table's ref.
function make_multitable_file(outfile)
    isfile(outfile) && rm(outfile)
    f = XLSX.newxlsx("First")

    sh1 = f["First"]
    sh1["A1"] = "a"; sh1["B1"] = "b"
    sh1["A2"] = 1;   sh1["B2"] = 10
    sh1["A3"] = 2;   sh1["B3"] = 20
    XLSX.addtable!(sh1, "A1:B3"; name="TableOne")

    sh2 = XLSX.addsheet!(f, "Second")
    sh2["A1"] = "id"; sh2["B1"] = "name"; sh2["C1"] = "score"
    sh2["A2"] = 1;    sh2["B2"] = "alice"; sh2["C2"] = 10.5
    sh2["A3"] = 2;    sh2["B3"] = "bob";   sh2["C3"] = 20.0
    sh2["A4"] = 3;    sh2["B4"] = "carol"; sh2["C4"] = 15.25
    XLSX.addtable!(sh2, "A1:C4"; name="TableTwo")
    # data outside the table, to confirm it's never picked up
    sh2["E1"] = "outside"; sh2["E2"] = "not in table"
    sh2["A6"] = "below";   sh2["B6"] = "also not in table"

    sh3 = XLSX.addsheet!(f, "Third")
    sh3["A1"] = "e"; sh3["B1"] = "f"
    sh3["A2"] = 5;   sh3["B2"] = 6
    XLSX.addtable!(sh3, "A1:B2"; name="TableThree")

    XLSX.writexlsx(outfile, f, overwrite=true)
    return outfile
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
        @test occursin("id=$(t.id)", s)
        @test occursin("1x3", s)   # 1 data row x 3 columns
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
        @test occursin("id=$(t.id)", s)
        @test occursin("1x2", s)   # data rows exclude header and totals row
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

    @testset "sheet-scoped: readtable(source, sheet; table_name=...)" begin
        outfile = make_multitable_file("readtable_tablename_scoped.xlsx")
        SAVE_FILES && save_outfile(outfile)

        dt = XLSX.readtable(outfile, "Second"; table_name="TableTwo")

        @test dt isa XLSX.DataTable
        @test dt.column_labels == [:id, :name, :score]
        @test length(dt.data) == 3
        @test dt.data[1] == [1, 2, 3]
        @test dt.data[2] == ["alice", "bob", "carol"]
        @test dt.data[3] == [10.5, 20.0, 15.25]

        isfile(outfile) && rm(outfile)
    end

    @testset "sheet-scoped: by sheet index rather than name" begin
        outfile = make_multitable_file("readtable_tablename_scoped_idx.xlsx")

        dt = XLSX.readtable(outfile, 2; table_name="TableTwo")
        @test dt.column_labels == [:id, :name, :score]
        @test dt.data[1] == [1, 2, 3]

        isfile(outfile) && rm(outfile)
    end

    @testset "sheet-scoped: excludes data outside the table's ref" begin
        outfile = make_multitable_file("readtable_tablename_bounds.xlsx")

        dt = XLSX.readtable(outfile, "Second"; table_name="TableTwo")

        # only the table's 3 columns / 3 data rows — nothing from E1:E2 or row 6
        @test length(dt.column_labels) == 3
        @test :outside ∉ dt.column_labels
        @test all(length(col) == 3 for col in dt.data)
        @test "not in table" ∉ dt.data[1]
        @test "also not in table" ∉ dt.data[2]

        isfile(outfile) && rm(outfile)
    end

    @testset "workbook-wide: readtable(source; table_name=...)" begin
        outfile = make_multitable_file("readtable_tablename_wide.xlsx")

        # table on the *second* sheet, found without naming the sheet
        dt = XLSX.readtable(outfile; table_name="TableTwo")
        @test dt.column_labels == [:id, :name, :score]
        @test dt.data[1] == [1, 2, 3]

        # table on the first sheet
        dt1 = XLSX.readtable(outfile; table_name="TableOne")
        @test dt1.column_labels == [:a, :b]
        @test dt1.data[1] == [1, 2]

        # table on the last sheet
        dt3 = XLSX.readtable(outfile; table_name="TableThree")
        @test dt3.column_labels == [:e, :f]
        @test dt3.data[1] == [5]

        isfile(outfile) && rm(outfile)
    end

    @testset "sheet-scoped and workbook-wide agree" begin
        outfile = make_multitable_file("readtable_tablename_agree.xlsx")

        dt_scoped = XLSX.readtable(outfile, "Second"; table_name="TableTwo")
        dt_wide   = XLSX.readtable(outfile; table_name="TableTwo")

        @test dt_scoped.column_labels == dt_wide.column_labels
        @test dt_scoped.data == dt_wide.data

        isfile(outfile) && rm(outfile)
    end

    @testset "enable_cache=false: sheet-scoped" begin
        outfile = make_multitable_file("readtable_tablename_nocache_scoped.xlsx")

        dt = XLSX.readtable(outfile, "Second"; table_name="TableTwo", enable_cache=false)
        @test dt.column_labels == [:id, :name, :score]
        @test dt.data[1] == [1, 2, 3]
        @test dt.data[2] == ["alice", "bob", "carol"]
        @test dt.data[3] == [10.5, 20.0, 15.25]

        isfile(outfile) && rm(outfile)
    end

    @testset "enable_cache=false: workbook-wide (scans every sheet's tableParts)" begin
        outfile = make_multitable_file("readtable_tablename_nocache_wide.xlsx")

        # This is the least-exercised corner: no target_sheet, no cache, so
        # every sheet's <tableParts> is cursor-scanned via
        # open_internal_file_stream rather than read from xf.data.
        dt = XLSX.readtable(outfile; table_name="TableTwo", enable_cache=false)
        @test dt.column_labels == [:id, :name, :score]
        @test dt.data[1] == [1, 2, 3]

        # also reach the first and last sheets under the same conditions
        dt1 = XLSX.readtable(outfile; table_name="TableOne", enable_cache=false)
        @test dt1.data[1] == [1, 2]

        dt3 = XLSX.readtable(outfile; table_name="TableThree", enable_cache=false)
        @test dt3.data[1] == [5]

        isfile(outfile) && rm(outfile)
    end

    @testset "enable_cache=false agrees with enable_cache=true" begin
        outfile = make_multitable_file("readtable_tablename_cachemodes.xlsx")

        dt_cached   = XLSX.readtable(outfile; table_name="TableTwo", enable_cache=true)
        dt_nocache  = XLSX.readtable(outfile; table_name="TableTwo", enable_cache=false)

        @test dt_cached.column_labels == dt_nocache.column_labels
        @test dt_cached.data == dt_nocache.data

        isfile(outfile) && rm(outfile)
    end

    @testset "infer_eltypes / normalizenames / missing_strings still apply" begin
        outfile = "readtable_tablename_kwargs.xlsx"
        isfile(outfile) && rm(outfile)

        f = XLSX.newxlsx()
        sh = f[1]
        sh["A1"] = "col one"; sh["B1"] = "value"
        sh["A2"] = "x";       sh["B2"] = 1.5
        sh["A3"] = "N/A";     sh["B3"] = 2.5
        XLSX.addtable!(sh, "A1:B3"; name="T")
        XLSX.writexlsx(outfile, f, overwrite=true)
        SAVE_FILES && save_outfile(outfile)

        # infer_eltypes=true (default) narrows columns
        dt = XLSX.readtable(outfile; table_name="T")
        @test eltype(dt.data[2]) != Any
        @test eltype(dt.data[2]) <: Union{Missing,Float64}

        # infer_eltypes=false leaves them as Any
        dt_any = XLSX.readtable(outfile; table_name="T", infer_eltypes=false)
        @test eltype(dt_any.data[2]) == Any

        # normalizenames turns "col one" into a valid identifier
        dt_norm = XLSX.readtable(outfile; table_name="T", normalizenames=true)
        @test dt_norm.column_labels[1] == :col_one

        # without normalizenames, the label keeps its space
        @test dt.column_labels[1] == Symbol("col one")

        # missing_strings converts "N/A" to missing
        dt_miss = XLSX.readtable(outfile; table_name="T", missing_strings="N/A")
        @test ismissing(dt_miss.data[1][2])
        @test dt_miss.data[1][1] == "x"

        isfile(outfile) && rm(outfile)
    end

    @testset "DataFrame round trip via readtable(; table_name)" begin
        outfile = make_multitable_file("readtable_tablename_df.xlsx")

        df = DataFrames.DataFrame(XLSX.readtable(outfile, "Second"; table_name="TableTwo"))
        @test size(df) == (3, 3)
        @test names(df) == ["id", "name", "score"]
        @test df.score == [10.5, 20.0, 15.25]
        @test eltype(df.score) != Any

        isfile(outfile) && rm(outfile)
    end

    @testset "errors: table not found" begin
        outfile = make_multitable_file("readtable_tablename_notfound.xlsx")

        # workbook-wide: no such table anywhere
        @test_throws XLSX.XLSXError XLSX.readtable(outfile; table_name="NoSuchTable")

        # sheet-scoped: table exists, but not on the named sheet
        @test_throws KeyError XLSX.readtable(outfile, "First"; table_name="TableTwo")

        isfile(outfile) && rm(outfile)
    end

    @testset "errors: table_name cannot combine with a columns range" begin
        outfile = make_multitable_file("readtable_tablename_conflict.xlsx")

        @test_throws XLSX.XLSXError XLSX.readtable(outfile, "Second", "A:C"; table_name="TableTwo")
        @test_throws XLSX.XLSXError XLSX.readtable(outfile, "Second", XLSX.ColumnRange("A:C"); table_name="TableTwo")

        isfile(outfile) && rm(outfile)
    end

    @testset "table with a totals row: totals excluded" begin
        outfile = "readtable_tablename_totals.xlsx"
        isfile(outfile) && rm(outfile)

        f = XLSX.newxlsx()
        sh = f[1]
        sh["A1"] = "item";   sh["B1"] = "amount"
        sh["A2"] = "Apples"; sh["B2"] = 12
        sh["A3"] = "Pears";  sh["B3"] = 8
        XLSX.addtable!(sh, "A1:B3"; name="Bulk")
        XLSX.settotals!(sh, "Bulk", "item" => "Total", "amount" => :sum)
        XLSX.writexlsx(outfile, f, overwrite=true)
        SAVE_FILES && save_outfile(outfile)

        for cache in (true, false)
            dt = XLSX.readtable(outfile; table_name="Bulk", enable_cache=cache)
            @test length(dt.data[1]) == 2  # totals row excluded
            @test dt.data[1] == ["Apples", "Pears"]
            @test dt.data[2] == [12, 8]
        end

        isfile(outfile) && rm(outfile)
    end

    @testset "table with a blank row: row preserved" begin
        outfile = "readtable_tablename_blank.xlsx"
        isfile(outfile) && rm(outfile)

        f = XLSX.newxlsx()
        sh = f[1]
        sh["A1"] = "a"; sh["B1"] = "b"
        sh["A2"] = 1;   sh["B2"] = 10
        # row 3 blank
        sh["A4"] = 3;   sh["B4"] = 30
        XLSX.addtable!(sh, "A1:B4"; name="T")
        XLSX.writexlsx(outfile, f, overwrite=true)

        for cache in (true, false)
            dt = XLSX.readtable(outfile; table_name="T", enable_cache=cache)
            @test length(dt.data[1]) == 3  # blank row still present
            @test ismissing(dt.data[1][2])
            @test ismissing(dt.data[2][2])
        end

        isfile(outfile) && rm(outfile)
    end

    @testset "existing readtable behaviour unaffected when table_name omitted" begin
        outfile = make_multitable_file("readtable_tablename_regression.xlsx")

        # plain range-inference read of the same sheet still works and, on
        # this fixture, picks up more than just the table (row 6 data too)
        dt_plain = XLSX.readtable(outfile, "Second")
        @test dt_plain isa XLSX.DataTable
        @test dt_plain.column_labels == [:id, :name, :score]

        isfile(outfile) && rm(outfile)
    end

    @testset "multiple spaces collapse to a single underscore" begin
        outfile = "readtable_normalize_multispace.xlsx"
        isfile(outfile) && rm(outfile)

        f = XLSX.newxlsx()
        sh = f[1]
        sh["A1"] = "col  one"; sh["B1"] = "col   two"   # 2 and 3 spaces
        sh["A2"] = 1;          sh["B2"] = 2
        XLSX.addtable!(sh, "A1:B2"; name="T")
        XLSX.writexlsx(outfile, f, overwrite=true)
        SAVE_FILES && save_outfile(outfile)

        dt = XLSX.readtable(outfile; table_name="T", normalizenames=true)
        @test dt.column_labels == [:col_one, :col_two]

        # without normalizenames, the original spacing is preserved
        dt_raw = XLSX.readtable(outfile; table_name="T")
        @test dt_raw.column_labels == [Symbol("col  one"), Symbol("col   two")]

        isfile(outfile) && rm(outfile)
    end

    @testset "header starting with a digit gets a leading underscore" begin
        outfile = "readtable_normalize_digit.xlsx"
        isfile(outfile) && rm(outfile)

        f = XLSX.newxlsx()
        sh = f[1]
        sh["A1"] = "2024"; sh["B1"] = "2025 total"
        sh["A2"] = 10;     sh["B2"] = 20
        XLSX.addtable!(sh, "A1:B2"; name="T")
        XLSX.writexlsx(outfile, f, overwrite=true)

        dt = XLSX.readtable(outfile; table_name="T", normalizenames=true)
        @test dt.column_labels == [:_2024, :_2025_total]

        isfile(outfile) && rm(outfile)
    end

    @testset "reserved Julia keyword header gets a leading underscore" begin
        outfile = "readtable_normalize_reserved.xlsx"
        isfile(outfile) && rm(outfile)

        f = XLSX.newxlsx()
        sh = f[1]
        sh["A1"] = "end"; sh["B1"] = "function"
        sh["A2"] = 1;     sh["B2"] = 2
        XLSX.addtable!(sh, "A1:B2"; name="T")
        XLSX.writexlsx(outfile, f, overwrite=true)

        dt = XLSX.readtable(outfile; table_name="T", normalizenames=true)
        @test dt.column_labels == [:_end, :_function]

        isfile(outfile) && rm(outfile)
    end

    @testset "gettable(t) directly honours normalizenames" begin
        f = XLSX.newxlsx()
        sh = f[1]
        sh["A1"] = "col one"; sh["B1"] = "already_ok"
        sh["A2"] = 1;         sh["B2"] = 2
        t = XLSX.addtable!(sh, "A1:B2"; name="T")

        dt = XLSX.gettable(t; normalizenames=true)
        @test dt.column_labels == [:col_one, :already_ok]

        dt_raw = XLSX.gettable(t)
        @test dt_raw.column_labels == [Symbol("col one"), :already_ok]
    end

    @testset "appends rows to a table with no totals row" begin
        f = XLSX.newxlsx()
        sh = f[1]
        sh["A1"] = "id"; sh["B1"] = "name"
        sh["A2"] = 1;    sh["B2"] = "alice"
        sh["A3"] = 2;    sh["B3"] = "bob"
        XLSX.addtable!(sh, "A1:B3"; name="People")

        t = XLSX.appendtable!(sh, "People", [(3, "carol"), (4, "dave")])

        @test t.ref == XLSX.CellRange("A1:B5")
        @test t.has_totals_row == false
        @test sh["A4"] == 3; @test sh["B4"] == "carol"
        @test sh["A5"] == 4; @test sh["B5"] == "dave"

        cols = Tables.columns(t)
        @test cols.id == [1, 2, 3, 4]
        @test cols.name == ["alice", "bob", "carol", "dave"]
    end

    @testset "appending zero rows is a no-op" begin
        f = XLSX.newxlsx()
        sh = f[1]
        sh["A1"] = "a"; sh["B1"] = "b"
        sh["A2"] = 1;   sh["B2"] = 2
        XLSX.addtable!(sh, "A1:B2"; name="T")

        t = XLSX.appendtable!(sh, "T", Tuple{Int,Int}[])
        @test t.ref == XLSX.CellRange("A1:B2")
    end

    @testset "autoFilter extends with ref (no totals row)" begin
        f = XLSX.newxlsx()
        sh = f[1]
        sh["A1"] = "a"; sh["B1"] = "b"
        sh["A2"] = 1;   sh["B2"] = 2
        XLSX.addtable!(sh, "A1:B2"; name="T")
        XLSX.appendtable!(sh, "T", [(3, 4)])

        table_doc = XLSX.get_xml_data(XLSX.get_xlsxfile(sh), XLSX._table_part_path(sh, "T"))
        root = XLSX.root_element(table_doc)
        af = XLSX.elements_with_tag(root, "autoFilter")
        @test !isempty(af)
        @test XLSX.get_attr(af[1], "ref") == "A1:B3"  # no totals row: same as ref
    end

    @testset "totals row moves down and is regenerated (sum + label)" begin
        f = XLSX.newxlsx()
        sh = f[1]
        sh["A1"] = "item";   sh["B1"] = "amount"
        sh["A2"] = "Apples"; sh["B2"] = 12
        sh["A3"] = "Pears";  sh["B3"] = 8
        XLSX.addtable!(sh, "A1:B3"; name="Bulk")
        XLSX.settotals!(sh, "Bulk", "item" => "Total", "amount" => :sum)

        t_before = XLSX.table(sh, "Bulk")
        @test t_before.ref == XLSX.CellRange("A1:B4")
        @test sh["A4"] == "Total"

        t = XLSX.appendtable!(sh, "Bulk", [("Cherries", 25), ("Damsons", 5)])

        # ref grew by 2; totals row is still last
        @test t.ref == XLSX.CellRange("A1:B6")
        @test t.has_totals_row == true

        # old totals position (row 4) now holds appended data
        @test sh["A4"] == "Cherries"; @test sh["B4"] == 25
        @test sh["A5"] == "Damsons";  @test sh["B5"] == 5

        # new totals row regenerated at row 6
        @test sh["A6"] == "Total"
        @test occursin("SUBTOTAL(109", XLSX.getFormula(sh, "B6"))
        @test occursin("Bulk[amount]", XLSX.getFormula(sh, "B6"))

        # data body excludes the totals row
        cols = Tables.columns(t)
        @test cols.item == ["Apples", "Pears", "Cherries", "Damsons"]
        @test cols.amount == [12, 8, 25, 5]
    end

    @testset "totals row: autoFilter excludes the totals row after append" begin
        f = XLSX.newxlsx()
        sh = f[1]
        sh["A1"] = "a"; sh["B1"] = "b"
        sh["A2"] = 1;   sh["B2"] = 10
        XLSX.addtable!(sh, "A1:B2"; name="T")
        XLSX.settotals!(sh, "T", "b" => :sum)   # ref becomes A1:B3

        XLSX.appendtable!(sh, "T", [(2, 20), (3, 30)])  # ref becomes A1:B5

        table_doc = XLSX.get_xml_data(XLSX.get_xlsxfile(sh), XLSX._table_part_path(sh, "T"))
        root = XLSX.root_element(table_doc)
        @test XLSX.get_attr(root, "ref") == "A1:B5"
        af = XLSX.elements_with_tag(root, "autoFilter")
        @test XLSX.get_attr(af[1], "ref") == "A1:B4"  # one row short of ref
    end

    @testset "totals row: preserves a custom formula" begin
        f = XLSX.newxlsx()
        sh = f[1]
        sh["A1"] = "revenue"; sh["B1"] = "cost"; sh["C1"] = "margin"
        sh["A2"] = 100;       sh["B2"] = 60;    sh["C2"] = 40
        XLSX.addtable!(sh, "A1:C2"; name="PnL")
        XLSX.settotals!(sh, "PnL",
            "revenue" => :sum,
            "margin"  => (:custom, "SUBTOTAL(109,PnL[revenue])-SUBTOTAL(109,PnL[cost])"),
        )

        f_before = XLSX.getFormula(sh, "C3")

        XLSX.appendtable!(sh, "PnL", [(200, 90, 110)])

        t = XLSX.table(sh, "PnL")
        @test t.ref == XLSX.CellRange("A1:C4")

        # custom formula survived the move, unchanged
        f_after = XLSX.getFormula(sh, "C4")
        @test f_after == f_before
        @test occursin("PnL[revenue]", f_after)
        @test occursin("PnL[cost]", f_after)

        # and the table part still records it as custom
        table_doc = XLSX.get_xml_data(XLSX.get_xlsxfile(sh), XLSX._table_part_path(sh, "PnL"))
        i, j = XLSX.get_idces(table_doc, "table", "tableColumns")
        margin_node = collect(XLSX.xml_elements(table_doc[i][j]))[3]
        @test XLSX.get_attr(margin_node, "totalsRowFunction") == "custom"
    end

    @testset "totals row: preserves every built-in function kind" begin
        f = XLSX.newxlsx()
        sh = f[1]
        cols = ["s", "av", "cnt", "cnta", "mx", "mn"]
        for (i, c) in enumerate(cols)
            sh[1, i] = c
        end
        for r in 2:4, i in 1:length(cols)
            sh[r, i] = r * 10 + i
        end
        XLSX.addtable!(sh, "A1:F4"; name="Funcs")
        XLSX.settotals!(sh, "Funcs"; s=:sum, av=:average, cnt=:countnums, cnta=:count, mx=:max, mn=:min)

        XLSX.appendtable!(sh, "Funcs", [(99, 99, 99, 99, 99, 99)])

        t = XLSX.table(sh, "Funcs")
        @test t.ref == XLSX.CellRange("A1:F6")

        table_doc = XLSX.get_xml_data(XLSX.get_xlsxfile(sh), XLSX._table_part_path(sh, "Funcs"))
        i, j = XLSX.get_idces(table_doc, "table", "tableColumns")
        nodes = collect(XLSX.xml_elements(table_doc[i][j]))

        @test XLSX.get_attr(nodes[1], "totalsRowFunction") == "sum"
        @test XLSX.get_attr(nodes[2], "totalsRowFunction") == "average"
        # NB: OOXML's attribute names for the two count variants are the
        # reverse of Excel's function names — :countnums (Excel COUNT, numeric
        # only) is OOXML "countNums", and :count (Excel COUNTA, non-blank)
        # is OOXML "count".
        @test XLSX.get_attr(nodes[3], "totalsRowFunction") == "countNums"
        @test XLSX.get_attr(nodes[4], "totalsRowFunction") == "count"
        @test XLSX.get_attr(nodes[5], "totalsRowFunction") == "max"
        @test XLSX.get_attr(nodes[6], "totalsRowFunction") == "min"
    end

    @testset "totals row: columns with no totals setting stay empty" begin
        f = XLSX.newxlsx()
        sh = f[1]
        sh["A1"] = "a"; sh["B1"] = "b"; sh["C1"] = "c"
        sh["A2"] = 1;   sh["B2"] = 2;   sh["C2"] = 3
        XLSX.addtable!(sh, "A1:C2"; name="T")
        XLSX.settotals!(sh, "T", "b" => :sum)   # only column b has totals

        XLSX.appendtable!(sh, "T", [(4, 5, 6)])

        t = XLSX.table(sh, "T")
        totals_row = t.ref.stop.row_number
        @test ismissing(sh[XLSX.CellRef(totals_row, 1)])   # a: no totals
        @test !ismissing(XLSX.getFormula(sh, XLSX.CellRef(totals_row, 2)))  # b: sum
        @test ismissing(sh[XLSX.CellRef(totals_row, 3)])   # c: no totals
    end

    @testset "input shapes: vector of tuples, vector of vectors, matrix, Tables.jl source" begin

        sh = fresh()
        t = XLSX.appendtable!(sh, "T", [(2, 20), (3, 30)])
        @test Tables.columns(t).x == [1, 2, 3]

        sh = fresh()
        t = XLSX.appendtable!(sh, "T", [[2, 20], [3, 30]])
        @test Tables.columns(t).x == [1, 2, 3]

        sh = fresh()
        t = XLSX.appendtable!(sh, "T", [2 20; 3 30])
        @test Tables.columns(t).x == [1, 2, 3]

        sh = fresh()
        t = XLSX.appendtable!(sh, "T", DataFrames.DataFrame(x=[2, 3], y=[20, 30]))
        @test Tables.columns(t).x == [1, 2, 3]
    end

    @testset "column count mismatch errors" begin
        f = XLSX.newxlsx()
        sh = f[1]
        sh["A1"] = "a"; sh["B1"] = "b"
        sh["A2"] = 1;   sh["B2"] = 2
        XLSX.addtable!(sh, "A1:B2"; name="T")

        @test_throws XLSX.XLSXError XLSX.appendtable!(sh, "T", [(1, 2, 3)])
        @test_throws XLSX.XLSXError XLSX.appendtable!(sh, "T", [(1,)])
    end

    @testset "refuses non-empty rows below the table; check_empty=false overrides" begin
        f = XLSX.newxlsx()
        sh = f[1]
        sh["A1"] = "a"; sh["B1"] = "b"
        sh["A2"] = 1;   sh["B2"] = 2
        XLSX.addtable!(sh, "A1:B2"; name="T")
        sh["A3"] = "in the way"

        @test_throws XLSX.XLSXError XLSX.appendtable!(sh, "T", [(9, 9)])

        # table unchanged after the failed attempt
        @test XLSX.table(sh, "T").ref == XLSX.CellRange("A1:B2")

        t = XLSX.appendtable!(sh, "T", [(9, 9)]; check_empty=false)
        @test t.ref == XLSX.CellRange("A1:B3")
        @test sh["A3"] == 9
    end

    @testset "table not found errors" begin
        f = XLSX.newxlsx()
        sh = f[1]
        sh["A1"] = "a"; sh["B1"] = "b"
        sh["A2"] = 1;   sh["B2"] = 2
        XLSX.addtable!(sh, "A1:B2"; name="T")

        @test_throws KeyError XLSX.appendtable!(sh, "NoSuchTable", [(1, 2)])
    end

    @testset "appending missing values" begin
        f = XLSX.newxlsx()
        sh = f[1]
        sh["A1"] = "a"; sh["B1"] = "b"
        sh["A2"] = 1;   sh["B2"] = 2
        XLSX.addtable!(sh, "A1:B2"; name="T")

        t = XLSX.appendtable!(sh, "T", [(3, missing), (missing, 6)])
        @test t.ref == XLSX.CellRange("A1:B4")
        @test ismissing(sh["B3"])
        @test ismissing(sh["A4"])
    end

    @testset "two tables on one sheet: appending to one leaves the other alone" begin
        f = XLSX.newxlsx()
        sh = f[1]
        sh["A1"] = "a"; sh["B1"] = "b"
        sh["A2"] = 1;   sh["B2"] = 2
        XLSX.addtable!(sh, "A1:B2"; name="Left")

        sh["D1"] = "c"; sh["E1"] = "d"
        sh["D2"] = 3;   sh["E2"] = 4
        XLSX.addtable!(sh, "D1:E2"; name="Right")

        XLSX.appendtable!(sh, "Left", [(5, 6)])

        @test XLSX.table(sh, "Left").ref == XLSX.CellRange("A1:B3")
        @test XLSX.table(sh, "Right").ref == XLSX.CellRange("D1:E2")  # untouched
        @test length(XLSX.tables(sh)) == 2
    end

    @testset "round trip: append then save and reopen" begin
        f = XLSX.newxlsx()
        sh = f[1]
        sh["A1"] = "item";   sh["B1"] = "amount"
        sh["A2"] = "Apples"; sh["B2"] = 12
        XLSX.addtable!(sh, "A1:B2"; name="Bulk")
        XLSX.settotals!(sh, "Bulk", "item" => "Total", "amount" => :sum)
        XLSX.appendtable!(sh, "Bulk", [("Pears", 8), ("Cherries", 25)])
        SAVE_FILES && save_outfile(f)

        outfile = "appendtable_roundtrip.xlsx"
        isfile(outfile) && rm(outfile)
        XLSX.writexlsx(outfile, f, overwrite=true)
        SAVE_FILES && save_outfile(outfile)

        XLSX.openxlsx(outfile) do xf2
            t = XLSX.table(xf2[1], "Bulk")
            @test t.ref == XLSX.CellRange("A1:B5")
            @test t.has_totals_row == true
            cols = Tables.columns(t)
            @test cols.item == ["Apples", "Pears", "Cherries"]
            @test cols.amount == [12, 8, 25]
        end

        XLSX.openxlsx(outfile, enable_cache=false) do xf2
            t = XLSX.table(xf2[1], "Bulk")
            @test t.ref == XLSX.CellRange("A1:B5")
            @test t.has_totals_row == true
        end

        isfile(outfile) && rm(outfile)
    end

    @testset "repeated appends accumulate correctly" begin
        f = XLSX.newxlsx()
        sh = f[1]
        sh["A1"] = "n"
        sh["A2"] = 1
        XLSX.addtable!(sh, "A1:A2"; name="T")
        XLSX.settotals!(sh, "T", "n" => :sum)

        for i in 2:5
            XLSX.appendtable!(sh, "T", [(i,)])
        end

        t = XLSX.table(sh, "T")
        @test t.ref == XLSX.CellRange("A1:A7")   # header + 5 data + totals
        @test Tables.columns(t).n == [1, 2, 3, 4, 5]
        @test occursin("SUBTOTAL(109", XLSX.getFormula(sh, "A7"))
    end

    @testset "not writable errors" begin
        XLSX.openxlsx("data/two_tables.xlsx") do xf   # read-only
            sh = xf["Sheet1"]
            @test_throws XLSX.XLSXError XLSX.appendtable!(sh, "IO_Table", [(1, 2, 3)])
        end
    end

    @testset "DataFrame columns matched by name, not position" begin
        sh = fresh_abc()

        # deliberately scrambled column order
        df = DataFrames.DataFrame(c=[30, 60], a=[10, 40], b=[20, 50])
        t = XLSX.appendtable!(sh, "T", df)

        @test t.ref == XLSX.CellRange("A1:C4")
        cols = Tables.columns(t)
        @test cols.a == [1, 10, 40]
        @test cols.b == [2, 20, 50]
        @test cols.c == [3, 30, 60]
    end

    @testset "DataTable source works and is name-matched" begin
        sh = fresh_abc()

        # build a DataTable with scrambled column order
        dt = XLSX.DataTable(Any[[30, 60], [10, 40], [20, 50]], [:c, :a, :b])
        t = XLSX.appendtable!(sh, "T", dt)

        @test t.ref == XLSX.CellRange("A1:C4")
        cols = Tables.columns(t)
        @test cols.a == [1, 10, 40]
        @test cols.b == [2, 20, 50]
        @test cols.c == [3, 30, 60]
    end

    @testset "DataTable read from one table appended to another" begin
        # read a table out, then append it to a second table with the same
        # columns in a different order
        f = XLSX.newxlsx()
        sh = f[1]
        sh["A1"] = "x"; sh["B1"] = "y"
        sh["A2"] = 1;   sh["B2"] = 10
        sh["A3"] = 2;   sh["B3"] = 20
        src = XLSX.addtable!(sh, "A1:B3"; name="Source")

        sh["E1"] = "y"; sh["F1"] = "x"      # reversed relative to Source
        sh["E2"] = 99;  sh["F2"] = 9
        XLSX.addtable!(sh, "E1:F2"; name="Dest")

        dt = XLSX.gettable(src)
        t = XLSX.appendtable!(sh, "Dest", dt)

        @test t.ref == XLSX.CellRange("E1:F4")
        cols = Tables.columns(t)
        @test cols.y == [99, 10, 20]   # matched by name despite reversal
        @test cols.x == [9, 1, 2]
    end

    @testset "missing column errors" begin
        sh = fresh_abc()
        df = DataFrames.DataFrame(a=[10], b=[20])   # no "c"
        @test_throws XLSX.XLSXError XLSX.appendtable!(sh, "T", df)
    end

    @testset "extra column errors" begin
        sh = fresh_abc()
        df = DataFrames.DataFrame(a=[10], b=[20], c=[30], d=[40])   # extra "d"
        @test_throws XLSX.XLSXError XLSX.appendtable!(sh, "T", df)
    end

    @testset "both missing and extra columns errors" begin
        sh = fresh_abc()
        df = DataFrames.DataFrame(a=[10], b=[20], z=[99])   # missing c, extra z
        @test_throws XLSX.XLSXError XLSX.appendtable!(sh, "T", df)
    end

    @testset "unnamed sources remain positional" begin
        # matrix
        sh = fresh_abc()
        t = XLSX.appendtable!(sh, "T", [10 20 30])
        cols = Tables.columns(t)
        @test cols.a == [1, 10]
        @test cols.b == [2, 20]
        @test cols.c == [3, 30]

        # vector of tuples
        sh = fresh_abc()
        t = XLSX.appendtable!(sh, "T", [(10, 20, 30)])
        cols = Tables.columns(t)
        @test cols.a == [1, 10]
        @test cols.c == [3, 30]

        # vector of vectors
        sh = fresh_abc()
        t = XLSX.appendtable!(sh, "T", [[10, 20, 30]])
        cols = Tables.columns(t)
        @test cols.a == [1, 10]
        @test cols.c == [3, 30]
    end

    @testset "name matching still works when a totals row is present" begin
        f = XLSX.newxlsx()
        sh = f[1]
        sh["A1"] = "item";   sh["B1"] = "amount"
        sh["A2"] = "Apples"; sh["B2"] = 12
        XLSX.addtable!(sh, "A1:B2"; name="Bulk")
        XLSX.settotals!(sh, "Bulk", "item" => "Total", "amount" => :sum)

        # reversed column order in the source
        df = DataFrames.DataFrame(amount=[8, 25], item=["Pears", "Cherries"])
        t = XLSX.appendtable!(sh, "Bulk", df)

        @test t.ref == XLSX.CellRange("A1:B5")
        @test t.has_totals_row == true
        cols = Tables.columns(t)
        @test cols.item == ["Apples", "Pears", "Cherries"]
        @test cols.amount == [12, 8, 25]
        @test occursin("SUBTOTAL(109", XLSX.getFormula(sh, "B5"))
    end

    @testset "readto(source, sheet, sink; table_name=...)" begin
        outfile = make_readto_file("readto_tablename_scoped.xlsx")
        SAVE_FILES && save_outfile(outfile)

        df = XLSX.readto(outfile, "Second", DataFrames.DataFrame; table_name="TableTwo")

        @test df isa DataFrames.DataFrame
        @test size(df) == (3, 3)
        @test DataFrames.names(df) == ["id", "name", "score"]
        @test df.id == [1, 2, 3]
        @test df.name == ["alice", "bob", "carol"]
        @test df.score == [10.5, 20.0, 15.25]
        @test eltype(df.score) != Any

        isfile(outfile) && rm(outfile)
    end

    @testset "readto(source, sink; table_name=...) - workbook-wide" begin
        outfile = make_readto_file("readto_tablename_wide.xlsx")

        # table on the second sheet, found without naming the sheet
        df = XLSX.readto(outfile, DataFrames.DataFrame; table_name="TableTwo")
        @test size(df) == (3, 3)
        @test df.id == [1, 2, 3]

        # table on the first sheet
        df1 = XLSX.readto(outfile, DataFrames.DataFrame; table_name="TableOne")
        @test size(df1) == (2, 2)
        @test df1.a == [1, 2]

        isfile(outfile) && rm(outfile)
    end

    @testset "readto by sheet index" begin
        outfile = make_readto_file("readto_tablename_idx.xlsx")

        df = XLSX.readto(outfile, 2, DataFrames.DataFrame; table_name="TableTwo")
        @test DataFrames.names(df) == ["id", "name", "score"]
        @test df.id == [1, 2, 3]

        isfile(outfile) && rm(outfile)
    end

    @testset "readto excludes data outside the table's ref" begin
        outfile = make_readto_file("readto_tablename_bounds.xlsx")

        df = XLSX.readto(outfile, "Second", DataFrames.DataFrame; table_name="TableTwo")
        @test DataFrames.ncol(df) == 3
        @test "outside" ∉ DataFrames.names(df)
        @test DataFrames.nrow(df) == 3

        isfile(outfile) && rm(outfile)
    end

    @testset "readto scoped and workbook-wide agree" begin
        outfile = make_readto_file("readto_tablename_agree.xlsx")

        df_scoped = XLSX.readto(outfile, "Second", DataFrames.DataFrame; table_name="TableTwo")
        df_wide   = XLSX.readto(outfile, DataFrames.DataFrame; table_name="TableTwo")
        @test isequal(df_scoped, df_wide)

        isfile(outfile) && rm(outfile)
    end

    @testset "readto with enable_cache=false" begin
        outfile = make_readto_file("readto_tablename_nocache.xlsx")

        df_scoped = XLSX.readto(outfile, "Second", DataFrames.DataFrame;
                                table_name="TableTwo", enable_cache=false)
        @test size(df_scoped) == (3, 3)
        @test df_scoped.name == ["alice", "bob", "carol"]

        df_wide = XLSX.readto(outfile, DataFrames.DataFrame;
                              table_name="TableTwo", enable_cache=false)
        @test isequal(df_scoped, df_wide)

        # and agrees with the cached read
        df_cached = XLSX.readto(outfile, DataFrames.DataFrame; table_name="TableTwo")
        @test isequal(df_cached, df_wide)

        isfile(outfile) && rm(outfile)
    end

    @testset "readto passes normalizenames / missing_strings through" begin
        outfile = "readto_tablename_kwargs.xlsx"
        isfile(outfile) && rm(outfile)

        f = XLSX.newxlsx()
        sh = f[1]
        sh["A1"] = "col one"; sh["B1"] = "value"
        sh["A2"] = "x";       sh["B2"] = 1.5
        sh["A3"] = "N/A";     sh["B3"] = 2.5
        XLSX.addtable!(sh, "A1:B3"; name="T")
        XLSX.writexlsx(outfile, f, overwrite=true)

        df_norm = XLSX.readto(outfile, DataFrames.DataFrame; table_name="T", normalizenames=true)
        @test DataFrames.names(df_norm)[1] == "col_one"

        df_miss = XLSX.readto(outfile, DataFrames.DataFrame; table_name="T", missing_strings="N/A")
        @test ismissing(df_miss[2, 1])

        isfile(outfile) && rm(outfile)
    end

    @testset "readto with a totals row: totals excluded" begin
        outfile = "readto_tablename_totals.xlsx"
        isfile(outfile) && rm(outfile)

        f = XLSX.newxlsx()
        sh = f[1]
        sh["A1"] = "item";   sh["B1"] = "amount"
        sh["A2"] = "Apples"; sh["B2"] = 12
        sh["A3"] = "Pears";  sh["B3"] = 8
        XLSX.addtable!(sh, "A1:B3"; name="Bulk")
        XLSX.settotals!(sh, "Bulk", "item" => "Total", "amount" => :sum)
        XLSX.writexlsx(outfile, f, overwrite=true)

        df = XLSX.readto(outfile, DataFrames.DataFrame; table_name="Bulk")
        @test DataFrames.nrow(df) == 2
        @test df.item == ["Apples", "Pears"]
        @test df.amount == [12, 8]

        isfile(outfile) && rm(outfile)
    end

    @testset "readto errors: table_name with a columns range" begin
        outfile = make_readto_file("readto_tablename_conflict.xlsx")

        @test_throws XLSX.XLSXError XLSX.readto(outfile, "Second", "A:C",
                                                DataFrames.DataFrame; table_name="TableTwo")

        isfile(outfile) && rm(outfile)
    end

    @testset "readto errors: table not found" begin
        outfile = make_readto_file("readto_tablename_notfound.xlsx")

        @test_throws XLSX.XLSXError XLSX.readto(outfile, DataFrames.DataFrame; table_name="NoSuchTable")
        @test_throws KeyError XLSX.readto(outfile, "First", DataFrames.DataFrame; table_name="TableTwo")

        isfile(outfile) && rm(outfile)
    end

    @testset "readto errors: missing sink" begin
        outfile = make_readto_file("readto_tablename_nosink.xlsx")

        @test_throws XLSX.XLSXError XLSX.readto(outfile; table_name="TableTwo")
        @test_throws XLSX.XLSXError XLSX.readto(outfile, "Second"; table_name="TableTwo")

        isfile(outfile) && rm(outfile)
    end

    @testset "readto agrees with readtable for the same table" begin
        outfile = make_readto_file("readto_tablename_vs_readtable.xlsx")

        df_readto = XLSX.readto(outfile, "Second", DataFrames.DataFrame; table_name="TableTwo")
        df_readtable = DataFrames.DataFrame(XLSX.readtable(outfile, "Second"; table_name="TableTwo"))
        @test isequal(df_readto, df_readtable)

        isfile(outfile) && rm(outfile)
    end

    @testset "existing readto behaviour unaffected when table_name omitted" begin
        outfile = make_readto_file("readto_tablename_regression.xlsx")

        df = XLSX.readto(outfile, "First", DataFrames.DataFrame)
        @test df isa DataFrames.DataFrame
        @test DataFrames.names(df) == ["a", "b"]

        df_cols = XLSX.readto(outfile, "Second", "A:B", DataFrames.DataFrame)
        @test DataFrames.names(df_cols) == ["id", "name"]

        isfile(outfile) && rm(outfile)
    end

    @testset "index by integer, symbol and string" begin
        f = XLSX.newxlsx()
        sh = f[1]
        sh["A1"] = "id"; sh["B1"] = "name"; sh["C1"] = "score"
        sh["A2"] = 1;    sh["B2"] = "alice"; sh["C2"] = 10.5
        sh["A3"] = 2;    sh["B3"] = "bob";   sh["C3"] = 20.0

        t = XLSX.addtable!(sh, "A1:C3"; name="People")
        rows = collect(XLSX.eachtablerow(t))

        # by integer position
        @test rows[1][1] == 1
        @test rows[1][2] == "alice"
        @test rows[1][3] == 10.5
        @test rows[2][1] == 2

        # by symbol
        @test rows[1][:id] == 1
        @test rows[1][:name] == "alice"
        @test rows[1][:score] == 10.5
        @test rows[2][:name] == "bob"

        # by string
        @test rows[1]["id"] == 1
        @test rows[1]["name"] == "alice"
        @test rows[2]["score"] == 20.0
    end

    @testset "all three index forms agree" begin
        f = XLSX.newxlsx()
        sh = f[1]
        sh["A1"] = "a"; sh["B1"] = "b"
        sh["A2"] = 10;  sh["B2"] = 20

        t = XLSX.addtable!(sh, "A1:B2"; name="T")
        r = first(XLSX.eachtablerow(t))

        @test r[1] == r[:a] == r["a"]
        @test r[2] == r[:b] == r["b"]
    end

    @testset "getindex agrees with Tables.getcolumn" begin
        f = XLSX.newxlsx()
        sh = f[1]
        sh["A1"] = "x"; sh["B1"] = "y"
        sh["A2"] = 1;   sh["B2"] = "one"
        sh["A3"] = 2;   sh["B3"] = "two"

        t = XLSX.addtable!(sh, "A1:B3"; name="T")

        for r in XLSX.eachtablerow(t)
            @test r[1] == Tables.getcolumn(r, 1)
            @test r[:x] == Tables.getcolumn(r, :x)
            @test r["y"] == Tables.getcolumn(r, :y)
        end
    end

    @testset "column names with spaces are reachable by string" begin
        f = XLSX.newxlsx()
        sh = f[1]
        sh["A1"] = "col one"; sh["B1"] = "col two"
        sh["A2"] = 1;         sh["B2"] = 2

        t = XLSX.addtable!(sh, "A1:B2"; name="T")
        r = first(XLSX.eachtablerow(t))

        @test r["col one"] == 1
        @test r["col two"] == 2
        @test r[Symbol("col one")] == 1   # symbol form still works, just awkward
    end

    @testset "missing values via getindex" begin
        f = XLSX.newxlsx()
        sh = f[1]
        sh["A1"] = "a"; sh["B1"] = "b"
        sh["A2"] = 1;   sh["B2"] = missing
        sh["A3"] = missing; sh["B3"] = 4

        t = XLSX.addtable!(sh, "A1:B3"; name="T")
        rows = collect(XLSX.eachtablerow(t))

        @test ismissing(rows[1][:b])
        @test ismissing(rows[1][2])
        @test ismissing(rows[2]["a"])
    end

    @testset "totals row excluded when indexing rows" begin
        f = XLSX.newxlsx()
        sh = f[1]
        sh["A1"] = "item";   sh["B1"] = "amount"
        sh["A2"] = "Apples"; sh["B2"] = 12
        sh["A3"] = "Pears";  sh["B3"] = 8
        XLSX.addtable!(sh, "A1:B3"; name="Bulk")
        t = XLSX.settotals!(sh, "Bulk", "item" => "Total", "amount" => :sum)

        rows = collect(XLSX.eachtablerow(t))
        @test length(rows) == 2
        @test rows[end][:item] == "Pears"   # not "Total"
        @test rows[end][:amount] == 8
    end

    @testset "real fixture (two_tables.xlsx)" begin
        XLSX.openxlsx("data/two_tables.xlsx") do xf
            sh = xf["Sheet1"]
            t = XLSX.table(sh, "IO_Table")
            r = first(XLSX.eachtablerow(t))

            for (i, colname) in enumerate(t.columns)
                @test isequal(r[i], r[Symbol(colname)])
                @test isequal(r[i], r[colname])
            end
        end
    end


# ---------------------------------------------------------------------------
# Test doubles for `_normalize_append_rows`' non-schema paths.
# ---------------------------------------------------------------------------


# ===========================================================================
# deletetable! — totals-row formula rewriting (largest uncovered block)
# ===========================================================================

@testset "deletetable! - totals formulas rewritten to static ranges" begin
    f = XLSX.newxlsx()
    sh = f[1]
    sh["A1"] = "revenue"; sh["B1"] = "cost"; sh["C1"] = "margin"
    sh["A2"] = 100;       sh["B2"] = 60;     sh["C2"] = 40
    sh["A3"] = 200;       sh["B3"] = 90;     sh["C3"] = 110

    XLSX.addtable!(sh, "A1:C3"; name="PnL")
    XLSX.settotals!(sh, "PnL",
        "revenue" => :sum,
        "cost"    => "Total",   # label only: no formula -> exercises the skip path
        "margin"  => (:custom, "SUBTOTAL(109,PnL[revenue])-SUBTOTAL(109,PnL[cost])"),
    )

    @test XLSX.table(sh, "PnL").ref == XLSX.CellRange("A1:C4")

    XLSX.deletetable!(sh, "PnL")
    @test isempty(XLSX.tables(sh))

    # Single-column reference: PnL[revenue] -> Sheet1!$A$2:$A$3
    fa = XLSX.getFormula(sh, "A4")
    @test !occursin("PnL[", fa)
    @test occursin("\$A\$2:\$A\$3", fa)
    @test occursin(sh.name, fa)

    # Custom formula referencing *two* columns: both must be substituted,
    # and the surrounding formula left intact.
    fc = XLSX.getFormula(sh, "C4")
    @test !occursin("PnL[", fc)
    @test occursin("\$A\$2:\$A\$3", fc)
    @test occursin("\$B\$2:\$B\$3", fc)
    @test occursin("SUBTOTAL(109,", fc)
    @test count("SUBTOTAL(109,", fc) == 2

    # Label column had no formula: value untouched, nothing rewritten.
    @test sh["B4"] == "Total"

    # Cell data survives the delete, as documented.
    @test sh["A2"] == 100
    @test sh["C3"] == 110
end

@testset "deletetable! - totals row with a column that has no totals at all" begin
    # Covers the `c isa EmptyCell` half of the skip condition: column "b"
    # has a genuinely empty totals cell.
    f = XLSX.newxlsx()
    sh = f[1]
    sh["A1"] = "a"; sh["B1"] = "b"
    sh["A2"] = 1;   sh["B2"] = 2

    XLSX.addtable!(sh, "A1:B2"; name="T")
    XLSX.settotals!(sh, "T", "a" => :sum)   # b left with no totals content

    XLSX.deletetable!(sh, "T")

    @test !occursin("T[", XLSX.getFormula(sh, "A3"))
    @test ismissing(sh["B3"])
    @test isempty(XLSX.tables(sh))
end

@testset "deletetable! - sheet name needing quoting" begin
    # Exercises quoteit's quoting branch in the static-range substitution.
    f = XLSX.newxlsx()
    sh = XLSX.addsheet!(f, "My Data")
    sh["A1"] = "amount"
    sh["A2"] = 10
    sh["A3"] = 20

    XLSX.addtable!(sh, "A1:A3"; name="Amounts")
    XLSX.settotals!(sh, "Amounts", "amount" => :sum)
    XLSX.deletetable!(sh, "Amounts")

    fa = XLSX.getFormula(sh, "A4")
    @test !occursin("Amounts[", fa)
    @test occursin("My Data", fa)
    @test occursin("\$A\$2:\$A\$3", fa)
end

@testset "deletetable! - round trip after totals rewrite" begin
    f = XLSX.newxlsx()
    sh = f[1]
    sh["A1"] = "n"
    sh["A2"] = 1
    sh["A3"] = 2
    XLSX.addtable!(sh, "A1:A3"; name="T")
    XLSX.settotals!(sh, "T", "n" => :sum)
    XLSX.deletetable!(sh, "T")
    SAVE_FILES && save_outfile(f)

    outfile = "deletetable_totals_roundtrip.xlsx"
    isfile(outfile) && rm(outfile)
    XLSX.writexlsx(outfile, f, overwrite=true)
    SAVE_FILES && save_outfile(outfile)

    XLSX.openxlsx(outfile) do xf2
        sh2 = xf2[1]
        @test isempty(XLSX.tables(sh2))
        @test !occursin("T[", XLSX.getFormula(sh2, "A4"))
        @test sh2["A2"] == 1
    end

    isfile(outfile) && rm(outfile)
end


# ===========================================================================
# settotals! — :none
# ===========================================================================

@testset "settotals! - :none clears a column's totals, row survives" begin
    f = XLSX.newxlsx()
    sh = f[1]
    sh["A1"] = "a"; sh["B1"] = "b"
    sh["A2"] = 1;   sh["B2"] = 2

    XLSX.addtable!(sh, "A1:B2"; name="T")
    XLSX.settotals!(sh, "T", "a" => :sum, "b" => "Label")
    @test sh["B3"] == "Label"

    t = XLSX.settotals!(sh, "T", "a" => :none, "b" => :none)

    func_a, label_a = _totals_col_attrs(sh, "T", "a")
    func_b, label_b = _totals_col_attrs(sh, "T", "b")
    @test isnothing(func_a); @test isnothing(label_a)
    @test isnothing(func_b); @test isnothing(label_b)
    @test ismissing(sh["A3"])
    @test ismissing(sh["B3"])

    # an entirely empty totals row is still a totals row
    @test t.has_totals_row == true
    @test t.ref == XLSX.CellRange("A1:B3")
end

@testset "settotals! - :none on a table with no totals row still adds one" begin
    f = XLSX.newxlsx()
    sh = f[1]
    sh["A1"] = "a"; sh["B1"] = "b"
    sh["A2"] = 1;   sh["B2"] = 2
    XLSX.addtable!(sh, "A1:B2"; name="T")

    t = XLSX.settotals!(sh, "T", "a" => :none)
    @test t.has_totals_row == true
    @test t.ref == XLSX.CellRange("A1:B3")
    @test ismissing(sh["A3"])
end

@testset "appendtable! - :none columns stay clear after the totals row moves" begin
    f = XLSX.newxlsx()
    sh = f[1]
    sh["A1"] = "a"; sh["B1"] = "b"
    sh["A2"] = 1;   sh["B2"] = 2
    XLSX.addtable!(sh, "A1:B2"; name="T")
    XLSX.settotals!(sh, "T", "a" => :sum, "b" => :none)

    XLSX.appendtable!(sh, "T", [(3, 4)])

    t = XLSX.table(sh, "T")
    totals_row = t.ref.stop.row_number
    @test occursin("SUBTOTAL(109", XLSX.getFormula(sh, XLSX.CellRef(totals_row, 1)))
    @test ismissing(sh[XLSX.CellRef(totals_row, 2)])
end


# ===========================================================================
# Workbook-scoped lookups (never compiled -> absent from lcov entirely)
# ===========================================================================

@testset "tables(xf) - every sheet, in order" begin
    XLSX.openxlsx("data/two_tables.xlsx") do xf
        all_t = XLSX.tables(xf)

        @test length(all_t) == 3
        @test [t.name for t in all_t] == ["IO_Table", "Age_height", "with_total"]
        @test [t.sheet.name for t in all_t] == ["Sheet1", "Sheet1", "Sheet2"]
        @test all(t -> t isa XLSX.Table, all_t)
    end
end

@testset "tables(xf) - workbook with no tables" begin
    xf = XLSX.newxlsx()
    @test XLSX.tables(xf) == XLSX.Table[]
end

@testset "table(xf, name) / table(xf, id)" begin
    XLSX.openxlsx("data/two_tables.xlsx") do xf
        t = XLSX.table(xf, "with_total")
        @test t.sheet.name == "Sheet2"
        @test t.ref == XLSX.CellRange("A3:C11")

        @test XLSX.table(xf, "IO_Table").id == 1
        @test XLSX.table(xf, 2).name == "Age_height"
        @test XLSX.table(xf, 3).name == "with_total"

        @test_throws KeyError XLSX.table(xf, "NoSuchTable")
        @test_throws KeyError XLSX.table(xf, 999)
    end
end

@testset "table(wb, name) / table(wb, id)" begin
    XLSX.openxlsx("data/two_tables.xlsx") do xf
        wb = XLSX.get_workbook(xf)

        @test XLSX.table(wb, "Age_height").ref == XLSX.CellRange("E1:G5")
        @test XLSX.table(wb, 1).name == "IO_Table"

        @test_throws KeyError XLSX.table(wb, "NoSuchTable")
        @test_throws KeyError XLSX.table(wb, 999)
    end
end

@testset "workbook-wide lookup agrees with sheet-scoped" begin
    XLSX.openxlsx("data/two_tables.xlsx") do xf
        @test tables_equal(XLSX.table(xf, "with_total"),
                           XLSX.table(xf["Sheet2"], "with_total"))
    end
end


# ===========================================================================
# _normalize_append_rows — non-schema Tables.jl sources
# ===========================================================================

@testset "appendtable! - source without a schema, matched by columnnames" begin
    sh = fresh_abc()   # table "T" over A1:C2, columns a, b, c

    # scrambled order, and no Tables.schema — forces the columnnames fallback
    src = NoSchemaCols((c=[30, 60], a=[10, 40], b=[20, 50]))
    t = XLSX.appendtable!(sh, "T", src)

    @test t.ref == XLSX.CellRange("A1:C4")
    cols = Tables.columns(t)
    @test cols.a == [1, 10, 40]
    @test cols.b == [2, 20, 50]
    @test cols.c == [3, 30, 60]
end

@testset "appendtable! - schemaless source with wrong columns still errors" begin
    sh = fresh_abc()
    @test_throws XLSX.XLSXError XLSX.appendtable!(sh, "T", NoSchemaCols((a=[1], b=[2])))
    @test_throws XLSX.XLSXError XLSX.appendtable!(sh, "T",
        NoSchemaCols((a=[1], b=[2], c=[3], d=[4])))
end

@testset "appendtable! - nameless Tables.jl source falls back to positional" begin
    sh = fresh_abc()

    src = NoNamesSource([Any[10, 20, 30], Any[40, 50, 60]])
    t = XLSX.appendtable!(sh, "T", src)

    @test t.ref == XLSX.CellRange("A1:C4")
    cols = Tables.columns(t)
    @test cols.a == [1, 10, 40]
    @test cols.b == [2, 20, 50]
    @test cols.c == [3, 30, 60]
end


# ===========================================================================
# Writability guards
# ===========================================================================

@testset "read-only file rejects every mutating table call" begin
    XLSX.openxlsx("data/two_tables.xlsx") do xf
        sh = xf["Sheet1"]

        @test_throws XLSX.XLSXError XLSX.addtable!(sh, "J1:K2"; name="Nope")
        @test_throws XLSX.XLSXError XLSX.deletetable!(sh, "IO_Table")
        @test_throws XLSX.XLSXError XLSX.deletetable!(sh, 1)
        @test_throws XLSX.XLSXError XLSX.settotals!(sh, "IO_Table", "id" => :sum)
        @test_throws XLSX.XLSXError XLSX.settotals!(sh, 1, "id" => :sum)

        # nothing was mutated by the failed attempts
        @test length(XLSX.tables(sh)) == 2
    end
end


# ===========================================================================
# Parser error paths (branch coverage on `cond && throw(...)` one-liners)
# ===========================================================================

@testset "parse_table_xml - malformed parts" begin
    f = XLSX.newxlsx()
    sh = f[1]

    # root element is not <table>
    @test_throws XLSX.XLSXError XLSX.parse_table_xml(
        _tbl_doc("<notATable/>"), "bad.xml", sh)

    # <table> with no attributes at all
    @test_throws XLSX.XLSXError XLSX.parse_table_xml(
        _tbl_doc("""<table xmlns="$_TBL_NS"/>"""), "bad.xml", sh)

    # missing each required attribute in turn
    @test_throws XLSX.XLSXError XLSX.parse_table_xml(
        _tbl_doc("""<table xmlns="$_TBL_NS" name="T" ref="A1:B2"/>"""), "bad.xml", sh)
    @test_throws XLSX.XLSXError XLSX.parse_table_xml(
        _tbl_doc("""<table xmlns="$_TBL_NS" id="1" ref="A1:B2"/>"""), "bad.xml", sh)
    @test_throws XLSX.XLSXError XLSX.parse_table_xml(
        _tbl_doc("""<table xmlns="$_TBL_NS" id="1" name="T"/>"""), "bad.xml", sh)

    # no <tableColumns> element
    @test_throws XLSX.XLSXError XLSX.parse_table_xml(
        _tbl_doc("""<table xmlns="$_TBL_NS" id="1" name="T" ref="A1:B2"/>"""),
        "bad.xml", sh)
end

@testset "parse_table_columns - <tableColumn> without a name" begin
    doc = _tbl_doc("""<table xmlns="$_TBL_NS" id="1" name="T" ref="A1:B2">
        <tableColumns count="1"><tableColumn id="1"/></tableColumns></table>""")
    @test_throws XLSX.XLSXError XLSX.parse_table_columns(doc)
end

@testset "parse_table_columns - non-tableColumn children are skipped" begin
    doc = _tbl_doc("""<table xmlns="$_TBL_NS" id="1" name="T" ref="A1:B2">
        <tableColumns count="2"><tableColumn id="1" name="a"/>
        <somethingElse/><tableColumn id="2" name="b"/></tableColumns></table>""")
    @test XLSX.parse_table_columns(doc) == ["a", "b"]
end

@testset "parse_table_xml - totalsRowShown without totalsRowCount" begin
    # Excel's own writer prefers totalsRowCount; hand-authored files often
    # carry only totalsRowShown. Both must be recognised.
    f = XLSX.newxlsx()
    sh = f[1]
    base = """<tableColumns count="1"><tableColumn id="1" name="a"/></tableColumns>"""

    shown = XLSX.parse_table_xml(_tbl_doc(
        """<table xmlns="$_TBL_NS" id="1" name="T" ref="A1:A3" totalsRowShown="1">$base</table>"""),
        "t.xml", sh)
    @test shown.has_totals_row == true

    count_only = XLSX.parse_table_xml(_tbl_doc(
        """<table xmlns="$_TBL_NS" id="1" name="T" ref="A1:A3" totalsRowCount="1">$base</table>"""),
        "t.xml", sh)
    @test count_only.has_totals_row == true

    neither = XLSX.parse_table_xml(_tbl_doc(
        """<table xmlns="$_TBL_NS" id="1" name="T" ref="A1:A3">$base</table>"""),
        "t.xml", sh)
    @test neither.has_totals_row == false

    # displayName defaults to name when absent
    @test neither.display_name == "T"
end

@testset "parse_table_style_info - element present but bare" begin
    doc = _tbl_doc("""<table xmlns="$_TBL_NS" id="1" name="T" ref="A1:B2">
        <tableColumns count="1"><tableColumn id="1" name="a"/></tableColumns>
        <tableStyleInfo/></table>""")
    s = XLSX.parse_table_style_info(doc)

    @test !isnothing(s)
    @test isnothing(s.name)
    @test s.show_first_column == false
    @test s.show_last_column == false
    @test s.show_row_stripes == false
    @test s.show_column_stripes == false
end

@testset "parse_table_style_info - absent returns nothing" begin
    doc = _tbl_doc("""<table xmlns="$_TBL_NS" id="1" name="T" ref="A1:B2">
        <tableColumns count="1"><tableColumn id="1" name="a"/></tableColumns></table>""")
    @test isnothing(XLSX.parse_table_style_info(doc))
end

@testset "parse_totals_settings - unrecognized totalsRowFunction errors" begin
    f = XLSX.newxlsx()
    sh = f[1]
    sh["A1"] = "a"; sh["B1"] = "b"
    sh["A2"] = 1;   sh["B2"] = 2
    XLSX.addtable!(sh, "A1:B2"; name="T")
    XLSX.settotals!(sh, "T", "b" => :sum)

    table_doc = XLSX.get_xml_data(XLSX.get_xlsxfile(sh), XLSX._table_part_path(sh, "T"))
    i, j = XLSX.get_idces(table_doc, "table", "tableColumns")
    node = collect(XLSX.xml_elements(table_doc[i][j]))[2]
    node["totalsRowFunction"] = "notAFunction"
    sh.tables_cache = nothing

    @test_throws XLSX.XLSXError XLSX.parse_totals_settings(sh, XLSX.table(sh, "T"))
end

@testset "parse_totals_settings - custom function with no formula errors" begin
    f = XLSX.newxlsx()
    sh = f[1]
    sh["A1"] = "a"; sh["B1"] = "b"
    sh["A2"] = 1;   sh["B2"] = 2
    XLSX.addtable!(sh, "A1:B2"; name="T")
    XLSX.settotals!(sh, "T", "b" => (:custom, "SUBTOTAL(109,T[a])"))

    # blank the formula cell while leaving totalsRowFunction="custom" in place
    sh["B3"] = missing
    sh.tables_cache = nothing

    @test_throws XLSX.XLSXError XLSX.parse_totals_settings(sh, XLSX.table(sh, "T"))
end

@testset "parse_totals_settings - label-only column is returned as a String" begin
    f = XLSX.newxlsx()
    sh = f[1]
    sh["A1"] = "a"; sh["B1"] = "b"
    sh["A2"] = 1;   sh["B2"] = 2
    XLSX.addtable!(sh, "A1:B2"; name="T")
    XLSX.settotals!(sh, "T", "a" => "Total", "b" => :sum)

    settings = Dict(XLSX.parse_totals_settings(sh, XLSX.table(sh, "T")))
    @test settings["a"] == "Total"
    @test settings["b"] === :sum
end


# ===========================================================================
# Small helpers
# ===========================================================================

@testset "_is_valid_table_display_name" begin
    @test XLSX._is_valid_table_display_name("Table1")
    @test XLSX._is_valid_table_display_name("_leading")
    @test XLSX._is_valid_table_display_name("with.periods")
    @test XLSX._is_valid_table_display_name("Ünicode")

    @test !XLSX._is_valid_table_display_name("")
    @test !XLSX._is_valid_table_display_name("1leading")
    @test !XLSX._is_valid_table_display_name("has space")
    @test !XLSX._is_valid_table_display_name("has-hyphen")
end

@testset "remove_attr! - node with no attributes, and absent key" begin
    bare = XLSX.XML.Element("x")
    @test isnothing(XLSX.remove_attr!(bare, "anything"))

    node = XLSX.XML.Element("x"; a="1", b="2")
    @test isnothing(XLSX.remove_attr!(node, "notThere"))
    @test XLSX.get_attr(node, "a") == "1"
    @test XLSX.get_attr(node, "b") == "2"

    XLSX.remove_attr!(node, "a")
    @test XLSX.get_attr(node, "a", "") == ""
    @test !haskey(XML.attributes(node), "a")
    @test XLSX.get_attr(node, "b") == "2"end

@testset "TableRow getindex with a vector or range of columns" begin
    f = XLSX.newxlsx()
    sh = f[1]
    sh["A1"] = "a"; sh["B1"] = "b"; sh["C1"] = "c"
    sh["A2"] = 1;   sh["B2"] = 2;   sh["C2"] = 3

    r = first(XLSX.eachtablerow(sh))
    @test r[1:2] == [1, 2]
    @test r[[1, 3]] == [1, 3]
    @test r[1:3] == [1, 2, 3]
end

@testset "eachtablerow(sheet) - leading row whose cells all read missing" begin
    # Covers the `isnothing(ci)` skip: row 1 has cells present in the XML but
    # every value reads back as `missing`, so it can't anchor the table.
    f = XLSX.newxlsx()
    sh = f[1]
    sh["A1"] = 1; sh["B1"] = 2
    sh["A1"] = missing; sh["B1"] = missing
    sh["A2"] = "a"; sh["B2"] = "b"
    sh["A3"] = 1;   sh["B3"] = 2

    dt = XLSX.gettable(sh)
    @test dt.column_labels == [:a, :b]
    @test dt.data[1] == [1]
end
end
