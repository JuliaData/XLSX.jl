#=
Streaming writes: `streamxlsx`, `writerow!` and `writetable(...; streaming=true)`,
which write rows straight into the file rather than through the cell cache. The
cached writer is the reference: a streamed table reads back as the same table.
=#

# The parts of an xlsx held in `bytes`, by name.
function stream_parts(bytes::Vector{UInt8})
    z = ZipReader(bytes)
    return Dict(n => zip_readentry(z, n, String) for n in zip_names(z))
end

function stream_written(f)
    io = IOBuffer()
    f(io)
    return take!(io)
end

const STREAM_MIXED_COLUMNS = Any[
    [1, 2, 3, 4],
    [1.5, NaN, -Inf, 2.25],
    ["plain", " padded ", "line\nbreak", "bad\u0001char"],
    [Date(2024, 1, 1), missing, Date(1900, 3, 1), Date(2024, 2, 29)],
    [DateTime(2024, 1, 1, 12), DateTime(2024, 1, 2), nothing, DateTime(2024, 1, 3, 23, 59, 59)],
    [true, false, missing, true],
    [Time(12, 30), Time(0), Time(23, 59, 59), missing],
    [:sym, 'c', 1 // 2, Int32(7)],
    ["", "a&b<c>\"d\"", missing, "ünïcödé"],
]
const STREAM_MIXED_NAMES = ["int", "float", "string", "date", "datetime", "bool", "time", "other", "text"]

@testset "same table as the cached writer" begin
    cached = stream_written(io -> XLSX.writetable(io, STREAM_MIXED_COLUMNS, STREAM_MIXED_NAMES))
    reference = XLSX.readtable(IOBuffer(cached), 1)
    for shared_strings in (false, true)
        streamed = stream_written(io -> XLSX.writetable(io, STREAM_MIXED_COLUMNS, STREAM_MIXED_NAMES;
                                                        streaming=true, shared_strings))
        t = XLSX.readtable(IOBuffer(streamed), 1)
        @test t.column_labels == reference.column_labels
        @test isequal(t.data, reference.data)
        # read again without the cache, as streamed reads see it
        XLSX.openxlsx(IOBuffer(streamed); enable_cache=false) do xf
            @test isequal(XLSX.gettable(xf[1]).data, reference.data)
        end
        parts = stream_parts(streamed)
        @test haskey(parts, "xl/sharedStrings.xml") == shared_strings
        @test occursin("t=\"inlineStr\"", parts["xl/worksheets/sheet1.xml"]) == !shared_strings
        @test occursin("<is><t xml:space=\"preserve\"> padded </t></is>", parts["xl/worksheets/sheet1.xml"]) == !shared_strings
    end

    # to a file, and a Tables.jl table by columns
    f = "streamed_table.xlsx"
    df = DataFrames.DataFrame(STREAM_MIXED_COLUMNS, Symbol.(STREAM_MIXED_NAMES))
    @test XLSX.writetable(f, df; streaming=true, overwrite=true) == abspath(f)
    SAVE_FILES && save_outfile(f)
    @test isequal(XLSX.readtable(f, 1).data, reference.data)
    isfile(f) && rm(f)
end

@testset "byte-identical to the cached writer" begin
    # With shared strings and no empty cells, every part of the package matches.
    n = 2 * XLSX.ROW_CHUNKSIZE + 17   # rows over several chunks
    cols = Any[collect(1:n), collect(1:n) ./ 7, ["s$(i % 40)" for i in 1:n], [isodd(i) for i in 1:n]]
    names = ["a", "b", "c", "d"]
    cached = stream_parts(stream_written(io -> XLSX.writetable(io, cols, names)))
    streamed = stream_parts(stream_written(io -> XLSX.writetable(io, cols, names; streaming=true, shared_strings=true)))
    @test keys(cached) == keys(streamed)
    for (name, xml) in cached
        @test streamed[name] == xml
    end
    @test occursin("<dimension ref=\"A1:D$(n + 1)\"/>", streamed["xl/worksheets/sheet1.xml"])

    # An empty cell is left out, as Excel leaves it out; the cached writer writes `<c r="…"/>`.
    sheet = stream_parts(stream_written(io -> XLSX.writetable(io, Any[[1, 2], [missing, "x"]], ["a", "b"]; streaming=true)))["xl/worksheets/sheet1.xml"]
    @test !occursin("r=\"B2\"", sheet)
    @test occursin("r=\"B3\"", sheet)
end

@testset "streamxlsx" begin
    f = "streamxlsx.xlsx"
    @test XLSX.streamxlsx(f; overwrite=true) do s
        @test s isa XLSX.StreamWriter
        sh = XLSX.addsheet!(s, "Data"; columns=["id", "name", "when", "amount"], freeze_header=true,
                            column_widths=[8, nothing, 14], column_formats=Dict("amount" => "0.00", 1 => "0"),
                            as_table=true, table_name="Orders", table_style="TableStyleMedium2")
        @test sh isa XLSX.StreamSheet
        for i in 1:2500
            XLSX.writerow!(sh, (i, "item $i", Date(2026, 1, 1) + Day(i), i / 4))
        end
        @test sprint(show, sh) == "StreamSheet(\"Data\", 2501 rows written)"
        more = XLSX.addsheet!(s)   # closes "Data"
        @test !sh.isopen
        @test_throws XLSX.XLSXError XLSX.writerow!(sh, (1,))
        XLSX.writerow!(more, [1, missing, "x"])
        XLSX.writerow!(more, [missing, nothing, ""])   # no cells: the row is left out
        XLSX.writerow!(more, (a=2, b=3))
        XLSX.writerow!(more, Any[])
        XLSX.writerow!(more, ["last"])
    end == abspath(f)
    SAVE_FILES && save_outfile(f)

    xf = XLSX.openxlsx(f; mode="rw")   # `getColumnWidth` needs a writable file
    @test XLSX.sheetnames(xf) == ["Data", "Sheet1"]   # the first free name, as `addsheet!` chooses
    data = xf["Data"]
    @test data["A1:D1"] == Any["id" "name" "when" "amount"]
    @test data["A2501"] == 2500
    @test data["B3"] == "item 2"
    @test data["C2"] == Date(2026, 1, 2)
    @test data["D3"] == 0.5
    @test XLSX.getFormat(data, "D2").format["numFmt"]["formatCode"] == "0.00"
    @test XLSX.getColumnWidth(data, "A1") ≈ 8 atol=1
    @test isnothing(XLSX.getColumnWidth(data, "B1"))
    t = only(XLSX.tables(data))
    @test t.name == "Orders"
    @test string(t.ref) == "A1:D2501"
    @test isequal(xf["Sheet1"][:], Any[1 missing "x"; missing missing missing; 2 3 missing; missing missing missing; "last" missing missing])

    parts = stream_parts(read(f))
    sheet1 = parts["xl/worksheets/sheet1.xml"]
    @test occursin("state=\"frozen\"", sheet1)
    @test occursin("<tableParts count=\"1\">", sheet1)
    @test !occursin("<dimension", parts["xl/worksheets/sheet2.xml"])   # rows not known in advance
    @test !occursin("r=\"4\"", parts["xl/worksheets/sheet2.xml"])      # the empty rows are left out
    isfile(f) && rm(f)

    # nothing written: an ordinary empty workbook
    empty = stream_written(io -> XLSX.streamxlsx(_ -> nothing, io))
    @test XLSX.sheetnames(XLSX.readxlsx(IOBuffer(empty))) == ["Sheet1"]

    # three sheets, the first keeping its default name
    three = stream_written(io -> XLSX.streamxlsx(io) do s
        for k in 1:3
            XLSX.writerow!(XLSX.addsheet!(s), (k,))
        end
    end)
    xf = XLSX.readxlsx(IOBuffer(three))
    @test XLSX.sheetnames(xf) == ["Sheet1", "Sheet2", "Sheet3"]
    @test [xf[k]["A1"] for k in 1:3] == [1, 2, 3]
end

@testset "row tables" begin
    rows = [(id=i, name="n$i", when=Date(2026, 1, i)) for i in 1:20]
    # a vector of NamedTuples is read a row at a time, its length known
    streamed = stream_written(io -> XLSX.writetable(io, rows; streaming=true, sheetname="Rows", as_table=true, table_name="RowTable"))
    t = XLSX.readtable(IOBuffer(streamed), "Rows")
    @test t.column_labels == [:id, :name, :when]
    @test t.data == Any[1:20, ["n$i" for i in 1:20], [Date(2026, 1, i) for i in 1:20]]
    @test occursin("<dimension ref=\"A1:C21\"/>", stream_parts(streamed)["xl/worksheets/sheet1.xml"])

    # a filtered iterator: its length isn't known, so there is no dimension
    filtered = Iterators.filter(r -> iseven(r.id), rows)
    streamed = stream_written(io -> XLSX.writetable(io, filtered; streaming=true))
    @test XLSX.readtable(IOBuffer(streamed), 1).data[1] == collect(2:2:20)
    @test !occursin("<dimension", stream_parts(streamed)["xl/worksheets/sheet1.xml"])

    # no rows: just the header
    none = Iterators.filter(_ -> false, rows)
    xf = XLSX.readxlsx(IOBuffer(stream_written(io -> XLSX.writetable(io, none; streaming=true))))
    @test xf[1]["A1:C1"] == Any["id" "name" "when"]
end

@testset "several sheets" begin
    df1 = DataFrames.DataFrame(a=[1, 2], b=["x", "y"])
    rows2 = [(c=1.5,), (c=2.5,)]
    for (label, write) in [
        "pairs"   => io -> XLSX.writetable(io, "First" => df1, :Second => rows2; streaming=true, as_table=true),
        "kwargs"  => io -> XLSX.writetable(io; First=(Any[[1, 2], ["x", "y"]], ["a", "b"]), Second=(Any[[1.5, 2.5]], ["c"]), streaming=true, as_table=true),
        "tuples"  => io -> XLSX.writetable(io, [("First", Any[[1, 2], ["x", "y"]], ["a", "b"]), ("Second", Any[[1.5, 2.5]], ["c"])]; streaming=true, as_table=true),
    ]
        @testset "$label" begin
            xf = XLSX.readxlsx(IOBuffer(stream_written(write)))
            @test XLSX.sheetnames(xf) == ["First", "Second"]
            @test XLSX.gettable(xf["First"]).data == Any[[1, 2], ["x", "y"]]
            @test xf["Second"]["A3"] == 2.5
            @test [only(XLSX.tables(xf[n])).name for n in 1:2] == ["First", "Second"]
        end
    end
end

@testset "streaming errors" begin
    f = "stream_errors.xlsx"
    touch(f)
    @test_throws XLSX.XLSXError XLSX.streamxlsx(_ -> nothing, f)
    @test_throws XLSX.XLSXError XLSX.writetable(f, Any[[1]], ["a"]; streaming=true)
    rm(f)

    # an error in the do block deletes the half-written file
    @test_throws ErrorException XLSX.streamxlsx(f) do s
        sh = XLSX.addsheet!(s, "A")
        for i in 1:(XLSX.ROW_CHUNKSIZE + 5)
            XLSX.writerow!(sh, (i,))
        end
        error("stop")
    end
    @test !isfile(f)

    stream_err(f) = @test_throws XLSX.XLSXError stream_written(io -> XLSX.streamxlsx(f, io))
    stream_err(s -> XLSX.writerow!(XLSX.addsheet!(s; columns=["a"]), (1, 2)))            # longer than the header
    stream_err(s -> XLSX.writerow!(XLSX.addsheet!(s), fill(1, XLSX.EXCEL_MAX_COLS + 1)))  # too many columns
    stream_err(s -> XLSX.writerow!(XLSX.addsheet!(s), 1))                                 # not a row
    stream_err(s -> XLSX.addsheet!(s; as_table=true))                                     # a table needs columns
    stream_err(s -> XLSX.addsheet!(s; columns=["a", "A"], as_table=true))                 # duplicate names
    stream_err(s -> XLSX.addsheet!(s; columns=["a"], as_table=true))                      # header only
    stream_err(s -> XLSX.addsheet!(s; column_widths=10))                                  # one width needs columns
    stream_err(s -> XLSX.addsheet!(s; columns=["a"], column_formats=Dict("b" => "0.0")))  # no column "b"
    stream_err(s -> XLSX.addsheet!(s; column_formats=["0.0"]))                            # not a Dict
    stream_err(function (s)
        sh = XLSX.addsheet!(s)
        sh.row = XLSX.EXCEL_MAX_ROWS
        XLSX.writerow!(sh, (1,))                                                   # beyond the last row
    end)
    @test_throws XLSX.XLSXError XLSX.writetable(IOBuffer(), Any[[1]], ["a"]; streaming=true, anchor_cell="B2")
    @test_throws XLSX.XLSXError XLSX.writetable(IOBuffer(), Any[[1]], ["a"]; streaming=true, as_table=true, totals=["a" => :sum])
end

@testset "writetable layout keywords" begin
    # the same keywords without streaming
    for streaming in (false, true)
        bytes = stream_written(io -> XLSX.writetable(io, Any[[1, 2], [0.5, 0.25]], ["a", "b"]; streaming,
                                                     column_widths=20, column_formats=Dict(:b => "0.0%"), freeze_header=true))
        xf = XLSX.openxlsx(IOBuffer(bytes); mode="rw")
        @test XLSX.getColumnWidth(xf[1], "A1") ≈ 20 atol=1
        @test XLSX.getColumnWidth(xf[1], "B1") ≈ 20 atol=1
        @test XLSX.getFormat(xf[1], "B2").format["numFmt"]["formatCode"] == "0.0%"
        @test isnothing(XLSX.getFormat(xf[1], "B1"))   # not the header
        @test occursin("state=\"frozen\"", stream_parts(bytes)["xl/worksheets/sheet1.xml"])
    end
    xf = XLSX.newxlsx()
    @test_throws XLSX.XLSXError XLSX.writetable!(xf[1], Any[[1]], ["a"]; write_columnnames=false, freeze_header=true)
    XLSX.writetable!(xf[1], Any[[1, 2]], ["a"]; anchor_cell=XLSX.CellRef("B3"), freeze_header=true, column_widths=[15])
    @test XLSX.getColumnWidth(xf[1], "B3") ≈ 15 atol=1
end
