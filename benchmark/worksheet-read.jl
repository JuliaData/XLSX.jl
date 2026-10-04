# Run with `julia --project=. --check-bounds=yes benchmark/worksheet-read.jl` after Pkg.instantiate().
# The compressed numeric worksheet has 20,000 rows and eight columns. Fixture
# construction and value checks are outside the warmed, five-sample measurements.
using XLSX, Test

function synthetic_workbook(nrows = 20_000, ncols = 8)
    xml = IOBuffer()
    print(xml, "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><dimension ref=\"A1:", XLSX.encode_column_number(ncols), nrows, "\"/><sheetData>")
    for row in 1:nrows
        print(xml, "<row r=\"", row, "\">")
        for col in 1:ncols
            print(xml, "<c r=\"", XLSX.encode_column_number(col), row, "\"><v>", row * 10 + col, "</v></c>")
        end
        print(xml, "</row>")
    end
    print(xml, "</sheetData><pageMargins left=\"0.7\" right=\"0.7\" top=\"0.75\" bottom=\"0.75\" header=\"0.3\" footer=\"0.3\"/></worksheet>")
    sheet = take!(xml)
    template = XLSX.ZipArchives.ZipReader(read(joinpath(dirname(pathof(XLSX)), "data", "blank.xlsx")))
    io = IOBuffer()
    XLSX.ZipArchives.ZipWriter(io) do writer
        for name in XLSX.ZipArchives.zip_names(template)
            XLSX.ZipArchives.zip_newfile(writer, name; compress = true)
            write(writer, name == "xl/worksheets/sheet1.xml" ? sheet : XLSX.ZipArchives.zip_readentry(template, name))
        end
    end
    return take!(io), length(sheet)
end

function samples(f)
    f()
    f()
    allocated = Int[]
    times = Float64[]
    for _ in 1:5
        GC.gc()
        result = @timed f()
        push!(allocated, result.bytes)
        push!(times, result.time)
    end
    return (median_bytes = sort(allocated)[3], median_seconds = sort(times)[3], bytes = allocated, seconds = times)
end

bytes, sheetbytes = synthetic_workbook()
println("ACTUAL_JULIA=", VERSION, " WORKER_THREADS=", Threads.nthreads(),
        " WORKSHEET_BYTES=", sheetbytes, " ZIP_BYTES=", length(bytes))
init() = XLSX.readxlsx(IOBuffer(bytes))
table() = XLSX.readtable(IOBuffer(bytes), 1; header = false)
@testset "Synthetic public worksheet controls" begin
    xf = init()
    @test xf[1]["A1"] == 11
    @test xf[1]["H20000"] == 200008
    @test size(xf[1][:]) == (20_000, 8)
    dt = table()
    @test length(dt.data) == 8
    @test dt.data[1][1] == 11
    @test dt.data[8][end] == 200008
end
println("READXLSX=", samples(init))
println("READTABLE=", samples(table))
