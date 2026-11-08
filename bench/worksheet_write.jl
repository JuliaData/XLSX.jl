# Run from bench/ with `julia --project=envs/dev worksheet_write.jl [name]` after
# instantiating envs/dev. Fixture-free, like worksheet_read.jl: the data is built in
# memory, and each write path is timed on its own (median of five warmed `@timed`
# samples), so a change to one shows clearly. With `name`, the results are also
# appended to results/write/<name>.txt, and earlier runs there are printed
# alongside for comparison.
#
# The paths timed:
#   writetable_numeric  `writetable` of 100,000 rows x 8 Float64 columns, to an IO
#   writetable_strings  the same shape of strings, half unique, half repeated
#   writetable_mixed    Int, Float64, String, Date, DateTime, Bool and missing columns
#   setindex_cells      a 20,000 x 8 block written cell by cell into a new workbook
#   setindex_block      the same block written with one `sheet[range] = matrix`
#   save_unchanged      a workbook of 100,000 x 8 numbers opened "rw" and saved unchanged
#   save_after_edit     the same, after changing one cell
#   format_cells        `setFont` and `setFill` on 2,000 x 8 cells, then save

using XLSX, Dates, Test

const NROWS = 100_000
const NCOLS = 8

numeric_columns(n=NROWS) = [collect(range(c, step=0.5, length=n)) for c in 1:NCOLS]
string_columns(n=NROWS) = [[isodd(c) ? "row $r col $c" : "label $(r % 50)" for r in 1:n] for c in 1:NCOLS]
function mixed_columns(n=NROWS)
    d0 = Date(2020, 1, 1)
    return Any[collect(1:n), collect(range(0.1, step=0.1, length=n)), ["name $(r % 1000)" for r in 1:n],
               [d0 + Day(r % 3000) for r in 1:n], [DateTime(d0) + Minute(r) for r in 1:n],
               [isodd(r) for r in 1:n], [r % 10 == 0 ? missing : r for r in 1:n], ["x" for _ in 1:n]]
end
names_for(cols) = ["col$i" for i in eachindex(cols)]

function write_table(cols)
    io = IOBuffer()
    XLSX.writetable(io, cols, names_for(cols))
    return take!(io)
end

function setindex_cells(n=20_000)
    xf = XLSX.newxlsx()
    s = xf[1]
    for r in 1:n, c in 1:NCOLS
        s[r, c] = r * 10 + c
    end
    io = IOBuffer()
    XLSX.writexlsx(io, xf)
    return take!(io)
end

function setindex_block(n=20_000)
    xf = XLSX.newxlsx()
    xf[1][1:n, 1:NCOLS] = [r * 10 + c for r in 1:n, c in 1:NCOLS]
    io = IOBuffer()
    XLSX.writexlsx(io, xf)
    return take!(io)
end

const SAVED = write_table(numeric_columns())

function save_unchanged()
    xf = XLSX.openxlsx(IOBuffer(SAVED); mode="rw")
    io = IOBuffer()
    XLSX.writexlsx(io, xf)
    return take!(io)
end

function save_after_edit()
    xf = XLSX.openxlsx(IOBuffer(SAVED); mode="rw")
    xf[1]["B2"] = "edited"
    io = IOBuffer()
    XLSX.writexlsx(io, xf)
    return take!(io)
end

function format_cells(n=2_000)
    xf = XLSX.newxlsx()
    s = xf[1]
    s[1:n, 1:NCOLS] = [r * 10 + c for r in 1:n, c in 1:NCOLS]
    XLSX.setFont(s, "A1:H$n"; bold=true, color="FF0000")
    XLSX.setFill(s, "A1:H$n"; pattern="solid", fgColor="FFFF00")
    io = IOBuffer()
    XLSX.writexlsx(io, xf)
    return take!(io)
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
    return (seconds = sort(times)[3], mib = round(sort(allocated)[3] / 2^20; digits=1))
end

const NUMERIC = numeric_columns()
const STRINGS = string_columns()
const MIXED = mixed_columns()

@testset "Write controls" begin
    r = XLSX.readxlsx(IOBuffer(write_table(MIXED)))[1]
    @test r["A2"] == 1 && r["C3"] == "name 2" && r["D2"] == Date(2020, 1, 2) && ismissing(r["G11"])
    @test XLSX.readxlsx(IOBuffer(setindex_cells(10)))[1]["H10"] == 108
    @test XLSX.readxlsx(IOBuffer(save_after_edit()))[1]["B2"] == "edited"
    @test XLSX.readxlsx(IOBuffer(save_unchanged()))[1][NROWS + 1, NCOLS] == NUMERIC[NCOLS][end]
end

const PATHS = [
    "writetable_numeric" => () -> write_table(NUMERIC),
    "writetable_strings" => () -> write_table(STRINGS),
    "writetable_mixed"   => () -> write_table(MIXED),
    "setindex_cells"     => setindex_cells,
    "setindex_block"     => setindex_block,
    "save_unchanged"     => save_unchanged,
    "save_after_edit"    => save_after_edit,
    "format_cells"       => format_cells,
]

println("JULIA=", VERSION, " THREADS=", Threads.nthreads())
lines = String[]
for (name, f) in PATHS
    s = samples(f)
    line = rpad(name, 22) * lpad(string(round(s.seconds; digits=3)), 9) * " s" * lpad(string(s.mib), 10) * " MiB"
    println(line)
    push!(lines, line)
end

if !isempty(ARGS)
    dir = joinpath(@__DIR__, "results", "write")
    mkpath(dir)
    for f in sort(readdir(dir))
        endswith(f, ".txt") && f != "$(ARGS[1]).txt" || continue
        println("\n-- earlier run: ", f[1:end-4])
        print(read(joinpath(dir, f), String))
    end
    write(joinpath(dir, "$(ARGS[1]).txt"), join(lines, "\n") * "\n")
end
