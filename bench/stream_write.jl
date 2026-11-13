# Run from bench/ with `julia --project=envs/dev -t auto stream_write.jl [name]` after
# instantiating envs/dev. Fixture-free: the data is built in memory.
#
# Compares `writetable` through the cell cache with `writetable(...; streaming=true)`,
# by peak memory (RSS) as well as time and allocations. Peak RSS only ever rises
# within a process, so each case runs in a fresh Julia process of its own, which
# writes once to compile and then once to measure. `data only` builds the data and
# writes nothing: the memory the data itself takes, to subtract from the others.
#
# The cases:
#   mixed 1M x 20     4 Int, 4 Float64, 3 Date, 1 Bool, 4 unique-string and
#                     4 repeated-string (50 values) columns, 1,000,000 rows
#   strings 100k x 8  the `writetable_strings` shape of worksheet_write.jl
#
# Each is written by `cached` (`writetable`), `streamed` (`streaming=true`, inline
# strings) and `streamed shared` (`streaming=true, shared_strings=true`).
# With `name`, the results are also appended to results/stream/<name>.txt, and
# earlier runs there are printed alongside.

const CASES = ["mixed 1M x 20", "strings 100k x 8"]
const MODES = ["data only", "cached", "streamed", "streamed shared"]

function mixed_columns(n)
    d0 = Date(2020, 1, 1)
    pool = ["category $i" for i in 1:50]
    return Any[
        (collect(r * c for r in 1:n) for c in 1:4)...,
        (collect(r * 0.25 + c for r in 1:n) for c in 1:4)...,
        ([d0 + Day((r + c) % 3000) for r in 1:n] for c in 1:3)...,
        [isodd(r) for r in 1:n],
        (["row $r text $c" for r in 1:n] for c in 1:4)...,
        ([pool[(7r + c) % 50 + 1] for r in 1:n] for c in 1:4)...,
    ]
end
string_columns(n) = [[isodd(c) ? "row $r col $c" : "label $(r % 50)" for r in 1:n] for c in 1:8]

case_columns(case) = case == "mixed 1M x 20" ? mixed_columns(1_000_000) : string_columns(100_000)

# One case and mode, in this process: prints `seconds MiB-allocated MiB-file MiB-peak-RSS`.
function run_one(case, mode)
    @eval using XLSX, Dates
    cols = Base.invokelatest(case_columns, case)
    names = ["c$i" for i in eachindex(cols)]
    path = tempname() * ".xlsx"
    write_it() = mode == "data only" ? nothing :
        XLSX.writetable(path, cols, names; overwrite=true,
                        (mode == "cached" ? (;) : (; streaming=true, shared_strings=(mode == "streamed shared")))...)
    Base.invokelatest(write_it)   # compile
    GC.gc()
    r = @timed Base.invokelatest(write_it)
    size = isfile(path) ? filesize(path) : 0
    isfile(path) && rm(path)
    println(round(r.time; digits=2), " ", round(r.bytes / 2^20; digits=1), " ",
            round(size / 2^20; digits=1), " ", round(Sys.maxrss() / 2^20; digits=1))
end

function main()
    if length(ARGS) >= 3 && ARGS[1] == "--one"
        return run_one(ARGS[2], ARGS[3])
    end
    name = isempty(ARGS) ? nothing : ARGS[1]
    lines = String[]
    header = rpad("case", 18) * rpad("mode", 17) * lpad("seconds", 9) * lpad("alloc MiB", 11) * lpad("file MiB", 10) * lpad("peak RSS MiB", 14)
    println(header)
    for case in CASES, mode in MODES
        cmd = `$(Base.julia_cmd()) --project=$(Base.active_project()) -t $(Threads.nthreads()) $(@__FILE__) --one $case $mode`
        t, alloc, size, rss = split(readchomp(cmd))
        line = rpad(case, 18) * rpad(mode, 17) * lpad(t, 9) * lpad(alloc, 11) * lpad(size, 10) * lpad(rss, 14)
        println(line)
        push!(lines, line)
    end
    isnothing(name) && return
    dir = joinpath(@__DIR__, "results", "stream")
    mkpath(dir)
    for f in filter(endswith(".txt"), readdir(dir))
        println("\n== earlier: ", f, "\n", read(joinpath(dir, f), String))
    end
    open(joinpath(dir, "$name.txt"), "w") do io
        println(io, "threads: ", Threads.nthreads())
        println(io, header)
        foreach(l -> println(io, l), lines)
    end
end

main()
