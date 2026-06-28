# bench_worker.jl
# Called as:
#   julia --project=envs/<ver> bench_worker.jl <version_label> <fixtures_dir> <output_json>

using BenchmarkTools
using XLSX
using Dates

label        = ARGS[1]
fixtures_dir = ARGS[2]
output_path  = ARGS[3]

BenchmarkTools.DEFAULT_PARAMETERS.seconds = 30
BenchmarkTools.DEFAULT_PARAMETERS.samples = 10
BenchmarkTools.DEFAULT_PARAMETERS.evals   = 1   # file I/O: 1 eval per sample

suite = BenchmarkGroup()

println("Version: $label  |  Threads: $(Threads.nthreads()) | XLSX: $(pkgversion(XLSX))")

# ── Helpers ───────────────────────────────────────────────────────────────────

function bench_readtable(path)
    XLSX.readtable(path, "Sheet1")
end

function bench_readxlsx(path)
    XLSX.readxlsx(path)
end

function bench_write(source_path, tmp_path)
    XLSX.writetable(tmp_path, XLSX.readtable(source_path, "Sheet1"); overwrite=true)
    nothing
end

function warm_cache!(sh)
    for row in XLSX.eachrow(sh)
        nothing
    end
    return nothing
end

# ── Build suite ───────────────────────────────────────────────────────────────

for fixture_name in [
    "small", "medium", "large", "wide_few", "tall_few",
    "sst_unique", "sst_repeated", "sst_mixed",
    "numeric_only", "dates_heavy", "multi_sheet",
]
    path = joinpath(fixtures_dir, "$(fixture_name).xlsx")
    isfile(path) || continue

    rw_path = joinpath(fixtures_dir, "$(fixture_name)_rw.xlsx")
    isfile(rw_path) || cp(path, rw_path)

    suite[fixture_name] = BenchmarkGroup()

    # Open cost only — no data access
    suite[fixture_name]["open"] = @benchmarkable(
        XLSX.openxlsx($path) do xf; nothing; end,
        seconds=30, evals=1
    )

    # readtable — full user-facing single-sheet read
    suite[fixture_name]["readtable"] = @benchmarkable bench_readtable($path) seconds=30 evals=1

    # readxlsx — open + parse, no iteration
    suite[fixture_name]["readxlsx"] = @benchmarkable bench_readxlsx($path) seconds=30 evals=1

    # eachrow — iteration only, cache pre-warmed in setup (excluded from timing)
    suite[fixture_name]["eachrow"] = @benchmarkable(
        begin
            for row in XLSX.eachrow(_sh)
                for col in _col_start:_col_stop
                    _ = XLSX.getdata(row, col)
                end
            end
        end,
        setup=(
            _xf = XLSX.openxlsx($path);
            _sh = _xf["Sheet1"];
            warm_cache!(_sh);
            _dim = XLSX.get_dimension(_sh);
            _col_start = XLSX.column_number(_dim.start);
            _col_stop = XLSX.column_number(_dim.stop)
        ),
        seconds=30, evals=1
    )

    # single_cell — random access, cache pre-warmed in setup (excluded from timing)
    suite[fixture_name]["single_cell"] = @benchmarkable(
        begin
            XLSX.getdata(_sh, _dim.start)
            XLSX.getdata(_sh, _dim.stop)
        end,
        setup=(
            _xf = XLSX.openxlsx($path);
            _sh = _xf["Sheet1"];
            warm_cache!(_sh);
            _dim = XLSX.get_dimension(_sh)
        ),
        seconds=30, evals=1
    )

    # writetable — read Sheet1 and write to temp file
    tmp = tempname() * ".xlsx"
    suite[fixture_name]["writetable"] = @benchmarkable bench_write($path, $tmp) seconds=60 evals=1

    # open_readwrite — open in rw mode (eager parallel fill)
    suite[fixture_name]["open_readwrite"] = @benchmarkable(
        begin
            tmp_rw = tempname() * ".xlsx"
            cp($path, tmp_rw)
            XLSX.openxlsx(tmp_rw, mode="rw") do xf; nothing; end
            rm(tmp_rw; force=true)
        end,
        seconds=30, evals=1
    )

    if fixture_name == "multi_sheet"
        suite[fixture_name]["readtable_all_sheets"] = @benchmarkable(
            XLSX.openxlsx($path) do xf
                for sheet_no in 1:5
                    sh = xf[sheet_no]
                    dim = XLSX.get_dimension(sh)
                    col_start = XLSX.column_number(dim.start)
                    col_stop  = XLSX.column_number(dim.stop)
                    for row in XLSX.eachrow(sh)
                        for col in col_start:col_stop
                            _ = XLSX.getdata(row, col)
                        end
                    end
                end
            end,
            seconds=30, evals=1
        )
    end
end

println("Warming up…")
warmup(suite)

println("Running benchmarks for version: $label")
results = run(suite; verbose=true)

println("Serialising results to $output_path")
BenchmarkTools.save(output_path, results)
println("Done.")