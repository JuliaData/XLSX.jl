# bench_worker.jl
# Called as:
#   julia --project=envs/<ver> bench_worker.jl <version_label> <fixtures_dir> <output_json>

using BenchmarkTools
using XLSX
using Dates

label        = ARGS[1]
fixtures_dir = ARGS[2]
output_path  = ARGS[3]

# Optional filters, used by quick.jl:
#   BENCH_FIXTURES   comma-separated fixture names (default: all present)
#   BENCH_BENCHMARKS comma-separated benchmark names (default: all)
#   BENCH_SECONDS    time budget per benchmark (default: 30)
_env_list(k) = (v = get(ENV, k, ""); isempty(v) ? nothing : Set(split(v, ',')))
const ONLY_FIXTURES   = _env_list("BENCH_FIXTURES")
const ONLY_BENCHMARKS = _env_list("BENCH_BENCHMARKS")
const SECONDS         = parse(Float64, get(ENV, "BENCH_SECONDS", "30"))
wanted(b) = isnothing(ONLY_BENCHMARKS) || b in ONLY_BENCHMARKS

BenchmarkTools.DEFAULT_PARAMETERS.seconds = SECONDS
BenchmarkTools.DEFAULT_PARAMETERS.samples = 10
BenchmarkTools.DEFAULT_PARAMETERS.evals   = 1   # file I/O: 1 eval per sample

suite = BenchmarkGroup()

println("Version: $label  |  Threads: $(Threads.nthreads()) | XLSX: $(pkgversion(XLSX))")

# ── Helpers ───────────────────────────────────────────────────────────────────

# The table's sheet per fixture. The Excel-like xl_* fixtures (generate_xl_fixtures.jl)
# hold their table on a second sheet, "Data", with headers on row 5 in A:CF.
table_sheet(fixture) = startswith(fixture, "xl_") ? "Data" : "Sheet1"

function bench_readtable(path, sheet)
    XLSX.readtable(path, sheet)
end

# The #462 call: a column subset with headers below title rows.
function bench_readtable_subset(path)
    XLSX.readtable(path, "Data", "A:CF"; first_row=5)
end

# Reading through gap rows (xl_gaps): absent rows come back as all-missing rows.
function bench_readtable_gaps(path)
    XLSX.readtable(path, "Data", "A:CF"; first_row=5, stop_in_empty_row=false, keep_empty_rows=true)
end

function bench_readxlsx(path)
    XLSX.readxlsx(path)
end

function bench_write(source_path, tmp_path, sheet)
    XLSX.writetable(tmp_path, XLSX.readtable(source_path, sheet); overwrite=true)
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
    "xl_medium", "xl_issue462", "xl_gaps",
]
    isnothing(ONLY_FIXTURES) || fixture_name in ONLY_FIXTURES || continue
    ext  = startswith(fixture_name, "xl_") ? ".xlsm" : ".xlsx"
    path = joinpath(fixtures_dir, "$(fixture_name)$(ext)")
    isfile(path) || continue
    # Skip a fixture this XLSX version can't open (v0.11 rejects the macro-enabled
    # xl_* files), rather than failing the whole run. report.jl shows it as N/A.
    try
        XLSX.openxlsx(_ -> nothing, path; enable_cache=false)
    catch e
        println("Skipping $fixture_name: XLSX v$(pkgversion(XLSX)) can't open it ($(sprint(showerror, e)))")
        continue
    end
    sheet = table_sheet(fixture_name)

    suite[fixture_name] = BenchmarkGroup()

    # Open cost only — no data access
    wanted("open") && (suite[fixture_name]["open"] = @benchmarkable(
        XLSX.openxlsx($path) do xf; nothing; end,
        seconds=SECONDS, evals=1
    ))

    # readtable — full user-facing single-sheet read
    wanted("readtable") && (suite[fixture_name]["readtable"] =
        @benchmarkable bench_readtable($path, $sheet) seconds=SECONDS evals=1)

    # readtable_subset — the #462 call (column subset, headers on row 5)
    if startswith(fixture_name, "xl_") && wanted("readtable_subset")
        suite[fixture_name]["readtable_subset"] =
            @benchmarkable bench_readtable_subset($path) seconds=SECONDS evals=1
    end

    # readtable_gaps — reading through gap rows (xl_gaps only)
    if fixture_name == "xl_gaps" && wanted("readtable_gaps")
        suite[fixture_name]["readtable_gaps"] =
            @benchmarkable bench_readtable_gaps($path) seconds=SECONDS evals=1
    end

    # readxlsx — open + parse, no iteration
    wanted("readxlsx") && (suite[fixture_name]["readxlsx"] =
        @benchmarkable bench_readxlsx($path) seconds=SECONDS evals=1)

    # eachrow — iteration only, cache pre-warmed in setup (excluded from timing)
    wanted("eachrow") && (suite[fixture_name]["eachrow"] = @benchmarkable(
        begin
            for row in XLSX.eachrow(_sh)
                for col in _col_start:_col_stop
                    _ = XLSX.getdata(row, col)
                end
            end
        end,
        setup=(
            _xf = XLSX.openxlsx($path);
            _sh = _xf[$sheet];
            warm_cache!(_sh);
            _dim = XLSX.get_dimension(_sh);
            _col_start = XLSX.column_number(_dim.start);
            _col_stop = XLSX.column_number(_dim.stop)
        ),
        seconds=SECONDS, evals=1
    ))

    # single_cell — random access, cache pre-warmed in setup (excluded from timing)
    wanted("single_cell") && (suite[fixture_name]["single_cell"] = @benchmarkable(
        begin
            XLSX.getdata(_sh, _dim.start)
            XLSX.getdata(_sh, _dim.stop)
        end,
        setup=(
            _xf = XLSX.openxlsx($path);
            _sh = _xf[$sheet];
            warm_cache!(_sh);
            _dim = XLSX.get_dimension(_sh)
        ),
        seconds=SECONDS, evals=1
    ))

    # writetable — read the table sheet and write to temp file
    tmp = tempname() * ".xlsx"
    wanted("writetable") && (suite[fixture_name]["writetable"] =
        @benchmarkable bench_write($path, $tmp, $sheet) seconds=2SECONDS evals=1)

    # open_readwrite — open in rw mode (eager parallel fill)
    wanted("open_readwrite") && (suite[fixture_name]["open_readwrite"] = @benchmarkable(
        begin
            tmp_rw = tempname() * $ext
            cp($path, tmp_rw)
            XLSX.openxlsx(tmp_rw, mode="rw") do xf; nothing; end
            rm(tmp_rw; force=true)
        end,
        seconds=SECONDS, evals=1
    ))

    if fixture_name == "multi_sheet" && wanted("readtable_all_sheets")
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
            seconds=SECONDS, evals=1
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