# external_read.jl
# Times Python's Excel readers (python-calamine, fastexcel, openpyxl) reading the same
# tables `readtable` reads in bench_worker.jl, so that report.jl can show where XLSX.jl's
# read speed stands. Optional, and not a test dependency: run it in its own environment,
# whose CondaPkg installs a private Python with the readers on first use.
#
#   julia --project=envs/python external_read.jl
#
# Results go to <results>/external.toml (BENCH_RESULTS as in run_benchmarks.jl).
# Each reader is timed in Python itself, after one warm-up read: the median of up to
# 10 reads within BENCH_SECONDS (default 30) per reader and fixture.

using PythonCall
import TOML

const ROOT        = @__DIR__
const FIXTURES    = joinpath(ROOT, "fixtures")
const RESULTS_DIR = get(ENV, "BENCH_RESULTS", joinpath(ROOT, "results"))
const SECONDS     = parse(Float64, get(ENV, "BENCH_SECONDS", "30"))
mkpath(RESULTS_DIR)

# The fixtures report.jl lists, with the sheet and header row `readtable` reads.
const fixtures = ["small", "medium", "large", "wide_few", "tall_few",
                  "sst_unique", "sst_repeated", "sst_mixed",
                  "numeric_only", "dates_heavy", "multi_sheet",
                  "xl_medium", "xl_issue462", "xl_gaps"]
fixture_path(f) = joinpath(FIXTURES, f * (startswith(f, "xl_") ? ".xlsm" : ".xlsx"))
table_sheet(f) = startswith(f, "xl_") ? "Data" : "Sheet1"
header_row(f) = startswith(f, "xl_") ? 5 : 1

pyexec("""
import time, statistics

def read_calamine(path, sheet, header):
    from python_calamine import CalamineWorkbook
    rows = CalamineWorkbook.from_path(path).get_sheet_by_name(sheet).to_python()
    return len(rows) - header

def read_fastexcel(path, sheet, header):
    import fastexcel
    return fastexcel.read_excel(path).load_sheet(sheet, header_row=header - 1).to_arrow().num_rows

def read_openpyxl(path, sheet, header):
    import openpyxl
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    rows = list(wb[sheet].iter_rows(values_only=True))
    wb.close()
    return len(rows) - header

READERS = {"calamine": read_calamine, "fastexcel": read_fastexcel, "openpyxl": read_openpyxl}

def bench(name, path, sheet, header, budget):
    fn = READERS[name]
    t = time.perf_counter(); rows = fn(path, sheet, header); first = time.perf_counter() - t
    if first > budget:          # too slow to repeat: the one read is the result
        return first, rows
    times, start = [], time.perf_counter()
    while len(times) < 10 and (len(times) < 3 or time.perf_counter() - start < budget):
        t = time.perf_counter(); fn(path, sheet, header); times.append(time.perf_counter() - t)
    return statistics.median(times), rows
""", Main)

const bench = pyeval("bench", Main)
const version = pyimport("importlib.metadata").version
results = Dict{String,Any}(
    "versions" => Dict(r => pyconvert(String, version(p)) for (r, p) in
                       (("calamine", "python-calamine"), ("fastexcel", "fastexcel"), ("openpyxl", "openpyxl"))),
    "python" => pyconvert(String, pyimport("platform").python_version()),
)
for f in fixtures
    path = fixture_path(f)
    isfile(path) || (println("Skipping $f: no fixture (run generate_fixtures.jl)"); continue)
    results[f] = Dict{String,Float64}()
    for reader in ("calamine", "fastexcel", "openpyxl")
        try
            secs, rows = pyconvert(Tuple{Float64,Int}, bench(reader, path, table_sheet(f), header_row(f), SECONDS))
            rows >= 100 || error("read only $rows rows")
            results[f][reader] = secs * 1000
            println(rpad("$f/$reader", 30), round(secs * 1000; digits=1), " ms  ($rows rows)")
        catch e
            println(rpad("$f/$reader", 30), "failed: ", sprint(showerror, e)[1:min(end, 120)])
        end
    end
end
open(io -> TOML.print(io, results), joinpath(RESULTS_DIR, "external.toml"), "w")
println("Wrote ", joinpath(RESULTS_DIR, "external.toml"))
