# external_write.jl
# Times Python's Excel writers (openpyxl, XlsxWriter, rustpy-xlsxwriter) writing the same
# tables bench_worker.jl's `writetable_only` writes, so that report.jl can show where
# XLSX.jl's write speed stands. Optional, and not a test dependency: run it in the same
# environment as external_read.jl, whose CondaPkg installs a private Python on first use.
#
#   julia --project=envs/python external_write.jl
#
# Results go to <results>/external_write.toml (BENCH_RESULTS as in run_benchmarks.jl).
# Each fixture's table is read once with python-calamine, outside the timing, into a list
# of rows (empty cells as None) holding the rows `readtable` reads: from the header row to
# the first empty row (to the end on xl_gaps). Each writer then writes it to a new file,
# timed in Python itself after one warm-up write: the median of up to 10 writes within
# BENCH_SECONDS (default 30) per writer and fixture.

using PythonCall
import TOML

const ROOT        = @__DIR__
const FIXTURES    = joinpath(ROOT, "fixtures")
const RESULTS_DIR = get(ENV, "BENCH_RESULTS", joinpath(ROOT, "results"))
const SECONDS     = parse(Float64, get(ENV, "BENCH_SECONDS", "30"))
mkpath(RESULTS_DIR)

const fixtures = ["small", "medium", "large", "wide_few", "tall_few",
                  "sst_unique", "sst_repeated", "sst_mixed",
                  "numeric_only", "dates_heavy", "multi_sheet",
                  "xl_medium", "xl_issue462", "xl_gaps"]
# BENCH_FIXTURES: comma-separated fixture names (default: all)
let only = get(ENV, "BENCH_FIXTURES", "")
    isempty(only) || filter!(in(split(only, ',')), fixtures)
end
fixture_path(f) = joinpath(FIXTURES, f * (startswith(f, "xl_") ? ".xlsm" : ".xlsx"))
table_sheet(f) = startswith(f, "xl_") ? "Data" : "Sheet1"
header_row(f) = startswith(f, "xl_") ? 5 : 1

const WRITERS = ("openpyxl", "openpyxl_write_only", "xlsxwriter", "xlsxwriter_constant_memory", "rustpy_xlsxwriter")

pyexec("""
import time, statistics, os

def load_table(path, sheet, header, keep_gaps):
    from python_calamine import CalamineWorkbook
    rows = CalamineWorkbook.from_path(path).get_sheet_by_name(sheet).to_python()[header - 1:]
    rows = [[None if v == "" else v for v in r] for r in rows]
    if not keep_gaps:
        for i in range(1, len(rows)):
            if all(v is None for v in rows[i]):
                rows = rows[:i]
                break
    names, seen = [], set()
    for j, v in enumerate(rows[0]):
        n = str(v) if v is not None else f"column_{j + 1}"
        while n in seen:
            n += "_"
        seen.add(n); names.append(n)
    return names, rows[1:]

def write_openpyxl(out, names, data, records):
    import openpyxl
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(names)
    for r in data:
        ws.append(r)
    wb.save(out)

def write_openpyxl_write_only(out, names, data, records):
    import openpyxl
    wb = openpyxl.Workbook(write_only=True)
    ws = wb.create_sheet()
    ws.append(names)
    for r in data:
        ws.append(r)
    wb.save(out)

def _xlsxwriter(out, names, data, options):
    import xlsxwriter
    wb = xlsxwriter.Workbook(out, dict(options, default_date_format="yyyy-mm-dd"))
    ws = wb.add_worksheet()
    ws.write_row(0, 0, names)
    for i, r in enumerate(data, 1):
        ws.write_row(i, 0, r)
    wb.close()

def write_xlsxwriter(out, names, data, records):
    _xlsxwriter(out, names, data, {})

def write_xlsxwriter_constant_memory(out, names, data, records):
    _xlsxwriter(out, names, data, {"constant_memory": True})

def write_rustpy_xlsxwriter(out, names, data, records):
    from rustpy_xlsxwriter import write_worksheet
    write_worksheet(records, out, sheet_name="Sheet1")

WRITERS = {"openpyxl": write_openpyxl, "openpyxl_write_only": write_openpyxl_write_only,
           "xlsxwriter": write_xlsxwriter, "xlsxwriter_constant_memory": write_xlsxwriter_constant_memory,
           "rustpy_xlsxwriter": write_rustpy_xlsxwriter}

def bench(name, out, names, data, records, budget):
    fn = WRITERS[name]
    t = time.perf_counter(); fn(out, names, data, records); first = time.perf_counter() - t
    size = os.path.getsize(out)
    if first > budget:          # too slow to repeat: the one write is the result
        return first, size
    times, start = [], time.perf_counter()
    while len(times) < 10 and (len(times) < 3 or time.perf_counter() - start < budget):
        t = time.perf_counter(); fn(out, names, data, records); times.append(time.perf_counter() - t)
    return statistics.median(times), size

def prepare(path, sheet, header, keep_gaps):
    names, data = load_table(path, sheet, header, keep_gaps)
    records = [dict(zip(names, r)) for r in data]   # rustpy-xlsxwriter takes a list of dicts
    return names, data, records, len(data), len(names)
""", Main)

const prepare = pyeval("prepare", Main)
const bench = pyeval("bench", Main)
const version = pyimport("importlib.metadata").version
results = Dict{String,Any}(
    "versions" => Dict(w => pyconvert(String, version(p)) for (w, p) in
                       (("openpyxl", "openpyxl"), ("xlsxwriter", "xlsxwriter"), ("rustpy_xlsxwriter", "rustpy-xlsxwriter"))),
    "python" => pyconvert(String, pyimport("platform").python_version()),
)
out = tempname() * ".xlsx"
for f in fixtures
    path = fixture_path(f)
    isfile(path) || (println("Skipping $f: no fixture (run generate_fixtures.jl)"); continue)
    names, data, records, nrows, ncols = prepare(path, table_sheet(f), header_row(f), f == "xl_gaps")
    nrows = pyconvert(Int, nrows)
    println("$f: table of $nrows rows × $(pyconvert(Int, ncols)) columns")
    nrows >= 100 || error("$f: read only $nrows rows")
    results[f] = Dict{String,Float64}()
    for w in WRITERS
        try
            secs, size = pyconvert(Tuple{Float64,Int}, bench(w, out, names, data, records, SECONDS))
            results[f][w] = secs * 1000
            println(rpad("$f/$w", 44), round(secs * 1000; digits=1), " ms  ($(round(size / 2^20; digits=2)) MiB)")
        catch e
            println(rpad("$f/$w", 44), "failed: ", sprint(showerror, e)[1:min(end, 160)])
        end
    end
    names = data = records = nothing
    PythonCall.GC.gc()
end
rm(out; force=true)
open(io -> TOML.print(io, results), joinpath(RESULTS_DIR, "external_write.toml"), "w")
println("Wrote ", joinpath(RESULTS_DIR, "external_write.toml"))
