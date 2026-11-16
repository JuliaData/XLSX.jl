using BenchmarkTools
using Printf

const ROOT        = @__DIR__
const RESULTS_DIR = get(ENV, "BENCH_RESULTS", joinpath(ROOT, "results"))  # as in run_benchmarks.jl

VERSIONS = [
    ("v0.10", joinpath(ROOT, "envs", "v0_10")),
    ("v0.11", joinpath(ROOT, "envs", "v0_11")),
    ("v0.12", joinpath(ROOT, "envs", "v0_12")),
    ("v0.13", joinpath(ROOT, "envs", "v0_13")),
    ("master", joinpath(ROOT, "envs", "master")), # upstream master on GitHub
    ("dev",   joinpath(ROOT, "envs", "dev")),     # the working tree
]

const fixtures = [
    "small", "medium", "large", "wide_few", "tall_few",
    "sst_unique", "sst_repeated", "sst_mixed",
    "numeric_only", "dates_heavy", "multi_sheet",
    "xl_medium", "xl_issue462", "xl_gaps",
]

const benchmarks = [
    "open", "readtable", "readtable_subset", "readtable_gaps", "readxlsx", "eachrow",
    "single_cell", "writetable", "writetable_only", "writetable_streaming", "open_readwrite",
    "readtable_all_sheets",
]

# ── Load ──────────────────────────────────────────────────────────────────────

all_results = Dict{String, BenchmarkGroup}()
for (ver_label, _) in VERSIONS
    outfile = joinpath(RESULTS_DIR, "$(ver_label).json")
    if !isfile(outfile)
        println("Missing results for $ver_label — run run_benchmarks.jl first.")
        continue
    end
    all_results[ver_label] = BenchmarkTools.load(outfile)[1]
    println("Loaded $ver_label")
end

isempty(all_results) && error("No results found in $(RESULTS_DIR)")

# ── Table ─────────────────────────────────────────────────────────────────────

println("\n" * "="^60)
println("RESULTS COMPARISON")
println("="^60)

let
    # Only the versions, and the ratios, this results directory has
    ver_labels = filter(v -> haskey(all_results, v), first.(VERSIONS))
    # (column label, numerator, denominator)
    ratios = [("v0.11/v0.10", "v0.11", "v0.10"), ("v0.12/v0.10", "v0.12", "v0.10"),
              ("v0.13/v0.10", "v0.13", "v0.10"), ("v0.13/v0.12", "v0.13", "v0.12"),
              ("dev/v0.10", "dev", "v0.10"), ("dev/v0.12", "dev", "v0.12"), ("dev/v0.13", "dev", "v0.13"),
              ("dev/master", "dev", "master")]
    filter!(r -> haskey(all_results, r[2]) && haskey(all_results, r[3]), ratios)

    header = @sprintf("%-30s", "fixture / benchmark")
    for v in ver_labels
        header *= @sprintf("%15s", v)
    end
    for (label, _, _) in ratios
        header *= @sprintf("%15s", label)
    end
    println(header)
    println("-"^(30 + 15*(length(ver_labels) + length(ratios))))

    for fix in fixtures, bench in benchmarks
        medians = Dict{String,Float64}()
        for v in ver_labels
            haskey(all_results, v)             || continue
            haskey(all_results[v], fix)        || continue
            haskey(all_results[v][fix], bench) || continue
            medians[v] = median(all_results[v][fix][bench]).time / 1e6
        end
        isempty(medians) && continue
        row = @sprintf("%-30s", "$(fix)/$(bench)")
        for v in ver_labels
            row *= haskey(medians, v) ? @sprintf("%13.1fms", medians[v]) : @sprintf("%15s", "N/A")
        end
        for (_, num, den) in ratios
            t, base = get(medians, num, NaN), get(medians, den, NaN)
            row *= isnan(base) || isnan(t) || base == 0 ?
                @sprintf("%15s", "N/A") :
                @sprintf("%14.2fx", t / base)
        end
        println(row)
    end
    println("\nMedian times in milliseconds. Ratio < 1.0x = numerator version is faster.")
end

# ── Against Python's readers ──────────────────────────────────────────────────
# external_read.jl's times for python-calamine, fastexcel and openpyxl reading the same
# tables, beside the newest XLSX.jl version's `readtable`.

import TOML
let path = joinpath(RESULTS_DIR, "external.toml")
    isfile(path) || return
    ext = TOML.parsefile(path)
    v = something(findlast(l -> haskey(all_results, l), first.(VERSIONS)) |> i -> isnothing(i) ? nothing : first(VERSIONS[i]), "")
    isempty(v) && return
    readers = ("calamine", "fastexcel", "openpyxl")
    println("\n" * "="^60)
    println("READING A TABLE: XLSX.jl $v AGAINST PYTHON'S READERS")
    println("="^60)
    vers = get(ext, "versions", Dict())
    println("Python ", get(ext, "python", "?"), "; ", join(("$r $(get(vers, r, "?"))" for r in readers), ", "))
    println(@sprintf("%-16s%14s", "fixture", "XLSX $v") * join(@sprintf("%14s%10s", r, "XLSX/" * first(r, 4)) for r in readers))
    for fix in fixtures
        haskey(ext, fix) && haskey(all_results[v], fix) && haskey(all_results[v][fix], "readtable") || continue
        ours = median(all_results[v][fix]["readtable"]).time / 1e6
        row = @sprintf("%-16s%12.1fms", fix, ours)
        for r in readers
            t = get(ext[fix], r, NaN)
            row *= isnan(t) ? @sprintf("%14s%10s", "N/A", "N/A") : @sprintf("%12.1fms%9.2fx", t, ours / t)
        end
        println(row)
    end
    println("\nMedian times in milliseconds. XLSX/reader < 1.0x = XLSX.jl is faster. The Python",
            "\nreaders are timed in Python itself; calamine and openpyxl read the whole sheet.")
end

# ── Against Python's writers ──────────────────────────────────────────────────
# external_write.jl's times for openpyxl, XlsxWriter and rustpy-xlsxwriter writing the
# same tables, beside the newest XLSX.jl version's `writetable_only` (and
# `writetable_streaming`, where it has it).

let path = joinpath(RESULTS_DIR, "external_write.toml")
    isfile(path) || return
    ext = TOML.parsefile(path)
    i = findlast(l -> haskey(all_results, l), first.(VERSIONS))
    isnothing(i) && return
    v = first(VERSIONS[i])
    writers = [("openpyxl", "openpyxl"), ("openpyxl_write_only", "opx wo"),
               ("xlsxwriter", "xlsxwr"), ("xlsxwriter_constant_memory", "xlsxwr cm"),
               ("rustpy_xlsxwriter", "rustpy")]
    println("\n" * "="^60)
    println("WRITING A TABLE: XLSX.jl $v AGAINST PYTHON'S WRITERS")
    println("="^60)
    vers = get(ext, "versions", Dict())
    println("Python ", get(ext, "python", "?"), "; ", join(("$r $(get(vers, r, "?"))" for r in ("openpyxl", "xlsxwriter", "rustpy_xlsxwriter")), ", "))
    println(@sprintf("%-16s%12s%12s", "fixture", "XLSX", "XLSX strm") * join(@sprintf("%12s", s) for (_, s) in writers) *
            @sprintf("%14s%14s", "XLSX/xlsxwr", "XLSX/rustpy"))
    for fix in fixtures
        haskey(ext, fix) && haskey(all_results[v], fix) && haskey(all_results[v][fix], "writetable_only") || continue
        ours = median(all_results[v][fix]["writetable_only"]).time / 1e6
        strm = haskey(all_results[v][fix], "writetable_streaming") ?
               median(all_results[v][fix]["writetable_streaming"]).time / 1e6 : NaN
        row = @sprintf("%-16s%10.1fms", fix, ours) * (isnan(strm) ? @sprintf("%12s", "N/A") : @sprintf("%10.1fms", strm))
        for (w, _) in writers
            t = get(ext[fix], w, NaN)
            row *= isnan(t) ? @sprintf("%12s", "N/A") : @sprintf("%10.1fms", t)
        end
        for w in ("xlsxwriter", "rustpy_xlsxwriter")
            t = get(ext[fix], w, NaN)
            row *= isnan(t) ? @sprintf("%14s", "N/A") : @sprintf("%13.2fx", ours / t)
        end
        println(row)
    end
    println("\nMedian times in milliseconds, writing alone (the table is read beforehand).",
            "\nXLSX/writer < 1.0x = XLSX.jl is faster. opx wo = openpyxl write_only=True;",
            "\nxlsxwr cm = XlsxWriter constant_memory=True. The Python writers are timed in Python itself.")
end
