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
    "single_cell", "writetable", "open_readwrite",
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
