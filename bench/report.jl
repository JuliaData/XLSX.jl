using BenchmarkTools
using Printf
#using UnicodePlots

const ROOT        = @__DIR__
const RESULTS_DIR = joinpath(ROOT, "results")

VERSIONS = [
    ("v0.10", joinpath(ROOT, "envs", "v0_10")),
    ("v0.11", joinpath(ROOT, "envs", "v0_11")),
    ("v0.12", joinpath(ROOT, "envs", "v0_12")),
]

const fixtures = [
    "small", "medium", "large", "wide_few", "tall_few",
    "sst_unique", "sst_repeated", "sst_mixed",
    "numeric_only", "dates_heavy", "multi_sheet",
]

const benchmarks = [
    "readtable", "readxlsx", "eachrow",
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
    ver_labels = first.(VERSIONS)

    header = @sprintf("%-30s", "fixture / benchmark")
    for v in ver_labels
        header *= @sprintf("%15s", v)
    end
    header *= @sprintf("%15s%15s", "v0.11/v0.10", "dev/v0.10")
    println(header)
    println("-"^(30 + 15*length(ver_labels) + 30))

    for fix in fixtures, bench in benchmarks
        row = @sprintf("%-30s", "$(fix)/$(bench)")
        medians = Dict{String,Float64}()
        for v in ver_labels
            haskey(all_results, v)             || continue
            haskey(all_results[v], fix)        || continue
            haskey(all_results[v][fix], bench) || continue
            t = median(all_results[v][fix][bench]).time / 1e6
            medians[v] = t
            row *= @sprintf("%14.1fms", t)
        end
        base = get(medians, "v0.10", NaN)
        for (_, key) in [("v0.11/v0.10", "v0.11"), ("v0.12/v0.10", "v0.12")]
            t = get(medians, key, NaN)
            row *= isnan(base) || isnan(t) || base == 0 ?
                @sprintf("%15s", "N/A") :
                @sprintf("%14.2fx", t / base)
        end
        println(row)
    end
    println("\nMedian times in milliseconds. Ratio < 1.0x = faster than v0.10.")
end

#=
# ── Bar charts ────────────────────────────────────────────────────────────────

println("\n" * "="^60)
println("BAR CHARTS")
println("="^60)

ver_labels = first.(VERSIONS)   # must be visible here — defined at top level

for bench in benchmarks, fix in fixtures
    times  = Float64[]
    labels = String[]
    for v in ver_labels
        haskey(all_results, v)             || continue
        haskey(all_results[v], fix)        || continue
        haskey(all_results[v][fix], bench) || continue
        push!(times,  median(all_results[v][fix][bench]).time / 1e6)
        push!(labels, v)
    end
    isempty(times) && continue
    println()
    display(barplot(labels, times; title="$(fix) / $(bench)", xlabel="milliseconds", width=60))
end
=#