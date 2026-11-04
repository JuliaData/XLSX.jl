# quick.jl
# A short benchmark for checking successive changes to the working tree: `open`,
# `readtable`, `readtable_subset` and (on xl_gaps) `readtable_gaps` on a few fixtures,
# for envs/dev against v0.12 (0.12.4) and v0.10 (0.10.4). Short time budgets, so expect
# a few percent of noise. v0.10.4 reads xl_gaps differently (it drops absent rows),
# so compare `readtable_gaps` with v0.12 only.
#
# Usage: julia --project=. quick.jl [label]
#   label   tag for this run's dev results (default "dev"); earlier runs are kept in
#           results/quick/, so stages can be compared: quick.jl stage0, quick.jl stage1, …
# The baselines (v0.12, v0.10) are run once and reused; delete results/quick/ to redo.

using BenchmarkTools, Printf

const ROOT      = @__DIR__
const OUT       = joinpath(ROOT, "results", "quick")
const FIXTURES  = "medium,large,xl_medium,xl_issue462,xl_gaps"
const BENCHES   = "open,readtable,readtable_subset,readtable_gaps"
mkpath(OUT)

label = isempty(ARGS) ? "dev" : ARGS[1]

runs = [("v0.10", "v0_10", false), ("v0.12", "v0_12", false), (label, "dev", true)]
for (lab, env, always) in runs
    out = joinpath(OUT, "$(lab).json")
    (!always && isfile(out)) && (println("Reusing $out"); continue)
    println("\n== $lab ($(env))")
    withenv("BENCH_FIXTURES" => FIXTURES, "BENCH_BENCHMARKS" => BENCHES, "BENCH_SECONDS" => "8") do
        run(`julia --project=$(joinpath(ROOT, "envs", env)) --threads=8 $(joinpath(ROOT, "bench_worker.jl")) $lab $(joinpath(ROOT, "fixtures")) $out`)
    end
end

# Table: every quick result present, oldest first, so stage-by-stage progress shows.
labels = [splitext(f)[1] for f in sort(readdir(OUT); by = f -> mtime(joinpath(OUT, f))) if endswith(f, ".json")]
res = Dict(l => BenchmarkTools.load(joinpath(OUT, "$l.json"))[1] for l in labels)

med(l, f, b) = (haskey(res[l], f) && haskey(res[l][f], b)) ? median(res[l][f][b]) : nothing

println()
@printf("%-30s", "fixture / benchmark")
foreach(l -> @printf("%14s", l), labels)
@printf("%16s%16s\n", "$label/v0.12", "$label/v0.10")
for f in split(FIXTURES, ','), b in split(BENCHES, ',')
    ms = [med(l, f, b) for l in labels]
    all(isnothing, ms) && continue
    @printf("%-30s", "$f/$b")
    foreach(m -> isnothing(m) ? @printf("%14s", "—") : @printf("%12.1fms", m.time / 1e6), ms)
    d, v12, v10 = med(label, f, b), med("v0.12", f, b), med("v0.10", f, b)
    for base in (v12, v10)
        isnothing(d) || isnothing(base) ? @printf("%16s", "—") : @printf("%15.2fx", d.time / base.time)
    end
    println()
end
println("\nAllocated memory (MB):")
for f in split(FIXTURES, ','), b in split(BENCHES, ',')
    ms = [med(l, f, b) for l in labels]
    all(isnothing, ms) && continue
    @printf("%-30s", "$f/$b")
    foreach(m -> isnothing(m) ? @printf("%14s", "—") : @printf("%12.1fMB", m.memory / 2^20), ms)
    println()
end
