# run_benchmarks.jl
# Orchestrates all three benchmark runs and prints a comparison table.
# Usage: julia --project=. run_benchmarks.jl

using Pkg

const ROOT         = @__DIR__
const FIXTURES_DIR = joinpath(ROOT, "fixtures")
const RESULTS_DIR  = joinpath(ROOT, "results")
mkpath(RESULTS_DIR)

VERSIONS = [
    ("v0.10", joinpath(ROOT, "envs", "v0_10")),
    ("v0.11", joinpath(ROOT, "envs", "v0_11")),
    ("v0.12", joinpath(ROOT, "envs", "v0_12")),
]

for (ver_label, env_path) in VERSIONS
    outfile = joinpath(RESULTS_DIR, "$(ver_label).json")

    if isfile(outfile)
        println("Results for $ver_label already exist, skipping run.")
        continue
    end

    println("\n" * "="^60)
    println("Benchmarking XLSX $ver_label")
    println("="^60)

    println("Instantiating environment…")
    run(`julia --project=$env_path -e "using Pkg; Pkg.instantiate()"`)

    cmd = `julia --project=$env_path --threads=8
                 $(joinpath(ROOT, "bench_worker.jl"))
                 $ver_label
                 $FIXTURES_DIR
                 $outfile`
    run(cmd)
end

println("\nAll done. Run report.jl to see results.")