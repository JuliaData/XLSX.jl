# run_benchmarks.jl
# Orchestrates all benchmark runs and prints a comparison table.
# Usage: julia --project=. run_benchmarks.jl [versions...]   e.g. `v0.12 dev`; default all

using Pkg

const ROOT         = @__DIR__
const FIXTURES_DIR = joinpath(ROOT, "fixtures")
# BENCH_RESULTS: another results directory, for a set of runs that must not mix with
# earlier ones (from another machine, say)
const RESULTS_DIR  = get(ENV, "BENCH_RESULTS", joinpath(ROOT, "results"))
mkpath(RESULTS_DIR)

VERSIONS = [
    ("v0.10", joinpath(ROOT, "envs", "v0_10")),
    ("v0.11", joinpath(ROOT, "envs", "v0_11")),
    ("v0.12", joinpath(ROOT, "envs", "v0_12")),
    ("v0.13", joinpath(ROOT, "envs", "v0_13")),
    ("master", joinpath(ROOT, "envs", "master")), # upstream master on GitHub
    ("dev",   joinpath(ROOT, "envs", "dev")),     # the working tree
]
isempty(ARGS) || filter!(v -> first(v) in ARGS, VERSIONS)

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
    # `master`, and XML.jl in `dev`, track a git branch: resolve/instantiate would keep the
    # commit pinned in the Manifest, so update to the current upstream tip instead.
    pkg_cmd = ver_label == "master" ? "Pkg.update(\"XLSX\")" :
              ver_label == "dev"    ? "Pkg.update(\"XML\")"  : "Pkg.resolve()"
    run(`julia --project=$env_path -e "using Pkg; $pkg_cmd; Pkg.instantiate()"`)

    cmd = `julia --project=$env_path --threads=8
                 $(joinpath(ROOT, "bench_worker.jl"))
                 $ver_label
                 $FIXTURES_DIR
                 $outfile`
    run(cmd)
end

println("\nAll done. Run report.jl to see results.")