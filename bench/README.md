## Benchmarks

This directory contains a benchmark suite comparing XLSX.jl performance across three versions.

### Setup

1. Create a benchmark folder with a structure like:

        path_to_bench_folder/bench
            fixtures/
            results/
            envs/
                dev/
                v0_10/
                v0_11/
                v0_12/

   This folder is your local standalone benchmark project.  
   It is **not** inside the XLSX.jl repository and does **not** contain a checkout of XLSX.jl.

2. Copy the benchmark scripts from the XLSX.jl repo into your benchmark folder:

        bench/bench_worker.jl
        bench/generate_fixtures.jl
        bench/report.jl
        bench/run_benchmarks.jl

   These scripts will load XLSX.jl **from the active environment**, i.e. the installed version.

3. Copy the appropriate `Project.toml` into each env sub-folder:

        env/dev/Project.toml
        env/v0_10/Project.toml
        env/v0_11/Project.toml
        env/v0_12/Project.toml
        
4. Instantiate all three environments (each pins a specific XLSX.jl version):

        julia --project=envs/dev   -e "using Pkg; Pkg.instantiate()"
        julia --project=envs/v0_12 -e "using Pkg; Pkg.instantiate()"
        julia --project=envs/v0_10 -e "using Pkg; Pkg.instantiate()"
        julia --project=envs/v0_11 -e "using Pkg; Pkg.instantiate()"

   Each environment contains a different XLSX.jl version.  
   Julia will load the correct version automatically when benchmarks run.

5. Generate fixtures (written to `bench/fixtures/`):

        julia --project=envs/dev generate_fixtures.jl

   Fixtures are shared across all versions and are not committed to git.

6. Run benchmarks (results written to `bench/results/`):

        julia --project=envs/dev run_benchmarks.jl

   This script internally calls `bench_worker.jl` once per version, activating the correct environment each time.

7. Report results:

        julia --project=envs/dev report.jl

### Fixture descriptions

| Fixture | Rows | Numeric cols | String cols | Notes |
|---------|------|-------------|-------------|-------|
| small | 100 | 5 | 3 | Small file baseline |
| medium | 5,000 | 20 | 10 | Medium file |
| large | 50,000 | 50 | 20 | Large mixed file |
| wide_few | 100 | 200 | 50 | Wide, few rows |
| tall_few | 10,000 | 3 | 2 | Tall, few columns |
| sst_unique | 50,000 | 0 | 10 | All unique strings |
| sst_repeated | 50,000 | 0 | 10 | Repeated strings from small pool |
| sst_mixed | 50,000 | 5 | 5 | Mix of unique and repeated strings |
| numeric_only | 50,000 | 20 | 0 | Pure numeric data |
| dates_heavy | 50,000 | 10 | 0 | Date values only |

### Versions compared

- `v0.10` — EzXML.jl based implementation
- `v0.11` — First XML.jl based implementation using XML.jl v0.3
- `v0.12` — Updated XLSX.jl implementation adopting XML.jl v0.4 
