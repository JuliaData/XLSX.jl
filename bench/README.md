## Benchmarks

This directory contains a benchmark suite comparing XLSX.jl performance across three versions.

### Setup

From the `bench/` directory:

1. Clone the repo and enter the bench directory:

        git clone -b Explicit-Test-Output https://github.com/TimG1964/XLSX.jl.git
        cd XLSX.jl/bench

2. Instantiate all three environments:

        julia --project=envs/dev -e "using Pkg; Pkg.instantiate()"
        julia --project=envs/v0_10_4 -e "using Pkg; Pkg.instantiate()"
        julia --project=envs/v0_11_10 -e "using Pkg; Pkg.instantiate()"

3. Generate fixtures (written to `bench/fixtures/`, not committed to git):

        julia --project=envs/dev generate_fixtures.jl

4. Run benchmarks (results written to `bench/results/`):

        julia --project=envs/dev run_benchmarks.jl

5. Report results:

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

- `v0.10.4` — EzXML.jl based implementation (JuliaIO/XLSX.jl tag v0.10.4)
- `v0.11.10` — XML.jl based implementation (JuliaIO/XLSX.jl tag v0.11.10)
- `dev` — WIP XLSX v0.12 based on WIP XML.jl v0.4 implementation