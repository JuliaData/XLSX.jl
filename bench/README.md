## Benchmarks

This directory contains a benchmark suite comparing XLSX.jl performance across the
released versions v0.10–v0.13, upstream `master` on GitHub, and the working tree
(`dev`), with a quick check of `dev` against the releases.

### Setup

1. Create a benchmark folder with a structure like:

        path_to_bench_folder/bench
            fixtures/
            results/
            envs/
                dev/
                master/
                v0_10/
                v0_11/
                v0_12/
                v0_13/

   This folder can be the repository's own `bench/` directory or a standalone
   benchmark project outside the repository. Two environments track unreleased code:

   - `envs/dev` is the checkout `bench/` sits in (`[sources] XLSX = {path = "../../.."}`):
     whatever branch you have checked out, including uncommitted changes. It works
     only when `bench/` is inside a clone of XLSX.jl. In a standalone folder, point its
     `path` at a checkout, or use `envs/master`.
   - `envs/master` is upstream `master` on GitHub
     (`[sources] XLSX = {url = "https://github.com/JuliaData/XLSX.jl", rev = "master"}`),
     fetched when the environment is instantiated. It works anywhere.

   Both need Julia 1.11 or later for `[sources]`. The `v0_*` environments pin
   registered releases.

2. If using a standalone folder, copy the benchmark scripts from the XLSX.jl repo into it:

        bench/bench_worker.jl
        bench/generate_fixtures.jl
        bench/generate_xl_fixtures.jl
        bench/xl_fixture_spec.jl
        bench/check_xl_fixtures.jl
        bench/quick.jl
        bench/report.jl
        bench/run_benchmarks.jl

   These scripts load XLSX.jl **from the active environment**, i.e. the installed version.

3. Copy the appropriate `Project.toml` into each env sub-folder:

        envs/dev/Project.toml
        envs/master/Project.toml
        envs/v0_10/Project.toml
        envs/v0_11/Project.toml
        envs/v0_12/Project.toml
        envs/v0_13/Project.toml

4. Instantiate the environments:

        julia --project=.          -e "using Pkg; Pkg.instantiate()"
        julia --project=envs/dev   -e "using Pkg; Pkg.instantiate()"
        julia --project=envs/master -e "using Pkg; Pkg.update()"    # fetch the latest master
        julia --project=envs/v0_10 -e "using Pkg; Pkg.instantiate()"
        julia --project=envs/v0_11 -e "using Pkg; Pkg.instantiate()"
        julia --project=envs/v0_12 -e "using Pkg; Pkg.instantiate()"
        julia --project=envs/v0_13 -e "using Pkg; Pkg.instantiate()"

   Each environment contains a different XLSX.jl version, and Julia loads the right one
   when the benchmarks run.

5. Generate fixtures (written to `bench/fixtures/`):

        julia --project=. generate_fixtures.jl
        julia --project=. generate_xl_fixtures.jl

   Fixtures are shared across all versions. They, and everything in `results/`, are
   ignored by git.

6. Run benchmarks (results written to `bench/results/`):

        julia --project=. run_benchmarks.jl                 # all versions
        julia --project=. run_benchmarks.jl master dev      # just these

   This script calls `bench_worker.jl` once per version, activating the right
   environment each time. A version whose results file already exists is skipped.

7. Report results:

        julia --project=. report.jl

### Quick check of the working tree

`quick.jl` runs a short subset — `open`, `readtable`, `readtable_subset` and
`readtable_gaps` on `medium`, `large` and the `xl_*` fixtures — for `envs/dev` against
v0.12.0 and v0.10.4, and prints every earlier quick run alongside, so successive
changes can be compared:

    julia --project=. quick.jl stage1

The baselines are run once and reused; delete `results/quick/` to redo them.

`bench_worker.jl` takes three optional environment variables, which `quick.jl` uses:
`BENCH_FIXTURES` and `BENCH_BENCHMARKS` (comma-separated names to run) and
`BENCH_SECONDS` (time budget per benchmark, default 30).

### Fixture descriptions

Written through XLSX.jl by `generate_fixtures.jl` (pretty-printed XML):

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
| multi_sheet | 5 × 20,000 | 10 | 5 | Plus 5 formula columns per sheet |

Written directly, in the compact form Excel writes, by `generate_xl_fixtures.jl`. These
model the workbook in issue #462, which can't be shared: a macro-enabled `.xlsm` whose
second sheet, `Data`, has title rows, headers on row 5, the table in A:CF and helper
columns beyond it, `s` on every cell, shared strings, dates, booleans, errors and 20%
formula cells with cached values (half shared, the rest long lookups).

| Fixture | Data rows | Sheet XML | Notes |
|---------|-----------|-----------|-------|
| xl_medium | 3,000 | 13 MB | Gap rows only after the table |
| xl_issue462 | 30,000 | 137 MB | The #462 shape; gap rows only after the table |
| xl_gaps | 3,000 | 13 MB | Absent, helper-only and empty rows inside the table |

`readtable_subset` is the #462 call, `readtable(path, "Data", "A:CF"; first_row=5)`,
which stops at the first empty row. `readtable_gaps` reads `xl_gaps` with
`stop_in_empty_row=false, keep_empty_rows=true`. v0.10.4 drops absent rows in that
mode, so compare `readtable_gaps` with v0.12 and later only.

Every cell of the `xl_*` fixtures is a pure function of its row and column
(`xl_fixture_spec.jl`). `check_xl_fixtures.jl` reads each one as the benchmarks do,
checks every value against the spec and prints a digest for comparing versions:

    julia --project=envs/dev check_xl_fixtures.jl

### Differential tests

`test/test_files/Differential_tests.jl` checks that `readtable` returns exactly what
`gettable` returns on an open file, across generated empty-row cases, randomised
sheets and every workbook in `test/data`. Two environment variables extend it:
`XLSX_DIFF_FIXTURES=<bench/fixtures path>` adds the `xl_*` fixtures, and
`XLSX_FULL_DIFF=1` runs the whole keyword covering array for every case (slow).

### Versions compared

- `v0.10` — EzXML.jl based implementation
- `v0.11` — First XML.jl based implementation using XML.jl v0.3. Pinned to its last
  patch, 0.11.11: 0.11.0 has a threading race in its style cache (fails
  intermittently with `--threads=8`) and rejects `.xlsm`, so it can't run the
  `xl_*` fixtures
- `v0.12` — Updated XLSX.jl implementation adopting XML.jl v0.4
- `v0.13` — Adds native Excel chart support
- `master` — upstream `master` on GitHub
- `dev` — the working tree
