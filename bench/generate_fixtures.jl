# generate_fixtures.jl
# Generates synthetic .xlsx files for benchmarking.
# Run once: julia --project=envs/dev generate_fixtures.jl

using XLSX
using Dates

const FIXTURES_DIR = joinpath(@__DIR__, "fixtures")
mkpath(FIXTURES_DIR)

# Long strings with minimal shared-string reuse: embed a unique counter.
long_string(i, j) = "LongUniqueString_row$(i)_col$(j)_" * "X"^80

# Pool of repeated strings for sst_repeated / sst_mixed
const REPEATED_POOL = ["Alpha", "Beta", "Gamma", "Delta", "Epsilon",
                       "Zeta", "Eta", "Theta", "Iota", "Kappa",
                       "Lambda", "Mu", "Nu", "Xi", "Omicron"]

# Each fixture is (label, nrows, ncols_numeric, ncols_string)
const FIXTURES = [
    ("small",           100,    5,   3),
    ("medium",        5_000,   20,  10),
    ("large",        50_000,   50,  20),
    ("wide_few",        100,  200,  50),
    ("tall_few",     10_000,    3,   2),
    ("sst_unique",   50_000,    0,  10),
    ("sst_repeated", 50_000,    0,  10),
    ("sst_mixed",    50_000,    5,   5),
    ("numeric_only", 50_000,   20,   0),
    ("dates_heavy",  50_000,   10,   0),
]

for (label, nrows, ncols_num, ncols_string) in FIXTURES
    path = joinpath(FIXTURES_DIR, "$(label).xlsx")
    isfile(path) && (println("Skipping $label (exists)"); continue)
    println("Generating $label ($nrows rows × $(ncols_num + ncols_string) cols)…")

    XLSX.openxlsx(path, mode="w") do xf
        sheet = xf[1]
        XLSX.rename!(sheet, "Sheet1")

        headers = Any[
            ["num_$i" for i in 1:ncols_num]...,
            ["str_$i" for i in 1:ncols_string]...,
        ]
        ncols_total = ncols_num + ncols_string
        ncols_total > 0 && (sheet[1, 1:ncols_total] = headers)

        CHUNK = 1000
        for chunk_start in 1:CHUNK:nrows
            chunk_end = min(chunk_start + CHUNK - 1, nrows)
            for row in chunk_start:chunk_end
                r = row + 1  # +1 for header
                for c in 1:ncols_num
                    sheet[r, c] = if label == "dates_heavy"
                        Dates.Date(2020, 1, 1) + Dates.Day((row - 1) % 3650)
                    else
                        Float64(row * 1000 + c) + 0.123456
                    end
                end
                for c in 1:ncols_string
                    sheet[r, ncols_num + c] = if label == "sst_repeated"
                        REPEATED_POOL[((row - 1) * ncols_string + c - 1) % length(REPEATED_POOL) + 1]
                    elseif label == "sst_mixed"
                        row % 5 == 0 ? "unique_$(row)_$(c)" :
                            REPEATED_POOL[(c - 1) % length(REPEATED_POOL) + 1]
                    elseif c % 3 == 0
                        long_string(row, c)
                    else
                        "s$(row)_$(c)"
                    end
                end
            end
        end
    end
    println("  → written $path")
end

# Multi-sheet fixture — 5 sheets × 20k rows × 10 numeric + 5 string + 5 formula cols
let
    label = "multi_sheet"
    path  = joinpath(FIXTURES_DIR, "$(label).xlsx")
    nsheets    = 5
    nrows      = 20_000
    ncols_num  = 10
    ncols_str  = 5
    ncols_form = 5

    if isfile(path)
        println("Skipping $label (exists)")
    else
        println("Generating $label ($nsheets sheets × $nrows rows × $(ncols_num + ncols_str + ncols_form) cols)…")
        try
            XLSX.openxlsx(path, mode="w") do xf
                for sheet_no in 1:nsheets
                    sheet = sheet_no == 1 ? xf[1] : XLSX.addsheet!(xf, "Sheet$sheet_no")
                    XLSX.rename!(sheet, "Sheet$sheet_no")

                    headers = Any[
                        ["num_$i"     for i in 1:ncols_num]...,
                        ["str_$i"     for i in 1:ncols_str]...,
                        ["formula_$i" for i in 1:ncols_form]...,
                    ]
                    ncols_total = ncols_num + ncols_str + ncols_form
                    sheet[1, 1:ncols_total] = headers

                    CHUNK = 1000
                    for chunk_start in 1:CHUNK:nrows
                        chunk_end = min(chunk_start + CHUNK - 1, nrows)
                        for row in chunk_start:chunk_end
                            r = row + 1  # +1 for header

                            # Numeric columns
                            for c in 1:ncols_num
                                sheet[r, c] = Float64(row * 1000 + c) + 0.123456
                            end

                            # String columns — repeated pool for SST stress
                            for c in 1:ncols_str
                                sheet[r, ncols_num + c] = REPEATED_POOL[((row - 1) * ncols_str + c - 1) % length(REPEATED_POOL) + 1]
                            end

                            # Formula columns — reference numeric columns
                            for c in 1:ncols_form
                                src_col = XLSX.encode_column_number(c)  # A, B, C, D, E
                                XLSX.setFormula(sheet, "$(XLSX.encode_column_number(ncols_num + ncols_str + c))$r",
                                    "=$(src_col)$r * 2 + $(c)")
                            end
                        end
                    end
                end
            end
            println("  → written $path")
        catch err
            println("ERROR generating $label: ", err)
            showerror(stdout, err, catch_backtrace())
            isfile(path) && rm(path)  # remove corrupt partial file
        end
    end
end

println("Done.")