#=
Differential tests (#462): new read paths checked against an oracle.

  - `readtable` against `gettable` (described below);
  - value rows against cache rows;
  - the filtered cell kernel against the node decoder;
  - uncached ranged reads (`getdata`/`getcellrange`/`getcell`) against the row decoders;
  - the row stream against the node walk.

Each differential is followed by the point tests for the same code path.

## readtable

`readtable(source, …)` builds its own `XLSXFile` and discards it. The oracle is
`gettable` on a file opened with `readxlsx`, which stays on the worksheet-cache path,
so any change to how `readtable` reads (or to `gettable` itself) must produce
exactly what the oracle produces: the same exception, or the same column labels,
column container types, row count, values (`isequal` and `typeof`), and the same
sequence of calls to `stop_in_row_function`. `eachtablerow` is checked against the
oracle's values too, so all three views of a table agree.

Corpora:
  1. a generated empty-row corpus: every sequence of 0–3 row states after a header,
     plus hand-made special sheets;
  2. every sheet of every workbook in `test/data`;
  3. randomised sparse sheets from fixed seeds;
  4. optionally, the bench fixtures (`XLSX_DIFF_FIXTURES=<dir>`).

Rows-deciding keywords (`header`, `stop_in_empty_row`, `keep_empty_rows`, `first_row`,
columns) are crossed in full on the generated sheets and pairwise on the test/data
workbooks; the others are drawn from a pairwise covering array. Each check computes
the oracle once and runs `readtable` both on value rows and on the cache.

The test/data workbooks are slow to open (each check reopens one), and the generated
sheets already cross every selection in full, so the kernel differentials below run
a test/data sheet with the full read plus a few other selections, rotated from sheet
to sheet so every selection still runs on the workbooks (see `_diff_choices`).

Set `XLSX_FULL_DIFF=1` to cross everything on every sheet: every row-deciding
combination with the whole covering array, and every selection (slow: tens of minutes).
=#

const _DIFF_FULL = get(ENV, "XLSX_FULL_DIFF", "0") == "1"
const _DIFF_CHECKS = Ref(0)        # comparisons run in the current section

# The workbook builder `_diff_build_xlsx` is in helpers.jl (Read-internals uses it too).

# ── Empty-row corpus ──────────────────────────────────────────────────────────
# The table sits in B:D with its header on `header_row`. Column F is always empty
# (so a range widened to F reaches no data) and column G holds "outside" cells.
# Shared strings 0–2 are the labels a, b, c; 3 is ""; 4 is "NA"; 5+ are data strings.

const _DIFF_ROW_STATES = (:absent, :norow_cells, :styled_empty, :outside_only,
                          :empty_strings, :missing_strings, :formula_nocache, :data)

_diff_row(r, cells; attrs = "") = "<row r=\"$r\" spans=\"2:7\" x14ac:dyDescent=\"0.25\"$attrs>$cells</row>"

function _diff_state_row(state::Symbol, r::Int, sst::Vector{String})::String
    if state === :absent
        return ""
    elseif state === :norow_cells
        return "<row r=\"$r\"/>"
    elseif state === :styled_empty
        return _diff_row(r, "<c r=\"B$r\" s=\"2\"/><c r=\"C$r\" s=\"2\"/><c r=\"D$r\" s=\"2\"/>")
    elseif state === :outside_only
        return _diff_row(r, "<c r=\"G$r\"><v>$(100 + r)</v></c>")
    elseif state === :empty_strings
        return _diff_row(r, "<c r=\"B$r\" t=\"s\"><v>3</v></c>" *
                            "<c r=\"C$r\" t=\"inlineStr\"><is><t></t></is></c>" *
                            "<c r=\"D$r\" t=\"str\"><f>\"\"</f><v></v></c>")
    elseif state === :missing_strings
        return _diff_row(r, "<c r=\"B$r\" t=\"s\"><v>4</v></c><c r=\"C$r\" t=\"s\"><v>4</v></c>" *
                            "<c r=\"D$r\" t=\"inlineStr\"><is><t>NA</t></is></c>")
    elseif state === :formula_nocache
        return _diff_row(r, "<c r=\"B$r\"><f>1+$r</f></c><c r=\"C$r\" t=\"str\"><f>\"x\"&amp;$r</f></c>" *
                            "<c r=\"D$r\" s=\"1\"><f>TODAY()</f></c>")
    else # :data — Int-valued number, shared string, date; plus an outside cell
        push!(sst, "s$r")
        return _diff_row(r, "<c r=\"B$r\"><v>$r</v></c><c r=\"C$r\" t=\"s\"><v>$(length(sst) - 1)</v></c>" *
                            "<c r=\"D$r\" s=\"1\"><v>$(45000 + r)</v></c><c r=\"G$r\"><v>1</v></c>")
    end
end

_diff_base_sst() = ["a", "b", "c", "", "NA"]
_diff_header(r) = _diff_row(r, "<c r=\"B$r\" t=\"s\"><v>0</v></c><c r=\"C$r\" t=\"s\"><v>1</v></c><c r=\"D$r\" t=\"s\"><v>2</v></c>")

function _diff_corpus_sheet(states)::Vector{UInt8}
    sst = _diff_base_sst()
    xml = _diff_header(1) * join(_diff_state_row(s, i + 1, sst) for (i, s) in enumerate(states))
    return _diff_build_xlsx(xml, sst)
end

# Hand-made sheets for the cases a slot sequence can't express.
function _diff_special_sheets()::Vector{Pair{String,Vector{UInt8}}}
    out = Pair{String,Vector{UInt8}}[]
    d(r) = _diff_state_row(:data, r, sst)
    sst = _diff_base_sst()
    push!(out, "no rows" => _diff_build_xlsx("", copy(sst)))
    push!(out, "header only" => _diff_build_xlsx(_diff_header(1), copy(sst)))
    sst = _diff_base_sst()
    push!(out, "header absent, data from row 2" => _diff_build_xlsx(d(2) * d(3), sst))
    sst = _diff_base_sst()
    push!(out, "header row empty" => _diff_build_xlsx(_diff_state_row(:styled_empty, 1, sst) * d(2) * d(3), sst))
    sst = _diff_base_sst()
    push!(out, "header with gap" => _diff_build_xlsx(
        _diff_row(1, "<c r=\"B1\" t=\"s\"><v>0</v></c><c r=\"D1\" t=\"s\"><v>2</v></c>") * d(2) * d(3), sst))
    sst = _diff_base_sst()
    push!(out, "title rows, header on 5" => _diff_build_xlsx(
        _diff_row(1, "<c r=\"A1\" t=\"inlineStr\"><is><t>Report</t></is></c>") *
        _diff_row(3, "<c r=\"A3\" t=\"inlineStr\"><is><t>notes</t></is></c>") *
        _diff_header(5) * d(6) * d(8) * d(9), sst))
    sst = _diff_base_sst()
    push!(out, "leading and trailing gaps" => _diff_build_xlsx(_diff_header(1) * d(5) * d(6) * d(12), sst))
    sst = _diff_base_sst()
    push!(out, "out-of-order rows" => _diff_build_xlsx(_diff_header(1) * d(4) * d(2) * d(3), sst))
    sst = _diff_base_sst()
    push!(out, "duplicated row" => _diff_build_xlsx(_diff_header(1) * d(2) * d(2) * d(3), sst))
    sst = _diff_base_sst()
    push!(out, "mixed types per column" => _diff_build_xlsx(_diff_header(1) *
        _diff_row(2, "<c r=\"B2\"><v>1</v></c><c r=\"C2\" s=\"1\"><v>45000</v></c><c r=\"D2\" t=\"b\"><v>1</v></c>") *
        _diff_row(3, "<c r=\"B3\"><v>1.5</v></c><c r=\"C3\" s=\"1\"><v>45000.5</v></c><c r=\"D3\" t=\"e\"><v>#N/A</v></c>") *
        _diff_row(4, "<c r=\"B4\" t=\"inlineStr\"><is><t>x</t></is></c><c r=\"C4\" s=\"1\"><v>45001</v></c><c r=\"D4\" t=\"b\"><v>0</v></c>"), sst))
    return out
end

# Columns and missing_strings values the readtable differentials cross.
const _DIFF_COLUMNS = (nothing, "B:D", "B:C", "B:F", "C:C", "A:D", "H:I")
const _DIFF_MS_OPTIONS = (nothing, "NA", ["NA", ""], "a")

# ── Running one comparison ────────────────────────────────────────────────────

struct _DiffCall
    row::Int
    values::Vector{Any}
end

function _diff_stop_fn(kind::Symbol, log::Vector{_DiffCall})
    kind === :none && return nothing
    pred = kind === :at2 ? (r -> XLSX.row_number(r) >= 2) :
           kind === :col1missing ? (r -> ismissing(r[1])) :
           kind === :always ? (r -> true) : (r -> false)
    # Values by index: `TableRow` iterates but has no `length`, so it can't be collected.
    return r -> (push!(log, _DiffCall(XLSX.row_number(r),
                                      Any[r[i] for i in XLSX.table_column_numbers(r)])); pred(r))
end

# Outcome of one call: the DataTable, or the exception it threw.
_diff_run(f) = try f() catch e; e end

function _diff_describe(o)
    o isa Exception && return "exception $(typeof(o)): $(sprint(showerror, o))"
    return "table $(o.column_labels) $(map(typeof, o.data)) rows=$(isempty(o.data) ? 0 : length(o.data[1]))"
end

# First difference between two outcomes, or `nothing` when they agree exactly.
function _diff_compare(a, b)::Union{Nothing,String}
    if a isa Exception || b isa Exception
        (a isa Exception && b isa Exception) || return "one threw: $(_diff_describe(a)) vs $(_diff_describe(b))"
        typeof(a) == typeof(b) || return "exception types differ: $(typeof(a)) vs $(typeof(b))"
        sprint(showerror, a) == sprint(showerror, b) || return "exception messages differ: $(sprint(showerror, a)) vs $(sprint(showerror, b))"
        return nothing
    end
    a.column_labels == b.column_labels || return "labels differ: $(a.column_labels) vs $(b.column_labels)"
    length(a.data) == length(b.data) || return "column counts differ"
    for (ci, (ca, cb)) in enumerate(zip(a.data, b.data))
        typeof(ca) == typeof(cb) || return "column $ci type differs: $(typeof(ca)) vs $(typeof(cb))"
        length(ca) == length(cb) || return "column $ci length differs: $(length(ca)) vs $(length(cb))"
        for ri in eachindex(ca)
            x, y = ca[ri], cb[ri]
            (isequal(x, y) && typeof(x) == typeof(y)) ||
                return "column $ci row $ri differs: $(repr(x))::$(typeof(x)) vs $(repr(y))::$(typeof(y))"
        end
    end
    return nothing
end

function _diff_compare_calls(a::Vector{_DiffCall}, b::Vector{_DiffCall})::Union{Nothing,String}
    length(a) == length(b) || return "stop_in_row_function called $(length(a)) vs $(length(b)) times"
    for (i, (x, y)) in enumerate(zip(a, b))
        x.row == y.row || return "callback $i saw table row $(x.row) vs $(y.row)"
        (length(x.values) == length(y.values) &&
         all(isequal(p, q) && typeof(p) == typeof(q) for (p, q) in zip(x.values, y.values))) ||
            return "callback $i saw values $(x.values) vs $(y.values)"
    end
    return nothing
end

# ── Pairwise covering array for the secondary keywords ────────────────────────

# Greedy pairwise cover over `levels` (number of values per factor); deterministic.
function _diff_pairwise(levels::Vector{Int})::Vector{Vector{Int}}
    nf = length(levels)
    uncovered = Set{NTuple{4,Int}}()
    for i in 1:nf, j in (i+1):nf, a in 1:levels[i], b in 1:levels[j]
        push!(uncovered, (i, a, j, b))
    end
    candidates = vec(collect(Iterators.product((1:l for l in levels)...)))
    rows = Vector{Vector{Int}}()
    while !isempty(uncovered)
        best, best_n = candidates[1], -1
        for c in candidates
            n = 0
            for i in 1:nf, j in (i+1):nf
                (i, c[i], j, c[j]) in uncovered && (n += 1)
            end
            n > best_n && ((best, best_n) = (c, n))
        end
        for i in 1:nf, j in (i+1):nf
            delete!(uncovered, (i, best[i], j, best[j]))
        end
        push!(rows, collect(best))
    end
    return rows
end

# Secondary factors: enable_cache, infer_eltypes, normalizenames, column_labels,
# missing_strings, stop_in_row_function.
const _DIFF_SECONDARY_LEVELS = [2, 2, 2, 3, 4, 5]
const _DIFF_STOP_KINDS = (:none, :at2, :col1missing, :always, :never)

# `eachtablerow` gives untyped rows; its values must equal the oracle's (`isequal`).
function _diff_compare_rows(rows, oracle)::Union{Nothing,String}
    rows isa Exception && return oracle isa Exception ? nothing : "eachtablerow threw: $(_diff_describe(rows))"
    oracle isa Exception && return "eachtablerow succeeded where gettable threw"
    n = isempty(oracle.data) ? 0 : length(oracle.data[1])
    length(rows) == n || return "eachtablerow gave $(length(rows)) rows vs $n"
    for (ri, r) in enumerate(rows), ci in eachindex(oracle.data)
        isequal(r[ci], oracle.data[ci][ri]) || return "eachtablerow row $ri column $ci: $(repr(r[ci])) vs $(repr(oracle.data[ci][ri]))"
    end
    return nothing
end

# With `readtable` on the cache path (`XLSX._READTABLE_VALUE_ROWS[] = false`),
# `enable_cache=false` streams the sheet, which reads sheets with out-of-order or
# duplicated `<row r>` (never written by Excel) differently from the cache. That is
# the one accepted difference. With value rows (the default), `enable_cache` has no
# effect and such sheets fall back to the cache, so agreement must be exact.
_diff_nocache_allowed(malformed::Bool, value_rows::Bool)::Bool = malformed && !value_rows

# The two paths `readtable` is checked on (#462), and their names in mismatch reports.
const _DIFF_PATHS = (true => "value rows", false => "the cache")

# `f()` with `readtable` on value rows (`true`) or on the worksheet cache (`false`).
function _with_value_rows(f, value_rows::Bool)
    XLSX._READTABLE_VALUE_ROWS[] = value_rows
    try
        return f()
    finally
        XLSX._READTABLE_VALUE_ROWS[] = true
    end
end

"""
Compare `readtable`, on both of its paths, with the `gettable` oracle for one keyword
set. `xf` is the oracle's open file; `bytes` the workbook. Returns the mismatches,
each starting with the path it was found on ("via value rows: …", "via the cache: …",
or "eachtablerow: …").
"""
function _diff_check(bytes::Vector{UInt8}, xf, sheet, cols, kw::NamedTuple, sec::Vector{Int};
                     check_rows::Bool = true, malformed::Bool = false)::Vector{String}
    _DIFF_CHECKS[] += 1
    enable_cache = sec[1] == 1
    infer_eltypes = sec[2] == 1
    normalizenames = sec[3] == 2
    missing_strings = _DIFF_MS_OPTIONS[sec[5]]
    stop_kind = _DIFF_STOP_KINDS[sec[6]]

    # column_labels: nothing, the right length, or one too many (must throw identically)
    column_labels = nothing
    if sec[4] != 1
        probe = _diff_run(() -> isnothing(cols) ?
            XLSX.gettable(xf[sheet]; kw...) : XLSX.gettable(xf[sheet], cols; kw...))
        probe isa Exception && return String[]    # nothing more to learn for this case
        n = length(probe.column_labels) + (sec[4] == 3 ? 1 : 0)
        column_labels = [Symbol("L$i") for i in 1:n]
    end

    log_a = _DiffCall[]
    common(log) = (; kw..., column_labels, normalizenames, missing_strings,
                    stop_in_row_function = _diff_stop_fn(stop_kind, log))

    oracle = _diff_run(() -> isnothing(cols) ?
        XLSX.gettable(xf[sheet]; common(log_a)..., infer_eltypes) :
        XLSX.gettable(xf[sheet], cols; common(log_a)..., infer_eltypes))

    mismatches = String[]
    for (value_rows, name) in _DIFF_PATHS
        # The one known difference of the streaming path (see `_diff_nocache_allowed`)
        !enable_cache && _diff_nocache_allowed(malformed, value_rows) && continue
        log_b = _DiffCall[]
        subject = _with_value_rows(value_rows) do
            _diff_run(() -> if isnothing(cols)
                sheet == 1 && sec[6] % 2 == 0 ?   # exercise readtable(source) too
                    XLSX.readtable(IOBuffer(bytes); common(log_b)..., infer_eltypes, enable_cache) :
                    XLSX.readtable(IOBuffer(bytes), sheet; common(log_b)..., infer_eltypes, enable_cache)
            else
                XLSX.readtable(IOBuffer(bytes), sheet, cols; common(log_b)..., infer_eltypes, enable_cache)
            end)
        end
        m = _diff_compare(oracle, subject)
        m === nothing && (m = _diff_compare_calls(log_a, log_b))
        m === nothing || push!(mismatches, "via $name: readtable: " * m)
    end

    if check_rows
        log_c = _DiffCall[]
        rows = _diff_run(() -> collect(isnothing(cols) ?
            XLSX.eachtablerow(xf[sheet]; common(log_c)...) :
            XLSX.eachtablerow(xf[sheet], cols; common(log_c)...)))
        m = _diff_compare_rows(rows, oracle)
        m === nothing || push!(mismatches, "eachtablerow: " * m)
    end
    return mismatches
end

"""
Run the row-deciding keywords on one workbook sheet: their full product, or with
`pairwise = true` a pairwise cover of it plus the default keywords (the full product
under `XLSX_FULL_DIFF=1`). Secondary keywords come from the covering array, rotated
by `offset` so different sheets see different pairings (all of them under
`XLSX_FULL_DIFF=1`). Returns the mismatches found.
"""
function _diff_product(label, bytes, sheet, first_rows, columns, cover, offset;
                       malformed::Bool = false, pairwise::Bool = false)::Vector{String}
    mismatches = String[]
    xf = _diff_run(() -> XLSX.readxlsx(IOBuffer(bytes)))
    if xf isa Exception
        # The oracle can't open it: readtable must fail the same way on both paths.
        for (value_rows, name) in _DIFF_PATHS
            e = _with_value_rows(() -> _diff_run(() -> XLSX.readtable(IOBuffer(bytes), sheet)), value_rows)
            m = _diff_compare(xf, e)
            m === nothing || push!(mismatches, "$label: open: via $name: $m")
        end
        return mismatches
    end
    bools = (true, false)
    levels = [2, 2, 2, length(first_rows), length(columns)]
    picks = if pairwise && !_DIFF_FULL
        p = _diff_pairwise(levels)
        [1, 1, 2, 1, 1] in p ? p : push!(p, [1, 1, 2, 1, 1])   # header, stop_in_empty_row, no keep_empty_rows
    else
        [collect(i) for i in Iterators.product((1:l for l in levels)...)]
    end
    k = offset
    for (h, s, e, f, c) in picks
        kw = (; header = bools[h], stop_in_empty_row = bools[s], keep_empty_rows = bools[e],
                first_row = first_rows[f])
        cols = columns[c]
        secs = _DIFF_FULL ? cover : [cover[mod1(k += 1, length(cover))]]
        for sec in secs, m in _diff_check(bytes, xf, sheet, cols, kw, sec; malformed)
            push!(mismatches, "$label: cols=$(repr(cols)) $(kw) sec=$(sec): $m")
        end
    end
    return mismatches
end

# The tests of one section, one per `readtable` path plus one for `eachtablerow`.
function _diff_test_paths(mismatches::Vector{String})
    for (_, name) in _DIFF_PATHS
        @testset "readtable via $name" begin
            @test !any(m -> occursin("via $name: ", m), mismatches)
        end
    end
    @testset "eachtablerow" begin
        @test !any(m -> occursin("eachtablerow: ", m), mismatches)
    end
    @test isempty(mismatches)
end

# Prints a summary and the first mismatches, if any; `XLSX_DIFF_LOG=<file>` appends all of them to a file.
function _diff_report(mismatches::Vector{String}, what::String, t0::Float64)
    isempty(mismatches) && return
    println("Differential $what: $(_DIFF_CHECKS[]) checks, $(length(mismatches)) mismatches, ",
            round(time() - t0, digits = 1), " s")
    log = get(ENV, "XLSX_DIFF_LOG", "")
    isempty(log) || open(io -> foreach(m -> println(io, what, " | ", m), mismatches), log, "a")
    println("Differential mismatches ($what), first $(min(20, length(mismatches))) of $(length(mismatches)):")
    foreach(m -> println("  ", m), first(mismatches, 20))
end

# ── Randomised sheets ─────────────────────────────────────────────────────────

function _diff_random_sheet(rng)::Vector{UInt8}
    sst = _diff_base_sst()
    header_row = rand(rng, 1:3)
    xml = IOBuffer()
    print(xml, _diff_header(header_row))
    for r in (header_row+1):(header_row+rand(rng, 0:12))
        state = rand(rng, (:absent, :absent, :norow_cells, :styled_empty, :outside_only,
                           :empty_strings, :missing_strings, :formula_nocache, :data, :data, :data, :sparse))
        if state === :sparse      # a random subset of B:D, mixed kinds
            cells = String[]
            for col in ("B", "C", "D")
                rand(rng) < 0.5 && continue
                push!(cells, rand(rng, (
                    "<c r=\"$col$r\"><v>$(rand(rng, -5:5))</v></c>",
                    "<c r=\"$col$r\"><v>$(rand(rng) * 10)</v></c>",
                    "<c r=\"$col$r\" s=\"1\"><v>$(45000 + rand(rng, 0:9))</v></c>",
                    "<c r=\"$col$r\" t=\"b\"><v>$(rand(rng, 0:1))</v></c>",
                    "<c r=\"$col$r\" t=\"inlineStr\"><is><t>i$r</t></is></c>",
                    "<c r=\"$col$r\" t=\"s\"><v>$(rand(rng, 3:4))</v></c>",
                    "<c r=\"$col$r\" t=\"e\"><v>#DIV/0!</v></c>",
                    "<c r=\"$col$r\" s=\"2\"/>")))
            end
            print(xml, isempty(cells) ? "" : _diff_row(r, join(cells)))
        else
            print(xml, _diff_state_row(state, r, sst))
        end
    end
    return _diff_build_xlsx(String(take!(xml)), sst)
end

# ── Tests ─────────────────────────────────────────────────────────────────────

@testset "Differential: readtable vs gettable" begin
    cover = _diff_pairwise(_DIFF_SECONDARY_LEVELS)

    @testset "pairwise cover is complete" begin
        levels = _DIFF_SECONDARY_LEVELS
        for i in eachindex(levels), j in (i+1):length(levels), a in 1:levels[i], b in 1:levels[j]
            @test any(r -> r[i] == a && r[j] == b, cover)
        end
    end

    # Every check runs readtable on its value rows (the default, #462) and on the
    # worksheet cache against one oracle, so neither path can drift from `gettable`.
    @testset "empty-row corpus" begin
        _DIFF_CHECKS[] = 0; t0 = time()
        mismatches = String[]
        n = 0
        for len in 0:3, states in Iterators.product(ntuple(_ -> _DIFF_ROW_STATES, len)...)
            n += 1
            bytes = _diff_corpus_sheet(collect(states))
            append!(mismatches, _diff_product("corpus $(collect(states))", bytes, 1,
                                              (nothing, 1, 2, 3, 4, 5), _DIFF_COLUMNS, cover, n))
        end
        @test n == 585
        for (name, bytes) in _diff_special_sheets()
            n += 1
            append!(mismatches, _diff_product("special \"$name\"", bytes, 1,
                                              (nothing, 1, 2, 3, 5, 6, 7, 13), _DIFF_COLUMNS, cover, n;
                                              malformed = name in ("out-of-order rows", "duplicated row")))
        end
        _diff_report(mismatches, "empty-row corpus", t0)
        _diff_test_paths(mismatches)
    end

    @testset "randomised sheets" begin
        _DIFF_CHECKS[] = 0; t0 = time()
        mismatches = String[]
        for seed in 1:500
            rng = Random.MersenneTwister(seed)
            bytes = _diff_random_sheet(rng)
            for _ in 1:8
                cols = rand(rng, _DIFF_COLUMNS)
                kw = (; header = rand(rng, Bool), stop_in_empty_row = rand(rng, Bool),
                        keep_empty_rows = rand(rng, Bool), first_row = rand(rng, (nothing, 1:8...)))
                sec = rand(rng, cover)
                xf = XLSX.readxlsx(IOBuffer(bytes))
                for m in _diff_check(bytes, xf, 1, cols, kw, sec)
                    push!(mismatches, "seed $seed cols=$(repr(cols)) $kw sec=$sec: $m")
                end
            end
        end
        _diff_report(mismatches, "randomised", t0)
        _diff_test_paths(mismatches)
    end

    # Each check reopens the workbook, so these get a pairwise cover of the
    # row-deciding keywords (the generated sheets above get their full product).
    @testset "test/data workbooks" begin
        _DIFF_CHECKS[] = 0; t0 = time()
        mismatches = String[]
        n = 0
        for file in sort(readdir(data_directory))
            any(ext -> endswith(lowercase(file), ext), (".xlsx", ".xlsm", ".xltx", ".xltm")) || continue
            bytes = read(joinpath(data_directory, file))
            xf = _diff_run(() -> XLSX.readxlsx(IOBuffer(bytes)))
            sheets = xf isa Exception ? [1] : collect(1:XLSX.sheetcount(xf))
            for sheet in sheets
                n += 1
                dim = xf isa Exception ? nothing : _diff_run(() -> XLSX.get_dimension(xf[sheet]))
                columns = Any[nothing]
                first_rows = Any[nothing, 1, 2]
                if dim isa XLSX.CellRange
                    c1, c2 = XLSX.column_number(dim.start), XLSX.column_number(dim.stop)
                    r2 = XLSX.row_number(dim.stop)
                    col(i) = XLSX.encode_column_number(i)
                    append!(columns, ["$(col(c1)):$(col(c2))", "$(col(c1)):$(col(c1))",
                                      "$(col(c1)):$(col(c2 + 2))"])
                    append!(first_rows, [r2, r2 + 1])
                end
                append!(mismatches, _diff_product("$file sheet $sheet", bytes, sheet,
                                                  first_rows, columns, cover, n; pairwise = true))
            end
        end
        _diff_report(mismatches, "test/data", t0)
        _diff_test_paths(mismatches)
    end

    fixtures = get(ENV, "XLSX_DIFF_FIXTURES", "")
    if !isempty(fixtures)
        @testset "bench fixtures ($fixtures)" begin
            _DIFF_CHECKS[] = 0; t0 = time()
            mismatches = String[]
            for file in sort(readdir(fixtures))
                startswith(file, "xl_") || continue
                bytes = read(joinpath(fixtures, file))
                xf = XLSX.readxlsx(IOBuffer(bytes))
                for (i, sec) in enumerate(cover)
                    kw = (; header = true, stop_in_empty_row = isodd(i), keep_empty_rows = i % 4 < 2, first_row = 5)
                    for m in _diff_check(bytes, xf, 2, "A:CF", kw, sec; check_rows = false)
                        push!(mismatches, "$file $kw sec=$sec: $m")
                    end
                end
            end
            _diff_report(mismatches, "bench fixtures", t0)
            _diff_test_paths(mismatches)
        end
    end
end

# ── Value rows vs cache rows, row by row (#462) ───────────────────────────────

# Rows of `bytes`' sheet `sheet` from a `ValueRowIterator` (decoding `cols`, or all
# columns) and from the worksheet cache, or the exceptions they threw.
function _both_row_sources(bytes, sheet, cols)
    value_rows = _diff_run() do
        xf = XLSX.open_or_read_xlsx(IOBuffer(bytes), true, true, false; split_sheets = false)
        ws = xf[sheet]
        raw = xf.data[XLSX.get_relationship_target_by_id("xl", XLSX.get_workbook(xf), ws.relationship_id)]
        collect(XLSX.ValueRowIterator(ws, raw, cols))
    end
    cache_rows = _diff_run(() -> collect(XLSX.eachrow(XLSX.readxlsx(IOBuffer(bytes))[sheet])))
    return value_rows, cache_rows
end

# First difference between value rows and cache rows, or `nothing`.
function _compare_row_sources(value_rows, cache_rows, cols)
    if value_rows isa Exception || cache_rows isa Exception
        value_rows isa XLSX._RowsNotAscending && return nothing   # readtable falls back to the cache
        return _diff_compare(cache_rows, value_rows)
    end
    length(value_rows) == length(cache_rows) || return "$(length(value_rows)) value rows vs $(length(cache_rows)) cache rows"
    for (v, c) in zip(value_rows, cache_rows)
        XLSX.row_number(v) == XLSX.row_number(c) || return "row numbers $(XLSX.row_number(v)) vs $(XLSX.row_number(c))"
        r = XLSX.row_number(c)
        isempty(v) == isempty(c) || return "row $r: isempty $(isempty(v)) vs $(isempty(c))"
        ws = XLSX.get_worksheet(c)
        check = if isnothing(cols)
            present = XLSX._present_columns(c)
            XLSX._present_columns(v) == present || return "row $r: present columns $(XLSX._present_columns(v)) vs $present"
            vcat(present, isempty(present) ? [1] : [maximum(present) + 1])   # plus an absent column
        else
            collect(cols)
        end
        for col in check
            a, b = XLSX._row_value(ws, v, col), XLSX._row_value(ws, c, col)
            (isequal(a, b) && typeof(a) == typeof(b)) || return "row $r col $col: $(repr(a)) vs $(repr(b))"
            la, lb = String[], String[]
            XLSX._push_header_label!(la, ws, v, col)
            XLSX._push_header_label!(lb, ws, c, col)
            la == lb || return "row $r col $col: header label $la vs $lb"
        end
    end
    return nothing
end

@testset "value rows match cache rows" begin
    sheets = Tuple{String,Vector{UInt8},Int}[]
    for len in 0:3, states in Iterators.product(ntuple(_ -> _DIFF_ROW_STATES, len)...)
        push!(sheets, ("corpus $(collect(states))", _diff_corpus_sheet(collect(states)), 1))
    end
    append!(sheets, [("special \"$k\"", v, 1) for (k, v) in _diff_special_sheets()])
    append!(sheets, [("random seed $s", _diff_random_sheet(Random.MersenneTwister(s)), 1) for s in 1:200])
    for file in sort(readdir(data_directory))
        any(ext -> endswith(lowercase(file), ext), (".xlsx", ".xlsm", ".xltx", ".xltm")) || continue
        bytes = read(joinpath(data_directory, file))
        xf = try XLSX.readxlsx(IOBuffer(bytes)) catch; continue end
        wb = XLSX.get_workbook(xf)
        for (i, ws) in enumerate(wb.sheets)
            XLSX.is_chartsheet(wb, ws.name) || push!(sheets, ("$file sheet $i", bytes, i))
        end
    end
    fixtures = get(ENV, "XLSX_DIFF_FIXTURES", "")
    if !isempty(fixtures)
        for file in sort(readdir(fixtures))
            startswith(file, "xl_") && push!(sheets, ("$file", read(joinpath(fixtures, file)), 2))
        end
    end

    mismatches = String[]
    for (label, bytes, sheet) in sheets, cols in (nothing, 2:4, 1:7, 3:3)
        m = _compare_row_sources(_both_row_sources(bytes, sheet, cols)..., cols)
        m === nothing || push!(mismatches, "$label cols=$cols: $m")
    end
    foreach(println, first(mismatches, 20))
    @test isempty(mismatches)
    @test length(sheets) > 900

    # the malformed sheets make value rows hand over to the cache
    for (name, bytes) in _diff_special_sheets()
        name in ("out-of-order rows", "duplicated row") || continue
        @test _both_row_sources(bytes, 1, nothing)[1] isa XLSX._RowsNotAscending
    end
end

@testset "value rows: invalid cells outside the read area" begin
    # By design (src/valuerows.jl), readtable decodes only the cells it returns, so an
    # invalid cell elsewhere raises no error; the cache path decodes every cell.
    header = "<row r=\"1\"><c r=\"B1\" t=\"inlineStr\"><is><t>a</t></is></c><c r=\"C1\" t=\"inlineStr\"><is><t>b</t></is></c></row>"
    good(r) = "<row r=\"$r\"><c r=\"B$r\"><v>$r</v></c><c r=\"C$r\"><v>$(10r)</v></c></row>"
    bad_bool(ref) = "<c r=\"$ref\" t=\"b\"><v>2</v></c>"
    outside_column = _diff_build_xlsx(header * good(2) * "<row r=\"3\"><c r=\"B3\"><v>3</v></c><c r=\"C3\"><v>30</v></c>$(bad_bool("E3"))</row>", String[])
    # the table ends at the gap before row 6; row 6 is read to find that end, row 7 never is
    after_table = _diff_build_xlsx(header * good(2) * good(3) * good(6) * "<row r=\"7\">$(bad_bool("B7"))</row>", String[])
    inside = _diff_build_xlsx(header * good(2) * "<row r=\"3\">$(bad_bool("B3"))<c r=\"C3\"><v>30</v></c></row>", String[])

    for (label, bytes) in ("outside the columns" => outside_column, "after the table" => after_table)
        t = XLSX.readtable(IOBuffer(bytes), 1, "B:C")
        @test t.column_labels == [:a, :b]
        @test t.data == Any[[2, 3], [20, 30]]
        XLSX._READTABLE_VALUE_ROWS[] = false
        try
            @test_throws XLSX.XLSXError XLSX.readtable(IOBuffer(bytes), 1, "B:C")
        finally
            XLSX._READTABLE_VALUE_ROWS[] = true
        end
    end
    # inside the read area: the same error on both paths
    e1 = try XLSX.readtable(IOBuffer(inside), 1, "B:C") catch e; e end
    XLSX._READTABLE_VALUE_ROWS[] = false
    e2 = try
        try XLSX.readtable(IOBuffer(inside), 1, "B:C") catch e; e end
    finally
        XLSX._READTABLE_VALUE_ROWS[] = true
    end
    @test e1 isa XLSX.XLSXError && e2 isa XLSX.XLSXError && e1.msg == e2.msg
end

# ── Filtered cell kernel vs the node decoder ─────────────────────────────────

# The cells, formulas and bounds of a sheet as read before `_read_cells` (each
# selected `<c>` decoded with `Cell(::LazyNode, …)` and then skipped), and whether
# its rows are in ascending order.
function _node_read_cells(ws, rows, cols)
    sel(s, x) = isnothing(s) || x in s
    wb = XLSX.get_workbook(ws)
    local_formulas = Dict{XLSX.SheetCellRef,XLSX.AbstractFormula}()
    cells = XLSX.Cell[]
    row_min = col_min = typemax(Int)
    row_max = col_max = 0
    in_rows = false
    ascending, last_row = true, 0
    c = XML.Cursor(XLSX._open_sheetdata(ws))
    while XML.next!(c) !== nothing
        d = XML.depth(c)
        d == 1 && continue
        if d == 2 && XML.nodetype(c) == XML.Element && XLSX.localname(c) == "row"
            row = parse(Int, XML.get(c, "r", nothing))
            ascending &= row > last_row; last_row = row
            row_min = min(row_min, row); row_max = max(row_max, row)
            in_rows = sel(rows, row)
        elseif d == 3 && XML.nodetype(c) == XML.Element && XLSX.localname(c) == "c"
            col = XLSX._ref_column_number(XML.get(c, "r", ""))
            col_min = min(col_min, col); col_max = max(col_max, col)
            if in_rows && sel(cols, col)
                push!(cells, XLSX.Cell(XML.LazyNode(c), ws, XLSX.get_sst_prefix(ws), local_formulas,
                                       XLSX.get_xlsxfile(ws).load_formulas))
            end
            XML.skip_element!(c)
        else
            XML.skip_element!(c)
        end
    end
    bounds = col_max > 0 ? XLSX.CellRange(XLSX.CellRef(row_min, col_min), XLSX.CellRef(row_max, col_max)) : nothing
    return cells, local_formulas, bounds, ascending
end

_formulas_of(ws) = filter(p -> first(p).sheet == ws.name, XLSX.get_workbook(ws).formulas)

# Cells the node decoder reads unusually: CDATA and repeated `<v>`, a second `<is>`,
# an `<f>` before `<is>`, comments and unknown children, empty and self-closing cells
# (also last in a row and last in the sheet), and a shared formula.
function _kernel_odd_cells_sheet(; pretty::Bool)
    rows = [
        "<row r=\"1\"><c r=\"A1\"><v><![CDATA[5]]></v></c><c r=\"B1\"><v>1</v><v>2</v></c>" *
            "<c r=\"C1\" t=\"inlineStr\"><is><t>a</t></is><is><t>b</t></is></c>" *
            "<c r=\"D1\" t=\"inlineStr\"><f>1</f><is><t>x</t></is></c>" *
            "<c r=\"E1\"><!-- note --><v>7</v><extLst><ext uri=\"u\"><y/></ext></extLst></c>" *
            "<c r=\"F1\" t=\"s\" s=\"1\" cm=\"1\"><v>0</v></c><c r=\"G1\" s=\"2\"/></row>",
        "<row r=\"2\"/>",
        "<row r=\"3\"><c r=\"B3\"><f t=\"shared\" ref=\"B3:B4\" si=\"0\">A1*2</f><v>10</v></c><c r=\"C3\"><v></v></c></row>",
        "<row r=\"4\"><c r=\"B4\"><f t=\"shared\" si=\"0\"/><v>4</v></c><c r=\"AB4\" t=\"str\"><f>\"q\"</f><v>q</v></c></row>",
        "<row r=\"6\"><c r=\"C6\"/></row>",
    ]
    xml = join(rows)
    # whitespace between rows, cells and a cell's children (never inside a value)
    pretty && (xml = replace(xml, (t => "\n  " * t for t in ("<row ", "</row>", "<c ", "<v>", "<f", "<is>", "<!--", "<extLst>"))...))
    return _diff_build_xlsx(xml, ["s0"])
end

# `(label, workbook bytes, sheet index)` of every sheet the kernel is checked on: the
# empty-row corpus, special and random sheets, the odd cells and the test/data worksheets.
# Built once and shared by the differentials below.
const _KERNEL_SHEETS = Tuple{String,Vector{UInt8},Int}[]
function _kernel_sheets()
    isempty(_KERNEL_SHEETS) || return _KERNEL_SHEETS
    sheets = _KERNEL_SHEETS
    for len in 0:2, states in Iterators.product(ntuple(_ -> _DIFF_ROW_STATES, len)...)
        push!(sheets, ("corpus $(collect(states))", _diff_corpus_sheet(collect(states)), 1))
    end
    append!(sheets, [("special \"$k\"", v, 1) for (k, v) in _diff_special_sheets()])
    append!(sheets, [("random seed $s", _diff_random_sheet(Random.MersenneTwister(s)), 1) for s in 1:60])
    push!(sheets, ("odd cells", _kernel_odd_cells_sheet(pretty = false), 1))
    push!(sheets, ("odd cells, pretty-printed", _kernel_odd_cells_sheet(pretty = true), 1))
    for file in sort(readdir(data_directory))
        any(ext -> endswith(lowercase(file), ext), (".xlsx", ".xlsm", ".xltx", ".xltm")) || continue
        bytes = read(joinpath(data_directory, file))
        xf = try XLSX.readxlsx(IOBuffer(bytes)) catch; continue end
        wb = XLSX.get_workbook(xf)
        for (i, ws) in enumerate(wb.sheets)
            XLSX.is_chartsheet(wb, ws.name) || push!(sheets, ("$file sheet $i", bytes, i))
        end
    end
    return sheets
end

"""
The choices from `all` to check on the `k`th (sheet, `load_formulas`) pair labelled
`label`: all of them on a generated sheet, but on a test/data sheet (slow to reopen
for every check) the first `keep` plus `n` others, rotated with `k` so that every
choice still runs on the test/data workbooks. All of them under `XLSX_FULL_DIFF=1`.
"""
function _diff_choices(all, label::String, k::Int; keep::Int = 1, n::Int)
    (_DIFF_FULL || !occursin(r"\.xl[st][xm] sheet \d+$"i, label)) && return collect(all)
    rest = collect(all)[keep+1:end]
    return vcat(collect(all)[1:keep], [rest[mod1((k - 1) * n + i, length(rest))] for i in 1:n])
end

@testset "filtered cell kernel matches the node decoder" begin
    sheets = _kernel_sheets()
    selectors = (nothing, 1:0, 2:4, 3:3, 1:2:9, [1, 3, 7], [2, 28])
    # every (rows, cols) pair; the first, (nothing, nothing), reads the whole sheet
    pairs = vec([(rows, cols) for rows in selectors, cols in selectors])
    mismatches = String[]
    n = k = 0
    for (label, bytes, sheet) in sheets, load_formulas in (true, false)
        openws() = XLSX.open_or_read_xlsx(IOBuffer(bytes), true, false, false; load_formulas)[sheet]
        for (rows, cols) in _diff_choices(pairs, label, k += 1; n = 6)
            n += 1
            tag = "$label lf=$load_formulas rows=$rows cols=$cols"
            ws = openws()
            expected = _diff_run(() -> _node_read_cells(ws, rows, cols))
            ws = openws()
            result = _diff_run(() -> XLSX._read_cells(ws, rows, cols; track_bounds = true))
            if expected isa Exception || result isa Exception
                # an invalid cell in the selection: the same error
                (typeof(expected) == typeof(result) && sprint(showerror, expected) == sprint(showerror, result)) ||
                    push!(mismatches, "$tag: $expected vs $result")
                continue
            end
            want, want_formulas, want_bounds, ascending = expected
            got, bounds = result
            got == want || push!(mismatches, "$tag: cells differ")
            isequal(bounds, want_bounds) || push!(mismatches, "$tag: bounds $bounds vs $want_bounds")
            _formulas_of(ws) == want_formulas || push!(mismatches, "$tag: formulas differ")

            # rows in ascending order: the kernel may stop early without tracking bounds
            ascending || continue
            ws = openws()
            got, bounds = XLSX._read_cells(ws, rows, cols; track_bounds = false)
            got == want || push!(mismatches, "$tag, no bounds: cells differ")
            bounds === nothing || push!(mismatches, "$tag, no bounds: bounds $bounds")
        end
    end
    foreach(println, first(mismatches, 20))
    @test isempty(mismatches)
    @test n > 10_000
end

@testset "filtered cell kernel: early stop and dimension scan" begin
    # Without tracking bounds, nothing after the last selected row is read: an invalid
    # cell there raises no error. With tracking, the pass reaches it.
    bad = "<row r=\"5\"><c r=\"B5\" t=\"b\"><v>2</v></c></row>"
    bytes = _diff_build_xlsx("<row r=\"1\"><c r=\"A1\"><v>1</v></c></row>" * bad, String[])
    ws = XLSX.openxlsx(IOBuffer(bytes); enable_cache = false)[1]
    cells, bounds = XLSX._read_cells(ws, 1:1, nothing; track_bounds = false)
    @test length(cells) == 1 && XLSX.getdata(ws, only(cells)) == 1 && bounds === nothing
    cells, bounds = XLSX._read_cells(ws, 1:1, nothing; track_bounds = true)
    @test length(cells) == 1 && bounds == XLSX.CellRange("A1:B5")
    @test isempty(first(XLSX._read_cells(ws, Int[], nothing; track_bounds = false)))
    @test_throws XLSX.XLSXError XLSX._read_cells(ws, nothing, nothing; track_bounds = false)

    # The odd cells, compact and pretty-printed, decode to the same values.
    for pretty in (false, true)
        odd = XLSX.openxlsx(IOBuffer(_kernel_odd_cells_sheet(; pretty)); enable_cache = false)[1]
        cells, bounds = XLSX._read_cells(odd, nothing, nothing; track_bounds = true)
        @test bounds == XLSX.CellRange("A1:AB6")
        values = Dict(string(c.ref) => XLSX.getdata(odd, c) for c in cells)
        @test isequal(values, Dict("A1" => 5, "B1" => 2, "C1" => "a", "D1" => "x", "E1" => 7, "F1" => "s0",
                                   "G1" => missing, "B3" => 10, "C3" => missing, "B4" => 4, "AB4" => "q", "C6" => missing))
        @test [c.formula for c in cells] == [false, false, false, false, false, false, false, true, false, true, true, false]
        @test sort([string(k.cellref) for k in keys(_formulas_of(odd))]) == ["AB4", "B3", "B4"]
    end

    # `_scan_dimension` decodes nothing, so it reads past the invalid cell.
    @test XLSX._scan_dimension(ws) == XLSX.CellRange("A1:B5")
    @test XLSX._scan_dimension(XLSX.openxlsx(IOBuffer(_diff_build_xlsx("<row r=\"3\"/>", String[]));
                                             enable_cache = false)[1]) === nothing
end

# ── Uncached ranged reads vs the row decoders they replaced ──────────────────

# The uncached reads as they were before `_read_cells`: `getcell` matched the row with
# `match_rows`, `getdata` iterated rows until past the range, and non-contiguous
# selections read one cell at a time. `getcellrange` matched the range's rows with
# `match_rows`, which returned no rows after one absent from the file; the oracle
# iterates rows as `getcellrange` did with an empty cache.
function _old_getcell(ws, r, c)
    sheetrows = XLSX.match_rows(ws, [r])
    length(sheetrows) == 1 && return XLSX.getcell(sheetrows[1], c)
    return XLSX.EmptyCell(XLSX.CellRef(r, c))
end
function _old_getdata(ws, rng::XLSX.CellRange)::Array{Any,2}
    result = Array{Any,2}(undef, size(rng))
    fill!(result, missing)
    top, bottom = XLSX.row_number(rng.start), XLSX.row_number(rng.stop)
    for sheetrow in XLSX.eachrow(ws)
        if top <= sheetrow.row <= bottom
            for column in XLSX.column_number(rng.start):XLSX.column_number(rng.stop)
                cell = XLSX.getcell(sheetrow, column)
                if !isempty(cell)
                    (r, c) = XLSX.relative_cell_position(cell, rng)
                    result[r, c] = XLSX.getdata(ws, cell)
                end
            end
        end
        sheetrow.row > bottom && break
    end
    return result
end
function _old_getcellrange(ws, rng::XLSX.CellRange)::Array{XLSX.AbstractCell,2}
    result = Array{Any,2}(undef, size(rng))
    for ref in rng
        (r, c) = XLSX.relative_cell_position(ref, rng)
        result[r, c] = XLSX.EmptyCell(ref)
    end
    top, bottom = XLSX.row_number(rng.start), XLSX.row_number(rng.stop)
    for sheetrow in XLSX.eachrow(ws)
        if top <= sheetrow.row <= bottom
            for column in XLSX.column_number(rng.start):XLSX.column_number(rng.stop)
                cell = XLSX.getcell(sheetrow, column)
                (r, c) = XLSX.relative_cell_position(cell, rng)
                result[r, c] = cell
            end
        end
        sheetrow.row > bottom && break
    end
    return result
end
_old_getdata(ws, rows, cols) = [XLSX.getdata(ws, _old_getcell(ws, a, b)) for a in rows, b in cols]
_old_getcellrange(ws, rows, cols) = [_old_getcell(ws, a, b) for a in rows, b in cols]

# A cell's contents, independent of the order inline strings were added to the
# workbook's string table (which depends on how many cells a read decoded).
_cell_contents(ws, c::XLSX.Cell) = (c.ref, c.datatype, c.style, c.meta, c.formula, XLSX.getdata(ws, c))
_cell_contents(ws, x) = x
_contents(ws, x) = x isa AbstractArray ? map(c -> _cell_contents(ws, c), x) : _cell_contents(ws, x)

@testset "uncached ranged reads match the row decoders" begin
    ranges = XLSX.CellRange.(["A1:A1", "B2:D4", "A1:G9", "C3:AB6", "E7:H12", "B1:B30"])
    singles = XLSX.CellRef.(["A1", "B3", "C1", "AB4", "G7", "Z99"])
    selections = (([1, 3, 7], [2, 4]), (1:2:9, 2:3), (9:-2:1, [7, 2, 2]), (2, [1, 3]),
                  ([3, 1, 3], 2:4), (2:5, 1:3:7), (Int[], [1]), (1:1:4, 3))
    mismatches = String[]
    n = 0
    # Same result (contents and type), and formulas recorded for exactly the selected
    # cells; an error only where the old path raised one. (The old paths lost formulas:
    # the row iterator never merged those after its last 500-row batch.)
    # Like the other streaming reads, these assume ascending rows: a sheet with
    # out-of-order rows (never written by Excel) may read differently.
    function check(tag, openws, new, old, (rows, cols); malformed = false)
        n += 1
        ws_new, ws_old = openws(), openws()
        got, want = _diff_run(() -> new(ws_new)), _diff_run(() -> old(ws_old))
        if got isa Exception
            want isa Exception || push!(mismatches, "$tag: raised $got")
            return
        end
        want isa Exception && return  # the old path decoded an invalid cell outside the selection
        (malformed || isequal(_contents(ws_new, got), _contents(ws_old, want)) && typeof(got) == typeof(want)) ||
            push!(mismatches, "$tag: $(typeof(got)) $got vs $(typeof(want)) $want")
        _formulas_of(ws_new) == _node_read_cells(openws(), rows, cols)[2] ||
            push!(mismatches, "$tag: formulas differ")
    end
    span(rng) = (XLSX.row_number(rng.start):XLSX.row_number(rng.stop), XLSX.column_number(rng.start):XLSX.column_number(rng.stop))
    k = 0
    for (label, bytes, sheet) in _kernel_sheets(), load_formulas in (true, false)
        openws() = XLSX.open_or_read_xlsx(IOBuffer(bytes), true, false, false; load_formulas)[sheet]
        tag = "$label lf=$load_formulas"
        malformed = label == "special \"out-of-order rows\""
        k += 1
        for rng in _diff_choices(ranges, label, k; keep = 0, n = 2)
            check("$tag getdata $rng", openws, ws -> XLSX.getdata(ws, rng), ws -> _old_getdata(ws, rng), span(rng);
                  malformed)
            check("$tag getcellrange $rng", openws, ws -> XLSX.getcellrange(ws, rng), ws -> _old_getcellrange(ws, rng),
                  span(rng); malformed)
        end
        for ref in _diff_choices(singles, label, k; keep = 0, n = 2)
            r, c = XLSX.row_number(ref), XLSX.column_number(ref)
            check("$tag getcell $ref", openws, ws -> XLSX.getcell(ws, ref), ws -> _old_getcell(ws, r, c), (r:r, c:c);
                  malformed)
        end
        for (rows, cols) in _diff_choices(selections, label, k; keep = 0, n = 2)
            check("$tag getdata $rows×$cols", openws, ws -> XLSX.getdata(ws, rows, cols),
                  ws -> _old_getdata(ws, rows, cols), (rows, cols); malformed)
            check("$tag getcellrange $rows×$cols", openws, ws -> XLSX.getcellrange(ws, rows, cols),
                  ws -> _old_getcellrange(ws, rows, cols), (rows, cols); malformed)
        end
    end
    foreach(println, first(mismatches, 20))
    @test isempty(mismatches)
    @test n > 10_000
end

@testset "uncached ranged reads: early stop and cached lookups" begin
    # Only the selected cells are decoded: an invalid cell outside them raises no error.
    bad = "<row r=\"5\"><c r=\"B5\" t=\"b\"><v>2</v></c></row>"
    rows = "<row r=\"1\"><c r=\"A1\"><v>1</v></c><c r=\"C1\" t=\"b\"><v>2</v></c></row>" *
           "<row r=\"2\"><c r=\"A2\"><v>3</v></c></row>"
    ws = XLSX.openxlsx(IOBuffer(_diff_build_xlsx(rows * bad, String[])); enable_cache = false)[1]
    @test XLSX.getcell(ws, "A2").ref == XLSX.CellRef("A2") && XLSX.getdata(ws, XLSX.getcell(ws, "A2")) == 3
    @test XLSX.getcell(ws, "B2") == XLSX.EmptyCell(XLSX.CellRef("B2"))
    @test XLSX.getcell(ws, "A3") == XLSX.EmptyCell(XLSX.CellRef("A3"))
    @test XLSX.getcellrange(ws, "A1:A2") == [XLSX.getcell(ws, "A1"); XLSX.getcell(ws, "A2");;]
    @test isequal(XLSX.getdata(ws, [2, 1], 1:1), Any[3; 1;;])
    @test isequal(XLSX.getcellrange(ws, [2], [1, 2]), [XLSX.getcell(ws, "A2") XLSX.EmptyCell(XLSX.CellRef("B2"))])
    @test_throws XLSX.XLSXError XLSX.getcell(ws, "C1")
    @test_throws XLSX.XLSXError XLSX.getcell(ws, [1, 2], 1:3)
    # Rows after one absent from the file are read (`match_rows` lost them).
    gap = XLSX.openxlsx(joinpath(data_directory, "two_tables.xlsx"); enable_cache = false)[2]
    @test XLSX.getcellrange(gap, "A1:C3")[3, :] == [XLSX.getcell(gap, "A3"), XLSX.getcell(gap, "B3"), XLSX.getcell(gap, "C3")]
    @test all(c -> c isa XLSX.Cell, XLSX.getcellrange(gap, "A3:C3"))

    # The range read stops after its last row, so it leaves an unknown dimension unknown.
    @test isnothing(ws.dimension)
    @test isequal(XLSX.getdata(ws, "A1:A2"), Any[1; 3;;])
    @test isequal(XLSX.getdata(ws, "A2:B3"), Any[3 missing; missing missing])
    @test isnothing(ws.dimension)

    # With the cache enabled, non-contiguous selections are looked up cell by cell, as before.
    for file in ("general.xlsx", "customXml.xlsx")
        path = joinpath(data_directory, file)
        cached = XLSX.readxlsx(path)[1]
        uncached = XLSX.openxlsx(path; enable_cache = false)[1]
        rows, cols = [3, 1, 2], 1:2:5
        @test isequal(XLSX.getdata(cached, rows, cols), XLSX.getdata(uncached, rows, cols))
        @test XLSX.getcellrange(cached, rows, cols) == XLSX.getcellrange(uncached, rows, cols)
    end
end

# ── Single-cursor row stream vs the node walk it replaced ────────────────────

# The rows of a sheet as the stream iterator read them before (each `<row>` node's
# `r`/`ht` attributes and `Cell(::LazyNode, …)` per `<c>`, the last of a repeated column),
# and the number of shared-string cells.
function _old_stream_rows(ws)
    rows = Tuple{Int,Union{Nothing,Float64},Dict{Int,XLSX.Cell}}[]
    sst = 0
    for row in XML.eachchildnode(XLSX._open_sheetdata(ws))
        (XML.nodetype(row) == XML.Element && XLSX.localname(row) == "row") || continue
        r, ht = nothing, nothing
        for (k, v) in XML.eachattribute(row)
            k == "r" && (r = parse(Int, v))
            k == "ht" && (ht = parse(Float64, v))
        end
        isnothing(r) && throw(XLSX.XLSXError("Row without 'r' attribute in worksheet $(ws.name)."))
        cells = Dict{Int,XLSX.Cell}()
        for c in XML.eachchildnode(row)
            (XML.nodetype(c) == XML.Element && XLSX.localname(c) == "c") || continue
            cell = XLSX.Cell(c, ws, XLSX.get_sst_prefix(ws), Dict{XLSX.SheetCellRef,XLSX.AbstractFormula}(), false)
            sst += cell.datatype == XLSX.CT_STRING
            cells[XLSX.column_number(cell)] = cell
        end
        push!(rows, (r, ht, cells))
    end
    return rows, sst
end

# Rows by contents (see `_cell_contents`): a targeted read numbers inline strings differently.
_row_contents(ws, rows) = [(r, ht, Dict(k => _cell_contents(ws, v) for (k, v) in cells)) for (r, ht, cells) in rows]

# The rows the stream iterator yields, copied (it reuses one `Dict` per pass); stops
# after `limit` rows.
function _stream_rows(ws; limit = typemax(Int))
    rows = Tuple{Int,Union{Nothing,Float64},Dict{Int,XLSX.Cell}}[]
    for r in XLSX.eachrow(ws)
        push!(rows, (XLSX.row_number(r), r.ht, copy(r.rowcells)))
        length(rows) >= limit && break
    end
    return rows
end

@testset "row stream matches the node walk" begin
    mismatches = String[]
    n = k = 0
    for (label, bytes, sheet) in _kernel_sheets(), load_formulas in (true, false)
        openws() = XLSX.open_or_read_xlsx(IOBuffer(bytes), true, false, false; load_formulas)[sheet]
        tag = "$label lf=$load_formulas"
        k += 1
        malformed = label in ("special \"out-of-order rows\"", "special \"duplicated row\"")
        n += 1
        oldws = openws()
        expected = _diff_run(() -> _old_stream_rows(oldws))
        ws = openws()
        sst0, dim0 = ws.sst_count, ws.dimension
        got = _diff_run(() -> _stream_rows(ws))
        if expected isa Exception || got isa Exception
            (typeof(expected) == typeof(got) && sprint(showerror, expected) == sprint(showerror, got)) ||
                push!(mismatches, "$tag: $expected vs $got")
            continue
        end
        want, sst = expected
        got == want || push!(mismatches, "$tag: rows differ")
        ws.sst_count == sst0 + sst || push!(mismatches, "$tag: sst_count $(ws.sst_count) vs $(sst0 + sst)")
        _, node_formulas, bounds, _ = _node_read_cells(openws(), nothing, nothing)
        isequal(ws.dimension, something(dim0, bounds, Some(nothing))) ||
            push!(mismatches, "$tag: dimension $(ws.dimension) vs $(something(dim0, bounds, Some(nothing)))")
        # every formula, which the old iterator lost after its last 500-row batch
        _formulas_of(ws) == node_formulas || push!(mismatches, "$tag: formulas differ")

        # a pass that stops early records the formulas of the rows it read
        if !malformed && length(want) > 1
            ws = openws()
            partial = _stream_rows(ws; limit = 2)
            partial == want[1:2] || push!(mismatches, "$tag, 2 rows: rows differ")
            _formulas_of(ws) == _node_read_cells(openws(), first.(partial), nothing)[2] ||
                push!(mismatches, "$tag, 2 rows: formulas differ")
        end

        # match_rows: the wanted rows present in the file, in ascending order
        malformed && continue
        present = first.(want)
        wanted_sets = ([1], [3, 1, 3], [2, 4, 5, 9], collect(1:12), [10_000], isempty(present) ? Int[] : [last(present)])
        for wanted in _diff_choices(wanted_sets, label, k; keep = 0, n = 2)
            n += 1
            ws = openws()
            matched = [(XLSX.row_number(r), r.ht, r.rowcells) for r in XLSX.match_rows(ws, wanted)]
            isequal(_row_contents(ws, matched), _row_contents(oldws, filter(r -> first(r) in wanted, want))) ||
                push!(mismatches, "$tag match_rows $wanted: rows differ")
            _formulas_of(ws) == _node_read_cells(openws(), first.(matched), nothing)[2] ||
                push!(mismatches, "$tag match_rows $wanted: formulas differ")
        end
    end
    foreach(println, first(mismatches, 20))
    @test isempty(mismatches)
    @test n > (_DIFF_FULL ? 4_000 : 2_500)
end

@testset "row stream: find_row and row heights" begin
    rows = "<row r=\"1\" ht=\"20.5\"><c r=\"A1\"><v>1</v></c></row><row r=\"3\" ht=\"x\"><c r=\"A3\"><v>3</v></c></row>" *
           "<row r=\"4\"><c r=\"A4\"><f>A1*4</f><v>4</v></c></row>"
    ws = XLSX.openxlsx(IOBuffer(_diff_build_xlsx(rows, String[])); enable_cache = false)[1]
    itr = XLSX.eachrow(ws)
    @test XLSX.find_row(itr, 1).ht == 20.5
    @test XLSX.find_row(itr, 4).ht === nothing && XLSX.getdata(ws, XLSX.getcell(XLSX.find_row(itr, 4), 1)) == 4
    @test_throws XLSX.XLSXError XLSX.find_row(itr, 2)
    @test_throws ArgumentError XLSX.find_row(itr, 3)                       # an invalid `ht` on a matched row
    @test XLSX.row_number.(XLSX.match_rows(ws, [1, 4])) == [1, 4]          # ... and not on a skipped one
    @test isempty(XLSX.match_rows(ws, Int[]))
    @test_throws ArgumentError collect(XLSX.eachrow(ws))
    @test haskey(XLSX.get_workbook(ws).formulas, XLSX.SheetCellRef(ws.name, XLSX.CellRef("A4")))
end

@testset "row stream: column window" begin
    mismatches = String[]
    n = k = 0
    for (label, bytes, sheet) in _kernel_sheets(), load_formulas in (true, false)
        openws() = XLSX.open_or_read_xlsx(IOBuffer(bytes), true, false, false; load_formulas)[sheet]
        full_ws = openws()
        full = _diff_run(() -> _stream_rows(full_ws))
        for cols in _diff_choices((1:1, 2:4, 3:3, 1:7, 28:28, 100:120), label, k += 1; keep = 0, n = 2)
            n += 1
            tag = "$label lf=$load_formulas cols=$cols"
            ws = openws()
            sst0 = ws.sst_count
            got = _diff_run(() -> begin
                rows = Tuple{Int,Union{Nothing,Float64},Dict{Int,XLSX.Cell}}[]
                for r in XLSX.SheetRowStreamIterator(ws, cols)
                    push!(rows, (XLSX.row_number(r), r.ht, copy(r.rowcells)))
                end
                rows
            end)
            if got isa Exception
                # an invalid cell inside the window raises as in the full pass
                full isa Exception || push!(mismatches, "$tag: raised $got")
                continue
            end
            full isa Exception && continue  # the full pass decoded an invalid cell outside the window
            want = [(r, ht, filter(p -> first(p) in cols, cells)) for (r, ht, cells) in full]
            isequal(_row_contents(ws, got), _row_contents(full_ws, want)) || push!(mismatches, "$tag: rows differ")
            ws.sst_count == sst0 + count(c -> c.datatype == XLSX.CT_STRING, (c for (_, _, cells) in want for c in values(cells))) ||
                push!(mismatches, "$tag: sst_count")
            isequal(ws.dimension, full_ws.dimension) || push!(mismatches, "$tag: dimension $(ws.dimension) vs $(full_ws.dimension)")
            _formulas_of(ws) == _node_read_cells(openws(), nothing, cols)[2] || push!(mismatches, "$tag: formulas differ")
        end
    end
    foreach(println, first(mismatches, 20))
    @test isempty(mismatches)
    @test n > (_DIFF_FULL ? 3_000 : 2_000)

    # Uncached table reads decode only the table's columns: an invalid cell elsewhere
    # raises no error, one inside does.
    bad = "<row r=\"1\"><c r=\"A1\" t=\"inlineStr\"><is><t>h</t></is></c></row>" *
          "<row r=\"2\"><c r=\"A2\"><v>1</v></c><c r=\"C2\" t=\"b\"><v>2</v></c></row>" *
          "<row r=\"3\"><c r=\"A3\"><v>2</v></c></row>"
    open_bad() = XLSX.openxlsx(IOBuffer(_diff_build_xlsx(bad, String[])); enable_cache = false)[1]
    @test [r[:h] for r in XLSX.eachtablerow(open_bad(), "A:A")] == [1, 2]
    @test XLSX.gettable(open_bad(), "A"; first_row = 1).data == [[1, 2]]
    @test_throws XLSX.XLSXError XLSX.gettable(open_bad(), "A:C")

    # A `<c>` without `r` can't be placed in or out of the window.
    nor = XLSX.openxlsx(IOBuffer(_diff_build_xlsx("<row r=\"1\"><c><v>1</v></c></row>", String[])); enable_cache = false)[1]
    @test_throws XLSX.XLSXError collect(XLSX.SheetRowStreamIterator(nor, 2:3))
end

#=
#462 Stage 4: `iterate(::TableRowIterator)` and `gettable` now share one row-stepping
function, and `gettable` writes values straight into columns. The table-row iterator
and `gettable` as they were before are kept below as the reference: the new code must
give exactly the same rows, values, types and `stop_in_row_function` calls.
=#

# The pre-#462 `iterate(::TableRowIterator)`, verbatim apart from qualified names.
function _old_handle_gap(itr, table_row_index, col_count, expected_row, actual_row, sheet_row, sheet_row_iterator_state)
    itr.stop_in_empty_row && return nothing
    itr.keep_empty_rows || return :skip
    table_row = XLSX.TableRow(table_row_index, itr.index, fill(missing, col_count))
    XLSX._should_stop(itr, table_row) && return nothing
    newstate = XLSX.TableRowIteratorState(table_row_index, expected_row, sheet_row_iterator_state,
                                          actual_row - expected_row - 1, sheet_row)
    return table_row, newstate
end

function _old_return_table_row(itr, table_row_index, actual_row, sheet_row, sheet_row_iterator_state)
    table_row = XLSX.TableRow(table_row_index, itr.index, sheet_row, itr.missing_strings)
    XLSX._should_stop(itr, table_row) && return nothing
    newstate = XLSX.TableRowIteratorState(table_row_index, actual_row, sheet_row_iterator_state, 0, nothing)
    return table_row, newstate
end

function _old_iterate(itr, state)
    table_row_index = state.table_row_index + 1
    col_count = length(XLSX.sheet_column_numbers(itr.index))
    if state.missing_rows > 0
        table_row = XLSX.TableRow(table_row_index, itr.index, fill(missing, col_count))
        XLSX._should_stop(itr, table_row) && return nothing
        newstate = XLSX.TableRowIteratorState(table_row_index, state.sheet_row_index + 1,
            state.sheet_row_iterator_state, state.missing_rows - 1, state.row_pending)
        return table_row, newstate
    end
    local sheet_row, sheet_row_iterator_state
    if !isnothing(state.row_pending)
        sheet_row = state.row_pending
        sheet_row_iterator_state = state.sheet_row_iterator_state
    else
        next = iterate(itr.itr, state.sheet_row_iterator_state)
        isnothing(next) && return nothing
        sheet_row, sheet_row_iterator_state = next
    end
    actual_row = XLSX.row_number(sheet_row)
    expected_row = state.sheet_row_index + 1
    if actual_row > expected_row
        result = _old_handle_gap(itr, table_row_index, col_count, expected_row, actual_row, sheet_row, sheet_row_iterator_state)
        result === :skip || return result
    end
    result = XLSX._skip_empty_rows(itr, sheet_row, sheet_row_iterator_state)
    isnothing(result) && return nothing
    sheet_row, sheet_row_iterator_state = result
    return _old_return_table_row(itr, table_row_index, XLSX.row_number(sheet_row), sheet_row, sheet_row_iterator_state)
end

# All rows the old iterator gives, as (table row, values).
function _old_tablerows(itr)
    out = Tuple{Int,Vector{Any}}[]
    next = isnothing(itr.resume) ? iterate(itr.itr) : itr.resume
    while !isnothing(next) && XLSX.row_number(next[1]) < itr.first_data_row
        next = iterate(itr.itr, next[2])
    end
    isnothing(next) && return out
    sheet_row, sheet_row_state = next
    state = XLSX.TableRowIteratorState(0, itr.first_data_row - 1, sheet_row_state, 0, sheet_row)
    while true
        r = _old_iterate(itr, state)
        isnothing(r) && break
        row, state = r
        push!(out, (row.row, Any[row.cell_values...]))
    end
    return out
end

# The pre-#462 `gettable(::TableRowIterator)`, over the old iterator's rows.
function _old_gettable(itr; infer_eltypes::Bool = true)
    n = XLSX.table_columns_count(itr)
    data = Any[Any[] for _ in 1:n]
    for (_, values) in _old_tablerows(itr), ci in 1:n
        push!(data[ci], values[ci])
    end
    if infer_eltypes
        for i in eachindex(data)
            T = XLSX.infer_eltype(data[i])
            T !== Any && (data[i] = convert(Vector{T}, data[i]))
        end
    end
    return XLSX.DataTable(data, XLSX.get_column_labels(itr))
end

_same_values(a, b) = length(a) == length(b) &&
    all(isequal(x, y) && typeof(x) == typeof(y) for (x, y) in zip(a, b))

# Old vs new for one sheet and keyword set: iteration and `gettable`, with the
# callback calls compared too. Returns a mismatch description or `nothing`.
function _stage4_check(ws, cols, kw::NamedTuple, stop_kind::Symbol, infer_eltypes::Bool)
    mk(log) = isnothing(cols) ?
        XLSX.eachtablerow(ws; kw..., stop_in_row_function = _diff_stop_fn(stop_kind, log)) :
        XLSX.eachtablerow(ws, cols; kw..., stop_in_row_function = _diff_stop_fn(stop_kind, log))
    probe = _diff_run(() -> mk(_DiffCall[]))
    probe isa Exception && return nothing          # construction is unchanged code

    # iteration
    la, lb = _DiffCall[], _DiffCall[]
    old_rows = _diff_run(() -> _old_tablerows(mk(la)))
    new_rows = _diff_run(() -> [(r.row, Any[r.cell_values...]) for r in mk(lb)])
    if old_rows isa Exception || new_rows isa Exception
        (old_rows isa Exception && new_rows isa Exception && typeof(old_rows) == typeof(new_rows)) ||
            return "iteration: $(_diff_describe(old_rows)) vs $(_diff_describe(new_rows))"
    else
        length(old_rows) == length(new_rows) || return "iteration: $(length(old_rows)) vs $(length(new_rows)) rows"
        for (i, (a, b)) in enumerate(zip(old_rows, new_rows))
            (a[1] == b[1] && _same_values(a[2], b[2])) || return "iteration row $i: $a vs $b"
        end
    end
    m = _diff_compare_calls(la, lb)
    m === nothing || return "iteration: " * m

    # gettable
    la, lb = _DiffCall[], _DiffCall[]
    old_t = _diff_run(() -> _old_gettable(mk(la); infer_eltypes))
    new_t = _diff_run(() -> XLSX.gettable(mk(lb); infer_eltypes))
    m = _diff_compare(old_t, new_t)
    m === nothing || return "gettable: " * m
    m = _diff_compare_calls(la, lb)
    m === nothing || return "gettable: " * m
    return nothing
end

@testset "table rows and gettable unchanged (#462 Stage 4)" begin
    stops = (:none, :at2, :col1missing, :always, :never)

    @testset "empty-row corpus" begin
        mismatches = String[]
        n = 0
        sheets = Pair{String,Vector{UInt8}}[]
        for len in 0:3, states in Iterators.product(ntuple(_ -> _DIFF_ROW_STATES, len)...)
            push!(sheets, "corpus $(collect(states))" => _diff_corpus_sheet(collect(states)))
        end
        append!(sheets, ["special \"$k\"" => v for (k, v) in _diff_special_sheets()])
        for (label, bytes) in sheets
            ws = XLSX.readxlsx(IOBuffer(bytes))[1]
            for header in (true, false), stop_in_empty_row in (true, false), keep_empty_rows in (true, false),
                first_row in (nothing, 1, 2, 3, 5), cols in _DIFF_COLUMNS
                n += 1
                kw = (; header, stop_in_empty_row, keep_empty_rows, first_row,
                        missing_strings = _DIFF_MS_OPTIONS[mod1(n, 4)])
                m = _stage4_check(ws, cols, kw, stops[mod1(n, 5)], isodd(n ÷ 5))
                m === nothing || push!(mismatches, "$label cols=$(repr(cols)) $kw: $m")
            end
        end
        foreach(println, first(mismatches, 20))
        @test isempty(mismatches)
        @test n > 100_000
    end

    @testset "test/data workbooks" begin
        mismatches = String[]
        n = 0
        for file in sort(readdir(data_directory))
            any(ext -> endswith(lowercase(file), ext), (".xlsx", ".xlsm", ".xltx", ".xltm")) || continue
            xf = try XLSX.readxlsx(joinpath(data_directory, file)) catch; continue end
            wb = XLSX.get_workbook(xf)
            for ws in wb.sheets
                XLSX.is_chartsheet(wb, ws.name) && continue
                dim = _diff_run(() -> XLSX.get_dimension(ws))
                columns, first_rows = Any[nothing], Any[nothing, 1, 2]
                if dim isa XLSX.CellRange
                    c1, c2 = XLSX.column_number(dim.start), XLSX.column_number(dim.stop)
                    col(i) = XLSX.encode_column_number(i)
                    append!(columns, ["$(col(c1)):$(col(c2))", "$(col(c1)):$(col(c1))", "$(col(c1)):$(col(c2 + 2))"])
                    push!(first_rows, XLSX.row_number(dim.stop))
                end
                for header in (true, false), stop_in_empty_row in (true, false), keep_empty_rows in (true, false),
                    first_row in first_rows, cols in columns
                    n += 1
                    kw = (; header, stop_in_empty_row, keep_empty_rows, first_row,
                            missing_strings = _DIFF_MS_OPTIONS[mod1(n, 4)])
                    m = _stage4_check(ws, cols, kw, stops[mod1(n, 5)], isodd(n ÷ 5))
                    m === nothing || push!(mismatches, "$file $(ws.name) cols=$(repr(cols)) $kw: $m")
                end
            end
        end
        foreach(println, first(mismatches, 20))
        @test isempty(mismatches)
        @test n > 10_000
    end

    fixtures = get(ENV, "XLSX_DIFF_FIXTURES", "")
    if !isempty(fixtures)
        @testset "bench fixtures" begin
            for file in sort(readdir(fixtures))
                startswith(file, "xl_") || continue
                ws = XLSX.readxlsx(joinpath(fixtures, file))["Data"]
                for stop_in_empty_row in (true, false), keep_empty_rows in (true, false)
                    kw = (; first_row = 5, stop_in_empty_row, keep_empty_rows)
                    m = _stage4_check(ws, "A:CF", kw, :none, true)
                    m === nothing || println("$file $kw: $m")
                    @test m === nothing
                end
            end
        end
    end

    @testset "column types" begin
        # cells for B2:B4, and the column eltype `gettable` should infer
        cases = [
            ("<c r=\"B2\"><v>1</v></c>", "<c r=\"B3\"><v>2.5</v></c>", "<c r=\"B4\"><v>3</v></c>") => Float64,
            ("<c r=\"B2\"><v>1</v></c>", "<c r=\"B3\"><v>2</v></c>", "<c r=\"B4\"><v>3</v></c>") => Int64,
            ("<c r=\"B2\" s=\"1\"><v>45000</v></c>", "<c r=\"B3\" s=\"1\"><v>45000.5</v></c>", "<c r=\"B4\" s=\"1\"><v>45001</v></c>") => Dates.DateTime,
            ("<c r=\"B2\" t=\"inlineStr\"><is><t>x</t></is></c>", "<c r=\"B3\"><v>2</v></c>", "<c r=\"B4\"><v>3</v></c>") => Any,
            ("<c r=\"B2\" t=\"b\"><v>1</v></c>", "<c r=\"B3\" s=\"2\"/>", "<c r=\"B4\" t=\"b\"><v>0</v></c>") => Union{Missing, Bool},
            ("<c r=\"B2\" s=\"2\"/>", "<c r=\"B3\" s=\"2\"/>", "<c r=\"B4\"><v>1</v></c>") => Union{Missing, Int64},
        ]
        for ((b2, b3, b4), T) in cases
            rows = "<row r=\"1\"><c r=\"B1\" t=\"inlineStr\"><is><t>h</t></is></c><c r=\"C1\" t=\"inlineStr\"><is><t>k</t></is></c></row>" *
                   "<row r=\"2\">$(b2)<c r=\"C2\"><v>1</v></c></row><row r=\"3\">$(b3)<c r=\"C3\"><v>1</v></c></row>" *
                   "<row r=\"4\">$(b4)<c r=\"C4\"><v>1</v></c></row>"
            ws = XLSX.readxlsx(IOBuffer(_diff_build_xlsx(rows, String[])))[1]
            for infer_eltypes in (true, false)
                @test _stage4_check(ws, "B:C", (;), :none, infer_eltypes) === nothing
            end
            @test eltype(XLSX.gettable(ws, "B:C").data[1]) == T
            @test eltype(XLSX.gettable(ws, "B:C"; infer_eltypes = false).data[1]) == Any
        end
        # all missing, and no data rows at all
        rows = "<row r=\"1\"><c r=\"B1\" t=\"inlineStr\"><is><t>h</t></is></c></row><row r=\"2\"><c r=\"B2\" s=\"2\"/></row><row r=\"3\"><c r=\"C3\"><v>1</v></c></row>"
        ws = XLSX.readxlsx(IOBuffer(_diff_build_xlsx(rows, String[])))[1]
        kw = (; stop_in_empty_row = false, keep_empty_rows = true)
        @test _stage4_check(ws, "B:B", kw, :none, true) === nothing
        @test eltype(XLSX.gettable(ws, "B:B"; kw...).data[1]) == Any
        rows = "<row r=\"1\"><c r=\"B1\" t=\"inlineStr\"><is><t>h</t></is></c></row>"
        ws = XLSX.readxlsx(IOBuffer(_diff_build_xlsx(rows, String[])))[1]
        @test _stage4_check(ws, "B:B", (;), :none, true) === nothing
        @test isempty(XLSX.gettable(ws, "B:B").data[1])
    end

    @testset "TableRow length and collect" begin
        ws = XLSX.readxlsx(joinpath(data_directory, "general.xlsx"))["table"]
        nrows = 0
        for r in XLSX.eachtablerow(ws)
            nrows += 1
            n = length(XLSX.get_column_labels(r))
            @test length(r) == n
            @test isequal(collect(r), Any[r[i] for i in 1:n])
            @test eltype(collect(r)) == XLSX.CellConcreteType
        end
        @test nrows > 0
    end
end
