# maintainability.jl
# Measures a few maintainability signals over src/ and compares them with the stored
# baseline (maintainability_baseline.toml), so a rise in copy-paste or error masking
# shows up as a number rather than in review. Lower is better for every metric.
#
# Usage: julia bench/maintainability.jl            # report; exit 1 if any metric rose
#        julia bench/maintainability.jl --update   # report and store the new baseline
#
# Needs only the standard library, so no environment is required.

using TOML

const SRC      = normpath(joinpath(@__DIR__, "..", "src"))
const BASELINE = joinpath(@__DIR__, "maintainability_baseline.toml")
const WINDOW   = 8   # lines in a duplicated block

source_files() = sort!([joinpath(r, f) for (r, _, fs) in walkdir(SRC) for f in fs if endswith(f, ".jl")])
rel(p) = replace(relpath(p, SRC), '\\' => '/')

# Lines as compared for duplication: trimmed, string and number literals blanked, and
# blank lines, comments and lone closers dropped. On a line defining a method in short
# form, `name(...) = ...`, `name` is blanked wherever it appears, so a family of methods
# copied from one function to the next counts as duplicated even though each names its
# own function. Returns (line number, text) pairs.
#
# Only code is compared. Left out, because blanking their literals makes unrelated
# ones look alike: the text of triple-quoted strings (docstrings), the body of a
# multi-line `const` table (`const T = Dict(` up to its closing `)` in column 1),
# lines of nothing but literals and punctuation, and `include` lines.
function code_lines(lines)
    out = Tuple{Int,String}[]
    in_string = false   # inside a triple-quoted string
    in_table = false    # inside a multi-line `const` table
    for (i, l) in enumerate(lines)
        quotes = count("\"\"\"", l)
        if in_string || quotes > 0
            isodd(quotes) && (in_string = !in_string)
            continue
        end
        if in_table
            occursin(r"^[\)\]\}]", l) && (in_table = false)
            continue
        end
        occursin(r"^const\b.*[\(\[\{]\s*(#.*)?$", l) && (in_table = true; continue)
        s = replace(strip(l), r"\"[^\"]*\"" => "S", r"\b\d+(\.\d+)?\b" => "N")
        (occursin(r"^[SN\s=>,:;()\[\]{}]*$", s) || startswith(s, "include(")) && continue
        m = match(r"^([A-Za-z_][A-Za-z0-9_]*!?)\(.*\)\s*(::[^=]+)?=\s", s)
        if m !== nothing
            s = replace(s, Regex("\\b" * replace(m[1], "!" => "\\!") * "(?![A-Za-z0-9_])") => "F")
        end
        (isempty(s) || startswith(s, "#") || s in ("end", "else", ")", "]", "\"\"\"")) && continue
        push!(out, (i, s))
    end
    return out
end

# Lines lying in a block of `WINDOW` consecutive code lines that also occurs elsewhere.
function duplicated_lines(files)
    seen = Dict{UInt,Vector{Tuple{String,Int}}}()
    for f in files
        cl = code_lines(readlines(f))
        for k in 1:length(cl)-WINDOW+1
            h = hash(join(last.(cl[k:k+WINDOW-1]), "\n"))
            push!(get!(seen, h, Tuple{String,Int}[]), (f, k))
        end
    end
    dup = Dict{String,Set{Int}}()
    for locs in values(seen)
        length(locs) > 1 || continue
        for (f, k) in locs
            cl = code_lines(readlines(f))
            union!(get!(dup, f, Set{Int}()), first.(cl[k:k+WINDOW-1]))
        end
    end
    return dup
end

# Top-level and one-level-nested functions written `function name ... end`, with lengths.
function long_functions(files; limit=100)
    long = Tuple{String,String,Int}[]
    for f in files
        L = readlines(f)
        for (i, l) in enumerate(L)
            m = match(r"^(    )?function ([^\s(]+)\(", l)   # `(` so docstring prose doesn't match
            (m === nothing || endswith(rstrip(l), " end")) && continue
            closer = m[1] === nothing ? "end" : "    end"
            j = findnext(x -> rstrip(x) == closer, L, i + 1)
            j === nothing || j - i <= limit || push!(long, (rel(f), m[2], j - i))
        end
    end
    return sort!(long, by=x -> -x[3])
end

count_matching(files, re) = sum(count(l -> occursin(re, l), readlines(f)) for f in files)

# Commented-out code: a comment holding a statement, with no space after `#` or with
# several (`#foo(x)`, `#    x = 1`). Prose comments have exactly one space, so don't match.
const COMMENTED_CODE = r"^\s*#(?![=!])(?:\S|\s{2,})\s*(?:[A-Za-z_.!]+\(.*\)\s*$|[A-Za-z_.\[\]\"]+ *= [^=]|@[a-z]+\b|for [a-z_(]|if |elseif |else\b|end\b|return\b|throw\()"

# Writability or cache checks that throw, written inline rather than with
# `_require_writable`/`_require_cache`: the flag tested on a line that throws, or on
# the line before one.
function inline_guards(files)
    flag = r"\.(is_writable|use_cache_for_sheet_data)\b|is_writable\("
    n = 0
    for f in files
        L = readlines(f)
        for (i, l) in enumerate(L)
            occursin(flag, l) || continue
            occursin("throw(", l) || (i < length(L) && occursin(r"^\s*throw\(", L[i+1])) || continue
            # the helpers' own checks
            (occursin("function _require_", l) || (i > 1 && occursin("function _require_", L[i-1]))) && continue
            n += 1
        end
    end
    return n
end

function measure()
    files = source_files()
    total = sum(length(readlines(f)) for f in files)
    dup = duplicated_lines(files)
    ndup = sum(length, values(dup); init=0)
    long = long_functions(files)
    sizes = sort!([(rel(f), length(readlines(f))) for f in files], by=x -> -x[2])
    metrics = Dict{String,Int}(
        "duplicated_lines"   => ndup,
        "bare_catch"         => count_matching(files, r"^\s*catch\s*$"),
        "something_wrong"    => count_matching(files, r"Something wrong"),
        "asserts"            => count_matching(files, r"@assert\b"),
        "commented_out_code" => count_matching(files, COMMENTED_CODE),
        "inline_guards"      => inline_guards(files),
        "functions_over_100" => length(long),
        "largest_file_lines" => first(sizes)[2],
    )
    return metrics, (; files, total, dup, long, sizes)
end

function report(metrics, detail, baseline)
    println("src/: ", length(detail.files), " files, ", detail.total, " lines\n")
    rose = String[]
    println(rpad("metric", 22), lpad("now", 7), lpad("baseline", 10))
    for k in sort(collect(keys(metrics)))
        b = isnothing(baseline) ? nothing : get(baseline, k, nothing)
        flag = (!isnothing(b) && metrics[k] > b) ? "  ▲ rose" : ""
        isempty(flag) || push!(rose, k)
        println(rpad(k, 22), lpad(metrics[k], 7), lpad(isnothing(b) ? "-" : string(b), 10), flag)
    end
    pct = round(100 * metrics["duplicated_lines"] / detail.total, digits=1)
    println("\nduplicated lines are $pct% of src/ ($(WINDOW)-line blocks)")
    if !isempty(detail.dup)
        println("most duplicated files:")
        for (f, s) in first(sort(collect(detail.dup), by=x -> -length(x[2])), 5)
            println("  ", rpad(rel(f), 40), length(s))
        end
    end
    if !isempty(detail.long)
        println("functions over 100 lines:")
        for (f, name, n) in detail.long
            println("  ", rpad("$f  $name", 60), n)
        end
    end
    println("largest files:")
    for (f, n) in first(detail.sizes, 5)
        println("  ", rpad(f, 40), n)
    end
    return rose
end

metrics, detail = measure()
baseline = isfile(BASELINE) ? Dict{String,Int}(k => v for (k, v) in TOML.parsefile(BASELINE)) : nothing
rose = report(metrics, detail, baseline)

if "--update" in ARGS
    open(BASELINE, "w") do io
        println(io, "# Written by bench/maintainability.jl --update. Lower is better for every metric.")
        TOML.print(io, metrics; sorted=true)
    end
    println("\nbaseline updated: ", BASELINE)
elseif isnothing(baseline)
    println("\nno baseline yet: run with --update to store one")
elseif !isempty(rose)
    println("\nrose since the baseline: ", join(rose, ", "))
    exit(1)
end
