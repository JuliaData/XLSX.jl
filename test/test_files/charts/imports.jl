# Every name the chart code uses that the core defines must be visible in
# XLSX.Charts, and be the same object there. A missing import otherwise fails
# only when the line that uses it runs, so it can reach a release unnoticed.
#
# The scan is static: it parses src/charts/*.jl and collects every bare symbol.
# A local or field that shares a name with a core global looks the same as a
# missing import, hence the allow-list. If this test fails on a name that is
# such a local, add it to `LOCALS`; otherwise add the import to Charts.jl.

module ChartImportCheck

const EXCLUDE  = Set([:iserror, :geterror])   # core generics Charts only extends
const LOCALS   = Set([:formula, :id])          # locals/fields named like core globals
const OWN_COPY = Set([:eval, :include])        # every module defines its own

function _signame(sig)
    while sig isa Expr && sig.head in (:where, :(::))
        sig = sig.args[1]
    end
    sig isa Expr && sig.head === :call && sig.args[1] isa Symbol ? sig.args[1] : nothing
end

function _typename(e)
    e isa Expr && e.head in (:(<:), :(::)) && (e = e.args[1])
    e isa Expr && e.head === :curly && (e = e.args[1])
    e isa Symbol ? e : nothing
end

function _defs!(out, ex)
    ex isa Expr || return out
    h = ex.head
    add(n) = n isa Symbol && push!(out, n)
    if h === :function && ex.args[1] isa Symbol
        add(ex.args[1]); return out
    elseif h === :function || h === :(=)
        n = _signame(ex.args[1])
        n === nothing || (add(n); return out)
        h === :function && return out
    elseif h === :struct
        add(_typename(ex.args[2])); return out
    elseif h === :abstract
        add(_typename(ex.args[1])); return out
    elseif h === :const && ex.args[1] isa Expr && ex.args[1].head === :(=)
        add(_typename(ex.args[1].args[1])); return out
    end
    foreach(a -> _defs!(out, a), ex.args)
    return out
end

function _bare!(s, ex)
    if ex isa Symbol
        push!(s, ex)
    elseif ex isa Expr
        ex.head === :kw && length(ex.args) == 2 && return _bare!(s, ex.args[2])
        ex.head === :. && length(ex.args) == 2 && ex.args[2] isa QuoteNode &&
            return _bare!(s, ex.args[1])
        foreach(a -> _bare!(s, a), ex.args)
    end
    return s                                                  # QuoteNodes skipped
end

jlfiles(dir) = [joinpath(d, f) for (d, _, fs) in walkdir(dir) for f in fs if endswith(f, ".jl")]
parsed(f)    = Meta.parseall(read(f, String); filename = f)

# Names defined in the chart files and nowhere in the core.
function chart_names(src, chartdir)
    charts = Set{Symbol}()
    foreach(f -> _defs!(charts, parsed(f)), jlfiles(chartdir))
    core = Set{Symbol}()
    for f in jlfiles(src)
        startswith(normpath(f), normpath(chartdir)) || _defs!(core, parsed(f))
    end
    return setdiff(charts, core, EXCLUDE)
end

# (unseen, shadowed), each a Dict of symbol => files using it.
function scan(X::Module)
    src      = joinpath(pkgdir(X), "src")
    chartdir = joinpath(src, "charts")
    C        = X.Charts
    own      = chart_names(src, chartdir)
    unseen   = Dict{Symbol,Set{String}}()
    shadowed = Dict{Symbol,Set{String}}()
    for f in jlfiles(chartdir), s in _bare!(Set{Symbol}(), parsed(f))
        isdefined(X, s) || continue
        if !isdefined(C, s)
            s in LOCALS || push!(get!(Set{String}, unseen, s), basename(f))
        elseif getglobal(C, s) !== getglobal(X, s) && !(s in own)
            s in OWN_COPY || push!(get!(Set{String}, shadowed, s), basename(f))
        end
    end
    return unseen, shadowed
end

end # module ChartImportCheck

# XLSX.Charts sees every core name it uses
@testset "Chart imports" begin
    unseen, shadowed = ChartImportCheck.scan(XLSX)
    # Defined in XLSX but not visible in Charts: a missing import in Charts.jl.
    @test isempty(unseen)
    # Visible in Charts but a different object from XLSX's.
    @test isempty(shadowed)

    @testset "XLSX-level chart names" begin
        for n in (:AbstractChart, :Chart, :ChartEx, :ChartRef, :ChartSeries,
            :getChartSchema, :getChartType, :getCharts, :getChart, :getChartData, :getChartRanges)
            @test getglobal(XLSX, n) === getglobal(XLSX.Charts, n)
            VERSION >= v"1.11" && @test Base.ispublic(XLSX, n)
        end
    end
end
