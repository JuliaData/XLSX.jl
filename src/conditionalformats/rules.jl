#
# -- Rule-based conditional formats: each matches cells by a rule and gives them a
#    differential format (`dxStyle`, or `format`, `font`, `border` and `fill`)
#

# The keywords every rule-based setter takes, after its own: whether later rules are
# skipped once this one applies, and the format given to matching cells.
const CF_FORMAT_KEYWORDS = (:stopIfTrue, :dxStyle, :format, :font, :border, :fill)

# A string keyword's value, given as a string or Symbol.
function _cf_string(k::Symbol, v)
    (isnothing(v) || v isa String) && return v
    (v isa AbstractString || v isa Symbol) && return String(v)
    throw(XLSXError("`$k` must be a string, not $(repr(v))."))
end

# Read a rule setter's keywords from `allkws`. `own` holds the setter's own string
# keywords with their defaults; the format keywords follow. Any other keyword is an error.
function _cf_rule_keywords(allkws, own::NamedTuple)
    valid = (keys(own)..., CF_FORMAT_KEYWORDS...)
    for k in keys(allkws)
        k in valid || throw(XLSXError("Invalid keyword argument: $k. Valid options are: $(join(("`$v`" for v in valid), ", "))."))
    end
    names = keys(own)
    strings = NamedTuple{names}(map(k -> _cf_string(k, get(allkws, k, own[k])), names))
    stopIfTrue = _cf_string(:stopIfTrue, get(allkws, :stopIfTrue, nothing))
    dxStyle = _cf_string(:dxStyle, get(allkws, :dxStyle, nothing))
    format::Union{Nothing,Vector{<:Pair{String}}} = get(allkws, :format, nothing)
    font::Union{Nothing,Vector{<:Pair{String}}} = get(allkws, :font, nothing)
    border::Union{Nothing,Vector{<:Pair{String}}} = get(allkws, :border, nothing)
    fill::Union{Nothing,Vector{<:Pair{String}}} = get(allkws, :fill, nothing)
    return (; strings..., stopIfTrue, dxStyle, format, font, border, fill)
end

# A rule's value or threshold: a number or a cell reference.
function _check_cf_value(name, v)
    isnothing(v) || is_valid_cellname(v) || is_valid_fixed_cellname(v) || !isnothing(tryparse(Float64, v)) ||
        throw(XLSXError("Invalid `$name`: $v. Must be a number or a CellRef."))
    return nothing
end

# The differential format given to the cells a rule matches, added to the workbook.
function _cf_rule_dx(ws::Worksheet, kw)
    wb = get_workbook(ws)
    return Add_Cf_Dx(wb, get_new_dx(wb, get_dx(kw.dxStyle, kw.format, kw.font, kw.border, kw.fill)))
end

# A new `<cfRule>` of `type`, giving matching cells format `dxid`, with the next
# priority. `attrs` follow the priority, and `stopIfTrue` comes last.
function _cf_rule_element(ws::Worksheet, pfx, type, dxid, kw; attrs...)
    cfx = XML.Element("$(pfx)cfRule"; type=type, dxfId=string(dxid.id))
    cfx["priority"] = next_cf_priority!(ws)
    for (k, v) in attrs
        cfx[string(k)] = v
    end
    if !isnothing(kw.stopIfTrue) && kw.stopIfTrue == "true"
        cfx["stopIfTrue"] = "1"
    end
    return cfx
end

_cf_prefix(ws::Worksheet) = (pfx = get_prefix(ws); pfx == "" ? pfx : pfx * ":")

const CELLIS_OPERATORS = ("greaterThan", "greaterThanOrEqual", "lessThan", "lessThanOrEqual", "equal", "notEqual", "between", "notBetween")
# Older spellings XLSX.jl once documented; they aren't OOXML names, so they're
# written as the names they stand for.
const CELLIS_OPERATOR_ALIASES = Dict("greaterEqual" => "greaterThanOrEqual", "lessEqual" => "lessThanOrEqual")

define_cf_address_methods(:setCfCellIs)
function setCfCellIs(ws::Worksheet, rng::CfRange; allkws::Dict{Symbol,Any}=Dict{Symbol,Any}())::Int
    kw = _cf_rule_keywords(allkws, (; operator="greaterThan", value=nothing, value2=nothing))
    _check_cf_range(ws, rng)
    pfx = _cf_prefix(ws)

    _check_cf_value("value", kw.value)
    _check_cf_value("value2", kw.value2)
    operator = get(CELLIS_OPERATOR_ALIASES, kw.operator, kw.operator)
    !isnothing(operator) && operator ∉ CELLIS_OPERATORS &&
        throw(XLSXError("Invalid `operator`: $(kw.operator). Valid options are: $(join(("`$o`" for o in CELLIS_OPERATORS), ", "))"))
    dxid = _cf_rule_dx(ws, kw)

    value = kw.value
    if isnothing(value)
        # The average of the range's numbers, as Excel's AVERAGE takes it: text,
        # booleans and empty cells are ignored.
        nums  = filter(v -> v isa Real && !(v isa Bool), _cf_values(ws, rng))
        value = isempty(nums) ? nothing : string(sum(nums) / length(nums))
    end

    cfx = _cf_rule_element(ws, pfx, "cellIs", dxid, kw)
    cfx["operator"] = operator
    push!(cfx, XML.Element("$(pfx)formula", XML.Text(value)))
    if !isnothing(kw.value2) && operator ∈ ["between", "notBetween"]
        push!(cfx, XML.Element("$(pfx)formula", XML.Text(kw.value2)))
    end

    update_worksheet_cfx!(cfx, ws, rng)
    return 0
end

# Operator => (the operator Excel stores, the rule's formula).
const CONTAINSTEXT_RULES = Dict(
    "containsText"    => ("containsText", "NOT(ISERROR(SEARCH(__txt__,__CR__)))"),
    "notContainsText" => ("notContains",  "ISERROR(SEARCH(__txt__,__CR__))"),
    "beginsWith"      => ("beginsWith",   "LEFT(__CR__,LEN(__txt__))=__txt__"),
    "endsWith"        => ("endsWith",     "RIGHT(__CR__,LEN(__txt__))=__txt__"),
)

# The rule's text operand: a cell reference as given, or literal text as a quoted
# formula string (an embedded `"` doubled).
_cf_text_operand(v::AbstractString) =
    (is_valid_cellname(v) || is_valid_fixed_cellname(v)) ? String(v) : "\"" * replace(v, "\"" => "\"\"") * "\""

define_cf_address_methods(:setCfContainsText)
function setCfContainsText(ws::Worksheet, rng::CfRange; allkws::Dict{Symbol,Any}=Dict{Symbol,Any}())::Int
    kw = _cf_rule_keywords(allkws, (; operator="containsText", value=nothing))
    _check_cf_range(ws, rng)
    pfx = _cf_prefix(ws)

    isnothing(kw.value) && throw(XLSXError("Invalid `value`: $(kw.value). Must contain text or a CellRef."))
    haskey(CONTAINSTEXT_RULES, kw.operator) ||
        throw(XLSXError("Invalid operator: $(kw.operator). Valid options are: `containsText`, `notContainsText`, `beginsWith`, `endsWith`."))
    operator, formula = CONTAINSTEXT_RULES[kw.operator]
    operand = _cf_text_operand(kw.value)
    formula = replace(formula, "__txt__" => operand, "__CR__" => string(_cf_anchor(rng)))
    dxid = _cf_rule_dx(ws, kw)

    cfx = _cf_rule_element(ws, pfx, kw.operator, dxid, kw)
    cfx["operator"] = operator
    startswith(operand, '"') && (cfx["text"] = kw.value)   # Excel stores `text` only for literal text
    push!(cfx, XML.Element("$(pfx)formula", XML.Text(formula)))

    update_worksheet_cfx!(cfx, ws, rng)
    return 0
end

# Operator => (percent, bottom) attributes; "" leaves one out.
const TOP10_RULES = Dict(
    "topN"     => ("",  ""),
    "topN%"    => ("1", ""),
    "bottomN"  => ("",  "1"),
    "bottomN%" => ("1", "1"),
)

define_cf_address_methods(:setCfTop10)
function setCfTop10(ws::Worksheet, rng::CfRange; allkws::Dict{Symbol,Any}=Dict{Symbol,Any}())::Int
    kw = _cf_rule_keywords(allkws, (; operator="topN", value="10"))
    _check_cf_range(ws, rng)
    pfx = _cf_prefix(ws)

    _check_cf_value("value", kw.value)
    haskey(TOP10_RULES, kw.operator) ||
        throw(XLSXError("Invalid operator: $(kw.operator). Valid options are: `topN`, `topN%`, `bottomN`, `bottomN%`."))
    percent, bottom = TOP10_RULES[kw.operator]
    dxid = _cf_rule_dx(ws, kw)

    cfx = _cf_rule_element(ws, pfx, "top10", dxid, kw)
    percent != "" && (cfx["percent"] = percent)
    bottom != "" && (cfx["bottom"] = bottom)
    cfx["rank"] = kw.value

    update_worksheet_cfx!(cfx, ws, rng)
    return 0
end

# Operator => the attributes Excel stores for it on an `aboveAverage` rule.
const ABOVEAVERAGE_RULES = Dict(
    "aboveAverage"   => (;),
    "aboveEqAverage" => (; equalAverage="1"),
    "plus1StdDev"    => (; stdDev="1"),
    "plus2StdDev"    => (; stdDev="2"),
    "plus3StdDev"    => (; stdDev="3"),
    "belowAverage"   => (; aboveAverage="0"),
    "belowEqAverage" => (; aboveAverage="0", equalAverage="1"),
    "minus1StdDev"   => (; aboveAverage="0", stdDev="1"),
    "minus2StdDev"   => (; aboveAverage="0", stdDev="2"),
    "minus3StdDev"   => (; aboveAverage="0", stdDev="3"),
)

define_cf_address_methods(:setCfAboveAverage)
function setCfAboveAverage(ws::Worksheet, rng::CfRange; allkws::Dict{Symbol,Any}=Dict{Symbol,Any}())::Int
    kw = _cf_rule_keywords(allkws, (; operator="aboveAverage"))
    _check_cf_range(ws, rng)
    pfx = _cf_prefix(ws)

    haskey(ABOVEAVERAGE_RULES, kw.operator) ||
        throw(XLSXError("Invalid operator: $(kw.operator). Valid options are: `aboveAverage`, `aboveEqAverage`, `plus1StdDev`, `plus2StdDev`, `plus3StdDev`, `belowAverage`, `belowEqAverage`, `minus1StdDev`, `minus2StdDev`, `minus3StdDev`."))
    dxid = _cf_rule_dx(ws, kw)

    cfx = _cf_rule_element(ws, pfx, "aboveAverage", dxid, kw; ABOVEAVERAGE_RULES[kw.operator]...)

    update_worksheet_cfx!(cfx, ws, rng)
    return 0
end

define_cf_address_methods(:setCfTimePeriod)
function setCfTimePeriod(ws::Worksheet, rng::CfRange; allkws::Dict{Symbol,Any}=Dict{Symbol,Any}())::Int
    kw = _cf_rule_keywords(allkws, (; operator="last7Days"))
    _check_cf_range(ws, rng)
    pfx = _cf_prefix(ws)

    haskey(timeperiods, kw.operator) ||
        throw(XLSXError("Invalid operator: $(kw.operator). Valid options are: `yesterday`, `today`, `tomorrow`, `last7Days`, `lastWeek`, `thisWeek`, `nextWeek`, `lastMonth`, `thisMonth`, `nextMonth`."))
    formula = replace(timeperiods[kw.operator], "__CR__" => string(_cf_anchor(rng)))
    dxid = _cf_rule_dx(ws, kw)

    cfx = _cf_rule_element(ws, pfx, "timePeriod", dxid, kw)
    cfx["timePeriod"] = kw.operator
    push!(cfx, XML.Element("$(pfx)formula", XML.Text(formula)))

    update_worksheet_cfx!(cfx, ws, rng)
    return 0
end

# Operator => the rule's formula; unique and duplicate values need none.
const BLANKERRORUNIQDUP_RULES = Dict(
    "containsBlanks"    => "LEN(TRIM(__CR__))=0",
    "notContainsBlanks" => "LEN(TRIM(__CR__))>0",
    "containsErrors"    => "ISERROR(__CR__)",
    "notContainsErrors" => "NOT(ISERROR(__CR__))",
    "uniqueValues"      => "",
    "duplicateValues"   => "",
)

define_cf_address_methods(:setCfContainsBlankErrorUniqDup)
function setCfContainsBlankErrorUniqDup(ws::Worksheet, rng::CfRange; allkws::Dict{Symbol,Any}=Dict{Symbol,Any}())::Int
    kw = _cf_rule_keywords(allkws, (; operator="containsBlanks"))
    _check_cf_range(ws, rng)
    pfx = _cf_prefix(ws)

    haskey(BLANKERRORUNIQDUP_RULES, kw.operator) ||
        throw(XLSXError("Invalid operator: $(kw.operator). Valid options are: containsBlanks, notContainsBlanks, containsErrors, notContainsErrors, uniqueValues, duplicateValues."))
    formula = replace(BLANKERRORUNIQDUP_RULES[kw.operator], "__CR__" => string(_cf_anchor(rng)))
    dxid = _cf_rule_dx(ws, kw)

    cfx = _cf_rule_element(ws, pfx, kw.operator, dxid, kw)
    formula != "" && push!(cfx, XML.Element("$(pfx)formula", XML.Text(formula)))

    update_worksheet_cfx!(cfx, ws, rng)
    return 0
end

define_cf_address_methods(:setCfFormula)
function setCfFormula(ws::Worksheet, rng::CfRange; allkws::Dict{Symbol,Any}=Dict{Symbol,Any}())::Int
    kw = _cf_rule_keywords(allkws, (; formula=nothing))
    isnothing(kw.formula) && throw(XLSXError("A `formula` must be provided as a keyword argument."))
    _check_cf_range(ws, rng)
    pfx = _cf_prefix(ws)

    dxid = _cf_rule_dx(ws, kw)

    cfx = _cf_rule_element(ws, pfx, "expression", dxid, kw)
    push!(cfx, XML.Element("$(pfx)formula", XML.Text("(" * uppercase_unquoted(kw.formula) * ")")))

    update_worksheet_cfx!(cfx, ws, rng)
    return 0
end
