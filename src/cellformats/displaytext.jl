#
# The text Excel shows for a cell: its value rendered by its number format, as
# Excel's `Range.Text` gives it. Checked against Excel 365's rendering of each case in
# test/data/number_formats_excel.tsv.
#
# A format code has up to four sections separated by `;` (positive; negative; zero;
# text). Each section is parsed into tokens once, by `_parse_numfmt`, which also
# decides whether a code is valid for `setFormat`.
#

struct _FmtToken
    kind::Symbol     # :lit :pad :fill :text :digit :point :comma :percent :exp :general
                     # :date :minute :elapsed :ampm :fracsec
    text::String
end

struct _FmtSection
    tokens::Vector{_FmtToken}
    condition::Union{Nothing,Tuple{String,Float64}}
    isdate::Bool
    hastext::Bool
end

const _FMT_COLORS = ("black", "blue", "cyan", "green", "magenta", "red", "white", "yellow")
const _FMT_CONDITION = r"^(<=|>=|<>|<|>|=)\s*(-?(?:\d+\.?\d*|\.\d+)(?:[eE][+-]?\d+)?)$"
const _FMT_ELAPSED = r"^(?:h+|m+|s+)$"

# A format code's sections: `nothing` if it is malformed (an unclosed quote or bracket,
# a trailing escape, or more than four sections).
function _split_sections(code::AbstractString)::Union{Nothing,Vector{String}}
    parts = String[]
    buf = IOBuffer()
    quoted = bracketed = escaped = false
    for c in code
        if escaped
            escaped = false
        elseif quoted
            c == '"' && (quoted = false)
        elseif bracketed
            c == ']' && (bracketed = false)
        elseif c in ('\\', '!', '_', '*')
            escaped = true
        elseif c == '"'
            quoted = true
        elseif c == '['
            bracketed = true
        elseif c == ';'
            push!(parts, String(take!(buf)))
            continue
        end
        print(buf, c)
    end
    (quoted || bracketed || escaped) && return nothing
    push!(parts, String(take!(buf)))
    return length(parts) > 4 ? nothing : parts
end

# The tokens of one section. With `strict`, a letter that is no format code makes the
# section invalid (`nothing`), as Excel refuses it; otherwise it is shown as it is.
function _tokenize_section(sec::AbstractString, strict::Bool)::Union{Nothing,_FmtSection}
    chars = collect(sec)
    n = length(chars)
    toks = _FmtToken[]
    condition = nothing
    lit(s) = push!(toks, _FmtToken(:lit, string(s)))
    i = 1
    while i <= n
        c = chars[i]
        lc = lowercase(c)
        step = 1                       # characters this token takes
        if c == '"'
            j = findnext(==('"'), chars, i + 1)
            lit(String(chars[i+1:j-1]))
            step = j - i + 1
        elseif c in ('\\', '!')
            lit(chars[i+1])
            step = 2
        elseif c == '_'
            push!(toks, _FmtToken(:pad, string(chars[i+1])))
            step = 2
        elseif c == '*'
            push!(toks, _FmtToken(:fill, string(chars[i+1])))
            step = 2
        elseif c == '['
            j = findnext(==(']'), chars, i + 1)
            b = String(chars[i+1:j-1])
            step = j - i + 1
            lb = lowercase(b)
            m = match(_FMT_CONDITION, b)
            if occursin(_FMT_ELAPSED, lb)
                push!(toks, _FmtToken(:elapsed, lb))
            elseif !isnothing(m)
                condition = (String(m.captures[1]), parse(Float64, m.captures[2]))
            elseif startswith(b, '$')
                symbol = first(split(b[2:end], '-'))     # [$€-x-euro2] shows €; [$-409] nothing
                isempty(symbol) || lit(symbol)
            elseif strict && !(lb in _FMT_COLORS || occursin(r"^(color\s*\d+|dbnum\d|natnum\d+)$", lb))
                return nothing                           # Excel refuses another word, such as [Purple]
            end                                          # colours ([Red], [Color10]) and others show nothing
        elseif c == '@'
            push!(toks, _FmtToken(:text, "@"))
        elseif c in ('0', '#', '?')
            push!(toks, _FmtToken(:digit, string(c)))
        elseif c == '.'
            push!(toks, _FmtToken(:point, "."))
        elseif c == '%'
            push!(toks, _FmtToken(:percent, "%"))
        elseif c == ','
            push!(toks, _FmtToken(:comma, ","))
        elseif lc == 'e' && i < n && chars[i+1] in ('+', '-')
            push!(toks, _FmtToken(:exp, string('E', chars[i+1])))
            step = 2
        elseif lc == 'g' && lowercase(String(chars[i:min(n, i + 6)])) == "general"
            push!(toks, _FmtToken(:general, "General"))
            step = 7
        elseif lc == 'a' && uppercase(String(chars[i:min(n, i + 4)])) == "AM/PM"
            push!(toks, _FmtToken(:ampm, String(chars[i:i+4])))
            step = 5
        elseif lc == 'a' && uppercase(String(chars[i:min(n, i + 2)])) == "A/P"
            push!(toks, _FmtToken(:ampm, String(chars[i:i+2])))
            step = 3
        elseif lc == 'b' && i < n && chars[i+1] in ('1', '2')
            push!(toks, _FmtToken(:date, ""))            # B1 or B2 chooses a calendar, showing nothing
            step = 2
        elseif lc in ('y', 'm', 'd', 'h', 's', 'e', 'g', 'b')
            while i + step <= n && lowercase(chars[i+step]) == lc
                step += 1
            end
            push!(toks, _FmtToken(:date, lc^step))
        else
            lit(c)                                       # Excel shows any other character as it is
        end
        i += step
    end
    isdate = any(t -> t.kind in (:date, :elapsed, :ampm), toks)
    isdate && _mark_time_tokens!(toks)
    # Excel refuses a section mixing date codes with digit placeholders, such as `0 h`
    strict && isdate && any(t -> t.kind == :digit, toks) && return nothing
    return _FmtSection(toks, condition, isdate, any(t -> t.kind == :text, toks))
end

# In a date section, `m` after an hour or before a second is minutes, and a point
# followed by zeros is a fraction of a second.
function _mark_time_tokens!(toks::Vector{_FmtToken})
    timeish = findall(t -> t.kind in (:date, :elapsed), toks)
    for (k, i) in enumerate(timeish)
        t = toks[i]
        (t.kind == :date && t.text in ("m", "mm")) || continue
        before = k > 1 && startswith(toks[timeish[k-1]].text, 'h')
        after = k < length(timeish) && startswith(toks[timeish[k+1]].text, 's')
        (before || after) && (toks[i] = _FmtToken(:minute, t.text))
    end
    i = 1
    while i <= length(toks)
        if toks[i].kind == :point && i < length(toks) && toks[i+1].kind == :digit
            j = i
            while j < length(toks) && toks[j+1].kind == :digit && toks[j+1].text == "0"
                j += 1
            end
            splice!(toks, i:j, [_FmtToken(:fracsec, string(j - i))])
        end
        i += 1
    end
    return toks
end

# A format code's sections, or `nothing` if it is not valid.
function _parse_numfmt(code::AbstractString; strict::Bool=false)::Union{Nothing,Vector{_FmtSection}}
    parts = _split_sections(code)
    isnothing(parts) && return nothing
    sections = _FmtSection[]
    for p in parts
        s = _tokenize_section(p, strict)
        isnothing(s) && return nothing
        push!(sections, s)
    end
    return sections
end

#
# Rendering
#

"""
    _display_text(v, code; date1904=false, for_width=false) -> String

The text Excel shows for `v` (a `Real`, `String`, `Bool`, or an error as its text, such
as `"#N/A"`) under the number format `code`. With `for_width`, padding is given as the
characters whose width it takes (`_x` as `x`, `?` as `0`), for measuring the text.
"""
function _display_text(v, code::AbstractString; date1904::Bool=false, for_width::Bool=false,
                       sections::Union{Nothing,Vector{_FmtSection}}=_parse_numfmt(code))::String
    isnothing(sections) && (sections = _parse_numfmt("General"))
    textsec = _text_section(sections)
    if v isa Bool || v isa AbstractString
        s = v isa Bool ? (v ? "TRUE" : "FALSE") : String(v)
        isnothing(textsec) && return s
        return _render_tokens(textsec.tokens, s, for_width)
    end
    x = Float64(v)
    sec, auto_minus = _number_section(sections, textsec, x)
    isnothing(sec) && return _general(x)
    return _render_section(sec, x, auto_minus, date1904, for_width)
end

# The section text is shown with: a fourth section, or a last section holding `@`.
function _text_section(sections::Vector{_FmtSection})
    length(sections) == 4 && return sections[4]
    s = sections[end]
    return s.hastext && !any(t -> t.kind == :digit, s.tokens) ? s : nothing
end

# The section a number is shown with, and whether a minus sign is added to it.
function _number_section(sections::Vector{_FmtSection}, textsec, x::Float64)
    numeric = isnothing(textsec) ? sections : (length(sections) == 1 ? _FmtSection[] : sections[1:end-1])
    isempty(numeric) && return nothing, true
    if any(s -> !isnothing(s.condition), numeric)
        for s in numeric
            !isnothing(s.condition) && _condition_holds(s.condition, x) && return s, false
        end
        for s in numeric
            isnothing(s.condition) && return s, true
        end
        return numeric[end], true
    end
    length(numeric) == 1 && return numeric[1], true
    x < 0 && return numeric[2], false
    x == 0 && length(numeric) >= 3 && return numeric[3], false
    return numeric[1], true
end

function _condition_holds((op, limit)::Tuple{String,Float64}, x::Float64)
    op == "<"  && return x < limit
    op == "<=" && return x <= limit
    op == ">"  && return x > limit
    op == ">=" && return x >= limit
    op == "="  && return x == limit
    return x != limit   # "<>"
end

function _render_section(sec::_FmtSection, x::Float64, auto_minus::Bool, date1904::Bool, for_width::Bool)::String
    toks = sec.tokens
    sec.isdate && return _render_date(toks, x, date1904, for_width)
    if any(t -> t.kind == :general, toks) || (sec.hastext && !any(t -> t.kind == :digit, toks))
        g = _general(abs(x))
        sign = auto_minus && x < 0 ? "-" : ""
        return sign * _render_tokens(toks, g, for_width; general=true)
    end
    any(t -> t.kind == :digit, toks) || return _render_tokens(toks, "", for_width)
    iexp = findfirst(t -> t.kind == :exp, toks)
    !isnothing(iexp) && return _render_scientific(toks, iexp, x, auto_minus, for_width)
    islash = _fraction_slash(toks)
    !isnothing(islash) && return _render_fraction(toks, islash, x, auto_minus, for_width)
    return _render_decimal(toks, x, auto_minus, for_width)
end

# Literals, padding and the text placeholder, for text or a value already formatted.
function _render_tokens(toks::Vector{_FmtToken}, value::String, for_width::Bool; general::Bool=false)
    io = IOBuffer()
    for t in toks
        if t.kind == :text || (general && t.kind == :general)
            print(io, value)
        elseif t.kind == :lit
            print(io, t.text)
        elseif t.kind == :pad
            print(io, for_width ? t.text : " ")
        elseif t.kind in (:digit, :point, :percent, :comma)
            print(io, t.text)       # a placeholder in a text section is shown as it is
        end                         # a fill is not repeated: there is no column width to fill
    end
    return String(take!(io))
end

#
# Numbers as Excel holds them: 15 significant decimal digits.
#

# `|x| = N × 10^(e - 14)`, with `N` the 15 significant digits of `|x|`.
function _sig15(x::Float64)::Tuple{Int64,Int}
    s = sprint(io -> @printf(io, "%.14e", abs(x)))
    mant, ex = split(s, 'e')
    return parse(Int64, replace(mant, "." => "")), parse(Int, ex)
end

# `N × 10^(e - 14)` rounded half away from zero to `ndec` decimals, as its integer and
# fraction digits.
function _round_digits(N::Int64, e::Int, ndec::Int)::Tuple{String,String}
    p = e - 14 + ndec
    s = if N == 0 || p < -16
        "0"
    elseif p >= 0
        string(N) * "0"^p
    else
        q, r = divrem(N, Int64(10)^(-p))
        string(q + (2r >= Int64(10)^(-p) ? 1 : 0))
    end
    s = lpad(s, ndec + 1, '0')
    i = s[1:end-ndec]
    i = lstrip(i, '0')
    return isempty(i) ? "0" : String(i), s[end-ndec+1:end]
end

_round_fixed(v::Float64, ndec::Int) = v == 0 ? ("0", "0"^ndec) : _round_digits(_sig15(v)..., ndec)

_is_zero_digits(i::AbstractString, f::AbstractString) = all(==('0'), i) && all(==('0'), f)

# Excel's General format: a number in at most 11 characters besides its sign, in
# scientific notation when it doesn't fit, or would show fewer significant digits than
# scientific notation.
function _general(x::Float64)::String
    x == 0 && return "0"
    neg = x < 0
    budget = 11
    N, e = _sig15(x)
    nsig = length(rstrip(string(N), '0'))
    decimal = nothing
    if e >= 0
        if e + 1 <= budget
            i, f = _round_digits(N, e, max(budget - e - 2, 0))
            length(i) <= budget && (decimal = _join_decimal(i, rstrip(f, '0')))
        end
    else
        i, f = _round_digits(N, e, budget - 2)
        if !_is_zero_digits(i, f)
            shown = length(rstrip(lstrip(f, '0'), '0'))
            sci = _general_scientific(N, e, budget)
            decimal = shown >= min(nsig, _sig_shown(sci)) ? _join_decimal(i, rstrip(f, '0')) : nothing
        end
    end
    s = isnothing(decimal) ? _general_scientific(N, e, budget) : decimal
    return neg ? "-" * s : s
end

_join_decimal(i, f) = isempty(f) ? String(i) : string(i, ".", f)
_sig_shown(sci::String) = length(rstrip(replace(first(split(sci, 'E')), "." => ""), '0'))

function _general_scientific(N::Int64, e::Int, budget::Int)::String
    for _ in 1:2   # once more if rounding carries into the next power of ten
        expstr = string(e < 0 ? '-' : '+', lpad(abs(e), 2, '0'))
        ndec = max(budget - 3 - length(expstr), 0)
        i, f = _round_digits(N, 0 + 14 - 14, ndec)   # the mantissa N × 10^-14
        if length(i) > 1
            e += 1
            N = N ÷ 10
            continue
        end
        return string(_join_decimal(i, rstrip(f, '0')), "E", expstr)
    end
    return "0E+00"
end

#
# Decimal numbers: 0, #, ? placeholders, thousands separators, scaling and percent
#

function _number_parts(toks::Vector{_FmtToken}, stop::Int)
    ipoint = findfirst(t -> t.kind == :point, toks[1:stop])
    intend = isnothing(ipoint) ? stop : ipoint - 1
    digits = findall(t -> t.kind == :digit, toks[1:stop])
    lastdigit = isempty(digits) ? 0 : last(digits)
    thousands = false
    scale = 0
    kinds = Symbol[]
    for (k, t) in enumerate(toks[1:stop])
        kind = t.kind
        if kind == :comma
            if k < intend && any(d -> k < d <= intend, digits) && any(<(k), digits)
                thousands = true
                kind = :skip
            elseif k > lastdigit && lastdigit > 0
                scale += 1
                kind = :skip
            else
                kind = :lit
            end
        end
        push!(kinds, kind)
    end
    npercent = count(t -> t.kind == :percent, toks)
    return (; ipoint, intend, thousands, scale, npercent, kinds)
end

function _render_decimal(toks::Vector{_FmtToken}, x::Float64, auto_minus::Bool, for_width::Bool)
    io = IOBuffer()
    _write_decimal!(io, toks, length(toks), abs(x) , x < 0 && auto_minus, for_width)
    return String(take!(io))
end

# The number part of `toks[1:stop]` for `v ≥ 0`, with a minus sign if `minus` and the
# number shown isn't zero. Returns how many integer digits the rounded number has, for
# scientific notation.
function _write_decimal!(io::IO, toks::Vector{_FmtToken}, stop::Int, v::Float64, minus::Bool, for_width::Bool)
    p = _number_parts(toks, stop)
    v *= 100.0^p.npercent / 1000.0^p.scale
    fracph = isnothing(p.ipoint) ? Int[] : [k for k in p.ipoint+1:stop if toks[k].kind == :digit]
    i, f = _round_fixed(v, length(fracph))
    minus && !_is_zero_digits(i, f) && print(io, '-')
    intdigits = i == "0" ? "" : i
    intph = [k for k in 1:p.intend if toks[k].kind == :digit]
    slots = _int_slots(toks, intph, intdigits, p.thousands, for_width)
    fracout = _frac_slots(toks, fracph, f, for_width)
    for k in 1:stop
        kind = p.kinds[k]
        t = toks[k]
        if kind == :digit
            print(io, get(slots, k, get(fracout, k, "")))
        elseif kind == :point
            isempty(intph) && print(io, intdigits)   # no integer placeholders: the digits go here
            print(io, '.')
        elseif kind == :lit
            print(io, t.text)
        elseif kind == :pad
            print(io, for_width ? t.text : " ")
        elseif kind == :percent
            print(io, '%')
        end
    end
    return length(intdigits)
end

# The text of each integer placeholder: digits from the right, the leftmost taking
# any digits left over; `0` pads with a zero, `?` with a space, `#` with nothing.
function _int_slots(toks, intph::Vector{Int}, digits::String, thousands::Bool, for_width::Bool)
    slots = Dict{Int,String}()
    isempty(intph) && return slots
    nph = length(intph)
    if thousands   # group the whole integer, placed at the first placeholder
        minlen = count(k -> toks[k].text == "0", intph)
        s = lpad(digits, minlen, '0')
        groups = String[]
        while length(s) > 3
            pushfirst!(groups, s[end-2:end])
            s = s[1:end-3]
        end
        isempty(s) || pushfirst!(groups, s)
        slots[intph[1]] = join(groups, ",")
        for k in intph[2:end]
            slots[k] = ""
        end
        return slots
    end
    nd = length(digits)
    for (j, k) in enumerate(intph)
        r = nph - j                       # placeholders to its right
        d = nd - r                        # its digit, counted from the left
        s = d >= 1 ? string(digits[d]) : _pad(toks[k].text, for_width)
        j == 1 && nd > nph && (s = digits[1:nd-nph] * s)
        slots[k] = s
    end
    return slots
end

_pad(placeholder::AbstractString, for_width::Bool) =
    placeholder == "0" ? "0" : placeholder == "?" ? (for_width ? "0" : " ") : ""

# The text of each decimal placeholder: trailing zeros are left out by `#` and shown as
# spaces by `?`.
function _frac_slots(toks, fracph::Vector{Int}, f::AbstractString, for_width::Bool)
    slots = Dict{Int,String}()
    trailing = true
    for j in length(fracph):-1:1
        k = fracph[j]
        d = f[j]
        if trailing && d == '0' && toks[k].text != "0"
            slots[k] = _pad(toks[k].text, for_width)
        else
            trailing = false
            slots[k] = string(d)
        end
    end
    return slots
end

#
# Scientific notation
#

function _render_scientific(toks::Vector{_FmtToken}, iexp::Int, x::Float64, auto_minus::Bool, for_width::Bool)
    v = abs(x)
    p = _number_parts(toks, iexp - 1)
    intph = [k for k in 1:p.intend if toks[k].kind == :digit]
    step = length(intph) > 1 && any(k -> toks[k].text == "#", intph) ? length(intph) : 1
    e = v == 0 ? 0 : last(_sig15(v))
    e = fld(e, step) * step
    io = IOBuffer()
    for _ in 1:2   # once more if rounding carries the mantissa past its range
        io = IOBuffer()
        mantissa = v == 0 ? 0.0 : v / exp10(e)
        nint = _write_decimal!(io, toks, iexp - 1, mantissa, x < 0 && auto_minus, for_width)
        nint <= step && break
        e += step
    end
    sign = e < 0 ? "-" : toks[iexp].text == "E+" ? "+" : ""
    expdigits = count(k -> toks[k].kind == :digit, iexp+1:length(toks))
    print(io, 'E', sign, lpad(abs(e), expdigits, '0'))
    for k in iexp+1:length(toks)
        t = toks[k]
        t.kind == :lit && print(io, t.text)
        t.kind == :pad && print(io, for_width ? t.text : " ")
    end
    return String(take!(io))
end

#
# Fractions
#

# The `/` of a fraction: a literal `/` with a digit placeholder before it and a
# placeholder or digit after it.
function _fraction_slash(toks::Vector{_FmtToken})
    for k in 2:length(toks)-1
        t = toks[k]
        t.kind == :lit && t.text == "/" || continue
        toks[k-1].kind == :digit || continue
        nxt = toks[k+1]
        (nxt.kind == :digit || (nxt.kind == :lit && all(isdigit, nxt.text))) && return k
    end
    return nothing
end

function _render_fraction(toks::Vector{_FmtToken}, islash::Int, x::Float64, auto_minus::Bool, for_width::Bool)
    v = abs(x)
    # numerator: the placeholders just before the slash; integer part: those before them
    numstart = islash - 1
    while numstart > 1 && toks[numstart-1].kind == :digit
        numstart -= 1
    end
    intph = [k for k in 1:numstart-1 if toks[k].kind == :digit]
    hasint = !isempty(intph)
    # denominator: placeholders, or a fixed number, after the slash
    denend = islash + 1
    while denend < length(toks) && (toks[denend+1].kind == :digit ||
                                     (toks[denend+1].kind == :lit && all(isdigit, toks[denend+1].text)))
        denend += 1
    end
    dentoks = toks[islash+1:denend]
    fixed = toks[islash+1].kind == :lit
    whole = hasint ? floor(v) : 0.0
    f = v - whole
    num, den = if fixed
        d = parse(Int, join(t.text for t in dentoks))
        round(Int, f * d, RoundNearestTiesAway), d
    else
        _best_fraction(f, 10^length(dentoks) - 1)
    end
    if hasint && num == den
        whole += 1
        num = 0
    end
    io = IOBuffer()
    x < 0 && auto_minus && (whole > 0 || num > 0) && print(io, '-')
    intdigits = whole == 0 ? "" : string(round(Int, whole))
    numph = numstart:islash-1
    if hasint && whole == 0 && num == 0
        intdigits = "0"                   # Excel shows the zero rather than nothing
    end
    slots = _int_slots(toks, intph, intdigits, false, for_width)
    for k in 1:numstart-1
        t = toks[k]
        t.kind == :digit ? print(io, slots[k]) : t.kind == :lit ? print(io, t.text) :
        t.kind == :pad ? print(io, for_width ? t.text : " ") : nothing
    end
    if num == 0 && hasint
        # the fraction is shown as spaces of its width
        print(io, " "^(length(numph) + 1 + length(dentoks)))
    else
        print(io, lpad(string(num), length(numph), for_width ? '0' : ' '))
        print(io, '/')
        print(io, fixed ? string(den) : rpad(string(den), length(dentoks), for_width ? '0' : ' '))
    end
    for k in denend+1:length(toks)
        t = toks[k]
        t.kind == :lit && print(io, t.text)
        t.kind == :pad && print(io, for_width ? t.text : " ")
    end
    return String(take!(io))
end

# The fraction nearest `f` with a denominator up to `maxden`, the smallest such.
function _best_fraction(f::Float64, maxden::Int)::Tuple{Int,Int}
    best = (round(Int, f), 1)
    besterr = abs(f - best[1])
    for d in 2:maxden
        n = round(Int, f * d)
        err = abs(f - n / d)
        if err < besterr - 1e-12
            best, besterr = (n, d), err
        end
    end
    return best
end

#
# Dates and times
#

const _MONTHS = ("January", "February", "March", "April", "May", "June", "July",
                 "August", "September", "October", "November", "December")
const _WEEKDAYS = ("Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday")

# The year, month, day and weekday (1 for Sunday) of Excel's day number `days`. In the
# 1900 system, day 0 is shown as 1900-01-00 and day 60 as 1900-02-29, as Excel shows
# them: it takes 1900 to be a leap year.
function _excel_ymd(days::Int, date1904::Bool)
    if date1904
        d = Dates.Date(1904, 1, 1) + Dates.Day(days)
        return Dates.year(d), Dates.month(d), Dates.day(d), Dates.dayofweek(d) % 7 + 1
    end
    weekday = mod(days + 6, 7) + 1
    days == 0 && return 1900, 1, 0, weekday
    days == 60 && return 1900, 2, 29, weekday
    d = Dates.Date(1899, 12, days < 60 ? 31 : 30) + Dates.Day(days)
    return Dates.year(d), Dates.month(d), Dates.day(d), weekday
end

function _render_date(toks::Vector{_FmtToken}, x::Float64, date1904::Bool, for_width::Bool)::String
    x < 0 && return "#"^11      # Excel can't show a negative date or time
    nfrac = min(maximum((parse(Int, t.text) for t in toks if t.kind == :fracsec); init=0), 3)
    per = 10^nfrac
    # The time to the millisecond, then to the digits shown, a half rounding down: the
    # rule that best matches Excel 365 (in all but 2 of 15 halves tried; Excel's own
    # rounding of halves depends on its floating-point arithmetic).
    days = floor(Int64, x)
    ms = round(Int64, (x - days) * 86_400_000)
    q, r = divrem(ms, 1000 ÷ per)
    units = days * 86400 * per + q + (2r > 1000 ÷ per ? 1 : 0)
    totalsec, frac = divrem(units, per)
    days, secofday = divrem(totalsec, 86400)
    h, rest = divrem(secofday, 3600)
    mi, s = divrem(rest, 60)
    y, mo, d, wd = _excel_ymd(Int(days), date1904)
    ampm = any(t -> t.kind == :ampm, toks)
    io = IOBuffer()
    for t in toks
        k, s_ = t.kind, t.text
        if k == :date && !isempty(s_)        # an empty one is a calendar choice
            c = s_[1]
            n = length(s_)
            if c == 'y' || c == 'b'          # b: the Buddhist year
                yr = c == 'b' ? y + 543 : y
                print(io, n <= 2 ? lpad(mod(yr, 100), 2, '0') : lpad(yr, 4, '0'))
            elseif c == 'g'                  # an era name, which has no Gregorian text
                nothing
            elseif c == 'e'
                print(io, y)
            elseif c == 'm'
                print(io, n == 1 ? string(mo) : n == 2 ? lpad(mo, 2, '0') : n == 3 ? _MONTHS[mo][1:3] :
                          n == 5 ? _MONTHS[mo][1:1] : _MONTHS[mo])
            elseif c == 'd'
                print(io, n == 1 ? string(d) : n == 2 ? lpad(d, 2, '0') : n == 3 ? _WEEKDAYS[wd][1:3] : _WEEKDAYS[wd])
            elseif c == 'h'
                hh = ampm ? (mod(h, 12) == 0 ? 12 : mod(h, 12)) : h
                print(io, n == 1 ? string(hh) : lpad(hh, 2, '0'))
            elseif c == 's'
                print(io, n == 1 ? string(s) : lpad(s, 2, '0'))
            end
        elseif k == :minute
            print(io, length(s_) == 1 ? string(mi) : lpad(mi, 2, '0'))
        elseif k == :elapsed
            total = s_[1] == 'h' ? totalsec ÷ 3600 : s_[1] == 'm' ? totalsec ÷ 60 : totalsec
            print(io, lpad(total, length(s_), '0'))
        elseif k == :fracsec
            n = parse(Int, s_)
            print(io, '.', lpad(frac, nfrac, '0')[1:n])
        elseif k == :ampm
            am = h < 12
            text = length(s_) == 5 ? (am ? s_[1:2] : string(s_[4], s_[5])) : (am ? s_[1:1] : s_[3:3])
            print(io, text)
        elseif k == :lit || k == :digit || k == :point || k == :comma
            print(io, s_)
        elseif k == :pad
            print(io, for_width ? s_ : " ")
        end
    end
    return String(take!(io))
end

#
# A cell's displayed text
#

# The format code of cell style `style`: one the file defines, or a built-in one. The
# built-ins Excel shows differently in each locale (23–36, 50–81) are shown as General.
function _format_code(wb::Workbook, style::Integer)::String
    style == 0 && return "General"
    id = styles_cell_xf_numFmtId(wb, style)
    id == 0 && return "General"
    cache = get_numFmt_cache(wb)
    haskey(cache, id) && return cache[id]
    return get(builtinFormats, string(id), "General")
end

# The text Excel shows for `cell`. `formats` caches parsed format codes across cells.
function _cell_text(ws::Worksheet, cell::AbstractCell; for_width::Bool=false,
                    formats::Dict{String,Union{Nothing,Vector{_FmtSection}}}=Dict{String,Union{Nothing,Vector{_FmtSection}}}())::String
    cell isa EmptyCell && return ""
    dt = cell.datatype
    dt == CT_EMPTY && return ""
    dt == CT_ERROR && return get_error_string(cell.value)
    dt == CT_RICHVALUE && return ""   # a picture shows no text
    wb = get_workbook(ws)
    v = if dt == CT_STRING
        sst_unformatted_string(wb, reinterpret(Int64, cell.value))
    elseif dt == CT_BOOL
        cell.value != 0
    elseif dt == CT_INT || dt == CT_DATE
        Float64(reinterpret(Int64, cell.value))
    else
        reinterpret(Float64, cell.value)
    end
    code = _format_code(wb, cell.style)
    sections = get!(() -> _parse_numfmt(code), formats, code)
    return _display_text(v, code; date1904=isdate1904(wb), for_width, sections)
end

"""
    getText(sh::Worksheet, cr::String) -> String
    getText(xf::XLSXFile, cr::String) -> String

    getText(sh::Worksheet, row::Int, col::Int) -> String

The text Excel shows in a cell: its value in its number format, as Excel's `Range.Text`
gives it when the column is wide enough. `1234.5` formatted `"#,##0.00"` is
`"1,234.50"`, a date formatted `"d mmm yyyy"` is `"7 Mar 2026"`, `0.25` formatted `"0%"`
is `"25%"`. A formula shows its cached value; an empty cell, or a formula without
one, is `""`.

The rendering was checked against Excel 365 for every part of a number format: General,
digit placeholders, thousands separators, scaling, percentages, scientific notation,
fractions, dates, times, elapsed times, text sections, literals, currency and locale
codes, colours and conditions. Excel's own results differ in a few cases:

- A fill (`*x`) repeats its character to fill the column; `getText` leaves it out, as
  it has no column width.
- The built-in formats Excel shows in the computer's locale (such as 14, the short date)
  are shown as Excel shows them in US English: `m/d/yyyy`. Formats 23–36 and 50–81,
  which Excel shows only in some locales, are shown as General.
- A time exactly half-way between two values shown to a fraction of a second may round
  the other way.

The cell may be read without the cell cache, and from a file opened for reading only.

# Examples
```julia
julia> sh["A1"] = 1234.5; XLSX.setFormat(sh, "A1"; format="#,##0.00");

julia> XLSX.getText(sh, "A1")
"1,234.50"
```

See also [`XLSX.autofitColumns`](@ref), which sizes columns to these texts, and
[`XLSX.setFormat`](@ref).
"""
function getText end
getText(xl::XLSXFile, sheetcell::AbstractString) = process_get_sheetcell(getText, xl, sheetcell)
getText(ws::Worksheet, cr::AbstractString) = process_get_cellname(getText, ws, cr)
getText(ws::Worksheet, row::Integer, col::Integer) = getText(ws, CellRef(row, col))
getText(ws::Worksheet, ref::CellRef)::String = _cell_text(ws, getcell(ws, ref))
