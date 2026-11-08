# ===========================================================================
# Shifting references as rows or columns are inserted or deleted
#
# A `RowColEdit` describes one edit: on sheet `sheet`, `count` rows (or columns)
# inserted before row (column) `at`, or with a negative `count`, `-count` of them
# deleted starting at `at`. The functions here say where a reference ends up after
# the edit, as Excel moves it:
#
# - A reference at or after `at` moves by the number inserted or deleted, whether
#   it is relative or absolute (`$`).
# - A range across the edit grows (insert) or shrinks (delete). A range, or a cell,
#   wholly inside deleted rows (columns) is gone: `nothing`, or `#REF!` in a formula.
# - A reference pushed past the last row (column) is gone too, except the end of a
#   range, which stops at the last row (column), so `A1:A1048576` stays whole.
# - Whole columns (`A:C`) don't move as rows are edited, nor whole rows (`3:5`) as
#   columns are.
#
# Only references to `sheet` move: qualified with its name (in any case, as Excel
# compares sheet names), or unqualified in a formula on that sheet. References in
# strings, to other workbooks (`[1]Sheet1!A1`), and 3-D references
# (`Sheet1:Sheet3!A1`) are left as they are; so is anything built from text, such
# as `INDIRECT("A1")`, as in Excel.
#
# These are pure functions of text and ranges: applying them to a workbook is the
# job of `insertRows` and the other editing functions.
# ===========================================================================

struct RowColEdit
    sheet::String
    axis::Symbol   # :row or :column
    at::Int        # the first row (column) moved or deleted
    count::Int     # > 0: inserted before `at`; < 0: `-count` deleted from `at`

    function RowColEdit(sheet::AbstractString, axis::Symbol, at::Integer, count::Integer)
        axis in (:row, :column) || throw(XLSXError("Internal error: axis must be :row or :column, not :$axis."))
        maxn = axis === :row ? EXCEL_MAX_ROWS : EXCEL_MAX_COLS
        1 <= at <= maxn || throw(XLSXError("Invalid $axis number: $at. Must be from 1 to $maxn."))
        count == 0 && throw(XLSXError("Internal error: an edit of no rows or columns."))
        count < 0 && at - count - 1 > maxn &&
            throw(XLSXError("Cannot delete $(-count) $(axis)s from $axis $at: the sheet has only $maxn."))
        return new(String(sheet), axis, Int(at), Int(count))
    end
end

_edit_max(e::RowColEdit) = e.axis === :row ? EXCEL_MAX_ROWS : EXCEL_MAX_COLS
_edit_last_deleted(e::RowColEdit) = e.at - e.count - 1   # for a deletion

# Where row (column) `p` goes, or `nothing` if it is deleted or pushed off the sheet.
function _shift_index(p::Int, e::RowColEdit)::Union{Nothing,Int}
    if e.count > 0
        p < e.at && return p
        q = p + e.count
        return q > _edit_max(e) ? nothing : q
    end
    p < e.at && return p
    p <= _edit_last_deleted(e) && return nothing
    return p + e.count
end

# Where the span `lo:hi` goes, or `nothing` if all of it is deleted or pushed off.
function _shift_span(lo::Int, hi::Int, e::RowColEdit)::Union{Nothing,Tuple{Int,Int}}
    if e.count > 0
        lo2 = lo >= e.at ? lo + e.count : lo
        hi2 = hi >= e.at ? hi + e.count : hi
        lo2 > _edit_max(e) && return nothing
        return (lo2, min(hi2, _edit_max(e)))
    end
    last = _edit_last_deleted(e)
    e.at <= lo && hi <= last && return nothing
    lo2 = lo < e.at ? lo : (lo > last ? lo + e.count : e.at)
    hi2 = hi < e.at ? hi : (hi > last ? hi + e.count : e.at - 1)
    return (lo2, hi2)
end

# One end of a reference: `$A$1`, `B7`, `C` (in `C:E`) or `5` (in `5:9`).
const _REF_PART = r"^(\$?)([A-Za-z]{1,3})?(\$?)([0-9]+)?$"

# A reference without its sheet (`B7`, `$A$1:C9`, `C:E`, `5:9`) after the edit,
# or `nothing` if it is gone.
function _shift_ref_body(body::AbstractString, e::RowColEdit)::Union{Nothing,String}
    parts = [match(_REF_PART, p) for p in split(body, ':')]
    any(isnothing, parts) && return String(body)   # not a reference we understand: leave it
    cap = e.axis === :row ? 4 : 2                     # the row digits, or the column letters
    any(m -> isnothing(m[cap]), parts) && return String(body)   # a whole row (column) range: unmoved
    pos = [e.axis === :row ? parse(Int, m[4]) : decode_column_number(m[2]) for m in parts]
    if length(pos) == 1
        p = _shift_index(pos[1], e)
        isnothing(p) && return nothing
        new = [p]
    else
        span = _shift_span(minimum(pos), maximum(pos), e)
        isnothing(span) && return nothing
        new = pos[1] <= pos[2] ? collect(span) : reverse(collect(span))
    end
    return join((_ref_part_text(m, e, p, p != old) for (m, p, old) in zip(parts, new, pos)), ':')
end

function _ref_part_text(m::RegexMatch, e::RowColEdit, p::Int, changed::Bool)::String
    cabs, col, rabs, row = (something(c, "") for c in m.captures)
    changed || return m.match
    e.axis === :row ? cabs * col * rabs * string(p) : cabs * encode_column_number(p) * rabs * row
end

"""
    _shift_cellref(ref::CellRef, e::RowColEdit) -> Union{Nothing, CellRef}

Where cell `ref` goes after edit `e`, or `nothing` if it is deleted.
"""
function _shift_cellref(ref::CellRef, e::RowColEdit)::Union{Nothing,CellRef}
    if e.axis === :row
        r = _shift_index(row_number(ref), e)
        return isnothing(r) ? nothing : CellRef(r, column_number(ref))
    end
    c = _shift_index(column_number(ref), e)
    return isnothing(c) ? nothing : CellRef(row_number(ref), c)
end

"""
    _shift_cellrange(rng::CellRange, e::RowColEdit) -> Union{Nothing, CellRange}

Where range `rng` goes after edit `e`: grown, shrunk or moved, or `nothing` if all
of it is deleted.
"""
function _shift_cellrange(rng::CellRange, e::RowColEdit)::Union{Nothing,CellRange}
    t, l, b, r = row_number(rng.start), column_number(rng.start), row_number(rng.stop), column_number(rng.stop)
    if e.axis === :row
        span = _shift_span(t, b, e)
        return isnothing(span) ? nothing : CellRange(span[1], l, span[2], r)
    end
    span = _shift_span(l, r, e)
    return isnothing(span) ? nothing : CellRange(t, span[1], b, span[2])
end

"""
    _shift_sqref(sqref::AbstractString, e::RowColEdit) -> Union{Nothing, String}

A space-separated list of cells and ranges (an OOXML `sqref`, all on the edited
sheet) after edit `e`, without the parts that are deleted; `nothing` if all are.
"""
function _shift_sqref(sqref::AbstractString, e::RowColEdit)::Union{Nothing,String}
    parts = String[]
    for p in split(sqref)
        q = _shift_ref_body(p, e)
        isnothing(q) || push!(parts, q)
    end
    return isempty(parts) ? nothing : join(parts, ' ')
end

# --- formulas

# A sheet name in a reference: quoted ('My Sheet', with '' for a quote) or plain.
const _RX_SHEET = raw"(?:'(?:[^']|'')+'|[\p{L}_\\][\p{L}\p{N}_.]*)"
const _RX_CELL  = raw"\$?[A-Za-z]{1,3}\$?[0-9]+"
const _RX_BODY  = "(?:$_RX_CELL(?::$_RX_CELL)?|\\\$?[A-Za-z]{1,3}:\\\$?[A-Za-z]{1,3}|\\\$?[0-9]+:\\\$?[0-9]+)"
# One of: a string literal, a bracketed part (a structured reference, or an
# external workbook's index with the reference that follows it), or a reference.
const _RX_FORMULA_TOKEN = Regex(
    "(?<str>\"(?:[^\"]|\"\")*\")" *
    "|(?<ext>(?<br>\\[(?:[^\\[\\]]|(?&br))*\\])(?:$_RX_SHEET(?::$_RX_SHEET)?!$_RX_BODY)?)" *
    "|(?<![\\p{L}\\p{N}_.\\\$])(?:(?<s1>$_RX_SHEET)(?<s2>:$_RX_SHEET)?!)?(?<body>$_RX_BODY)(?![\\p{L}\\p{N}_.(\\[!])")

# Whether `formula`, on sheet `own_sheet` (or `nothing`), refers to a cell of `sheet`.
function _formula_refers_to(formula::AbstractString, sheet::AbstractString, own_sheet::Union{Nothing,AbstractString})::Bool
    for m in eachmatch(_RX_FORMULA_TOKEN, String(formula))
        (isnothing(m[:body]) || !isnothing(m[:s2])) && continue
        s = isnothing(m[:s1]) ? own_sheet : unquoteit(m[:s1])
        !isnothing(s) && _same_sheetname(s, sheet) && return true
    end
    return false
end

"""
    _shift_formula(formula::AbstractString, e::RowColEdit; own_sheet=nothing) -> String

`formula` with its references to `e.sheet` moved by edit `e`, and those to deleted
cells made `#REF!`. `own_sheet` is the sheet the formula is on, which its
unqualified references refer to; `nothing` for one not on a sheet, such as a
defined name, whose references count only if qualified.
"""
function _shift_formula(formula::AbstractString, e::RowColEdit; own_sheet::Union{Nothing,AbstractString}=nothing)::String
    s = String(formula)
    out = IOBuffer()
    i = 1
    for m in eachmatch(_RX_FORMULA_TOKEN, s)
        write(out, SubString(s, i, prevind(s, m.offset)))
        i = m.offset + ncodeunits(m.match)
        body = m[:body]
        if isnothing(body) || !isnothing(m[:s2])   # a string, a bracketed part or a 3-D reference
            write(out, m.match)
            continue
        end
        sheet = isnothing(m[:s1]) ? own_sheet : unquoteit(m[:s1])
        if isnothing(sheet) || !_same_sheetname(sheet, e.sheet)
            write(out, m.match)
            continue
        end
        new = _shift_ref_body(body, e)
        prefix = SubString(m.match, 1, prevind(m.match, ncodeunits(m.match) - ncodeunits(body) + 1))
        write(out, prefix, something(new, "#REF!"))
    end
    write(out, SubString(s, i))
    return String(take!(out))
end
