#
# -- Ignored errors: the error-checking warnings Excel is told not to show for some cells
#

# The error checks that can be ignored, in the schema's order, as Excel's Error
# Checking rules name them.
const IGNORABLE_ERRORS = (
    evalError          = "a formula that results in an error",
    twoDigitTextYear   = "a date written as text with a two-digit year",
    numberStoredAsText = "a number stored as text",
    formula            = "a formula inconsistent with the formulas next to it",
    formulaRange       = "a formula that omits cells next to its range",
    unlockedFormula    = "an unlocked cell holding a formula",
    emptyCellReference = "a formula referring to empty cells",
    listDataValidation = "a value not allowed by the cell's list data validation",
    calculatedColumn   = "a formula inconsistent with its Excel Table's calculated column",
)

"""
    setIgnoredErrors(ws::Worksheet, ref; evalError, twoDigitTextYear, numberStoredAsText,
                     formula, formulaRange, unlockedFormula, emptyCellReference,
                     listDataValidation, calculatedColumn)

Stop Excel marking cells with the green triangle of an error-checking warning, as its
*Ignore Error* option does, or mark them again. Each keyword is a `Bool`: `true` ignores
that warning in the cells of `ref`, and `false` shows it again. Warnings not given are
left as they are.

| Keyword              | The warning for                                              |
|:---------------------|:-------------------------------------------------------------|
$(join(("| `$k` | $v |" for (k, v) in pairs(IGNORABLE_ERRORS)), "\n"))

`ref` is a cell or range (`"A1:A100"`), several separated by commas, whole columns or rows
(`"B:B"`, `"2:5"`), a defined name, or a `CellRange`.

The file must be open for writing (`mode="rw"`).

# Examples
```julia
julia> XLSX.setIgnoredErrors(sheet, "A2:A100"; numberStoredAsText=true)   # IDs kept as text

julia> XLSX.setIgnoredErrors(sheet, "C:C"; formula=true, formulaRange=true)
```

See also [`getIgnoredErrors`](@ref) and [`removeIgnoredErrors`](@ref).
"""
function setIgnoredErrors(ws::Worksheet, ref::Union{AbstractString,CellRef,CellRange}; kw...)
    _require_writable(ws, "set ignored errors")
    _require_worksheet(ws, "set ignored errors")
    isempty(kw) && throw(XLSXError("Give the warnings to ignore as keywords, such as `numberStoredAsText=true`."))
    for (k, v) in kw
        haskey(IGNORABLE_ERRORS, k) || throw(XLSXError("Unknown error check `$k`. Must be one of $(join(keys(IGNORABLE_ERRORS), ", "))."))
        v isa Bool || throw(XLSXError("`$k` must be `true` or `false`. Found `$(repr(v))`."))
    end
    region = _ignored_error_region(ws, ref)
    _update_ignored_errors!(ws, region) do flags
        for (k, v) in kw
            v ? push!(flags, k) : delete!(flags, k)
        end
    end
    return nothing
end

"""
    getIgnoredErrors(ws::Worksheet) -> Vector{Pair{String,Vector{Symbol}}}
    getIgnoredErrors(ws::Worksheet, cell) -> Vector{Symbol}

The error-checking warnings Excel is told not to show. For the whole sheet, each set of
cells, as a comma-separated string such as `"A1:A3,C5"`, with the warnings ignored in
them; for one cell, the warnings ignored in it. The warnings are named as the keywords of
[`setIgnoredErrors`](@ref), in that order.

Works on files opened for reading only.

# Examples
```julia
julia> XLSX.getIgnoredErrors(sheet)
1-element Vector{Pair{String, Vector{Symbol}}}:
 "A2:A100" => [:numberStoredAsText]

julia> XLSX.getIgnoredErrors(sheet, "A5")
1-element Vector{Symbol}:
 :numberStoredAsText
```
"""
function getIgnoredErrors(ws::Worksheet)::Vector{Pair{String,Vector{Symbol}}}
    return [join(_sqref_part.(rngs), ',') => _ordered_flags(flags) for (rngs, flags) in _read_ignored_errors(ws)]
end
function getIgnoredErrors(ws::Worksheet, cell::Union{AbstractString,CellRef})::Vector{Symbol}
    c = cell isa CellRef ? cell : is_valid_cellname(cell) ? CellRef(cell) :
        throw(XLSXError("Give a single cell, such as `\"A1\"`. Found `$(repr(cell))`."))
    flags = Set{Symbol}()
    for (rngs, f) in _read_ignored_errors(ws)
        any(r -> c ∈ r, rngs) && union!(flags, f)
    end
    return _ordered_flags(flags)
end

"""
    removeIgnoredErrors(ws::Worksheet) -> Bool
    removeIgnoredErrors(ws::Worksheet, ref) -> Bool

Show every error-checking warning again, in the whole sheet or in the cells of `ref`
(given as for [`setIgnoredErrors`](@ref)). Returns `false` if no warning was ignored there.

The file must be open for writing (`mode="rw"`).
"""
function removeIgnoredErrors(ws::Worksheet)::Bool
    _require_writable(ws, "remove ignored errors")
    _require_worksheet(ws, "remove ignored errors")
    isempty(_read_ignored_errors(ws)) && return false
    _set_sheet_element!(ws, "ignoredErrors", Pair{String,String}[]; replace=true, kids=XML.Node[])
    return true
end
function removeIgnoredErrors(ws::Worksheet, ref::Union{AbstractString,CellRef,CellRange})::Bool
    _require_writable(ws, "remove ignored errors")
    _require_worksheet(ws, "remove ignored errors")
    region = _ignored_error_region(ws, ref)
    any(((rngs, _),) -> any(r -> any(s -> _overlap(r, s) !== nothing, region), rngs), _read_ignored_errors(ws)) || return false
    _update_ignored_errors!(empty!, ws, region)
    return true
end

_ordered_flags(flags) = Symbol[k for k in keys(IGNORABLE_ERRORS) if k in flags]

# The cells of `ref` as ranges; whole columns and rows reach the edges of the sheet, as
# Excel writes them.
function _ignored_error_region(ws::Worksheet, ref)::Vector{CellRange}
    parts = ref isa AbstractString ? _rowcol_parts(ws, ref) : Any[ref]
    return map(parts) do p
        p isa CellRef && return CellRange(p, p)
        p isa CellRange && return p
        p isa ColumnRange && return CellRange(CellRef(1, p.start), CellRef(EXCEL_MAX_ROWS, p.stop))
        p isa RowRange && return CellRange(CellRef(p.start, 1), CellRef(p.stop, EXCEL_MAX_COLS))
        throw(XLSXError("Give cells, such as `A1:A100`, columns or rows. Found `$p`."))
    end
end

# The sheet's `<ignoredError>`s: each one's ranges and the warnings it ignores.
function _read_ignored_errors(ws::Worksheet)::Vector{Tuple{Vector{CellRange},Set{Symbol}}}
    out = Tuple{Vector{CellRange},Set{Symbol}}[]
    for block in _sheet_top_elements(get_xlsxfile(ws), _sheet_path(ws), ("ignoredErrors",)), el in xml_elements(block)
        localname(el) == "ignoredError" || continue
        rngs = CellRange[]
        for t in split(get(el, "sqref", ""))
            is_valid_cellname(t) ? push!(rngs, CellRange(CellRef(t), CellRef(t))) :
                is_valid_cellrange(t) && push!(rngs, CellRange(t))
        end
        flags = Set{Symbol}(k for k in keys(IGNORABLE_ERRORS) if _xml_bool(Dict(_atts_vector(el)), String(k), false))
        isempty(rngs) || isempty(flags) || push!(out, (rngs, flags))
    end
    return out
end

# Change the warnings ignored in the cells of `region` with `change!(flags)`, keeping each
# cell in one `<ignoredError>`, as Excel does: one for each set of warnings, listing the
# cells that have exactly that set, in the order the sets first appear.
function _update_ignored_errors!(change!::Function, ws::Worksheet, region::Vector{CellRange})
    pieces = Tuple{CellRange,Set{Symbol}}[]
    for (rngs, flags) in _read_ignored_errors(ws), r in rngs
        outside = [r]
        for s in region
            outside = reduce(vcat, (_subtract(o, s) for o in outside); init=CellRange[])
        end
        append!(pieces, ((o, flags) for o in outside))
        new = copy(flags)
        change!(new)
        for (i, s) in enumerate(region)   # the part of `r` in each range of the region, counted once
            o = _overlap(r, s)
            isnothing(o) && continue
            append!(pieces, ((p, new) for p in _minus_all(o, region[1:i-1])))
        end
    end
    # cells of the region no `<ignoredError>` had yet
    fresh = Set{Symbol}()
    change!(fresh)
    if !isempty(fresh)
        covered = CellRange[r for (rngs, _) in _read_ignored_errors(ws) for r in rngs]
        for (i, s) in enumerate(region), p in _minus_all(s, vcat(covered, region[1:i-1]))
            push!(pieces, (p, fresh))
        end
    end

    groups = Tuple{Set{Symbol},Vector{CellRange}}[]
    for (r, flags) in pieces
        isempty(flags) && continue
        g = findfirst(x -> x[1] == flags, groups)
        isnothing(g) ? push!(groups, (flags, [r])) : push!(groups[g][2], r)
    end
    pfx = something(_ns_prefix(XML.tag(_sheet_root(ws))), "")
    kids = XML.Node[_xml_element(pfx * "ignoredError",
                                 ["sqref" => join(_sqref_part.(rngs), ' '), (String(k) => "1" for k in _ordered_flags(flags))...])
                    for (flags, rngs) in groups]
    _set_sheet_element!(ws, "ignoredErrors", Pair{String,String}[]; replace=true, kids)
    return nothing
end

# Rectangle arithmetic on cell ranges.
function _overlap(a::CellRange, b::CellRange)::Union{Nothing,CellRange}
    top, bottom = max(a.start.row_number, b.start.row_number), min(a.stop.row_number, b.stop.row_number)
    left, right = max(a.start.column_number, b.start.column_number), min(a.stop.column_number, b.stop.column_number)
    (top > bottom || left > right) && return nothing
    return CellRange(CellRef(top, left), CellRef(bottom, right))
end

# The cells of `a` outside `b`, as at most four ranges: above, below, left and right of `b`.
function _subtract(a::CellRange, b::CellRange)::Vector{CellRange}
    o = _overlap(a, b)
    isnothing(o) && return [a]
    out = CellRange[]
    rect(t, l, b_, r) = (t <= b_ && l <= r) && push!(out, CellRange(CellRef(t, l), CellRef(b_, r)))
    rect(a.start.row_number, a.start.column_number, o.start.row_number - 1, a.stop.column_number)
    rect(o.stop.row_number + 1, a.start.column_number, a.stop.row_number, a.stop.column_number)
    rect(o.start.row_number, a.start.column_number, o.stop.row_number, o.start.column_number - 1)
    rect(o.start.row_number, o.stop.column_number + 1, o.stop.row_number, a.stop.column_number)
    return out
end

_minus_all(a::CellRange, bs) = foldl((acc, b) -> reduce(vcat, (_subtract(x, b) for x in acc); init=CellRange[]), bs; init=[a])
