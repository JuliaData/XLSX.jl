# ===========================================================================
# Data validation
#
# A validation is a `<dataValidation>` in the worksheet's `<dataValidations>`
# block, applying to the cells in its `sqref`. Excel 2010 also wrote validations
# whose list came from another sheet as `<x14:dataValidation>`s under `<extLst>`.
# Both are read; only the first kind is written, which every Excel since 2010
# reads, wherever its list comes from.
#
# Formulas are stored without a leading `=`. `renamesheet!`, `deletesheet!` and
# `copysheet!` already keep the references in them up to date (see
# `SHEET_XML_FORMULA_CONTAINERS` in formula.jl).
# ===========================================================================

const DV_TYPES        = (:any, :whole, :decimal, :list, :date, :time, :textLength, :custom)
const DV_OPERATORS    = (:between, :notBetween, :equal, :notEqual, :greaterThan, :lessThan,
                         :greaterThanOrEqual, :lessThanOrEqual)
const DV_ERROR_STYLES = (:stop, :warning, :information)
const DV_MAX_TITLE    = 32    # characters Excel allows in `promptTitle` and `errorTitle`
const DV_MAX_MESSAGE  = 255   # ... in `prompt` and `error`, and in a list's own items

const DataValidationInfo = NamedTuple{
    (:sheet, :range, :type, :operator, :formula1, :formula2, :allowBlank, :dropdown,
     :showInputMessage, :promptTitle, :prompt, :showErrorMessage, :errorTitle, :error, :errorStyle),
    Tuple{String, Union{CellRange,NonContiguousRange}, Symbol, Union{Nothing,Symbol},
          Union{Nothing,String}, Union{Nothing,String}, Bool, Bool,
          Bool, Union{Nothing,String}, Union{Nothing,String}, Bool, Union{Nothing,String}, Union{Nothing,String}, Symbol},
}

# ===========================================================================
# Public API
# ===========================================================================

"""
    addDataValidation(ws::Worksheet, rng; type, kw...) -> NamedTuple
    addDataValidation(xf::XLSXFile, rng; type, kw...) -> NamedTuple

Restrict what may be typed into the cells of `rng`, as Excel's *Data Validation*
dialog does. `rng` may be a cell, a range, a column or row range (`"B:B"` is
the whole column), a non-contiguous range (`"A1:A5,C1:C5"`) or a defined name.
With an `XLSXFile`, `rng` names its sheet (`"Sheet1!B2:B20"`).

A cell can have only one validation, so `rng` may not overlap a range that
already has one. Remove that first with [`deleteDataValidation`](@ref).

# Keyword arguments

`type` is the kind of value allowed:

| `type`        | Allows                                              |
| :------------ | :-------------------------------------------------- |
| `:list`       | one of a list of values, chosen from a drop-down    |
| `:whole`      | a whole number                                      |
| `:decimal`    | a number                                            |
| `:date`       | a date                                              |
| `:time`       | a time                                              |
| `:textLength` | text of a given length                              |
| `:custom`     | anything for which the formula in `formula1` is true |
| `:any`        | any value: shows only the input message             |

`type` defaults to `:list` when `values` is given.

For `:whole`, `:decimal`, `:date`, `:time` and `:textLength`, `operator` says how
the value compares with `formula1` (and `formula2`). It is one of `:between`
(the default), `:notBetween`, `:equal`, `:notEqual`, `:greaterThan`, `:lessThan`,
`:greaterThanOrEqual` or `:lessThanOrEqual`. `:between` and `:notBetween` need
both `formula1` and `formula2`; the others need `formula1` only.

`formula1` and `formula2` may each be a number, a `Date`, `Time` or `DateTime`,
or a string holding a cell reference, a defined name or a formula. A leading `=`
is optional. Use absolute references (`"\\\$A\\\$1"`) for a list taken from cells,
as in Excel: a relative reference moves with each validated cell.

`values` gives a list's items directly, e.g. `["Red", "Green", "Blue"]`, instead
of `formula1`. Items may not contain commas, and the whole list may be at most
255 characters, as in Excel. For a longer list, put the items in cells and give
their range as `formula1`.

The other keywords set the remaining options in Excel's dialog:

- `allowBlank::Bool = true`: whether an empty cell is valid.
- `dropdown::Bool = true`: for a `:list`, whether the cell shows a drop-down arrow.
- `prompt`, `promptTitle`: a message shown when the cell is selected, and its title.
  `showInputMessage::Bool = true` turns it on.
- `error`, `errorTitle`: the message shown when an invalid value is entered, and its
  title. `showErrorMessage::Bool = true` turns this check on.
- `errorStyle`: `:stop` (the default) refuses an invalid value, `:warning` asks
  whether to keep it, and `:information` only tells the user.

Titles may be at most 32 characters long and messages 255.

# Return Value

The validation added, as [`getDataValidations`](@ref) reads it.

# Examples

```julia
XLSX.addDataValidation(sheet, "B2:B100"; values=["Yes", "No", "Maybe"])
XLSX.addDataValidation(sheet, "C2:C100"; type=:list, formula1="Lists!\\\$A\\\$1:\\\$A\\\$12")
XLSX.addDataValidation(sheet, "D:D"; type=:whole, operator=:between, formula1=1, formula2=10,
                       error="Enter a whole number from 1 to 10.")
XLSX.addDataValidation(sheet, "E2"; type=:date, operator=:greaterThan, formula1=Date(2025, 1, 1))
XLSX.addDataValidation(sheet, "F2:F50"; type=:custom, formula1="=ISNUMBER(F2)")
XLSX.addDataValidation(sheet, "G1"; type=:any, prompt="Type your name here")
```

See also [`getDataValidations`](@ref), [`getDataValidation`](@ref), [`deleteDataValidation`](@ref).
"""
function addDataValidation end

"""
    deleteDataValidation(ws::Worksheet, rng) -> Int
    deleteDataValidation(xf::XLSXFile, rng) -> Int

Remove every data validation that applies only to cells inside `rng`, and return
how many were removed. A validation that also applies to cells outside `rng` is
left alone. `rng` takes the same forms as in [`addDataValidation`](@ref), so
`deleteDataValidation(ws, :)` removes all of a sheet's validations.

See also [`addDataValidation`](@ref).
"""
function deleteDataValidation end

for F in (:addDataValidation, :deleteDataValidation)
    @eval begin
        $F(ws::Worksheet, ref::AbstractString; kw...) = process_ranges($F, ws, ref; kw...)
        $F(xf::XLSXFile, ref::AbstractString; kw...) = process_sheetcell($F, xf, ref; kw...)
        $F(ws::Worksheet, c::CellRef; kw...) = $F(ws, CellRange(c, c); kw...)
        $F(ws::Worksheet, r::ColumnRange; kw...) =
            $F(ws, CellRange(CellRef(1, r.start), CellRef(EXCEL_MAX_ROWS, r.stop)); kw...)
        $F(ws::Worksheet, r::RowRange; kw...) =
            $F(ws, CellRange(CellRef(r.start, 1), CellRef(r.stop, EXCEL_MAX_COLS)); kw...)
        $F(ws::Worksheet, r::SheetCellRef; kw...) = do_sheet_names_match(ws, r) && $F(ws, r.cellref; kw...)
        $F(ws::Worksheet, r::SheetCellRange; kw...) = do_sheet_names_match(ws, r) && $F(ws, r.rng; kw...)
        $F(ws::Worksheet, r::SheetColumnRange; kw...) = do_sheet_names_match(ws, r) && $F(ws, r.colrng; kw...)
        $F(ws::Worksheet, r::SheetRowRange; kw...) = do_sheet_names_match(ws, r) && $F(ws, r.rowrng; kw...)
        $F(ws::Worksheet, ::Colon; kw...) = $F(ws, CellRange(CellRef(1, 1), CellRef(EXCEL_MAX_ROWS, EXCEL_MAX_COLS)); kw...)
    end
end

function addDataValidation(
    ws::Worksheet,
    rng::CfRange;
    type::Union{Nothing,Symbol} = nothing,
    operator::Union{Nothing,Symbol} = nothing,
    formula1 = nothing,
    formula2 = nothing,
    values::Union{Nothing,AbstractVector} = nothing,
    allowBlank::Bool = true,
    dropdown::Union{Nothing,Bool} = nothing,
    showInputMessage::Bool = true,
    promptTitle::Union{Nothing,AbstractString} = nothing,
    prompt::Union{Nothing,AbstractString} = nothing,
    showErrorMessage::Bool = true,
    errorTitle::Union{Nothing,AbstractString} = nothing,
    error::Union{Nothing,AbstractString} = nothing,
    errorStyle::Symbol = :stop,
)::DataValidationInfo
    _require_writable(ws, "add a data validation")
    _require_worksheet(ws, "add a data validation")
    rng isa NonContiguousRange && do_sheet_names_match(ws, rng)

    # --- what is allowed
    isnothing(type) && !isnothing(values) && (type = :list)
    isnothing(type) && throw(XLSXError("Give the kind of value allowed as `type`: one of $(join(DV_TYPES, ", "))."))
    type in DV_TYPES || throw(XLSXError("Invalid `type` `:$type`. Must be one of $(join(DV_TYPES, ", "))."))
    if !isnothing(values)
        type === :list || throw(XLSXError("`values` gives the items of a `:list` validation, not a `:$type` one."))
        isnothing(formula1) || throw(XLSXError("Give a list's items as `values` or as `formula1`, not both."))
        formula1 = _dv_list_literal(values)
    end
    f1 = _dv_formula_text(ws, formula1, "formula1")
    f2 = _dv_formula_text(ws, formula2, "formula2")

    if type === :any
        (isnothing(f1) && isnothing(f2) && isnothing(operator)) ||
            throw(XLSXError("An `:any` validation allows any value, so it takes no `formula1`, `formula2` or `operator`."))
    elseif type in (:list, :custom)
        isnothing(f1) && throw(XLSXError("A `:$type` validation needs `formula1`$(type === :list ? " or `values`" : "")."))
        isnothing(operator) || throw(XLSXError("A `:$type` validation takes no `operator`."))
        isnothing(f2) || throw(XLSXError("A `:$type` validation takes no `formula2`."))
    else
        operator = something(operator, :between)
        operator in DV_OPERATORS ||
            throw(XLSXError("Invalid `operator` `:$operator`. Must be one of $(join(DV_OPERATORS, ", "))."))
        isnothing(f1) && throw(XLSXError("A `:$type` validation needs `formula1`."))
        two = operator in (:between, :notBetween)
        two && isnothing(f2) && throw(XLSXError("Operator `:$operator` needs `formula2` as well as `formula1`."))
        !two && !isnothing(f2) && throw(XLSXError("Operator `:$operator` takes `formula1` only, not `formula2`."))
    end
    isnothing(dropdown) || type === :list ||
        throw(XLSXError("`dropdown` applies to a `:list` validation, not a `:$type` one."))

    # --- messages
    errorStyle in DV_ERROR_STYLES ||
        throw(XLSXError("Invalid `errorStyle` `:$errorStyle`. Must be one of $(join(DV_ERROR_STYLES, ", "))."))
    for (k, v, n) in (("promptTitle", promptTitle, DV_MAX_TITLE), ("errorTitle", errorTitle, DV_MAX_TITLE),
                      ("prompt", prompt, DV_MAX_MESSAGE), ("error", error, DV_MAX_MESSAGE))
        isnothing(v) || length(v) <= n ||
            throw(XLSXError("`$k` may be at most $n characters long. Found $(length(v))."))
    end

    # --- one validation per cell
    xf = get_xlsxfile(ws)
    sheet_path = _sheet_path(ws)
    for dv in getDataValidations(ws)
        _dv_overlap(dv.range, rng) &&
            throw(XLSXError("`$(_cf_sqref_user(rng))` overlaps `$(_cf_sqref_user(dv.range))`, which already has a data validation. Delete that first."))
    end

    el = XML.Element(prefixed_tag(_root_prefix(xml_root_element(get_xml_data(xf, sheet_path))), "dataValidation"))
    type === :any || (el["type"] = string(type))
    errorStyle === :stop || (el["errorStyle"] = string(errorStyle))
    isnothing(operator) || operator === :between || (el["operator"] = string(operator))
    allowBlank && (el["allowBlank"] = "1")
    dropdown === false && (el["showDropDown"] = "1")   # sic: OOXML's showDropDown="1" *hides* the arrow
    showInputMessage && (el["showInputMessage"] = "1")
    showErrorMessage && (el["showErrorMessage"] = "1")
    isnothing(errorTitle) || (el["errorTitle"] = errorTitle)
    isnothing(error) || (el["error"] = error)
    isnothing(promptTitle) || (el["promptTitle"] = promptTitle)
    isnothing(prompt) || (el["prompt"] = prompt)
    el["sqref"] = join((_dv_area_string(a) for a in _cf_areas(rng)), " ")
    pfx = _root_prefix(xml_root_element(get_xml_data(xf, sheet_path)))
    isnothing(f1) || push!(el, XML.Element(prefixed_tag(pfx, "formula1"), XML.Text(f1)))
    isnothing(f2) || push!(el, XML.Element(prefixed_tag(pfx, "formula2"), XML.Text(f2)))

    block = _ensure_dv_block!(xml_root_element(get_xml_data(xf, sheet_path)))
    push!(block, el)
    block["count"] = string(length(elements_with_tag(block, "dataValidation")))
    return _dv_info(ws, el)
end

function deleteDataValidation(ws::Worksheet, rng::CfRange)::Int
    _require_writable(ws, "delete a data validation")
    _require_worksheet(ws, "delete a data validation")
    rng isa NonContiguousRange && do_sheet_names_match(ws, rng)
    root = xml_root_element(get_xml_data(get_xlsxfile(ws), _sheet_path(ws)))
    n = 0
    for (parent, block) in _dv_blocks(root)
        kids = XML.children(block)
        for i in reverse(eachindex(kids))
            d = kids[i]
            XML.nodetype(d) == XML.Element && localname(d) == "dataValidation" || continue
            sq = _dv_sqref(d)
            (isnothing(sq) || !_dv_within(_dv_range(ws, sq), rng)) && continue
            deleteat!(kids, i)
            n += 1
        end
        left = length(elements_with_tag(block, "dataValidation"))
        if left > 0
            block["count"] = string(left)
        else
            filter!(c -> c !== block, XML.children(parent))
            # An x14 block sits in its own `<ext>`, which is dropped with it, and
            # `<extLst>` with that if nothing is left in it.
            if localname(parent) == "ext"
                ext_lst = first_element_with_tag(root, "extLst")
                isempty(xml_elements(parent)) && filter!(c -> c !== parent, XML.children(ext_lst))
                isempty(xml_elements(ext_lst)) && filter!(c -> c !== ext_lst, XML.children(root))
            end
        end
    end
    return n
end

"""
    getDataValidations(ws::Worksheet) -> Vector{NamedTuple}
    getDataValidations(xf::XLSXFile) -> Vector{NamedTuple}

Return every data validation on a worksheet, or on every worksheet of a workbook,
in the order the file holds them. Each is a `NamedTuple` with fields:

- `sheet`, and `range`: a `CellRange`, or a `NonContiguousRange` for several areas
- `type` and `operator`, as `Symbol`s (`operator` is `nothing` for `:list`,
  `:custom` and `:any`)
- `formula1` and `formula2`, as `String`s without a leading `=`, or `nothing`. A list
  given as items reads as a quoted string, e.g. `"\\"Yes,No\\""`
- `allowBlank`, `dropdown`, `showInputMessage` and `showErrorMessage`, as `Bool`s
- `promptTitle`, `prompt`, `errorTitle` and `error`, as `String`s or `nothing`
- `errorStyle`, as `:stop`, `:warning` or `:information`

Validations Excel 2010 wrote in its extended form, for a list taken from another
sheet, are read too. The file doesn't need to be opened for writing.

See also [`addDataValidation`](@ref), [`getDataValidation`](@ref).
"""
function getDataValidations(ws::Worksheet)::Vector{DataValidationInfo}
    is_chartsheet(get_workbook(ws), ws.name) && return DataValidationInfo[]
    out = DataValidationInfo[]
    for el in _sheet_top_elements(get_xlsxfile(ws), _sheet_path(ws), ("dataValidations", "extLst"))
        holder = XML.Node[el]
        localname(el) == "extLst" && (holder = XML.Node[b for e in xml_elements(el) for b in elements_with_tag(e, "dataValidations")])
        for block in holder, d in elements_with_tag(block, "dataValidation")
            isnothing(_dv_sqref(d)) || push!(out, _dv_info(ws, d))
        end
    end
    return out
end

function getDataValidations(xf::XLSXFile)::Vector{DataValidationInfo}
    wb = get_workbook(xf)
    return reduce(vcat, (getDataValidations(getsheet(wb, i)) for i in 1:sheetcount(wb)); init=DataValidationInfo[])
end

"""
    getDataValidation(ws::Worksheet, ref) -> Union{Nothing, NamedTuple}
    getDataValidation(ws::Worksheet, row::Integer, col::Integer) -> Union{Nothing, NamedTuple}

Return the data validation that applies to the given cell, as
[`getDataValidations`](@ref) reads it, or `nothing` if the cell has none.

See also [`addDataValidation`](@ref).
"""
getDataValidation(ws::Worksheet, row::Integer, col::Integer) = getDataValidation(ws, CellRef(row, col))
getDataValidation(ws::Worksheet, ref::AbstractString) = getDataValidation(ws, _single_cellref(ref, "Data validations are looked up for"))
function getDataValidation(ws::Worksheet, ref::CellRef)::Union{Nothing,DataValidationInfo}
    cell = CellRange(ref, ref)
    for dv in getDataValidations(ws)
        _dv_overlap(dv.range, cell) && return dv
    end
    return nothing
end

# ===========================================================================
# Helpers
# ===========================================================================

_cf_sqref_user(rng) = replace(_cf_sqref(rng), ' ' => ',')

_dv_area(a) = a isa CellRef ? CellRange(a, a) : a
# One area of an `sqref`: a single cell as Excel writes it, `B2` rather than `B2:B2`.
_dv_area_string(a) = (r = _dv_area(a); r.start == r.stop ? string(r.start) : string(r))
_dv_overlap(a::CfRange, b::CfRange) =
    any(intersects(_dv_area(x), _dv_area(y)) for x in _cf_areas(a) for y in _cf_areas(b))
# Every area of `inner` lies inside one of the areas of `outer`.
_dv_within(inner::CfRange, outer::CfRange) =
    all(any(issubset(_dv_area(x), _dv_area(y)) for y in _cf_areas(outer)) for x in _cf_areas(inner))

function _dv_list_literal(values::AbstractVector)::String
    isempty(values) && throw(XLSXError("A list needs at least one item in `values`."))
    items = String[string(v) for v in values]
    for s in items
        occursin(',', s) && throw(XLSXError("A list item may not contain a comma: `$s`. Put the items in cells and give their range as `formula1`."))
    end
    joined = join(items, ",")
    length(joined) > DV_MAX_MESSAGE && throw(XLSXError(
        "A list given as `values` may be at most $DV_MAX_MESSAGE characters long, commas included. Found $(length(joined)). " *
        "Put the items in cells and give their range as `formula1`."))
    return "\"" * replace(joined, "\"" => "\"\"") * "\""
end

_dv_formula_text(::Worksheet, ::Nothing, ::String) = nothing
_dv_formula_text(::Worksheet, ::Bool, k::String) = throw(XLSXError("`$k` can't be a Bool."))
_dv_formula_text(::Worksheet, v::Real, ::String) = string(isinteger(v) ? Int(v) : float(v))
_dv_formula_text(ws::Worksheet, v::Dates.Date, ::String) = string(date_to_excel_value(v, isdate1904(ws)))
_dv_formula_text(ws::Worksheet, v::Dates.DateTime, ::String) = string(datetime_to_excel_value(v, isdate1904(ws)))
_dv_formula_text(::Worksheet, v::Dates.Time, ::String) = string(time_to_excel_value(v))
function _dv_formula_text(::Worksheet, v::AbstractString, k::String)
    s = strip(v)
    startswith(s, '=') && (s = strip(s[nextind(s, 1):end]))
    isempty(s) && throw(XLSXError("`$k` can't be empty."))
    return String(s)
end
_dv_formula_text(::Worksheet, v, k::String) =
    throw(XLSXError("`$k` must be a number, a date or time, or a string holding a reference or formula. Got a $(typeof(v))."))

# The sheet's validation blocks, each with the element holding it: `<dataValidations>`
# in `<worksheet>`, and `<x14:dataValidations>` each in an `<ext>` under `<extLst>`.
function _dv_blocks(root::XML.Node)::Vector{Tuple{XML.Node,XML.Node}}
    out = Tuple{XML.Node,XML.Node}[(root, b) for b in elements_with_tag(root, "dataValidations")]
    ext_lst = first_element_with_tag(root, "extLst")
    isnothing(ext_lst) && return out
    for e in xml_elements(ext_lst), b in elements_with_tag(e, "dataValidations")
        push!(out, (e, b))
    end
    return out
end

function _ensure_dv_block!(root::XML.Node)::XML.Node
    block = first_element_with_tag(root, "dataValidations")
    isnothing(block) || return block
    block = XML.Element(prefixed_tag(_root_prefix(root), "dataValidations"))
    insert!(XML.children(root), insert_index(root, "dataValidations", WORKSHEET_ORDER) + 1, block)
    return block
end

# `sqref` is an attribute of a 2007 validation and an `<xm:sqref>` child of an x14 one.
function _dv_sqref(d::XML.Node)::Union{Nothing,String}
    a = get_attr(d, "sqref")
    isempty(a) || return a
    s = first_element_with_tag(d, "sqref")
    return isnothing(s) ? nothing : XML.simple_value(s)
end

# The text of `<formula1>` or `<formula2>`: their own text in a 2007 validation, and
# the text of their `<xm:f>` in an x14 one.
function _dv_formula(d::XML.Node, tag::String)::Union{Nothing,String}
    f = first_element_with_tag(d, tag)
    isnothing(f) && return nothing
    inner = first_element_with_tag(f, "f")
    t = XML.is_simple_value(something(inner, f))
    return isnothing(t) ? nothing : String(t)
end

# Areas of an `sqref`, which Excel may write as whole columns or rows (`B:B`, `3:5`).
function _dv_range(ws::Worksheet, sqref::AbstractString)::CfRange
    areas = Union{CellRef,CellRange}[]
    for p in split(strip(sqref))
        if is_valid_cellname(p)
            push!(areas, CellRef(p))
        elseif is_valid_cellrange(p)
            push!(areas, CellRange(p))
        elseif is_valid_column_range(p)
            c = ColumnRange(p)
            push!(areas, CellRange(CellRef(1, c.start), CellRef(EXCEL_MAX_ROWS, c.stop)))
        elseif is_valid_row_range(p)
            r = RowRange(p)
            push!(areas, CellRange(CellRef(r.start, 1), CellRef(r.stop, EXCEL_MAX_COLS)))
        else
            throw(XLSXError("Malformed worksheet `$(ws.name)`: a data validation applies to `$sqref`."))
        end
    end
    length(areas) == 1 && return _dv_area(only(areas))
    return NonContiguousRange(ws.name, areas)
end

_dv_flag(d::XML.Node, k::String) = get_attr(d, k) in ("1", "true")
_dv_text(d::XML.Node, k::String) = (v = get_attr(d, k); isempty(v) ? nothing : v)

function _dv_info(ws::Worksheet, d::XML.Node)::DataValidationInfo
    t = Symbol(get_attr(d, "type", "none"))
    t === :none && (t = :any)
    op = t in (:any, :list, :custom) ? nothing : Symbol(get_attr(d, "operator", "between"))
    return DataValidationInfo((
        ws.name, _dv_range(ws, _dv_sqref(d)), t, op,
        _dv_formula(d, "formula1"), _dv_formula(d, "formula2"),
        _dv_flag(d, "allowBlank"), !_dv_flag(d, "showDropDown"),
        _dv_flag(d, "showInputMessage"), _dv_text(d, "promptTitle"), _dv_text(d, "prompt"),
        _dv_flag(d, "showErrorMessage"), _dv_text(d, "errorTitle"), _dv_text(d, "error"),
        Symbol(get_attr(d, "errorStyle", "stop")),
    ))
end
