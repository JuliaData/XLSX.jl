# ===========================================================================
# Sheet-level settings: visibility, tab colour, protection and page setup
#
# A sheet's visibility is the `state` of its `<sheet>` in workbook.xml, held in
# `ws.visibility` and written by `update_workbook_xml!`. Its tab colour is the
# `<tabColor>` in the sheet's own `<sheetPr>`, which worksheets and chartsheets
# both have as their first element.
# ===========================================================================

const SHEET_VISIBILITIES = (:visible, :hidden, :veryHidden)

# `ws.is_hidden` was a field before `visibility` replaced it; keep it working.
function Base.getproperty(ws::Worksheet, s::Symbol)
    s === :is_hidden && return getfield(ws, :visibility) !== :visible
    return getfield(ws, s)
end
function Base.setproperty!(ws::Worksheet, s::Symbol, v)
    if s === :is_hidden
        v || return setfield!(ws, :visibility, :visible)
        getfield(ws, :visibility) === :visible && setfield!(ws, :visibility, :hidden)
        return v
    end
    return setfield!(ws, s, convert(fieldtype(Worksheet, s), v))
end

"""
    setSheetVisibility(ws::Worksheet, visibility::Symbol)

Show or hide a sheet's tab. `visibility` is one of

- `:visible`: the tab is shown;
- `:hidden`: the tab is hidden, and can be shown again from Excel's *Unhide* menu;
- `:veryHidden`: the tab is hidden and not listed in the *Unhide* menu. It can
  be shown again only from VBA, or with this function.

A workbook must keep at least one visible sheet, so hiding the last one is an
error. Hiding the sheet the workbook opens on makes the first visible sheet the
one it opens on instead.

Works on worksheets and chartsheets. The file must be open for writing
(`mode="rw"`).

# Examples
```julia
julia> XLSX.setSheetVisibility(xf["Lookups"], :hidden)

julia> XLSX.setSheetVisibility(xf["Lookups"], :visible)
```

See also [`getSheetVisibility`](@ref).
"""
function setSheetVisibility(ws::Worksheet, visibility::Symbol)
    _require_writable(ws, "change a sheet's visibility")
    visibility in SHEET_VISIBILITIES ||
        throw(XLSXError("Invalid visibility `:$visibility`. Must be one of $(join(SHEET_VISIBILITIES, ", "))."))
    visibility === ws.visibility && return nothing
    wb = get_workbook(ws)
    if visibility !== :visible && ws.visibility === :visible &&
       count(s -> s.visibility === :visible, wb.sheets) == 1
        throw(XLSXError("Cannot hide `$(ws.name)`: it is the workbook's only visible sheet, and Excel needs one."))
    end
    ws.visibility = visibility
    visibility === :visible || _move_active_tab_from!(ws)
    update_workbook_xml!(get_xlsxfile(ws))
    return nothing
end

"""
    getSheetVisibility(ws::Worksheet) -> Symbol

Whether a sheet's tab is shown: `:visible`, `:hidden` or `:veryHidden`. See
[`setSheetVisibility`](@ref).
"""
getSheetVisibility(ws::Worksheet)::Symbol = ws.visibility

# A hidden sheet cannot be the one the workbook opens on: if `ws` is, make the
# first visible sheet that one, moving the selected tab with it.
function _move_active_tab_from!(ws::Worksheet)
    xf = get_xlsxfile(ws)
    wb = get_workbook(ws)
    pos = findfirst(s -> s === ws, wb.sheets) - 1
    root = xml_root_element(xmlroot(xf, "xl/workbook.xml"))
    views = first_element_with_tag(root, "bookViews")
    view = isnothing(views) ? nothing : first_element_with_tag(views, "workbookView")
    active = isnothing(view) ? 0 : parse(Int, get(view, "activeTab", "0"))
    active == pos || return nothing

    new = findfirst(s -> s.visibility === :visible, wb.sheets) - 1
    if isnothing(view)
        pfx = get_prefix("xl/workbook.xml", xf)
        pfx = pfx == "" ? pfx : pfx * ":"
        if isnothing(views)
            views = XML.Element(pfx * "bookViews")
            insert!(XML.children(root), insert_index(root, "bookViews", WORKBOOK_ORDER) + 1, views)
        end
        view = XML.Element(pfx * "workbookView")
        push!(views, view)
    end
    view["activeTab"] = string(new)
    _set_tab_selected!(ws, false)
    _set_tab_selected!(wb.sheets[new + 1], true)
    return nothing
end

function _set_tab_selected!(ws::Worksheet, selected::Bool)
    root = xml_root_element(get_xml_data(get_xlsxfile(ws), _sheet_path(ws)))
    views = first_element_with_tag(root, "sheetViews")
    isnothing(views) && return nothing
    for v in xml_elements(views)
        localname(v) == "sheetView" || continue
        (selected || haskey(v, "tabSelected")) && (v["tabSelected"] = selected ? "1" : "0")
    end
    update_worksheets_xml!(get_xlsxfile(ws))
    return nothing
end

"""
    setTabColor(ws::Worksheet, color)

Colour a sheet's tab. The `color` can be $COLOR_FORMS_DOC

Works on worksheets and chartsheets. The file must be open for writing
(`mode="rw"`).

# Examples
```julia
julia> XLSX.setTabColor(sheet, "FF0000")

julia> XLSX.setTabColor(sheet, :accent2)
```

See also [`getTabColor`](@ref) and [`removeTabColor`](@ref).
"""
function setTabColor(ws::Worksheet, color)
    _require_writable(ws, "set a tab colour")
    atts = _color_atts(color)   # checks the colour before anything changes
    _set_sheet_pr_child!(ws, "tabColor", atts)
    return nothing
end

"""
    getTabColor(ws::Worksheet) -> Union{Nothing, String}

The colour of a sheet's tab as an `"FFRRGGBB"` hex string, or `nothing` if it
has none. A theme colour is given as the colour it has in the workbook's theme.

See also [`setTabColor`](@ref).
"""
function getTabColor(ws::Worksheet)::Union{Nothing,String}
    atts = _sheet_pr_child_atts(ws, "tabColor")
    isnothing(atts) && return nothing
    return resolveColor(get_workbook(ws), Dict{String,String}(atts))
end

"""
    removeTabColor(ws::Worksheet) -> Bool

Remove a sheet's tab colour. Returns `false` if it had none.

The file must be open for writing (`mode="rw"`).
"""
function removeTabColor(ws::Worksheet)::Bool
    _require_writable(ws, "remove a tab colour")
    isnothing(_sheet_pr_child_atts(ws, "tabColor")) && return false
    _set_sheet_pr_child!(ws, "tabColor", nothing)
    return true
end

# ===========================================================================
# Sheet protection
# ===========================================================================

# What a user may do on a protected sheet, as in Excel's *Protect Sheet* dialog,
# and whether Excel allows it by default. Each is an OOXML attribute of
# `<sheetProtection>`, but inverted there: `formatCells="1"` *prevents* formatting.
const SHEET_PROTECTION_ALLOWS = (
    selectLockedCells = true, selectUnlockedCells = true,
    formatCells = false, formatColumns = false, formatRows = false,
    insertColumns = false, insertRows = false, insertHyperlinks = false,
    deleteColumns = false, deleteRows = false,
    sort = false, autoFilter = false, pivotTables = false,
    objects = false, scenarios = false,
)
# The attributes whose schema default is false (allowed); the others default to
# true (prevented). An attribute is written only where it differs from its default.
const _PROTECTION_ATTR_DEFAULT_FALSE = (:selectLockedCells, :selectUnlockedCells, :objects, :scenarios)

"""
    protectSheet(ws::Worksheet; password=nothing, allow...)

Protect a worksheet, as Excel's *Review → Protect Sheet* does, so that its
locked cells can't be changed in Excel. Every cell is locked unless its format
says otherwise.

`password`, if given, must be typed in Excel to unprotect the sheet. It is
stored as Excel's legacy 16-bit hash, which Excel accepts but which is easily
broken: sheet protection guards against mistakes, not against anyone determined
to change the sheet. The password must be ASCII text of at most 15 characters,
the most the legacy hash uses.

The other keywords say what a user may still do, as in Excel's dialog. Each is
a `Bool`:

| Keyword               | Allows                          | Default |
| :-------------------- | :------------------------------ | :------ |
| `selectLockedCells`   | selecting locked cells          | `true`  |
| `selectUnlockedCells` | selecting unlocked cells        | `true`  |
| `formatCells`         | formatting cells                | `false` |
| `formatColumns`       | formatting columns              | `false` |
| `formatRows`          | formatting rows                 | `false` |
| `insertColumns`       | inserting columns               | `false` |
| `insertRows`          | inserting rows                  | `false` |
| `insertHyperlinks`    | inserting hyperlinks            | `false` |
| `deleteColumns`       | deleting columns                | `false` |
| `deleteRows`          | deleting rows                   | `false` |
| `sort`                | sorting                         | `false` |
| `autoFilter`          | using AutoFilter                | `false` |
| `pivotTables`         | using PivotTables               | `false` |
| `objects`             | editing objects such as charts  | `false` |
| `scenarios`           | editing scenarios               | `false` |

Protecting a sheet that is already protected replaces its settings. Protection
applies in Excel only: XLSX.jl reads and writes a protected sheet as any other.

The file must be open for writing (`mode="rw"`).

# Examples
```julia
julia> XLSX.protectSheet(sheet)

julia> XLSX.protectSheet(sheet; password="secret", sort=true, autoFilter=true)
```

See also [`unprotectSheet`](@ref), [`isSheetProtected`](@ref) and [`getSheetProtection`](@ref).
"""
function protectSheet(ws::Worksheet; password::Union{Nothing,AbstractString}=nothing, allow...)
    _require_writable(ws, "protect a sheet")
    _require_worksheet(ws, "protect a sheet")
    for (k, v) in allow
        haskey(SHEET_PROTECTION_ALLOWS, k) ||
            throw(XLSXError("Unknown keyword `$k`. Must be `password` or one of $(join(keys(SHEET_PROTECTION_ALLOWS), ", "))."))
        v isa Bool || throw(XLSXError("`$k` must be `true` or `false`. Found `$(repr(v))`."))
    end
    atts = Pair{String,String}[]
    isnothing(password) || push!(atts, "password" => _excel_password_hash(password))
    push!(atts, "sheet" => "1")
    for (k, default) in pairs(SHEET_PROTECTION_ALLOWS)
        prevented = !get(allow, k, default)
        prevented == (k ∉ _PROTECTION_ATTR_DEFAULT_FALSE) || push!(atts, String(k) => prevented ? "1" : "0")
    end
    _set_sheet_element!(ws, "sheetProtection", atts; replace=true)
    return nothing
end

"""
    unprotectSheet(ws::Worksheet) -> Bool

Remove a worksheet's protection, whatever its password. Returns `false` if it
wasn't protected.

The file must be open for writing (`mode="rw"`).
"""
function unprotectSheet(ws::Worksheet)::Bool
    _require_writable(ws, "unprotect a sheet")
    _require_worksheet(ws, "unprotect a sheet")
    isnothing(_sheet_element_atts(ws, "sheetProtection")) && return false
    _set_sheet_element!(ws, "sheetProtection", Pair{String,String}[]; replace=true)
    return true
end

"""
    isSheetProtected(ws::Worksheet) -> Bool

Whether a worksheet is protected. See [`protectSheet`](@ref).
"""
function isSheetProtected(ws::Worksheet)::Bool
    is_chartsheet(get_workbook(ws), ws.name) && return false
    atts = _sheet_element_atts(ws, "sheetProtection")
    return !isnothing(atts) && _xml_bool(Dict(atts), "sheet", false)
end

"""
    getSheetProtection(ws::Worksheet) -> Union{Nothing, NamedTuple}

How a worksheet is protected, or `nothing` if it isn't. The `NamedTuple` has a
field `password`, whether a password is needed to unprotect the sheet, and a
`Bool` field for each thing a user may do, named as the keywords of
[`protectSheet`](@ref).

# Examples
```julia
julia> p = XLSX.getSheetProtection(sheet);

julia> p.password, p.sort
(true, true)
```
"""
function getSheetProtection(ws::Worksheet)
    isSheetProtected(ws) || return nothing
    atts = Dict(_sheet_element_atts(ws, "sheetProtection"))
    pw = haskey(atts, "password") || haskey(atts, "hashValue")
    allows = map(k -> !_xml_bool(atts, String(k), k ∉ _PROTECTION_ATTR_DEFAULT_FALSE), keys(SHEET_PROTECTION_ALLOWS))
    return merge((password = pw,), NamedTuple{keys(SHEET_PROTECTION_ALLOWS)}(allows))
end

# Excel's legacy password hash (ECMA-376 Part 4, 14.7.1), as 4 hex digits.
function _excel_password_hash(password::AbstractString)::String
    isascii(password) || throw(XLSXError("A sheet password must be ASCII text."))
    n = length(password)
    # The legacy hash uses only the first 15 characters, so a longer password would
    # unlock with its first 15 alone. Refuse it rather than mislead.
    1 <= n <= 15 || throw(XLSXError("A sheet password must be from 1 to 15 characters long, the most Excel's legacy hash uses."))
    h = 0x0000
    for (i, c) in enumerate(password)
        v = UInt16(c)
        for _ in 1:i   # rotate left within 15 bits
            v = ((v << 1) & 0x7fff) | (v >> 14)
        end
        h ⊻= v
    end
    h ⊻= UInt16(n)
    h ⊻= 0xCE4B
    return uppercase(string(h; base=16, pad=4))
end

_xml_bool(atts::AbstractDict, k::AbstractString, default::Bool) =
    haskey(atts, k) ? atts[k] in ("1", "true") : default

# ===========================================================================
# Page setup and printing
# ===========================================================================

const PAPER_SIZES = (letter = 1, tabloid = 3, ledger = 4, legal = 5, executive = 7,
                     A3 = 8, A4 = 9, A5 = 11, B4 = 12, B5 = 13)

"""
    setPageSetup(ws::Worksheet; orientation, paperSize, scale, fitToWidth, fitToHeight)

Set how a worksheet is printed, as Excel's *Page Setup* dialog does. Only the
keywords given are changed.

- `orientation`: `:portrait` or `:landscape`.
- `paperSize`: one of $(join((":$k" for k in keys(PAPER_SIZES)), ", ")), or Excel's
  number for any other size.
- `scale`: print at this percentage of normal size, from 10 to 400.
- `fitToWidth`, `fitToHeight`: shrink the sheet to fit this many pages wide and
  tall. `0` means as many as needed, and so does leaving one out:
  `fitToWidth=1` alone fits the width to one page and lets the sheet run on to
  as many pages as it needs.

`scale` and fitting to pages are alternatives, as in Excel: giving one turns the
other off.

The file must be open for writing (`mode="rw"`).

# Examples
```julia
julia> XLSX.setPageSetup(sheet; orientation=:landscape, paperSize=:A4, fitToWidth=1)
```

See also [`getPageSetup`](@ref), [`setPageMargins`](@ref), [`setPrintOptions`](@ref),
[`setPrintArea`](@ref) and [`setPrintTitles`](@ref).
"""
function setPageSetup(ws::Worksheet; orientation::Union{Nothing,Symbol}=nothing,
                      paperSize::Union{Nothing,Symbol,Integer}=nothing, scale::Union{Nothing,Integer}=nothing,
                      fitToWidth::Union{Nothing,Integer}=nothing, fitToHeight::Union{Nothing,Integer}=nothing)
    _require_writable(ws, "set up the page")
    _require_worksheet(ws, "set up the page")
    fit = !isnothing(fitToWidth) || !isnothing(fitToHeight)
    changes = Pair{String,Union{Nothing,String}}[]
    if !isnothing(orientation)
        orientation in (:portrait, :landscape) ||
            throw(XLSXError("Invalid `orientation` `:$orientation`. Must be `:portrait` or `:landscape`."))
        push!(changes, "orientation" => String(orientation))
    end
    if !isnothing(paperSize)
        if paperSize isa Symbol
            haskey(PAPER_SIZES, paperSize) ||
                throw(XLSXError("Unknown `paperSize` `:$paperSize`. Must be one of $(join(keys(PAPER_SIZES), ", ")), or Excel's number for the size."))
            paperSize = PAPER_SIZES[paperSize]
        end
        paperSize >= 1 || throw(XLSXError("Invalid `paperSize` $paperSize. Excel numbers paper sizes from 1."))
        push!(changes, "paperSize" => string(paperSize))
    end
    if !isnothing(scale)
        fit && throw(XLSXError("Give `scale` or `fitToWidth`/`fitToHeight`, not both: Excel uses one or the other."))
        10 <= scale <= 400 || throw(XLSXError("Invalid `scale` $scale. Must be from 10 to 400 (percent)."))
        push!(changes, "scale" => string(scale))
    end
    if fit
        for (k, v) in (("fitToWidth", fitToWidth), ("fitToHeight", fitToHeight))
            v = something(v, 0)
            0 <= v <= 32767 || throw(XLSXError("Invalid `$k` $v. Must be from 0 (as many pages as needed) to 32767."))
            push!(changes, k => string(v))
        end
    end
    _set_sheet_element!(ws, "pageSetup", changes)
    if fit || !isnothing(scale)
        atts = something(_sheet_pr_child_atts(ws, "pageSetUpPr"), Pair{String,String}[])
        atts = _merge_atts(atts, ["fitToPage" => (fit ? "1" : nothing)])
        _set_sheet_pr_child!(ws, "pageSetUpPr", isempty(atts) ? nothing : atts)
    end
    return nothing
end

"""
    getPageSetup(ws::Worksheet) -> NamedTuple

How a worksheet is printed: a `NamedTuple` with fields `orientation`,
`paperSize`, `scale`, `fitToWidth` and `fitToHeight`, each `nothing` where the
sheet doesn't set it. `fitToWidth` and `fitToHeight` are given only if the sheet
is set to fit to pages, and `scale` only if it isn't. See [`setPageSetup`](@ref).
"""
function getPageSetup(ws::Worksheet)
    atts = Dict(something(_sheet_element_atts(ws, "pageSetup"), Pair{String,String}[]))
    fit = _xml_bool(Dict(something(_sheet_pr_child_atts(ws, "pageSetUpPr"), Pair{String,String}[])), "fitToPage", false)
    int(k) = haskey(atts, k) ? parse(Int, atts[k]) : nothing
    orientation = get(atts, "orientation", "default")
    return (orientation = orientation == "default" ? nothing : Symbol(orientation),
            paperSize = int("paperSize"),
            scale = fit ? nothing : int("scale"),
            fitToWidth = fit ? something(int("fitToWidth"), 1) : nothing,
            fitToHeight = fit ? something(int("fitToHeight"), 1) : nothing)
end

const _PAGE_MARGINS = (left = 0.7, right = 0.7, top = 0.75, bottom = 0.75, header = 0.3, footer = 0.3)

"""
    setPageMargins(ws::Worksheet; left, right, top, bottom, header, footer)

Set a worksheet's printed margins, in inches. `header` and `footer` are the
distances of the header and footer from the edge of the page. Only the keywords
given are changed; a sheet with no margins starts from Excel's defaults,
$(join(("`$k=$v`" for (k, v) in pairs(_PAGE_MARGINS)), ", ")).

The file must be open for writing (`mode="rw"`).

# Examples
```julia
julia> XLSX.setPageMargins(sheet; left=0.5, right=0.5)
```

See also [`getPageMargins`](@ref).
"""
function setPageMargins(ws::Worksheet; kw...)
    _require_writable(ws, "set page margins")
    _require_worksheet(ws, "set page margins")
    for (k, v) in kw
        haskey(_PAGE_MARGINS, k) || throw(XLSXError("Unknown keyword `$k`. Must be one of $(join(keys(_PAGE_MARGINS), ", "))."))
        (v isa Real && isfinite(v) && v >= 0) || throw(XLSXError("`$k` must be a number of inches, 0 or more. Found `$(repr(v))`."))
    end
    # All six are required, so a sheet without them takes the defaults for the rest.
    old = Dict(something(_sheet_element_atts(ws, "pageMargins"), Pair{String,String}[]))
    atts = [String(k) => (haskey(kw, k) ? string(Float64(kw[k])) : get(old, String(k), string(d))) for (k, d) in pairs(_PAGE_MARGINS)]
    _set_sheet_element!(ws, "pageMargins", atts)
    return nothing
end

"""
    getPageMargins(ws::Worksheet) -> Union{Nothing, NamedTuple}

A worksheet's printed margins in inches, as a `NamedTuple` with fields `left`,
`right`, `top`, `bottom`, `header` and `footer`, or `nothing` if the sheet
doesn't set them. See [`setPageMargins`](@ref).
"""
function getPageMargins(ws::Worksheet)
    atts = _sheet_element_atts(ws, "pageMargins")
    isnothing(atts) && return nothing
    d = Dict(atts)
    return NamedTuple{keys(_PAGE_MARGINS)}(map(k -> parse(Float64, get(d, String(k), string(_PAGE_MARGINS[k]))), keys(_PAGE_MARGINS)))
end

const _PRINT_OPTIONS = (:gridLines, :headings, :horizontalCentered, :verticalCentered)

"""
    setPrintOptions(ws::Worksheet; gridLines, headings, horizontalCentered, verticalCentered)

Choose what is printed with a worksheet. Each keyword is a `Bool`, and only
those given are changed:

- `gridLines`: print the cell gridlines;
- `headings`: print the row and column headings (`1, 2, 3...` and `A, B, C...`);
- `horizontalCentered`, `verticalCentered`: centre the sheet on the page.

The file must be open for writing (`mode="rw"`).

See also [`getPrintOptions`](@ref).
"""
function setPrintOptions(ws::Worksheet; kw...)
    _require_writable(ws, "set print options")
    _require_worksheet(ws, "set print options")
    for (k, v) in kw
        k in _PRINT_OPTIONS || throw(XLSXError("Unknown keyword `$k`. Must be one of $(join(_PRINT_OPTIONS, ", "))."))
        v isa Bool || throw(XLSXError("`$k` must be `true` or `false`. Found `$(repr(v))`."))
    end
    _set_sheet_element!(ws, "printOptions", [String(k) => (v ? "1" : nothing) for (k, v) in kw])
    return nothing
end

"""
    getPrintOptions(ws::Worksheet) -> NamedTuple

What is printed with a worksheet: a `NamedTuple` of `Bool`s with fields
`gridLines`, `headings`, `horizontalCentered` and `verticalCentered`. See
[`setPrintOptions`](@ref).
"""
function getPrintOptions(ws::Worksheet)
    d = Dict(something(_sheet_element_atts(ws, "printOptions"), Pair{String,String}[]))
    return NamedTuple{_PRINT_OPTIONS}(map(k -> _xml_bool(d, String(k), false), _PRINT_OPTIONS))
end

const PRINT_AREA = "_xlnm.Print_Area"
const PRINT_TITLES = "_xlnm.Print_Titles"

"""
    setPrintArea(ws::Worksheet, rng)

Print only the cells in `rng`, as Excel's *Page Layout → Print Area* does. `rng`
is a cell range (`"A1:F40"`), several separated by commas (`"A1:F40,H1:K10"`,
each printed on its own pages), or a defined name. It replaces any print area
the sheet has.

Excel keeps the print area as a name scoped to the sheet, `_xlnm.Print_Area`,
which is kept up to date as sheets are renamed, copied and deleted.

The file must be open for writing (`mode="rw"`).

See also [`getPrintArea`](@ref) and [`removePrintArea`](@ref).
"""
function setPrintArea(ws::Worksheet, rng::Union{AbstractString,CellRange})
    _require_writable(ws, "set a print area")
    _require_worksheet(ws, "set a print area")
    parts = rng isa CellRange ? Any[rng] : _rowcol_parts(ws, rng)
    for p in parts
        p isa Union{CellRef,CellRange} || throw(XLSXError("A print area is made of cells, such as `A1:F40`. Found `$p`."))
    end
    value = length(parts) == 1 ?
        SheetCellRange(ws.name, only(parts) isa CellRef ? CellRange(only(parts), only(parts)) : only(parts)) :
        NonContiguousRange(ws.name, Vector{Union{CellRef,CellRange}}(parts))
    get_workbook(ws).worksheet_names[(ws.sheetId, PRINT_AREA)] =
        DefinedNameValue(value, value isa NonContiguousRange ? fill(true, length(parts)) : true, false)
    return nothing
end

"""
    getPrintArea(ws::Worksheet) -> Union{Nothing, CellRange, NonContiguousRange, String}

A worksheet's print area: a `CellRange`, a `NonContiguousRange` of several, or
`nothing` if it has none. A print area Excel wrote in another form, such as
whole columns, is given as its text. See [`setPrintArea`](@ref).
"""
function getPrintArea(ws::Worksheet)
    v = get(get_workbook(ws).worksheet_names, (ws.sheetId, PRINT_AREA), nothing)
    isnothing(v) && return nothing
    x = v.value
    x isa SheetCellRange && return x.rng
    x isa SheetCellRef && return CellRange(x.cellref, x.cellref)
    x isa NonContiguousRange && return x
    s = string(x)
    return startswith(s, '=') ? s[2:end] : s
end

"""
    removePrintArea(ws::Worksheet) -> Bool

Remove a worksheet's print area, so that the whole sheet is printed. Returns
`false` if it had none.

The file must be open for writing (`mode="rw"`).
"""
function removePrintArea(ws::Worksheet)::Bool
    _require_writable(ws, "remove a print area")
    return !isnothing(pop!(get_workbook(ws).worksheet_names, (ws.sheetId, PRINT_AREA), nothing))
end

"""
    setPrintTitles(ws::Worksheet; rows=nothing, cols=nothing)

Repeat rows at the top of every printed page, and columns at the left, as
Excel's *Page Layout → Print Titles* does. `rows` is a row number, a range of
rows (`1:2`) or a string (`"1:2"`); `cols` is a column number, a range or a
string (`"A"`, `"A:B"`). Give either or both; they replace any titles the sheet
has.

Excel keeps the titles as a name scoped to the sheet, `_xlnm.Print_Titles`.

The file must be open for writing (`mode="rw"`).

# Examples
```julia
julia> XLSX.setPrintTitles(sheet; rows=1)              # repeat the header row

julia> XLSX.setPrintTitles(sheet; rows=1:2, cols="A")
```

See also [`getPrintTitles`](@ref) and [`removePrintTitles`](@ref).
"""
function setPrintTitles(ws::Worksheet; rows=nothing, cols=nothing)
    _require_writable(ws, "set print titles")
    _require_worksheet(ws, "set print titles")
    isnothing(rows) && isnothing(cols) &&
        throw(XLSXError("Give the rows or columns to repeat as `rows`, `cols` or both. To remove the titles, use `removePrintTitles`."))
    sheet = quoteit(ws.name)
    parts = String[]
    if !isnothing(cols)   # Excel writes the columns first
        c = _one_span(_col_spans(ws, cols), "columns")
        push!(parts, "$(sheet)!\$$(encode_column_number(first(c))):\$$(encode_column_number(last(c)))")
    end
    if !isnothing(rows)
        r = _one_span(_row_spans(ws, rows), "rows")
        push!(parts, "$(sheet)!\$$(first(r)):\$$(last(r))")
    end
    get_workbook(ws).worksheet_names[(ws.sheetId, PRINT_TITLES)] = DefinedNameValue("=" * join(parts, ","), false, false)
    return nothing
end

function _one_span(spans::Vector{UnitRange{Int}}, what::String)
    length(spans) == 1 || throw(XLSXError("Print titles must be adjacent $what. Found $(join(spans, ", "))."))
    return only(spans)
end

"""
    getPrintTitles(ws::Worksheet) -> Union{Nothing, NamedTuple}

The rows and columns repeated on every printed page, as a `NamedTuple`
`(rows, cols)` of `UnitRange`s (each `nothing` if not set), or `nothing` if the
sheet has no print titles. See [`setPrintTitles`](@ref).
"""
function getPrintTitles(ws::Worksheet)
    v = get(get_workbook(ws).worksheet_names, (ws.sheetId, PRINT_TITLES), nothing)
    isnothing(v) && return nothing
    rows = cols = nothing
    for part in split(lstrip(string(v.value), '='), ',')
        ref = replace(last(rsplit(part, '!'; limit=2)), "\$" => "")
        if (m = match(r"^(\d+):(\d+)$", ref)) !== nothing
            rows = parse(Int, m[1]):parse(Int, m[2])
        elseif (m = match(r"^([A-Za-z]+):([A-Za-z]+)$", ref)) !== nothing
            cols = decode_column_number(m[1]):decode_column_number(m[2])
        end
    end
    return (rows = rows, cols = cols)
end

"""
    removePrintTitles(ws::Worksheet) -> Bool

Stop repeating rows and columns on every printed page. Returns `false` if the
sheet had no print titles.

The file must be open for writing (`mode="rw"`).
"""
function removePrintTitles(ws::Worksheet)::Bool
    _require_writable(ws, "remove print titles")
    return !isnothing(pop!(get_workbook(ws).worksheet_names, (ws.sheetId, PRINT_TITLES), nothing))
end

# ===========================================================================
# Sheet XML helpers
# ===========================================================================

_sheet_root(ws::Worksheet) = xml_root_element(get_xml_data(get_xlsxfile(ws), _sheet_path(ws)))
_atts_vector(n::XML.Node) = Pair{String,String}[String(k) => String(v) for (k, v) in something(XML.attributes(n), Dict{String,String}())]

# `atts` with each change applied in place: a `nothing` value removes the attribute.
function _merge_atts(atts::Vector{Pair{String,String}}, changes)
    out = copy(atts)
    for (k, v) in changes
        i = findfirst(p -> first(p) == k, out)
        if isnothing(v)
            isnothing(i) || deleteat!(out, i)
        else
            isnothing(i) ? push!(out, k => v) : (out[i] = k => v)
        end
    end
    return out
end

# An element with these attributes and children. Built anew rather than changed
# in place: XML.jl can't remove an attribute, and an element made with no
# children takes none.
function _xml_element(tag::AbstractString, atts, kids=XML.Node[])
    el = isempty(kids) ? XML.Element(tag) : XML.Element(tag, kids...)
    for (k, v) in atts
        el[k] = v
    end
    return el
end

# The attributes of the sheet's top-level element `name`, or `nothing` if it has none.
function _sheet_element_atts(ws::Worksheet, name::String)
    els = _sheet_top_elements(get_xlsxfile(ws), _sheet_path(ws), (name,))
    return isempty(els) ? nothing : _atts_vector(first(els))
end

# Change the attributes of the sheet's top-level element `name`, keeping its
# children, or with `replace=true` set exactly `changes`. Made where the schema
# puts it if missing, and removed if left with nothing.
function _set_sheet_element!(ws::Worksheet, name::String, changes; replace::Bool=false)
    root = _sheet_root(ws)
    chn = XML.children(root)
    idx = findfirst(c -> localname(c) == name, chn)
    old = isnothing(idx) ? nothing : chn[idx]
    atts = _merge_atts(isnothing(old) || replace ? Pair{String,String}[] : _atts_vector(old), changes)
    kids = isnothing(old) || replace ? XML.Node[] : collect(XML.children(old))
    if isempty(atts) && isempty(kids)
        isnothing(idx) || deleteat!(chn, idx)
    else
        tag = isnothing(old) ? something(_ns_prefix(XML.tag(root)), "") * name : XML.tag(old)
        el = _xml_element(tag, atts, kids)
        isnothing(idx) ? insert!(chn, insert_index(root, name, WORKSHEET_ORDER) + 1, el) : (chn[idx] = el)
    end
    update_worksheets_xml!(get_xlsxfile(ws))
    return nothing
end

const _SHEET_PR_ORDER = ("tabColor", "outlinePr", "pageSetUpPr")

function _sheet_pr_child_atts(ws::Worksheet, name::String)
    for pr in _sheet_top_elements(get_xlsxfile(ws), _sheet_path(ws), ("sheetPr",)), c in xml_elements(pr)
        localname(c) == name && return _atts_vector(c)
    end
    return nothing
end

# Set the child `name` of the sheet's `<sheetPr>` to an element with attributes
# `atts`, or remove it if `atts` is `nothing`. `<sheetPr>` comes first in both
# worksheets and chartsheets, and its children in the schema's order.
function _set_sheet_pr_child!(ws::Worksheet, name::String, atts)
    root = _sheet_root(ws)
    chn = XML.children(root)
    idx = findfirst(c -> localname(c) == "sheetPr", chn)
    pr = isnothing(idx) ? nothing : chn[idx]
    pfx = something(_ns_prefix(XML.tag(isnothing(pr) ? root : pr)), "")
    kids = XML.Node[c for c in (isnothing(pr) ? XML.Node[] : XML.children(pr)) if localname(c) != name]
    isnothing(atts) || push!(kids, _xml_element(pfx * name, atts))
    sort!(kids; by = c -> something(findfirst(==(localname(c)), _SHEET_PR_ORDER), length(_SHEET_PR_ORDER) + 1))
    pr_atts = isnothing(pr) ? Pair{String,String}[] : _atts_vector(pr)
    if isempty(kids) && isempty(pr_atts)
        isnothing(idx) || deleteat!(chn, idx)
    else
        el = _xml_element(isnothing(pr) ? pfx * "sheetPr" : XML.tag(pr), pr_atts, kids)
        isnothing(idx) ? pushfirst!(chn, el) : (chn[idx] = el)
    end
    update_worksheets_xml!(get_xlsxfile(ws))
    return nothing
end

# "x:" for the tag "x:sheetPr", or nothing for an unprefixed one.
_ns_prefix(tag::AbstractString) = (i = findfirst(':', tag); isnothing(i) ? nothing : tag[1:i])
