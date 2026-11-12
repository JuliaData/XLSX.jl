#
# autofitColumns: column widths from the text cells show (see `getText`), measured with
# the character widths Excel 365 uses, from src/data/charwidths.tsv.
#
# Excel measures a column in widths of the widest digit of the workbook's default
# font, and its autofit leaves a little room beyond the text. Both were measured from
# Excel's own autofit; across 50 cells of different fonts, sizes, number formats and
# lengths, the widths found here are within 0.2 characters of Excel's at the median.
#

const _CHAR_WIDTHS_PATH = joinpath(@__DIR__, "..", "data", "charwidths.tsv")
include_dependency(_CHAR_WIDTHS_PATH)

function _read_char_widths(path::AbstractString)::Dict{Tuple{String,Bool},Vector{Float64}}
    widths = Dict{Tuple{String,Bool},Vector{Float64}}()
    for line in eachline(path)
        startswith(line, '#') && continue
        font, bold, ems = split(line, '\t')
        widths[(lowercase(font), bold == "true")] = parse.(Float64, split(ems))
    end
    return widths
end

const _CHAR_WIDTHS = _read_char_widths(_CHAR_WIDTHS_PATH)

# The room Excel's autofit leaves beyond the text, in digit widths (measured).
const _AUTOFIT_PADDING = 0.29

# The character widths of a font: its own, those of the longest family name it begins
# with (Calibri's for "Calibri Light"), or else Calibri's.
function _font_widths(name::AbstractString, bold::Bool)::Vector{Float64}
    key = lowercase(name)
    best, bestlen = _CHAR_WIDTHS[("calibri", bold)], 0
    for ((font, b), w) in _CHAR_WIDTHS
        b == bold && startswith(key, font) && length(font) > bestlen && ((best, bestlen) = (w, length(font)))
    end
    return best
end

# A character's width as a fraction of the font size. Characters outside the table are
# taken as wide as `n`, or as two digits if they take two columns (as CJK ones do).
function _char_em(widths::Vector{Float64}, c::Char)::Float64
    ' ' <= c <= '~' && return widths[Int(c) - 31]
    tw = textwidth(c)
    tw == 0 && return 0.0
    return tw >= 2 ? 2 * widths[Int('0') - 31] : widths[Int('n') - 31]
end

# The width of the widest line of `text`, as a fraction of the font size.
_text_em(widths::Vector{Float64}, text::AbstractString) =
    maximum(line -> sum(c -> _char_em(widths, c), line; init=0.0), eachsplit(text, '\n'))

# The font name, size and boldness of cell style `style`.
function _style_font(wb::Workbook, style::Integer)::Tuple{String,Float64,Bool}
    xf = styles_cell_xf(wb, style)
    fontid = haskey(xf, "fontId") ? parse(Int, xf["fontId"]) : 0
    name, size, bold = "Calibri", 11.0, false
    for c in xml_elements(get_fonts_nodes(wb)[fontid + 1])
        tag = localname(c)
        if tag == "name" && haskey(c, "val")
            name = c["val"]
        elseif tag == "sz" && haskey(c, "val")
            size = parse(Float64, c["val"])
        elseif tag == "b"
            bold = !haskey(c, "val") || c["val"] in ("1", "true")
        end
    end
    return name, size, bold
end

# The columns `cols` names, or `nothing` for all.
_autofit_columns(::Colon) = nothing
_autofit_columns(c::Integer) = Set(Int(c))
_autofit_columns(cs::AbstractVector{<:Integer}) = Set{Int}(cs)
function _autofit_columns(s::AbstractString)
    cols = Set{Int}()
    for part in split(s, ',')
        part = strip(part)
        if occursin(':', part)
            r = ColumnRange(part)
            union!(cols, r.start:r.stop)
        else
            is_valid_column_name(part) || throw(XLSXError("Invalid column: `$part`. Give columns as letters (`\"B\"`, `\"B:D\"`), numbers or `:`."))
            push!(cols, decode_column_number(part))
        end
    end
    return cols
end
_autofit_columns(cols) = throw(XLSXError("Columns are given as `:`, a number, a range or vector of numbers, or letters such as `\"B:D\"`, not `$(typeof(cols))`."))

"""
    autofitColumns(sh::Worksheet, [cols=:]; min_width=0, max_width=255) -> Dict{Int,Float64}

Set the width of columns to fit the text their cells show, as Excel's *Format →
AutoFit Column Width* does. The text is what [`XLSX.getText`](@ref) gives: a cell's
value in its number format, so a date or a number with separators is measured as it
is shown. Each cell's font, size and boldness are taken into account.

`cols` is `:` (every column with something in it), a column number, a range or vector
of numbers, or letters such as `"B"`, `"B:D"` or `"A,C:E"`. Widths are in characters,
as [`XLSX.setColumnWidth`](@ref) takes them, and are kept between `min_width` and
`max_width` (at most 255, Excel's widest). A column whose cells show nothing keeps its
width. Returns the width set for each column, by column number.

The widths are approximate. Characters are measured with the widths Excel 365 uses for
Calibri, Aptos, Aptos Narrow, Arial, Times New Roman and Courier New (other fonts as the
nearest of these by name, or else Calibri), which were measured at one screen
resolution; Excel's own autofit varies a little with the screen's. Cells in merged
ranges are left out, as Excel leaves them out. A cell with several lines is measured by
its longest. Rich text is measured in the cell's own font, and indents and rotated text
are not allowed for.

# Examples
```julia
julia> XLSX.autofitColumns(sh)                   # every column

julia> XLSX.autofitColumns(sh, "B:D"; max_width=40)

julia> XLSX.autofitColumns(sh, [1, 3]; min_width=8)
```
"""
function autofitColumns(ws::Worksheet, cols=:; min_width::Real=0, max_width::Real=255)::Dict{Int,Float64}
    _require_writable(ws, "autofit columns")
    0 <= min_width <= max_width <= 255 ||
        throw(XLSXError("Widths must have 0 ≤ `min_width` ≤ `max_width` ≤ 255 (Excel's widest column), got $min_width and $max_width."))
    wanted = _autofit_columns(cols)
    wb = get_workbook(ws)
    merged = something(getMergedCells(ws), CellRange[])
    normal_name, normal_size, _ = _style_font(wb, 0)
    normal = _font_widths(normal_name, false)
    digit = maximum(normal[Int('0')-31:Int('9')-31]) * normal_size    # Excel's unit of width, in points

    fonts = Dict{UInt32,Tuple{Vector{Float64},Float64}}()
    formats = Dict{String,Union{Nothing,Vector{_FmtSection}}}()
    widest = Dict{Int,Float64}()
    for row in eachrow(ws), (c, cell) in row.rowcells
        (isnothing(wanted) || c in wanted) || continue
        any(m -> cell.ref in m, merged) && continue
        text = _cell_text(ws, cell; for_width=true, formats)
        isempty(text) && continue
        widths, size = get!(fonts, cell.style) do
            name, size, bold = _style_font(wb, cell.style)
            (_font_widths(name, bold), size)
        end
        width = _text_em(widths, text) * size / digit + _AUTOFIT_PADDING
        widest[c] = max(get(widest, c, 0.0), width)
    end
    for (c, width) in widest
        widest[c] = round(clamp(width, min_width, max_width); digits=2)
        setColumnWidth(ws, CellRange(CellRef(1, c), CellRef(1, c)); width=widest[c])
    end
    return widest
end
