#
# Writing cell values: setdata! and setindex! for cells and ranges.
#

function add_cell_to_worksheet_dimension!(ws::Worksheet, cell::Cell)
    # update worksheet dimension
    ws_dimension = get_dimension(ws)

    top = row_number(ws_dimension.start)
    left = column_number(ws_dimension.start)

    bottom = row_number(ws_dimension.stop)
    right = column_number(ws_dimension.stop)

    r = row_number(cell)
    c = column_number(cell)

    if r < top || c < left || r > bottom || c > right
        set_dimension!(ws, CellRange(min(r, top), min(c, left), max(r, bottom), max(c, right)))
    end

    nothing
end

function setdata!(ws::Worksheet, cell::Cell)
    _require_writable(ws, "write to a cell")
    ws.cache === nothing && throw(XLSXError("Can't write data to a Worksheet with empty cache."))
    cache = ws.cache

    r = row_number(cell)
    c = column_number(cell)

    _ensure_cache_row!(ws, r)
    cache.cells[r][c] = cell
    add_cell_to_worksheet_dimension!(ws, cell)

    nothing
end
 
# Issue #284
# Whether `c` is a character `strip_illegal_chars` removes.
@inline function _is_illegal_xml_char(c::Char, strip_discouraged::Bool)::Bool
    isvalid(c) || return false
    u = UInt32(c)
    (u <= 0x1F && u != 0x09 && u != 0x0A && u != 0x0D) && return true
    (u == 0xFFFE || u == 0xFFFF) && return true
    strip_discouraged || return false
    return (0xFDD0 <= u <= 0xFDEF) || (u >= 0x1FFFE && (u & 0xFFFE) == 0xFFFE)
end

function strip_illegal_chars(s::String; strip_discouraged::Bool=true)
    # Nearly every string is clean: one pass to confirm it, rather than the
    # many full-string `replace`s below.
    any(c -> _is_illegal_xml_char(c, strip_discouraged), s) || return s
    # remove XML 1.0 illegal C0 controls and U+0000
    s = replace(s, r"[\x00-\x08\x0B\x0C\x0E-\x1F]" => "")
    # remove BMP non-characters U+FFFE U+FFFF
    s = replace(s, "\uFFFE" => "", "\uFFFF" => "")
    # optionally remove U+FDD0..U+FDEF and plane-ending non-characters
    if strip_discouraged
        s = replace(s, r"[\uFDD0-\uFDEF]" => "")
        for cp in (0x1FFFE:0x10000:0x10FFFF)
            # remove both XFFFE and XFFFF for each plane
            s = replace(s, Char(cp) => "")
            s = replace(s, Char(cp+1) => "")
        end
    end
    return s
end

# Returns the datatype and value for `val` to be inserted into `ws`.
function xlsx_encode(ws::Worksheet, val::AbstractString)
    if isempty(val)
        return (CT_EMPTY, UInt64(0))
    end
    sst_ind = add_shared_string!(get_workbook(ws), strip_illegal_chars(val))
    ws.sst_count+=1

    return (CT_STRING, UInt64(sst_ind))
end

xlsx_encode(::Worksheet, val::Missing) = (CT_EMPTY, UInt64(0))
xlsx_encode(::Worksheet, val::Bool) = (CT_BOOL, UInt64(val))
xlsx_encode(::Worksheet, val::Float64) = (CT_FLOAT, reinterpret(UInt64, val))
xlsx_encode(::Worksheet, val::Int64) = (CT_INT, reinterpret(UInt64, val))
xlsx_encode(ws::Worksheet, val::Dates.Date) = (CT_DATE, reinterpret(UInt64, date_to_excel_value(val, isdate1904(get_xlsxfile(ws)))))
xlsx_encode(ws::Worksheet, val::Dates.DateTime) = (CT_DATETIME, reinterpret(UInt64, datetime_to_excel_value(val, isdate1904(get_xlsxfile(ws)))))
xlsx_encode(::Worksheet, val::Dates.Time) = (CT_TIME, reinterpret(UInt64, time_to_excel_value(val)))
# `ws[ref] = v` and `ws[r, c] = v`. A vector is written along a row or column; any other
# value goes in every cell addressed. Each has one entry point, the vector one more
# specific, which then dispatches on the indices; typing the value and the indices on
# the same methods made most vector forms ambiguous.
Base.setindex!(ws::Worksheet, v::AbstractVector, ref; dim::Integer=2) = setdata!(ws, ref, v, dim)
Base.setindex!(ws::Worksheet, v, ref) = setdata!(ws, ref, v)
Base.setindex!(ws::Worksheet, v::AbstractVector, r, c; dim::Integer=2) = _setindex_vector!(ws, v, r, c, dim)
Base.setindex!(ws::Worksheet, v, r, c) = _setindex_value!(ws, v, r, c)

const _RowsOrCols = Union{Integer,UnitRange{<:Integer}}
const _SparseRowsOrCols = Union{Vector{Int},StepRange{<:Integer}}

# A vector from an anchor cell (along `dim`, as for `ws[ref] = v`), or along one row
# or column.
_setindex_vector!(ws::Worksheet, v, r::Integer, c::Integer, dim) = setdata!(ws, r, c, v, dim)
_setindex_vector!(ws::Worksheet, v, r::Integer, c::Union{Colon,UnitRange{<:Integer}}, dim) = setdata!(ws, r, c, v)
_setindex_vector!(ws::Worksheet, v, r::Union{Colon,UnitRange{<:Integer}}, c::Integer, dim) = setdata!(ws, r, c, v)
_setindex_vector!(ws::Worksheet, v, r, c, dim) = throw(XLSXError(
    "Can't write a vector to `ws[$r, $c]`: a vector is written along one row or column " *
    "(`ws[row, :]`, `ws[:, col]`, `ws[row, c1:c2]`, `ws[r1:r2, col]`) or across from one cell (`ws[row, col]`)."))

# Any other value, in every cell addressed.
_setindex_value!(ws::Worksheet, v, row::Integer, col::Integer) = setdata!(ws, CellRef(row, col), v)
_setindex_value!(ws::Worksheet, v, row::_RowsOrCols, col::_RowsOrCols) = setdata!(ws, CellRange(CellRef(first(row), first(col)), CellRef(last(row), last(col))), v)
function _setindex_value!(ws::Worksheet, v, row::Union{_RowsOrCols,_SparseRowsOrCols}, col::Union{_RowsOrCols,_SparseRowsOrCols})
    for a in row, b in col
        setdata!(ws, CellRef(a, b), v)
    end
end
_setindex_value!(ws::Worksheet, v, r, c) = setdata!(ws, r, c, v)

function setdata!(ws::Worksheet, ref::CellRef, val::CellFormula)
    c = getcell(ws, ref)
    if !(c isa EmptyCell) && c.formula && get_formula_from_cache(ws, ref) isa ReferencedFormula
        rereference_formulae(ws, c)
    end
    v = ""
    t = ""
    m = ""
    cell = Cell(get_workbook(ws), ref, t, string(id(val.styleid)), v, m, true)
    add_formula_to_cache(ws, ref, val.value)
    setdata!(ws, cell)
end
function setdata!(ws::Worksheet, ref::CellRef, val::CellValue, convert_to_string::Bool)
    # Convert Float64 values NaN, Inf and -Inf to strings. Excel can't handle them as floats
    # Addresses #179 & #342
    @assert convert_to_string == true "Converting values other than NaN, -Inf, Inf is not permitted"
    val = CellValue(string(val.value), val.styleid)
    setdata!(ws, ref, val)
end
function setdata!(ws::Worksheet, ref::CellRef, val::CellValue)
    if val.value isa Float64 && (isnan(val.value) || isinf(val.value))
        return setdata!(ws, ref, val, true)
    end
    t, v = xlsx_encode(ws, val.value)
    m = UInt64(0)
    cell = Cell(ref, v, id(val.styleid), m, t, false)
    setdata!(ws, cell)
end

# convert AbstractTypes to concrete
setdata!(ws::Worksheet, ref::CellRef, val::AbstractString) = setdata!(ws, ref, CellValue(ws, convert(String, val)))
setdata!(ws::Worksheet, ref::CellRef, val::Integer) = setdata!(ws, ref, convert(Int64, val))
setdata!(ws::Worksheet, ref::CellRef, val::Real) = setdata!(ws, ref, convert(Float64, val))

function setdata!(sheet::Worksheet, ref::CellRange, data::AbstractVector) # vector to row or column range
    rows = ref.start.row_number:ref.stop.row_number
    cols = ref.start.column_number:ref.stop.column_number
    length(rows)==1 || length(cols)==1 || throw(XLSXError("Can only write 1D data to a row or column range."))
    if length(cols)==1
        length(rows) != length(data) && throw(XLSXError("Row count mismatch between `data` ($(length(data)) rows) and row range $rows ($(length(rows)) rows)."))
        anchor_cell_ref = CellRef(first(rows), first(cols))
        setdata!(sheet, anchor_cell_ref, data, 1)
    elseif length(rows)==1
        length(cols) != length(data) && throw(XLSXError("Column count mismatch between `data` ($(length(data)) columns) and column range $cols ($(length(cols)) columns)."))
        anchor_cell_ref = CellRef(first(rows), first(cols))
        setdata!(sheet, anchor_cell_ref, data, 2)
    end
end

# convert nothing to missing when writing
setdata!(ws::Worksheet, ref::CellRef, ::Nothing) = setdata!(ws, ref, CellValue(ws, missing))

setdata!(ws::Worksheet, row::Integer, col::Integer, val::CellValue) = setdata!(ws, CellRef(row, col), val)

setdata!(ws::Worksheet, row::Union{Integer,UnitRange{<:Integer}}, col::Union{Integer,UnitRange{<:Integer}}, v) = setdata!(ws, CellRange(CellRef(first(row), first(col)), CellRef(last(row), last(col))), v)

function setdata!(ws::Worksheet, ref::CellRef, val::Union{AbstractFormula,CellConcreteType}) # use existing cell format if it exists
    c = getcell(ws, ref)
    if !(c isa EmptyCell) && c.formula && get_formula_from_cache(ws, ref) isa ReferencedFormula
        rereference_formulae(ws, c)
    end
    if c isa EmptyCell || c.style == 0
        if val isa AbstractFormula
            return setdata!(ws, ref, CellFormula(ws, val))
        else
            return setdata!(ws, ref, CellValue(ws, val))
        end
    else
        c.style = _style_for_value(ws, c.style, val)
        if val isa AbstractFormula
            return setdata!(ws, ref, CellFormula(val, CellDataFormat(c.style)))
        else
            return setdata!(ws, ref, CellValue(val, CellDataFormat(c.style)))
        end
    end
end
# The style for `val` in a cell of style `style`, keeping its other formatting: a date
# or time takes a date or time format if it has none, and a number or `Bool` loses a
# date or time format, which would misdisplay it.
function _style_for_value(ws::Worksheet, style::Integer, val)::UInt32
    style == 0 && return id(default_cell_format(ws, val))
    existing_style = CellDataFormat(style)
    isa_dt = styles_is_datetime(get_workbook(ws), existing_style)
    numfmt = if val isa Dates.Date
        isa_dt ? nothing : DEFAULT_DATE_numFmtId
    elseif val isa Dates.Time
        isa_dt ? nothing : DEFAULT_TIME_numFmtId
    elseif val isa Dates.DateTime
        isa_dt ? nothing : DEFAULT_DATETIME_numFmtId
    elseif val isa Float64 || val isa Int64
        isa_dt ? DEFAULT_NUMBER_numFmtId : nothing
    elseif val isa Bool
        isa_dt ? DEFAULT_BOOL_numFmtId : nothing
    else
        nothing
    end
    isnothing(numfmt) && return style
    return update_template_xf(ws, existing_style, ["numFmtId", "applyNumberFormat"], [string(numfmt), "1"]).id
end

# A value as a cell holds it: `missing` for none, an `Int64`, `Float64`, `Bool`,
# `String`, date or time, or else its `string`. Excel has no NaN or Inf, so they are
# written as text.
_as_cell_value(::Union{Missing,Nothing}) = missing
_as_cell_value(v::Union{Bool,String,Dates.Date,Dates.DateTime,Dates.Time}) = v
_as_cell_value(v::Integer) = Int64(v)
_as_cell_value(v::Real) = (f = Float64(v); isfinite(f) ? f : string(v))
_as_cell_value(v::AbstractString) = String(v)
_as_cell_value(v::CellImage) = v
_as_cell_value(v) = string(v)

function setdata!(ws::Worksheet, ref::AbstractString, value)
    # A string starting with `=` is stored as text, not as a formula: formulas are
    # set only with `setFormula`.
    if is_worksheet_defined_name(ws, ref)
        v = get_defined_name_value(ws, ref)
        if is_defined_name_value_a_reference(v)
            return setdata!(ws, v, value)
        else
            throw(XLSXError("`$ref` is not a valid cell or range reference."))
        end
    elseif is_workbook_defined_name(get_workbook(ws), ref)
        wb = get_workbook(ws)
        v = get_defined_name_value(wb, ref)
        if is_defined_name_value_a_reference(v)
            return setdata!(ws, v, value)
        else
            throw(XLSXError("`$ref` is not a valid cell or range reference."))
        end
    elseif is_valid_cellname(ref)
        return setdata!(ws, CellRef(ref), value)
    elseif is_valid_sheet_cellname(ref)
        return setdata!(ws, SheetCellRef(ref), value)
    elseif is_valid_cellrange(ref)
        return setdata!(ws, CellRange(ref), value)
    elseif is_valid_column_range(ref)
        return setdata!(ws, ColumnRange(ref), value)
    elseif is_valid_row_range(ref)
        return setdata!(ws, RowRange(ref), value)
    elseif is_valid_non_contiguous_range(ref)
        return setdata!(ws, NonContiguousRange(ws, ref), value)
    elseif is_valid_sheet_cellrange(ref)
        return setdata!(ws, SheetCellRange(ref), value)
    elseif is_valid_sheet_column_range(ref)
        return setdata!(ws, SheetColumnRange(ref), value)
    elseif is_valid_sheet_row_range(ref)
        return setdata!(ws, SheetRowRange(ref), value)
    elseif is_valid_non_contiguous_cellrange(ref)
        return setdata!(ws, NonContiguousRange(ws, ref), value)
    end
    throw(XLSXError("`$ref` is not a valid cell or range reference."))
end
function setdata!(ws::Worksheet, rng::CellRange, value)
    for row in rng.start.row_number:rng.stop.row_number
        for col in rng.start.column_number:rng.stop.column_number
            setdata!(ws, row, col, value)
        end
    end
end
function setdata!(ws::Worksheet, rng::RowRange, value)
    dim = get_dimension(ws)
    start = CellRef(rng.start, dim.start.column_number,)
    stop = CellRef(rng.stop, dim.stop.column_number)
    setdata!(ws, CellRange(start, stop), value)
end
function setdata!(ws::Worksheet, rng::ColumnRange, value)
    dim = get_dimension(ws)
    start = CellRef(dim.start.row_number, rng.start)
    stop = CellRef(dim.stop.row_number, rng.stop)
    setdata!(ws, CellRange(start, stop), value)
end
function setdata!(ws::Worksheet, rng::NonContiguousRange, value)
    do_sheet_names_match(ws, rng)
    for r in rng.rng
        setdata!(ws, r, value) # r may be a single Cell or a CellRange
    end
end
setdata!(ws::Worksheet, ::Colon, ::Colon, v) = setdata!(ws::Worksheet, :, v)
function setdata!(ws::Worksheet, ::Colon, v)
    dim = get_dimension(ws)
    setdata!(ws, dim, v)
end
function setdata!(ws::Worksheet, row::Union{Integer,UnitRange{<:Integer}}, ::Colon, v)
    dim = get_dimension(ws)
    setdata!(ws, CellRange(CellRef(first(row), dim.start.column_number), CellRef(last(row), dim.stop.column_number)), v)
end
function setdata!(ws::Worksheet, ::Colon, col::Union{Integer,UnitRange{<:Integer}}, v)
    dim = get_dimension(ws)
    setdata!(ws, CellRange(CellRef(dim.start.row_number, first(col)), CellRef(dim.stop.row_number, last(col))), v)
end
function setdata!(ws::Worksheet, row::Union{Vector{Int},StepRange{<:Integer}}, ::Colon, v)
    dim = get_dimension(ws)
    for a in row
        for b in dim.start.column_number:dim.stop.column_number
            setdata!(ws, CellRef(a, b), v)
        end
    end
end
function setdata!(ws::Worksheet, ::Colon, col::Union{Vector{Int},StepRange{<:Integer}}, v)
    dim = get_dimension(ws)
    for b in col
        for a in dim.start.row_number:dim.stop.row_number
            setdata!(ws, CellRef(a, b), v)
        end
    end
end
setdata!(ws::Worksheet, ref::SheetCellRef, value) = do_sheet_names_match(ws, ref) && setdata!(ws, ref.cellref, value)
setdata!(ws::Worksheet, rng::SheetCellRange, value) = do_sheet_names_match(ws, rng) && setdata!(ws, rng.rng, value)
setdata!(ws::Worksheet, rng::SheetColumnRange, value) = do_sheet_names_match(ws, rng) && setdata!(ws, rng.colrng, value)
setdata!(ws::Worksheet, rng::SheetRowRange, value) = do_sheet_names_match(ws, rng) && setdata!(ws, rng.rowrng, value)
setdata!(ws::Worksheet, row::Integer, col::Integer, data) = setdata!(ws, CellRef(row, col), data)
setdata!(ws::Worksheet, ref::CellRef, value) = throw(XLSXError("$value: Unsupported datatype $(typeof(value)) for writing data to Excel file. Supported data types are $(CellConcreteType) or $(CellValue)."))
setdata!(ws::Worksheet, row::Integer, col::Integer, data::AbstractVector, dim::Integer) = setdata!(ws, CellRef(row, col), data, dim)
function setdata!(ws::Worksheet, ref_str::AbstractString, value::AbstractVector, dim::Integer)
    if is_valid_cellrange(ref_str)
        setdata!(ws, CellRange(ref_str), value)
    elseif is_valid_cellname(ref_str)
        setdata!(ws, CellRef(ref_str), value, dim)
    else
        throw(XLSXError("Invalid cell reference or range: $ref_str"))
    end
end

function setdata!(sheet::Worksheet, ref::CellRef, data::AbstractVector, dim::Integer)
    for (i, val) in enumerate(data)
        target_cell_ref = target_cell_ref_from_offset(ref, i - 1, dim)
        setdata!(sheet, target_cell_ref, val)
    end
end

setdata!(sheet::Worksheet, ::Colon, col::Integer, data::AbstractVector) = setdata!(sheet, 1, col, data, 1)
setdata!(sheet::Worksheet, row::Integer, ::Colon, data::AbstractVector) = setdata!(sheet, row, 1, data, 2)

function setdata!(sheet::Worksheet, row::Integer, cols::UnitRange{T}, data::AbstractVector) where {T<:Integer}
    length(data) != length(cols) && throw(XLSXError("Column count mismatch between `data` ($(length(data)) columns) and column range $cols ($(length(cols)) columns)."))
    anchor_cell_ref = CellRef(row, first(cols))

    # since cols is the unit range, this is a column-based operation
    setdata!(sheet, anchor_cell_ref, data, 2)
end

function setdata!(sheet::Worksheet, rows::UnitRange{T}, col::Integer, data::AbstractVector) where {T<:Integer}
    length(data) != length(rows) && throw(XLSXError("Row count mismatch between `data` ($(length(data)) rows) and row range $rows ($(length(rows)) rows)."))
    anchor_cell_ref = CellRef(first(rows), col)

    # since rows is the unit range, this is a row-based operation
    setdata!(sheet, anchor_cell_ref, data, 1)
end

function setdata!(sheet::Worksheet, ref_or_rng::AbstractString, matrix::AbstractArray{T,2}) where {T}
    if is_valid_cellrange(ref_or_rng)
        setdata!(sheet, CellRange(ref_or_rng), matrix)
    elseif is_valid_cellname(ref_or_rng)
        setdata!(sheet, CellRef(ref_or_rng), matrix)
    else
        throw(XLSXError("Invalid cell reference or range: $ref_or_rng"))
    end
end

# Generalise to AbstractArray (#158)
function setdata!(sheet::Worksheet, ref::CellRef, matrix::AbstractArray{T,2}) where {T}
    row_ax, col_ax = axes(matrix)
    anchor_row = row_number(ref)
    anchor_col = column_number(ref)

    @inbounds for c in col_ax, r in row_ax
        setdata!(sheet, anchor_row + (r - first(row_ax)), anchor_col + (c - first(col_ax)), matrix[r, c])
    end
end

function setdata!(sheet::Worksheet, rng::CellRange, matrix::AbstractArray{T,2}) where {T}
    size(rng) != size(matrix) && throw(XLSXError("Target range $rng size ($(size(rng))) must be equal to the input matrix size ($(size(matrix)))"))
    setdata!(sheet, rng.start, matrix)
end

function setdata!(sheet::Worksheet, ref::CellRef, rts::RichTextString)
    isempty(rts.text) && throw(XLSXError("Cannot set a cell value to an empty RichTextString."))
    c=getcell(sheet, ref)
    if c isa EmptyCell
        sheet[ref]=""
        c=getcell(sheet, ref)
    end

    pfx = get_prefix(sheet)
    pfx = pfx == "" ? pfx : pfx * ":"

    if length(rts.runs) > 1 # add RichTextString as a rich sharedString
        sheet.sst_count += 1
        c.value = add_formatted_string!(get_workbook(sheet), richTextStringtoXML(rts, pfx))
        c.datatype = CT_STRING
    else # add single run as a normal cell value with a cell level font style
        sheet[ref] = rts.text
        kw = _run_font_kwargs(rts.runs[1].atts)
        isempty(kw) || setFont(sheet, ref; kw...)
    end

    return sheet[ref]
end

