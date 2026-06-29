
@inline Base.isempty(::EmptyCell) = true
@inline Base.isempty(::AbstractCell) = false
@inline row_number(c::EmptyCell) = row_number(c.ref)
@inline column_number(c::EmptyCell) = column_number(c.ref)
@inline row_number(c::Cell) = row_number(c.ref)
@inline column_number(c::Cell) = column_number(c.ref)
@inline relative_cell_position(c::Cell, rng::CellRange) = relative_cell_position(c.ref, rng)
@inline relative_cell_position(c::EmptyCell, rng::CellRange) = relative_cell_position(c.ref, rng)
@inline relative_column_position(c::Cell, rng::ColumnRange) = relative_column_position(c.ref, rng)
@inline relative_column_position(c::EmptyCell, rng::ColumnRange) = relative_column_position(c.ref, rng)

Base.:(==)(c1::Cell, c2::Cell) = c1.ref == c2.ref && c1.datatype == c2.datatype && c1.style == c2.style && c1.value == c2.value && c1.meta == c2.meta && c1.formula == c2.formula
Base.hash(c::Cell, h::UInt) = hash(c.formula, hash(c.meta, hash(c.value, hash(c.style, hash(c.datatype, hash(c.ref, h))))))

Base.:(==)(c1::EmptyCell, c2::EmptyCell) = c1.ref == c2.ref
Base.hash(c::EmptyCell, h::UInt) = hash(c.ref, h)

const RGX_INTEGER = r"^\-?[0-9]+$"

const ERROR_STRING_TO_CODE = Dict{String, UInt64}(
    "#NULL!"  => UInt64(XL_NULL),
    "#DIV/0!" => UInt64(XL_DIV0),
    "#VALUE!" => UInt64(XL_VALUE),
    "#REF!"   => UInt64(XL_REF),
    "#NAME?"  => UInt64(XL_NAME),
    "#NUM!"   => UInt64(XL_NUM),
    "#N/A"    => UInt64(XL_NA),
    "#SPILL!" => UInt64(XL_SPILL), # Won't happen - #SPILL isn't an actual error. Returns #VALUE! instead.
)

const ERROR_CODE_TO_STRING = Dict{UInt64, String}(v => k for (k, v) in ERROR_STRING_TO_CODE)

function get_error_type(v::AbstractString)::UInt64
    get(ERROR_STRING_TO_CODE, v) do
        throw(XLSXError("Unknown error value: $v"))
    end
end

function get_error_string(e::UInt64)::String
    get(ERROR_CODE_TO_STRING, e) do
        throw(XLSXError("Unknown error code: $e"))
    end
end

get_error_string(::Nothing) = ""

"""
    iserror(s::Worksheet, ref::AbstractString)
    iserror(s::Worksheet, rows, cols)

Returns `true` if the cell(s) at the given reference contain an error, `false` otherwise.
An `EmptyCell` is not considered an error and returns `false`.

The return type depends on the type of `ref` and is the same shape as the return 
type of `getcell` for the same `ref`:
- If `ref` or (`row`, `col`) refers to a single cell, returns a single Bool.
- If `ref` or (`rows`, `cols`) refers to a range of cells, returns a matrix of Bools.
- If `ref` or (`rows`, `cols`) refers to a non-contiguous range of cells, returns a vector of matrices of Bools.

# Examples
```julia
julia> XLSX.iserror(sh, "A1") # Cell
true

julia> XLSX.iserror(sh, "I1") # EmptyCell
false

julia> XLSX.iserror(sh, "A1:I1") # CellRange - note that I1 is an EmptyCell, which is not an error
1×9 Matrix{Bool}:
 1  1  1  1  1  1  1  1  0

julia> XLSX.iserror(sh, "A1:B1,D1:E1") # non-contiguous range
2-element Vector{Matrix{Bool}}:
 [1 1]
 [1 1]
```

See also [`XLSX.geterror`](@ref), [`XLSX.getcell`](@ref).
"""
iserror(s::Worksheet, ref::AbstractString) = iserror(getcell(s, ref))
iserror(s::Worksheet, ::Colon) = iserror(getcell(s, :))
iserror(s::Worksheet, r, c) = iserror(getcell(s, r, c))
iserror(c::AbstractVector) = collect(iserror.(c))
iserror(c::AbstractMatrix) = collect(iserror.(c))
iserror(c::AbstractVector{<:AbstractMatrix}) = [collect(iserror.(M)) for M in c]
@inline iserror(c::Cell) = c.datatype == CT_ERROR
@inline iserror(::EmptyCell) = false

getval(x) = hasproperty(x, :datatype) && x.datatype == CT_ERROR && hasproperty(x, :value) ? x.value : nothing

"""
    geterror(s::Worksheet, ref::AbstractString)
    geterror(s::Worksheet, rows, cols)

Returns the error value (e.g. `#DIV/0!`) for the cell(s) at the given reference, if any. 
If there is no error, returns an empty string.

The return type depends on the type of `ref` and is the same shape as the return 
type of `getcell` for the same `ref`:
- If `ref` or (`row`, `col`) refers to a single cell, returns a single Bool.
- If `ref` or (`rows`, `cols`) refers to a range of cells, returns a matrix of Bools.
- If `ref` or (`rows`, `cols`) refers to a non-contiguous range of cells, returns a vector of matrices of Bools.

# Examples
```julia
julia> XLSX.geterror(sh, "A1") # Cell
"#NULL!"

julia> XLSX.geterror(sh, "I1") # EmptyCell
""

julia> XLSX.geterror(sh, "A1:I1") # CellRange - note that I1 is an EmptyCell, which returns an empty string
1×9 Matrix{String}:
 "#NULL!"  "#DIV/0!"  "#VALUE!"  "#REF!"  "#NAME?"  "#NUM!"  "#N/A"  "#VALUE!"  ""

julia> XLSX.geterror(sh, "A1:B1,D1:E1") # non-contiguous range
2-element Vector{Matrix{String}}:
 ["#NULL!" "#DIV/0!"]
 ["#REF!" "#NAME?"]
```

See also [`XLSX.iserror`](@ref), [`XLSX.getcell`](@ref).
"""
geterror(s::Worksheet, ref::AbstractString) = geterror(getcell(s, ref))
geterror(s::Worksheet, ::Colon) = geterror(getcell(s, :))
geterror(s::Worksheet, r, c) = geterror(getcell(s, r, c))
geterror(c::AbstractVector) = collect(geterror.(c))
geterror(c::AbstractMatrix) = collect(geterror.(c))
geterror(c::AbstractVector{<:AbstractMatrix}) = [collect(geterror.(M)) for M in c]
geterror(c::AbstractCell) = get_error_string(getval(c))

# Extracts the unformatted text from an inlineStr "is" XML element as a <si> XML string.
function _rewrite_node(io::IOBuffer, node::XML.LazyNode, pfx::String)
    XML.nodetype(node) == XML.Element || return nothing
    tag = localname(node)

    write(io, '<', pfx, tag)
    for (k, v) in XML.eachattribute(node)
        write(io, ' ', k, '=', '"', v, '"')
    end

    first_child = iterate(XML.eachchildnode(node))

    if isnothing(first_child)
        txt = XML.value(node)
        if txt isa AbstractString && !isempty(txt)
            write(io, '>', txt, '<', '/', pfx, tag, '>')
        else
            write(io, '/', '>')
        end
    elseif tag == "t"
        write(io, '>')
        for child in XML.eachchildnode(node)
            sv = XML.is_simple_value(child)
            if !isnothing(sv)
                write(io, sv)
            else
                v = XML.value(child)
                !isnothing(v) && write(io, v)
            end
        end
        write(io, '<', '/', pfx, tag, '>')
    else
        write(io, '>')
        for child in XML.eachchildnode(node)
            XML.nodetype(child) == XML.Element || continue
            write(io, "\n  ")
            _rewrite_node(io, child, pfx)
        end
        write(io, '\n', '<', '/', pfx, tag, '>')
    end
    return nothing
end

function _rewrite_node(node::XML.LazyNode, pfx::Union{String,Nothing})::String
    io = IOBuffer()
    _rewrite_node(io, node, something(pfx, ""))
    return String(take!(io))
end

function _build_si_xml(si_node::XML.LazyNode, pfx::String)::String
    io = IOBuffer()
    prefix_part = isempty(pfx) ? "si" : "$(pfx)si"
    write(io, '<', prefix_part, '>')
    for child in XML.eachchildnode(si_node)
        XML.nodetype(child) == XML.Element || continue
        write(io, "\n  ")
        _rewrite_node(io, child, pfx)
    end
    write(io, '\n', '<', '/', prefix_part, '>')
    return String(take!(io))
end

# Parses a style string to (UInt32, Int) for use as style and num_style.
function _parse_style(s::AbstractString)
    isempty(s) && return UInt32(0), 0
    n = parse(Int, s)
    return UInt32(n), n
end

# Resolves unhandled_attributes to nothing if empty, for compact Formula construction.
_extra_attrs(d::Dict) = isempty(d) ? nothing : d

function Cell(c::XML.LazyNode, ws::Worksheet, sst_pfx::String,
              local_formulas::Union{Nothing, Dict{SheetCellRef, AbstractFormula}}=nothing,
              load_formulas::Bool=true)::Union{Cell,EmptyCell}
    wb = get_workbook(ws)
    @assert localname(c) == "c" "`Cell` expects a `c` (cell) XML node."
    ref_str::Union{SubString{String},String} = ""
    t::Union{SubString{String},String}       = ""
    s_str::Union{SubString{String},String}   = ""
    m_str::Union{SubString{String},String}   = ""
    XML.foreach_attr(c) do name_tok, val_tok
        k = XML.XMLTokenizer.raw(name_tok, c.data)
        if k == "r";      ref_str = XML.XMLTokenizer.attr_value(val_tok, c.data)
        elseif k == "t";  t       = XML.XMLTokenizer.attr_value(val_tok, c.data)
        elseif k == "s";  s_str   = XML.XMLTokenizer.attr_value(val_tok, c.data)
        elseif k == "cm"; m_str   = XML.XMLTokenizer.attr_value(val_tok, c.data)
        end
    end
    ref   = CellRef(ref_str)
    style, num_style = _parse_style(s_str)
    meta::UInt32 = isempty(m_str) ? UInt32(0) : parse(UInt32, m_str)
    datatype::CellValueType = CT_EMPTY
    value::UInt64           = UInt64(0)
    formula::Bool           = false
    for child in XML.eachchildnode(c)
        XML.nodetype(child) == XML.Element || continue
        tag = localname(child)
        if t == "inlineStr"
            tag == "is" || continue
            uft = unformatted_text(wb, child)
            if !isempty(uft)
                ft = _build_si_xml(child, sst_pfx)
                datatype = CT_STRING
                value = reinterpret(UInt64, Int64(add_formatted_string!(wb, ft)))
            end
            break
        else
            if tag == "v"
                sv = XML.is_simple_value(child)
                if !isnothing(sv) && !isempty(sv)
                    datatype, value = process_tv(wb, t, sv, num_style)
                end
            elseif tag == "f"
                if load_formulas
                    f = parse_formula_from_element(wb, child)
                    if isnothing(local_formulas)
                        # streaming path — write directly under lock
                        lock(wb.formulas_lock) do
                            wb.formulas[SheetCellRef(combine_sheet_ref(ws, ref))] = f
                        end
                    else
                        # cache fill path — write to local dict, merged later
                        local_formulas[SheetCellRef(combine_sheet_ref(ws, ref))] = f
                    end
                end
                formula = true
            end
        end
    end
    return Cell(ref, value, style, meta, datatype, formula)
end

function parse_formula_from_element(wb, c_child_element)::AbstractFormula
    localname(c_child_element) == "f" ||
        throw(XLSXError("Expected nodename `f`. Found: `$(localname(c_child_element))`"))

    # Single pass over attributes — no Attributes dict allocated
    t_val  = nothing
    si_val = nothing
    ref_val = nothing
    unhandled = Dict{String,String}()
    for (k, v) in XML.eachattribute(c_child_element)
        if k == "t"
            t_val = String(v)
        elseif k == "si"
            si_val = String(v)
        elseif k == "ref"
            ref_val = String(v)
        else
            unhandled[String(k)] = String(v)
        end
    end

    # Extract text content — is_simple_value handles the common no-attribute case,
    # fall back to eachchildnode for <f> elements that have attributes
    formula_string = something(XML.is_simple_value(c_child_element), "")
    if isempty(formula_string)
        for ch in XML.eachchildnode(c_child_element)
            if XML.nodetype(ch) === XML.Text
                v = XML.value(ch)
                isnothing(v) || (formula_string = v)
                break
            end
        end
    end

    if !isnothing(t_val)
        if t_val == "shared"
            isnothing(si_val) && throw(XLSXError("Expected shared formula to have an index. `si` attribute is missing: $c_child_element"))
            si = parse(Int, si_val)
            extra = _extra_attrs(unhandled)
            return isnothing(ref_val) ?
                FormulaReference(si, extra) :
                ReferencedFormula(formula_string, si, ref_val, extra)
        elseif t_val == "array"
            return Formula(formula_string, "array", ref_val, _extra_attrs(unhandled))
        end
    end

    return Formula(formula_string, nothing, ref_val, _extra_attrs(unhandled))
end

# Returns (raw_value::UInt64, datatype::CellValueType) for datetime strings,
# keeping the value in its Excel numeric form for storage in Cell.
function _parse_excel_datetime_raw(v::AbstractString)
    isempty(v) && throw(XLSXError("Cannot convert an empty string into a datetime value."))
    if occursin('.', v) || v == "0"
        time_value = parse(Float64, v)
        time_value >= 0 || throw(XLSXError("Cannot have a datetime value < 0. Got $time_value"))
        datatype = time_value < 1.0 ? CT_TIME : CT_DATETIME
        return reinterpret(UInt64, time_value), datatype
    else
        date_value = parse(Int64, v)
        date_value >= 0 || throw(XLSXError("Cannot have a datetime value < 0. Got $date_value"))
        return reinterpret(UInt64, date_value), CT_DATE
    end
end

function process_tv(wb::Workbook, t::AbstractString, v::AbstractString, num_style::Int)
    datatype::CellValueType = CT_EMPTY
    value::UInt64           = UInt64(0)
    isempty(v) && return datatype, value

    if t == "n" || t == ""
        if styles_is_datetime(wb, num_style)
            value, datatype = _parse_excel_datetime_raw(v)
        elseif styles_is_float(wb, num_style)
            datatype = CT_FLOAT
            value = reinterpret(UInt64, parse(Float64, v))
        else
            parsed_int = tryparse(Int64, v)
            if parsed_int !== nothing
                datatype = CT_INT
                value = reinterpret(UInt64, parsed_int)
            else
                datatype = CT_FLOAT
                value = reinterpret(UInt64, parse(Float64, v))
            end
        end

    elseif t == "s"
        parsed = tryparse(Int64, v)
        parsed === nothing && throw(XLSXError("Expected SST index in cell value, got: $v"))
        datatype = CT_STRING
        value = reinterpret(UInt64, parsed::Int64)
    elseif t == "b"
        datatype = CT_BOOL
        value = v == "1" ? UInt64(1) :
                v == "0" ? UInt64(0) :
                throw(XLSXError("Unknown boolean value: $v"))

    elseif t == "str"
        datatype = CT_STRING
        value = reinterpret(UInt64, Int64(add_shared_string!(wb, v)))

    elseif t == "e"
        datatype = CT_ERROR
        value = get_error_type(v)

    elseif t == "d"
        # Check for 'T' separator (datetime) vs '-' (date) vs time-only
        T_pos = findfirst(==('T'), v)
        if T_pos !== nothing && T_pos > firstindex(v)
            # Datetime: strip trailing Z, truncate sub-milliseconds
            v2 = rstrip(v, 'Z')
            dot_pos = findfirst(==('.'), v2)
            if dot_pos !== nothing && length(v2) - dot_pos > 3
                v2 = v2[begin:dot_pos+3]
            end
            dt = Dates.DateTime(v2, Dates.dateformat"yyyy-mm-ddTHH:MM:SS.sss")
            value = reinterpret(UInt64, datetime_to_excel_value(dt, wb.date1904))
            datatype = CT_DATETIME
        elseif findfirst(==('-'), v) !== nothing
            d = Dates.Date(v, Dates.dateformat"yyyy-mm-dd")
            value = reinterpret(UInt64, date_to_excel_value(d, wb.date1904))
            datatype = CT_DATE
        else
            # Time-only HH:MM:SS[.frac][Z] — parse without allocating via split
            v2 = rstrip(v, 'Z')
            c1 = findfirst(==(':'), v2)
            c2 = findnext(==(':'), v2, c1+1)
            h  = parse(Float64, @view v2[begin:c1-1])
            m  = parse(Float64, @view v2[c1+1:c2-1])
            s  = parse(Float64, @view v2[c2+1:end])
            value = reinterpret(UInt64, (h * 3600.0 + m * 60.0 + s) / 86400.0)
            datatype = CT_TIME
        end

    else
        throw(XLSXError("Cannot parse cell value: $v"))
    end

    return datatype, value
end

# Constructor with simple formula string, for backward compatibility and tests.
function Cell(wb::Workbook, ref::CellRef, t::String, s::String, v::String, m::String, f::Bool)
    style, num_style = _parse_style(s)
    meta::UInt32     = isempty(m) ? UInt32(0) : parse(UInt32, m)
    datatype, value  = process_tv(wb, t, v, num_style)
    return Cell(ref, value, style, meta, datatype, f)
end

const EXCEL_DATE_OFFSET_1904 = 695056
const EXCEL_DATE_OFFSET_1900 = 693594
const NANOSECONDS_PER_DAY    = Int64(86_400) * Int64(1_000_000_000)

# Converts Excel number to Time.
# x must be in [0, 1), where 1 represents one full day.
# The decimal part of a floating point number represents the time fraction of a day.
function excel_value_to_time(x::Float64)::Dates.Time
    0.0 <= x < 1.0 || throw(XLSXError("A value must be in [0, 1) to be converted to time. Got $x"))
    return Dates.Time(Dates.Nanosecond(round(Int64, x * NANOSECONDS_PER_DAY)))
end

time_to_excel_value(x::Dates.Time)::Float64 = Dates.value(x) / NANOSECONDS_PER_DAY

# Converts Excel number to Date. See also XLSX.isdate1904.
function excel_value_to_date(x::Integer, is1904::Bool)::Dates.Date
    offset = is1904 ? EXCEL_DATE_OFFSET_1904 : EXCEL_DATE_OFFSET_1900
    return Dates.Date(Dates.rata2datetime(x + offset))
end

function date_to_excel_value(date::Dates.Date, is1904::Bool)::Int64
    offset = is1904 ? EXCEL_DATE_OFFSET_1904 : EXCEL_DATE_OFFSET_1900
    return Dates.datetime2rata(date) - offset
end

# Converts Excel number to DateTime.
# The integer part represents the Date, the decimal part the Time.
# See also XLSX.isdate1904.
function excel_value_to_datetime(x::Float64, is1904::Bool)::Dates.DateTime
    x >= 0 || throw(XLSXError("Cannot have a datetime value < 0. Got $x"))
    dt_part = trunc(Int64, x)
    # Round to nearest second to absorb float precision drift
    hr_ns = round(Int64, (x - dt_part) * NANOSECONDS_PER_DAY / 1_000_000_000) * 1_000_000_000
    return excel_value_to_date(dt_part, is1904) + Dates.Time(Dates.Nanosecond(hr_ns))
end

function datetime_to_excel_value(dt::Dates.DateTime, is1904::Bool)::Float64
    date_part = date_to_excel_value(Dates.Date(dt), is1904)
    time_part = Dates.value(Dates.Time(dt)) / NANOSECONDS_PER_DAY  # integer nanoseconds / const
    return date_part + time_part
end

@inline getdata(ws::Worksheet, empty::EmptyCell) = missing

"""
    getdata(ws::Worksheet, cell::Cell) :: CellValue

Returns a Julia representation of a given cell value.
The result data type is chosen based on the value of the cell as well as its style.

For example, date is stored as integers inside the spreadsheet, and the style is the
information that is taken into account to chose `Date` as the result type.

For numbers, if the style implies that the number is visualized with decimals,
the method will return a float, even if the underlying number is stored
as an integer inside the spreadsheet XML.

If `cell` has empty value or empty `String`, this function will return `missing`.
"""
function getdata(ws::Worksheet, cell::Cell)
    dt = cell.datatype
    v  = cell.value

    # Fast path for common non-date types — avoids fetching workbook date mode
    dt == CT_EMPTY  && return missing
    dt == CT_ERROR  && return missing
    dt == CT_STRING && return sst_unformatted_string(ws, reinterpret(Int64, v))
    dt == CT_BOOL   && return v != 0
    dt == CT_INT    && return reinterpret(Int64, v)
    dt == CT_FLOAT  && return reinterpret(Float64, v)

    # Date types require workbook date mode — fetch only when needed
    is1904 = isdate1904(get_workbook(ws))
    dt == CT_DATE     && return excel_value_to_date(reinterpret(Int64, v), is1904)
    dt == CT_DATETIME && return excel_value_to_datetime(reinterpret(Float64, v), is1904)
    dt == CT_TIME     && return excel_value_to_time(reinterpret(Float64, v))

    throw(XLSXError("Couldn't parse data for $cell."))
end

# Extract cells from a <row> LazyNode and push them (in place) into a Dict(column -> Cell)
# Extract cells from a <row> LazyNode and push them (in place) into a Dict(column -> Cell)
function get_rowcells!(rowcells::Dict{Int,Cell}, row::XML.LazyNode, ws::Worksheet, sst_pfx::String,
                        local_formulas::Dict{SheetCellRef,AbstractFormula}, load_formulas::Bool=true)
    sst_count = 0
    for child in XML.eachchildnode(row)
        XML.nodetype(child) == XML.Element || continue
        localname(child) == "c" || continue
        cell = Cell(child, ws, sst_pfx, local_formulas, load_formulas)
        sst_count += cell.datatype == CT_STRING ? 1 : 0
        rowcells[column_number(cell)] = cell
    end
    return nothing, sst_count
end

