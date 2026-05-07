
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
function get_error_type(v::AbstractString)::UInt64
    if v == "#NULL!"
        return UInt64(XL_NULL)
    elseif v == "#DIV/0!"
        return UInt64(XL_DIV0)
    elseif v == "#VALUE!"
        return UInt64(XL_VALUE)
    elseif v == "#REF!"
        return UInt64(XL_REF)
    elseif v == "#NAME?"
        return UInt64(XL_NAME)
    elseif v == "#NUM!"
        return UInt64(XL_NUM)
    elseif v == "#N/A"
        return UInt64(XL_NA)
    elseif v == "#SPILL!"
        return UInt64(XL_SPILL)
    else
        throw(XLSXError("Unknown error value: $v"))
    end
end

function get_error_string(e::UInt64)::String
    if e == UInt64(XL_NULL)
        return "#NULL!"
    elseif e == UInt64(XL_DIV0)
        return "#DIV/0!"
    elseif e == UInt64(XL_VALUE)
        return "#VALUE!"
    elseif e == UInt64(XL_REF)
        return "#REF!"
    elseif e == UInt64(XL_NAME)
        return "#NAME?"
    elseif e == UInt64(XL_NUM)
        return "#NUM!"
    elseif e == UInt64(XL_NA)
        return "#N/A"
    elseif e == UInt64(XL_SPILL) # Won't happen - #SPILL isn't an actual error. Returns #VALUE! instead.
        return "#SPILL!"
     else
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
#=
# Returns the enums directly rather than the strings:
# julia> XLSX.geterror(s, "A1")
# XL_NULL::CellErrorType = 0x0000000000000001
#
function geterror(c::AbstractCell)
    c = getval(c)
    isnothing(c) && return ""
    return CellErrorType(c)
end
=#

# Extracts the unformatted text from an inlineStr "is" XML element as a <si> XML string.
function _rewrite_node(node::XML.LazyNode, pfx::Union{String,Nothing})::String
    pfx = something(pfx, "")
    tag = localname(node)
    
    attrs = XML.attributes(node)
    attr_str = isnothing(attrs) || isempty(attrs) ? "" : " " * join(
        ("$(k)=\"$(v)\"" for (k, v) in attrs), " "
    )
    
    children = XML.children(node)
    
    if tag == "t"
        # Emit text inline to avoid injecting whitespace text nodes
        txt = if isempty(children)
            XML.value(node)
        else
            join((XML.is_simple(c) ? XML.simple_value(c) : something(XML.value(c), "") for c in children), "")
        end
        return "<$(pfx)$(tag)$(attr_str)>$(something(txt, ""))</$(pfx)$(tag)>"
    elseif isempty(children)
        txt = XML.value(node)
        if txt !== nothing && !isempty(txt)
            return "<$(pfx)$(tag)$(attr_str)>$(txt)</$(pfx)$(tag)>"
        else
            return "<$(pfx)$(tag)$(attr_str)/>"
        end
    else
        inner = join(_rewrite_node.(children, pfx), "\n  ")
        return "<$(pfx)$(tag)$(attr_str)>\n  $(inner)\n</$(pfx)$(tag)>"
    end
end

function _build_si_xml(si_node::XML.LazyNode, pfx::String)::String
    children = XML.children(si_node)
    inner = join(_rewrite_node.(children, pfx), "\n  ")
    prefix_part = isempty(pfx) ? "si" : "$(pfx)si"
    return "<$(prefix_part)>\n  $(inner)\n</$(prefix_part)>"
end
#=
function _build_si_xml(child::XML.LazyNode, pfx::String)::String
    inner = join(XML.write.(XML.children(child)), "\n")
    return "<$(pfx)si>\n  $inner\n</$(pfx)si>"
end
=#

# Parses a style string to (UInt32, Int) for use as style and num_style.
function _parse_style(s::String)
    isempty(s) && return UInt32(0), 0
    n = parse(Int, s)
    return UInt32(n), n
end

# Resolves unhandled_attributes to nothing if empty, for compact Formula construction.
_extra_attrs(d::Dict) = isempty(d) ? nothing : d

function Cell(c::XML.LazyNode, ws::Worksheet; mylock::Union{ReentrantLock,Nothing}=nothing)::Union{Cell,EmptyCell}
    wb = get_workbook(ws)
    sst_pfx = get_prefix("xl/SharedStrings.xml", get_xlsxfile(ws))
    if isnothing(sst_pfx) || sst_pfx == ""
        sst_pfx = ""
    else
        sst_pfx = ":"
    end

    localname(c) == "c" || throw(XLSXError("`Cell` expects a `c` (cell) XML node."))

    a = XML.attributes(c)
    chn = XML.children(c)
    ref = CellRef(a["r"])

    t     = get(a, "t", "")
    s_str = get(a, "s", "")
    m_str = get(a, "cm", "")

    # Parse style once, reuse for both UInt32 style field and Int num_style
    style, num_style = _parse_style(s_str)
    meta::UInt32     = isempty(m_str) ? UInt32(0) : parse(UInt32, m_str)

    datatype::CellValueType = CT_EMPTY
    value::UInt64           = UInt64(0)
    formula::Bool           = false

    if t == "inlineStr"
        for child in chn
            localname(child) == "is" || continue
            uft = unformatted_text(wb, child)
            if !isempty(uft)
                ft = _build_si_xml(child, sst_pfx)
                datatype = CT_STRING
                value = reinterpret(UInt64, Int64(add_formatted_string!(wb, ft; mylock)))
            end
            break
        end
    else
        for child in chn
            tag = localname(child)
            if tag == "v"
                ch = XML.children(child)
                isempty(ch) && continue
                raw = XML.value(ch[1])
                v = occursin('&', raw) ? XLSX.unescape(raw) : raw
                datatype, value = process_tv(wb, t, v, num_style; mylock)
            elseif tag == "f"
                if get_xlsxfile(wb).is_writable
                    f = parse_formula_from_element(wb,child)
                    wb.formulas[SheetCellRef(combine_sheet_ref(ws, ref))] = f
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

    # Extract formula string
    formula_string = if XML.is_simple(c_child_element)
        XLSX.unescape(XML.simple_value(c_child_element))
    else
        text_nodes = filter(x -> XML.nodetype(x) == XML.Text, XML.children(c_child_element))
        isempty(text_nodes) ? "" : XLSX.unescape(XML.value(text_nodes[1]))
    end

    a = XML.attributes(c_child_element)

    # Collect unhandled attributes
    unhandled = Dict{String,String}()
    if !isnothing(a)
        for (k, v) in a
            k ∉ ("t", "si", "ref") && push!(unhandled, k => v)
        end
    end

    is_array = false
    ref      = nothing

    if !isnothing(a) && haskey(a, "t")
        if a["t"] == "shared"
            haskey(a, "si") || throw(XLSXError("Expected shared formula to have an index. `si` attribute is missing: $c_child_element"))
            si = parse(Int, a["si"])
            extra = _extra_attrs(unhandled)
            return haskey(a, "ref") ?
                ReferencedFormula(formula_string, si, a["ref"], extra) :
                FormulaReference(si, extra)
        elseif a["t"] == "array"
            is_array = true
            ref = get(a, "ref", nothing)
        end
    end

    return Formula(
        formula_string,
        is_array ? "array" : nothing,
        ref,
        _extra_attrs(unhandled)
    )
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

function process_tv(wb::Workbook, t::String, v::String, num_style::Int; mylock::Union{ReentrantLock,Nothing}=nothing)
    datatype::CellValueType = CT_EMPTY
    value::UInt64           = UInt64(0)
    isempty(v) && return datatype, value

    if t == "b"
        datatype = CT_BOOL
        value = v == "1" ? UInt64(1) :
                v == "0" ? UInt64(0) :
                throw(XLSXError("Unknown boolean value: $v"))

    elseif t == "s"
        datatype = CT_STRING
        value = reinterpret(UInt64, parse(Int64, v))

    elseif t == "d" # ISO 8601 date/datetime/time string
        if occursin("T", v) && !startswith(v, "T")
            dt = Dates.DateTime(replace(rstrip(v, 'Z'), r"(\.\d{3})\d+$" => s"\1"), Dates.dateformat"yyyy-mm-ddTHH:MM:SS.sss")
            serial = datetime_to_excel_value(dt, wb.date1904)
            datatype = CT_DATETIME
        elseif occursin("-", v)
            d = Dates.Date(v, Dates.dateformat"yyyy-mm-dd")
            serial = date_to_excel_value(d, wb.date1904)
            datatype = CT_DATE
        else
            # Time-only: parse HH:MM:SS.fractional directly to fractional day
            parts = split(rstrip(v, 'Z'), ":")
            seconds = parse(Float64, parts[1]) * 3600.0 +
                      parse(Float64, parts[2]) * 60.0 +
                      parse(Float64, parts[3])
            serial = seconds / 86400.0
            datatype = CT_TIME
        end
        value = reinterpret(UInt64, serial)

    elseif t == "str"
        datatype = CT_STRING
        value = reinterpret(UInt64, Int64(add_shared_string!(wb, v; mylock)))

    elseif t == "e"
        datatype = CT_ERROR
        value = get_error_type(v)

    elseif t == "n" || t == ""
        if styles_is_datetime(wb, num_style)
            value, datatype = _parse_excel_datetime_raw(v)
        elseif styles_is_float(wb, num_style)
            datatype = CT_FLOAT
            value = reinterpret(UInt64, parse(Float64, v))
        else
            # Use tryparse to distinguish integers from floats, avoiding manual byte scanning
            parsed_int = tryparse(Int64, v)
            if !isnothing(parsed_int)
                datatype = CT_INT
                value = reinterpret(UInt64, parsed_int)
            else
                datatype = CT_FLOAT
                value = reinterpret(UInt64, parse(Float64, v))
            end
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

#=
# Shared helper for parsing a raw Excel datetime string into a value and CellValueType.
function _parse_excel_datetime(v::AbstractString, is1904::Bool)
    isempty(v) && throw(XLSXError("Cannot convert an empty string into a datetime value."))
    if occursin('.', v) || v == "0"
        time_value = parse(Float64, v)
        time_value >= 0 || throw(XLSXError("Cannot have a datetime value < 0. Got $time_value"))
        return time_value < 1.0 ?
            (excel_value_to_time(time_value), CT_TIME) :
            (excel_value_to_datetime(time_value, is1904), CT_DATETIME)
    else
        return excel_value_to_date(parse(Int64, v), is1904), CT_DATE
    end
end
=#

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
function get_rowcells!(rowcells::Dict{Int,Cell}, row::XML.LazyNode, ws::Worksheet; mylock::Union{ReentrantLock,Nothing}=nothing)

    #=
        # threaded cell extraction causes hugely more lock conflicts for low cellchunk size.
        # may be worthwhile if many columns (hundreds+), with a cellchunk size > ~10 or ~20, but this is unverified.

        # debug
        # @assert row.tag == "row" "Not a row node"
        cellchunk=8 # bigger chunks, fewer lock conflicts but columns are generally relatively few.
        sst_count=0
        d=row.depth

        row_cellnodes = Channel{Vector{XML.LazyNode}}(1 << 8)
        row_cells = Channel{Vector{XLSX.Cell}}(1 << 8)

        # consumer task
        consumer = @async begin
            for cells in row_cells  
                for cell in cells      
                    sst_count += cell.datatype == "s" ? 1 : 0
                    rowcells[column_number(cell)] = cell
                end
            end
        end

        # Feed row_cellnodes
        cellnodes = Vector{XML.LazyNode}(undef, cellchunk)
        pos=0
        cellnode=XML.next(row)
        while !isnothing(cellnode) && cellnode.depth > d
            if cellnode.tag == "c" # This is a cell
                pos += 1
                cellnodes[pos] = cellnode
            end
            if pos >= cellchunk
                put!(row_cellnodes, copy(cellnodes))
                pos=0
            end
            cellnode = XML.next(cellnode)
        end
        if pos>0 # handle last incomplete chunk
            put!(row_cellnodes, cellnodes[1:pos])
        end
        close(row_cellnodes)

        # Producer tasks
        mylock = ReentrantLock() # lock for thread-safe access to shared string table in case of inlineStrings
        @sync for _ in 1:Threads.nthreads()
            Threads.@spawn begin
                chunk = Vector{XLSX.Cell}(undef, cellchunk)
                for cns in row_cellnodes
                    cell_count=0
                    for cn in cns
                        cell_count += 1
                        chunk[cell_count] = Cell(cn, ws; mylock)
                        if cell_count >= cellchunk
                            put!(row_cells, copy(chunk))
                            cell_count=0
                        end
                    end
                    if cell_count > 0 # handle last incomplete chunk
                        put!(row_cells, chunk[1:cell_count])
                    end
                end
            end
        end
        close(row_cells)

        wait(consumer)  # ensure consumer is done

        if !isnothing(cellnode) && cellnode.tag == "row" # have reached the end of last row, beginning of next
            return cellnode, sst_count
        else                                             # no more rows
            return nothing, sst_count
        end
    =#
    # unthreaded cell extraction is (exceedingly marginally) slower but no lock conflicts introduced.

    # debug
    # @assert row.tag == "row" "Not a row node"

    sst_count = 0

    d = row.depth

    cellnode = XML.next(row)

    while !isnothing(cellnode) && cellnode.depth > d
        if localname(cellnode) == "c" # This is a cell
            cell = Cell(cellnode, ws; mylock) # construct an XLSX.Cell from an XML.LazyNode
            sst_count += cell.datatype == CT_STRING ? 1 : 0
            rowcells[column_number(cell)] = cell
        end
        cellnode = XML.next(cellnode)
    end
    if !isnothing(cellnode) && localname(cellnode) == "row" # have reached the beginning of next row
        return cellnode, sst_count
    else                                             # no more rows
        return nothing, sst_count
    end

end
