
function CellRef(n::AbstractString)
    i = 1
    len = ncodeunits(n)
    len == 0 && throw(XLSXError("$n is not a valid CellRef."))
    
    # scan letters for column
    while i <= len && isascii(n[i]) && isuppercase(n[i])
        i += 1
    end
    col_end = i - 1
    col_end == 0 && throw(XLSXError("$n is not a valid CellRef."))
    
    # decode column number inline (no SubString, no decode_column_number call)
    col = 0
    for j in 1:col_end
        col = col * 26 + (codeunit(n, j) - UInt8('A') + 1)
    end
    
    # parse row number inline (no SubString, no parse() call)
    row = 0
    while i <= len
        c = codeunit(n, i)
        (c < UInt8('0') || c > UInt8('9')) && throw(XLSXError("$n is not a valid CellRef."))
        row = row * 10 + (c - UInt8('0'))
        i += 1
    end
    row == 0 && throw(XLSXError("$n is not a valid CellRef."))
    
    return CellRef(Int32(row), Int32(col))
end
#=
function CellRef(n::AbstractString)
    !is_valid_cellname(n) && throw(XLSXError("$n is not a valid CellRef."))
    column_name, row_number = split_cellname(n)
    return CellRef(Int32(row_number), Int32(decode_column_number(column_name)))
end
=#
@inline CellRef(row::Integer, col::Integer) = CellRef(Int32(row), Int32(col))
@inline CellPosition(ref::CellRef) = CellPosition(row_number(ref), column_number(ref))
@inline row_number(p::CellPosition) = p.row
@inline column_number(p::CellPosition) = p.column
@inline CellRef(p::CellPosition) = CellRef(row_number(p), column_number(p))

# Converts column name to a column number. See also XLSX.encode_column_number.
function decode_column_number(column_name::AbstractString) :: Int
    local result::Int = 0

    !isascii(column_name) && throw(XLSXError("$column_name is not a valid column name."))
    num_characters = length(column_name) # this is safe, since `column_name` is encoded as ASCII

    iteration = 1
    for i in num_characters:-1:1
        column_char_as_int = Int(column_name[i])
        result += (26^(iteration-1)) * (column_char_as_int - 64) # From A to Z we have 26 values. 'A' Char is ASCII 65.
        iteration += 1
    end

    return result
end

# Converts column number to a column name. See also XLSX.decode_column_number.
function encode_column_number(column_number::Integer) :: String
    if column_number <= 0 && column_number > EXCEL_MAX_COLS
        throw(XLSXError("Column number should be in the range from 1 to $EXCEL_MAX_COLS."))
    end

    third_letter_sequence = div(column_number - 26 - 1, 26*26)
    column_number = column_number - third_letter_sequence*(26*26)

    second_letter_sequence = div(column_number - 1, 26) # 26^1
    column_number = column_number - second_letter_sequence*(26)

    first_letter_sequence = column_number # 26^0

    if third_letter_sequence > 0
        # result will have 3 letters
        return String([ Char(third_letter_sequence+64), Char(second_letter_sequence+64), Char(first_letter_sequence+64) ])

    elseif second_letter_sequence > 0
        # result will have 2 letters
        return String([ Char(second_letter_sequence+64), Char(first_letter_sequence+64) ])

    else
        # result will have 1 letter
        return String([ Char(first_letter_sequence+64) ])
    end
end

Base.string(c::CellRef) = string(encode_column_number(column_number(c))) * string(row_number(c))
Base.show(io::IO, c::CellRef) = print(io, string(c))
Base.:(==)(c1::CellRef, c2::CellRef) = c1.row_number == c2.row_number && c1.column_number == c2.column_number
Base.hash(c::CellRef, h::UInt) = hash(c.column_number, hash(c.row_number, h))
Base.isless(c1::CellRef, c2::CellRef) = Base.isless(string(c1), string(c2))
cellname(c::CellRef) :: String = string(encode_column_number(column_number(c))) * string(row_number(c))

Base.:(==)(rf1::ReferencedFormula, rf2::ReferencedFormula) = (
    rf1.formula == rf2.formula &&
    rf1.id == rf2.id &&
    rf1.ref == rf2.ref &&
    rf1.unhandled == rf2.unhandled
)

Base.:(==)(rf1::FormulaReference, rf2::FormulaReference) = (
    rf1.id == rf2.id &&
    rf1.unhandled == rf2.unhandled
)

Base.:(==)(rf1::Formula, rf2::Formula) = (
    rf1.formula == rf2.formula &&
    rf1.type == rf2.type &&
    rf1.ref == rf2.ref &&
    rf1.unhandled == rf2.unhandled
)

const RGX_COLUMN_NAME = r"^[A-Z]?[A-Z]?[A-Z]$"
const RGX_ROW_NAME = r"^[1-9][0-9]*$"
const RGX_CELLNAME = r"^[A-Z]+[0-9]+$"
const RGX_CELLRANGE = r"^[A-Z]+[0-9]+:[A-Z]+[0-9]+$"

function is_valid_column_name(n::AbstractString) :: Bool
    if !occursin(RGX_COLUMN_NAME, n)
        return false
    end

    column_number = decode_column_number(n)
    if column_number < 1 || column_number > EXCEL_MAX_COLS
        return false
    end

    return true
end
function is_valid_row_name(n::AbstractString) :: Bool
    if !occursin(RGX_ROW_NAME, n)
        return false
    end

    row_number = parse(Int, n)
    if row_number < 1 || row_number > EXCEL_MAX_ROWS
        return false
    end

    return true
end

const RGX_CELLNAME_LEFT = r"^[A-Z]+"
const RGX_CELLNAME_RIGHT = r"[0-9]+$"

# Splits a string representing a cell name to its column name and row number.
@inline function split_cellname(n::AbstractString)
    !isascii(n) && throw(XLSXError("$n is not a valid cell name."))
    for (i, c) in enumerate(n)
        if isdigit(c) # this block is safe since n is encoded as ASCII
            column_name = SubString(n, 1, i-1)
            row = parse(Int, SubString(n, i, length(n)))
            return column_name, row
        end
    end

    throw(XLSXError("Couldn't split (column_name, row) for cellname $n."))
end

# Checks whether `n` is a valid name for a cell.
function is_valid_cellname(n::AbstractString) :: Bool

    if is_valid_non_contiguous_range(n) # Non-contiguous ranges are comma separated `CellRef-like` or `CellRange-like` strings
        return false
    end

    if !occursin(RGX_CELLNAME, n)
        return false
    end

    column_name, row = split_cellname(n)

    if row < 1 || row > EXCEL_MAX_ROWS
        return false
    end

    if !is_valid_column_name(column_name)
        return false
    end

    return true
end

const RGX_CELLRANGE_START = r"^[A-Z]+[0-9]+"
const RGX_CELLRANGE_STOP = r"[A-Z]+[0-9]+$"

#=
    split_cellrange(n::AbstractString) -> start_name, stop_name

Splits a string representing a cell range into its cell names.

# Example

```julia
julia> XLSX.split_cellrange("AB12:CD24")
("AB12", "CD24")
```
=#
@inline function split_cellrange(n::AbstractString)
    s = split(n, ":")
    length(s) != 2 && throw(XLSXError("$n is not a valid cell range."))
    return s[1], s[2]
end

function is_valid_cellrange(n::AbstractString) :: Bool
    if is_valid_non_contiguous_range(n) # Non-contiguous ranges are comma separated `CellRef-like` or `CellRange-like` strings
        return false
    end

    if !occursin(RGX_CELLRANGE, n)
        return false
    end

    start_name, stop_name = split_cellrange(n)

    if !is_valid_cellname(start_name)
        return false
    end

    if !is_valid_cellname(stop_name)
        return false
    end

    return true
end

macro ref_str(ref)
    CellRef(ref)
end

function CellRange(r::AbstractString)
    !occursin(RGX_CELLRANGE, r) && throw(XLSXError("Invalid cell range: $r."))
    start_name, stop_name = split_cellrange(r)
    return CellRange(CellRef(start_name), CellRef(stop_name))
end

CellRange(start_row::Integer, start_column::Integer, stop_row::Integer, stop_column::Integer) = CellRange(CellRef(start_row, start_column), CellRef(stop_row, stop_column))

Base.string(cr::CellRange) = string(cr.start)*":"*string(cr.stop)
Base.show(io::IO, cr::CellRange) = print(io, string(cr))
Base.:(==)(cr1::CellRange, cr2::CellRange) = cr1.start == cr2.start && cr1.stop == cr2.stop
Base.hash(cr::CellRange, h::UInt) = hash(cr.stop, hash(cr.start, h))
Base.isless(cr1::CellRange, cr2::CellRange) = Base.isless(string(cr1), string(cr2)) # needed for tests

macro range_str(cellrange)
    CellRange(cellrange)
end

# Checks whether `ref` is a cell reference inside a range given by `rng`.
function Base.in(ref::CellRef, rng::CellRange) :: Bool
    top = row_number(rng.start)
    bottom = row_number(rng.stop)
    r = row_number(ref)

    if top <= r && r <= bottom
        left = column_number(rng.start)
        right = column_number(rng.stop)
        c = column_number(ref)

        if left <= c && c <= right
            return true
        end
    end

    return false
end

# Checks whether `subrng` is a cell range contained in `rng`.
Base.issubset(subrng::CellRange, rng::CellRange) :: Bool = in(subrng.start, rng) && in(subrng.stop, rng)
 
function intersects(rng1::CellRange, rng2::CellRange) :: Bool
    if row_number(rng1.start) <= row_number(rng2.stop) && row_number(rng1.start) >= row_number(rng2.start) && 
        column_number(rng1.stop) <= column_number(rng2.stop) && column_number(rng1.stop) >= column_number(rng2.start)
#        println("offset 1")
        return true
    end
    if row_number(rng2.start) <= row_number(rng1.stop) && row_number(rng2.start) >= row_number(rng1.start) && 
        column_number(rng2.stop) <= column_number(rng1.stop) && column_number(rng2.stop) >= column_number(rng1.start)
#        println("offset 2")
        return true
    end
    if row_number(rng1.start)>=row_number(rng2.start) && column_number(rng1.start)<=column_number(rng2.start) && 
        row_number(rng1.stop)<=row_number(rng2.stop) && column_number(rng1.stop)>=column_number(rng2.stop)
#        println("cruciform 1")
        return true
    end
    if row_number(rng2.start)>=row_number(rng1.start) && column_number(rng2.start)<=column_number(rng1.start) && 
        row_number(rng2.stop)<=row_number(rng1.stop) && column_number(rng2.stop)>=column_number(rng1.stop)
#        println("cruciform 2")
        return true
    end
    if in(rng1.start, rng2) || in(rng1.stop, rng2)
#        println("inside corner 1")
        return true
    end
    if in(rng2.start, rng1) || in(rng2.stop, rng1)
#        println("inside corner 2")
        return true
    end
    if issubset(rng1, rng2) || issubset(rng2, rng1)
#        println("subset")
        return true
    end
    return false
end


function Base.size(rng::CellRange)
    top = row_number(rng.start)
    bottom = row_number(rng.stop)
    left = column_number(rng.start)
    right = column_number(rng.stop)

    return ( bottom - top + 1, right - left + 1 )
end

"""
    row_number(c::CellRef) :: Int

Returns the row number of a given cell reference.
"""
row_number(c::CellRef) :: Int = c.row_number

"""
    column_number(c::CellRef) :: Int

Returns the column number of a given cell reference.
"""
column_number(c::CellRef) :: Int = c.column_number

column_name(c::CellRef) :: String = encode_column_number(column_number(c))

#=
Returns (row, column) representing a `ref` position relative to `rng`.

For example, for a range "B2:D4", we have:

* "C3" relative position is (2, 2)

* "B2" relative position is (1, 1)

* "C4" relative position is (3, 2)

* "D4" relative position is (3, 3)
=#

function relative_cell_position(ref::CellRef, rng::CellRange)
    ref ∉ rng && throw(XLSXError("$ref is outside range $rng."))

    top = row_number(rng.start)
    left = column_number(rng.start)

    r, c = row_number(ref), column_number(ref)

    return ( r - top + 1 , c - left + 1 )
end

#
# ColumnRange and RowRange
#

Base.string(cr::ColumnRange) = encode_column_number(cr.start)*":"*encode_column_number(cr.stop)
Base.show(io::IO, cr::ColumnRange) = print(io, string(cr))
Base.:(==)(cr1::ColumnRange, cr2::ColumnRange) = cr1.start == cr2.start && cr1.stop == cr2.stop
Base.hash(cr::ColumnRange, h::UInt) = hash(cr.stop, hash(cr.start, h))
Base.in(column_number::Integer, rng::ColumnRange) = rng.start <= column_number && column_number <= rng.stop

Base.string(cr::RowRange) = string(cr.start)*":"*string(cr.stop)
Base.show(io::IO, cr::RowRange) = print(io, string(cr))
Base.:(==)(cr1::RowRange, cr2::RowRange) = cr1.start == cr2.start && cr1.stop == cr2.stop
Base.hash(cr::RowRange, h::UInt) = hash(cr.stop, hash(cr.start, h))
Base.in(row_number::Integer, rng::RowRange) = rng.start <= row_number && row_number <= rng.stop

function cell_offset(from::CellRef, to::CellRef) # return tuple (row_offset, column_offset) between `from` and `to` cells
    offset = (to.row_number - from.row_number, to.column_number - from.column_number)
    (first(offset)>=0 && last(offset) >=0) || throw(XLSXError("The 'to' cell cannot be above or to the left of the 'from' cell."))
    return offset
end

function relative_column_position(column_number::Integer, rng::ColumnRange)
    column_number ∉ rng && throw(XLSXError("Column $column_number is outside range $rng."))
    return column_number - rng.start + 1
end

@inline relative_column_position(ref::CellRef, rng::ColumnRange) = relative_column_position(column_number(ref), rng)

const RGX_COLUMN_RANGE = r"^[A-Z]?[A-Z]?[A-Z]:[A-Z]?[A-Z]?[A-Z]$"
const RGX_COLUMN_RANGE_START = r"^[A-Z]+"
const RGX_COLUMN_RANGE_STOP = r"[A-Z]+$"
const RGX_SINGLE_COLUMN = r"^[A-Z]+$"
const RGX_ROW_RANGE = r"^[1-9][0-9]*:[1-9][0-9]*$"
const RGX_ROW_RANGE_START = r"^[1-9][0-9]*"
const RGX_ROW_RANGE_STOP = r"[1-9][0-9]*$"
const RGX_SINGLE_ROW = r"^[1-9][0-9]*$"

# Returns tuple (column_name_start, column_name_stop).
# Also works for row ranges (row_name_start, row_name_stop)!
@inline function split_sheet_range(n::AbstractString)
    if !occursin(":", n)
        return n, n
    else
        s = split(n, ":")
        return s[1], s[2]
    end
end

@inline function combine_sheet_range(s::Worksheet, rng::CellRange)
    return string(s.name, "!", encode_column_number(rng.start.column_number), rng.start.row_number, ":", encode_column_number(rng.stop.column_number), rng.stop.row_number)
end
@inline function combine_sheet_ref(s::Worksheet, ref::CellRef)
    return string(s.name, "!", encode_column_number(ref.column_number), ref.row_number)
end

function is_valid_column_range(r::AbstractString) :: Bool
    if occursin(RGX_SINGLE_COLUMN, r)
        return true
    end
    if !occursin(RGX_COLUMN_RANGE, r)
        return false
    end
    start_name, stop_name = split_sheet_range(r)
    if !is_valid_column_name(start_name) || !is_valid_column_name(stop_name)
        return false
    end
    return true
end
function is_valid_row_range(r::AbstractString) :: Bool
    if occursin(RGX_SINGLE_ROW, r)
        row_number = parse(Int, r)
        if row_number <= 0 && row_number > EXCEL_MAX_ROWS
            throw(XLSXError("Row number should be in the range from 1 to $EXCEL_MAX_ROWS."))
        end
        return true
    end
    if !occursin(RGX_ROW_RANGE, r)
        return false
    end
    start_name, stop_name = split_sheet_range(r) # Function works for row ranges too.
    if !is_valid_row_name(start_name) || !is_valid_row_name(stop_name)
        return false
    end
    return true
end

function RowRange(r::AbstractString)
    if is_valid_cellrange(r)
        cr=CellRange(r)
        return RowRange(row_number(cr.start), row_number(cr.stop))
    end
    !is_valid_row_range(r) && throw(XLSXError("Invalid row range: $r."))
    start_name, stop_name = split_sheet_range(r) 
    return RowRange(parse(Int, start_name), parse(Int, stop_name))
end
function ColumnRange(r::AbstractString)
    if is_valid_cellrange(r)
        cr=CellRange(r)
        return ColumnRange(column_number(cr.start), column_number(cr.stop))
    end
        !is_valid_column_range(r) && throw(XLSXError("Invalid column range: $r."))
    start_name, stop_name = split_sheet_range(r)
    return ColumnRange(decode_column_number(start_name), decode_column_number(stop_name))
end

convert(::Type{ColumnRange}, str::AbstractString) = ColumnRange(str)
convert(::Type{ColumnRange}, column_range::ColumnRange) = column_range
convert(::Type{RowRange}, str::AbstractString) = RowRange(str)
convert(::Type{RowRange}, row_range::RowRange) = row_range

column_bounds(r::ColumnRange) = (r.start, r.stop)
Base.length(r::ColumnRange) = r.stop - r.start + 1
row_bounds(r::RowRange) = (r.start, r.stop)
Base.length(r::RowRange) = r.stop - r.start + 1

# ColumnRange iterator: element is a String with the column name, the state is the column number.
function Base.iterate(itr::ColumnRange, state::Int=itr.start)
    if state > itr.stop
        return nothing
    end

    return encode_column_number(state), state + 1
end

# RowRange iterator: element is a String with the row name (e.g. "1"). The state is the row number.
function Base.iterate(itr::RowRange, state::Int=itr.start)
    if state > itr.stop
        return nothing
    end

    return string(state), state + 1
end

# CellRange iterator: element is a CellRef, the state is a CellPosition.
function Base.iterate(rng::CellRange, state::CellPosition=CellPosition(rng.start))

    if row_number(state) > row_number(rng.stop)
        return nothing
    elseif column_number(state) == column_number(rng.stop)
        # reached last column. Go to the next row.
        next_state = CellPosition(row_number(state) + 1, column_number(rng.start))
    else
        # go to the next column
        next_state = CellPosition(row_number(state), column_number(state) + 1)
    end

    return CellRef(state), next_state
end

function Base.length(rng::CellRange)
    (r, c) = size(rng)
    return r * c
end

#
# SheetCellRef, SheetCellRange, SheetColumnRange, SheetRowRange, NonContiguousRange
#

Base.string(cr::SheetCellRef) = string(quoteit(cr.sheet), "!", cr.cellref)
Base.show(io::IO, cr::SheetCellRef) = print(io, string(cr))
Base.:(==)(cr1::SheetCellRef, cr2::SheetCellRef) = cr1.sheet == cr2.sheet && cr1.cellref == cr2.cellref
Base.hash(cr::SheetCellRef, h::UInt) = hash(cr.cellref, hash(cr.sheet, h))

Base.string(cr::SheetCellRange) = string(quoteit(cr.sheet), "!", cr.rng)
Base.show(io::IO, cr::SheetCellRange) = print(io, string(cr))
Base.:(==)(cr1::SheetCellRange, cr2::SheetCellRange) = cr1.sheet == cr2.sheet && cr1.rng == cr2.rng
Base.hash(cr::SheetCellRange, h::UInt) = hash(cr.rng, hash(cr.sheet, h))

Base.string(cr::SheetColumnRange) = string(quoteit(cr.sheet), "!", cr.colrng)
Base.show(io::IO, cr::SheetColumnRange) = print(io, string(cr))
Base.:(==)(cr1::SheetColumnRange, cr2::SheetColumnRange) = cr1.sheet == cr2.sheet && cr1.colrng == cr2.colrng
Base.hash(cr::SheetColumnRange, h::UInt) = hash(cr.colrng, hash(cr.sheet, h))

Base.string(cr::SheetRowRange) = string(quoteit(cr.sheet), "!", cr.rowrng)
Base.show(io::IO, cr::SheetRowRange) = print(io, string(cr))
Base.:(==)(cr1::SheetRowRange, cr2::SheetRowRange) = cr1.sheet == cr2.sheet && cr1.rowrng == cr2.rowrng
Base.hash(cr::SheetRowRange, h::UInt) = hash(cr.rowrng, hash(cr.sheet, h))

Base.string(cr::NonContiguousRange) = join([string(quoteit(cr.sheet), "!", x) for x in cr.rng],",")
Base.show(io::IO, cr::NonContiguousRange) = print(io, string(cr))
Base.:(==)(cr1::NonContiguousRange, cr2::NonContiguousRange) = cr1.sheet == cr2.sheet && cr1.rng == cr2.rng
Base.hash(cr::NonContiguousRange, h::UInt) = hash(cr.rng, hash(cr.sheet, h))

function Base.in(ref::SheetCellRef, ncrng::NonContiguousRange) :: Bool
    if ref.sheet != ncrng.sheet
        return false
    end
    for r in ncrng.rng
        if r isa CellRef
            if ref.cellref == r
                return true
            end
        else
            if ref.cellref in r
                return true
            end
        end
    end
    return false
end

function nc_bounds(r::NonContiguousRange)::CellRange # Smallest rectangualar `CellRange` that contains all the elements in `r`.
    top = EXCEL_MAX_ROWS
    bottom = 0
    left = EXCEL_MAX_COLS
    right = 0
    for rng in r.rng
        if isa(rng, CellRef)
            top = min(top, row_number(rng))
            bottom = max(bottom, row_number(rng))
            left = min(left, column_number(rng))
            right = max(right, column_number(rng))
        else
            top = min(top, row_number(rng.start))
            bottom = max(bottom, row_number(rng.stop))
            left = min(left, column_number(rng.start))
            right = max(right, column_number(rng.stop))
        end
    end
    return CellRange(CellRef(top, left), CellRef(bottom, right))
end
function Base.length(r::NonContiguousRange)::Int # Number of cells in `r`, eliminating duplicates.
#    s = 0
    allcells= Vector{String}()
    for rng in r.rng
        if rng isa CellRef
            push!(allcells, cellname(rng))
        else
            for cell in rng
                push!(allcells, cellname(cell))
            end
        end
    end
#        if rng isa CellRef
#            s += 1
#        else
#            s += length(rng)
#        end
#    end
    return length(unique(allcells))
end

const RGX_SHEET_CELLNAME = r"^.+![A-Z]+[0-9]+$"
const RGX_SHEET_CELLRANGE = r"^.+![A-Z]+[0-9]+:[A-Z]+[0-9]+$"
const RGX_SHEET_COLUMN_RANGE = r"^.+![A-Z]?[A-Z]?[A-Z]:[A-Z]?[A-Z]?[A-Z]$"
const RGX_SHEET_ROW_RANGE = r"^.+![1-9][0-9]*:[1-9][0-9]*$"

const RGX_SHEET_CELLNAME_RIGHT = r"[A-Z]+[0-9]+$"
const RGX_SHEET_CELLRANGE_RIGHT = r"[A-Z]+[0-9]+:[A-Z]+[0-9]+$"
const RGX_SHEET_COLUMN_RANGE_RIGHT = r"[A-Z]?[A-Z]?[A-Z]:[A-Z]?[A-Z]?[A-Z]$"
const RGX_SHEET_ROW_RANGE_RIGHT = r"[1-9][0-9]*:[1-9][0-9]*$"

function is_valid_sheet_cellname(n::AbstractString) :: Bool

    if is_valid_non_contiguous_range(n) # Non-contiguous ranges are comma separated `CellRef-like` or `CellRange-like` strings
        return false
    end

    if !occursin(RGX_SHEET_CELLNAME, n)
        return false
    end

    cellname = match(RGX_SHEET_CELLNAME_RIGHT, n).match
    if !is_valid_cellname(cellname)
        return false
    end

    return true
end

function is_valid_sheet_cellrange(n::AbstractString) :: Bool

    if is_valid_non_contiguous_range(n) # Non-contiguous ranges are comma separated `CellRef-like` or `CellRange-like` strings
        return false
    end

    if !occursin(RGX_SHEET_CELLRANGE, n)
        return false
    end

    cellrange = match(RGX_SHEET_CELLRANGE_RIGHT, n).match
    if !is_valid_cellrange(cellrange)
        return false
    end

    return true
end

function is_valid_sheet_column_range(n::AbstractString) :: Bool
    if !occursin(RGX_SHEET_COLUMN_RANGE, n)
        return false
    end

    column_range = match(RGX_SHEET_COLUMN_RANGE_RIGHT, n).match
    if !is_valid_column_range(column_range)
        return false
    end

    return true
end
function is_valid_sheet_row_range(n::AbstractString) :: Bool
    if !occursin(RGX_SHEET_ROW_RANGE, n)
        return false
    end

    row_range = match(RGX_SHEET_ROW_RANGE_RIGHT, n).match
    if !is_valid_row_range(row_range)
        return false
    end

    return true
end

const RGX_SHEET_PREFIX = r"^.+!"
const RGX_CELLNAME_RIGHT_FIXED = r"\$[A-Z]+\$[0-9]+$"
const RGX_SHEET_CELLNAME_RIGHT_FIXED = r"\$[A-Z]+\$[0-9]+:\$[A-Z]+\$[0-9]+$"

function parse_sheetname_from_sheetcell_name(n::AbstractString) :: SubString
    !occursin(RGX_SHEET_PREFIX, n) && throw(XLSXError("$n is not a SheetCell reference."))
    sheetname = match(RGX_SHEET_PREFIX, n).match
    sheetname = SubString(sheetname, firstindex(sheetname), prevind(sheetname, lastindex(sheetname)))
    return sheetname
end

function SheetCellRef(n::AbstractString)
    local cellref::CellRef

    if is_valid_fixed_sheet_cellname(n)
        fixed_cellname = match(RGX_CELLNAME_RIGHT_FIXED, n).match
        cellref = CellRef(replace(fixed_cellname, "\$" => ""))
    else
        !is_valid_sheet_cellname(n) && throw(XLSXError("$n is not a valid SheetCellRef."))
        cellref = CellRef(match(RGX_SHEET_CELLNAME_RIGHT, n).match)
    end
    sheetname = parse_sheetname_from_sheetcell_name(n)
    return SheetCellRef(sheetname, cellref)
end

function SheetCellRange(n::AbstractString)
    local cellrange::CellRange

    if is_valid_fixed_sheet_cellrange(n)
        fixed_cellrange = match(RGX_SHEET_CELLNAME_RIGHT_FIXED, n).match
        cellrange = CellRange(replace(fixed_cellrange, "\$" => ""))
    else
        !is_valid_sheet_cellrange(n) && throw(XLSXError("$n is not a valid SheetCellRange."))
        cellrange = CellRange(match(RGX_SHEET_CELLRANGE_RIGHT, n).match)
    end

    sheetname = parse_sheetname_from_sheetcell_name(n)
    return SheetCellRange(sheetname, cellrange)
end

function SheetColumnRange(n::AbstractString)
    !is_valid_sheet_column_range(n) && throw(XLSXError("$n is not a valid SheetColumnRange."))
    column_range = match(RGX_SHEET_COLUMN_RANGE_RIGHT, n).match
    sheetname = parse_sheetname_from_sheetcell_name(n)
    return SheetColumnRange(sheetname, ColumnRange(column_range))
end
function SheetRowRange(n::AbstractString)
    !is_valid_sheet_row_range(n) && throw(XLSXError("$n is not a valid SheetRowRange."))
    row_range = match(RGX_SHEET_ROW_RANGE_RIGHT, n).match
    sheetname = parse_sheetname_from_sheetcell_name(n)
    return SheetRowRange(sheetname, RowRange(row_range))
end

# Named ranges
const RGX_FIXED_CELLNAME = r"^\$[A-Z]+\$[0-9]+$"
const RGX_FIXED_CELLRANGE = r"^\$[A-Z]+\$[0-9]+:\$[A-Z]+\$[0-9]+$"
const RGX_FIXED_SHEET_CELLNAME = r"^.+!\$[A-Z]+\$[0-9]+$"
const RGX_FIXED_SHEET_CELLRANGE = r"^.+!\$[A-Z]+\$[0-9]+:\$[A-Z]+\$[0-9]+$"

is_valid_fixed_cellname(s::AbstractString) = occursin(RGX_FIXED_CELLNAME, s)
is_valid_fixed_cellrange(s::AbstractString) = occursin(RGX_FIXED_CELLRANGE, s)
is_valid_fixed_sheet_cellname(s::AbstractString) = occursin(RGX_FIXED_SHEET_CELLNAME, s)
is_valid_fixed_sheet_cellrange(s::AbstractString) = occursin(RGX_FIXED_SHEET_CELLRANGE, s)

is_valid_non_contiguous_range(v::AbstractString) :: Bool = is_valid_non_contiguous_cellrange(v) || is_valid_non_contiguous_sheetcellrange(v)

function is_valid_non_contiguous_sheetcellrange(v::AbstractString) :: Bool

    if !occursin(",", string(v)) # Non-contiguous ranges are comma separated `SheetCellRef-like` or `SheetCellRange-like` strings
        return false
    end

    ranges = split(v, ",")
    for r in ranges
        if !is_valid_sheet_cellname(r) && !is_valid_sheet_cellrange(r) && !is_valid_fixed_sheet_cellname(r) && !is_valid_fixed_sheet_cellrange(r)
            return false
        end
    end

    firstsheet = parse_sheetname_from_sheetcell_name(ranges[begin])

    if any(parse_sheetname_from_sheetcell_name(r) != firstsheet for r in ranges) # All `SheetCellRef`s and `SheetCellRange`s should have the same sheet name
        return false
    end

    return true
end

function is_valid_non_contiguous_cellrange(v::AbstractString) :: Bool

    if !occursin(",", string(v)) # Non-contiguous ranges are comma separated `SheetCellRef-like` or `SheetCellRange-like` strings

        return false
    end

    ranges = split(v, ",")

    for r in ranges
        if !is_valid_cellname(r) && !is_valid_cellrange(r) &&!is_valid_fixed_cellname(r) && !is_valid_fixed_cellrange(r)
            return false
        end
    end

    return true
end

NonContiguousRange(s::Worksheet, v::AbstractString)::NonContiguousRange = nCR(s.name, string.(split(v, ",")))
function NonContiguousRange(v::AbstractString)::NonContiguousRange

    !is_valid_non_contiguous_sheetcellrange(v) && throw(XLSXError("$v is not a valid non-contiguous sheetcell range."))
    
    ranges = string.(split(v, ","))
    firstsheet = parse_sheetname_from_sheetcell_name(ranges[1])
    if !all(parse_sheetname_from_sheetcell_name(r) == firstsheet for r in ranges)
        throw(XLSXError("All `CellRef`s and `CellRange`s should have the same sheet name."))
    end

    return nCR(unquoteit(firstsheet), ranges)
end

function nCR(s::AbstractString, ranges::Vector{String}) :: NonContiguousRange
    noncontig = Vector{Union{CellRef, CellRange}}()
    
    for n in ranges
        if is_valid_fixed_sheet_cellname(n)
            fixed_cellname = match(RGX_CELLNAME_RIGHT_FIXED, n).match
            push!(noncontig, CellRef(replace(fixed_cellname, "\$" => "")))
        elseif is_valid_sheet_cellname(n)
            push!(noncontig, CellRef(match(RGX_SHEET_CELLNAME_RIGHT, n).match))
        elseif is_valid_fixed_sheet_cellrange(n)
            fixed_cellrange = match(RGX_SHEET_CELLNAME_RIGHT_FIXED, n).match
            push!(noncontig, CellRange(replace(fixed_cellrange, "\$" => "")))
        elseif is_valid_sheet_cellrange(n)
            push!(noncontig, CellRange(match(RGX_SHEET_CELLRANGE_RIGHT, n).match))
        elseif is_valid_fixed_cellname(n)
            push!(noncontig, CellRef(replace(n, "\$" => "")))
        elseif is_valid_fixed_cellrange(n)
            push!(noncontig, CellRange(replace(n, "\$" => "")))
        elseif is_valid_cellname(n)
            push!(noncontig, CellRef(n))
        elseif is_valid_cellrange(n)
            push!(noncontig, CellRange(n))
        else
            throw(XLSXError("Invalid non-contiguous range: $n."))
        end
    end

    return NonContiguousRange(s, noncontig)
end