
struct CellPosition
    row::Int
    column::Int
end

"""
    CellRef(n::AbstractString)
    CellRef(row::Int, col::Int)

A `CellRef` represents a cell location given by row and column identifiers.

`CellRef("B6")` indicates a cell located at column `2` and row `6`.

These row and column integers can also be passed directly to the `CellRef` constructor: `CellRef(6,2) == CellRef("B6")`.

Finally, a convenience macro `@ref_str` is provided: `ref"B6" == CellRef("B6")`.

# Examples

```julia
cn = XLSX.CellRef("AB1")
println( XLSX.row_number(cn) ) # will print 1
println( XLSX.column_number(cn) ) # will print 28
println( string(cn) ) # will print out AB1
println( cellname(cn) ) # will print out AB1

cn = XLSX.CellRef(1, 28)
println( XLSX.row_number(cn) ) # will print 1
println( XLSX.column_number(cn) ) # will print 28
println( string(cn) ) # will print out AB1
println( cellname(cn) ) # will print out AB1

cn = XLSX.ref"AB1"
println( XLSX.row_number(cn) ) # will print 1
println( XLSX.column_number(cn) ) # will print 28
println( string(cn) ) # will print out AB1
println( cellname(cn) ) # will print out AB1
```

"""
struct CellRef
    row_number::Int32
    column_number::Int32
end

abstract type AbstractCellDataFormat end

struct EmptyCellDataFormat <: AbstractCellDataFormat end

# Keeps track of formatting information.
struct CellDataFormat <: AbstractCellDataFormat
    id::UInt32
end

abstract type AbstractFormula end
abstract type ExplicitFormula <: AbstractFormula end

struct EmptyFormula <: AbstractFormula end

"""
A default formula simply storing the formula string.
"""
mutable struct Formula <: ExplicitFormula
    formula::String
    type::Union{String,Nothing} # usually nothing but has value "array" for dynamic array functions.
    ref::Union{String,Nothing} # usually nothing but refers to the "spill" range for dynamic array functions.
    unhandled::Union{Dict{String,String},Nothing}
end
function Formula()
    return EmptyFormula()
end
function Formula(s::String)
    return Formula(s, nothing, nothing, nothing)
end


"""
The formula in this cell was defined somewhere else; we simply reference its ID.
"""
mutable struct FormulaReference <: AbstractFormula
    id::Int
    unhandled::Union{Dict{String,String},Nothing}
end

"""
Formula that is defined once and referenced in all cells given by the cell range given in `ref` and using the same `id`.
"""
mutable struct ReferencedFormula <: ExplicitFormula
    formula::String
    id::Int
    ref::String # actually a CellRange, but defined later --> change if at some point we want to actively change formulae
    unhandled::Union{Dict{String,String},Nothing}
end

struct CellFormula# <: AbstractFormula
    value::T where T<:AbstractFormula
    styleid::AbstractCellDataFormat
end

# Keeps track of external references in formulas.
struct ExternalRef
    index::Int          # the [n] index in the formula
    sheet::String       # sheet name
    full::String        # raw "[n]Sheet!$A$1" formula element
end


mutable struct CellFont
    fontId::Int
    font::Dict{String, Union{Dict{String, String}, Nothing}} # fontAttribute -> (attribute -> value)
    applyFont::String

    function CellFont(fontid::Int, font::Dict{String, Union{Dict{String, String}, Nothing}}, applyFont::String)
        return new(fontid, font, applyFont)
    end
end

# A border position element (e.g. `top` or `left`) has a style attribute, but `color` is a child element.
# The `color` element has an attribute (e.g. `rgb`) that defines the color of the border.
# These are both stored in the `border` field of `CellBorder`. The key for the color element
# will vary depending on how the color is defined (e.g. `rgb`, `indexed`, `auto`, etc.).
# Thus, for example, `"top" => Dict("style" => "thin", "rgb" => "FF000000")`
mutable struct CellBorder
    borderId::Int
    border::Dict{String, Union{Dict{String, String}, Nothing}} # borderAttribute -> (attribute -> value)
    applyBorder::String

    function CellBorder(borderid::Int, border::Dict{String, Union{Dict{String, String}, Nothing}}, applyBorder::String)
        return new(borderid, border, applyBorder)
    end
end

# A fill has a pattern type attribute and two children fgColor and bgColor, each with 
# one or two attributes of their own. These color attributes are pushed in to the Dict 
# of attributes with either `fg` or `bg` prepended to their name to support later 
# reconstruction of the xml element.
mutable struct CellFill
    fillId::Int
    fill::Dict{String, Union{Dict{String, String}, Nothing}} # fillAttribute -> (attribute -> value)
    applyFill::String

    function CellFill(fillid::Int, fill::Dict{String, Union{Dict{String, String}, Nothing}}, applyfill::String)
        return new(fillid, fill, applyfill)
    end
end
mutable struct CellFormat
    numFmtId::Int
    format::Dict{String, Union{Dict{String, String}, Nothing}} # fillAttribute -> (attribute -> value)
    applyNumberFormat::String

    function CellFormat(formatid::Int, format::Dict{String, Union{Dict{String, String}, Nothing}}, applynumberformat::String)
        return new(formatid, format, applynumberformat)
    end
end

mutable struct CellAlignment # Alignment is part of the cell style `xf` so doesn't need an Id
    alignment::Dict{String, Union{Dict{String, String}, Nothing}} # alignmentAttribute -> (attribute -> value)
    applyAlignment::String

    function CellAlignment(alignment::Dict{String, Union{Dict{String, String}, Nothing}}, applyalignment::String)
        return new(alignment, applyalignment)
    end
end

# Anything that can be given as a color in cell formatting, conditional formats and rich text.
const CellColor = Union{AbstractString,Symbol,Colors.Colorant,SchemeColor}

# The value of a border side keyword, e.g. `["style" => "thin", "color" => "red"]`.
const BorderKw = Union{Nothing,Vector{<:Pair{String}}}

# The keywords of a formatting setter (`setFont`, `setFill`, ...), gathered once at
# the public entry (see `define_percell_spec_methods`). The address methods forward
# their keywords through many helpers, each compiled again for every new set of
# keywords; forwarding a single `spec` lets them compile once. Each is made by
# `S(caller; kw...)`, which takes the setter's keywords with their usual types and
# defaults, and names `caller` in the error for an unrecognized keyword.
struct FontSpec
    bold::Union{Nothing,Bool}
    italic::Union{Nothing,Bool}
    under::Union{Nothing,String}
    strike::Union{Nothing,Bool}
    size::Union{Nothing,Real}
    color::Union{Nothing,CellColor}
    name::Union{Nothing,String}
    vertAlign::Union{Nothing,String}
end
function FontSpec(caller::Symbol;
    bold::Union{Nothing,Bool}=nothing,
    italic::Union{Nothing,Bool}=nothing,
    under::Union{Nothing,String}=nothing,
    strike::Union{Nothing,Bool}=nothing,
    size::Union{Nothing,Real}=nothing,
    color::Union{Nothing,CellColor}=nothing,
    name::Union{Nothing,String}=nothing,
    vertAlign::Union{Nothing,String}=nothing,
    kw...
)
    _check_keywords(caller, FontSpec, kw)
    return FontSpec(bold, italic, under, strike, size, color, name, vertAlign)
end

struct FillSpec
    pattern::Union{Nothing,String}
    fgColor::Union{Nothing,CellColor}
    bgColor::Union{Nothing,CellColor}
end
function FillSpec(caller::Symbol;
    pattern::Union{Nothing,String}=nothing,
    fgColor::Union{Nothing,CellColor}=nothing,
    bgColor::Union{Nothing,CellColor}=nothing,
    kw...
)
    _check_keywords(caller, FillSpec, kw)
    return FillSpec(pattern, fgColor, bgColor)
end

struct AlignmentSpec
    horizontal::Union{Nothing,String}
    vertical::Union{Nothing,String}
    wrapText::Union{Nothing,Bool}
    shrink::Union{Nothing,Bool}
    indent::Union{Nothing,Int}
    rotation::Union{Nothing,Int}
end
function AlignmentSpec(caller::Symbol;
    horizontal::Union{Nothing,String}=nothing,
    vertical::Union{Nothing,String}=nothing,
    wrapText::Union{Nothing,Bool}=nothing,
    shrink::Union{Nothing,Bool}=nothing,
    indent::Union{Nothing,Int}=nothing,
    rotation::Union{Nothing,Int}=nothing,
    kw...
)
    _check_keywords(caller, AlignmentSpec, kw)
    return AlignmentSpec(horizontal, vertical, wrapText, shrink, indent, rotation)
end

struct FormatSpec
    format::Union{Nothing,String}
end
function FormatSpec(caller::Symbol; format::Union{Nothing,String}=nothing, kw...)
    _check_keywords(caller, FormatSpec, kw)
    return FormatSpec(format)
end

struct BorderSpec
    outside::BorderKw
    allsides::BorderKw
    left::BorderKw
    right::BorderKw
    top::BorderKw
    bottom::BorderKw
    diagonal::BorderKw
end
function BorderSpec(caller::Symbol;
    outside::BorderKw=nothing, allsides::BorderKw=nothing,
    left::BorderKw=nothing, right::BorderKw=nothing,
    top::BorderKw=nothing, bottom::BorderKw=nothing,
    diagonal::BorderKw=nothing,
    kw...
)
    _check_keywords(caller, BorderSpec, kw)
    return BorderSpec(outside, allsides, left, right, top, bottom, diagonal)
end

function _check_keywords(caller::Symbol, ::Type{S}, kw) where {S}
    isempty(kw) && return nothing
    bad = join(("`$k`" for k in keys(kw)), ", ")
    valid = join(("`$k`" for k in fieldnames(S)), ", ")
    throw(XLSXError("Unrecognized keyword argument$(length(kw) == 1 ? "" : "s") $bad for `$caller`. Valid keywords are $valid."))
end

@enum CellValueType::UInt8 begin
    CT_EMPTY = 0
    CT_STRING = 1
    CT_FLOAT = 2
    CT_INT = 3
    CT_BOOL = 4
    CT_DATE = 5
    CT_TIME = 6
    CT_DATETIME = 7
    CT_ERROR = 8
end

"""
    CellErrorType

Excel's error values, such as `#REF!`. [`getChartRanges`](@ref) returns
`XL_REF` for a chart reference that Excel has replaced with `#REF!`, as it does
when the referenced sheet is deleted.
"""
@enum CellErrorType::UInt64 begin
    XL_NULL = 1
    XL_DIV0 = 2
    XL_VALUE = 3
    XL_REF = 4
    XL_NAME = 5
    XL_NUM = 6
    XL_NA = 7
    XL_SPILL = 8 # Turns out #SPILL is not an official error. These will return #VALUE errors
end

"`#NULL!`: an intersection of ranges that do not intersect."
XL_NULL
"`#DIV/0!`: division by zero."
XL_DIV0
"`#VALUE!`: an argument of the wrong type."
XL_VALUE
"`#REF!`: a reference to a cell or sheet that no longer exists."
XL_REF
"`#NAME?`: an unrecognised name."
XL_NAME
"`#NUM!`: an invalid numeric value."
XL_NUM
"`#N/A`: a value that is not available."
XL_NA
"`#SPILL!`: a dynamic array that cannot spill. Not an official stored error; read as `#VALUE!`."
XL_SPILL

abstract type AbstractCell end

mutable struct Cell <: AbstractCell
    ref::CellRef 
    value::UInt64 # Needs to be `reinterpret`ed according to the `Cell.datatype`
    style::UInt32
    meta::UInt16
    datatype::CellValueType # ENUM determines how `Cell.value` needs to be `reinterpret`ed
    formula::Bool # has a formula in Workbook.formulas
end

struct EmptyCell <: AbstractCell
    ref::CellRef
end

# Keeps track of conditional formatting information.
struct DxFormat <: AbstractCellDataFormat
    id::UInt
end

"""
    CellConcreteType

Concrete supported data-types.

```julia
Union{String, Missing, Float64, Int64, Bool, Dates.Date, Dates.Time, Dates.DateTime}
```

!!! note

    In julia, the values `Inf`, `-Inf` and `NaN` are of type `Float64`. However, there is 
    no way to represent these values as numbers in Excel. Instead, these specific values 
    are eagerly converted to string representation (`"Inf"`, `"-Inf"` and `"NaN"`) as they 
    are added to an XLSXFile and they are written out as such to any saved `.xlsx` file.

"""
const CellConcreteType = Union{String, Missing, Float64, Int64, Bool, Dates.Date, Dates.Time, Dates.DateTime}
# `Int64`, not `Int`: the set of types a cell can hold must not depend on the
# host word size, or 32-bit builds would silently narrow it to `Int32`.
# `setdata!` normalises every `Integer` to `Int64`, so consumers testing a cell
# value's type must test against `Int64` too (see write.jl).

# CellValue is a Julia type of a value read from a Spreadsheet.
struct CellValue
    value::CellConcreteType
    styleid::AbstractCellDataFormat
    CellValue(value::CellConcreteType, styleid::AbstractCellDataFormat) = new(value, styleid)
    CellValue(value::Integer, styleid::AbstractCellDataFormat) = new(Int64(value), styleid)
end

#=
A `CellRange` represents a rectangular range of cells in a spreadsheet.

`CellRange("A1:C4")` denotes cells ranging from `A1` (upper left corner) to `C4` (bottom right corner).

As a convenience, `@range_str` macro is provided.

```julia
cr = XLSX.range"A1:C4"
```
=#

abstract type AbstractCellRange end
abstract type ContiguousCellRange <: AbstractCellRange end
abstract type AbstractSheetCellRange <: AbstractCellRange end
abstract type ContiguousSheetCellRange <: AbstractSheetCellRange end

struct CellRange <: ContiguousCellRange
    start::CellRef
    stop::CellRef

    function CellRange(a::CellRef, b::CellRef)

        top = row_number(a)
        bottom = row_number(b)
        left = column_number(a)
        right = column_number(b)

        if left > right || top > bottom
            throw(XLSXError("Invalid CellRange. Start cell should be at the top left corner of the range."))
        end

        return new(a, b)
    end
end

struct ColumnRange <: ContiguousCellRange
    start::Int # column number
    stop::Int  # column number

    function ColumnRange(a::Integer, b::Integer)
        if a > b 
            throw(XLSXError("Invalid ColumnRange. Start column must be located before end column."))
        end
        return new(a, b)
    end
end
struct RowRange <: ContiguousCellRange
    start::Int # row number
    stop::Int  # row number

    function RowRange(a::Integer, b::Integer)
        if a > b
            throw(XLSXError("Invalid RowRange. Start row must be located before end row."))
        end
        return new(a, b)
    end
end

"""
    SheetCellRef(ref::AbstractString)
    SheetCellRef(sheet::AbstractString, cellref::CellRef)

A cell on a named sheet: a [`CellRef`](@ref) together with the name of its sheet, in
the fields `sheet` and `cellref`.

`SheetCellRef("Sheet1!B2")` refers to cell `B2` on sheet `Sheet1`. A sheet name
containing spaces or other special characters is quoted, as in Excel:
`SheetCellRef("'Sheet 1'!B2")`, which `SheetCellRef("Sheet 1", CellRef("B2"))` also
gives.

# Examples

```julia
julia> r = XLSX.SheetCellRef("Sheet1!B2")
Sheet1!B2

julia> r.sheet
"Sheet1"

julia> r.cellref
B2
```
"""
struct SheetCellRef
    sheet::String
    cellref::CellRef
end

struct SheetCellRange <: ContiguousSheetCellRange
   sheet::String
   rng::CellRange
end

struct NonContiguousRange <: AbstractSheetCellRange
    sheet::String
    rng::Vector{Union{CellRef, CellRange}}
end

struct SheetColumnRange <: ContiguousSheetCellRange
    sheet::String
    colrng::ColumnRange
end
struct SheetRowRange <: ContiguousSheetCellRange
    sheet::String
    rowrng::RowRange
end

abstract type MSOfficePackage end

struct EmptyMSOfficePackage <: MSOfficePackage
end

#=
Relationships are defined in ECMA-376-1 Section 9.2.
This struct matches the `Relationship` tag attribute names.

A `Relationship` defines relations between the files inside a MSOffice package.
Regarding Spreadsheets, there are two kinds of relationships:

    * package level: defined in `_rels/.rels`.
    * workbook level: defined in `xl/_rels/workbook.xml.rels`.

The function `parse_relationships!(xf::XLSXFile)` is used to parse
package and workbook level relationships.
=#
struct Relationship
    Id::String
    Type::String
    Target::String
end

const CellCache = Dict{Int, Dict{Int, Cell}} # row -> ( column -> cell )

#=
Iterates over Worksheet cells. See `eachrow` method docs.
Each element is a `SheetRow`.

Implementations: SheetRowStreamIterator, WorksheetCache.
=#
abstract type SheetRowIterator end

mutable struct SheetRowStreamIteratorState{C}
    cursor::C          # in `<sheetData>`, after the last row read (single pass: advance a state once)
    rowcells::Dict{Int,Cell}
    local_formulas::Dict{SheetCellRef,AbstractFormula}
    # Sheet bounds seen so far, tracked only when the worksheet had no dimension
    # at the start of iteration. Set as the dimension when iteration completes.
    track::Bool
    row_min::Int
    row_max::Int
    col_min::Int
    col_max::Int
end

mutable struct WorksheetCacheIteratorState
    row_from_last_iteration::Int
end

mutable struct WorksheetCache{I<:SheetRowIterator} <: SheetRowIterator
    is_full::Bool # false until iterator runs to completion
    cells::CellCache # SheetRowNumber -> Dict{column_number, Cell}
    rows_in_cache::Vector{Int} # row numbers stored in cache; sorted only when `dirty` is false (see `ensure_sorted!`)
    row_ht::Dict{Int, Union{Float64, Nothing}} # Maps a row number to a row height
    row_index::Dict{Int, Int} # maps a row number to the index of the row number in rows_in_cache
    stream_iterator::I
    stream_state::Union{Nothing, SheetRowStreamIteratorState}
    dirty::Bool #indicate that data are not sorted, avoid sorting if we dont use the iterator
end

Base.isempty(wc::WorksheetCache) = isempty(wc.rows_in_cache)


# Excel Tables
"""
    TableStyleInfo

The style of an Excel [`XLSX.Table`](@ref), as the table's `style` field holds it: the
table style's `name` (e.g. `"TableStyleMedium2"`, or `nothing` for none) and whether
the first column, last column, row stripes and column stripes are emphasized, in the
fields `show_first_column`, `show_last_column`, `show_row_stripes` and
`show_column_stripes`.

Give a new table a style with the `style` keyword of [`XLSX.addtable!`](@ref).
"""
struct TableStyleInfo
    name::Union{String,Nothing}
    show_first_column::Bool
    show_last_column::Bool
    show_row_stripes::Bool
    show_column_stripes::Bool
end

"""
    Table

A native Excel Table: a named, structured range within a worksheet, created in Excel
with *Insert → Table* (or `Ctrl+T`), or in XLSX.jl with [`XLSX.addtable!`](@ref).

Obtain one with [`XLSX.table`](@ref) or [`XLSX.tables`](@ref). A `Table` conforms to the
`Tables.jl` interface, so it can be passed directly to any compatible sink (e.g.
`DataFrame`), or read with [`XLSX.gettable`](@ref), [`XLSX.getdata`](@ref) or
[`XLSX.eachtablerow`](@ref). In every case only the Table's data rows are returned: the
header row and, if present, the totals row are excluded.

# Fields
- `id::Int`: the Table's id, unique within the workbook.
- `name::String`: the Table's name, unique within the workbook and sharing a namespace
  with defined names.
- `display_name::String`: the name Excel displays. Usually identical to `name`.
- `ref::CellRange`: the Table's full extent, including its header row and, if present,
  its totals row.
- `columns::Vector{String}`: the column names, in order, taken from the header row.
- `has_totals_row::Bool`: whether the last row of `ref` is a totals row.
- `style::Union{TableStyleInfo,Nothing}`: the Table's style, or `nothing` if it has none.
- `sheet`: the `Worksheet` the Table belongs to. Cell values are read from it on demand.

A `Table` is an immutable snapshot of the Table's structure at the time it was read.
Functions that modify a Table ([`XLSX.settotals!`](@ref), [`XLSX.appendtable!`](@ref))
return an updated `Table`; any earlier one goes stale and should be discarded.

See also [`XLSX.table`](@ref), [`XLSX.tables`](@ref), [`XLSX.addtable!`](@ref),
[`XLSX.deletetable!`](@ref), [`XLSX.settotals!`](@ref), [`XLSX.appendtable!`](@ref).
"""
struct Table
    id::Int
    name::String
    display_name::String
    ref::CellRange
    columns::Vector{String}
    has_header_row::Bool          # false when the header row is hidden (headerRowCount="0"); ref then starts at the data
    has_totals_row::Bool
    style::Union{TableStyleInfo,Nothing}
    sheet # untyped to resolve circular dependency with Worksheet

end


"""
A `Worksheet` represents a reference to an Excel Worksheet.

From a `Worksheet` you can query for Cells, cell values and ranges.

# Example

```julia
xf = XLSX.readxlsx("myfile.xlsx")
sh = xf["mysheet"] # get a reference to a Worksheet
println( sh[2, 2] ) # access element "B2" (2nd row, 2nd column)
println( sh["B2"] ) # you can also use the cell name
println( sh["A2:B4"] ) # or a cell range
println( sh[:] ) # all data inside worksheet's dimension
```
"""
mutable struct Worksheet
    package::MSOfficePackage
    sheetId::Int
    relationship_id::String
    name::String
    dimension::Union{Nothing, CellRange}
    is_hidden::Bool
    cache::Union{WorksheetCache, Nothing}
    next_formula_id::Int
    unhandled_attributes::Union{Nothing,Dict{Int,Dict{String,String}}}
    sst_count::Int
    next_cf_priority::Union{Int, Nothing}
    tables_cache::Union{Vector{Table}, Nothing}   # nothing until first access to `tables(ws)`

    function Worksheet(package::MSOfficePackage, sheetId::Int, relationship_id::String, name::String, dimension::Union{Nothing, CellRange}, is_hidden::Bool)
        return new(package, sheetId, relationship_id, name, dimension, is_hidden, nothing, 0, nothing, 0, nothing, nothing)
    end
end

struct SheetRowStreamIterator <: SheetRowIterator
    sheet::Worksheet
    # Columns to decode; `nothing` = every column. A table read decodes only its own,
    # so its rows' `rowcells` hold no cell outside them.
    cols::Union{Nothing,UnitRange{Int}}
end
SheetRowStreamIterator(ws::Worksheet) = SheetRowStreamIterator(ws, nothing)

#------------------------------------------------------------------------------ sharedStrings
mutable struct SharedStringTable
    shared_strings::Vector{String}
    unformatted::Vector{String}
    index::Dict{String, Int64} # for search optimisation. Tuple of indices to handle hash collisions.
    is_loaded::Bool
end

const ValidRichTextAttributes = [:bold, :italic, :under, :strike, :vertAlign, :color, :size, :name]

"""
    RichTextRun(text::String, pairs::Union{Nothing,Vector{Pair{Symbol,Any}}}=nothing)     -> RichTextRun
    RichTextRun(text::String)                                                             -> RichTextRun

Create an instance of a RichTextRun, representing a formatted substring element (run) to form part 
of a RichTextString. Each RichTextRun defines none, one or several font attributes to apply to its text.

- `text` specifies the text of the run's substring element.
- `pairs` is a vector of formatting attributes to apply to `text` (default = `nothing`).

Valid attributes that can be defined in `pairs` are:
- `:bold` - set `:bold => true` for this run to be emboldened. Omit otherwise.
- `:italic` - set `:italic => true` for this run to be italicised. Omit otherwise.
- `:under` - set `:under => true` to underline this run. Omit otherwise.
- `:strike` - set `:strike => true` to apply strikethrough to this run. Omit otherwise.
- `:vertAlign` - whether this run is `subscript` or `superscript` (eg `:vertAlign => "superscript"`). Omit otherwise.
- `:color` - the color of this run (eg `:color => "red"`), in any of the forms listed in
  [Specifying colors](@ref). An invalid color is an error here, not when the run is written.
- `:size` - the size of the font to be used (eg `:size => 12`).
- `:name` - the name of the font to be used (e.g. `:name => "Arial"`).

Omit `pairs` to specify a run without formatting.

See also [`XLSX.RichTextString`](@ref).

# Examples
```julia
julia> rt1 = XLSX.RichTextRun("Water is H")
RichTextRun ("Water is H"  [ ])

julia> rt2 = XLSX.RichTextRun("2", [:vertAlign => "subscript"])
RichTextRun ("2"  [:vertAlign => "subscript"])

julia> rt3 = XLSX.RichTextRun("O!")
RichTextRun ("O!"  [ ])

julia> rt = XLSX.RichTextString(rt1, rt2, rt3)
RichTextString: "Water is H2O!" 
 containing 3 runs:
 Run text                 Run attributes
 -------------------------------------------------------------------------------------------
 "Water is H"             [ ]
 "2"                      [:vertAlign => "subscript"]
 "O!"                     [ ]

julia> s["A1"] = rt
RichTextString: "Water is H2O!" 
 containing 3 runs:
 Run text                 Run attributes
 -------------------------------------------------------------------------------------------
 "Water is H"             [ ]
 "2"                      [:vertAlign => "subscript"]
 "O!"                     [ ]

```
![image|320x500](../images/H2O.png)
"""
struct RichTextRun
    text::String
    atts::Union{Nothing, Dict{Symbol,Any}}
    
    function RichTextRun(text::String, pairs::Union{Nothing,Vector{Pair{Symbol,Any}}}=nothing)
        isempty(text) && throw(XLSXError("Cannot create a RichTextRun with no text."))
        if isnothing(pairs)
            new(text, nothing)
        else
            atts=Dict(pairs)
            for x in keys(atts)
                in(x, ValidRichTextAttributes) || throw(XLSXError("Unknown Rich Text Attribute: ':$x'. Valid attributes are :bold, :italic, :under, :strike, :vertAlign, :color, :size, :name."))
            end
            haskey(atts, :color) && _check_color(atts[:color])
            new(text, atts)
        end
    end

end

"""
    RichTextString(runs::RichTextRun...)      -> RichTextString
    RichTextString(runs::Vector{RichTextRun}) -> RichTextString

Create an instance of a RichTextString from a set of RichTextRuns. A RichTextString supports rich text 
formatting within a single cell and is made up of multiple substrings (runs), each with different font 
attributes. The text in the cell is the simple concatenation of the text of each run but Excel will display 
each run with its own distinct font formatting within the cell. See also [`XLSX.RichTextRun`](@ref).

If a `RichTextString` containing only one run is assigned to a cell, the text will be assigned as plain 
text and the formatting attributes will be implemented on the whole cell using [`XLSX.setFont`](@ref).

See also [`XLSX.RichTextRun`](@ref).

# Examples
```julia
julia> rt = XLSX.RichTextString(rtf1, rtf2, rtf3, rtf4) # Create a RichTextString from four separate RichTextRuns.

julia> rt = XLSX.RichTextString([rtf1, rtf2, rtf3, rtf4]) # Create a RichTextString from a vector of four RichTextRuns.

```
"""
struct RichTextString <: AbstractString
    text::String
    runs::Vector{RichTextRun}

    function RichTextString(text::String, runs::Vector{RichTextRun })
        (isempty(text) || isempty(runs)) && throw(XLSXError("Cannot create an empty RichTextString"))
        new(text, runs)
    end
end

const DefinedNameValueTypes = Union{SheetCellRef, SheetCellRange, NonContiguousRange, CellConcreteType}#Int, Float64, String, Missing}
const DefinedNameRangeTypes = Union{SheetCellRef, SheetCellRange, NonContiguousRange}

struct DefinedNameValue
    value::DefinedNameValueTypes
    isabs::Union{Bool, Vector{Bool}}
    hidden::Bool
end

"""
    DefinedName

A defined name and its definition, as returned by [`getDefinedNames`](@ref) and
[`getAllDefinedNames`](@ref).

# Fields
- `name::String` — the defined name.
- `scope::Union{Nothing,String}` — `nothing` for a workbook-scoped name, or the
  display name of the worksheet it is scoped to.
- `value::DefinedNameValueTypes` — a `SheetCellRef`, `SheetCellRange` or
  `NonContiguousRange` for a name referring to a range, or the constant itself.
- `absolute::Union{Bool,Vector{Bool}}` — whether the reference is written as an
  absolute one (`\$A\$1` rather than `A1`); a vector, one entry per part, for a
  `NonContiguousRange`. Always `false` for a constant.
- `hidden::Bool` - whether the reference is hidden (system defined).

Together these are enough to recreate the name: addDefinedName(x, dn.name, dn.value; absolute=dn.absolute). 
This holds for the names getDefinedNames returns; names Excel generates for itself cannot be recreated, 
and are excluded from that result.

This is a snapshot of the definition at the time it was read, not a live handle:
it does not track later edits, and renaming a worksheet does not update the
`scope` or `value` of a `DefinedName` already in hand.
"""
struct DefinedName
    name::String
    scope::Union{Nothing,String}
    value::DefinedNameValueTypes
    absolute::Union{Bool,Vector{Bool}}
    hidden::Bool
end

# Workbook is the result of parsing file `xl/workbook.xml`.
# The `xl/workbook.xml` will need to be updated using the Workbook_names and 
# worksheet_names from here when a workbook is saved in case any new defined 
# names have been created.
"""
    Workbook

The workbook part of an [`XLSXFile`](@ref): its sheets, shared strings, styles and
defined names, parsed from `xl/workbook.xml`. An `XLSXFile` holds one, as
`xf.workbook`. Its fields are internal: work through the file and its worksheets.
"""
mutable struct Workbook
    package::MSOfficePackage
    sheets::Vector{Worksheet}
    date1904::Bool
    relationships::Vector{Relationship}
    formulas::Dict{SheetCellRef, AbstractFormula}
    sst::SharedStringTable
    buffer_styles_is_float::Dict{Int, Bool}
    buffer_styles_is_datetime::Dict{Int, Bool}
    styles_lock::ReentrantLock
    formulas_lock::ReentrantLock
    sst_lock::ReentrantLock
    workbook_names::Dict{String, DefinedNameValue}
    # Keyed by (sheetId, name). NOT by localSheetId: the `localSheetId`
    # attribute of <definedName> is a zero-based index into <sheets>, so it
    # shifts whenever sheets are added, deleted or reordered. sheetId is stable
    # for the life of the sheet, so it is what we key on; the conversion to and
    # from localSheetId happens only at the read and write boundaries.
    worksheet_names::Dict{Tuple{Int, String}, DefinedNameValue}
    styles_xroot::Union{XML.Node, Nothing}
    num_style_index_cache::Dict{Int, CellDataFormat}
    theme_xroot::Union{XML.Node, Nothing}
    theme_colors::Union{Vector{String}, Nothing}
    theme_color_map::Union{Nothing,Dict{String,String}}
    theme_font_map::Union{Nothing,Dict{String,String}}
    cellXfs_cache::Union{Vector{XML.Node}, Nothing}   # cache for get_cellXfs_nodes
    numFmt_cache::Union{Dict{Int, String}, Nothing}   # cache for get_numFmt_cache
    style_table_cache::Dict{String, Vector{XML.Node}} # cache for fonts/borders/fills, keyed by tag ("fonts","borders","fills")
    next_table_id::Union{Int,Nothing}   # nothing until first computed
end

@enum TemplateType begin
    NotATemplate
    XLTXTemplate   # .xltx → save as .xlsx
    XLTMTemplate   # .xltm → save as .xlsm
end

"""
`XLSXFile` represents a reference to an Excel file.

It is created by using [`XLSX.readxlsx`](@ref), [`XLSX.openxlsx`](@ref), 
[`XLSX.opentemplate`](@ref) or [`XLSX.newxlsx`](@ref).

From an `XLSXFile` you can navigate to an `XLSX.Worksheet` reference
as shown in the example below.

# Example

```julia
xf = XLSX.readxlsx("myfile.xlsx")
sh = xf["mysheet"] # get a reference to a Worksheet
```
"""
mutable struct XLSXFile <: MSOfficePackage
    source::Union{AbstractString, IO}
    use_cache_for_sheet_data::Bool
    load_formulas::Bool                          # ← new
    files::Dict{String, Bool}
    data::Dict{String, Union{XML.Node, String}}
    namespace::Dict{String, Union{String, Nothing}}
    binary_data::Dict{String, Vector{UInt8}}
    # Read-only, cache-on: the `<sheetData>`-stripped stub of each worksheet, made at
    # open while `data` keeps the full XML. The first `eachrow` swaps it in after filling
    # the cache, instead of splitting the XML a second time (#462).
    sheet_stubs::Dict{String, String}
    workbook::Workbook
    relationships::Vector{Relationship}
    is_writable::Bool
    template_type::TemplateType
    uuid_rng::Random.Xoshiro

    function XLSXFile(source::Union{AbstractString, IO}, use_cache::Bool, is_writable::Bool, load_formulas::Bool=true)
        check_for_xlsx_file_format(source)
        xl = new(
            source,
            use_cache,
            load_formulas,
            Dict{String, Bool}(),
            Dict{String, Union{XML.Node, String}}(),
            Dict{String, Union{String, Nothing}}(),
            Dict{String, Vector{UInt8}}(),
            Dict{String, String}(),
            EmptyWorkbook(),
            Vector{Relationship}(),
            is_writable,
            NotATemplate,
            Random.Xoshiro(2468)
        )
        xl.workbook.package = xl
        return xl
    end
end

struct ReadFile
    node::Union{Nothing,XML.Node,String}
    raw::Union{Nothing,String}
    bin::Union{Nothing,Vector{UInt8}}
    name::String
end

#
# Iterators
#

# A worksheet row as the table code reads it: `SheetRow` (from the worksheet cache) or
# `ValueRow` (readtable's value rows, src/valuerows.jl).
abstract type AbstractSheetRow end

struct SheetRow <: AbstractSheetRow
    sheet::Worksheet
    row::Int                     # index of the row in the worksheet
    ht::Union{Float64, Nothing}  # row height
    rowcells::Dict{Int, Cell}    # column -> value
end

struct Index # for TableRowIterator - based on DataFrames.jl
    lookup::Dict{Symbol, Int} # column label -> table column index
    column_labels::Vector{Symbol}
    column_map::Dict{Int, Int} # table column index (1-based) -> sheet column index (cellref based)

    function Index(column_range::Union{ColumnRange, AbstractString}, column_labels)
        column_labels_as_syms = [ Symbol(i) for i in column_labels ]
        column_range = convert(ColumnRange, column_range)
        if length(unique(column_labels_as_syms)) != length(column_labels_as_syms)
            throw(XLSXError("Column labels must be unique."))
        end

        lookup = Dict{Symbol, Int}()
        for (i, n) in enumerate(column_labels_as_syms)
            lookup[n] = i
        end

        column_map = Dict{Int, Int}()
        for (i, n) in enumerate(column_range)
            column_map[i] = decode_column_number(n)
        end
        return new(lookup, column_labels_as_syms, column_map)
    end
end

struct TableRowIterator{I<:SheetRowIterator}
    itr::I
    index::Index
    first_data_row::Int
    stop_in_empty_row::Bool
    stop_in_row_function::Union{Nothing, Function}
    keep_empty_rows::Bool
    missing_strings::Set{String}
    resume::Union{Nothing, Tuple{SheetRow, Any}}  # pre-fetched (row, state) to start from, or nothing
end

struct TableRow
    row::Int # Index of the row in the table. This is not relative to the worksheet cell row.
    index::Index
    cell_values::Vector{CellConcreteType}
end

struct TableRowIteratorState{S}
    table_row_index::Int
    sheet_row_index::Int
    sheet_row_iterator_state::S
    missing_rows::Int # number of completely empty rows between the last row and the current row
    row_pending::Union{Nothing, AbstractSheetRow} # if the last row was empty, this is the row that was pending to be returned
end

struct XLSXTableRow
    table::Table
    row_number::Int
end

struct XLSXTableRowIterator
    table::Table
end

"""
`XLSX.DataTable` is a simple `Tables.jl` compatibledata structure to hold tabular 
data extracted from an Excel worksheet.

It is created by [`XLSX.gettable`](@ref) or, direct from a file, with [`XLSX.readtable`](@ref).

Pass a `DataTable` to any `Tables.jl` compatible sink, e.g. `DataFrame(dt)`.

"""
struct DataTable
    data::Vector{Any} # columns
    column_labels::Vector{Symbol}
    column_label_index::Dict{Symbol, Int} # column_label -> column_index

    function DataTable(
            data::Vector{Any}, # columns
            column_labels::Vector{Symbol},
        )

        if length(data) != length(column_labels)
            throw(XLSXError("Data has $(length(data)) columns but $(length(column_labels)) column labels."))
        end

        column_label_index = Dict{Symbol, Int}()
        for (i, sym) in enumerate(column_labels)
            if haskey(column_label_index, sym)
                throw(XLSXError("DataTable has repeated label for column `$sym`"))
            end
            column_label_index[sym] = i
        end

        return new(data, column_labels, column_label_index)
    end
end

struct XPathInfo
    node::XML.Node
    path::String

    function XPathInfo(node::XML.Node, path::String)
        new(node, path)
    end
end

"""
    XLSXError(msg::String)

The exception type XLSX.jl throws for every error it raises itself: a malformed
or unsupported file, an argument that names something the workbook does not
contain, a value outside what the format allows, or an operation the file's mode
does not permit.

One type covers all of them, so `e isa XLSXError` in a `catch` distinguishes an
error the package raised from one escaping out of Julia or a dependency. The
message says what went wrong and, where it helps, what was found instead.

# Fields
- `msg::String` — the message, printed after `XLSXError: `.

# Example

    julia> XLSX.Charts.getChartAxis(c, 99)
    ERROR: XLSXError: Chart `chart1` has no axis with axId 99. Found: 1, 2.
"""
struct XLSXError <: Exception
    msg::String
end
Base.showerror(io::IO, e::XLSXError) = print(io, "XLSXError: ",e.msg)

mutable struct FileArray <: AbstractVector{UInt8}
    const filename::String
    const offset::Int64
    const len::Int64
    const lock::ReentrantLock
    block_start::Int64      # position in this array of `block[1]`; -1 while nothing is cached
    const block::Vector{UInt8}
end