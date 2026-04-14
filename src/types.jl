
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

# A border postion element (e.g. `top` or `left`) has a style attribute, but `color` is a child element.
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

abstract type AbstractCell end

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
Union{String, Missing, Float64, Int, Bool, Dates.Date, Dates.Time, Dates.DateTime}
```

!!! note

    In julia, the values `Inf`, `-Inf` and `NaN` are of type `Float64`. However, there is 
    no way to represent these values as numbers in Excel. Instead, on read, these specific 
    values are eagerly converted to string representation (`"Inf"`, `"-Inf"` and `"NaN"`). 
    They are represented as strings in the XLSXFile and are written out as such to any saved 
    `.xlsx` file.

"""
const CellConcreteType = Union{String, Missing, Float64, Int64, Bool, Dates.Date, Dates.Time, Dates.DateTime}

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

mutable struct SheetRowStreamIteratorState
    next_rownode::Union{Nothing, XML.LazyNode} # Worksheet row being processed
    rowcells::Dict{Int,Cell}
    lock::ReentrantLock
end

mutable struct WorksheetCacheIteratorState
    row_from_last_iteration::Int
end

mutable struct WorksheetCache{I<:SheetRowIterator} <: SheetRowIterator
    is_full::Bool # false until iterator runs to completion
    cells::CellCache # SheetRowNumber -> Dict{column_number, Cell}
    rows_in_cache::Vector{Int} # ordered vector with row numbers that are stored in cache
    row_ht::Dict{Int, Union{Float64, Nothing}} # Maps a row number to a row height
    row_index::Dict{Int, Int} # maps a row number to the index of the row number in rows_in_cache
    stream_iterator::I
    stream_state::Union{Nothing, SheetRowStreamIteratorState}
    dirty::Bool #indicate that data are not sorted, avoid sorting if we dont use the iterator
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
    package::MSOfficePackage # parent XLSXFile
    sheetId::Int
    relationship_id::String # r:id="rId1"
    name::String
    dimension::Union{Nothing, CellRange}
    is_hidden::Bool
    cache::Union{WorksheetCache, Nothing}
    unhandled_attributes::Union{Nothing,Dict{Int,Dict{String,String}}} # row => attributes(name=>value)
    sst_count::Int # number of cells containing a shared string

    function Worksheet(package::MSOfficePackage, sheetId::Int, relationship_id::String, name::String, dimension::Union{Nothing, CellRange}, is_hidden::Bool)
        return new(package, sheetId, relationship_id, name, dimension, is_hidden, nothing, nothing, 0)
    end
end

struct SheetRowStreamIterator <: SheetRowIterator
    sheet::Worksheet
end

#------------------------------------------------------------------------------ sharedStrings
mutable struct SharedStringTable
    shared_strings::Vector{String}
    index::Dict{String, Int64} # for search optimisation. Tuple of indices to handle hash collisions.
    is_loaded::Bool
end
struct SstToken
    n::XML.LazyNode
    idx::Int
end
struct Sst
    formatted::String
    idx::Int
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
- `:vertAlign` - whether this run is `subscript` or `superscript` (eg `:vertAling => "superscript"`). Omit otherwise.
- `:color` - the color of this run (eg `:color => "red"`).
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
end

# Workbook is the result of parsing file `xl/workbook.xml`.
# The `xl/workbook.xml` will need to be updated using the Workbook_names and 
# worksheet_names from here when a workbook is saved in case any new defined 
# names have been created.
mutable struct Workbook
    package::MSOfficePackage # parent XLSXFile
    sheets::Vector{Worksheet} # workbook -> sheets -> <sheet name="Sheet1" r:id="rId1" sheetId="1"/>. sheetId determines the index of the WorkSheet in this vector.
    date1904::Bool              # workbook -> workbookPr -> attribute date1904 = "1" or absent
    relationships::Vector{Relationship} # contains workbook level relationships
    formulas::Dict{SheetCellRef, AbstractFormula} # eg SheetCellRef("mysheet!A1") => formula (not deduped)
    sst::SharedStringTable # shared string table
    buffer_styles_is_float::Dict{Int, Bool}      # cell style -> true if is float
    buffer_styles_is_datetime::Dict{Int, Bool}   # cell style -> true if is datetime
    styles_lock::ReentrantLock # Prevent race conditions when accessing Styles
    workbook_names::Dict{String, DefinedNameValue} # definedName
    worksheet_names::Dict{Tuple{Int, String}, DefinedNameValue} # definedName. (sheetId, name) -> value.
    styles_xroot::Union{XML.Node, Nothing}
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
    use_cache_for_sheet_data::Bool # indicates whether Worksheet.cache will be fed while reading worksheet cells.
    files::Dict{String, Bool} # maps filename => isread bool
    data::Dict{String, XML.Node} # maps filename => XMLDocument (with row/sst elements removed)
    binary_data::Dict{String, Vector{UInt8}} # maps filename => file content in bytes
    workbook::Workbook
    relationships::Vector{Relationship} # contains package level relationships
    is_writable::Bool # indicates whether this XLSX file can be edited
    uuid_rng::Random.Xoshiro # rng used to generate uuids for this file

    function XLSXFile(source::Union{AbstractString, IO}, use_cache::Bool, is_writable::Bool)
        check_for_xlsx_file_format(source)
        xl = new(source, use_cache, Dict{String, Bool}(), Dict{String, XML.Node}(), Dict{String, Vector{UInt8}}(), EmptyWorkbook(), Vector{Relationship}(), is_writable, Random.Xoshiro(2468))
        xl.workbook.package = xl
        return xl
    end
end


struct ReadFile
    node::Union{Nothing,XML.Node}
    raw::Union{Nothing,XML.Raw}
    bin::Union{Nothing,Vector{UInt8}}
    name::String
end

#
# Iterators
#

struct SheetRow
    sheet::Worksheet
    row::Int                  # index of the row in the worksheet
    ht::Union{Float64, Nothing}   # row height
    rowcells::Dict{Int, Cell} # column -> value
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
    missing_strings::Set{String} # issue 90
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
    row_pending::Union{Nothing, SheetRow} # if the last row was empty, this is the row that was pending to be returned
end

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

struct xpath
    node::XML.Node
    path::String

    function xpath(node::XML.Node, path::String)
        new(node, path)
    end
end

struct XLSXError <: Exception
    msg::String
end
Base.showerror(io::IO, e::XLSXError) = print(io, "XLSXError: ",e.msg)

struct FileArray <: AbstractVector{UInt8}
    filename::String
    offset::Int64
    len::Int64
end

mutable struct Locked{T}
    value::T
    lock::ReentrantLock
    Locked(x::T) where {T} = new{T}(x, ReentrantLock())
end

function withlock(f, obj::Locked)
    lock(obj.lock) do
        f(obj.value)
    end
end
