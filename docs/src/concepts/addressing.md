```@meta
DocTestSetup = :(import XLSX)
```

# Addressing cells and ranges

Most XLSX.jl functions take the cells they act on in the same ways. This page collects
those ways in one place. The examples use this sheet:

```julia
import XLSX
xf = XLSX.newxlsx("Data")
s = xf["Data"]
s["A1"] = ["Name", "Score"]
s["A2:A4"] = ["Ann", "Bob", "Cy"]
s["B2:B4"] = [7, 9, 4]
XLSX.addDefinedName(xf, "Scores", "Data!B2:B4")
```

## Reading cells

Index a worksheet with a cell name, a range, or row and column numbers:

| Form | Example | Gives |
|:-----|:--------|:------|
| a cell | `s["B2"]` or `s[2, 2]` | the value, `7` |
| a range | `s["A1:B2"]` or `s[1:2, 1:2]` | a `2×2` matrix |
| whole columns | `s["A:A"]`, `s[:, 2]` | the column, within the sheet's used range |
| whole rows | `s["2:3"]`, `s[2, :]` | the rows, within the sheet's used range |
| the used range | `s[:]` | every cell from the top-left to the bottom-right value |
| chosen rows or columns | `s[[2, 4], 1]`, `s[2:2:4, 1]` | the values in those cells |
| several ranges | `s["A1,B3"]` | a vector with one matrix per range |
| a defined name | `s["Scores"]` | the cells it refers to |

Index the `XLSXFile` instead to name the sheet in the reference, or to use a
workbook-scoped name:

```julia
xf["Data!A1:B2"]      # a range on the sheet Data
xf["Data!B:B"]        # a column
xf["Scores"]          # a defined name
```

[`XLSX.readdata`](@ref) does the same straight from a file:
`XLSX.readdata("file.xlsx", "Data!A1:B2")`. To read a whole table without knowing
its range, use [`XLSX.readtable`](@ref) instead (see [Read Tabular Data](@ref)).

## Writing cells

Assignment takes the same forms. What is written depends on the value:

```julia
s["A1"] = 1                       # one value into one cell
s["A1"] = [1, 2, 3]               # a vector fills across a row, from the cell: A1:C1
s["A1", dim=1] = [1, 2, 3]        # dim=1 fills down a column instead: A1:A3
s["A1"] = [1 2; 3 4]              # a matrix fills a block: A1:B2
s["H1:H3"] = ["a", "b", "c"]      # a vector into a named range, of the same length
s["J1:K2"] = 0                    # one value fills every cell in the range
s[1:3, 13] = 1:3                  # rows and columns by number
```

Writing needs a workbook you can change: see
[Creating and editing workbooks](../tutorial/creatingWorkbooks.md).

## Formatting and other functions

The functions that format cells (`setFont`, `setFill`, `setBorder`, `setAlignment`,
`setFormat` and their `setUniform…` forms), and [`XLSX.setCellProtection`](@ref),
accept all of these. Whole columns and rows, and `:`, reach only as far as the sheet's
used range:

```julia
XLSX.setFont(s, "A1"; bold=true)              # a cell
XLSX.setFont(s, "A1:B2"; bold=true)           # a range
XLSX.setFont(s, "A:B"; bold=true)             # whole columns
XLSX.setFont(s, "2:3"; bold=true)             # whole rows
XLSX.setFont(s, "A1,B3"; bold=true)           # several cells or ranges
XLSX.setFont(s, "A1:A2,B3:B4"; bold=true)
XLSX.setFont(s, "Scores"; bold=true)          # a defined name
XLSX.setFont(s, 2, 2; bold=true)              # row and column numbers
XLSX.setFont(s, 1:2, 1:2; bold=true)          # ranges of numbers
XLSX.setFont(s, [2, 4], 1; bold=true)         # vectors of numbers
XLSX.setFont(s, 2:2:4, 1; bold=true)          # step ranges
XLSX.setFont(s, :, 1; bold=true)              # every row of column 1
XLSX.setFont(xf, "Data!A1:B2"; bold=true)     # given the file, name the sheet
```

Applied to more than one cell, the formatting functions return `-1`; applied to one,
they return the id of the format used. [Formatting multiple cells at once](@ref) shows
how the formats of the cells in a range combine.

## Rows and columns

Functions that act on whole rows or columns, such as [`XLSX.setColumnWidth`](@ref),
[`XLSX.setRowHidden`](@ref), [`XLSX.groupRows`](@ref) and [`XLSX.deleteColumns`](@ref),
take the rows or columns as numbers, ranges, vectors or strings. A string can name
rows (`"3:5"`), columns (`"B:D"`), cells (whose rows or columns are used) or a defined
name. See [Which rows and columns](@ref) for the full list.

```julia
XLSX.setColumnWidth(s, "A:B"; width=20)
XLSX.setColumnWidth(s, 1:2; width=20)
XLSX.setRowHeight(s, "A2:A5"; height=25)     # rows 2 to 5; the columns are ignored
XLSX.setRowHeight(s, :; height=15)           # every row
```

## Cell references in code

[`XLSX.CellRef`](@ref) holds a cell address, made from a name or from row and column
numbers. [`XLSX.row_number`](@ref) and [`XLSX.column_number`](@ref) take it apart:

```jldoctest cellref
julia> ref = XLSX.CellRef(3, 2)
B3

julia> XLSX.row_number(ref), XLSX.column_number(ref)
(3, 2)
```

## See also

- API: [Cells and data](../api/data.md)
- Related guides: [Named ranges](../formulae/namedRanges.md), [Cell formats](../formatting/cellFormatting.md)
