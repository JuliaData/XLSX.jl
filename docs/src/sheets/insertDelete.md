# Inserting and deleting rows and columns

[`XLSX.insertRows`](@ref), [`XLSX.deleteRows`](@ref),
[`XLSX.insertColumns`](@ref) and [`XLSX.deleteColumns`](@ref) work as Excel's
*Insert* and *Delete* commands do: the cells below (or to the right) move, and
everything in the workbook that refers to them follows. Open the file with
`mode="rw"` (or start from [`XLSX.newxlsx`](@ref)).

```julia
XLSX.openxlsx("budget.xlsx", mode="rw") do xf
    sheet = xf["Budget"]
    XLSX.insertRows(sheet, 5)              # a new, empty row 5
    XLSX.insertRows(sheet, 10; count=3)    # three new rows, 10 to 12
    XLSX.deleteRows(sheet, "20:22")
    XLSX.insertColumns(sheet, "C")
    XLSX.deleteColumns(sheet, [2, 6])      # columns B and F
end
```

Rows and columns to delete can be given as numbers, ranges, vectors, or strings
naming rows (`"3:5,9"`), columns (`"B:C"`), cells or a defined name.

## What follows the moved cells

- Formulas in any sheet of the workbook, whether their references are relative
  or absolute (`$A$5` moves too, as in Excel).
- Row heights, hidden and grouped rows, and column widths.
- Merged cells, conditional formats, data validations, hyperlinks, the AutoFilter,
  frozen panes and page breaks.
- Defined names, including the print area and print titles.
- The ranges plotted by charts.

A range that spans the place where rows are inserted grows to include them:
`=SUM(A1:A10)` becomes `=SUM(A1:A11)` when a row is inserted at row 5. Inserting
just after a range leaves it as it is, as in Excel.

## Deleting

A range that loses some of its rows shrinks. A reference to a cell that is
deleted becomes `#REF!`:

```julia
julia> XLSX.getFormula(sheet, "F2")
"=A5*2"

julia> XLSX.deleteRows(sheet, 5);

julia> XLSX.getFormula(sheet, "F2")
"=#REF!*2"
```

A merged range, conditional format, data validation, hyperlink or AutoFilter
whose cells are all deleted goes too.

## Limits

- The new rows and columns are empty and unformatted. (Excel gives them the
  format of the row above, or the column to the left.)
- Inserting rows or columns that would push data past the last row or column
  Excel allows is an error, as in Excel.
- References written as text, such as `INDIRECT("A5")`, and references in other
  workbooks don't change, as in Excel. Nor do 3-D references (`Sheet1:Sheet3!A5`).
- A shared formula (one Excel stores once for a block of cells) stays shared when
  the edit moves all its cells alike. One the edit would change unevenly, say
  because its cells refer across the inserted rows, or rows inside its block are
  deleted, is stored as a separate formula in each of its cells afterwards. The
  formulas are the same; the file may be a little larger.
- Sheets with Excel Tables, images, charts or comments can't be edited this way
  yet.
