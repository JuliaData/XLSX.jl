# FAQ and troubleshooting

## Opening and saving

### Excel says "We found a problem with some content" when opening my file

Excel found something in the file it doesn't accept, and offers to repair it. The usual
cause is a formula Excel can't read, such as a missing bracket or a function name it
doesn't know: [`XLSX.setFormula`](@ref) writes formulas as given, without checking them.
Test the formula in Excel and copy it from the formula bar, and see the note at the end
of [Using Formulas](formulae/formulas.md).

If no formula is to blame, please
[open an issue](https://github.com/juliadata/XLSX.jl/issues) with the code that made the
file, or the file itself.

### "looks like an old XLS file (not XLSX)"

XLSX.jl reads the Office Open XML formats (`.xlsx`, `.xlsm`, `.xltx` and `.xltm`), not
the older `.xls` format. Open the file in Excel or LibreOffice and save it as `.xlsx`.

A `.xlsx` file protected with a password to open is stored in the same container as an
`.xls` file, so it gives the same message. Decrypt it first, for example with
[XLSXDecrypt.jl](https://github.com/TimG1964/XLSXDecrypt.jl).

### "Cannot write to a cell: the XLSXFile is not writable"

The file was opened read-only, by [`XLSX.readxlsx`](@ref) or by
[`XLSX.openxlsx`](@ref) in its default `mode="r"`. Open it with `mode="rw"`, or with
[`XLSX.opentemplate`](@ref). See [Creating and editing workbooks](tutorial/creatingWorkbooks.md).

### "Output file … already exists"

[`XLSX.writexlsx`](@ref) and [`XLSX.writetable`](@ref) don't replace an existing file
unless you pass `overwrite=true`.

### My file lost something after I edited it in `mode="rw"`

`openxlsx(...; mode="rw") do … end` saves over the original when the block ends. Parts
of a workbook that XLSX.jl doesn't handle (see [Supported features](features.md)) may not
survive. To be safe, read with [`XLSX.opentemplate`](@ref) and save to a new name with
[`XLSX.writexlsx`](@ref), then check the result before replacing the original.

If the block throws an error, the file is not saved and stays as it was.

### Reading a large file is slow, or runs out of memory

Open it with `enable_cache=false` and read it row by row with [`XLSX.eachrow`](@ref) or
[`XLSX.eachtablerow`](@ref): see
[Reading Large Excel Files and Caching](@ref).

## Values

### Cells with formulas read as `missing`

XLSX.jl doesn't calculate formulas. A formula written with [`XLSX.setFormula`](@ref) has
no value until Excel calculates it, which it does when it opens the file. A file saved by
Excel holds the values it last calculated, and those are read normally.

### Dates read as numbers

Excel stores a date as a number of days, and shows it as a date because of its number
format. XLSX.jl returns a `Date`, `DateTime` or `Time` when the cell has a date or time
format, and the number otherwise. To show the number as a date in Excel, give the cell a
date format (see [Number formats](@ref)). To convert it in Julia, for a workbook using
Excel's usual 1900 date system:

```julia
using Dates
Date(1899, 12, 30) + Day(round(Int, x))    # x is the number read from the cell
```

### Excel shows `#NAME?` in a cell

Excel doesn't recognise a name in the formula. Either the formula uses a defined name
that has been deleted (see [Deleting names](@ref)), or a function Excel doesn't know by
that name. Most newer functions, such as `SORT` and `UNIQUE`, are handled by
`setFormula`; for others, see [Newer functions](@ref).

### `xf["Name"]` gives a worksheet, not my named range

A string given to an `XLSXFile` that is a sheet's name finds the sheet. Choose defined
names that differ from the sheet names. See [Reading names](@ref).

## Formatting

### A formatting function returned `-1`

The formatting functions return the id of the format they applied when given one cell,
and `-1` when given several, since the cells may end up with different formats. It isn't
an error.

### Column widths and row heights are slightly different from what I set

Excel adds padding to the widths and heights it shows, and XLSX.jl approximates it. See
[Column width and row height](formatting/widthAndHeight.md).

### My image is stretched or squashed

An image's size is worked out assuming Excel's default cell size, so over resized rows
or columns it is stretched to fit the cells it spans. See [Images](@ref).

### Rows stay hidden after I remove a filter

Rows that a filter hid in Excel stay hidden when the filter is removed, as in Excel.
Show them with [`XLSX.setRowHidden`](@ref)`(sheet, rows; hidden=false)`.

### A chart I created uses older Office colours

A chart follows the workbook's theme. A workbook from [`XLSX.newxlsx`](@ref) has an
older Office theme than current Excel, so a chart created in it uses that theme's
colours. A chart added to a workbook Excel created matches Excel's own. See
[Themes](@ref).

## Getting help

If your question isn't here, [open an issue](https://github.com/juliadata/XLSX.jl/issues)
with a short piece of code that shows the problem.
