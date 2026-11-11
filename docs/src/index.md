
# XLSX.jl

## Introduction

**XLSX.jl** is a Julia package to read and write
[Excel](https://www.microsoft.com/microsoft-365/excel) spreadsheet files.

Internally, an Excel XLSX file is just a
[Zip](https://en.wikipedia.org/wiki/Zip_(file_format)) file with a set of
[XML](https://en.wikipedia.org/wiki/XML) files inside.
The formats for these XML files are described in
the [Standard ECMA-376](https://ecma-international.org/publications-and-standards/standards/ecma-376/).

This package follows ECMA-376 to parse and generate XLSX files.

## Features

* Read and write cell values, ranges and whole tables through the
  [Tables.jl](https://github.com/JuliaData/Tables.jl) interface, so data goes straight to and from a `DataFrame`
  (see the [XLSX Tutorial](tutorial/XLSXtutorial.md))
* Stream the rows of large files with `eachrow` and `eachtablerow`, without loading the whole sheet
* [Cell formatting](formatting/cellFormatting.md) and [conditional formats](formatting/conditionalFormatting.md)
* [Formulas](formulae/formulas.md) and [named ranges](formulae/namedRanges.md)
* [Insert and delete rows and columns](sheets/insertDelete.md), with formulas and everything else that refers to them updated
* [Data validation](sheets/dataValidation.md), including drop-down lists
* [AutoFilter](sheets/autoFilter.md) buttons on plain ranges
* [Sheet visibility and tab colours](sheets/worksheets.md), [protection](sheets/protection.md) and [page setup](sheets/printing.md)
* Native [Excel Tables](tables/excelTables.md)
* [Merged cells](formatting/mergedCells.md), [column widths and row heights](formatting/widthAndHeight.md),
  [hidden and grouped rows and columns](formatting/hideAndGroup.md)
  and [freeze/split panes](formatting/freezeAndSplitPanes.md)
* [Adding, copying, renaming, moving and deleting sheets](sheets/worksheets.md)
* [Hyperlinks](sheets/hyperlinks.md), [images](sheets/images.md) and [cell comments](sheets/comments.md)
* Native [Excel charts](charts/excelCharts.md): read, format and create (experimental)
* [FileIO.jl](https://github.com/JuliaIO/FileIO.jl) integration (see the [FileIO Tutorial](tutorial/FileIOtutorial.md))

[Supported features](features.md) lists every Excel feature, what XLSX.jl can do with it,
and what it can't do.

## Quick Start

```julia
using XLSX, DataFrames

df = DataFrame(name = ["a", "b", "c"], value = [1.5, 2.5, 3.5])
XLSX.writetable("data.xlsx", df; sheetname = "Data", overwrite = true)

df2 = DataFrame(XLSX.readtable("data.xlsx", "Data"))

XLSX.openxlsx("data.xlsx", mode = "rw") do xf
    sheet = xf["Data"]
    sheet["C1"] = "Total"
    XLSX.setFormula(sheet, "C2", "=SUM(B2:B4)")
    XLSX.setFont(sheet, "C1"; bold = true)
end
```

## I want to…

The [Cookbook](cookbook.md) answers the most common of these with a few lines of code each.

### Read and write data

| I want to… | Where to look |
|:-----------|:--------------|
| read a sheet into a `DataFrame` | [Read Tabular Data](@ref) |
| create a new workbook, or edit and save an existing one | [Creating and editing workbooks](tutorial/creatingWorkbooks.md) |
| write a `DataFrame` (or several) to a file | [Writing Excel Files](@ref) |
| read individual cells or a range | [Getting Started](@ref) |
| edit an existing file | [Edit Existing Files](@ref) |
| read a file too big for memory | [Reading Large Excel Files and Caching](@ref) |
| use FileIO's `load` and `save` | [FileIO Tutorial](tutorial/FileIOtutorial.md) |

### Data features

| I want to… | Where to look |
|:-----------|:--------------|
| create an Excel Table | [Creating Tables](@ref) |
| add a totals row to a Table | [Totals rows](@ref) |
| write a formula | [Using a simple formula](@ref) |
| use `SORT`, `UNIQUE` and other spilling functions | [Spill ranges](@ref) |
| define or read a named range | [Named ranges](formulae/namedRanges.md) |
| make a drop-down list | [Drop-down lists](@ref) |
| restrict a cell to numbers, dates or text length | [Numbers, dates, times and text length](@ref) |
| add filter buttons to a header row | [Filtering](sheets/autoFilter.md) |

### Formatting

| I want to… | Where to look |
|:-----------|:--------------|
| make text bold, change font, fill or border | [Setting format attributes of a cell](@ref) |
| format numbers as currency, percentage, dates or a custom format | [Number formats](@ref) |
| refer to cells, ranges, rows, columns or names | [Addressing cells and ranges](concepts/addressing.md) |
| format a whole range at once | [Formatting multiple cells at once](@ref) |
| choose a color, including theme colors | [Specifying colors](@ref) |
| color part of the text in a cell | [Rich text formats](@ref) |
| highlight cells by value (conditional formatting) | [Dynamic conditional formats](@ref) |
| list or remove conditional formats | [Reading and clearing conditional formats](@ref) |
| add data bars, color scales or icon sets | [Data Bar](@ref), [Color Scale](@ref), [Icon Set](@ref) |
| merge cells | [Creating merged cells](@ref) |
| set column widths and row heights | [Column width and row height](formatting/widthAndHeight.md) |

### Sheet layout and view

| I want to… | Where to look |
|:-----------|:--------------|
| freeze the header row or first column | [Freeze Panes](@ref) |
| insert or delete rows and columns | [Inserting and deleting rows and columns](sheets/insertDelete.md) |
| hide rows or columns | [Hiding](@ref) |
| group rows or columns (outline) | [Grouping](@ref) |

### Charts, images, links and comments

| I want to… | Where to look |
|:-----------|:--------------|
| create a chart | [A chart on a worksheet](@ref) |
| create a waterfall, funnel, treemap, histogram or other new chart type | [Creating chartEx charts](charts/creatingChartEx.md) |
| change a chart's colors, title, legend or axes | [Formatting charts](charts/formattingCharts.md) |
| read the data behind an existing chart | [Getting the cached data](@ref) |
| insert an image or logo | [Images](@ref) |
| insert a plot made in Julia | [Adding a plot image](@ref) |
| add a hyperlink | [Hyperlinks](@ref) |
| add a comment (note) to a cell | [Comments](@ref) |

### Worksheets, protection and printing

| I want to… | Where to look |
|:-----------|:--------------|
| add, rename, copy or delete a sheet | [Managing worksheets](sheets/worksheets.md) |
| reorder sheets | [Sheet order](@ref) |
| hide a sheet or color its tab | [Hiding a sheet](@ref), [Tab colour](@ref) |
| protect a sheet, leaving some cells editable | [Protecting a sheet](@ref), [Unlocking cells](@ref) |
| set orientation, paper size, margins or print area | [Printing](@ref) |

## Where to find things

* **Tutorials** for a first walk-through: [XLSX Tutorial](tutorial/XLSXtutorial.md), [Creating and editing workbooks](tutorial/creatingWorkbooks.md) and [FileIO Tutorial](tutorial/FileIOtutorial.md).
* [Cookbook](cookbook.md): short recipes for common tasks, from writing a `DataFrame` to protecting a sheet.
* **Guides** for each feature, in the sidebar.
* [Examples](examples.md): longer worked examples, from a formatted table to a native chart.
* [Supported features](features.md): what XLSX.jl can and can't do with each Excel feature.
* [API Reference](api/index.md): every function, with an [alphabetical index](@ref "API index").
* [FAQ and troubleshooting](faq.md): error messages and surprises, and what to do about them.
* [Migration Guide](migration.md): changes between versions.

## Requirements

* Julia v1.10 (LTS) or more recent

* Linux, macOS or Windows.

## Installation

From a Julia session, run:

```julia
julia> using Pkg

julia> Pkg.add("XLSX")
```

## Source Code

The source code for this package is hosted at
[https://github.com/juliadata/XLSX.jl](https://github.com/juliadata/XLSX.jl).

## License

The source code for the package **XLSX.jl** is licensed under
the [MIT License](https://raw.githubusercontent.com/juliadata/XLSX.jl/master/LICENSE).

## Getting Help

If you're having any trouble, have any questions about this package
or want to ask for a new feature,
just open a new [issue](https://github.com/juliadata/XLSX.jl/issues).

## Contributing

Contributions are always welcome!

To contribute, fork the project on [GitHub](https://github.com/juliadata/XLSX.jl)
and send a Pull Request.

## References

* [ECMA Open XML White Paper](https://www.ecma-international.org/wp-content/uploads/OpenXML-White-Paper.pdf)

* [ECMA-376](https://ecma-international.org/publications-and-standards/standards/ecma-376/)

* [Excel file limits](https://support.microsoft.com/en-us/excel/excel-specifications-and-limits)

## Alternative Packages

* [FileIO.jl](https://github.com/JuliaIO/FileIO.jl)

* [LibXLSXWriter.jl](https://github.com/jaakkor2/LibXLSXWriter.jl)
