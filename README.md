
# XLSX.jl

[![License][license-img]](LICENSE)
[![CI][ci-img]][ci-url]
[![codecov][codecov-img]][codecov-url]
[![dev][docs-dev-img]][docs-dev-url]
[![stable][docs-stable-img]][docs-stable-url]
[![DOI][doi-img]][doi-url]

[license-img]: http://img.shields.io/badge/license-MIT-brightgreen.svg?style=flat-square
[ci-img]: https://github.com/juliadata/XLSX.jl/workflows/CI/badge.svg
[ci-url]: https://github.com/juliadata/XLSX.jl/actions?query=workflow%3ACI
[codecov-img]: https://img.shields.io/codecov/c/github/juliadata/XLSX.jl/master.svg?label=codecov&style=flat-square
[codecov-url]: http://codecov.io/github/juliadata/XLSX.jl?branch=master
[docs-dev-img]: https://img.shields.io/badge/docs-dev-blue.svg?style=flat-square
[docs-dev-url]: https://juliadata.github.io/XLSX.jl/dev
[docs-stable-img]: https://img.shields.io/badge/docs-stable-blue.svg?style=flat-square
[docs-stable-url]: https://juliadata.github.io/XLSX.jl/stable
[doi-img]: https://zenodo.org/badge/DOI/10.5281/zenodo.19894703.svg
[doi-url]: https://doi.org/10.5281/zenodo.19894703

Excel file reader/writer for the Julia language.

## Introduction

**XLSX.jl** is a Julia package to read and write
[Excel](https://www.microsoft.com/microsoft-365/excel) spreadsheet files.

Internally, an Excel XLSX file is just a
[Zip](https://en.wikipedia.org/wiki/Zip_(file_format)) file with a set of
[XML](https://en.wikipedia.org/wiki/XML) files inside.
The formats for these XML files are described in
the [Standard ECMA-376](https://ecma-international.org/publications-and-standards/standards/ecma-376/).

This package follows ECMA-376 to parse and generate XLSX files.

## Requirements

* Julia v1.10 (LTS) or more recent

* Linux, macOS or Windows.

## Installation

From a Julia session, run:

```julia
julia> using Pkg

julia> Pkg.add("XLSX")
```

## Features

* Read and write cell values, ranges and whole tables through the
  [Tables.jl](https://github.com/JuliaData/Tables.jl) interface, so data goes straight to and from a `DataFrame`
* Stream the rows of large files with `eachrow` and `eachtablerow`, without loading the whole sheet
* Cell formatting (fonts, fills, borders, alignment, number formats) and conditional formats
* Formulas and defined names
* Data validation, including drop-down lists
* AutoFilter buttons on plain ranges
* Sheet visibility (hidden and very hidden), tab colours, sheet protection and page setup for printing
* Native Excel Tables
* Merged cells, column widths, row heights, hidden and grouped rows and columns, and freeze/split panes
* Hyperlinks, images and cell comments
* Native Excel charts: read, format and create (experimental)
* [FileIO.jl](https://github.com/JuliaIO/FileIO.jl) integration

See the [documentation](https://juliadata.github.io/XLSX.jl/stable) for guides and the full API.

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

* [Excel file limits](https://support.microsoft.com/en-us/office/excel-specifications-and-limits-1672b34d-7043-467e-8e27-269d656771c3)

## Alternative Packages

* [FileIO.jl](https://github.com/JuliaIO/FileIO.jl)

* [LibXLSXWriter.jl](https://github.com/jaakkor2/LibXLSXWriter.jl)
