
# XLSX.jl

## Introduction

**XLSX.jl** is a Julia package to read and write
[Excel](https://products.office.com/excel) spreadsheet files.

Internally, an Excel XLSX file is just a
[Zip](https://en.wikipedia.org/wiki/Zip_(file_format)) file with a set of
[XML](https://en.wikipedia.org/wiki/XML) files inside.
The formats for these XML files are described in
the [Standard ECMA-376](https://ecma-international.org/publications-and-standards/standards/ecma-376/).

This package follows the EMCA-376 to parse and generate XLSX files.

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

* [Excel file limits](https://support.microsoft.com/en-us/office/excel-specifications-and-limits-1672b34d-7043-467e-8e27-269d656771c3)

## Alternative Packages

* [FileIO.jl](https://github.com/JuliaIO/FileIO.jl)

* [LibXLSXWriter.jl](https://github.com/jaakkor2/LibXLSXWriter.jl)
