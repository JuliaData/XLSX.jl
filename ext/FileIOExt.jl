"""
# Introduction

This package provides support for Excel files under the
[FileIO.jl](https://github.com/JuliaIO/FileIO.jl) package.

It provides functionality to read simple tabular data from 
an Excel (.xlsx) file and to save simple tabular data to an 
Excel file.

For more extensive functionality when reading and writing Excel files,
consider using [XLSX.jl](https://juliadata.github.io/XLSX.jl/stable/).
Under the hood, `ExcelFiles.jl` uses the `XLSX.jl` functions `readtable` 
and `writetable`. 

# Usage

## Load an Excel file

To read an Excel file into a `DataFrame`, use the following julia code:

```julia
using ExcelFiles, DataFrames

df = DataFrame(load("data.xlsx", "Sheet1"))
```

The call to `load` returns an object that is a [Tables.jl](https://github.com/JuliaData/Tables.jl) table, so it can be passed to any function that can handle Tables.jl tables. Here are some examples of materializing an Excel file into such data structures:

```julia
using ExcelFiles, DataFrames, PrettyTables

# Load into a DataFrame
julia> DataFrame(load("HTable.xlsx"))
5×10 DataFrame
 Row │ Year    1940   1950        1960     1970     1980   1990        2000     2010     2020    
     │ String  Any    Any         Float64  Float64  Any    Any         Float64  Float64  Float64 
─────┼───────────────────────────────────────────────────────────────────────────────────────────
   1 │ Col A   1      2               3.0     4.0   5      6               7.0     8.0       9.0
   2 │ Col B   10     20             30.0    40.0   50     60             70.0    80.0      90.0
   3 │ Col C   100    200           300.0   400.0   500    600           700.0   800.0     900.0
   4 │ Col D   0.1    0.2             0.3     0.4   0.5    0.6             0.7     0.8       0.9
   5 │ Col E   Hello  2025-12-19      3.0     3.33  Hello  2025-12-19      3.0     3.33      1.0

julia> DataFrame(load("HTable.xlsx"; transpose=true))
9×6 DataFrame
 Row │ Year   Col A  Col B  Col C  Col D    Col E      
     │ Int64  Int64  Int64  Int64  Float64  Any        
─────┼─────────────────────────────────────────────────
   1 │  1940      1     10    100      0.1  Hello
   2 │  1950      2     20    200      0.2  2025-12-19
   3 │  1960      3     30    300      0.3  3
   4 │  1970      4     40    400      0.4  3.33
   5 │  1980      5     50    500      0.5  Hello
   6 │  1990      6     60    600      0.6  2025-12-19
   7 │  2000      7     70    700      0.7  3
   8 │  2010      8     80    800      0.8  3.33
   9 │  2020      9     90    900      0.9  true


# Load into a PrettyTable
julia> PrettyTable(load("HTable.xlsx"))
┌───────┬───────┬────────────┬───────┬───────┬───────┬────────────┬───────┬───────┬───────┐
│  Year │  1940 │       1950 │  1960 │  1970 │  1980 │       1990 │  2000 │  2010 │  2020 │
├───────┼───────┼────────────┼───────┼───────┼───────┼────────────┼───────┼───────┼───────┤
│ Col A │     1 │          2 │   3.0 │   4.0 │     5 │          6 │   7.0 │   8.0 │   9.0 │
│ Col B │    10 │         20 │  30.0 │  40.0 │    50 │         60 │  70.0 │  80.0 │  90.0 │
│ Col C │   100 │        200 │ 300.0 │ 400.0 │   500 │        600 │ 700.0 │ 800.0 │ 900.0 │
│ Col D │   0.1 │        0.2 │   0.3 │   0.4 │   0.5 │        0.6 │   0.7 │   0.8 │   0.9 │
│ Col E │ Hello │ 2025-12-19 │   3.0 │  3.33 │ Hello │ 2025-12-19 │   3.0 │  3.33 │   1.0 │
└───────┴───────┴────────────┴───────┴───────┴───────┴────────────┴───────┴───────┴───────┘

julia> PrettyTable(load("HTable.xlsx"; transpose=true))
┌──────┬───────┬───────┬───────┬───────┬────────────┐
│ Year │ Col A │ Col B │ Col C │ Col D │      Col E │
├──────┼───────┼───────┼───────┼───────┼────────────┤
│ 1940 │     1 │    10 │   100 │   0.1 │      Hello │
│ 1950 │     2 │    20 │   200 │   0.2 │ 2025-12-19 │
│ 1960 │     3 │    30 │   300 │   0.3 │          3 │
│ 1970 │     4 │    40 │   400 │   0.4 │       3.33 │
│ 1980 │     5 │    50 │   500 │   0.5 │      Hello │
│ 1990 │     6 │    60 │   600 │   0.6 │ 2025-12-19 │
│ 2000 │     7 │    70 │   700 │   0.7 │          3 │
│ 2010 │     8 │    80 │   800 │   0.8 │       3.33 │
│ 2020 │     9 │    90 │   900 │   0.9 │       true │
└──────┴───────┴───────┴───────┴───────┴────────────┘

```

The `load` function takes a number of arguments and keywords:

```julia
    FileIO.load(
        source::String,
        [sheet::String,
        [columns::String]];
        [first_row::Int],
        [first_column::String]
        [column_labels::Vector{String}],
        [header::Bool],
        [normalizenames::Bool],
        [transpose::Bool]
    )
```

### Arguments:

* `source`: The name of the file to be loaded.
* `sheet`: Specifies the sheet name to be loaded. If `sheet` is not given, the first Excel sheet in the file will be used.
* `columns`: Determines which columns to read. For example, `"B:D"` will select columns B, C and D. If columns is not given, the algorithm will find the first sequence of consecutive non-empty cells. A valid sheet **must** be specified when specifying columns. If `transpose = true` or is omitted, `columns` should be used to specify rows. For example, specifying `"2:4"` with `transpose = true` will read only from these rows.

### Keywords:

* `first_row`: Indicates the first row of the data table to be read. For example, `first_row=5` will look for a table starting at sheet row 5. If first_row is not given, the algorithm will look for the first non-empty row in the sheet (ignored if `transpose = true`).
* `first_column`: Indicates the first row of the data table to be read. For example, `first_column="B"` will look for a table starting at sheet row 5. If first_row is not given, the algorithm will look for the first non-empty row in the sheet (ignored if `transpose = false` or is omitted).
* `column_labels`: Specifies column names for the header of the table. If `column_labels` are given and `header=true`, the headers given by `column_labels` will be used, and the first row of the table (containing headers) will be ignored.
* `header`: Indicates if the first row (column if `transpose = true`) is a header. If `header=true` and `column_labels` is not specified, the column labels for the table will be read from the first row (column) of the table. If `header=false` and `column_labels` is not specified, the algorithm will generate column labels. The default value is `header=true`.
* `normalizenames`: Set to `true` to normalize column names to valid Julia identifiers. Default=`false`.
* `transpose`: Set to `true` to transpose the table to read data from rows not columns.

### Examples

```julia
julia> PrettyTable(load("HTable.xlsx", "Offset"; first_row=2))

julia> df = DataFrame(load("HTable.xlsx", "Offset", "2:7"; transpose=true, first_column="B"))

julia> df = DataFrame(load("HTable.xlsx"; normalizenames=true, transpose=true, column_labels=["Date", "Name1", "Name2", "Name3", "Name4", "Name5"]))

```
## Save an Excel file

The following code saves any Tables.jl table (such as a `DataFrame`) as an Excel file:
```julia
using ExcelFiles

save("output.xlsx", tbl)
```

The `save` function takes a number of arguments and keywords:

```julia
    FileIO.save(
        source::String;
        [sheetname::String],
        [overwrite::Bool]
    )
```

### Arguments:

* `source`: The name of the file to be created on save.

### Keywords:

* `sheetname`: Specify the sheetname to be used in the created file. By default, the sheetname will be `Sheet1`.
* `overwrite`: Set `overwrite=true` to overwite any existing file of the same name. Default = `false`.

### Examples

```julia
julia> save("myfile.xlsx", df; sheetname="myname", overwrite=true)
```

## Using the pipe syntax

The `load` and `save` functions also support the pipe syntax. For example, to load an Excel file into a `DataFrame`, one can use the following code:

```julia
using ExcelFiles, DataFrame

df = load("data.xlsx", "Sheet1") |> DataFrame
```

To save any Tables.jl compatible table (such as a DataFrame), one can use the following form:

```julia
using ExcelFiles, DataFrame

df = # Aquire a DataFrame somehow

df |> save("output.xlsx")
```
"""

module FileIOloaderExt
# Provides hooks for FileIO.jl to load XLSX files. This is not a dependency of XLSX.jl, so that users who do not want to use FileIO.jl can avoid installing it.

using FileIO

using XLSX

import XLSX: load, save

function load(f::File{format"Excel"}; transpose::Bool=false, kw...)
    filename = FileIO.filename(f)
    if transpose
        return XLSX.readtransposedtable(filename; kw...)
    else
        return XLSX.readtable(filename; kw...)
    end
end

function load(f::File{format"Excel"}, sheet; transpose::Bool=false, kw...)
    filename = FileIO.filename(f)
    if transpose
        return XLSX.readtransposedtable(filename, sheet; kw...)
    else
        return XLSX.readtable(filename, sheet; kw...)
    end
end

function load(f::File{format"Excel"}, sheet, columns; transpose::Bool=false, kw...)
    filename = FileIO.filename(f)
    if transpose
        return XLSX.readtransposedtable(filename, sheet, columns; kw...)
    else
        return XLSX.readtable(filename, sheet, columns; kw...)
    end
end

function save(f::File{format"Excel"}, data; kw...)
    XLSX.writetable(FileIO.filename(f), data; kw...)
end    

end # module