# FileIO Tutorial

## Introduction

A package extension to XLSX.jl provides support for Excel files 
under the [FileIO.jl](https://github.com/JuliaIO/FileIO.jl) package.

[FileIO.jl](https://github.com/JuliaIO/FileIO.jl) aims to provide a common 
framework for detecting file formats and dispatching to appropriate readers/writers.

Through [FileIO.jl](https://github.com/JuliaIO/FileIO.jl), you can read 
simple tabular data from an Excel (.xlsx) file and save tabular data 
to an Excel file using simple `load` and `save` functions without needing 
to know anything about XLSX.jl itself.

XLSX.jl provides much more extensive functionality if you need it.
Check out the rest of the documentation for full details.

## Setup

First, make sure you have the **FileIO.jl** and **XLSX.jl** packages installed.

```julia
julia> using Pkg

julia> Pkg.add(["FileIO", "XLSX"])
```

!!! note

    [FileIO.jl](https://github.com/JuliaIO/FileIO.jl) support requires a version of [FileIO.jl](https://github.com/JuliaIO/FileIO.jl) greater than v1.19.0.

## Usage

### Load an Excel file

To read tabular data from an Excel file into a `DataFrame`, use the following julia code:

```julia
using FileIO, DataFrames

df = DataFrame(load("data.xlsx", "Sheet1"))
```

The call to `load` returns an object that is a [Tables.jl](https://github.com/JuliaData/Tables.jl) table, 
so it can be passed to any function that can handle Tables.jl tables. Here are some examples of 
materializing an Excel file into such data structures:

```julia
using FileIO, DataFrames, PrettyTables

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

For more information, see [`XLSX.load`](@ref)

### Save an Excel file

The following code saves any Tables.jl table (such as a `DataFrame`) as an Excel file:

```julia
using FileIO 

save("output.xlsx", myTable)
```

For more information, see [`XLSX.save`](@ref)

### Using the pipe syntax

The `load` and `save` functions also support the pipe syntax. For example, to load an 
Excel file into a `DataFrame`, one can use the following code:

```julia
using FileIO, DataFrame

df = load("data.xlsx", "Sheet1") |> DataFrame
```

To save any Tables.jl compatible table (such as a DataFrame), one can use the following form:

```julia
using FileIO, DataFrame

df = # Aquire a DataFrame somehow

df |> save("output.xlsx")
```
