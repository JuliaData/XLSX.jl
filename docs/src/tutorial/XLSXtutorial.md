
# XLSX Tutorial

## Setup

First, make sure you have the **XLSX.jl** package installed.

```julia
julia> using Pkg

julia> Pkg.add("XLSX")
```

## Getting Started

The basic usage is to read an Excel file and read values.

```julia
julia> import XLSX

julia> xf = XLSX.readxlsx("myfile.xlsx")
XLSXFile("myfile.xlsx") containing 3 Worksheets
            sheetname size          range
-------------------------------------------------
              mysheet 4x2           A1:B4
           othersheet 1x1           A1:A1
                named 1x1           B4:B4

julia> XLSX.sheetnames(xf)
3-element Vector{String}:
 "mysheet"
 "othersheet"
 "named"

julia> sh = xf["mysheet"] # get a reference to a Worksheet
4×2 XLSX.Worksheet: ["mysheet"](A1:B4)

julia> sh[2, 2] # access element "B2" (2nd row, 2nd column)
"first"

julia> sh["B2"] # you can also use the cell name
"first"

julia> sh["A2:B4"] # or a cell range
3×2 Matrix{Any}:
 1  "first"
 2  "second"
 3  "third"

julia> XLSX.readdata("myfile.xlsx", "mysheet", "A2:B4") # shorthand for all above
3×2 Matrix{Any}:
 1  "first"
 2  "second"
 3  "third"

julia> sh[:] # all data inside worksheet's dimension
4×2 Matrix{Any}:
  "HeaderA"  "HeaderB"
 1           "first"
 2           "second"
 3           "third"

julia> xf["mysheet!A2:B4"] # you can also query values using a sheet reference
3×2 Matrix{Any}:
 1  "first"
 2  "second"
 3  "third"

julia> xf["mysheet!A:B"] # Column ranges are also supported
4×2 Matrix{Any}:
  "HeaderA"  "HeaderB"
 1           "first"
 2           "second"
 3           "third"

julia> xf["mysheet!2:4"] # And so are row ranges
3×2 Matrix{Any}:
 1  "first"
 2  "second"
 3  "third"
 
 julia> xf["NAMED_CELL"] # you can even read named ranges
"B4 is a named cell from sheet \"named\""

```

To inspect the internal representation of each cell, use the `getcell` or `getcellrange` methods.

The example above used `xf = XLSX.readxlsx(filename)` to open a file, so all file contents are fetched at once from disk.

You can also use `XLSX.openxlsx` to read file contents as needed (see [Reading Large Excel Files and Caching](@ref)).

If you need to read an Excel file that has been protected with a password, consider using [XLSXDecrypt.jl](https://github.com/TimG1964/XLSXDecrypt.jl) 
to decrypt the file so it can be read by XLSX.jl.

## Data Types

This package uses the following concrete types when handling XLSX files.

```@docs
XLSX.CellConcreteType
```

- Abstract types of these concrete types are converted to the appropriate concrete type when writing.

- `Nothing` values are converted to `Missing` when writing.

## Read Tabular Data

The [`XLSX.gettable`](@ref) method returns tabular data from a spreadsheet as a struct `XLSX.DataTable`
that implements [`Tables.jl`](https://github.com/JuliaData/Tables.jl) interface.
You can use it to create a `DataFrame` from [DataFrames.jl](https://github.com/JuliaData/DataFrames.jl).
Check the docstring for `gettable` method for more advanced options.

There are also two helper methods [`XLSX.readtable`](@ref) and [`XLSX.readto`](@ref) to read from file 
directly, as shown in the following examples.

```julia
julia> using DataFrames, XLSX

julia> df = DataFrame(XLSX.readtable("myfile.xlsx", "mysheet")) # Returns a `Tables.jl` table that `DataFrame` can accept
3×2 DataFrame
 Row │ HeaderA  HeaderB 
     │ Int64    String  
─────┼──────────────────
   1 │       1  first
   2 │       2  second
   3 │       3  third

julia> df = XLSX.readto("myfile.xlsx", "mysheet", DataFrame) # Returns a `DataFrame` directly.
3×2 DataFrame
 Row │ HeaderA  HeaderB 
     │ Int64    String  
─────┼──────────────────
   1 │       1  first
   2 │       2  second
   3 │       3  third


```

## Reading Cells as a Julia Matrix

Use [`XLSX.readdata`](@ref) or [`XLSX.getdata`](@ref) to read content as a Julia matrix.

```julia
julia> import XLSX

julia> m = XLSX.readdata("myfile.xlsx", "mysheet!A1:B3")
3×2 Array{Any,2}:
  "HeaderA"  "HeaderB"
 1           "first"
 2           "second"
```

Indexing in a `Worksheet` will dispatch to [`XLSX.getdata`](@ref) method.

```julia
julia> xf = XLSX.readxlsx("myfile.xlsx")
XLSXFile("myfile.xlsx") containing 3 Worksheets
            sheetname size          range
-------------------------------------------------
              mysheet 4x2           A1:B4
           othersheet 1x1           A1:A1
                named 1x1           B4:B4

julia> xf["mysheet!A1:B3"]
3×2 Array{Any,2}:
  "HeaderA"  "HeaderB"
 1           "first"
 2           "second"

julia> sheet = xf["mysheet"]
4×2 XLSX.Worksheet: ["mysheet"](A1:B4)

julia> sheet["A1:B3"]
3×2 Array{Any,2}:
  "HeaderA"  "HeaderB"
 1           "first"
 2           "second"
```

But indexing in a single cell will return a single value instead of a matrix.

```julia
julia> sheet["A1"]
"HeaderA"
```

If you don't know the desired range in advance, you can take advantage of the
[`XLSX.readtable`](@ref) and [`XLSX.gettable`](@ref) methods.

```julia
julia> dtable = XLSX.readtable("myfile.xlsx", "mysheet")
XLSX.DataTable(Any[Any[1, 2, 3], Any["first", "second", "third"]], [:HeaderA, :HeaderB], Dict(:HeaderB => 2, :HeaderA => 1))

julia> m = hcat(dtable.data...)
3×2 Matrix{Any}:
 1  "first"
 2  "second"
 3  "third"
```

## Reading Large Excel Files and Caching

The method `XLSX.openxlsx` has a `enable_cache` option to control worksheet cells caching.

Cache is enabled by default, so if you read a worksheet cell twice it will use the cached value instead of reading from disk 
the second time.

If `enable_cache=false`, worksheet cells will always be read from disk. In addition, if `enable_cache=false` 
and `openxlsx` is used with do-syntax, the xlsx file itself will be opened as a `FileArray` so that the 
zip archives themselves are not read into memory. This is useful when you want to read a spreadsheet that 
doesn't fit into memory.

The following example shows how you would read worksheet cells, one row at a time,
where `myfile.xlsx` is a spreadsheet that doesn't fit into memory.

```julia
julia> XLSX.openxlsx("myfile.xlsx", enable_cache=false) do f
           sheet = f["mysheet"]
           for r in XLSX.eachrow(sheet)
              # r is a `SheetRow`. Values are read using column references
              rn = XLSX.row_number(r) # `SheetRow` row number
              v1 = r[1]    # will read value at column 1
              v2 = r["B"]  # will read value at column 2

              println("v1=$v1, v2=$v2")
           end
      end
v1=HeaderA, v2=HeaderB
v1=1, v2=first
v1=2, v2=second
v1=3, v2=third
```

You could also stream tabular data using `XLSX.eachtablerow(sheet)`, which is the underlying iterator in the `gettable` method.
Check docstrings for `XLSX.eachtablerow` for more advanced options.

```julia
julia> XLSX.openxlsx("myfile.xlsx", enable_cache=false) do f
           sheet = f["mysheet"]
           for r in XLSX.eachtablerow(sheet)
               # r is a `TableRow`. Values are read using column labels or numbers
               rn = XLSX.row_number(r) # `TableRow` row number
               v1 = r[1] # will read value at table column 1
               v2 = r[:HeaderB] # will read value at column labeled `:HeaderB`

               println("v1=$v1, v2=$v2")
            end
       end
v1=1, v2=first
v1=2, v2=second
v1=3, v2=third
```

## Writing Excel Files

### Create New Files

Opening a file in `write` mode with `XLSX.openxlsx` will open a new (blank) Excel file for editing.

```julia
XLSX.openxlsx("my_new_file.xlsx", mode="w") do xf
    sheet = xf[1]
    XLSX.renamesheet!(sheet, "new_sheet")
    sheet["A1"] = "this"
    sheet["A2"] = "is a"
    sheet["A3"] = "new file"
    sheet["A4"] = 100

    # will add a row from "A5" to "E5"
    sheet["A5"] = collect(1:5) # equivalent to `sheet["A5", dim=2] = collect(1:4)`

    # will add a column from "B1" to "B4"
    sheet["B1", dim=1] = collect(1:4)

    # will add a matrix from "A7" to "C9"
    sheet["A7:C9"] = [ 1 2 3 ; 4 5 6 ; 7 8 9 ]
end
```

### Edit Existing Files

Opening a file in `read-write` mode with `XLSX.openxlsx` will open an existing Excel file for editing.
This will preserve existing data and formatting in the original file.

```julia
XLSX.openxlsx("my_new_file.xlsx", mode="rw") do xf
    sheet = xf[1]
    sheet["B1"] = "new data"
end
```

!!! warning

    Using do-block syntax in "rw" mode will overwrite the file you read in with the modified data when the do block ends. Care is needed to ensure data are not inadvertantly overwritten, especially if the xlsx file contains any elements that `XLSX.jl` cannot process (such as charts, pivot tables, etc), but that would otherwise be preserved if not overwritten. You may avoid this risk by choosing to open files in "rw" mode without using do-block syntax, in which case it becomes necessary explicitly to write the `XLSXFile` out again, providing the option to write to another file name.

### Export Tabular Data from a Worksheet

Given a sheet reference, use the `XLSX.writetable!` method. Anchor cell defaults to cell `"A1"`.

```julia
using XLSX, Test

filename = "myfile.xlsx"

columns = Vector()
push!(columns, [1, 2, 3])
push!(columns, ["a", "b", "c"])

labels = [ "column_1", "column_2"]

XLSX.openxlsx(filename, mode="w") do xf
    sheet = xf[1]
    XLSX.writetable!(sheet, columns, labels, anchor_cell=XLSX.CellRef("B2"))
end

# read data back
XLSX.openxlsx(filename) do xf
    sheet = xf[1]
    @test sheet["B2"] == "column_1"
    @test sheet["C2"] == "column_2"
    @test sheet["B3"] == 1
    @test sheet["B4"] == 2
    @test sheet["B5"] == 3
    @test sheet["C3"] == "a"
    @test sheet["C4"] == "b"
    @test sheet["C5"] == "c"
end
```

You can also use `XLSX.writetable` to write directly to a new file (see next section).

### Export Tabular Data from any `Tables.jl` compatible source

To export tabular data to Excel, use `XLSX.writetable` method, which accepts either columns and column names,
or any `Tables.jl` table.

```julia
julia> using Dates

julia> import DataFrames, XLSX

julia> df = DataFrames.DataFrame(integers=[1, 2, 3, 4], strings=["Hey", "You", "Out", "There"], floats=[10.2, 20.3, 30.4, 40.5], dates=[Date(2018,2,20), Date(2018,2,21), Date(2018,2,22), Date(2018,2,23)], times=[Dates.Time(19,10), Dates.Time(19,20), Dates.Time(19,30), Dates.Time(19,40)], datetimes=[Dates.DateTime(2018,5,20,19,10), Dates.DateTime(2018,5,20,19,20), Dates.DateTime(2018,5,20,19,30), Dates.DateTime(2018,5,20,19,40)])
4×6 DataFrames.DataFrame
│ Row │ integers │ strings │ floats │ dates      │ times    │ datetimes           │
├─────┼──────────┼─────────┼────────┼────────────┼──────────┼─────────────────────┤
│ 1   │ 1        │ Hey     │ 10.2   │ 2018-02-20 │ 19:10:00 │ 2018-05-20T19:10:00 │
│ 2   │ 2        │ You     │ 20.3   │ 2018-02-21 │ 19:20:00 │ 2018-05-20T19:20:00 │
│ 3   │ 3        │ Out     │ 30.4   │ 2018-02-22 │ 19:30:00 │ 2018-05-20T19:30:00 │
│ 4   │ 4        │ There   │ 40.5   │ 2018-02-23 │ 19:40:00 │ 2018-05-20T19:40:00 │

julia> XLSX.writetable("df.xlsx", df)
```

You can also export multiple tables to Excel, each table in a separate worksheet, by either passing a tuple (columns, names)
to a keyword argument for each sheet name, or a list `"sheet name" => table` pairs for any Tables.jl compatible source.

```julia
julia> import DataFrames, XLSX

julia> df1 = DataFrames.DataFrame(COL1=[10,20,30], COL2=["Fist", "Sec", "Third"])
3×2 DataFrames.DataFrame
│ Row │ COL1 │ COL2  │
├─────┼──────┼───────┤
│ 1   │ 10   │ Fist  │
│ 2   │ 20   │ Sec   │
│ 3   │ 30   │ Third │

julia> df2 = DataFrames.DataFrame(AA=["aa", "bb"], AB=[10.1, 10.2])
2×2 DataFrames.DataFrame
│ Row │ AA │ AB   │
├─────┼────┼──────┤
│ 1   │ aa │ 10.1 │
│ 2   │ bb │ 10.2 │

julia> XLSX.writetable("report.xlsx", "REPORT_A" => df1, "REPORT_B" => df2)
```

This last example shows how to do the same thing, but when you don't know how many tables you'll be exporting in advance.

```julia
df1 = DataFrame(A=[1,2], B=[3,4])
df2 = DataFrame(C=["Hey", "you"], D=["out", "there"])

sheet_names = [ "1st", "2nd" ]
dataframes = [ df1, df2 ]

@assert length(sheet_names) == length(dataframes)

XLSX.openxlsx("report.xlsx", mode="w") do xf
    for i in eachindex(sheet_names)
        sheet_name = sheet_names[i]
        df = dataframes[i]
        
        if i == firstindex(sheet_names)
            sheet = xf[1]
            XLSX.renamesheet!(sheet, sheet_name)
            XLSX.writetable!(sheet, df)
        else
            sheet = XLSX.addsheet!(xf, sheet_name)
            XLSX.writetable!(sheet, df)        
        end
    end
end
```

## Tables.jl interface

Both types `XLSX.DataTable` and `XLSX.TableRowIterator` conforms to [Tables.jl](https://github.com/JuliaData/Tables.jl) interface.
An instance of `XLSX.TableRowIterator` is created by the function `XLSX.eachtablerow`.

Also, both `XLSX.writetable` and `XLSX.XLSXFile` accept an argument that conforms to the `Tables.jl` interface.

As an example, the type `DataFrame` from [DataFrames](https://github.com/JuliaData/DataFrames.jl) package
supports the `Tables.jl` interface. The following code writes and reads back a `DataFrame` to an Excel file.

```julia
julia> using Dates

julia> import DataFrames, XLSX

julia> df = DataFrames.DataFrame(integers=[1, 2, 3, 4], strings=["Hey", "You", "Out", "There"], floats=[10.2, 20.3, 30.4, 40.5], dates=[Date(2018,2,20), Date(2018,2,21), Date(2018,2,22), Date(2018,2,23)], times=[Dates.Time(19,10), Dates.Time(19,20), Dates.Time(19,30), Dates.Time(19,40)], datetimes=[Dates.DateTime(2018,5,20,19,10), Dates.DateTime(2018,5,20,19,20), Dates.DateTime(2018,5,20,19,30), Dates.DateTime(2018,5,20,19,40)])
4×6 DataFrames.DataFrame
│ Row │ integers │ strings │ floats  │ dates      │ times    │ datetimes           │
│     │ Int64    │ String  │ Float64 │ Date       │ Time     │ DateTime            │
├─────┼──────────┼─────────┼─────────┼────────────┼──────────┼─────────────────────┤
│ 1   │ 1        │ Hey     │ 10.2    │ 2018-02-20 │ 19:10:00 │ 2018-05-20T19:10:00 │
│ 2   │ 2        │ You     │ 20.3    │ 2018-02-21 │ 19:20:00 │ 2018-05-20T19:20:00 │
│ 3   │ 3        │ Out     │ 30.4    │ 2018-02-22 │ 19:30:00 │ 2018-05-20T19:30:00 │
│ 4   │ 4        │ There   │ 40.5    │ 2018-02-23 │ 19:40:00 │ 2018-05-20T19:40:00 │

julia> XLSX.writetable("output_table.xlsx", df, overwrite=true, sheetname="report", anchor_cell="B2")

julia> f = XLSX.readxlsx("output_table.xlsx")
XLSXFile("output_table.xlsx") containing 1 Worksheet
            sheetname size          range
-------------------------------------------------
               report 6x7           A1:G6


julia> s = f["report"]
6×7 XLSX.Worksheet: ["report"](A1:G6)

julia> df2 = XLSX.eachtablerow(s) |> DataFrames.DataFrame
4×6 DataFrames.DataFrame
│ Row │ integers │ strings │ floats  │ dates      │ times    │ datetimes           │
│     │ Int64    │ String  │ Float64 │ Date       │ Time     │ DateTime            │
├─────┼──────────┼─────────┼─────────┼────────────┼──────────┼─────────────────────┤
│ 1   │ 1        │ Hey     │ 10.2    │ 2018-02-20 │ 19:10:00 │ 2018-05-20T19:10:00 │
│ 2   │ 2        │ You     │ 20.3    │ 2018-02-21 │ 19:20:00 │ 2018-05-20T19:20:00 │
│ 3   │ 3        │ Out     │ 30.4    │ 2018-02-22 │ 19:30:00 │ 2018-05-20T19:30:00 │
│ 4   │ 4        │ There   │ 40.5    │ 2018-02-23 │ 19:40:00 │ 2018-05-20T19:40:00 │
```
Some other packages provide functions (such as `CSV.read`) that can take a sink function 
to accept a `Tables.jl` table. `XLSXFile` can serve as such a sink function, allowing 
`XLSXFile` objects to be created directly.

For example, using the example of `CSV.read`:
```
julia> using CSV, XLSX

julia> xf = CSV.read("iris.csv", XLSXFile)
XLSXFile("blank.xlsx") containing 1 Worksheet
            sheetname size          range
-------------------------------------------------
               Sheet1 151x5         A1:E151
```