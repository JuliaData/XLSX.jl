# Creating and editing workbooks

The [XLSX Tutorial](XLSXtutorial.md) concentrates on reading. This one shows the ways of
making a workbook you can change, and of saving it. Most of the guides build on these.

## Which function to use

| To… | Use | Save with |
|:----|:----|:----------|
| start a new workbook | [`XLSX.newxlsx`](@ref) | [`XLSX.writexlsx`](@ref) |
| build a new file from an existing one, leaving that one unchanged | [`XLSX.opentemplate`](@ref) | [`XLSX.writexlsx`](@ref) |
| change a file and save it in place, all in one block of code | [`XLSX.openxlsx`](@ref) with `mode="rw"` and `do` | saved automatically at the end of the block |
| change a file over several steps, then decide where to save it | [`XLSX.openxlsx`](@ref) with `mode="rw"`, without `do` | [`XLSX.savexlsx`](@ref) or [`XLSX.writexlsx`](@ref) |
| only read | [`XLSX.readxlsx`](@ref), or [`XLSX.openxlsx`](@ref) with the default `mode="r"` | — |

A file opened read-only can't be changed. Trying to gives an error such as
`Cannot write to a cell: the XLSXFile is not writable. Open it with mode="rw"`.

## A new workbook

[`XLSX.newxlsx`](@ref) returns an empty workbook in memory, with one sheet, `Sheet1`, or
the name you give. Write to it, format it, then save it with [`XLSX.writexlsx`](@ref):

```julia
using XLSX, Dates

xf = XLSX.newxlsx("Report")
s = xf["Report"]

s["A1"] = ["Month", "Sales"]                      # a vector fills across a row
s["A2:A4"] = [Date(2025, 1, 1), Date(2025, 2, 1), Date(2025, 3, 1)]
s["B2:B4"] = [1200.0, 1350.5, 990.25]

XLSX.setFont(s, "A1:B1"; bold=true)
XLSX.setFormat(s, "A2:A4"; format="mmm yyyy")
XLSX.setFormat(s, "B2:B4"; format="#,##0.00")
XLSX.freezePanes(s)                               # keep the header row in view

XLSX.writexlsx("report.xlsx", xf)
```

`writexlsx` returns the path of the file it wrote. It won't replace a file that already
exists unless you pass `overwrite=true`:

```julia
julia> XLSX.writexlsx("report.xlsx", xf)
ERROR: XLSXError: Output file report.xlsx already exists.

julia> XLSX.writexlsx("report.xlsx", xf; overwrite=true)
"C:\\Users\\...\\report.xlsx"
```

More sheets are added with [`XLSX.addsheet!`](@ref): see
[Managing worksheets](../sheets/worksheets.md). Ways of writing values, from single
cells to whole tables, are in [Writing Excel Files](@ref).

## A new file from an existing one

[`XLSX.opentemplate`](@ref) reads a whole file into memory, ready to change. The file on
disk isn't touched: write the result to a new name.

```julia
xf = XLSX.opentemplate("report.xlsx")
xf["Report"]["B5"] = 1500.0
XLSX.writexlsx("report_v2.xlsx", xf)
```

This is the safest way to start from a workbook made in Excel, such as a formatted
template. Excel template files (`.xltx`) work too.

## Editing a file in place

[`XLSX.openxlsx`](@ref) with `mode="rw"` and a `do` block opens a file, runs the block,
and saves the file over the original when the block ends:

```julia
XLSX.openxlsx("report.xlsx", mode="rw") do xf
    s = xf["Report"]
    s["C1"] = "Note"
end
```

If the block throws an error, the file isn't saved and stays as it was.

!!! warning
    The file is overwritten when the block ends. If it holds something XLSX.jl doesn't
    handle (see [Supported features](../features.md)), keep a copy, or use
    `opentemplate` and write to a new name instead.

To work over several steps, or to choose where to save at the end, open the file without
`do`. Then save it yourself, with [`XLSX.savexlsx`](@ref) back to the same file, or with
`writexlsx` to another:

```julia
xf = XLSX.openxlsx("report.xlsx", mode="rw")
xf["Report"]["C2"] = "Checked"
XLSX.savexlsx(xf)                    # or XLSX.writexlsx("other.xlsx", xf)
```

`savexlsx` needs a file to save back to, so it can't save a workbook made by `newxlsx`:
use `writexlsx` for that.

## A new file in one block

`mode="w"` starts a new, empty workbook and writes it to the given path when the `do`
block ends, replacing any file already there:

```julia
XLSX.openxlsx("new_file.xlsx", mode="w") do xf
    xf[1]["A1"] = "Hello"
end
```

## Saving to memory

Both `writexlsx` and `opentemplate` take an `IO` as well as a file name, so a workbook can
be built and sent without touching the disk:

```julia
io = IOBuffer()
XLSX.writexlsx(io, xf)
bytes = take!(io)                    # the .xlsx file's contents
```

## Where next

- Reading data back: [XLSX Tutorial](XLSXtutorial.md)
- Formatting cells: [Cell formats](../formatting/cellFormatting.md)
- Adding sheets: [Managing worksheets](../sheets/worksheets.md)
- Everything XLSX.jl can do: [Supported features](../features.md)
