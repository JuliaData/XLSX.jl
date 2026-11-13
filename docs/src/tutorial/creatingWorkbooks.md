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
| write a new file too large to hold in memory | [`XLSX.writetable`](@ref) with `streaming=true`, or [`XLSX.streamxlsx`](@ref) (see [Writing large files](@ref)) | written as it goes |
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

## Writing large files

Every function above keeps the whole workbook in memory until it is saved, which takes
several GiB for a sheet of a million rows. Streaming writes rows straight into the
file instead, a chunk at a time, so memory stays about the same however many rows are
written.

To stream a whole table, give [`XLSX.writetable`](@ref) `streaming=true`. A table that
can be read a row at a time, such as a vector of `NamedTuple`s, a database query or a
`CSV.Rows`, is never held in memory as a whole:

```jldoctest streaming; setup = :(import XLSX; using Dates)
julia> data = (id=1:1000, value=(1:1000) ./ 8, when=[DateTime(2026, 1, 1) + Hour(i) for i in 1:1000]);

julia> io = IOBuffer();

julia> XLSX.writetable(io, data; streaming=true, as_table=true, table_name="Readings",
                       freeze_header=true, column_widths=[6, 10, 18], column_formats=Dict(:value => "0.00"))

julia> XLSX.readxlsx(seekstart(io))[1]["A1:C3"]
3×3 Matrix{Any}:
  "id"   "value"  "when"
 1      0.125     2026-01-01T01:00:00
 2      0.25      2026-01-01T02:00:00
```

To write rows as you produce them, use [`XLSX.streamxlsx`](@ref): add a sheet with
[`XLSX.addsheet!`](@ref) and write its rows in order with [`XLSX.writerow!`](@ref).
Adding the next sheet finishes the one before, and the file is finished when the `do`
block ends:

```jldoctest streaming
julia> io = IOBuffer();

julia> XLSX.streamxlsx(io) do s
           sh = XLSX.addsheet!(s, "Data"; columns=["id", "name", "when"], freeze_header=true)
           for i in 1:3
               XLSX.writerow!(sh, (i, "item $i", Date(2026, 1, i)))
           end
       end

julia> XLSX.readxlsx(seekstart(io))["Data"]["A1:C4"]
4×3 Matrix{Any}:
  "id"  "name"    "when"
 1      "item 1"  2026-01-01
 2      "item 2"  2026-01-02
 3      "item 3"  2026-01-03
```

Writing a million rows of 20 mixed columns this way needs about 160 MiB beyond the
memory the data itself takes, where writing it through the cell cache needs about
3.8 GiB, and it is faster too (see `bench/stream_write.jl` in the repository).

Strings are written in each cell ("inline") by default. With `shared_strings=true`
they go in one shared table, as Excel writes them. That makes a file of many repeated
strings a little smaller, but keeps every different string in memory until the file is
finished. Excel opens both, and stores strings its own way when it saves.

A streamed sheet is written in one pass, so it has limits:

- Rows are written in order, from row 1, each once. A row can't be read or changed
  once written, nor can the file be opened until it is finished.
- A sheet can have column widths, column number formats, a frozen header row and an
  Excel Table, set when it is added. It can't have formulas, cell formatting, merged
  cells, conditional formats, data validation, comments, images, charts or a totals
  row. To add these, open the finished file with `mode="rw"` (which reads it into
  memory) or write it the ordinary way.
- With `writetable`, the table starts in cell A1 (no `anchor_cell`), and has no
  `totals`.
- When the number of rows isn't known in advance (with `streamxlsx`, or a table read
  a row at a time whose length is unknown), the sheet's `<dimension>`, the range Excel
  records as used, is left out. Excel and XLSX.jl work it out from the rows.

## Document properties

A workbook's title, author, tags and other properties, shown in Excel's
*File → Info*, are set with [`XLSX.setDocumentProperties`](@ref) and read with
[`XLSX.getDocumentProperties`](@ref):

```jldoctest docprops; setup = :(import XLSX; using Dates)
julia> xf = XLSX.newxlsx();

julia> XLSX.setDocumentProperties(xf; title="Quarterly sales", creator="Finance team",
                                      keywords="sales; Q3", company="Acme Ltd")

julia> p = XLSX.getDocumentProperties(xf);

julia> (p.title, p.creator, p.company)
("Quarterly sales", "Finance team", "Acme Ltd")
```

The properties are `title`, `subject`, `creator` (Excel's Author), `keywords` (Tags),
`description` (Comments), `lastModifiedBy`, `category`, `contentStatus` (Status),
`company`, `manager`, and the dates `created` and `modified`, which are in UTC. Only
those given are changed, and `""` removes one.

Custom properties, from Excel's *Advanced Properties → Custom*, hold a string, number,
`Bool` or date each:

```jldoctest docprops
julia> XLSX.setCustomProperty(xf, "Project", "Apollo")

julia> XLSX.setCustomProperty(xf, "Approved", true)

julia> XLSX.setCustomProperty(xf, "Due", Date(2026, 6, 30))

julia> XLSX.getCustomProperty(xf, "approved")    # names ignore case, as in Excel
true

julia> collect(keys(XLSX.getCustomProperties(xf)))
3-element Vector{String}:
 "Project"
 "Approved"
 "Due"
```

[`XLSX.deleteCustomProperty`](@ref) deletes one. Properties other programs add, such as
sensitivity labels, are kept as they are. XLSX.jl doesn't change `modified` or
`lastModifiedBy` when it saves; set them if you need to.

## Where next

- Reading data back: [XLSX Tutorial](XLSXtutorial.md)
- Formatting cells: [Cell formats](../formatting/cellFormatting.md)
- Adding sheets: [Managing worksheets](../sheets/worksheets.md)
- Everything XLSX.jl can do: [Supported features](../features.md)
