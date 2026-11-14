# Migration Guide

## Behaviour changes in v0.13 patch releases

The v0.13.x patch releases fixed many edge cases. A few of the fixes change
what existing code sees. Each is listed here with the version that made it.

### Defined names holding a formula read with a leading `=` (v0.13.2)

A defined name whose value is a formula, or any expression that is not a
reference or a constant, is now read with a leading `=`, as Excel's Name
Manager shows it. A plain string is a text constant and is written quoted.

```julia
# Before
dn = only(filter(d -> d.name == "Total", XLSX.getDefinedNames(xf)))
dn.value                                                   # "COUNTA(Sheet1!$A:$A)"
XLSX.addDefinedName(xf, "Total", "COUNTA(Sheet1!$A:$A)")   # was treated as a formula

# After
dn.value                                                   # "=COUNTA(Sheet1!$A:$A)"
XLSX.addDefinedName(xf, "Total", "=COUNTA(Sheet1!$A:$A)")  # the `=` makes it a formula
XLSX.addDefinedName(xf, "Label", "Quarterly")              # a text constant
```

### `openxlsx` no longer saves after an error (v0.13.2)

`openxlsx(f, source; mode="rw")` (and `mode="w"`) no longer writes `source`
when the do block throws. A failure part-way through can no longer overwrite
the original file with a half-edited workbook. Code that relied on a partial
save must catch the error inside the do block.

### `deletesheet!` keeps defined names as `=#REF!` (v0.13.2)

Defined names that refer to a deleted sheet are kept with the value
`"=#REF!"`, as Excel keeps them. Workbook-scoped names used to be deleted.
Delete them yourself with [`XLSX.deleteDefinedName`](@ref) if you don't want them.

### Negative numbers in date-formatted cells (v0.13.2)

A negative number in a date-formatted cell is read as the number. It used to
throw, and that failed the read of the whole sheet.

### Unknown formatting keywords throw an `XLSXError` (v0.13.2)

An unrecognised keyword to `setFont`, `setFill`, `setAlignment`, `setFormat`,
`setBorder` or their `setUniform*` versions is an [`XLSX.XLSXError`](@ref)
naming the valid keywords. It used to be a `MethodError`. Code that caught
`MethodError` should catch `XLSXError`.

### Literal line ends in text read as LF (v0.13.3)

Text written in a file with a literal CR LF or CR (rather than `&#13;`) now
reads as LF, as XML 1.0 requires. Excel writes a CR as `_x000D_`, which is
unaffected, so files written by Excel read as before.

### `copysheet!` points self-references at the copy (v0.13.3)

`copysheet!` points the copy's references to the original sheet at the copy,
as Excel does. This applies in cell formulas, conditional formats, data
validations and sheet-scoped defined names holding a formula. For example,
`=Sheet1!A1` on `Sheet1` becomes `=Copy!A1` on the copy. References to other
sheets are unchanged.

### Built-in number formats 6, 8, 38, 40 and 47 use Excel's codes (v0.13.4)

[`XLSX.getFormat`](@ref) gives these built-in formats the codes Excel gives
them: `$#,##0_);[Red]($#,##0)` and the like for 6, 8, 38 and 40, and
`mm:ss.0` for 47. It used to give `Red` without brackets and `mmss.0`. Code
that compared against the old codes should compare against Excel's.
[`XLSX.setFormat`](@ref) given one of Excel's codes now uses the built-in
format, where it used to add a custom one.

### `setFormat` checks format codes as Excel does (v0.13.4)

[`XLSX.setFormat`](@ref) accepts the number format codes Excel accepts and
refuses the ones Excel refuses. It now accepts codes it used to refuse, such
as `0.0E-0`, `# ?/4` and `[Color12]`. It refuses codes it used to accept, such
as an unknown colour (`[Purple]`) or a date code among digit placeholders
(`0 h`), with an [`XLSX.XLSXError`](@ref).

## Migrating from v0.12 to v0.13

Chart code moved to the `XLSX.Charts` sub-module (`using XLSX, XLSX.Charts`).
The chart names released in v0.12 are still reachable as `XLSX.getCharts` and
so on. These changes may need changes to existing code:

- [`getCharts`](@ref XLSX.Charts.getCharts) also returns `chartEx` charts, as a
  `Vector{AbstractChart}`. Code that assumed every element is a `Chart` should
  check with `isa`.
- [`XLSX.getDefinedNames`](@ref) and [`XLSX.getAllDefinedNames`](@ref) exclude
  the names Excel keeps for itself (hidden names and the `_xlnm.` built-ins).
  Pass `include_system=true` for the old behaviour. [`XLSX.deleteDefinedName`](@ref)
  and [`XLSX.deleteAllDefinedNames`](@ref) refuse to delete those names unless
  `force=true`, and [`XLSX.addDefinedName`](@ref) refuses to create them.
- [`XLSX.getConditionalFormats`](@ref) returns
  `Vector{Pair{Union{CellRange,NonContiguousRange},...}}` rather than
  `Vector{Pair{CellRange,...}}`.

Deprecated, and still working:

| Deprecated | Use instead |
| --- | --- |
| `chart.title`, `chart.charttypes`, `chart.series` | `getChartTitle`, `getChartTypes`, `getChartSeries` |
| `chartType`, `chartSchema` | `getChartType`, `getChartSchema` |

## Migrating from v0.11 to v0.12

There are no changes to the public API. v0.12 requires Julia v1.10 (LTS) or
later and adds a [FileIO.jl](https://github.com/JuliaIO/FileIO.jl) extension
(see the FileIO tutorial).

## Migrating from v0.10 to v0.11

There is only one breaking change in v0.11.0.

- `infer_eltypes` now defaults to `true` (e.g. in `gettable` and `readtable`). This is the more common use case but,
  if it is not *your* use case you will need explicitly to set `infer_eltypes = false` in the relevant functions.

All other changes either introduce new functionality (documented elsewhere) or relate to internals only.


## Migrating Legacy Code to v0.8+

!!! note

    The sections below were written as a guide to describe migrating from a pre v0.8 version 
    of XLSX.jl to v0.8. They are largely only historic now.

Version `v0.8` introduced a breaking change on methods [`XLSX.gettable`](@ref) and [`XLSX.readtable`](@ref).

These methods used to return a tuple `data, column_labels`.
On XLSX `v0.8` these methods return a `XLSX.DataTable` struct that implements [`Tables.jl`](https://github.com/JuliaData/Tables.jl) interface.

### Basic code replacement

Before

```julia
data, col_names = XLSX.readtable(joinpath(data_directory, "general.xlsx"), "table4")
```

After

```julia
dtable = XLSX.readtable(joinpath(data_directory, "general.xlsx"), "table4")
data, col_names = dtable.data, dtable.column_labels
```

### Reading DataFrames

Since `XLSX.DataTable` implements `Tables.jl` interface,
the result of `XLSX.gettable` or `XLSX.readtable` can be
passed to a `DataFrame` constructor.

Before

```julia
df = DataFrame(XLSX.readtable("myfile.xlsx", "mysheet")...)
```

After

```julia
df = DataFrame(XLSX.readtable("myfile.xlsx", "mysheet"))
```
