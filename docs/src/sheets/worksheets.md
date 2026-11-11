```@meta
DocTestSetup = :(import XLSX)
```

# Managing worksheets

This page covers the sheets in a workbook: finding them, and adding, renaming, copying,
moving and deleting them. Functions that change the workbook need a file opened with
`mode="rw"` (or started from [`XLSX.newxlsx`](@ref)).

## Finding sheets

```jldoctest sheets
julia> xf = XLSX.newxlsx("Summary"); XLSX.addsheet!(xf, "Data");

julia> XLSX.sheetnames(xf)
2-element Vector{String}:
 "Summary"
 "Data"

julia> XLSX.sheetcount(xf)
2

julia> XLSX.hassheet(xf, "Data")
true
```

A sheet is fetched by name or by its position in the tab order:

```jldoctest sheets
julia> xf["Data"]
1×1 Worksheet: ["Data"](A1:A1)

julia> xf[2]
1×1 Worksheet: ["Data"](A1:A1)
```

Sheet names are matched without regard to case, as Excel matches them, so
`xf["data"]` and `XLSX.hassheet(xf, "DATA")` both find the sheet `Data`. A workbook
can't hold two sheets whose names differ only in case.

## Adding sheets

[`XLSX.addsheet!`](@ref) adds an empty worksheet at the end and returns it. Without a
name, it is given one Excel would choose (`Sheet1`, `Sheet2`, …) that isn't already used:

```julia
data  = XLSX.addsheet!(xf, "Prices")
other = XLSX.addsheet!(xf)            # e.g. "Sheet1"
```

To write a whole table to a new sheet in one step, see
[Export Tabular Data from any `Tables.jl` compatible source](@ref).

## Renaming sheets

```julia
XLSX.renamesheet!(xf["Prices"], "Price list")
```

Excel compares sheet names without regard to case, so a name already used in another
case is refused. Renaming updates what refers to the sheet by name:

- formulas on every sheet, and [defined names](../formulae/namedRanges.md);
- the ranges plotted by charts (see [Sheets with charts](../charts/chartSheets.md));
- data validation formulas, and the hidden name Excel uses for an AutoFilter;
- the print area and print titles.

A hyperlink into the workbook stores its target as text, so it isn't updated: see
[Hyperlinks](@ref).

## Copying sheets

[`XLSX.copysheet!`](@ref) copies a sheet to the end of the workbook and returns the copy.
Without a name, the copy is called `"Data (copy)"`, with a number added if that is taken:

```julia
XLSX.copysheet!(xf["Data"])               # "Data (copy)"
XLSX.copysheet!(xf["Data"], "Data 2025")
```

The copy has the original's cells, formats, conditional formats, data validations,
AutoFilter and the defined names scoped to it. References in it to the original sheet
are pointed at the copy, as Excel does.

!!! warning "Experimental"
    `copysheet!` may not copy every feature of complex sheets. Check the result, and
    please report any problems.

## Sheet order

[`XLSX.movesheet!`](@ref) changes the order of the tabs. Give the sheet to move, and
either `before` or `after` to say where it goes. Either sheet can be named by its name,
its tab position or the `Worksheet` itself:

```jldoctest order
julia> xf = XLSX.newxlsx("A"); XLSX.addsheet!(xf, "B"); XLSX.addsheet!(xf, "C");

julia> XLSX.movesheet!(xf, "C"; before="A");

julia> XLSX.sheetnames(xf)
3-element Vector{String}:
 "C"
 "A"
 "B"

julia> XLSX.movesheet!(xf[1]; after=XLSX.sheetcount(xf));   # move to the end
```

Moving a sheet doesn't rewrite anything else. Formulas, charts and defined names refer
to sheets by name, and names scoped to a sheet move with it. The one thing that changes
is a 3D reference such as `=SUM(Jan:Mar!A1)`. As in Excel, it covers whatever sheets lie
between `Jan` and `Mar` in tab order.

## Deleting sheets

[`XLSX.deletesheet!`](@ref) deletes a sheet, given as a `Worksheet`, a name or a tab
position:

```julia
XLSX.deletesheet!(xf, "Data (copy)")
XLSX.deletesheet!(xf["Notes"])
```

As in Excel:

- formulas on other sheets that referred to the deleted sheet get `#REF!` in place of
  the reference;
- defined names that referred to it are kept, with the value `=#REF!`, and names scoped
  to it are deleted;
- data validation formulas that referred to it get `#REF!` too.

A workbook must keep at least one sheet, so the only sheet can't be deleted.

## Hiding a sheet

[`XLSX.setSheetVisibility`](@ref) shows or hides a sheet's tab:

```julia
XLSX.setSheetVisibility(xf["Lookups"], :hidden)       # listed in Excel's Unhide menu
XLSX.setSheetVisibility(xf["Workings"], :veryHidden)  # not listed: only VBA can show it
XLSX.setSheetVisibility(xf["Lookups"], :visible)
```

```jldoctest visible; setup = :(import XLSX; xf = XLSX.newxlsx("Summary"); XLSX.addsheet!(xf, "Workings"); XLSX.setSheetVisibility(xf["Workings"], :veryHidden))
julia> XLSX.getSheetVisibility(xf["Workings"])
:veryHidden
```

A hidden sheet still works as any other: formulas can refer to its cells, and
XLSX.jl reads and writes it as before.

A workbook needs at least one visible sheet, so hiding the last one is an error.
If the hidden sheet is the one the workbook opens on, the first visible sheet
becomes the one it opens on instead.

## Tab colour

```julia
XLSX.setTabColor(xf["Summary"], "FF0000")   # or "red", :accent2, ...
```

The colour can be given in any of the forms described in
[Specifying colors](@ref). [`XLSX.getTabColor`](@ref) returns it as an
`"FFRRGGBB"` hex string, and [`XLSX.removeTabColor`](@ref) removes it:

```jldoctest tab; setup = :(import XLSX; xf = XLSX.newxlsx("Summary"); XLSX.setTabColor(xf["Summary"], "FF0000"))
julia> XLSX.getTabColor(xf["Summary"])
"FFFF0000"

julia> XLSX.removeTabColor(xf["Summary"])
true
```

Visibility and tab colour work on chartsheets as well as worksheets.

## Chartsheets

A chartsheet is a sheet holding a single chart. [`addChart`](@ref XLSX.Charts.addChart)
creates one when given the file rather than a worksheet: see
[A chart on its own sheet](@ref). Chartsheets are renamed, moved, hidden and deleted
with the same functions as worksheets, but can't be copied.
[Sheets with charts](../charts/chartSheets.md)
explains what happens to charts when the sheets they refer to are copied, renamed and
deleted.

## See also

- API: [Worksheets](@ref) in Files and worksheets
- Related guides: [Protecting a sheet](protection.md), [Page setup and printing](printing.md),
  [Named ranges](../formulae/namedRanges.md)
