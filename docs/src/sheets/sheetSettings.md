# Sheet settings

These functions change settings that apply to a whole sheet. They change the
workbook, so open the file with `mode="rw"` (or start from
[`XLSX.newxlsx`](@ref)). The settings can be read from any file. They work on
chartsheets as well as worksheets.

## Hiding a sheet

[`XLSX.setSheetVisibility`](@ref) shows or hides a sheet's tab:

```julia
XLSX.setSheetVisibility(xf["Lookups"], :hidden)       # listed in Excel's Unhide menu
XLSX.setSheetVisibility(xf["Workings"], :veryHidden)  # not listed: only VBA can show it
XLSX.setSheetVisibility(xf["Lookups"], :visible)
```

```julia
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

```julia
julia> XLSX.getTabColor(xf["Summary"])
"FFFF0000"

julia> XLSX.removeTabColor(xf["Summary"])
true
```
