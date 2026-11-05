# Sheet settings

These functions change settings that apply to a whole sheet: its visibility,
tab colour, protection and printing. They change the workbook, so open the file
with `mode="rw"` (or start from [`XLSX.newxlsx`](@ref)). The settings can be
read from any file. Visibility and tab colour work on chartsheets as well as
worksheets.

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

## Protecting a sheet

[`XLSX.protectSheet`](@ref) protects a worksheet, as Excel's
*Review → Protect Sheet* does, so that its cells can't be changed in Excel:

```julia
XLSX.protectSheet(sheet)
XLSX.protectSheet(sheet; password="secret", sort=true, autoFilter=true)
```

Keywords say what users may still do, as the check boxes in Excel's dialog do:
selecting cells is allowed by default, and everything else (formatting,
inserting and deleting rows and columns, sorting, filtering, ...) is not. See
[`XLSX.protectSheet`](@ref) for the full list.

!!! warning

    Sheet protection guards against mistakes, not against anyone determined to
    change the sheet. The password is stored as Excel's legacy 16-bit hash,
    which is easily broken, and other programs, XLSX.jl among them, read and
    write a protected sheet as any other.

```julia
julia> XLSX.isSheetProtected(sheet)
true

julia> XLSX.getSheetProtection(sheet).sort
true

julia> XLSX.unprotectSheet(sheet)
true
```

### Unlocking cells

Every cell is locked by default, so protection locks the whole sheet. To let
users fill in some cells of a protected sheet, unlock them with
[`XLSX.setCellProtection`](@ref), as the *Protection* tab of Excel's
*Format Cells* dialog does:

```julia
XLSX.setCellProtection(sheet, "B2:B20"; locked=false)   # the cells to fill in
XLSX.setCellProtection(sheet, "D2:D20"; hidden=true)    # hide these cells' formulas
XLSX.protectSheet(sheet)
```

Empty cells can be unlocked too: they are given the format without a value.
A hidden cell still shows its value; only its formula is kept from the formula
bar. Neither setting has any effect until the sheet is protected.

```julia
julia> XLSX.getCellProtection(sheet, "B5")
(locked = false, hidden = false)
```

## Printing

[`XLSX.setPageSetup`](@ref) sets the orientation, paper size and scaling, as
Excel's *Page Setup* dialog does:

```julia
XLSX.setPageSetup(sheet; orientation=:landscape, paperSize=:A4)
XLSX.setPageSetup(sheet; fitToWidth=1)   # one page wide, as many pages tall as needed
XLSX.setPageSetup(sheet; scale=75)       # or print at 75%
```

The margins are set in inches, and print options choose whether gridlines and
row and column headings are printed, and whether the sheet is centred on the page:

```julia
XLSX.setPageMargins(sheet; left=0.5, right=0.5, top=0.75, bottom=0.75)
XLSX.setPrintOptions(sheet; gridLines=true, horizontalCentered=true)
```

A print area limits what is printed, and print titles repeat rows at the top of
each page, or columns at its left:

```julia
XLSX.setPrintArea(sheet, "A1:H60")
XLSX.setPrintTitles(sheet; rows=1)          # repeat the header row on every page
XLSX.setPrintTitles(sheet; rows=1:2, cols="A")
```

Excel keeps these two as names scoped to the sheet, `_xlnm.Print_Area` and
`_xlnm.Print_Titles`, which are kept up to date as sheets are renamed, copied
and deleted.

Each setting can be read back with [`XLSX.getPageSetup`](@ref),
[`XLSX.getPageMargins`](@ref), [`XLSX.getPrintOptions`](@ref),
[`XLSX.getPrintArea`](@ref) and [`XLSX.getPrintTitles`](@ref), and the print
area and titles removed with [`XLSX.removePrintArea`](@ref) and
[`XLSX.removePrintTitles`](@ref).
