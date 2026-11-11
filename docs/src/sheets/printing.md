# Printing

Setting up a sheet for printing changes the workbook, so open the file with `mode="rw"` (or start from
[`XLSX.newxlsx`](@ref)). The settings can be read from any file.

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

## See also

- API: [Page setup and printing](@ref) in Files and worksheets
- Related guides: [Freeze and split panes](../formatting/freezeAndSplitPanes.md), which keep header rows in view on screen as print titles do on paper
