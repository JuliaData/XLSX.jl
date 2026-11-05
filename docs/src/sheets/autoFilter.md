# Filtering

An AutoFilter puts drop-down buttons on the header row of a range, as Excel's
*Data → Filter* does, so that its rows can be filtered and sorted in Excel.
Adding and removing a filter changes the workbook, so open the file with
`mode="rw"` (or start from [`XLSX.newxlsx`](@ref)). Filters can be read from any
file.

```julia
XLSX.openxlsx("sales.xlsx", mode="rw") do xf
    sheet = xf["Sales"]
    XLSX.setAutoFilter(sheet, "A1:F200")   # row 1 holds the headers
end
```

`XLSX.setAutoFilter(sheet, :)` filters the sheet's whole used range. The range
may also be given as a defined name, or with an `XLSXFile` as
`"Sales!A1:F200"`.

A sheet has a single filter, so setting one replaces any filter the sheet
already has, together with the criteria chosen in Excel. Setting the same
range again keeps them.

```julia
julia> XLSX.getAutoFilter(sheet)
A1:F200

julia> XLSX.removeAutoFilter(sheet)
true
```

Rows that a filter hid in Excel stay hidden when the filter is removed. Show
them with [`XLSX.setRowHidden`](@ref)`(sheet, rows; hidden=false)`.

## Filters in Excel Tables

An [Excel Table](../tables/excelTables.md) has filter buttons of its own, so a
sheet's filter and a table may not overlap: `setAutoFilter` refuses a range on
a table, and `addtable!`, `writetable!` with `as_table=true`, `appendtable!` and
`settotals!` refuse to place or grow a table onto the filter. [`XLSX.getAutoFilter`](@ref) returns
only the sheet's filter, not those of its tables.

## Sheet operations

The filter goes with its sheet: [`XLSX.copysheet!`](@ref) gives the copy the same
filter, and [`XLSX.renamesheet!`](@ref) and [`XLSX.deletesheet!`](@ref) keep the
hidden name Excel uses to record the filter's range in step.
