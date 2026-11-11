# Column width and row height

Two functions offer the ability to set the column width and row height within a worksheet. These can use 
all of the ways of naming rows and columns described in [Addressing cells and ranges](../concepts/addressing.md). For example:

```julia
julia> XLSX.setRowHeight(s, "A2:A5"; height=25)  # Rows 2 to 5 (columns ignored)

julia> XLSX.setColumnWidth(s, 5:5:100; width=50) # Every 5th column up to column 100.
```

Excel applies some padding to user specified widths and heights. The two functions described here attempt 
to do something similar but it is not an exact match to what Excel does. User specified row heights and 
column widths will therefore differ by a small amount from the values you would see setting the same 
widths in Excel itself.

## See also

- API: [Sheet layout](../api/sheetLayout.md)
- Related guides: [Hiding and grouping rows and columns](hideAndGroup.md), [Cell formats](cellFormatting.md), [Merged cells](mergedCells.md)
- Examples: [Creating a formatted form](@ref)
