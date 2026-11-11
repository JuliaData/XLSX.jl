# Tables

XLSX.jl uses "table" in two senses: tabular data held in plain cells, and native Excel
Tables (*Insert → Table*). The [Excel Tables](../tables/excelTables.md) guide explains the
difference and covers both; the [XLSX Tutorial](../tutorial/XLSXtutorial.md) covers
reading and writing tabular data.

## Tabular data in cell ranges

```@docs
XLSX.DataTable
XLSX.gettable
XLSX.readtable
XLSX.gettransposedtable
XLSX.readtransposedtable
XLSX.writetable!
```

## Native Excel Tables

```@docs
XLSX.Table
XLSX.TableStyleInfo
XLSX.tables
XLSX.table
XLSX.addtable!
XLSX.appendtable!
XLSX.deletetable!
XLSX.settotals!
XLSX.gettablerange
XLSX.gettotals
XLSX.removetotals!
```
