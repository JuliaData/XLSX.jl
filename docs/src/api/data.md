# Cells and data

## Cell referencing

```@docs
XLSX.CellRef
XLSX.row_number
XLSX.column_number
XLSX.eachrow
XLSX.eachtablerow
```

## Cell data

```@docs
#XLSX.CellConcreteType
XLSX.readdata
XLSX.getdata
XLSX.getcell
XLSX.getcellrange
XLSX.iserror(::XLSX.Worksheet, ::AbstractString)
XLSX.geterror(::XLSX.Worksheet, ::AbstractString)
```

## Tables
```@docs
XLSX.DataTable
XLSX.gettable
XLSX.readtable
XLSX.gettransposedtable
XLSX.readtransposedtable
XLSX.writetable!
XLSX.Table
XLSX.tables
XLSX.table
XLSX.addtable!
XLSX.appendtable!
XLSX.deletetable!
XLSX.settotals!
XLSX.gettotals
XLSX.removetotals!
```

## Cell formulas

```@docs
XLSX.setFormula
XLSX.getFormula
```

## Defined names

Defined names have two scopes. A workbook-scoped name is visible throughout the
file; a worksheet-scoped name belongs to one sheet, and a name may exist at both
scopes at once with different definitions.

Every function below takes its scope from its first argument: an `XLSXFile`
means workbook scope, a `Worksheet` means that worksheet's own scope. So
`getDefinedNames(ws)` returns only the names scoped to `ws` — not the
workbook-scoped names, even though those can be used from `ws` in a formula.
Use [`XLSX.getAllDefinedNames`](@ref) to see every name in a file at once.

Names are matched case-insensitively, as in Excel. Defined names also share a
namespace with Excel Table names — see [Excel Tables](@ref).

```@docs
XLSX.DefinedName
XLSX.getDefinedNames
XLSX.getAllDefinedNames
XLSX.addDefinedName
XLSX.deleteDefinedName
XLSX.deleteAllDefinedNames
```
