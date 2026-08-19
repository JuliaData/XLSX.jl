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

```@docs
XLSX.addDefinedName
```
