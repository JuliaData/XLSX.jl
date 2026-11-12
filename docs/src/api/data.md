# Cells and data

Functions for reading and writing cell values, formulas and defined names, and for
data validation. For a walk-through, see the [XLSX Tutorial](../tutorial/XLSXtutorial.md). The ways of naming cells and
ranges are described in [Addressing cells and ranges](../concepts/addressing.md).

## Cell referencing

```@docs
XLSX.CellRef
XLSX.SheetCellRef
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

## Cell formulas

Guide: [Using Formulas](../formulae/formulas.md).

```@docs
XLSX.setFormula
XLSX.getFormula
```

## Defined names

Guide: [Named ranges](../formulae/namedRanges.md).

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

## Data validation

See [Data validation](../sheets/dataValidation.md) for a guide.

```@docs
XLSX.addDataValidation
XLSX.getDataValidations
XLSX.getDataValidation
XLSX.deleteDataValidation
```

## Ignored errors

Guide: [Ignoring error-checking warnings](@ref).

```@docs
XLSX.setIgnoredErrors
XLSX.getIgnoredErrors
XLSX.removeIgnoredErrors
```

## Error Values

```@docs
XLSX.CellErrorType
XLSX.XL_NULL
XLSX.XL_DIV0
XLSX.XL_VALUE
XLSX.XL_REF
XLSX.XL_NAME
XLSX.XL_NUM
XLSX.XL_NA
XLSX.XL_SPILL
```
