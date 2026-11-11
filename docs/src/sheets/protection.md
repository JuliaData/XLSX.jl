```@meta
DocTestSetup = :(import XLSX)
```

# Protecting a sheet

Protecting a sheet changes the workbook, so open the file with `mode="rw"` (or start from
[`XLSX.newxlsx`](@ref)). The settings can be read from any file.

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

```jldoctest protect; setup = :(import XLSX; sheet = XLSX.newxlsx()[1]; XLSX.protectSheet(sheet; password="secret", sort=true, autoFilter=true))
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

Whole columns (`"B:B"`) and rows (`"5:6"`), and `:` with row or column numbers
(`setCellProtection(sheet, :, 2; locked=false)`), reach only as far as the sheet's
used range, as with the other format functions. Rows added below it later are
locked, so to leave room for new entries give an explicit range, such as
`"B2:B500"`.

```jldoctest unlock; setup = :(import XLSX; sheet = XLSX.newxlsx()[1]; XLSX.setCellProtection(sheet, "B2:B20"; locked=false))
julia> XLSX.getCellProtection(sheet, "B5")
(locked = false, hidden = false)
```

## See also

- API: [Sheet protection](@ref) in Files and worksheets, and [`XLSX.setCellProtection`](@ref) on the Cell formats page
- Related guides: [Data validation](dataValidation.md), to restrict what can be typed into the unlocked cells
