```@meta
DocTestSetup = :(import XLSX)
```

# Named ranges

A *defined name*, often called a named range, gives a name to a cell, a range, a
constant or a formula, as Excel's *Formulas → Name Manager* does. Formulas can then use
`=SUM(PriceList)` rather than `=SUM(Prices!$B$2:$B$5)`, and XLSX.jl functions accept the
name wherever they take a range.

Adding and deleting names changes the workbook, so open the file with `mode="rw"` (or
start from [`XLSX.newxlsx`](@ref)). Names can be read from any file.

## Adding names

[`XLSX.addDefinedName`](@ref) adds a name. What the name refers to depends on the value
given:

```julia
xf = XLSX.newxlsx("Prices")
ws = xf["Prices"]
ws["A1"] = ["Item", "Price", "Price inc VAT"]
ws["A2:A5"] = ["Tea", "Coffee", "Cake", "Toast"]
ws["B2:B5"] = [1.5, 2.2, 3.0, 1.8]

XLSX.addDefinedName(xf, "PriceList", "Prices!B2:B5")                  # a range
XLSX.addDefinedName(xf, "VAT", 0.2)                                   # a number
XLSX.addDefinedName(xf, "Shop", "The Corner Cafe")                    # a text constant
XLSX.addDefinedName(xf, "ItemCount", "=COUNTA(Prices!\$A\$2:\$A\$5)")  # a formula
```

A string that is a cell or range reference becomes a reference. One that begins with `=`
is a formula, as in the Name Manager. Any other string is a text constant. A reference
may be a single cell, a range, or several ranges separated by commas
(`"Prices!A1,Prices!C3:C5"`).

References are written as absolute references (`$B$2:$B$5`), as Excel's Name Manager
writes them. Pass `absolute=false` for a relative reference.

A name may contain letters, numbers, `_` and `\`, and must begin with a letter or `_`. It
can't look like a cell address, such as `A1` or `R1C1`. Names are matched
case-insensitively, as in Excel, and they share a namespace with
[Excel Tables](../tables/excelTables.md): a name can't be the same as a Table's name.

### Workbook and worksheet scope

A name can belong to the whole workbook or to one worksheet. The first argument decides
which:

- given the `XLSXFile`, the name is *workbook-scoped* and can be used from any sheet. A
  reference must then include the sheet name: `"Prices!B2:B5"`, not `"B2:B5"`;
- given a `Worksheet`, the name is *worksheet-scoped*. It belongs to that sheet, and a
  reference without a sheet name refers to that sheet:

```julia
XLSX.addDefinedName(ws, "Header", "A1:C1")   # scoped to the sheet "Prices"
```

The same name may be defined at both scopes, with different values. Within the sheet,
a formula uses the sheet's own definition.

## Reading names

[`XLSX.getDefinedNames`](@ref) lists the names in one scope, using the same rule: an
`XLSXFile` for the workbook's names, a `Worksheet` for that sheet's own names.
[`XLSX.getAllDefinedNames`](@ref) lists every name in the file, with its scope:

```jldoctest names; setup = :(import XLSX; xf = XLSX.newxlsx("Prices"); ws = xf["Prices"]; ws["A1"] = ["Item", "Price", "Price inc VAT"]; ws["A2:A5"] = ["Tea", "Coffee", "Cake", "Toast"]; ws["B2:B5"] = [1.5, 2.2, 3.0, 1.8]; XLSX.addDefinedName(xf, "PriceList", "Prices!B2:B5"); XLSX.addDefinedName(xf, "VAT", 0.2); XLSX.addDefinedName(xf, "Shop", "The Corner Cafe"); XLSX.addDefinedName(xf, "ItemCount", "=COUNTA(Prices!\$A\$2:\$A\$5)"); XLSX.addDefinedName(ws, "Header", "A1:C1"))
julia> XLSX.getDefinedNames(xf)
4-element Vector{DefinedName}:
 Name                       Value
 -------------------------- ------------------------
 "ItemCount"                "=COUNTA(Prices!\$A\$2:…
 "PriceList"                Prices!$B$2:$B$5
 "Shop"                     "The Corner Cafe"
 "VAT"                      0.2

julia> XLSX.getAllDefinedNames(xf)
5-element Vector{DefinedName}:
 Name                       Value                    [Scope]
 -------------------------- ------------------------ --------------------------------
 "ItemCount"                "=COUNTA(Prices!\$A\$2:… [Workbook]
 "PriceList"                Prices!$B$2:$B$5         [Workbook]
 "Shop"                     "The Corner Cafe"        [Workbook]
 "VAT"                      0.2                      [Workbook]
 "Header"                   Prices!$A$1:$C$1         [Worksheet: "Prices"]
```

Each is an [`XLSX.DefinedName`](@ref), with fields `name`, `scope` (`nothing` for the
workbook, or the sheet's name), `value` and `absolute`. Names Excel creates for its own
use, such as print areas and filter ranges, are left out. Pass `include_system=true` to
see them.

To read what a name refers to, index the file or sheet with it. A range gives its cell
values, and a constant or formula gives the value as stored:

```jldoctest names
julia> xf["PriceList"]
4×1 Matrix{Any}:
 1.5
 2.2
 3.0
 1.8

julia> xf["VAT"]
0.2

julia> xf["ItemCount"]          # a formula is not calculated
"=COUNTA(Prices!\$A\$2:\$A\$5)"

julia> ws["Header"]             # a worksheet-scoped name, from its sheet
1×3 Matrix{Any}:
 "Item"  "Price"  "Price inc VAT"
```

!!! note
    A name that is also a sheet's name, given to an `XLSXFile`, finds the sheet:
    `xf["Prices"]` is the worksheet `Prices`, not a defined name. Give names that differ
    from the sheet names to avoid this.

## Using names

In a formula, write the name as Excel would:

```julia
XLSX.setFormula(ws, "C2:C5", "=B2*(1+VAT)")
XLSX.setFormula(ws, "B7", "=SUM(PriceList)")
```

The functions that format cells, and the ones that hide, group, size and filter rows
and columns, take a name in place of a range (see
[Addressing cells and ranges](../concepts/addressing.md)):

```julia
XLSX.setFont(ws, "Header"; bold=true)
XLSX.setFormat(xf, "PriceList"; format="Currency")
```

A name is a good way to give a drop-down list its items, since the list can then be
changed in one place (see [Drop-down lists](@ref)):

```julia
lists = XLSX.addsheet!(xf, "Lists")
lists["A1:A3"] = ["North", "South", "East"]
XLSX.addDefinedName(xf, "Regions", "Lists!A1:A3")
XLSX.addDataValidation(ws, "D2:D5"; type=:list, formula1="Regions")
```

## Deleting names

[`XLSX.deleteDefinedName`](@ref) deletes names from one scope, again chosen by the first
argument. It takes a name, a vector of names, or the result of `getDefinedNames`:

```julia
XLSX.deleteDefinedName(xf, "Shop")                         # a workbook name
XLSX.deleteDefinedName(ws, XLSX.getDefinedNames(ws))       # every name scoped to ws
XLSX.deleteAllDefinedNames(xf)                             # every name, both scopes
```

If any name isn't defined in that scope, an error is thrown and nothing is deleted.
Formulas that use a deleted name aren't changed: Excel shows `#NAME?` for them, as it
does when a name is deleted in the Name Manager.

## When sheets change

Names follow the sheets they refer to:

- [`XLSX.renamesheet!`](@ref) updates references in names. After renaming `Prices` to
  `Menu`, `PriceList` refers to `Menu!$B$2:$B$5`;
- [`XLSX.deletesheet!`](@ref) keeps names that referred to the deleted sheet, with the
  value `=#REF!` as in Excel, and removes names scoped to it;
- [`XLSX.copysheet!`](@ref) gives the copy its own copies of the names scoped to the
  original sheet;
- [`XLSX.insertRows`](@ref) and the other insert and delete functions move and resize
  the ranges names refer to (see [Inserting and deleting rows and columns](../sheets/insertDelete.md)).

A `DefinedName` already read is a snapshot: it doesn't change when the sheet does. Read
the names again after such a change.

## See also

- API: [Defined names](@ref) in Cells and data
- Related guides: [Using Formulas](formulas.md), [Data validation](../sheets/dataValidation.md)
