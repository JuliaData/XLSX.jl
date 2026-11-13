# Files and worksheets

Functions for opening, creating and saving files, and for the sheets in them. For a
walk-through, see the [XLSX Tutorial](../tutorial/XLSXtutorial.md) and [Creating and editing workbooks](../tutorial/creatingWorkbooks.md).

## Files

```@docs
XLSX.XLSXFile
XLSX.XLSXFile(::Any)
XLSX.Workbook
XLSX.readxlsx
XLSX.readto
XLSX.openxlsx
XLSX.savexlsx
XLSX.opentemplate
XLSX.newxlsx
XLSX.writexlsx
XLSX.writetable
```

## Streaming writes

Guide: [Writing large files](@ref). [`XLSX.writetable`](@ref) streams a whole table with
`streaming=true`; `XLSX.addsheet!` adds a sheet to a `StreamWriter`.

```@docs
XLSX.streamxlsx
XLSX.writerow!
XLSX.StreamWriter
XLSX.StreamSheet
```

## Files (using FileIO)

Guide: [FileIO Tutorial](../tutorial/FileIOtutorial.md).

!!! note

    These functions extend `FileIO.load` and `FileIO.save`. Call them as
    `FileIO.load(...)` and `FileIO.save(...)` after doing `using FileIO`.

```@docs
XLSX.load
XLSX.save
```

## Worksheets

Guide: [Managing worksheets](../sheets/worksheets.md). Images, comments and hyperlinks
are on their [own API page](objects.md).

```@docs
XLSX.Worksheet
XLSX.sheetnames
XLSX.sheetcount
XLSX.hassheet
XLSX.addsheet!
XLSX.renamesheet!
XLSX.copysheet!
XLSX.deletesheet!
XLSX.movesheet!
```

## Sheet visibility and tab colour

Guide: [Hiding a sheet](@ref), [Tab colour](@ref).

```@docs
XLSX.setSheetVisibility
XLSX.getSheetVisibility
XLSX.setTabColor
XLSX.getTabColor
XLSX.removeTabColor
```

## Sheet protection

Guide: [Protecting a sheet](@ref). Cell locking is set with [`XLSX.setCellProtection`](@ref).

```@docs
XLSX.protectSheet
XLSX.unprotectSheet
XLSX.isSheetProtected
XLSX.getSheetProtection
XLSX.protectWorkbook
XLSX.unprotectWorkbook
XLSX.isWorkbookProtected
XLSX.getWorkbookProtection
```

## Page setup and printing

Guide: [Printing](@ref).

```@docs
XLSX.setPageSetup
XLSX.getPageSetup
XLSX.setPageMargins
XLSX.getPageMargins
XLSX.setPrintOptions
XLSX.getPrintOptions
XLSX.setPrintArea
XLSX.getPrintArea
XLSX.removePrintArea
XLSX.setPrintTitles
XLSX.getPrintTitles
XLSX.removePrintTitles
XLSX.setHeaderFooter
XLSX.getHeaderFooter
XLSX.removeHeaderFooter
XLSX.headerFooterField
XLSX.headerFooterFont
```

## Document properties

Guide: [Document properties](../tutorial/creatingWorkbooks.md#Document-properties).

```@docs
XLSX.setDocumentProperties
XLSX.getDocumentProperties
XLSX.setCustomProperty
XLSX.getCustomProperty
XLSX.getCustomProperties
XLSX.deleteCustomProperty
```

## Errors

```@docs
XLSX.XLSXError
```