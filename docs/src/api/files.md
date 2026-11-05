# Files and worksheets

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

## Files (using FileIO)

!!! note

    These functions extend `FileIO.load` and `FileIO.save`. Call them as
    `FileIO.load(...)` and `FileIO.save(...)` after doing `using FileIO`.

```@docs
XLSX.load
XLSX.save
```

## Worksheets

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
XLSX.addImage
XLSX.getImages
XLSX.addComment
XLSX.getComment
XLSX.getComments
XLSX.deleteComment
XLSX.addHyperlink
XLSX.getHyperlink
XLSX.getHyperlinks
XLSX.deleteHyperlink
```

## Sheet visibility and tab colour

```@docs
XLSX.setSheetVisibility
XLSX.getSheetVisibility
XLSX.setTabColor
XLSX.getTabColor
XLSX.removeTabColor
```

## Errors

```@docs
XLSX.XLSXError
```