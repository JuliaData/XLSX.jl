# Files and worksheets

## Files

```@docs
XLSX.XLSXFile
XLSX.XLSXFile(::Any)
XLSX.readxlsx
XLSX.openxlsx
XLSX.opentemplate
XLSX.newxlsx
XLSX.writexlsx
XLSX.savexlsx
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
XLSX.addImage
```
