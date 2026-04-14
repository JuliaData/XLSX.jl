# Release Notes

All notable changes to this package will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/).

## [Unreleased]
- Add support for custom missing strings (issue #90)
- Widen `setdata!` to accept an AbstractArray (issue #158)
- Fix intermittent bug with race condition when reading styles

## [v0.11.2](https://github.com/JuliaData/XLSX.jl/tree/v0.11.2) - 2026-04-01
Bug-fix to support 32-bit Julia

## [v0.11.1](https://github.com/JuliaData/XLSX.jl/tree/v0.11.1) - 2026-03-30
Minor bug-fix to get the StyledStrings extension working in Julia v1.14

## [v0.11.0](https://github.com/JuliaData/XLSX.jl/tree/v0.11.0) - 2026-03-23
This release introduces significant new functionality as set out below.

There are almost no changes in existing functional APIs in v0.11.0 compared with v0.10.4. Those changes that have been made are described briefly here.

This version drops support for Julia v1.6, and requires at least Julia v1.8.

### Breaking changes
There is only one breaking change in this version:

- `infer_eltypes` now defaults to `true` (e.g. in `gettable` and `readtable`). This is the more common use case but,
  if it is not *your* use case, you will need explicitly to set `infer_eltypes = false` in the relevant functions.

All other changes either introduce new functionality (documented elsewhere) or relate to internals only.

### New Functions
A number of new functions have been added compared with v0.10.4.

These include 18 new functions to support formatting of cells and cell values together with functions to copy or delete a sheet, to merge cells and to add new defined names for cells or cell ranges. In addition, it is now also possible to assign `AnnotatedStrings` (from [StyledStrings.jl](https://github.com/JuliaLang/StyledStrings.jl)) to cells to create content using Excel's rich text formatting.

A new function, `XLSXFile`, is provided that takes a `Tables.jl` compatible table and creates a new `XLSXFile` object for writing and which can act as a sink for functions such as `CSV.read`.

A new function, `renamesheet!` is created to replace `rename!` for consistency in naming with `addsheet!`, `copysheet!` and `deletesheet!` and to avoid potential name conflicts when exported (e.g. with `DataFrames.rename!`). However, the existing function `XLSX.rename!` is retained (but not exported) to avoid a breaking change.

Two new functions, `gettransposedtable` and `readtransposedtable`, mirror `gettable` and `readtable` for worksheet tables that have data organised in rows rather than columns.

Some additional convenience functions have also been added to streamline functions that were already available (such as `newxlsx`, `savexlsx`).

A wide range of additional indexing options is now widely supported by most functions. Most functions now support indexing rows and columns using vectors, ranges and step ranges and will accept a colon.

### Exported Functions
Most useful functions are now public, and can be used without the `XLSX.` prefix. The following function names are now exported:

- Files and worksheets
    `XLSXFile`, `readxlsx`, `openxlsx`, `opentemplate`, `newxlsx`, `writexlsx`, `savexlsx`,
    `Worksheet`, `sheetnames`, `sheetcount`, `hassheet`,
    `addsheet!`, `renamesheet!`, `copysheet!`, `deletesheet!` 

- Cells & data
    `CellRef`, `row_number`, `column_number`, `eachtablerow`,
    `readdata`, `getdata`, `gettable`, `readtable`, `readto`,
    `gettransposedtable`, `readtransposedtable`,
    `writetable`, `writetable!`,
    `setFormula`,
    `addDefinedName`

- Formats
    `setFormat`, `setFont`, `setBorder`, `setFill`, `setAlignment`,
    `setUniformFormat`, `setUniformFont`, `setUniformBorder`, `setUniformFill`, `setUniformAlignment`, `setUniformStyle`,
    `setConditionalFormat`,
    `RichTextString`, `RichTextRun`,
    `setColumnWidth`, `setRowHeight`,
    `getMergedCells`, `isMergedCell`, `getMergedBaseCell`, `mergeCells`

The iterator `XLSX.eachrow` has retained the XLSX prefix to avoid making a breaking change. However, `Base.eachrow` now refers to `XLSX.eachrow` for `XLSX.Worksheet` data types, meaning that  `eachrow` can be used without qualification, too.

### Fixed issues

This release addresses the following issues:
https://github.com/JuliaData/XLSX.jl/issues/52, https://github.com/JuliaData/XLSX.jl/issues/61, https://github.com/JuliaData/XLSX.jl/issues/63, https://github.com/JuliaData/XLSX.jl/issues/80, https://github.com/JuliaData/XLSX.jl/issues/88, https://github.com/JuliaData/XLSX.jl/issues/120, https://github.com/JuliaData/XLSX.jl/issues/147, https://github.com/JuliaData/XLSX.jl/issues/148, https://github.com/JuliaData/XLSX.jl/issues/150, https://github.com/JuliaData/XLSX.jl/issues/155, https://github.com/JuliaData/XLSX.jl/issues/156, https://github.com/JuliaData/XLSX.jl/issues/159, https://github.com/JuliaData/XLSX.jl/issues/165, https://github.com/JuliaData/XLSX.jl/issues/172, https://github.com/JuliaData/XLSX.jl/issues/179, https://github.com/JuliaData/XLSX.jl/issues/184, https://github.com/JuliaData/XLSX.jl/issues/189, https://github.com/JuliaData/XLSX.jl/issues/190, https://github.com/JuliaData/XLSX.jl/issues/198, https://github.com/JuliaData/XLSX.jl/issues/222, https://github.com/JuliaData/XLSX.jl/issues/224, https://github.com/JuliaData/XLSX.jl/issues/232, https://github.com/JuliaData/XLSX.jl/issues/234, https://github.com/JuliaData/XLSX.jl/issues/235, https://github.com/JuliaData/XLSX.jl/issues/238, https://github.com/JuliaData/XLSX.jl/issues/239, https://github.com/JuliaData/XLSX.jl/issues/241, https://github.com/JuliaData/XLSX.jl/issues/243, https://github.com/JuliaData/XLSX.jl/issues/251, https://github.com/JuliaData/XLSX.jl/issues/252, https://github.com/JuliaData/XLSX.jl/issues/253, https://github.com/JuliaData/XLSX.jl/issues/258, https://github.com/JuliaData/XLSX.jl/issues/259, https://github.com/JuliaData/XLSX.jl/issues/260, https://github.com/JuliaData/XLSX.jl/issues/275, https://github.com/JuliaData/XLSX.jl/issues/276, https://github.com/JuliaData/XLSX.jl/issues/277, https://github.com/JuliaData/XLSX.jl/issues/278, https://github.com/JuliaData/XLSX.jl/issues/281, https://github.com/JuliaData/XLSX.jl/issues/284, https://github.com/JuliaData/XLSX.jl/issues/299, https://github.com/JuliaData/XLSX.jl/issues/301, https://github.com/JuliaData/XLSX.jl/issues/305, https://github.com/JuliaData/XLSX.jl/issues/308,  https://github.com/JuliaData/XLSX.jl/issues/311, https://github.com/JuliaData/XLSX.jl/issues/314, https://github.com/JuliaData/XLSX.jl/issues/316, https://github.com/JuliaData/XLSX.jl/issues/324, https://github.com/JuliaData/XLSX.jl/issues/331, https://github.com/JuliaData/XLSX.jl/issues/335, https://github.com/JuliaData/XLSX.jl/issues/338, https://github.com/JuliaData/XLSX.jl/issues/342.

### Documentation
The documentation for this package has been extended substantially to cover the new functionality and all changes are (should be) reflected therein. In particular, a detailed guide to using the new formatting functions has been added.

### Internal changes
A number of changes to package internals have been made. Specifically, changes have been made to the following data `struct`s:

- `SheetRowStreamIteratorState`
- `WorksheetCacheIteratorState`
- `WorksheetCache`
- `XLSXFile`
- `Workbook`
- `Worksheet`
- `SheetRow`
- `Cell`

In particular, the internal memory configuration of an `XLSXFile` object and its components has been changed significantly, nearly halving the package's memory footprint.

### Changed dependencies
v0.11.0 has now fully migrated to `ZipArchives.jl` whereas v0.10.4 relied upon both this and `ZipFiles.jl`. In addition, xml support is now from `XML.jl` rather than `EzXML.jl`.

The use of `AnnotatedStrings` is supported through a package extension. This requires `StyledStrings.jl` to be in the active environment. 

New functionality that has been added has brought the following additional dependencies compared with v0.10.4:
- `Colors.jl`
- `UUIDs.jl`
- `Random.jl`

In addition, the test suite now has dependencies on `CSV.jl`,  `Distributions.jl` and `StyledStrings.jl`.

### Precompilation
v0.11.0 now makes use of `PrecompileTools.jl` (initially only in a small way).

## [v0.10.4](https://github.com/JuliaData/XLSX.jl/tree/v0.10.4) - 2024-09-29

This is a bugfix release.

- Update table.jl: promoting type for columns mixing integer and float values [#269](https://github.com/JuliaData/XLSX.jl/pull/269) (@rcqls)
- Remove the gc call in write.jl [#271](https://github.com/JuliaData/XLSX.jl/pull/271) (@TimG1964)

## [v0.10.3](https://github.com/JuliaData/XLSX.jl/tree/v0.10.3) - 2024-09-09

- support files without default namespace [#267](https://github.com/JuliaData/XLSX.jl/pull/267) (@hhaensel)
- Faster writing using ZipArchives.jl [#266](https://github.com/JuliaData/XLSX.jl/pull/266) (@nhz2)

This version drops support for Julia v1.3, and requires at least Julia v1.6.

## [v0.10.2](https://github.com/JuliaData/XLSX.jl/tree/v0.10.2) - 2024-08-31

- Document CellRef [#257](https://github.com/JuliaData/XLSX.jl/pull/257) (@nathanrboyer)
- Update read.jl to pass through Custom XML internal files [#261](https://github.com/JuliaData/XLSX.jl/pull/261) (@TimG1964)
- Added option to not write column names when writing dataframes to xlsx [#265](https://github.com/JuliaData/XLSX.jl/pull/265) (@ST2-EV)

## [v0.10.1](https://github.com/JuliaData/XLSX.jl/tree/v0.10.1) - 2023-12-30

This is a bugfix release.

- weaker assertion of relationship [#249](https://github.com/JuliaData/XLSX.jl/pull/249)
- support IO for writetable [#245](https://github.com/JuliaData/XLSX.jl/pull/245)
- add consts for max size and assert in writetable! [#247](https://github.com/JuliaData/XLSX.jl/pull/247)

Many thanks to @hhaensel and @nathanrboyer!

## [v0.10.0](https://github.com/JuliaData/XLSX.jl/tree/v0.10.0) - 2023-08-22

This release contains no breaking changes regarding the public XLSX API.

There's a breaking change regarding Cell struct: formula field changed type from String to AbstractFormula.

There's a breaking change regarding TableRowIterator struct: added field keep_empty_rows.

- Fixes row formatting that was previously lost [#227](https://github.com/JuliaData/XLSX.jl/pull/227) (@best4innio)
- Add keep_empty_rows kwarg [#220](https://github.com/JuliaData/XLSX.jl/pull/220) (@rben01)
- add colon indexing [#216](https://github.com/JuliaData/XLSX.jl/pull/216) (@BeastyBlacksmith, @divbyzerofordummies)
- Docs review [#229](https://github.com/JuliaData/XLSX.jl/pull/229) (@goggle)

## [v0.9.0](https://github.com/JuliaData/XLSX.jl/tree/v0.9.0) - 2023-03-08

This release contains no breaking changes regarding the public XLSX API.

It contains a breaking changing on an internal struct: `XLSXFile` `filepath::AbstractString` field was replaced by `source::Union{AbstractString, IO}`.

- support reading from IO as well as a file path [#217](https://github.com/JuliaData/XLSX.jl/pull/217) (@tecosaur)
