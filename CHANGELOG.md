# Release Notes

All notable changes to this package will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/).

## Unreleased

### BREAKING

- `infer_eltypes` now defaults to `true` (e.g. in `gettable` and `readtable`). This is the more common use case but,
  if it is not *your* use case you will need explicitly to set `infer_eltypes = false` in the relevant functions.

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