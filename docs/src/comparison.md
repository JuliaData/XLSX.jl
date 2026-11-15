# Compared with Python's readers

This page sets XLSX.jl's features beside those of three Python packages often used to
read Excel files: [python-calamine](https://github.com/dimastbk/python-calamine)
(Python bindings for the Rust [calamine](https://github.com/tafia/calamine) reader),
[fastexcel](https://github.com/ToucanToco/fastexcel) (built on calamine and Apache Arrow)
and [openpyxl](https://openpyxl.readthedocs.io/). It is meant to help you choose, or to
help you move from one to another.

The Python columns were checked on 2026-11-14 against each project's documentation and
type stubs, for python-calamine 0.9.1, fastexcel 0.21.0 and openpyxl 3.1. "Not
documented" means the project's documentation and public API don't mention the feature.
It may still exist in a form they don't describe. In short:

- **python-calamine** and **fastexcel** only read. They read more file formats than
  XLSX.jl, but they can't write files or read formatting, formulas or objects. On tables,
  XLSX.jl's `readtable` is faster than python-calamine and about as fast as fastexcel
  (see [Speed](#Speed)).
- **openpyxl** reads and writes `.xlsx` files, as XLSX.jl does, with a similar range of
  features, but it is much slower to read.
- **XLSX.jl** reads and writes `.xlsx` files, edits them in place, and covers formatting,
  charts and Excel's newer features. It reads only the Office Open XML formats.

## File formats and sources

| Feature | XLSX.jl | python-calamine | fastexcel | openpyxl |
|:--------|:--------|:----------------|:----------|:---------|
| `.xlsx`, `.xlsm` | read and write | read | read | read and write |
| Templates (`.xltx`, `.xltm`) | read and write | not documented | not documented | read and write |
| `.xlsb`, `.xls`, `.ods` | — | read | read¹ | — |
| Password-encrypted files | — (decrypt first with [XLSXDecrypt.jl](https://github.com/TimG1964/XLSXDecrypt.jl)) | — (raises `PasswordError`) | not documented | — |
| From a path | ✓ | ✓ | ✓ | ✓ |
| From an `IO` stream or bytes | ✓ | ✓ (file-like objects) | bytes only | ✓ (file-like objects) |

¹ fastexcel doesn't list the formats it reads. It opens files through calamine, which
reads all of these.

## Reading data

| Feature | XLSX.jl | python-calamine | fastexcel | openpyxl |
|:--------|:--------|:----------------|:----------|:---------|
| Cell values | ✓ `sheet["A1"]`, [`XLSX.readdata`](@ref) | ✓ | through a table | ✓ |
| A table, typed by column | ✓ [`XLSX.readtable`](@ref), any Tables.jl sink such as a `DataFrame` | lists of rows | ✓ to Arrow, pandas or Polars | lists of rows (`ws.values`) |
| Dates, times and Booleans | ✓ `Date`, `DateTime`, `Time`, `Bool` | ✓ `date`, `datetime`, `time`, `timedelta`, `bool` | ✓ date, datetime, duration and Boolean columns | ✓ `datetime` and others, from the number format |
| Error values (`#N/A`, …) | ✓ [`XLSX.iserror`](@ref), [`XLSX.geterror`](@ref); `missing` in tables | not documented | `None`, with the cell's position from `to_arrow_with_errors` | as text (`"#N/A"`) |
| Cached values of formulas | ✓ | ✓ | ✓ | ✓ with `data_only=True` |
| Formula text | ✓ [`XLSX.getFormula`](@ref) | — | — | ✓ (the default) |
| Calculating formulas | — | — | — | — |
| Choosing the header row and first row | ✓ `first_row`, `header`, `column_labels` | — | ✓ `header_row`, `skip_rows`, `column_names` | — |
| Choosing columns | ✓ `"B:D"` | — | ✓ by name, index or `"A:C,E"` | — |
| Stopping at an empty row or a condition | ✓ `stop_in_empty_row`, `stop_in_row_function` | — | `n_rows` | — |
| Setting column types | — (inferred, or `infer_eltypes=false`) | — | ✓ `dtypes`, `dtype_coercion` | — |
| Text read as missing | ✓ `missing_strings` | — | `whitespace_as_null` | — |
| Reading row by row without loading the sheet | ✓ [`XLSX.eachtablerow`](@ref), `enable_cache=false` | ✓ `iter_rows` | lazy loading | ✓ `read_only=True` |
| Excel Tables | ✓ [`XLSX.tables`](@ref), read and write | ✓ read (`load_tables=True`) | ✓ read (`load_table`) | ✓ read and write |
| Named ranges | ✓ [`XLSX.getDefinedNames`](@ref), read and write | — | ✓ read (`defined_names`) | ✓ read and write |
| Sheet visibility | ✓ [`XLSX.getSheetVisibility`](@ref), read and write | ✓ read | ✓ read | ✓ read and write |
| Merged cells | ✓ [`XLSX.getMergedCells`](@ref), read and write | ✓ read (`.xlsx`, `.xls`) | — | ✓ read and write |

## Writing and editing

| Feature | XLSX.jl | python-calamine | fastexcel | openpyxl |
|:--------|:--------|:----------------|:----------|:---------|
| New files | ✓ [`XLSX.writetable`](@ref), [`XLSX.newxlsx`](@ref) | — | — | ✓ |
| Editing an existing file in place | ✓ `mode="rw"` (see the warning on [`XLSX.openxlsx`](@ref)) | — | — | ✓ (shapes are lost) |
| Writing large files in constant memory | ✓ `streaming=true`, [`XLSX.streamxlsx`](@ref) | — | — | ✓ `write_only=True` |
| Formulas with a cached value | ✓ [`XLSX.setFormula`](@ref) with `value` | — | — | formula only |
| Inserting and deleting rows and columns | ✓ moving formulas, formats and objects | — | — | moves cells; formulas aren't updated |
| VBA macros | kept, not edited | — | — | kept with `keep_vba=True`, not edited |

## Formatting and objects

Neither python-calamine nor fastexcel reads or writes anything in this table.

| Feature | XLSX.jl | openpyxl |
|:--------|:--------|:---------|
| Fonts, fills, borders, alignment | ✓ read and write | ✓ read and write |
| Number formats | ✓ read and write, and [`XLSX.getText`](@ref) for the text Excel shows | ✓ read and write |
| Rich text in a cell | ✓ | ✓ `rich_text=True` |
| Conditional formatting | ✓ read and write | ✓ read and write |
| Data validation | ✓ read and write | ✓ read and write |
| AutoFilter | ✓ filters and sorts the rows, as Excel does | stores the criteria only; Excel applies them |
| Column widths, row heights, freeze panes | ✓, and [`XLSX.autofitColumns`](@ref) | ✓ |
| Charts | ✓ read, write and edit, including chartEx charts (experimental) | ✓ write; existing charts are kept |
| Images | ✓ read and write, and pictures in cells | ✓ write (needs Pillow); existing images are kept |
| Text boxes and shapes | ✓ read and write | — (lost on save) |
| Comments | ✓ notes and threaded comments | ✓ notes (formatting lost on save) |
| Hyperlinks | ✓ read and write | ✓ read and write |
| Sparklines, checkboxes | ✓ read and write | — |
| Sheet and workbook protection | ✓ | ✓ |
| Page setup, print areas, headers and footers | ✓ | ✓ |
| Document properties | ✓ | ✓ |

## Speed

On the benchmark fixtures, XLSX.jl's [`XLSX.readtable`](@ref) reads a table 11 to 14 times
as fast as openpyxl, in 0.4 to 0.8 times python-calamine's time, and in about the same
time as fastexcel (0.75 to 1.4 times its time). The exception is the smallest table
(100 rows), where fixed costs dominate and the Python readers are 2.6 to 3.2 times as fast.
Times in milliseconds (medians), for a few of the fixtures:

| Table | XLSX.jl | python-calamine | fastexcel | openpyxl |
|:------|--------:|----------------:|----------:|---------:|
| 5,000 rows × 30 columns, mixed | 49 | 77 | 45 | 626 |
| 50,000 rows × 70 columns, mixed | 1,095 | 1,775 | 1,073 | 13,650 |
| 50,000 rows × 20 columns, numbers | 227 | 387 | 194 | 2,677 |
| 50,000 rows × 10 columns, unique strings | 289 | 393 | 316 | 3,741 |
| 30,000 rows × 100 columns, as Excel writes them, 20% formulas | 783 | 1,394 | 809 | 8,680 |

These are XLSX.jl's times with the reading changes in the next release, made on 2026-11-14;
the Python readers' times were made on 2026-11-13 on the same machine (Windows, 8 threads;
Python 3.13, calamine 0.9.1, fastexcel 0.21.0, openpyxl 3.1.5). The method is in
[`bench/README.md`](https://github.com/JuliaData/XLSX.jl/blob/master/bench/README.md#Against-Python's-readers).
Opening a workbook with [`XLSX.openxlsx`](@ref) to read or edit it does more work than
`readtable`, and isn't compared here.

See also: [Supported features](features.md) for every XLSX.jl feature and its functions.
