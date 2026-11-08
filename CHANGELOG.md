# Release Notes

All notable changes to this package will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/).

## [Unreleased]

- A column inserted inside a hidden column span is shown, with the span split around it, as a row inserted among hidden rows already was; and a group that gains a visible row or column loses the collapsed marker on its summary row or column. Both as Excel 365 does; the hidden span used to grow over the new column. Checked against Excel with the fixtures made for P1–P3, now used by the tests.
- `setLabelFlag(c::ChartEx, i, flag, on)` shows or hides a chartEx series' label parts (`:seriesName`, `:categoryName`, `:value`), writing labels as Excel 365 does. As in Excel, histogram and Pareto labels have no category name and a Pareto line has no labels.
- Reading conditional formats, and adding a number format, no longer walk every node of the sheet or styles part to find one element. `getConditionalFormats` on a 50,000-cell sheet read with `readxlsx` takes half the time (111 ms to 58 ms), and each new number format takes under a microsecond, from 58 μs.
- `setLabelFlag(c, i, flag, on)` shows or hides one part of a `c:` series' data labels (`:value`, `:seriesName`, `:categoryName`, `:legendKey`, `:percent` or `:bubbleSize`), and `getLabelFlag(c::Chart, i, flag)` reads it back. A series with no labels is given the ones Excel 365 writes when labels are turned on, checked for every chart kind. There was no way to make a `c:` chart's labels appear before.
- `insertRows` and `insertColumns` format the new rows and columns as Excel does by default ("Format Same As Above"): each new cell takes its neighbour's style, without its border; a row takes the height, style and outline level of the row above, and a column the width, style and outline level of the column to its left; and a conditional format or data validation ending just before the insert grows over it. Checked against Excel 365's own results. The new rows and columns used to be unformatted.
- Deleting all the rows or columns under an image or chart set to move and size with its cells deletes it, with its picture or chart part, as Excel does. It used to shrink to nothing and stay in the file.
- `makeCombo(c)` makes a `:column`, `:stackedColumn` or `:line` chart a combo chart, as Excel's Change Chart Type → Combo → "Clustered Column – Line on Secondary Axis" does: the last series becomes a line on new secondary axes and the others clustered columns. As in Excel, the title, legend, axes and the formatting of series that keep their type are kept, and a series that changes type takes the default formatting of its new group, keeping its name, data and data labels. Checked against Excel 365's own results.
- `addSeries` and `addChartEx` take absolute references, such as `Data!$B$2:$B$5`, `'My Data'!$B$1` or `$B$2:$B$5`, so a reference read from a series with `getChartSeries` can be passed straight back. They threw before.
- `addSeries` writes a series' children in schema order. A line or scatter series had its `c:smooth` written before its data, where the schema places it after.
- `addChart(ws, :combo; …)` creates Excel's "Clustered Column – Line on Secondary Axis" combo chart, from a chart Excel 365 wrote, and `addSeries` takes a `group` keyword, `:column` or `:line`, saying which of a combo chart's groups a series joins. It works on combo charts Excel made too, which `addSeries` refused before. `group` is required on a chart with more than one plot group and an error on one with a single group. Combo series are formatted as Excel formats them, with markers on the line giving Excel's Line with Markers formatting.
- `addSeries` numbers a new series (`c:idx`, `c:order` and its automatic accent colour) across all of a chart's series rather than within its group. This changes nothing for a chart with one group.
- `addChart` creates nine more kinds, from charts Excel 365 wrote: `:percentStackedColumn`, `:percentStackedBar`, `:stackedLine`, `:percentStackedLine`, `:stackedLineMarkers`, `:percentStackedLineMarkers`, `:stackedArea`, `:percentStackedArea` and `:filledRadar`. Series added with `addSeries` are formatted as Excel formats them, the second series and on included. `markers` and `smooth` apply to every line kind, and markers on a stacked line series give it the formatting of the kind with markers, as they do for `:line`
- `addSeries` reads which kind a chart is from its plot group's `barDir`, `grouping`, marker flag and `radarStyle`, compared with each kind's template, so an existing stacked or 100% stacked chart, from Excel or elsewhere, takes the matching kind's series formatting. A stacked bar chart was read as a clustered one; the two format their series alike, so files written before are unaffected

- `getSeriesFill` on a chartEx chart reports the colour Excel draws where the chart itself writes none, which until now was nearly always (Excel writes no fill for a `cx:` series until it is formatted). The colour is read from the chart's style and colours parts as a last rung of the cascade, `site.level === :style`, and was checked colour for colour against Excel 365: series of box & whisker, funnel, histogram and Pareto charts, a waterfall's increase, decrease and total points, and a treemap's or sunburst's points by top-level category. Lines, text, monochromatic colour schemes and `c:` charts are unchanged (a `c:` series without formatting is drawn from its `c:style`, not from these parts). A waterfall, treemap or sunburst point's `chain` now ends in the `:style` rung
- Test fixtures written by Excel 365, with SVG exports of their charts in `test/data/chart_style_svg`, record the colours Excel draws

- Internal: `src/charts/chartprops.jl` (2,972 lines) is split by object into `chartnodes.jl` (design notes, node locators and shared helpers), `charttext.jl`, `chartseries.jl`, `chartaxes.jl`, `chartlegend.jl` and `chartcascade.jl`. Every function moves unchanged. No change in behaviour

- Writing a worksheet allocates less and is 10–25% faster: `writetable` of 100,000 × 8 strings allocates 772 → 336 MiB, and of numbers 986 → 666 MiB (0.55 → 0.44 s). Integers and cell references are written without making strings, the tags are built once, rows are written a chunk at a time into one buffer per thread, and the sheet goes into the file in pieces rather than being copied whole twice. The files written are byte-for-byte the same
- `bench/worksheet_write.jl` times each write path on its own, without fixtures, and keeps earlier runs to compare against

- Add `deleteImage`, which deletes the images whose top-left corner is in a cell, or one image from `getImages`. The picture file goes too, unless another image still shows it, and a drawing left empty is removed. `getImages` is now declared public
- Add `Charts.deleteChart`, which deletes a chart on a worksheet with the parts only it uses, and for a chartEx chart its hidden `_xlchart` names. A chart on a chartsheet is refused: delete the sheet instead
- Add `getThreadedComments` and `getThreadedComment`, which read Excel 365's threaded comments: each thread's author, text, date and whether it is resolved, with its replies. Threaded comments are still not written

- `insertRows`, `deleteRows`, `insertColumns` and `deleteColumns` move Excel Tables, so sheets with tables can now be edited this way. A table grows to take in rows or columns inserted inside it (inserted columns become table columns named `Column1`, `Column2`, ...) and shrinks as they are deleted from it; one deleted whole is deleted. References to a deleted table or table column (`Sales[Amount]`) become `#REF!` everywhere, and formulas in calculated columns and totals rows move. Deleting a table's header row, totals row or all its data rows on their own is refused
- Internal: `addtable!` is broken into steps (choosing the name, reading the header, checking the totals row, attaching the table part)

- `insertRows`, `deleteRows`, `insertColumns` and `deleteColumns` move the sheet's images, charts and shapes, as each one's *move and size with cells* setting says (an object all of whose rows or columns are deleted shrinks to nothing), and its comments and their boxes. A comment on a deleted cell is deleted; a threaded comment is moved, but deleting its cell is refused. Sheets with images, charts or comments can now be edited this way

- Add `insertRows`, `deleteRows`, `insertColumns` and `deleteColumns`, which work as Excel's *Insert* and *Delete* commands do. Cells move, with their row heights and attributes, and everything that refers to them follows: formulas on every sheet (relative or absolute), merged cells, conditional formats and data validations (their ranges and formulas, x14 ones included), hyperlinks, the AutoFilter and its filter columns, frozen panes (deleting all the frozen rows and columns unfreezes the sheet) and selections, column widths, page breaks, defined names (including the print area and titles) and chart series. Ranges across an insert grow and ranges cut by a delete shrink; references to deleted cells become `#REF!`. A shared formula stays shared when the edit moves all its cells alike; one it would change unevenly (references across the edited rows, or a delete inside its block) is stored as one formula per cell afterwards.

- Internal: the `readtable` and `eachtablerow` methods share one keyword list rather than repeating it, and `mergeCells`/`removeMergedCells` and `setColumnWidth`/`setRowHeight` define their shared address methods once. No change in behaviour, except that an unknown keyword to `readtable` or `eachtablerow` is reported against the internal function the keywords are passed to

- Internal: `write.jl` is split into `save.jl` (writing the package), `setdata.jl` (writing cells), `writetable.jl` and `sheetops.jl` (adding, renaming, copying, deleting and moving sheets), and the Excel Tables half of `table.jl` moves to `exceltables.jl`. `copysheet!` and `deletesheet!` are broken into steps, and repeated code in the cell format setters, range handling, conditional format data bars and the sheet cache fill is shared. No change in behaviour
- `bench/maintainability.jl` counts duplication in code only: it leaves out docstrings, the bodies of `const` tables, lines of nothing but literals, and `include` lines, which it counted as duplicated because blanking their literals made unrelated ones look alike

- Add `setCellProtection` and `getCellProtection` to unlock cells, so they can be filled in on a protected sheet, and to hide cells' formulas, as the *Protection* tab of Excel's *Format Cells* dialog does. Empty cells are given the format without a value, and the cells' other formatting is kept
- A cell format with both an alignment and a protection setting writes them in the order the schema requires. `setAlignment` added the alignment after any protection Excel had written

- Add `protectSheet`, `unprotectSheet`, `isSheetProtected` and `getSheetProtection` to protect worksheets, with an optional password (Excel's legacy hash, so up to 15 ASCII characters) and the same choice of what users may still do as Excel's *Protect Sheet* dialog
- Add `setPageSetup`/`getPageSetup` (orientation, paper size, scaling or fit to pages), `setPageMargins`/`getPageMargins`, `setPrintOptions`/`getPrintOptions` (gridlines, headings, centring), and `setPrintArea`, `getPrintArea`, `removePrintArea`, `setPrintTitles`, `getPrintTitles` and `removePrintTitles` for printing. A `pageSetup` that links to printer settings keeps the link

- Add `setSheetVisibility` and `getSheetVisibility` to show and hide sheets (`:visible`, `:hidden` or `:veryHidden`), and `setTabColor`, `getTabColor` and `removeTabColor` for tab colours, on worksheets and chartsheets. Hiding the workbook's last visible sheet is an error, and hiding the sheet the workbook opens on makes the first visible sheet the one it opens on
- `Worksheet`'s internal `is_hidden` field is replaced by `visibility`, which tells hidden and very hidden sheets apart. `ws.is_hidden` still reads and sets it

- Add `setAutoFilter`, `getAutoFilter` and `removeAutoFilter` to add, read and remove the filter buttons on a plain range of a sheet, outside an Excel Table. The hidden `_xlnm._FilterDatabase` name Excel keeps for the filter is written and removed with it, and follows the sheet through `copysheet!`, `renamesheet!` and `deletesheet!`. A filter may not overlap an Excel Table
- `addtable!`, `writetable!` with `as_table=true`, `appendtable!` and `settotals!` throw an error rather than place or grow a table onto the sheet's AutoFilter, as Excel does not allow the two to overlap

- Add `setColumnHidden`, `setRowHidden`, `isColumnHidden` and `isRowHidden` to hide and show rows and columns, and `groupRows`, `groupColumns`, `ungroupRows`, `ungroupColumns`, `getRowOutlineLevel` and `getColumnOutlineLevel` to group them in an outline, nested up to Excel's seven levels and optionally collapsed. Rows and columns needn't hold data, and can be read from a file opened only for reading
- `setRowHeight` sets the height of a row with no cells, which is written with just its height. It skipped such rows, returning -1 if every row was empty
- `copysheet!` copies the attributes of the original's rows: their styles, and whether they are hidden or grouped. The copy's rows had none
- `setColumnWidth` joins adjacent column definitions left with the same settings

- Add `addDataValidation`, `getDataValidations`, `getDataValidation` and `deleteDataValidation` to add, read and delete data validation: drop-down lists (from `values` or a range), whole and decimal numbers, dates, times, text length and custom formulas, with input and error messages. Validations Excel 2010 wrote in its extended (`x14`) form are read and deleted too, and validations can be read from a file opened only for reading
- `renamesheet!`, `deletesheet!` and `copysheet!` now update sheet references in Excel 2010 (`x14`) data validations, whose formula is held in an `<xm:f>` inside `<x14:formula1>`. Those references were left unchanged

- Add `addHyperlink`, `getHyperlink`, `getHyperlinks` and `deleteHyperlink` to add, read and delete cell hyperlinks: to a web page, email address or file (`url`), or to a place in the workbook (`location`). An empty cell given a link takes its display text and a blue, underlined font, as in Excel. Links can be read from a file opened only for reading, with or without the cell cache
- `copysheet!` gives the copy's hyperlinks to web pages and files relationships of their own. The copy kept the original's relationship ids, which pointed at relationships the copy didn't have

- Docs: a *Chart subtypes* section in *Creating charts* maps each of Excel's chart subtypes to the kind or `addSeries` keywords that create it, and the *Limitations* list of kinds that can't be created is completed
- Docs: the *Migration Guide* covers v0.12 and v0.13, including every behaviour change in the v0.13 patch releases
- Docs: the README and home page list the package's features and give a quick start
- Docs: a new *Images, comments and sheet order* guide

## [v0.13.3](https://github.com/JuliaData/XLSX.jl/tree/v0.13.3) - 2026-11-04

- Requires XML.jl v0.4.7, and reads through its public API only: `XML.splicetext` and `XML.sourcetext` replace the XML.jl internals XLSX.jl read, including `XMLTokenizer._skip_element_raw`, which XML.jl v0.4.7 deprecates and v0.5 removes [#482](https://github.com/JuliaData/XLSX.jl/issues/482)
- XML.jl v0.4.7 reads each part in place, whatever its line ends ([XML.jl#194](https://github.com/JuliaData/XML.jl/pull/194)): a part with CR line ends, as Excel writes after the XML declaration, is not copied to normalise them, and the `<sheetData>`-less stub of a worksheet and the raw XML of each rich shared string keep the bytes they were read from, CRs included
- Each XML part is decompressed straight into a `String` rather than into a byte vector that was then copied, so reading a workbook allocates the size of its XML once less: `readxlsx` on a workbook with a 137 MiB worksheet allocates 170 MiB rather than 310 MiB [#482](https://github.com/JuliaData/XLSX.jl/issues/482)
- Opening a file by name reads it in blocks of 64 KiB, growing to 1 MiB for long reads, rather than opening it again for each small read, which was slow on Windows. The file is still not held open between reads, so it can be overwritten straight after it is read [#482](https://github.com/JuliaData/XLSX.jl/issues/482)
- The shared string table keeps the raw XML of a string only when it has rich text runs or phonetic hints. A plain string, one `<si>` holding a single `<t>`, is kept as its text alone and its XML is rebuilt when the workbook is written, so the table takes about half the memory it did for a workbook of plain strings (67 MiB rather than 153 MiB for one million). Writing a workbook of many unique strings allocates up to a third less, and is up to 25% faster [#482](https://github.com/JuliaData/XLSX.jl/issues/482)
- A string written to a cell reuses the entry for the same string already in a file's shared string table. It was added again, because the table was searched by the XML of each entry, and the XML XLSX.jl built never matched the XML Excel writes
- Streamed reads (`enable_cache=false`) and the shared string table now drop a leading BOM and line feed as other reads do
- A CR in cell text, a rich text run or a formula is written as `&#13;`, and a CR, LF or tab in an attribute value (such as a table column name) as `&#13;`, `&#10;` or `&#9;`, so each reads back unchanged. With XML.jl v0.4.7, which normalises line ends as XML 1.0 requires, a CR written as it was read back as LF; a LF or tab in an attribute value already read back as a space. XLSX.jl writes these through `XML.escape`, which does this from XML.jl v0.4.7 ([XML.jl#186](https://github.com/JuliaData/XML.jl/pull/186))
- **Behavior change:** text written in a file with a literal CR LF or CR, rather than `&#13;`, now reads as LF, as XML 1.0 requires (XML.jl v0.4.7). Excel itself writes a CR as `_x000D_`, which is unaffected
- An attribute of `<sheetData>` holding `&`, `<` or `"` no longer makes the stored worksheet malformed
- Add a `Base.size(::Worksheet)` that returns the worksheet dimension (#483)
- Add `addComment`, `getComment`, `getComments` and `deleteComment` to write, read and delete cell comments (Excel "Notes"). Threaded comments (Excel 365) are not written, and replacing or deleting one already in a file throws an error
- Add `movesheet!` to change a sheet's place in the tab order, before or after another sheet. Names scoped to the sheet and the active tab stay with it
- `renamesheet!` updates references to the renamed sheet in conditional format and data validation formulas on every sheet, and in defined names holding a formula; it updated only cell formulas, charts and defined names holding a reference. References are matched ignoring case, as Excel does, so `SHEET1!A1` (as `setConditionalFormat` writes it) is renamed too
- `deletesheet!` replaces references to the deleted sheet with `#REF!` in conditional format and data validation formulas on the remaining sheets, and in defined names holding a formula, as it already did in cell formulas
- **Behavior change:** `copysheet!` points the copy's references to the original sheet at the copy, as Excel does, in cell formulas, conditional formats, data validations and sheet-scoped defined names holding a formula: `=Sheet1!A1` on `Sheet1` becomes `=Copy!A1` on the copy. They were copied unchanged, so the copy read the original's cells. References to other sheets are unchanged

## [v0.13.2](https://github.com/JuliaData/XLSX.jl/tree/v0.13.2) - 2026-10-21

This release is largely a bug fixing release. There are a lot of small changes, mainly addressing edge cases. Some of these introduce minor behaviour changes. There is quite a long list under the fold.

<details>
<summary>Consolidated maintainability and bug fixing changes</summary>

- **Behavior change:** `openxlsx(f, source; mode="rw")` (and `mode="w"`) no longer writes `source` when the do block throws, so a failure part-way through cannot overwrite the original file with a half-edited workbook
- **Behavior change:** defined names holding a formula (or any expression that is not a reference or constant) are now read with a leading `=`, as Excel's Name Manager shows them, e.g. `"=COUNTA(Sheet1!\$A:\$A)"`. A plain string is a text constant and is written quoted; it was written unquoted, so Excel read it as a formula, and a constant read from a file was written back as a reference to a name. Pass a string starting with `=` to `addDefinedName` to define a name by a formula
- **Behavior change:** a negative number in a date-formatted cell is read as the number, rather than throwing and failing the read of the whole sheet
- Shifted cell references in shared formulas are correct: `getFormula` on a shared-formula cell could return `=A2+A20` for `=A1+A10` shifted one row, and shifted function names such as `LOG10` and `ATAN2` as if they were references
- `addImage` no longer empties a writable `IOBuffer` passed as the image. The buffer's contents and position are left as they were
- `openxlsx(io; mode="rw") do … end` and `savexlsx` on a workbook opened from an `IO` now replace the `IO`'s content with the saved workbook. They previously threw an `ArgumentError` from ZipArchives, because the workbook was written after the bytes already read. An `IO` that can't be overwritten, such as a read-only `IOBuffer`, gives an `XLSXError`
- Renaming a sheet no longer rewrites references to other sheets whose names contain it (renaming `Sheet1` turned `MySheet1!A1` into `MyData!A1` and lost that sheet's formulas). Deleting a sheet likewise no longer deletes or breaks formulas on such sheets, no longer throws when the deleted sheet has a shared formula, and replaces references to a sheet whose name contains characters such as `(` with `#REF!`
- Hidden and very hidden sheets stay hidden when the workbook is saved
- Writing a cell above or left of the sheet's range that is also below or right of it now extends the recorded dimension on every side, so `ws[:]` and the written `<dimension>` include it
- Sheet names that start with a digit or look like a cell address (`2024`, `A1`) are quoted in the references XLSX.jl writes
- A new custom number format takes an id not already in use, and an existing identical format is reused under its own id; ids were derived from positions, which collided or pointed at the wrong format when a file's custom ids were not contiguous from 164. `setFormat` accepts built-in format names in any case (`"percentage"`)
- `addImage` no longer overwrites an image still in use after another sheet's images were deleted
- `setColumnWidth` splits a column definition spanning several columns rather than writing an overlapping one, and writes column definitions in order; `getColumnWidth` reads a width from a spanning definition
- `mergeCells` inserts `<mergeCells>` where the schema requires it (after `autoFilter`, `sheetProtection` and the like)
- Conditional formats with `font=["bold"=>"true"]` or `["italic"=>"true"]` make text bold or italic; they wrote `val="0"`, which turns it off. `"under"` writes the requested underline (`single` or `double`) and rejects other values
- `setConditionalFormat(…, :notContainsText / :beginsWith / :endsWith; …)` makes the rule named; all three made a `containsText` rule. A text rule's `value` may contain `"`, and a cell reference is used as one, as documented
- Overwriting the first cell of a shared formula at least three columns wide and two rows high no longer throws a `KeyError`, and leaves the other cells' cached values unchanged
- `copysheet!` points the copy's sheet-scoped defined names at the copy (a non-contiguous one threw), and gives formulas set on the copy shared-formula ids of their own
- **Behavior change:** `deletesheet!` keeps defined names that refer to the deleted sheet with the value `"=#REF!"`, as Excel does. Workbook-scoped ones were deleted, and ones scoped to other sheets were left pointing at a sheet that no longer exists
- `setFormula` no longer throws for a two-argument `IF` or an `OFFSET` whose size is not a literal number; a `GROUPBY` or `PIVOTBY` call inside a longer formula keeps the rest of the formula; `raw` and `spill` apply to every cell of a range
- Worksheet parts not named `sheetN.xml` (e.g. `Sheet1.xml`, as some tools write) can be read and written
- A single-run `RichTextString` with no attributes, `:under => true` or a fractional `:size` can be written to a cell; `setFont` takes a fractional `size`. `getRichTextString` reads fractional sizes and `<b val="0"/>` / `<u val="none"/>` as off, and `setFont` no longer turns an explicit `<b val="0"/>` into bold
- Slicing a `RichTextString` uses the indices of its text, as for a `String`, so non-ASCII text slices correctly. Equal `RichTextRun`s hash equally
- `resolveColor` resolves `indexed="64"` and `"65"` (system foreground and background colors)
- Datetimes that round up to midnight move to the next day rather than back to the start of the same day
- Defined names that are cell addresses in lower case (`a1`) are refused, as Excel refuses them; sheet, table and table column names are compared ignoring case, as Excel does
- Totals-row formulas escape table column names containing `[`, `]`, `#` or `'`
- **Behavior change:** an unrecognized keyword argument to `setFont`, `setFill`, `setAlignment`, `setFormat`, `setBorder` or their `setUniform*` versions is an `XLSXError` naming the function, the keyword and the valid keywords, rather than a `MethodError`. `outside` given to `setBorder` for a single cell, or to `setUniformBorder`, is likewise an `XLSXError`
- The formatting setters (`setFont`, `setFill`, `setAlignment`, `setFormat`, `setBorder` and their `setUniform*` versions) compile once rather than again for each new combination of keywords, so a first call with keywords not used before takes tens of milliseconds rather than about half a second
- The precompile workload covers the common read and write paths (`writetable`, editing and saving a file, `readtable`, `readxlsx`, `eachrow`, `readdata`, Excel Tables, cell formatting), for file names as well as `IO`, so their first call in a session is much faster (about 1 s in all rather than about 5.5 s). It no longer precompiles conditional formats. Set the PrecompileTools preference `precompile_workload = false` to turn it off
- Faster: `getFormula` on shared-formula cells no longer re-reads the whole sheet for each call; reading an Excel Table without the cache reads the sheet once rather than once per cell; writing strings skips the illegal-character pass for clean text; filling each sheet's cache no longer scans every formula in the workbook
- Reading a workbook on several threads no longer reads and writes the style lookup tables at the same time
- An empty or truncated file, or a sheet index beyond the last sheet, gives an `XLSXError` rather than a `BoundsError`. A misnamed column of an Excel Table row gives an `XLSXError`
- `partition` with `breaks` excludes `Inf` and booleans, as documented
- The rule-based conditional formats (`:cellIs`, `:containsText`, `:top10`, `:aboveAverage`, `:timePeriod`, `:expression` and the blanks, errors, unique and duplicate rules) accept a Symbol for any string keyword, as `:containsText` already did, and a value of another type gives an `XLSXError` rather than a `MethodError`. An invalid `operator` no longer adds an unused format to the workbook's styles before the error
- `ws[row, col] = v` with a vector `v` now writes it across from that cell, as `ws["B2"] = v` does. Most other two-index vector assignments, such as `ws[1:2, 1:2] = v` or `ws[1, [1, 3]] = v`, threw an "ambiguous" `MethodError`; they now throw an `XLSXError` saying a vector is written along one row or column, or across from one cell
- Removed `getGapWidth`, `setGapWidth`, `getOverlap` and `setOverlap` from the `XLSX.Charts` export list. They were never defined: the functions are `getGroupGapWidth`, `setGroupGapWidth`, `getGroupOverlap` and `setGroupOverlap`
- `SheetCellRef`, `TableStyleInfo` and `Workbook` now have docstrings
- `setConditionalFormat(...; :iconSet)` now rejects an invalid `mid_val` (one that is neither a number nor a fixed cell reference, unless `mid_type="formula"`), as it does every other threshold value; any `mid_val` was previously accepted. Errors for `mid2_type` and `mid2_val` now name those keywords rather than `mid_type` and `mid2_type`
- `setConditionalFormat(...; :iconSet, iconset="Custom", reverse="true")` no longer reverses the caller's `icon_list` vector in place
- Errors that said only "Something wrong here!" now say what is wrong, e.g. which part of a malformed file is unexpected. A file part that fails to parse now reports the parser's own error alongside the part's name, and failing to save back to a source `IO` says why. Interrupting (Ctrl-C) while a file is read or saved is no longer reported as a parse or save error
- Rich-text cells now take `vertAlign` like any other font property: `setFont(...; vertAlign)` removes the runs' own superscript/subscript so the cell's shows, where it previously had no effect on those runs, and `setUniformFont(...; vertAlign)` no longer throws a `MethodError`
- `setUniformStyle` over rich-text cells no longer throws when the first cell's font has `vertAlign`, a fractional size, or font details that Excel writes but XLSX.jl can't set (such as the font family and scheme)
- Formatting a rich-text cell keeps the run properties XLSX.jl can't set but which files written by Excel or other tools may contain (such as a character set or outline). Every run property not on a fixed list was previously dropped
- `setUniformAlignment(ws, :, cols)` with a vector or step range of columns aligned every column of the sheet rather than only `cols`
- Errors for an operation on a file that isn't writable, or whose cell cache isn't enabled, now share one wording that names the operation and says how to open the file (`mode="rw"` or `enable_cache=true`)
- `copysheet!` gives the copy its own cells. The copy shared the original's cell cache, so a value, format, row height or formula set on either sheet also appeared in the other, and a formula set on the copy left the original with a cell that `getFormula` couldn't read
- `getdata` (and `ws[range]`) over a range of a writable sheet whose rows were written out of order could return wrong values, or write outside its result and corrupt memory, crashing Julia later in the garbage collector. The written rows are now sorted before the range is read
- `setUniformFont`, `setUniformBorder`, `setUniformFill`, `setUniformAlignment` and `setUniformFormat` given a whole row or column as `(ws, row, :)` or `(ws, :, col)` now apply one attribute to every cell, as the other target forms do. They previously applied the per-cell setter (`setFont` etc.), so cells that started with different formatting kept their differences
- Chart titles, axis titles and data labels set from a `RichTextString` keep each run's formatting (bold, italic, underline, strikethrough, superscript/subscript, size, font and color), with newlines handled as for plain strings. Previously this threw an `ArgumentError`. `addChart(...; title = rt)` keeps the run formatting on top of the new title's own
- On Julia 1.11+, the same chart text setters and `addChart`'s `title` keyword accept a StyledStrings `styled"…"` string, converted to a `RichTextString` as for cells (in the StyledStrings extension)
- `setChartTitleText` and `setAxisTitleText` accept a `CellRef` or `SheetCellRef`, binding the title to that cell (`c:tx/c:strRef`, with the cell's text cached). `addChart`'s `title` keyword accepts the same. A bound title has one style, so a typed title's formatting is moved into `c:txPr` and a rich-text cell lends the formatting its runs share; per-run formatting can't be kept, as Excel's file format doesn't allow it
- `ChartEx` titles: `setChartTitleText(c::ChartEx, …)` accepts a `RichTextString`, a styled string or a cell reference, and the new `setAxisTitleText(c::ChartEx, id, …)` sets an axis title's text the same ways. Run formatting is written to `cx:txPr`, which Excel renders; a newline now starts a new paragraph rather than being written inside one run. `addChartEx`'s `title` keyword accepts the same values
- `addChart(...; title = "A\nB")` now writes one paragraph per line, as `setChartTitleText` does, rather than a newline inside one run
- `addChart` and `addChartEx` resolve a cell-reference title before adding anything, so a bad reference leaves the workbook unchanged
- `getRichTextString` returns `nothing` for an empty cell, as documented, rather than throwing a `FieldError`
- Cell getters treat a cell outside the sheet's dimension as empty, as `ws[ref]` does, rather than throwing: `getFont`, `getBorder`, `getFill`, `getAlignment`, `getFormat`, `getRichTextString`, `getFormula` and `getMergedBaseCell` return `nothing`, and `isMergedCell` returns `false`. `getColumnWidth` and `getRowHeight` still throw, since a row or column beyond the dimension can't have a size
- Fixed a missing space in `setLabelText`'s error for a deleted label, and a typo in the `RichTextRun` docstring
- Title getters for both chart schemas: new `getAxisTitle` (plain text) and `getAxisTitleRange` (the resolved cell of a bound axis title) for `Chart` and `ChartEx`; `getAxisTitleRef` and `getChartTitleRef` now have `ChartEx` methods, and `getChartTitleRange` a `Chart` method
- `addSeries` and cell-bound titles now write elements in the chart part's own namespace prefixes rather than assuming Excel's `c:` and `a:`; a new chart's anchor likewise takes the drawing part's prefixes
- `addSeries(...; name_ref)` pointing at a number cell now writes the series name as a `c:strRef`; it wrote a `c:numRef`, which the schema doesn't allow there, and Excel refused to open the file
- Chart setters work on a chart part indented with whitespace between elements, as some tools other than Excel write; they previously threw a `MethodError`
- An empty `ChartEx` title, axis title or series name reads as `nothing`, as for a `c:` chart, rather than `""`
- `getAxisTitleTextProps` now has a method for `c:` charts, `getAxisTitleTextProps(c::Chart, ax::ChartAxis)`, reading the axis title's `c:txPr`. It previously existed only for `ChartEx`
- `SheetCellRef` is now exported
- `addChart` no longer writes `lang="en-US"` on a typed title's run
- Removed a duplicate `setAxisTitleText` from the `XLSX.Charts` export list
- New `setLegendPos` and `setLegendOverlay` for `Chart` and `ChartEx`: place the legend at Excel's five positions (`:b`, `:t`, `:l`, `:r`, `:tr`, or `:bottom`, `:top`, `:left`, `:right`, `:topRight`), drawn over the plot area or beside it. Both create the legend if the chart has none. On a `Chart`, setting a position removes any manual legend layout, as Excel does. `getLegendPos` and `getLegendOverlay` now have `ChartEx` methods, reading `cx:legend`'s `pos`, `align` and `overlay`
- Every chart color setter — `setSeriesFill`, the line and marker setters, and the text `:fill` and `:line` fields — accepts a theme color named as a String or Symbol (`"accent2"`, `:tx1`), as `addSeries(...; color)` already did. It is written as `a:schemeClr`, as `SchemeColor(:accent2)` is
- `setLabelTextProp(c::ChartEx, …)` accepts Excel's field names `:font`, `:color`, `:underline` and `:strikethrough`, as the other text-property setters do; it threw an error
- The text-property setters check `:under`, `:strike` and `:caps` values against the DrawingML vocabularies, and throw on any other word rather than writing a file Excel refuses to open. `:under` and `:strike` also take a `Bool`, and Excel's names `:single`/`:double` (and `:none` for `:strike`)
- Chart setter docstrings state the values each argument and keyword accepts. The color values, line keywords and text-property fields are each listed in full once — at `setSeriesFill`, `setSeriesLine` and `setChartTitleTextProp` — and the `ChartEx` single-property line setters and `addSeries(c::ChartEx, …)` have docstrings. The Charts API page gains a Conventions section
- Cell formatting (`setFont`, `setFill`, `setBorder`, their `Uniform` forms and `setOutsideBorder`), conditional formats (color scales, data bars, `setColoredDataBars` and the `font`/`fill`/`border` keywords) and `RichTextRun` colors accept every color form charts do: a hex or Colors.jl name as a String or Symbol, a Colors.jl `Colorant`, or a theme color (`"accent1"`, `:tx1` or a `SchemeColor`). A theme color is written as `<color theme="N" tint="t"/>`, so it follows the workbook's theme. Outside charts, a `SchemeColor` can be lightened or darkened (`lumMod`, or `lumMod` with `lumOff` adding to 100, as Excel's own palette does); other transforms throw. Existing String arguments behave as before. A new "Specifying colors" section in the formatting guide covers all of this
- `SchemeColor` is now defined in `XLSX` (and is `public` there), and is the same type as `XLSX.Charts.SchemeColor`
- A `RichTextRun` with an invalid `:color` now throws when it is created, not when it is written
- Unblock XML fix for #479 by updating Base.eachregion to StyledStrings.eachregion.
- The test suite no longer depends on Distributions.jl

</details>

## [v0.13.1](https://github.com/JuliaData/XLSX.jl/tree/v0.13.1) - 2026-10-07

- Fix [#462](https://github.com/JuliaData/XLSX.jl/issues/462) (Read-performance regression vs v0.10.4 on large worksheets)
- Fix [#472](https://github.com/JuliaData/XLSX.jl/issues/472) (Worksheets with CRLF line endings were corrupted when split, and precompilation failed, with XML.jl versions that normalise line endings)
- Faster reads of worksheets that have no `<dimension>` element. [#470](https://github.com/JuliaData/XLSX.jl/pull/470) found the missing dimension with an extra pass that built every cell; it is now recorded during the sheet's first read, so uncached `ws[:]` on a 20,000-row sheet takes 99 ms instead of 266 ms. Sheets with a dimension are unaffected
- Faster uncached reads (`enable_cache=false`). Each read now makes one pass over the worksheet and decodes only the cells it returns. Measured on a 50,000 × 70 sheet:
  - Ranges: `ws["A25000:C25010"]` takes 0.4 s instead of 2.4 s.
  - Non-contiguous selections: `ws[1:10000:50001, 1:10:70]` takes 0.9 s instead of 20 s, because it no longer re-reads the file for every cell.
  - `eachrow`: about twice as fast.
  - `eachtablerow` and `gettable`: 1.5× faster, or 2-2.5× faster over a few columns.
- Fix uncached `getcellrange` and `getcell` with several rows, which returned empty cells for every row after one absent from the file
- Fix uncached reads that lost formulas recorded in `wb.formulas`: the row iterator merged them only every 500 rows, so those after the last merge were dropped
- `find_row` on an uncached sheet now raises `XLSXError` for an absent row, rather than `ArgumentError`
- Uncached range, cell and table reads no longer raise an error for an invalid cell outside the cells they return, as `readtable` already did. Uncached reads assume rows are in ascending order, as the spec requires and Excel writes
- Numeric cell values are now parsed with [Parsers.jl](https://github.com/JuliaData/Parsers.jl) (new dependency, v2.8 or v3), giving results identical to `Base.parse`
- Fix `Tables.getcolumn(row, i)` on `eachtablerow` rows, which threw a `MethodError` for every integer index
- Column numbers and row ranges outside Excel's limits are now rejected (e.g. `CellRef(1, 20000)` no longer gives `]OF1`)
- An invalid reference passed to `setdata!` with a vector now raises `XLSXError` rather than `UndefVarError`
- Assigning an `AnnotatedString` whose colour names a face without its own foreground (e.g. `styled"{(foreground=highlight):x}"`) no longer throws a `MethodError`
- Faces whose foregrounds name each other in a cycle no longer cause a stack overflow when assigned as an `AnnotatedString`; the text is written without a colour
- Modest tidy-up of Charts docs
- Fix `setConditionalFormat` crash on files with a trailing newline or comment after `</worksheet>`
- Fix adding a dataBar or icon set to a reopened file that already has one
- Fix the first rule added to an existing range of a reopened file being lost
- Excel 2010 conditional formats now go in the correct `<ext>`
- Ids generated after reopening no longer duplicate existing ones
- Fix files Excel refused after `clearConditionalFormats` removed every dataBar
- `addtable!` rejects table names that are cell addresses (e.g. `T1`, `R1C1`)
- `addDefinedName` rejects R1C1-style names (e.g. `RC`, `R`)
- `writetable(...; as_table=true)` auto-names a table whose sheet name is a cell address
- `addtable!` with `has_totals_row=true` now errors if the last row has content
- `setSeriesSubtotals` and `setSeriesQuartileMethod` reject series of the wrong layout
- `addtable!` and `writetable!(...; as_table=true)` now raise an error if the table would overlap a merged range or another table, and `mergeCells` if the range would overlap a table. Excel used to discard such tables when it repaired the file on opening. `writetable!` checks before writing, so a refused table leaves the sheet unchanged

## [v0.13.0](https://github.com/JuliaData/XLSX.jl/tree/v0.13.0) - 2026-09-02

Chart support, read-only in v0.12, now covers formatting and creation, in a new
`XLSX.Charts` sub-module (`using XLSX, XLSX.Charts`). Chart support remains
experimental; see *Limitations* in the manual.

### Breaking changes
These may need changes to existing code:
- `getCharts` now also returns `chartEx` charts, as a `Vector{AbstractChart}`.
  Code that assumed every element is a `Chart` should check with `isa`.
- `getDefinedNames` and `getAllDefinedNames` exclude names Excel maintains for
  itself (hidden names and the `_xlnm.` built-ins); pass `include_system=true` for
  the old behaviour. `deleteDefinedName` and `deleteAllDefinedNames` refuse them
  unless `force=true`, and `addDefinedName` refuses to create them.
- `getConditionalFormats` returns `Vector{Pair{Union{CellRange,NonContiguousRange},...}}`
  rather than `Vector{Pair{CellRange,...}}`.

### Added
- **Charts**
  - Chart code moves to the `XLSX.Charts` sub-module. The names released in
    v0.12 remain reachable as `XLSX.getCharts` and so on.
  - Reading of `chartEx` charts: waterfall, funnel, treemap, sunburst, histogram,
    Pareto and box & whisker, region map returned as `ChartEx`.
  - Formatting: fills, lines, markers, data labels, text, axes, gridlines and
    legends can be read and set. Getters resolve the formatting cascade and report
    where each value came from.
  - Creation: `addChart` and `addSeries` for `c:` charts, `addChartEx` for chartEx
    charts, on a worksheet or a new chartsheet.
  - A seven-page charts guide in the manual.
- `deleteDefinedName` and `removeMergedCells`.
- Conditional formats on non-contiguous ranges (`"A1:A5,C1:C5"`, non-contiguous
  defined names, and vectors or step ranges of rows/columns), for every conditional
  format type.
- `setColoredDataBars` (experimental, not exported): data bars whose colour changes
  with value as well as their length.
- `clearConditionalFormats` (experimental, not exported): removes conditional
  formats lying entirely within a given range.
- `XLSXError` and `CellErrorType` (with `XL_NULL` … `XL_SPILL`) are now public.

### Changed
- `deletesheet!` invalidates chart references to the deleted sheet as Excel does:
  `#REF!` for a `c:` chart, which keeps its cached values, and the reference
  removed for a chartEx chart. `getChartRanges` reports such a reference as
  `XL_REF`.
- Assigning an `AnnotatedString` no longer changes the cell's font, and runs no
  longer force a 12pt size. A single-run string merges its formatting into the
  cell font, and unstyled text in a multi-run string takes the cell's font, so a
  template's formatting survives.
- A non-contiguous defined name now yields a single conditional format rule
  covering all its areas, rather than one rule per area.
- `DefinedName` gains a `hidden` field.

### Deprecated
- `chart.title`, `chart.charttypes` and `chart.series`: use `getChartTitle`,
  `getChartTypes` and `getChartSeries`.
- `chartType` and `chartSchema`: use `getChartType` and `getChartSchema`.

### Fixed
- System (Excel-private) defined names now survive a read/write round trip.
- `deletesheet!` no longer leaves chart parts orphaned in the package.
- `copysheet!` no longer shares chart parts with the original: charts are cloned
  and repointed at the copied sheet, for both `c:` and chartEx charts.
- Empty and `SubString` annotated strings are handled.
- Data bar `gradient`, `border` and axis settings were silently discarded by
  Excel: the `<extLst>` linking a rule to its Excel 2010 counterpart was written
  inside `<dataBar>` instead of `<cfRule>`.
- `setCfDataBar` cleared `max_val` when `min_type` was `"automatic"`, and failed
  to clear it when `max_type` was `"highest"`.
- A custom three-colour `colorScale` without an explicit `mid_col` threw an error;
  `mid_col` now defaults to white.

## [v0.12.4](https://github.com/JuliaData/XLSX.jl/tree/v0.12.4) - 2026-09-24
- Fixed a regression (since v0.11.0) where cells referencing an empty shared string were read as "" instead of missing.
- fix [#460](https://github.com/JuliaData/XLSX.jl/issues/460) (partial)
- fix [#454](https://github.com/JuliaData/XLSX.jl/issues/454) (retain number formats)
- When given range that resolves to a single cell, `setFormat`, `setFont`, `setFill`, `setBorder` and `setAlignment` now return the cell's attribute id, instead of -1.


## [v0.12.3](https://github.com/JuliaData/XLSX.jl/tree/v0.12.3) - 2026-08-20
- Address [#263](https://github.com/JuliaData/XLSX.jl/issues/263) (Retrieve data from chart)
- Add docs for Charts (WIP)
- Tidy `writetable` API and docs (non-breaking)
- Add show method for `XLSX.Workbook`
- Bump XML.jl compat to v0.4.6, fixing [#448](https://github.com/JuliaData/XLSX.jl/issues/448) (Thanks to @mathieu17g for a quick fix for this)
- Fix [#447](https://github.com/JuliaData/XLSX.jl/issues/447) - escaping inline strings correctly.
- Fix [#446](https://github.com/JuliaData/XLSX.jl/issues/446) (Fix sheet row stream iterator on resumption)
- Extend Excel-native Table methods and tests
- The `cache` keyword in several functions handling Excel charts (`getChart`, `getCharts`) is changed to `read_cached_values` which is more descriptive and less likely to create confusion with the cell cache. The impact of this should be very low as this functionality was only first added a couple of weeks ago and it was not in any release before this one (so technically not breaking).
  
## [v0.12.2](https://github.com/JuliaData/XLSX.jl/tree/v0.12.2) - 2026-07-31
- Add read and write support for native Excel tables
- Adopt XML.jl v0.4.4 for reduced allocations
- Improve test coverage

## [v0.12.1](https://github.com/JuliaData/XLSX.jl/tree/v0.12.1) - 2026-07-26
- Adopt [XML.splicetext](https://github.com/JuliaData/XML.jl/issues/92) from XML.jl v0.4.3 [#439](https://github.com/JuliaData/XLSX.jl/issues/439)
- Fix [#434](https://github.com/JuliaData/XLSX.jl/issues/434) to support freeze/split panes

## [v0.12.0](https://github.com/JuliaData/XLSX.jl/tree/v0.12.0) - 2026-07-12
- Adopt XML.jl v0.4
- Add a package extension to support [FileIO.jl](https://github.com/JuliaIO/FileIO.jl)
- Fix issues [#425](https://github.com/JuliaData/XLSX.jl/issues/425), [#426](https://github.com/JuliaData/XLSX.jl/issues/426), [#427](https://github.com/JuliaData/XLSX.jl/issues/427) and [#428](https://github.com/JuliaData/XLSX.jl/issues/428)
- Update copyright notice end-date
- Modularize tests
- Julia floor raised to 1.10 (LTS)
- Apply dependabot updates

Adoption of XML.jl v0.4 involves a major rewrite of package internals but provides a
significant speed improvement in most use cases. Thanks to @joshday and @mathieu17g for support with this.
Please [report any issues](https://github.com/JuliaData/XLSX.jl/issues).

[FileIO.jl](https://github.com/JuliaIO/FileIO.jl) support requires a [FileIO.jl](https://github.com/JuliaIO/FileIO.jl) version greater than v1.19.0 ([release pending](https://github.com/JuliaIO/FileIO.jl/pull/439)).

There are **no changes to the existing public API.**

## [v0.11.11](https://github.com/JuliaData/XLSX.jl/tree/v0.11.11) - 2026-06-18
- Fix [#410](https://github.com/JuliaData/XLSX.jl/issues/410) by making `is_binary_path` case insensitive

## [v0.11.10](https://github.com/JuliaData/XLSX.jl/tree/v0.11.10) - 2026-05-28
- support macro-enabled files ([#401](https://github.com/JuliaData/XLSX.jl/issues/401))
- support pass-through of customXml files (again). ([#403](https://github.com/JuliaData/XLSX.jl/issues/403))

## [v0.11.9](https://github.com/JuliaData/XLSX.jl/tree/v0.11.9) - 2026-05-26
- fix bug in `setFormula` when a function name occured inside a quoted string
- fix issue [#395](https://github.com/JuliaData/XLSX.jl/issues/395)(@mathieu17g)
- add tests for multi-threaded reading formulas and shared strings

## [v0.11.8](https://github.com/JuliaData/XLSX.jl/tree/v0.11.8) - 2026-05-15
- add `dependabot` support ([#130](https://github.com/JuliaData/XLSX.jl/issues/130))
- add ability to read native template (`.xltx`) files ([#293](https://github.com/JuliaData/XLSX.jl/issues/293))
- add `getDefinedNames` to mirror `addDefinedName`
- support adding png, jpg or gif images into a sheet ([#134](https://github.com/JuliaData/XLSX.jl/issues/134))

## [v0.11.7](https://github.com/JuliaData/XLSX.jl/tree/v0.11.7) - 2026-05-07
- Fix issue [#368](https://github.com/JuliaData/XLSX.jl/issues/368) for both reading and writing

## [v0.11.6](https://github.com/JuliaData/XLSX.jl/tree/v0.11.6) - 2026-04-29
- Correct citation details

## [v0.11.5](https://github.com/JuliaData/XLSX.jl/tree/v0.11.5) - 2026-04-29
- Add basic support for Chartsheets (read/write)
- Modestly widen support for strict format files
- Add hooks for Zenodo citation

## [v0.11.4](https://github.com/JuliaData/XLSX.jl/tree/v0.11.4) - 2026-04-22
- Add `iserror` and `geterror` functions ([#327](https://github.com/JuliaData/XLSX.jl/issues/327))
- Add support for strict (ISO/IEC 29500) .xlsx files ([#328](https://github.com/JuliaData/XLSX.jl/issues/328), [#233](https://github.com/JuliaData/XLSX.jl/issues/233))
- Widen Compat on Colors.jl to "0.12, 0.13" ([[#370](https://github.com/JuliaData/XLSX.jl/issues/370)])

## [v0.11.3](https://github.com/JuliaData/XLSX.jl/tree/v0.11.3) - 2026-04-14
- Bake in relocatable data (issue [#351](https://github.com/JuliaData/XLSX.jl/issues/351) @nhz2)
- Add type inference to `eachtablerow` (issue [#225](https://github.com/JuliaData/XLSX.jl/issues/225))
- Add support for custom missing strings (issue [#90](https://github.com/JuliaData/XLSX.jl/issues/90))
- Widen `setdata!` to accept an AbstractArray (issue [#158](https://github.com/JuliaData/XLSX.jl/issues/158))
- Fix intermittent bug with race condition when reading styles
- Move test data files to /test/data

## [v0.11.2](https://github.com/JuliaData/XLSX.jl/tree/v0.11.2) - 2026-04-01
Bug-fix to support 32-bit Julia (@davidanthoff)

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
