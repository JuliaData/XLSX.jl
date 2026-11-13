# Supported features

This page lists the Excel features XLSX.jl works with, the functions for each, and
where each is documented. The second table lists features XLSX.jl doesn't create or
edit, so you needn't search for them.

Functions that add, change or remove something need a file opened with `mode="rw"`,
or one started with [`XLSX.newxlsx`](@ref). Functions that only read work on any file.

## What XLSX.jl can do

### Data

| Excel feature | In Excel | Read | Create or change | Remove | Guide |
|:--------------|:---------|:-----|:-----------------|:-------|:------|
| Cell values | typing in cells | `sheet["A1"]`, [`XLSX.getdata`](@ref), [`XLSX.readdata`](@ref) | `sheet["A1"] = value` | — | [XLSX Tutorial](tutorial/XLSXtutorial.md) |
| Tabular data | — | [`XLSX.readtable`](@ref), [`XLSX.gettable`](@ref), [`XLSX.eachtablerow`](@ref) | [`XLSX.writetable`](@ref), [`XLSX.writetable!`](@ref) | — | [XLSX Tutorial](tutorial/XLSXtutorial.md) |
| Writing large files in constant memory | — | — | [`XLSX.writetable`](@ref) with `streaming=true`, [`XLSX.streamxlsx`](@ref), [`XLSX.writerow!`](@ref) | — | [Writing large files](@ref) |
| Excel Tables | *Insert → Table* | [`XLSX.tables`](@ref), [`XLSX.table`](@ref), [`XLSX.gettotals`](@ref) | [`XLSX.addtable!`](@ref), [`XLSX.appendtable!`](@ref), [`XLSX.settotals!`](@ref) | [`XLSX.deletetable!`](@ref), [`XLSX.removetotals!`](@ref) | [Excel Tables](tables/excelTables.md) |
| Formulas, with their cached values | formula bar | [`XLSX.getFormula`](@ref), `sheet["A1"]` for the value | [`XLSX.setFormula`](@ref), with `value` | — | [Using Formulas](formulae/formulas.md) |
| Named ranges | *Formulas → Name Manager* | [`XLSX.getDefinedNames`](@ref), [`XLSX.getAllDefinedNames`](@ref) | [`XLSX.addDefinedName`](@ref) | [`XLSX.deleteDefinedName`](@ref), [`XLSX.deleteAllDefinedNames`](@ref) | [Named ranges](formulae/namedRanges.md) |
| Data validation and drop-downs | *Data → Data Validation* | [`XLSX.getDataValidations`](@ref), [`XLSX.getDataValidation`](@ref) | [`XLSX.addDataValidation`](@ref) | [`XLSX.deleteDataValidation`](@ref) | [Data validation](sheets/dataValidation.md) |
| AutoFilter buttons, filter criteria and sorting | *Data → Filter*, *Data → Sort* | [`XLSX.getAutoFilter`](@ref) | [`XLSX.setAutoFilter`](@ref), with `filters` and `sort` | [`XLSX.removeAutoFilter`](@ref) | [Filtering](sheets/autoFilter.md) |
| Error values (`#N/A`, `#DIV/0!`, …) | — | [`XLSX.iserror`](@ref), [`XLSX.geterror`](@ref) | — | — | [Error Values](@ref) (API) |
| Ignoring error-checking warnings | green triangle → *Ignore Error* | [`XLSX.getIgnoredErrors`](@ref) | [`XLSX.setIgnoredErrors`](@ref) | [`XLSX.removeIgnoredErrors`](@ref) | [Ignoring error-checking warnings](@ref) |

### Formatting

| Excel feature | In Excel | Read | Create or change | Remove | Guide |
|:--------------|:---------|:-----|:-----------------|:-------|:------|
| Fonts, fills, borders, alignment | *Home → Font / Alignment* | [`XLSX.getFont`](@ref), [`XLSX.getFill`](@ref), [`XLSX.getBorder`](@ref), [`XLSX.getAlignment`](@ref) | [`XLSX.setFont`](@ref), [`XLSX.setFill`](@ref), [`XLSX.setBorder`](@ref), [`XLSX.setAlignment`](@ref) and their `setUniform…` forms, [`XLSX.setUniformStyle`](@ref) | — | [Cell formats](formatting/cellFormatting.md) |
| Number formats | *Home → Number* | [`XLSX.getFormat`](@ref), [`XLSX.getText`](@ref) (the text shown) | [`XLSX.setFormat`](@ref), [`XLSX.setUniformFormat`](@ref) | — | [Number formats](@ref) |
| Rich text (mixed formats in one cell) | selecting part of the text | [`XLSX.getRichTextString`](@ref) | [`XLSX.RichTextString`](@ref), StyledStrings.jl | — | [Rich text formats](@ref) |
| Conditional formatting | *Home → Conditional Formatting* | [`XLSX.getConditionalFormats`](@ref) | [`XLSX.setConditionalFormat`](@ref), [`XLSX.setColoredDataBars`](@ref) (experimental) | [`XLSX.clearConditionalFormats`](@ref) | [Conditional formats](formatting/conditionalFormatting.md) |
| Checkboxes in cells (Excel 365) | *Insert → Checkbox* | [`XLSX.getCheckboxes`](@ref), [`XLSX.hasCheckbox`](@ref) | [`XLSX.setCheckbox`](@ref) | [`XLSX.removeCheckbox`](@ref) | [Checkboxes](@ref) |
| Merged cells | *Home → Merge & Center* | [`XLSX.getMergedCells`](@ref), [`XLSX.isMergedCell`](@ref), [`XLSX.getMergedBaseCell`](@ref) | [`XLSX.mergeCells`](@ref) | [`XLSX.removeMergedCells`](@ref) | [Merged cells](formatting/mergedCells.md) |
| Column widths and row heights | *Home → Format* | [`XLSX.getColumnWidth`](@ref), [`XLSX.getRowHeight`](@ref) | [`XLSX.setColumnWidth`](@ref), [`XLSX.setRowHeight`](@ref), [`XLSX.autofitColumns`](@ref) | — | [Column width and row height](formatting/widthAndHeight.md) |

### Sheet layout and view

| Excel feature | In Excel | Read | Create or change | Remove | Guide |
|:--------------|:---------|:-----|:-----------------|:-------|:------|
| Hidden rows and columns | *Home → Format → Hide & Unhide* | [`XLSX.isRowHidden`](@ref), [`XLSX.isColumnHidden`](@ref) | [`XLSX.setRowHidden`](@ref), [`XLSX.setColumnHidden`](@ref) | the same, with `hidden=false` | [Hiding and grouping](formatting/hideAndGroup.md) |
| Grouping and outlines | *Data → Group* | [`XLSX.getRowOutlineLevel`](@ref), [`XLSX.getColumnOutlineLevel`](@ref) | [`XLSX.groupRows`](@ref), [`XLSX.groupColumns`](@ref) | [`XLSX.ungroupRows`](@ref), [`XLSX.ungroupColumns`](@ref) | [Hiding and grouping](formatting/hideAndGroup.md) |
| Freeze and split panes | *View → Freeze Panes / Split* | — | [`XLSX.freezePanes`](@ref), [`XLSX.splitPanes`](@ref), [`XLSX.splitFreeze`](@ref) | [`XLSX.removePanes`](@ref) | [Freeze and split panes](formatting/freezeAndSplitPanes.md) |
| Zoom, gridlines, headings, zeros, right-to-left, scroll position and selection | *View* | [`XLSX.getSheetView`](@ref) | [`XLSX.setSheetView`](@ref) | — | [Zoom, gridlines and the view](formatting/sheetView.md) |
| Inserting and deleting rows and columns | *Home → Insert / Delete* | — | [`XLSX.insertRows`](@ref), [`XLSX.insertColumns`](@ref) | [`XLSX.deleteRows`](@ref), [`XLSX.deleteColumns`](@ref) | [Inserting and deleting rows and columns](sheets/insertDelete.md) |

### Charts, images, links and comments

| Excel feature | In Excel | Read | Create or change | Remove | Guide |
|:--------------|:---------|:-----|:-----------------|:-------|:------|
| Charts (experimental) | *Insert → Charts* | [`getCharts`](@ref XLSX.Charts.getCharts), [`getChartData`](@ref XLSX.Charts.getChartData) and many more | [`addChart`](@ref XLSX.Charts.addChart), [`addChartEx`](@ref XLSX.Charts.addChartEx), [`addSeries`](@ref XLSX.Charts.addSeries), [`makeCombo`](@ref XLSX.Charts.makeCombo), the `set…` functions | [`deleteChart`](@ref XLSX.Charts.deleteChart) | [Charts](charts/excelCharts.md) |
| Chartsheets | *Move Chart → New sheet* | [`getCharts`](@ref XLSX.Charts.getCharts) | [`addChart`](@ref XLSX.Charts.addChart) given the file | [`XLSX.deletesheet!`](@ref) | [Sheets with charts](charts/chartSheets.md) |
| Sparklines | *Insert → Sparklines* | [`XLSX.getSparklines`](@ref) | [`XLSX.addSparklines`](@ref) | [`XLSX.deleteSparklines`](@ref) | [Sparklines](charts/sparklines.md) |
| Images (PNG, JPEG, GIF) | *Insert → Pictures* | [`XLSX.getImages`](@ref) | [`XLSX.addImage`](@ref) | [`XLSX.deleteImage`](@ref) | [Images](@ref) |
| Pictures in cells (Excel 365) | *Insert → Pictures → Place in Cell* | [`XLSX.getCellImage`](@ref), [`XLSX.getCellImages`](@ref) | [`XLSX.CellImage`](@ref) written to a cell, [`XLSX.addImage`](@ref) with `inCell=true` | [`XLSX.deleteImage`](@ref) | [Pictures in cells](@ref) |
| Text boxes and shapes | *Insert → Text Box*, *Insert → Shapes* | [`XLSX.getShapes`](@ref) | [`XLSX.addTextBox`](@ref), [`XLSX.addShape`](@ref) | [`XLSX.deleteShape`](@ref) | [Text boxes and shapes](sheets/shapes.md) |
| Hyperlinks | *Insert → Link* | [`XLSX.getHyperlink`](@ref), [`XLSX.getHyperlinks`](@ref) | [`XLSX.addHyperlink`](@ref) | [`XLSX.deleteHyperlink`](@ref) | [Hyperlinks](@ref) |
| Comments (notes) | *Review → Notes* | [`XLSX.getComment`](@ref), [`XLSX.getComments`](@ref) | [`XLSX.addComment`](@ref) | [`XLSX.deleteComment`](@ref) | [Comments](@ref) |
| Threaded comments | *Review → New Comment* | [`XLSX.getThreadedComment`](@ref), [`XLSX.getThreadedComments`](@ref) | [`XLSX.addThreadedComment`](@ref), [`XLSX.addThreadedReply`](@ref), [`XLSX.setThreadedCommentResolved`](@ref) | [`XLSX.deleteThreadedComment`](@ref) | [Threaded comments](@ref) |

### Worksheets, protection and printing

| Excel feature | In Excel | Read | Create or change | Remove | Guide |
|:--------------|:---------|:-----|:-----------------|:-------|:------|
| Document properties | *File → Info → Properties* | [`XLSX.getDocumentProperties`](@ref), [`XLSX.getCustomProperty`](@ref), [`XLSX.getCustomProperties`](@ref) | [`XLSX.setDocumentProperties`](@ref), [`XLSX.setCustomProperty`](@ref) | [`XLSX.deleteCustomProperty`](@ref) | [Document properties](@ref) |
| Worksheets | sheet tabs | [`XLSX.sheetnames`](@ref), [`XLSX.sheetcount`](@ref), [`XLSX.hassheet`](@ref) | [`XLSX.addsheet!`](@ref), [`XLSX.copysheet!`](@ref), [`XLSX.renamesheet!`](@ref), [`XLSX.movesheet!`](@ref) | [`XLSX.deletesheet!`](@ref) | [Managing worksheets](sheets/worksheets.md) |
| Sheet visibility | right-click a tab → *Hide* | [`XLSX.getSheetVisibility`](@ref) | [`XLSX.setSheetVisibility`](@ref) | — | [Hiding a sheet](@ref) |
| Tab colour | right-click a tab → *Tab Color* | [`XLSX.getTabColor`](@ref) | [`XLSX.setTabColor`](@ref) | [`XLSX.removeTabColor`](@ref) | [Tab colour](@ref) |
| Sheet protection | *Review → Protect Sheet* | [`XLSX.isSheetProtected`](@ref), [`XLSX.getSheetProtection`](@ref), [`XLSX.getCellProtection`](@ref) | [`XLSX.protectSheet`](@ref), [`XLSX.setCellProtection`](@ref) | [`XLSX.unprotectSheet`](@ref) | [Protecting a sheet](@ref) |
| Workbook protection | *Review → Protect Workbook* | [`XLSX.isWorkbookProtected`](@ref), [`XLSX.getWorkbookProtection`](@ref) | [`XLSX.protectWorkbook`](@ref) | [`XLSX.unprotectWorkbook`](@ref) | [Protecting the workbook](@ref) |
| Page setup and printing | *Page Layout* | [`XLSX.getPageSetup`](@ref), [`XLSX.getPageMargins`](@ref), [`XLSX.getPrintOptions`](@ref), [`XLSX.getPrintArea`](@ref), [`XLSX.getPrintTitles`](@ref) | [`XLSX.setPageSetup`](@ref), [`XLSX.setPageMargins`](@ref), [`XLSX.setPrintOptions`](@ref), [`XLSX.setPrintArea`](@ref), [`XLSX.setPrintTitles`](@ref) | [`XLSX.removePrintArea`](@ref), [`XLSX.removePrintTitles`](@ref) | [Printing](@ref) |
| Headers and footers | *Page Layout → Page Setup → Header/Footer* | [`XLSX.getHeaderFooter`](@ref) | [`XLSX.setHeaderFooter`](@ref), [`XLSX.headerFooterField`](@ref), [`XLSX.headerFooterFont`](@ref) | [`XLSX.removeHeaderFooter`](@ref) | [Headers and footers](@ref) |

## What XLSX.jl doesn't do

| Feature | Notes |
|:--------|:------|
| Calculating formulas | [`XLSX.setFormula`](@ref) writes the formula, and Excel recalculates it when the file is opened. Its value reads as `missing` until then, unless you give the value yourself: see [Cached values](@ref). |
| Filtering by colour or by icon | [`XLSX.setAutoFilter`](@ref) filters by values, conditions, top or bottom items, the average and dates. |
| Creating region map charts, and some chart subtypes | See the charts [Limitations](charts/chartLimitations.md). |
| Pivot tables and slicers | No functions to create or edit them. |
| Shapes filled with pictures, connectors, grouped shapes and SmartArt | [`XLSX.addShape`](@ref) adds single preset shapes with text, fill and border. |
| Pictures in headers and footers | [`XLSX.setHeaderFooter`](@ref) sets their text, fields and fonts only. |
| Password-encrypted files | To read an encrypted file, decrypt it first with [XLSXDecrypt.jl](https://github.com/TimG1964/XLSXDecrypt.jl). |
| VBA macros | Not created or edited. |

Opening a file that holds any of these in `mode="rw"` and saving it over the
original deserves care: see the warning on [`XLSX.openxlsx`](@ref).
