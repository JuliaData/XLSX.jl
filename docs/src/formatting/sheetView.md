# Zoom, gridlines and the view

[`XLSX.setSheetView`](@ref) sets how a worksheet looks when Excel opens it, as Excel's
*View* tab does: the zoom, whether gridlines, headings and zeros are shown, the
direction of the sheet, where it is scrolled to, and which cells are selected. These
change the workbook, so open the file with `mode="rw"` (or start from
[`XLSX.newxlsx`](@ref)). [`XLSX.getSheetView`](@ref) reads them from any file.

```jldoctest sheetview; setup = :(import XLSX)
julia> s = XLSX.newxlsx()[1];

julia> XLSX.setSheetView(s; zoom=150, showGridLines=false)

julia> v = XLSX.getSheetView(s);

julia> (v.zoom, v.showGridLines)
(150, false)
```

Only the settings given are changed. They are:

| Keyword | Excel | Default |
|:--------|:------|:--------|
| `zoom` | *View → Zoom*, 10 to 400 percent | `100` |
| `view` | *View → Normal / Page Layout / Page Break Preview*: `:normal`, `:pageLayout`, `:pageBreakPreview` | `:normal` |
| `showGridLines` | *View → Gridlines* | `true` |
| `showRowColHeaders` | *View → Headings*, the `1, 2, 3` and `A, B, C` | `true` |
| `showZeros` | *File → Options → Advanced → Show a zero in cells that have zero value* | `true` |
| `rightToLeft` | *Page Layout → Sheet Right-to-Left* | `false` |
| `topLeftCell` | the cell scrolled to the top left of the window | `A1` |
| `selection`, `activeCell` | the selected cells, and the one with the cursor | `A1` |

Excel keeps a zoom for each view, so setting `zoom` changes the zoom of the sheet's
current view, and switching view brings back the zoom that view had.

To open a sheet at a particular place, scroll to it and select the cells:

```jldoctest sheetview
julia> XLSX.setSheetView(s; topLeftCell="A20", selection="B22:D30", activeCell="C25")

julia> v = XLSX.getSheetView(s);

julia> (v.topLeftCell, v.activeCell, v.selection)
(A20, C25, "B22:D30")
```

A selection can be several ranges separated by commas (`"B2:D5,F7"`) or a defined name.
A `selection` given alone has its first cell active, and an `activeCell` given alone is
selected on its own.

## With frozen panes

With frozen or split panes ([Freeze and split panes](freezeAndSplitPanes.md)),
`topLeftCell` is the top-left cell of the part that scrolls, and the selection is the
active pane's, as Excel keeps them:

```julia
XLSX.freezePanes(sheet, "B3")                                # freeze rows 1-2 and column A
XLSX.setSheetView(sheet; topLeftCell="B40", activeCell="D45")  # scrolled down to row 40
```

With frozen panes, `topLeftCell` must be below and to the right of the frozen rows and
columns. Freezing counts from row 1 and column A, so [`XLSX.freezePanes`](@ref) scrolls
the view back to the top first.

## See also

- API: [Sheet view](@ref) in Sheet layout
- Related guides: [Freeze and split panes](freezeAndSplitPanes.md); [Printing](../sheets/printing.md), whose [`XLSX.setPrintOptions`](@ref) prints gridlines and headings
