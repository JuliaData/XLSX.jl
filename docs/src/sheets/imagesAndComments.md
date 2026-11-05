# Images, comments and sheet order

This page covers three things you can do to a worksheet beyond its cell values:
adding pictures, attaching comments to cells, and arranging the sheets in a
workbook. Each function used here changes the workbook, so open the file with
`mode="rw"` (or start from [`XLSX.newxlsx`](@ref)) and save it afterwards.

## Images

[`XLSX.addImage`](@ref) places a PNG, JPEG or GIF image on a worksheet. The image
floats above the grid, as a picture inserted in Excel does. It does not change
any cell, and it can be moved and resized in Excel as usual.

The image's top-left corner is anchored to a cell. With a single cell, the image
keeps its own pixel size, or the size you give:

```julia
XLSX.openxlsx("report.xlsx", mode="rw") do xf
    sheet = xf["Summary"]
    XLSX.addImage(sheet, "B2", "logo.png")                     # native size
    XLSX.addImage(sheet, "H2", "photo.jpg"; size=(200, 150))   # width, height in pixels
end
```

With a cell range, the image is stretched to fill it and `size` is ignored:

```julia
XLSX.addImage(sheet, "B10:F20", "chart.png")
```

The image can also come from an `IOBuffer`, for example a plot rendered in
memory. The buffer is read without being consumed:

```julia
io = IOBuffer()
# ... write PNG bytes to io, e.g. with a plotting package ...
XLSX.addImage(sheet, "D2:H12", io)
```

An image's size snaps to whole cells. The calculation assumes Excel's default
cell size of 64 × 20 pixels, so an image over resized columns or rows is
stretched or squeezed to fit the cells it spans.

[`XLSX.getImages`](@ref) lists the images on a sheet, or in the whole workbook.
Each one is a `NamedTuple` giving the sheet, the internal media name and the
cells at its top-left and bottom-right corners:

```julia
julia> XLSX.getImages(sheet)
2-element Vector{@NamedTuple{sheet::String, media_name::String, from::String, to::String}}:
 (sheet = "Summary", media_name = "image1.png", from = "B2", to = "D5")
 (sheet = "Summary", media_name = "image2.png", from = "H2", to = "J9")
```

Images can't yet be deleted. They are kept when a workbook is read and saved
again.

## Comments

A comment, called a *Note* in recent Excel versions, is text attached to one
cell. The cell shows a small red triangle, and the comment appears when the
mouse is over it. [`XLSX.addComment`](@ref) adds one, or replaces the cell's
existing comment:

```julia
XLSX.addComment(sheet, "B3", "Checked against the 2024 accounts"; author="Finance")
XLSX.addComment(sheet, 5, 2, "Line one\nLine two"; visible=true)   # always shown
```

`width` and `height` set the size of the pop-up box in pixels.

To read comments back, use [`XLSX.getComment`](@ref) for one cell or
[`XLSX.getComments`](@ref) for a sheet or workbook. Each comment is a
`NamedTuple` of `(sheet, ref, author, text)`, and a cell with no comment gives
`nothing`:

```julia
julia> XLSX.getComment(sheet, "B3")
(sheet = "Summary", ref = "B3", author = "Finance", text = "Checked against the 2024 accounts")

julia> XLSX.getComment(sheet, "C3") === nothing
true
```

Notes that Excel writes usually begin with the author's name (`"Name:\n…"`),
because Excel puts it in the note text itself. `getComment` returns the text as
stored.

[`XLSX.deleteComment`](@ref) removes a comment. It returns `true` if there was
one to remove:

```julia
XLSX.deleteComment(sheet, "B3")   # true
```

!!! note "Threaded comments"

    Excel 365 also has *threaded comments*, which hold a conversation.
    XLSX.jl does not write them. For a cell that has one, `getComment` returns the
    plain note Excel stores alongside it. Replacing or deleting a threaded
    comment throws an error rather than leaving the two out of step.

## Sheet order

The functions [`XLSX.addsheet!`](@ref), [`XLSX.copysheet!`](@ref),
[`XLSX.renamesheet!`](@ref) and [`XLSX.deletesheet!`](@ref) manage the sheets in a
workbook. [`XLSX.movesheet!`](@ref) changes their order. Give the sheet to move,
and either `before` or `after` to say where it goes. Either sheet can be named
by its name, its tab position or the `Worksheet` itself:

```julia
julia> xf = XLSX.newxlsx("A"); XLSX.addsheet!(xf, "B"); XLSX.addsheet!(xf, "C");

julia> XLSX.movesheet!(xf, "C"; before="A");

julia> XLSX.sheetnames(xf)
3-element Vector{String}:
 "C"
 "A"
 "B"

julia> XLSX.movesheet!(xf[1]; after=XLSX.sheetcount(xf));   # move to the end
```

Moving a sheet doesn't rewrite anything else. Formulas, charts and defined names
refer to sheets by name, and names scoped to a sheet move with it. The one thing
that changes is a 3D reference such as `=SUM(Jan:Mar!A1)`. As in Excel, it covers
whatever sheets lie between `Jan` and `Mar` in tab order.

See [Files and worksheets](../api/files.md) in the API reference for the full
signatures.
