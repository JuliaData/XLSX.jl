# Links, images, comments and sheet order

This page covers things you can do to a worksheet beyond its cell values:
linking cells to web pages or other places in the workbook, adding pictures,
attaching comments to cells, and arranging the sheets in a workbook. Each
function that adds or deletes something changes the workbook, so open the file
with `mode="rw"` (or start from [`XLSX.newxlsx`](@ref)) and save it afterwards.
The functions that only read work on any file.

## Hyperlinks

[`XLSX.addHyperlink`](@ref) makes a cell a link. Give `url` for an address
outside the workbook (a web page, a `mailto:` address or a file), or `location`
for a place in the workbook (a cell or a defined name):

```julia
XLSX.openxlsx("report.xlsx", mode="rw") do xf
    sheet = xf["Summary"]
    XLSX.addHyperlink(sheet, "A1"; url="https://julialang.org", display="Julia")
    XLSX.addHyperlink(sheet, "A2"; url="mailto:team@example.com", tooltip="Email the team")
    XLSX.addHyperlink(sheet, "A3"; location="Data!B2", display="See the data")
end
```

When the cell is empty, as here, its value becomes `display` (or the address if
there is no `display`) and it is given a blue, underlined font, as Excel does. A
cell that already holds a value keeps it and its formatting, so a link can be
added to existing text. `tooltip` is shown when the mouse is over the link.

[`XLSX.getHyperlink`](@ref) reads the link on one cell and
[`XLSX.getHyperlinks`](@ref) every link on a sheet or in a workbook. Each link
is a `NamedTuple` of `(sheet, ref, url, location, display, tooltip)`, with
`nothing` for the fields it doesn't use:

```julia
julia> XLSX.getHyperlink(sheet, "A3")
(sheet = "Summary", ref = "A3", url = nothing, location = "Data!B2", display = "See the data", tooltip = nothing)
```

[`XLSX.deleteHyperlink`](@ref) removes a link and returns `true` if there was
one. The cell keeps its value and formatting.

Adding a link to a cell that has one replaces it. Excel can also put one link on
a range of cells. XLSX.jl reads and deletes such links, but adds links to single
cells only.

!!! note "Links into the workbook"

    A `location` is stored as text. Renaming or deleting the sheet it names
    doesn't change it, so check links into a sheet after renaming it.

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

[`XLSX.deleteImage`](@ref) deletes the images whose top-left corner is in a
cell, or one image from `getImages`, and returns how many it deleted. The
picture file goes too, unless another image in the workbook shows it:

```julia
XLSX.deleteImage(sheet, "B2")                                # 1
foreach(img -> XLSX.deleteImage(sheet, img), XLSX.getImages(sheet))   # every image
```

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

### Threaded comments

Excel 365 also has *threaded comments*, which hold a conversation: a comment and
its replies, each with its author and date, and the thread can be marked
resolved. [`XLSX.getThreadedComments`](@ref) and
[`XLSX.getThreadedComment`](@ref) read them:

```julia
julia> t = XLSX.getThreadedComment(sheet, "B3");

julia> t.author, t.text, t.resolved
("Ana", "Is this right?", false)

julia> [(r.author, r.text) for r in t.replies]
1-element Vector{Tuple{String, String}}:
 ("Ben", "Yes, checked.")
```

XLSX.jl reads threaded comments but doesn't write them. For a cell that has one,
`getComment` returns the plain note Excel stores alongside it for older versions
of Excel. Replacing or deleting a threaded comment throws an error rather than
leaving the two out of step.

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
