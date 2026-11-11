# Images

Adding or deleting images changes the workbook, so open the file with `mode="rw"` (or start
from [`XLSX.newxlsx`](@ref)) and save it afterwards. The functions that only read work on
any file.

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

[Adding a plot image](@ref) has a complete example with a Makie plot. To add a chart that
Excel draws from the cells, and that updates when they change, see
[Adding a simple native Excel chart](@ref) instead.

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

## See also

- API: [Images, comments and hyperlinks](../api/objects.md)
- Related guides: [Creating charts](../charts/creatingCharts.md), for charts Excel draws from the cells
- Examples: [Adding a plot image](@ref), [Adding a simple native Excel chart](@ref)
