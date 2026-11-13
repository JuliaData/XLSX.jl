```@meta
DocTestSetup = :(import XLSX)
```

# Hyperlinks

Adding or deleting links changes the workbook, so open the file with `mode="rw"` (or start
from [`XLSX.newxlsx`](@ref)) and save it afterwards. The functions that only read work on
any file.

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

```jldoctest links; setup = :(import XLSX; xf = XLSX.newxlsx("Summary"); XLSX.addsheet!(xf, "Data"); sheet = xf["Summary"]; XLSX.addHyperlink(sheet, "A3"; location="Data!B2", display="See the data"))
julia> link = XLSX.getHyperlink(sheet, "A3");

julia> link.location, link.display, link.url
("Data!B2", "See the data", nothing)
```

[`XLSX.deleteHyperlink`](@ref) removes a link and returns `true` if there was
one. The cell keeps its value and formatting. Given a range, it removes every
link that overlaps it.

Adding a link to a cell that has one replaces it.

## A link on a range of cells

A link can cover a range, so that clicking any of its cells follows it. As Excel
does, the text goes in the range's top-left cell, if that is empty, and every
empty cell in the range gets the link font:

```julia
XLSX.addHyperlink(sheet, "B5:D5"; url="https://example.com/report", display="Full report")
```

[`XLSX.getHyperlink`](@ref) returns the range's link for any of its cells, with
the range as its `ref`. A link may not overlap another: delete the other one
first.

!!! note "Links into the workbook"

    A `location` is stored as text. Renaming or deleting the sheet it names
    doesn't change it, so check links into a sheet after renaming it.

## See also

- API: [Images, comments and hyperlinks](../api/objects.md)
- Related guides: [Using Formulas](../formulae/formulas.md), for the `HYPERLINK` function; [Managing worksheets](worksheets.md)
