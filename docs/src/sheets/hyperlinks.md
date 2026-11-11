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
one. The cell keeps its value and formatting.

Adding a link to a cell that has one replaces it. Excel can also put one link on
a range of cells. XLSX.jl reads and deletes such links, but adds links to single
cells only.

!!! note "Links into the workbook"

    A `location` is stored as text. Renaming or deleting the sheet it names
    doesn't change it, so check links into a sheet after renaming it.

## See also

- API: [Images, comments and hyperlinks](../api/objects.md)
- Related guides: [Using Formulas](../formulae/formulas.md), for the `HYPERLINK` function; [Managing worksheets](worksheets.md)
