```@meta
DocTestSetup = :(import XLSX)
```

# Comments

Adding or deleting comments changes the workbook, so open the file with `mode="rw"` (or start
from [`XLSX.newxlsx`](@ref)) and save it afterwards. The functions that only read work on
any file.

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

```jldoctest comments; setup = :(import XLSX; xf = XLSX.newxlsx("Summary"); sheet = xf["Summary"]; XLSX.addComment(sheet, "B3", "Checked against the 2024 accounts"; author="Finance"))
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

## Threaded comments

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

## See also

- API: [Images, comments and hyperlinks](../api/objects.md)
- Related guides: [Inserting and deleting rows and columns](insertDelete.md), which moves comments with their cells
