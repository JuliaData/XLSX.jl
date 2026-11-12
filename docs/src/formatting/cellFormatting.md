
# Cell formats

## Excel formatting

Each cell in an Excel spreadsheet may refer to an Excel `style`. Multiple cells can 
refer to the same `style` and therefore have a uniform appearance. A `style` defines
the cell's `alignment` directly (as part of the `style` definition), but it may also 
refer to further formatting definitions for `font`, `fill`, `border` and number `format`. 
Multiple `style`s may each refer to the same `fill` definition or the same `font` 
definition, etc, and therefore share these formatting characteristics.
This hierarchy can be shown like this:
```
                `Cell`
                  │
               `Style` => `Alignment`
                  │
  ┌──────────┬────┴─────┬─────────┐
  │          │          │         │
`font`     `fill`    `border`  `format`
```
A family of setter functions is provided to set each of the formatting characteristics 
Excel uses. These are applied to cells, and the functions deal with the relationships 
between the individual characteristics, the overarching `style` and the cell(s) themselves.

## Setting format attributes of a cell

Set the font attributes of a cell using [`XLSX.setFont`](@ref). For example, to set cells `A1` and 
`A5` in the `general` sheet of a workbook to specific `font` values, use:

```julia

julia> using XLSX

julia> f=XLSX.opentemplate("general.xlsx")
XLSXFile("general.xlsx") containing 13 Worksheets
            sheetname size          range        
-------------------------------------------------
              general 10x6          A1:F10       
               table3 5x6           A2:F6        
               table4 4x3           E12:G15
                table 12x8          A2:H13
               table2 5x3           A1:C5
                empty 1x1           A1:A1
               table5 6x1           C3:C8
               table6 8x2           B1:C8
               table7 7x2           B2:C8
               lookup 4x9           B2:J5
         header_error 3x4           B2:E4
       named_ranges_2 4x5           A1:E4
         named_ranges 14x6          A2:F15

julia> s=f["general"]
10×6 Worksheet: ["general"](A1:F10)

julia> XLSX.setFont(s, "A1"; name="Arial", size=24, color="blue", bold=true)
2

julia> XLSX.setFont(s, "A5"; name="Arial", size=24, color="blue", bold=true)
2
```

The function returns the `fontId` that has been used to define this combination 
of attributes.

There are more `font` attributes that can be set. Setting attributes for a cell 
that already has some, merges the new attributes with the old. Thus:

```julia
julia> XLSX.setFont(s, "A5"; italic=true, under="double", bold=false)
3
```

will over-ride the `bold` setting that was previously defined and add a double 
underline and make the font italic. However, the color, font name and size will 
all remain unchanged from before. This new combination of attributes is unique, 
so a new `fontId` has been created.

The other set attribute functions behave in similar ways. See [`XLSX.setBorder`](@ref), 
[`XLSX.setFill`](@ref), [`XLSX.setFormat`](@ref) and [`XLSX.setAlignment`](@ref).

## Specifying colors

Anywhere a color is set — `setFont`, `setFill`, `setBorder` and their `Uniform` forms,
conditional formats, rich text runs and charts — these forms are accepted:

```julia
julia> XLSX.setFont(s, "A1"; color="FF0000")                       # hex RGB, "#RRGGBB" or AARRGGBB

julia> XLSX.setFont(s, "A1"; color="coral")                        # any Colors.jl named color

julia> XLSX.setFont(s, "A1"; color=:coral)                         # the same name as a symbol

julia> XLSX.setFont(s, "A1"; color=RGB(0.2, 0.4, 0.6))             # a Colors.jl value

julia> XLSX.setFont(s, "A1"; color="accent2")                      # a theme color, by name

julia> XLSX.setFont(s, "A1"; color=XLSX.SchemeColor(:accent2; lumMod=75))   # a theme color, darker
```

Named colors are any of those provided by [Colors.jl](https://juliagraphics.github.io/Colors.jl/stable/namedcolors/).
A Colors.jl value is useful for colors computed in code, such as a palette from
`distinguishable_colors` or `range(colorant"red", colorant"blue"; length=5)`.

Excel ignores the alpha (transparency) part of a cell color, so an `AARRGGBB` value or a
Colors.jl value with alpha is accepted but always appears opaque. Charts are different:
there, alpha is honored.

A theme color refers to a slot in the workbook's theme rather than to a fixed RGB value,
so it changes when the theme does, as colors picked from Excel's theme palette do. Name
it as a string or symbol — `"bg1"`, `"tx1"`, `"bg2"`, `"tx2"`, `"accent1"` … `"accent6"`,
`"hlink"` or `"folHlink"` — or as an [`XLSX.SchemeColor`](@ref) to make it lighter or
darker. A cell color can only hold the lighter and darker shades Excel's own palette
offers:

```julia
julia> XLSX.SchemeColor(:accent1)                          # Accent 1

julia> XLSX.SchemeColor(:accent1; lumMod=75)               # Accent 1, Darker 25%

julia> XLSX.SchemeColor(:accent1; lumMod=60, lumOff=40)    # Accent 1, Lighter 40%
```

That is, `lumMod` alone (darker) or `lumMod` and `lumOff` adding to 100 (lighter). Other
`SchemeColor` transforms, such as `shade` or `satMod`, can only be used in charts and are
an error elsewhere.

When a theme color is read back, `getFont`, `getFill` and `getBorder` report it as
`"theme"` and `"tint"` attributes (e.g. `Dict("theme" => "4", "tint" => "0.4")`).
[`XLSX.getRichTextString`](@ref) reports every color as its `"FFRRGGBB"` value in the
current theme.

## Number formats

[`XLSX.setFormat`](@ref) sets how a cell's number, date or time is shown, as the *Number*
tab of Excel's *Format Cells* dialog does. The value itself doesn't change: a cell
holding `0.256` shown as `26%` still reads as `0.256`.

```julia
XLSX.setFormat(s, "A1"; format="Percentage")               # 26%
XLSX.setFormat(s, "A2"; format="Currency")                 # in the local currency of whoever opens the file
XLSX.setFormat(s, "A3"; format="LongDate")                 # 9-Mar-25
XLSX.setFormat(s, "A4"; format="yyyy-mm-dd hh:mm")         # 2025-03-09 14:30
XLSX.setFormat(s, "A5"; format="0.0%")                     # 25.6%
XLSX.setFormat(s, "A6"; format="#,##0.00;[Red]-#,##0.00")  # negative numbers in red
```

`format` is one of these names, the number of one of Excel's built-in formats (as a
string, such as `"10"` for `0.00%`, listed under [`XLSX.setFormat`](@ref)), or a format
code of your own, written as in the *Custom* category of the Format Cells dialog.

| Name | Format code |
|:-----|:------------|
| `"General"` | `General` |
| `"Number"` | `0.00` |
| `"Currency"` | `$#,##0.00_);($#,##0.00)` |
| `"Percentage"` | `0%` |
| `"ShortDate"` | `m/d/yyyy` |
| `"LongDate"` | `d-mmm-yy` |
| `"Time"` | `h:mm:ss` |
| `"Scientific"` | `##0.0E+0` |

A `Date`, `DateTime` or `Time` written to a cell is given a date or time format
automatically (Excel's built-in `m/d/yyyy`, `m/d/yyyy h:mm` and `h:mm`). Excel shows
those built-in formats in the date style of the computer opening the file. Use a format
code of your own, such as `"dd/mm/yyyy"`, `"yyyy-mm-dd"` or `"d mmmm yyyy"`, for a fixed
style.

[`XLSX.getFormat`](@ref) reads a cell's format back, and [`XLSX.setUniformFormat`](@ref)
applies one format to a range faster (see [Setting uniform attributes](@ref)).

### The text a cell shows

[`XLSX.getText`](@ref) gives the text Excel shows in a cell: its value in its number
format, as it appears on screen. It works on a file opened only for reading too.

```jldoctest celltext; setup = :(import XLSX; using Dates)
julia> s = XLSX.newxlsx()[1];

julia> s["A1"] = 1234.5; XLSX.setFormat(s, "A1"; format="#,##0.00");

julia> s["A2"] = Date(2026, 3, 7); XLSX.setFormat(s, "A2"; format="d mmmm yyyy");

julia> s["A3"] = 0.256; XLSX.setFormat(s, "A3"; format="0.0%");

julia> [XLSX.getText(s, "A\$i") for i in 1:3]
3-element Vector{String}:
 "1,234.50"
 "7 March 2026"
 "25.6%"
```

The text matches Excel 365's for every part of a format code: General, digit
placeholders, separators, percentages, scientific notation, fractions, dates, times,
text sections, literals, currency codes, colours and conditions. Excel shows its
built-in date formats in the date style of the computer opening the file; `getText`
shows them as Excel does in US English (`3/7/2026`).

## Formatting multiple cells at once

[Addressing cells and ranges](../concepts/addressing.md) lists every way of naming the cells to
format. This section shows how formats combine when they are applied to many cells.

### Applying `setAttribute` to multiple cells

Each of the setter functions can be applied to multiple cells at once using cell-ranges, 
row- or column-ranges or non-contiguous ranges. Additionally, indexing can use integer
indices for rows and columns, vectors of index values, unit- or step-ranges. This makes 
it easy to apply formatting to many cells at once.

Thus, for example:

```julia

julia> using XLSX

julia> f=XLSX.newxlsx()
XLSXFile("C:\...\blank.xlsx") containing 1 Worksheet
            sheetname size          range        
-------------------------------------------------
               Sheet1 1x1           A1:A1

julia> s=f[1]
1×1 Worksheet: ["Sheet1"](A1:A1)

julia> s[1:100, 1:100] = "" # Ensure these aren't `EmptyCell`s.
""

julia> XLSX.setFont(s, "A1:CV100"; name="Arial", size=24, color="blue", bold=true)
-1                          # Returns -1 on a range.

julia> XLSX.setBorder(s, "A1:CV100"; allsides = ["style" => "thin", "color" => "black"])
-1

julia> XLSX.setAlignment(s, [10, 50, 90], 1:100; wrapText=true) # Wrap text in the specified rows.
-1

julia>  XLSX.setAlignment(s, 1:100, 2:2:100; rotation=90) # Rotate text 90° every second column in the first 100 rows.
-1
```

It is even possible to use defined names to index these functions:

```julia

julia> XLSX.addDefinedName(s, "my_name", "A1,B20,C30") # Define a non-contiguous named range.
XLSX.DefinedNameValue(Sheet1!A1,Sheet1!B20,Sheet1!C30, Bool[1, 1, 1])

julia> XLSX.setFill(s, "my_name"; pattern="solid", fgColor="coral")
-1
```

When setting format attributes over a range of cells as decribed, the new attributes are merged 
with existing on a cell by cell basis. If you set the font name on a range of cells that previously 
all had different font colors, the color differences will persist even as the font name is applied 
to the range consistently.

### Setting uniform attributes

Sometimes it is useful to be able to apply a fully consistent set of format attributes to a range of 
cells, over-riding any pre-existing differences. This is the purpose of the `setUniformAttribute` 
family of functions. These functions update the attributes of the first cell in the range and then 
apply the relevant attribute Id to the rest of the cells in the range. Thus:

```julia
julia> XLSX.setUniformBorder(s, "A1:CV100"; allsides = ["color" => "green"], diagonal = ["direction"=>"both", "color"=>"red"])
2 # This is the `borderId` that has now been uniformly applied to every cell.
```

This sets the border color in cell `A1` to be green and adds red diagonal lines across the cell. 
It then applies all the `Border` attributes of cell `A1` uniformly to all the other cells in the range, 
overriding their previous attributes.

All the format setter functions have `setUniformAttribute` versions, too. See [`XLSX.setUniformBorder`](@ref), 
[`XLSX.setUniformFill`](@ref), [`XLSX.setUniformFormat`](@ref) and [`XLSX.setUniformAlignment`](@ref).

### Setting uniform styles

It is possible to use each of the `setUniformAttribute` functions in turn to ensure every possible 
attribute is consistently applied to a range of cells. However, if perfect uniformity is required, 
then `setUniformStyle` is considerably more efficient. It will simply take the `styleId` of the 
first cell in the range and apply it uniformly to each cell in the range. This ensures that all 
of font, fill, border, format, and alignment are all completely consistent across the range:

```julia
julia> XLSX.setUniformStyle(s, "A1:CV100") # set all formatting attributes to be uniformly the same as cell A1.
7    # this is the `styleId` that has now been applied to all cells in the range
```

### Illustrating the different approaches

To illustrate the differences between applying `setAttribute`, `setUniformAttribute` and `setUinformStyle`,
consider the following worksheet, which has very hetrogeneous formatting across the three cells:

![Three cells, B2, D2 and F2, with different fonts, fills and borders](../images/multicell.png)

We can apply `setBorder()` to add a top border to each cell:

```julia
julia> XLSX.setBorder(s, "B2,D2,F2"; top=["style"=>"thick", "color"=>"red"])
-1
```
This merges the new top border definition with the other, existing border attributes, to get

![The three cells after setBorder: each keeps its own borders and gains a thick red top border](../images/multicell2.png)

Alternatively, we can apply `setUniformBorder()`, which will update the borders of cell `B2` 
and then apply all the border attributes of `B2` to the other cells, overwriting the previous 
settings:

```julia
julia> XLSX.setUniformBorder(s, "B2,D2,F2"; top=["style"=>"thick", "color"=>"red"])
4
```

This makes the border formatting entirely consistent across the cells but leaves the other formatting 
attributes (font, fill, format, alignment) as they were.

![The three cells after setUniformBorder: all have B2's borders, with a thick red top, and keep their other formats](../images/multicell3.png)

Finally, we can set `B2` to have the formatting we want, and then apply a uniform style to all three cells.

```julia
julia> XLSX.setBorder(s, "B2"; top=["style"=>"thick", "color"=>"red"])
4

julia> XLSX.setUniformStyle(s, "B2,D2,F2")
19
```
Which results in all formatting attributes being entirely consistent across the cells.

![The three cells after setUniformStyle: all are formatted exactly like B2](../images/multicell4.png)

### Performance differences between methods

To illustrtate the relative performance of these three methods, applied to a million cells:
```julia
using XLSX
function setup()
    f = XLSX.newxlsx()
    s = f[1]
    s[1:1000, 1:1000] = pi
    return f
end
do_format(f) = XLSX.setFormat(f[1], 1:1000, 1:1000; format="0.0000")
do_uniform_format(f) = XLSX.setUniformFormat(f[1], 1:1000, 1:1000; format="0.0000")
function do_format_styles(f)
    XLSX.setFormat(f[1], "A1"; format="0.0000")
    XLSX.setUniformStyle(f[1], 1:1000, 1:1000)
end
function timeit()
    f = setup()
    do_format(f)
    do_uniform_format(f)
    do_format_styles(f)
    f = setup()
    print("Using `setFormat`        : ")
    @time do_format(f)
    f = setup()
    print("Using `setUniformFormat` : ")
    @time do_uniform_format(f)
    f = setup()
    print("Using `setUniformStyle` : ")
    @time do_format_styles(f)
    return f
end
f=timeit()
```

which yields the following timings:

```
Using `setFormat`        :  10.966803 seconds (256.00 M allocations: 19.771 GiB, 18.81% gc time)
Using `setUniformFormat` :   2.222868 seconds (31.00 M allocations: 1.137 GiB, 19.48% gc time)
Using `setUniformStyles` :   0.519658 seconds (14.00 M allocations: 416.587 MiB)
```

The same test, using the more involved `setBorder` function

```julia
do_format(f) = XLSX.setBorder(f[1], 1:1000, 1:1000;
        left     = ["style" => "dotted", "color" => "FF000FF0"],
        right    = ["style" => "medium", "color" => "firebrick2"],
        top      = ["style" => "thick",  "color" => "FF230000"],
        bottom   = ["style" => "medium", "color" => "goldenrod3"],
        diagonal = ["style" => "dotted", "color" => "FF00D4D4", "direction" => "both"]
    )
```

gives

```
Using `setBorder`        :  29.536010 seconds (759.00 M allocations: 64.286 GiB, 22.01% gc time)
Using `setUniformBorder` :   2.052018 seconds (31.00 M allocations: 1.197 GiB, 13.18% gc time)
Using `setUniformStyles` :   0.599491 seconds (14.00 M allocations: 416.586 MiB, 15.20% gc time)
```

If maintaining heterogeneous formatting attributes is not important, it is more efficient to 
apply `setUinformAttribute` functions rather than `setAttribute` functions, especially on large 
cell ranges, and more efficient still to use `setUniformStyle`.

## Copying formatting attributes

It is possible to use non-contiguous ranges to copy format attributes from any cell to any other cells, 
whether you are also updating the source cell's format or not.

```julia
julia> XLSX.setBorder(s, "BB50"; allsides = ["style" => "medium", "color" => "yellow"])
3 # Cell BB50 now has the border format I want!

julia> XLSX.setUniformBorder(s, "BB50,A1:CV100") # Make cell BB50 the first (reference) cell in a non-contiguous range.
3

julia> XLSX.setUniformStyle(s, "BB50,A1:CV100") # Or if I want to apply all formatting attributes from BB50 to the range.
11
```

## Rich text formats

Text in a cell can have font formats applied to any substring element as in this example:

![Cells whose text mixes colours, sizes, bold, italic, underline, superscript and subscript](../images/complexFontFormats.png)

This is referred to here as rich text formatting.

These rich text formats are stored separately from the cell Style and over-ride any Style-based formatting.

### Creating a rich text cell value

It is possible to create a rich text value in a cell in two ways: using an `XLSX.RichTextString` or using a `StyledString` from StyledStrings.jl. If the latter, `StyledString`s are translated to `RichTextStrings` before being written to cells.

#### Using RichTextStrings

A rich text string includes formatting that changes for different substring elements of the string. Each substring element is described as a `RichTextRun`, and multiple `RichTextRuns` together form a `RichTextString`. A `RichTextRun` can describe any or all of the following font attributes: `:bold`, `:italic`, `:under`, `:strike`, `:vertAlign`, `:color`, `:size`, `:name`. These attributes have the same effect as described in [`XLSX.setFont`](@ref), but, unlike `setFont`, they do not apply to the whole cell, they only apply to the current rich text run substring.

Thus,

```
julia> rtf1=XLSX.RichTextRun("Hello", [:color => "red", :size => 24])
RichTextRun ("Hello"  [:color => "red", :size => 24])

julia> rtf2=XLSX.RichTextRun("o", [:color => "green", :size => 24, :vertAlign => "superscript"])
RichTextRun ("o"  [:color => "green", :size => 24, :vertAlign => "superscript"])

julia> rtf3=XLSX.RichTextRun(" Kitt", [:color => "blue", :size => 14])
RichTextRun (" Kitt"  [:color => "blue", :size => 14])

julia> rtf4=XLSX.RichTextRun("y", [:color => "green", :size => 14, :vertAlign => "subscript"])
RichTextRun ("y"  [:color => "green", :size => 14, :vertAlign => "subscript"])

julia> rt = XLSX.RichTextString(rtf1, rtf2, rtf3, rtf4)
RichTextString: "Helloo Kitty"
 containing 4 runs:
 Run text                 Run attributes
 -------------------------------------------------------------------------------------------
 "Hello"                  [:color => "red", :size => 24]
 "o"                      [:color => "green", :size => 24, :vertAlign => "superscript"]
 " Kitt"                  [:color => "blue", :size => 14]
 "y"                      [:color => "green", :size => 14, :vertAlign => "subscript"]

julia> s["A1"] = rt
RichTextString: "Helloo Kitty"
 containing 4 runs:
 Run text                 Run attributes
 -------------------------------------------------------------------------------------------
 "Hello"                  [:color => "red", :size => 24]
 "o"                      [:color => "green", :size => 24, :vertAlign => "superscript"]
 " Kitt"                  [:color => "blue", :size => 14]
 "y"                      [:color => "green", :size => 14, :vertAlign => "subscript"]       
```
![A cell showing Helloo Kitty in rich text: red, green superscript, blue and green subscript runs](../images/richTextString.png)

A `RichTextString` with only a single run will be converted to a simple text string with font attributes set through `setFont`.

For more information, refer to [`XLSX.RichTextString`](@ref).

#### Using StyledStrings

An XLSX.jl package extension provides support for `AnnotatedStrings` from StyledStrings.jl when that package is being used in the current environment. For example:
```
julia> using StyledStrings, XLSX

julia> f = newxlsx()
XLSXFile("blank.xlsx") containing 1 Worksheet
            sheetname size          range
-------------------------------------------------
               Sheet1 1x1           A1:A1

julia> s = f[1]
1×1 Worksheet: ["Sheet1"](A1:A1) 

julia> s["A1"] = styled"{yellow:hello} {blue:there}"
"hello there"

julia> s["A2"] = styled"The {bold:{italic:quick {(foreground=#cd853f):brown} fox} jumps over the {(foreground=#FFC000):lazy} dog}"
"The quick brown fox jumps over the lazy dog"

julia> XLSX.getRichTextString(s, "A1")
RichTextString: "hello there"
 containing 3 runs:
 Run text                 Run attributes
 -------------------------------------------------------------------------------------------
 "hello"                  [:color => "FFE5A509", :size => 12.0]
 " "                      [:size => 12.0]
 "there"                  [:color => "FF195EB3", :size => 12.0]


julia> XLSX.getRichTextString(s, "A2")
RichTextString: "The quick brown fox jumps over the lazy dog"
 containing 7 runs:
 Run text                 Run attributes
 -------------------------------------------------------------------------------------------
 "The "                   [:size => 12.0]
 "quick "                 [:bold => true, :italic => true, :size => 12.0]
 "brown"                  [:bold => true, :color => "FFCD853F", :italic => true, :size => …]
 " fox"                   [:bold => true, :italic => true, :size => 12.0]
 " jumps over the "       [:bold => true, :size => 12.0]
 "lazy"                   [:bold => true, :color => "FFFFC000", :size => 12.0]
 " dog"                   [:bold => true, :size => 12.0]

```

![Two cells written from StyledStrings: hello there in yellow and blue, and a sentence with bold, italic and coloured words](../images/styledString.png)

When an `AnnotatedString` is written to a cell, the cell's font is first reset
to the workbook default, so any font applied to the cell beforehand does not
leak into the new value. Other cell formatting — alignment, borders, fills,
number format — is left untouched.

Each styled region of the string then becomes a rich text run. A run's
attributes come from its `StyledStrings` face, resolved against the default
face, so named faces and their `inherit` chains are followed and any
customisations in the user's `faces.toml` are applied.

Attributes matching the default face are omitted from the run. A run that sets
no font name or size therefore takes them from the cell's font, which is the
workbook default until `setFont` is called on the cell afterwards.

For more information on the use of styled strings, refer to the documentation for the [StyledStrings.jl](https://github.com/JuliaLang/StyledStrings.jl) package.

!!! note

    The use of `AnnotatedStrings` requires Julia v1.11 or higher


### Updating a rich text cell value

The `setFont`, `setUniformFont` and `setUniformStyle` functions operate at the cell level and cannot apply 
formatting at a substring level. Instead, using any of these functions will remove from a cell's rich text format 
any attributes that these functions are themselves applying or updating on the whole cell, but will leave the remaining 
rich text format attributes in place.

For example, to remove the underlining in column B from the "Hello Kitty" example, above:

```julia
julia> setFont(s, "B"; under="none")
-1
```
![The rich text cells after setFont with under=none: the underlining is gone and the other run formats remain](../images/complexNoUnder.png)

The underlining has been removed, but the other rich text formats remain in place.

Applying `setUniformFont` to a range of cells will remove those elements of the rich text format that are now being applied 
to the whole cell so that those elements become uniform. Any other heterogeneity arising from rich text formats 
(which over-ride Style-based formatting) will remain.

For example:
```julia
julia> setUniformFont(s, "A1:F2"; color="orange")
23
```
![The rich text cells after setUniformFont with an orange colour: all text is orange, other run formats remain](../images/complexUniformColor.png)

Similarly, applying `setUniformStyle` to a range of cells will remove those elements of the rich text format that are now being applied to the whole cells so that those elements become uniform. Any other heterogeneity arising from rich text formats will remain.

Rich text formats only affect the font settings; other elements of the style (number format, borders, fill, alignment) will be unaffected and will be applied uniformly to the range.

To illustrate, starting from the initial Excel file above, and using `setUniformStyle` to apply the Style in cell `A1` to the cells `D1:F2`:
```julia
julia> setFont(s, "A1"; size=20) # make sure cell A1 has the formatting we want
22

julia> XLSX.getFont(s, "A1") # just to illustrate the cell A1 Style's font attributes
XLSX.CellFont(22, Dict{String, Union{Nothing, Dict{String, String}}}("name" => Dict("val" => "Aptos Narrow"), "sz" => Dict("val" => "20"), "color" => Dict("theme" => "1")), "1")

julia> setUniformStyle(s, "A1,D1:F2") # apply cell A1's Style to cells D1:F2
1
```
![The rich text cells after setUniformStyle from A1: font, size and colour are uniform, bold, italic and sub- and superscript remain](../images/complexUniformStyle.png)

The Style for cell `A1` has a specified font, size and color. These have been uniformly applied to the range `D1:F2`, but the other rich text attributes (in this case, bold, italic, underline, super- and sub-script) retain the original heterogeneity.

If the first cell in the specified range has no defined Style (`s=""`), all cells will be given the same undefined style and any rich text formatting will remain unchanged.

In Excel, subscript and superscript formatting to a substring (for example) can only be made using this kind of rich text formatting, and cannot be set at cell level (neither in Excel nor in `setFont`). Retaining rich text format heterogeneity in the ways described here preserves these string formatting elements through format changes.

To clear any rich text formatting and remove the heterogeneity it introduces, simply copy the cells over themselves.
This strips out all rich text formatting information from the cells, leaving only the cell content and Style-base formating.
For example, for arbitrary cells:

```julia
julia> for row=1:2, col=1:6; s[row,col] = s[row,col]; end
```
or, for tabular data:

```julia
julia> writetable!(s, gettable(s, "A:F"; header=false); write_columnnames=false)
```
![The cells after being copied over themselves: plain text with no rich text formatting](../images/complexNone.png)

An `XLSXFile` must be opened in write mode for rich text formatting to be editable, otherwise an error is thrown.

There is no way in XLSX.jl to edit the individual runs of a rich text value. To make changes to individual runs beyond what `setFont` can achieve, it is neseccary to overwrite the cell with a new `RichTextString` or `AnnotatedString`.

## See also

- API: [Cell formats](../api/formats.md)
- Related guides: [Conditional formats](conditionalFormatting.md), [Merged cells](mergedCells.md), [Column width and row height](widthAndHeight.md)
- Examples: [Applying cell format to an existing table](@ref), [Creating a formatted form](@ref)
