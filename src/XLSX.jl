
module XLSX

import Base.convert
import Base.Threads
import Colors
import Dates
import Parsers
import Printf.@printf
import Random
import Tables
import Unicode
import UUIDs
import XML
import ZipArchives

import PrecompileTools as PCT    # this is a small dependency.

# ---------------------------------------------------------------------------
# Naming conventions
#
# Keywords: unqualified when the keyword names a property of the function's own
# subject (`addtable!(...; name, style)` -> the new Table; `setFont(...; name,
# size)` -> the font). Qualified only where a competing entity is in scope:
# `writetable(...; table_name, table_style)`, because `sheetname` and cell
# styles are also in that namespace. Qualify to disambiguate, not by default —
# `writetable`'s `totals` stays bare because nothing competes with it.
#
# Parameters that inject pre-computed state rather than describe the 
# subject (mergedCells, enable_cache, sheet_template_data) are named 
# for what they supply, not what they configure.
#
# Formatting keywords mirror the OOXML attribute name verbatim when it is
# already readable (`bgColor`, `vertAlign`, `wrapText`); single-letter OOXML
# names are expanded (`b` -> `bold`, `sz` -> `size`).
#
# Keywords borrowed wholesale from another package keep that package's
# spelling: `normalizenames` and `header` behave exactly as in CSV.jl, so they
# are spelled as in CSV.jl. Keywords that are merely inspired by one are ours
# and use underscores — `missing_strings` takes a set and defaults to nothing,
# unlike CSV.jl's `missingstring`. Grandfathered: `sheetname`, `overwrite`.
#
# Verbs: `delete*` destroys a first-class entity (`deletesheet!`,
# `deletetable!`); `remove*` clears a property of an entity that survives
# (`removetotals!`, `removePanes`). So a future image deletion is
# `deleteImage`, but unmerging cells is `removeMergedCells`.
#
# Function casing follows the family a function joins, not a global rule:
# `*table`/`*sheet`/`*data` are lowercase (`addtable!`, `getdata`), the
# formatting and feature layer is camelCase (`setFont`, `freezePanes`,
# `getCharts`). Match the neighbours.
#
# Chart naming and the chart export rule are documented in src/charts/Charts.jl.
# ---------------------------------------------------------------------------

export
    # Files and worksheets
    XLSXFile,
    readxlsx, openxlsx, opentemplate, newxlsx,
    writexlsx, savexlsx,
    Worksheet, sheetnames, sheetcount, hassheet,
    addsheet!, renamesheet!, copysheet!, deletesheet!, movesheet!,
    addImage,
    addComment, deleteComment,
    # Cells & data
    CellRef, SheetCellRef, row_number, column_number, eachtablerow,
    readdata, getdata, gettable, readtable, readto,
    iserror, geterror,
    gettransposedtable, readtransposedtable,
    writetable, writetable!,
    addDefinedName, deleteDefinedName, deleteAllDefinedNames,
    setFormula,
    # Formats
    setFormat, setFont, setBorder, setFill, setAlignment,
    setUniformFormat, setUniformFont, setUniformBorder, setUniformFill, setUniformAlignment, setUniformStyle,
    setConditionalFormat,
    RichTextString, RichTextRun,
    setColumnWidth, setRowHeight,
    getMergedCells, isMergedCell, getMergedBaseCell, mergeCells, removeMergedCells,
    freezePanes, splitFreeze, splitPanes, removePanes,
    # Excel Tables
    addtable!, deletetable!, settotals!, gettotals, removetotals!, appendtable!, gettablerange

@static if VERSION >= v"1.11"
    eval(Meta.parse("""
    public getcell, getcellrange, getFormula, getRichTextString,
           getConditionalFormats, getColumnWidth, getRowHeight,
           getFormat, getFont, getBorder, getFill, getAlignment,
           DataTable, Table, TableStyleInfo, table, tables,
           getDefinedNames, getAllDefinedNames,
           getComment, getComments,
           Workbook, XLSXError, SchemeColor,
           CellErrorType, XL_NULL, XL_DIV0, XL_VALUE, XL_REF, 
           XL_NAME, XL_NUM, XL_NA, XL_SPILL
    """))
end

const SPREADSHEET_NAMESPACE_XPATH_ARG = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"

const EXCEL_MAX_COLS = 16_384           # total columns supported by Excel per sheet
const EXCEL_MAX_ROWS = 1_048_576           # total rows supported by Excel per sheet (including headers)
const ROW_CHUNKSIZE = 1_000           # number of rows to be processed in each thread

include("scheme-color.jl") # must load before types.jl (`CellColor`)
include("types.jl")
include("xmlutil.jl")
include("xlsx-colors.jl") # must load before sst.jl and cellformat-helpers.jl
include("formula.jl")
include("cellref.jl")
include("sst.jl")
include("stream.jl")
include("valuerows.jl")
include("table.jl")
include("tables_interface.jl")
include("relationship.jl")
include("read.jl")
include("workbook.jl")
include("worksheet.jl")
include("cell.jl")
include("styles.jl")
include("cellformat-helpers.jl") # must load before cellformats/
include("cellformats/font.jl")
include("cellformats/border.jl")
include("cellformats/fill.jl")
include("cellformats/alignment.jl")
include("cellformats/numberformat.jl")
include("cellformats/uniformstyle.jl")
include("cellformats/widthheight.jl")
include("cellformats/mergedcells.jl")
include("panes.jl")
include("conditional-format-helpers.jl") # must load before conditionalformats/
include("conditionalformats/presets.jl")
include("conditionalformats/read.jl")
include("conditionalformats/setconditionalformat.jl")
include("conditionalformats/rules.jl")
include("conditionalformats/colorscale.jl")
include("conditionalformats/iconset.jl")
include("conditionalformats/databar.jl")
include("conditionalformats/clear.jl")
include("images.jl")
include("comments.jl")
include("write.jl")
include("fileArray.jl")

# Charts depends on the core and must load after it; the core never refers to Charts.
include("charts/Charts.jl")

# Names released as `XLSX.x` before the sub-module existed.
using .Charts: AbstractChart, Chart, ChartEx, ChartRef, ChartSeries,
    chartSchema, getChartSchema, chartType, getChartType,
    getCharts, getChart, getChartData, getChartRanges

@static if VERSION >= v"1.11"
    eval(Meta.parse("""
    public Charts,
            AbstractChart, Chart, ChartEx, ChartRef, ChartSeries,
            chartSchema, getChartSchema, chartType, getChartType, 
            getCharts, getChart, getChartData, getChartRanges
    """))
end

# Writes one small workbook holding each common cell type, then reads it back through
# every read path. Compile cost grows with the number of distinct methods and types
# run, not with the number of cells, so the data stays tiny. Paths are exercised
# with a file name (`String`) as well as an `IOBuffer`: the entry points specialize
# on each. Niche features (charts, images, conditional formats) are left out — they
# cost seconds of precompile time each, paid by every user.
# Users can turn this off with the PrecompileTools preference `precompile_workload = false`.
PCT.@setup_workload begin
    let
        dir = mktempdir()
        path = joinpath(dir, "precompile.xlsx")
        io = IOBuffer()
        nt = (i = [1, 2, 3], f = [1.5, 2.5, 3.5], s = ["a", "b", "c"], b = [true, false, true],
              d = [Dates.Date(2024, 1, 1), Dates.Date(2024, 1, 2), Dates.Date(2024, 1, 3)],
              dt = [Dates.DateTime(2024, 1, 1, 12), Dates.DateTime(2024, 1, 2, 12), Dates.DateTime(2024, 1, 3, 12)],
              m = Union{Missing,Int}[1, missing, 3])
        try
            PCT.@compile_workload begin
                # Writing
                writetable(path, nt; overwrite=true)
                writetable(path, "A" => nt, "B" => nt; overwrite=true)
                xf = openxlsx(path, mode="rw")
                sh = xf["A"]
                # Cell assignment compiles once per value type and reference form.
                vals = Any[1, 2.5, "x", true, Dates.Date(2024, 1, 1), Dates.DateTime(2024, 1, 1, 12), missing]
                for (c, v) in enumerate(vals)
                    sh[string("J", c)] = v
                    sh[CellRef(2, 9 + c)] = v
                    sh[3, 9 + c] = v
                end
                sh[5, :] = vals
                sh["J6:K7"] = "y"
                writetable!(addsheet!(xf, "C"), Any[nt.i, nt.s], [:i, :s])
                setFont(sh, "A1"; bold=true, color="red")
                setFill(sh, "A2"; pattern="solid", fgColor="yellow")
                setAlignment(sh, "A3"; horizontal="center", wrapText=true)
                setFormat(sh, "B2"; format="0.00")
                setBorder(sh, "A1:C3"; allsides=["style" => "thin", "color" => "black"])
                setBorder(sh, "A1:C3"; outside=["style" => "thick"])
                # The formatting setters take their keywords as one spec (see `FontSpec`),
                # so each call here covers every set of keywords that setter is given.
                setUniformFont(sh, "A2:C3"; italic=true)
                setUniformFill(sh, "A2:C3"; pattern="solid", fgColor="yellow")
                setUniformAlignment(sh, "A2:C3"; horizontal="center")
                setUniformFormat(sh, "B2:B3"; format="0.00")
                setUniformBorder(sh, "A2:C3"; bottom=["style" => "thin"])
                setColumnWidth(sh, 1; width=12)
                mergeCells(sh, "J6:K7")
                savexlsx(xf)
                writexlsx(io, xf)
                openxlsx(IOBuffer(), mode="w") do xf
                    xf[1]["A1:B2"] = 1.5
                end

                # Reading
                dt = readtable(path, 1)
                Tables.columns(dt)
                Tables.schema(dt)
                readtable(path, "A", "A:C")
                readtable(seekstart(io), 1)
                xf = readxlsx(path)
                sheetnames(xf)
                sh = xf["A"]
                sh["A1:C4"]
                getdata(sh)
                sh[2, 1]
                getcell(sh, "B2")
                openxlsx(path) do xf
                    gettable(xf[1])
                    for r in eachtablerow(xf[1])
                        r[1]
                    end
                end
                openxlsx(path; enable_cache=false) do xf
                    for r in eachrow(xf[1])
                        r[1]
                    end
                end
                readdata(path, 1, "A1:C3")
                readdata(path, "A!A1:B2")

                # Excel Tables
                writetable(path, nt; overwrite=true, as_table=true, table_name="T")
                readtable(path, 1; table_name="T")
            end
        finally
            rm(dir; recursive=true, force=true)
        end
    end
end

end # module XLSX

