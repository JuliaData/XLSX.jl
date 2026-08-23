
module XLSX

import Base.convert
import Base.Threads
import Colors
import Dates
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
# ---------------------------------------------------------------------------

export
    # Files and worksheets
    XLSXFile,
    readxlsx, openxlsx, opentemplate, newxlsx,
    writexlsx, savexlsx,
    Worksheet, sheetnames, sheetcount, hassheet, 
    addsheet!, renamesheet!, copysheet!, deletesheet!, 
    addImage,
    # Cells & data
    CellRef, row_number, column_number, eachtablerow,
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
    getMergedCells, isMergedCell, getMergedBaseCell, mergeCells,
    freezePanes, splitFreeze, splitPanes, removePanes,
    # Excel Tables
    addtable!, deletetable!, settotals!, gettotals, removetotals!, appendtable!

@static if VERSION >= v"1.11"
    eval(Meta.parse("""
    public getcell, getcellrange, getFormula, getRichTextString,
           getConditionalFormats, getColumnWidth, getRowHeight,
           getFormat, getFont, getBorder, getFill, getAlignment,
           DataTable, Table, TableStyleInfo, table, tables,
           getCharts, getChart, getChartData, Chart, ChartSeries, ChartRef,
           getDefinedNames, getAllDefinedNames,
           Workbook
    """))
end

const SPREADSHEET_NAMESPACE_XPATH_ARG = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"

const EXCEL_MAX_COLS =    16_384           # total columns supported by Excel per sheet
const EXCEL_MAX_ROWS = 1_048_576           # total rows supported by Excel per sheet (including headers)
const ROW_CHUNKSIZE  =     1_000           # number of rows to be processed in each thread

include("types.jl")
include("xmlutil.jl")
include("xlsx-colors.jl") # must load before sst.jl and cellformat-helpers.jl
include("formula.jl")
include("cellref.jl")
include("sst.jl")
include("stream.jl")
include("table.jl")
include("tables_interface.jl")
include("relationship.jl")
include("read.jl")
include("workbook.jl")
include("worksheet.jl")
include("cell.jl")
include("styles.jl")
include("cellformat-helpers.jl") # must load before cellformats.jl
include("cellformats.jl")
include("panes.jl")
include("conditional-format-helpers.jl") # must load before conditional-formats.jl
include("conditional-formats.jl")
include("images.jl")
include("charts.jl")
include("write.jl")
include("fileArray.jl")

PCT.@setup_workload begin
    # Putting some things in `@setup_workload` instead of `@compile_workload` can reduce the size of the
    # precompile file and potentially make loading faster.
    s=IOBuffer()
    t=IOBuffer()
    PCT.@compile_workload begin
        # all calls in this block will be precompiled, regardless of whether
        # they belong to your package or not (on Julia 1.8 and higher)
        f=openxlsx(joinpath(@__DIR__, "data", "blank.xlsx"), mode="rw")
        f[1]["A1:Z26"] = "hello World"
        openxlsx(s, mode="w") do xf
            xf[1][1:26, 1:26] = pi
        end
        _ = readtable(seekstart(s), 1, "A:Z")
        f= openxlsx(seekstart(s), mode="rw")
        f[1][1:26, 1:26] = pi
        setConditionalFormat(f[1], :, :cellIs)
        setConditionalFormat(f[1], "A1:Z26", :colorScale)
        setBorder(f[1], collect(1:26), 1:26, allsides=["style"=>"thin", "color"=>"black"])
        _ = getdata(f[1], "A1:A20")
        writexlsx(t, f)
    end
end

end # module XLSX
