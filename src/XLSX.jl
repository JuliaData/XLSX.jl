
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
using OrderedCollections: OrderedDict
import ZipArchives

import PrecompileTools as PCT    # this is a small dependency.

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
    addDefinedName, setFormula,
    # Formats
    setFormat, setFont, setBorder, setFill, setAlignment,
    setUniformFormat, setUniformFont, setUniformBorder, setUniformFill, setUniformAlignment, setUniformStyle,
    setConditionalFormat,
    RichTextString, RichTextRun,
    setColumnWidth, setRowHeight,
    getMergedCells, isMergedCell, getMergedBaseCell, mergeCells
  
const SPREADSHEET_NAMESPACE_XPATH_ARG = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
const EXCEL_MAX_COLS = 16_384     # total columns supported by Excel per sheet
const EXCEL_MAX_ROWS = 1_048_576  # total rows supported by Excel per sheet (including headers)
const ROW_CHUNKSIZE = 1000        # number of rows to be processed in each thread

include("types.jl")
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
include("conditional-format-helpers.jl") # must load before conditional-formats.jl
include("conditional-formats.jl")
include("images.jl")
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
        _ = XLSX.readtable(seekstart(s), 1, "A:Z")
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
