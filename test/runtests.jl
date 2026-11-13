import XLSX
import Tables
using Test, Dates, XML
using OrderedCollections: OrderedDict
import DataFrames, Random
import CSV
using StyledStrings
using ZipArchives
using Logging
using Colors

# If SAVE_FILES is true then every testset that creates an XLSXFile in memory
# or that writes an xlsx file will have a file created in `outdir`.
# It may be possible to create files that XLSX.jl can handle but Excel, which 
# has stricter rules, rejects. By creating the output of every testset as a 
# file explicitly, it is possible to confirm that this never happens in practice 
# by manually opening each output file.
# If TRUE, approx 700 files are created in `outdir`.
# This flag should be reserved for local use and **never** left as TRUE for CI.
const SAVE_FILES = false

const src_data_directory = joinpath(dirname(pathof(XLSX)), "data")
const data_directory = joinpath(dirname(pathof(XLSX)), "..", "test", "data")
const files_directory = joinpath(dirname(pathof(XLSX)), "..", "test", "test_files")

@assert isdir(src_data_directory)
@assert isdir(data_directory)
@assert isdir(files_directory)

const outdir = joinpath(data_directory, "output_files")

if SAVE_FILES
    if !isdir(outdir)
        mkdir(outdir)
    else
        for f in readdir(outdir; join=true)
            try
                isfile(f) && rm(f; force=true)
            catch e
                @warn "Could not remove old output file $f" exception=e
            end
        end
    end
end

# Tests write their scratch files to the working directory. Run them in a fresh
# temporary directory, removed when the run ends, so nothing is left in the
# repository whether a test passes or fails. Fixtures are always read through
# `data_directory`/`src_data_directory`, which are absolute.
const test_run_directory = pwd()
cd(mktempdir())

#=
The tests are grouped by theme, one folder of `test_files` per theme, run from the
lowest level to the highest. Each theme prints one summary table listing its files;
a file's own testsets are shown only when one of them fails (or with `verbose=true`).
The separation is not perfect, and need not be: any function can be used with any
other.

Every summary table's title is padded to `NAME_WIDTH`, so the Pass, Total and Time
columns line up across the whole run. Julia aligns the rows within a table itself, so
nested names only need to fit within `NAME_WIDTH - 2`. Add a new test file to the
right theme's list below rather than including it on its own.

`include` evaluates each file at global scope, so the constants and helper functions a
file defines work just as they would if it were included directly.
=#
const NAME_WIDTH = 26

function testgroup(theme::AbstractString, files::Vector{Pair{String,String}})
    @testset "$(rpad(theme, NAME_WIDTH))" verbose=true begin
        for (label, path) in files
            @assert textwidth(label) <= NAME_WIDTH - 2 "test label \"$label\" is too long"
            @testset "$label" begin
                include(joinpath(files_directory, path))
            end
        end
    end
end

include(joinpath(files_directory, "helpers.jl"))  # shared helpers first

testgroup("XML", [
    "XML helpers"             => "xml/xml.jl",
    "namespaces"              => "xml/namespaces.jl",
    "Strict OOXML"            => "xml/strict.jl",
])
testgroup("Files", [
    "fixture files"           => "files/fixtures.jl",
    "Excel-made files"        => "files/excel-made.jl",
    "file modes"              => "files/modes.jl",
    "FileIO"                  => "files/fileio.jl",
    "write and save"          => "files/write-save.jl",
    "streaming writes"        => "files/streaming.jl",
    "document properties"     => "files/docprops.jl",
    "read internals"          => "files/read-internals.jl",
    "differential reads"      => "files/differential.jl",
])
testgroup("Cells", [
    "cell references"         => "cells/references.jl",
    "ranges"                  => "cells/ranges.jl",
    "getcell"                 => "cells/getcell.jl",
    "getindex/setindex"       => "cells/getindex-setindex.jl",
    "error values"            => "cells/error-values.jl",
    "strings"                 => "cells/strings.jl",
    "formulas"                => "cells/formulas.jl",
    "shifting references"     => "cells/shift-refs.jl",
])
testgroup("Tables", [
    "readtable/writetable"    => "tables/tables.jl",
    "Excel tables"            => "tables/excel-tables.jl",
])
testgroup("Sheets", [
    "add/copy/move sheets"    => "sheets/add-copy-move.jl",
    "defined names"           => "sheets/defined-names.jl",
    "insert/delete rows/cols" => "sheets/row-col-edit.jl",
    "merged cells"            => "sheets/merge.jl",
    "hide and group"          => "sheets/hide-group.jl",
    "panes"                   => "sheets/panes.jl",
    "sheet view"              => "sheets/sheet-view.jl",
    "sheet settings"          => "sheets/settings.jl",
])
testgroup("Formatting", [
    "colors"                  => "formatting/colors.jl",
    "styles"                  => "formatting/styles.jl",
    "displayed text"          => "formatting/display-text.jl",
    "autofit columns"         => "formatting/autofit.jl",
    "checkboxes"              => "formatting/checkboxes.jl",
    "conditional formats"     => "formatting/conditional.jl",
])
testgroup("Features", [
    "AutoFilter"              => "features/autofilter.jl",
    "data validation"         => "features/data-validation.jl",
    "hyperlinks"              => "features/hyperlinks.jl",
    "comments"                => "features/comments.jl",
    "ignored errors"          => "features/ignored-errors.jl",
    "sparklines"              => "features/sparklines.jl",
])
testgroup("Drawings", [
    "DrawingML"               => "drawings/drawingml.jl",
    "images"                  => "drawings/images.jl",
    "text boxes and shapes"   => "drawings/shapes.jl",
    "pictures in cells"       => "drawings/cell-images.jl",
    "deleting objects"        => "drawings/delete-objects.jl",
])
testgroup("Charts", [
    "chart schema"            => "charts/schema.jl",
    "charts"                  => "charts/charts.jl",
    "c: properties"           => "charts/c-props.jl",
    "cx: properties"          => "charts/cx-props.jl",
    "chart imports"           => "charts/imports.jl",
])
testgroup("Quality", [
    "Aqua"                    => "quality/aqua.jl",
])

# leave the temporary directory so it can be removed
cd(test_run_directory)
