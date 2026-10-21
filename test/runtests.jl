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
Tests have been split across different files each essentially focused on 
a different subset of functionality. Note, however, that all functions can 
be applied in combination with any others, so the perfect separation of 
functions in tests is not complete and it wouuld be undesirable to make 
it so.
=#
include(joinpath(files_directory, "helpers.jl"))  # shared helpers first
include(joinpath(files_directory, "Aqua_tests.jl"))
include(joinpath(files_directory, "Cell-names_tests.jl"))
include(joinpath(files_directory, "ChartExProps_tests.jl"))
include(joinpath(files_directory, "ChartImports_tests.jl"))
include(joinpath(files_directory, "ChartProps_tests.jl"))
include(joinpath(files_directory, "ChartSchema_tests.jl"))
include(joinpath(files_directory, "Charts_tests.jl"))
include(joinpath(files_directory, "Colors_tests.jl"))
include(joinpath(files_directory, "Conditional-format_tests.jl"))
include(joinpath(files_directory, "Copy-Add_tests.jl"))
include(joinpath(files_directory, "Defined-names_tests.jl"))
include(joinpath(files_directory, "Differential_tests.jl"))
include(joinpath(files_directory, "DrawingML_tests.jl"))
include(joinpath(files_directory, "Error-values_tests.jl"))
include(joinpath(files_directory, "Excel_tables_tests.jl"))
include(joinpath(files_directory, "FileIO_tests.jl"))
include(joinpath(files_directory, "Filemodes_tests.jl"))
include(joinpath(files_directory, "Fixture-files_tests.jl"))
include(joinpath(files_directory, "Formulas_tests.jl"))
include(joinpath(files_directory, "Getcell_tests.jl"))
include(joinpath(files_directory, "Getindex-Setindex_tests.jl"))
include(joinpath(files_directory, "Images_tests.jl"))
include(joinpath(files_directory, "Merge_tests.jl"))
include(joinpath(files_directory, "Namespace_tests.jl"))
include(joinpath(files_directory, "Panes_tests.jl"))
include(joinpath(files_directory, "Ranges_tests.jl"))
include(joinpath(files_directory, "Read-internals_tests.jl"))
include(joinpath(files_directory, "Strict-format_tests.jl"))
include(joinpath(files_directory, "Strings_tests.jl"))
include(joinpath(files_directory, "Styles_tests.jl"))
include(joinpath(files_directory, "Tables_tests.jl"))
include(joinpath(files_directory, "Write-Save_tests.jl"))
include(joinpath(files_directory, "XML_tests.jl"))

# leave the temporary directory so it can be removed
cd(test_run_directory)
