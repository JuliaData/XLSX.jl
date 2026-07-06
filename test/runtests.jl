import XLSX
import Tables
using Test, Dates, XML
using OrderedCollections: OrderedDict
import DataFrames, Random
import Distributions as Dist
import CSV
using StyledStrings
using ZipArchives: ZipReader, zip_names, zip_readentry

# If SAVE_FILES is true then every testset that creates an XLSXFile in memory
# or that writes an xlsx file will have a file created in `outdir`.
# It may be possible to create files that XLSX.jl can handle but Excel, which 
# has stricter rules, rejects. By creating the output of every testset as a 
# file explicitly, it is possible to confirm that this never happens in practice 
# by manually opening each output file.
# If TRUE, approx 220 files are created in `outdir`.
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

function current_testset_label()
    ts = Test.get_testset()
    return ts === nothing ? "untitled" : ts.description
end

function sanitize_filename(s::AbstractString)
    return replace(s, r"[^A-Za-z0-9_-]" => "_")
end

const outfile_counter = Ref(0)

function save_outfile(xf::XLSX.XLSXFile)
    outfile_counter[] += 1
    label = sanitize_filename(current_testset_label())
    fname = "outfile_$(lpad(outfile_counter[], 3, '0'))_$(label).xlsx"
    XLSX.writexlsx(joinpath(outdir, fname), xf)
end

function save_outfile(name::AbstractString)
    outfile_counter[] += 1
    label = sanitize_filename(current_testset_label())
    ext = splitext(name)[2]  # preserves ".xlsx", ".xlsm", ".xltx", ".xltm", etc.
    fname = "outfile_$(lpad(outfile_counter[], 3, '0'))_$(label)$(ext)"
    cp(name, joinpath(outdir, fname))
end

function save_outfile(io::IO)
    outfile_counter[] += 1
    label = sanitize_filename(current_testset_label())
    fname = "outfile_$(lpad(outfile_counter[], 3, '0'))_$(label).xlsx"
    pos = position(io)
    seekstart(io)
    open(joinpath(outdir, fname), "w") do out
        write(out, io)
    end
    seek(io, pos)
end

# Checks whether `data` equals `test_data`
function check_test_data(data::Vector{S}, test_data::Vector{T}) where {S,T}

    @test length(data) == length(test_data)

    function size_of_data(d::Vector{T}) where {T}
        isempty(d) && return (0, 0)
        return length(d[1]), length(d)
    end

    rows, cols = size_of_data(test_data)

    for col in 1:cols
        @test length(data[col]) == length(test_data[col])
    end

    for row in 1:rows, col in 1:cols
        test_value = test_data[col][row]
        value = data[col][row]

        if test_value === nothing
            @test ismissing(value)
        elseif ismissing(test_value) || (isa(test_value, AbstractString) && isempty(test_value))
            @test ismissing(value) || (isa(value, AbstractString) && isempty(value))
        else
            if isa(test_value, Integer) || isa(value, Integer)
                @test isa(test_value, Integer)
                @test isa(value, Integer)
            end

            if isa(test_value, Real) && !isa(test_value, Integer)
                @test isapprox(value, test_value)
            else
                @test value == test_value
            end
        end
    end

    nothing
end

#=
Tests have been split across different files each essentially focused on 
a different subset of functionality. Note, however, that all functions can 
be applied in combination with any others, so the perfect separation of 
functions is tests is not complete and it wouuld be undesirable to make 
it so.
=#
include(joinpath(files_directory, "Cell-names_tests.jl"))
include(joinpath(files_directory, "Colors-tests.jl"))
include(joinpath(files_directory, "Conditional-format_tests.jl"))
include(joinpath(files_directory, "Copy-Add_tests.jl"))
include(joinpath(files_directory, "Defined-names_tests.jl"))
include(joinpath(files_directory, "Edit_tests.jl"))
include(joinpath(files_directory, "Errors_tests.jl"))
include(joinpath(files_directory, "FileIO_tests.jl"))
include(joinpath(files_directory, "Filemodes_tests.jl"))
include(joinpath(files_directory, "Formulas_tests.jl"))
include(joinpath(files_directory, "Getcell_tests.jl"))
include(joinpath(files_directory, "Getindex-Setindex_tests.jl"))
include(joinpath(files_directory, "Images_tests.jl"))
include(joinpath(files_directory, "Merge_tests.jl"))
include(joinpath(files_directory, "Namespace_tests.jl"))
include(joinpath(files_directory, "Ranges_tests.jl"))
include(joinpath(files_directory, "Strict-format_tests.jl"))
include(joinpath(files_directory, "Strings_tests.jl"))
include(joinpath(files_directory, "Styles_tests.jl"))
include(joinpath(files_directory, "Tables_tests.jl"))
include(joinpath(files_directory, "Test-files_tests.jl"))
include(joinpath(files_directory, "Write-Save_tests.jl"))
include(joinpath(files_directory, "XML_tests.jl"))

#Some residual tests still here!
@testset "Time and DateTime" begin
    @test XLSX.excel_value_to_time(0.82291666666666663) == Dates.Time(Dates.Hour(19), Dates.Minute(45))
    @test XLSX.time_to_excel_value(XLSX.excel_value_to_time(0.2)) == 0.2
    @test XLSX.excel_value_to_datetime(43206.805447106482, false) == Date(2018, 4, 16) + Dates.Time(Dates.Hour(19), Dates.Minute(19), Dates.Second(51))
    @test XLSX.excel_value_to_datetime(XLSX.datetime_to_excel_value(Date(2018, 4, 16) + Dates.Time(Dates.Hour(19), Dates.Minute(19), Dates.Second(51)), false), false) == Date(2018, 4, 16) + Dates.Time(Dates.Hour(19), Dates.Minute(19), Dates.Second(51))

    dt = Date(2018, 4, 1)
    @test XLSX.excel_value_to_date(XLSX.date_to_excel_value(dt, false), false) == dt
    @test XLSX.excel_value_to_date(XLSX.date_to_excel_value(dt, true), true) == dt
end

@testset "No Dimension" begin
    noDim = XLSX.openxlsx(joinpath(data_directory, "NoDim.xlsx"), mode="rw")
    Dim = XLSX.readxlsx(joinpath(data_directory, "customXml.xlsx"))
    @test noDim[1].dimension == Dim[1].dimension
    @test noDim[2].dimension == Dim[2].dimension

    f = XLSX.newxlsx()
    s = f[1]
    for i = 10:20, j = 10:20
        s[i, j] = i + j
    end
    XLSX.set_dimension!(s, XLSX.CellRange(XLSX.CellRef("J10"), XLSX.CellRef("T20")))
    @test XLSX.get_dimension(s) == XLSX.CellRange(XLSX.CellRef("J10"), XLSX.CellRef("T20"))
    s["A1"] = 2
    @test XLSX.get_dimension(s) == XLSX.CellRange(XLSX.CellRef("A1"), XLSX.CellRef("T20"))
    SAVE_FILES && save_outfile(f)
end

# issue #67
@testset "row_index" begin
    filename = "test_pr67.xlsx"
    XLSX.openxlsx(filename, mode="w") do xf
        xf[1]["A2"] = 5
        xf[1]["A1"] = 7
    end
    @test isfile(filename)
    SAVE_FILES && save_outfile(filename)
    isfile(filename) && rm(filename)
end

@testset "show xlsx" begin
    @testset "single sheet" begin
        xf = XLSX.readxlsx(joinpath(src_data_directory, "blank.xlsx"))
        io = IOBuffer()
        show(io, xf)
        lines = split(String(take!(io)), '\n')

        expected_tail = "            sheetname size          range        \n-------------------------------------------------\n               Sheet1 1x1           A1:A1        \n"

        @test join(lines[2:end], '\n') == expected_tail
    end

    @testset "multiple sheets" begin
        xf = XLSX.readxlsx(joinpath(data_directory, "Book1.xlsx"))
        io = IOBuffer()
        show(io, xf)
        lines = split(String(take!(io)), '\n')

        expected_tail = "            sheetname size          range        \n-------------------------------------------------\n               Sheet1 7x2           B2:C8        \n               Sheet2 3x3           A1:C3        \n"

        @test join(lines[2:end], '\n') == expected_tail
    end

    @testset "IOBuffer source" begin
        source = IOBuffer()
        write(source, read(joinpath(data_directory, "Book1.xlsx")))
        seekstart(source)
        xf = XLSX.readxlsx(source)
        io=IOBuffer()
        show(io, xf)
        result = String(take!(io))
        expected = "XLSXFile(IOBuffer) containing 2 Worksheets\n            sheetname size          range        \n-------------------------------------------------\n               Sheet1 7x2           B2:C8        \n               Sheet2 3x3           A1:C3        \n"

        @test result == expected
    end
end

# issues #62, #75
@testset "relative paths" begin
    let
        xf = XLSX.readxlsx(joinpath(data_directory, "openpyxl.xlsx"))
        @test XLSX.sheetnames(xf) == ["Sheet", "Test1"]
        @test xf["Test1"]["A1"] == "One"
        @test xf["Test1"]["A2"] == 1
        show(IOBuffer(), xf)
        show(IOBuffer(), xf["Sheet"])
        show(IOBuffer(), xf["Test1"])
    end

    let
        dtable = XLSX.readtable(joinpath(data_directory, "openpyxl.xlsx"), "Test1")
        data, col_names = dtable.data, dtable.column_labels
        @test data == [[1, 3], [2, 4]]
        @test col_names == [:One, :Two]
    end
end


# issues #62, #71
@testset "windows compatibility" begin
    xf = XLSX.open_xlsx_template(joinpath(data_directory, "issue62_71.xlsx"))
    @test xf["Sheet1"]["A1"] == "One"
    @test xf["Sheet1"]["A2"] == 1

    @test collect(keys(xf.binary_data)) == ["xl/printerSettings/printerSettings1.bin"]
end

@testset "stream iterator" begin
    f = XLSX.openxlsx(joinpath(data_directory, "general.xlsx"), enable_cache=false)
    s = f["table"]
    for sheetrow in XLSX.eachrow(s)
        for column in 2:4
            cell = XLSX.getcell(sheetrow, column)
            if XLSX.row_number(cell) == 2 && XLSX.column_number(cell) == 2
                @test XLSX.getdata(s, cell) == "Column B"
            end
            if XLSX.row_number(cell) == 12 && XLSX.column_number(cell) == 2
                @test XLSX.getdata(s, cell) == "trash"
            end
        end
    end
end