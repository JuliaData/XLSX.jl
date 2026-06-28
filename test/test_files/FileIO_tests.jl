function get_cols(source::XLSX.DataTable)
    return source.data, source.column_labels
end


@testset "No FileIO" verbose=true begin

    filename = joinpath(data_directory, "TestData.xlsx")

    try
        XLSX.load(filename, "Sheet1")
        @test false  # should error before this line
    catch e
        @test  e isa XLSX.XLSXError && occursin("requires the FileIO.jl package", e.msg)
    end
    try
        XLSX.save(filename, "Sheet1")
        @test false  # should error before this line
    catch e
        @test  e isa XLSX.XLSXError && occursin("requires the FileIO.jl package", e.msg)
    end
end

using Pkg
using FileIO

if Pkg.pkgversion(FileIO) > v"1.19.0"

    @testset "FileIO" begin

        filename = joinpath(data_directory, "TestData.xlsx")

        efile = load(filename, "Sheet1")

        @test Tables.istable(efile) == true # Defined in XLSX.jl

        # Test show renders expected number of rows and columns.
        @testset "show plain text" begin
            s = sprint(show, efile)
            @test s == "XLSX.DataTable with 13 columns and 4 rows."
        end

        @testset "read table" begin
            for source in [load(filename, "Sheet1", "C:O"; first_row=3), load(filename, "Sheet1")]
                df, names = get_cols(source)
                @test length(df) == 13
                @test length(df[1]) == 4

                @test df[1]  == [1., 1.5, 2., 2.5]
                @test df[2]  == ["A", "BB", "CCC", "DDDD"]
                @test df[3]  == [true, false, false, true]
                @test isequal(df[4],  [2, "EEEEE", false, 1.5])
                @test isequal(df[5],  [9., "III", missing, true])
                @test isequal(df[6],  [3., missing, 3.5, 4.])
                @test isequal(df[7],  ["FF", missing, "GGG", "HHHH"])
                @test isequal(df[8],  [missing, true, missing, false])
                @test df[9]  == [Date(2015, 3, 3), DateTime(2015, 2, 4, 10, 14), Date(1988, 4, 9), Dates.Time(15, 2, 0)]
                @test isequal(df[10], [Date(1965, 4, 3), DateTime(1950, 8, 9, 18, 40), Dates.Time(19, 0, 0), missing])
                @test all(ismissing, df[11])
                @test isequal(df[12], [missing, missing, missing, missing])
                @test isequal(df[13], [missing, 3.4, "HKEJW", missing])
            end

            df, names = get_cols(load(filename, "Sheet1", "C:O"; first_row=4, header=false))
            @test names == [:C, :D, :E, :F, :G, :H, :I, :J, :K, :L, :M, :N, :O]
            @test length(df[1]) == 4
            @test length(df) == 13
            @test df[1]  == [1., 1.5, 2., 2.5]
            @test df[2]  == ["A", "BB", "CCC", "DDDD"]
            @test df[3]  == [true, false, false, true]
            @test isequal(df[4],  [2, "EEEEE", false, 1.5])
            @test isequal(df[5],  [9., "III", missing, true])
            @test isequal(df[6],  [3., missing, 3.5, 4.])
            @test isequal(df[7],  ["FF", missing, "GGG", "HHHH"])
            @test isequal(df[8],  [missing, true, missing, false])
            @test df[9]  == [Date(2015, 3, 3), DateTime(2015, 2, 4, 10, 14), DateTime(1988, 4, 9), Dates.Time(15, 2, 0)]
            @test isequal(df[10], [Date(1965, 4, 3), DateTime(1950, 8, 9, 18, 40), Dates.Time(19, 0, 0), missing])
            @test all(ismissing, df[11])
            @test all(ismissing, df[12])
            @test isequal(df[13], [missing, 3.4, "HKEJW", missing])
            @test ismissing(df[12][4])

            good_colnames = [:c1, :c2, :c3, :c4, :c5, :c6, :c7, :c8, :c9, :c10, :c11, :c12, :c13]

            df, names = get_cols(load(filename, "Sheet1", "C:O"; first_row=4, header=false, column_labels=good_colnames))
            @test names == good_colnames
            @test length(df[1]) == 4
            @test length(df) == 13
            @test df[1]  == [1., 1.5, 2., 2.5]
            @test df[2]  == ["A", "BB", "CCC", "DDDD"]
            @test df[3]  == [true, false, false, true]
            @test isequal(df[4],  [2, "EEEEE", false, 1.5])
            @test isequal(df[5],  [9., "III", missing, true])
            @test isequal(df[6],  [3., missing, 3.5, 4.])
            @test isequal(df[7],  ["FF", missing, "GGG", "HHHH"])
            @test isequal(df[8],  [missing, true, missing, false])
            @test df[9]  == [Date(2015, 3, 3), DateTime(2015, 2, 4, 10, 14), DateTime(1988, 4, 9), Dates.Time(15, 2, 0)]
            @test isequal(df[10], [Date(1965, 4, 3), DateTime(1950, 8, 9, 18, 40), Dates.Time(19, 0, 0), missing])
            @test all(ismissing, df[11])
            @test all(ismissing, df[12])
            @test isequal(df[13], [missing, 3.4, "HKEJW", missing])
            @test ismissing(df[12][4])

            # Test for saving DataFrame to XLSX
            input = (Day = ["Nov. 27", "Nov. 28", "Nov. 29"], Highest = [78, 79, 75]) |> DataFrames.DataFrame
            save("file.xlsx", input)
            SAVE_FILES && save_outfile("file.xlsx")
            output = load("file.xlsx", "Sheet1") |> DataFrames.DataFrame
            @test input == output
            rm("file.xlsx")

            # Test for saving DataFrame to XLSX with sheetname keyword
            input = (Day = ["Nov. 27", "Nov. 28", "Nov. 29"], Highest = [78, 79, 75]) |> DataFrames.DataFrame
            save("file.xlsx", input, sheetname="SheetName")
            SAVE_FILES && save_outfile("file.xlsx")
            output = load("file.xlsx", "SheetName") |> DataFrames.DataFrame
            @test input == output
            rm("file.xlsx")

            df, names = get_cols(load(filename, "Sheet1"; column_labels=good_colnames))
            @test names == good_colnames
            @test length(df[1]) == 4
            @test length(df) == 13
            @test df[1]  == [1., 1.5, 2., 2.5]
            @test df[2]  == ["A", "BB", "CCC", "DDDD"]
            @test df[3]  == [true, false, false, true]
            @test isequal(df[4],  [2, "EEEEE", false, 1.5])
            @test isequal(df[5],  [9., "III", missing, true])
            @test isequal(df[6],  [3., missing, 3.5, 4.])
            @test isequal(df[7],  ["FF", missing, "GGG", "HHHH"])
            @test isequal(df[8],  [missing, true, missing, false])
            @test df[9]  == [Date(2015, 3, 3), DateTime(2015, 2, 4, 10, 14), DateTime(1988, 4, 9), Dates.Time(15, 2, 0)]
            @test isequal(df[10], [Date(1965, 4, 3), DateTime(1950, 8, 9, 18, 40), Dates.Time(19, 0, 0), missing])
            @test all(ismissing, df[11])
            @test all(ismissing, df[12])
            @test isequal(df[13], [missing, 3.4, "HKEJW", missing])
            @test ismissing(df[12][4])

            # Too few column labels - Note: Bypass FileIO here to avoid false "Fatal Error" from FileIO when the error is correctly thrown by ExcelFiles for mismatched column_labels length.
            try
                XLSX.load(File{FileIO.format"Excel"}(filename), "Sheet1", "C:O"; header=true, column_labels=[:c1, :c2, :c3, :c4])
                @test false  # should error before this line
            catch e
                @test  e isa XLSX.XLSXError && occursin("`column_range` (length=13) and `column_labels` (length=4) must have the same length.", e.msg)
            end

            # Test for constructing DataFrame with empty header cell
            data, names = get_cols(load(filename, "Sheet2", "C:E"))
            @test names == [:Col1, Symbol("#Empty"), :Col3]

            # normalizenames keyword (XLSX.jl v0.11 only)
            data, names = get_cols(load(filename, "Sheet2", "C:E"; normalizenames=true))
            @test names == [:Col1, :_Empty, :Col3]
        end
        @testset "transposed tables" begin
            # Note: readtransposedtable cannot handle entirely empty rows/columns,
            # so the Transpose sheet omits those from the original Sheet1 data.
            # Note: eltype of mixed date columns is Dates.TimeType (not Any) when
            # there are no missing values, since a common supertype can be inferred.

            df, names = get_cols(load(filename, "Transpose"; transpose=true, first_column=2))
            @test length(df) == 5
            @test length(df[1]) == 4
            @test names == [Symbol("Some Float64s"), Symbol("Some Strings"), Symbol("Some Bools"), Symbol("Mixed with NA"), Symbol("Some dates")]

            @test df[1] == [1.0, 1.5, 2.0, 2.5]
            @test df[2] == ["A", "BB", "CCC", "DDDD"]
            @test df[3] == Bool[true, false, false, true]
            @test isequal(df[4], Any[9, "III", missing, true])
            @test df[5] == Dates.TimeType[Date(2015, 3, 3), DateTime(2015, 2, 4, 10, 14), Date(1988, 4, 9), Dates.Time(15, 2, 0)]
        end

        @testset "template and macro files" begin
            tbl = load(joinpath(data_directory, "Template File.xltx"), "Sheet1", "K:N")
            @test length(tbl.data) == 4
            @test length(tbl.data[1]) == 9
            tbl = load(joinpath(data_directory, "macro-enabled2.xltm"))
            @test length(tbl.data) == 1
            @test length(tbl.data[1]) == 0
            @test tbl.column_labels == [:hello]
            tbl = load(joinpath(data_directory, "macro-enabled.xlsm"))
            @test length(tbl.data) == 1
            @test length(tbl.data[1]) == 0
            @test tbl.column_labels == [:hello]
        end

    end
else
    @info "Skipping FileIO tests (requires FileIO > v1.19.0, got $(pkgversion(FileIO)))"
end
