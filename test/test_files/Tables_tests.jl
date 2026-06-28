@testset "Tables Helpers" begin

    test_data = Vector{Any}(undef, 3)
    test_data[1] = [missing, missing, "B5"]
    test_data[2] = ["C3", missing, missing]
    test_data[3] = [missing, "D4", missing]

    dtable = XLSX.readtable(joinpath(data_directory, "general.xlsx"), "table4")
    data, col_names = dtable.data, dtable.column_labels
    @test col_names == [:H1, :H2, :H3]

    dtable = XLSX.readtable(joinpath(data_directory, "general.xlsx"), "table4"; enable_cache=false)
    data, col_names = dtable.data, dtable.column_labels
    @test col_names == [:H1, :H2, :H3]

    @testset "Tables.jl DataTable interface" begin
        df = DataFrames.DataFrame(dtable)
        @test DataFrames.names(df) == ["H1", "H2", "H3"]
        @test size(df) == (3, 3)
        @test df[1, :H2] == "C3"
        @test df[3, :H1] == "B5"
        @test df[2, :H3] == "D4"
        @test ismissing(df[1, 1])
    end

    check_test_data(data, test_data)

    @test XLSX.readdata(joinpath(data_directory, "general.xlsx"), "table4", "E12") == "H1"
    test_data = Array{Any,2}(undef, 2, 1)
    test_data[1, 1] = "H2"
    test_data[2, 1] = "C3"

    @test XLSX.readdata(joinpath(data_directory, "general.xlsx"), "table4", "F12:F13") == test_data

    @testset "readtable select single column" begin
        dtable = XLSX.readtable(joinpath(data_directory, "general.xlsx"), "table4", "F")
        data, col_names = dtable.data, dtable.column_labels
        @test col_names == [:H2]
        @test data == Any[Any["C3"]]
    end

    @testset "readtable select column range" begin

        dtable = XLSX.readtable(joinpath(data_directory, "general.xlsx"), "table4", "F:G")
        data, col_names = dtable.data, dtable.column_labels
        @test col_names == [:H2, :H3]
        test_data = Any[Any["C3", missing], Any[missing, "D4"]]
        check_test_data(data, test_data)

        @test_throws XLSX.XLSXError XLSX.readtable(joinpath(data_directory, "general.xlsx"), "table4", "F:G"; header=false, column_labels=["a", "b", "c"])

    end

    @testset "readtable empty rows" begin

        t = XLSX.readtable(joinpath(data_directory, "EmptyTableRows.xlsx"), "EmptyRow", "B:C"; first_row=2, stop_in_empty_row=false, keep_empty_rows=true)
        test_data = Vector{Any}(undef, 2)
        test_data[1] = [1, 2, missing, missing, 3, 4, 5]
        test_data[2] = ["a", "b", missing, missing, "c", "d", "e"]
        check_test_data(t.data, test_data)
        @test t.column_labels == [Symbol("Col A"), Symbol("Col B")]
        @test t.column_label_index == Dict(Symbol("Col A") => 1, Symbol("Col B") => 2)

        t = XLSX.readtable(joinpath(data_directory, "EmptyTableRows.xlsx"), "EmptyCols", "B:C"; first_row=2, stop_in_empty_row=false, keep_empty_rows=true)
        test_data = Vector{Any}(undef, 2)
        test_data[1] = [1, 2, missing, missing, 3, 4, 5]
        test_data[2] = ["a", "b", missing, missing, "c", "d", "e"]
        check_test_data(t.data, test_data)
        @test t.column_labels == [Symbol("Col A"), Symbol("Col B")]
        @test t.column_label_index == Dict(Symbol("Col A") => 1, Symbol("Col B") => 2)

        t = XLSX.readtable(joinpath(data_directory, "EmptyTableRows.xlsx"), "MixedEmpty", "B:C"; first_row=2, stop_in_empty_row=false, keep_empty_rows=true)
        test_data = Vector{Any}(undef, 2)
        test_data[1] = [1, 2, missing, missing, 3, 4, 5, missing, missing, missing, missing, missing, 6, 7, 8, missing, missing, missing, missing, missing, missing, missing, 9, 10, 11, 12, 13, 14, 15, 16]
        test_data[2] = ["a", "b", missing, missing, "c", "d", "e", missing, missing, missing, missing, missing, "c", "d", "e", missing, missing, missing, missing, missing, missing, missing, "c", "d", "e", "c", "d", "e", "c", "d"]
        check_test_data(t.data, test_data)
        @test t.column_labels == [Symbol("Col A"), Symbol("Col B")]
        @test t.column_label_index == Dict(Symbol("Col A") => 1, Symbol("Col B") => 2)

        t = XLSX.readtable(joinpath(data_directory, "EmptyTableRows.xlsx"), "EmptyRow", "B:C"; first_row=2, stop_in_empty_row=false, keep_empty_rows=false)
        test_data = Vector{Any}(undef, 2)
        test_data[1] = [1, 2, 3, 4, 5]
        test_data[2] = ["a", "b", "c", "d", "e"]
        check_test_data(t.data, test_data)
        @test t.column_labels == [Symbol("Col A"), Symbol("Col B")]
        @test t.column_label_index == Dict(Symbol("Col A") => 1, Symbol("Col B") => 2)

        t = XLSX.readtable(joinpath(data_directory, "EmptyTableRows.xlsx"), "EmptyCols", "B:C"; first_row=2, stop_in_empty_row=false, keep_empty_rows=false)
        test_data = Vector{Any}(undef, 2)
        test_data[1] = [1, 2, 3, 4, 5]
        test_data[2] = ["a", "b", "c", "d", "e"]
        check_test_data(t.data, test_data)
        @test t.column_labels == [Symbol("Col A"), Symbol("Col B")]
        @test t.column_label_index == Dict(Symbol("Col A") => 1, Symbol("Col B") => 2)

        t = XLSX.readtable(joinpath(data_directory, "EmptyTableRows.xlsx"), "MixedEmpty", "B:C"; first_row=2, stop_in_empty_row=false, keep_empty_rows=false)
        test_data = Vector{Any}(undef, 2)
        test_data[1] = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16]
        test_data[2] = ["a", "b", "c", "d", "e", "c", "d", "e", "c", "d", "e", "c", "d", "e", "c", "d"]
        check_test_data(t.data, test_data)
        @test t.column_labels == [Symbol("Col A"), Symbol("Col B")]
        @test t.column_label_index == Dict(Symbol("Col A") => 1, Symbol("Col B") => 2)

        t = XLSX.readtable(joinpath(data_directory, "EmptyTableRows.xlsx"), "LeadingEmpty", "B:C"; first_row=2, stop_in_empty_row=false, keep_empty_rows=true)
        test_data = Vector{Any}(undef, 2)
        test_data[1] = [missing, missing, missing, 1, 2, missing, missing, 3, 4, 5, missing, missing, missing, missing, missing, 6, 7, 8, missing, missing, missing, missing, missing, missing, missing, 9, 10, 11, 12, 13, 14, 15, 16]
        test_data[2] = [missing, missing, missing, "a", "b", missing, missing, "c", "d", "e", missing, missing, missing, missing, missing, "c", "d", "e", missing, missing, missing, missing, missing, missing, missing, "c", "d", "e", "c", "d", "e", "c", "d"]
        check_test_data(t.data, test_data)
        @test t.column_labels == [Symbol("col A"), Symbol("col B")]
        @test t.column_label_index == Dict(Symbol("col A") => 1, Symbol("col B") => 2)

        t = XLSX.readtable(joinpath(data_directory, "EmptyTableRows.xlsx"), "LeadingEmpty", "B:C"; first_row=2, stop_in_empty_row=false, keep_empty_rows=false)
        test_data = Vector{Any}(undef, 2)
        test_data[1] = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16]
        test_data[2] = ["a", "b", "c", "d", "e", "c", "d", "e", "c", "d", "e", "c", "d", "e", "c", "d"]
        check_test_data(t.data, test_data)
        @test t.column_labels == [Symbol("col A"), Symbol("col B")]
        @test t.column_label_index == Dict(Symbol("col A") => 1, Symbol("col B") => 2)

        t = XLSX.readtable(joinpath(data_directory, "EmptyTableRows.xlsx"), "LeadingMixed", "B:C"; first_row=2, stop_in_empty_row=false, keep_empty_rows=true)
        test_data = Vector{Any}(undef, 2)
        test_data[1] = [missing, missing, missing, 1, 2, missing, missing, 3, 4, 5, missing, missing, missing, missing, missing, 6, 7, 8, missing, missing, missing, missing, missing, missing, missing, 9, 10, 11, 12, 13, 14, 15, 16]
        test_data[2] = [missing, missing, missing, "a", "b", missing, missing, "c", "d", "e", missing, missing, missing, missing, missing, "c", "d", "e", missing, missing, missing, missing, missing, missing, missing, "c", "d", "e", "c", "d", "e", "c", "d"]
        check_test_data(t.data, test_data)
        @test t.column_labels == [Symbol("col A"), Symbol("col B")]
        @test t.column_label_index == Dict(Symbol("col A") => 1, Symbol("col B") => 2)

        t = XLSX.readtable(joinpath(data_directory, "EmptyTableRows.xlsx"), "LeadingMixed", "B:C"; first_row=2, stop_in_empty_row=false, keep_empty_rows=false)
        test_data = Vector{Any}(undef, 2)
        test_data[1] = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16]
        test_data[2] = ["a", "b", "c", "d", "e", "c", "d", "e", "c", "d", "e", "c", "d", "e", "c", "d"]
        check_test_data(t.data, test_data)
        @test t.column_labels == [Symbol("col A"), Symbol("col B")]
        @test t.column_label_index == Dict(Symbol("col A") => 1, Symbol("col B") => 2)
    end

    @testset "stop function" begin

        t = XLSX.readtable(joinpath(data_directory, "EmptyTableRows.xlsx"), "LeadingEmpty", "B:C"; first_row=2, stop_in_empty_row=false, keep_empty_rows=true, stop_in_row_function=x -> !ismissing(x.cell_values[1]) && x.cell_values[1] == 2)
        test_data = Vector{Any}(undef, 2)
        test_data[1] = [missing, missing, missing, 1]
        test_data[2] = [missing, missing, missing, "a"]
        check_test_data(t.data, test_data)
        @test t.column_labels == [Symbol("col A"), Symbol("col B")]
        @test t.column_label_index == Dict(Symbol("col A") => 1, Symbol("col B") => 2)

        t = XLSX.readtable(joinpath(data_directory, "EmptyTableRows.xlsx"), "LeadingEmpty", "B:C"; first_row=2, stop_in_empty_row=false, keep_empty_rows=false, stop_in_row_function=x -> !ismissing(x.cell_values[1]) && x.cell_values[1] == 14)
        test_data = Vector{Any}(undef, 2)
        test_data[1] = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13]
        test_data[2] = ["a", "b", "c", "d", "e", "c", "d", "e", "c", "d", "e", "c", "d"]
        check_test_data(t.data, test_data)
        @test t.column_labels == [Symbol("col A"), Symbol("col B")]
        @test t.column_label_index == Dict(Symbol("col A") => 1, Symbol("col B") => 2)

        t = XLSX.readtable(joinpath(data_directory, "EmptyTableRows.xlsx"), "LeadingEmpty", "B:C"; first_row=2, stop_in_empty_row=false, keep_empty_rows=true, stop_in_row_function=x -> ismissing(x.cell_values[1]))
        @test isempty(t.data[1])
        @test isempty(t.data[2])
        @test t.column_labels == [Symbol("col A"), Symbol("col B")]
        @test t.column_label_index == Dict(Symbol("col A") => 1, Symbol("col B") => 2)

        t = XLSX.readtable(joinpath(data_directory, "EmptyTableRows.xlsx"), "LeadingEmpty", "B:C"; first_row=2, stop_in_empty_row=false, keep_empty_rows=false, stop_in_row_function=x -> !ismissing(x.cell_values[1]) && x.cell_values[1] == 1)
        @test isempty(t.data[1])
        @test isempty(t.data[2])
        @test t.column_labels == [Symbol("col A"), Symbol("col B")]
        @test t.column_label_index == Dict(Symbol("col A") => 1, Symbol("col B") => 2)

        t = XLSX.readtable(joinpath(data_directory, "EmptyTableRows.xlsx"), "LeadingMixed", "B:C"; first_row=2, stop_in_empty_row=false, keep_empty_rows=true, stop_in_row_function=x -> ismissing(x.cell_values[1]))
        @test isempty(t.data[1])
        @test isempty(t.data[2])
        @test t.column_labels == [Symbol("col A"), Symbol("col B")]
        @test t.column_label_index == Dict(Symbol("col A") => 1, Symbol("col B") => 2)

        t = XLSX.readtable(joinpath(data_directory, "EmptyTableRows.xlsx"), "LeadingMixed", "B:C"; first_row=2, stop_in_empty_row=true)
        @test isempty(t.data[1])
        @test isempty(t.data[2])
        @test t.column_labels == [Symbol("col A"), Symbol("col B")]
        @test t.column_label_index == Dict(Symbol("col A") => 1, Symbol("col B") => 2)

    end

    @testset "normalizenames" begin
        test_data = ["hello", "Hello 1", "123", Symbol("name")]
        @test XLSX.normalizename.(test_data) == [:hello, :Hello_1, :_123, :name]

        data = Vector{Any}()
        push!(data, [:sym1, :sym2, :sym3])
        push!(data, [1.0, 2.0, 3.0])
        push!(data, ["abc", "DeF", "gHi"])
        push!(data, [true, true, false])
        cols = ["1 col", "col \$2", "local", "col:4"]

        XLSX.writetable("mytest.xlsx", data, cols; overwrite=true)
        SAVE_FILES && save_outfile("mytest.xlsx")
        df = DataFrames.DataFrame(XLSX.readtable("mytest.xlsx", "Sheet1", normalizenames=true))
        @test DataFrames.names(df) == Any["_1_col", "col_2", "_local", "col_4"]

        isfile("mytest.xlsx") && rm("mytest.xlsx")

    end

    @testset "missing_strings" begin # issue #90
        t = XLSX.readtable(joinpath(data_directory, "missing_strings.xlsx"); missing_strings="N/A", stop_in_empty_row=false)
        test_data = Vector{Any}(undef, 2)
        test_data[1] = [1, 2, 3, missing, 4, 5]
        test_data[2] = [4, 5, 6, "Null", 7, 8]
        check_test_data(t.data, test_data)

        t = XLSX.readtable(joinpath(data_directory, "missing_strings.xlsx"); missing_strings="Null", stop_in_empty_row=false)
        test_data = Vector{Any}(undef, 2)
        test_data[1] = [1, 2, 3, "N/A", 4, 5]
        test_data[2] = [4, 5, 6, missing, 7, 8]
        check_test_data(t.data, test_data)

        t = XLSX.readtable(joinpath(data_directory, "missing_strings.xlsx"); missing_strings=["Null", "N/A"], stop_in_empty_row=false)
        test_data = Vector{Any}(undef, 2)
        test_data[1] = [1, 2, 3, missing, 4, 5]
        test_data[2] = [4, 5, 6, missing, 7, 8]
        check_test_data(t.data, test_data)

        t = XLSX.readtable(joinpath(data_directory, "missing_strings.xlsx"); missing_strings=["Null", "N/A"], stop_in_empty_row=false, keep_empty_rows=true)
        test_data = Vector{Any}(undef, 2)
        test_data[1] = [missing, 1, 2, 3, missing, missing, 4, 5]
        test_data[2] = [missing, 4, 5, 6, missing, missing, 7, 8]
        check_test_data(t.data, test_data)

    end

    @testset "Read DataFrame" begin

        df = XLSX.readto(joinpath(data_directory, "general.xlsx"), "table4", "F:G", DataFrames.DataFrame)
        @test names(df) == ["H2", "H3"]
        @test size(df) == (2, 2)
        @test df[1, :H2] == "C3"
        @test df[2, :H3] == "D4"
        @test ismissing(df[1, 2])
        @test ismissing(df[2, 1])

        df = XLSX.readto(joinpath(data_directory, "general.xlsx"), "table4", DataFrames.DataFrame)
        @test names(df) == ["H1", "H2", "H3"]
        @test size(df) == (3, 3)
        @test df[1, :H2] == "C3"
        @test df[2, :H3] == "D4"
        @test ismissing(df[1, :H1])
        @test ismissing(df[2, :H2])

        df = XLSX.readto(joinpath(data_directory, "general.xlsx"), DataFrames.DataFrame)
        @test names(df) == ["text", "regular text"]
        @test size(df) == (9, 2)
        @test df[1, "text"] == "integer"
        @test df[2, "regular text"] == 102.2
        @test df[3, 2] == Dates.Date(1983, 04, 16)
        @test df[5, 2] == Dates.DateTime(2018, 04, 16, 19, 19, 51)

        @test_throws XLSX.XLSXError XLSX.readto(joinpath(data_directory, "general.xlsx"))           # No sink
        @test_throws XLSX.XLSXError df = XLSX.readto(joinpath(data_directory, "general.xlsx"), 3)        # No sink
        @test_throws XLSX.XLSXError df = XLSX.readto(joinpath(data_directory, "general.xlsx"), 3, "F:G") # No sink

    end

end

@testset "Table" begin

    @test Tables.istable(XLSX.DataTable)

    @testset "Index" begin
        index = XLSX.Index("A:B", ["First", "Second"])
        @test index.column_labels == [:First, :Second]
        @test index.lookup[:First] == 1
        @test index.lookup[:Second] == 2
    end

    @testset "Bounds" begin
        f = XLSX.readxlsx(joinpath(data_directory, "book_sparse.xlsx"))
        s = f["Sheet1"]

        report = Vector{String}()
        for r in XLSX.eachrow(s)
            if !isempty(r)
                push!(report, string(XLSX.row_number(r), " - ", XLSX.column_bounds(r)))

                if XLSX.row_number(r) == 2
                    @test XLSX.last_column_index(r, 2) == 2
                elseif XLSX.row_number(r) == 3
                    @test XLSX.last_column_index(r, 3) == 4
                elseif XLSX.row_number(r) == 6
                    @test XLSX.last_column_index(r, 1) == 4
                    @test XLSX.last_column_index(r, 2) == 4
                    @test XLSX.last_column_index(r, 3) == 4
                    @test XLSX.last_column_index(r, 4) == 4
                    @test_throws XLSX.XLSXError XLSX.last_column_index(r, 5)
                elseif XLSX.row_number(r) == 9
                    @test XLSX.last_column_index(r, 2) == 3
                    @test XLSX.last_column_index(r, 3) == 3
                    @test XLSX.last_column_index(r, 5) == 5
                end
            end
        end

        @test report == ["2 - (2, 2)", "3 - (3, 4)", "6 - (1, 4)", "9 - (2, 5)"]
    end

    XLSX.openxlsx(joinpath(data_directory, "general.xlsx")) do f
        f["general"][:]
    end

    f = XLSX.readxlsx(joinpath(data_directory, "general.xlsx"))
    s = f["table"]
    s[:]
    dtable = XLSX.gettable(s)

    plaintext = sprint(show, dtable)
    @test plaintext == "XLSX.DataTable with 6 columns and 8 rows."

    data, col_names = dtable.data, dtable.column_labels
    @test col_names == [Symbol("Column B"), Symbol("Column C"), Symbol("Column D"), Symbol("Column E"), Symbol("Column F"), Symbol("Column G")]

    test_data = Vector{Any}(undef, 6)
    test_data[1] = collect(1:8)
    test_data[2] = ["Str1", missing, "Str1", "Str1", "Str2", "Str2", "Str2", "Str2"]
    test_data[3] = [Date(2018, 4, 21) + Dates.Day(i) for i in 0:7]
    test_data[4] = [missing, missing, missing, missing, missing, "a", "b", missing]
    test_data[5] = [0.2001132319, 0.2793987377, 0.0950591677, 0.0744023067, 0.8242278091, 0.6205883578, 0.9174151018, 0.6749604883]
    test_data[6] = [missing for i in 1:8]

    check_test_data(data, test_data)

    @test XLSX.infer_eltype(data[1]) == Int64
    @test XLSX.infer_eltype(data[2]) == Union{Missing,String}
    @test XLSX.infer_eltype(data[3]) == Date
    @test XLSX.infer_eltype(data[4]) == Union{Missing,String}
    @test XLSX.infer_eltype(data[5]) == Float64
    @test XLSX.infer_eltype(data[6]) == Any
    @test XLSX.infer_eltype(Vector{Int}()) == Int
    @test XLSX.infer_eltype(Vector{Float64}()) == Float64
    @test XLSX.infer_eltype(Vector{Any}()) == Any
    @test XLSX.infer_eltype([1, "1", 10.2]) == Any
    @test XLSX.infer_eltype([1, 10]) == Int
    @test XLSX.infer_eltype([1.0, 10.0]) == Float64
    @test XLSX.infer_eltype([1, 10.2]) == Float64 # Promote mixed int/float columns to float (#192)

    dtable_inferred = XLSX.gettable(s, infer_eltypes=true)
    data_inferred, col_names = dtable_inferred.data, dtable_inferred.column_labels
    @test eltype(data_inferred[1]) == Int64
    @test eltype(data_inferred[2]) == Union{Missing,String}
    @test eltype(data_inferred[3]) == Date
    @test eltype(data_inferred[4]) == Union{Missing,String}
    @test eltype(data_inferred[5]) == Float64
    @test eltype(data_inferred[6]) == Any

    function stop_function(r::XLSX.TableRow)
        v = r[Symbol("Column C")]
        return !ismissing(v) && v == "Str2"
    end

    function never_reaches_stop(r::XLSX.TableRow)
        v = r[Symbol("Column C")]
        return !ismissing(v) && v == "never was found"
    end

    dtable = XLSX.gettable(s, stop_in_row_function=never_reaches_stop)
    data, col_names = dtable.data, dtable.column_labels
    check_test_data(data, test_data)

    dtable = XLSX.gettable(s, stop_in_row_function=stop_function)
    data, col_names = dtable.data, dtable.column_labels
    @test col_names == [Symbol("Column B"), Symbol("Column C"), Symbol("Column D"), Symbol("Column E"), Symbol("Column F"), Symbol("Column G")]

    test_data = Vector{Any}(undef, 6)
    test_data[1] = collect(1:4)
    test_data[2] = ["Str1", missing, "Str1", "Str1"]
    test_data[3] = [Date(2018, 4, 21) + Dates.Day(i) for i in 0:3]
    test_data[4] = [missing, missing, missing, missing]
    test_data[5] = [0.2001132319, 0.2793987377, 0.0950591677, 0.0744023067]
    test_data[6] = [missing for i in 1:4]

    check_test_data(data, test_data)

    dtable = XLSX.gettable(s, stop_in_empty_row=false)
    data, col_names = dtable.data, dtable.column_labels
    @test col_names == [Symbol("Column B"), Symbol("Column C"), Symbol("Column D"), Symbol("Column E"), Symbol("Column F"), Symbol("Column G")]

    # test keep_empty_rows
    for (stop_in_empty_row, keep_empty_rows, n_rows) in [
        (false, false, 9),
        (false, true, 11),
        (true, false, 8),
        (true, true, 8)
    ]
        dtable = XLSX.gettable(s; stop_in_empty_row=stop_in_empty_row, keep_empty_rows=keep_empty_rows)
        @test all(col_name -> length(Tables.getcolumn(dtable, col_name)) == n_rows, Tables.columnnames(dtable))
    end

    test_data = Vector{Any}(undef, 6)
    test_data[1] = [1, 2, 3, 4, 5, 6, 7, 8, "trash"]
    test_data[2] = ["Str1", missing, "Str1", "Str1", "Str2", "Str2", "Str2", "Str2", missing]
    test_data[3] = Any[Date(2018, 4, 21) + Dates.Day(i) for i in 0:7]
    push!(test_data[3], "trash")
    test_data[4] = [missing, missing, missing, missing, missing, "a", "b", missing, missing]
    test_data[5] = [0.2001132319, 0.2793987377, 0.0950591677, 0.0744023067, 0.8242278091, 0.6205883578, 0.9174151018, 0.6749604883, "trash"]
    test_data[6] = Any[missing for i in 1:8]
    push!(test_data[6], "trash")

    check_test_data(data, test_data)

    # queries based on ColumnRange
    x = XLSX.getcellrange(s, XLSX.ColumnRange("B:D"))
    @test size(x) == (12, 3)
    y = XLSX.getcellrange(s, "B:D")
    @test size(y) == (12, 3)
    @test x == y
    @test_throws XLSX.XLSXError XLSX.getcellrange(s, "D:B")
    @test_throws XLSX.XLSXError XLSX.getcellrange(s, "A:C1")

    d = XLSX.getdata(s, "B:D")
    @test size(d) == (12, 3)
    @test_throws XLSX.XLSXError XLSX.getdata(s, "A:C1")
    @test d[1, 1] == "Column B"
    @test d[1, 2] == "Column C"
    @test d[1, 3] == "Column D"
    @test d[9, 1] == 8
    @test d[9, 2] == "Str2"
    @test d[9, 3] == Date(2018, 4, 28)
    @test d[11, 1] == "trash"
    @test ismissing(d[11, 2])
    @test d[11, 3] == "trash"
    @test ismissing(d[12, 1])
    @test ismissing(d[12, 2])
    @test ismissing(d[12, 3])

    d1 = XLSX.getdata(s, "2:3")
    @test size(d1) == (2, 8)
    @test d1[1, 2] == "Column B"
    @test d1[1, 4] == "Column D"
    @test d1[2, 2] == 1
    @test d1[2, 4] == Date(2018, 4, 21)

    d2 = f["table!B:D"]
    @test size(d) == size(d2)
    @test all(d .=== d2)

    @test_throws XLSX.XLSXError f["table!B1:D"]
    @test_throws XLSX.XLSXError f["table!D:B"]

    s = f["table2"]
    test_data = Vector{Any}(undef, 3)
    test_data[1] = ["A1", "A2", "A3", missing]
    test_data[2] = ["B1", "B2", missing, "B4"]
    test_data[3] = [missing, missing, missing, missing]

    dtable = XLSX.gettable(s)
    data, col_names = dtable.data, dtable.column_labels

    @test col_names == [:HA, :HB, :HC]
    check_test_data(data, test_data)

    for (ri, rowdata) in enumerate(XLSX.eachtablerow(s))
        if ismissing(test_data[1][ri])
            @test ismissing(rowdata[:HA])
        else
            @test rowdata[:HA] == test_data[1][ri]
        end

        @test XLSX.table_columns_count(rowdata) == 3
        @test XLSX.row_number(rowdata) == ri
        @test XLSX.get_column_labels(rowdata) == col_names
        @test XLSX.get_column_label(rowdata, 1) == :HA
        @test XLSX.get_column_label(rowdata, 2) == :HB
        @test XLSX.get_column_label(rowdata, 3) == :HC

        @test_throws XLSX.XLSXError XLSX.getdata(rowdata, :INVALID_COLUMN)
    end

    override_col_names_strs = ["ColumnA", "ColumnB", "ColumnC"]
    override_col_names = [Symbol(i) for i in override_col_names_strs]

    dtable = XLSX.gettable(s, column_labels=override_col_names_strs)
    data, col_names = dtable.data, dtable.column_labels

    @test col_names == override_col_names
    check_test_data(data, test_data)

    dtable = XLSX.gettable(s, "A:B", first_row=1)
    data, col_names = dtable.data, dtable.column_labels
    test_data_AB_cols = Vector{Any}(undef, 2)
    test_data_AB_cols[1] = test_data[1]
    test_data_AB_cols[2] = test_data[2]
    @test col_names == [:HA, :HB]
    check_test_data(data, test_data_AB_cols)

    dtable = XLSX.gettable(s, "A:B")
    data, col_names = dtable.data, dtable.column_labels
    test_data_AB_cols = Vector{Any}(undef, 2)
    test_data_AB_cols[1] = test_data[1]
    test_data_AB_cols[2] = test_data[2]
    @test col_names == [:HA, :HB]
    check_test_data(data, test_data_AB_cols)

    dtable = XLSX.gettable(s, "B:B", first_row=2)
    data, col_names = dtable.data, dtable.column_labels
    @test col_names == [:B1]
    @test length(data) == 1
    @test length(data[1]) == 1
    @test data[1][1] == "B2"

    dtable = XLSX.gettable(s, "B:C")
    data, col_names = dtable.data, dtable.column_labels
    @test col_names == [:HB, :HC]
    test_data_BC_cols = Vector{Any}(undef, 2)
    test_data_BC_cols[1] = ["B1", "B2"]
    test_data_BC_cols[2] = [missing, missing]
    check_test_data(data, test_data_BC_cols)

    dtable = XLSX.gettable(s, "B:C", first_row=2, header=false)
    data, col_names = dtable.data, dtable.column_labels
    @test col_names == [:B, :C]
    check_test_data(data, test_data_BC_cols)

    s = f["table3"]
    test_data = Vector{Any}(undef, 3)
    test_data[1] = [missing, missing, "B5"]
    test_data[2] = ["C3", missing, missing]
    test_data[3] = [missing, "D4", missing]
    dtable = XLSX.gettable(s)
    data, col_names = dtable.data, dtable.column_labels
    @test col_names == [:H1, :H2, :H3]
    check_test_data(data, test_data)
    @test_throws XLSX.XLSXError XLSX.find_row(XLSX.eachrow(s), 20)

    for r in XLSX.eachrow(s)
        @test isempty(XLSX.getcell(r, "A"))
        @test XLSX.getdata(s, XLSX.getcell(r, "B")) == "H1"
        @test r[2] == "H1"
        @test r["B"] == "H1"
        break
    end

    @test XLSX._find_first_row_with_data(s, 5) == 5
    @test_throws XLSX.XLSXError XLSX._find_first_row_with_data(s, 7)

    s = f["table4"]
    dtable = XLSX.gettable(s)
    data, col_names = dtable.data, dtable.column_labels
    @test col_names == [:H1, :H2, :H3]
    check_test_data(data, test_data)

    @testset "empty/invalid" begin
        XLSX.openxlsx(joinpath(data_directory, "general.xlsx")) do xf
            empty_sheet = XLSX.getsheet(xf, "empty")
            @test_throws XLSX.XLSXError XLSX.gettable(empty_sheet)
            itr = XLSX.eachrow(empty_sheet)
            @test_throws XLSX.XLSXError XLSX.find_row(itr, 1)
            @test_throws XLSX.XLSXError XLSX.getsheet(xf, "invalid_sheet")
        end
    end

    @testset "sheets 6/7/lookup/header_error" begin
        f = XLSX.readxlsx(joinpath(data_directory, "general.xlsx"))
        tb5 = f["table5"]
        test_data = Vector{Any}(undef, 1)
        test_data[1] = [1, 2, 3, 4, 5]
        dtable = XLSX.gettable(tb5)
        data, col_names = dtable.data, dtable.column_labels
        @test col_names == [:HEADER]
        check_test_data(data, test_data)
        tb6 = f["table6"]
        dtable = XLSX.gettable(tb6, first_row=3)
        data, col_names = dtable.data, dtable.column_labels
        @test col_names == [:HEADER]
        check_test_data(data, test_data)
        tb7 = f["table7"]
        dtable = XLSX.gettable(tb7, first_row=3)
        data, col_names = dtable.data, dtable.column_labels
        @test col_names == [:HEADER]
        check_test_data(data, test_data)

        sheet_lookup = f["lookup"]
        test_data = Vector{Any}(undef, 3)
        test_data[1] = [10, 20, 30]
        test_data[2] = ["name1", "name2", "name3"]
        test_data[3] = [100, 200, 300]
        dtable = XLSX.gettable(sheet_lookup)
        data, col_names = dtable.data, dtable.column_labels
        @test col_names == [:ID, :NAME, :VALUE]
        check_test_data(data, test_data)

        header_error_sheet = f["header_error"]
        dtable = XLSX.gettable(header_error_sheet, "B:E")
        data, col_names = dtable.data, dtable.column_labels
        @test col_names == [:COLUMN_A, :COLUMN_B, Symbol("COLUMN_A_2"), Symbol("#Empty")]
    end

    @testset "Consecutive passes" begin
        # Consecutive passes should yield the same results
        XLSX.openxlsx(joinpath(data_directory, "general.xlsx")) do f
            sl = f["lookup"]
            dtable = XLSX.gettable(sl)
            data, col_names = dtable.data, dtable.column_labels
            @test col_names == [:ID, :NAME, :VALUE]
            check_test_data(data, test_data)

            dtable = XLSX.gettable(sl)
            data, col_names = dtable.data, dtable.column_labels
            @test col_names == [:ID, :NAME, :VALUE]
            check_test_data(data, test_data)
        end
    end

    @testset "Transposed Tables" begin
        test_data = Vector{Any}(undef, 6)
        test_data[1] = [1940, 1950, 1960, 1970, 1980, 1990, 2000, 2010, 2020]
        test_data[2] = [1, 2, 3, 4, 5, 6, 7, 8, 9]
        test_data[3] = [10, 20, 30, 40, 50, 60, 70, 80, 90]
        test_data[4] = [100, 200, 300, 400, 500, 600, 700, 800, 900]
        test_data[5] = [0.1, 0.2, 0.3, 0.4, 0.5, 0.6, 0.7, 0.8, 0.9]
        test_data[6] = Any["Hello", Date(2025, 12, 19), 3, 3.33, "Hello", Date(2025, 12, 19), 3, 3.33, true]

        f = XLSX.openxlsx(joinpath(data_directory, "HTable.xlsx"), mode="rw")
        dtable = XLSX.gettransposedtable(f[1])
        data, colnames = dtable.data, dtable.column_labels
        @test colnames == Symbol.(["Year", "Col A", "Col B", "Col C", "Col D", "Col E"])
        dtable = XLSX.gettransposedtable(f[1]; normalizenames=true)
        data, colnames = dtable.data, dtable.column_labels
        @test colnames == Symbol.(["Year", "Col_A", "Col_B", "Col_C", "Col_D", "Col_E"])
        check_test_data(data, test_data)

        dtable = XLSX.readtransposedtable(joinpath(data_directory, "HTable.xlsx"))
        data, colnames = dtable.data, dtable.column_labels
        @test colnames == Symbol.(["Year", "Col A", "Col B", "Col C", "Col D", "Col E"])
        dtable = XLSX.readtransposedtable(joinpath(data_directory, "HTable.xlsx"); normalizenames=true)
        data, colnames = dtable.data, dtable.column_labels
        @test colnames == Symbol.(["Year", "Col_A", "Col_B", "Col_C", "Col_D", "Col_E"])
        check_test_data(data, test_data)

        dtable = XLSX.readtransposedtable(joinpath(data_directory, "HTable.xlsx"), "Offset", "2:7")
        data, colnames = dtable.data, dtable.column_labels
        @test colnames == Symbol.(["Year", "Col A", "Col B", "Col C", "Col D", "Col E"])
        check_test_data(data, test_data)

        dtable = XLSX.readtransposedtable(joinpath(data_directory, "HTable.xlsx"), "Multiple", "2:7")
        data, colnames = dtable.data, dtable.column_labels
        @test colnames == Symbol.(["Year", "Col A", "Col B", "Col C", "Col D", "Col E"])
        check_test_data(data, test_data)

        @test_throws XLSX.XLSXError XLSX.readtransposedtable(joinpath(data_directory, "HTable.xlsx"), "Multiple", "2:9") # Beyond worksheet dimension
        @test_throws XLSX.XLSXError XLSX.readtransposedtable(joinpath(data_directory, "HTable.xlsx"), "Multiple", "b:d") # Invalid row range
        test_data[1] = [1840, 1841, 1842, 1843, 1844, 1845, 1846, 1847, 1848]
        test_data[2] = [12.4, 12.6, 12.8, 13.0, 13.2, 13.4, 13.6, 13.8, 14.0]
        test_data[3] = [0.045, 0.046, 0.047, 0.048, 0.049, 0.050, 0.051, 0.052, 0.053]
        test_data[4] = [true, true, false, true, false, true, true, true, false]
        test_data[5] = [Time(10, 22), Time(10, 23), Time(10, 24), Time(10, 25), Time(10, 26), Time(10, 27), Time(10, 28), Time(10, 29), Time(10, 30)]
        test_data[6] = Any["Hello", Date(2025, 12, 19), 3, 3.33, "Hello", Date(2025, 12, 19), 3, 3.33, true]
        dtable = XLSX.readtransposedtable(joinpath(data_directory, "HTable.xlsx"), "Multiple", "2:7"; first_column="M")
        data, colnames = dtable.data, dtable.column_labels
        @test colnames == Symbol.(["date", "name1", "name2", "name3", "name4", "name5"])
        check_test_data(data, test_data)

        @test_throws XLSX.XLSXError XLSX.readtransposedtable(joinpath(data_directory, "HTable.xlsx"), "Multiple", "2:7"; first_column=4.2) # Invalid first_column
        @test_throws XLSX.XLSXError XLSX.readtransposedtable(joinpath(data_directory, "HTable.xlsx"), "Multiple", "2:7"; first_column=200) # first_column beyond worksheet dimension

        dtable = XLSX.readtransposedtable(joinpath(data_directory, "HTable.xlsx"), "Multiple", "2:7"; first_column=13, column_labels=["year", "Col_2", "Col_3", "Col_4", "Col_5", "Col_6"])
        data, colnames = dtable.data, dtable.column_labels
        @test colnames == Symbol.(["year", "Col_2", "Col_3", "Col_4", "Col_5", "Col_6"])
        check_test_data(data, test_data)

        @test_throws XLSX.XLSXError XLSX.readtransposedtable(joinpath(data_directory, "HTable.xlsx"), "Multiple", "2:7"; first_column=13, column_labels=["Col_2", "Col_3", "Col_4", "Col_5", "Col_6"]) # Missing one label

        test_data[1] = Any["date", 1840, 1841, 1842, 1843, 1844, 1845, 1846, 1847, 1848]
        test_data[2] = Any["name1", 12.4, 12.6, 12.8, 13.0, 13.2, 13.4, 13.6, 13.8, 14.0]
        test_data[3] = Any["name2", 0.045, 0.046, 0.047, 0.048, 0.049, 0.050, 0.051, 0.052, 0.053]
        test_data[4] = Any["name3", true, true, false, true, false, true, true, true, false]
        test_data[5] = Any["name4", Time(10, 22), Time(10, 23), Time(10, 24), Time(10, 25), Time(10, 26), Time(10, 27), Time(10, 28), Time(10, 29), Time(10, 30)]
        test_data[6] = Any["name5", "Hello", Date(2025, 12, 19), 3, 3.33, "Hello", Date(2025, 12, 19), 3, 3.33, true]
        dtable = XLSX.readtransposedtable(joinpath(data_directory, "HTable.xlsx"), "Multiple", "2:7"; first_column=13, header=false, column_labels=["year", "Col2", "Col3", "Col4", "Col5", "Col6"])
        data, colnames = dtable.data, dtable.column_labels
        @test colnames == Symbol.(["year", "Col2", "Col3", "Col4", "Col5", "Col6"])

        dtable = XLSX.readtransposedtable(joinpath(data_directory, "HTable.xlsx"), "Multiple", "2:7"; first_column=13, header=false)
        data, colnames = dtable.data, dtable.column_labels
        @test colnames == Symbol.(["Col_1", "Col_2", "Col_3", "Col_4", "Col_5", "Col_6"])

        test_data = Vector{Any}(undef, 4)
        test_data[1] = ["A", "B", "C", "D"]
        test_data[2] = [10, 20, 30, 40]
        test_data[3] = [15, 25, 35, 40]
        test_data[4] = [20, 30, 40, 50]
        dtable = XLSX.readtransposedtable(joinpath(data_directory, "HTable.xlsx"), "Example")
        data, colnames = dtable.data, dtable.column_labels
        @test colnames == Symbol.(["Category", "Variable 1", "Variable 2", "Variable 3"])
        check_test_data(data, test_data)

        XLSX.addsheet!(f, "ExpandedDim")
        s = f["ExpandedDim"]
        s["B2"] = "Category"
        s["B3"] = "Variable 1"
        s["B4"] = "Variable 2"
        s["B5"] = "Variable 3"
        s["C2"] = test_data[1]
        s["C3"] = test_data[2]
        s["C4"] = test_data[3]
        s["C5"] = test_data[4]
        s["M10"] = "ExtraData"
        @test XLSX.get_dimension(s) == XLSX.CellRange("A1:M10")
        dtable = XLSX.gettransposedtable(s)
        data, colnames = dtable.data, dtable.column_labels
        @test colnames == Symbol.(["Category", "Variable 1", "Variable 2", "Variable 3"])
        check_test_data(data, test_data)
        SAVE_FILES && save_outfile(f)

    end
end

@testset "writetable" begin

    @testset "single" begin
        col_names = ["Integers", "Strings", "Floats", "Booleans", "Dates", "Times", "DateTimes", "AbstractStrings", "Rational", "Irrationals", "MixedStringNothingMissing"]
        data = Vector{Any}(undef, 11)
        data[1] = [1, 2, missing, UInt8(4)]
        data[2] = ["Hey", "You", "Out", "There"]
        data[3] = [101.5, 102.5, missing, 104.5]
        data[4] = [true, false, missing, true]
        data[5] = [Date(2018, 2, 1), Date(2018, 3, 1), Date(2018, 5, 20), Date(2018, 6, 2)]
        data[6] = [Dates.Time(19, 10), Dates.Time(19, 20), Dates.Time(19, 30), Dates.Time(0, 0)]
        data[7] = [Dates.DateTime(2018, 5, 20, 19, 10), Dates.DateTime(2018, 5, 20, 19, 20), Dates.DateTime(2018, 5, 20, 19, 30), Dates.DateTime(2018, 5, 20, 19, 40)]
        data[8] = SubString.(["Hey", "You", "Out", "There"], 1, 2)
        data[9] = [1 // 2, 1 // 3, missing, 22 // 3]
        data[10] = [pi, sqrt(2), missing, sqrt(5)]
        data[11] = [nothing, "middle", missing, nothing]

        XLSX.writetable("output_table.xlsx", data, col_names, overwrite=true, sheetname="report", anchor_cell="B2")
        SAVE_FILES && save_outfile("output_table.xlsx")
        @test isfile("output_table.xlsx")

        dtable = XLSX.readtable("output_table.xlsx", "report")
        read_data, read_column_names = dtable.data, dtable.column_labels
        @test length(read_column_names) == length(col_names)
        for c in axes(col_names, 1)
            @test Symbol(col_names[c]) == read_column_names[c]
        end
        check_test_data(read_data, data)
    end

    @testset "multiple" begin
        report_1_column_names = ["HEADER_A", "HEADER_B"]
        report_1_data = Vector{Any}(undef, 2)
        report_1_data[1] = [1, 2, 3]
        report_1_data[2] = ["A", "B", ""]

        report_2_column_names = ["COLUMN_A", "COLUMN_B"]
        report_2_data = Vector{Any}(undef, 2)
        report_2_data[1] = [Date(2017, 2, 1), Date(2018, 2, 1)]
        report_2_data[2] = [10.2, 10.3]

        XLSX.writetable("output_tables.xlsx", overwrite=true, REPORT_A=(report_1_data, report_1_column_names), REPORT_B=(report_2_data, report_2_column_names))
        SAVE_FILES && save_outfile("output_tables.xlsx")
        @test isfile("output_tables.xlsx")

        dtable = XLSX.readtable("output_tables.xlsx", "REPORT_A")
        data, labels = dtable.data, dtable.column_labels
        @test labels[1] == :HEADER_A
        @test labels[2] == :HEADER_B
        check_test_data(data, report_1_data)

        dtable = XLSX.readtable("output_tables.xlsx", "REPORT_B")
        data, labels = dtable.data, dtable.column_labels
        @test labels[1] == :COLUMN_A
        @test labels[2] == :COLUMN_B
        check_test_data(data, report_2_data)
        XLSX.writetable("output_tables.xlsx", [("REPORT_A", report_1_data, report_1_column_names), ("REPORT_B", report_2_data, report_2_column_names)], overwrite=true)
        SAVE_FILES && save_outfile("output_tables.xlsx")

        dtable = XLSX.readtable("output_tables.xlsx", "REPORT_A")
        data, labels = dtable.data, dtable.column_labels
        @test labels[1] == :HEADER_A
        @test labels[2] == :HEADER_B
        check_test_data(data, report_1_data)

        dtable = XLSX.readtable("output_tables.xlsx", "REPORT_B")
        data, labels = dtable.data, dtable.column_labels
        @test labels[1] == :COLUMN_A
        @test labels[2] == :COLUMN_B
        check_test_data(data, report_2_data)

        report_1_column_names = [:HEADER_A, :HEADER_B]
        report_2_column_names = [:COLUMN_A, :COLUMN_B]
        XLSX.writetable("output_tables.xlsx", [("REPORT_A", report_1_data, report_1_column_names), ("REPORT_B", report_2_data, report_2_column_names)], overwrite=true)
        SAVE_FILES && save_outfile("output_tables.xlsx")

        dtable = XLSX.readtable("output_tables.xlsx", "REPORT_A")
        data, labels = dtable.data, dtable.column_labels
        @test labels[1] == :HEADER_A
        @test labels[2] == :HEADER_B
        check_test_data(data, report_1_data)

        dtable = XLSX.readtable("output_tables.xlsx", "REPORT_B")
        data, labels = dtable.data, dtable.column_labels
        @test labels[1] == :COLUMN_A
        @test labels[2] == :COLUMN_B
        check_test_data(data, report_2_data)

        report_1_column_names = ["HEADER_A", "HEADER_B"]
        report_1_data = [["1", "2", "3"], ["A", "B", ""]]

        report_2_column_names = ["COLUMN_A", "COLUMN_B"]
        report_2_data = Vector{Any}(undef, 2)
        report_2_data[1] = [Date(2017, 2, 1), Date(2018, 2, 1)]
        report_2_data[2] = [10.2, 10.3]
        XLSX.writetable("output_tables.xlsx", [("REPORT_A", report_1_data, report_1_column_names), ("REPORT_B", report_2_data, report_2_column_names)], overwrite=true)
        SAVE_FILES && save_outfile("output_tables.xlsx")

        dtable = XLSX.readtable("output_tables.xlsx", "REPORT_A")
        data, labels = dtable.data, dtable.column_labels
        @test labels[1] == :HEADER_A
        @test labels[2] == :HEADER_B
        check_test_data(data, report_1_data)

        dtable = XLSX.readtable("output_tables.xlsx", "REPORT_B")
        data, labels = dtable.data, dtable.column_labels
        @test labels[1] == :COLUMN_A
        @test labels[2] == :COLUMN_B
        check_test_data(data, report_2_data)
    end

    @testset "writetable to IO" begin
        dt = XLSX.DataTable(Any[Any[1, 2, 3], Any[4, 5, 6]], [:a, :b])
        io = IOBuffer()
        XLSX.writetable(io, "Test" => dt)
        SAVE_FILES && save_outfile(io)
        seek(io, 0)
        dt_read = XLSX.readtable(io, "Test")
        @test dt_read.data == dt.data
        @test dt_read.column_labels == dt.column_labels
        @test dt_read.column_label_index == dt.column_label_index
    end

    @testset "extended types" begin # Issue #239
        @enum enums begin
            enum1
            enum2
            enum3
        end

        data = Vector{Any}()
        push!(data, [:sym1, :sym2, :sym3])
        push!(data, [1.0, 2.0, 3.0])
        push!(data, ["abc", "DeF", "gHi"])
        push!(data, [true, true, false])
        push!(data, [XLSX.CellRef("A1"), XLSX.CellRef("B2"), XLSX.CellRef("CCC34000")])
        push!(data, collect(instances(enums)))
        cols = [string(eltype(x)) for x in data]

        XLSX.writetable("mytest.xlsx", data, cols; overwrite=true)
        SAVE_FILES && save_outfile("mytest.xlsx")

        f = XLSX.readxlsx("mytest.xlsx")
        @test f[1]["A1"] == "Symbol"
        @test f[1]["A1:A4"] == Any["Symbol"; "sym1"; "sym2"; "sym3";;] # A 2D Array, size (4, 1)
        @test f[1]["B1"] == "Float64"
        @test f[1]["B1:B4"] == Any["Float64"; 1.0; 2.0; 3.0;;]
        @test f[1]["C1"] == "String"
        @test f[1]["C1:C4"] == Any["String"; "abc"; "DeF"; "gHi";;]
        @test f[1]["D1"] == "Bool"
        @test f[1]["D1:D4"] == Any["Bool"; true; true; false;;]
        @test f[1]["E1"] == "CellRef"
        @test f[1]["E1:E4"] == Any["CellRef"; "A1"; "B2"; "CCC34000";;]
        @test f[1]["F1"] == "enums"
        @test f[1]["F1:F4"] == Any["enums"; "enum1"; "enum2"; "enum3";;]
    end

    @testset "NaN and Inf" begin # Issues #342 and #179
        col_names = ["Integers", "Strings", "Floats", "Booleans", "Dates", "Times", "DateTimes"]
        data = Vector{Any}(undef, 7)
        data[1] = [1, 2, missing, 4]
        data[2] = ["Hey", "You", "Out", "There"]
        data[3] = [-Inf, Inf, missing, NaN]
        data[4] = [true, false, missing, true]
        data[5] = [Date(2018, 2, 1), Date(2018, 3, 1), Date(2018, 5, 20), Date(2018, 6, 2)]
        data[6] = [Dates.Time(19, 10), Dates.Time(19, 20), Dates.Time(19, 30), Dates.Time(19, 40)]
        data[7] = [Dates.DateTime(2018, 5, 20, 19, 10), Dates.DateTime(2018, 5, 20, 19, 20), Dates.DateTime(2018, 5, 20, 19, 30), Dates.DateTime(2018, 5, 20, 19, 40)]
        table = NamedTuple{Tuple(Symbol(x) for x in col_names)}(Tuple(data))

        f=XLSX.newxlsx()
        s=f[1]
        XLSX.writetable!(s, table)
        @test s["C2"] == "-Inf"
        @test s["C3"] == "Inf"
        @test ismissing(s["C4"])
        @test s["C5"] == "NaN"
        SAVE_FILES && save_outfile(f)

        XLSX.writetable("mytest.xlsx", data, col_names; overwrite=true)
        SAVE_FILES && save_outfile("mytest.xlsx")

        f2=XLSX.readxlsx("mytest.xlsx")
        s2=f2[1]
        @test s2["C2"] == "-Inf"
        @test s2["C3"] == "Inf"
        @test ismissing(s2["C4"])
        @test s2["C5"] == "NaN"

        df=XLSX.readto("mytest.xlsx", DataFrames.DataFrame)
        @test all(isequal.(df.Floats, ["-Inf", "Inf", missing, "NaN"]))

    end

    # delete files created by this testset
    delete_files = ["output_table.xlsx", "output_tables.xlsx", "mytest.xlsx"]
    for f in delete_files
        isfile(f) && rm(f)
    end
end

@testset "Tables.jl integration" begin
    f = XLSX.readxlsx(joinpath(data_directory, "general.xlsx"))
    s = f["table"]
    ct = XLSX.eachtablerow(s) |> Tables.columntable
    @test isequal(ct, NamedTuple{(Symbol("Column B"), Symbol("Column C"), Symbol("Column D"), Symbol("Column E"), Symbol("Column F"), Symbol("Column G"))}(([1, 2, 3, 4, 5, 6, 7, 8], Union{Missing,String}["Str1", missing, "Str1", "Str1", "Str2", "Str2", "Str2", "Str2"], Date[Date(2018, 04, 21), Date(2018, 04, 22), Date(2018, 04, 23), Date(2018, 04, 24), Date(2018, 04, 25), Date(2018, 04, 26), Date(2018, 04, 27), Date(2018, 04, 28)], Union{Missing,String}[missing, missing, missing, missing, missing, "a", "b", missing], [0.2001132319106511, 0.27939873773400004, 0.09505916768351352, 0.07440230673248627, 0.82422780912015, 0.620588357787471, 0.9174151017732964, 0.6749604882690108], Missing[missing, missing, missing, missing, missing, missing, missing, missing])))
    rt = XLSX.eachtablerow(s) |> Tables.rowtable
    rt2 = [
        NamedTuple{(Symbol("Column B"), Symbol("Column C"), Symbol("Column D"), Symbol("Column E"), Symbol("Column F"), Symbol("Column G"))}((1, "Str1", Date(2018, 04, 21), missing, 0.2001132319106511, missing)),
        NamedTuple{(Symbol("Column B"), Symbol("Column C"), Symbol("Column D"), Symbol("Column E"), Symbol("Column F"), Symbol("Column G"))}((2, missing, Date(2018, 04, 22), missing, 0.27939873773400004, missing)),
        NamedTuple{(Symbol("Column B"), Symbol("Column C"), Symbol("Column D"), Symbol("Column E"), Symbol("Column F"), Symbol("Column G"))}((3, "Str1", Date(2018, 04, 23), missing, 0.09505916768351352, missing)),
        NamedTuple{(Symbol("Column B"), Symbol("Column C"), Symbol("Column D"), Symbol("Column E"), Symbol("Column F"), Symbol("Column G"))}((4, "Str1", Date(2018, 04, 24), missing, 0.07440230673248627, missing)),
        NamedTuple{(Symbol("Column B"), Symbol("Column C"), Symbol("Column D"), Symbol("Column E"), Symbol("Column F"), Symbol("Column G"))}((5, "Str2", Date(2018, 04, 25), missing, 0.82422780912015, missing)),
        NamedTuple{(Symbol("Column B"), Symbol("Column C"), Symbol("Column D"), Symbol("Column E"), Symbol("Column F"), Symbol("Column G"))}((6, "Str2", Date(2018, 04, 26), "a", 0.620588357787471, missing)),
        NamedTuple{(Symbol("Column B"), Symbol("Column C"), Symbol("Column D"), Symbol("Column E"), Symbol("Column F"), Symbol("Column G"))}((7, "Str2", Date(2018, 04, 27), "b", 0.9174151017732964, missing)),
        NamedTuple{(Symbol("Column B"), Symbol("Column C"), Symbol("Column D"), Symbol("Column E"), Symbol("Column F"), Symbol("Column G"))}((8, "Str2", Date(2018, 04, 28), missing, 0.6749604882690108, missing))
    ]
    @test isequal(rt, rt2)
    ct = XLSX.eachtablerow(f["table2"]) |> Tables.columntable
    @test length(ct) == 3
    @test length(ct[1]) == 4
    ct = XLSX.eachtablerow(f["general"]) |> Tables.columntable
    @test length(ct) == 2
    @test length(ct[1]) == 9
    ct = XLSX.eachtablerow(f["table3"]) |> Tables.columntable
    @test length(ct) == 3
    @test length(ct[1]) == 3
    ct = XLSX.eachtablerow(f["table4"]) |> Tables.columntable
    @test length(ct) == 3
    @test length(ct[1]) == 3
    ct = XLSX.eachtablerow(f["table5"]) |> Tables.columntable
    @test length(ct) == 1
    @test length(ct[1]) == 5
    ct = XLSX.eachtablerow(f["table6"]) |> Tables.columntable
    @test length(ct) == 1
    @test isempty(ct.hey)
    #    @test ct.hey == Any[missing]
    ct = XLSX.eachtablerow(f["table7"]; header=false) |> Tables.columntable
    @test length(ct) == 1
    @test length(ct[1]) == 1

    # write
    col_names = ["Integers", "Strings", "Floats", "Booleans", "Dates", "Times", "DateTimes"]
    data = Vector{Any}(undef, 7)
    data[1] = [1, 2, missing, 4]
    data[2] = ["Hey", "You", "Out", "There"]
    data[3] = [101.5, 102.5, missing, 104.5]
    data[4] = [true, false, missing, true]
    data[5] = [Date(2018, 2, 1), Date(2018, 3, 1), Date(2018, 5, 20), Date(2018, 6, 2)]
    data[6] = [Dates.Time(19, 10), Dates.Time(19, 20), Dates.Time(19, 30), Dates.Time(19, 40)]
    data[7] = [Dates.DateTime(2018, 5, 20, 19, 10), Dates.DateTime(2018, 5, 20, 19, 20), Dates.DateTime(2018, 5, 20, 19, 30), Dates.DateTime(2018, 5, 20, 19, 40)]
    table = NamedTuple{Tuple(Symbol(x) for x in col_names)}(Tuple(data))

    XLSX.writetable("output_table.xlsx", table, overwrite=true, sheetname="report", anchor_cell="B2")
    SAVE_FILES && save_outfile("output_table.xlsx")
    @test isfile("output_table.xlsx")

    XLSX.openxlsx("output_table2.xlsx", mode="w") do xf
        sheet = XLSX.getsheet(xf, 1)
        XLSX.renamesheet!(sheet, "report")
        XLSX.writetable!(sheet, table)
    end
    SAVE_FILES && save_outfile("output_table2.xlsx")

    for file in ["output_table.xlsx", "output_table2.xlsx"]
        try
            f = XLSX.readxlsx(file)
            s = f["report"]
            table2 = XLSX.eachtablerow(s) |> Tables.columntable
            @test isequal(table, table2)
        finally
            isfile(file) && rm(file)
        end
    end

    # multiple tables in same file
    table2 = (a=[1, 2, 3, 4], b=["a", "b", "c", "d"])
    XLSX.writetable("output_table3.xlsx", "report1" => table, "report2" => table2)
    SAVE_FILES && save_outfile("output_table3.xlsx")
    XLSX.writetable("output_table4.xlsx", ["report1" => table, "report2" => table2])
    SAVE_FILES && save_outfile("output_table4.xlsx")
    for file in ["output_table4.xlsx", "output_table3.xlsx"]
        try
            f = XLSX.readxlsx(file)
            result1 = Tables.columntable(XLSX.eachtablerow(f["report1"]))
            result2 = Tables.columntable(XLSX.eachtablerow(f["report2"]))
            @test isequal(table, result1)
            @test isequal(table2, result2)
        finally
            isfile(file) && rm(file)
        end
    end

    @testset "Tables.jl with DataFrames" begin
        f = XLSX.readxlsx(joinpath(data_directory, "general.xlsx"))
        s = f["table"]
        df = XLSX.eachtablerow(s) |> DataFrames.DataFrame
        @test size(df) == (8, 6)
        @test df[!, "Column B"] == collect(1:8)
        @test df[!, "Column D"] == collect(Date(2018, 4, 21):Dates.Day(1):Date(2018, 4, 28))
        @test all(ismissing.(df[!, "Column G"]))

        file = joinpath(@__DIR__, "test_report.xlsx")

        try
            df1 = DataFrames.DataFrame(COL1=[10, 20, 30], COL2=["Fist", "Sec", "Third"])
            df2 = DataFrames.DataFrame(AA=["aa", "bb"], AB=[10.1, 10.2])
            XLSX.writetable(file, "REPORT_A" => df1, "REPORT_B" => df2, overwrite=true)
            SAVE_FILES && save_outfile(file)
        finally
            isfile(file) && rm(file)
        end
    end

    @testset "Tables.jl as sink" begin
        f = CSV.read(joinpath(data_directory, "iris.csv"), XLSX.XLSXFile)
        @test XLSX.hassheet(f, "Sheet1") == true
        sheet = f["Sheet1"]
        @test sheet["A1"] == "sepal_length"
        @test sheet["B1"] == "sepal_width"
        @test sheet["C1"] == "petal_length"
        @test sheet["D1"] == "petal_width"
        @test sheet["E1"] == "species"
        @test sheet["A150"] ≈ 6.2
        @test sheet["B150"] ≈ 3.4
        @test sheet["C150"] ≈ 5.4
        @test sheet["D150"] ≈ 2.3
        @test sheet["E150"] == "virginica"
        XLSX.setFormula(f[1], "G1", "=GROUPBY(E1:E151,A1:D151,AVERAGE,3,1)")
        @test XLSX.get_formula_from_cache(sheet, XLSX.CellRef("G1")) == XLSX.Formula("_xlfn.GROUPBY(E1:E151,A1:D151,_xleta.AVERAGE,3,1)", "array", "G1:G1", nothing)
        f[1]["M1"] = "versicolor"
        XLSX.setFormula(f[1], "M2", "=VLOOKUP(M1,G1#,3,FALSE)")
        @test XLSX.get_formula_from_cache(sheet, XLSX.CellRef("M2")) == XLSX.Formula("=VLOOKUP(M1,_xlfn.ANCHORARRAY(G1),3,FALSE)", "array", "M2:M2", nothing)
        XLSX.setFormula(f[1], "G1", "=GROUPBY(E1:E151,A1:D151,STDEV.P,3,1)")
        @test XLSX.get_formula_from_cache(sheet, XLSX.CellRef("G1")) == XLSX.Formula("_xlfn.GROUPBY(E1:E151,A1:D151,_xleta.STDEV.P,3,1)", "array", "G1:G1", nothing)
        XLSX.setFormula(f[1], "G10", "_xlfn.GROUPBY(E1:E151,A1:D151,_xlfn.LAMBDA(_xlpm.x,AVERAGE(_xlpm.x)),3,1)"; raw=true)
        @test XLSX.get_formula_from_cache(sheet, XLSX.CellRef("G10")) == XLSX.Formula("_xlfn.GROUPBY(E1:E151,A1:D151,_xlfn.LAMBDA(_xlpm.x,AVERAGE(_xlpm.x)),3,1)", "array", "G10:G10", nothing)
        SAVE_FILES && save_outfile(f)
    end

    @testset "type inference in `eachtablerow`" begin
        f = XLSX.readxlsx(joinpath(data_directory, "general.xlsx"))
        df = XLSX.eachtablerow(f["lookup"], "B:J") |> DataFrames.DataFrame
        @test eltype.(eachcol(df)) == [Int64, String, Int64, Any, Int64, String, Any, Int64, Int64]
    end
end
