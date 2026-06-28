@testset "Write" begin
    f = XLSX.open_xlsx_template(joinpath(data_directory, "general.xlsx"))
    filename_copy = "general_copy.xlsx"

    XLSX.writexlsx(filename_copy, f)
    SAVE_FILES && save_outfile(filename_copy)
    @test isfile(filename_copy)

    f_copy = XLSX.readxlsx(filename_copy)

    s = f_copy["table"]
    s[:]
    dtable = XLSX.gettable(s)
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
    isfile(filename_copy) && rm(filename_copy)
end

@testset "Save" begin
    f = XLSX.openxlsx("saveable.xlsx", mode="w")
    XLSX.renamesheet!(f["Sheet1"], "new_name")
    s = f[1]
    s[1:10, 1:10] = "hello world"
    @test XLSX.savexlsx(f) == abspath("saveable.xlsx")
    SAVE_FILES && save_outfile("saveable.xlsx")
    f2 = XLSX.openxlsx("saveable.xlsx", mode="rw")
    @test XLSX.hassheet(f2, "new_name")
    @test f2["new_name"][1, 1] == "hello world"
    @test f2["new_name"][10, 10] == "hello world"
    f2["new_name"][1:5, 1:5] = "goodbye world"
    XLSX.savexlsx(f2)
    SAVE_FILES && save_outfile("saveable.xlsx")
    f3 = XLSX.openxlsx("saveable.xlsx", mode="r")
    @test f3["new_name"][1, 1] == "goodbye world"
    @test f3["new_name"][5, 5] == "goodbye world"
    @test f3["new_name"][10, 10] == "hello world"
    isfile("saveable.xlsx") && rm("saveable.xlsx")
end