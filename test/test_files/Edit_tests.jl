@testset "Edit Template" begin
    new_filename = "new_file_from_empty_template.xlsx"
    isfile(new_filename) && rm(new_filename)
    f = XLSX.open_empty_template()
    f["Sheet1"]["A1"] = "Hello"
    f["Sheet1"]["A2"] = 10
    XLSX.writexlsx(new_filename, f, overwrite=true)
    SAVE_FILES && save_outfile(new_filename)

    f = XLSX.readxlsx(new_filename)
    @test f["Sheet1"]["A1"] == "Hello"
    @test f["Sheet1"]["A2"] == 10

    rm(new_filename)
end

@testset "Edit" begin
    f = XLSX.open_xlsx_template(joinpath(data_directory, "general.xlsx"))
    s = f["general"]
    @test_throws XLSX.XLSXError s["A1"] = :sym
    XLSX.renamesheet!(s, "general") # no-op
    @test_throws XLSX.XLSXError XLSX.renamesheet!(s, "table") # name is taken
    XLSX.renamesheet!(s, "renamed_sheet") # retain old function to avoid breaking change
    @test s.name == "renamed_sheet"
    s["A1"] = "Hey You!"
    s["B1"] = "Out there in the cold..."
    s["A2"] = "Getting lonely getting old..."
    s["B2"] = "Can you feel me?"
    s["A3"] = 1000
    s["B3"] = 99.99

    # create a new sheet
    s = XLSX.addsheet!(f, "my_new_sheet_1")
    s = XLSX.addsheet!(f, "my_new_sheet_2")
    s["B1"] = "This is a new sheet"
    s["B2"] = "This is a new sheet"
    s = XLSX.addsheet!(f)
    s["B1"] = "unnamed sheet"

    XLSX.writexlsx("general_copy_2.xlsx", f, overwrite=true)
    SAVE_FILES && save_outfile("general_copy_2.xlsx")
    @test isfile("general_copy_2.xlsx")

    XLSX.openxlsx("general_copy_2.xlsx") do f
        s = f["renamed_sheet"]
        @test s["A1"] == "Hey You!"
        @test s["B1"] == "Out there in the cold..."
        @test s["A2"] == "Getting lonely getting old..."
        @test s["B2"] == "Can you feel me?"
        @test s["A3"] == 1000
        @test s["B3"] == 99.99
        f["my_new_sheet_1"]
        @test f["my_new_sheet_2"]["B1"] == "This is a new sheet"
        @test f["my_new_sheet_2"]["B2"] == "This is a new sheet"
        @test f["Sheet1"]["B1"] == "unnamed sheet"
    end

    isfile("general_copy_2.xlsx") && rm("general_copy_2.xlsx")
end
