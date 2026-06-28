# issue #117
@testset "whitespace nodes" begin
    xf = XLSX.readxlsx(joinpath(data_directory, "noutput_first_second_third.xlsx"))
    @test XLSX.sheetnames(xf) == ["NOTES", "DATA"]
    @test xf["NOTES"]["A1"] == "Nominal GNP/GDP"
    @test xf["NOTES"]["A9"] == "Last updated on: August 29, 2019"
    @test xf["DATA"]["A5"] == "Date"
    @test xf["DATA"]["A6"] == "1965:Q3"
    @test xf["DATA"]["B6"] ≈ 6.7731
    @test xf["DATA"]["E5"] == "Most_Recent"
    @test xf["DATA"]["E7"] ≈ 12.6215
end

# issue #303
@testset "xml:space" begin
    f = XLSX.openxlsx(joinpath(data_directory, "sstTest.xlsx"), mode="rw")
    s = f[1]
    @test XLSX.getdata(s, :) == ["  hello" "    "; "  hello  " "    "; " hello\">" "    "; "hello\">" "    "; "  hello" "    "]
    s["C1"] = " "
    s["C2"] = " hello"
    s["C3"] = "hello "
    s["C4"] = " hello "
    s["C5"] = " \"hello\" "
    @test XLSX.getdata(s, "C1:C5") == Any[" "; " hello"; "hello "; " hello "; " \"hello\" ";;]
    XLSX.writexlsx("mydata.xlsx", f, overwrite=true)
    SAVE_FILES && save_outfile("mydata.xlsx")
    @test XLSX.readdata("mydata.xlsx", 1, :) == ["  hello" "    " " "; "  hello  " "    " " hello"; " hello\">" "    " "hello "; "hello\">" "    " " hello "; "  hello" "    " " \"hello\" "]
    XLSX.writetable("mydata.xlsx", [["  hello", "  hello  ", " hello\">", "hello\">", "  hello"], ["    ", "    ", "    ", "    ", "    "], [" ", " hello", "hello ", " hello ", " \"hello\" "]], ["Col_A", "Col_B", "Col_C"]; overwrite=true)
    SAVE_FILES && save_outfile("mydata.xlsx")
    @test XLSX.readdata("mydata.xlsx", 1, :) == ["Col_A" "Col_B" "Col_C"; "  hello" "    " " "; "  hello  " "    " " hello"; " hello\">" "    " "hello "; "hello\">" "    " " hello "; "  hello" "    " " \"hello\" "]
    isfile("mydata.xlsx") && rm("mydata.xlsx")
end

# issue #243
@testset "xml bom" begin
    xf = XLSX.readxlsx(joinpath(data_directory, "Bom - issue243.xlsx"))
    @test XLSX.sheetnames(xf) == ["QMJ Factors", "Definition", "Data Sources", "--> Additional Global Factors", "Disclosures"]
    @test XLSX.sheetcount(xf) == 5
    @test XLSX.hassheet(xf, "QMJ Factors") == true
    @test xf["QMJ Factors"]["H833"] ≈ -0.0686846616503713
end

@testset "escape" begin

    @test XML.escape("hello&world<'") == "hello&amp;world&lt;&apos;"
    @test XML.unescape("hello&amp;world&lt;&apos;") == "hello&world<'"

    esc_filename = "output_table_escape_test.xlsx"
    isfile(esc_filename) && rm(esc_filename)

    esc_col_names = ["&' & \" < > '", "I❤Julia", "\"<'&O-O&'>\"", "<&>"]
    esc_sheetname = "& & \" > < "
    esc_data = Vector{Any}(undef, 4)
    esc_data[1] = ["11&&", "12\"&", "13<&", "14>&", "15'&"]
    esc_data[2] = ["21&&&&", "22&\"&&", "23&<&&", "24&>&&", "25&'&&"]
    esc_data[3] = ["31&&&&&&", "32&&\"&&&", "33&&<&&&", "34&&>&&&", "35&&'&&&"]
    esc_data[4] = ["41& &; &&", "42\" \"; \"\"", "43< <; <<", "44> >; >>", "45' '; ''"]
    XLSX.writetable(esc_filename, esc_data, esc_col_names, overwrite=true, sheetname=esc_sheetname)
    SAVE_FILES && save_outfile(esc_filename)

    dtable = XLSX.readtable(esc_filename, esc_sheetname)
    r1_data, r1_col_names = dtable.data, dtable.column_labels
    check_test_data(r1_data, esc_data)
    @test r1_col_names[4] == Symbol(esc_col_names[4])
    @test r1_col_names[3] == Symbol(esc_col_names[3])
    @test r1_col_names[2] == Symbol(esc_col_names[2])
    @test r1_col_names[1] == Symbol(esc_col_names[1])
    isfile(esc_filename) && rm(esc_filename)

    # compare to the backup version: escape.xlsx
    dtable = XLSX.readtable(joinpath(data_directory, "escape.xlsx"), esc_sheetname)
    r2_data, r2_col_names = dtable.data, dtable.column_labels
    check_test_data(r2_data, esc_data)
    check_test_data(r2_data, r1_data)
    @test string(r2_col_names[4]) == esc_col_names[4]
    @test string(r2_col_names[3]) == esc_col_names[3]
    @test string(r2_col_names[2]) == esc_col_names[2]
    @test string(r2_col_names[1]) == esc_col_names[1]

    esc_col_names = ["&; &amp; &quot; &lt; &gt; &apos; ", "I❤Julia", "\"<'&O-O&'>\"", "<&>"]
    esc_sheetname = string( esc_col_names[1],esc_col_names[2],esc_col_names[3],esc_col_names[4])[1:30] # There is a hard limit in Excel
    esc_data = Vector{Any}(undef, 4)
    esc_data[1] = ["11&amp;&",    "12&quot;&",    "13&lt;&",    "14&gt;&",    "15&apos;&"    ]
    esc_data[2] = ["21&&amp;&&",  "22&&quot;&&",  "23&&lt;&&",  "24&&gt;&&",  "25&&apos;&&"  ]
    esc_data[3] = ["31&&&amp;&&&","32&&&quot;&&&","33&&&lt;&&&","34&&&gt;&&&","35&&&apos;&&&"]
    esc_data[4] = ["41& &; &&",   "42\" \"; \"\"","43< <; <<",  "44> >; >>",  "45' '; ''"    ]
    XLSX.writetable(esc_filename, esc_data, esc_col_names, overwrite=true, sheetname=esc_sheetname)
    SAVE_FILES && save_outfile(esc_filename)

    dtable = XLSX.readtable(esc_filename, esc_sheetname)
    r3_data, r3_col_names = dtable.data, dtable.column_labels
    check_test_data(r3_data, esc_data)
    @test r3_col_names[4] == Symbol( esc_col_names[4] )
    @test r3_col_names[3] == Symbol( esc_col_names[3] )
    @test r3_col_names[2] == Symbol( esc_col_names[2] )
    @test r3_col_names[1] == Symbol( esc_col_names[1] )
    isfile(esc_filename) && rm(esc_filename)

    # compare to the backup version: escape2.xlsx
    dtable = XLSX.readtable(joinpath(data_directory, "escape2.xlsx"), esc_sheetname)
    r4_data, r4_col_names = dtable.data, dtable.column_labels
    check_test_data(r4_data, esc_data)
    check_test_data(r4_data, r3_data)
    @test r4_col_names[4] == Symbol( esc_col_names[4] )
    @test r4_col_names[3] == Symbol( esc_col_names[3] )
    @test r4_col_names[2] == Symbol( esc_col_names[2] )
    @test r4_col_names[1] == Symbol( esc_col_names[1] )

end