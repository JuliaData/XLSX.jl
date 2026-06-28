@testset "Error Values" begin
    f = XLSX.openxlsx(joinpath(data_directory, "Errors.xlsx"), mode="rw")
    s=f[1]
    @test Int64(XLSX.getcell(s, "A1").value) == 1
    @test Int64(XLSX.getcell(s, "B1").value) == 2
    @test Int64(XLSX.getcell(s, "C1").value) == 3
    @test Int64(XLSX.getcell(s, "D1").value) == 4
    @test Int64(XLSX.getcell(s, "E1").value) == 5
    @test Int64(XLSX.getcell(s, "F1").value) == 6
    @test Int64(XLSX.getcell(s, "G1").value) == 7
    @test Int64(XLSX.getcell(s, "H1").value) == 3

    XLSX.writexlsx("mytest.xlsx", f, overwrite=true)
    SAVE_FILES && save_outfile("mytest.xlsx")

    f2 = XLSX.openxlsx("mytest.xlsx", mode="rw")
    s=f[1]
    @test Int64(XLSX.getcell(s, "A1").value) == 1
    @test Int64(XLSX.getcell(s, "B1").value) == 2
    @test Int64(XLSX.getcell(s, "C1").value) == 3
    @test Int64(XLSX.getcell(s, "D1").value) == 4
    @test Int64(XLSX.getcell(s, "E1").value) == 5
    @test Int64(XLSX.getcell(s, "F1").value) == 6
    @test Int64(XLSX.getcell(s, "G1").value) == 7
    @test Int64(XLSX.getcell(s, "H1").value) == 3

    isfile("mytest.xlsx") && rm("mytest.xlsx")

    f = XLSX.readxlsx(joinpath(data_directory, "Errors.xlsx"))
    sh = f[1]
    @test XLSX.iserror(sh, "A1") == true
    @test XLSX.iserror(sh, 1, 1) == true
    @test XLSX.iserror(sh, "I1") == false
    @test XLSX.iserror(sh, 1, 9) == false
    @test XLSX.iserror(sh, "A1:I1") == [true true true true true true true true false]
    @test XLSX.iserror(sh, 1, 1:9) == [true true true true true true true true false]
    @test XLSX.iserror(sh, "A1:B1,D1:E1") == [[true true], [true true]]
    @test XLSX.iserror(sh, 1, [1, 2, 4, 5]) == [true, true, true, true]
    @test XLSX.iserror(sh, :) == Bool[
                                        1 1 1 1 1 1 1 1
                                        0 0 0 0 0 0 0 0
                                        0 0 0 0 0 0 0 0
                                        0 0 0 0 0 0 0 0
                                        0 0 0 0 0 0 0 0
                                     ]
    @test XLSX.geterror(sh, "A1") == "#NULL!"
    @test XLSX.geterror(sh, 1, 1) == "#NULL!"
    @test XLSX.geterror(sh, "I1") == ""
    @test XLSX.geterror(sh, 1, 9) == ""
    @test XLSX.geterror(s, "A1:I1") == ["#NULL!"  "#DIV/0!"  "#VALUE!"  "#REF!"  "#NAME?"  "#NUM!"  "#N/A"  "#VALUE!"  ""]
    @test XLSX.geterror(s, 1, 1:9) == ["#NULL!"  "#DIV/0!"  "#VALUE!"  "#REF!"  "#NAME?"  "#NUM!"  "#N/A"  "#VALUE!"  ""]
    @test XLSX.geterror(sh, "A1:B1,D1:E1") == [["#NULL!" "#DIV/0!"], ["#REF!" "#NAME?"]]
    @test XLSX.geterror(sh, 1, [1, 2, 4, 5]) == ["#NULL!", "#DIV/0!", "#REF!", "#NAME?"]
    @test XLSX.geterror(sh, :) == [ "#NULL!"  "#DIV/0!"  "#VALUE!"  "#REF!"  "#NAME?"  "#NUM!"  "#N/A"  "#VALUE!"
                                    ""        ""         ""         ""       ""        ""       ""      ""
                                    ""        ""         ""         ""       ""        ""       ""      ""
                                    ""        ""         ""         ""       ""        ""       ""      ""
                                    ""        ""         ""         ""       ""        ""       ""      ""
                                 ]

end