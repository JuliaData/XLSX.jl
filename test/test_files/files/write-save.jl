@testset "Write" begin
    f = XLSX.open_xlsx_template(joinpath(data_directory, "general.xlsx"))
    filename_copy = "general_copy.xlsx"

    XLSX.writexlsx(filename_copy, f; overwrite=true)
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
    SAVE_FILES && save_outfile(f2)
    SAVE_FILES && save_outfile(f)

    @testset "save back to an IO" begin
        src = XLSX.newxlsx(); src[1]["A1"] = "original"
        XLSX.addsheet!(src, "big")
        src["big"]["A1:A2000"] = [string(i, "-", i^2) for i in 1:2000, _ in 1:1]
        bytes = (io = IOBuffer(); XLSX.writexlsx(io, src); take!(io))

        # do-block "rw" replaces the buffer's content, truncating when the result is shorter
        buf = IOBuffer(copy(bytes); read=true, write=true)
        XLSX.openxlsx(buf; mode="rw") do xf
            XLSX.deletesheet!(xf, "big")
            xf[1]["A1"] = "short"
            xf[1]["B1"] = 2
        end
        @test buf.size < length(bytes)
        seekstart(buf)
        f4 = XLSX.readxlsx(buf)
        @test f4[1]["A1"] == "short"
        @test f4[1]["B1"] == 2

        # savexlsx does the same
        buf = IOBuffer(copy(bytes); read=true, write=true)
        xf = XLSX.openxlsx(buf; mode="rw")
        xf[1]["C1"] = 3
        @test XLSX.savexlsx(xf) === nothing
        seekstart(buf)
        @test XLSX.readxlsx(buf)[1]["C1"] == 3

        # a read-only buffer can't be overwritten
        xf = XLSX.openxlsx(IOBuffer(copy(bytes)); mode="rw")
        @test_throws XLSX.XLSXError XLSX.savexlsx(xf)
    end
end

# issue #67
@testset "row_index" begin
    filename = "test_pr67.xlsx"
    XLSX.openxlsx(filename, mode="w") do xf
        xf[1]["A2"] = 5
        xf[1]["A1"] = 7
    end
    @test isfile(filename)
    @test XLSX.readxlsx(filename)[1]["A1:A2"] == Any[7; 5;;]   # rows written out of order read back in order
    SAVE_FILES && save_outfile(filename)
    isfile(filename) && rm(filename)
end

@testset "Edit Template" begin
    new_filename = "new_file_from_empty_template.xlsx"
    f = XLSX.open_empty_template()
    f["Sheet1"]["A1"] = "Hello"
    f["Sheet1"]["A2"] = 10
    XLSX.writexlsx(new_filename, f, overwrite=true)
    SAVE_FILES && save_outfile(new_filename)

    f = XLSX.readxlsx(new_filename)
    @test f["Sheet1"]["A1"] == "Hello"
    @test f["Sheet1"]["A2"] == 10

    isfile(new_filename) && rm(new_filename)
end

@testset "Fix timestamp" begin
    t = Dates.now(Dates.UTC) - Dates.Second(1)
    xf = XLSX.newxlsx()
    f = "docProps/core.xml"
    date_format = Dates.dateformat"yyyy-mm-ddTHH:MM:SSZ"
    i, j = XLSX.get_idces(xf.data[f], "cp:coreProperties", "dcterms:created") # i, j should always be found unless the `blank.xlsx` file is updated
    @test t < DateTime(XML.value(xf.data[f][i][j][1]), date_format)
    i, j = XLSX.get_idces(xf.data[f], "cp:coreProperties", "dcterms:modified")
    @test t < DateTime(XML.value(xf.data[f][i][j][1]), date_format)
    SAVE_FILES && save_outfile(xf)

    xf = XLSX.newxlsx(; update_timestamp=false) # do not update timestamp
    i, j = XLSX.get_idces(xf.data[f], "cp:coreProperties", "dcterms:created") # i, j should always be found unless the `blank.xlsx` file is updated
    @test XML.value(xf.data[f][i][j][1]) == "2018-05-22T02:41:32Z"
    i, j = XLSX.get_idces(xf.data[f], "cp:coreProperties", "dcterms:modified")
    @test XML.value(xf.data[f][i][j][1]) == "2018-05-22T02:42:04Z"
    SAVE_FILES && save_outfile(xf)
end

@testset "writetable refuses to overwrite an existing file" begin
    @test_throws XLSX.XLSXError XLSX.savexlsx(XLSX.newxlsx())   # blank: no file name

    outfile = "errbranch_exists.xlsx"
    XLSX.writetable(outfile, [[1, 2]], ["a"]; overwrite=true)
    @test_throws XLSX.XLSXError XLSX.writetable(outfile; Sheet1=([[1]], ["a"]))
    @test_throws XLSX.XLSXError XLSX.writetable(outfile, [("Sheet1", [[1]], ["a"])])
    SAVE_FILES && save_outfile(outfile)
    isfile(outfile) && rm(outfile)
end

@testset "review fixes: write" begin
    @testset "hidden sheets stay hidden" begin
        f = XLSX.newxlsx("A"); s2 = XLSX.addsheet!(f, "B")
        s2.is_hidden = true
        XLSX.writexlsx("hidden.xlsx", f; overwrite=true)
        @test XLSX.readxlsx("hidden.xlsx")["B"].is_hidden
        SAVE_FILES && save_outfile("hidden.xlsx")
    end
    @testset "hidden and veryHidden sheets read from a file survive an edit" begin
        f = XLSX.newxlsx("Visible"); XLSX.addsheet!(f, "Hidden"); XLSX.addsheet!(f, "VeryHidden")
        XLSX.writexlsx("states_base.xlsx", f; overwrite=true)
        r = ZipReader(read("states_base.xlsx"))
        open("states.xlsx", "w") do io                      # mark them as Excel does
            ZipWriter(io) do w
                for n in zip_names(r)
                    bytes = zip_readentry(r, n)
                    n == "xl/workbook.xml" && (bytes = Vector{UInt8}(replace(String(bytes),
                        r"(<sheet [^>]*name=\"Hidden\")" => s"\1 state=\"hidden\"",
                        r"(<sheet [^>]*name=\"VeryHidden\")" => s"\1 state=\"veryHidden\"")))
                    zip_newfile(w, n); write(w, bytes)
                end
            end
        end
        XLSX.openxlsx("states.xlsx"; mode="rw") do xf
            xf[1]["A1"] = "edited"
        end
        states = Dict(m[1] => m[2] for m in eachmatch(r"<sheet [^>]*name=\"([^\"]+)\"[^>]*state=\"([^\"]+)\"",
                      String(ZipReader(read("states.xlsx")) |> z -> zip_readentry(z, "xl/workbook.xml"))))
        @test states == Dict("Hidden" => "hidden", "VeryHidden" => "veryHidden")
        SAVE_FILES && save_outfile("states.xlsx")
    end
    @testset "dimension grows on all sides at once" begin
        f = XLSX.newxlsx(); s = f[1]
        s["B2:D10"] = 1
        XLSX.set_dimension!(s, XLSX.CellRange("B2:D10"))
        s["A20"] = 2
        @test XLSX.get_dimension(s) == XLSX.CellRange("A2:D20")
        SAVE_FILES && save_outfile(f)
    end
    @testset "rw do-block that throws leaves the file unchanged" begin
        XLSX.writetable("rw_throw.xlsx", [[1, 2]], ["a"]; overwrite=true)
        @test_throws ErrorException XLSX.openxlsx("rw_throw.xlsx"; mode="rw") do xf
            xf[1]["A2"] = "partially edited"
            error("boom")
        end
        @test XLSX.readxlsx("rw_throw.xlsx")[1]["A2"] == 1
    end
    @testset "sheet names are unique ignoring case" begin
        f = XLSX.newxlsx("Sheet1")
        @test XLSX.addsheet!(f, "SHEET1").name == "SHEET1 1"
        @test_throws XLSX.XLSXError XLSX.renamesheet!(f[2], "sheet1")
        XLSX.renamesheet!(f[1], "SHEET1")                    # a new case of its own name
        @test XLSX.sheetnames(f)[1] == "SHEET1"
        SAVE_FILES && save_outfile(f)
    end
end
