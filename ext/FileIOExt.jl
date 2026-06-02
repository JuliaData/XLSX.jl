module FileIOExt
# Provides hooks for FileIO.jl to save and load XLSX files.

using FileIO

using XLSX

import XLSX: load, save

function load(f::File{FileIO.format"Excel"}; transpose::Bool=false, kw...)
    filename = FileIO.filename(f)
    if transpose
        return XLSX.readtransposedtable(filename; kw...)
    else
        return XLSX.readtable(filename; kw...)
    end
end

function load(f::File{FileIO.format"Excel"}, sheet; transpose::Bool=false, kw...)
    filename = FileIO.filename(f)
    if transpose
        return XLSX.readtransposedtable(filename, sheet; kw...)
    else
        return XLSX.readtable(filename, sheet; kw...)
    end
end

function load(f::File{FileIO.format"Excel"}, sheet, rows_or_columns; transpose::Bool=false, kw...)
    filename = FileIO.filename(f)
    if transpose
        return XLSX.readtransposedtable(filename, sheet, rows_or_columns; kw...)
    else
        return XLSX.readtable(filename, sheet, rows_or_columns; kw...)
    end
end

function save(f::File{FileIO.format"Excel"}, data; kw...)
    XLSX.writetable(FileIO.filename(f), data; kw...)
end    

end # module