#=
Helpers shared by more than one test file. Included first by `runtests.jl`;
helpers used by a single file live in that file.
=#

# ── Saving test outputs (see `SAVE_FILES` in runtests.jl) ─────────────────────

function current_testset_label()
    ts = Test.get_testset()
    return ts === nothing ? "untitled" : strip(ts.description)
end

function sanitize_filename(s::AbstractString)
    return replace(s, r"[^A-Za-z0-9_-]" => "_")
end

const outfile_counter = Ref(0)

function save_outfile(xf::XLSX.XLSXFile)
    outfile_counter[] += 1
    label = sanitize_filename(current_testset_label())
    # Excel won't open a macro-enabled workbook saved with an `.xlsx` extension
    src_ext = xf.source isa AbstractString ? lowercase(splitext(xf.source)[2]) : ""
    ext = (src_ext in (".xlsm", ".xltm") || xf.template_type == XLSX.XLTMTemplate) ? ".xlsm" : ".xlsx"
    fname = "outfile_$(lpad(outfile_counter[], 3, '0'))_$(label)$(ext)"
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

# ── Checking table data ───────────────────────────────────────────────────────

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

# ── Patching xlsx files ───────────────────────────────────────────────────────

# Rewrite a single entry of an xlsx zip into a new file in the working directory. Used to inject XML that
# Excel itself never produces (inlineStr cells), so the `_rewrite_node` path in
# cell.jl actually gets exercised.
function patch_xlsx_entry(src::AbstractString, entry::AbstractString, patch::Function)
    r = ZipReader(read(src))
    names = zip_names(r)
    entry in names || error("no `$entry` in $src")
    dst = "patched_" * basename(src)
    ZipWriter(dst) do w
        for n in names
            bytes = zip_readentry(r, n)
            n == entry && (bytes = Vector{UInt8}(patch(String(bytes))))
            zip_newfile(w, n; compress=true)
            write(w, bytes)
        end
    end
    return dst
end

# ── Minimal workbook builder (differential tests, Read-internals) ───────────
# Compact XML, as Excel writes it. Style 0 = General, 1 = date (numFmt 14),
# 2 = a fill only (for styled-but-empty cells).

const _DIFF_CT = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/><Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/></Types>"""
const _DIFF_RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>"""
const _DIFF_WB = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>"""
const _DIFF_WBRELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/></Relationships>"""
const _DIFF_STYLES = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts><fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FFFFFF00"/></patternFill></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="3"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="14" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/><xf numFmtId="0" fontId="0" fillId="2" borderId="0" xfId="0" applyFill="1"/></cellXfs></styleSheet>"""

function _diff_build_xlsx(rows_xml::AbstractString, sst::Vector{String})::Vector{UInt8}
    sst_xml = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n" *
        "<sst xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" count=\"$(length(sst))\" uniqueCount=\"$(length(sst))\">" *
        join(isempty(s) ? "<si><t/></si>" : "<si><t xml:space=\"preserve\">$s</t></si>" for s in sst) * "</sst>"
    sheet_xml = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n" *
        "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" " *
        "xmlns:x14ac=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac\">" *
        "<sheetData>$rows_xml</sheetData></worksheet>"
    io = IOBuffer()
    ZipArchives.ZipWriter(io) do w
        ZipArchives.zip_writefile(w, "[Content_Types].xml", codeunits(_DIFF_CT))
        ZipArchives.zip_writefile(w, "_rels/.rels", codeunits(_DIFF_RELS))
        ZipArchives.zip_writefile(w, "xl/workbook.xml", codeunits(_DIFF_WB))
        ZipArchives.zip_writefile(w, "xl/_rels/workbook.xml.rels", codeunits(_DIFF_WBRELS))
        ZipArchives.zip_writefile(w, "xl/styles.xml", codeunits(_DIFF_STYLES))
        ZipArchives.zip_writefile(w, "xl/sharedStrings.xml", codeunits(sst_xml))
        ZipArchives.zip_writefile(w, "xl/worksheets/sheet1.xml", codeunits(sheet_xml))
    end
    return take!(io)
end

# ── OOXML namespaces ──────────────────────────────────────────────────────────

const _TRANS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
