
SharedStringTable() = SharedStringTable(Vector{String}(), Dict{String, Int64}(), false)

@inline get_sst(wb::Workbook) = wb.sst
@inline get_sst(xl::XLSXFile) = get_sst(get_workbook(xl))
@inline Base.length(sst::SharedStringTable) = length(sst.shared_strings)
@inline Base.isempty(sst::SharedStringTable) = isempty(sst.shared_strings)

# Checks if string is inside shared string table.
# Returns `nothing` if it's not in the shared string table.
# Returns the index of the string in the shared string table. The index is 0-based.
function get_shared_string_index(sst::SharedStringTable, str::String)# :: Union{Nothing, Int}
    !sst.is_loaded && throw(XLSXError("Can't query shared string table because it's not loaded into memory."))

    #using a Dict is much more efficient than the findfirst approach especially on large datasets
    return get(sst.index, str, nothing)
end
function create_new_sst(wb::Workbook, sst::SharedStringTable)
    if !sst.is_loaded
        sst.is_loaded = true

        # add relationship
        #<Relationship Id="rId16" Target="sharedStrings.xml" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings"/>
        add_relationship!(wb, "sharedStrings.xml", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings")

        # add Content Type <Override ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml" PartName="/xl/sharedStrings.xml"/>
        ctype_root = xmlroot(get_xlsxfile(wb), "[Content_Types].xml")[end]
        localname(ctype_root) != "Types" && throw(XLSXError("Something wrong here!"))
        override_node = XML.Element("Override";
            ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml",
            PartName = "/xl/sharedStrings.xml"
        )
        push!(ctype_root, override_node)
        init_sst_index(sst)
    end
end
function add_to_sst!(ss::SharedStringTable, si_xml::String)::Int64
    
    # Check for match
    ind = get(ss.index, si_xml, nothing)
    if ind !== nothing
        return ind  # Found exact match
    end
    
    # No match found, add new entry
    new_idx = length(ss.shared_strings)  # 0-based index
    push!(ss.shared_strings, si_xml)

    ss.index[si_xml] = new_idx

    return new_idx
end

function add_formatted_string!(sst::SharedStringTable, str::String; mylock::Union{Nothing,ReentrantLock}=nothing) :: Int64
    if isnothing(mylock)
        return add_to_sst!(sst, str)
    else
        lock(mylock) do
            return add_to_sst!(sst, str)
        end
    end
end

# Adds a string to shared string table. Returns the 0-based index of the shared string in the shared string table.
function add_formatted_string!(wb::Workbook, str_formatted::String; mylock::Union{Nothing,ReentrantLock}=nothing) :: Int64
    if isempty(str_formatted)
        throw(XLSXError("Can't add empty string to Shared String Table."))
    end
    sst = get_sst(wb)

        # if got to this point, the file was opened as template but doesn't have a Shared String Table.
        # Will create a new one.
    if !sst.is_loaded
        if isnothing(mylock)
            create_new_sst(wb, sst)
        else
            lock(mylock) do # ensure thread-safety if multiple threads are trying to add inlineStrings
                create_new_sst(wb, sst)
            end
        end
    end
    
    return add_formatted_string!(sst, str_formatted; mylock)
end

# check if unformatted shared string needs xml:space="preserve"
needs_preserve(s::String) = startswith(s, ' ') || endswith(s, ' ') || contains(s, '\n')  || contains(s, "  ")

# allow to write cells containing only whitespace characters or with leading or trailing whitespace.
function add_shared_string!(wb::Workbook, str_unformatted::AbstractString; mylock::Union{Nothing,ReentrantLock}=nothing) :: Int
    escaped = XLSX.escape(str_unformatted)

    pfx = get_prefix("xl/sharedStrings.xml", get_xlsxfile(wb))
    pfx = pfx == "" ? "" : "$(pfx):"
    
    io = IOBuffer()
    write(io, "<$(pfx)si>\n  <$(pfx)t")
    if needs_preserve(str_unformatted)
        write(io, " xml:space=\"preserve\"")
    end
    write(io, ">", escaped, "</$(pfx)t>\n</$(pfx)si>")
    str_formatted = String(take!(io))
    return add_formatted_string!(wb, str_formatted; mylock)
end

function sst_load!(workbook::Workbook)
    chunksize = ROW_CHUNKSIZE
    sst = get_sst(workbook)
    if !sst.is_loaded

        relationship_type = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings"
        if has_relationship_by_type(workbook, relationship_type)
            sst_chan = stream_ssts(open_internal_file_stream(get_xlsxfile(workbook), "xl/sharedStrings.xml")[end], chunksize)
            load_sst_table!(workbook, sst_chan, Threads.nthreads())
            init_sst_index(sst)
            
            return
        end

        throw(XLSXError("Shared Strings Table not found for this workbook."))
    end
end
@inline _is_tag(n::String, tag::String) = n == tag
@inline _is_tag(n::Nothing, tag::String) = false
 function produce_sstchunks(out, n, ssts, chunksize)
    i = 0           # Position within current chunk
    global_idx = 0  # Global position in SST table
   
    while !isnothing(n)
        if _is_tag(localname(n), "si")
            i += 1
            global_idx += 1
            ssts[i] = SstToken(n, global_idx)  # ← Use global index
        end
        if i >= chunksize
            put!(out, copy(ssts))
            i = 0  # Reset chunk position, but global_idx keeps going
        end
        n = XML.next(n)
    end
    if i > 0
        put!(out, copy(ssts[1:i]))
    end
end

function stream_ssts(n::XML.LazyNode, chunksize::Int; channel_size::Int=1 << 8)
    n = XML.next(n)
    ssts = Vector{SstToken}(undef, chunksize)
    Channel{Vector{SstToken}}(channel_size) do out
        produce_sstchunks(out, n, ssts, chunksize)
    end
end

function process_sst(wb, sst::SstToken)
    el = sst.n
    i = sst.idx

    if XML.nodetype(el) != XML.Text
        localname(el) != "si" && throw(XLSXError("Unsupported node $(localname(el)) in sst table."))
        sst = Sst(XML.write(el), i)
        return sst

    end

end

 function load_sst_table!(wb::Workbook, chan::Channel, nthreads::Int)
    sst_table = get_sst(wb)
    sst_table.is_loaded = true
    sst_results = Channel{Vector{Sst}}(1 << 8)
    all_ssts = Vector{Tuple{Int,Sst}}()
   
    consumer = @async begin
        for ssts in sst_results        
            for sst in ssts
                push!(all_ssts, (sst.idx, sst))
            end
        end    
        sort!(all_ssts, by = x -> x[1])
   
        empty!(sst_table.shared_strings)
        empty!(sst_table.index)

        for (i, sst) in all_ssts
            push!(sst_table.shared_strings, sst.formatted)
            sst_table.index[sst.formatted] = i - 1   # 0-based
        end

    end
   
    # Producer tasks
    @sync for _ in 1:nthreads
        Threads.@spawn begin
            for ssts in chan
                # ssts is already a chunk - just process it
                processed = [process_sst(wb, tok) for tok in ssts]
                put!(sst_results, processed)
            end
        end
    end
   
    close(sst_results)
    wait(consumer)
end

# Checks whether this workbook has a Shared String Table.
function has_sst(workbook::Workbook) :: Bool
    relationship_type = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings"
    return has_relationship_by_type(workbook, relationship_type)
end

# Helper function to gather unformatted text from Excel data files.
# It looks at all children of `el` for tag name `t` and returns
# a join of all the strings found.
function unformatted_text(wb::Workbook, el::XML.LazyNode) :: String
    io = IOBuffer()
    gather_strings!(wb, io, el)
    s = XLSX.unescape(String(take!(io)))
    return s
end

function gather_strings!(wb::Workbook, io::IOBuffer, e::XML.LazyNode)
    tag = localname(e)

    # Skip phonetic hints entirely
    tag == "rPh" && return nothing

    if tag == "t"
        for c in XML.children(e)
            if XML.nodetype(c) == XML.Text
                val = XML.is_simple(c) ? XML.simple_value(c) : XML.value(c)
                !isnothing(val) && write(io, val)
            end
        end

        # Fallback for truly empty <t>
        if isempty(XML.children(e))
            val = XML.is_simple(e) ? XML.simple_value(e) : XML.value(e)
            !isnothing(val) && write(io, val)
        end
    else
        for ch in XML.children(e)
            gather_strings!(wb, io, ch)
        end
    end

    return nothing
end

# Looks for a string inside the Shared Strings Table (sst).
# `index` starts at 0.
@inline function sst_unformatted_string(wb::Workbook, index::Int64)::String
    sst_load!(wb)
    uss = get_sst(wb).shared_strings[index+1]
    return unformatted_text(wb, parse(XML.LazyNode, uss))
end

@inline sst_unformatted_string(xl::XLSXFile, index::Int64) :: String = sst_unformatted_string(get_workbook(xl), index)
@inline sst_unformatted_string(ws::Worksheet, index::Int64) :: String = sst_unformatted_string(get_xlsxfile(ws), index)
@inline sst_unformatted_string(target::Union{Workbook, XLSXFile, Worksheet}, index_str::String) :: String = sst_unformatted_string(target, parse(Int64, index_str))

# init the index table
function init_sst_index(sst::SharedStringTable)
    empty!(sst.index)
    for i in 1:length(sst.shared_strings)
       sst.index[sst.shared_strings[i]] = i - 1
    end
end

#=
This is the required order of attributes in the xml:
- <rFont> — Font name
- <charset> — Character set # Not required/used here
- <family> — Font family # Not required/used here
- <b> — Bold
- <i> — Italic
- <strike> — Strikethrough
- <outline>
- <shadow>
- <condense>
- <extend>
- <color> — Color
- <sz> — Font size
- <u> — Underline
- <vertAlign> — Superscript/subscript
- <scheme> — Font scheme (major/minor) # Not required/used here
=#

 """
    richTextRunToXML!(io::IO, run::RichTextRun, pfx::String) -> IO

Convert an RichTextRun to XML format for Excel shared strings.
Each rich text shared string may have multiple runs to allow 
heterogeneous formatting within a single cell.
"""
function richTextRunToXML!(io::IO, run::RichTextRun, pfx)
    write(io, "<$(pfx)r>")

    atts = run.atts
    if !isnothing(atts)
        props = IOBuffer()

        if (v = get(atts, :name, nothing)) !== nothing
            write(props, "<$(pfx)rFont val=\"", v, "\"/>")
        end
        if get(atts, :bold, false)  in (true, 1)
            write(props, "<$(pfx)b/>")
        end
        if get(atts, :italic, false)  in (true, 1)
            write(props, "<$(pfx)i/>")
        end
        if get(atts, :strike, false)  in (true, 1)
            write(props, "<$(pfx)strike/>")
        end
        if (v = get(atts, :color, nothing)) !== nothing
            write(props, "<$(pfx)color rgb=\"", get_color(v), "\"/>")
        end
        if (v = get(atts, :size, nothing)) !== nothing
            write(props, "<$(pfx)sz val=\"", string(v), "\"/>") # size read as a float, output rounded to nearest half point.
        end
        if get(atts, :under, false)  in (true, 1)
            write(props, "<$(pfx)u/>")
        end
        if (v = get(atts, :vertAlign, nothing)) !== nothing
            write(props, "<$(pfx)vertAlign val=\"", v, "\"/>")
        end

        if position(props) > 0
            write(io, "<$(pfx)rPr>")
            write(io, take!(props))
            write(io, "</$(pfx)rPr>")
        end
    end

    needs_preserve =
        startswith(run.text, " ") ||
        endswith(run.text, " ") ||
        contains(run.text, '\n') ||
        contains(run.text, "  ")

    escaped = XLSX.escape(run.text)

    if needs_preserve
        write(io, "<$(pfx)t xml:space=\"preserve\">", escaped, "</$(pfx)t>")
    else
        write(io, "<$(pfx)t>", escaped, "</$(pfx)t>")
    end

    write(io, "</$(pfx)r>")
    return nothing
end

function richTextStringtoXML(rts::RichTextString, pfx::String)
    xml = IOBuffer()
    write(xml, "<$(pfx)si>")
    for r in rts.runs
        richTextRunToXML!(xml, r, pfx)
    end
    write(xml, "</$(pfx)si>")
    return String(take!(xml))
end
function RichTextString(runs::Vector{RichTextRun})
    isempty(runs) && throw(XLSXError("Cannot create a RichTextString with no RichTextRuns"))
    t = join((x.text for x in runs))
    return RichTextString(t, runs)
end

Base.:(==)(rts1::RichTextString, rts2::RichTextString) = rts1.text == rts2.text && length(rts1.runs) == length(rts2.runs) && all(==(true), rts1.runs .== rts2.runs)
Base.hash(rts::RichTextString, h::UInt) = hash(rts.runs, hash(rts.text, h))
Base.length(rts::RichTextString) = length(rts.text)
Base.iterate(rts::RichTextString, i::Integer=firstindex(rts.text)) = Base.iterate(rts.text, i)
Base.ncodeunits(rts::RichTextString) = ncodeunits(rts.text)
Base.codeunit(rts::RichTextString) = codeunit(rts.text)
Base.codeunit(rts::RichTextString, i::Integer) = codeunit(rts.text, i)
Base.String(rts::RichTextString) = rts.text
Base.isvalid(rts::RichTextString, i::Integer) = isvalid(rts.text, i)

function Base.show(io::IO, rts::RichTextString)
    maxlen_txt = 22
    maxlen_atts = 64

    print(io, "RichTextString: \"$(rts.text)\" \n containing $(length(rts.runs)) runs:\n")
    @printf(io, " %-24s %-14s\n", "Run text", "Run attributes")
    println(io, " "*"-"^(24 + 1 + 66))
    for run in rts.runs

        if isnothing(run.atts)
            s=" "
        else
            # Convert pairs to "key=value"
            parts = (":"*string(k) * " => " * sprint(show, v) for (k,v) in sort(collect(run.atts), by=first))

            s = join(parts, ", ")

            # Truncate if too long
            if length(s) > maxlen_atts
                s = s[1:prevind(s, maxlen_atts)] * "…"
            end
        end

        t = if length(run.text) > maxlen_txt
            run.text[1:prevind(run.text, maxlen_txt)] * "…"
        else
            run.text
        end

        @printf(io, " %-24s %-66s\n", "\""*t*"\"", "["*s*"]")
    end
end

# Concatenate two RichTextStrings into a single RichTextString
function Base.:*(s1::RichTextString, s2::RichTextString)::RichTextString
    RichTextString(s1.text * s2.text, vcat(s1.runs, s2.runs))
end

# Take a substring from a RichTextString as another RichTextString
Base.getindex(rts::RichTextString, i::Int) = getindex(rts, i:i)
function Base.getindex(rts::RichTextString, r::UnitRange{Int})::RichTextString
    substr = rts.text[r]

    new_runs = RichTextRun[]
    global_char_pos = 1

    for run in rts.runs
        run_text = run.text
        run_chars = length(run_text)
        run_range = global_char_pos:(global_char_pos + run_chars - 1)

        overlap_start = max(first(r), first(run_range))
        overlap_end   = min(last(r), last(run_range))

        if overlap_start <= overlap_end
            # Convert global character indices to local character indices
            local_start = overlap_start - global_char_pos + 1
            local_end   = overlap_end   - global_char_pos + 1

            # Slice using character indexing
            sliced_text = run_text[local_start:local_end]

            push!(new_runs, RichTextRun(sliced_text, run.atts))
        end

        global_char_pos += run_chars
    end

    return RichTextString(substr, new_runs)
end

function Base.:(==)(r1::RichTextRun, r2::RichTextRun)
    r1.text == r2.text || return false

    # Handle Nothing vs Dict
    a1 = r1.atts
    a2 = r2.atts
    a1 === a2 && return true
    (a1 === nothing || a2 === nothing) && return false

    # Same keys?
    keys(a1) == keys(a2) || return false

    # Compare values with special color handling
    for k in keys(a1)
        v1 = a1[k]
        v2 = a2[k]
        if k === :color
            get_color(v1) == get_color(v2) || return false
        else
            v1 == v2 || return false
        end
    end

    return true
end
function Base.hash(r::RichTextRun, h::UInt)
    # Hash the run text first
    h = hash(r.text, h)
    atts = r.atts

    # No attributes
    if atts === nothing
        return hash(0xdebdceef, h)  # any constant
    end

    # Order‑insensitive hash over attributes
    # (color values canonicalized via get_color)
    for (k, v) in atts
        v′ = (k === :color ? get_color(v) : v)
        h = hash((k, v′), h)
    end
    return h
end
Base.length(run::RichTextRun) = length(run.text)
function Base.show(io::IO, run::RichTextRun)
    maxlen_txt = 22
    maxlen_atts = 66

    # Convert pairs to "key=value"
    if isnothing(run.atts)
        s = " "
    else
        parts = (":"*string(k) * " => " * sprint(show, v) for (k,v) in sort(collect(run.atts), by=first))
        s = join(parts, ", ")

        # Truncate if too long
        if length(s) > maxlen_atts
            s = s[1:prevind(s, maxlen_atts)] * "…"
        end
    end

    t = if length(run.text) > maxlen_txt
        run.text[1:prevind(run.text, maxlen_txt)] * "…"
    else
        run.text
    end

    print(io,"RichTextRun (","\""*t*"\"  [", s,"])")
end

RichTextString(runs::RichTextRun...) = RichTextString(collect(runs))

# Normalize any attribute input into Vector{Pair{Symbol,Any}}
_to_pairs(::Nothing) = nothing

_to_pairs(nt::NamedTuple) =
    [Pair{Symbol,Any}(k, v) for (k, v) in pairs(nt)]

_to_pairs(p::Pair) =
    [Pair{Symbol,Any}(p.first, p.second)]

_to_pairs(v::Vector{<:Pair}) =
    [Pair{Symbol,Any}(p.first, p.second) for p in v]

_to_pairs(d::Dict{Symbol,Any}) =
    [Pair{Symbol,Any}(k, v) for (k, v) in d]

_to_pairs(t::Tuple{Vararg{Pair}}) =
    [Pair{Symbol,Any}(p.first, p.second) for p in t]

RichTextRun(text::String, atts) = RichTextRun(text, _to_pairs(atts))

"""
    getRichTextString(ws::Worksheet, cr::String)                 -> Union{RichTextString, Nothing}
    getRichTextString(xl::XLSXFile, sheetcell::String)           -> Union{RichTextString, Nothing}
    getRichTextString(ws::Worksheet, row::Integer, col::Integer) -> Union{RichTextString, Nothing}

Create a RichTextString object from a cell value.

If the cell value is not a string, return nothing.

If the cell contains simple text with no rich text formatting, return nothing.

# Examples
```julia

julia> rtf1=XLSX.RichTextRun("Hello", [:vertAlign => "subscript"])
RichTextRun ("Hello"  [:vertAlign => "subscript"] )

julia> rtf2=XLSX.RichTextRun(" Kitty ", [:color => "green", :size => 14, :bold => true, :under => true])
RichTextRun (" Kitty "  [:color => "green", :bold => true, :size => 14, :under => true] )

julia> rtf3=XLSX.RichTextRun("Hello", [:color => "green", :size => 14, :under => true])
RichTextRun ("Hello"  [:color => "green", :size => 14, :under => true] )

julia> r = XLSX.RichTextString(rtf1, rtf2, rtf3)
RichTextString: "Hello Kitty Hello" 
 containing 3 runs:
 Run text                 Run attributes
 -------------------------------------------------------------------------------------------
 "Hello"                  [:vertAlign => "subscript"]
 " Kitty "                [:color => "green", :bold => true, :size => 14, :under => true]
 "Hello"                  [:color => "green", :size => 14, :under => true]

julia> f = newxlsx()
XLSXFile("blank.xlsx") containing 1 Worksheet
            sheetname size          range
-------------------------------------------------
               Sheet1 1x1           A1:A1


julia> s=f[1]
1×1 Worksheet: ["Sheet1"](A1:A1)

julia> s["A1"] = r
RichTextString: "Hello Kitty Hello" 
 containing 3 runs:
 Run text                 Run attributes
 -------------------------------------------------------------------------------------------
 "Hello"                  [:vertAlign => "subscript"]
 " Kitty "                [:color => "green", :bold => true, :size => 14, :under => true]
 "Hello"                  [:color => "green", :size => 14, :under => true]

julia> s["A2"] = styled"The {bold:{italic:quick {(foreground=#cd853f):brown} fox} jumps over the {(foreground=#FFC000):lazy} dog}"
"The quick brown fox jumps over the lazy dog"

julia> XLSX.getRichTextString(s, "A2")
RichTextString: "The quick brown fox jumps over the lazy dog" 
 containing 7 runs:
 Run text                 Run attributes
 -------------------------------------------------------------------------------------------
 "The "                   [:size => 12.0]
 "quick "                 [:bold => true, :italic => true, :size => 12.0]
 "brown"                  [:bold => true, :color => "FFCD853F", :italic => true, :size => …]
 " fox"                   [:bold => true, :italic => true, :size => 12.0]
 " jumps over the "       [:bold => true, :size => 12.0]
 "lazy"                   [:bold => true, :color => "FFFFC000", :size => 12.0]
 " dog"                   [:bold => true, :size => 12.0]

julia> a = XLSX.getRichTextString(s, "A1")
RichTextString: "Hello Kitty Hello" 
 containing 3 runs:
 Run text                 Run attributes
 -------------------------------------------------------------------------------------------
 "Hello"                  [:vertAlign => "subscript"]
 " Kitty "                [:bold => true, :color => "FF008000", :size => 14.0, :under => t…]
 "Hello"                  [:color => "FF008000", :size => 14.0, :under => true]
```

When they are written to a cell, named colors are converted to RGB values for Excel. However, 
two RichTextStrings will be considered equal regardless of this representation so long as the 
colors are identical. So:

```julia
julia> a == r
true
```

"""
getRichTextString(ws::Worksheet, cr::String) = process_get_cellname(getRichTextString, ws, cr)
getRichTextString(xl::XLSXFile, sheetcell::String) = process_get_sheetcell(getRichTextString, xl, sheetcell)
getRichTextString(ws::Worksheet, row::Integer, col::Integer) = getRichTextString(ws, CellRef(row, col))
function getRichTextString(s::Worksheet, c::CellRef)::Union{RichTextString, Nothing}
    cell = getcell(s, c)
    cell.datatype == CT_STRING || return nothing
    sst_load!(get_workbook(s))
    uss = get_sst(get_workbook(s)).shared_strings[reinterpret(Int64, cell.value)+1]
    return getRichTextString(get_workbook(s), uss)
end

const INDEXED_PALETTE = [
    "000000", "FFFFFF", "FF0000", "00FF00", "0000FF", "FFFF00", "FF00FF", "00FFFF",
    "000000", "FFFFFF", "FF0000", "00FF00", "0000FF", "FFFF00", "FF00FF", "00FFFF",
    "800000", "008000", "000080", "808000", "800080", "008080", "C0C0C0", "808080",
    "9999FF", "993366", "FFFFCC", "CCFFFF", "660066", "FF8080", "0066CC", "CCCCFF",
    "000080", "FF00FF", "FFFF00", "00FFFF", "800080", "800000", "008080", "0000FF",
    "00CCFF", "CCFFFF", "CCFFCC", "FFFF99", "99CCFF", "FF99CC", "CC99FF", "FFCC99",
    "3366FF", "33CCCC", "99CC00", "FFCC00", "FF9900", "FF6600", "666699", "969696",
    "003366", "339966", "003300", "333300", "993300", "993366", "333399", "333333"
]

# Excel tint algorithm
@inline function apply_tint(channel::UInt8, tint::Float64)::UInt8
    c = Float64(channel)
    if tint > 0
        c = c + (255 - c) * tint
    else
        c = c * (1 + tint)
    end
    return UInt8(clamp(round(Int, c), 0, 255))
end

# Convert theme + tint to RGB
function resolve_theme_color(theme_index::Int, tint::Float64)
    # Default Excel theme colors - assume these are never customised.
     theme = [
    0x000000, 0xFFFFFF, 0x1F497D, 0xEEECE1,
    0x4F81BD, 0xC0504D, 0x9BBB59, 0x8064A2,
    0x4BACC6, 0xF79646,
    0x0000FF,  # hyperlink
    0x800080   # followed hyperlink
]

    base = theme[theme_index + 1]
    r = apply_tint(UInt8(base >> 16), tint)
    g = apply_tint(UInt8((base >> 8) & 0xFF), tint)
    b = apply_tint(UInt8(base & 0xFF), tint)

   buf = IOBuffer()
    print(buf, "FF")
    print(buf, uppercase(string(r, base=16, pad=2)))
    print(buf, uppercase(string(g, base=16, pad=2)))
    print(buf, uppercase(string(b, base=16, pad=2)))
    return String(take!(buf))
end

# Create a RichTextString from a shared string with multiple runs (or nothing if a simple text)
function getRichTextString(wb::Workbook, xml_string::String)::Union{RichTextString, Nothing}
    doc = parse(XML.Node, xml_string)
    si = doc[end]
    
    # Check for rich text runs <r> elements
    runs = [child for child in XML.children(si) if localname(child) == "r"]
    
    # No rich text runs — plain string, return nothing
    isempty(runs) && return nothing
    
    rts_runs = RichTextRun[]
    
    for run in runs
        children = XML.children(run)
        
        t_node = findfirst(c -> localname(c) == "t", children)
        isnothing(t_node) && continue

        text = XML.is_simple(children[t_node]) ? XML.simple_value(children[t_node]) : XML.value(children[t_node][1])
        isempty(text) && continue
        
        rpr = findfirst(c -> localname(c) == "rPr", children)
        atts = if isnothing(rpr)
            nothing
        else
            rpr_node = children[rpr]
            rpr_children = XML.children(rpr_node)
            pairs = Pair{Symbol, Any}[]
            
            any(c -> localname(c) == "b",      rpr_children) && push!(pairs, :bold      => true)
            any(c -> localname(c) == "i",      rpr_children) && push!(pairs, :italic    => true)
            any(c -> localname(c) == "strike", rpr_children) && push!(pairs, :strike    => true)
            any(c -> localname(c) == "u",      rpr_children) && push!(pairs, :under     => true)
            
            sz_node = findfirst(c -> localname(c) == "sz", rpr_children)
            !isnothing(sz_node) && push!(pairs, :size => parse(Int, XML.attributes(rpr_children[sz_node])["val"]))
            
            color_node = findfirst(c -> localname(c) == "color", rpr_children)
            if !isnothing(color_node)
                atts = XML.attributes(rpr_children[color_node])

                if haskey(atts, "rgb")
                    push!(pairs, :color => atts["rgb"])
                elseif haskey(atts, "theme")
                    theme = parse(Int, atts["theme"])
                    tint  = haskey(atts, "tint") ? parse(Float64, atts["tint"]) : 0.0
                    rgb = resolve_theme_color(theme, tint)
                    push!(pairs, :color => rgb)
                elseif haskey(atts, "indexed")
                    idx = parse(Int, atts["indexed"])
                    idx = clamp(idx, 0, length(INDEXED_PALETTE)-1)
                    rgb = INDEXED_PALETTE[idx + 1]
                    push!(pairs, :color => "FF" * rgb)
                elseif haskey(atts, "auto")
                    push!(pairs, :color => "000000")  # Excel default
                end
            end            
            font_node = findfirst(c -> localname(c) == "rFont", rpr_children)
            !isnothing(font_node) && push!(pairs, :name => XML.attributes(rpr_children[font_node])["val"])
            
            va_node = findfirst(c -> localname(c) == "vertAlign", rpr_children)
            !isnothing(va_node) && push!(pairs, :vertAlign => XML.attributes(rpr_children[va_node])["val"])
            
            isempty(pairs) ? nothing : pairs
        end
        
        push!(rts_runs, RichTextRun(text, atts))
    end
    
    isempty(rts_runs) && return nothing
    full_text = join(r.text for r in rts_runs)
    return RichTextString(full_text, rts_runs)
end