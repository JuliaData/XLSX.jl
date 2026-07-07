
SharedStringTable() = SharedStringTable(Vector{String}(), Vector{String}(), Dict{String, Int64}(), false)

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
        ctype_root = xml_root_element(xmlroot(get_xlsxfile(wb), "[Content_Types].xml"))
        XML.tag(ctype_root) != "Types" && throw(XLSXError("Something wrong here!"))
        override_node = XML.Element("Override";
            ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml",
            PartName = "/xl/sharedStrings.xml"
        )
        push!(ctype_root, override_node)
        init_sst_index(sst)
    end
end

function add_to_sst!(ss::SharedStringTable, si_xml::String)::Int64
    isempty(ss.index) && init_sst_index(ss)
    ind = get(ss.index, si_xml, nothing)
    ind !== nothing && return ind
    new_idx = length(ss.shared_strings)
    push!(ss.shared_strings, si_xml)
    push!(ss.unformatted, unformatted_text(xml_root_element(parse(si_xml, XML.LazyNode))))
    ss.index[si_xml] = new_idx
    return new_idx
end

# Adds a string to shared string table. Returns the 0-based index of the shared string in the shared string table.
function add_formatted_string!(wb::Workbook, sst::SharedStringTable, str::String)::Int64
    lock(wb.sst_lock) do
        add_to_sst!(sst, str)
    end
end
function add_formatted_string!(wb::Workbook, str_formatted::String)::Int64
    isempty(str_formatted) && throw(XLSXError("Can't add empty string to Shared String Table."))
    sst = get_sst(wb)
    
    lock(wb.sst_lock) do
        if has_sst(wb) && !sst.is_loaded
            sst_load!(wb)
        elseif !sst.is_loaded
            create_new_sst(wb, sst)
        end
        add_to_sst!(sst, str_formatted)
    end
end

# check if unformatted shared string needs xml:space="preserve"
needs_preserve(s::AbstractString) = startswith(s, ' ') || endswith(s, ' ') || contains(s, '\n')  || contains(s, "  ")

# allow to write cells containing only whitespace characters or with leading or trailing whitespace.
function add_shared_string!(wb::Workbook, str_unformatted::AbstractString) :: Int
    escaped = XML.escape(str_unformatted)
    # pfx is stable for the lifetime of the workbook — could be cached on wb
    pfx = get_prefix("xl/sharedStrings.xml", get_xlsxfile(wb))
    pfx_colon = pfx == "" ? "" : "$(pfx):"

    str_formatted = if needs_preserve(str_unformatted)
        "<$(pfx_colon)si>\n  <$(pfx_colon)t xml:space=\"preserve\">$(escaped)</$(pfx_colon)t>\n</$(pfx_colon)si>"
    else
        "<$(pfx_colon)si>\n  <$(pfx_colon)t>$(escaped)</$(pfx_colon)t>\n</$(pfx_colon)si>"
    end

    return add_formatted_string!(wb, str_formatted)
end

function _si_unformatted(child::XML.LazyNode)::String
    # Fast path: find first <t> child and try is_simple_value on it
    for t_node in XML.eachchildnode(child)
        XML.nodetype(t_node) == XML.Element || continue
        localname(t_node) == "t" || continue
        sv = XML.is_simple_value(t_node)
        isnothing(sv) || return String(sv)
        break  # <t> exists but isn't simple — fall through
    end
    # Fallback for rich text / phonetic hints / complex structure
    unformatted_text(child)
end

function sst_load!(workbook::Workbook)
    sst = get_sst(workbook)
    sst.is_loaded && return
    has_sst(workbook) || return

    xlsxfile = get_xlsxfile(workbook)
    if !internal_xml_file_exists(xlsxfile, "xl/sharedStrings.xml")
        sst.is_loaded = true
        return
    end

    doc = open_internal_file_stream(xlsxfile, "xl/sharedStrings.xml")
    empty!(sst.shared_strings)
    empty!(sst.unformatted)

    sst_root = xml_root_element(doc)
    xml_str = sst_root.data   # underlying string for zero-copy slice

    uc = XML.get(sst_root, "uniqueCount", nothing)
    if !isnothing(uc)
        n = parse(Int, uc)
        sizehint!(sst.shared_strings, n)
        sizehint!(sst.unformatted, n)
    end

    c = XML.Cursor(sst_root)
    while XML.next!(c) !== nothing
        XML.depth(c) == 1 && continue
        XML.depth(c) != 2 && (XML.skip_element!(c); continue)
        XML.nodetype(c) == XML.Element || continue
        localname(c) == "si" || (XML.skip_element!(c); continue)
        child = XML.LazyNode(c)
        si_start = c.token.offset + 1
        XML.skip_element!(c)
        si_end = c.st.state.pos
        push!(sst.shared_strings, xml_str[si_start:si_end-1])
        push!(sst.unformatted, _si_unformatted(child))
    end

    sst.is_loaded = true
end

# Checks whether this workbook has a Shared String Table.
function has_sst(workbook::Workbook) :: Bool
    relationship_type = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings"
    return has_relationship_by_type(workbook, relationship_type)
end

# Helper function to gather unformatted text from Excel data files.
# It looks at all children of `el` for tag name `t` and returns
# a join of all the strings found.
function unformatted_text(el::XML.LazyNode) :: String
    io = IOBuffer()
    gather_strings!(io, el)
    # XML.jl 0.4 `LazyNode` already unescapes entity references in `XML.value`,
    # so no extra unescape is needed here.
    return String(take!(io))
end

# 2-arg form retained for call sites that thread the workbook (e.g. inlineStr cells).
unformatted_text(::Workbook, el::XML.LazyNode) :: String = unformatted_text(el)

function gather_strings!(io::IOBuffer, e::XML.LazyNode)
    XML.nodetype(e) == XML.Element || return
    tag = localname(e)
    tag == "rPh" && return

    if tag == "t"
        for ch in XML.eachchildnode(e)
            nt = XML.nodetype(ch)
            if nt === XML.Text || nt === XML.CData
                v = XML.value(ch)
                isnothing(v) || write(io, v)
            end
        end
    else
        for ch in XML.eachchildnode(e)
            gather_strings!(io, ch)
        end
    end
end

# Looks for a string inside the Shared Strings Table (sst).
# `index` starts at 0.
@inline function sst_unformatted_string(wb::Workbook, index::Int64)::String
    sst = get_sst(wb)
    sst.is_loaded || sst_load!(wb)
    return sst.unformatted[index+1]
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

    escaped = XML.escape(run.text)

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

A rich text cell value created in Excel may have its colors defined using an Excel theme. Reading 
such a value to a RichTextString will convert the theme color to the actual RGB color in the current 
theme. If this RichTextString is subsequently written back to the same or a different cell, the 
color will be written as an RGB color and the link to the theme will be lost.

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

# Create a RichTextString from a shared string with multiple runs (or nothing if a simple text)
function getRichTextString(wb::Workbook, xml_string::String)::Union{RichTextString, Nothing}
    doc = parse(xml_string, XML.Node)
    si = xml_root_element(doc)

    # No rich text runs — plain string, return nothing.
    # (single existence check, no intermediate array of runs)
    any(c -> localname(c) == "r", XML.children(si)) || return nothing

    rts_runs = RichTextRun[]

    for run in XML.children(si)
        localname(run) == "r" || continue
        children = XML.children(run)

        t_child   = nothing
        rpr_child = nothing
        for c in children
            ln = localname(c)
            if ln == "t" && isnothing(t_child)
                t_child = c
            elseif ln == "rPr" && isnothing(rpr_child)
                rpr_child = c
            end
        end

        isnothing(t_child) && continue
        text = XML.is_simple(t_child) ? XML.simple_value(t_child) : XML.value(t_child[1])
        isempty(text) && continue

        atts = if isnothing(rpr_child)
            nothing
        else
            bold = italic = strike = under = false
            size_val = color_val = name_val = vertAlign_val = nothing

            for c in XML.children(rpr_child)
                ln = localname(c)
                if ln == "b"
                    bold = true
                elseif ln == "i"
                    italic = true
                elseif ln == "strike"
                    strike = true
                elseif ln == "u"
                    under = true
                elseif ln == "sz" && isnothing(size_val)
                    size_val = parse(Int, XML.attributes(c)["val"])
                elseif ln == "color" && isnothing(color_val)
                    color_val = resolveColor(wb, XML.attributes(c))
                elseif ln == "rFont" && isnothing(name_val)
                    name_val = XML.attributes(c)["val"]
                elseif ln == "vertAlign" && isnothing(vertAlign_val)
                    vertAlign_val = XML.attributes(c)["val"]
                end
            end

            # fixed output order, matching the original's push! sequence exactly
            pairs = Pair{Symbol, Any}[]
            bold                      && push!(pairs, :bold      => true)
            italic                    && push!(pairs, :italic    => true)
            strike                    && push!(pairs, :strike    => true)
            under                     && push!(pairs, :under     => true)
            !isnothing(size_val)      && push!(pairs, :size      => size_val)
            !isnothing(color_val)     && push!(pairs, :color     => color_val)
            !isnothing(name_val)      && push!(pairs, :name      => name_val)
            !isnothing(vertAlign_val) && push!(pairs, :vertAlign => vertAlign_val)

            isempty(pairs) ? nothing : pairs
        end

        push!(rts_runs, RichTextRun(text, atts))
    end

    isempty(rts_runs) && return nothing
    full_text = join(r.text for r in rts_runs)
    return RichTextString(full_text, rts_runs)
end