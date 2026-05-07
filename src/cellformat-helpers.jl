
const font_tags = ["b", "i", "u", "strike", "outline", "shadow", "condense", "extend", "sz", "color", "name", "vertAlign", "scheme"]
const border_tags = ["left", "right", "top", "bottom", "diagonal"]
const fill_tags = ["patternFill"]
const builtinFormats = Dict(
    "0" => "General",
    "1" => "0",
    "2" => "0.00",
    "3" => "#,##0",
    "4" => "#,##0.00",
    "5" => "\$#,##0_);(\$#,##0)",
    "6" => "\$#,##0_);Red",
    "7" => "\$#,##0.00_);(\$#,##0.00)",
    "8" => "\$#,##0.00_);Red",
    "9" => "0%",
    "10" => "0.00%",
    "11" => "0.00E+00",
    "12" => "# ?/?",
    "13" => "# ??/??",
    "14" => "m/d/yyyy",
    "15" => "d-mmm-yy",
    "16" => "d-mmm",
    "17" => "mmm-yy",
    "18" => "h:mm AM/PM",
    "19" => "h:mm:ss AM/PM",
    "20" => "h:mm",
    "21" => "h:mm:ss",
    "22" => "m/d/yyyy h:mm",
    "37" => "#,##0_);(#,##0)",
    "38" => "#,##0_);Red",
    "39" => "#,##0.00_);(#,##0.00)",
    "40" => "#,##0.00_);Red",
    "45" => "mm:ss",
    "46" => "[h]:mm:ss",
    "47" => "mmss.0",
    "48" => "##0.0E+0",
    "49" => "@"
)
const builtinFormatNames = Dict(
    "General" => 0,
    "Number" => 2,
    "Currency" => 7,
    "Percentage" => 9,
    "ShortDate" => 14,
    "LongDate" => 15,
    "Time" => 21,
    "Scientific" => 48
)

# Regex fragments for canonical tokens (from Claude)
const LITERAL      = raw"\"[^\"]*\""       # quoted text (check first)
const CONDITION    = raw"\[[<>=].+?\]"
const COLOR        = raw"\[[A-Za-z]+\]"
const DATETIME     = raw"(?:AM/PM|A/P|am/pm|a/p|d{1,4}|m{1,5}|y{2,4}|h{1,2}|s{1,2})"
const DECIMAL      = raw"\.[0#?]+"         
const EXPONENT     = raw"[0#?]+[eE][+-]?[0#?]+"
const FRACTION     = raw"\?+/\?+"          # fraction with multiple ?
const PERCENT      = raw"%"
const ESCAPE       = raw"\\."              
const ALIGN        = raw"_."               
const FILL         = raw"\*."              
const TEXTPLACE    = raw"@"
const DIGIT        = raw"[0#?]"            # single digit placeholder
const COMMA        = raw","                # thousand separator
const PAREN        = raw"[\(\)]"
const COLON        = raw":"                # time separator
const SPACE        = raw" +"               # one or more spaces
const DASH         = raw"-"                # minus/dash
const CURRENCY     = raw"[\$£€¥₹]"
const PLUS         = raw"\+"

# Combine into a master regex - ORDER MATTERS!
const RGX_FMT = Regex(
    join([
        LITERAL, CONDITION, COLOR, DATETIME,
        DECIMAL, EXPONENT, FRACTION, PERCENT,
        ESCAPE, ALIGN, FILL, TEXTPLACE,
        DIGIT, COMMA, COLON, DASH, PAREN, SPACE, 
        CURRENCY, PLUS
    ], "|")
)

const VALID_FILL_PATTERNS = (
    "none", "solid", "mediumGray", "darkGray", "lightGray",
    "darkHorizontal", "darkVertical", "darkDown", "darkUp", "darkGrid", "darkTrellis",
    "lightHorizontal", "lightVertical", "lightDown", "lightUp", "lightGrid", "lightTrellis",
    "gray125", "gray0625"
)

const VALID_BORDER_STYLES = (
    "none", "thin", "medium", "dashed", "dotted", "thick", "double", "hair",
    "mediumDashed", "dashDot", "mediumDashDot", "dashDotDot", "mediumDashDotDot", "slantDashDot"
)
const VALID_DIAGONAL_DIRECTIONS = ("up", "down", "both")

# Shared kwarg type alias for border sides
const BorderKw = Union{Nothing,Vector{Pair{String,String}}}

const VALID_HORIZONTAL_ALIGNMENTS = ("left", "center", "right", "fill", "justify", "centerContinuous", "distributed")
const VALID_VERTICAL_ALIGNMENTS   = ("top", "center", "bottom", "justify", "distributed")

const BUILTIN_NUMFMT_RANGES = (0:22, 37:40, 45:49)

# Reverse lookup: format code string -> built-in numFmtId integer.
# Avoids linear scan of builtinFormats on every setFormat call.
const BUILTIN_FORMAT_CODES = Dict{String,Int}(v => parse(Int, k) for (k, v) in builtinFormats)

# Excel adds this padding to any user-specified column width
const EXCEL_COLUMN_WIDTH_PADDING = 0.7109375

#
# -- A bunch of helper functions ...
#

function copynode(o::XML.Node)
    n = XML.Node(o.nodetype, o.tag, o.attributes, o.value, isnothing(o.children) ? nothing : [copynode(x) for x in o.children])
    return n
end
function do_sheet_names_match(ws::Worksheet, rng::T) where {T<:Union{SheetCellRef,AbstractSheetCellRange}}
    if ws.name == rng.sheet
        return true
    else
        throw(XLSXError("Worksheet `$(ws.name)` does not match sheet in cell reference: `$(rng.sheet)`"))
    end
end

function make_child_node(tag::String, name::String, pfx::String)::XML.Node
    children = tag ∈ ("border", "fill") ? Vector{XML.Node}() : nothing
    return XML.Node(XML.Element, pfx*name, OrderedDict{String,String}(), nothing, children)
end

function build_font_child!(new_node::XML.Node, tag::String, name::String, attrs::Union{Nothing,Dict{String,String}}, pfx::String)
    cnode = isnothing(attrs) ? XML.Element(pfx*name) : make_child_node(tag, name, pfx)
    if !isnothing(attrs)
        for (k, v) in attrs
            cnode[k] = v
        end
    end
    push!(new_node, cnode)
end

function build_border_child!(new_node::XML.Node, tag::String, name::String, attrs::Union{Nothing,Dict{String,String}}, pfx::String)
    cnode = isnothing(attrs) ? XML.Element(pfx*name) : make_child_node(tag, name, pfx)
    if !isnothing(attrs)
        color = XML.Element("$(pfx)color")
        for (k, v) in attrs
            if k == "style" && v != "none"
                cnode[k] = v
            elseif k == "direction"
                v ∈ ("up",   "both") && (new_node["diagonalUp"]   = "1")
                v ∈ ("down", "both") && (new_node["diagonalDown"] = "1")
            else
                color[k] = v
            end
        end
        !isempty(XML.attributes(color)) && push!(cnode, color)
    end
    push!(new_node, cnode)
end

function build_fill_child!(new_node::XML.Node, tag::String, name::String, attrs::Union{Nothing,Dict{String,String}}, pfx)
    if isnothing(attrs)
        push!(new_node, XML.Element(pfx*name))
        return
    end
    patternfill = XML.Element("$(pfx)patternFill")
    fgcolor     = XML.Element("$(pfx)fgColor")
    bgcolor     = XML.Element("$(pfx)bgColor")
    for (k, v) in attrs
        if k == "patternType"
            patternfill[k] = v
        elseif startswith(k, "fg")
            fgcolor[k[3:end]] = v
        elseif startswith(k, "bg")
            bgcolor[k[3:end]] = v
        end
    end
    haskey(patternfill, "patternType") || throw(XLSXError("No `patternType` attribute found."))
    !isempty(XML.attributes(fgcolor)) && push!(patternfill, fgcolor)
    !isempty(XML.attributes(bgcolor)) && push!(patternfill, bgcolor)
    push!(new_node, patternfill)  # patternfill goes directly onto new_node
end

function buildNode(tag::String, attributes::Dict{String,Union{Nothing,Dict{String,String}}}, pfx::String)::XML.Node
    attribute_tags, build_child! = if tag == "font"
        font_tags,   build_font_child!
    elseif tag == "border"
        border_tags, build_border_child!
    elseif tag == "fill"
        fill_tags,   build_fill_child!
    else
        throw(XLSXError("Unknown tag: $tag"))
    end

    new_node = XML.Element(pfx*tag)
    for name in attribute_tags
        haskey(attributes, name) && build_child!(new_node, tag, name, attributes[name], pfx)
    end
    return new_node
end

#=function buildNode(tag::String, attributes::Dict{String,Union{Nothing,Dict{String,String}}})::XML.Node
    if tag == "font"
        attribute_tags = font_tags
    elseif tag == "border"
        attribute_tags = border_tags
    elseif tag == "fill"
        attribute_tags = fill_tags
    else
        throw(XLSXError("Unknown tag: $tag"))
    end
    new_node = XML.Element(tag)
    for a in attribute_tags # Use this as a device to keep ordering constant for Excel
        if tag == "font"
            if haskey(attributes, a)
                if isnothing(attributes[a])
                    cnode = XML.Element(a)
                else
                    cnode = XML.Node(XML.Element, a, OrderedDict{String,String}(), nothing, tag ∈ ["border", "fill"] ? Vector{XML.Node}() : nothing)
                    for (k, v) in attributes[a]
                        cnode[k] = v
                    end
                end
                push!(new_node, cnode)
            end
        elseif tag == "border"
            if haskey(attributes, a)
                if isnothing(attributes[a])
                    cnode = XML.Element(a)
                else
                    cnode = XML.Node(XML.Element, a, OrderedDict{String,String}(), nothing, tag ∈ ["border", "fill"] ? Vector{XML.Node}() : nothing)
                    color = XML.Element("color")
                    for (k, v) in attributes[a]
                        if k == "style" && v != "none"
                            cnode[k] = v
                        elseif k == "direction"
                            if v in ["up", "both"]
                                new_node["diagonalUp"] = "1"
                            end
                            if v in ["down", "both"]
                                new_node["diagonalDown"] = "1"
                            end
                        else
                            color[k] = v
                        end
                    end
                    if length(XML.attributes(color)) > 0 # Don't push an empty color.
                        push!(cnode, color)
                    end
                end
                push!(new_node, cnode)
            end
        elseif tag == "fill"
            if haskey(attributes, a)
                if isnothing(attributes[a])
                    cnode = XML.Element(a)
                else
                    cnode = XML.Node(XML.Element, a, OrderedDict{String,String}(), nothing, tag ∈ ["border", "fill"] ? Vector{XML.Node}() : nothing)
                    patternfill = XML.Element("patternFill")
                    fgcolor = XML.Element("fgColor")
                    bgcolor = XML.Element("bgColor")
                    for (k, v) in attributes[a]
                        if k == "patternType"
                            patternfill[k] = v
                        elseif first(k, 2) == "fg"
                            fgcolor[k[3:end]] = v
                        elseif first(k, 2) == "bg"
                            bgcolor[k[3:end]] = v
                        end
                    end
                    if !haskey(patternfill, "patternType")
                        throw(XLSXError("No `patternType` attribute found."))
                    end
                    length(XML.attributes(fgcolor)) > 0 && push!(patternfill, fgcolor)
                    length(XML.attributes(bgcolor)) > 0 && push!(patternfill, bgcolor)
                end
                push!(new_node, patternfill)
            end
            #else
        end
    end
    return new_node
end
=#
function isInDim(ws::Worksheet, dim::CellRange, rng::CellRange)
    if !issubset(rng, dim)
        throw(XLSXError("Cell range $rng is out of bounds. Worksheet `$(ws.name)` only has dimension `$dim`."))
    end
    return true
end
function isInDim(ws::Worksheet, dim::CellRange, row, col)
    if maximum(row) > dim.stop.row_number || minimum(row) < dim.start.row_number
        throw(XLSXError("Row range $row is out of bounds. Worksheet `$(ws.name)` only has dimension `$dim`."))
    end
    if maximum(col) > dim.stop.column_number || minimum(col) < dim.start.column_number
        throw(XLSXError("Column range $col is out of bounds. Worksheet `$(ws.name)` only has dimension `$dim`."))
    end
    return true
end

_xpath(parts...) = join(["/$SPREADSHEET_NAMESPACE_XPATH_ARG:$p" for p in parts], "")

# Merges a boolean/flag font tag (e.g. "b", "i", "strike").
# Returns nothing (present, no attributes) if kept/set, or missing to omit.
function _merge_flag_tag(tag::String, new_val::Union{Nothing,Bool}, old_atts::Dict)
    (isnothing(new_val) ? haskey(old_atts, tag) : new_val) ? nothing : missing
end

# Merges a single-value font tag (e.g. "sz", "name", "vertAlign").
# Returns Dict("val" => ...) if set, preserves old value, or missing to omit.
function _merge_val_tag(tag::String, new_val, old_atts::Dict)
    isnothing(new_val) ? get(old_atts, tag, missing) : Dict("val" => string(new_val))
end

# Merges a dict-valued tag (e.g. "color") using a provided constructor for new values.
# Returns a Dict of attributes if set/preserved, or missing to omit.
function _merge_dict_tag(tag::String, new_val, old_atts::Dict, build_fn::Function)
    isnothing(new_val) ? get(old_atts, tag, missing) : build_fn(new_val)
end

"""
    is_valid_format(fmt::AbstractString) -> Bool

Check if `fmt` is a syntactically valid Excel number format string.
"""
function is_valid_format(fmt::AbstractString) # From Claude
    # Split into up to 4 sections
    sections = split(fmt, ';')
    length(sections) > 4 && return false

    for sec in sections
        pos = 1
        while pos <= lastindex(sec)
            # Use SubString to match from current position
            m = match(RGX_FMT, SubString(sec, pos))

            # No token matches at this position
            if m === nothing
                return false
            end

            # Token must start at beginning of substring (offset should be 1)
            if m.offset != 1
                return false
            end

            # Zero-length matches are invalid (avoid infinite loops)
            tok = m.match
            if isempty(tok)
                return false
            end

            # Advance by the number of characters in the match
            pos = nextind(sec, pos, length(tok))
        end
    end

    return true
end

function first2_after_colon(tag::AbstractString)
    parts = split(tag, ':', limit=2)
    s = length(parts) == 1 ? parts[1] : parts[2]
    chars = collect(s)
    return join(chars[1:min(2, length(chars))])
end

# Parses a patternFill XML node into a flat dict of fill attributes.
# patternType is stored directly, fg/bg color attributes are prefixed with "fg"/"bg".
function _parse_pattern_fill(pattern::XML.Node)::Dict{String,String}
    atts = Dict{String,String}()
    a = XML.attributes(pattern)
    if !isnothing(a)
        for (k, v) in a
            atts[k] = v  # e.g. "patternType" => "solid"
        end
    end
    for subc in XML.children(pattern)
        XML.nodetype(subc) == XML.Element || continue
        tag_prefix = first2_after_colon(localname(subc))  # "fg" or "bg"
        sub_atts = XML.attributes(subc)
        if isnothing(sub_atts) || isempty(sub_atts)
            throw(XLSXError("Expected attributes on fill sub-element <$(localname(subc))>, found none."))
        end
        for (k, v) in sub_atts
            atts[tag_prefix * k] = v  # e.g. "fgrgb" => "FFFF0000"
        end
    end
    return atts
end

# Dispatch boilerplate
# Merges a single alignment attribute into atts.
# Preserves the old value if new_val is nothing, otherwise applies convert_fn to new_val.
function _merge_alignment_att(
    atts::AbstractDict, xml_key::String, new_val,
    old_atts::Dict{String,String}, convert_fn=string
)
    if isnothing(new_val)
        haskey(old_atts, xml_key) && (atts[xml_key] = old_atts[xml_key])
    else
        atts[xml_key] = convert_fn(new_val)
    end
end

# Coerces a Vector{Pair{String,String}} kwarg to Dict{String,String}, or returns nothing.
_to_border_dict(v) = isnothing(v) ? nothing : Dict{String,String}(p for p in v)

# Validates that `outside` is not combined with any other border kwargs.
function _check_outside_conflict(outside, left, right, top, bottom, diagonal, allsides)
    !isnothing(outside) && !all(isnothing, [left, right, top, bottom, diagonal, allsides]) &&
        throw(XLSXError("Keyword `outside` is incompatible with any other keywords."))
end

# Merges new and old attributes for a single border side.
# Handles style, color, and (for diagonal) direction.
function _merge_border_side(
    side::String,
    new_val::Union{Nothing,Dict{String,String}},
    old_atts::Union{Nothing,Dict{String,Union{Dict{String,String},Nothing}}}
)::Dict{String,String}

    result = Dict{String,String}()
    old_side = (!isnothing(old_atts) && haskey(old_atts, side)) ? old_atts[side] : nothing

    # If no new value, preserve old side entirely
    isnothing(new_val) && return isnothing(old_side) ? result : old_side

    # Style: use new if provided, else inherit from old
    if haskey(new_val, "style")
        new_val["style"] ∈ VALID_BORDER_STYLES ||
            throw(XLSXError("Invalid border style: $(new_val["style"]). Must be one of: $(join(VALID_BORDER_STYLES, ", "))."))
        result["style"] = new_val["style"]
    elseif !isnothing(old_side) && haskey(old_side, "style")
        result["style"] = old_side["style"]
    end

    # Color: use new if provided, else inherit all non-style keys from old
    if haskey(new_val, "color")
        result["rgb"] = get_color(new_val["color"])
    elseif !isnothing(old_side)
        for (k, v) in old_side
            k != "style" && (result[k] = v)
        end
    end

    # Diagonal direction: use new if provided, else inherit, else default to "both"
    if side == "diagonal"
        if haskey(new_val, "direction")
            new_val["direction"] ∈ VALID_DIAGONAL_DIRECTIONS ||
                throw(XLSXError("Invalid diagonal direction: $(new_val["direction"]). Must be one of: $(join(VALID_DIAGONAL_DIRECTIONS, ", "))."))
            result["direction"] = new_val["direction"]
        elseif !isnothing(old_side) && haskey(old_side, "direction")
            result["direction"] = old_side["direction"]
        else
            result["direction"] = "both"
        end
    end

    return result
end

function get_new_formatId(wb::Workbook, format::String)::Int
    if haskey(builtinFormatNames, uppercasefirst(format)) # User specified a format by name
        return builtinFormatNames[format]
    elseif haskey(builtinFormats, format)                 # User specified a built-in format by ID
        return parse(Int64, format)
    else                                                  # user specified a format code
        if !is_valid_format(format)
            throw(XLSXError("Specified format is not a valid numFmt: $format"))
        end

        xroot = styles_xmlroot(wb)
        i, j = get_idces(xroot, "styleSheet", "numFmts")
        if isnothing(j) # There are no existing custom formats
            return styles_add_numFmt(wb, format)
        else
            existing_elements_count = length(XML.children(xroot[i][j]))
            if parse(Int, xroot[i][j]["count"]) != existing_elements_count
                throw(XLSXError("Wrong number of font elements found: $existing_elements_count. Expected $(parse(Int, xroot[i][j]["count"]))."))
            end

            format_node = XML.Element("numFmt";
                numFmtId=string(existing_elements_count + PREDEFINED_NUMFMT_COUNT),
                formatCode=XLSX.escape(format)
            )

            return styles_add_cell_attribute(wb, format_node, "numFmts") + PREDEFINED_NUMFMT_COUNT
        end
    end
end
function update_template_xf(ws::Worksheet, existing_style::CellDataFormat, attributes::Vector{String}, vals::Vector{String})::CellDataFormat
    old_cell_xf = styles_cell_xf(ws.package.workbook, Int(existing_style.id))
    new_cell_xf = copynode(old_cell_xf)
    if length(attributes) != length(vals)
        throw(XLSXError("Attributes and values must be of the same length."))
    end
    for (a, v) in zip(attributes, vals)
        new_cell_xf[a] = v
    end
    return styles_add_cell_xf(ws.package.workbook, new_cell_xf)
end
function update_template_xf(ws::Worksheet, allXfNodes::Vector{XML.Node}, existing_style::CellDataFormat, attributes::Vector{String}, vals::Vector{String})::CellDataFormat
    old_cell_xf = styles_cell_xf(allXfNodes, Int(existing_style.id))
    new_cell_xf = copynode(old_cell_xf)
    if length(attributes) != length(vals)
        throw(XLSXError("Attributes and values must be of the same length."))
    end
    for (a, v) in zip(attributes, vals)
        new_cell_xf[a] = v
    end
    return styles_add_cell_xf(ws.package.workbook, new_cell_xf)
end
#=
function update_template_xf(ws::Worksheet, existing_style::CellDataFormat, alignment::XML.Node)::CellDataFormat
    old_cell_xf = styles_cell_xf(ws.package.workbook, Int(existing_style.id))
    new_cell_xf = copynode(old_cell_xf)
    if isnothing(new_cell_xf.children)
        new_cell_xf=XML.Node(new_cell_xf, alignment)
    elseif length(XML.children(new_cell_xf)) == 0
        push!(new_cell_xf, alignment)
    else
        new_cell_xf[1] = alignment
    end
    return styles_add_cell_xf(ws.package.workbook, new_cell_xf)
end
=#
function update_template_xf(ws::Worksheet, allXfNodes::Vector{XML.Node}, existing_style::CellDataFormat, alignment::XML.Node)::CellDataFormat
    old_cell_xf = styles_cell_xf(allXfNodes, Int(existing_style.id))
    new_cell_xf = copynode(old_cell_xf)
    if isnothing(new_cell_xf.children)
        new_cell_xf=XML.Node(new_cell_xf, alignment)
    elseif length(XML.children(new_cell_xf)) == 0
        push!(new_cell_xf, alignment)
    else
        new_cell_xf[1] = alignment
    end
    return styles_add_cell_xf(ws.package.workbook, new_cell_xf)
end

# Only used in testing!
function styles_add_cell_font(wb::Workbook, attributes::Dict{String,Union{Dict{String,String},Nothing}})::Int
    pfx = get_prefix("xl/styles.xml", get_xlsxfile(wb))
    pfx = pfx == "" ? pfx : pfx * ":"
    new_font = buildNode("font", attributes, pfx)
    return styles_add_cell_attribute(wb, new_font, "fonts")
end

# Used by setFont(), setBorder(), setFill(), setAlignment() and setNumFmt()
function styles_add_cell_attribute(wb::Workbook, new_att::XML.Node, att::String)::Int
    xroot = styles_xmlroot(wb)
    i, j = get_idces(xroot, "styleSheet", att)
    existing_elements_count = length(XML.children(xroot[i][j]))
    if parse(Int, xroot[i][j]["count"]) != existing_elements_count
        throw(XLSXError("Wrong number of elements elements found: $existing_elements_count. Expected $(parse(Int, xroot[i][j]["count"]))."))
    end

    # Check new_att doesn't duplicate any existing att. If yes, use that rather than create new.
    for (k, node) in enumerate(XML.children(xroot[i][j]))
        if localname(new_att) == "numFmt" # mustn't compare numFmtId attribute for formats
            if node["formatCode"] == new_att["formatCode"]
                return k - 1 # CellDataFormat is zero-indexed
            end
        else
            if node == new_att
                return k - 1 # CellDataFormat is zero-indexed
            end
        end
    end

    push!(xroot[i][j], new_att)
    xroot[i][j]["count"] = string(existing_elements_count + 1)

    return existing_elements_count # turns out this is the new index (because it's zero-based)
end
function process_sheetcell(f::Function, xl::XLSXFile, sheetcell::String; kw...)
    if is_workbook_defined_name(xl, sheetcell)
        v = get_defined_name_value(xl.workbook, sheetcell)
        if is_defined_name_value_a_constant(v)
            throw(XLSXError("Can only assign attributes to cells but `$(sheetcell)` is a constant: $(sheetcell)=$v."))
        elseif is_defined_name_value_a_reference(v)
            newid = process_sheetcell(f, xl, string(v); kw...)
        else
            throw(XLSXError("Unexpected defined name value: $v."))
        end
    elseif is_valid_non_contiguous_sheetcellrange(sheetcell)
        sheetncrng = NonContiguousRange(sheetcell)
        !hassheet(xl, sheetncrng.sheet) && throw(XLSXError("Sheet $(sheetncrng.sheet) not found."))
        newid = f(xl[sheetncrng.sheet], sheetncrng; kw...)
    elseif is_valid_sheet_column_range(sheetcell)
        sheetcolrng = SheetColumnRange(sheetcell)
        !hassheet(xl, sheetcolrng.sheet) && throw(XLSXError("Sheet $(sheetcolrng.sheet) not found."))
        newid = f(xl[sheetcolrng.sheet], sheetcolrng.colrng; kw...)
    elseif is_valid_sheet_row_range(sheetcell)
        sheetrowrng = SheetRowRange(sheetcell)
        !hassheet(xl, sheetrowrng.sheet) && throw(XLSXError("Sheet $(sheetrowrng.sheet) not found."))
        newid = f(xl[sheetrowrng.sheet], sheetrowrng.rowrng; kw...)
    elseif is_valid_sheet_cellrange(sheetcell)
        sheetcellrng = SheetCellRange(sheetcell)
        !hassheet(xl, sheetcellrng.sheet) && throw(XLSXError("Sheet $(sheetcellrng.sheet) not found."))
        newid = f(xl[sheetcellrng.sheet], sheetcellrng.rng; kw...)
    elseif is_valid_sheet_cellname(sheetcell)
        ref = SheetCellRef(sheetcell)
        !hassheet(xl, ref.sheet) && throw(XLSXError("Sheet $(ref.sheet) not found."))
        newid = f(getsheet(xl, ref.sheet), ref.cellref; kw...)
    else
        throw(XLSXError("Invalid sheet cell reference: $sheetcell"))
    end
    return newid
end
function process_ranges(f::Function, ws::Worksheet, ref_or_rng::AbstractString; kw...)
    # Moved the tests for defined names to be first in case a name looks like a column name (e.g. "ID")
    if is_worksheet_defined_name(ws, ref_or_rng)
        v = get_defined_name_value(ws, ref_or_rng)
        if is_defined_name_value_a_constant(v)
            throw(XLSXError("Can only assign attributes to cells but `$(ref_or_rng)` is a constant: $(ref_or_rng)=$v."))
        elseif is_defined_name_value_a_reference(v)
            wb = get_workbook(ws)
            newid = f(get_xlsxfile(wb), string(v); kw...)
        else
            throw(XLSXError("Unexpected defined name value: $v."))
        end
    elseif is_workbook_defined_name(get_workbook(ws), ref_or_rng)
        wb = get_workbook(ws)
        v = get_defined_name_value(wb, ref_or_rng)
        if is_defined_name_value_a_constant(v)
            throw(XLSXError("Can only assign attributes to cells but `$(ref_or_rng)` is a constant: $(ref_or_rng)=$v."))
        elseif is_defined_name_value_a_reference(v)
            if is_valid_non_contiguous_range(string(v))
                _ = f.(Ref(get_xlsxfile(wb)), replace.(split(string(v), ","), "'" => "", "\$" => ""); kw...)
                newid = -1
            else
                newid = f(get_xlsxfile(wb), replace(string(v), "'" => "", "\$" => ""); kw...)
            end
        else
            throw(XLSXError("Unexpected defined name value: $v."))
        end
    elseif is_valid_column_range(ref_or_rng)
        newid = f(ws, ColumnRange(ref_or_rng); kw...)
    elseif is_valid_row_range(ref_or_rng)
        newid = f(ws, RowRange(ref_or_rng); kw...)
    elseif is_valid_cellrange(ref_or_rng)
        newid = f(ws, CellRange(ref_or_rng); kw...)
    elseif is_valid_cellname(ref_or_rng)
        newid = f(ws, CellRef(ref_or_rng); kw...)
    elseif is_valid_sheet_cellname(ref_or_rng)
        newid = f(ws, SheetCellRef(ref_or_rng); kw...)
    elseif is_valid_sheet_cellrange(ref_or_rng)
        newid = f(ws, SheetCellRange(ref_or_rng); kw...)
    elseif is_valid_sheet_column_range(ref_or_rng)
        newid = f(ws, SheetColumnRange(ref_or_rng); kw...)
    elseif is_valid_sheet_row_range(ref_or_rng)
        newid = f(ws, SheetRowRange(ref_or_rng); kw...)
    elseif is_valid_non_contiguous_cellrange(ref_or_rng)
        newid = f(ws, NonContiguousRange(ws, ref_or_rng); kw...)
    elseif is_valid_non_contiguous_sheetcellrange(ref_or_rng)
        nc = NonContiguousRange(ref_or_rng)
        newid = do_sheet_names_match(ws, nc) && f(ws, nc; kw...)
    else
        throw(XLSXError("Invalid cell reference or range: $ref_or_rng"))
    end
    return newid
end
function process_columnranges(f::Function, ws::Worksheet, colrng::ColumnRange; kw...)
    bounds = column_bounds(colrng)
    dim = (get_dimension(ws))
    left = bounds[begin]
    right = bounds[end]
    top = dim.start.row_number
    bottom = dim.stop.row_number

    OK = dim.start.column_number <= left
    OK &= dim.stop.column_number >= right
    OK &= dim.start.row_number <= top
    OK &= dim.stop.row_number >= bottom

    if OK
        rng = CellRange(top, left, bottom, right)
        return f(ws, rng; kw...)
    else
        throw(XLSXError("Column range $colrng is out of bounds. Worksheet `$(ws.name)` only has dimension `$dim`."))
    end
end
function process_rowranges(f::Function, ws::Worksheet, rowrng::RowRange; kw...)
    bounds = row_bounds(rowrng)
    dim = (get_dimension(ws))
    top = bounds[begin]
    bottom = bounds[end]
    left = dim.start.column_number
    right = dim.stop.column_number

    OK = dim.start.column_number <= left
    OK &= dim.stop.column_number >= right
    OK &= dim.start.row_number <= top
    OK &= dim.stop.row_number >= bottom

    if OK
        rng = CellRange(top, left, bottom, right)
        return f(ws, rng; kw...)
    else
        throw(XLSXError("Row range $rowrng is out of bounds. Worksheet `$(ws.name)` only has dimension `$dim`."))
    end
end
function process_ncranges(f::Function, ws::Worksheet, ncrng::NonContiguousRange; kw...)::Int
    bounds = nc_bounds(ncrng)
    if length(ncrng) == 1
        single = true
    else
        single = false
    end
    dim = (get_dimension(ws))
    OK = dim.start.column_number <= bounds.start.column_number
    OK &= dim.stop.column_number >= bounds.stop.column_number
    OK &= dim.start.row_number <= bounds.start.row_number
    OK &= dim.stop.row_number >= bounds.stop.row_number
    if OK
        for r in ncrng.rng
            if r isa CellRef && getcell(ws, r) isa EmptyCell
                single && throw(XLSXError("Cannot set format for an `EmptyCell`: $(cellname(r)). Set the value first."))
                continue
            end
            _ = f(ws, r; kw...)
        end
        return -1
    else
        throw(XLSXError("Non-contiguous range $ncrng is out of bounds. Worksheet `$(ws.name)` only has dimension `$dim`."))
    end
end
function process_cellranges(f::Function, ws::Worksheet, rng::CellRange; kw...)::Int
    if length(rng) == 1
        single = true
    else
        single = false
    end
    isInDim(ws, get_dimension(ws), rng)
    for cellref in rng
        if getcell(ws, cellref) isa EmptyCell
            single && throw(XLSXError("Cannot set format for an `EmptyCell`: $(cellname(cellref)). Set the value first."))
            continue
        end
        _ = f(ws, cellref; kw...)
    end
    return -1 # Each cell may have a different attribute Id so we can't return a single value.
end
function process_get_sheetcell(f::Function, xl::XLSXFile, sheetcell::String; kw...)
    ref = SheetCellRef(sheetcell)
    !hassheet(xl, ref.sheet) && throw(XLSXError("Sheet $(ref.sheet) not found."))
    ws = getsheet(xl, ref.sheet)
    d = get_dimension(ws)
    if ref.cellref ∉ d
        throw(XLSXError("Cell specified is outside sheet dimension \"$d\""))
    end
    return f(ws, ref.cellref; kw...)
end
function process_get_cellref(f::Function, ws::Worksheet, cellref::CellRef; kw...)
    wb = get_workbook(ws)
    cell = getcell(ws, cellref)
    d = get_dimension(ws)
    if cellref ∉ d
        throw(XLSXError("Cell specified is outside sheet dimension \"$d\""))
    end
    if cell isa EmptyCell || cell.style == UInt64(0)
        return nothing
    end
    cell_style = styles_cell_xf(wb, Int(cell.style))
    return f(wb, cell_style; kw...)
end
function process_get_cellname(f::Function, ws::Worksheet, ref_or_rng::AbstractString; kw...)
    if is_workbook_defined_name(get_workbook(ws), ref_or_rng)
        wb = get_workbook(ws)
        v = get_defined_name_value(wb, ref_or_rng)
        if is_defined_name_value_a_constant(v)
            throw(XLSXError("Can only assign attributes to cells but `$(ref_or_rng)` is a constant: $(ref_or_rng)=$v."))
        elseif is_defined_name_value_a_reference(v)
            new_att = f(get_xlsxfile(wb), unquoteit(string(v)); kw...)
#            new_att = f(get_xlsxfile(wb), replace(string(v), "'" => ""); kw...)
        else
            throw(XLSXError("Unexpected defined name value: $v."))
        end
    elseif is_valid_cellname(ref_or_rng)
        new_att = f(ws, CellRef(ref_or_rng); kw...)
    else
        throw(XLSXError("Invalid cell reference: $ref_or_rng"))
    end
    return new_att
end

#
# - Used for indexing `setAttribute` family of functions
#
function process_colon(f::Function, ws::Worksheet, row, col; kw...)
    dim = get_dimension(ws)
    @assert isnothing(row) || isnothing(col) "Something wrong here!"
    if isnothing(row) && isnothing(col)
        return f(ws, dim; kw...)
    elseif isnothing(col)
        rng = CellRange(CellRef(first(row), dim.start.column_number), CellRef(last(row), dim.stop.column_number))
    else
        rng = CellRange(CellRef(dim.start.row_number, first(col)), CellRef(dim.stop.row_number, last(col)))
#    else
#        throw(XLSXError("Something wrong here!"))
    end

    return f(ws, rng; kw...)
end
function process_veccolon(f::Function, ws::Worksheet, row, col; kw...)
    dim = get_dimension(ws)
    @assert isnothing(row) || isnothing(col) "Something wrong here!"
    if isnothing(col)
        col = dim.start.column_number:dim.stop.column_number
    else
        row = dim.start.row_number:dim.stop.row_number
#    else
#        throw(XLSXError("Something wrong here!"))
    end
    isInDim(ws, dim, row, col)
    if length(row) == 1 && length(col) == 1
        single = true
    else
        single = false
    end
    for a in row
        for b in col
            cellref = CellRef(a, b)
            if getcell(ws, cellref) isa EmptyCell
                single && throw(XLSXError("Cannot set attribute for an `EmptyCell`: $(cellname(cellref)). Set the value first."))
                continue
            end
            f(ws, cellref; kw...)
        end
    end
    return -1
end
function process_vecint(f::Function, ws::Worksheet, row, col; kw...)
    if length(col) == 1 && length(row) == 1
        single = true
    else
        single = false
    end
    dim = get_dimension(ws)
    isInDim(ws, dim, row, col)
    for a in row, b in col
        cellref = CellRef(a, b)
        if getcell(ws, cellref) isa EmptyCell
            single && throw(XLSXError("Cannot set format for an `EmptyCell`: $(cellname(cellref)). Set the value first."))
            continue
        end
        f(ws, cellref; kw...)
    end
    return -1
end

#
# - Used for indexing `setUniformAttribute` family of functions
#

#
# Most setUniform functions (but not Style or Alignment - see below)
#
function get_all_xf_nodes(ws::Worksheet)
    find_all_nodes(
        "/" * SPREADSHEET_NAMESPACE_XPATH_ARG * ":styleSheet/" *
        SPREADSHEET_NAMESPACE_XPATH_ARG * ":cellXfs/" *
        SPREADSHEET_NAMESPACE_XPATH_ARG * ":xf",
        styles_xmlroot(get_workbook(ws))
    )
end

function maybe_update_font!(f::Function, ws::Worksheet, cell, cellref; kw...)
    f == setFont || return
    cell isa EmptyCell && return
    cell.datatype == CT_STRING || return
    v = update_sharedString_font(ws, cell; kw...)
    cell.value = isnothing(v) ? cell.value : reinterpret(UInt64, Int64(v))
end

function process_uniform_core(f::Function, ws::Worksheet, allXfNodes::Vector{XML.Node}, cellref::CellRef, atts::Vector{String}, newid::Union{Int,Nothing}, first::Bool; kw...)
    cell = getcell(ws, cellref)
    cell isa EmptyCell && return newid, first
    if first
        newid = f(ws, cellref; kw...)
        first = false
    else
        if cell.style == UInt64(0)
            cell.style = get_num_style_index(ws, allXfNodes, 0).id
        end
        cell.style = update_template_xf(ws, allXfNodes, CellDataFormat(cell.style), atts, [string(newid), "1"]).id
    end
    return newid, first
end

function process_uniform_attribute(f::Function, ws::Worksheet, rng::CellRange, atts::Vector{String}; kw...)
    get_xlsxfile(ws).use_cache_for_sheet_data || throw(XLSXError("Cannot set uniform attributes because cache is not enabled."))
    allXfNodes = get_all_xf_nodes(ws)
    newid, first = nothing, true
    isInDim(ws, get_dimension(ws), rng)
    for cellref in rng
        newid, first = process_uniform_core(f, ws, allXfNodes, cellref, atts, newid, first; kw...)
        maybe_update_font!(f, ws, getcell(ws, cellref), cellref; kw...)
    end
    return first ? -1 : newid
end

function process_uniform_ncranges(f::Function, ws::Worksheet, ncrng::NonContiguousRange, atts::Vector{String}; kw...)::Int
    allXfNodes = get_all_xf_nodes(ws)
    single = length(ncrng) == 1
    bounds = nc_bounds(ncrng)
    dim = get_dimension(ws)

    dim.start.column_number <= bounds.start.column_number &&
    dim.stop.column_number  >= bounds.stop.column_number  &&
    dim.start.row_number    <= bounds.start.row_number    &&
    dim.stop.row_number     >= bounds.stop.row_number     ||
        throw(XLSXError("Non-contiguous range $ncrng is out of bounds. Worksheet `$(ws.name)` only has dimension `$dim`."))

    newid, first = nothing, true
    for r in ncrng.rng
        @assert r isa CellRef || r isa CellRange "Something wrong here"
        if r isa CellRef
            cell = getcell(ws, r)
            if cell isa EmptyCell
                single && throw(XLSXError("Cannot set format for an `EmptyCell`: $(cellname(r)). Set the value first."))
                continue
            end
            newid, first = process_uniform_core(f, ws, allXfNodes, r, atts, newid, first; kw...)
            maybe_update_font!(f, ws, cell, r; kw...)
        else
            for c in r
                newid, first = process_uniform_core(f, ws, allXfNodes, c, atts, newid, first; kw...)
                maybe_update_font!(f, ws, getcell(ws, c), c; kw...)
            end
        end
    end
    return first ? -1 : newid
end

function process_uniform_veccolon(f::Function, ws::Worksheet, row, col, atts::Vector{String}; kw...)
    @assert isnothing(row) || isnothing(col) "Something wrong here!"
    allXfNodes = get_all_xf_nodes(ws)
    dim = get_dimension(ws)
    isnothing(col) ? (col = dim.start.column_number:dim.stop.column_number) :
                     (row = dim.start.row_number:dim.stop.row_number)
    isInDim(ws, dim, row, col)
    newid, first = nothing, true
    for a in row, b in col
        cellref = CellRef(a, b)
        cell = getcell(ws, cellref)
        cell isa EmptyCell && continue
        newid, first = process_uniform_core(f, ws, allXfNodes, cellref, atts, newid, first; kw...)
        maybe_update_font!(f, ws, cell, cellref; kw...)
    end
    return first ? -1 : newid
end

function process_uniform_vecint(f::Function, ws::Worksheet, row, col, atts::Vector{String}; kw...)
    allXfNodes = get_all_xf_nodes(ws)
    dim = get_dimension(ws)
    isInDim(ws, dim, row, col)
    newid, first = nothing, true
    for a in row, b in col
        cellref = CellRef(a, b)
        cell = getcell(ws, cellref)
        cell isa EmptyCell && continue
        newid, first = process_uniform_core(f, ws, allXfNodes, cellref, atts, newid, first; kw...)
        maybe_update_font!(f, ws, cell, cellref; kw...)
    end
    return first ? -1 : newid
end

#
# UniformStyles
#
function process_uniform_core(ws::Worksheet, cellref::CellRef, newid::Union{Int,Nothing}, first::Bool, firstFont::Union{CellFont,Nothing})
    cell = getcell(ws, cellref)
    if cell isa EmptyCell # Can't add a attribute to an empty cell.
        return newid, first, firstFont
    end
    if first                           # Get the style of the first cell in the range.
        if cell.style !== UInt64(0)
            newid = Int(cell.style)
        end
        firstFont=getFont(ws, cellref)
        first = false
    else                               # Apply the same style to the rest of the cells in the range.
        cell.style = isnothing(newid) ? UInt64(0) : UInt64(newid)
    end
    if cell.datatype == CT_STRING && !isnothing(firstFont)
        v=update_sharedString_font(ws, cell, firstFont)
        cell.value= isnothing(v) ? cell.value : reinterpret(UInt64, Int64(v))
    end

    return newid, first, firstFont
end
function process_uniform_ncranges(ws::Worksheet, ncrng::NonContiguousRange)
    bounds = nc_bounds(ncrng)
    if length(ncrng) == 1
        single = true
    else
        single = false
    end
    dim = (get_dimension(ws))
    OK = dim.start.column_number <= bounds.start.column_number
    OK &= dim.stop.column_number >= bounds.stop.column_number
    OK &= dim.start.row_number <= bounds.start.row_number
    OK &= dim.stop.row_number >= bounds.stop.row_number
    if OK
        let newid::Union{Int,Nothing}, first::Bool
            newid = nothing
            first = true
            firstFont=nothing
            for r in ncrng.rng
                @assert r isa CellRef || r isa CellRange "Something wrong here"
                if r isa CellRef
                    if getcell(ws, r) isa EmptyCell
                        single && throw(XLSXError("Cannot set format for an `EmptyCell`: $(cellname(r)). Set the value first."))
                        continue
                    end
                    newid, first, firstFont = process_uniform_core(ws, r, newid, first, firstFont)
                else
                    for c in r
                        newid, first, firstFont = process_uniform_core(ws, c, newid, first, firstFont)
                    end
#                else
#                    throw(XLSXError("Something wrong here!"))
                end
            end
            if first
                newid = -1
            end
            return newid
        end
    else
        throw(XLSXError("Non-contiguous range $ncrng is out of bounds. Worksheet `$(ws.name)` only has dimension `$dim`."))
    end
end
function process_colon(ws::Worksheet, row, col)
    dim = get_dimension(ws)
    @assert isnothing(row) || isnothing(col) "Something wrong here!"
    if isnothing(row) && isnothing(col)
        return setUniformStyle(ws, dim)
    elseif isnothing(col)
        rng = CellRange(CellRef(first(row), dim.start.column_number), CellRef(last(row), dim.stop.column_number))
    else
        rng = CellRange(CellRef(dim.start.row_number, first(col)), CellRef(dim.stop.row_number, last(col)))
#    else
#        throw(XLSXError("Something wrong here!"))
    end

    return setUniformStyle(ws, rng)
end
function process_uniform_veccolon(ws::Worksheet, row, col)
    dim = get_dimension(ws)
    @assert isnothing(row) || isnothing(col) "Something wrong here!"
    if isnothing(col)
        col = dim.start.column_number:dim.stop.column_number
    else
        row = dim.start.row_number:dim.stop.row_number
#    else
#        throw(XLSXError("Something wrong here!"))
    end
    isInDim(ws, dim, row, col)
    let newid::Union{Int,Nothing}, first::Bool
        newid = nothing
        first = true
        firstFont = nothing
        for a in row
            for b in col
                cellref = CellRef(a, b)
                if getcell(ws, cellref) isa EmptyCell
                    continue
                end
                newid, first, firstFont = process_uniform_core(ws, cellref, newid, first, firstFont)
            end
        end
        if first
            newid = -1
        end
        return newid
    end
end
function process_uniform_vecint(ws::Worksheet, row, col)
    let newid::Union{Int,Nothing}, first::Bool
        dim = get_dimension(ws)
        newid = nothing
        first = true
        firstFont = nothing
        isInDim(ws, dim, row, col)
        for a in row, b in col
            cellref = CellRef(a, b)
            if getcell(ws, cellref) isa EmptyCell
                continue
            end
            newid, first, firstFont = process_uniform_core(ws, cellref, newid, first, firstFont)
        end
        if first
            newid = -1
        end
        return newid
    end
end

#
# Alignment is different
#
function process_uniform_core(f::Function, ws::Worksheet, allXfNodes::Vector{XML.Node}, cellref::CellRef, newid::Union{Int,Nothing}, first::Bool, alignment_node::Union{XML.Node,Nothing}; kw...) # setUniformAlignment is different
    cell = getcell(ws, cellref)
    if cell isa EmptyCell # Can't add a attribute to an empty cell.
        return newid, first, alignment_node
    end
    if first                           # Get the attribute of the first cell in the range.
        newid = f(ws, cellref; kw...)
        new_alignment = getAlignment(ws, cellref).alignment["alignment"]
        alignment_node = XML.Node(XML.Element, "alignment", new_alignment, nothing, nothing)
        first = false
    else                               # Apply the same attribute to the rest of the cells in the range.
        if cell.style == UInt64(0)
            cell.style = get_num_style_index(ws, allXfNodes, 0).id
        end
        cell.style = update_template_xf(ws, allXfNodes, CellDataFormat(cell.style), alignment_node).id
    end
    return newid, first, alignment_node
end
function process_uniform_attribute(f::Function, ws::Worksheet, rng::CellRange; kw...)
    if !get_xlsxfile(ws).use_cache_for_sheet_data
        throw(XLSXError("Cannot set uniform attributes because cache is not enabled."))
    end
    allXfNodes=find_all_nodes("/" * SPREADSHEET_NAMESPACE_XPATH_ARG * ":styleSheet/" * SPREADSHEET_NAMESPACE_XPATH_ARG * ":cellXfs/" * SPREADSHEET_NAMESPACE_XPATH_ARG * ":xf", styles_xmlroot(get_workbook(ws)))
    let newid::Union{Int,Nothing}, first::Bool, alignment_node::Union{XML.Node,Nothing}
        newid = nothing
        first = true
        alignment_node = nothing
        isInDim(ws, get_dimension(ws), rng)
        for cellref in rng
            if getcell(ws, cellref) isa EmptyCell
                continue
            end
            newid, first, alignment_node = process_uniform_core(f, ws, allXfNodes, cellref, newid, first, alignment_node; kw...)
        end
        if first
            newid = -1
        end
        return newid
    end
end
function process_uniform_ncranges(f::Function, ws::Worksheet, ncrng::NonContiguousRange; kw...)::Int
    allXfNodes=find_all_nodes("/" * SPREADSHEET_NAMESPACE_XPATH_ARG * ":styleSheet/" * SPREADSHEET_NAMESPACE_XPATH_ARG * ":cellXfs/" * SPREADSHEET_NAMESPACE_XPATH_ARG * ":xf", styles_xmlroot(get_workbook(ws)))
    bounds = nc_bounds(ncrng)
    if length(ncrng) == 1
        single = true
    else
        single = false
    end
    dim = (get_dimension(ws))
    OK = dim.start.column_number <= bounds.start.column_number
    OK &= dim.stop.column_number >= bounds.stop.column_number
    OK &= dim.start.row_number <= bounds.start.row_number
    OK &= dim.stop.row_number >= bounds.stop.row_number
    if OK
        let newid::Union{Int,Nothing}, first::Bool, alignment_node::Union{XML.Node,Nothing}
            newid = nothing
            first = true
            alignment_node = nothing
            for r in ncrng.rng
                @assert r isa CellRef || r isa CellRange "Something wrong here"
                if r isa CellRef && getcell(ws, r) isa EmptyCell
                    single && throw(XLSXError("Cannot set format for an `EmptyCell`: $(cellname(r)). Set the value first."))
                    continue
                end
                if r isa CellRef
                    if getcell(ws, r) isa EmptyCell
                        single && throw(XLSXError("Cannot set format for an `EmptyCell`: $(cellname(r)). Set the value first."))
                        continue
                    end
                    newid, first, alignment_node = process_uniform_core(f, ws, allXfNodes, r, newid, first, alignment_node; kw...)
                else
                    for c in r
                        newid, first, alignment_node = process_uniform_core(f, ws, allXfNodes, c, newid, first, alignment_node; kw...)
                    end
#                else
#                    throw(XLSXError("Something wrong here!"))
                end
            end
            if first
                newid = -1
            end
            return newid
        end
    else
        throw(XLSXError("Non-contiguous range $ncrng is out of bounds. Worksheet `$(ws.name)` only has dimension `$dim`."))
    end
end
function process_uniform_veccolon(f::Function, ws::Worksheet, row, col; kw...)
    dim = get_dimension(ws)
    if dim === nothing
        throw(XLSXError("No worksheet dimension found"))
    else
        @assert isnothing(row) || isnothing(col) "Something wrong here!"
        allXfNodes=find_all_nodes("/" * SPREADSHEET_NAMESPACE_XPATH_ARG * ":styleSheet/" * SPREADSHEET_NAMESPACE_XPATH_ARG * ":cellXfs/" * SPREADSHEET_NAMESPACE_XPATH_ARG * ":xf", styles_xmlroot(get_workbook(ws)))
        if isnothing(col)
            col = dim.start.column_number:dim.stop.column_number
        else
            row = dim.start.row_number:dim.stop.row_number
#        else
#            throw(XLSXError("Something wrong here!"))
        end
        isInDim(ws, dim, row, col)
        let newid::Union{Int,Nothing}, first::Bool, alignment_node::Union{XML.Node,Nothing}
            newid = nothing
            first = true
            alignment_node = nothing
            for a in row
                for b in dim.start.column_number:dim.stop.column_number
                    cellref = CellRef(a, b)
                    if getcell(ws, cellref) isa EmptyCell
                        continue
                    end
                    newid, first, alignment_node = process_uniform_core(f, ws, allXfNodes, cellref, newid, first, alignment_node; kw...)
                end
            end
            if first
                newid = -1
            end
            return newid
        end
    end
end
function process_uniform_vecint(f::Function, ws::Worksheet, row, col; kw...)
    let newid::Union{Int,Nothing}, first::Bool, alignment_node::Union{XML.Node,Nothing}
        dim = get_dimension(ws)
        if dim === nothing
            throw(XLSXError("No worksheet dimension found"))
        end
        allXfNodes=find_all_nodes("/" * SPREADSHEET_NAMESPACE_XPATH_ARG * ":styleSheet/" * SPREADSHEET_NAMESPACE_XPATH_ARG * ":cellXfs/" * SPREADSHEET_NAMESPACE_XPATH_ARG * ":xf", styles_xmlroot(get_workbook(ws)))
        newid = nothing
        first = true
        alignment_node = nothing
        isInDim(ws, dim, row, col)
        for a in row, b in col
            cellref = CellRef(a, b)
            if getcell(ws, cellref) isa EmptyCell
                continue
            end
            newid, first, alignment_node = process_uniform_core(f, ws, allXfNodes, cellref, newid, first, alignment_node; kw...)
        end
        if first
            newid = -1
        end
        return newid
    end
end

# Check if a string is a valid named color in Colors.jl and convert to "FFRRGGBB" if it is.
get_colorant(color_symb::Symbol) = get_colorant(String(color_symb))
function get_colorant(color_string::String)
    try
        c = parse(Colors.Colorant, color_string)
        rgb = Colors.hex(c, :RRGGBB)
        return "FF" * rgb
    catch
        return nothing
    end
end
get_color(s::Symbol)::String = get_color(String(s))
function get_color(str::String)::String
    if occursin(r"^[0-9A-F]{8}$", str) # is a valid 8 digit hexadecimal color
        return str
    end
    s = replace(lowercase(str), "grey" => "gray")
    c = get_colorant(s)
    if isnothing(c)
        throw(XLSXError("Invalid color specified: $s. Either give a valid color name (from Colors.jl) or an 8-digit rgb color in the form FFRRGGBB"))
    end
    return c
end
function update_sharedString_font(ws::Worksheet, cell::Cell, firstFont::CellFont) :: Union{Nothing,Int64}
    let bold=nothing, italic=nothing, under=nothing, strike=nothing, size=nothing, color=nothing, name=nothing
        for (k, v) in firstFont.font
            if k=="b"
                bold = true
            elseif k=="i"
                italic = true
            elseif k=="strike"
                strike = true
            elseif k=="u"
                under = v === nothing ? nothing : v["val"]
            elseif k=="sz"
                size = v === nothing ? nothing : parse(Int, v["val"])
            elseif k=="color"
                color = v
            elseif k=="name"
                name = v === nothing ? nothing : v["val"]
            else
                throw(XLSXError("Something wrong here!"))
            end
        end
        return update_sharedString_font(ws, cell; bold, italic, under, strike, size, color, name)
    end
end
        
function update_sharedString_font(ws::Worksheet, cell::Cell;
    bold::Union{Nothing,Bool}=nothing,
    italic::Union{Nothing,Bool}=nothing,
    under::Union{Nothing,String}=nothing,
    strike::Union{Nothing,Bool}=nothing,
    size::Union{Nothing,Int}=nothing,
    color::Union{Nothing,String,Dict{String,String}}=nothing,
    name::Union{Nothing,String}=nothing
) :: Union{Nothing,Int64}
    # <rPr> elements in a sharedString override any font attributes in the cell Style.
    # If setFont is called, we need to replace any of the attributes it is setting in the <rPr> elements.
    # When this makes successive <rPr> elements identical, the <r> elements that contain them can be merged.

    # starting values
    wb=get_workbook(ws)
    index=reinterpret(Int64,cell.value)
    sst=get_sst(wb)
    str_formatted=sst.shared_strings[index+1]

    isnothing(findfirst("<r>", str_formatted)) && return nothing # no <r> elements to manage

    is = parse(str_formatted, XML.Node)[1] # Convert to XML.Node for ease of handling

    all_r = filter(z -> z.tag == "r", XML.children(is))
    run_elements = reduce(vcat, [XML.children(z) for z in all_r])
    rPr_elements=filter(z -> z.tag == "rPr", run_elements) # rPr elements

    t=String[] # text elements
    for i in filter(z -> z.tag == "t", run_elements)
        push!(t, XML.is_simple(i[1]) ? XML.simple_value(i[1]) : XML.value(i[1]))
    end

    for rPr in rPr_elements
        # Delete rPr attributes for any attributes (kw...) given in setFont
        atts = ["b", "i", "strike", "u", "vertAlign", "sz", "color", "rFont", "family", "scheme"] # set of all possible rPr attributes in required order

        new_rPr = fill(XML.Element("DeleteMe"), length(atts)) # to collect new rPr elements

        for att in XML.children(rPr) # first copy existing attributes
            for i in 1:length(atts)
                if att.tag == atts[i]
                    new_rPr[i] = att
                end
            end
        end

        # then mark any elements for deletion that are in the keywords given in setFont/setUniformFont
        if !isnothing(bold)
            new_rPr[1] = XML.Element("DeleteMe")
        end
        if !isnothing(italic)
            new_rPr[2] = XML.Element("DeleteMe")
        end
        if !isnothing(strike)
            new_rPr[3] = XML.Element("DeleteMe")
        end
        if !isnothing(under)
            new_rPr[4] = XML.Element("DeleteMe")
        end
        if !isnothing(size)
            new_rPr[6] = XML.Element("DeleteMe")
        end
        if !isnothing(color)
            new_rPr[7] = XML.Element("DeleteMe")
        end
        if !isnothing(name)
            new_rPr[8] = XML.Element("DeleteMe")
            new_rPr[9] = XML.Element("DeleteMe")
            new_rPr[10] = XML.Element("DeleteMe")
        end

        # finally push merged elements back to rPr
        if !isnothing(rPr.children)
            empty!(rPr.children)
            foreach(new_rPr) do element 
                element.tag != "DeleteMe" && push!(rPr.children, element)
            end
        end
    end

    # now need to merge any adjacent <r> elements that have identical <rPr> elements
    if length(t) == length(rPr_elements) # first <r> may or may not have an <rPr> element but always has a <t> element.
        inc_first=0
    elseif length(t) == length(rPr_elements)+1
        inc_first=1
    else
        throw(XLSXError("Something wrong here!"))
    end
    for i in length(rPr_elements):-1:2 # merge adjacent <r> elements that have identical <rPr> elements
        if rPr_elements[i] == rPr_elements[i-1]
            t[i+inc_first-1] *= t[i+inc_first]
            t[i+inc_first] = ")___DeleteMe___("
        end
    end

    # if first <r> has no <rPr> and the only remaining <rPr> is empty, merge text
    if inc_first==1 && (length(t) == 2 || all([x == ")___DeleteMe___(" for x in t[3:end]]))
        if isnothing(XML.attributes(rPr_elements[1]))
            t[1] *= t[2]
            t[2] = ")___DeleteMe___("
        end
    end


    pattern = ")___DeleteMe___("
    valid = s -> !occursin(pattern, s)
    only_first_valid = valid(t[1]) && count(valid, t) == 1
    if only_first_valid
        # only one <r>, so convert to cell level Font
        if inc_first == 1
            # no atts => no op
        else
            # move single run attributes to cell Font attributes
            setFont(ws, cell.ref; XLSX.getRichTextString(ws, cell.ref).runs[1].atts...)
        end
        new_index=add_shared_string!(wb, t[1])
    else
        # reconstruct updated str_formatted
        new_r = IOBuffer()
        write(new_r, "<si>\n")
        for r in 1:length(all_r)
            if t[r] != ")___DeleteMe___(" # signals a merged <r> element to be skipped
                write(new_r, "  <r>\n")
                r > inc_first && write(new_r, XML.write(rPr_elements[r-inc_first];depth=3) * "\n")
                write(new_r, "    <t" * (needs_preserve(t[r]) ? " xml:space=\"preserve\"" : "") * ">" *t[r] * "</t>\n")
                write(new_r, "  </r>\n")
            end
        end
        write(new_r, "</si>")

        str_formatted = String(take!(new_r))

        ind = get(sst.index, str_formatted, nothing)
        if ind !== nothing
            return ind  # Found exact match
        end

        new_index=add_formatted_string!(sst, str_formatted) # can't update existing sharded string in case it is used by another cell
    end
    return new_index

end