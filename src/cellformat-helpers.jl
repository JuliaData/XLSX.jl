
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
    "6" => "\$#,##0_);[Red](\$#,##0)",
    "7" => "\$#,##0.00_);(\$#,##0.00)",
    "8" => "\$#,##0.00_);[Red](\$#,##0.00)",
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
    "38" => "#,##0_);[Red](#,##0)",
    "39" => "#,##0.00_);(#,##0.00)",
    "40" => "#,##0.00_);[Red](#,##0.00)",
    "45" => "mm:ss",
    "46" => "[h]:mm:ss",
    "47" => "mm:ss.0",
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

# A Markdown bullet list of `vals`, for docstrings to list the values a setter
# validates against from the same constant, so the two can't drift apart.
_doc_list(vals) = join(("- `$v`" for v in vals), "\n")

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

function do_sheet_names_match(ws::Worksheet, rng::T) where {T<:Union{SheetCellRef,AbstractSheetCellRange}}
    if _same_sheetname(ws.name, rng.sheet)
        return true
    else
        throw(XLSXError("Worksheet `$(ws.name)` does not match sheet in cell reference: `$(rng.sheet)`"))
    end
end

function make_child_node(tag::String, name::String, pfx::String)::XML.Node
    children = tag ∈ ("border", "fill") ? Vector{XML.Node{String}}() : nothing
    return XML.Node{String}(XML.Element, pfx*name, Pair{String,String}[], nothing, children)
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


# Merges a boolean/flag font tag (e.g. "b", "i", "strike").
# Returns nothing (present, no attributes) if kept/set, or missing to omit.
function _merge_flag_tag(tag::String, new_val::Union{Nothing,Bool}, old_atts::Dict)
    (isnothing(new_val) ? _flag_set(old_atts, tag) : new_val) ? nothing : missing
end

# Whether a font's flag tag is on: present, and not `<b val="0"/>` / `val="false"`.
function _flag_set(atts::Dict, tag::String)::Bool
    haskey(atts, tag) || return false
    a = atts[tag]
    return isnothing(a) || !(get(a, "val", "1") in ("0", "false"))
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

Whether `fmt` is a number format code Excel accepts: checked against Excel 365, which
refuses, for example, an unclosed quote or bracket, more than four sections, a colour
it doesn't know (`[Purple]`) or date codes mixed with digit placeholders (`0 h`).
"""
is_valid_format(fmt::AbstractString) = !isnothing(_parse_numfmt(fmt; strict=true))

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

# Coerces a border-side kwarg (a vector of `String => value` pairs) to a Dict, or returns nothing.
_to_border_dict(v) = isnothing(v) ? nothing : Dict{String,Any}(p for p in v)

# Validates that `outside` is not combined with any other border kwargs.
function _check_outside_conflict(s::BorderSpec)
    !isnothing(s.outside) && !all(isnothing, (s.left, s.right, s.top, s.bottom, s.diagonal, s.allsides)) &&
        throw(XLSXError("Keyword `outside` is incompatible with any other keywords."))
end

# Merges new and old attributes for a single border side.
# Handles style, color, and (for diagonal) direction.
function _merge_border_side(
    side::String,
    new_val::Union{Nothing,AbstractDict{String}},
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
        merge!(result, Dict(_color_atts(new_val["color"])))
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
    for (name, id) in builtinFormatNames                  # User specified a format by name
        lowercase(name) == lowercase(format) && return id
    end
    if haskey(builtinFormats, format)                     # User specified a built-in format by ID
        return parse(Int64, format)
    else                                                  # user specified a format code
        if !is_valid_format(format)
            throw(XLSXError("Specified format is not a valid numFmt: $format"))
        end
        return styles_add_numFmt(wb, format)              # reuses an identical format
    end
end
# The start of each per-cell format setter: the cell at `cellref`, which must have
# a value, given a style of its own if it uses the default one, and the cellXfs nodes.
function _format_target(sh::Worksheet, cellref::CellRef, what::String)
    cell = getcell(sh, cellref)
    cell isa EmptyCell &&
        throw(XLSXError("Cannot set $what for an `EmptyCell`: $(cellname(cellref)). Set the value first."))
    allXfNodes = get_cellXfs_nodes(get_workbook(sh))
    cell.style == UInt32(0) && (cell.style = get_num_style_index(sh, allXfNodes, 0).id)
    return cell, allXfNodes
end

# The namespace prefix of styles.xml with its colon, e.g. "x:", or "" for none.
_styles_prefix(x) = (p = get_prefix("xl/styles.xml", get_xlsxfile(x)); p == "" ? p : p * ":")

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

# The children of a cellXfs `<xf>`, in the order the schema requires.
const XF_CHILD_ORDER = ("alignment", "protection", "extLst")

# A copy of the style with its `<alignment>` or `<protection>` replaced by `child`,
# and attribute `apply` (e.g. `"applyProtection"`) set to "1" if given.
function update_template_xf(ws::Worksheet, allXfNodes::Vector{XML.Node}, existing_style::CellDataFormat, child::XML.Node;
                            apply::Union{Nothing,String}=nothing)::CellDataFormat
    old_cell_xf = styles_cell_xf(allXfNodes, Int(existing_style.id))
    new_cell_xf = copynode(old_cell_xf)

    existing_children = isnothing(new_cell_xf.children) ? XML.Node{String}[] : new_cell_xf.children
    remaining_children = filter(c -> localname(c) != localname(child), existing_children)
    push!(remaining_children, child)
    sort!(remaining_children; by = c -> something(findfirst(==(localname(c)), XF_CHILD_ORDER), length(XF_CHILD_ORDER) + 1))

    new_cell_xf = XML.Node{String}(XML.Element, XML.tag(new_cell_xf), new_cell_xf.attributes, nothing, remaining_children)
    isnothing(apply) || (new_cell_xf[apply] = "1")

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
    existing_elements_count = length(xml_elements(xroot[i][j]))
    if parse(Int, xroot[i][j]["count"]) != existing_elements_count
        throw(XLSXError("Wrong number of elements elements found: $existing_elements_count. Expected $(parse(Int, xroot[i][j]["count"]))."))
    end

    # Check new_att doesn't duplicate any existing att. If yes, use that rather than create new.
    for (k, node) in enumerate(xml_elements(xroot[i][j]))
        if XML.tag(new_att) == "numFmt" # mustn't compare numFmtId attribute for formats
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

    if att == "numFmts" && wb.numFmt_cache !== nothing
        wb.numFmt_cache[parse(Int, new_att["numFmtId"])] = new_att["formatCode"]
    elseif haskey(wb.style_table_cache, att)
        push!(wb.style_table_cache[att], new_att)
    end

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
            cleaned = replace(string(v), "'" => "", "\$" => "")
            if is_valid_non_contiguous_range(cleaned)
                newid = f(ws, NonContiguousRange(ws, cleaned); kw...)
            else
                newid = f(get_xlsxfile(wb), cleaned; kw...)
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
    elseif is_valid_non_contiguous_range(ref_or_rng)
        return f(ws, NonContiguousRange(ws, ref_or_rng); kw...)
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
    dim = get_dimension(ws)
    left, right = column_bounds(colrng)
    rng = CellRange(dim.start.row_number, left, dim.stop.row_number, right)
    issubset(rng, dim) ||
        throw(XLSXError("Column range $colrng is out of bounds. Worksheet `$(ws.name)` only has dimension `$dim`."))
    return f(ws, rng; kw...)
end
function process_rowranges(f::Function, ws::Worksheet, rowrng::RowRange; kw...)
    dim = get_dimension(ws)
    top, bottom = row_bounds(rowrng)
    rng = CellRange(top, dim.start.column_number, bottom, dim.stop.column_number)
    issubset(rng, dim) ||
        throw(XLSXError("Row range $rowrng is out of bounds. Worksheet `$(ws.name)` only has dimension `$dim`."))
    return f(ws, rng; kw...)
end
function process_ncranges(f::Function, ws::Worksheet, ncrng::NonContiguousRange; kw...)::Int
    single = length(ncrng) == 1
    dim = get_dimension(ws)
    OK = issubset(nc_bounds(ncrng), dim)
    if OK
        newid = -1 # Each cell may have a different attribute Id so we can't return a single value...
        for r in ncrng.rng
            if r isa CellRef && getcell(ws, r) isa EmptyCell
                single && throw(XLSXError("Cannot set format for an `EmptyCell`: $(cellname(r)). Set the value first."))
                continue
            end
            id = f(ws, r; kw...)
            single && (newid = id) # ...unless the range is a single cell.

        end
        return newid
    else
        throw(XLSXError("Non-contiguous range $ncrng is out of bounds. Worksheet `$(ws.name)` only has dimension `$dim`."))
    end
end
function process_cellranges(f::Function, ws::Worksheet, rng::CellRange; kw...)::Int
    single = length(rng) == 1
    isInDim(ws, get_dimension(ws), rng)
    newid = -1 # Each cell may have a different attribute Id so we can't return a single value...
    for cellref in rng
        if getcell(ws, cellref) isa EmptyCell
            single && throw(XLSXError("Cannot set format for an `EmptyCell`: $(cellname(cellref)). Set the value first."))
            continue
        end
        id = f(ws, cellref; kw...)
        single && (newid = id) # ...unless the range is a single cell.
    end
    return newid
end
function process_get_sheetcell(f::Function, xl::XLSXFile, sheetcell::String; kw...)
    ref = SheetCellRef(sheetcell)
    !hassheet(xl, ref.sheet) && throw(XLSXError("Sheet $(ref.sheet) not found."))
    ws = getsheet(xl, ref.sheet)
    return f(ws, ref.cellref; kw...)
end
# A cell outside the sheet's dimension is empty, as `ws[ref]` treats it, so it has
# no format to get. (getColumnWidth and getRowHeight keep their own check: a row or
# column beyond the dimension can still have a size.)
function process_get_cellref(f::Function, ws::Worksheet, cellref::CellRef; kw...)
    wb = get_workbook(ws)
    cell = getcell(ws, cellref)
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
# In a `(row, col)` target, `nothing` stands for `:`. Expand it to the full extent of
# the sheet's dimension `dim`. A target with neither a colon is not a colon form,
# so has no method.
_expand_colon(dim::CellRange, row, ::Nothing) = (row, dim.start.column_number:dim.stop.column_number)
_expand_colon(dim::CellRange, ::Nothing, col) = (dim.start.row_number:dim.stop.row_number, col)
_expand_colon(dim::CellRange, ::Nothing, ::Nothing) = (dim.start.row_number:dim.stop.row_number, dim.start.column_number:dim.stop.column_number)

function process_colon(f::Function, ws::Worksheet, row, col; kw...)
    row, col = _expand_colon(get_dimension(ws), row, col)
    return f(ws, CellRange(CellRef(first(row), first(col)), CellRef(last(row), last(col))); kw...)
end
function process_veccolon(f::Function, ws::Worksheet, row, col; kw...)
    dim = get_dimension(ws)
    row, col = _expand_colon(dim, row, col)
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
function maybe_update_font!(f::Function, ws::Worksheet, cell, cellref; spec)
    f === _setFont || return
    cell isa EmptyCell && return
    cell.datatype == CT_STRING || return
    v = update_sharedString_font(ws, cell, spec)
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
    _require_cache(ws, "set uniform attributes")
    allXfNodes = get_cellXfs_nodes(get_workbook(ws))
    newid, first = nothing, true
    isInDim(ws, get_dimension(ws), rng)
    for cellref in rng
        newid, first = process_uniform_core(f, ws, allXfNodes, cellref, atts, newid, first; kw...)
        maybe_update_font!(f, ws, getcell(ws, cellref), cellref; kw...)
    end
    return first ? -1 : newid
end

function process_uniform_ncranges(f::Function, ws::Worksheet, ncrng::NonContiguousRange, atts::Vector{String}; kw...)::Int
    allXfNodes = get_cellXfs_nodes(get_workbook(ws))
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
    row, col = _expand_colon(get_dimension(ws), row, col)
    return process_uniform_vecint(f, ws, row, col, atts; kw...)
end

function process_uniform_vecint(f::Function, ws::Worksheet, row, col, atts::Vector{String}; kw...)
    allXfNodes = get_cellXfs_nodes(get_workbook(ws))
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
        if cell.style != 0
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
    single = length(ncrng) == 1
    dim = get_dimension(ws)
    OK = issubset(nc_bounds(ncrng), dim)
    if OK
        let newid::Union{Int,Nothing}, first::Bool
            newid = nothing
            first = true
            firstFont=nothing
            for r in ncrng.rng
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
process_colon(ws::Worksheet, row, col) = process_colon(setUniformStyle, ws, row, col)
function process_uniform_veccolon(ws::Worksheet, row, col)
    dim = get_dimension(ws)
    row, col = _expand_colon(dim, row, col)
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
        alignment_node = XML.Node{String}(XML.Element, "alignment", isnothing(new_alignment) ? Pair{String,String}[] : Pair{String,String}[k => v for (k,v) in new_alignment], nothing, nothing)
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
    _require_cache(ws, "set uniform attributes")
allXfNodes = get_cellXfs_nodes(get_workbook(ws))
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
    allXfNodes = get_cellXfs_nodes(get_workbook(ws))
    single = length(ncrng) == 1
    dim = get_dimension(ws)
    OK = issubset(nc_bounds(ncrng), dim)
    if OK
        let newid::Union{Int,Nothing}, first::Bool, alignment_node::Union{XML.Node,Nothing}
            newid = nothing
            first = true
            alignment_node = nothing
            for r in ncrng.rng
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
        allXfNodes = get_cellXfs_nodes(get_workbook(ws))
        row, col = _expand_colon(dim, row, col)
        isInDim(ws, dim, row, col)
        let newid::Union{Int,Nothing}, first::Bool, alignment_node::Union{XML.Node,Nothing}
            newid = nothing
            first = true
            alignment_node = nothing
            for a in row
                for b in col
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
        allXfNodes = get_cellXfs_nodes(get_workbook(ws))
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


# The run properties (`<rPr>` children) that name a run's typeface. A cell font that
# sets any of them sets the typeface as a whole, so overrides all of them. XLSX.jl only
# writes `rFont`, but files from Excel carry `family` and `scheme` beside it (and
# `charset` for many fonts). `scheme` must go when the name changes, or Excel shows
# the theme font in place of the new one.
const RUN_FACE_TAGS = ("rFont", "family", "scheme", "charset")

# A cell font's tags are the run properties it overrides: the same names, except the
# cell font's `name` (or any other typeface tag) stands for all of `RUN_FACE_TAGS`.
_run_tags(font_tag::AbstractString) = font_tag in ("name", RUN_FACE_TAGS...) ? RUN_FACE_TAGS : (font_tag,)

# Each cell of a uniform range takes the first cell's font, so a rich-text cell's runs
# lose every property that font sets.
function update_sharedString_font(ws::Worksheet, cell::Cell, firstFont::CellFont) :: Union{Nothing,Int64}
    tags = Set{String}()
    for k in keys(firstFont.font)
        union!(tags, _run_tags(k))
    end
    return _strip_run_props!(ws, cell, tags)
end

# `setFont` and `setUniformFont` give the font as keywords; each one given overrides
# the matching run properties.
function update_sharedString_font(ws::Worksheet, cell::Cell, s::FontSpec) :: Union{Nothing,Int64}
    tags = Set{String}()
    for (given, font_tag) in ((!isnothing(s.bold), "b"), (!isnothing(s.italic), "i"), (!isnothing(s.strike), "strike"),
                              (!isnothing(s.under), "u"), (!isnothing(s.size), "sz"), (!isnothing(s.color), "color"),
                              (!isnothing(s.name), "name"), (!isnothing(s.vertAlign), "vertAlign"))
        given && union!(tags, _run_tags(font_tag))
    end
    return _strip_run_props!(ws, cell, tags)
end

# <rPr> elements in a sharedString override any font attributes in the cell Style.
# When the cell's font changes, remove from each run the properties named in `tags`,
# so the new cell font shows through, and keep every other property. When this makes
# successive <rPr> elements identical, the <r> elements that contain them are merged.
function _strip_run_props!(ws::Worksheet, cell::Cell, tags::Set{String}) :: Union{Nothing,Int64}
    wb=get_workbook(ws)
    index=reinterpret(Int64,cell.value)
    sst=get_sst(wb)
    str_formatted = get(sst.rich, index, nothing)
    str_formatted === nothing && return nothing # plain string

    isnothing(findfirst("<r>", str_formatted)) && return nothing # no <r> elements to manage

    is = parse(str_formatted, XML.Node)[1] # Convert to XML.Node for ease of handling

    all_r = filter(z -> localname(z) == "r", XML.children(is))

    # One entry per <r>: its <rPr> (or `nothing`) and its text.
    rPrs = Union{Nothing, XML.Node}[]
    t    = String[]
    for r in all_r
        kids = XML.children(r)
        i_rPr = findfirst(z -> localname(z) == "rPr", kids)
        i_t   = findfirst(z -> localname(z) == "t", kids)
        push!(rPrs, isnothing(i_rPr) ? nothing : kids[i_rPr])
        tel = kids[i_t]
        push!(t, XML.is_simple(tel[1]) ? XML.simple_value(tel[1]) : XML.value(tel[1]))
    end

    first_had_rPr = !isnothing(rPrs[1])

    for rPr in rPrs
        isnothing(rPr) && continue    # nothing to strip from an unattributed run
        # Whitespace between the properties is dropped too, so that a stripped-bare
        # <rPr> is empty below.
        isnothing(rPr.children) ||
            filter!(k -> XML.nodetype(k) == XML.Element && localname(k) ∉ tags, rPr.children)
    end

    # An <rPr> stripped down to nothing means the same as no <rPr> at all:
    # inherit the cell font. Normalise so the merge comparison below is total.
    for i in eachindex(rPrs)
        r = rPrs[i]
        if !isnothing(r) && (isnothing(XML.children(r)) || isempty(XML.children(r)))
            rPrs[i] = nothing
        end
    end

    # Merge adjacent runs with identical formatting. `nothing == nothing` is
    # true, so runs that both inherit the cell font merge as expected.
    for i in length(rPrs):-1:2
        if rPrs[i] == rPrs[i-1]
            t[i-1] *= t[i]
            t[i] = ")___DeleteMe___("
        end
    end

    pattern = ")___DeleteMe___("
    valid = s -> !occursin(pattern, s)

    if valid(t[1]) && count(valid, t) == 1
        # Read the run attributes while the cell still points at the rich text.
        atts1 = first_had_rPr ? getRichTextString(ws, cell.ref).runs[1].atts : nothing

        new_index = add_shared_string!(wb, t[1])

        kw = _run_font_kwargs(atts1)
        if !isempty(kw)
            # Point the cell at the plain string first, otherwise the setFont
            # below re-enters this function on the same rich text and recurses.
            cell.value = reinterpret(UInt64, Int64(new_index))
            setFont(ws, cell.ref; kw...)
        end
    else
        new_r = IOBuffer()
        write(new_r, "<si>\n")
        for r in eachindex(t)
            t[r] == pattern && continue
            write(new_r, "  <r>\n")
            isnothing(rPrs[r]) || write(new_r, XML.write(rPrs[r]) * "\n")
            write(new_r, "    <t" * (needs_preserve(t[r]) ? " xml:space=\"preserve\"" : "") *
                ">" * XML.escape(t[r]) * "</t>\n")
            write(new_r, "  </r>\n")
        end
        write(new_r, "</si>")

        str_formatted = String(take!(new_r))

        new_index = add_formatted_string!(wb, sst, str_formatted)
    end
    return new_index
end

# A rich-text run's attributes as `setFont` keywords: `:under => true` is a single
# underline, and false or absent flags set nothing.
function _run_font_kwargs(atts::Union{Nothing,AbstractDict})
    kw = Pair{Symbol,Any}[]
    isnothing(atts) && return kw
    for (k, v) in atts
        if k in (:bold, :italic, :strike)
            v in (true, 1) && push!(kw, k => true)
        elseif k === :under
            v in (true, 1) ? push!(kw, :under => "single") : v isa AbstractString && push!(kw, :under => String(v))
        elseif k === :name || k === :vertAlign
            push!(kw, k => String(v))
        else                                            # :size, :color
            push!(kw, k => v)
        end
    end
    return kw
end

#
# --- Address forms shared by the formatting setters
#
# Each setter is written by hand for one target: a `CellRef` for the per-cell
# setters, a `CellRange` for the uniform ones. These define every other way of
# naming the target (sheet-qualified references, strings, rows and columns,
# colons, vectors and step ranges), each reducing to that one. Generating them
# keeps each family identical, so a form added or a routing fixed here applies
# to every setter in it.

# Per-cell setters (`setFont`, `setFill`, ...). A setter that defines the
# `CellRange`, `ColumnRange`, `RowRange` and `XLSXFile` forms itself, as
# `setBorder` does for its `outside` keyword, passes `ranges=false`.
# Each entry is (arguments, body, return type or `nothing`).
function _percell_address_methods(F::Symbol, ranges::Bool)
    IntOrUR = :(Union{Integer,UnitRange{<:Integer}})
    VecOrSR = :(Union{Vector{Int},StepRange{<:Integer}})
    ms = Any[
        ((:(ws::Worksheet), :(ref::SheetCellRef)),      :(do_sheet_names_match(ws, ref) && $F(ws, ref.cellref; kw...)), nothing),
        ((:(ws::Worksheet), :(rng::SheetCellRange)),    :(do_sheet_names_match(ws, rng) && $F(ws, rng.rng; kw...)), nothing),
        ((:(ws::Worksheet), :(rng::SheetColumnRange)),  :(do_sheet_names_match(ws, rng) && $F(ws, rng.colrng; kw...)), nothing),
        ((:(ws::Worksheet), :(rng::SheetRowRange)),     :(do_sheet_names_match(ws, rng) && $F(ws, rng.rowrng; kw...)), nothing),
        ((:(ws::Worksheet), :(ncrng::NonContiguousRange)),   :(process_ncranges($F, ws, ncrng; kw...)), :Int),
        ((:(ws::Worksheet), :(ref_or_rng::AbstractString)),  :(process_ranges($F, ws, ref_or_rng; kw...)), :Int),
        ((:(ws::Worksheet), :(row::$IntOrUR), :(c::Colon)),  :(process_colon($F, ws, row, nothing; kw...)), nothing),
        ((:(ws::Worksheet), :(r::Colon), :(col::$IntOrUR)),  :(process_colon($F, ws, nothing, col; kw...)), nothing),
        ((:(ws::Worksheet), :(r::Colon), :(c::Colon)),       :(process_colon($F, ws, nothing, nothing; kw...)), nothing),
        ((:(ws::Worksheet), :(rc::Colon)),                   :(process_colon($F, ws, nothing, nothing; kw...)), nothing),
        ((:(ws::Worksheet), :(row::$VecOrSR), :(c::Colon)),  :(process_veccolon($F, ws, row, nothing; kw...)), nothing),
        ((:(ws::Worksheet), :(r::Colon), :(col::$VecOrSR)),  :(process_veccolon($F, ws, nothing, col; kw...)), nothing),
        ((:(ws::Worksheet), :(row::$IntOrUR), :(col::$VecOrSR)), :(process_vecint($F, ws, row, col; kw...)), nothing),
        ((:(ws::Worksheet), :(row::$VecOrSR), :(col::$IntOrUR)), :(process_vecint($F, ws, row, col; kw...)), nothing),
        ((:(ws::Worksheet), :(row::$VecOrSR), :(col::$VecOrSR)), :(process_vecint($F, ws, row, col; kw...)), nothing),
        ((:(ws::Worksheet), :(row::$IntOrUR), :(col::$IntOrUR)),
            :($F(ws, CellRange(CellRef(first(row), first(col)), CellRef(last(row), last(col))); kw...)), nothing),
    ]
    ranges && append!(ms, Any[
        ((:(ws::Worksheet), :(rng::CellRange)),       :(process_cellranges($F, ws, rng; kw...)), :Int),
        ((:(ws::Worksheet), :(colrng::ColumnRange)),  :(process_columnranges($F, ws, colrng; kw...)), :Int),
        ((:(ws::Worksheet), :(rowrng::RowRange)),     :(process_rowranges($F, ws, rowrng; kw...)), :Int),
        ((:(xl::XLSXFile), :(sheetcell::String)),     :(process_sheetcell($F, xl, sheetcell; kw...)), :Int),
    ])
    return ms
end

function define_percell_address_methods(F::Symbol; ranges::Bool=true)
    for (args, body, T) in _percell_address_methods(F, ranges)
        sig = :($F($(args...); kw...))
        isnothing(T) || (sig = :($sig::$T))
        @eval $sig = $body
    end
    return nothing
end

# Defines the public setter `F` over the internal setter `G`: each method of `F`
# gathers its keywords into one `S(:F; kw...)` and passes it to the same method
# of `G` as the keyword `spec`. The methods of `F` are those in `methods`, so its
# signatures, and the errors for a target it doesn't take, are those of any
# other setter in its family; but the many helpers behind `G` see only the
# keyword `spec`, so they compile once, rather than again for every new set of
# keywords. An unrecognized keyword is an `XLSXError` naming `F`.
function _define_spec_wrappers(F::Symbol, G::Symbol, S::Symbol, methods)
    for (args, _, _) in methods
        names = [a.args[1] for a in args]
        @eval $F($(args...); kw...) = $G($(names...); spec=$S($(QuoteNode(F)); kw...))
    end
    return nothing
end

# A per-cell setter `F` taking its keywords as an `S` (see `_define_spec_wrappers`).
# The internal setter `G` defines the `CellRef` method, taking `spec`, and with
# `ranges=false` the range methods too.
function define_percell_spec_methods(F::Symbol, G::Symbol, S::Symbol; ranges::Bool=true)
    define_percell_address_methods(G; ranges)
    cellref = ((:(ws::Worksheet), :(ref::CellRef)), nothing, nothing)
    _define_spec_wrappers(F, G, S, [_percell_address_methods(G, true); [cellref]])
    return nothing
end

# Uniform setters (`setUniformFont`, ...). `P` is the per-cell setter applied to
# the first cell, and `atts` the cellXfs attributes then copied to the others
# (e.g. `["fontId", "applyFont"]`). `setUniformAlignment` copies an alignment
# node rather than attributes, so passes `atts=nothing`.
# Each entry is (arguments, body, return type or `nothing`).
function _uniform_address_methods(F::Symbol, P::Symbol, atts::Union{Nothing,Vector{String}})
    a = isnothing(atts) ? () : (:([$(atts...)]),)   # a fresh vector literal at each call
    IntOrUR = :(Union{Integer,UnitRange{<:Integer}})
    VecOrSR = :(Union{Vector{Int},StepRange{<:Integer}})
    return Any[
        ((:(ws::Worksheet), :(rng::SheetCellRange)),    :(do_sheet_names_match(ws, rng) && $F(ws, rng.rng; kw...)), nothing),
        ((:(ws::Worksheet), :(rng::SheetColumnRange)),  :(do_sheet_names_match(ws, rng) && $F(ws, rng.colrng; kw...)), nothing),
        ((:(ws::Worksheet), :(rng::SheetRowRange)),     :(do_sheet_names_match(ws, rng) && $F(ws, rng.rowrng; kw...)), nothing),
        ((:(ws::Worksheet), :(colrng::ColumnRange)),    :(process_columnranges($F, ws, colrng; kw...)), :Int),
        ((:(ws::Worksheet), :(rowrng::RowRange)),       :(process_rowranges($F, ws, rowrng; kw...)), :Int),
        ((:(ws::Worksheet), :(ncrng::NonContiguousRange)),     :(process_uniform_ncranges($P, ws, ncrng, $(a...); kw...)), :Int),
        ((:(xl::XLSXFile), :(sheetcell::AbstractString)),      :(process_sheetcell($F, xl, sheetcell; kw...)), :Int),
        ((:(ws::Worksheet), :(ref_or_rng::AbstractString)),    :(process_ranges($F, ws, ref_or_rng; kw...)), :Int),
        ((:(ws::Worksheet), :(row::$IntOrUR), :(c::Colon)),    :(process_colon($F, ws, row, nothing; kw...)), nothing),
        ((:(ws::Worksheet), :(r::Colon), :(col::$IntOrUR)),    :(process_colon($F, ws, nothing, col; kw...)), nothing),
        ((:(ws::Worksheet), :(r::Colon), :(c::Colon)),         :(process_colon($F, ws, nothing, nothing; kw...)), nothing),
        ((:(ws::Worksheet), :(rc::Colon)),                     :(process_colon($F, ws, nothing, nothing; kw...)), nothing),
        ((:(ws::Worksheet), :(row::$VecOrSR), :(c::Colon)),    :(process_uniform_veccolon($P, ws, row, nothing, $(a...); kw...)), nothing),
        ((:(ws::Worksheet), :(r::Colon), :(col::$VecOrSR)),    :(process_uniform_veccolon($P, ws, nothing, col, $(a...); kw...)), nothing),
        ((:(ws::Worksheet), :(row::$IntOrUR), :(col::$VecOrSR)), :(process_uniform_vecint($P, ws, row, col, $(a...); kw...)), nothing),
        ((:(ws::Worksheet), :(row::$VecOrSR), :(col::$IntOrUR)), :(process_uniform_vecint($P, ws, row, col, $(a...); kw...)), nothing),
        ((:(ws::Worksheet), :(row::$VecOrSR), :(col::$VecOrSR)), :(process_uniform_vecint($P, ws, row, col, $(a...); kw...)), nothing),
        ((:(ws::Worksheet), :(row::$IntOrUR), :(col::$IntOrUR)),
            :($F(ws, CellRange(CellRef(first(row), first(col)), CellRef(last(row), last(col))); kw...)), nothing),
        ((:(ws::Worksheet), :(rng::CellRange)),         :(process_uniform_attribute($P, ws, rng, $(a...); kw...)), :Int),
    ]
end

function define_uniform_address_methods(F::Symbol, P::Symbol, atts::Union{Nothing,Vector{String}})
    for (args, body, T) in _uniform_address_methods(F, P, atts)
        sig = :($F($(args...); kw...))
        isnothing(T) || (sig = :($sig::$T))
        @eval $sig = $body
    end
    return nothing
end

# A uniform setter `F` taking its keywords as an `S` (see `_define_spec_wrappers`),
# over the internal uniform setter `G` and internal per-cell setter `P`.
function define_uniform_spec_methods(F::Symbol, G::Symbol, P::Symbol, S::Symbol, atts::Union{Nothing,Vector{String}})
    define_uniform_address_methods(G, P, atts)
    _define_spec_wrappers(F, G, S, _uniform_address_methods(G, P, atts))
    # A string naming a single cell reaches `G` as a `CellRef`, which a uniform
    # setter doesn't take: report it against `F`, as for any other target it doesn't take.
    @eval $G(ws::Worksheet, ref::CellRef; _...) = throw(MethodError($F, (ws, ref)))
    return nothing
end