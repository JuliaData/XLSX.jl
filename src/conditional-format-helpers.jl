#
# ---- Some random helper functions
#
function isValidKw(kw::String, val::Union{String, Nothing}, valid::Vector{String})
    if isnothing(val) || val ∈ valid
        return true
    else
        throw(XLSXError("Invalid keyword $kw: $val. Valid values are $valid"))
    end
end
function uppercase_unquoted(s::AbstractString)
    result = IOBuffer()
    i = firstindex(s)
    inside_quote = false
    while i <= lastindex(s)
        c = s[i]
        if c == '\\' && nextind(s, i) <= lastindex(s)
            # Handle escaped character
            next_i = nextind(s, i)
            print(result, s[i:next_i])
            i = nextind(s, next_i)
        elseif c == '"'
            inside_quote = !inside_quote
            print(result, c)
            i = nextind(s, i)
        else
            if inside_quote
                print(result, c)
            else
                print(result, uppercase(c))
            end
            i = nextind(s, i)
        end
    end
    return String(take!(result))
end

#
# --- Standard conditional formats
#

function process_cf_vecint(f::Function, ws::Worksheet, row, col; kw...)
    dim = get_dimension(ws)
    isInDim(ws, dim, row, col)
    cells = [CellRef(a, b) for a in row for b in col]
    return f(ws, NonContiguousRange(ws.name, _compress(cells)); kw...)
end
function process_cf_veccolon(f::Function, ws::Worksheet, row, col; kw...)
    dim = get_dimension(ws)
    row, col = _expand_colon(dim, row, col)
    isInDim(ws, dim, row, col)
    cells = [CellRef(a, b) for a in row for b in col]
    return f(ws, NonContiguousRange(ws.name, _compress(cells)); kw...)
end

# Each `setCf*` setter is written by hand for a `CfRange`. This defines every
# other way of naming the target, each reducing to that one, so that all the
# setters accept the same forms.
function define_cf_address_methods(F::Symbol)
    @eval begin
        $F(ws::Worksheet, row::Union{Integer,UnitRange{<:Integer}}, ::Colon; kw...) = process_colon($F, ws, row, nothing; kw...)
        $F(ws::Worksheet, ::Colon, col::Union{Integer,UnitRange{<:Integer}}; kw...) = process_colon($F, ws, nothing, col; kw...)
        $F(ws::Worksheet, ::Colon, ::Colon; kw...) = process_colon($F, ws, nothing, nothing; kw...)
        $F(ws::Worksheet, ::Colon; kw...) = process_colon($F, ws, nothing, nothing; kw...)
        $F(ws::Worksheet, row::Union{Integer,UnitRange{<:Integer}}, col::Union{Integer,UnitRange{<:Integer}}; kw...) = $F(ws, CellRange(CellRef(first(row), first(col)), CellRef(last(row), last(col))); kw...)
        $F(ws::Worksheet, cell::CellRef; kw...) = $F(ws, CellRange(cell, cell); kw...)
        $F(ws::Worksheet, cell::SheetCellRef; kw...) = do_sheet_names_match(ws, cell) && $F(ws, CellRange(cell.cellref, cell.cellref); kw...)
        $F(ws::Worksheet, rng::SheetCellRange; kw...) = do_sheet_names_match(ws, rng) && $F(ws, rng.rng; kw...)
        $F(ws::Worksheet, rng::SheetColumnRange; kw...) = do_sheet_names_match(ws, rng) && $F(ws, rng.colrng; kw...)
        $F(ws::Worksheet, rng::SheetRowRange; kw...) = do_sheet_names_match(ws, rng) && $F(ws, rng.rowrng; kw...)
        $F(ws::Worksheet, rng::RowRange; kw...) = process_rowranges($F, ws, rng; kw...)
        $F(ws::Worksheet, rng::ColumnRange; kw...) = process_columnranges($F, ws, rng; kw...)
        $F(xl::XLSXFile, sheetcell::AbstractString; kw...)::Int = process_sheetcell($F, xl, sheetcell; kw...)
        $F(ws::Worksheet, ref_or_rng::AbstractString; kw...)::Int = process_ranges($F, ws, ref_or_rng; kw...)
        $F(ws::Worksheet, row::Union{Vector{Int},StepRange{<:Integer}}, ::Colon; kw...) = process_cf_veccolon($F, ws, row, nothing; kw...)
        $F(ws::Worksheet, ::Colon, col::Union{Vector{Int},StepRange{<:Integer}}; kw...) = process_cf_veccolon($F, ws, nothing, col; kw...)
        $F(ws::Worksheet, row::Union{Integer,UnitRange{<:Integer}}, col::Union{Vector{Int},StepRange{<:Integer}}; kw...) = process_cf_vecint($F, ws, row, col; kw...)
        $F(ws::Worksheet, row::Union{Vector{Int},StepRange{<:Integer}}, col::Union{Integer,UnitRange{<:Integer}}; kw...) = process_cf_vecint($F, ws, row, col; kw...)
        $F(ws::Worksheet, row::Union{Vector{Int},StepRange{<:Integer}}, col::Union{Vector{Int},StepRange{<:Integer}}; kw...) = process_cf_vecint($F, ws, row, col; kw...)
    end
    return nothing
end

function allCfs(ws::Worksheet)::Vector{XML.Node}
    wb = get_workbook(ws)
    xf = get_xlsxfile(ws)
    target_file = get_relationship_target_by_id("xl", wb, ws.relationship_id)
    v = xf.data[target_file]
    sheetdoc = v isa String ? parse(v, XML.Node) : xmlroot(wb, ws.relationship_id)
    return _cfs_in(sheetdoc)
end
# The sheet's conditional formats: children of its root, matched by local name. (Not a
# path search over the whole document, which visits every cell of a parsed sheet.)
_cfs_in(sheetdoc::XML.Node)::Vector{XML.Node} =
    elements_with_tag(xml_root_element(sheetdoc), "conditionalFormatting")
function add_cf_to_XML(ws, new_cf)
    wb = get_workbook(ws)
    sheetdoc = xmlroot(get_workbook(ws), ws.relationship_id)
    # Not `sheetdoc[end]`: text or comments may follow `</worksheet>`.
    root = sheetdoc[find_child_index(XML.children(sheetdoc), "worksheet")]
    l = insert_index(root, "conditionalFormatting", WORKSHEET_ORDER)
    len = length(root)
    if l != len
        insert!(root.children, l+1, new_cf)
    else
        push!(root, new_cf)
    end
end
function next_cf_priority!(ws::Worksheet)::Int
    if ws.next_cf_priority === nothing
        # One-time O(n) scan over whatever rules already exist (e.g. in a
        # file opened with existing conditional formatting). Every
        # subsequent call is O(1).
        allcfs    = allCfs(ws)
        allextcfs = allExtCfs(ws)
        old_cf    = append!(getConditionalFormats(ws, allcfs), getConditionalExtFormats(ws, allextcfs))
        ws.next_cf_priority = isempty(old_cf) ? 1 : maximum(last(x).priority for x in old_cf) + 1
    end
    pr = ws.next_cf_priority
    ws.next_cf_priority += 1
    return pr
end
function update_worksheet_cfx!(cfx, ws, rng)
    pfx = get_prefix(ws)
    pfx = pfx == "" ? pfx : pfx * ":"

    sq = _cf_sqref(rng)

    # Match against the live tree, not `allCfs`, which may return a throwaway parse.
    allcfs = _cfs_in(xmlroot(get_workbook(ws), ws.relationship_id))
    matchcfs = filter(x -> x["sqref"] == sq, allcfs)
    l = length(matchcfs)
    if l == 0
        new_cf = XML.Element("conditionalFormatting"; sqref=sq)
        push!(new_cf, cfx)
        add_cf_to_XML(ws, new_cf)
    elseif l == 1
        push!(matchcfs[1], cfx)
    else
        throw(XLSXError("Too many conditional formatting blocks for range `$rng`. Must be one or none, found `$l`."))
    end
    update_worksheets_xml!(get_xlsxfile(ws))
end
#
# --- Conditional formats relying on Excel 2010 extensions
#
function allExtCfs(ws::Worksheet)::Vector{XML.Node}
    wb = get_workbook(ws)
    xf = get_xlsxfile(ws)
    target_file = get_relationship_target_by_id("xl", wb, ws.relationship_id)
    v = xf.data[target_file]
    sheetdoc = v isa String ? parse(v, XML.Node) : xmlroot(wb, ws.relationship_id)
    return _extcfs_in(sheetdoc)
end
function _extcfs_in(sheetdoc::XML.Node)::Vector{XML.Node}
    blk = _x14_cf_block(sheetdoc)
    return isnothing(blk) ? Vector{XML.Node}() : xml_elements(blk)
end

# `<extLst>` may hold other extensions (sparklines, data validations, ...), in any order.
# Excel reads x14 conditional formats only from the `<ext>` with this uri.
const X14_CF_EXT_URI = "{78C0D931-6437-407d-A8EE-F0AAD7539E65}"

function _x14_cf_ext(extlst::XML.Node)
    i = findfirst(e -> localname(e) == "ext" && uppercase(get(e, "uri", "")) == uppercase(X14_CF_EXT_URI), xml_elements(extlst))
    return isnothing(i) ? nothing : xml_elements(extlst)[i]
end

# The worksheet's `<x14:conditionalFormattings>` block, or `nothing`.
function _x14_cf_block(sheetdoc::XML.Node)
    i, j = get_idces(sheetdoc, "worksheet", "extLst")
    isnothing(j) && return nothing
    ext = _x14_cf_ext(sheetdoc[i][j])
    isnothing(ext) && return nothing
    k = findfirst(c -> localname(c) == "conditionalFormattings", xml_elements(ext))
    return isnothing(k) ? nothing : xml_elements(ext)[k]
end

# Ids of the worksheet's x14 rules, which pair each with its `<cfRule>` (via `x14:id`).
function _x14_cf_ids(ws::Worksheet)::Set{String}
    ids = Set{String}()
    for b in _extcfs_in(xmlroot(get_workbook(ws), ws.relationship_id)), r in xml_elements(b)
        localname(r) == "cfRule" && haskey(r, "id") && push!(ids, uppercase(r["id"]))
    end
    return ids
end

# As `_x14_cf_block`, creating `<extLst>`, the `<ext>` and the block as needed.
function _x14_cf_block!(sheetdoc::XML.Node)
    blk = _x14_cf_block(sheetdoc)
    isnothing(blk) || return blk
    ext = _worksheet_ext!(xml_root_element(sheetdoc), X14_CF_EXT_URI)
    blk = XML.Element("x14:conditionalFormattings")
    push!(ext, blk)
    return blk
end
function make_extCfsBlock()
    extCf = XML.Element("x14:conditionalFormatting")
    extCf["xmlns:xm"] = "http://schemas.microsoft.com/office/excel/2006/main"
    return extCf
end
function update_worksheet_ext_cfx!(cfx, ws, rng)
    wb = get_workbook(ws)
    sq = _cf_sqref(rng)
    sheetdoc = xmlroot(get_workbook(ws), ws.relationship_id)
    allcfs = _extcfs_in(sheetdoc)   # live tree, not `allExtCfs`, which may return a throwaway parse
    blk = _x14_cf_block!(sheetdoc)
    # Match range with existing conditional formatting blocks. Find `xm:sqref` by name: it
    # need not be the last child, as indentation may leave whitespace text nodes after it.
    matchcfs = filter(allcfs) do x
        els = xml_elements(x)
        s = findlast(e -> localname(e) == "sqref", els)
        !isnothing(s) && XML.simple_value(els[s]) == sq
    end
    l = length(matchcfs)
    if l == 0                                                   # No existing conditional formatting blocks for this range so create a new one.
        new_cf = make_extCfsBlock()
        push!(new_cf, cfx)
        push!(new_cf, XML.Element("xm:sqref", XML.Text(sq)))
        push!(blk, new_cf)                                      # Add the new conditional formatting block to the worksheet XML.
    elseif l == 1                                               # Existing conditional formatting block found for this range so add new rule to that block.
        pushfirst!(matchcfs[1], cfx)
    else
        throw(XLSXError("Too many conditional formatting blocks for range `$rng`. Must be one or none, found `$l`."))
    end
    update_worksheets_xml!(get_xlsxfile(ws))
end
function get_x14_icon(x14set)
    rule = XML.Element("x14:cfRule", type="iconSet", priority="1", id="XXXX-xxxx-XXXX") # replace id with UUID generated at time of use.
    if x14set == "Custom"
        icon = XML.Element("x14:iconSet", iconSet="3Arrows", custom="1")
    else
        icon = XML.Element("x14:iconSet", iconSet=x14set)
    end
    if x14set=="5Boxes"
        vals=[0, 20, 40, 60, 80]
    elseif x14set=="Custom"
        vals=[0]
    else
        vals=[0, 33, 67]
    end
    for v in vals
        cfvo = XML.Element("x14:cfvo", type="percent")
        push!(cfvo, XML.Element("xm:f", XML.Text(v)))
        push!(icon, cfvo)
    end
    push!(rule, icon)
    return rule
end

function _normalise_cfvo_val(ws, val)
    isnothing(val) && return nothing
    if is_valid_fixed_sheet_cellname(val)
        do_sheet_names_match(ws, SheetCellRef(val))
        val = string(SheetCellRef(val).cellref)
    end
    return uppercase_unquoted(val)
end

# Check one threshold (`cfvo`) of a color scale, data bar or icon set: its type must be
# one of `types` (or, if `optional`, absent), and its value, unless the type is
# `formula`, a fixed cell reference or a number. `name` is the keywords' prefix
# (`min`, `mid`, ...), so errors name the keyword at fault.
function _check_cfvo(name::AbstractString, type, val, types; optional::Bool=true)
    (optional && isnothing(type)) || type in types ||
        throw(XLSXError("Invalid $(name)_type: $type. Valid options are: $(join(types, ", "))."))
    (!isnothing(type) && type == "formula") || isnothing(val) || is_valid_fixed_cellname(val) ||
        is_valid_fixed_sheet_cellname(val) || !isnothing(tryparse(Float64, val)) ||
        throw(XLSXError("Invalid $(name)_val: `$val`. Valid options (unless $(name)_type is `formula`) are a CellRef (e.g. `\$A\$1`) or a number."))
    return nothing
end
#
# ---- Formatting (styles) definitions for conditional formats
#
function Add_Cf_Dx(wb::Workbook, new_dx::XML.Node)::DxFormat
    # Check if the workbook already has a dxfs element. If not, add one.
    xroot = styles_xmlroot(wb)
    pfx = get_prefix("xl/styles.xml", get_xlsxfile(wb))
    pfx = pfx == "" ? pfx : pfx * ":"
    
    i, j = get_idces(xroot, "styleSheet", "dxfs")

    if isnothing(j) # No existing conditional formats so need to add a block (is this even possible?). Push everything lower down one.
        throw(XLSXError("No <dxfs> block found in the styles.xml file. Please submit an issue to report this and attach the Excel file you were working with."))
    else
        existing_dxf_elements_count = length(xml_elements(xroot[i][j]))

        if parse(Int, xroot[i][j]["count"]) != existing_dxf_elements_count
            throw(XLSXError("Wrong number of xf elements found: $existing_dxf_elements_count. Expected $(parse(Int, xroot[i][j]["count"]))."))
        end
    end

    #   Don't reuse duplicates here. Always create new!
    existingdx = xml_elements(xroot[i][j])
    dxfs = unlink(xroot[i][j], ("dxfs", "dxf")) # Create the new <dxfs> Node
    if length(existingdx) > 0
        for c in existingdx
            push!(dxfs, c) # Copy each existing <dxf> into the new <dxfs> Node
        end
    end
    push!(dxfs, new_dx)

    xroot[i][j] = dxfs # Update the worksheet with the new cols.

    xroot[i][j]["count"] = string(existing_dxf_elements_count + 1)

    return DxFormat(existing_dxf_elements_count) # turns out this is the new index (because it's zero-based)

end
function get_dx(dxStyle::Union{Nothing,String}, format::Union{Nothing,Vector{<:Pair{String}}}, font::Union{Nothing,Vector{<:Pair{String}}}, border::Union{Nothing,Vector{<:Pair{String}}}, fill::Union{Nothing,Vector{<:Pair{String}}})::AbstractDict{String,<:AbstractDict{String}}
    if isnothing(dxStyle)
        if all(isnothing.([border, fill, font, format]))
            dx = highlights["redfilltext"]
        else
            dx = Dict{String,Dict{String,Any}}()
            for att in ["font" => font, "fill" => fill, "border" => border, "format" => format]
                if !isnothing(last(att))
                    dxx = Dict{String,Any}()
                    for i in last(att)
                        push!(dxx, first(i) => last(i))
                    end
                    push!(dx, first(att) => dxx)
                end
            end
        end
    elseif haskey(highlights, dxStyle)
        dx = highlights[dxStyle]
    else
        throw(XLSXError("Invalid dxStyle: $dxStyle. Valid options are: $(keys(highlights))."))
    end
    return dx
end
function get_new_dx(wb::Workbook, dx::AbstractDict{String,<:AbstractDict{String}})::XML.Node
    new_dx = XML.Element("dxf")

    for k in ["font", "format", "fill", "border"]
        haskey(dx, k) || continue
        v = dx[k]

        if k == "fill"
            filldx = XML.Element("fill")
            patterndx = XML.Element("patternFill")
            for (y, z) in v
                y in ["pattern", "bgColor", "fgColor"] || throw(XLSXError("Invalid fill attribute: $y. Valid options are: `pattern`, `bgColor`, `fgColor`."))
                if y in ["fgColor", "bgColor"]
                    push!(patterndx, _color_element(y, z))
                elseif y == "pattern" && z != "none"
                    patterndx["patternType"] = z
                end
            end
            push!(filldx, patterndx)
            push!(new_dx, filldx)

        elseif k == "font"
            fontdx = XML.Element("font")
            for (y, z) in v
                y in ["color", "bold", "italic", "under", "strike"] || throw(XLSXError("Invalid font attribute: $y. Valid options are: `color`, `bold`, `italic`, `under`, `strike`."))
                if     y == "color"                  ; push!(fontdx, _color_element(y, z))
                elseif y == "bold"   && z == "true"  ; push!(fontdx, XML.Element("b"))   # `val="0"` would mean not bold
                elseif y == "italic" && z == "true"  ; push!(fontdx, XML.Element("i"))
                elseif y == "under"
                    z ∈ ("none", "single", "double") || throw(XLSXError("Invalid font attribute value: under=$z. Valid options are: `none`, `single`, `double`."))
                    z == "single" && push!(fontdx, XML.Element("u"))
                    z == "double" && push!(fontdx, XML.Element("u"; val="double"))
                elseif y == "strike" && z == "true"  ; push!(fontdx, XML.Element(y))
                end
            end
            push!(new_dx, fontdx)

        elseif k == "border"
            all(y in ["color", "style"] for y in keys(v)) || throw(XLSXError("Invalid border attribute. Valid options are: `color`, `style`."))
            borderdx = XML.Element("border")
            cdx = haskey(v, "color") ? _color_element("color", v["color"]) : nothing
            sdx = get(v, "style", nothing)
            for side in ["left", "right", "top", "bottom"]
                el = XML.Element(side)
                isnothing(sdx) || (el["style"] = sdx)
                isnothing(cdx) || push!(el, cdx)
                push!(borderdx, el)
            end
            push!(new_dx, borderdx)

        elseif k == "format"
            if haskey(v, "format")
                new_formatId = get_new_formatId(wb, String(v["format"]))   # any AbstractString (#498)
                push!(new_dx, XML.Element("numFmt"; numFmtId=string(new_formatId), formatCode=styles_numFmt_formatCode(wb, new_formatId)))
            end
        end
    end

    return new_dx
end

# ---------------------------------------------------------------------------
# Non-contiguous range support for conditional formats
# ---------------------------------------------------------------------------

const CfRange = Union{CellRange,NonContiguousRange}

_cf_areas(rng::CellRange) = (rng,)
_cf_areas(ncr::NonContiguousRange) = ncr.rng

# OOXML `sqref` is space-separated; XLSX.jl's user-facing ranges are comma-separated.
_cf_sqref(rng::CellRange) = string(rng)
_cf_sqref(ncr::NonContiguousRange) = join((string(a) for a in ncr.rng), " ")

_cf_anchor(rng::CellRange) = rng.start
_cf_anchor(ncr::NonContiguousRange) =
    (a = first(ncr.rng); a isa CellRef ? a : a.start)

function _cf_range(ws::Worksheet, sqref::AbstractString)
    s = join(split(strip(sqref)), ",")
    occursin(',', s) && return NonContiguousRange(ws, s)
    is_valid_cellname(s) && return CellRange(CellRef(s), CellRef(s))
    return CellRange(s)
end

function _check_cf_areas(ws::Worksheet, rng::CfRange)
    dim = get_dimension(ws)
    for a in _cf_areas(rng)
        r = a isa CellRef ? CellRange(a, a) : a
        issubset(r, dim) || throw(XLSXError("Range `$a` goes outside worksheet dimension ($dim)."))
    end
    return nothing
end

_check_cf_range(ws::Worksheet, rng::CellRange) = _check_cf_areas(ws, rng)
function _check_cf_range(ws::Worksheet, ncr::NonContiguousRange)
    do_sheet_names_match(ws, ncr)
    isempty(ncr.rng) && throw(XLSXError("Cannot apply a conditional format to an empty range."))
    _check_cf_areas(ws, ncr)
    sq = _cf_sqref(ncr)
    length(sq) > 2000 && throw(XLSXError(
        "Range has too many separate areas for a conditional format (sqref is $(length(sq)) characters)."))
    return nothing
end

# ---------------------------------------------------------------------------
# Color bands
# ---------------------------------------------------------------------------

# "FFRRGGBB" -> Colorant, for interpolation. `get_color` guarantees 8 hex digits,
# so `s[3:end]` is always a clean RRGGBB. Alpha is dropped: data bar fills ignore it.
_to_colorant(s::AbstractString) = parse(Colors.Colorant, "#" * s[3:end])

# Colorant -> "FFRRGGBB". LCHab interpolation can leave the sRGB gamut, so clamp.
function _from_colorant(c)
    r = convert(Colors.RGB{Float64}, c)
    return "FF" * Colors.hex(Colors.RGB{Float64}(clamp(r.r, 0, 1),
                                                 clamp(r.g, 0, 1),
                                                 clamp(r.b, 0, 1)), :RRGGBB)
end

# Interpolation is done in LCHab rather than RGB: an RGB path from green to red
# passes through a muddy olive, which defeats the point of visually ordered bands.
# RGB colors come back as "AARRGGBB"; theme colors are kept as given, so they stay
# theme colors, but cannot be interpolated between.
function _band_colors(spec, n::Int)::Vector
    cols = [_is_theme_color(c) ? (_check_color(c); c) : get_color(c)
            for c in (spec isa CellColor ? [spec] : spec)]

    if length(cols) == n
        return cols
    elseif length(cols) == 2
        n == 1 && return [cols[1]]
        any(_is_theme_color, cols) && throw(XLSXError(
            "Theme colors cannot be interpolated between. Give RGB colors, or one color per band."))
        c1, c2 = convert.(Colors.LCHab, _to_colorant.(cols))
        return _from_colorant.(range(c1, c2; length=n))
    end
    throw(XLSXError("`colors` must give either $n colors (one per band) or " *
                    "2 colors to interpolate between, got $(length(cols))."))
end

# Every value in a range as one flat vector. getdata gives a matrix for a
# contiguous range and a vector of matrices, one per part, for a non-contiguous one.
_cf_values(ws::Worksheet, rng) =
    (d = getdata(ws, rng);
     d isa AbstractMatrix ? vec(d) : reduce(vcat, vec.(d); init = Any[]))
