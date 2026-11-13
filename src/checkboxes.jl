#
# -- Checkboxes in cells (Excel 365's Insert → Checkbox)
#
# A checkbox is part of a cell's format: its cellXfs `<xf>` has an extension
#   <ext uri="{C7286773-…}"><xfpb:xfComplement i="0"/></ext>
# whose `i` indexes the list of mapped bags in xl/featurePropertyBag/featurePropertyBag.xml.
# That bag (an XFComplement) leads through an XFControls bag to the Checkbox bag. The
# value is the cell's own: TRUE is ticked and FALSE isn't. Excel shows any other value
# as it is, without a box.
#

const REL_FEATURE_BAG = "http://schemas.microsoft.com/office/2022/11/relationships/FeaturePropertyBag"
const NS_FEATURE_BAG  = "http://schemas.microsoft.com/office/spreadsheetml/2022/featurepropertybag"
const EXT_XF_COMPLEMENT = "{C7286773-470A-42A8-94C5-96B5CB345126}"

const _FEATURE_BAG_PART = (rel=REL_FEATURE_BAG, path="xl/featurePropertyBag/featurePropertyBag.xml",
    ctype="application/vnd.ms-excel.featurepropertybag+xml",
    root=() -> _container("FeaturePropertyBags", ["xmlns" => NS_FEATURE_BAG]))

# The text of the `bagId` child of `bag` with key `k`, as an Int.
function _bag_ref(bag::XML.Node, k::String)::Union{Nothing,Int}
    for el in elements_with_tag(bag, "bagId")
        get_attr(el, "k") == k && return tryparse(Int, something(text_value(el), ""))
    end
    return nothing
end

_find_bag(bags, pick) = (i = findfirst(pick, bags); isnothing(i) ? nothing : i - 1)

# The 0-based ids of the bags for checkboxes, `(checkbox, controls, complement)`, as far
# as the file has them, and the bag list's XFComplements mapper.
function _checkbox_bags(root::XML.Node)
    bags = elements_with_tag(root, "bag")
    box = _find_bag(bags, b -> get_attr(b, "type") == "Checkbox")
    ctl = isnothing(box) ? nothing :
          _find_bag(bags, b -> get_attr(b, "type") == "XFControls" && _bag_ref(b, "CellControl") == box)
    cmp = isnothing(ctl) ? nothing :
          _find_bag(bags, b -> get_attr(b, "type") == "XFComplement" && _bag_ref(b, "XFControls") == ctl)
    mapper = findfirst(b -> get_attr(b, "type") == "XFComplements", bags)
    return box, ctl, cmp, isnothing(mapper) ? nothing : bags[mapper]
end

# The mapped bags' list in the XFComplements mapper.
function _mapped_list(mapper::XML.Node)::Union{Nothing,XML.Node}
    for a in elements_with_tag(mapper, "a")
        get_attr(a, "k") == "MappedFeaturePropertyBags" && return a
    end
    return nothing
end

# The `i` of a checkbox's xfComplement, or `nothing` if the file has no checkboxes.
function _checkbox_complement(xf::XLSXFile)::Union{Nothing,Int}
    root = _part_root(xf, REL_FEATURE_BAG)
    isnothing(root) && return nothing
    _, _, cmp, mapper = _checkbox_bags(root)
    (isnothing(cmp) || isnothing(mapper)) && return nothing
    list = _mapped_list(mapper)
    isnothing(list) && return nothing
    i = findfirst(b -> tryparse(Int, something(text_value(b), "")) == cmp, elements_with_tag(list, "bagId"))
    return isnothing(i) ? nothing : i - 1
end

# The `i` of a checkbox's xfComplement, adding to the feature property bags (and making
# their part) what a checkbox needs, as Excel writes it.
function _checkbox_complement!(xf::XLSXFile)::Int
    i = _checkbox_complement(xf)
    isnothing(i) || return i
    root = _workbook_part!(xf, _FEATURE_BAG_PART)
    pfx = _root_pfx(root)
    kids = XML.children(root)
    # New bags go before the mapper, which Excel writes last; nothing refers to it by id.
    function add_bag!(type, refs...)
        bag = _container(pfx * "bag", ["type" => type])
        for (k, id) in refs
            push!(bag, _xml_element(pfx * "bagId", ["k" => k], [XML.Text(string(id))]))
        end
        m = findfirst(k -> XML.nodetype(k) == XML.Element && get_attr(k, "type") == "XFComplements", kids)
        isnothing(m) ? push!(kids, bag) : insert!(kids, m, bag)
        return findfirst(b -> b === bag, elements_with_tag(root, "bag")) - 1
    end
    box, ctl, cmp, mapper = _checkbox_bags(root)
    box = something(box, add_bag!("Checkbox"))
    isnothing(ctl) && (ctl = add_bag!("XFControls", "CellControl" => box))
    isnothing(cmp) && (cmp = add_bag!("XFComplement", "XFControls" => ctl))
    if isnothing(mapper)
        mapper = _container(pfx * "bag", ["type" => "XFComplements", "extRef" => "XFComplementsMapperExtRef"])
        push!(root, mapper)
    end
    list = _mapped_list(mapper)
    if isnothing(list)
        list = _container(pfx * "a", ["k" => "MappedFeaturePropertyBags"])
        push!(mapper, list)
    end
    push!(list, _xml_element(pfx * "bagId", Pair{String,String}[], [XML.Text(string(cmp))]))
    return length(elements_with_tag(list, "bagId")) - 1
end

# Whether an xf's extension list has a checkbox's xfComplement `i`.
function _xf_has_checkbox(xf_node::XML.Node, i::Int)::Bool
    lst = first_element_with_tag(xf_node, "extLst")
    isnothing(lst) && return false
    for ext in elements_with_tag(lst, "ext")
        get_attr(ext, "uri") == EXT_XF_COMPLEMENT || continue
        c = first_element_with_tag(ext, "xfComplement")
        !isnothing(c) && get_attr(c, "i") == string(i) && return true
    end
    return false
end

# A copy of style `style` with the checkbox's extension added (`on`) or removed.
function _checkbox_style(ws::Worksheet, allXfNodes::Vector{XML.Node}, style::Int, on::Bool, i::Int)::Int
    style == 0 && (style = Int(get_num_style_index(ws, allXfNodes, 0).id))
    old = copynode(styles_cell_xf(allXfNodes, style))
    _xf_has_checkbox(old, i) == on && return style
    pfx = something(_ns_prefix(XML.tag(old)), "")
    lst = first_element_with_tag(old, "extLst")
    exts = isnothing(lst) ? XML.Node[] :
           XML.Node[e for e in xml_elements(lst) if get_attr(e, "uri") != EXT_XF_COMPLEMENT]
    if on
        ext = _container(pfx * "ext", ["uri" => EXT_XF_COMPLEMENT, "xmlns:xfpb" => NS_FEATURE_BAG])
        push!(ext, _xml_element("xfpb:xfComplement", ["i" => string(i)]))
        push!(exts, ext)
    end
    kids = XML.Node[c for c in something(XML.children(old), XML.Node[])
                    if !(XML.nodetype(c) == XML.Element && localname(c) == "extLst")]
    if !isempty(exts)
        lst = _container(pfx * "extLst")
        foreach(e -> push!(lst, e), exts)
        push!(kids, lst)   # extLst is an xf's last child
    end
    new = XML.Node{String}(XML.Element, XML.tag(old), old.attributes, nothing, kids)
    return Int(styles_add_cell_xf(get_workbook(ws), new).id)
end

function _set_checkbox!(ws::Worksheet, parts::Vector, on::Bool)
    what = on ? "add checkboxes" : "remove checkboxes"
    _require_writable(ws, what)
    _require_worksheet(ws, what)
    _require_cache(ws, what)
    parts = Any[_clip_to_dimension(ws, p) for p in parts]
    xf = get_xlsxfile(ws)
    i = on ? _checkbox_complement!(xf) : _checkbox_complement(xf)
    isnothing(i) && return nothing   # no checkboxes to remove
    allXfNodes = get_cellXfs_nodes(get_workbook(ws))
    _restyle_cells!(ws, parts; empty=on ? false : missing) do style
        _checkbox_style(ws, allXfNodes, style, on, i)
    end
    return nothing
end

"""
    setCheckbox(ws::Worksheet, ref)
    setCheckbox(xf::XLSXFile, "Sheet1!B2:B10")
    setCheckbox(ws::Worksheet, row, col)

Put a checkbox in each cell of `ref`, as Excel 365's *Insert → Checkbox* does. A
checkbox shows the cell's value: ticked for `TRUE` and clear for `FALSE`. Tick one by
writing `true` to its cell, and read it as a `Bool` like any other value.

An empty cell is given the value `false`. A cell that already holds `true` or `false`
keeps it. A cell holding anything else keeps its value too, and Excel shows that value
instead of a box until the cell holds `TRUE` or `FALSE`, as it does for a checkbox
added in Excel. The cells' other formatting is kept.

`ref` is a cell (`"B2"`), a range (`"B2:B10"`), whole columns (`"B:C"`) or rows
(`"2:5"`), several of these separated by commas, a defined name, a `CellRef` or a
`CellRange`. Whole rows and columns reach only as far as the sheet's used range.

Checkboxes need Excel 365; older versions of Excel show `TRUE` and `FALSE`.

The file must be open for writing (`mode="rw"`).

# Examples
```julia
julia> XLSX.setCheckbox(sheet, "B2:B10")

julia> sheet["B3"] = true            # tick B3

julia> [c.cell for c in XLSX.getCheckboxes(sheet) if c.checked]
1-element Vector{String}:
 "B3"
```

See also [`removeCheckbox`](@ref), [`hasCheckbox`](@ref) and [`getCheckboxes`](@ref).
"""
function setCheckbox end

"""
    removeCheckbox(ws::Worksheet, ref)
    removeCheckbox(xf::XLSXFile, "Sheet1!B2:B10")
    removeCheckbox(ws::Worksheet, row, col)

Take the checkboxes out of the cells of `ref`, leaving their values, `TRUE` or `FALSE`,
and their other formatting. `ref` is given as for [`setCheckbox`](@ref).

The file must be open for writing (`mode="rw"`).
"""
function removeCheckbox end

for (f, on) in ((:setCheckbox, true), (:removeCheckbox, false))
    @eval begin
        $f(xf::XLSXFile, ref::AbstractString) = $f(_rowcol_sheet(xf, ref)...)
        $f(ws::Worksheet, row::Integer, col::Integer) = $f(ws, CellRef(row, col))
        $f(ws::Worksheet, ref::CellRef) = $f(ws, CellRange(ref, ref))
        $f(ws::Worksheet, rng::CellRange) = _set_checkbox!(ws, Any[rng], $on)
        $f(ws::Worksheet, ref::AbstractString) = _set_checkbox!(ws, _rowcol_parts(ws, ref), $on)
    end
end

"""
    hasCheckbox(ws::Worksheet, ref) -> Bool
    hasCheckbox(xf::XLSXFile, "Sheet1!B2") -> Bool
    hasCheckbox(ws::Worksheet, row, col) -> Bool

Whether cell `ref` has a checkbox. Its value says whether it's ticked.

See also [`setCheckbox`](@ref) and [`getCheckboxes`](@ref).
"""
function hasCheckbox end
hasCheckbox(xf::XLSXFile, ref::AbstractString) = hasCheckbox(_rowcol_sheet(xf, ref)...)
hasCheckbox(ws::Worksheet, row::Integer, col::Integer) = hasCheckbox(ws, CellRef(row, col))
hasCheckbox(ws::Worksheet, ref::AbstractString) = hasCheckbox(ws, CellRef(ref))
function hasCheckbox(ws::Worksheet, ref::CellRef)::Bool
    _require_worksheet(ws, "read a checkbox")
    i = _checkbox_complement(get_xlsxfile(ws))
    isnothing(i) && return false
    cell = getcell(ws, ref)
    (cell isa EmptyCell || cell.style == 0) && return false
    return _xf_has_checkbox(styles_cell_xf(get_workbook(ws), Int(cell.style)), i)
end

"""
    getCheckboxes(ws::Worksheet) -> Vector{NamedTuple}

Every cell of the sheet with a checkbox, by row and then column, as `(cell, checked)`
`NamedTuple`s of the cell's reference (`"B2"`) and whether it's ticked: whether the
cell holds `true`.

See also [`setCheckbox`](@ref) and [`hasCheckbox`](@ref).
"""
function getCheckboxes(ws::Worksheet)
    _require_worksheet(ws, "read checkboxes")
    out = NamedTuple{(:cell, :checked),Tuple{String,Bool}}[]
    i = _checkbox_complement(get_xlsxfile(ws))
    isnothing(i) && return out
    wb = get_workbook(ws)
    boxed = Dict{UInt32,Bool}()   # by style
    for r in eachrow(ws)
        for col in sort!(collect(keys(r.rowcells)))
            cell = r.rowcells[col]
            cell.style == 0 && continue
            get!(() -> _xf_has_checkbox(styles_cell_xf(wb, Int(cell.style)), i), boxed, cell.style) || continue
            push!(out, (cell=string(cell.ref), checked=cell.datatype == CT_BOOL && cell.value != 0))
        end
    end
    return out
end
