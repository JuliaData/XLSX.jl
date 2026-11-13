#
# -- Rich values: pictures in cells
#
# A cell holding a picture (Excel's *Place in Cell*, or an `IMAGE` formula) has the
# value `#VALUE!` with a `vm` attribute: a 1-based index into the `valueMetadata` of
# xl/metadata.xml. That points, through a `futureMetadata` block, to a rich value in
# xl/richData/rdrichvalue.xml, whose structure (rdrichvaluestructure.xml) names its
# values. A picture placed in a cell is a `_localImage`, whose identifier indexes
# richValueRel.xml, a list of relationships to the media. A picture an `IMAGE` formula
# fetched is a `_webimage`, whose identifier indexes rdRichValueWebImage.xml.
#
# Dynamic-array formulas share xl/metadata.xml: their `cm` indexes its `cellMetadata`.
#

const REL_SHEET_METADATA = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sheetMetadata"
const REL_RD_RICHVALUE   = "http://schemas.microsoft.com/office/2017/06/relationships/rdRichValue"
const REL_RD_STRUCTURE   = "http://schemas.microsoft.com/office/2017/06/relationships/rdRichValueStructure"
const REL_RD_TYPES       = "http://schemas.microsoft.com/office/2017/06/relationships/rdRichValueTypes"
const REL_RICHVALUEREL   = "http://schemas.microsoft.com/office/2022/10/relationships/richValueRel"
const REL_RD_WEBIMAGE    = "http://schemas.microsoft.com/office/2020/07/relationships/rdRichValueWebImage"

const NS_RICHDATA     = "http://schemas.microsoft.com/office/spreadsheetml/2017/richdata"
const NS_RICHVALUEREL = "http://schemas.microsoft.com/office/spreadsheetml/2022/richvaluerel"
const NS_DYNAMICARRAY = "http://schemas.microsoft.com/office/spreadsheetml/2017/dynamicarray"

const EXT_RICHVALUE_BLOCK = "{3e2802c4-a4d2-4d8b-9148-e3be6c30e623}"
const EXT_DYNAMIC_ARRAY   = "{bdbb8cdc-fa1e-496e-a857-3c3f30c029c3}"

include_dependency(joinpath(@__DIR__, "data", "rdRichValueTypes.xml"))
const RICH_VALUE_TYPES_XML = read(joinpath(@__DIR__, "data", "rdRichValueTypes.xml"), String)

# An element with these attributes whose children can be added to: XML.jl gives an
# element made with none no children vector.
_container(tag::AbstractString, atts::Vector{Pair{String,String}}=Pair{String,String}[]) =
    XML.Node{String}(XML.Element, String(tag), atts, nothing, XML.Node{String}[])

# The prefix of a part's root element ("x:"), for its children, or "".
_root_pfx(root::XML.Node) = something(_ns_prefix(XML.tag(root)), "")

_part_doc(root::XML.Node) =
    XML.Document(XML.Declaration(; version="1.0", encoding="UTF-8", standalone="yes"), root)

# The workbook parts rich values use: relationship type, the path a new one takes, its
# content type, and its empty root.
const _METADATA_PART = (rel=REL_SHEET_METADATA, path="xl/metadata.xml",
    ctype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheetMetadata+xml",
    root=() -> _container("metadata", ["xmlns" => SPREADSHEET_NAMESPACE_XPATH_ARG]))
const _RICHVALUE_PART = (rel=REL_RD_RICHVALUE, path="xl/richData/rdrichvalue.xml",
    ctype="application/vnd.ms-excel.rdrichvalue+xml",
    root=() -> _container("rvData", ["xmlns" => NS_RICHDATA, "count" => "0"]))
const _STRUCTURE_PART = (rel=REL_RD_STRUCTURE, path="xl/richData/rdrichvaluestructure.xml",
    ctype="application/vnd.ms-excel.rdrichvaluestructure+xml",
    root=() -> _container("rvStructures", ["xmlns" => NS_RICHDATA, "count" => "0"]))
const _RICHVALUEREL_PART = (rel=REL_RICHVALUEREL, path="xl/richData/richValueRel.xml",
    ctype="application/vnd.ms-excel.richvaluerel+xml",
    root=() -> _container("richValueRels", ["xmlns" => NS_RICHVALUEREL, "xmlns:r" => NS_R]))
const _TYPES_PART = (rel=REL_RD_TYPES, path="xl/richData/rdRichValueTypes.xml",
    ctype="application/vnd.ms-excel.rdrichvaluetypes+xml",
    root=() -> xml_root_element(parse(RICH_VALUE_TYPES_XML, XML.Node)))

# The path of the workbook's part with relationship type `rel`, or `nothing`.
function _workbook_part(xf::XLSXFile, rel::String)::Union{Nothing,String}
    for r in get_workbook(xf).relationships
        r.Type == rel || continue
        path = parse_relationship_target("xl", r.Target)
        return haskey(xf.data, path) ? path : nothing
    end
    return nothing
end

# The root of the workbook's part `part`, made with its relationship and content type
# if the file has none.
function _workbook_part!(xf::XLSXFile, part)::XML.Node
    path = _workbook_part(xf, part.rel)
    if isnothing(path)
        path = part.path
        xf.data[path] = _part_doc(part.root())
        xf.files[path] = true
        add_relationship!(get_workbook(xf), make_relative_target("xl", path), part.rel)
        register_content_type!(xf, "[Content_Types].xml"; tag="Override", key="PartName",
                               val="/$path", content_type=part.ctype)
    end
    return xml_root_element(get_xml_data(xf, path))
end

_part_root(xf::XLSXFile, rel::String) =
    (path = _workbook_part(xf, rel); isnothing(path) ? nothing : xml_root_element(get_xml_data(xf, path)))

# Add `child` to `parent` and update its `count`. Returns the child's 1-based position.
function _push_counted!(parent::XML.Node, child::XML.Node)::Int
    push!(parent, child)
    n = length(xml_elements(parent))
    parent["count"] = string(n)
    return n
end

#
# -- xl/metadata.xml
#

const METADATA_ORDER = ("metadataTypes", "metadataStrings", "mdxMetadata", "futureMetadata",
                        "cellMetadata", "valueMetadata", "extLst")

# Excel's flags for its two kinds of metadata.
const _METADATA_TYPE_FLAGS = ["minSupportedVersion" => "120000", "copy" => "1", "pasteAll" => "1",
    "pasteValues" => "1", "merge" => "1", "splitFirst" => "1", "rowColShift" => "1",
    "clearFormats" => "1", "clearComments" => "1", "assign" => "1", "coerce" => "1"]

# The child `name` of metadata.xml's root (with `atts`, to tell `futureMetadata`
# blocks apart), added in the schema's order if it isn't there.
function _metadata_child!(root::XML.Node, name::String, atts::Vector{Pair{String,String}}=Pair{String,String}[])::XML.Node
    for el in xml_elements(root)
        localname(el) == name && all(get_attr(el, k) == v for (k, v) in atts) && return el
    end
    el = _container(_root_pfx(root) * name, vcat(atts, ["count" => "0"]))
    rank(n) = something(findfirst(==(localname(n)), METADATA_ORDER), length(METADATA_ORDER) + 1)
    kids = XML.children(root)
    i = findfirst(k -> XML.nodetype(k) == XML.Element && rank(k) > rank(el), kids)
    isnothing(i) ? push!(kids, el) : insert!(kids, i, el)
    return el
end

# The 1-based index of metadata type `name`, added with Excel's flags if missing.
function _metadata_type!(root::XML.Node, name::String, extra::Vector{Pair{String,String}})::Int
    types = _metadata_child!(root, "metadataTypes")
    i = findfirst(t -> get_attr(t, "name") == name, xml_elements(types))
    isnothing(i) || return i
    return _push_counted!(types, _xml_element(_root_pfx(root) * "metadataType",
                                              vcat(["name" => name], _METADATA_TYPE_FLAGS, extra)))
end

# Declare namespace `ns` on the root as `xmlns:pfx` if it isn't declared; returns its prefix.
function _declare_ns!(root::XML.Node, pfx::String, ns::String)::String
    p = get(ns_prefixes(root), ns, nothing)
    isnothing(p) || return p
    root["xmlns:$pfx"] = ns
    return pfx
end

# A `bk` holding an `extLst` with one `ext` of `uri` around `inner`.
function _future_bk(pfx::String, uri::String, inner::XML.Node)::XML.Node
    ext = _container(pfx * "ext", ["uri" => uri]); push!(ext, inner)
    lst = _container(pfx * "extLst"); push!(lst, ext)
    bk = _container(pfx * "bk"); push!(bk, lst)
    return bk
end

# Add a block for metadata type `name` (index `t`) whose future data is `inner`, and a
# record of it in `cellMetadata` or `valueMetadata` (`which`). Returns the record's
# 1-based index: a cell's `cm` or `vm`.
function _add_metadata_record!(root::XML.Node, name::String, t::Int, uri::String, inner::XML.Node, which::String)::Int
    pfx = _root_pfx(root)
    future = _metadata_child!(root, "futureMetadata", ["name" => name])
    v = _push_counted!(future, _future_bk(pfx, uri, inner)) - 1
    records = _metadata_child!(root, which)
    bk = _container(pfx * "bk")
    push!(bk, _xml_element(pfx * "rc", ["t" => string(t), "v" => string(v)]))
    return _push_counted!(records, bk)
end

# The `cm` of a dynamic-array formula: Excel's single `XLDAPR` record, added to
# metadata.xml if it has none, and metadata.xml made if the file has none.
function _dynamic_array_cm!(xf::XLSXFile)::Int
    root = _workbook_part!(xf, _METADATA_PART)
    t = _metadata_type!(root, "XLDAPR", ["cellMeta" => "1"])
    cells = first_element_with_tag(root, "cellMetadata")
    if !isnothing(cells)
        for (i, bk) in enumerate(xml_elements(cells))
            rc = first_element_with_tag(bk, "rc")
            !isnothing(rc) && get_attr(rc, "t") == string(t) && return i
        end
    end
    xda = _declare_ns!(root, "xda", NS_DYNAMICARRAY)
    inner = _xml_element("$xda:dynamicArrayProperties", ["fDynamic" => "1", "fCollapsed" => "0"])
    return _add_metadata_record!(root, "XLDAPR", t, EXT_DYNAMIC_ARRAY, inner, "cellMetadata")
end

# The `vm` of a new cell whose value is rich value `rv` (0-based).
function _rich_value_vm!(xf::XLSXFile, rv::Int)::Int
    root = _workbook_part!(xf, _METADATA_PART)
    t = _metadata_type!(root, "XLRICHVALUE", Pair{String,String}[])
    xlrd = _declare_ns!(root, "xlrd", NS_RICHDATA)
    inner = _xml_element("$xlrd:rvb", ["i" => string(rv)])
    return _add_metadata_record!(root, "XLRICHVALUE", t, EXT_RICHVALUE_BLOCK, inner, "valueMetadata")
end

#
# -- Reading rich values
#

# The workbook's rich values, read once for any number of cells.
struct _RichValues
    rv_of_vm::Vector{Int}                                  # vm => rich value (0-based), or -1
    values::Vector{Tuple{Int,Vector{String}}}              # rich value => (structure, its values)
    structures::Vector{Tuple{String,Vector{String}}}       # structure => (type, key names)
    rel_media::Vector{Union{Nothing,String}}               # richValueRel index => media part
    web::Vector{Tuple{Union{Nothing,String},Union{Nothing,String}}}   # web picture => (url, media part)
end

# The targets of a part's relationships, by Id: resolved part paths, or the URL itself
# for an external one.
function _part_rel_targets(xf::XLSXFile, part::String)::Dict{String,String}
    out = Dict{String,String}()
    rels = _sheet_rels_path(part)
    haskey(xf.data, rels) || return out
    dir = first(_split_zip_path(part))
    for n in elements_with_tag(xml_root_element(get_xml_data(xf, rels)), "Relationship")
        t = get_attr(n, "Target")
        out[get_attr(n, "Id")] = get_attr(n, "TargetMode") == "External" ? t : resolve_relative_target(dir, t)
    end
    return out
end

function _rich_values(xf::XLSXFile)::_RichValues
    rv_of_vm = Int[]
    meta = _part_root(xf, REL_SHEET_METADATA)
    if !isnothing(meta)
        types = first_element_with_tag(meta, "metadataTypes")
        type_names = isnothing(types) ? String[] : [get_attr(t, "name") for t in xml_elements(types)]
        futures = Dict(get_attr(f, "name") => xml_elements(f) for f in elements_with_tag(meta, "futureMetadata"))
        vals = first_element_with_tag(meta, "valueMetadata")
        for bk in (isnothing(vals) ? XML.Node[] : xml_elements(vals))
            rc = first_element_with_tag(bk, "rc")
            t = isnothing(rc) ? 0 : something(tryparse(Int, get_attr(rc, "t")), 0)
            v = isnothing(rc) ? -1 : something(tryparse(Int, get_attr(rc, "v")), -1)
            rv = -1
            if 1 <= t <= length(type_names) && type_names[t] == "XLRICHVALUE"
                bks = get(futures, "XLRICHVALUE", XML.Node[])
                if 0 <= v < length(bks)
                    rvb = _find_descendant(bks[v+1], "rvb")
                    rv = isnothing(rvb) ? -1 : something(tryparse(Int, get_attr(rvb, "i")), -1)
                end
            end
            push!(rv_of_vm, rv)
        end
    end
    values = Tuple{Int,Vector{String}}[]
    root = _part_root(xf, REL_RD_RICHVALUE)
    for rv in (isnothing(root) ? XML.Node[] : elements_with_tag(root, "rv"))
        push!(values, (parse(Int, get_attr(rv, "s", "0")),
                       [something(text_value(v), "") for v in elements_with_tag(rv, "v")]))
    end
    structures = Tuple{String,Vector{String}}[]
    root = _part_root(xf, REL_RD_STRUCTURE)
    for s in (isnothing(root) ? XML.Node[] : elements_with_tag(root, "s"))
        push!(structures, (get_attr(s, "t"), [get_attr(k, "n") for k in elements_with_tag(s, "k")]))
    end
    rel_media = Union{Nothing,String}[]
    path = _workbook_part(xf, REL_RICHVALUEREL)
    if !isnothing(path)
        targets = _part_rel_targets(xf, path)
        for rel in elements_with_tag(xml_root_element(get_xml_data(xf, path)), "rel")
            push!(rel_media, get(targets, something(get_prefixed_attr(rel, "id"), ""), nothing))
        end
    end
    web = Tuple{Union{Nothing,String},Union{Nothing,String}}[]
    path = _workbook_part(xf, REL_RD_WEBIMAGE)
    if !isnothing(path)
        targets = _part_rel_targets(xf, path)
        rid(el) = isnothing(el) ? nothing : get(targets, something(get_prefixed_attr(el, "id"), ""), nothing)
        for w in elements_with_tag(xml_root_element(get_xml_data(xf, path)), "webImageSrd")
            push!(web, (rid(first_element_with_tag(w, "address")), rid(first_element_with_tag(w, "blip"))))
        end
    end
    return _RichValues(rv_of_vm, values, structures, rel_media, web)
end

function _find_descendant(node::XML.Node, name::String)::Union{Nothing,XML.Node}
    for el in xml_elements(node)
        localname(el) == name && return el
        found = _find_descendant(el, name)
        isnothing(found) || return found
    end
    return nothing
end

# The rich value `vm` refers to, as its structure's type and a key => value Dict.
function _rich_value_fields(rich::_RichValues, vm::Integer)
    1 <= vm <= length(rich.rv_of_vm) || return nothing
    rv = rich.rv_of_vm[vm]
    0 <= rv < length(rich.values) || return nothing
    s, vals = rich.values[rv+1]
    0 <= s < length(rich.structures) || return nothing
    type, keys = rich.structures[s+1]
    return type, Dict(zip(keys, vals))
end

# The picture format of a media part, from its name.
function _media_format(path::AbstractString)::Symbol
    ext = lowercase(last(splitext(path)))
    return ext in (".jpg", ".jpeg") ? :jpeg : Symbol(lstrip(ext, '.'))
end

# The picture rich value `vm` holds, or `nothing` if it holds something else.
function _cell_image(xf::XLSXFile, rich::_RichValues, vm::Integer)::Union{Nothing,CellImage}
    fields = _rich_value_fields(rich, vm)
    isnothing(fields) && return nothing
    type, f = fields
    media(path) = isnothing(path) ? UInt8[] : get(xf.binary_data, path, UInt8[])
    if type == "_localImage"
        i = tryparse(Int, get(f, "_rvRel:LocalImageIdentifier", ""))
        (isnothing(i) || !(0 <= i < length(rich.rel_media))) && return nothing
        path = rich.rel_media[i+1]
        isnothing(path) && return nothing
        return CellImage(media(path), _media_format(path), get(f, "Text", ""), nothing)
    elseif type == "_webimage"
        i = tryparse(Int, get(f, "WebImageIdentifier", ""))
        (isnothing(i) || !(0 <= i < length(rich.web))) && return nothing
        url, path = rich.web[i+1]
        return CellImage(media(path), isnothing(path) ? :unknown : _media_format(path), get(f, "Text", ""), url)
    end
    return nothing
end

#
# -- Writing pictures into cells
#

# Excel's `_localImage` structures, without and with alternative text, and the
# CalcOrigin of a picture placed in a cell.
const _LOCAL_IMAGE_KEYS     = [("_rvRel:LocalImageIdentifier", "i"), ("CalcOrigin", "i")]
const _LOCAL_IMAGE_KEYS_ALT = [_LOCAL_IMAGE_KEYS; ("Text", "s")]
const CALC_ORIGIN_PLACED = "5"

# The 0-based index of the structure of `type` with exactly these keys, added if missing.
function _rv_structure!(xf::XLSXFile, type::String, keys::Vector{Tuple{String,String}})::Int
    root = _workbook_part!(xf, _STRUCTURE_PART)
    for (i, s) in enumerate(elements_with_tag(root, "s"))
        get_attr(s, "t") == type || continue
        ks = elements_with_tag(s, "k")
        [(get_attr(k, "n"), get_attr(k, "t")) for k in ks] == keys && return i - 1
    end
    pfx = _root_pfx(root)
    s = _container(pfx * "s", ["t" => type])
    foreach(((n, t),) -> push!(s, _xml_element(pfx * "k", ["n" => n, "t" => t])), keys)
    return _push_counted!(root, s) - 1
end

# The 0-based richValueRel index of a picture with these bytes: an existing one if the
# workbook has the same picture, as Excel shares them. Returns it with the media name.
function _rich_value_media!(xf::XLSXFile, bytes::Vector{UInt8})::Tuple{Int,String}
    root = _workbook_part!(xf, _RICHVALUEREL_PART)
    path = _workbook_part(xf, REL_RICHVALUEREL)
    targets = _part_rel_targets(xf, path)
    for (i, rel) in enumerate(elements_with_tag(root, "rel"))
        target = get(targets, something(get_prefixed_attr(rel, "id"), ""), "")
        get(xf.binary_data, target, nothing) == bytes && return i - 1, last(_split_zip_path(target))
    end
    name = _add_media_bytes!(xf, bytes)
    rid = add_part_rel!(xf, path, "xl/media/$name", REL_IMAGE)
    r = _declare_ns!(root, "r", NS_R)
    pfx = _root_pfx(root)
    return _push_counted_rel!(root, _xml_element(pfx * "rel", ["$r:id" => rid])), name
end

# richValueRels has no `count`.
function _push_counted_rel!(root::XML.Node, rel::XML.Node)::Int
    push!(root, rel)
    return length(xml_elements(root)) - 1
end

# Add picture `img` to the workbook's rich values. Returns the `vm` of a cell showing
# it, and the name of its media file.
function _add_cell_image!(xf::XLSXFile, img::CellImage)::Tuple{Int,String}
    isnothing(img.url) || throw(XLSXError("Can't write a picture from the web into a cell. Write an `IMAGE` formula with `setFormula`, or give the picture's file."))
    media, name = _rich_value_media!(xf, img.bytes)
    keys = isempty(img.alt) ? _LOCAL_IMAGE_KEYS : _LOCAL_IMAGE_KEYS_ALT
    s = _rv_structure!(xf, "_localImage", keys)
    root = _workbook_part!(xf, _RICHVALUE_PART)
    pfx = _root_pfx(root)
    rv = _container(pfx * "rv", ["s" => string(s)])
    for v in (string(media), CALC_ORIGIN_PLACED)
        push!(rv, _xml_element(pfx * "v", Pair{String,String}[], [XML.Text(v)]))
    end
    if !isempty(img.alt)
        keep = occursin(r"^\s|\s$", img.alt)
        push!(rv, _xml_element(pfx * "v", keep ? ["xml:space" => "preserve"] : Pair{String,String}[], [XML.Text(img.alt)]))
    end
    i = _push_counted!(root, rv) - 1
    _workbook_part!(xf, _TYPES_PART)
    return _rich_value_vm!(xf, i), name
end

#
# -- Public API
#

CellImage(image::AbstractString; alt::AbstractString="") = CellImage(read(image); alt)
CellImage(image::IOBuffer; alt::AbstractString="") = CellImage(take!(copy(image)); alt)
function CellImage(bytes::Vector{UInt8}; alt::AbstractString="")
    ext = detect_image_ext(bytes)
    return CellImage(bytes, ext == ".jpg" ? :jpeg : Symbol(lstrip(ext, '.')), String(alt), nothing)
end

Base.:(==)(a::CellImage, b::CellImage) =
    a.bytes == b.bytes && a.format == b.format && a.alt == b.alt && a.url == b.url
Base.hash(a::CellImage, h::UInt) = hash(a.url, hash(a.alt, hash(a.format, hash(a.bytes, h))))

function Base.show(io::IO, img::CellImage)
    print(io, "CellImage(:", img.format, ", ", length(img.bytes), " bytes")
    isempty(img.alt) || print(io, ", alt=", repr(img.alt))
    isnothing(img.url) || print(io, ", url=", repr(img.url))
    print(io, ")")
end

function setdata!(ws::Worksheet, ref::CellRef, img::CellImage)
    _require_writable(ws, "put a picture in a cell")
    _require_worksheet(ws, "put a picture in a cell")
    vm, _ = _add_cell_image!(get_xlsxfile(ws), img)
    _set_rich_value_cell!(ws, ref, vm)
    return nothing
end

# Make the cell at `ref` show rich value `vm`, keeping its format.
function _set_rich_value_cell!(ws::Worksheet, ref::CellRef, vm::Integer)
    c = getcell(ws, ref)
    if !(c isa EmptyCell) && c.formula && get_formula_from_cache(ws, ref) isa ReferencedFormula
        rereference_formulae(ws, c)
    end
    style = c isa EmptyCell ? UInt32(0) : c.style
    setdata!(ws, Cell(ref, UInt64(vm), style, UInt16(0), CT_RICHVALUE, false))
    return nothing
end

"""
    getCellImage(ws::Worksheet, ref) -> Union{CellImage, Nothing}
    getCellImage(xf::XLSXFile, "Sheet1!B2") -> Union{CellImage, Nothing}

The picture in cell `ref`, as a [`CellImage`](@ref), or `nothing` if the cell holds
none. The picture is one placed in the cell, by Excel's *Place in Cell* or by
`sheet[ref] = CellImage(...)`, or one an `IMAGE` formula shows, for which `url` gives
its web address. Pictures floating over the cells are read with [`getImages`](@ref).

A cell holding a picture reads as `missing` through `getdata`, `readtable` and the
other functions that read values, since a picture isn't data.

# Examples
```julia
julia> img = XLSX.getCellImage(sheet, "B2")
CellImage(:png, 5120 bytes, alt="Logo")

julia> write("logo.png", img.bytes);
```

See also [`getCellImages`](@ref).
"""
function getCellImage end
getCellImage(xf::XLSXFile, ref::AbstractString) = getCellImage(_rowcol_sheet(xf, ref)...)
getCellImage(ws::Worksheet, ref::AbstractString) = getCellImage(ws, CellRef(ref))
getCellImage(ws::Worksheet, row::Integer, col::Integer) = getCellImage(ws, CellRef(row, col))
function getCellImage(ws::Worksheet, ref::CellRef)::Union{Nothing,CellImage}
    _require_worksheet(ws, "read a picture in a cell")
    cell = getcell(ws, ref)
    (cell isa Cell && cell.datatype == CT_RICHVALUE) || return nothing
    xf = get_xlsxfile(ws)
    return _cell_image(xf, _rich_values(xf), cell.value)
end

"""
    getCellImages(ws::Worksheet) -> Vector{NamedTuple}

Every picture in the sheet's cells, by row and then column, as `(cell, image)`
`NamedTuple`s of the cell's reference (`"B2"`) and its [`CellImage`](@ref).

See also [`getCellImage`](@ref).
"""
function getCellImages(ws::Worksheet)
    _require_worksheet(ws, "read the pictures in cells")
    xf = get_xlsxfile(ws)
    out = NamedTuple{(:cell, :image),Tuple{String,CellImage}}[]
    rich = nothing
    for r in eachrow(ws)
        for col in sort!(collect(keys(r.rowcells)))
            cell = r.rowcells[col]
            cell.datatype == CT_RICHVALUE || continue
            rich = something(rich, _rich_values(xf))
            img = _cell_image(xf, rich, cell.value)
            isnothing(img) || push!(out, (cell=string(cell.ref), image=img))
        end
    end
    return out
end

# Remove the picture in cell `ref`, keeping its format. Whether there was one.
function _delete_cell_image!(ws::Worksheet, ref::CellRef)::Bool
    isnothing(getCellImage(ws, ref)) && return false
    ws[ref] = missing
    return true
end
