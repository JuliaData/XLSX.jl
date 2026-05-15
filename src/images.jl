# ===========================================================================
# Constants
# ===========================================================================

const REL_DRAWING =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing"
const REL_IMAGE =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
const NS_RELATIONSHIPS =
    "http://schemas.openxmlformats.org/package/2006/relationships"
const NS_XDR =
    "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
const NS_A =
    "http://schemas.openxmlformats.org/drawingml/2006/main"
const NS_R =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

const MIME_DRAWING =
    "application/vnd.openxmlformats-officedocument.drawing+xml"
const EXT_MIME = Dict(
    ".png"  => "image/png",
    ".jpg"  => "image/jpeg",
    ".jpeg" => "image/jpeg",
    ".gif"  => "image/gif",
)

const ImageInfo = NamedTuple{
    (:sheet, :media_name, :from, :to),
    Tuple{String, String, String, String},
}

# ===========================================================================
# Traversal helpers  (eliminate the repeated nodetype/tag/attributes pattern)
# ===========================================================================

element_children(node::XML.Node) =
    filter(n -> XML.nodetype(n) === XML.Element, something(XML.children(node), []))

# Match on local name only (ignores namespace prefix)
elements_with_tag(node::XML.Node, tag::String) =
    filter(n -> localname(XML.tag(n)) == tag, element_children(node))

get_attr(node::XML.Node, key::AbstractString, default::AbstractString = "") =
    something(get(XML.attributes(node), key, nothing), default)

function root_element(doc::XML.Node)::XML.Node
    children = something(XML.children(doc), [])
    idx = findfirst(n -> XML.nodetype(n) === XML.Element, children)
    idx !== nothing ? children[idx] : throw(XLSXError("Document has no root element"))
end

function _text_value(node::XML.Node)::Union{Nothing,String}
    for c in something(XML.children(node), [])
        XML.nodetype(c) === XML.Text && return XML.value(c)
    end
    return nothing
end

# Prepends prefix if non-empty: prefixed_tag("pkg", "Relationship") → "pkg:Relationship"
prefixed_tag(prefix::AbstractString, name::AbstractString) =
    isempty(prefix) ? name : "$prefix:$name"

# ===========================================================================
# Document templates
# ===========================================================================

empty_rels_doc() = XML.Document(
    XML.Declaration(; version="1.0", encoding="UTF-8"),
    XML.Element("Relationships"; xmlns=NS_RELATIONSHIPS),
)

empty_drawing_doc() = XML.Document(
    XML.Declaration(; version="1.0", encoding="UTF-8"),
    XML.Element("xdr:wsDr";
        var"xmlns:xdr" = NS_XDR,
        var"xmlns:a"   = NS_A,
        var"xmlns:r"   = NS_R,
    ),
)

# ===========================================================================
# addImage — public API
# ===========================================================================

"""
    addImage(s::Worksheet, ref::AbstractString, image::Union{AbstractString, IOBuffer}; size::Union{Nothing, Tuple{<:Integer, <:Integer}}=nothing)
    addImage(s::Worksheet, row::Integer, col::Integer, image::Union{AbstractString, IOBuffer}; size::Union{Nothing, Tuple{<:Integer, <:Integer}}=nothing)


Insert an image into a worksheet at the given cell reference. The image "floats" above the 
grid and does not affect cell contents or dimensions. In Excel, the image may be resized 
and repositioned by the user as normal.
Supports file paths and `IOBuffer` sources.

If multiple, overlapping images are added, newer images overly older ones.

# Arguments

- `s::Worksheet`: the target worksheet.
- `ref::AbstractString`: Either a valid cell reference (e.g. `"A1"`) or a valid cell range (e.g. `"B2:D4"`). 
The image will be anchored to the top left of the reference and sized to fit within the reference bounds. 
If a cell range is given, the `size` keyword argument is ignored.

- `image::Union{AbstractString, IOBuffer}`: Specifies the image to be inserted. Either:  
  - a file path (`String`)  
  - an `IOBuffer` containing raw image bytes  

Supported formats (auto-detected): PNG, JPEG, GIF.

# Keyword Arguments

- `size`: provide the desired size of the image as a tuple of integers: `(width_px, height_px)`. Actual size 
will snap to the nearest actual cell boundaries. If `nothing` (default), the image's native pixel size is used. 
Ignored if `ref` is a cell range.

# Return Value

Returns a structured summary describing where and how the image was placed as a `NamedTuple` of `String` values:

```julia
(
    sheet      = sheet name,
    media_name = internal media file name,
    from       = Start cell (top left),
    to         = End cell (bottom right),
)
```

# Examples

Insert from a file:

```julia
info = XLSX.addImage(sheet, "B2", "photo.jpg")
```

Insert from an `IOBuffer`:

```julia
buf = IOBuffer(read("logo.png"))
info = XLSX.addImage(sheet, "C5", buf)
```

Insert with explicit size:

```julia
info = XLSX.addImage(sheet, "A1", "icon.png"; size=(128, 128))
```

"""
addImage(s::Worksheet, row::Integer, col::Integer, image; kw...) =
    addImage(s, CellRef(row, col), image; kw...)

function addImage(s::Worksheet, ref::AbstractString, image; kw...)
    if is_valid_cellname(ref)
        addImage(s, CellRef(ref), image; kw...)
    elseif is_valid_cellrange(ref)
        addImage(s, CellRange(ref), image; kw...)
    else
        throw(ArgumentError("Invalid cell reference: $ref"))
    end
end

function addImage(
    s::Worksheet,
    cellref::Union{CellRef, CellRange},
    image::Union{AbstractString, IOBuffer};
    size::Union{Nothing, Tuple{<:Integer,<:Integer}} = nothing,
)
    xf         = get_xlsxfile(s)
    sheet_path = get_relationship_target_by_id("xl", get_workbook(s), s.relationship_id)

    media_name   = add_media!(xf, image)
    drawing_path = ensure_drawing!(xf, sheet_path)
    img_rid      = add_image_rel!(xf, drawing_path, media_name)
    col, row, col_to, row_to =
        add_anchor!(xf, drawing_path, img_rid, media_name, cellref; size)

    return (
        sheet      = s.name,
        media_name = media_name,
        from       = string(CellRef(row,    col)),
        to         = string(CellRef(row_to, col_to)),
    )
end

# ===========================================================================
# Media
# ===========================================================================

add_media!(xf::XLSXFile, path::AbstractString) = _add_media_bytes!(xf, read(path))
add_media!(xf::XLSXFile, io::IOBuffer)         = _add_media_bytes!(xf, take!(io))

function _add_media_bytes!(xf::XLSXFile, bytes::Vector{UInt8})::String
    ext      = detect_image_ext(bytes)
    existing = count(k -> startswith(k, "xl/media/"), keys(xf.binary_data))
    name     = "image$(existing + 1)$ext"
    xf.binary_data["xl/media/$name"] = bytes
    ext_no_dot = String(lstrip(ext, '.'))
    register_content_type!(xf, "[Content_Types].xml";
                           tag="Default", key="Extension", val=ext_no_dot,
                           content_type=get(EXT_MIME, ext, "image/$ext_no_dot"))
    return name
end

# ===========================================================================
# Drawing setup
# ===========================================================================

function ensure_drawing!(xf::XLSXFile, sheet_path::String)::String
    sheet_dir, sheet_file = rsplit(sheet_path, "/"; limit=2)
    rels_path = "$sheet_dir/_rels/$sheet_file.rels"

    if !haskey(xf.data, rels_path)
        xf.data[rels_path]  = empty_rels_doc()
        xf.files[rels_path] = true
    end
    rels_root = root_element(xf.data[rels_path])

    # Return existing drawing path if already linked
    for node in elements_with_tag(rels_root, "Relationship")
        if get_attr(node, "Type") == REL_DRAWING
            drawing_file = rsplit(get_attr(node, "Target"), "/"; limit=2)[2]
            return "xl/drawings/$drawing_file"
        end
    end

    # Create a new drawing
    i = 1
    while haskey(xf.data, "xl/drawings/drawing$i.xml"); i += 1; end
    drawing_file = "drawing$i.xml"
    drawing_path = "xl/drawings/$drawing_file"

    xf.data[drawing_path]  = empty_drawing_doc()
    xf.files[drawing_path] = true

    rid = new_relationship_id(rels_root)
    pfx = get_prefix(rels_path, xf)
    push!(rels_root, XML.Element(prefixed_tag(pfx, "Relationship");
        Id     = rid,
        Type   = REL_DRAWING,
        Target = "../drawings/$drawing_file",
    ))

    ensure_drawing_element!(xf, xf.data[sheet_path], sheet_path, rid)
    register_content_type!(xf, "[Content_Types].xml";
                           tag="Override", key="PartName", val="/$drawing_path",
                           content_type=MIME_DRAWING)
    return drawing_path
end

function add_image_rel!(xf::XLSXFile, drawing_path::String, media_name::String)::String
    drawing_file = rsplit(drawing_path, "/"; limit=2)[2]
    rels_path    = "xl/drawings/_rels/$drawing_file.rels"

    if !haskey(xf.data, rels_path)
        xf.data[rels_path]  = empty_rels_doc()
        xf.files[rels_path] = true
    end
    rels_root = root_element(xf.data[rels_path])

    # Reuse existing rel if the same media is already referenced
    for node in elements_with_tag(rels_root, "Relationship")
        get_attr(node, "Target") == "../media/$media_name" && return get_attr(node, "Id")
    end

    rid = new_relationship_id(rels_root)
    pfx = get_prefix(rels_path, xf)
    push!(rels_root, XML.Element(prefixed_tag(pfx, "Relationship");
        Id     = rid,
        Type   = REL_IMAGE,
        Target = "../media/$media_name",
    ))
    return rid
end

# ===========================================================================
# Anchor
# ===========================================================================

function add_anchor!(
    xf::XLSXFile,
    drawing_path::String,
    img_rid::String,
    media_name::String,
    cellref::Union{CellRef, CellRange};
    size::Union{Nothing,Tuple{<:Integer,<:Integer}} = nothing,
)
    # Convention: col/row/col_to/row_to are 1-based inclusive throughout.
    # build_two_cell_anchor takes 0-based (from inclusive, to exclusive).
    # 1-based inclusive → 0-based inclusive:  n - 1
    # 1-based inclusive → 0-based exclusive:  n  (unchanged, since excl = incl + 1 - 1)

    if cellref isa CellRef
        col, row   = column_number(cellref), row_number(cellref)
        bytes      = xf.binary_data["xl/media/$media_name"]
        w_px, h_px = size !== nothing ? size : image_dimensions(bytes)
        col_to     = col + max(1, round(Int, w_px / 64)) - 1
        row_to     = row + max(1, round(Int, h_px / 20)) - 1
    else
        col,    row    = column_number(cellref.start), row_number(cellref.start)
        col_to, row_to = column_number(cellref.stop),  row_number(cellref.stop)
    end

    root_el   = root_element(xf.data[drawing_path])
    n_anchors = count(_ -> true, element_children(root_el))

    push!(root_el, build_two_cell_anchor(
        col - 1, row - 1,  # 0-based inclusive from
        col_to,  row_to,   # 0-based exclusive to
        img_rid;
        shape_id = n_anchors + 2,
    ))

    return col, row, col_to, row_to
end

# ===========================================================================
# Relationship / content-type helpers
# ===========================================================================

function register_content_type!(
    xf::XLSXFile,
    path::AbstractString;
    tag::AbstractString, key::AbstractString, val::AbstractString, content_type::AbstractString,
)::Nothing
    ct_root = root_element(xf.data[path])
    pfx     = get_prefix(path, xf)
    any(n -> localname(XML.tag(n)) == tag && get_attr(n, key) == val,
        element_children(ct_root)) && return nothing
    push!(ct_root, XML.Element(prefixed_tag(pfx, tag); Symbol(key) => val, ContentType=content_type))
    return nothing
end

function ensure_drawing_element!(xf::XLSXFile, sheet_doc::XML.Node, sheet_path::String, rid::String)
    sheet_root = root_element(sheet_doc)
    any(n -> localname(XML.tag(n)) == "drawing", element_children(sheet_root)) && return nothing
    if !haskey(something(XML.attributes(sheet_root), Dict()), "xmlns:r")
        sheet_root["xmlns:r"] = NS_R
    end
    pfx = get_prefix(sheet_path, xf)
    el  = XML.Element(prefixed_tag(pfx, "drawing"))
    el["r:id"] = rid
    push!(sheet_root, el)
    return nothing
end

# ===========================================================================
# Low-level XML builder
# ===========================================================================

function build_two_cell_anchor(
    col::Int, row::Int,       # 0-based inclusive
    col_to::Int, row_to::Int, # 0-based exclusive
    img_rid::String;
    shape_id::Int,
)::XML.Node
    tel(tag, text) = XML.Element(tag, XML.Text(text))

    function cell_marker(tag, c, r)
        XML.Element(tag,
            tel("xdr:col",    string(c)),
            tel("xdr:colOff", "0"),
            tel("xdr:row",    string(r)),
            tel("xdr:rowOff", "0"),
        )
    end

    blip = XML.Element("a:blip")
    blip["r:embed"] = img_rid

    return XML.Element("xdr:twoCellAnchor",
        cell_marker("xdr:from", col,    row),
        cell_marker("xdr:to",   col_to, row_to),
        XML.Element("xdr:pic",
            XML.Element("xdr:nvPicPr",
                XML.Element("xdr:cNvPr"; id=string(shape_id), name="Image $shape_id"),
                XML.Element("xdr:cNvPicPr",
                    XML.Element("a:picLocks"; noChangeAspect="1"),
                ),
            ),
            XML.Element("xdr:blipFill",
                blip,
                XML.Element("a:stretch", XML.Element("a:fillRect")),
            ),
            XML.Element("xdr:spPr",
                XML.Element("a:xfrm",
                    XML.Element("a:off";  x="0", y="0"),
                    XML.Element("a:ext"; cx="0", cy="0"),
                ),
                XML.Element("a:prstGeom", XML.Element("a:avLst"); prst="rect"),
            ),
        ),
        XML.Element("xdr:clientData"),
    )
end

# ===========================================================================
# Image format detection
# ===========================================================================

function detect_image_ext(bytes::Vector{UInt8})::String
    length(bytes) ≥ 8 &&
        bytes[1:8] == UInt8[0x89,0x50,0x4E,0x47,0x0D,0x0A,0x1A,0x0A] && return ".png"
    length(bytes) ≥ 2 &&
        bytes[1] == 0xFF && bytes[2] == 0xD8                           && return ".jpg"
    length(bytes) ≥ 4 &&
        bytes[1:4] == UInt8[0x47,0x49,0x46,0x38]                       && return ".gif"
    throw(XLSXError("Unsupported or unknown image format"))
end

function image_dimensions(bytes::Vector{UInt8})::Tuple{Int,Int}
    # PNG: width/height in bytes 17–20 and 21–24
    if length(bytes) ≥ 24 &&
            bytes[1:8] == UInt8[0x89,0x50,0x4E,0x47,0x0D,0x0A,0x1A,0x0A]
        w = Int(bytes[17]) << 24 | Int(bytes[18]) << 16 |
            Int(bytes[19]) << 8  | Int(bytes[20])
        h = Int(bytes[21]) << 24 | Int(bytes[22]) << 16 |
            Int(bytes[23]) << 8  | Int(bytes[24])
        return (w, h)
    end
    # GIF: little-endian 16-bit at bytes 7–10
    if length(bytes) ≥ 10 && bytes[1:4] == UInt8[0x47,0x49,0x46,0x38]
        return (Int(bytes[7]) | Int(bytes[8]) << 8,
                Int(bytes[9]) | Int(bytes[10]) << 8)
    end
    # JPEG: scan for SOF marker
    if length(bytes) ≥ 2 && bytes[1] == 0xFF && bytes[2] == 0xD8
        i = 3
        while i + 8 ≤ length(bytes)
            bytes[i] == 0xFF || break
            marker = bytes[i+1]
            if marker in 0xC0:0xC3
                h = Int(bytes[i+5]) << 8 | Int(bytes[i+6])
                w = Int(bytes[i+7]) << 8 | Int(bytes[i+8])
                return (w, h)
            end
            i += 2 + (Int(bytes[i+2]) << 8 | Int(bytes[i+3]))
        end
        throw(XLSXError("Could not find JPEG SOF marker"))
    end
    throw(XLSXError("Unsupported image format for dimension extraction"))
end

# ===========================================================================
# getImages — public API
# ===========================================================================

function getImages(s::Worksheet)::Vector{ImageInfo}
    xf         = get_xlsxfile(s)
    sheet_path = get_relationship_target_by_id("xl", get_workbook(s), s.relationship_id)
    return _images_for_sheet(xf, sheet_path, s.name)
end

function getImages(xf::XLSXFile)::Vector{ImageInfo}
    wb = get_workbook(xf)
    return reduce(vcat, [
        _images_for_sheet(xf,
            get_relationship_target_by_id("xl", wb, sheet.relationship_id),
            sheet.name)
        for sheet in wb.sheets
    ]; init=ImageInfo[])
end

function _images_for_sheet(xf::XLSXFile, sheet_path::String, sheet_name::String)::Vector{ImageInfo}
    drawing_path = _drawing_path_for_sheet(xf, sheet_path)
    drawing_path === nothing && return ImageInfo[]
    return _images_for_drawing(xf, drawing_path, sheet_name)
end

function _drawing_path_for_sheet(xf::XLSXFile, sheet_path::String)::Union{Nothing,String}
    sheet_dir, sheet_file = rsplit(sheet_path, "/"; limit=2)
    rels_path = "$sheet_dir/_rels/$sheet_file.rels"
    haskey(xf.data, rels_path) || return nothing

    for node in elements_with_tag(root_element(xf.data[rels_path]), "Relationship")
        if get_attr(node, "Type") == REL_DRAWING
            drawing_file = rsplit(get_attr(node, "Target"), "/"; limit=2)[2]
            return "xl/drawings/$drawing_file"
        end
    end
    return nothing
end

function _images_for_drawing(xf::XLSXFile, drawing_path::String, sheet_name::String)::Vector{ImageInfo}
    haskey(xf.data, drawing_path) || return ImageInfo[]
    drawing_file = rsplit(drawing_path, "/"; limit=2)[2]
    rels_path    = "xl/drawings/_rels/$drawing_file.rels"
    haskey(xf.data, rels_path) || return ImageInfo[]

    rid_to_media = _rid_to_media(xf.data[rels_path])
    return filter(!isnothing, [
        _parse_anchor(node, rid_to_media, sheet_name)
        for node in elements_with_tag(root_element(xf.data[drawing_path]), "twoCellAnchor")
    ])
end

function _rid_to_media(rels_doc::XML.Node)::Dict{String,String}
    Dict(
        get_attr(n, "Id") => rsplit(get_attr(n, "Target"), "/"; limit=2)[2]
        for n in elements_with_tag(root_element(rels_doc), "Relationship")
        if get_attr(n, "Type") == REL_IMAGE && !isempty(get_attr(n, "Id"))
    )
end

function _parse_anchor(
    anchor::XML.Node,
    rid_to_media::Dict{String,String},
    sheet_name::String,
)::Union{Nothing,ImageInfo}
    from_ref   = _parse_cell_marker(anchor, "from"; is_to=false)
    to_ref     = _parse_cell_marker(anchor, "to";   is_to=true)
    rid        = _find_blip_rid(anchor)
    media_name = rid !== nothing ? get(rid_to_media, rid, nothing) : nothing
    (from_ref === nothing || to_ref === nothing || media_name === nothing) && return nothing
    return (sheet=sheet_name, media_name=media_name, from=from_ref, to=to_ref)
end

function _parse_cell_marker(anchor::XML.Node, tag::String; is_to::Bool)::Union{Nothing,String}
    marker = nothing
    for n in element_children(anchor)
        localname(XML.tag(n)) == tag && (marker = n; break)
    end
    marker === nothing && return nothing
    vals = Dict(localname(XML.tag(c)) => _text_value(c) for c in element_children(marker))
    col  = get(vals, "col", nothing)
    row  = get(vals, "row", nothing)
    (col === nothing || row === nothing) && return nothing
    adj = is_to ? 0 : 1
    return string(CellRef(parse(Int, row) + adj, parse(Int, col) + adj))
end

function _find_blip_rid(node::XML.Node)::Union{Nothing,String}
    XML.nodetype(node) === XML.Element || return nothing
    if localname(XML.tag(node)) == "blip"
        attrs = XML.attributes(node)
        attrs === nothing && return nothing
        return something(get(attrs, "r:embed", nothing),
                         get(attrs, "{$(NS_R)}embed", nothing),
                         nothing)
    end
    for child in something(XML.children(node), [])
        rid = _find_blip_rid(child)
        rid !== nothing && return rid
    end
    return nothing
end