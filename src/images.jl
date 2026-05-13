"""
    addImage(s::Worksheet, ref, image; size=nothing, z_layer=nothing)


Insert an image into a worksheet at the given cell reference.  
Supports file paths and `IOBuffer` sources, auto‑detects image format, creates drawing parts and relationships as needed, and computes the anchor size using either native pixel dimensions or a user‑supplied `(width_px, height_px)`.

Returns a structured summary describing where and how the image was placed.

---

# **Arguments**

- **`s::Worksheet`**  
  Target worksheet.

- **`ref`**  
  Excel cell reference (`"A1"`, `"B2"`, …).  
  Must be valid; otherwise an error is thrown.

- **`image`**  
  Either:  
  - a file path (`String`)  
  - an `IOBuffer` containing raw image bytes  

  Supported formats (auto‑detected): PNG, JPEG, GIF.

---

# **Keyword Arguments**

- **`size = nothing`**  
  If `nothing`, the image’s native pixel size is used.  
  Otherwise supply `(width_px, height_px)`. Actual size will snap to the nearest actual cell boundaries.

- **`z_layer = nothing`**  
  Optional stacking order inside the drawing.

---

# **Return Value**

A `NamedTuple`:

```julia
(
    sheet      = sheet_name::String,
    media_name = media_name::String,
    from       = Start cell (top left)
    to         = End cell (bottom right),
)
```

Where:

- **`media_name`** is the filename created in `xl/media/`
- **`from`** and **`to`** are zero‑based anchor coordinates
- **`width_px`**, **`height_px`** are the pixel dimensions used for sizing

---

# **Examples**

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
addImage(
    s::Worksheet,
    ref::AbstractString,
    image::Union{AbstractString, IOBuffer};
    size::Union{Nothing, Tuple{<:Integer, <:Integer}} = nothing,
    z_layer::Union{Nothing, Integer} = nothing
) = addImage(s, CellRef(ref), image; size, z_layer)
addImage(
    s::Worksheet,
    row::Integer, col::Integer,
    image::Union{AbstractString, IOBuffer};
    size::Union{Nothing, Tuple{<:Integer, <:Integer}} = nothing,
    z_layer::Union{Nothing, Integer} = nothing
) = addImage(s, CellRef(row, col), image; size, z_layer)
function addImage(
    s::Worksheet,
    cellref::CellRef,
    image::Union{AbstractString, IOBuffer};
    size::Union{Nothing, Tuple{<:Integer, <:Integer}} = nothing,
    z_layer::Union{Nothing, Integer} = nothing
)
    xf         = s.package
    sheet_path = get_relationship_target_by_id("xl", get_workbook(s), s.relationship_id)
    sheet_name = s.name

    media_name   = add_media!(xf, image)
    drawing_path = ensure_drawing!(xf, sheet_path)
    img_rid      = add_image_rel!(xf, drawing_path, media_name)

    # Pass media_name directly — no need to re-derive it from the rels XML
    col, row, col_to, row_to =
        add_anchor!(xf, drawing_path, img_rid, media_name, cellref; size, z_layer)

    return (
        sheet      = sheet_name,
        media_name = media_name,
        from       = "$(CellRef(col, row))",
        to         = "$(CellRef(col_to, row_to))",
    )
end

function add_media!(xf::XLSXFile, image_path::AbstractString)::String
    return _add_media_bytes!(xf, read(image_path))
end
function add_media!(xf::XLSXFile, io::IOBuffer)::String
    return _add_media_bytes!(xf, take!(io))
end
function _add_media_bytes!(xf::XLSXFile, bytes::Vector{UInt8})::String
    ext      = detect_image_ext(bytes)
    existing = count(k -> startswith(k, "xl/media/"), keys(xf.binary_data))
    name     = "image$(existing + 1)$(ext)"
    xf.binary_data["xl/media/$name"] = bytes
    # binary_data has its own write loop in writexlsx — no xf.files entry needed
    register_image_content_type!(xf, ext)
    return name
end

function ensure_drawing!(xf::XLSXFile, sheet_path::String)::String
    # All keys in xf.data / xf.files use forward slashes — use rsplit, not
    # dirname/basename/joinpath, which use the OS separator on Windows.
    sheet_dir  = rsplit(sheet_path, "/"; limit=2)[1]   # "xl/worksheets"
    sheet_file = rsplit(sheet_path, "/"; limit=2)[2]   # "sheet1.xml"
    sheet_rels_path = sheet_dir * "/_rels/" * sheet_file * ".rels"

    if !haskey(xf.data, sheet_rels_path)
        xf.data[sheet_rels_path]  = empty_rels_doc()
        xf.files[sheet_rels_path] = true
    end
    rels_root = root_element(xf.data[sheet_rels_path])

    # Return existing drawing path if one is already linked to this sheet
    for node in something(XML.children(rels_root), [])
        XML.nodetype(node) === XML.Element || continue
        XML.tag(node) == "Relationship"    || continue
        attrs = XML.attributes(node)
        attrs === nothing && continue
        if get(attrs, "Type", "") == REL_DRAWING
            drawing_file = rsplit(get(attrs, "Target", ""), "/"; limit=2)[2]
            return "xl/drawings/" * drawing_file
        end
    end

    # Find a free drawing filename (avoids collisions with charts etc.)
    drawing_file = let i = 1
        while haskey(xf.data, "xl/drawings/drawing$i.xml"); i += 1; end
        "drawing$i.xml"
    end
    drawing_path = "xl/drawings/" * drawing_file

    xf.data[drawing_path]  = empty_drawing_doc()
    xf.files[drawing_path] = true

    rid = new_relationship_id(rels_root)
    push!(rels_root, XML.Element("Relationship";
        Id     = rid,
        Type   = REL_DRAWING,
        Target = "../drawings/" * drawing_file,
    ))

    ensure_drawing_element!(xf.data[sheet_path], rid)
    register_drawing_content_type!(xf, drawing_path)

    return drawing_path
end

function add_image_rel!(xf::XLSXFile, drawing_path::String, media_name::String)::String
    drawing_file = rsplit(drawing_path, "/"; limit=2)[2]
    rels_path    = "xl/drawings/_rels/" * drawing_file * ".rels"

    if !haskey(xf.data, rels_path)
        xf.data[rels_path]  = empty_rels_doc()
        xf.files[rels_path] = true
    end
    rels_root = root_element(xf.data[rels_path])

    # Return existing rid if this media is already referenced
    for node in something(XML.children(rels_root), [])
        XML.nodetype(node) === XML.Element || continue
        attrs = XML.attributes(node)
        attrs === nothing && continue
        get(attrs, "Target", "") == "../media/$media_name" && return get(attrs, "Id", "")
    end

    rid = new_relationship_id(rels_root)
    push!(rels_root, XML.Element("Relationship";
        Id     = rid,
        Type   = REL_IMAGE,
        Target = "../media/$media_name",
    ))
    return rid
end

function add_anchor!(
    xf::XLSXFile,
    drawing_path::String,
    img_rid::String,
    media_name::String,          # passed directly — no rels lookup needed
    cellref::CellRef;
    size::Union{Nothing,Tuple{<:Integer,<:Integer}} = nothing,
    z_layer::Union{Nothing,Integer} = nothing,
)
    root_el = root_element(xf.data[drawing_path])

    col = column_number(cellref) - 1
    row = row_number(cellref) - 1

    bytes      = xf.binary_data["xl/media/" * media_name]
    w_px, h_px = size === nothing ? image_dimensions(bytes) : size

    col_to = col + max(1, round(Int, w_px / 64))
    row_to = row + max(1, round(Int, h_px / 20))

    # Unique id for cNvPr: count existing anchors in this drawing
    n_anchors = count(n -> XML.nodetype(n) === XML.Element,
                      something(XML.children(root_el), []))

    push!(root_el, build_two_cell_anchor(
        col, row, col_to, row_to, img_rid;
        shape_id = n_anchors + 2,   # cNvPr id must be ≥2 and unique per drawing
        z_layer  = z_layer,
    ))

    return col, row, col_to, row_to
end


# ---------------------------------------------------------------------------
# Low-level helpers
# ---------------------------------------------------------------------------

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

# ── XML document factories ───────────────────────────────────────────────────

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

# ── Root-element accessor ────────────────────────────────────────────────────

function root_element(doc::XML.Node)::XML.Node
    for node in something(XML.children(doc), [])
        XML.nodetype(node) === XML.Element && return node
    end
    throw(XLSXError("Document has no root element"))
end

# ── Relationship helpers ─────────────────────────────────────────────────────

"""Choose the next available `rId<N>` by scanning existing Id attributes."""
function new_relationship_id(rels_root::XML.Node)::String
    ids = Int[]
    for node in something(XML.children(rels_root), [])
        XML.nodetype(node) === XML.Element || continue
        attrs = XML.attributes(node)
        attrs === nothing && continue
        m = match(r"rId(\d+)", get(attrs, "Id", ""))
        m !== nothing && push!(ids, parse(Int, m[1]))
    end
    return "rId$(isempty(ids) ? 1 : maximum(ids) + 1)"
end

# ── Content-type registration (split into two focused functions) ─────────────

function register_drawing_content_type!(xf::XLSXFile, drawing_path::String)::Nothing
    ct_root  = root_element(xf.data["[Content_Types].xml"])
    abs_path = "/" * drawing_path
    for node in something(XML.children(ct_root), [])
        XML.nodetype(node) === XML.Element || continue
        XML.tag(node) == "Override"        || continue
        attrs = XML.attributes(node)
        attrs !== nothing &&
            get(attrs, "PartName", "") == abs_path && return nothing
    end
    push!(ct_root, XML.Element("Override";
        PartName    = abs_path,
        ContentType = MIME_DRAWING,
    ))
    return nothing
end

function register_image_content_type!(xf::XLSXFile, image_ext::String)::Nothing
    ct_root    = root_element(xf.data["[Content_Types].xml"])
    ext_no_dot = lstrip(image_ext, '.')
    mime       = get(EXT_MIME, image_ext, "image/$ext_no_dot")
    for node in something(XML.children(ct_root), [])
        XML.nodetype(node) === XML.Element || continue
        XML.tag(node) == "Default"         || continue
        attrs = XML.attributes(node)
        attrs !== nothing &&
            get(attrs, "Extension", "") == ext_no_dot && return nothing
    end
    push!(ct_root, XML.Element("Default";
        Extension   = ext_no_dot,
        ContentType = mime,
    ))
    return nothing
end

# ── Sheet XML ────────────────────────────────────────────────────────────────

"""Append `<drawing r:id="rid"/>` to the sheet root if not already present."""
function ensure_drawing_element!(sheet_doc::XML.Node, rid::String)
    sheet_root = root_element(sheet_doc)
    for node in something(XML.children(sheet_root), [])
        XML.nodetype(node) === XML.Element || continue
        XML.tag(node) == "drawing" && return nothing
    end
    attrs = XML.attributes(sheet_root)
    if attrs === nothing || !haskey(attrs, "xmlns:r")
        sheet_root["xmlns:r"] = NS_R
    end
    el = XML.Element("drawing")
    el["r:id"] = rid
    push!(sheet_root, el)
    return nothing
end

# ── Anchor XML ───────────────────────────────────────────────────────────────

function build_two_cell_anchor(
        col::Int, row::Int,
        col_to::Int, row_to::Int,
        img_rid::String;
        shape_id::Int,
        z_layer = nothing)::XML.Node

    tel(tag, text) = XML.Element(tag, XML.Text(text))

    function cell_marker(tag, c, r)
        XML.Element(tag,
            tel("xdr:col",    string(c)),
            tel("xdr:colOff", "0"),
            tel("xdr:row",    string(r)),
            tel("xdr:rowOff", "0"),
        )
    end

    nvpicpr = XML.Element("xdr:nvPicPr",
        XML.Element("xdr:cNvPr"; id=string(shape_id), name="Image $shape_id"),
        XML.Element("xdr:cNvPicPr",
            XML.Element("a:picLocks"; noChangeAspect="1"),
        ),
    )

    blip = XML.Element("a:blip")
    blip["r:embed"] = img_rid   # colon in attribute name — set after construction
    blipfill = XML.Element("xdr:blipFill",
        blip,
        XML.Element("a:stretch", XML.Element("a:fillRect")),
    )

    sppr = XML.Element("xdr:spPr",
        XML.Element("a:xfrm",
            XML.Element("a:off";  x="0", y="0"),
            XML.Element("a:ext"; cx="0", cy="0"),
        ),
        XML.Element("a:prstGeom", XML.Element("a:avLst"); prst="rect"),
    )

    return XML.Element("xdr:twoCellAnchor",
        cell_marker("xdr:from", col,    row),
        cell_marker("xdr:to",   col_to, row_to),
        XML.Element("xdr:pic", nvpicpr, blipfill, sppr),
        XML.Element("xdr:clientData"),
    )
end

# ── Image format detection ───────────────────────────────────────────────────

function detect_image_ext(bytes::Vector{UInt8})::String
    length(bytes) ≥ 8 &&
        bytes[1:8] == UInt8[0x89,0x50,0x4E,0x47,0x0D,0x0A,0x1A,0x0A] && return ".png"
    length(bytes) ≥ 2 &&
        bytes[1] == 0xFF && bytes[2] == 0xD8                           && return ".jpg"
    length(bytes) ≥ 4 &&
        bytes[1:4] == UInt8[0x47,0x49,0x46,0x38]                       && return ".gif"
    throw(XLSXError("Unsupported or unknown image format"))
end

function image_dimensions(bytes::Vector{UInt8})
    # PNG: IHDR chunk begins at byte 9; width at 17-20, height at 21-24 (big-endian)
    if length(bytes) ≥ 24 &&
            bytes[1:8] == UInt8[0x89,0x50,0x4E,0x47,0x0D,0x0A,0x1A,0x0A]
        w = Int(bytes[17]) << 24 | Int(bytes[18]) << 16 | Int(bytes[19]) << 8 | Int(bytes[20])
        h = Int(bytes[21]) << 24 | Int(bytes[22]) << 16 | Int(bytes[23]) << 8 | Int(bytes[24])
        return (w, h)
    end
    # GIF: little-endian width at bytes 7-8, height at 9-10
    if length(bytes) ≥ 10 && bytes[1:4] == UInt8[0x47,0x49,0x46,0x38]
        return (Int(bytes[7]) | Int(bytes[8]) << 8,
                Int(bytes[9]) | Int(bytes[10]) << 8)
    end
    # JPEG: scan for SOF0/SOF1/SOF2 (0xFFC0–0xFFC3) markers
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