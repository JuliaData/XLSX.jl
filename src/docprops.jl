#
# -- Document properties: docProps/core.xml, docProps/app.xml and docProps/custom.xml
#

const _NS_CP      = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
const _NS_DC      = "http://purl.org/dc/elements/1.1/"
const _NS_DCTERMS = "http://purl.org/dc/terms/"
const _NS_XSI     = "http://www.w3.org/2001/XMLSchema-instance"
const _NS_APP     = "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"
const _NS_CUSTOM  = "http://schemas.openxmlformats.org/officeDocument/2006/custom-properties"
const _NS_PREFIXES = Dict(_NS_CP => "cp", _NS_DC => "dc", _NS_DCTERMS => "dcterms", _NS_XSI => "xsi",
                          DOCPROPS_VTYPES_NS => "vt")

# Each document part: its package relationship, usual path, content type and root element.
const _CORE_PART = (rel = "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties",
                    path = "docProps/core.xml", content_type = "application/vnd.openxmlformats-package.core-properties+xml",
                    root = """<cp:coreProperties xmlns:cp="$_NS_CP" xmlns:dc="$_NS_DC" xmlns:dcterms="$_NS_DCTERMS" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="$_NS_XSI"/>""")
const _APP_PART = (rel = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties",
                   path = "docProps/app.xml", content_type = "application/vnd.openxmlformats-officedocument.extended-properties+xml",
                   root = """<Properties xmlns="$_NS_APP" xmlns:vt="$DOCPROPS_VTYPES_NS"/>""")
const _CUSTOM_PART = (rel = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties",
                      path = "docProps/custom.xml", content_type = "application/vnd.openxmlformats-officedocument.custom-properties+xml",
                      root = """<Properties xmlns="$_NS_CUSTOM" xmlns:vt="$DOCPROPS_VTYPES_NS"/>""")

# The keywords of `setDocumentProperties` and their elements, each part's in the order Excel writes them.
const _CORE_PROPERTIES = (title = (_NS_DC, "title"), subject = (_NS_DC, "subject"), creator = (_NS_DC, "creator"),
                          keywords = (_NS_CP, "keywords"), description = (_NS_DC, "description"),
                          lastModifiedBy = (_NS_CP, "lastModifiedBy"), created = (_NS_DCTERMS, "created"),
                          modified = (_NS_DCTERMS, "modified"), category = (_NS_CP, "category"),
                          contentStatus = (_NS_CP, "contentStatus"))
const _CORE_ORDER = [(_NS_DC, "title"), (_NS_DC, "subject"), (_NS_DC, "creator"), (_NS_CP, "keywords"),
                     (_NS_DC, "description"), (_NS_CP, "lastModifiedBy"), (_NS_CP, "revision"), (_NS_CP, "lastPrinted"),
                     (_NS_DCTERMS, "created"), (_NS_DCTERMS, "modified"), (_NS_CP, "category"), (_NS_CP, "contentStatus")]
const _CORE_DATES = (:created, :modified)
const _APP_PROPERTIES = (manager = (_NS_APP, "Manager"), company = (_NS_APP, "Company"))
const _APP_ORDER = [(_NS_APP, n) for n in ("Template", "TotalTime", "Application", "DocSecurity", "ScaleCrop",
                    "HeadingPairs", "TitlesOfParts", "Manager", "Company", "LinksUpToDate", "SharedDoc",
                    "HyperlinkBase", "HyperlinksChanged", "AppVersion")]

# The format id Excel gives every custom property.
const _CUSTOM_FMTID = "{D5CDD505-2E9C-101B-9397-08002B2CF9AE}"
const _CUSTOM_NAME_MAX = 255

"""
    setDocumentProperties(xf::XLSXFile; title, subject, creator, keywords, description,
                          lastModifiedBy, category, contentStatus, created, modified,
                          company, manager)

Set a workbook's document properties, as Excel's *File → Info → Properties* does. Only the
keywords given are changed, and an empty string `""` removes a property.

| Keyword          | In Excel        |
|:-----------------|:----------------|
| `title`          | Title           |
| `subject`        | Subject         |
| `creator`        | Author          |
| `keywords`       | Tags            |
| `description`    | Comments        |
| `lastModifiedBy` | Last Modified By|
| `category`       | Categories      |
| `contentStatus`  | Status          |
| `created`        | Created         |
| `modified`       | Last Modified   |
| `company`        | Company         |
| `manager`        | Manager         |

`created` and `modified` are a `DateTime` or `Date` in UTC (Coordinated Universal Time), as
the file keeps them; Excel shows them in local time. The others are strings.

XLSX.jl doesn't change `modified` or `lastModifiedBy` when it saves a file; Excel does.

The file must be open for writing (`mode="rw"`).

# Examples
```julia
julia> XLSX.setDocumentProperties(xf; title="Quarterly sales", creator="Finance team", keywords="sales; Q3")

julia> XLSX.setDocumentProperties(xf; created=DateTime(2026, 1, 15, 9, 30))
```

See also [`getDocumentProperties`](@ref) and [`setCustomProperty`](@ref).
"""
function setDocumentProperties(xf::XLSXFile; kw...)
    _require_writable(xf, "set document properties")
    isempty(kw) && throw(XLSXError("Give the properties to set as keywords, such as `title=\"Sales\"`."))
    core = Pair{Tuple{String,String},Any}[]
    app = Pair{Tuple{String,String},Any}[]
    for (k, v) in kw
        if k in _CORE_DATES
            v isa Union{Dates.Date,Dates.DateTime} || throw(XLSXError("`$k` must be a `DateTime` or `Date`, in UTC. Found `$(repr(v))`."))
            push!(core, _CORE_PROPERTIES[k] => (_w3cdtf(v), true))
        elseif haskey(_CORE_PROPERTIES, k) || haskey(_APP_PROPERTIES, k)
            v isa AbstractString || throw(XLSXError("`$k` must be a string. Found `$(repr(v))`."))
            value = isempty(v) ? nothing : (String(v), false)
            haskey(_CORE_PROPERTIES, k) ? push!(core, _CORE_PROPERTIES[k] => value) : push!(app, _APP_PROPERTIES[k] => value)
        else
            throw(XLSXError("Unknown keyword `$k`. Must be one of $(join((keys(_CORE_PROPERTIES)..., keys(_APP_PROPERTIES)...), ", "))."))
        end
    end
    isempty(core) || _set_property_elements!(xf, _document_part!(xf, _CORE_PART), core, _CORE_ORDER)
    isempty(app) || _set_property_elements!(xf, _document_part!(xf, _APP_PART), app, _APP_ORDER)
    return nothing
end

"""
    getDocumentProperties(xf::XLSXFile) -> NamedTuple

A workbook's document properties: a `NamedTuple` with the fields described in
[`setDocumentProperties`](@ref), each `nothing` if the workbook doesn't have it. `created`
and `modified` are `DateTime`s in UTC, or the text as written if it isn't a date.

Works on files opened for reading only.

# Examples
```julia
julia> XLSX.getDocumentProperties(xf).title
"Quarterly sales"
```
"""
function getDocumentProperties(xf::XLSXFile)
    core = _property_texts(xf, _CORE_PART)
    app = _property_texts(xf, _APP_PART)
    function value(k)
        text = haskey(_CORE_PROPERTIES, k) ? get(core, _CORE_PROPERTIES[k], nothing) : get(app, _APP_PROPERTIES[k], nothing)
        (isnothing(text) || isempty(text)) && return nothing
        return k in _CORE_DATES ? something(_parse_w3cdtf(text), text) : text
    end
    ks = (keys(_CORE_PROPERTIES)..., keys(_APP_PROPERTIES)...)
    return NamedTuple{ks}(map(value, ks))
end

"""
    setCustomProperty(xf::XLSXFile, name::AbstractString, value)

Set the custom document property `name` to `value`, as Excel's *File → Info → Properties →
Advanced Properties → Custom* does, adding it if the workbook doesn't have it. A property
of the same name, ignoring case as Excel does, is replaced.

`value` is a string, a whole number, a real number, a `Bool`, or a `Date` or `DateTime` in
UTC. A whole number Excel can't keep as one (beyond ±2,147,483,647) is kept as a real number.

Other custom properties are kept as they are, including ones other programs write, such as
sensitivity labels.

The file must be open for writing (`mode="rw"`).

# Examples
```julia
julia> XLSX.setCustomProperty(xf, "Project", "Apollo")

julia> XLSX.setCustomProperty(xf, "Approved", true)
```

See also [`getCustomProperty`](@ref), [`getCustomProperties`](@ref) and
[`deleteCustomProperty`](@ref).
"""
function setCustomProperty(xf::XLSXFile, name::AbstractString, value)
    _require_writable(xf, "set a custom property")
    1 <= length(name) <= _CUSTOM_NAME_MAX ||
        throw(XLSXError("A custom property's name must be 1 to $_CUSTOM_NAME_MAX characters long. Found `$(repr(name))`."))
    vtype, text = _custom_vt(value)
    path = _document_part!(xf, _CUSTOM_PART)
    _update_part_root!(xf, path) do atts, nss, kids
        vt = _ns_prefix!(atts, nss, DOCPROPS_VTYPES_NS)
        i = _custom_property_index(kids, name)
        old = isnothing(i) ? Pair{String,String}[] : _atts_vector(kids[i])
        pid = isnothing(i) ? string(max(1, maximum((something(tryparse(Int, get(c, "pid", "")), 1) for c in kids
                                                    if XML.nodetype(c) == XML.Element); init=1)) + 1) : nothing
        new_atts = isnothing(i) ? ["fmtid" => _CUSTOM_FMTID, "pid" => pid, "name" => String(name)] : old
        tag = isnothing(i) ? _ns_prefix!(atts, nss, _NS_CUSTOM) * "property" : XML.tag(kids[i])
        el = _xml_element(tag, new_atts, [_xml_element(vt * vtype, (), [XML.Text(text)])])
        isnothing(i) ? push!(kids, el) : (kids[i] = el)
    end
    return nothing
end

"""
    getCustomProperty(xf::XLSXFile, name::AbstractString)

The value of the custom document property `name`, ignoring case, or `nothing` if the workbook
doesn't have it. See [`getCustomProperties`](@ref) for the types of the values.

Works on files opened for reading only.
"""
function getCustomProperty(xf::XLSXFile, name::AbstractString)
    for (k, v) in getCustomProperties(xf)
        _same_property_name(k, name) && return v
    end
    return nothing
end

"""
    getCustomProperties(xf::XLSXFile) -> OrderedDict{String,Any}

All of a workbook's custom document properties, in the order the file keeps them, by name.
Each value is a `String`, an `Int`, a `Float64`, a `Bool` or a `DateTime` (in UTC), as the
property is typed, and a list (which other programs can write) is a `Vector` of these. A
value of another type, such as currency, is given as its text.

Works on files opened for reading only.

# Examples
```julia
julia> XLSX.getCustomProperties(xf)
OrderedCollections.OrderedDict{String, Any} with 2 entries:
  "Project"  => "Apollo"
  "Approved" => true
```
"""
function getCustomProperties(xf::XLSXFile)::OrderedCollections.OrderedDict{String,Any}
    out = OrderedCollections.OrderedDict{String,Any}()
    root = _document_root(xf, _CUSTOM_PART)
    isnothing(root) && return out
    for p in xml_elements(root)
        localname(p) == "property" && haskey(p, "name") || continue
        els = xml_elements(p)
        v = isempty(els) ? nothing : first(els)
        out[p["name"]] = isnothing(v) ? nothing : _custom_value(v)
    end
    return out
end

"""
    deleteCustomProperty(xf::XLSXFile, name::AbstractString) -> Bool

Delete the custom document property `name`, ignoring case. Returns `false` if the workbook
doesn't have it.

The file must be open for writing (`mode="rw"`).
"""
function deleteCustomProperty(xf::XLSXFile, name::AbstractString)::Bool
    _require_writable(xf, "delete a custom property")
    path = _document_part_path(xf, _CUSTOM_PART)
    isnothing(path) && return false
    found = false
    _update_part_root!(xf, path) do _, _, kids
        i = _custom_property_index(kids, name)
        isnothing(i) || (deleteat!(kids, i); found = true)
    end
    return found
end

#
# -- Helpers
#

_same_property_name(a::AbstractString, b::AbstractString) = lowercase(a) == lowercase(b)

_custom_property_index(kids, name) = findfirst(c -> XML.nodetype(c) == XML.Element && localname(c) == "property" &&
                                                    _same_property_name(get(c, "name", ""), name), kids)

# The `vt:` type and text of a custom property's value, as Excel writes them.
_custom_vt(v::AbstractString) = ("lpwstr", String(v))
_custom_vt(v::Bool) = ("bool", v ? "true" : "false")
_custom_vt(v::Integer) = typemin(Int32) <= v <= typemax(Int32) ? ("i4", string(v)) : _custom_vt(Float64(v))
function _custom_vt(v::Real)
    isfinite(v) || throw(XLSXError("A custom property can't be `$v`."))
    return ("r8", string(Float64(v)))
end
_custom_vt(v::Union{Dates.Date,Dates.DateTime}) = ("filetime", _w3cdtf(v))
_custom_vt(v) = throw(XLSXError(
    "A custom property is a string, a number, a `Bool`, or a `Date` or `DateTime`. Found `$(repr(v))`."))

# A `vt:` value: a list's items, each in a `vt:variant` if they may differ in type.
function _custom_value(v::XML.Node)
    localname(v) == "vector" || return _custom_value(localname(v), _element_text(v))
    return Any[_custom_value(localname(e) == "variant" ? something(first(xml_elements(e)), e) : e) for e in xml_elements(v)]
end
function _custom_value(vtype::AbstractString, text::AbstractString)
    vtype in ("lpwstr", "lpstr", "bstr") && return String(text)
    vtype in ("i1", "i2", "i4", "i8", "int", "ui1", "ui2", "ui4", "ui8", "uint") &&
        return something(tryparse(Int, text), String(text))
    vtype in ("r4", "r8", "decimal") && return something(tryparse(Float64, text), String(text))
    vtype == "bool" && return lowercase(text) in ("true", "1")
    vtype in ("filetime", "date") && return something(_parse_w3cdtf(text), String(text))
    return String(text)
end

_w3cdtf(d::Dates.Date) = _w3cdtf(Dates.DateTime(d))
_w3cdtf(d::Dates.DateTime) = Dates.format(d, Dates.dateformat"yyyy-mm-ddTHH:MM:SS") * "Z"

# A W3C date and time, as document properties are written, as a `DateTime` in UTC.
function _parse_w3cdtf(s::AbstractString)::Union{Nothing,Dates.DateTime}
    m = match(r"^(\d{4})-(\d\d)-(\d\d)(?:T(\d\d):(\d\d)(?::(\d\d)(?:\.(\d+))?)?)?\s*(Z|[+-]\d\d:\d\d)?$", strip(s))
    isnothing(m) && return nothing
    n(i) = isnothing(m[i]) ? 0 : parse(Int, m[i])
    ms = isnothing(m[7]) ? 0 : parse(Int, rpad(first(m[7], 3), 3, '0'))
    d = try
        Dates.DateTime(n(1), n(2), n(3), n(4), n(5), n(6), ms)
    catch e
        e isa ArgumentError || rethrow()
        return nothing
    end
    z = m[8]
    if !isnothing(z) && z != "Z"   # a local time with its offset from UTC
        sgn = z[1] == '-' ? -1 : 1
        d -= Dates.Minute(sgn * (60 * parse(Int, z[2:3]) + parse(Int, z[5:6])))
    end
    return d
end

_element_text(el::XML.Node) = join(XML.value(t) for t in XML.children(el) if XML.nodetype(t) == XML.Text)

# The path of a document part, from the package relationships, or `nothing` if the file has none.
function _document_part_path(xf::XLSXFile, part)::Union{Nothing,String}
    for r in xf.relationships
        r.Type == part.rel || continue
        path = String(lstrip(r.Target, '/'))
        return haskey(xf.data, path) ? path : nothing
    end
    return nothing
end

function _document_root(xf::XLSXFile, part)
    path = _document_part_path(xf, part)
    return isnothing(path) ? nothing : xml_root_element(xf.data[path])
end

# The text of each element of a document part, by namespace and name.
function _property_texts(xf::XLSXFile, part)::Dict{Tuple{String,String},String}
    root = _document_root(xf, part)
    isnothing(root) && return Dict{Tuple{String,String},String}()
    nss = get_namespaces(root)
    return Dict(_element_key(c, nss) => _element_text(c) for c in xml_elements(root))
end

function _element_key(el::XML.Node, nss)::Tuple{String,String}
    pfx = _ns_prefix(XML.tag(el))
    return (get(nss, isnothing(pfx) ? "" : chop(pfx), ""), localname(el))
end

# The path of a document part, made with its relationship and content type if the file has none.
function _document_part!(xf::XLSXFile, part)::String
    path = _document_part_path(xf, part)
    isnothing(path) || return path
    path = part.path
    xf.data[path] = parse("""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n""" * part.root, XML.Node)
    xf.files[path] = true
    rels = get_package_relationship_root(xf)
    rid = new_relationship_id(rels)
    push!(rels, XML.Element(something(_ns_prefix(XML.tag(rels)), "") * "Relationship"; Id=rid, Type=part.rel, Target=path))
    push!(xf.relationships, Relationship(rid, part.rel, path))
    register_content_type!(xf, "[Content_Types].xml"; tag="Override", key="PartName", val="/$path", content_type=part.content_type)
    return path
end

# Rebuild the root element of the part at `path` from `f(atts, nss, kids)`, which changes its
# attributes, namespaces (prefix => URI) and children in place.
function _update_part_root!(f, xf::XLSXFile, path::String)
    chn = XML.children(xf.data[path])
    ri = findfirst(c -> XML.nodetype(c) == XML.Element, chn)
    root = chn[ri]
    atts = _atts_vector(root)
    nss = get_namespaces(root)
    kids = XML.Node[c for c in XML.children(root)]
    f(atts, nss, kids)
    chn[ri] = _xml_element(XML.tag(root), atts, kids)
    return nothing
end

# The prefix ("dc:", or "" for the default namespace) of namespace `ns` in a part's root,
# declared on the root with its usual prefix if it isn't yet.
function _ns_prefix!(atts::Vector{Pair{String,String}}, nss::Dict{String,String}, ns::String)::String
    for (p, uri) in nss
        uri == ns && return isempty(p) ? "" : p * ":"
    end
    p = _NS_PREFIXES[ns]
    push!(atts, "xmlns:$p" => ns)
    nss[p] = ns
    return p * ":"
end

# Set the text of elements of a document part, each given by namespace and name, a
# `nothing` value removing it. A date is written with its type, as Excel writes it. A new
# element goes where `order` puts it.
function _set_property_elements!(xf::XLSXFile, path::String, changes, order)
    _update_part_root!(xf, path) do atts, nss, kids
        rank(c) = XML.nodetype(c) == XML.Element ? findfirst(==(_element_key(c, nss)), order) : nothing
        for ((ns, name), value) in changes
            i = findfirst(c -> XML.nodetype(c) == XML.Element && _element_key(c, nss) == (ns, name), kids)
            if isnothing(value)
                isnothing(i) || deleteat!(kids, i)
                continue
            end
            text, isdate = value
            el_atts = isdate ? [_ns_prefix!(atts, nss, _NS_XSI) * "type" => _ns_prefix!(atts, nss, _NS_DCTERMS) * "W3CDTF"] :
                               Pair{String,String}[]
            el = _xml_element(_ns_prefix!(atts, nss, ns) * name, el_atts, [XML.Text(text)])
            if !isnothing(i)
                kids[i] = el
                continue
            end
            r = findfirst(==((ns, name)), order)
            j = findlast(c -> (k = rank(c); !isnothing(k) && k < r), kids)
            if isnothing(j)
                k = findfirst(c -> (k = rank(c); !isnothing(k) && k > r), kids)
                insert!(kids, something(k, length(kids) + 1), el)
            else
                insert!(kids, j + 1, el)
            end
        end
    end
    return nothing
end
