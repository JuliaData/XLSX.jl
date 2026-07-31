
# Resolves a relationship Target that's relative to the directory containing
# the referencing part (e.g. "xl/worksheets"), collapsing ".." segments.
# Operates on zip-entry-style forward-slash strings only — deliberately not
# OS-path-aware, since these are in-memory dictionary keys, not filesystem paths.
function resolve_relative_target(base_dir::AbstractString, target::AbstractString)::String
    startswith(target, "/") && return String(target[nextind(target, begin):end])
    parts = split(base_dir, "/")
    for seg in split(target, "/")
        if seg == ".."
            isempty(parts) && throw(XLSXError("Malformed relationship target `$target` relative to `$base_dir`."))
            pop!(parts)
        elseif seg != "." && !isempty(seg)
            push!(parts, seg)
        end
    end
    return join(parts, "/")
end

function Relationship(wb::Workbook, e::XML.Node)::Relationship
    localname(e) !=   "Relationship" && throw(XLSXError("Unexpected XML Element: $(localname(e)). Expected: \"Relationship\"."))
    a = XML.attributes(e)
    return Relationship(
        a["Id"],
        a["Type"],
        a["Target"]
    )
end

function parse_relationship_target(prefix::String, target::String)::String
    isempty(prefix) || isempty(target) && throw(XLSXError("Something wrong here!"))
    if target[begin] == '/'
        sizeof(target) <= 1 && throw(XLSXError("Incomplete target path $target."))
        return target[nextind(target, begin):end]
    else
        return prefix * '/' * target
    end
end

function get_relationship_target_by_id(prefix::String, wb::Workbook, Id::String)::String
    for r in wb.relationships
        if Id == r.Id
            return parse_relationship_target(prefix, r.Target)
        end
    end
    throw(XLSXError("Relationship Id=$(Id) not found"))
end

function get_relationship_id_by_target(wb::Workbook, target::String)::String
    for r in wb.relationships
        if r.Type == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"
            if endswith(target, r.Target)
                return r.Id
            end
        end
    end
    throw(XLSXError("Target=$(target) not found"))
end

function get_relationship_target_by_type(prefix::String, wb::Workbook, _type_::String)::String
    for r in wb.relationships
        if _type_ == r.Type
            return parse_relationship_target(prefix, r.Target)
        end
    end
    throw(XLSXError("Relationship Type=$(_type_) not found"))
end

function has_relationship_by_type(wb::Workbook, _type_::String)::Bool
    for r in wb.relationships
        if _type_ == r.Type
            return true
        end
    end
    false
end

function get_package_relationship_root(xf::XLSXFile)::XML.Node
    xroot = xml_root_element(xmlroot(xf, "_rels/.rels"))
    XML.tag(xroot) != "Relationships" && throw(XLSXError("Malformed XLSX file $(xf.source). _rels/.rels root node name should be `Relationships`. Found $(XML.tag(xroot))."))
    if ("" => "http://schemas.openxmlformats.org/package/2006/relationships") ∉ get_namespaces(xroot)
        throw(XLSXError("Unexpected namespace at workbook relationship file: `$(get_namespaces(xroot))`."))
    end
    return xroot
end

function get_workbook_relationship_root(xf::XLSXFile)::XML.Node
    xroot = xml_root_element(xmlroot(xf, "xl/_rels/workbook.xml.rels"))
    XML.tag(xroot) != "Relationships" && throw(XLSXError("Malformed XLSX file $(xf.source). xl/_rels/workbook.xml.rels root node name should be `Relationships`. Found $(XML.tag(xroot))."))
    if ("" => "http://schemas.openxmlformats.org/package/2006/relationships") ∉ get_namespaces(xroot)
        throw(XLSXError("Unexpected namespace at workbook relationship file: `$(get_namespaces(xroot))`."))
    end
    return xroot
end

function new_relationship_id(rels_root::XML.Node)::String
    ids = [parse(Int, m[1])
           for n in element_children(rels_root)
           for m in [match(r"rId(\d+)", get_attr(n, "Id"))]
           if m !== nothing]
    return "rId$(isempty(ids) ? 1 : maximum(ids) + 1)"
end

# Adds new relationship. Returns new generated rId.
function add_relationship!(wb::Workbook, target::String, _type::String)::String
    xf    = get_xlsxfile(wb)
    xroot = get_workbook_relationship_root(xf)
    rId   = new_relationship_id(xroot) 

    push!(wb.relationships, Relationship(rId, _type, target))
    push!(xroot, XML.Element("Relationship"; Id=rId, Type=_type, Target=target))
    return rId
end

function delete_relationships!(xf::XLSXFile, rel::Relationship)
    #TODO renumber worksheet files in relationships - if necessary.

    xroot = xmlroot(xf, "xl/_rels/workbook.xml.rels")
    root_el = xml_root_element(xroot)

    c=XML.children(root_el)
    d = findfirst(r -> XML.nodetype(r) == XML.Element && r["Target"] == rel.Target, c)
    deleteat!(c, d)
    new_rels=XML.Element("Relationships",  xmlns="http://schemas.openxmlformats.org/package/2006/relationships")
    for child in xml_elements(root_el)
        push!(new_rels, child)
    end
    root_idx = findfirst(n -> XML.nodetype(n) == XML.Element, XML.children(xroot))
    xroot[root_idx]=new_rels
    xf.data["xl/_rels/workbook.xml.rels"]=xroot

end

#is_chartsheet(wb::Workbook, rid::String) = any(r.Id == rid && occursin("chartsheet", r.Type) for r in wb.relationships)
function is_chartsheet(wb::Workbook, sheetname::AbstractString)::Bool
    name = unquoteit(sheetname)
    xroot = xml_root_element(get_xlsxfile(wb).data["xl/workbook.xml"])
    for node in xml_elements(xroot)
        localname(node) != "sheets" && continue
        for sheet_node in xml_elements(node)
            attrs = XML.attributes(sheet_node)
            isnothing(attrs) && continue
            get(attrs, "name", "") == name || continue
            rid = get(attrs, "r:id", "")
            return any(r.Id == rid && occursin("chartsheet", r.Type) for r in wb.relationships)
        end
    end
    return false
end

# Splits "xl/worksheets/sheet1.xml" into ("xl/worksheets", "sheet1.xml").
# Manual split rather than Base.dirname/basename, matching resolve_relative_target's
# deliberate non-OS-path-aware treatment of these as forward-slash zip-entry keys.
function _split_zip_path(path::AbstractString)::Tuple{String,String}
    idx = findlast('/', path)
    isnothing(idx) && return ("", String(path))
    return (String(path[1:prevind(path, idx)]), String(path[nextind(path, idx):end]))
end

"""
    get_worksheet_relationship_target(xf::XLSXFile, ws::Worksheet, r_id::String) -> String

Resolve an `r:id` found inside `ws`'s own XML (e.g. a `<tablePart r:id="rId1"/>`)
to its target part path, via `ws`'s own relationship file
(`xl/worksheets/_rels/sheetN.xml.rels`), not the workbook-level relationships.
"""
function get_worksheet_relationship_target(xf::XLSXFile, ws::Worksheet, r_id::String)::String
    wb = get_workbook(xf)
    sheet_file = get_relationship_target_by_id("xl", wb, ws.relationship_id)
    dir, fname = _split_zip_path(sheet_file)
    rels_file = isempty(dir) ? "_rels/$fname.rels" : "$dir/_rels/$fname.rels"

    !internal_xml_file_exists(xf, rels_file) &&
        throw(XLSXError("Worksheet $sheet_file references relationship `$r_id` but no relationship file `$rels_file` exists in the package."))

    rels_root = xml_root_element(xmlroot(xf, rels_file))
    XML.tag(rels_root) != "Relationships" &&
        throw(XLSXError("Malformed $rels_file: root node name should be `Relationships`. Found $(XML.tag(rels_root))."))

    for el in xml_elements(rels_root)
        localname(el) != "Relationship" && continue
        attrs = XML.attributes(el)
        (isnothing(attrs) || get(attrs, "Id", nothing) != r_id) && continue
        return resolve_relative_target(dir, attrs["Target"])
    end

    throw(XLSXError("Relationship Id=$r_id not found in $rels_file"))
end

function next_table_id!(wb::Workbook)::Int
    if isnothing(wb.next_table_id)
        max_id = 0
        for ws in wb.sheets
            is_chartsheet(wb, ws.name) && continue
            for t in tables(ws)
                max_id = max(max_id, t.id)
            end
        end
        wb.next_table_id = max_id
    end
    wb.next_table_id += 1
    return wb.next_table_id
end

function new_table_filename(xf::XLSXFile)::String
    i = 1
    while haskey(xf.files, "xl/tables/table$(i).xml")
        i += 1
    end
    return "xl/tables/table$(i).xml"
end

function get_or_create_worksheet_rels!(xf::XLSXFile, sheet_path::String)
    sheet_dir, sheet_file = rsplit(sheet_path, "/"; limit=2)
    rels_path = "$sheet_dir/_rels/$sheet_file.rels"
    if !haskey(xf.data, rels_path)
        xf.data[rels_path]  = empty_rels_doc()
        xf.files[rels_path] = true
    end
    return rels_path, root_element(xf.data[rels_path])
end

function make_relative_target(base_dir::AbstractString, target_path::AbstractString)::String
    base_parts   = split(base_dir, "/")
    target_parts = split(target_path, "/")
    n = 0
    while n < length(base_parts) && n < length(target_parts) - 1 && base_parts[n+1] == target_parts[n+1]
        n += 1
    end
    ups = length(base_parts) - n
    return join(vcat(fill("..", ups), target_parts[n+1:end]), "/")
end
