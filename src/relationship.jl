
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
    xroot = xmlroot(xf, "_rels/.rels")[end]
    localname(xroot) != "Relationships" && throw(XLSXError("Malformed XLSX file $(xf.source). _rels/.rels root node name should be `Relationships`. Found $(XML.tag(xroot))."))
    if ("" => "http://schemas.openxmlformats.org/package/2006/relationships") ∉ get_namespaces(xroot)
        throw(XLSXError("Unexpected namespace at workbook relationship file: `$(get_namespaces(xroot))`."))
    end
    return xroot
end

function get_workbook_relationship_root(xf::XLSXFile)::XML.Node
    xroot = xmlroot(xf, "xl/_rels/workbook.xml.rels")[end]
    localname(xroot) != "Relationships" && throw(XLSXError("Malformed XLSX file $(xf.source). xl/_rels/workbook.xml.rels root node name should be `Relationships`. Found $(localname(xroot))."))
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

    c=XML.children(xroot[end])
    d = findfirst(r -> r["Target"] == rel.Target, c)
    deleteat!(c, d)
    new_rels=XML.Element("Relationships",  xmlns="http://schemas.openxmlformats.org/package/2006/relationships")
    for child in c
        push!(new_rels, child)
    end
    xroot[end]=new_rels
    xf.data["xl/_rels/workbook.xml.rels"]=xroot

end

#is_chartsheet(wb::Workbook, rid::String) = any(r.Id == rid && occursin("chartsheet", r.Type) for r in wb.relationships)
function is_chartsheet(wb::Workbook, sheetname::AbstractString)::Bool
    name = unquoteit(sheetname)
    xroot = get_xlsxfile(wb).data["xl/workbook.xml"][end]
    for node in XML.children(xroot)
        localname(node) != "sheets" && continue
        for sheet_node in XML.children(node)
            attrs = XML.attributes(sheet_node)
            isnothing(attrs) && continue
            get(attrs, "name", "") == name || continue
            rid = get(attrs, "r:id", "")
            return any(r.Id == rid && occursin("chartsheet", r.Type) for r in wb.relationships)
        end
    end
    return false
end

