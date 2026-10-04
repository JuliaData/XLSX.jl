
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
           for n in xml_elements(rels_root)
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

get_or_create_worksheet_rels!(xf::XLSXFile, sheet_path::String) = get_or_create_rels!(xf, sheet_path)

function get_or_create_rels!(xf::XLSXFile, sheet_path::String)
    sheet_dir, sheet_file = rsplit(sheet_path, "/"; limit=2)
    rels_path = "$sheet_dir/_rels/$sheet_file.rels"
    if !haskey(xf.data, rels_path)
        xf.data[rels_path]  = empty_rels_doc()
        xf.files[rels_path] = true
    end
    return rels_path, xml_root_element(xf.data[rels_path])
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

# The final path segment of a relationship Target, swapped for a new filename.
# Avoids relative-path arithmetic: a clone always sits beside its original.
function _retarget(target::AbstractString, new_fname::AbstractString)::String
    i = findlast('/', target)
    return isnothing(i) ? String(new_fname) : target[1:i] * new_fname
end

# chart1.xml -> chart2.xml, chartEx1.xml -> chartEx2.xml, style1.xml -> style2.xml
function _next_part_name(xl::XLSXFile, dir::AbstractString, fname::AbstractString)::String
    stem = replace(fname, r"\d*\.xml$" => "")
    taken(p) = haskey(xl.data, p) || haskey(xl.files, p)
    i = 1
    while taken("$dir/$stem$i.xml"); i += 1; end
    return "$stem$i.xml"
end

function content_type_for_part(xf::XLSXFile, path::AbstractString)::Union{Nothing,String}
    haskey(xf.data, "[Content_Types].xml") || return nothing
    for n in elements_with_tag(xml_root_element(xf.data["[Content_Types].xml"]), "Override")
        String(lstrip(get_attr(n, "PartName"), '/')) == path && return get_attr(n, "ContentType")
    end
    return nothing
end

"""
Clone `path` and every part it exclusively owns, returning the clone's package
path.

Used when copying a sheet: a chart part belongs to exactly one drawing, so the
copy needs its own. The chart's own relationships (style, colors, theme
override) are owned the same way and are cloned with it. Images are not — media
is shared across drawings — so image relationships are left pointing at the
original.
"""
function clone_owned_part!(xl::XLSXFile, path::String)::String
    dir, fname = _split_zip_path(path)
    new_fname = _next_part_name(xl, dir, fname)
    new_path  = "$dir/$new_fname"

    xl.data[new_path]  = copynode(xl.data[path])
    xl.files[new_path] = true

    ct = content_type_for_part(xl, path)
    isnothing(ct) || register_content_type!(xl, "[Content_Types].xml";
                        tag="Override", key="PartName", val="/$new_path",
                        content_type=ct)

    old_rels = "$dir/_rels/$fname.rels"
    haskey(xl.data, old_rels) || return new_path

    new_rels = "$dir/_rels/$new_fname.rels"
    xl.data[new_rels]  = copynode(xl.data[old_rels])
    xl.files[new_rels] = true

    for n in elements_with_tag(xml_root_element(xl.data[new_rels]), "Relationship")
        get_attr(n, "TargetMode") == "External" && continue
        endswith(get_attr(n, "Type"), "/image") && continue
        target = get_attr(n, "Target")
        old_t  = resolve_relative_target(dir, target)
        haskey(xl.data, old_t) || continue
        new_t  = clone_owned_part!(xl, old_t)
        n["Target"] = _retarget(target, last(_split_zip_path(new_t)))
    end

    return new_path
end

# ---- chart source references: shared helpers --------------------------------

# A reference names the sheet in `prefix` (e.g. `Data!` or `'My Sheet'!`) when
# the prefix starts it: not preceded by a name character or quote, so `Data!`
# does not match inside `MyData!`. Formulas naming an external workbook never
# name a local sheet.
_sheet_ref_regex(prefix::String) = Regex("(?<![\\w.'])\\Q" * prefix * "\\E")
_names_sheet(s::AbstractString, prefix::String) =
    !occursin('[', s) && occursin(_sheet_ref_regex(prefix), s)
    
# `s` with every reference to the sheet in `old_prefix` pointed at `new_prefix`
# instead. Formulas naming an external workbook are left alone.
function _repoint_ref_string(s::AbstractString, old_prefix::String, new_prefix::String)::String
    occursin('[', s) && return String(s)
    return replace(s, _sheet_ref_regex(old_prefix) => new_prefix)
end

# Any element named `f` in a chart part is a formula reference: series data,
# titles, data labels in extLst, trendlines. Walking the whole tree rather than
# enumerating paths means the extLst cases are covered too. `g` maps each
# formula to its replacement; returning it unchanged leaves the element alone.
function _rewrite_chart_formulas!(g, node::XML.Node)
    for child in XML.eachelement(node)
        if localname(child) == "f"
            s = XML.is_simple_value(child)
            isnothing(s) && continue
            t = g(s)
            t == s || (child[end] = XML.Text(t))
        else
            _rewrite_chart_formulas!(g, child)
        end
    end
    return nothing
end

# ---- copying a sheet ---------------------------------------------------------

"""
Repoint every source reference in a chart part from `old_sheet` to `new_sheet`.

Used when copying a sheet: the cloned chart part still names the original
sheet, so without this the copy plots the original's data. References to other
sheets and to external workbooks are left alone.

The cached values are deliberately not touched: immediately after a copy they
are correct for both sheets, and clearing them would leave `getChartData`
empty for the copy.
"""
function repoint_chart_refs!(xl::XLSXFile, chart_path::String,
                             old_sheet::String, new_sheet::String)
    haskey(xl.data, chart_path) || return nothing
    old, new = quoteit(old_sheet) * "!", quoteit(new_sheet) * "!"
    _rewrite_chart_formulas!(s -> _repoint_ref_string(s, old, new),
                             xml_root_element(xl.data[chart_path]))
    return nothing
end

# Excel names a chartEx chart's sources `_xlchart.v<N>.<M>`, where N is a
# counter it bumps per batch of charts and M runs within the batch. Pick a
# fresh N so a copied chart's names cannot collide with the original's.
function _next_xlchart_series(wb::Workbook)::Int
    n = 0
    for name in keys(wb.workbook_names)
        m = match(r"^_xlchart\.v(\d+)\.\d+$", name)
        isnothing(m) && continue
        n = max(n, parse(Int, m.captures[1]))
    end
    return n + 1
end

"""
Repoint a cloned `chartEx` part's source references at `new_sheet`.

Unlike a `c:` chart, a chartEx chart does not name its ranges directly: each
`cx:f` holds a hidden defined name (`_xlchart.v1.0`) that resolves to the
range. Cloning the part alone therefore leaves the copy plotting the original
sheet, so each referenced name is cloned under a fresh series index with its
value repointed, and the `cx:f` rewritten to the new name.

References to other sheets, and anything that is not a resolvable workbook
defined name, are left alone.
"""
function repoint_chartex_refs!(xl::XLSXFile, chart_path::String,
                               old_sheet::String, new_sheet::String)
    haskey(xl.data, chart_path) || return nothing
    wb = get_workbook(xl)
    series = _next_xlchart_series(wb)
    counter = Ref(0)
    # Excel shares one _xlchart name between cx:f elements in the same part
    # (a Pareto chart's data blocks share their categories), so each name is
    # cloned once and every occurrence repointed to the same clone.
    cloned = Dict{String,String}()
    old, new = quoteit(old_sheet) * "!", quoteit(new_sheet) * "!"
    _rewrite_chart_formulas!(xml_root_element(xl.data[chart_path])) do s
        n = _clone_xlchart_name!(wb, String(s), old_sheet, new_sheet, series, counter, cloned)
        isnothing(n) ? _repoint_ref_string(s, old, new) : n    # direct reference, as stage 6 writes
    end
    return nothing
end

# ---- deleting a sheet --------------------------------------------------------

"""
Invalidate every chart reference to `sheet`, as Excel does when a sheet is deleted.

A `c:` chart keeps the reference's element and its cache: the whole `c:f` becomes
`#REF!`. A `chartEx` chart loses the reference entirely: a dimension whose `cx:f`
names the sheet is emptied to one `cx:lvl ptCount="0"` per level, and a `cx:tx`
whose `cx:f` names it is removed. `dead` maps each `_xlchart` defined name that
resolved to the sheet to its former value; the names themselves are deleted with
the sheet's other defined names, as Excel drops them too.
"""
function invalidate_chart_refs!(xf::XLSXFile, sheet::String, dead::Dict{String,String})
    prefix = quoteit(sheet) * "!"
    for path in parts_with_content_type(xf, MIME_CHART)
        haskey(xf.files, path) || continue        # removed with the sheet
        _rewrite_chart_formulas!(s -> _names_sheet(s, prefix) ? "#REF!" : s,
                                 xml_root_element(get_xml_data(xf, path)))
    end
    for path in parts_with_content_type(xf, MIME_CHARTEX)
        haskey(xf.files, path) || continue
        _empty_cx_refs!(xml_root_element(get_xml_data(xf, path)), prefix, dead)
    end
    return nothing
end

# cx: — the reference an `f` resolves to, if it names the deleted sheet, either
# through a dead `_xlchart` name or directly (as stage 6 writes); else `nothing`.
function _dead_cx_ref(f, prefix::String, dead::Dict{String,String})
    isnothing(f) && return nothing
    s = XML.is_simple_value(f)
    isnothing(s) && return nothing
    haskey(dead, s) && return dead[s]
    _names_sheet(s, prefix) && return String(s)
    return nothing
end

# Levels in a reference: its columns, or its rows when the `cx:f` has dir="row".
function _ref_levels(ref::AbstractString, dir)::Int
    m = match(r"\$?([A-Z]{1,3})\$?(\d*)(?::\$?([A-Z]{1,3})\$?(\d*))?$", ref)
    (isnothing(m) || isnothing(m[3])) && return 1
    if dir == "row"
        (isempty(m[2]) || isempty(m[4])) && return 1
        return abs(parse(Int, m[4]) - parse(Int, m[2])) + 1
    end
    return abs(decode_column_number(m[3]) - decode_column_number(m[1])) + 1
end

# `name` with the prefix `node`'s own tag carries.
function _same_prefix_tag(node::XML.Node, name::String)::String
    t = XML.tag(node)
    i = findfirst(':', t)
    return isnothing(i) ? name : t[1:i] * name
end

function _empty_cx_refs!(node::XML.Node, prefix::String, dead::Dict{String,String})
    drop = XML.Node[]
    for child in XML.eachelement(node)
        ln = localname(child)
        if ln == "strDim" || ln == "numDim"
            f = first_element_with_tag(child, "f")
            ref = _dead_cx_ref(f, prefix, dead)
            if !isnothing(ref)
                n   = _ref_levels(ref, get_attr(f, "dir"))
                lvl = _same_prefix_tag(f, "lvl")
                empty!(XML.children(child))
                for _ in 1:n
                    push!(child, XML.Element(lvl; ptCount = "0"))
                end
            end
        elseif ln == "tx"
            txd = first_element_with_tag(child, "txData")
            f   = isnothing(txd) ? nothing : first_element_with_tag(txd, "f")
            isnothing(_dead_cx_ref(f, prefix, dead)) || push!(drop, child)
        else
            _empty_cx_refs!(child, prefix, dead)
        end
    end
    isempty(drop) || filter!(c -> !any(d -> d === c, drop), XML.children(node))
    return nothing
end

# `nothing` when the reference is not a workbook defined name pointing at
# `old_sheet`, so the caller leaves it untouched.
function _clone_xlchart_name!(wb::Workbook, ref::String, old_sheet::String,
                              new_sheet::String, series::Int, counter::Ref{Int},
                              cloned::Dict{String,String})
    haskey(cloned, ref) && return cloned[ref]

    k = find_workbook_defined_name(wb, ref)
    isnothing(k) && return nothing
    dnv = wb.workbook_names[k]
    is_defined_name_value_a_reference(dnv.value) || return nothing
    dnv.value.sheet == old_sheet || return nothing

    new_name = "_xlchart.v$series.$(counter[])"
    counter[] += 1
    # Written straight into the store: addDefinedName refuses reserved
    # prefixes, and rightly so — this is the package generating a system
    # name, not a user creating one.
    wb.workbook_names[new_name] =
        DefinedNameValue(rename_sheet(dnv.value, new_sheet), dnv.isabs, dnv.hidden)
    cloned[ref] = new_name
    return new_name
end

# Add a relationship from part `from` to part `to`, reusing an existing one with
# the same type and target. Returns its Id.
function add_part_rel!(xf::XLSXFile, from::String, to::String, type::String)::String
    rels_path, root = get_or_create_rels!(xf, from)
    target = make_relative_target(first(_split_zip_path(from)), to)
    for n in elements_with_tag(root, "Relationship")
        get_attr(n, "Type") == type && get_attr(n, "Target") == target && return get_attr(n, "Id")
    end
    rid = new_relationship_id(root)
    push!(root, XML.Element(prefixed_tag(get_prefix(rels_path, xf), "Relationship");
                            Id = rid, Type = type, Target = target))
    return rid
end
