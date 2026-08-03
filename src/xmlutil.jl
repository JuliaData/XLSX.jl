# =============================================================================
# XML helpers
# =============================================================================

# ── elements ─────────────────────────────────────────────────────────────────
xml_elements(node) = XML.elements(node)

has_localname(n, tag::AbstractString) = localname(n) == tag

function first_element_with_tag(node::Union{Nothing,XML.Node}, tag::AbstractString)::Union{Nothing,XML.Node}
    isnothing(node) && return nothing
    for n in XML.eachelement(node)
        has_localname(n, tag) && return n
    end
    return nothing
end

elements_with_tag(node::XML.Node, tag::AbstractString) =
    XML.Node[n for n in XML.eachelement(node) if has_localname(n, tag)]

function xml_root_element(doc::XML.Node)::XML.Node
    for n in XML.eachelement(doc)
        return n
    end
    throw(XLSXError("Document has no root element"))
end

function xml_root_element(lz::XML.LazyNode)
    c = XML.Cursor(lz)
    while XML.next!(c) !== nothing
        XML.nodetype(c) == XML.Element && return XML.LazyNode(c)
    end
    throw(XLSXError("Document has no root element"))
end

# ── text ─────────────────────────────────────────────────────────────────────
function text_value(node::XML.Node)::Union{Nothing,String}
    v = XML.is_simple_value(node)
    isnothing(v) || return String(v)
    for c in XML.children(node)
        XML.nodetype(c) in (XML.Text, XML.CData) && return String(XML.value(c))
    end
    return nothing
end

child_text(node::Union{Nothing,XML.Node}, tag::AbstractString)::Union{Nothing,String} =
    (el = first_element_with_tag(node, tag); isnothing(el) ? nothing : text_value(el))

# ── attributes ───────────────────────────────────────────────────────────────
get_attr(node::XML.Node, key::AbstractString, default::AbstractString="") =
    get(node, key, default)

"""Value of a *namespace-prefixed* attribute matched by local name (`r:id`, `x:id`, …).
Unprefixed attributes are skipped deliberately: a bare `id` on `<c:chart>` or `embed`
on `<a:blip>` is a different attribute, not the relationship reference."""
function get_prefixed_attr(node::XML.Node, key::AbstractString)::Union{Nothing,String}
    atts = XML.attributes(node)
    isnothing(atts) && return nothing
    for (k, v) in atts
        occursin(':', k) && localname(k) == key && return v
    end
    return nothing
end

function child_val(node::Union{Nothing,XML.Node}, tag::AbstractString, default::Int)::Int
    el = first_element_with_tag(node, tag)
    isnothing(el) && return default
    v = get(el, "val", nothing)
    isnothing(v) && return default
    return something(tryparse(Int, v), default)
end