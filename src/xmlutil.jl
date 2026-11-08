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

"""
    _attr(node, key) -> Union{Nothing,String}

Read an attribute, normalising absence to `nothing`. `get_attr` returns `""`
for a missing attribute, which is fine where a default is being applied but
loses information wherever absent and explicitly-empty must stay distinct —
throughout DrawingML, where an absent attribute means "inherit".

Accepts `nothing` as the node, so an optional child element can be passed
straight through: `_attr(first_element_with_tag(el, "latin"), "typeface")`.
"""
_attr(node::XML.Node, key::AbstractString) = (v = get_attr(node, key); isempty(v) ? nothing : v)
_attr(::Nothing, ::AbstractString) = nothing

# DrawingML percentage attributes are in thousandths of a percent. Returns a
# fraction, not a percentage: 60000 -> 0.6. Used for color transforms
# (lumMod, satMod, shade, tint), text baseline, normAutofit scaling and line
# spacing — anything reading `_attr_pct` or `_attr_pct_opt` lands here.
@inline _pct(v::Int) = v / 100_000

"""Value of a *namespace-prefixed* attribute matched by local name (`r:id`, `x:id`, …).
Unprefixed attributes are skipped deliberately: a bare `id` on `<c:chart>` or `embed`
on `<a:blip>` is a different attribute, not the relationship reference."""
function get_prefixed_attr(node::XML.Node, key::AbstractString)::Union{Nothing,String}
    atts = XML.attributes(node)
    isnothing(atts) && return nothing
    for (k, v) in atts
        occursin(':', k) && has_localname(k, key) && return v
    end
    return nothing
end

"""
    ns_prefixes(root) -> Dict{String,String}

Map namespace URI to the prefix declared for it on `root`, from the `xmlns:*`
attributes of the chart part's `c:chartSpace`. Read from the file rather than
assumed, because a prefix is a per-document choice even though Excel always
writes `c` and `a`.
"""
function ns_prefixes(root::XML.Node)
    out = Dict{String,String}()
    for (k, v) in something(root.attributes, Pair{String,String}[])
        if k == "xmlns"
            out[v] = ""
        elseif startswith(k, "xmlns:")
            out[v] = k[7:end]
        end
    end
    return out
end

function child_val(node::Union{Nothing,XML.Node}, tag::AbstractString, default::Int)::Int
    el = first_element_with_tag(node, tag)
    isnothing(el) && return default
    v = get(el, "val", nothing)
    isnothing(v) && return default
    return something(tryparse(Int, v), default)
end

"""
    with_attribute(node, name, value) -> XML.Node

Return a node equal to `node` with attribute `name` set to `value`, replacing
any existing value in place so attribute order is preserved. A `nothing` value
removes the attribute.

`XML.Node` is immutable and `XML.attributes` returns a wrapper with no
`setindex!`, so this rebuilds rather than mutating — see [`insert_child`](@ref).
"""
function with_attribute(node::XML.Node, name::AbstractString, value)
    attrs = isnothing(node.attributes) ? Pair{String,String}[] : copy(node.attributes)
    i = findfirst(p -> first(p) == name, attrs)

    if isnothing(value)
        isnothing(i) && return node
        deleteat!(attrs, i)
    elseif isnothing(i)
        push!(attrs, String(name) => string(value))
    else
        attrs[i] = String(name) => string(value)
    end

    return typeof(node)(XML.nodetype(node), XML.tag(node),
                        isempty(attrs) ? nothing : attrs,
                        XML.value(node), node.children)
end

# ── nodes ──────────────────────────────────────────────────────────────────
function copynode(o::XML.Node)
    attrs = isnothing(o.attributes) ? nothing : copy(o.attributes)
    children = isnothing(o.children) ? nothing : XML.Node{String}[copynode(x) for x in o.children]
    return XML.Node{String}(o.nodetype, o.tag, attrs, o.value, children)
end

# -- XML tree helpers

# Remove all children with tag given by att[2] from a parent XML node with a tag given by att[1].
function unlink(node::XML.Node, att::Tuple{String,String}, pfx::String="")
    new_node = XML.Element(pfx * first(att))

    atts = XML.attributes(node)
    isnothing(atts) || foreach(((k, v),) -> new_node[k] = v, atts)

    for child in XML.children(node)
        localname(child) != last(att) && push!(new_node, child)
    end

    return new_node
end

function keep_prefix(node, pfx)
    pfx && return XML.tag(node)
    return localname(node)
end

# Find the index location of child and grandchild in a parent XML node.
function find_child_index(children, target)
    pfx = occursin(":", target)
    findfirst(c -> keep_prefix(c, pfx) == target, children)
end
function get_idces(doc::XML.Node, t, b)
    chn = XML.children(doc)
    i = find_child_index(chn, t)
    i === nothing && return nothing, nothing

    j = find_child_index(XML.children(chn[i]), b)
    return i, j
end
