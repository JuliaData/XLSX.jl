define_cf_address_methods(:setCfIconSet)
function setCfIconSet(ws::Worksheet, rng::CfRange; allkws::Dict{Symbol,Any}=Dict{Symbol,Any}())::Int
    kw = _iconset_keywords(allkws)

    _check_cf_range(ws, rng)

    new_pr = string(next_cf_priority!(ws))

    types = ("percentile", "percent", "num", "formula")
    _check_cfvo("min",  kw.min_type,  kw.min_val,  types)
    _check_cfvo("mid",  kw.mid_type,  kw.mid_val,  types)
    _check_cfvo("mid2", kw.mid2_type, kw.mid2_val, types)
    _check_cfvo("max",  kw.max_type,  kw.max_val,  types)

    # Each threshold as (type, value, gte), lowest first.
    lo = (kw.min_type,  _normalise_cfvo_val(ws, kw.min_val),  kw.min_gte)
    m1 = (kw.mid_type,  _normalise_cfvo_val(ws, kw.mid_val),  kw.mid_gte)
    m2 = (kw.mid2_type, _normalise_cfvo_val(ws, kw.mid2_val), kw.mid2_gte)
    hi = (kw.max_type,  _normalise_cfvo_val(ws, kw.max_val),  kw.max_gte)

    if !haskey(iconsets, kw.iconset)
        throw(XLSXError("Invalid iconset option chosen: $(kw.iconset). Valid options are: $(keys(iconsets))"))
    end
    cfx, thresholds, nicons = _iconset_template(kw.iconset, lo, m1, m2, hi)

    if kw.iconset in X14_ICONSETS
        cfx["id"] = _unique_guid(ws.package.uuid_rng, _x14_cf_ids(ws))
        cfx["priority"] = new_pr
        _set_x14_iconset!(cfx, kw, thresholds, nicons)
        update_worksheet_ext_cfx!(cfx, ws, rng)
    else
        cfx["priority"] = new_pr
        _set_iconset!(cfx, kw, thresholds)
        update_worksheet_cfx!(cfx, ws, rng)
    end
    return 0
end

# Icon sets Excel stores only in the x14 extension; the rest are 2007 rules.
const X14_ICONSETS = ("3Triangles", "3Stars", "5Boxes", "Custom")

# The keywords `setCfIconSet` accepts.
const ICONSET_KEYWORDS = (:iconset, :reverse, :showVal, :min_type, :min_val, :min_gte,
    :mid_type, :mid_val, :mid_gte, :mid2_type, :mid2_val, :mid2_gte,
    :max_type, :max_val, :max_gte, :icon_list)

# Read `setCfIconSet`'s keywords from `allkws`, rejecting any it doesn't accept.
function _iconset_keywords(allkws)
    for k in keys(allkws)
        k in ICONSET_KEYWORDS ||
            throw(XLSXError("Invalid keyword argument: $k. Valid options are: $(join(("`$v`" for v in ICONSET_KEYWORDS), ", "))."))
    end
    iconset::Union{Nothing,String} = get(allkws, :iconset, "3TrafficLights")
    iconset == "3TrafficLights1" && (iconset = "3TrafficLights") # the OOXML name for it
    reverse::Union{Nothing,String} = get(allkws, :reverse, nothing)
    showVal::Union{Nothing,String} = get(allkws, :showVal, nothing)
    min_type::Union{Nothing,String} = get(allkws, :min_type, nothing)
    min_val::Union{Nothing,String} = get(allkws, :min_val, nothing)
    min_gte::Union{Nothing,String} = get(allkws, :min_gte, nothing)
    mid_type::Union{Nothing,String} = get(allkws, :mid_type, nothing)
    mid_val::Union{Nothing,String} = get(allkws, :mid_val, nothing)
    mid_gte::Union{Nothing,String} = get(allkws, :mid_gte, nothing)
    mid2_type::Union{Nothing,String} = get(allkws, :mid2_type, nothing)
    mid2_val::Union{Nothing,String} = get(allkws, :mid2_val, nothing)
    mid2_gte::Union{Nothing,String} = get(allkws, :mid2_gte, nothing)
    max_type::Union{Nothing,String} = get(allkws, :max_type, nothing)
    max_val::Union{Nothing,String} = get(allkws, :max_val, nothing)
    max_gte::Union{Nothing,String} = get(allkws, :max_gte, nothing)
    icon_list::Union{Nothing,Vector{Int64}} = get(allkws, :icon_list, nothing)
    return (; iconset, reverse, showVal, min_type, min_val, min_gte, mid_type, mid_val, mid_gte,
              mid2_type, mid2_val, mid2_gte, max_type, max_val, max_gte, icon_list)
end

# The rule for `iconset`, copied from its preset, with the thresholds it uses and, for a
# custom set, the number of icons. A custom set takes as many thresholds as are given
# (two to four), and the preset's `<cfvo>`s are added to match.
function _iconset_template(iconset, lo, m1, m2, hi)
    cfx = copynode(iconsets[iconset])
    nicons = nothing
    l = first(iconset)
    if l == 'C'
        cfvo = XML.Element("x14:cfvo", type="percent")
        push!(cfvo, XML.Element("xm:f", XML.Text("dummy")))
        push!(cfx[1], copynode(cfvo)) # for min_val
        push!(cfx[1], copynode(cfvo)) # for max_val
        if isnothing(lo[1]) || isnothing(lo[2]) || isnothing(hi[1]) || isnothing(hi[2])
            throw(XLSXError("No type or val keywords defined. Must define at least `min_type`, `min_val`, `max_type` and `max_val` for a custom iconSet"))
        elseif isnothing(m1[1]) || isnothing(m1[2])
            thresholds = [lo, hi]
            nicons = 3
        elseif isnothing(m2[1]) || isnothing(m2[2])
            push!(cfx[1], copynode(cfvo)) # for mid_val
            cfx[1]["iconSet"] = "4Arrows"
            nicons = 4
            thresholds = [lo, m1, hi]
        else
            push!(cfx[1], copynode(cfvo)) # for mid_val
            push!(cfx[1], copynode(cfvo)) # for mid2_val
            cfx[1]["iconSet"] = "5Quarters"
            nicons = 5
            thresholds = [lo, m1, m2, hi]
        end
    elseif l == '5'
        thresholds = [lo, m1, m2, hi]
    elseif l == '4'
        thresholds = [lo, m1, hi]
    else
        thresholds = [lo, hi]
    end
    return cfx, thresholds, nicons
end

# Apply the options to an x14 icon set rule: thresholds are `<xm:f>` children, and a
# custom set lists its icons.
function _set_x14_iconset!(cfx, kw, thresholds, nicons)
    if !isnothing(kw.showVal) && kw.showVal == "false"
        cfx[1]["showValue"] = "0"
    end
    icon_list = kw.icon_list
    if !isnothing(kw.reverse) && kw.reverse == "true"
        if kw.iconset == "Custom"
            isnothing(icon_list) || (icon_list = reverse(icon_list))   # leaves the caller's vector alone
        else
            cfx[1]["reverse"] = "1"
        end
    end
    for (i, (type, val, gte)) in enumerate(thresholds)
        if !isnothing(type)
            cfx[1][i+1]["type"] = type # Need +1 because the first <cfvo> is always 0 percent.
        end
        if !isnothing(val)
            if !isnothing(type) && type == "formula"
                c = XML.Element("xm:f", XML.Text("(" * val * ")"))
            else
                c = XML.Element("xm:f", XML.Text(val))
            end
            if isempty(XML.children(cfx[1][i+1]))
                old = cfx[1][i+1]
                cfx[1][i+1] = XML.Node{String}(XML.Element, XML.tag(old), old.attributes, nothing, XML.Node{String}[c])
            else
                cfx[1][i+1][1] = c
            end
        end
        if !isnothing(gte) && gte == "false"
            cfx[1][i+1]["gte"] = "0"
        end
    end
    if kw.iconset == "Custom"
        if isnothing(icon_list)
            throw(XLSXError("No custom icons specified. Must specify between two and four icons."))
        elseif length(icon_list) < nicons
            throw(XLSXError("Too few custom icons specified: $(length(icon_list)). Expected $nicons"))
        end
        for (count, icon) in enumerate(string.(icon_list))
            if !isnothing(icon)
                if !haskey(allIcons, icon)
                    throw(XLSXError("Invalid custom icon specified: $icon. Valid values are \"1\" to \"52\"."))
                end
                i = allIcons[icon]
                push!(cfx[1], XML.Element("x14:cfIcon", iconSet=first(i), iconId=last(i)))
                count == nicons && break
            end
        end
    end
    return cfx
end

# Apply the options to a 2007 icon set rule: thresholds are `val` attributes.
function _set_iconset!(cfx, kw, thresholds)
    if !isnothing(kw.showVal) && kw.showVal == "false"
        cfx[1]["showValue"] = "0"
    end
    if !isnothing(kw.reverse) && kw.reverse == "true"
        cfx[1]["reverse"] = "1"
    end
    for (i, (type, val, gte)) in enumerate(thresholds)
        if !isnothing(val)
            if !isnothing(type) && type == "formula"
                cfx[1][i+1]["val"] = "(" * val * ")"
            else
                cfx[1][i+1]["val"] = val
            end
        end
        if !isnothing(type)
            cfx[1][i+1]["type"] = type
        end
        if !isnothing(gte) && gte == "false"
            cfx[1][i+1]["gte"] = "0"
        end
    end
    return cfx
end
