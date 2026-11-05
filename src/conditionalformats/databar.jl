define_cf_address_methods(:setCfDataBar)
function setCfDataBar(ws::Worksheet, rng::CfRange; allkws::Dict{Symbol,Any}=Dict{Symbol,Any}())::Int
    kw = _databar_keywords(allkws)

    _check_cf_range(ws, rng)

    pfx = _cf_prefix(ws)

    # `priority` is for internal callers only and is deliberately undocumented: it is
    # omitted from the valid-keyword list so a user passing it gets an error.
    # `setColoredDataBars` reserves a contiguous block of priorities up front and passes
    # them in band order, so the bands are numbered predictably rather than in whatever
    # order the partition happens to yield. Note it bypasses `next_cf_priority!`, so a
    # caller supplying it owns the number: no uniqueness check, and the counter does not
    # advance.
    new_pr = string(something(kw.priority, next_cf_priority!(ws)))

    settings = _databar_settings(ws, kw)

    # The 2007 rule and its x14 extension are linked by this id.
    id = _unique_guid(ws.package.uuid_rng, _x14_cf_ids(ws))

    update_worksheet_cfx!(_databar_rule(pfx, settings, new_pr, id), ws, rng)  # Add basic elements to worksheet xml file
    update_worksheet_ext_cfx!(_databar_ext_rule(settings, id), ws, rng)     # Add extension elements to worksheet xml file
    return 0
end

# The keywords `setCfDataBar` accepts (`priority` aside: see `setCfDataBar`).
const DATABAR_KEYWORDS = (:databar, :showVal, :gradient, :borders, :sameNegFill, :sameNegBorders,
    :direction, :axis_pos, :axis_col, :min_type, :min_val, :max_type, :max_val,
    :fill_col, :border_col, :neg_fill_col, :neg_border_col)

# Read `setCfDataBar`'s keywords from `allkws`, rejecting any it doesn't accept.
function _databar_keywords(allkws)
    for k in keys(allkws)
        k in DATABAR_KEYWORDS || k == :priority ||
            throw(XLSXError("Invalid keyword argument: $k. Valid options are: $(join(("`$v`" for v in DATABAR_KEYWORDS), ", "))."))
    end
    databar::Union{Nothing,String} = get(allkws, :databar, "bluegrad")
    priority::Union{Nothing,Int} = get(allkws, :priority, nothing)
    showVal::Union{Nothing,String} = get(allkws, :showVal, nothing)
    gradient::Union{Nothing,String} = get(allkws, :gradient, nothing)
    borders::Union{Nothing,String} = get(allkws, :borders, nothing)
    sameNegFill::Union{Nothing,String} = get(allkws, :sameNegFill, nothing)
    sameNegBorders::Union{Nothing,String} = get(allkws, :sameNegBorders, nothing)
    direction::Union{Nothing,String} = get(allkws, :direction, nothing)
    axis_pos::Union{Nothing,String} = get(allkws, :axis_pos, nothing)
    axis_col::Union{Nothing,CellColor} = get(allkws, :axis_col, nothing)
    min_type::Union{Nothing,String} = get(allkws, :min_type, nothing)
    min_val::Union{Nothing,String} = get(allkws, :min_val, nothing)
    max_type::Union{Nothing,String} = get(allkws, :max_type, nothing)
    max_val::Union{Nothing,String} = get(allkws, :max_val, nothing)
    fill_col::Union{Nothing,CellColor} = get(allkws, :fill_col, nothing)
    border_col::Union{Nothing,CellColor} = get(allkws, :border_col, nothing)
    neg_fill_col::Union{Nothing,CellColor} = get(allkws, :neg_fill_col, nothing)
    neg_border_col::Union{Nothing,CellColor} = get(allkws, :neg_border_col, nothing)
    return (; databar, priority, showVal, gradient, borders, sameNegFill, sameNegBorders,
              direction, axis_pos, axis_col, min_type, min_val, max_type, max_val,
              fill_col, border_col, neg_fill_col, neg_border_col)
end

# Validate the keywords and fill any not given from the `databar` preset, returning the
# settings the rule is built from, keyed by name.
function _databar_settings(ws::Worksheet, kw)
    (; databar, min_type, min_val, max_type, max_val) = kw

    min_type in ("least", "automatic") && (min_val = nothing)
    max_type in ("highest", "automatic") && (max_val = nothing)
    _check_cfvo("min", min_type, min_val, ("least", "percentile", "percent", "num", "formula", "automatic"))
    _check_cfvo("max", max_type, max_val, ("highest", "percentile", "percent", "num", "formula", "automatic"))

    min_val = _normalise_cfvo_val(ws, min_val)
    max_val = _normalise_cfvo_val(ws, max_val)

    if !haskey(databars, databar)
        throw(XLSXError("Invalid dataBar option chosen: $databar. Valid options are: $(keys(databars))"))
    end

    settings = Dict{String,Any}(
        "showVal" => kw.showVal,
        "gradient" => kw.gradient,
        "borders" => kw.borders,
        "sameNegFill" => kw.sameNegFill,
        "sameNegBorders" => kw.sameNegBorders,
        "direction" => kw.direction,
        "min_type" => min_type,
        "min_val" => min_val,
        "max_type" => max_type,
        "max_val" => max_val,
        "fill_col" => kw.fill_col,
        "border_col" => kw.border_col,
        "neg_fill_col" => kw.neg_fill_col,
        "neg_border_col" => kw.neg_border_col,
        "axis_pos" => kw.axis_pos,
        "axis_col" => kw.axis_col
    )

    for (k, w) in settings # Allow user input to override any default value
        if isnothing(w)
            if haskey(databars[databar], k)
                settings[k] = databars[databar][k]
            end
        end
    end
    for k in ["showVal", "gradient", "borders", "sameNegFill", "sameNegBorders"]
        haskey(settings, k) && isValidKw(k, settings[k], ["true", "false"])
    end
    haskey(settings, "direction") && isValidKw("direction", settings["direction"], ["leftToRight", "rightToLeft"])
    haskey(settings, "axis_pos") && isValidKw("axis_pos", settings["axis_pos"], ["middle", "none"])
    return settings
end

# The 2007 `<cfRule>`: the bar's range and fill color, and an `<extLst>` linking it by
# `id` to the x14 rule that carries everything else.
function _databar_rule(pfx::String, s::Dict{String,Any}, priority::String, id::AbstractString)
    mnt = s["min_type"] ∈ ["automatic", "least"] ? "min" : s["min_type"]
    mxt = s["max_type"] ∈ ["automatic", "highest"] ? "max" : s["max_type"]

    cfx = XML.Element("$(pfx)cfRule"; type="dataBar", priority=priority)
    cdb = XML.Element("$(pfx)dataBar")
    push!(cdb, isnothing(s["min_val"]) ?
            XML.Element("$(pfx)cfvo"; type=mnt) :
            XML.Element("$(pfx)cfvo"; type=mnt, val=s["min_val"])
    )

    push!(cdb, isnothing(s["max_val"]) ?
            XML.Element("$(pfx)cfvo"; type=mxt) :
            XML.Element("$(pfx)cfvo"; type=mxt, val=s["max_val"])
    )

    push!(cdb, _color_element("$(pfx)color", s["fill_col"]))

    push!(cfx, cdb)

    if haskey(s, "showVal") && !isnothing(s["showVal"]) && s["showVal"] == "false"
        cdb["showValue"] = "0"
    end
    # Links the 2007 rule to its x14 counterpart by id. Excel expects the <ext> in an
    # <extLst> inside <cfRule>, as a sibling following <dataBar> - not inside <dataBar>
    # itself, where it is silently discarded on open along with the x14 rule it links to.
    cfx_ext = XML.Element("$(pfx)ext")
    cfx_ext["uri"] = "{B025F937-C7B1-47D3-B67F-A62EFF666E3E}"
    cfx_ext["xmlns:x14"] = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"
    push!(cfx_ext, XML.Element("x14:id", XML.Text(id)))
    cfx_extlst = XML.Element("$(pfx)extLst")
    push!(cfx_extlst, cfx_ext)
    push!(cfx, cfx_extlst)
    return cfx
end

# The x14 `<cfRule>`: gradient, borders, direction, negative bars and axis.
function _databar_ext_rule(s::Dict{String,Any}, id::AbstractString)
    emnt = s["min_type"] == "automatic" ? "autoMin" : s["min_type"] == "least" ? "min" : s["min_type"]
    emxt = s["max_type"] == "automatic" ? "autoMax" : s["max_type"] == "highest" ? "max" : s["max_type"]
    emnv = s["min_type"] == "formula" ? "(" * s["min_val"] * ")" : s["min_val"]
    emxv = s["max_type"] == "formula" ? "(" * s["max_val"] * ")" : s["max_val"]
    ext_cfx = XML.Element("x14:cfRule", type="dataBar", id=id)
    ext_db = XML.Element("x14:dataBar", minLength="0", maxLength="100")
    valmin = XML.Element("x14:cfvo", type=emnt)
    !isnothing(emnv) && push!(valmin, XML.Element("xm:f", XML.Text(emnv)))
    valmax = XML.Element("x14:cfvo", type=emxt)
    !isnothing(emxv) && push!(valmax, XML.Element("xm:f", XML.Text(emxv)))
    push!(ext_db, valmin)
    push!(ext_db, valmax)
    if s["gradient"] == "false"
        ext_db["gradient"] = "0"
    end
    do_borders = haskey(s, "borders") && s["borders"] == "true"
    if do_borders
        ext_db["border"] = "1"
        push!(ext_db, _color_element("x14:borderColor", something(get(s, "border_col", nothing), "FF638EC6")))
    end
    if haskey(s, "direction")
        if s["direction"] == "leftToRight"
            ext_db["direction"] = "leftToRight"
        elseif s["direction"] == "rightToLeft"
            ext_db["direction"] = "rightToLeft"
        end
    end
    if haskey(s, "sameNegFill") && s["sameNegFill"] == "true"
        ext_db["negativeBarColorSameAsPositive"] = "1"
    else
        push!(ext_db, _color_element("x14:negativeFillColor", something(get(s, "neg_fill_col", nothing), "FFFF0000")))
    end
    if do_borders && haskey(s, "sameNegBorders") && s["sameNegBorders"] == "false"
        ext_db["negativeBarBorderColorSameAsPositive"] = "0"
        push!(ext_db, _color_element("x14:negativeBorderColor", something(get(s, "neg_border_col", nothing), "FFFF0000")))
    end
    if haskey(s, "axis_pos")
        if s["axis_pos"] == "none"
            ext_db["axisPosition"] = "none"
        elseif s["axis_pos"] == "middle"
            ext_db["axisPosition"] = "middle"
        end
    end
    haskey(s, "axis_col") && push!(ext_db, _color_element("x14:axisColor", s["axis_col"]))
    push!(ext_cfx, ext_db)
    return ext_cfx
end



# ---------------------------------------------------------------------------
# setColoredDataBars
# ---------------------------------------------------------------------------

"""
    setColoredDataBars(sheet, rng; [bands], [colors], [breaks], [gte], [showVal],
                       [direction], [min_type], [min_val], [max_type], [max_val],
                       [clear]) -> Vector{Pair}

Apply data bars to `rng` whose color changes with value as well as their length.

Excel has no such conditional format: a `dataBar` rule carries a single color and
applies to every cell in its range. `setColoredDataBars` splits `rng` into value-based
bands and writes one data bar rule per band, each over the non-contiguous range of
cells falling in that band. The result is a valid workbook that Excel reads without
complaint, but it is not something Excel's own interface can produce, and it is not
maintained by Excel: see the note on editing below.

`bands` (default 5) sets how many bands to use, with boundaries placed at equal
intervals between the smallest and largest value in `rng`. Give `breaks` instead to
place the boundaries yourself, as a sorted vector of `length(bands)-1` values; `breaks`
overrides `bands`. By default a cell whose value equals a break falls in the band above
it; set `gte=false` to place it in the band below, or give a vector of `Bool` with one
entry per break.

`colors` accepts either one color per band, or exactly two colors to interpolate
between (in LCHab, so intermediate hues stay clean). Colors may be given in any form
described in [Specifying colors](@ref), but theme colors can only be used one per band,
not interpolated between. Note that Colors.jl's named colormaps are not accepted; for a
red-amber-green ramp give the three colors explicitly.

`showVal` and `direction` are passed through to each band's rule unchanged. Bars are
drawn with a solid fill rather than a gradient.

`min_type`, `min_val`, `max_type` and `max_val` set the value axis shared by every band,
and default to `num` with the minimum and maximum of `rng`. Only `num` and `formula` are
accepted: the relative types (`percent`, `percentile`, `min`, `max`, `automatic`) are
evaluated over each rule's own cells, which for a band is only that band's values, so
every band would span the full bar width and bar lengths would carry no meaning.

By default (clear=true) any conditional format lying entirely within rng is removed
before the new bands are written, so repeated calls replace rather than accumulate. This
matters when re-banding after the data have changed: bands from an earlier call name the
cells they covered then, and left in place they would overlay the new ones. Set
clear=false to add the bands alongside whatever is already there.

Returns the band label => range pairs written. Bands containing no cells are dropped, so
the result may be shorter than bands. Labels keep their position on the scale rather
than being renumbered: banding data with only two distinct values across bands=5 may
return labels 1 and 5, using the first and last colors of the ramp.

!!! note
    `setColoredDataBars` is a static formatting function. The bands reflect cell values 
    at the time of writing. If a value is later edited in Excel its bar length updates, 
    but its color does not, because the cell stays in the band it was originally 
    assigned to. Rewrite the conditional formatting to re-band it.

    This is a limitation of Excel itself, which does not support dataBars whose color 
    varies with value. `setColoredDataBars` is a workaround for this.


!!! note "Experimental"

    This function is experimental. It has a full set of tests and Documentation is 
    published but it is not yet marked as public. It may be better included in 
    future only as a documented example rather than as a package function.

# Examples

```julia
julia> XLSX.setColoredDataBars(sheet, "B2:B101")

julia> XLSX.setColoredDataBars(sheet, "B2:B101"; bands=3, colors=["green", "orange", "red"])

julia> XLSX.setColoredDataBars(sheet, "B2:B101"; breaks=[10, 50], colors=["green", "red"])

julia> XLSX.setColoredDataBars(sheet, "B2:B101"; bands=8, min_val="0", max_val="100")
```

See also [`setConditionalFormat`](@ref), [`clearConditionalFormats`](@ref).
"""
function setColoredDataBars end

setColoredDataBars(ws::Worksheet, rng::AbstractString; kw...) =
    setColoredDataBars(ws, CellRange(rng); kw...)
setColoredDataBars(ws::Worksheet, cell::CellRef; kw...) =
    setColoredDataBars(ws, CellRange(cell, cell); kw...)
setColoredDataBars(ws::Worksheet, rng::SheetCellRange; kw...) =
    do_sheet_names_match(ws, rng) && setColoredDataBars(ws, rng.rng; kw...)

function setColoredDataBars(ws::Worksheet, rng::CellRange;
                            bands::Int=5,
                            colors=["green", "red"],
                            breaks=nothing,
                            gte::Union{Bool,AbstractVector{Bool}}=true,
                            showVal::Union{Nothing,String}=nothing,
                            direction::Union{Nothing,String}=nothing,
                            min_type::Union{Nothing,String}=nothing,
                            min_val::Union{Nothing,String}=nothing,
                            max_type::Union{Nothing,String}=nothing,
                            max_val::Union{Nothing,String}=nothing,
                            clear::Bool=true)

    _check_cf_areas(ws, rng)
    
    clear && clearConditionalFormats(ws, rng)

    vals = filter(v -> v isa Real && !isnan(v), vec(getdata(ws, rng)))
    isempty(vals) && throw(XLSXError("No numeric values in `$rng` to band."))
    lo, hi = float(minimum(vals)), float(maximum(vals))

    brk = if isnothing(breaks)
        bands < 1 && throw(XLSXError("`bands` must be at least 1, got $bands."))
        bands == 1 && return _single_bar(ws, rng, colors, lo, hi,
                                         showVal, direction, min_type, min_val, max_type, max_val)
        hi > lo || throw(XLSXError("All numeric values in `$rng` are equal; cannot form bands."))
        collect(range(lo, hi; length=bands + 1))[2:end-1]
    else
        bands = length(breaks) + 1
        collect(breaks)
    end

    cols = _band_colors(colors, bands)

    for (t, name) in ((min_type, "min_type"), (max_type, "max_type"))
        isnothing(t) || t in ["num", "formula"] || throw(XLSXError(
            "`$name` must be `num` or `formula` for colored data bars; `$t` is evaluated " *
            "separately for each band, which makes bar lengths incomparable between bands."))
    end
    mnt = something(min_type, "num")
    mxt = something(max_type, "num")
    mnv = something(min_val, string(lo))
    mxv = something(max_val, string(hi))

    p = partition(ws, rng, brk; labels=1:bands, gte=gte)

    base = next_cf_priority!(ws)
    for _ in 2:length(p)
        next_cf_priority!(ws)          # reserve the rest
    end

    for (n, (i, ncr)) in enumerate(p)
        kws = Dict{Symbol,Any}(
            :priority => base + n - 1,
            :fill_col => cols[i],
            :gradient => "false",
            :borders => "false",
            :sameNegFill => "true",
            :min_type => mnt, :min_val => mnv,
            :max_type => mxt, :max_val => mxv,
        )
        isnothing(showVal) || (kws[:showVal] = showVal)
        isnothing(direction) || (kws[:direction] = direction)
        setCfDataBar(ws, ncr; allkws=kws)
    end
    return p
end

# `bands=1` degenerates to an ordinary data bar over the whole contiguous range:
# no partition, no non-contiguous sqref.
function _single_bar(ws, rng, colors, lo, hi, showVal, direction,
                     min_type, min_val, max_type, max_val)
    kws = Dict{Symbol,Any}(
        :fill_col => only(_band_colors(colors, 1)),
        :gradient => "false",
        :borders => "false",
        :sameNegFill => "true",
        :min_type => something(min_type, "num"), :min_val => something(min_val, string(lo)),
        :max_type => something(max_type, "num"), :max_val => something(max_val, string(hi)),
    )
    isnothing(showVal) || (kws[:showVal] = showVal)
    isnothing(direction) || (kws[:direction] = direction)
    setCfDataBar(ws, rng; allkws=kws)
    return [1 => rng]
end
