# Making a one-group chart a combo chart, as Excel's Change Chart Type → Combo →
# "Clustered Column – Line on Secondary Axis" does. Checked against Excel 365's
# own results: chart_combo_changes.xlsx (charts 1–3, from Clustered Column, Line
# and Stacked Column) and chart_combo_custom.xlsx (a customised column chart, and
# one with two series). What Excel does, and so what makeCombo does:
#
# - the last series by order becomes the line, on new secondary axes; the others
#   become a clustered column group on the chart's own axes, which keep their
#   formatting and titles;
# - a series that stays the same type is kept as it is; one that changes type
#   takes Excel's default formatting for its new group, coloured by its position,
#   and keeps its name, data and data labels;
# - the column group keeps a bar chart's gapWidth (150 from a line chart) and
#   loses its overlap and its stacking;
# - the title, legend and everything else outside the plot area, and the chart's
#   style and colour parts, are left as they are.

const _COMBO_SOURCES = (:column, :stackedColumn, :line)

# What a series keeps when it changes type: all but its formatting (spPr, marker,
# smooth, invertIfNegative), its per-point formatting and the bar-only settings.
const _RETYPE_KEEP = ("idx", "order", "tx", "dLbls", "trendline", "errBars", "cat", "val", "extLst")

# The data label positions each group allows (ECMA-376 Part 1, 21.2.2.48).
const _DLBL_POSITIONS = Dict("barChart"  => ("ctr", "inEnd", "inBase", "outEnd"),
                             "lineChart" => ("t", "b", "l", "r", "ctr"))

# Data labels with any position the new group doesn't allow removed, so Excel
# places them by default rather than finding the file invalid.
_fit_label_positions(dlbls::XML.Node, gtag::AbstractString) =
    _rewrite(dlbls) do e
        localname(e) in ("dLbls", "dLbl") || return e
        _with_children(e, XML.Node[k for k in XML.eachelement(e)
                                   if !(localname(k) == "dLblPos" &&
                                        !(_attr(k, "val") in _DLBL_POSITIONS[gtag]))])
    end

# A series moved to group `to` of the combo template (1 the columns, 2 the line),
# formatted as a series of that group in that position. A line series takes at
# most one set of error bars.
function _retype_series(ser::XML.Node, to::Integer, pfx)
    gtag = to == 1 ? "barChart" : "lineChart"
    key  = (NS_C, to == 1 ? "barSer" : "lineSer")
    kids = XML.Node[]
    for k in XML.eachelement(ser)
        t = localname(k)
        t in _RETYPE_KEEP || continue
        t == "errBars" && to == 2 && any(x -> localname(x) == "errBars", kids) && continue
        push!(kids, t == "dLbls" ? _fit_label_positions(k, gtag) : k)
    end
    out = _with_children(ser, kids)
    pat, from = _series_pattern(:combo, to)
    col = _accent(something(_int_val(ser, "idx"), 0) + 1)
    for k in pat
        out = insert_child(out, key, _with_color(_in_prefixes(k, pfx), col, pfx; from))
    end
    return out
end

_axis_ids(n::XML.Node) = [_attr(a, "val") for a in elements_with_tag(n, "axId")]

"""
    makeCombo(c::Chart) -> Chart

Make `c` a combo chart, as Excel's Change Chart Type → Combo → "Clustered Column
– Line on Secondary Axis" does. The last series becomes a line on new secondary
axes, and the others clustered columns on the chart's own axes. Afterwards `c`
is a `:combo` chart, and [`addSeries`](@ref) takes `group = :column` or
`group = :line` to add to it.

`c` must be a `:column`, `:stackedColumn` or `:line` chart with at least two
series. As in Excel, a series that changes type takes the default formatting of
its new group, keeping its name, data and data labels; the others keep theirs.
The axes, title, legend and the rest of the chart are kept, and a stacked
chart's columns become clustered.
"""
function makeCombo(c::Chart)::Chart
    groups = _group_nodes(chart_root(c))
    length(groups) == 1 || throw(XLSXError(isempty(groups) ? "Chart `$(c.name)` has no plot group." :
        "Chart `$(c.name)` already has $(length(groups)) plot groups; " *
        "`makeCombo` makes a combo chart from a chart with one."))
    kind = _chart_kind(c)
    kind in _COMBO_SOURCES || throw(XLSXError(
        "`makeCombo` makes a combo chart from a " * join((":$k" for k in _COMBO_SOURCES), ", ", " or ") *
        " chart, not a :$kind one."))
    root = chart_root(c)
    pfx  = ns_prefixes(root)
    g    = only(groups)
    sers = sort(collect(elements_with_tag(g, "ser")); by = s -> something(_int_val(s, "order"), 0))
    length(sers) >= 2 || throw(XLSXError(
        "Chart `$(c.name)` has $(length(sers)) series, and a combo chart is made from two or more, " *
        "the last becoming the line. Add series first, or create one with `addChart(ws, :combo; …)`."))
    frombar = localname(g) == "barChart"
    cols = frombar ? sers[1:end-1] : [_retype_series(s, 1, pfx) for s in sers[1:end-1]]
    line = frombar ? _retype_series(sers[end], 2, pfx) : sers[end]

    # The combo template's groups and secondary axes, renumbered after the chart's own.
    troot = xml_root_element(parse(CHART_KIND_TEMPLATES[:combo], XML.Node))
    tbar, tline = _group_nodes(troot)
    taxes = [k for k in _plotarea_children(troot) if localname(k) in ("catAx", "valAx")]
    used  = [parse(Int, v) for n in _plotarea_children(root) for v in _axis_ids(n)]
    top   = maximum(used; init = 0)
    newid = Dict(v => string(top + i) for (i, v) in enumerate(_axis_ids(tline)))
    renum(n) = _rewrite(n) do e
        localname(e) in ("axId", "crossAx") && haskey(newid, _attr(e, "val")) || return e
        with_attribute(e, "val", newid[_attr(e, "val")])
    end
    secondary = [_in_prefixes(renum(a), pfx) for a in taxes
                 if _attr(first_element_with_tag(a, "axId"), "val") in keys(newid)]

    without(n, tags) = _with_children(n, XML.Node[k for k in XML.eachelement(n) if !(localname(k) in tags)])
    gap = frombar ? something(first_element_with_tag(g, "gapWidth"),
                              _c_el(pfx, "gapWidth"; val = "150")) : _c_el(pfx, "gapWidth"; val = "150")
    bar = without(_in_prefixes(tbar, pfx), ("ser", "overlap", "gapWidth", "axId"))
    for k in (cols..., gap, (_c_el(pfx, "axId"; val = v) for v in _axis_ids(g))...)
        bar = insert_child(bar, (NS_C, "barChart"), k)
    end
    ln = insert_child(without(_in_prefixes(renum(tline), pfx), ("ser",)), (NS_C, "lineChart"), line)

    function newplot(pa)
        kids = XML.Node[]
        for k in XML.eachelement(pa)
            k === g ? append!(kids, (bar, ln)) : push!(kids, k)
        end
        lastax = findlast(k -> localname(k) in ("catAx", "valAx", "dateAx", "serAx"), kids)
        for (i, a) in enumerate(secondary)
            insert!(kids, lastax + i, a)
        end
        return _with_children(pa, kids)
    end
    set_chart_root!(c, rebuild_path(root, [(NS_C, "chart") => "chart", (NS_C, "plotArea") => "plotArea"],
                                    newplot; prefixes = pfx))
    return c
end
