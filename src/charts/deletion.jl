# Deleting a chart from a worksheet.

"""
    deleteChart(c::AbstractChart)

Delete chart `c`, a [`Chart`](@ref) or [`ChartEx`](@ref) from [`getCharts`](@ref)
or [`getChart`](@ref), from the worksheet it is on: its frame in the sheet's
drawing, and the chart part with everything only it uses (its style and colors).
A chartEx chart's hidden `_xlchart` names go too. The data it plotted is not
changed.

A chart on a chartsheet is the sheet: delete the sheet with
[`XLSX.deletesheet!`](@ref) instead.

The file must be open for writing (`mode="rw"`).

# Examples
```julia
julia> c = getChart(xf, "Sales by month");

julia> deleteChart(c)
```

See also [`addChart`](@ref) and [`getCharts`](@ref).
"""
function deleteChart(c::AbstractChart)
    xf = c.package
    XLSX._require_writable(xf, "delete a chart")
    (isnothing(c.sheet) || isnothing(c.rId)) &&
        throw(XLSXError("Chart `$(c.name)` isn't placed on a sheet, so there is nothing to delete it from."))
    wb = get_workbook(xf)
    is_chartsheet(wb, c.sheet) &&
        throw(XLSXError("Chart `$(c.name)` is the chartsheet `$(c.sheet)`. Delete the sheet with `deletesheet!` instead."))
    sheet_path = get_worksheet_internal_file(getsheet(wb, c.sheet))
    drawing_path = _drawing_path_for_sheet(xf, sheet_path)
    isnothing(drawing_path) && throw(XLSXError("Chart `$(c.name)` is no longer on sheet `$(c.sheet)`."))

    # The chart's frame: the drawing's top-level element that refers to it (an
    # anchor, or for a chartEx chart the mc:AlternateContent around one).
    root = xml_root_element(get_xml_data(xf, drawing_path))
    i = findfirst(n -> XML.nodetype(n) == XML.Element && XLSX._refers_to_rel(n, c.rId), XML.children(root))
    isnothing(i) && throw(XLSXError("Chart `$(c.name)` is no longer on sheet `$(c.sheet)`."))
    deleteat!(XML.children(root), i)

    # A chartEx chart finds its data through hidden `_xlchart` names, which only it uses.
    if c isa ChartEx && haskey(xf.data, c.path)
        names = String[]
        XLSX._rewrite_chart_formulas!(f -> (startswith(f, "_xlchart.") && push!(names, String(f)); f),
                                      xml_root_element(get_xml_data(xf, c.path)))
        foreach(n -> delete!(wb.workbook_names, n), names)
    end

    XLSX._drop_drawing_rel!(xf, drawing_path, c.rId)
    XLSX._remove_drawing_if_empty!(xf, sheet_path, drawing_path)
    return nothing
end
