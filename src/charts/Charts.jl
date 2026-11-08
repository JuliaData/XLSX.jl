"""
    XLSX.Charts

Reading, formatting and creating Excel charts, in both chart schemas: the `c:` schema
([`Chart`](@ref XLSX.Charts.Chart)) and the `cx:` chartEx schema
([`ChartEx`](@ref XLSX.Charts.ChartEx)).

`using XLSX.Charts` brings the chart functions into scope. The handle and value types are
public but not exported; refer to them as `XLSX.Charts.ChartSeries` and so on.

Conventions shared by the functions:

- `i` is a series' position, counting from 1 in [`getChartSeries`](@ref) order, and
  `point` a data point's position in the series, also from 1.
- `ax` is a `ChartAxis` from [`getChartAxes`](@ref) or [`getChartAxis`](@ref), and `g` a
  `ChartGroup` from [`getChartGroups`](@ref). On a `ChartEx`, an axis is named by its `id`
  from [`getChartAxisIds`](@ref).
- Setters return the chart. Most take `:inherit` to remove a setting so it is inherited
  from the chart style; each docstring says where it does.
- Color values are listed at [`setSeriesFill`](@ref), line keywords at
  [`setSeriesLine`](@ref), and text-property fields at [`setChartTitleTextProp`](@ref).
"""
module Charts

# ---------------------------------------------------------------------------
# Exports
#
# Export what users write in their own code: every function they call, and the
# types they construct as arguments (`SchemeColor`). Types users only receive —
# handles (`Chart`, `ChartEx`), values (`ChartSeries`, `ChartAxis`, …), the
# DrawingML structs, `Effective` — are `public` but not exported.
#
# Naming
#
# Chart property accessors follow `get<Subject><Property>`, and their setters
# `set<Subject><Property>` — `getSeriesFill`/`setSeriesFill`,
# `getSeriesLine`/`setSeriesLine`. The subject is the thing the property
# belongs to, not the argument used to reach it: `getMarkerFill(c, i, point)`
# takes a series index but describes the marker.
#
# Positions count from 1 throughout — series in document order, data points as
# the user sees them — and are resolved afresh on each call. Excel's own
# identifiers (`c:idx`, `c:axId`) are the keys the value types carry, and need
# not agree with a position.
#
# A data point is a keyword, `point`, where it is optional, because the series
# form means something on its own: `setMarkerSize(c, i, size; point)`,
# `getSeriesFill(c, i; point)`. It is positional where it is required, because
# the function means nothing without it: `setLabelDeleted(c, i, point, deleted)`,
# `getMarkerFill(c, i, point)`.
#
# `get…TextProp` (singular) resolves one run property up the cascade and returns
# an `Effective`; `get…TextProps` (plural) returns the text-properties struct as
# written at that site.
#
# The DrawingML layer keeps snake_case — `parse_drawing_fill`, `text_content`, 
# `resolve_color_base` — because those are parsers and
# predicates rather than accessors, and are internal.
#
# The predicates users call, `hasFill` and `hasLine`, are camelCase like the rest of
# the exported API.
#
# Two senses of "resolve" are kept apart. `resolve_color_base` and
# `apply_drawingml_transforms` resolve a color to RGB. The `get<Subject>`
# accessors that return an `Effective` resolve a property up the inheritance
# cascade. A cascade-resolved fill may still hold a scheme color, which is then
# a separate step.
# ---------------------------------------------------------------------------

import ..XLSX
import ..XLSX: iserror, geterror      # core generics; discovery.jl adds methods
import Colors
import Dates
import XML

using ..XLSX:
    # types
    XLSXFile, Workbook, Worksheet, XLSXError, DataTable,
    CellRef, CellRange, NonContiguousRange,
    SheetCellRef, SheetCellRange, SheetColumnRange, SheetRowRange,
    XL_REF, CellErrorType, RichTextString, RichTextRun,
    # package parts and relationships
    REL_CHART, REL_CHARTEX, REL_CHARTSHEET, MIME_CHART, MIME_CHARTEX, MIME_CHARTSHEET,
    parts_with_content_type, register_content_type!, add_part_rel!,
    get_relationship_target_by_id, resolve_relative_target,
    _next_part_name, _split_zip_path, _register_sheet!, _next_xlchart_series,
    # drawings
    ensure_drawing!, _drawing_path_for_sheet, build_two_cell_anchor, _next_shape_id,
    # XML helpers
    NS_A, NS_R, NS_XDR, _attr, _pct, child_text, child_val, elements_with_tag,
    first_element_with_tag, get_attr, get_prefixed_attr, has_localname, localname,
    ns_prefixes, prefixed_tag, with_attribute, xml_root_element, truncate_len,
    # workbook, sheets and cells
    get_workbook, get_xlsxfile, get_xml_data, get_worksheet_internal_file,
    getsheet, getdata, getRichTextString, is_chartsheet, isdate1904, date_to_excel_value, time_to_excel_value,
    # references and defined names
    abscell, mkabs, quoteit, ref_chooser, row_number, column_number, _parse_cell_marker,
    get_defined_name_value, is_defined_name_value_a_reference, is_workbook_defined_name,
    get_ext_refs, get_external_workbook_path,
    is_valid_cellname, is_valid_cellrange,
    is_valid_column_range, is_valid_row_range,
    is_valid_sheet_cellname, is_valid_sheet_cellrange,
    is_valid_sheet_column_range, is_valid_sheet_row_range,
    is_valid_fixed_sheet_cellname, is_valid_fixed_sheet_cellrange,
    is_valid_fixed_sheet_column_range, is_valid_fixed_sheet_row_range,
    # cell errors
    ERROR_STRING_TO_CODE, get_error_string,
    # theme and colors (xlsx-colors.jl)
    get_color, get_colorant, _linear_to_srgb,
    SchemeColor, SCHEME_TOKENS, COLOR_TRANSFORMS, SCHEME_ALIASES, canonical_token,
    get_theme_color_map, get_theme_fonts, resolve_theme_font, apply_drawingml_transforms

export
    # Finding and reading charts
    getCharts, getChart, getChartData, getChartRanges, getChartTitle,
    getChartType, getChartSchema,
    # Creating charts
    addChart, addChartEx, addSeries, deleteChart,
    # Chart level
    getChartSeries, getChartTypes, getChartGroups, getChartAxes, getChartAxis,
    getChartSeriesCount, getChartDataBlocks, getChartAxisIds,
    getChartTitleText, getChartTitleRef, getChartTitleRange,
    getChartTitleShapeProps, getChartTitleTextProps, getAutoTitleDeleted,
    getPlotAreaShapeProps, getChartSpaceShapeProps, getChartSpaceTextProps,
    getChartLegend, getLegendPos, getLegendOverlay, getLegendShapeProps, getLegendTextProps,
    setChartTitleText, setChartTitleTextProp, setChartSpaceTextProp, setLegendTextProp,
    setLegendEntryDeleted, getLegendEntryDeleted, setLegendPos, setLegendOverlay,
    setPlotAreaLine,
    # Series
    getSeriesName, getSeriesNameRange, getSeriesData, getSeriesHidden,
    getSeriesGroup, getSeriesAxes, getSeriesAxisIds, getSeriesLayout,
    getSeriesShapeProps, getSeriesFill, getSeriesLine, getSeriesMarker,
    getSeriesLabelTextProps, getSeriesDataPoints, getSeriesDataPoint,
    getSeriesDataLabels, getSeriesDataLabel, getSeriesTrendlines, getSeriesErrorBars,
    getSeriesSubtotals, getSeriesBinning, getSeriesAggregation, getSeriesQuartileMethod,
    getSeriesParentLabelLayout, getSeriesLayoutFlag,
    setSeriesName, setSeriesFill, setSeriesLine, setSeriesLineColor, setSeriesLineWidth,
    setSeriesLineDash, setSeriesLineCap, setSeriesLineCompound, setSeriesLineJoin,
    setSeriesLineMiterLimit,
    setSeriesSubtotals, setSeriesBinning, setSeriesAggregation, setSeriesQuartileMethod,
    setSeriesParentLabelLayout, setSeriesLayoutFlag,
    # Groups
    getGroupAxes, getGroupLabelTextProps, getGroupDropLines, getGroupHiLowLines,
    getGroupSeriesLines, getGroupUpDownBars, getUpBarShapeProps, getDownBarShapeProps,
    setGroupLabelTextProp, getGroupGapWidth, setGroupGapWidth, getGroupOverlap, setGroupOverlap,
    # Axes
    getAxisShapeProps, getAxisTextProps, getAxisTitle, getAxisTitleText, getAxisTitleRef,
    getAxisTitleRange,
    getAxisTitleTextProps, getAxisGridlines, getAxisPartner,
    getAxisNumberFormatCode, getAxisNumberFormatLinked,
    getAxisMajorTickMark, getAxisMinorTickMark, getAxisTickLabelPos, getAxisOrientation,
    getAxisMin, getAxisMax, getAxisLogBase, getAxisCrosses, getAxisCrossesAt,
    getAxisMajorUnit, getAxisMinorUnit, getAxisLabelAlign, getAxisLabelOffset,
    getAxisMultiLevelLabels, getAxisCrossBetween,
    setAxisTitleText, setAxisTitleTextProp, setAxisTickLabelPos, 
    setAxisScaling, setAxisMin, setAxisMax, setAxisOrientation, 
    setAxisCrosses, setAxisCrossesAt, setAxisNumberFormatCode, setAxisNumberFormatLinked, 
    setAxisMajorTickMark, setAxisMinorTickMark, setAxisGridlines, setAxisTextProp, setAxisLine,
    setAxisMajorUnit, setAxisMinorUnit,
    # Data points, markers and labels
    getDataPointShapeProps, getDataPointMarker, getMarkerFill,
    getDataLabelText, getDataLabelTextProps, getDataLabelShapeProps,
    getDataLabelPosition, getDataLabelOffset,
    getLabelTextProp, getLabelFlag, getLabelPosition,
    setMarker, setMarkerSymbol, setMarkerSize, setMarkerFill,
    setMarkerLineColor, setMarkerLineWidth,
    setLabelText, setLabelTextProp, setLabelDeleted,
    # Trendlines and error bars
    getTrendlineShapeProps, getTrendlineLabelText, getTrendlineLabelTextProps,
    getTrendlineLabelShapeProps, getErrorBarsShapeProps, getErrorBarsCustomRefs,
    # Types users construct
    SchemeColor,
    # Predicates
    hasFill, hasLine

@static if VERSION >= v"1.11"
    eval(Meta.parse("""
    public AbstractChart, Chart, ChartEx, ChartRef, ChartRange, ChartRanges,
           ChartSeries, ChartAxis, ChartGroup, ChartDataPoint, ChartDataLabel,
           ChartMarker, ChartTrendline, ChartErrorBars, ChartUpDownBars,
           ChartExData, ChartExDimension, ChartExBinning,
           ColorSpec, DrawingColor, DrawingFill, DrawingLine, DrawingShapeProps,
           DrawingRunProps, DrawingParaProps, DrawingRun, DrawingParagraph,
           DrawingBodyProps, DrawingText,
           Effective, FormatSite
    """))
end

# Templates stay in src/data; this file is in src/charts.
const DATA_DIR = joinpath(dirname(@__DIR__), "data")

include("types.jl")
include("drawingml.jl")
include("discovery.jl")
include("creation.jl")
include("deletion.jl")
include("chartschema.jl")
include("chartprops.jl")
include("chartexprops.jl")

end # module Charts
