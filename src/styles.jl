
# References
#
# 18.8.30 numFmt (Number Format)
# lists a table with predefined numFmtId.
# In this case, a numFmtId value is written on the xf record,
# but no corresponding numFmt element is written.
# Some of these Ids can be interpreted differently,
# depending on the UI language of the implementing application.
#
# General formatCodes ranges from 0 to 81.
#
# 18.8.10 cellXfs (Cell Formats)
import Base: isempty

function CellFormula(ws::Worksheet, val::AbstractFormula)
    CellFormula(val, default_cell_format(ws, val))
end
function CellValue(ws::Worksheet, val::CellConcreteType)
    CellValue(val, default_cell_format(ws, val))
end

id(format::CellDataFormat) = format.id
id(::EmptyCellDataFormat) = UInt64(0)
isempty(::CellDataFormat) = false
isempty(::EmptyCellDataFormat) = true

# The number of predefined number formats in XLSX
# Any custom number formats must have an id >= PREDEFINED_NUMFMT_COUNT
const PREDEFINED_NUMFMT_COUNT = 164

# these formats may appear differently in different editors
const DEFAULT_DATE_numFmtId = 14 # dd-mm-yyyy
const DEFAULT_TIME_numFmtId = 20 # h:mm
const DEFAULT_DATETIME_numFmtId = 22 # dd-mm-yyyy h:mm
const DEFAULT_NUMBER_numFmtId = 0 # General - seems like an OK default for now
const DEFAULT_BOOL_numFmtId = 0 # General - seems like an OK default for now

# Returns the default `CellDataFormat` for a type
default_cell_format(ws::Worksheet, ::AbstractFormula) = EmptyCellDataFormat()
default_cell_format(ws::Worksheet, ::CellConcreteType) = EmptyCellDataFormat()
default_cell_format(ws::Worksheet, ::Dates.Date) = get_num_style_index(ws, DEFAULT_DATE_numFmtId)
default_cell_format(ws::Worksheet, ::Dates.Time) = get_num_style_index(ws, DEFAULT_TIME_numFmtId)
default_cell_format(ws::Worksheet, ::Dates.DateTime) = get_num_style_index(ws, DEFAULT_DATETIME_numFmtId)

# Attempts to get CellDataFormat associated with a numFmtId and sets a default style if it is not found
# Use for ensuring default formats exist
function get_num_style_index(ws::Worksheet, numformatid::Integer)
    numformatid < 0 && throw(XLSXError("Invalid number format id"))

    wb = get_workbook(ws)
    style_index = styles_get_cellXf_with_numFmtId(wb, numformatid)
    if isempty(style_index)
        # adds default style <xf applyNumberFormat="1" borderId="0" fillId="0" fontId="0" numFmtId=formatid xfId="0"/>
        style_index = styles_add_cell_xf(wb, Dict("applyNumberFormat" => "1", "borderId" => "0", "fillId" => "0", "fontId" => "0", "numFmtId" => string(numformatid), "xfId" => "0"))
    end

    return style_index
end
function get_num_style_index(ws::Worksheet, allXfNodes::Vector{XML.Node}, numformatid::Integer)
    numformatid < 0 && throw(XLSXError("Invalid number format id"))

    wb = get_workbook(ws)
    style_index = styles_get_cellXf_with_numFmtId(allXfNodes, numformatid)
    if isempty(style_index)
        # adds default style <xf applyNumberFormat="1" borderId="0" fillId="0" fontId="0" numFmtId=formatid xfId="0"/>
        style_index = styles_add_cell_xf(wb, Dict("applyNumberFormat" => "1", "borderId" => "0", "fillId" => "0", "fontId" => "0", "numFmtId" => string(numformatid), "xfId" => "0"))
    end

    return style_index
end

# get styles document for workbook
function styles_xmlroot(workbook::Workbook)
#    xf=get_xlsxfile(workbook)
    if workbook.styles_xroot === nothing
        STYLES_RELATIONSHIP_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"
        if has_relationship_by_type(workbook, STYLES_RELATIONSHIP_TYPE)
            styles_target = get_relationship_target_by_type("xl", workbook, STYLES_RELATIONSHIP_TYPE)
            styles_root = xmlroot(get_xlsxfile(workbook), styles_target)

            # check root node name for styles.xml
            if get_default_namespace(styles_root[end]) != SPREADSHEET_NAMESPACE_XPATH_ARG
                throw(XLSXError("Unsupported styles XML namespace $(get_default_namespace(styles_root[end]))."))
            end
            XML.tag(styles_root[end]) != "styleSheet" && throw(XLSXError("Malformed package. Expected root node named `styleSheet` in `styles.xml`."))
            workbook.styles_xroot = styles_root
        else
            throw(XLSXError("Styles not found for this workbook."))
        end
    end

    return workbook.styles_xroot
end


# Returns the xf XML node element for style `index`.
# `index` is 0-based.
function styles_cell_xf(wb::Workbook, index::Integer)::XML.Node
    xroot = styles_xmlroot(wb)
    xf_elements = find_all_nodes("/" * SPREADSHEET_NAMESPACE_XPATH_ARG * ":styleSheet/" * SPREADSHEET_NAMESPACE_XPATH_ARG * ":cellXfs/" * SPREADSHEET_NAMESPACE_XPATH_ARG * ":xf", xroot)
    return xf_elements[index+1]
end
function styles_cell_xf(allXfNodes::Vector{XML.Node}, index::Integer)::XML.Node
    return allXfNodes[index+1]
end

# Queries numFmtId from cellXfs -> xf nodes."
function styles_cell_xf_numFmtId(wb::Workbook, index::Integer)::Int
    el = styles_cell_xf(wb, index)
    if !haskey(el, "numFmtId")
        return 0
    end
    return parse(Int, el["numFmtId"])
end
function styles_cell_xf_numFmtId(allXfNodes::Vector{XML.Node}, index::Integer)::Int
    el = styles_cell_xf(allXfNodes, index)
    if !haskey(el, "numFmtId")
        return 0
    end
    return parse(Int, el["numFmtId"])
end

# Defines a custom number format to render numbers, dates or text.
# Returns the index to be used as the `numFmtId` in a cellXf definition.
function styles_add_numFmt(wb::Workbook, format_code::AbstractString)::Integer
    xroot = styles_xmlroot(wb)

    numfmts = find_all_nodes("/" * SPREADSHEET_NAMESPACE_XPATH_ARG * ":styleSheet/" * SPREADSHEET_NAMESPACE_XPATH_ARG * ":numFmts", xroot)
    if isempty(numfmts)
        stylesheet = find_all_nodes("/" * SPREADSHEET_NAMESPACE_XPATH_ARG * ":styleSheet", xroot)[begin] # find first

        # We need to add the numFmts node directly after the styleSheet node
        # Move everything down one and then insert the new node at the top
        numfmts = XML.Element("numFmts", count="1")
        pushfirst!(stylesheet, numfmts)
    else
        numfmts = numfmts[1]
    end

    existing_numFmt_elements_count = length(XML.children(numfmts))
    fmt_code = existing_numFmt_elements_count + PREDEFINED_NUMFMT_COUNT
    new_fmt = XML.Element("numFmt";
        numFmtId=fmt_code,
        formatCode=XLSX.escape(format_code)
    )
    push!(numfmts, new_fmt)
    return fmt_code
end

const FontAttribute = Union{String,Pair{String,Pair{String,String}}}

# Queries numFmt formatCode field by numFmtId.
function styles_numFmt_formatCode(wb::Workbook, numFmtId::AbstractString)::String
    if haskey(builtinFormats, numFmtId)
        return builtinFormats[numFmtId]
    end
    xroot = styles_xmlroot(wb)
    nodes_found = find_all_nodes("/" * SPREADSHEET_NAMESPACE_XPATH_ARG * ":styleSheet/" * SPREADSHEET_NAMESPACE_XPATH_ARG * ":numFmts/" * SPREADSHEET_NAMESPACE_XPATH_ARG * ":numFmt", xroot)
    elements_found = filter(x -> XML.attributes(x)["numFmtId"] == numFmtId, nodes_found)
    length(elements_found) != 1 && throw(XLSXError("numFmtId $numFmtId not found."))
    return XML.attributes(elements_found[1])["formatCode"]
end

styles_numFmt_formatCode(wb::Workbook, numFmtId::Int)::String = styles_numFmt_formatCode(wb, string(numFmtId))

const DATETIME_CODES = ["d", "m", "yy", "h", "s", "a/p", "am/pm"]

function remove_formatting(code)
    # this regex should cover all the formatting cases found here(colors/conditionals/quotes/spacing):
    # https://support.office.com/en-us/article/create-or-delete-a-custom-number-format-78f2a361-936b-4c03-8772-09fab54be7f4
    ignoredformatting = r"""\[.{2,}?\]|&quot;.+?&quot;|_.|\\.|\*."""x # Had to add ? to "&quot;.+&quot;" to make it work. Don't understand what made this necessary!
    replace(code, ignoredformatting => "")
end

function styles_is_datetime(wb::Workbook, index::Integer)::Bool
    haskey(wb.buffer_styles_is_datetime, index) && return wb.buffer_styles_is_datetime[index]
    lock(wb.styles_lock) do
        if !haskey(wb.buffer_styles_is_datetime, index)
            isdatetime = false
            numFmtId = styles_cell_xf_numFmtId(wb, index)

            if (14 <= numFmtId && numFmtId <= 22) || (45 <= numFmtId && numFmtId <= 47)
                isdatetime = true
            elseif numFmtId > 81
                code = lowercase(styles_numFmt_formatCode(wb, numFmtId))
                code = remove_formatting(code)
                if any(map(x -> occursin(x, code), DATETIME_CODES))
                    isdatetime = true
                end
            end

            wb.buffer_styles_is_datetime[index] = isdatetime
        end
        return wb.buffer_styles_is_datetime[index]
    end
end

styles_is_datetime(wb::Workbook, fmt::CellDataFormat) = styles_is_datetime(wb, Int(fmt.id))

function styles_is_datetime(wb::Workbook, index::AbstractString)
    isempty(index) && throw(XLSXError("Something wrong here!"))
    styles_is_datetime(wb, parse(Int, index))
end

styles_is_datetime(ws::Worksheet, index) = styles_is_datetime(get_workbook(ws), index)

function styles_is_float(wb::Workbook, index::Integer)::Bool
    haskey(wb.buffer_styles_is_float, index) && return wb.buffer_styles_is_float[index]
    lock(wb.styles_lock) do
        if !haskey(wb.buffer_styles_is_float, index)
            isfloat = false
            numFmtId = styles_cell_xf_numFmtId(wb, index)

            if numFmtId == 2 || numFmtId == 4 || (7 <= numFmtId && numFmtId <= 11) || numFmtId == 39 || numFmtId == 40 || numFmtId == 44 || numFmtId == 48
                isfloat = true
            elseif numFmtId > 81
                code = styles_numFmt_formatCode(wb, numFmtId)
                code = remove_formatting(code)

                floatformats = r"""
                    \.[0#?]|
                    [0#?]e[+-]?[0#?]|
                    [0#?]/[0#?]|
                    %
                    """ix
                if occursin(floatformats, code)
                    isfloat = true
                end
            end

            wb.buffer_styles_is_float[index] = isfloat
        end

        return wb.buffer_styles_is_float[index]
    end
end

function styles_is_float(wb::Workbook, index::AbstractString)
    isempty(index) && throw(XLSXError("Something wrong here!"))
    styles_is_float(wb, parse(Int, index))
end

styles_is_float(wb::Workbook, fmt::CellDataFormat) = styles_is_float(wb, Int(fmt.id))
styles_is_float(ws::Worksheet, index) = styles_is_float(get_workbook(ws), index)

#=
Cell Xf element follows the XML format below.
This function queries the 0-based index of the first xf element that has the provided numFmtId.
Returns -1 if not found.

```
<styleSheet ...
    <cellXfs count="5">
            <xf borderId="0" fillId="0" fontId="0" numFmtId="0" xfId="0"/>
            <xf applyNumberFormat="1" borderId="0" fillId="0" fontId="0" numFmtId="14" xfId="0"/>
            <xf applyNumberFormat="1" borderId="0" fillId="0" fontId="0" numFmtId="20" xfId="0"/>
            <xf applyNumberFormat="1" borderId="0" fillId="0" fontId="0" numFmtId="22" xfId="0"/>
```
=#
function styles_get_cellXf_with_numFmtId(wb::Workbook, numFmtId::Int)::AbstractCellDataFormat
    xroot = styles_xmlroot(wb)
    allXfNodes = find_all_nodes("/" * SPREADSHEET_NAMESPACE_XPATH_ARG * ":styleSheet/" * SPREADSHEET_NAMESPACE_XPATH_ARG * ":cellXfs/" * SPREADSHEET_NAMESPACE_XPATH_ARG * ":xf", xroot)
    return styles_get_cellXf_with_numFmtId(allXfNodes, numFmtId)
end
function styles_get_cellXf_with_numFmtId(allXfNodes::Vector{XML.Node}, numFmtId::Int)::AbstractCellDataFormat
    if isempty(allXfNodes)
        return EmptyCellDataFormat()
    else
        for i in 1:length(allXfNodes)
            el = XML.attributes(allXfNodes[i])
            if !isnothing(el) && haskey(el, "numFmtId")
                if parse(Int, el["numFmtId"]) == numFmtId
                    return CellDataFormat(i - 1) # CellDataFormat is zero-indexed
                end
            end
        end

        # not found
        return EmptyCellDataFormat()
    end
end

function styles_add_cell_xf(wb::Workbook, attributes::Dict{String,String})::CellDataFormat
    new_xf = XML.Node(XML.Element, "xf", OrderedDict{String,String}(), nothing, nothing)
    for k in keys(attributes)
        new_xf[k] = attributes[k]
    end
    return styles_add_cell_xf(wb, new_xf)
end

function styles_add_cell_xf(wb::Workbook, new_xf::XML.Node)::CellDataFormat
    xroot = styles_xmlroot(wb)
    i, j = get_idces(xroot, "styleSheet", "cellXfs")
    existing_cellxf_elements_count = length(XML.children(xroot[i][j]))
    if parse(Int, xroot[i][j]["count"]) != existing_cellxf_elements_count
        throw(XLSXError("Wrong number of xf elements found: $existing_cellxf_elements_count. Expected $(parse(Int, xroot[i][j]["count"]))."))
    end
    # Check new_xf doesn't duplicate any existing xf. If yes, use that rather than create new.
    for (k, node) in enumerate(XML.children(xroot[i][j]))
        if node == new_xf
            return CellDataFormat(k - 1) # CellDataFormat is zero-indexed
        end
    end
    push!(xroot[i][j], new_xf)
    xroot[i][j]["count"] = string(existing_cellxf_elements_count + 1)

    return CellDataFormat(existing_cellxf_elements_count) # turns out this is the new index (because it's zero-based)

end
