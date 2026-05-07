#
# ---- Some random helper functions
#
function isValidKw(kw::String, val::Union{String, Nothing}, valid::Vector{String})
    if isnothing(val) || val ∈ valid
        return true
    else
        throw(XLSXError("Invalid keyword $kw: $val. Valid values are $valid"))
    end
end
function uppercase_unquoted(s::AbstractString)
    result = IOBuffer()
    i = firstindex(s)
    inside_quote = false
    while i <= lastindex(s)
        c = s[i]
        if c == '\\' && nextind(s, i) <= lastindex(s)
            # Handle escaped character
            next_i = nextind(s, i)
            print(result, s[i:next_i])
            i = nextind(s, next_i)
        elseif c == '"'
            inside_quote = !inside_quote
            print(result, c)
            i = nextind(s, i)
        else
            if inside_quote
                print(result, c)
            else
                print(result, uppercase(c))
            end
            i = nextind(s, i)
        end
    end
    return String(take!(result))
end

#
# --- Standard conditional formats
#
function allCfs(ws::Worksheet)::Vector{XML.Node}
    sheetdoc = xmlroot(get_workbook(ws), ws.relationship_id) # find all the <conditionalFormatting> blocks in the worksheet's xml file
#    sheetdoc = xmlroot(ws.package, "xl/worksheets/sheet" * string(ws.sheetId) * ".xml") # find all the <conditionalFormatting> blocks in the worksheet's xml file
    return find_all_nodes("/" * SPREADSHEET_NAMESPACE_XPATH_ARG * ":worksheet/" * SPREADSHEET_NAMESPACE_XPATH_ARG * ":conditionalFormatting", sheetdoc)
end
function add_cf_to_XML(ws, new_cf) # Add a new conditional formatting to the worksheet XML.
    wb = get_workbook(ws)
    sheetdoc = xmlroot(get_workbook(ws), ws.relationship_id)
    l = insert_index(sheetdoc[end], "conditionalFormatting", WORKSHEET_ORDER)
    len = length(sheetdoc[end])
    if l != len
        insert!(sheetdoc[end].children, l+1, new_cf)
    else
        push!(sheetdoc[end], new_cf)
    end
end
function update_worksheet_cfx!(allcfs, cfx, ws, rng)
    pfx = get_prefix(ws)
    pfx = pfx == "" ? pfx : pfx * ":"

    matchcfs = filter(x -> x["sqref"] == string(rng), allcfs)   # Match range with existing conditional formatting blocks.
    l = length(matchcfs)
    if l == 0                                                   # No existing conditional formatting blocks for this range so create a new one.
        new_cf = XML.Element("$(pfx)conditionalFormatting"; sqref=rng)
        push!(new_cf, cfx)
        add_cf_to_XML(ws, new_cf)                               # Add the new conditional formatting block to the worksheet XML.
    elseif l == 1                                               # Existing conditional formatting block found for this range so add new rule to that block.
        push!(matchcfs[1], cfx)
    else
        throw(XLSXError("Too many conditional formatting blocks for range `$rng`. Must be one or none, found `$l`."))
    end
    update_worksheets_xml!(get_xlsxfile(ws))
end

#
# --- Conditional formats relying on Excel 2010 extensions
#
function allExtCfs(ws::Worksheet)::Vector{XML.Node}
    wb = get_workbook(ws)
    sheetdoc = xmlroot(get_workbook(ws), ws.relationship_id)    
#    sheetdoc = xmlroot(ws.package, "xl/worksheets/sheet" * string(ws.sheetId) * ".xml")
    i, j = get_idces(sheetdoc, "worksheet", "extLst")
    if isnothing(j)
        return Vector{XML.Node}()
    end
    extlst = sheetdoc[i][j]
    exts = XML.children(extlst)
    let cfs = nothing
        for ext in exts
            for c in XML.children(ext)
                if XML.tag(c) == "x14:conditionalFormattings"
                    cfs = c
                    break
                end
            end
        end
        return isnothing(cfs) ? Vector{XML.Node}() : XML.children(cfs)
    end
end
function make_extLst!(s)
    ext_list = XML.Element("extLst")
    ext_element = XML.Element("ext")
    ext_element["xmlns:x14"] = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"
    ext_element["uri"] = "{78C0D931-6437-407d-A8EE-F0AAD7539E65}"
    push!(ext_list, ext_element)
    push!(s, ext_list)
end
function make_extCfsBlock()
    extCf = XML.Element("x14:conditionalFormatting")
    extCf["xmlns:xm"] = "http://schemas.microsoft.com/office/excel/2006/main"
    return extCf
end
function update_worksheet_ext_cfx!(allcfs, cfx, ws, rng)
    wb = get_workbook(ws)
    sheetdoc = xmlroot(get_workbook(ws), ws.relationship_id)
#    sheetdoc = xmlroot(ws.package, "xl/worksheets/sheet" * string(ws.sheetId) * ".xml")
    i, j = get_idces(sheetdoc, "worksheet", "extLst")
    if isnothing(j)
        make_extLst!(sheetdoc[i])
        j = length(XML.children(sheetdoc[i]))
    end
    m, n = get_idces(sheetdoc[i], "extLst", "ext")
    @assert m==j
    if length(allcfs)==0                                        # No <conditionalFormattings> block. Need to create one.
        extcfs=XML.Element("x14:conditionalFormattings")
        push!(sheetdoc[i][j][n], extcfs)
    end
    matchcfs = filter(x -> XML.simple_value(x[end]) == string(rng), allcfs)   # Match range with existing conditional formatting blocks.
    o, p = get_idces(sheetdoc[i][j], "ext", "x14:conditionalFormattings")
    @assert o==n
    l = length(matchcfs)
    if l == 0                                                   # No existing conditional formatting blocks for this range so create a new one.
        new_cf = make_extCfsBlock()
        push!(new_cf, cfx)
        push!(new_cf, XML.Element("xm:sqref", XML.Text(string(rng))))
        push!(sheetdoc[i][j][n][p], new_cf)                        # Add the new conditional formatting block to the worksheet XML.
    elseif l == 1                                               # Existing conditional formatting block found for this range so add new rule to that block.
        pushfirst!(matchcfs[1], cfx)
    else
        throw(XLSXError("Too many conditional formatting blocks for range `$rng`. Must be one or none, found `$l`."))
    end
    update_worksheets_xml!(get_xlsxfile(ws))
end
function get_x14_icon(x14set)
    rule = XML.Element("x14:cfRule", type="iconSet", priority="1", id="XXXX-xxxx-XXXX") # replace id with UUID generated at time of use.
    if x14set == "Custom"
        icon = XML.Element("x14:iconSet", iconSet="3Arrows", custom="1")
    else
        icon = XML.Element("x14:iconSet", iconSet=x14set)
    end
    if x14set=="5Boxes"
        vals=[0, 20, 40, 60, 80]
    elseif x14set=="Custom"
        vals=[0]
    else
        vals=[0, 33, 67]
    end
    for v in vals
        cfvo = XML.Element("x14:cfvo", type="percent")
        push!(cfvo, XML.Element("xm:f", XML.Text(v)))
        push!(icon, cfvo)
    end
    push!(rule, icon)
    return rule
end

#
# ---- Formatting (styles) definitions for conditional formats
#
function Add_Cf_Dx(wb::Workbook, new_dx::XML.Node)::DxFormat
    # Check if the workbook already has a dxfs element. If not, add one.
    xroot = styles_xmlroot(wb)
    pfx = get_prefix("xl/styles.xml", get_xlsxfile(wb))
    pfx = pfx == "" ? pfx : pfx * ":"
    
    i, j = get_idces(xroot, "styleSheet", "dxfs")

    if isnothing(j) # No existing conditional formats so need to add a block (is this even possible?). Push everything lower down one.
        throw(XLSXError("No <dxfs> block found in the styles.xml file. Please submit an issue to report this and attach the Excel file you were working with."))
        #=  I don't think this can ever happen, so I've commented it out to improve coverage.
            k, l = get_idces(xroot, "styleSheet", "cellStyles")
            l += 1 # The dxfs block comes after the cellXfs block.
            len = length(xroot[k])
            i != k && throw(XLSXError("Some problem here!"))
            push!(xroot[k], xroot[k][end]) # duplicate last element then move everything else down one
            if l < len
                for pos = len-1:-1:l
                    xroot[k][pos+1] = xroot[k][pos]
                end
            end
            xroot[k][l] = XML.Element("dxsf", count="0")
            j = l
            println(XML.write(xroot[i][j]))
        =#
    else
        existing_dxf_elements_count = length(XML.children(xroot[i][j]))

        if parse(Int, xroot[i][j]["count"]) != existing_dxf_elements_count
            throw(XLSXError("Wrong number of xf elements found: $existing_cellxf_elements_count. Expected $(parse(Int, xroot[i][j]["count"]))."))
        end
    end

    #   Don't reuse duplicates here. Always create new!
    existingdx = XML.children(xroot[i][j])
    dxfs = unlink(xroot[i][j], ("dxfs", "dxf"), pfx) # Create the new <dxfs> Node
    if length(existingdx) > 0
        for c in existingdx
            push!(dxfs, c) # Copy each existing <dxf> into the new <dxfs> Node
        end
    end
    push!(dxfs, new_dx)

    xroot[i][j] = dxfs # Update the worksheet with the new cols.

    xroot[i][j]["count"] = string(existing_dxf_elements_count + 1)

    return DxFormat(existing_dxf_elements_count) # turns out this is the new index (because it's zero-based)

end
function get_dx(dxStyle::Union{Nothing,String}, format::Union{Nothing,Vector{Pair{String,String}}}, font::Union{Nothing,Vector{Pair{String,String}}}, border::Union{Nothing,Vector{Pair{String,String}}}, fill::Union{Nothing,Vector{Pair{String,String}}})::Dict{String,Dict{String,String}}
    if isnothing(dxStyle)
        if all(isnothing.([border, fill, font, format]))
            dx = highlights["redfilltext"]
        else
            dx = Dict{String,Dict{String,String}}()
            for att in ["font" => font, "fill" => fill, "border" => border, "format" => format]
                if !isnothing(last(att))
                    dxx = Dict{String,String}()
                    for i in last(att)
                        push!(dxx, first(i) => last(i))
                    end
                    push!(dx, first(att) => dxx)
                end
            end
        end
    elseif haskey(highlights, dxStyle)
        dx = highlights[dxStyle]
    else
        throw(XLSXError("Invalid dxStyle: $dxStyle. Valid options are: $(keys(highlights))."))
    end
    return dx
end
function get_new_dx(wb::Workbook, dx::Dict{String,Dict{String,String}})::XML.Node
    new_dx = XML.Element("dxf")
    for k in ["font", "format", "fill", "border"] # Order seems to be important to Excel.
        if haskey(dx, k)
            v = dx[k]
            if k == "fill"
                if !isnothing(v)
                    filldx = XML.Element("fill")
                    patterndx = XML.Element("patternFill")
                    for (y, z) in v
                        y in ["pattern", "bgColor", "fgColor"] || throw(XLSXError("Invalid fill attribute: $k. Valid options are: `pattern`, `bgColor`, `fgColor`."))
                        if y in ["fgColor", "bgColor"]
                            push!(patterndx, XML.Element(y, rgb=get_color(z)))
                        elseif y == "pattern" && z != "none"
                            patterndx["patternType"] = z
                        end
                    end
                    push!(filldx, patterndx)
                end
                push!(new_dx, filldx)
            elseif k == "font"
                if !isnothing(v)
                    fontdx = XML.Element("font")
                    for (y, z) in v
                        y in ["color", "bold", "italic", "under", "strike"] || throw(XLSXError("Invalid font attribute: $y. Valid options are: `color`, `bold`, `italic`, `under`, `strike`."))
                        if y == "color"
                            push!(fontdx, XML.Element(y, rgb=get_color(z)))
                        elseif y == "bold"
                            z == "true" && push!(fontdx, XML.Element("b", val="0"))
                        elseif y == "italic"
                            z == "true" && push!(fontdx, XML.Element("i", val="0"))
                        elseif y == "under"
                            z != "none" && push!(fontdx, XML.Element("u"; val="v"))
                        elseif y == "strike"
                            z == "true" && push!(fontdx, XML.Element(y))
                        end
                    end
                end
                push!(new_dx, fontdx)
            elseif k == "border"
                if !isnothing(v)
                    all([y in ["color", "style"] for y in keys(v)]) || throw(XLSXError("Invalid border attribute. Valid options are: `color`, `style`."))
                    borderdx = XML.Element("border")
                    cdx = haskey(v, "color") ? XML.Element("color", rgb=get_color(v["color"])) : nothing
                    sdx = haskey(v, "style") ? v["style"] : nothing
                    leftdx = XML.Element("left")
                    rightdx = XML.Element("right")
                    topdx = XML.Element("top")
                    bottomdx = XML.Element("bottom")
                    if !isnothing(sdx)
                        leftdx["style"] = sdx
                        rightdx["style"] = sdx
                        topdx["style"] = sdx
                        bottomdx["style"] = sdx
                    end
                    if !isnothing(cdx)
                        push!(leftdx, cdx)
                        push!(rightdx, cdx)
                        push!(topdx, cdx)
                        push!(bottomdx, cdx)
                    end
                end
                push!(borderdx, leftdx)
                push!(borderdx, rightdx)
                push!(borderdx, topdx)
                push!(borderdx, bottomdx)
                push!(new_dx, borderdx)
            elseif k == "format"
                if !isnothing(v)
                    if haskey(v, "format")
                        fmtCode = v["format"]
                        new_formatId = get_new_formatId(wb, fmtCode)
                        new_fmtCode = styles_numFmt_formatCode(wb, new_formatId)
                        fmtdx = XML.Element("numFmt"; numFmtId=string(new_formatId), formatCode=new_fmtCode)
                        push!(new_dx, fmtdx)
                    end
                end
            end
        end
    end
    return new_dx
end