include_dependency(joinpath(@__DIR__, "data", "metadata.xml"))
const METADATA_XML_DATA = read(joinpath(@__DIR__, "data", "metadata.xml"))

const RGX_FORMULA_SHEET_CELL = r"!\$?[A-Z]+\$?[0-9]" # to recognise sheetcell references like "otherSheet!A1"

# Prefixes needed for newer Excel functions - previously two different prefixes (hence Dict) but now only one.
# Retain as Dict in case Excel introduces other namespace prefixes in future
# name => prefix
const EXCEL_FUNCTION_PREFIX = Dict(
    # Core dynamic array + higher-order
    "MAKEARRAY" => "_xlfn.",
    "SEQUENCE" => "_xlfn.",
    "RANDARRAY" => "_xlfn.",
    "ANCHORARRAY" => "_xlfn.", # used internally to handle spill references like A1#
    "LAMBDA" => "_xlfn.", # not well supported at present
    "MAP" => "_xlfn.",
    "REDUCE" => "_xlfn.",
    "SCAN" => "_xlfn.",
    "BYROW" => "_xlfn.",
    "BYCOL" => "_xlfn.",
    "LET" => "_xlfn.",  # not well supported at present. Parameters may be tagged with _xlpm.

    # Array shaping/stacking
    "VSTACK" => "_xlfn.",
    "HSTACK" => "_xlfn.",
    "TOCOL" => "_xlfn.",
    "TOROW" => "_xlfn.",
    "WRAPROWS" => "_xlfn.",
    "WRAPCOLS" => "_xlfn.",
    "TAKE" => "_xlfn.",
    "DROP" => "_xlfn.",
    "EXPAND" => "_xlfn.",
    "CHOOSECOLS" => "_xlfn.",
    "CHOOSEROWS" => "_xlfn.",

    # Sort/filter/distinct/group/pivot
    "SORT" => "_xlfn.",  # historically also _xlws.
    "SORTBY" => "_xlfn.",  # historically also _xlws.
    "FILTER" => "_xlfn.",  # historically also _xlws.
    "UNIQUE" => "_xlfn.",  # historically also _xlws.
    "GROUPBY" => "_xlfn.",
    "PIVOTBY" => "_xlfn.",

    # Text spill functions
    "TEXTSPLIT" => "_xlfn.",
    "TEXTBEFORE" => "_xlfn.",
    "TEXTAFTER" => "_xlfn.",

    # Lookup
    "XLOOKUP" => "_xlfn.",
    "XMATCH" => "_xlfn.",
    "STOCKHISTORY" => "_xlfn.",
    "FIELDVALUE" => "_xlfn.",

    # Other modern functions commonly serialized with _xlfn.
    "IFS" => "_xlfn.",
    "MAXIFS" => "_xlfn.",
    "MINIFS" => "_xlfn.",
    "SWITCH" => "_xlfn.",
    "IFNA" => "_xlfn.",
    "SINGLE" => "_xlfn.",
    "CONCAT" => "_xlfn.",
    "TEXTJOIN" => "_xlfn.",

    # Image insertion (modern Excel)
    "IMAGE" => "_xlfn."
)

const SPILL_FUNCTIONS = Set([
    "SEQUENCE",
    "RANDARRAY",
    "MAKEARRAY",
    "UNIQUE",
    "SORT",
    "SORTBY",
    "FILTER",
    "TOCOL",
    "TOROW",
    "VSTACK",
    "HSTACK",
    "WRAPROWS",
    "WRAPCOLS",
    "EXPAND",
    "TEXTSPLIT",
    "STOCKHISTORY",
    "GROUPBY",
    "PIVOTBY",
    "ANCHORARRAY"
    #    "XLOOKUP",
    #    "TEXTBEFORE",
    #    "TEXTAFTER"
])
# Map of aggregator functions used in LAMBDA functions as name => prefix
# Retain as Dict in case Excel introduces other namespace prefixes in future
const GROUPBY_AGGREGATORS = Dict(
    # Eta-reduced aggregators (become _xleta.FUNC)
    "SUM" => "_xleta.",
    "AVERAGE" => "_xleta.",
    "COUNT" => "_xleta.",
    "COUNTA" => "_xleta.",
    "COUNTBLANK" => "_xleta.",
    "MIN" => "_xleta.",
    "MAX" => "_xleta.",
    "MEDIAN" => "_xleta.",
    "MODE" => "_xleta.",
    "MODE.SNGL" => "_xleta.",
    "MODE.MULT" => "_xleta.",
    "PRODUCT" => "_xleta.",
    "STDEV" => "_xleta.",
    "STDEVP" => "_xleta.",
    "STDEV.S" => "_xleta.",
    "STDEV.P" => "_xleta.",
    "VAR" => "_xleta.",
    "VARP" => "_xleta.",
    "VAR.S" => "_xleta.",
    "VAR.P" => "_xleta."
    # Everything else stays as _xlfn.LAMBDA(...)
    # (or _xlfn.FUNC if used directly)
)

Base.isempty(f::Formula) = f.formula == ""
Base.isempty(f::ReferencedFormula) = f.formula == ""
Base.isempty(f::FormulaReference) = false # always links to another formula
Base.hash(f::Formula, h::UInt) = hash(f.unhandled, hash(f.formula, h))
Base.hash(f::FormulaReference, h::UInt) = hash(f.unhandled, hash(f.id, h))
Base.hash(f::ReferencedFormula, h::UInt) = hash(f.unhandled, hash(f.ref, hash(f.id, hash(f.formula, h))))

function new_ReferencedFormula_Id(ws::Worksheet)
    # return the first positive integer (or 0) not currently used as a ReferencedFormula Id

    ids = Set{Int}()
    for r in eachrow(ws) # first find all Ids currently in use
        for cell in values(r.rowcells)
            if cell.formula
                rf = get_formula_from_cache(ws, cell.ref)
                if rf isa ReferencedFormula
                    push!(ids, rf.id)
                end
            end
        end
    end

    id = 0
    while id ∈ ids # then find first Id not currently in use
        id += 1
    end
    return id
end

function build_reference_index(ws::Worksheet)
    # Create Dict of all ReferencedFormulae in worksheet
    refs = Dict{Int,ReferencedFormula}() # Id => ReferencedFormula
    for r in eachrow(ws)
        for cell in values(r.rowcells)
            if cell.formula
                f = get_formula_from_cache(ws, cell.ref)
                if f isa ReferencedFormula
                    refs[f.id] = f
                end
            end
        end
    end
    return refs
end

formula(f::ReferencedFormula)::Formula = Formula(f.formula, nothing, f.ref, f.unhandled) 

function get_referenced_formula(ws::Worksheet, cellref::CellRef; refs::Union{Nothing,Dict{Int,ReferencedFormula}}=nothing)
    # find the actual formula a cell's FormulaReference refers to
    if isnothing(refs)
        refs = build_reference_index(ws)
    end
    cell = getcell(ws, cellref)
    if cell.formula
        f = get_formula_from_cache(ws, cellref)
        if isa(f, FormulaReference)
            id = f.id
            haskey(refs, id) || throw(XLSXError("No ReferencedFormula found for id=$id"))
            offset = cell_offset(CellRange(refs[id].ref).start, cellref)
            f = shift_excel_references(refs[id].formula, offset)
            return f
        else
            throw(XLSXError("Cell `$CellRef` does not contain a formula reference!"))
        end
    end
end

# If overwriting a cell containing a ReferencedFormula, need to re-reference all referring cells.
# The ReferencedFormula will be in the top left cell of the referenced block. Need to re-reference 
# the rest of the block on this top row (without the first, overwritten cell) and then the rest of 
# the block without this top row. Need to do this as two new, separate rectangular blocks with the 
# ReferencedFormula in the first cell of each and the other cells set to FormulaReferences.
# 
# overwritten newRF1    FR1       FR1       FR1
# newRF2      FR2       FR2       FR2       FR2
# FR2         FR2       FR2       FR2       FR2
#
# Note that a block of referencedFormulas can have a separate referencedFormula block set within it! 
# 
function rereference_formulae(ws::Worksheet, cell::Cell)
    wb = get_workbook(ws)
    old_range = CellRange(get_formula_from_cache(ws, cell.ref).ref)
    ranges = CellRange[]
    if old_range.stop.column_number > old_range.start.column_number
        push!(ranges, CellRange(CellRef(old_range.start.row_number, old_range.start.column_number + 1), CellRef(old_range.start.row_number, old_range.stop.column_number)))
    end
    if old_range.stop.row_number > old_range.start.row_number
        push!(ranges, CellRange(CellRef(old_range.start.row_number + 1, old_range.start.column_number), CellRef(old_range.stop.row_number, old_range.stop.column_number)))
    end

    for newrng in ranges
        if size(newrng) == (1, 1) # replace with simple formula for single cell
            f=get_formula_from_cache(ws, cell.ref)
            f.ref=string(newrng.stop)
            add_formula_to_cache(ws, newrng.stop, formula(f)) 
         else # update reference and re-index all referring formulas
            f = get_formula_from_cache(ws, cell.ref)
            delete!(wb.formulas, SheetCellRef(ws.name, cell.ref))
            cell.formula=false
            newid = new_ReferencedFormula_Id(ws)
            rereference_formulae(ws, cell, f, newrng, newid)
        end
    end
end

function rereference_formulae(ws::Worksheet, oldcell::Cell, f::AbstractFormula, newrng::CellRange, newid::Int)
    wb = get_workbook(ws)
    oldform = f.formula
    oldunhandled = f.unhandled
    offset = cell_offset(oldcell.ref, newrng.start)
    newform = ReferencedFormula(shift_excel_references(oldform, offset), newid, string(newrng), oldunhandled)
    for fr in newrng
        newfr = getcell(ws, fr)
        if fr != newrng.start
            if newfr.formula
                formula = get_formula_from_cache(ws, fr)
                if formula isa FormulaReference && formula.id == f.id
                    t = encode(newfr.datatype)
                    add_formula_to_cache(ws, fr, FormulaReference(newid, oldunhandled))
                    setdata!(ws, Cell(wb, fr, t, string(newfr.style), string(newfr.value), string(newfr.meta), true))
                end
            end
        else
            t = encode(oldcell.datatype)
            add_formula_to_cache(ws, fr, newform)
            setdata!(ws, Cell(wb, fr, t, string(newfr.style), string(newfr.value), string(newfr.meta), true))
        end
    end
    return nothing
end

# shift the relative cell references ina formula when shifting a ReferencedFormula
function shift_excel_references(formula::String, offset::Tuple{Int32,Int32})
    # Regex to match Excel-style cell references (e.g., A1, $A$1, A$1, $A1)
    pattern = r"\$?[A-Z]{1,3}\$?[1-9][0-9]*"
    row_shift, col_shift = offset

    initial = [string(x.match) for x in eachmatch(pattern, formula)]
    result = Vector{String}()

    for ref in eachmatch(pattern, formula)
        # Extract parts using regex
        m = match(r"(\$?)([A-Z]{1,3})(\$?)([1-9][0-9]*)", ref.match)
        col_abs, col_letters, row_abs, row_digits = m.captures

        col_num = decode_column_number(col_letters)
        row_num = parse(Int, row_digits)

        # Apply shifts only if not absolute
        new_col = col_abs == "\$" ? col_letters : encode_column_number(col_num + col_shift)
        new_row = row_abs == "\$" ? row_digits : string(row_num + row_shift)

        push!(result, col_abs * new_col * row_abs * new_row)
    end

    pairs = Dict(zip(initial, result))
    if !isempty(pairs)
        formula = replace(formula, pairs...)
    end

    return formula
end

# Replace formula references to a sheet that has been deleted with #REF errors
function update_formulas_missing_sheet!(wb::Workbook, name::String)
    xf = get_xlsxfile(wb)
    quoted_name = quoteit(name)
    sheet_prefix = quoted_name * "!"
    
    # Pre-compile regex pattern
    ref_pattern = Regex(sheet_prefix * raw"\$?[A-Z]{1,3}\$?[1-9][0-9]*")
    
    deleted_formula_ids = Set{Int}()
    error_formula_ids = Set{Int}()
    formulas_to_delete = Set{SheetCellRef}()
    
    # Single pass: process non-reference formulas
    for (ref, f) in wb.formulas
        if f isa FormulaReference
            continue
        end
        
        ref_str = string(ref)
        
        # Check if formula is in the deleted sheet
        if occursin(sheet_prefix, ref_str)
            if f isa ReferencedFormula
                push!(deleted_formula_ids, f.formula.id)
            end
            push!(formulas_to_delete, ref)
            continue
        end
        
        # Check if formula refers to the deleted sheet
        if occursin(sheet_prefix, f.formula)
            old_formula = f.formula
            # Replace references to deleted sheet
            f.formula = replace(f.formula, ref_pattern => "#REF!")
            # Update cell if formula changed
            if old_formula != f.formula
                cell = getcell(xf, ref)
                cell.datatype = CT_ERROR
                cell.value = UInt64(XL_REF)
                f isa ReferencedFormula &&push!(error_formula_ids, f.id) # all FormulaReferences to this formula need updating
            end
        end
    end
    
    # Second pass: handle FormulaReferences
    for (ref, f) in wb.formulas
        if f isa FormulaReference
            if f.id ∈ error_formula_ids && f.id ∉ deleted_formula_ids
                cell = getcell(xf, ref)
                cell.datatype = CT_ERROR
                cell.value = UInt64(XL_REF)
            end
            if occursin(sheet_prefix, string(ref)) || f.id ∈ deleted_formula_ids
                push!(formulas_to_delete, ref)
            end
        end
    end
    
    # Delete all marked formulas
    for ref in formulas_to_delete
        delete!(wb.formulas, ref)
    end
    
    return nothing
end

# Update formula references to a sheet that has been renamed with new names
function update_formulas_renamed_sheet!(wb::Workbook, old_name::String, new_name::String)
    xf = get_xlsxfile(wb)
    old_quoted = quoteit(old_name)
    new_quoted = quoteit(new_name)
    old_prefix = old_quoted * "!"
    new_prefix = new_quoted * "!"
    
    formulas_to_move = Dict{SheetCellRef, Tuple{SheetCellRef, AbstractFormula}}()  # old_ref => (new_ref, formula)
    
    # First pass: update formula values and identify keys that need updating
    for (ref, f) in wb.formulas
        ref_str = string(ref)
        needs_key_update = occursin(old_prefix, ref_str)
        
        # Calculate new reference if key needs updating
        new_ref = if needs_key_update
            new_ref_str = replace(ref_str, old_quoted => new_quoted)
            SheetCellRef(new_ref_str)  
        else
            nothing
        end
        
        # Update formula values for non-FormulaReference types
        if f isa ExplicitFormula
            needs_value_update = occursin(old_prefix, f.formula)
            
            if needs_value_update
                f.formula = replace(f.formula, old_prefix => new_prefix)
            end
        end
        
        # Store formulas that need key updates
        if needs_key_update
            formulas_to_move[ref] = (new_ref, f)
        end
    end
    
    # Second pass: update keys by deleting old and inserting with new keys
    for (old_ref, (new_ref, f)) in formulas_to_move
        delete!(wb.formulas, old_ref)
        wb.formulas[new_ref] = f
    end
    
    return nothing
end

"""
    split_function_args(formula::String; fname::Union{Nothing,String}=nothing) -> Vector{String}

Given a formula string like `=GROUPBY(E1:E151,A1:D151,LAMBDA(x,AVERAGE(x)),3,1)`,
return the arguments as a vector of strings:
["E1:E151", "A1:D151", "LAMBDA(x,AVERAGE(x))", "3", "1"].

If `fname` is provided, it will look specifically for that function name.
If not, it will match the first identifier followed by '('.
"""
function split_function_args(formula::String; fname::Union{Nothing,String}=nothing)
    # Build regex for the function name
    pat = isnothing(fname) ?
          r"([A-Za-z_][A-Za-z0-9_]*)\s*\(" :
          Regex("\\b$(fname)\\s*\\(", "i")

    m = match(pat, formula)
    isnothing(m) && return String[]
    start = m.offset + length(m.match)
    depth = 1
    buf = IOBuffer()
    args = String[]
    i = start
    while i <= lastindex(formula) && depth > 0
        c = formula[i]
        if c == '('
            depth += 1
            print(buf, c)
        elseif c == ')'
            depth -= 1
            if depth > 0
                print(buf, c)
            end
        elseif c == ',' && depth == 1
            push!(args, strip(String(take!(buf))))
        else
            print(buf, c)
        end
        i += 1
    end
    if position(buf) > 0
        push!(args, strip(String(take!(buf))))
    end
    return args
end

# Simplified regex for cell/range references
const RGX_RANGE_RE = r"[A-Z]+\$?\d+:[A-Z]+\$?\d+"

"""
    needs_array_attr(fname::String, args::Vector{String}) -> Bool

Determine if a function call will spill and thus require t="array".
"""
function needs_array_attr(fname::AbstractString, args::Vector{String})
    f = uppercase(fname)

    # Helper: does an argument look like a range?
    is_range(arg) = occursin(RGX_RANGE_RE, arg)

    # INDEX(array, row_num, [col_num])
    if f == "INDEX"
        # If row_num or col_num are omitted or themselves arrays, INDEX can return multiple cells
        return is_range(args[1]) && (length(args) < 3 || args[2] == "" || args[3] == "")
    end

    # OFFSET(reference, rows, cols, [height], [width])
    if f == "OFFSET"
        # If first arg is a range, OFFSET can spill
        if is_range(args[1])
            return true
        end
        # If height or width > 1, OFFSET returns a multi-cell reference
        if length(args) >= 5
            return (tryparse(Int, args[4]) |> x -> x > 1 ? true : false) ||
                   (tryparse(Int, args[5]) |> x -> x > 1 ? true : false)
        end
        return false
    end

    # IF(logical_test, value_if_true, value_if_false)
    if f == "IF"
        # IF spills if either branch is a range/array
        return any(is_range, args[1:3])
    end

    # CHOOSE(index_num, value1, value2, ...)
    if f == "CHOOSE"
        # If any chosen value is a range/array, CHOOSE can spill
        return any(is_range, args[2:end])
    end

    return false
end

"""
    is_array_formula(formula::String) -> Bool

Detects whether a formula implies an array spill (thus needing t="array").
Currently flags:
  - Binary operators (+, -, *, /, ^) applied to ranges
  - Functions with range arguments that return arrays (e.g. MMULT, TRANSPOSE)
"""
function is_array_formula(formula::String)
    # Simplistic regex for a cell/range reference like A1:B13
    range_re = "[\$]?[A-Z]+[\$]?\\d+:[\$]?[A-Z]+[\$]?\\d+"

    # Case 1: is a bare range
    if occursin(Regex("^=$(range_re)\$"), formula)
        return true
    end
    # Case 2: binary operator applied to a range
    if occursin(Regex("[+\\-*/^]\\s*$(range_re)|$(range_re)\\s*[+\\-*/^]"), formula)
        return true
    end

    # Case 3: known array-returning functions
    array_funcs = ["FREQUENCY", "LINEST", "LOGEST", "MINVERSE", "MMULT", "MUNIT", "MODE.MULT", "TRANSPOSE", "TREND", "GROWTH"]
    for f in array_funcs
        if occursin(Regex("\\b$f\\s*\\("), formula)
            return true
        end
    end

    # Case 4: _xlfn functions
    #    array_funcs = SPILL_FUNCTIONS # not all of these do? May need a different set here!
    for f in SPILL_FUNCTIONS
        if occursin(Regex("\\b$f\\s*\\("), formula)
            return true
        end
    end

    # Case 5: functions that spill conditional on given range
    array_funcs = ["INDEX", "OFFSET", "IF", "CHOOSE"]
    for f in array_funcs
        args = split_function_args(formula; fname=f)
        !isempty(args) && return needs_array_attr(f, args)
    end
    return false
end

# Finally got one that works from Claude
const SPILL_REF_RE = r"""
    (                                       # capture group 1: full reference before #
        (?:                                 # optional sheet/workbook prefix
            (?:
                (?:'[^']+'|\[[^\]]+\])      # quoted sheet or [Book] - already handles Unicode
                | [\p{L}\p{N}_]+            # unquoted sheet - Unicode letters/numbers
            )
            !
        )?
        (?:
            \$?[A-Za-z]{1,3}\$?\d+          # cell reference (A1 notation is ASCII-only)
          | [\p{L}_][\p{L}\p{N}_.\\/]*      # named range - Unicode support
            (?!\[)                          # negative lookahead: not followed by [
          | (?:                             # structured reference
                (?:'[^']+'|[\p{L}_][\p{L}\p{N}_ ]*)   # table name with Unicode
                (?:
                    \[(?:[^\[\]]|\[[^\]]*\])*\]      # single [...] with potential nested []
                  | \[\[(?:[^\[\]]|\[[^\]]*\])+\]\]  # [[...]] with nested elements
                )
            )
        )
    )
    \#                                      # literal spill operator
"""x

function anchor_spill_refs(formula::String)
    return replace(formula, SPILL_REF_RE => s -> "ANCHORARRAY($(s[1:end-1]))")
end

function process_dynamic_array_functions(xf::XLSXFile, cellref::CellRef, val::String; raw::Bool, spill::Union{Nothing,Bool})

    t = ""
    ref = ""
    cm = ""

    formula = val
    if !raw
        if occursin("#", formula) # handle spill references like A1# or myName#
            formula = anchor_spill_refs(formula)
        end

        # Handle the 3rd/4th (function) arg of GROUPBY and PIVOTBY (NOTE. Can't handle anything but simple aggregation functions yet)
        m = match(r"(?i)\b(GROUPBY|PIVOTBY)\s*\(", formula)
        if !isnothing(m)
            fname = uppercase(m.captures[1])
            idx = fname == "GROUPBY" ? 3 : 4
            # Extract arguments
            args = split_function_args(formula; fname=fname)
            if length(args) >= idx

                # Transform 3rd/4th argument if it's a bare identifier
                agg = args[idx]
                if haskey(GROUPBY_AGGREGATORS, uppercase(agg))
                    prefix = GROUPBY_AGGREGATORS[uppercase(agg)]
                    args[idx] = prefix * agg
                else
                    throw(XLSXError("Currently only simple aggregation functions like sum or average are supported with `raw=false`."))
                end

                # Reconstruct
                formula = fname * "(" * join(args, ",") * ")"
            end
        end

        for (k, v) in EXCEL_FUNCTION_PREFIX # add prefixes to any array functions
            r = "(?i)\\b" * k
            formula = replace(formula, Regex(r) => v * k) # replace any dynamic array function name with its prefixed name
        end
    end

    iaf = is_array_formula(formula)
    if !isnothing(spill) && (spill || !spill && iaf) || isnothing(spill) && iaf
        t = "array"
        ref = cellname(cellref) * ":" * cellname(cellref)
        cm = "1"
        if !haskey(xf.files, "xl/metadata.xml") # add metadata.xml on first use of a dynamicArray formula
            xf.data["xl/metadata.xml"] = XML.Node(XML.Raw(copy(METADATA_XML_DATA)))
            xf.files["xl/metadata.xml"] = true # set file as read
            add_override!(xf, "/xl/metadata.xml", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheetMetadata+xml")
            rId = add_relationship!(get_workbook(xf), "metadata.xml", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sheetMetadata")
        end
    end
    return formula, t, ref, cm
end

const EXTERNAL_REF_RE = r"\[(\d+)\]([\p{L}\p{N}_]+)!\$?[A-Za-z]+\$?\d+"

# Extract all external references from a formula string (eg like "[1]Sheet1!$A$1")
function get_ext_refs(formula::String)
    [ExternalRef(parse(Int, m.captures[1]),
        m.captures[2],
        m.match) for m in eachmatch(EXTERNAL_REF_RE, formula)] # workbook_path to be filled in later
end

# Lookup an external file reference from its index in the workbook's externalReferences
function get_external_workbook_path(xf::XLSXFile, id::Int)
    extRef = get_wb_ext_refs(xf)
    rel = get_relationship_target_by_id("xl", get_workbook(xf), extRef[id])
    extXml = xmlroot(xf, rel)
    i, j = get_idces(extXml, "externalLink", "externalBook") # we are looking for ExternalBook to find an external filename
    isnothing(i) && throw(XLSXError("Malformed external reference in workbook. Missing externalLink node."))
    isnothing(j) && throw(XLSXError("Malformed external reference in workbook. Missing externalBook node."))
    k, l = get_idces(extXml[i], "externalBook", "externalBookPr")
    k == j || throw(XLSXError("Something wrong here!"))

    # find the file name directly, if present, searching in order:
    # 1. externalBook filename attribute
    # 2. externalBookPr filename attribute
    # 3. first alternateUrls r:id attribute (to be further resolved via relationships)
    atts = XML.attributes(extXml[i][k])
    haskey(atts, "filename") && return atts["filename"] # externalBook filename attribute
    if !isnothing(l)
        atts = XML.attributes(extXml[i][k][l])
        haskey(atts, "filename") && return atts["filename"] # externalBookPr filename attribute
    end
    k, l = get_idces(extXml[i], "externalBook", "xxl21:alternateUrls")
    if !isnothing(l)
        atts = XML.attributes(extXml[i][k][l][1]) # prefer the first alternateUrls r:id if multiple
        haskey(atts, "r:id") || throw(XLSXError("Something wrong here!"))
        rId = atts["r:id"]
        # now need a second lookup of this further r:id
        altUrls = XML.children(xmlroot(xf, "xl/externalLinks/_rels/$(basename(rel)).rels")[end])
        for c in altUrls
            atts = XML.attributes(c)
            if haskey(atts, "Id") && atts["Id"] == rId
                haskey(atts, "Target") || throw(XLSXError("Something wrong here!"))
                return atts["Target"]
            end
        end
    end
    throw(XLSXError("Unreachable reached!"))
end

function get_formula_from_cache(ws::Worksheet, ref::CellRef)
    #    get_xlsxfile(ws).is_writable || throw(XLSXError("Workbook is not writable - cannot get formulae!"))
    wb = get_workbook(ws)
    return wb.formulas[SheetCellRef(combine_sheet_ref(ws, ref))]
end
function add_formula_to_cache(ws::Worksheet, ref::CellRef, f::AbstractFormula)
    wb = get_workbook(ws)
    wb.formulas[SheetCellRef(combine_sheet_ref(ws, ref))] = f
end

"""
    setFormula(ws::Worksheet, RefOrRange::AbstractString, formula::String; raw=false, spill=false)
    setFormula(xf::XLSXFile,  RefOrRange::AbstractString, formula::String; raw=false, spill=false)

    setFormula(sh::Worksheet, row, col, formula::String; raw=false, spill=false)

Set the Excel formula to be used in the given cell or cell range.

Formulae must be valid Excel formulae and written in US english with comma
separators. Cell references may be absolute or relative references in either 
the row or the column or both (e.g. `\$A\$2`). No validation of the specified 
formula is made by `XLSX.jl` and formulae are usually stored verbatim, as given.

Non-contiguous ranges are not supported by `setFormula`. Set the formula in 
each cell or contiguous range separately.

Use `raw=true` if entering a formula in xml-ready format to prevent any processing 
by `setFormula`.

Use `spill=true` to force the formula to be treated as an array formula that spills 
and `spill=false` to prevent it being treated as such. By default `spill=nothing` and
`setFormula` will determine whether a formula should spill or not automatically.

Keyword options should be rarely needed - `setFormula` should handle most formulae.

Since `XLSX.jl` does not and cannot replicate all the functions built in to Excel, 
setting a formula in a cell does not permit the cell's value to be re-calculated 
within `XLSX.jl`. Instead, although the formula is properly added to the cell, the 
value is set to `missing`. However, the saved `XLSXFile` is set to force Excel to 
re-calculate on opening. 

If a cell spills but any of the cells in the spill range already contains a value, Excel will 
show a `#SPILL` error.

More details can be found in the section [Using Formulas](@ref).

See also [XLSX.getFormula](@ref).

# Examples:

```julia

julia> using XLSX

julia> f=newxlsx("setting formulas")
XLSXFile("blank.xlsx") containing 1 Worksheet
            sheetname size          range        
-------------------------------------------------
     setting formulas 1x1           A1:A1        


julia> s=f[1]
1×1 Worksheet: ["setting formulas"](A1:A1) 

julia> s["A2:A10"]=1
1

julia> s["A1:J1"]=1
1

julia> setFormula(s, "B2:J10", "=A2+B1") # adds formulae but cannot update calculated values
"=A2+B1"

julia> addsheet!(f, "trig functions")
1×1 Worksheet: ["trig functions"](A1:A1) 

julia> f
XLSXFile("mytest.xlsx") containing 2 Worksheets
            sheetname size          range        
-------------------------------------------------
     setting formulas 10x10         A1:J10
       trig functions 1x1           A1:A1


julia> s2=f[2]
1×1 Worksheet: ["trig functions"](A1:A1)

julia> for i=1:100, s2[i, 1] = 2.0*pi*i/100.0; end

julia> setFormula(s2, "B1:B100", "=sin(A1)")

julia> setFormula(s2, "C1:C100", "=cos(A1)")

julia> setFormula(s2, "D1:D100", "=sin(A1)^2 + cos(A1)^2")

julia> XLSX.getFormula(s2, "D100")
"=sin(A100)^2 + cos(A100)^2"

julia> f=newxlsx("mysheet")
XLSXFile("blank.xlsx") containing 1 Worksheet
            sheetname size          range
-------------------------------------------------
              mysheet 1x1           A1:A1

julia> s=f[1]
1×1 Worksheet: ["mysheet"](A1:A1)

julia> s["A1"]=["Header1" "Header2" "Header3"; 1 2 3; 4 5 6; 7 8 9; 1 2 3; 4 5 6; 7 8 9]
7×3 Matrix{Any}:
  "Header1"   "Header2"   "Header3"
 1           2           3
 4           5           6
 7           8           9
 1           2           3
 4           5           6
 7           8           9

julia> setFormula(s, "E1:G1", "=sort(unique(A2:A7),,-1)") # using dynamic array functions
```
![image|320x500](../images/SortUnique.png)

```julia

f = CSV.read("iris.csv", XLSXFile) # read a CSV file into an XLSXFile

XLSX.setFormula(f[1], "G1", "=GROUPBY(E1:E151,A1:D151,AVERAGE,3,1)") # Find average of each characteristic by species
"_xlfn.GROUPBY(E1:E151,A1:D151,_xleta.AVERAGE,3,1)"

f[1]["M1"] = "versicolor"
XLSX.setFormula(f[1], "M2", "=VLOOKUP(M1,G1#,3,FALSE)") # Lookup average sepal width for versicolor using the spill range of G1
"=VLOOKUP(M1,_xlfn.ANCHORARRAY(G1),3,FALSE)"

XLSX.setFormula(f[1], "G10", "_xlfn.GROUPBY(E1:E151,A1:D151,_xlfn.LAMBDA(_xlpm.x,AVERAGE(_xlpm.x)),3,1)"; raw=true) # using `raw` format
"_xlfn.GROUPBY(E1:E151,A1:D151,_xlfn.LAMBDA(_xlpm.x,AVERAGE(_xlpm.x)),3,1)"
```

!!! note

    It is not yet possible for `setFormula` to create external references in formulas. 
    It is therefore not possible to set a formula that refers to a sheet in another Excel 
    file.

"""
setFormula(w, r, f::AbstractString; raw::Bool=false, spill::Union{Nothing,Bool}=nothing) = setFormula(w, r; val=f, raw=raw, spill=spill) # move formula to a kw to take advantage of cellformat-helpers functions
setFormula(w, r, c, f::AbstractString; raw::Bool=false, spill::Bool=false) = setFormula(w, r, c; val=f, raw=raw, spill=spill) # move formula to a kw to take advantage of cellformat-helpers functions
setFormula(ws::Worksheet, ref::SheetCellRef; kw...) = do_sheet_names_match(ws, ref) && setFormula(ws, ref.cellref; kw...)
setFormula(ws::Worksheet, rng::SheetCellRange; kw...) = do_sheet_names_match(ws, rng) && setFormula(ws, rng.rng; kw...)
setFormula(ws::Worksheet, rng::SheetColumnRange; kw...) = do_sheet_names_match(ws, rng) && setFormula(ws, rng.colrng; kw...)
setFormula(ws::Worksheet, rng::SheetRowRange; kw...) = do_sheet_names_match(ws, rng) && setFormula(ws, rng.rowrng; kw...)
setFormula(ws::Worksheet, colrng::ColumnRange; kw...) = process_columnranges(setFormula, ws, colrng; kw...)
setFormula(ws::Worksheet, rowrng::RowRange; kw...) = process_rowranges(setFormula, ws, rowrng; kw...)
setFormula(ws::Worksheet, ref_or_rng::AbstractString; kw...) = process_ranges(setFormula, ws, ref_or_rng; kw...)
setFormula(xl::XLSXFile, sheetcell::String; kw...) = process_sheetcell(setFormula, xl, sheetcell; kw...)
setFormula(ws::Worksheet, row::Integer, col::Integer; kw...) = setFormula(ws, CellRef(row, col); kw...)
setFormula(ws::Worksheet, row::Union{Integer,UnitRange{<:Integer}}, col::Union{Integer,UnitRange{<:Integer}}; kw...) = setFormula(ws, CellRange(CellRef(first(row), first(col)), CellRef(last(row), last(col))); kw...)
setFormula(ws::Worksheet, row::Union{Integer,UnitRange{<:Integer}}, ::Colon; kw...) = process_colon(setFormula, ws, row, nothing; kw...)
setFormula(ws::Worksheet, ::Colon, col::Union{Integer,UnitRange{<:Integer}}; kw...) = process_colon(setFormula, ws, nothing, col; kw...)
setFormula(ws::Worksheet, ::Colon, ::Colon; kw...) = setFormula(ws, :; kw...)
setFormula(ws::Worksheet, ::Colon; kw...) = process_colon(setFormula, ws, nothing, nothing; kw...)
# These all give rise to (potentially) non-contiguous ranges so would need special handling if implemented. The cellformat-helpers functions won't work here because the kw will need processing.
#setFormula(ws::Worksheet, ncrng::NonContiguousRange; kw...) = process_ncranges(setFormula, ws, ncrng; kw...)
#setFormula(ws::Worksheet, row::Union{Vector{Int},StepRange{<:Integer}}, ::Colon; kw...) = process_veccolon(setFormula, ws, row, nothing; kw...)
#setFormula(ws::Worksheet, ::Colon, col::Union{Vector{Int},StepRange{<:Integer}}; kw...) = process_veccolon(setFormula, ws, nothing, col; kw...)
#setFormula(ws::Worksheet, row::Union{Integer,UnitRange{<:Integer}}, col::Union{Vector{Int},StepRange{<:Integer}}; kw...) = process_vecint(setFormula, ws, row, col; kw...)
#setFormula(ws::Worksheet, row::Union{Vector{Int},StepRange{<:Integer}}, col::Union{Integer,UnitRange{<:Integer}}; kw...) = process_vecint(setFormula, ws, row, col; kw...)
#setFormula(ws::Worksheet, row::Union{Vector{Int},StepRange{<:Integer}}, col::Union{Vector{Int},StepRange{<:Integer}}; kw...) = process_vecint(setFormula, ws, row, col; kw...)
function setFormula(ws::Worksheet, rng::CellRange; val::AbstractString, raw::Bool, spill::Union{Nothing,Bool}=nothing)

    xf = get_xlsxfile(ws)

    if xf.is_writable == false
        throw(XLSXError("Cannot set formula because because XLSXFile is not writable."))
    end

    is_array = is_array_formula(val)

    is_sheetcell = occursin(RGX_FORMULA_SHEET_CELL, val)

    if is_array || is_sheetcell || occursin("#", val) # Don't use ReferencedFormulas for sheetcell formulas or dynamic array functions. Set each cell individually.
        start = rng.start
        first = true
        f1 = ""
        for c in rng
            offset = (c.row_number - start.row_number, c.column_number - start.column_number)
            newval = shift_excel_references(val, offset)
            f = setFormula(ws, c, newval)
            first && (first = false; f1 = f) # return the formula from the first cell only
        end
        return f1
    end

    # now we know the formula does not include a dynamic array function, so no need to process for prefixes
    first_cell = getcell(ws, rng.start)
    if !(first_cell isa EmptyCell) && first_cell.formula
        formula = get_formula_from_cache(ws, first_cell.ref)
        if formula isa ReferencedFormula
            if CellRange(formula.ref) == rng # range matches, so just need to change the referenced formula
                formula.formula = val
                return val
            end
        end
    end

    newid = new_ReferencedFormula_Id(ws)
    f = ""
    for c in rng
        if c == rng.start
            newform = ReferencedFormula(val, newid, string(rng), nothing)
            f = newform.formula
        else
            newform = FormulaReference(newid, nothing)
        end
        cell = getcell(ws, c)
        if cell isa EmptyCell || cell.style == UInt8(0)
            setdata!(ws, c, CellFormula(ws, newform))
        else
            setdata!(ws, c, CellFormula(newform, CellDataFormat(cell.style)))
        end
    end
    return f
end
function setFormula(ws::Worksheet, cellref::CellRef; val::AbstractString, raw::Bool=false, spill::Union{Nothing,Bool}=nothing)
    # cell references in formulas have already been adjusted for offset in a range before here

    xf = get_xlsxfile(ws)

    if xf.is_writable == false
        throw(XLSXError("Cannot set formula because because XLSXFile is not writable."))
    end

    c = getcell(ws, cellref)

    formula, t, ref, cm = process_dynamic_array_functions(xf, cellref, val; raw, spill)
    f = raw ? val : formula
    if c isa EmptyCell || c.style == UInt8(0)
        setdata!(ws, cellref, CellFormula(ws, Formula(f, t == "" ? nothing : t, ref == "" ? nothing : ref, nothing)))
    else
        setdata!(ws, cellref, CellFormula(Formula(f, t == "" ? nothing : t, ref == "" ? nothing : ref, nothing), CellDataFormat(c.style)))
    end
    c = getcell(ws, cellref)
    c.meta = cm == "" ? UInt16(0) : parse(UInt16, cm)
    return f
end

"""
    getFormula(sh::Worksheet, cr::String; get_external_refs::Bool=false) -> Union{String,Nothing}
    getFormula(xf::XLSXFile, cr::String; get_external_refs::Bool=false) -> Union{String,Nothing}

    getFormula(sh::Worksheet, row::Int, col::Int; get_external_refs::Bool=false) -> Union{String,Nothing}

Get the formula for a single cell reference in a worksheet `sh` or XLSXFile `xf`.
The specified cell must be within the sheet dimension.

If the cell does not contain any formula (but is not an `EmptyCell`), return an empty string ("").
If the cell is an EmptyCell, return `nothing`.

If the cell contains a `FormulaReference`, look up the actual formula.

A formula may contain references to cells in external workbooks, in the form
`[index]SheetName!A1` where `index` is an integer providing an internal Excel reference 
to the external workbook. Use the keyword option `get_external_refs=true` to replace
the index with the actual workbook path (as stored in the workbook's externalReferences).
By default, `get_external_refs=false` and the formula is returned unchanged.

See also [XLSX.setFormula](@ref).

# Examples:
```julia

julia> setFormula(s, "B2:B5", "=A2+2")
"=A2+2"

julia> XLSX.getcell(s, "B2")
XLSX.Cell(B2, "", "", "", "", XLSX.ReferencedFormula("=A2+2", 0, "B2:B5", nothing))

julia> XLSX.getcell(s, "B3")
XLSX.Cell(B3, "", "", "", "", XLSX.FormulaReference(0, nothing))

julia> XLSX.getFormula(s, XLSX.CellRef("B3"))
"=A3+2"

julia> XLSX.getFormula(s, XLSX.CellRef("A1"))
"HYPERLINK(\"https://www.bbc.co.uk/news\", \"BBC News\")"

julia> XLSX.getFormula(s, XLSX.CellRef("B1"))
"[1]Sheet1!\$A\$1"

julia> XLSX.getFormula(s, XLSX.CellRef("B1"); get_external_refs=true)
"[https://d.docs.live.net/.../Documents/Julia/XLSX/linked-2.xlsx]Sheet1!\$A\$1"
```

"""
getFormula(ws::Worksheet, cr::String; kw...) = process_get_cellname(getFormula, ws, cr; kw...)
getFormula(xl::XLSXFile, sheetcell::String; kw...) = process_get_sheetcell(getFormula, xl, sheetcell; kw...)
getFormula(ws::Worksheet, row::Integer, col::Integer; kw...) = getFormula(ws, CellRef(row, col); kw...)
function getFormula(ws::Worksheet, cellref::CellRef; get_external_refs::Bool=false)
    cellref ∉ get_dimension(ws) && throw(XLSXError("Cell $cellref is out of range for worksheet '$(ws.name)'"))
    xf = get_xlsxfile(ws)
    if !xf.use_cache_for_sheet_data
        throw(XLSXError("Cannot get formula because cache is not enabled."))
    end
    cell = getcell(ws, cellref)
    (isa(cell, EmptyCell) || !cell.formula) && return nothing
    f = get_formula_from_cache(ws, cellref)
    if f isa FormulaReference
        # need to look up the ReferencedFormula this FormulaReference points to
        f = get_referenced_formula(ws, cell.ref; refs=build_reference_index(ws))
    else
        f = f.formula
    end

    if get_external_refs # to address #224
        iserror(cell) && cell.value == XL_REF && throw(XLSXError("Cell $cellref contains a `#REF!` error - external reference not found"))
        ext = get_ext_refs(f)
        for e in ext
            extLink = get_external_workbook_path(get_xlsxfile(ws), e.index)
            f = replace(f, "[" * string(e.index) * "]" => "[" * extLink * "]") # replace index with actual workbook path
        end
    end

    if !startswith(f, "=")
        f = "=" * f
    end

    return f
end

