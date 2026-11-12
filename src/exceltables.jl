#
# ====================================================================================== Excel Tables
#

const REL_TABLE  = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/table"
const MIME_TABLE = "application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"

# SUBTOTAL function codes in the 100s range (ignore manually-hidden rows),
# which is what Excel writes for table totals rows.
const TOTALS_ROW_FUNCTIONS = Dict{Symbol,Tuple{String,Int}}(
    :sum       => ("sum",       109),
    :average   => ("average",   101),
    :count     => ("count",     103),  # OOXML "count" == Excel COUNTA (counts non-blank, incl. text)
    :countnums => ("countNums", 102),  # OOXML "countNums" == Excel COUNT (counts numeric cells only)
    :max       => ("max",       104),
    :min       => ("min",       105),
    :stddev    => ("stdDev",    107),
    :var       => ("var",       110),
)

const TOTALS_FUNCTION_BY_NAME = Dict(first(v) => k for (k, v) in TOTALS_ROW_FUNCTIONS)

function parse_totals_settings(sheet::Worksheet, t::Table)
    t.has_totals_row || return Pair[]

    table_doc = get_xml_data(get_xlsxfile(sheet), _table_part_path(sheet, t.name))
    i, j = get_idces(table_doc, "table", "tableColumns")
    column_nodes = collect(xml_elements(table_doc[i][j]))

    totals_row = t.ref.stop.row_number
    col0 = _col_start(t)
    settings = Pair[]

    for (idx, col_name) in enumerate(t.columns)
        node  = column_nodes[idx]
        func  = get_attr(node, "totalsRowFunction", "")
        label = get_attr(node, "totalsRowLabel", "")

        if func == "custom"
            # A custom function's formula lives only in the cell, so read it
            # before the totals row is overwritten. getFormula prepends "=",
            # but setFormula (used by settotals!) wants a bare formula, so
            # strip it back off.
            fstr = getFormula(sheet, CellRef(totals_row, col0 + idx - 1))
            isnothing(fstr) && throw(XLSXError(
                "Column `$col_name` of table `$(t.name)` is marked as a custom totals function but its totals cell holds no formula."))
            push!(settings, col_name => (:custom, lstrip(fstr, '=')))
        elseif func != ""
            sym = get(TOTALS_FUNCTION_BY_NAME, func, nothing)
            isnothing(sym) && throw(XLSXError("Unrecognized totalsRowFunction `$func` on column `$col_name` of table `$(t.name)`."))
            push!(settings, col_name => sym)
        elseif label != ""
            push!(settings, col_name => label)
        end
    end
    return settings
end

# Compact one-line form — used implicitly by default Vector{Table} printing,
# error messages, and anywhere a Table appears embedded in other output.
function Base.show(io::IO, t::Table)
    t = _resolve(t)
    nrows = _last_data_row(t) - _first_data_row(t) + 1
    print(io, nrows, "x", length(t.columns), " Table (id=", t.id, ", \"", t.name, "\", ", t.ref)
    t.has_totals_row && print(io, ", +totals")
    print(io, ")")
end

# Richer multi-line form — used when a single Table is the direct display
# value, e.g. at the REPL: `julia> XLSX.table(sheet, "Sales")`
function Base.show(io::IO, ::MIME"text/plain", t::Table)
    t = _resolve(t)
    println(io, "XLSX.Table: \"", t.name, "\"", t.name == t.display_name ? "" : " (displayName: \"$(t.display_name)\")")
    println(io, "  id      : ", t.id)
    println(io, "  range   : ", t.ref)
    println(io, "  columns : ", join(t.columns, ", "))
    if isnothing(t.style)
        println(io, "  style   : none")
    else
        style_desc = something(t.style.name, "(unnamed)")
        flags = String[]
        t.style.show_first_column   && push!(flags, "first col")
        t.style.show_last_column    && push!(flags, "last col")
        t.style.show_row_stripes    && push!(flags, "row stripes")
        t.style.show_column_stripes && push!(flags, "col stripes")
        println(io, "  style   : ", style_desc, isempty(flags) ? "" : " (" * join(flags, ", ") * ")")
    end
    print(io, "  totals  : ", t.has_totals_row ? "yes" : "no")
end

# Compact form for the style struct on its own (e.g. `t.style` at the REPL)
function Base.show(io::IO, s::TableStyleInfo)
    print(io, "TableStyleInfo(", something(s.name, "none"))
    flags = String[]
    s.show_first_column   && push!(flags, "first")
    s.show_last_column    && push!(flags, "last")
    s.show_row_stripes    && push!(flags, "rows")
    s.show_column_stripes && push!(flags, "cols")
    isempty(flags) || print(io, ", ", join(flags, "+"))
    print(io, ")")
end

"""
    _find_table_part(sheet::Worksheet, name::AbstractString) -> (rid, path)

Locate the table part for the Excel Table `name` on `sheet`, by walking the
worksheet's `<tableParts>` → `r:id` → worksheet `.rels` → target path. Returns
the relationship id and the package path of `xl/tables/tableN.xml`.

Throws an `XLSXError` if `sheet` has no `<tableParts>` element, no relationship
file, or no table part whose `name` attribute matches.
"""
function _find_table_part(sheet::Worksheet, name::AbstractString)::Tuple{String,String}
    xf = get_xlsxfile(sheet)
    sheet_path = get_relationship_target_by_id("xl", get_workbook(sheet), sheet.relationship_id)
    sheet_dir, sheet_file = rsplit(sheet_path, "/"; limit=2)
    rels_path = "$sheet_dir/_rels/$sheet_file.rels"

    !haskey(xf.data, rels_path) &&
        throw(XLSXError("Internal error: `$(sheet.name)` has table `$name` but no relationship file."))

    rels_root  = xml_root_element(xf.data[rels_path])
    sheet_root = xml_root_element(get_xml_data(xf, sheet_path))

    tp_container_els = elements_with_tag(sheet_root, "tableParts")
    isempty(tp_container_els) &&
        throw(XLSXError("Internal error: `$(sheet.name)` has table `$name` in cache but no <tableParts> element."))
    tp_container = tp_container_els[1]

    rel_els = elements_with_tag(rels_root, "Relationship")
    for tp in elements_with_tag(tp_container, "tablePart")
        rid = get_attr(tp, "r:id")
        k = findfirst(r -> get_attr(r, "Id") == rid, rel_els)
        isnothing(k) && continue
        path = resolve_relative_target(sheet_dir, get_attr(rel_els[k], "Target"))
        get_attr(xml_root_element(get_xml_data(xf, path)), "name") == name && return (rid, path)
    end

    throw(XLSXError("Internal error: could not locate table part for `$name`."))
end

# Convenience wrapper for callers that don't need the relationship id.
_table_part_path(sheet::Worksheet, name::AbstractString)::String = last(_find_table_part(sheet, name))

function parse_table_style_info(table_doc::XML.Node)
    i, j = get_idces(table_doc, "table", "tableStyleInfo")
    isnothing(j) && return nothing

    node = table_doc[i][j]
    attrs = XML.attributes(node)
    isnothing(attrs) && return TableStyleInfo(nothing, false, false, false, false)

    return TableStyleInfo(
        get(attrs, "name", nothing),
        get(attrs, "showFirstColumn", "0") == "1",
        get(attrs, "showLastColumn", "0") == "1",
        get(attrs, "showRowStripes", "0") == "1",
        get(attrs, "showColumnStripes", "0") == "1",
    )
end

"""
    remove_attr!(node::XML.Node, key::String)

Remove the attribute `key` from `node`, if present. No-op if `node` has no
attributes at all, or doesn't have `key`.
"""
function remove_attr!(node::XML.Node, key::String)
    attrs = XML.attributes(node)
    isnothing(attrs) && return nothing
    filter!(p -> first(p) != key, node.attributes)
    return nothing
end

"""
    _is_valid_table_display_name(name::AbstractString) -> Bool

Whether `name` is a valid Excel Table `displayName`: starts with a letter,
underscore, or other Unicode identifier-start character; contains only
identifier characters, underscores, or periods thereafter; no spaces.
Uses the same (Unicode-aware) character classification as `normalizename`,
so a name normalized via `normalizename` is always accepted here.
"""
function _is_valid_table_display_name(name::AbstractString)::Bool
    isempty(name) && return false
    Base.is_id_start_char(first(name)) || return false
    return all(c -> c == '.' || Base.is_id_char(c), name)
end

function parse_table_columns(table_doc::XML.Node)
    i, j = get_idces(table_doc, "table", "tableColumns")
    isnothing(j) && throw(XLSXError("Malformed table part: missing <tableColumns>."))

    columns = String[]
    for col_node in xml_elements(table_doc[i][j])
        localname(col_node) != "tableColumn" && continue
        attrs = XML.attributes(col_node)
        (isnothing(attrs) || !haskey(attrs, "name")) &&
            throw(XLSXError("Malformed <tableColumn>: missing required `name` attribute."))
        push!(columns, attrs["name"])
    end
    return columns
end

function parse_table_xml(table_doc::XML.Node, filename::AbstractString, sheet::Worksheet)::Table
    root_els = xml_elements(table_doc)
    isempty(root_els) && throw(XLSXError("Malformed table part $filename: no root element."))
    table_node = last(root_els)
    localname(table_node) != "table" &&
        throw(XLSXError("Malformed table part $filename: root should be <table>, got <$(localname(table_node))>."))

    attrs = XML.attributes(table_node)
    isnothing(attrs) && throw(XLSXError("Malformed table part $filename: <table> has no attributes."))
    haskey(attrs, "id")   || throw(XLSXError("<table> in $filename missing required `id` attribute."))
    haskey(attrs, "name") || throw(XLSXError("<table> in $filename missing required `name` attribute."))
    haskey(attrs, "ref")  || throw(XLSXError("<table> in $filename missing required `ref` attribute."))

    # totalsRowCount is the number of totals rows in `ref` now. totalsRowShown is
    # history — whether one has ever been shown — and defaults to true, so Excel
    # writes "0" for a table that never had one and omits it otherwise. It says
    # nothing about the current range. (Checked against Excel: a table with its
    # totals row switched on, then off, carries neither attribute.)
    has_totals = something(tryparse(Int, get(attrs, "totalsRowCount", "0")), 0) > 0
    has_header_row = get(attrs, "headerRowCount", "1") != "0"

return Table(
        parse(Int, attrs["id"]),
        attrs["name"],
        get(attrs, "displayName", attrs["name"]),
        CellRange(attrs["ref"]),
        parse_table_columns(table_doc),
        has_header_row,
        has_totals,
        parse_table_style_info(table_doc),
        sheet,
    )
end

function get_worksheet_tables(xf::XLSXFile, ws::Worksheet)::Vector{Table}
    r_ids = get_worksheet_table_rids(xf, ws)
    isempty(r_ids) && return Table[]

    tables = Table[]
    for r_id in r_ids
        target = get_worksheet_relationship_target(xf, ws, r_id)
        table_doc = xmlroot(xf, target)  # table parts are fully parsed like any other part
        push!(tables, parse_table_xml(table_doc, target, ws))
    end
    return tables
end

"""
    tables(ws::Worksheet) -> Vector{Table}
    tables(xf::XLSXFile)  -> Vector{Table}

All Excel Tables defined on `ws` or on `xf` across every worksheet, in sheet order and then in
document order within each sheet. Empty if the worksheet or workbook has none. Chartsheets are
skipped.

Each `Table` knows the worksheet it belongs to (its `sheet` field), so the sheet is
recoverable from the result.

```julia
# Examples
```julia
julia> XLSX.tables(sheet)
2-element Vector{XLSX.Table}:
 Table(id=1, "IO_Table", A1:C8, 3 cols)
 Table(id=2, "Age_height", E1:G6, 3 cols, +totals)
 
julia> for t in XLSX.tables(f)
           println(t.sheet.name, ": ", t.name, " ", t.ref)
       end
Sheet1: IO_Table A1:C8
Sheet1: Age_height E1:G6
Sheet2: with_total A3:C11
```

!!! note

    [`XLSX.tables`](@ref)`(xf::XLSXFile)` must examine every worksheet, so
    on a large workbook opened lazily it will cause each sheet's Table metadata to be
    read (Cell data are not read in this process).

See also [`XLSX.table`](@ref).
"""
function tables(ws::Worksheet)::Vector{Table}
    if isnothing(ws.tables_cache)
        ws.tables_cache = get_worksheet_tables(get_xlsxfile(ws), ws)
    end
    return ws.tables_cache
end
function tables(xf::XLSXFile)::Vector{Table}
    wb = get_workbook(xf)
    result = Table[]
    for ws in wb.sheets
        is_chartsheet(wb, ws.name) && continue
        append!(result, tables(ws))
    end
    return result
end

# A `Table` is a handle (sheet + name), not a snapshot: it is immutable, so a
# handle taken before an `appendtable!`/`settotals!` still carries the old
# `ref` and `has_totals_row`. Public entry points that read a Table's extent
# resolve it afresh rather than trusting the caller's copy.
_resolve(t::Table)::Table = table(t.sheet, t.name)

"""
    table(ws::Worksheet, name::AbstractString) -> Table
    table(wb::Workbook, name::AbstractString) -> Table
    table(xf::XLSXFile, name::AbstractString) -> Table
    table(ws::Worksheet, id::Integer) -> Table
    table(wb::Workbook, id::Integer) -> Table
    table(xf::XLSXFile, id::Integer) -> Table

Look up a single table by name or workbook-scoped numeric id, searching
a single worksheet or across every worksheet in the workbook. 

Throws `KeyError` if not found.

# Examples
```julia
julia> XLSX.table(sheet, "Age_height")
XLSX.Table: "Age_height"
  id      : 2
  range   : E1:G6
  columns : name, age, height
  style   : TableStyleMedium2 (row stripes)
  totals  : yes

julia> XLSX.table(sheet, 2)  # same table, looked up by id
XLSX.Table: "Age_height"
  id      : 2
  range   : E1:G6
  columns : name, age, height
  style   : TableStyleMedium2 (row stripes)
  totals  : yes
```
"""
function table(ws::Worksheet, name::AbstractString)::Table
    idx = findfirst(t -> t.name == name, tables(ws))
    isnothing(idx) && throw(KeyError(name))
    return tables(ws)[idx]
end

function table(ws::Worksheet, id::Integer)::Table
    idx = findfirst(t -> t.id == id, tables(ws))
    isnothing(idx) && throw(KeyError(id))
    return tables(ws)[idx]
end
function table(wb::Workbook, name::AbstractString)::Table
    for ws in wb.sheets
        is_chartsheet(wb, ws.name) && continue
        idx = findfirst(t -> t.name == name, tables(ws))
        idx !== nothing && return tables(ws)[idx]
    end
    throw(KeyError(name))
end
function table(wb::Workbook, id::Integer)::Table
    for ws in wb.sheets
        is_chartsheet(wb, ws.name) && continue
        idx = findfirst(t -> t.id == id, tables(ws))
        idx !== nothing && return tables(ws)[idx]
    end
    throw(KeyError(id))
end
table(xf::XLSXFile, name::AbstractString) = table(get_workbook(xf), name)
table(xf::XLSXFile, id::Integer) = table(get_workbook(xf), id)

function build_table_xml(id::Int, name::String, display_name::String, ref::CellRange,
                          columns::Vector{String}, has_totals_row::Bool,
                          style::Union{TableStyleInfo,Nothing})::XML.Node

    buf = IOBuffer()
    print(buf, """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>""")
    print(buf, """<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" """)
    print(buf, """id="$(id)" name="$(XML.escape(name; attribute = true))" displayName="$(XML.escape(display_name; attribute = true))" ref="$(ref)\"""")
    # in build_table_xml:
    if has_totals_row
        print(buf, """ totalsRowShown="1" totalsRowCount="1\"""")
    end
    print(buf, ">")

    # autoFilter covers header+data rows only, excluding any totals row
    filter_stop_row = has_totals_row ? ref.stop.row_number - 1 : ref.stop.row_number
    filter_ref = CellRange(CellRef(ref.start.row_number, ref.start.column_number),
                            CellRef(filter_stop_row, ref.stop.column_number))
    print(buf, """<autoFilter ref="$(filter_ref)"/>""")

    print(buf, """<tableColumns count="$(length(columns))">""")
    for (i, col_name) in enumerate(columns)
        print(buf, """<tableColumn id="$(i)" name="$(XML.escape(col_name; attribute = true))"/>""")
    end
    print(buf, "</tableColumns>")

    if !isnothing(style)
        print(buf, "<tableStyleInfo")
        !isnothing(style.name) && print(buf, """ name="$(style.name)\"""")
        print(buf, """ showFirstColumn="$(style.show_first_column ? "1" : "0")\"""")
        print(buf, """ showLastColumn="$(style.show_last_column ? "1" : "0")\"""")
        print(buf, """ showRowStripes="$(style.show_row_stripes ? "1" : "0")\"""")
        print(buf, """ showColumnStripes="$(style.show_column_stripes ? "1" : "0")\"""")
        print(buf, "/>")
    end

    print(buf, "</table>")
    return parse(String(take!(buf)), XML.Node)
end

"""
    addtable!(
        sheet::Worksheet,
        ref::Union{CellRange,AbstractString};
        name::AbstractString="",
        style::Union{AbstractString,Nothing}=nothing,
        has_totals_row::Bool=false,
    ) -> Table

Create a new Excel Table over `ref` on `sheet`, turning an existing range of
cells into a Table object (banding, filter dropdowns, structured references,
and the `Table` metadata itself) without changing any of the underlying data.

`ref` must already contain data before calling `addtable!`: specifically, the
first row of `ref` is read as the table's header row, so every cell in that
row must already contain the column name you want (write these first, e.g.
with `sheet[...] = ...` or [`XLSX.writetable!`](@ref)). `addtable!` only wraps
existing cells in a Table; it does not write any header, data, or totals
values into the sheet itself.

`ref` must span at least two rows — a header row plus at least one data row.
Excel does not support header-only tables (a single-row `ref` is rejected).

`ref` must not overlap a merged range or another table on the sheet. Excel
discards such a table when it opens the file, so an error is thrown instead.
Nor may it overlap the sheet's AutoFilter (see [`setAutoFilter`](@ref)), since a
table has filter buttons of its own.

If `name` is not given, a unique name is generated (`"Table1"`, `"Table2"`,
...). Table names are workbook-scoped and must not collide with another
table's name or with a defined name anywhere in the workbook. Nor may a name
be a cell address, in either reference style and in any case (`T1`, `xfd1`,
`R1C1`, `RC`, or just `R` or `C`): Excel refuses a file with such a table.

`style` sets the table's visual style and should be one of Excel's built-in
table style names, matching the "Table Styles" gallery in Excel's Table
Design ribbon:
- `"TableStyleLightnn"` where nn is between 1 and 21
- `"TableStyleMediumnn"` where nn is between 1 and 28
- `"TableStyleDarknn"` where nn is between 1 and 11
- `"None"` for no style. 

`style` is not validated and any string is passed straight through as the
`tableStyleInfo`'s `name` attribute; an unrecognized name will make Excel 
fall back to its default table appearance rather than causing an error. 
If omitted, Excel treats the table as having no explicit style (its own 
default appearance applies).

`has_totals_row=true` marks the last row of `ref` as the table's totals row. 
This only sets the flag that tells Excel to *reserve and display* that row as 
a totals row; it does not populate any totals formula or label into the cells
themselves — write whatever content you want into that row's cells yourself
(or leave them blank) before or after calling `addtable!`, or use
[`XLSX.settotals!`](@ref) afterward to set per-column totals functions or
labels. With `has_totals_row=true` the last row of `ref` must be empty: if it
already has content, an error is thrown and no table is added. (Excel refuses
a file whose totals row holds plain values.)

# Examples
```julia
julia> sheet[1, :] = ["id", "name", "score"]
julia> sheet[2, :] = [1, "alice", 10.5]
julia> sheet[3, :] = [2, "bob", 20.0]

julia> XLSX.addtable!(sheet, "A1:C3"; name="Results", style="TableStyleMedium2")
```

See also [`XLSX.tables`](@ref), [`XLSX.table`](@ref), [`XLSX.deletetable!`](@ref), [`XLSX.settotals!`](@ref).
"""
function addtable!(sheet::Worksheet, ref::CellRange;
                    name::AbstractString="",
                    style::Union{AbstractString,Nothing}=nothing,
                    has_totals_row::Bool=false)::Table

    _require_writable(get_xlsxfile(sheet), "add a table")
    name = _checked_new_table_name(sheet, name, ref)

    columns = _table_header_columns(sheet, ref)
    has_totals_row && _check_totals_row_empty(sheet, ref, name, columns)

    _check_table_overlap(sheet, ref, name)
    _add_table_part!(sheet, ref, name, columns, style, has_totals_row)
    return table(sheet, name)
end

# The name of a new table over `ref` (see `_new_table_name`), checked to be one Excel
# accepts, and `ref` checked to have a data row.
function _checked_new_table_name(sheet::Worksheet, name::AbstractString, ref::CellRange)::String
    name = _new_table_name(get_workbook(sheet), sheet, name)
    !_is_valid_table_display_name(name) &&
        throw(XLSXError("Table displayName `$name` is not a valid identifier (must start with a letter or underscore, contain only letters, digits, underscores, or periods thereafter, and no spaces)."))
    _is_cell_address_like(name) &&
        throw(XLSXError("Table name `$name` is a cell address (such as `T1` or `R1C1`), which Excel does not allow."))

    ref.stop.row_number == ref.start.row_number &&
        throw(XLSXError("Table `ref` must span at least two rows (a header row plus at least one data row) — Excel does not support header-only tables. Got `$ref`."))
    return name
end

# Write the part of a new table, checked already, and link it to `sheet`.
function _add_table_part!(sheet::Worksheet, ref::CellRange, name::String, columns::Vector{String},
                          style::Union{AbstractString,Nothing}, has_totals_row::Bool)
    xf = get_xlsxfile(sheet)
    style_info = isnothing(style) ? nothing : TableStyleInfo(style, false, false, true, false)

    id = next_table_id!(get_workbook(sheet))
    table_doc  = build_table_xml(id, name, name, ref, columns, has_totals_row, style_info)
    table_path = new_table_filename(xf)
    xf.data[table_path]  = table_doc
    xf.files[table_path] = true
    add_override!(xf, "/$table_path", MIME_TABLE)

    _attach_table_part!(xf, sheet, table_path)

    sheet.tables_cache = nothing
    return nothing
end

addtable!(sheet::Worksheet, ref::AbstractString; kw...) = addtable!(sheet, CellRange(ref); kw...)

# A name for a new table: `name`, checked unused (Excel compares table names
# ignoring case, and they share a namespace with defined names), or for `""` the
# first free `TableN`.
function _new_table_name(wb::Workbook, sheet::Worksheet, name::AbstractString)::String
    existing_names = Set{String}()
    for ws in wb.sheets
        is_chartsheet(wb, ws.name) && continue
        for t in tables(ws)
            push!(existing_names, uppercase(t.name))
        end
    end

    if name == ""
        i = 1
        candidate = "Table$i"
        while uppercase(candidate) ∈ existing_names ||
              is_workbook_defined_name(wb, candidate) || is_worksheet_defined_name(sheet, candidate)
            i += 1
            candidate = "Table$i"
        end
        name = candidate
    else
        uppercase(name) ∈ existing_names && throw(XLSXError("Table name `$name` is already in use."))
        (is_workbook_defined_name(wb, name) || is_worksheet_defined_name(sheet, name)) &&
            throw(XLSXError("Table name `$name` collides with an existing defined name."))
    end

    return name
end

# The column names of a new table, from the header row of `ref`: each must be
# given, and unique ignoring case, as in Excel.
function _table_header_columns(sheet::Worksheet, ref::CellRange)::Vector{String}
    header_row = ref.start.row_number
    columns = String[]
    for c in column_number(ref.start):column_number(ref.stop)
        v = getdata(sheet, CellRef(header_row, c))
        (ismissing(v) || v == "") &&
            throw(XLSXError("Table header cell $(CellRef(header_row, c)) is empty; every column needs a header value before calling `addtable!`."))
        push!(columns, string(v))
    end
    _check_table_columns(columns)
    return columns
end

function _check_table_columns(columns::Vector{String})
    any(isempty, columns) &&
        throw(XLSXError("Table header row has an empty column name; every column of a table needs one."))
    length(unique(uppercase.(columns))) != length(columns) &&   # case-insensitive, as in Excel
        throw(XLSXError("Table header row contains duplicate column names."))
    return nothing
end

# The last row of `ref`, about to be a totals row, must be empty, as `settotals!`
# requires. Excel refuses a file whose totals row holds a plain value, and content
# there is most often a data row (e.g. written by `writetable!`) that the caller
# forgot to exclude from `ref`.
function _check_totals_row_empty(sheet::Worksheet, ref::CellRange, name::AbstractString, columns::Vector{String})
    totals_row_num = ref.stop.row_number
    nonempty_cols = String[]
    for c in column_number(ref.start):column_number(ref.stop)
        v = getdata(sheet, CellRef(totals_row_num, c))
        (!ismissing(v) && v != "") && push!(nonempty_cols, columns[c - column_number(ref.start) + 1])
    end
    isempty(nonempty_cols) ||
        throw(XLSXError("Table `$name`: cannot mark the last row of `ref` (row $totals_row_num) as the " *
              "totals row, because it already has content in column(s) $(join(nonempty_cols, ", ")). " *
              "If that row is table data (e.g. written by `writetable!`), exclude it from `ref` " *
              "or pass `has_totals_row=false`; otherwise clear it first, then use `settotals!` " *
              "to set totals."))
    return nothing
end

# Link a new table part at `table_path` to `sheet`: a relationship, and a
# `<tablePart>` in the sheet's `<tableParts>`.
function _attach_table_part!(xf::XLSXFile, sheet::Worksheet, table_path::String)
    wb = get_workbook(sheet)
    sheet_path = get_relationship_target_by_id("xl", wb, sheet.relationship_id)
    sheet_dir, _ = rsplit(sheet_path, "/"; limit=2)
    rels_path, rels_root = get_or_create_worksheet_rels!(xf, sheet_path)

    rid = new_relationship_id(rels_root)
    pfx_rels = get_prefix(rels_path, xf)
    push!(rels_root, XML.Element(prefixed_tag(pfx_rels, "Relationship");
        Id=rid, Type=REL_TABLE, Target=make_relative_target(sheet_dir, table_path)))

    sheet_doc  = get_xml_data(xf, sheet_path)
    sheet_root = xml_root_element(sheet_doc)
    pfx   = get_prefix(sheet)
    pfx_c = pfx == "" ? "" : "$(pfx):"

    tp_container_els = elements_with_tag(sheet_root, "tableParts")
    if isempty(tp_container_els)
        tp_container = XML.Element("$(pfx_c)tableParts"; count="1")
        tp_node = XML.Element("$(pfx_c)tablePart")
        tp_node["r:id"] = rid
        push!(tp_container, tp_node)
        children = XML.children(sheet_root)
        ext_idx = findfirst(c -> localname(c) == "extLst", children)
        isnothing(ext_idx) ? push!(sheet_root, tp_container) : insert!(children, ext_idx, tp_container)
    else
        tp_container = tp_container_els[1]
        tp_node = XML.Element("$(pfx_c)tablePart")
        tp_node["r:id"] = rid
        push!(tp_container, tp_node)
        tp_container["count"] = string(length(XML.children(tp_container)))
    end

    return nothing
end

# Excel drops ("repairs") a table that overlaps a merged range or another table,
# so refuse to create one. Nor may a table overlap the sheet's AutoFilter.
function _check_table_overlap(sheet::Worksheet, ref::CellRange, name::AbstractString)
    for m in something(getMergedCells(sheet), CellRange[])
        intersects(ref, m) &&
            throw(XLSXError("Table `$name` (`$ref`) would overlap merged range `$m`, which Excel does not allow. Unmerge the range or place the table elsewhere."))
    end
    for t in tables(sheet)
        intersects(ref, t.ref) &&
            throw(XLSXError("Table `$name` (`$ref`) would overlap table `$(t.name)` (`$(t.ref)`), which Excel does not allow."))
    end
    _check_table_autofilter(sheet, ref, name)
    return nothing
end

# A table has its own filter, so it may not overlap the sheet's (see `setAutoFilter`).
function _check_table_autofilter(sheet::Worksheet, ref::CellRange, name::AbstractString)
    af = getAutoFilter(sheet)
    !isnothing(af) && intersects(ref, af) &&
        throw(XLSXError("Table `$name` (`$ref`) would overlap the sheet's AutoFilter (`$af`). Remove it first with `removeAutoFilter`."))
    return nothing
end

# `Table[Column]` as a formula writes it: `[`, `]`, `#` and `'` in a column name are
# escaped with `'`, so `Price [USD]` gives `Sales[Price '[USD']]`.
_structured_ref(table::AbstractString, column::AbstractString) =
    table * "[" * replace(column, r"(['\[\]#])" => s"'\1") * "]"

"""
    deletetable!(t::Table)
    deletetable!(sheet::Worksheet, name::AbstractString)
    deletetable!(sheet::Worksheet, id::Integer)

Delete the given Excel Table, provided directly as `t` or by fron `sheet` by 
its `name` or workbook-scoped numeric `id`.

This removes the table *object* only: its `xl/tables/tableN.xml` part, its
worksheet-level relationship, its `<tablePart>` entry, and its
`[Content_Types].xml` override. It does **not** clear, delete, or modify any
of the underlying cell data — the header row, data rows, and any totals row
are left completely untouched, still holding whatever values they held while
the table existed.

This mirrors what Excel itself does when you use **Table Design → Convert to
Range** (or right-click → Table → Convert to Range): the table's structure
(banding, filter dropdowns, structured references, and the Table object
itself) is removed, but the cells and their values remain in place as an
ordinary range. There is no single-step Excel operation that removes a table
and its data together; if you want the data gone too, clear or delete those
cells yourself as a separate step, e.g.:

```julia
julia> t = XLSX.table(sheet, "MyTable")

julia> XLSX.deletetable!(sheet, "MyTable")

julia> sheet[t.ref] = missing   # optional: also clear the data
```

Other tables on the same sheet, and tables on other sheets, are unaffected.

See also [`XLSX.addtable!`](@ref), [`XLSX.tables`](@ref), [`XLSX.table`](@ref).
"""
deletetable!(t::Table)                                = _deletetable!(table(t.sheet, t.name))
deletetable!(sheet::Worksheet, name::AbstractString)  = _deletetable!(table(sheet, name))
deletetable!(sheet::Worksheet, id::Integer)           = _deletetable!(table(sheet, id))

function _deletetable!(t::Table)
    sheet = t.sheet
    name = t.name
    xf = get_xlsxfile(sheet)
    _require_writable(xf, "delete a table")

    t=table(sheet, name)  # throws KeyError if not found — fail fast

    # Before removing the Table: totals-row formulas refer to the Table by name
    # (e.g. SUBTOTAL(109,Sales[amount])), which becomes #REF! once it's gone.
    # Excel rewrites such references to a static, sheet-qualified absolute range
    # (SUBTOTAL(109,Sheet1!$B$2:$B$5)) when converting a Table to a range; do the
    # same here. Only the structured reference is substituted — the surrounding
    # formula is left alone.
    if t.has_totals_row
        col0 = _col_start(t)
        totals_row = t.ref.stop.row_number
        first_data = _first_data_row(t)
        last_data  = _last_data_row(t)
        qsheet = quoteit(sheet.name)

        for (idx, col_name) in enumerate(t.columns)
            cell_ref = CellRef(totals_row, col0 + idx - 1)
            c = getcell(sheet, cell_ref)
            (c isa EmptyCell || !c.formula) && continue

            f = getFormula(sheet, cell_ref)
            isnothing(f) && continue

            col_letter = encode_column_number(col0 + idx - 1)
            static_range = "$(qsheet)!\$$(col_letter)\$$(first_data):\$$(col_letter)\$$(last_data)"

            # substitute every column reference of this Table, not just this
            # column's own — a custom formula may reference several.
            new_f = f
            for (j, other) in enumerate(t.columns)
                other_letter = encode_column_number(col0 + j - 1)
                other_range = "$(qsheet)!\$$(other_letter)\$$(first_data):\$$(other_letter)\$$(last_data)"
                new_f = replace(new_f, _structured_ref(name, other) => other_range)
            end

            setFormula(sheet, cell_ref; val=lstrip(new_f, '='))
        end
    end

    target_rid, target_path = _find_table_part(sheet, name)

    sheet_path = get_relationship_target_by_id("xl", get_workbook(sheet), sheet.relationship_id)
    sheet_dir, sheet_file = rsplit(sheet_path, "/"; limit=2)
    rels_path  = "$sheet_dir/_rels/$sheet_file.rels"
    rels_root  = xml_root_element(xf.data[rels_path])
    sheet_root = xml_root_element(get_xml_data(xf, sheet_path))

    # remove <tablePart>, shrink/drop <tableParts>
    tp_container = elements_with_tag(sheet_root, "tableParts")[1]
    tp_children = XML.children(tp_container)
    tp_idx = findfirst(tp -> XML.nodetype(tp) === XML.Element && get_attr(tp, "r:id") == target_rid, tp_children)
    isnothing(tp_idx) && throw(XLSXError("Internal error: <tablePart> for `$name` not found."))
    deleteat!(tp_children, tp_idx)

    if isempty(tp_children)
        deleteat!(XML.children(sheet_root), findfirst(c -> localname(c) == "tableParts", XML.children(sheet_root)))
    else
        tp_container["count"] = string(length(tp_children))
    end

    # remove the relationship entry itself
    rel_children = XML.children(rels_root)
    rel_idx = findfirst(r -> XML.nodetype(r) === XML.Element && get_attr(r, "Id") == target_rid, rel_children)
    isnothing(rel_idx) && throw(XLSXError("Internal error: relationship `$target_rid` not found in $rels_path."))
    deleteat!(rel_children, rel_idx)

    delete_part_and_orphans!(xf, target_path)  # handles the table part + its Content_Types override

    sheet.tables_cache = nothing
    return nothing
end

"""
    settotals!(t::Table, settings::Pair...)
    settotals!(t::Table; kwargs...)
    settotals!(sheet::Worksheet, name::AbstractString, settings::Pair...)
    settotals!(sheet::Worksheet, name::AbstractString; kwargs...)
    settotals!(sheet::Worksheet, id::Integer, settings::Pair...)
    settotals!(sheet::Worksheet, id::Integer; kwargs...)

Add or update the totals row for the Excel Table `t` or the table in `sheet` 
with the specified `name` or workbook-scoped numeric `id`.

Each element of `settings` is `"ColumnName" => value` (or, in the kwarg form,
`ColumnName=value` for identifier-safe column names), where `value` is one of:

- a `Symbol` naming a built-in totals function — `:sum`, `:average`, `:count`,
  `:countnums`, `:max`, `:min`, `:stddev`, `:var` or `:none`. Writes both the
  `totalsRowFunction` attribute and an actual `SUBTOTAL(...)` formula into
  the totals row cell for that column; Excel performs the calculation from
  this formula exactly as it would for any other `SUBTOTAL` formula.
- a `(:custom, formula::AbstractString)` tuple, for a custom totals function
  — Excel's "More Functions..." option. `formula` is written verbatim into
  the cell (it need not be `SUBTOTAL`-based at all), and
  `totalsRowFunction="custom"` is set. XLSX.jl does not validate or evaluate
  the formula; Excel computes it on open, same as any other formula cell.
- an `AbstractString`, written as a plain text label in that column's totals
  row cell (e.g. `"Grand Total"`), with no function or formula attached.

Columns not mentioned in `settings` are left untouched: if the table already
has a totals row, their existing totals content (function, label, or blank)
is preserved as-is. To *remove* a column's totals, pass `:none` explicitly:

```julia
julia> XLSX.settotals!(s, "Sales", "margin" => :none)   # margin's totals cell cleared
```

The totals row itself remains, even if every column's totals is cleared — an empty
totals row is valid, and Excel displays it.

If the table does not already have a totals row, one is added by extending
the table by one row — the row immediately following its current last row.
That row must be completely empty; an `XLSXError` is thrown otherwise (clear
it first, or use [`XLSX.addtable!`](@ref) with `has_totals_row=true` if that
row was always meant to be part of the table).

As with every formula-writing path in XLSX.jl, no cached value is written
alongside a totals-row formula — Excel recalculates it on open
(`update_workbook_xml!` forces `fullCalcOnLoad="1"`). Formulas are written
via [`XLSX.setFormula`](@ref), which fully replaces the cell's formula/value
while preserving its existing style, so calling `settotals!` again on a
column that already has totals content (function, custom formula, or label)
cleanly replaces it.

# Examples
```julia
julia> XLSX.settotals!(sheet, "Sales",
           "Revenue" => :sum,
           "Notes"   => "Grand Total",
           "Margin"  => (:custom, "SUBTOTAL(109,Sales[Revenue])-SUBTOTAL(109,Sales[Cost])"),
       )

julia> XLSX.settotals!(sheet, "Sales"; Revenue=:sum, Margin=:average)

julia> XLSX.settotals!(sheet, tbl.id, "Revenue" => :sum)  # by workbook-scoped table id

julia> # Add a totals row to a table that doesn't have one yet — the row
       # immediately following the table's current last row must be empty.
       XLSX.settotals!(sheet, "Sales", "Revenue" => :sum)

julia> # Update just one column's totals function on a table that already
       # has a totals row — other columns' existing totals are untouched.
       XLSX.settotals!(sheet, "Sales", "Revenue" => :max)

julia> # A custom formula MUST aggregate each column reference itself
       # (e.g. via SUBTOTAL or SUM) — a bare `Sales[Revenue]` in a totals
       # cell is rewritten by Excel to a "this row" reference, which has
       # no valid row to intersect against in the totals row and raises
       # #VALUE!.
       XLSX.settotals!(sheet, "Sales",
           "Margin" => (:custom, "SUBTOTAL(109,Sales[Revenue])-SUBTOTAL(109,Sales[Cost])"),
       )
```

See also [`XLSX.addtable!`](@ref), [`XLSX.tables`](@ref), [`XLSX.table`](@ref).
"""
settotals!(t::Table, settings::Pair...)                                = _settotals!(table(t.sheet, t.name), settings...)
settotals!(t::Table; kwargs...) =
    _settotals!(t, (String(k) => v for (k, v) in kwargs)...)

settotals!(sheet::Worksheet, name::AbstractString, settings::Pair...)  = _settotals!(table(sheet, name), settings...)
settotals!(sheet::Worksheet, name::AbstractString; kwargs...) =
    _settotals!(table(sheet, name), (String(k) => v for (k, v) in kwargs)...)

settotals!(sheet::Worksheet, id::Integer, settings::Pair...)           = _settotals!(table(sheet, id), settings...)
settotals!(sheet::Worksheet, id::Integer; kwargs...) =
    _settotals!(table(sheet, id), (String(k) => v for (k, v) in kwargs)...)


function _settotals!(t::Table, settings::Pair...)
    sheet = t.sheet
    name = t.name
    xf = get_xlsxfile(sheet)
    _require_writable(xf, "set a totals row")

    sheet = t.sheet 
    name = t.name

    for (col, _) in settings
        col ∈ t.columns || throw(XLSXError("Column `$col` not found in table `$name`. Available columns: $(join(t.columns, ", "))."))
    end

    table_path = _table_part_path(sheet, name)
    table_doc  = get_xml_data(xf, table_path)
    table_root = xml_root_element(table_doc)

    old_ref = t.ref
    local totals_row_num::Int

    if !t.has_totals_row
        new_last_row = old_ref.stop.row_number + 1
        new_last_row > EXCEL_MAX_ROWS &&
            throw(XLSXError("Cannot add a totals row to table `$name`: table already reaches the last worksheet row."))
        _check_table_autofilter(sheet, CellRange(old_ref.start, CellRef(new_last_row, old_ref.stop.column_number)), name)

        for c in column_number(old_ref.start):column_number(old_ref.stop)
            v = getdata(sheet, CellRef(new_last_row, c))
            (!ismissing(v) && v != "") &&
                throw(XLSXError("Cannot add a totals row to table `$name`: row $new_last_row already has content at $(CellRef(new_last_row, c)). Clear that row first, or use `addtable!` with `has_totals_row=true` if this row was always meant to be part of the table."))
        end

        new_ref = CellRange(old_ref.start, CellRef(new_last_row, old_ref.stop.column_number))
        table_root["ref"] = string(new_ref)
        table_root["totalsRowShown"] = "1"
        table_root["totalsRowCount"] = "1"

        totals_row_num = new_last_row
    else
        totals_row_num = old_ref.stop.row_number
    end

    i, j = get_idces(table_doc, "table", "tableColumns")
    column_nodes = collect(xml_elements(table_doc[i][j]))

    for (col_name, value) in settings
        col_idx  = findfirst(==(col_name), t.columns)
        col_node = column_nodes[col_idx]
        cell_ref = CellRef(totals_row_num, column_number(old_ref.start) + col_idx - 1)

        remove_attr!(col_node, "totalsRowFunction")
        remove_attr!(col_node, "totalsRowLabel")

        if value === :none
            # Clear this column's totals setting entirely. Both attributes were
            # already removed above; just blank the cell.
            sheet[cell_ref] = missing

        elseif value isa Symbol
            haskey(TOTALS_ROW_FUNCTIONS, value) ||
                throw(XLSXError("Unknown totals function `:$value`. Supported: $(join(sort(string.(keys(TOTALS_ROW_FUNCTIONS))), ", ")), or pass `(:custom, \"formula\")` for a custom function."))
            func_name, subtotal_code = TOTALS_ROW_FUNCTIONS[value]
            col_node["totalsRowFunction"] = func_name
            setFormula(sheet, cell_ref; val="SUBTOTAL($subtotal_code,$(_structured_ref(name, col_name)))")

        elseif value isa Tuple{Symbol,S} where {S<:AbstractString}
            fn, formula_str = value
            fn == :custom || throw(XLSXError("Tuple totals setting for column `$col_name` must be `(:custom, formula)`; got `(:$fn, ...)`."))
            col_node["totalsRowFunction"] = "custom"
            setFormula(sheet, cell_ref; val=formula_str)

        elseif value isa AbstractString
            col_node["totalsRowLabel"] = value
            sheet[cell_ref] = value

        else
            throw(XLSXError("Totals setting for column `$col_name` must be a Symbol (built-in function), a `(:custom, formula)` tuple, or a String (label); got $(typeof(value))."))
        end
    end

    sheet.tables_cache = nothing
    return table(sheet, name)
end

"""
    gettotals(t::Table) -> NamedTuple
    gettotals(sheet::Worksheet, name::AbstractString)
    gettotals(sheet::Worksheet, id::Integer)

Return the Excel Table's totals row as a `NamedTuple` keyed by column name. Provide the
Table `t` directly or specify the table in `sheet` by its `name` or workbook-scoped 
numeric `id`. Each entry is itself a `NamedTuple` with fields:

- `setting`: the column's totals setting — a `Symbol` naming a built-in function
  (`:sum`, `:average`, …), `:custom` for a custom formula, `:label` for a text label, or
  `:none` if the column has no totals setting.
- `formula`: the totals cell's formula as a `String`, or `nothing` if the cell holds no
  formula.
- `value`: the totals cell's current value. For a label column this is the label text.
  For a function or custom column this is `missing` unless the file was written by Excel
  (XLSX.jl does not evaluate formulas, so a totals cell it wrote holds no cached value
  until Excel recalculates and saves the file).

Returns an empty `NamedTuple` if the Table has no totals row.

# Example
```julia
julia> tot = XLSX.gettotals(t)
(region = (setting = :label, formula = nothing, value = "Total"), revenue = (setting = :sum, formula = "=SUBTOTAL(109,Sales[revenue])", value = 3400), margin = (setting = :average, formula = "=SUBTOTAL(101,Sales[margin])", value = 243.33))

julia> tot.revenue.setting
:sum

julia> tot.revenue.value
3400
```

`gettotals` requires the `XLSXFile` to have `enable_cache=true` (the default). It throws otherwise.

See also [`XLSX.settotals!`](@ref), [`XLSX.removetotals!`](@ref), [`XLSX.table`](@ref).
"""
gettotals(t::Table)                                = _gettotals(table(t.sheet, t.name))
gettotals(sheet::Worksheet, name::AbstractString)  = _gettotals(table(sheet, name))
gettotals(sheet::Worksheet, id::Integer)           = _gettotals(table(sheet, id))

function _gettotals(t::Table)
    t.has_totals_row || return NamedTuple()

    sheet = t.sheet
    totals_row = t.ref.stop.row_number
    col0 = _col_start(t)
    ncols = length(t.columns)

    # Read the totals cells FIRST, while the worksheet XML is still in
    # whatever state the cache machinery expects. `_table_part_path` below
    # calls `get_xml_data` on the sheet, which permanently promotes a
    # read-only file's raw XML string to a parsed node and breaks the lazy
    # per-sheet cache fill that `getcell`/`getFormula` rely on.
    formulas = Vector{Union{Nothing,String}}(undef, ncols)
    values   = Vector{Any}(undef, ncols)
    for idx in 1:ncols
        cell_ref = CellRef(totals_row, col0 + idx - 1)
        raw_f = getFormula(sheet, cell_ref)
        formulas[idx] = (isnothing(raw_f) || isempty(raw_f)) ? nothing : raw_f
        values[idx] = getdata(sheet, cell_ref)
    end

    # Now the table part, whose metadata gives each column's setting.
    table_doc = get_xml_data(get_xlsxfile(sheet), _table_part_path(sheet, t.name))
    i, j = get_idces(table_doc, "table", "tableColumns")
    column_nodes = collect(xml_elements(table_doc[i][j]))

    names = Symbol[]
    entries = Any[]
    for (idx, col_name) in enumerate(t.columns)
        node  = column_nodes[idx]
        func  = get_attr(node, "totalsRowFunction", "")
        label = get_attr(node, "totalsRowLabel", "")

        setting = if func == "custom"
            :custom
        elseif func != ""
            get(TOTALS_FUNCTION_BY_NAME, func, Symbol(func))
        elseif label != ""
            :label
        else
            :none
        end

        push!(names, Symbol(col_name))
        push!(entries, (setting=setting, formula=formulas[idx], value=values[idx]))
    end

    return NamedTuple{Tuple(names)}(Tuple(entries))
end

"""
    removetotals!(t::Table) -> Table
    removetotals!(sheet::Worksheet, name::AbstractString) -> Table
    removetotals!(sheet::Worksheet, id::Integer) -> Table

Remove the totals row from the Excel Table given directly as `t` or by `name` or `id` on 
`sheet`, shrinking the Table by one row. Every column's totals function or label is 
cleared, and the cells that held the totals row are emptied.

Does nothing if the Table has no totals row.

See also [`XLSX.settotals!`](@ref), [`XLSX.addtable!`](@ref).
"""
removetotals!(t::Table)                                = _removetotals!(table(t.sheet, t.name))
removetotals!(sheet::Worksheet, name::AbstractString)  = _removetotals!(table(sheet, name))
removetotals!(sheet::Worksheet, id::Integer)           = _removetotals!(table(sheet, id))

function _removetotals!(t::Table)
    sheet = t.sheet
    name = t.name
    xf = get_xlsxfile(sheet)
    _require_writable(xf, "remove a totals row")

    t.has_totals_row || return t

    table_path = _table_part_path(sheet, name)
    table_doc  = get_xml_data(xf, table_path)
    table_root = xml_root_element(table_doc)

    totals_row = t.ref.stop.row_number
    col0 = _col_start(t)

    # clear the cells that held the totals row
    for c in col0:(col0 + length(t.columns) - 1)
        sheet[CellRef(totals_row, c)] = missing
    end

    # drop per-column totals metadata
    i, j = get_idces(table_doc, "table", "tableColumns")
    for col_node in xml_elements(table_doc[i][j])
        localname(col_node) == "tableColumn" || continue
        remove_attr!(col_node, "totalsRowFunction")
        remove_attr!(col_node, "totalsRowLabel")
    end

    # shrink ref and drop the totals-row flags
    new_ref = CellRange(t.ref.start, CellRef(totals_row - 1, t.ref.stop.column_number))
    table_root["ref"] = string(new_ref)
    remove_attr!(table_root, "totalsRowShown")
    remove_attr!(table_root, "totalsRowCount")

    sheet.tables_cache = nothing
    return table(sheet, name)
end

"""
    gettable(t::Table; [infer_eltypes], [normalizenames], [missing_strings]) -> DataTable

Returns data from an Excel Table `t` (as returned by [`XLSX.table`](@ref)) as
a struct `XLSX.DataTable`, which can be passed directly to any function that
accepts `Tables.jl` data (e.g. `DataFrame` from package `DataFrames.jl`).

!!! note "Two `gettable` methods"

    Different from `XLSX.gettable(sheet, ...)`, which infers a
    table's row/column bounds heuristically from cell content on a plain
    `Worksheet`. Here, `t.ref` is authoritative, so there is no
    `columns`/`first_row`/`header`/`stop_in_empty_row`/`stop_in_row_function`/
    `keep_empty_rows` equivalent — the header row and totals row (if any)
    are always excluded, and any blank row within `t.ref` is returned as
    ordinary data.

Use `normalizenames=true` to normalize column names to valid Julia
identifiers.

Use `missing_strings` to specify strings that should be interpreted as
`missing` values in the resulting table. `missing_strings` can be a single
string or a vector of strings. The default value is `missing_strings=nothing`.

Use `infer_eltypes=true` (the default) to have each column narrowed to its
own concrete type (e.g. `Vector{Float64}` rather than `Vector{Any}`), the
same narrowing [`XLSX.eachtablerow`](@ref)/`Tables.columns` apply. Set
`infer_eltypes=false` to skip narrowing and leave every column as `Any`.

# Example
```julia
julia> using DataFrames

julia> t = XLSX.table(sheet, "Sales")

julia> df = DataFrame(XLSX.gettable(t))
```

See also: [`XLSX.table`](@ref), [`XLSX.tables`](@ref), [`XLSX.eachtablerow`](@ref), [`XLSX.readtable`](@ref).
"""
function gettable(t::Table;
    infer_eltypes::Bool=true,
    normalizenames::Bool=false,
    missing_strings::Union{AbstractString,AbstractVector{<:AbstractString},Nothing}=nothing
)::DataTable
    t = _resolve(t)
    missing_set = missing_strings === nothing ? nothing :
                  missing_strings isa AbstractString ? Set([missing_strings]) : Set(missing_strings)

    m = _table_data_matrix(t)

    data = Vector{Any}(undef, length(t.columns))
    for i in eachindex(t.columns)
        col = Vector{Any}(m[:, i])
        if !isnothing(missing_set)
            col = Any[(x isa AbstractString && x in missing_set) ? missing : x for x in col]
        end
        data[i] = infer_eltypes ? typed_column(col) : col
    end

    column_labels = normalizenames ? normalizename.(t.columns) : Symbol.(t.columns)

    return DataTable(data, column_labels)
end

"""
    appendtable!(t::Table, data; kw...) -> Table
    appendtable!(sheet::Worksheet, name::AbstractString, data; [check_empty]) -> Table
    appendtable!(sheet::Worksheet, id::Integer, data; kw...) -> Table


Append rows to the existing Excel Table. Provide the Table `t` directly or 
specify the table in `sheet` by its `name` or workbook-scoped numeric `id`.

`data` may be any `Tables.jl`-compatible source (e.g. an `XLSX.DataTable` or a
`DataFrame`), an `AbstractMatrix`, or a vector of row vectors/tuples.

If `data` exposes column names, columns are matched **by name** and reordered
to the table's own column order; a source missing any of the table's columns,
or carrying any column the table doesn't have, is an error. Sources without
column names (matrices, vectors of tuples/vectors) are matched
**positionally**, so their column order must match the table's.

If the table has a totals row, it moves down to remain the last row of the
table, and its content is regenerated from the table's own per-column
totals settings. Functions, custom formulas and labels are all preserved 
but values are reset to `missing`.

The rows immediately below the table must be empty; an `XLSXError` is thrown
otherwise. Pass `check_empty=false` to overwrite whatever is there.

See also [`XLSX.addtable!`](@ref), [`XLSX.settotals!`](@ref).
"""
appendtable!(t::Table, data; kw...)                                = _appendtable!(table(t.sheet, t.name), data; kw...)
appendtable!(sheet::Worksheet, name::AbstractString, data; kw...)  = _appendtable!(table(sheet, name), data; kw...)
appendtable!(sheet::Worksheet, id::Integer, data; kw...)           = _appendtable!(table(sheet, id), data; kw...)

function _appendtable!(t::Table, data; check_empty::Bool=true, kw... )
    sheet = t.sheet 
    name = t.name
    xf = get_xlsxfile(sheet)
    _require_writable(xf, "append to a table")

    ncols = length(t.columns)
    rows = _normalize_append_rows(data, t, name)
    n = length(rows)
    n == 0 && return t

    col0 = _col_start(t)
    old_stop = t.ref.stop.row_number
    new_stop = old_stop + n
    new_stop > EXCEL_MAX_ROWS &&
        throw(XLSXError("Appending $n rows to table `$name` would exceed Excel's row limit."))
    _check_table_autofilter(sheet, CellRange(t.ref.start, CellRef(new_stop, t.ref.stop.column_number)), name)

    # capture totals settings before we overwrite the old totals row
    totals_settings = parse_totals_settings(sheet, t)

    if check_empty
        for r in (old_stop + 1):new_stop, c in col0:(col0 + ncols - 1)
            v = getdata(sheet, CellRef(r, c))
            (!ismissing(v) && v != "") && throw(XLSXError(
                "Cannot append to table `$name`: cell $(CellRef(r, c)) is not empty. " *
                "Clear the rows below the table, or pass `check_empty=false` to overwrite."))
        end
    end

    # first appended row lands on the old totals row position when one exists,
    # otherwise immediately below the current last data row
    first_new = t.has_totals_row ? old_stop : old_stop + 1
    for (ri, row) in enumerate(rows)
        for ci in 1:ncols
            sheet[CellRef(first_new + ri - 1, col0 + ci - 1)] = row[ci]
        end
    end

    # clear the new totals row position before settotals! repopulates it
    if t.has_totals_row
        for c in col0:(col0 + ncols - 1)
            sheet[CellRef(new_stop, c)] = missing
        end
    end

    # extend ref and autoFilter in the table part
    table_doc  = get_xml_data(xf, _table_part_path(sheet, name))
    table_root = xml_root_element(table_doc)
    new_ref = CellRange(t.ref.start, CellRef(new_stop, t.ref.stop.column_number))
    table_root["ref"] = string(new_ref)

    af = elements_with_tag(table_root, "autoFilter")
    if !isempty(af)
        af_stop = t.has_totals_row ? new_stop - 1 : new_stop
        af[1]["ref"] = string(CellRange(t.ref.start, CellRef(af_stop, t.ref.stop.column_number)))
    end

    sheet.tables_cache = nothing

    # replay the totals row at its new position
    isempty(totals_settings) || settotals!(sheet, name, totals_settings...)

    sheet.tables_cache = nothing
    return table(sheet, name)
end


"""
Normalize whatever was passed as `data` into a vector of row-value vectors,
each of length `length(t.columns)`, ordered to match the table's columns.

For `Tables.jl` sources that expose column names (a `DataTable`, `DataFrame`,
etc.), columns are matched **by name** and reordered into the table's own
column order; both missing and extra columns are errors. Matrices, vectors of
tuples and vectors of vectors carry no names, so those are matched
**positionally**.
"""
function _normalize_append_rows(data, t::Table, name::AbstractString)
    ncols = length(t.columns)

    rows = if data isa AbstractMatrix
        [collect(data[r, :]) for r in axes(data, 1)]

    elseif data isa AbstractVector && (isempty(data) || first(data) isa Union{AbstractVector,Tuple})
        [collect(r) for r in data]

    else
        sch = Tables.schema(data)
        src_names = if !isnothing(sch) && !isnothing(sch.names)
            sch.names
        else
            # Some Tables.jl sources expose names via columnnames without
            # implementing schema (e.g. anything relying on the column-access
            # interface alone). Ask a source that offers column access before
            # giving up and going positional; asking any other is an error.
            Tables.columnaccess(data) ? Tables.columnnames(Tables.columns(data)) : nothing
        end
        if isnothing(src_names)
            # Unnamed Tables.jl source — fall back to positional matching.
            [collect(Tables.getcolumn(r, i) for i in 1:ncols) for r in Tables.rows(data)]
        else
            src = collect(Symbol.(src_names))
            tbl = Symbol.(t.columns)

            missing_cols = setdiff(tbl, src)
            isempty(missing_cols) || throw(XLSXError(
                "Source is missing column(s) $(join(missing_cols, ", ")) required by table `$name`. " *
                "Table columns are: $(join(t.columns, ", "))."))

            extra_cols = setdiff(src, tbl)
            isempty(extra_cols) || throw(XLSXError(
                "Source has column(s) $(join(extra_cols, ", ")) that table `$name` does not have. " *
                "Table columns are: $(join(t.columns, ", "))."))

            # Match by name, emitting values in the table's column order.
            [collect(Tables.getcolumn(r, nm) for nm in tbl) for r in Tables.rows(data)]
        end
    end

    for (ri, r) in enumerate(rows)
        length(r) == ncols || throw(XLSXError(
            "Row $ri has $(length(r)) values but table `$name` has $ncols columns."))
    end

    return rows
end

"""
    gettablerange(t::Table) -> SheetCellRange
    gettablerange(t::Table, column; header=false) -> Union{SheetCellRange,SheetCellRef}
    gettablerange(t::Table, first, last) -> SheetCellRange

The cells of a table's data rows as a sheet-qualified range, excluding the header
row and any totals row: the whole table, one column, or the adjacent columns from
`first` to `last`. A column is given by name or by its 1-based position in the
table. With `header = true`, the one-column form returns that column's header cell.

The range is the table's extent when this is called; it does not follow the table
if rows are added later. The table is looked up afresh, so `t` may be one obtained
before the table was changed.

These are the ranges to pass to a chart:

    t = table(ws, "Sales")
    addSeries(c, gettablerange(t, "Amount"); categories = gettablerange(t, "Region"),
              name_ref = gettablerange(t, "Amount"; header = true))
"""
function gettablerange(t::Table)
    t = table(t.sheet, t.name)
    top, bottom = _table_data_rows(t)
    return SheetCellRange(t.sheet.name, CellRange(CellRef(top, column_number(t.ref.start)),
                                                  CellRef(bottom, column_number(t.ref.stop))))
end

function gettablerange(t::Table, column; header::Bool = false)
    t   = table(t.sheet, t.name)
    col = column_number(t.ref.start) + _table_column_index(t, column) - 1
     if header
        t.has_header_row || throw(XLSXError(
            "Table `$(t.name)` has its header row hidden, so column `$column` has no header cell on the sheet."))
        return SheetCellRef(t.sheet.name, CellRef(row_number(t.ref.start), col))
    end
    top, bottom = _table_data_rows(t)
    return SheetCellRange(t.sheet.name, CellRange(CellRef(top, col), CellRef(bottom, col)))
end

function gettablerange(t::Table, first, last)
    t    = table(t.sheet, t.name)
    a, b = minmax(_table_column_index(t, first), _table_column_index(t, last))
    c0   = column_number(t.ref.start) - 1
    top, bottom = _table_data_rows(t)
    return SheetCellRange(t.sheet.name, CellRange(CellRef(top, c0 + a), CellRef(bottom, c0 + b)))
end

# First and last data rows: below the header row, above any totals row.
function _table_data_rows(t::Table)
    top, bottom = _first_data_row(t), _last_data_row(t)
    bottom >= top || throw(XLSXError("Table `$(t.name)` has no data rows."))
    return top, bottom
end

function _table_column_index(t::Table, column)::Int
    if column isa Integer
        1 <= column <= length(t.columns) || throw(XLSXError(
            "Table `$(t.name)` has $(length(t.columns)) columns; asked for column $column."))
        return column
    end
    i = findfirst(==(String(column)), t.columns)
    isnothing(i) && throw(XLSXError(
        "Table `$(t.name)` has no column `$column`. Its columns are: " * join(t.columns, ", ") * "."))
    return i
end

