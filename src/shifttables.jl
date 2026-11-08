# ===========================================================================
# Moving a sheet's Excel Tables as rows or columns are inserted or deleted
#
# A table's range moves as a range does (see shiftrefs.jl): rows or columns
# inserted inside it are added to it, and ones deleted from it are taken out. As
# in Excel:
#
# - Columns inserted inside a table become table columns, named `Column1`,
#   `Column2`, ... (unused in the table) in its header row.
# - A table all of whose rows (columns) are deleted is deleted.
# - Deleting a table's header row, its totals row or all its data rows, without
#   the rest of it, is refused: a table must keep them. (Excel turns off the
#   totals row; here, use `removetotals!` first.)
#
# References to a deleted table, or to its deleted columns (`Sales[Amount]`),
# become #REF! in every formula, name, conditional format, validation and chart.
# ===========================================================================

# What edit `e` does to each table on `ws`: `:delete` or `:move`. Refuses an edit
# a table can't take, before anything changes.
function _plan_table_edits(ws::Worksheet, e::RowColEdit)::Vector{Tuple{Table,Symbol}}
    plan = Tuple{Table,Symbol}[]
    for t in tables(ws)
        lo, hi = e.axis === :row ? (row_number(t.ref.start), row_number(t.ref.stop)) :
                                   (column_number(t.ref.start), column_number(t.ref.stop))
        if e.count < 0
            last = _edit_last_deleted(e)
            if e.at <= lo && hi <= last
                push!(plan, (t, :delete))
                continue
            end
            if e.axis === :row && !(hi < e.at || lo > last)
                gone(r) = e.at <= r <= last
                t.has_header_row && gone(lo) &&
                    throw(XLSXError("Cannot delete row $lo: it is the header row of table `$(t.name)`. Delete the whole table's rows, or the table, instead."))
                t.has_totals_row && gone(hi) &&
                    throw(XLSXError("Cannot delete row $hi: it is the totals row of table `$(t.name)`. Remove it with `removetotals!` first."))
                first_data, last_data = lo + t.has_header_row, hi - t.has_totals_row
                all(gone, first_data:last_data) &&
                    throw(XLSXError("Cannot delete rows $first_data to $last_data: they are all the data rows of table `$(t.name)`, which must keep at least one."))
            end
        end
        push!(plan, (t, :move))
    end
    return plan
end

# Delete the tables edit `e` deletes whole, noting them in `dead` (see `_kill_structured_refs!`).
function _delete_tables!(plan, dead::Dict{String,Union{Nothing,Set{String}}})
    for (t, action) in plan
        action === :delete || continue
        dead[uppercase(t.name)] = nothing
        _deletetable!(t)
    end
    return nothing
end

# Move the other tables' parts for edit `e`, after the cells have moved. Columns
# deleted from a table are noted in `dead`.
function _move_tables!(ws::Worksheet, plan, e::RowColEdit, dead::Dict{String,Union{Nothing,Set{String}}})
    xf = get_xlsxfile(ws)
    for (t, action) in plan
        action === :move || continue
        root = xml_root_element(get_xml_data(xf, _table_part_path(ws, t.name)))
        e.axis === :column && _edit_table_columns!(ws, t, root, e, dead)
        _shift_attr!(root, "ref", e)
        for el in xml_elements(root)
            n = localname(el)
            n in ("autoFilter", "sortState") || continue
            if n == "autoFilter"   # the table's header and data rows: all but a totals row
                rng = CellRange(root["ref"])
                el["ref"] = string(CellRange(rng.start, CellRef(row_number(rng.stop) - t.has_totals_row, column_number(rng.stop))))
            else
                _shift_attr!(el, "ref", e)
            end
            for s in xml_elements(el)   # an autoFilter's sortState, and their sort conditions
                localname(s) in ("sortState", "sortCondition") && _shift_attr!(s, "ref", e)
                for c in xml_elements(s)
                    localname(c) == "sortCondition" && _shift_attr!(c, "ref", e)
                end
            end
        end
    end
    ws.tables_cache = nothing
    return nothing
end

# Add the columns inserted inside table `t`, or take out those deleted from it:
# its `<tableColumn>`s, and its filter columns, which count from its first column.
function _edit_table_columns!(ws::Worksheet, t::Table, root::XML.Node, e::RowColEdit, dead)
    left, right = column_number(t.ref.start), column_number(t.ref.stop)
    cols_el = first_element_with_tag(root, "tableColumns")
    isnothing(cols_el) && return nothing
    tcs = XML.Node[c for c in XML.children(cols_el) if XML.nodetype(c) == XML.Element]
    if e.count > 0
        left < e.at <= right || return nothing   # a column inserted inside the table
        j = e.at - left                           # how many table columns come before it
        names = Set(uppercase(get(c, "name", "")) for c in tcs)
        next_id = maximum(parse(Int, c["id"]) for c in tcs; init=0) + 1
        pfx = something(_ns_prefix(XML.tag(cols_el)), "")
        new = XML.Node[]
        k = 1
        for _ in 1:e.count
            while uppercase("Column$k") in names
                k += 1
            end
            name = "Column$k"
            push!(names, uppercase(name))
            push!(new, _xml_element(pfx * "tableColumn", ["id" => string(next_id), "name" => name]))
            t.has_header_row && (ws[CellRef(row_number(t.ref.start), e.at + length(new) - 1)] = name)
            next_id += 1
        end
        kids = XML.children(cols_el)
        pos = findall(c -> XML.nodetype(c) == XML.Element, kids)
        insert_at = j < length(pos) ? pos[j+1] : length(kids) + 1
        for (i, c) in enumerate(new)
            insert!(kids, insert_at + i - 1, c)
        end
        _shift_filter_columns!(root, i -> i >= j ? i + e.count : i)
    else
        last = _edit_last_deleted(e)
        gone = [left + i - 1 for i in eachindex(tcs) if e.at <= left + i - 1 <= last]
        isempty(gone) && return nothing
        names = Set(uppercase(get(tcs[c-left+1], "name", "")) for c in gone)
        union!(get!(() -> Set{String}(), dead, uppercase(t.name)), names)
        _filter_elements!(cols_el, c -> uppercase(get(c, "name", "")) ∉ names)
        before(i) = count(c -> c - left < i, gone)
        _shift_filter_columns!(root, i -> (left + i) in gone ? nothing : i - before(i))
    end
    cols_el["count"] = string(length(xml_elements(cols_el)))
    return nothing
end

# Renumber a table's filter columns by `f` (0-based colId in, new colId or `nothing` to drop out).
function _shift_filter_columns!(root::XML.Node, f::Function)
    af = first_element_with_tag(root, "autoFilter")
    isnothing(af) && return nothing
    _filter_elements!(af) do fc
        localname(fc) == "filterColumn" || return true
        i = f(parse(Int, fc["colId"]))
        isnothing(i) && return false
        fc["colId"] = string(i)
        return true
    end
    return nothing
end

# --- formulas in tables, and structured references

const _TABLE_FORMULA_ELEMENTS = ("calculatedColumnFormula", "totalsRowFormula")

# Apply `g` to the formulas held in every table part: calculated columns and totals.
function _rewrite_table_formulas!(g::Function, xf::XLSXFile)
    wb = get_workbook(xf)
    for s in wb.sheets
        is_chartsheet(wb, s.name) && continue
        for t in tables(s), c in xml_elements(first_element_with_tag(
                xml_root_element(get_xml_data(xf, _table_part_path(s, t.name))), "tableColumns"))
            for f in xml_elements(c)
                localname(f) in _TABLE_FORMULA_ELEMENTS || continue
                v = XML.is_simple_value(f)
                isnothing(v) && continue
                w = g(v, s.name)
                w == v || (f[end] = XML.Text(w))
            end
        end
    end
    return nothing
end

# `formula` with each structured reference to a dead table, or to a dead column of
# one, made #REF!. `dead` maps a table's name (upper case) to its dead columns (upper
# case), or `nothing` if the whole table is gone.
function _kill_structured_refs(formula::AbstractString, dead::Dict{String,Union{Nothing,Set{String}}})::String
    isempty(dead) && return String(formula)
    s = String(formula)
    out = IOBuffer()
    i = 1
    for m in eachmatch(_RX_STRUCTURED_REF, s)
        write(out, SubString(s, i, prevind(s, m.offset)))
        i = m.offset + ncodeunits(m.match)
        name = m[:table]
        if isnothing(name) || !haskey(dead, uppercase(name))   # a string, or another table
            write(out, m.match)
            continue
        end
        cols = dead[uppercase(name)]
        killed = isnothing(cols) || any(c -> uppercase(c) in cols, _structured_ref_columns(m[:spec]))
        write(out, killed ? "#REF!" : m.match)
    end
    write(out, SubString(s, i))
    return String(take!(out))
end

# A string literal, or a table name followed by its bracketed specifier.
const _RX_STRUCTURED_REF = Regex(
    "\"(?:[^\"]|\"\")*\"" *
    "|(?<![\\p{L}\\p{N}_.\\\\])(?<table>[\\p{L}_\\\\][\\p{L}\\p{N}_.]*)(?<spec>(?<br>\\[(?:[^\\[\\]']|'.|(?&br))*\\]))")

# The column names in a structured reference's specifier: `[Col]`, `[[#This Row],[Col]]`,
# `[[Col1]:[Col3]]`, with `'` escaping `[`, `]`, `#` and `'`. Special items (`#All`,
# `#Data`, `@`...) are not columns.
function _structured_ref_columns(spec::AbstractString)::Vector{String}
    out = String[]
    for m in eachmatch(r"\[((?:[^\[\]']|'.)*)\]", spec)
        raw = strip(m[1])
        startswith(raw, '@') && (raw = strip(raw[2:end]))
        (isempty(raw) || startswith(raw, '#')) && continue   # #All, #Data, ...; an escaped '# is a name
        push!(out, replace(raw, r"'(.)" => s"\1"))
    end
    return out
end

# Make every structured reference to what `dead` names #REF!, wherever formulas are.
function _kill_structured_refs!(xf::XLSXFile, dead::Dict{String,Union{Nothing,Set{String}}})
    isempty(dead) && return nothing
    wb = get_workbook(xf)
    kill(f) = _kill_structured_refs(f, dead)
    for f in values(wb.formulas)
        f isa ExplicitFormula && (f.formula = kill(f.formula))
    end
    for s in wb.sheets
        is_chartsheet(wb, s.name) || _repoint_sheet_xml_formulas!(s, r"(?s)\A.+\z" => kill)
    end
    for names in (wb.workbook_names, wb.worksheet_names), (k, v) in collect(names)
        v.value isa String && startswith(v.value, '=') && (names[k] = DefinedNameValue(kill(v.value), v.isabs, v.hidden))
    end
    for mime in (MIME_CHART, MIME_CHARTEX), path in parts_with_content_type(xf, mime)
        _rewrite_chart_formulas!(kill, xml_root_element(get_xml_data(xf, path)))
    end
    _rewrite_table_formulas!((f, _) -> kill(f), xf)
    return nothing
end
