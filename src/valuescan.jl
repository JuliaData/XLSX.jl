#=
The byte scanner: a worksheet's cells read straight from the bytes of `<sheetData>`, for
`readtable`'s value rows and for the worksheet cache that `openxlsx` and `readxlsx` fill.

`ValueRowIterator` (src/valuerows.jl) and `first_cache_fill!` (src/stream.jl) can read a
sheet with an `XML.Cursor`, and most of their time then goes on XML.jl's general
tokenizer: a token per tag, attribute and text, each ready for anything XML allows. The
scanner reads the few shapes that `<sheetData>` is made of, a byte at a time, into the
same `ValueRow`s and `Cell`s: so everything above the row source is shared, and so are
the decoders (`process_tv`, `_cell_attribute_values`, `_rich_value`, `_set_cell!`,
`_push_filled_row!`).

It is a second parser, kept narrow so it never decides what unusual XML means:
  - XML.jl still finds `<sheetData>` (`_find_sheetdata`) and gives its text
    (`XML.sourcetext`), already rewritten by any DOCTYPE, so entities it declares are
    expanded and defaults supplied before the scanner sees a byte.
  - The fast path takes a `<c>` whose wanted attributes (`r`, `t`, `s`, `cm`, `vm`) need no
    normalization (no `&`, tab or line end), holding any `<f>` with only text and at most
    one `<v>` with only plain text (no `&`, CR or CDATA); or, for an inline string, one
    `<is>` holding one `<t>` of plain text, added to the shared strings by its text, as
    `_inline_string` and `add_to_sst!` add it. Element names may have any prefix, as
    `localname` reads them; comments, processing instructions and text between elements
    are skipped, as the cursor walk skips them; end tags aren't matched against start
    tags, as `XML.LazyNode` doesn't match them.
  - Any other `<c>` (a rich inline string, entities, a comment or element inside it, a
    second `<v>`, ...) is read by XML.jl: its text is parsed on its own and decoded by the
    cursor walk's own code (`_cell_attributes` and `_cursor_cell_contents`, or
    `_cursor_cell`). That is exact, at the cost of one small parse per such cell. So is
    a cell with an `<f>` when the cache loads formulas, to record the formula.
  - Anything that isn't markup the scanner can step over (a tag or comment cut off by the
    end of the text, a missing quote, a `<row>` attribute with an entity) throws
    `_ScanFallback`, and the sheet is read again with the cursor, which then does whatever
    it does today, errors included. For `readtable`, the pass that checks row order
    (`_check_rows_ascending!`) steps over rows with the same tag reader, so with
    `stop_in_row_function`, whose rows are checked before any is read, a fallback comes
    before the function is first called.

Malformed XML that both parsers step over can still differ in its errors.

Not scanned: `eachrow` with `enable_cache=false` (`SheetRowStreamIterator`) and the
uncached ranged reads (`_read_cells`), which read with a cursor.

Keeping the two in step: the tests compare scanned rows with cursor value rows and cache
rows, and caches filled by the scanner with caches filled by the cursor, for every test
workbook, and read sheets written in each shape above
(test/test_files/files/differential.jl). Set `_READTABLE_SCAN[] = false` or
`_CACHE_SCAN[] = false` to read with the cursor; or delete this file, its include, the
`scan` branches in src/valuerows.jl and `_readtable_value_rows`, and the `_CACHE_SCAN`
branch in `first_cache_fill!`.
=#

# `readtable` reads value rows with the byte scanner; `false` reads them with a cursor.
const _READTABLE_SCAN = Ref(true)

# `first_cache_fill!` reads rows with the byte scanner; `false` reads them with a cursor.
const _CACHE_SCAN = Ref(true)

# Thrown by the scanner for markup it can't step over; the sheet is then read with a cursor.
struct _ScanFallback <: Exception end

# Cells the scanner has handed to XML.jl, for the tests.
const _SCAN_CELL_FALLBACKS = Threads.Atomic{Int}(0)

mutable struct _ScanState
    text::SubString{String}   # `<sheetData …>…</sheetData>`
    pos::Int                  # where the next row is looked for
    last_row::Int
    done::Bool
    width::Int                # cells in the last row read, to size the next (all-columns mode)
end

# A start tag: where its name and wanted attribute values are in the text (`_NO_RANGE` when
# absent, which a present, empty value is not; for a repeated attribute, the last one,
# as `_cell_attributes` reads them).
struct _ScanTag
    name::UnitRange{Int}
    stop::Int                 # the index of its `>`
    empty::Bool               # `<…/>`
    r::UnitRange{Int}
    t::UnitRange{Int}
    s::UnitRange{Int}
    cm::UnitRange{Int}
    vm::UnitRange{Int}
end

const _NO_RANGE = 0:-1

@inline _byte(s::SubString{String}, i::Int) = 1 <= i <= ncodeunits(s) ? (@inbounds codeunit(s, i)) : 0x00
@inline _is_space(b::UInt8) = b == 0x20 || b == 0x0a || b == 0x09 || b == 0x0d

# The index of the next byte `b` from `i`, or 0.
@inline function _scan_find(s::SubString{String}, b::UInt8, i::Int)
    n = ncodeunits(s)
    while i <= n
        @inbounds codeunit(s, i) == b && return i
        i += 1
    end
    return 0
end

# The index after the next run of spaces from `i`.
@inline function _skip_spaces(s::SubString{String}, i::Int)
    while _is_space(_byte(s, i))
        i += 1
    end
    return i
end

# Whether `range` of `s` holds `lit`.
@inline function _scan_eq(s::SubString{String}, range::UnitRange{Int}, lit::String)
    length(range) == ncodeunits(lit) || return false
    for (k, i) in enumerate(range)
        @inbounds codeunit(s, i) == codeunit(lit, k) || return false
    end
    return true
end

# Whether the local name (after any prefix, as `localname` reads it) of `tag` is `lit`.
@inline function _scan_name_is(s::SubString{String}, tag::_ScanTag, lit::String)
    for i in tag.name
        @inbounds codeunit(s, i) == UInt8(':') && return _scan_eq(s, i+1:last(tag.name), lit)
    end
    return _scan_eq(s, tag.name, lit)
end

# Whether an attribute value is read as written: XML.jl normalizes one with a reference, a
# tab or a line end.
@inline function _scan_plain_attr(s::SubString{String}, range::UnitRange{Int})
    for i in range
        b = @inbounds codeunit(s, i)
        (b == UInt8('&') || b == 0x09 || b == 0x0a || b == 0x0d) && return false
    end
    return true
end

# The text of `range`, empty when the attribute is absent. Ranges end at the byte before
# an ASCII delimiter, which may be inside a character: the text ends at that character.
@inline _scan_text(s::SubString{String}, range::UnitRange{Int}) =
    isempty(range) ? SubString(s, 1, 0) : SubString(s, first(range), prevind(s, last(range) + 1))

# The attribute at `j` in a start tag: where its name and value are, and the index after it.
@inline function _scan_attribute(s::SubString{String}, j::Int)::Tuple{UnitRange{Int},UnitRange{Int},Int}
    k = j
    while (b = _byte(s, j); !(_is_space(b) || b == UInt8('=') || b == UInt8('>') || b == 0x00))
        j += 1
    end
    key = k:j-1
    j = _skip_spaces(s, j)
    (isempty(key) || _byte(s, j) != UInt8('=')) && throw(_ScanFallback())
    j = _skip_spaces(s, j + 1)
    q = _byte(s, j)
    (q == UInt8('"') || q == UInt8('\'')) || throw(_ScanFallback())
    e = _scan_find(s, q, j + 1)
    e == 0 && throw(_ScanFallback())
    return key, j+1:e-1, e + 1
end

# Whether the start tag's attributes end at `j`, with `>` or `/>`.
@inline _scan_attributes_end(s::SubString{String}, j::Int) =
    (b = _byte(s, j); b == UInt8('>') || b == UInt8('/'))

# The start tag whose `<` is at `i`.
function _scan_tag(s::SubString{String}, i::Int)::_ScanTag
    j = i + 1
    while (b = _byte(s, j); !(_is_space(b) || b == UInt8('>') || b == UInt8('/') || b == 0x00))
        j += 1
    end
    name = i+1:j-1
    isempty(name) && throw(_ScanFallback())
    r = t = st = cm = vm = _NO_RANGE
    while true
        j = _skip_spaces(s, j)
        _scan_attributes_end(s, j) && break
        key, val, j = _scan_attribute(s, j)
        if length(key) == 1
            c = codeunit(s, first(key))
            c == UInt8('r') ? (r = val) : c == UInt8('t') ? (t = val) : c == UInt8('s') ? (st = val) : nothing
        elseif _scan_eq(s, key, "cm")
            cm = val
        elseif _scan_eq(s, key, "vm")
            vm = val
        end
    end
    _byte(s, j) == UInt8('>') && return _ScanTag(name, j, false, r, t, st, cm, vm)
    _byte(s, j + 1) == UInt8('>') || throw(_ScanFallback())
    return _ScanTag(name, j + 1, true, r, t, st, cm, vm)
end

# The index of the last byte of the comment, CDATA section or processing instruction at `i`.
function _scan_markup_end(s::SubString{String}, i::Int)::Int
    close = startswith(SubString(s, i), "<!--") ? "-->" :
            startswith(SubString(s, i), "<![CDATA[") ? "]]>" :
            _byte(s, i + 1) == UInt8('?') ? "?>" : throw(_ScanFallback())
    found = findnext(close, s, i + 2)
    found === nothing && throw(_ScanFallback())
    return last(found)
end

# The index of the `>` that ends the end tag at `i`.
@inline function _scan_end_tag(s::SubString{String}, i::Int)::Int
    k = _scan_find(s, UInt8('>'), i)
    k == 0 && throw(_ScanFallback())
    return k
end

# The index of the last byte of the element whose start tag is `tag`.
function _scan_element_end(s::SubString{String}, tag::_ScanTag)::Int
    tag.empty && return tag.stop
    depth = 1
    j = tag.stop + 1
    while true
        j = _scan_find(s, UInt8('<'), j)
        j == 0 && throw(_ScanFallback())
        b = _byte(s, j + 1)
        if b == UInt8('/')
            j = _scan_end_tag(s, j)
            (depth -= 1) == 0 && return j
        elseif b == UInt8('!') || b == UInt8('?')
            j = _scan_markup_end(s, j)
        else
            inner = _scan_tag(s, j)
            inner.empty || (depth += 1)
            j = inner.stop
        end
        j += 1
    end
end

# The next element at `i` and below, among text, comments and processing instructions:
# its start tag, or `nothing` at an end tag, whose `>` is then at the returned index.
function _scan_next_tag(s::SubString{String}, i::Int)::Tuple{Union{Nothing,_ScanTag},Int}
    while true
        i = _scan_find(s, UInt8('<'), i)
        i == 0 && throw(_ScanFallback())
        b = _byte(s, i + 1)
        b == UInt8('/') && return nothing, _scan_end_tag(s, i)
        if b == UInt8('!') || b == UInt8('?')
            i = _scan_markup_end(s, i) + 1
        else
            tag = _scan_tag(s, i)
            return tag, tag.stop
        end
    end
end

# The text of `<sheetData>`; an unterminated one is left to the cursor, which reads what is there.
_scan_sheetdata_text(sheetdata::XML.LazyNode)::SubString{String} =
    try XML.sourcetext(sheetdata) catch e; e isa ErrorException ? throw(_ScanFallback()) : rethrow() end

# Call `f(ctag)` for each `<c>` in the `<row>` whose start tag is `tag`, stepping over
# anything else; `f` returns the index of the cell's last byte. Returns the index of the
# row's last byte and whether it has any `<c>`.
@inline function _scan_each_cell(f, s::SubString{String}, tag::_ScanTag)::Tuple{Int,Bool}
    i = tag.stop
    tag.empty && return i, false
    has_cells = false
    while true
        ctag, i = _scan_next_tag(s, i + 1)
        ctag === nothing && return i, has_cells                 # `</row>`
        if _scan_name_is(s, ctag, "c")
            has_cells = true
            i = f(ctag)
        else
            i = _scan_element_end(s, ctag)
        end
    end
end

#--------------------------------------------------------------------- One cell

# Whether the wanted attributes of a `<c>` are read as written (see `_scan_plain_attr`).
@inline _scan_plain_cell(s::SubString{String}, ctag::_ScanTag) =
    _scan_plain_attr(s, ctag.r) && _scan_plain_attr(s, ctag.t) && _scan_plain_attr(s, ctag.s) &&
    _scan_plain_attr(s, ctag.cm) && _scan_plain_attr(s, ctag.vm)

# `_cell_attributes` of a `<c>` that passed `_scan_plain_cell`.
@inline _scan_cell_attributes(s::SubString{String}, ctag::_ScanTag) =
    _cell_attribute_values(_scan_text(s, ctag.r), _scan_text(s, ctag.t), _scan_text(s, ctag.s),
                           _scan_text(s, ctag.cm), _scan_text(s, ctag.vm))

# The datatype and value of the `<c>` whose start tag is `ctag`, given its `t`, number
# style and `vm`, as `_cursor_cell_contents` reads them; the index of its last byte; and
# whether it has a formula, as the cursor walk flags one. The index is 0 for a cell outside
# the fast path, and for one with a formula when `formulas_by_xml`, which is then left
# undecoded.
function _scan_cell_value(s::SubString{String}, ctag::_ScanTag, wb::Workbook, t::SubString{String},
                          num_style::Int, vm::UInt32, formulas_by_xml::Bool)::Tuple{CellValueType,UInt64,Int,Bool}
    inline = t == "inlineStr"
    text, stop, has_f = _scan_cell_text(s, ctag, inline)
    (stop == 0 || (has_f && formulas_by_xml)) && return CT_EMPTY, UInt64(0), 0, has_f
    datatype, value = isempty(text) ? (CT_EMPTY, UInt64(0)) :
        # as `_inline_string` adds a plain `<is>`: by its text, as `add_to_sst!` keeps it
        inline ? (CT_STRING, reinterpret(UInt64, Int64(add_shared_string!(wb, text)))) :
        process_tv(wb, t, text, num_style)
    datatype, value = _rich_value(datatype, value, vm)
    return datatype, value, stop, has_f
end

# The value text of the `<c>` whose start tag is `ctag`, empty if none; the index of the
# cell's last byte, which is 0 for a cell outside the fast path; and whether it has an
# `<f>` the cursor walk flags (not in an inline string's cell). The text is that of its
# one `<v>`, or for an inline string (`inline`), as the cursor walk reads it, of the one
# `<t>` in its first `<is>`; any `<f>` is skipped.
function _scan_cell_text(s::SubString{String}, ctag::_ScanTag, inline::Bool)::Tuple{SubString{String},Int,Bool}
    text = SubString(s, 1, 0)
    ctag.empty && return text, ctag.stop, false
    seen = has_f = false
    i = ctag.stop + 1
    while true
        i = _scan_find(s, UInt8('<'), i)
        i == 0 && return text, 0, has_f
        b = _byte(s, i + 1)
        b == UInt8('/') && return text, _scan_find(s, UInt8('>'), i), has_f     # `</c>`
        (b == UInt8('!') || b == UInt8('?')) && return text, 0, has_f
        tag = _scan_tag(s, i)
        if _scan_name_is(s, tag, "f")
            has_f |= !inline
            _, i = _scan_text_element(s, tag, false)
        elseif !seen && _scan_name_is(s, tag, inline ? "is" : "v")
            seen = true
            text, i = inline ? _scan_inline_text(s, tag) : _scan_text_element(s, tag, true)
        else
            return text, 0, has_f
        end
        i == 0 && return text, 0, has_f
        i += 1
    end
end

# The text of the element whose start tag is `tag` and the index of its last byte, or 0
# for an element holding anything but text, or, with `plain`, text XML.jl would change
# (a reference or a CR).
function _scan_text_element(s::SubString{String}, tag::_ScanTag, plain::Bool)::Tuple{SubString{String},Int}
    tag.empty && return SubString(s, 1, 0), tag.stop
    e = _scan_find(s, UInt8('<'), tag.stop + 1)
    (e == 0 || _byte(s, e + 1) != UInt8('/')) && return SubString(s, 1, 0), 0
    text = SubString(s, tag.stop + 1, prevind(s, e))
    plain && (_scan_find(text, UInt8('&'), 1) != 0 || _scan_find(text, 0x0d, 1) != 0) && return text, 0
    return text, _scan_find(s, UInt8('>'), e)
end

# The text of an `<is>` holding one `<t>` of plain text, and the index of its last byte,
# or 0 for any other `<is>` (rich text runs, phonetic hints); text between them is ignored,
# as `_inline_string` ignores it.
function _scan_inline_text(s::SubString{String}, is::_ScanTag)::Tuple{SubString{String},Int}
    none = SubString(s, 1, 0)
    is.empty && return none, is.stop
    i = _scan_find(s, UInt8('<'), is.stop + 1)
    (i == 0 || _byte(s, i + 1) in (UInt8('/'), UInt8('!'), UInt8('?'))) && return none, 0
    t = _scan_tag(s, i)
    _scan_name_is(s, t, "t") || return none, 0
    text, k = _scan_text_element(s, t, true)
    k == 0 && return none, 0
    e = _scan_find(s, UInt8('<'), k + 1)                                  # `</is>`
    (e == 0 || _byte(s, e + 1) != UInt8('/')) && return none, 0
    return text, _scan_find(s, UInt8('>'), e)
end

# A cursor on the `<c>` whose start tag is `ctag`, its text parsed by XML.jl on its own,
# and the index of the cell's last byte: for a cell outside the fast path.
function _scan_cell_cursor(s::SubString{String}, ctag::_ScanTag)
    stop = _scan_element_end(s, ctag)
    Threads.atomic_add!(_SCAN_CELL_FALLBACKS, 1)
    c = XML.Cursor(parse(SubString(s, first(ctag.name) - 1, stop), XML.LazyNode))
    while XML.next!(c) !== nothing && XML.nodetype(c) != XML.Element
    end
    return c, stop
end

#--------------------------------------------------------------------- readtable's value rows

# A state at the start of `<sheetData>`, found as `_value_row_start` finds it.
function _scan_start(itr::ValueRowIterator)
    text = _scan_sheetdata_text(_find_sheetdata(parse(itr.xml, XML.LazyNode), itr.ws.name))
    tag = _scan_tag(text, 1)
    return _ScanState(text, tag.stop + 1, 0, tag.empty, 0)
end

# The next `<row>`'s start tag, other elements skipped, or `nothing` at the end of `<sheetData>`.
function _scan_next_row!(st::_ScanState)::Union{Nothing,_ScanTag}
    st.done && return nothing
    s = st.text
    while true
        tag, _ = _scan_next_tag(s, st.pos)
        if tag === nothing
            st.done = true
            return nothing
        end
        _scan_name_is(s, tag, "row") && return tag
        st.pos = _scan_element_end(s, tag) + 1
    end
end

# The row number of a `<row>`, as `_row_number` reads it.
function _scan_row_number(s::SubString{String}, tag::_ScanTag, wsname::String)::Int
    first(tag.r) == 0 && return _row_number_from(nothing, wsname)     # no `r`
    _scan_plain_attr(s, tag.r) || throw(_ScanFallback())
    return _row_number_from(_scan_text(s, tag.r), wsname)
end

function Base.iterate(itr::ValueRowIterator, st::_ScanState)
    itr.last_state[] = st
    tag = _scan_next_row!(st)
    tag === nothing && return nothing
    s = st.text
    row = _scan_row_number(s, tag, itr.ws.name)
    row <= st.last_row && throw(_RowsNotAscending())
    st.last_row = row
    firstcol, kinds, raws = _new_row_cells(itr.cols, st.width)
    stop, has_cells = _scan_each_cell(ctag -> _scan_cell!(kinds, raws, firstcol, itr, s, ctag), s, tag)
    st.pos = stop + 1
    st.width = length(kinds)
    return ValueRow(itr.ws, row, firstcol, kinds, raws, has_cells), st
end

# Read the `<c>` whose start tag is `ctag` into the row, if in the iterator's columns;
# returns the index of its last byte.
function _scan_cell!(kinds::Vector{UInt8}, raws::Vector{UInt64}, firstcol::Int, itr::ValueRowIterator,
                     s::SubString{String}, ctag::_ScanTag)::Int
    _scan_plain_cell(s, ctag) || return _scan_cell_by_xml!(kinds, raws, firstcol, itr, s, ctag)
    ref, _, t, _, num_style, _, vm = _scan_cell_attributes(s, ctag)
    col = column_number(ref)
    cols = itr.cols
    (isnothing(cols) || col in cols) || return _scan_element_end(s, ctag)
    datatype, value, stop, _ = _scan_cell_value(s, ctag, get_workbook(itr.ws), t, num_style, vm, false)
    stop == 0 && return _scan_cell_by_xml!(kinds, raws, firstcol, itr, s, ctag)
    _set_cell!(kinds, raws, firstcol, col, datatype, value)
    return stop
end

# Read the `<c>` whose start tag is `ctag` with XML.jl, as the cursor walk reads it;
# returns the index of its last byte.
function _scan_cell_by_xml!(kinds::Vector{UInt8}, raws::Vector{UInt64}, firstcol::Int, itr::ValueRowIterator,
                            s::SubString{String}, ctag::_ScanTag)::Int
    c, stop = _scan_cell_cursor(s, ctag)
    attrs = _cell_attributes(c)
    col = column_number(first(attrs))
    if isnothing(itr.cols) || col in itr.cols
        ws = itr.ws
        datatype, value, _ = _cursor_cell_contents(c, attrs, ws, get_workbook(ws), itr.sst_pfx, _NO_FORMULAS, false)
        _set_cell!(kinds, raws, firstcol, col, datatype, value)
    end
    return stop
end

# `_check_rows_ascending!` for the scanner: the `r` of each row not yet read, its
# contents stepped over.
function _check_rows_ascending!(itr::ValueRowIterator, st::_ScanState)
    s = st.text
    last = st.last_row
    while (tag = _scan_next_row!(st)) !== nothing
        if first(tag.r) != 0
            _scan_plain_attr(s, tag.r) || throw(_ScanFallback())
            r = tryparse(Int, _scan_text(s, tag.r))
            if !isnothing(r)
                r <= last && throw(_RowsNotAscending())
                last = r
            end
        end
        st.pos = _scan_element_end(s, tag) + 1
    end
    return nothing
end

#--------------------------------------------------------------------- The worksheet cache

# Read `<sheetData>`'s rows into the cache with the byte scanner, as `_cursor_fill_rows!`
# reads them.
function _scan_fill_rows!(fill::_CacheFill, sheetdata::XML.LazyNode)
    s = _scan_sheetdata_text(sheetdata)
    tag = _scan_tag(s, 1)
    tag.empty && return nothing
    pos = tag.stop + 1
    while true
        tag, _ = _scan_next_tag(s, pos)
        tag === nothing && return nothing                       # `</sheetData>`
        pos = (_scan_name_is(s, tag, "row") ? _scan_fill_row!(fill, s, tag) : _scan_element_end(s, tag)) + 1
    end
end

# Read the `<row>` whose start tag is `tag` into the cache; returns the index of its last byte.
function _scan_fill_row!(fill::_CacheFill, s::SubString{String}, tag::_ScanTag)::Int
    row_num, row_ht, unhandled = _scan_row_attributes(s, tag, fill.ws.name)
    rowcells = _new_rowcells(fill)
    stop, _ = _scan_each_cell(s, tag) do ctag
        cell, cell_stop = _scan_cache_cell(fill, s, ctag)
        _add_filled_cell!(fill, rowcells, cell)
        cell_stop
    end
    _push_filled_row!(fill, row_num, row_ht, rowcells, unhandled)
    return stop
end

# The number, height and unhandled attributes of a `<row>`, as `_cursor_fill_rows!` reads
# them. A value XML.jl would normalize hands the sheet back to the cursor.
function _scan_row_attributes(s::SubString{String}, tag::_ScanTag, wsname::String)
    r_val = ht_val = nothing
    unhandled = _EMPTY_ROW_ATTRS
    j = last(tag.name) + 1
    while true
        j = _skip_spaces(s, j)
        _scan_attributes_end(s, j) && break
        key, val, j = _scan_attribute(s, j)
        _scan_plain_attr(s, val) || throw(_ScanFallback())
        k, v = _scan_text(s, key), _scan_text(s, val)
        if k == "r"
            r_val = v
        elseif k == "ht"
            ht_val = v
        elseif !(k in _HANDLED_ROW_ATTRIBUTES)
            unhandled === _EMPTY_ROW_ATTRS && (unhandled = Dict{String,String}())
            unhandled[String(k)] = String(v)
        end
    end
    return _row_number_from(r_val, wsname), (isnothing(ht_val) ? nothing : parse(Float64, ht_val)), unhandled
end

# The `Cell` of the `<c>` whose start tag is `ctag`, as `_cursor_cell` reads it, and the
# index of its last byte. When formulas are loaded, a cell with an `<f>` is read by XML.jl,
# which records the formula.
function _scan_cache_cell(fill::_CacheFill, s::SubString{String}, ctag::_ScanTag)::Tuple{Cell,Int}
    if _scan_plain_cell(s, ctag)
        ref, _, t, style, num_style, meta, vm = _scan_cell_attributes(s, ctag)
        datatype, value, stop, has_f = _scan_cell_value(s, ctag, fill.wb, t, num_style, vm, fill.load_formulas)
        stop == 0 || return Cell(ref, value, style, meta, datatype, has_f), stop
    end
    c, stop = _scan_cell_cursor(s, ctag)
    return _cursor_cell(c, fill.ws, fill.wb, fill.sst_pfx, fill.local_formulas, fill.load_formulas), stop
end
