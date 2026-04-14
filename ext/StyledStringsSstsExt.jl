module StyledStringsSstsExt

@static if VERSION >= v"1.11-"

    using XLSX
    import XLSX: setdata!, RichTextString, RichTextRun

    import StyledStrings: StyledStrings, AnnotatedString
    import StyledStrings: load_customisations!, getface, Face, FACES, SimpleColor

    @static if VERSION < v"1.14-"
        import StyledStrings: HTML_BASIC_COLORS
    end

    setdata!(sheet::Worksheet, ref::CellRef, ss::AnnotatedString{T}) where T =
        setdata!(sheet, ref, RichTextString(_ssToRuns(ss)))

"""
    _ssToRuns(s::Union{<:AnnotatedString, SubString{<:AnnotatedString}}) -> Vector{XLSX.RichTextRun}

Converts an `AnnotatedString` to a vector of `RichTextRun`s.
"""
    function _ssToRuns(s::Union{<:AnnotatedString, SubString{<:AnnotatedString}})
        runs = RichTextRun[]
        load_customisations!()
        for (str, styles) in Base.eachregion(s)
            push!(runs, RichTextRun(String(str), collect(_ss_style(getface(styles)))))
        end
        return runs
    end

"""
    _ss_color(color::SimpleColor) -> String

Convert a `SimpleColor` to an Excel-compatible hex color string (e.g. `"#ff0000"`),
or `""` if the color represents the default foreground.
"""
    function _ss_color(color::SimpleColor)
        if color.value isa Symbol
            if color.value in (:default, :foreground)
                return ""
            elseif (fg = get(FACES.current[], color.value, getface()).foreground) != SimpleColor(color.value)
                return _ss_color(fg)
            else
                @static if VERSION >= v"1.14-"
                    rgb = get(FACES.basecolors, color.value, nothing)
                    isnothing(rgb) && return ""
                    r, g, b = rgb.r, rgb.g, rgb.b
                else
                    return _ss_color(get(HTML_BASIC_COLORS, color.value, SimpleColor(:default)))
                end
            end
        else
            (; r, g, b) = color.value
        end
        io = IOBuffer()
        print(io, '#')
        r < 0x10 && print(io, '0')
        print(io, string(r, base=16))
        g < 0x10 && print(io, '0')
        print(io, string(g, base=16))
        b < 0x10 && print(io, '0')
        print(io, string(b, base=16))
        return String(take!(io))
    end

"""
    _ss_style(face::Face) -> Dict{Symbol, Any}

Creates a dictionary of Excel font attributes from a StyledString `face`.

Returns a Dict of (attribute => value).
"""
    function _ss_style(face::Face)
        d = Dict{Symbol, Any}()

        if !isnothing(face.font) && face.font != "monospace"
            d[:name] = face.font
        end

        if !isnothing(face.height)
            d[:size] = face.height ÷ 10
        end

        if !isnothing(face.weight) && face.weight in (:medium, :semibold, :bold, :extrabold, :black)
            d[:bold] = true
        end

        if !isnothing(face.slant) && face.slant in (:italic, :oblique)
            d[:italic] = true
        end

        if !isnothing(face.foreground)
            c = _ss_color(face.foreground)
            if c != ""
                d[:color] = c
            end
        end

        if !isnothing(face.underline)
            if face.underline isa Tuple || face.underline isa SimpleColor || face.underline === true
                d[:under] = true
            end
        end

        if !isnothing(face.strikethrough) && face.strikethrough
            d[:strike] = true
        end

        return d
    end

end # @static if VERSION >= v"1.11-"

end # module
