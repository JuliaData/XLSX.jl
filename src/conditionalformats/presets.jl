const highlights::Dict{String,Dict{String,Dict{String,String}}} = Dict(
    "redfilltext" => Dict(
        "font" => Dict("color" => "FF9C0006"),
        "fill" => Dict("pattern" => "solid", "bgColor" => "FFFFC7CE")
    ),
    "yellowfilltext" => Dict(
        "font" => Dict("color" => "FF9C5700"),
        "fill" => Dict("pattern" => "solid", "bgColor" => "FFFFEB9C")
    ),
    "greenfilltext" => Dict(
        "font" => Dict("color" => "FF006100"),
        "fill" => Dict("pattern" => "solid", "bgColor" => "FFC6EFCE")
    ),
    "redfill" => Dict(
        "fill" => Dict("pattern" => "solid", "bgColor" => "FFFFC7CE")
    ),
    "redtext" => Dict(
        "font" => Dict("color" => "FF9C0006"),
    ),
    "redborder" => Dict(
        "border" => Dict("color" => "FF9C0006", "style" => "thin")
    )
)
const databars::Dict{String,Dict{String,String}} = Dict(
    "bluegrad" => Dict(
        "min_type" => "automatic",
        "max_type" => "automatic",
        "fill_col" => "FF638EC6",
        "borders" => "true",
        "sameNegBorders" => "false",
        "border_col" => "FF638EC6",
        "neg_fill_col" => "FFFF0000",
        "neg_border_col" => "FFFF0000",
        "axis_col" => "FF000000"
    ),
    "greengrad" => Dict(
        "min_type" => "automatic",
        "max_type" => "automatic",
        "fill_col" => "FF63C384",
        "borders" => "true",
        "sameNegBorders" => "false",
        "border_col" => "FF63C384",
        "neg_fill_col" => "FFFF0000",
        "neg_border_col" => "FFFF0000",
        "axis_col" => "FF000000"
    ),
    "redgrad" => Dict(
        "min_type" => "automatic",
        "max_type" => "automatic",
        "fill_col" => "FFFF555A",
        "borders" => "true",
        "sameNegBorders" => "false",
        "border_col" => "FFFF555A",
        "neg_fill_col" => "FFFF0000",
        "neg_border_col" => "FFFF0000",
        "axis_col" => "FF000000"
    ),
    "orangegrad" => Dict(
        "min_type" => "automatic",
        "max_type" => "automatic",
        "fill_col" => "FFFFB628",
        "borders" => "true",
        "sameNegBorders" => "false",
        "border_col" => "FFFFB628",
        "neg_fill_col" => "FFFF0000",
        "neg_border_col" => "FFFF0000",
        "axis_col" => "FF000000"
    ),
    "lightbluegrad" => Dict(
        "min_type" => "automatic",
        "max_type" => "automatic",
        "fill_col" => "FF008AEF",
        "borders" => "true",
        "sameNegBorders" => "false",
        "border_col" => "FF008AEF",
        "neg_fill_col" => "FFFF0000",
        "neg_border_col" => "FFFF0000",
        "axis_col" => "FF000000"
    ),
    "purplegrad" => Dict(
        "min_type" => "automatic",
        "max_type" => "automatic",
        "fill_col" => "FFD6007B",
        "borders" => "true",
        "sameNegBorders" => "false",
        "border_col" => "FFD6007B",
        "neg_fill_col" => "FFFF0000",
        "neg_border_col" => "FFFF0000",
        "axis_col" => "FF000000"
    ),
    "blue" => Dict(
        "min_type" => "automatic",
        "max_type" => "automatic",
        "fill_col" => "FF638EC6",
        "gradient" => "false",
        "neg_fill_col" => "FFFF0000",
        "axis_col" => "FF000000"
    ),
    "green" => Dict(
        "min_type" => "automatic",
        "max_type" => "automatic",
        "fill_col" => "FF63C384",
        "gradient" => "false",
        "neg_fill_col" => "FFFF0000",
        "axis_col" => "FF000000"
    ),
    "red" => Dict(
        "min_type" => "automatic",
        "max_type" => "automatic",
        "fill_col" => "FFFF555A",
        "gradient" => "false",
        "neg_fill_col" => "FFFF0000",
        "axis_col" => "FF000000"
    ),
    "orange" => Dict(
        "min_type" => "automatic",
        "max_type" => "automatic",
        "fill_col" => "FFFFB628",
        "gradient" => "false",
        "neg_fill_col" => "FFFF0000",
        "axis_col" => "FF000000"
    ),
    "lightblue" => Dict(
        "min_type" => "automatic",
        "max_type" => "automatic",
        "fill_col" => "FF008AEF",
        "gradient" => "false",
        "neg_fill_col" => "FFFF0000",
        "axis_col" => "FF000000"
    ),
    "purple" => Dict(
        "min_type" => "automatic",
        "max_type" => "automatic",
        "fill_col" => "FFD6007B",
        "gradient" => "false",
        "neg_fill_col" => "FFFF0000",
        "axis_col" => "FF000000"
    ),
)
const colorscales::Dict{String,XML.Node} = Dict(    # Defines the 12 standard, built-in Excel color scales for conditional formatting.
    "greenyellowred" => XML.h.cfRule(type="colorScale", priority="1",
        XML.h.colorScale(
            XML.h.cfvo(type="min"),
            XML.h.cfvo(type="percentile", val="50"),
            XML.h.cfvo(type="max"),
            XML.h.color(rgb="FFF8696B"),
            XML.h.color(rgb="FFFFEB84"),
            XML.h.color(rgb="FF63BE7B")
        )
    ),
    "redyellowgreen" => XML.h.cfRule(type="colorScale", priority="1",
        XML.h.colorScale(
            XML.h.cfvo(type="min"),
            XML.h.cfvo(type="percentile", val="50"),
            XML.h.cfvo(type="max"),
            XML.h.color(rgb="FF63BE7B"),
            XML.h.color(rgb="FFFFEB84"),
            XML.h.color(rgb="FFF8696B")
        )
    ),
    "greenwhitered" => XML.h.cfRule(type="colorScale", priority="1",
        XML.h.colorScale(
            XML.h.cfvo(type="min"),
            XML.h.cfvo(type="percentile", val="50"),
            XML.h.cfvo(type="max"),
            XML.h.color(rgb="FFF8696B"),
            XML.h.color(rgb="FFFCFCFF"),
            XML.h.color(rgb="FF63BE7B")
        )
    ),
    "redwhitegreen" => XML.h.cfRule(type="colorScale", priority="1",
        XML.h.colorScale(
            XML.h.cfvo(type="min"),
            XML.h.cfvo(type="percentile", val="50"),
            XML.h.cfvo(type="max"),
            XML.h.color(rgb="FF63BE7B"),
            XML.h.color(rgb="FFFCFCFF"),
            XML.h.color(rgb="FFF8696B")
        )
    ),
    "bluewhitered" => XML.h.cfRule(type="colorScale", priority="1",
        XML.h.colorScale(
            XML.h.cfvo(type="min"),
            XML.h.cfvo(type="percentile", val="50"),
            XML.h.cfvo(type="max"),
            XML.h.color(rgb="FFF8696B"),
            XML.h.color(rgb="FFFCFCFF"),
            XML.h.color(rgb="FF5A8AC6")
        )
    ),
    "redwhiteblue" => XML.h.cfRule(type="colorScale", priority="1",
        XML.h.colorScale(
            XML.h.cfvo(type="min"),
            XML.h.cfvo(type="percentile", val="50"),
            XML.h.cfvo(type="max"),
            XML.h.color(rgb="FF5A8AC6"),
            XML.h.color(rgb="FFFCFCFF"),
            XML.h.color(rgb="FFF8696B")
        )
    ),
    "whitered" => XML.h.cfRule(type="colorScale", priority="1",
        XML.h.colorScale(
            XML.h.cfvo(type="min"),
            XML.h.cfvo(type="max"),
            XML.h.color(rgb="FFF8696B"),
            XML.h.color(rgb="FFFCFCFF")
        )
    ),
    "redwhite" => XML.h.cfRule(type="colorScale", priority="1",
        XML.h.colorScale(
            XML.h.cfvo(type="min"),
            XML.h.cfvo(type="max"),
            XML.h.color(rgb="FFFCFCFF"),
            XML.h.color(rgb="FFF8696B")
        )
    ),
    "greenwhite" => XML.h.cfRule(type="colorScale", priority="1",
        XML.h.colorScale(
            XML.h.cfvo(type="min"),
            XML.h.cfvo(type="max"),
            XML.h.color(rgb="FFFCFCFF"),
            XML.h.color(rgb="FF63BE7B")
        )
    ),
    "whitegreen" => XML.h.cfRule(type="colorScale", priority="1",
        XML.h.colorScale(
            XML.h.cfvo(type="min"),
            XML.h.cfvo(type="max"),
            XML.h.color(rgb="FF63BE7B"),
            XML.h.color(rgb="FFFCFCFF")
        )
    ),
    "greenyellow" => XML.h.cfRule(type="colorScale", priority="1",
        XML.h.colorScale(
            XML.h.cfvo(type="min"),
            XML.h.cfvo(type="max"),
            XML.h.color(rgb="FFFFEF9C"),
            XML.h.color(rgb="FF63BE7B")
        )
    ),
    "yellowgreen" => XML.h.cfRule(type="colorScale", priority="1",
        XML.h.colorScale(
            XML.h.cfvo(type="min"),
            XML.h.cfvo(type="max"),
            XML.h.color(rgb="FF63BE7B"),
            XML.h.color(rgb="FFFFEF9C")
        )
    )
)
const iconsets::Dict{String,XML.Node} = Dict(    # Defines the 20 standard, built-in Excel icon sets for conditional formatting.
    "3Arrows" => XML.h.cfRule(type="iconSet", priority="1",
        XML.h.iconSet(iconSet="3Arrows",
            XML.h.cfvo(type="percent", val="0"),
            XML.h.cfvo(type="percent", val="33"),
            XML.h.cfvo(type="percent", val="67"),
        )
    ),
    "4Arrows" => XML.h.cfRule(type="iconSet", priority="1",
        XML.h.iconSet(iconSet="4Arrows",
            XML.h.cfvo(type="percent", val="0"),
            XML.h.cfvo(type="percent", val="25"),
            XML.h.cfvo(type="percent", val="50"),
            XML.h.cfvo(type="percent", val="75"),
        )
    ),
    "5Arrows" => XML.h.cfRule(type="iconSet", priority="1",
        XML.h.iconSet(iconSet="5Arrows",
            XML.h.cfvo(type="percent", val="0"),
            XML.h.cfvo(type="percent", val="20"),
            XML.h.cfvo(type="percent", val="40"),
            XML.h.cfvo(type="percent", val="60"),
            XML.h.cfvo(type="percent", val="80"),
        )
    ),
    "3ArrowsGray" => XML.h.cfRule(type="iconSet", priority="1",
        XML.h.iconSet(iconSet="3ArrowsGray",
            XML.h.cfvo(type="percent", val="0"),
            XML.h.cfvo(type="percent", val="33"),
            XML.h.cfvo(type="percent", val="67"),
        )
    ),
    "4ArrowsGray" => XML.h.cfRule(type="iconSet", priority="1",
        XML.h.iconSet(iconSet="4ArrowsGray",
            XML.h.cfvo(type="percent", val="0"),
            XML.h.cfvo(type="percent", val="25"),
            XML.h.cfvo(type="percent", val="50"),
            XML.h.cfvo(type="percent", val="75"),
        )
    ),
    "5ArrowsGray" => XML.h.cfRule(type="iconSet", priority="1",
        XML.h.iconSet(iconSet="5ArrowsGray",
            XML.h.cfvo(type="percent", val="0"),
            XML.h.cfvo(type="percent", val="20"),
            XML.h.cfvo(type="percent", val="40"),
            XML.h.cfvo(type="percent", val="60"),
            XML.h.cfvo(type="percent", val="80"),
        )
    ),
    # The OOXML name is `3TrafficLights1`; XLSX.jl calls it `3TrafficLights`.
    "3TrafficLights" => XML.h.cfRule(type="iconSet", priority="1",
        XML.h.iconSet(iconSet="3TrafficLights1",
            XML.h.cfvo(type="percent", val="0"),
            XML.h.cfvo(type="percent", val="33"),
            XML.h.cfvo(type="percent", val="67"),
        )
    ),
    "3Signs" => XML.h.cfRule(type="iconSet", priority="1",
        XML.h.iconSet(iconSet="3Signs",
            XML.h.cfvo(type="percent", val="0"),
            XML.h.cfvo(type="percent", val="33"),
            XML.h.cfvo(type="percent", val="67"),
        )
    ),
    "4RedToBlack" => XML.h.cfRule(type="iconSet", priority="1",
        XML.h.iconSet(iconSet="4RedToBlack",
            XML.h.cfvo(type="percent", val="0"),
            XML.h.cfvo(type="percent", val="25"),
            XML.h.cfvo(type="percent", val="50"),
            XML.h.cfvo(type="percent", val="75"),
        )
    ),
    "3TrafficLights2" => XML.h.cfRule(type="iconSet", priority="1",
        XML.h.iconSet(iconSet="3TrafficLights2",
            XML.h.cfvo(type="percent", val="0"),
            XML.h.cfvo(type="percent", val="33"),
            XML.h.cfvo(type="percent", val="67"),
        )
    ),
    "4TrafficLights" => XML.h.cfRule(type="iconSet", priority="1",
        XML.h.iconSet(iconSet="4TrafficLights",
            XML.h.cfvo(type="percent", val="0"),
            XML.h.cfvo(type="percent", val="25"),
            XML.h.cfvo(type="percent", val="50"),
            XML.h.cfvo(type="percent", val="75"),
        )
    ),
    "3Symbols" => XML.h.cfRule(type="iconSet", priority="1",
        XML.h.iconSet(iconSet="3Symbols",
            XML.h.cfvo(type="percent", val="0"),
            XML.h.cfvo(type="percent", val="33"),
            XML.h.cfvo(type="percent", val="67"),
        )
    ),
    "3Symbols2" => XML.h.cfRule(type="iconSet", priority="1",
        XML.h.iconSet(iconSet="3Symbols2",
            XML.h.cfvo(type="percent", val="0"),
            XML.h.cfvo(type="percent", val="33"),
            XML.h.cfvo(type="percent", val="67"),
        )
    ),
    "3Flags" => XML.h.cfRule(type="iconSet", priority="1",
        XML.h.iconSet(iconSet="3Flags",
            XML.h.cfvo(type="percent", val="0"),
            XML.h.cfvo(type="percent", val="33"),
            XML.h.cfvo(type="percent", val="67"),
        )
    ),
    "5Quarters" => XML.h.cfRule(type="iconSet", priority="1",
        XML.h.iconSet(iconSet="5Quarters",
            XML.h.cfvo(type="percent", val="0"),
            XML.h.cfvo(type="percent", val="20"),
            XML.h.cfvo(type="percent", val="40"),
            XML.h.cfvo(type="percent", val="60"),
            XML.h.cfvo(type="percent", val="80"),
        )
    ),
    "4Rating" => XML.h.cfRule(type="iconSet", priority="1",
        XML.h.iconSet(iconSet="4Rating",
            XML.h.cfvo(type="percent", val="0"),
            XML.h.cfvo(type="percent", val="25"),
            XML.h.cfvo(type="percent", val="50"),
            XML.h.cfvo(type="percent", val="75"),
        )
    ),
    "5Rating" => XML.h.cfRule(type="iconSet", priority="1",
        XML.h.iconSet(iconSet="5Rating",
            XML.h.cfvo(type="percent", val="0"),
            XML.h.cfvo(type="percent", val="20"),
            XML.h.cfvo(type="percent", val="40"),
            XML.h.cfvo(type="percent", val="60"),
            XML.h.cfvo(type="percent", val="80"),
        )
    ),
    # These three require Excel 2010 extensions and will be ignored by earlier versions of Excel.
    "3Triangles" => get_x14_icon("3Triangles"),
    "3Stars" => get_x14_icon("3Stars"),
    "5Boxes" => get_x14_icon("5Boxes"),
    "Custom" => get_x14_icon("Custom")
)
const allIcons::Dict{String,Tuple{String,String}} = Dict(
    "1" => ("3Arrows", "0"),
    "2" => ("3Arrows", "1"),
    "3" => ("3Arrows", "2"),
    "4" => ("4Arrows", "1"),
    "5" => ("4Arrows", "2"),
    "6" => ("3ArrowsGray", "0"),
    "7" => ("3ArrowsGray", "1"),
    "8" => ("3ArrowsGray", "2"),
    "9" => ("4ArrowsGray", "1"),
    "10" => ("4ArrowsGray", "2"),
    "11" => ("3Flags", "0"),
    "12" => ("3Flags", "1"),
    "13" => ("3Flags", "2"),
    "14" => ("3TrafficLights1", "0"),
    "15" => ("3TrafficLights1", "1"),
    "16" => ("3TrafficLights1", "2"),
    "17" => ("3TrafficLights2", "0"),
    "18" => ("3TrafficLights2", "1"),
    "19" => ("3TrafficLights2", "2"),
    "20" => ("4TrafficLights", "0"),
    "21" => ("3Signs", "0"),
    "22" => ("3Signs", "1"),
    "23" => ("3Symbols", "0"),
    "24" => ("3Symbols", "1"),
    "25" => ("3Symbols", "2"),
    "26" => ("3Symbols2", "0"),
    "27" => ("3Symbols2", "1"),
    "28" => ("3Symbols2", "2"),
    "29" => ("4RedToBlack", "0"),
    "30" => ("4RedToBlack", "1"),
    "31" => ("4RedToBlack", "2"),
    "32" => ("4RedToBlack", "3"),
    "33" => ("5Quarters", "0"),
    "34" => ("5Quarters", "1"),
    "35" => ("5Quarters", "2"),
    "36" => ("5Quarters", "3"),
    "37" => ("5Rating", "0"),
    "38" => ("5Rating", "1"),
    "39" => ("5Rating", "2"),
    "40" => ("5Rating", "3"),
    "41" => ("5Rating", "4"),
    "42" => ("3Stars", "0"),
    "43" => ("3Stars", "1"),
    "44" => ("3Stars", "2"),
    "45" => ("3Triangles", "0"),
    "46" => ("3Triangles", "1"),
    "47" => ("3Triangles", "2"),
    "48" => ("5Boxes", "0"),
    "49" => ("5Boxes", "1"),
    "50" => ("5Boxes", "2"),
    "51" => ("5Boxes", "3"),
    "52" => ("5Boxes", "4")
)

const timeperiods::Dict{String,String} = Dict(
    "last7Days" => "AND(TODAY()-FLOOR(__CR__,1)<=6,FLOOR(__CR__,1)<=TODAY())",
    "yesterday" => "FLOOR(__CR__,1)=TODAY()-1",
    "today" => "FLOOR(__CR__,1)=TODAY()",
    "tomorrow" => "FLOOR(__CR__,1)=TODAY()+1",
    "lastWeek" => "AND(TODAY()-ROUNDDOWN(__CR__,0)>=(WEEKDAY(TODAY())),TODAY()-ROUNDDOWN(__CR__,0)<(WEEKDAY(TODAY())+7))",
    "thisWeek" => "AND(TODAY()-ROUNDDOWN(__CR__,0)<=WEEKDAY(TODAY())-1,ROUNDDOWN(__CR__,0)-TODAY()<=7-WEEKDAY(TODAY()))",
    "nextWeek" => "AND(ROUNDDOWN(__CR__,0)-TODAY()>(7-WEEKDAY(TODAY())),ROUNDDOWN(__CR__,0)-TODAY()<(15-WEEKDAY(TODAY())))",
    "lastMonth" => "AND(MONTH(__CR__)=MONTH(EDATE(TODAY(),0-1)),YEAR(__CR__)=YEAR(EDATE(TODAY(),0-1)))",
    "thisMonth" => "AND(MONTH(__CR__)=MONTH(TODAY()),YEAR(__CR__)=YEAR(TODAY()))",
    "nextMonth" => "AND(MONTH(__CR__)=MONTH(EDATE(TODAY(),0+1)),YEAR(__CR__)=YEAR(EDATE(TODAY(),0+1)))"
)

const CfKey = Union{CellRange,NonContiguousRange}
