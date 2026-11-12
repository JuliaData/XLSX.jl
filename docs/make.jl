using Documenter, XLSX
makedocs(
    sitename = "XLSX.jl",
    modules = [XLSX, XLSX.Charts],
    pages = [
        "Home" => "index.md",
        "Supported features" => "features.md",
        "Getting started" => Any[
            "XLSX Tutorial" => "tutorial/XLSXtutorial.md",
            "Creating and editing workbooks" => "tutorial/creatingWorkbooks.md",
            "Addressing cells and ranges" => "concepts/addressing.md",
            "FileIO Tutorial" => "tutorial/FileIOtutorial.md",
        ],
        "Cookbook" => "cookbook.md",
        "Data" => Any[
            "Excel Tables" => "tables/excelTables.md",
            "Formulas" => "formulae/formulas.md",
            "Named ranges" => "formulae/namedRanges.md",
            "Data validation and drop-downs" => "sheets/dataValidation.md",
            "AutoFilter (filter buttons)" => "sheets/autoFilter.md",
        ],
        "Formatting" => Any[
            "Cell formats" => "formatting/cellFormatting.md",
            "Conditional formats" => "formatting/conditionalFormatting.md",
            "Merged cells" => "formatting/mergedCells.md",
        ],
        "Layout and view" => Any[
            "Column width and row height" => "formatting/widthAndHeight.md",
            "Inserting and deleting rows and columns" => "sheets/insertDelete.md",
            "Hide and group rows/columns (outlines)" => "formatting/hideAndGroup.md",
            "Freeze and split panes" => "formatting/freezeAndSplitPanes.md",
            "Zoom, gridlines and the view" => "formatting/sheetView.md",
        ],
        "Worksheets" => Any[
            "Managing worksheets" => "sheets/worksheets.md",
            "Protecting a sheet" => "sheets/protection.md",
            "Page setup and printing" => "sheets/printing.md",
        ],
        "Images, comments and links" => Any[
            "Images" => "sheets/images.md",
            "Comments" => "sheets/comments.md",
            "Hyperlinks" => "sheets/hyperlinks.md",
        ],
        "Charts" => Any[
            "Overview" => "charts/excelCharts.md",
            "Creating charts" => "charts/creatingCharts.md",
            "Creating chartEx charts" => "charts/creatingChartEx.md",
            "Formatting charts" => "charts/formattingCharts.md",
            "Reading charts" => "charts/readingCharts.md",
            "Sheets with charts" => "charts/chartSheets.md",
            "Limitations" => "charts/chartLimitations.md",
        ],
        "Examples" => "examples.md",
        "API Reference" => Any[
            "Index" => "api/index.md",
            "Files and worksheets" => "api/files.md",
            "Cells and data" => "api/data.md",
            "Tables" => "api/tables.md",
            "Cell formats" => "api/formats.md",
            "Conditional formats" => "api/conditionalFormats.md",
            "Sheet layout" => "api/sheetLayout.md",
            "Images, comments and hyperlinks" => "api/objects.md",
            "Charts" => "api/charts.md",
            "Chart axes and elements" => "api/chartElements.md",
            "chartEx charts" => "api/chartEx.md",
            "Chart types" => "api/chartTypes.md",
        ],
        "FAQ and troubleshooting" => "faq.md",
        "Migration Guide" => "migration.md",
    ],
    checkdocs = :public,
    # External links are checked only when asked, as they depend on other sites being up:
    # the CI job "Documentation links" sets DOCUMENTER_LINKCHECK=true.
    linkcheck = get(ENV, "DOCUMENTER_LINKCHECK", "false") == "true",
)
deploydocs(
    repo = "github.com/JuliaData/XLSX.jl.git",
    target = "build",
    versions = [
        "stable" => "v^",
        "dev" => "dev"
    ],
)
