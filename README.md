# RK PDF-Annotator Excel Add-in

RK PDF-Annotator is a Windows Excel-DNA add-in for annotating P&ID PDF documents from the currently open Excel workbook. It reads tag rows from the active worksheet, searches a selected PDF for matching tag text, highlights matches, and adds PDF note annotations with selected Excel data.

The add-in is implemented in C# and runs locally inside desktop Excel. It does not require the Flask web app, Office.js, a browser task pane, or a cloud service.

## Features

- Read tag data directly from the active Excel worksheet.
- Select the header row, tag column, and note/comment columns.
- Configure the PDF tag pattern by part count and separators, such as 3-5 parts using `-` or `.`.
- Preview worksheet data before running annotation.
- Build a PDF tag index before annotation and report PDF tags that are not present in Excel.
- Highlight matching PDF text.
- Add PDF note annotations with selected workbook data.
- Draw native PDF watermarks beneath matched tags using one or more Excel column values.
- Apply per-tag highlight color rules based on a tag part or an Excel column value, with a configurable default color.
- Show annotation progress while indexing, annotating, saving, and creating the report.
- Create a formatted Excel report sheet in the active workbook after annotation.
- Open a built-in HTML user guide from the Excel ribbon.
- Show version, author, contact, and GitHub information from the `About` button.

## Requirements

- Windows.
- Microsoft Excel desktop.
- .NET Framework 4.8 runtime.
- 64-bit Excel is recommended.

## Build From Source

From the repository root:

```powershell
dotnet restore .\excel-addin\RKPdfAnnotator.csproj
dotnet build .\excel-addin\RKPdfAnnotator.csproj -c Release /p:Platform=x64
```

The packed add-ins are created in:

```text
excel-addin\bin\x64\Release\net48\publish\
```

For 64-bit Excel, load:

```text
excel-addin\bin\x64\Release\net48\publish\RKPdfAnnotator64-packed.xll
```

For 32-bit Excel, load:

```text
excel-addin\bin\x64\Release\net48\publish\RKPdfAnnotator-packed.xll
```

## Basic Use

1. Open the workbook that contains the tag list.
2. Select the worksheet with the component data.
3. Click `PDF Annotator` -> `Annotate PDFs`.
4. Choose the header row, tag column, and note columns.
5. Set the tag format. The default is 3-5 parts with `-` or `.` separators.
6. To add visible labels beneath tags, enable watermarking and select one or more watermark columns.
7. Optionally reorder the watermark columns, change the font size or text color, and enable a white background.
8. Optionally add color rules to vary the highlight color by tag part or Excel column value, and set a default highlight color.
9. Select the source PDF and output PDF path.
10. Click `Annotate`.

The add-in creates a new annotated PDF and leaves the source PDF unchanged.

Before annotation, the add-in indexes all tags in the PDF that match the configured tag format. When annotation finishes, it creates a formatted `PDF Tag Report` worksheet in the active workbook with a summary, duplicate Excel tags, found tags, Excel tags not found in the PDF, and per-page tag counts. It also writes a CSV copy next to the output PDF as a file-based fallback.

Watermark values are joined in the selected order with ` / `. Empty cells are skipped. Watermarks are written directly into the PDF by the add-in and do not require the web app or Python.

## Create a Release Package

Run:

```powershell
.\excel-addin\scripts\package-release.ps1 -Version 1.0.0
```

The package is written to:

```text
excel-addin\dist\RKPdfAnnotator-v1.0.0.zip
```

## Notes

The current C# engine works with PDFs that contain selectable text. Scanned drawing PDFs need OCR before tags can be found.
