# RK PID Annotator Excel Add-in

RK PID Annotator is a Windows Excel-DNA add-in for annotating P&ID PDF documents from the currently open Excel workbook. It reads tag rows from the active worksheet, searches a selected PDF for matching tag text, highlights matches, and adds PDF note annotations with selected Excel data.

The add-in is implemented in C# and runs locally inside desktop Excel. It does not require the Flask web app, Office.js, a browser task pane, or a cloud service.

## Features

- Read tag data directly from the active Excel worksheet.
- Select the header row, tag column, and note/comment columns.
- Preview worksheet data before running annotation.
- Highlight matching PDF text.
- Add PDF note annotations with selected workbook data.
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
dotnet restore .\excel-addin\RKPidAnnotator.csproj
dotnet build .\excel-addin\RKPidAnnotator.csproj -c Release /p:Platform=x64
```

The packed add-ins are created in:

```text
excel-addin\bin\x64\Release\net48\publish\
```

For 64-bit Excel, load:

```text
excel-addin\bin\x64\Release\net48\publish\RKPidAnnotator64-packed.xll
```

For 32-bit Excel, load:

```text
excel-addin\bin\x64\Release\net48\publish\RKPidAnnotator-packed.xll
```

## Basic Use

1. Open the workbook that contains the tag list.
2. Select the worksheet with the component data.
3. Click `PID Annotator` -> `Annotate PDFs`.
4. Choose the header row, tag column, and note columns.
5. Select the source PDF and output PDF path.
6. Click `Annotate`.

The add-in creates a new annotated PDF and leaves the source PDF unchanged.

## Create a Release Package

Run:

```powershell
.\excel-addin\scripts\package-release.ps1 -Version 1.0.0
```

The package is written to:

```text
excel-addin\dist\RKPidAnnotator-v1.0.0.zip
```

## Notes

The current C# engine works with PDFs that contain selectable text. Scanned drawing PDFs need OCR before tags can be found.
