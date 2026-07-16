param(
    [Parameter(Mandatory = $false)]
    [string]$SamplePdf = ""
)

$ErrorActionPreference = "Stop"

$scriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$addinRoot = Split-Path -Parent $scriptRoot
$repoRoot = Split-Path -Parent $addinRoot
$bin = Join-Path $addinRoot "bin\x64\Release\net48"
$testRoot = Join-Path $addinRoot "tmp-smoke"
$inputPdf = Join-Path $testRoot "sample.pdf"
$outputPdf = Join-Path $testRoot "sample_annotated.pdf"

New-Item -ItemType Directory -Force -Path $testRoot | Out-Null

$pdfSharpPath = Join-Path $bin "PdfSharp.dll"
[System.Reflection.Assembly]::LoadFrom($pdfSharpPath) | Out-Null

$makerCode = @"
using System;
using System.IO;
using PdfSharp.Drawing;
using PdfSharp.Pdf;

public static class SmokePdfMaker
{
    public static string Make(string path)
    {
        PdfDocument document = new PdfDocument();
        PdfPage page = document.AddPage();
        using (XGraphics graphics = XGraphics.FromPdfPage(page))
        {
            XFont font = new XFont("Arial", 16);
            graphics.DrawString("PUMP-101-A is installed near the inlet.", font, XBrushes.Black, new XPoint(72, 120));
        }

        PdfPage rotatedPage = document.AddPage();
        using (XGraphics graphics = XGraphics.FromPdfPage(rotatedPage))
        {
            XFont font = new XFont("Arial", 16);
            graphics.DrawString("PUMP-101-A is shown on a rotated page.", font, XBrushes.Black, new XPoint(72, 120));
        }
        rotatedPage.Rotate = 90;

        document.Save(path);
        return path;
    }
}
"@

Add-Type -ReferencedAssemblies @($pdfSharpPath, "System.Drawing") -TypeDefinition $makerCode
[SmokePdfMaker]::Make($inputPdf) | Out-Null

if (Test-Path $outputPdf) {
    Remove-Item -LiteralPath $outputPdf -Force
}

$dlls = @(
    "PdfSharp.dll",
    "PdfSharp.Charting.dll",
    "Microsoft.Bcl.HashCode.dll",
    "System.Buffers.dll",
    "System.Memory.dll",
    "System.Numerics.Vectors.dll",
    "System.Runtime.CompilerServices.Unsafe.dll",
    "UglyToad.PdfPig.dll",
    "UglyToad.PdfPig.Core.dll",
    "UglyToad.PdfPig.DocumentLayoutAnalysis.dll",
    "UglyToad.PdfPig.Fonts.dll",
    "UglyToad.PdfPig.Package.dll",
    "UglyToad.PdfPig.Tokenization.dll",
    "UglyToad.PdfPig.Tokens.dll",
    "RKPdfAnnotator.dll"
)

foreach ($dll in $dlls) {
    [System.Reflection.Assembly]::LoadFrom((Join-Path $bin $dll)) | Out-Null
}

$assembly = [System.Reflection.Assembly]::LoadFrom((Join-Path $bin "RKPdfAnnotator.dll"))
$tagType = $assembly.GetType("RKPdfAnnotator.TagRecord", $true)
$engineType = $assembly.GetType("RKPdfAnnotator.PdfAnnotationEngine", $true)
$optionsType = $assembly.GetType("RKPdfAnnotator.WatermarkOptions", $true)
$tagMatchingOptionsType = $assembly.GetType("RKPdfAnnotator.TagMatchingOptions", $true)
$rowType = $assembly.GetType("RKPdfAnnotator.RowData", $true)
$sheetType = $assembly.GetType("RKPdfAnnotator.SheetData", $true)
$readerType = $assembly.GetType("RKPdfAnnotator.ExcelSheetReader", $true)

$values = [System.Collections.Generic.Dictionary[string,string]]::new([System.StringComparer]::OrdinalIgnoreCase)
$values["Tag"] = "PUMP-101-A"
$values["Type"] = "Pump"
$values["Blank"] = ""
$values["Location"] = "Inlet"
$row = [System.Activator]::CreateInstance($rowType, @([int]2, $values))
$rowListType = [System.Collections.Generic.List``1].MakeGenericType($rowType)
$rows = [System.Activator]::CreateInstance($rowListType)
$rows.Add($row) | Out-Null
$headers = [System.Collections.Generic.List[string]]::new()
@("Tag", "Type", "Blank", "Location") | ForEach-Object { $headers.Add($_) }
$sheet = [System.Activator]::CreateInstance($sheetType, @("Smoke.xlsx", "Tags", $headers, $rows))
$noteColumns = [System.Collections.Generic.List[string]]::new()
$watermarkColumns = [System.Collections.Generic.List[string]]::new()
@("Type", "Blank", "Tag", "Location") | ForEach-Object { $watermarkColumns.Add($_) }
$builtRecords = $readerType.GetMethod("BuildTagRecords").Invoke($null, @($sheet, "Tag", $noteColumns, $watermarkColumns, $null, $null))
$builtWatermark = $tagType.GetProperty("WatermarkText").GetValue($builtRecords[0])
if ($builtWatermark -ne "Pump / Inlet") {
    throw "Watermark formatting failed: expected 'Pump / Inlet', got '$builtWatermark'."
}

$record = [System.Activator]::CreateInstance($tagType, @("PUMP-101-A", "Description: Smoke test component", "Pump / Inlet", 2, [System.Drawing.Color]::Yellow))
$listType = [System.Collections.Generic.List``1].MakeGenericType($tagType)
$records = [System.Activator]::CreateInstance($listType)
$records.Add($record) | Out-Null

$columnList = [System.Collections.Generic.List[string]]::new()
$columnList.Add("Description")
$watermarkOptions = [System.Activator]::CreateInstance($optionsType, @($true, $columnList, [single]9, [System.Drawing.Color]::DarkBlue, $true))
$tagMatchingOptions = $tagMatchingOptionsType.GetProperty("Default", [System.Reflection.BindingFlags]"Public, Static").GetValue($null)

$method = $engineType.GetMethod("Annotate", [System.Reflection.BindingFlags]"Public, Static")
$parameters = [object[]]@([string]$inputPdf, [string]$outputPdf, $records, $watermarkOptions, $tagMatchingOptions, $null)
$result = $method.Invoke($null, $parameters)
$resultType = $result.GetType()
$totalTags = $resultType.GetProperty("TotalTags").GetValue($result)
$matchedTags = $resultType.GetProperty("MatchedTags").GetValue($result)
$annotations = $resultType.GetProperty("Annotations").GetValue($result)
$watermarks = $resultType.GetProperty("Watermarks").GetValue($result)

[PSCustomObject]@{
    InputPdf = $inputPdf
    OutputPdf = $outputPdf
    Exists = Test-Path $outputPdf
    TotalTags = $totalTags
    MatchedTags = $matchedTags
    Annotations = $annotations
    Watermarks = $watermarks
}

if (-not (Test-Path $outputPdf) -or $totalTags -ne 1 -or $matchedTags -ne 1 -or $annotations -ne 2 -or $watermarks -ne 2) {
    throw "Smoke test failed: expected one matched tag with two annotations and two watermarks."
}

if ($SamplePdf -and (Test-Path $SamplePdf)) {
    $sampleOutputPdf = Join-Path $testRoot "SLK_PID_annotated_test.pdf"
    if (Test-Path $sampleOutputPdf) {
        Remove-Item -LiteralPath $sampleOutputPdf -Force
    }

    $sampleRecords = [System.Activator]::CreateInstance($listType)
    $sampleRecords.Add([System.Activator]::CreateInstance($tagType, @("301.TA.NON.1", "Description: punctuation normalization test", "Non-return / Area 301", 2, [System.Drawing.Color]::Yellow))) | Out-Null
    $sampleRecords.Add([System.Activator]::CreateInstance($tagType, @("201-LM-01", "Description: dash tag sample", "Level monitor", 3, [System.Drawing.Color]::Yellow))) | Out-Null
    $sampleRecords.Add([System.Activator]::CreateInstance($tagType, @("111.LT.CSB1.407", "Description: screenshot placement test", "Level transmitter / CSB1", 4, [System.Drawing.Color]::Yellow))) | Out-Null
    $sampleRecords.Add([System.Activator]::CreateInstance($tagType, @("311.PV.PRL1.201", "Description: landscape placement regression", "LB3051252157", 5, [System.Drawing.Color]::Yellow))) | Out-Null

    $sampleParameters = [object[]]@([string]$SamplePdf, [string]$sampleOutputPdf, $sampleRecords, $watermarkOptions, $tagMatchingOptions, $null)
    $sampleResult = $method.Invoke($null, $sampleParameters)
    $sampleResultType = $sampleResult.GetType()

    [PSCustomObject]@{
        InputPdf = $SamplePdf
        OutputPdf = $sampleOutputPdf
        Exists = Test-Path $sampleOutputPdf
        TotalTags = $sampleResultType.GetProperty("TotalTags").GetValue($sampleResult)
        MatchedTags = $sampleResultType.GetProperty("MatchedTags").GetValue($sampleResult)
        Annotations = $sampleResultType.GetProperty("Annotations").GetValue($sampleResult)
        Watermarks = $sampleResultType.GetProperty("Watermarks").GetValue($sampleResult)
    }
}
