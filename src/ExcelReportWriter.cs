using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Reflection;
using ExcelDna.Integration;

namespace RKPdfAnnotator;

internal static class ExcelReportWriter
{
    private const int XlContinuous = 1;
    private const int XlCenter = -4108;
    private const int XlLeft = -4131;
    private const int XlTop = -4160;

    public static string WriteToNewSheet(TagReport report)
    {
        dynamic app = ExcelDnaUtil.Application;
        dynamic workbook = app.ActiveWorkbook ?? throw new InvalidOperationException("No active workbook is open.");
        dynamic worksheets = workbook.Worksheets;
        dynamic lastWorksheet = worksheets[worksheets.Count];
        dynamic worksheet = worksheets.Add(Missing.Value, lastWorksheet, 1, Missing.Value);
        string sheetName = UniqueSheetName(workbook, "PDF Tag Report");
        worksheet.Name = sheetName;

        int row = 1;
        WriteTitle(worksheet, row, "PID Annotator - PDF Tag Report", 5);
        row++;
        WriteMetadata(worksheet, row, "Exported: " + report.CreatedAt.ToString("yyyy-MM-dd HH:mm:ss"));
        row += 2;

        row = WriteSection(worksheet, row, "Summary", new[] { "Metric", "Value" },
            report.SummaryRows.Select(item => new object[] { item.Metric, item.Value }));
        row++;

        int firstTableHeaderRow = row + 1;
        row = WriteSection(worksheet, row, "Duplicate Excel Tags", new[] { "Tag", "Excel Rows", "Count" },
            report.DuplicateTags.Select(item => new object[] { item.Tag, item.ExcelRows, item.Count }));
        row++;

        row = WriteSection(worksheet, row, "Found Tags", new[] { "Tag", "Pages", "Occurrences", "Excel Row" },
            report.FoundTags.Select(item => new object[] { item.Tag, item.Pages, item.Occurrences, item.ExcelRow }));
        row++;

        row = WriteSection(worksheet, row, "Excel Tags Not Found In PDF", new[] { "Tag", "Excel Row", "Reason" },
            report.ExcelTagsNotFound.Select(item => new object[] { item.Tag, item.ExcelRow, item.Reason }));
        row++;

        WriteSection(worksheet, row, "Page Statistics", new[] { "Page", "Total PDF Tag Occurrences", "Unique PDF Tags" },
            report.PageStatistics.Select(item => new object[] { item.Page, item.TotalPdfTagOccurrences, item.UniquePdfTags }));

        dynamic usedRange = worksheet.UsedRange;
        usedRange.Columns.AutoFit();
        usedRange.VerticalAlignment = XlTop;

        worksheet.Activate();
        app.ActiveWindow.SplitRow = firstTableHeaderRow - 1;
        app.ActiveWindow.SplitColumn = 0;
        app.ActiveWindow.FreezePanes = true;

        return sheetName;
    }

    private static int WriteSection(dynamic worksheet, int row, string title, IReadOnlyList<string> headers, IEnumerable<object[]> rows)
    {
        WriteSectionTitle(worksheet, row, title, headers.Count);
        row++;

        int headerRow = row;
        for (int column = 1; column <= headers.Count; column++)
            worksheet.Cells[row, column].Value2 = headers[column - 1];

        FormatHeader(worksheet.Range[worksheet.Cells[row, 1], worksheet.Cells[row, headers.Count]]);
        row++;

        int firstDataRow = row;
        int dataRows = 0;
        foreach (object[] values in rows)
        {
            for (int column = 1; column <= headers.Count; column++)
                worksheet.Cells[row, column].Value2 = column <= values.Length ? values[column - 1] : string.Empty;

            dataRows++;
            row++;
        }

        if (dataRows == 0)
        {
            worksheet.Cells[row, 1].Value2 = "No rows";
            worksheet.Range[worksheet.Cells[row, 1], worksheet.Cells[row, headers.Count]].Merge();
            worksheet.Cells[row, 1].Font.Italic = true;
            worksheet.Cells[row, 1].Font.Color = ToOle(Color.FromArgb(107, 114, 128));
            dataRows = 1;
            row++;
        }

        dynamic tableRange = worksheet.Range[worksheet.Cells[headerRow, 1], worksheet.Cells[row - 1, headers.Count]];
        tableRange.Borders.LineStyle = XlContinuous;
        tableRange.WrapText = false;
        tableRange.VerticalAlignment = XlTop;

        if (headers.Count > 0 && row > firstDataRow)
        {
            dynamic tagRange = worksheet.Range[worksheet.Cells[firstDataRow, 1], worksheet.Cells[row - 1, 1]];
            tagRange.Font.Name = "Courier New";
            tagRange.Font.Bold = true;
            tagRange.Font.Color = ToOle(Color.FromArgb(37, 99, 235));
        }

        return row;
    }

    private static void WriteTitle(dynamic worksheet, int row, string title, int columns)
    {
        dynamic range = worksheet.Range[worksheet.Cells[row, 1], worksheet.Cells[row, columns]];
        range.Merge();
        range.Value2 = title;
        range.Font.Bold = true;
        range.Font.Size = 14;
        range.Font.Color = ToOle(Color.FromArgb(37, 99, 235));
        range.HorizontalAlignment = XlCenter;
    }

    private static void WriteMetadata(dynamic worksheet, int row, string text)
    {
        dynamic range = worksheet.Range[worksheet.Cells[row, 1], worksheet.Cells[row, 5]];
        range.Merge();
        range.Value2 = text;
        range.Font.Size = 10;
        range.Font.Color = ToOle(Color.FromArgb(107, 114, 128));
        range.HorizontalAlignment = XlCenter;
    }

    private static void WriteSectionTitle(dynamic worksheet, int row, string title, int columns)
    {
        dynamic range = worksheet.Range[worksheet.Cells[row, 1], worksheet.Cells[row, Math.Max(1, columns)]];
        range.Merge();
        range.Value2 = title;
        range.Font.Bold = true;
        range.Font.Size = 11;
        range.Interior.Color = ToOle(Color.FromArgb(243, 244, 246));
        range.HorizontalAlignment = XlLeft;
    }

    private static void FormatHeader(dynamic range)
    {
        range.Interior.Color = ToOle(Color.FromArgb(37, 99, 235));
        range.Font.Bold = true;
        range.Font.Color = ToOle(Color.White);
        range.HorizontalAlignment = XlCenter;
        range.VerticalAlignment = XlTop;
    }

    private static string UniqueSheetName(dynamic workbook, string baseName)
    {
        string trimmedBaseName = baseName.Length > 31 ? baseName.Substring(0, 31) : baseName;
        if (!SheetExists(workbook, trimmedBaseName))
            return trimmedBaseName;

        for (int i = 2; i < 1000; i++)
        {
            string suffix = " " + i;
            int maxBaseLength = 31 - suffix.Length;
            string candidate = trimmedBaseName.Length > maxBaseLength
                ? trimmedBaseName.Substring(0, maxBaseLength) + suffix
                : trimmedBaseName + suffix;
            if (!SheetExists(workbook, candidate))
                return candidate;
        }

        throw new InvalidOperationException("Could not create a unique report worksheet name.");
    }

    private static bool SheetExists(dynamic workbook, string sheetName)
    {
        try
        {
            _ = workbook.Worksheets[sheetName];
            return true;
        }
        catch
        {
            return false;
        }
    }

    private static int ToOle(Color color)
        => ColorTranslator.ToOle(color);
}
