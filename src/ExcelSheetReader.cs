using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using ExcelDna.Integration;

namespace RKPdfAnnotator;

internal sealed class SheetData
{
    public SheetData(string workbookName, string sheetName, IReadOnlyList<string> headers, IReadOnlyList<RowData> rows)
    {
        WorkbookName = workbookName;
        SheetName = sheetName;
        Headers = headers;
        Rows = rows;
    }

    public string WorkbookName { get; }
    public string SheetName { get; }
    public IReadOnlyList<string> Headers { get; }
    public IReadOnlyList<RowData> Rows { get; }
}

internal sealed class RowData
{
    public RowData(int excelRowNumber, IReadOnlyDictionary<string, string> values)
    {
        ExcelRowNumber = excelRowNumber;
        Values = values;
    }

    public int ExcelRowNumber { get; }
    public IReadOnlyDictionary<string, string> Values { get; }
}

internal sealed class TagRecord
{
    public TagRecord(string tag, string note, int rowNumber)
        : this(tag, note, string.Empty, rowNumber, Color.Yellow)
    {
    }

    public TagRecord(string tag, string note, string watermarkText, int rowNumber, Color highlightColor)
    {
        Tag = tag;
        Note = note;
        WatermarkText = watermarkText;
        RowNumber = rowNumber;
        HighlightColor = highlightColor;
    }

    public string Tag { get; }
    public string Note { get; }
    public string WatermarkText { get; }
    public int RowNumber { get; }
    public Color HighlightColor { get; }
}

internal static class ExcelSheetReader
{
    public static SheetData ReadActiveSheet(int headerRow = 1)
    {
        if (headerRow < 1)
            throw new ArgumentOutOfRangeException(nameof(headerRow), "Header row must be 1 or higher.");

        dynamic app = ExcelDnaUtil.Application;
        dynamic workbook = app.ActiveWorkbook ?? throw new InvalidOperationException("No active workbook is open.");
        dynamic sheet = app.ActiveSheet ?? throw new InvalidOperationException("No active worksheet is selected.");
        dynamic usedRange = sheet.UsedRange ?? throw new InvalidOperationException("The active worksheet is empty.");

        object?[,] values = ToArray(usedRange.Value2);
        int rowCount = values.GetLength(0);
        int columnCount = values.GetLength(1);
        if (rowCount < headerRow)
            throw new InvalidOperationException("The active worksheet does not contain the selected header row.");

        var headers = new List<string>();
        var usedNames = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
        for (int column = 1; column <= columnCount; column++)
        {
            string header = Convert.ToString(values[headerRow, column])?.Trim() ?? string.Empty;
            if (string.IsNullOrWhiteSpace(header))
                header = "Column " + column;

            if (usedNames.TryGetValue(header, out int duplicateCount))
            {
                duplicateCount++;
                usedNames[header] = duplicateCount;
                header = header + " " + duplicateCount;
            }
            else
            {
                usedNames[header] = 1;
            }

            headers.Add(header);
        }

        var rows = new List<RowData>();
        for (int row = headerRow + 1; row <= rowCount; row++)
        {
            var dictionary = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            bool hasAnyValue = false;
            for (int column = 1; column <= columnCount; column++)
            {
                string cellValue = Convert.ToString(values[row, column])?.Trim() ?? string.Empty;
                dictionary[headers[column - 1]] = cellValue;
                hasAnyValue |= cellValue.Length > 0;
            }

            if (hasAnyValue)
                rows.Add(new RowData(row, dictionary));
        }

        return new SheetData((string)workbook.Name, (string)sheet.Name, headers, rows);
    }

    public static IReadOnlyList<TagRecord> BuildTagRecords(
        SheetData sheet,
        string tagColumn,
        IReadOnlyCollection<string> noteColumns,
        IReadOnlyCollection<string>? watermarkColumns = null,
        IReadOnlyList<ColorRule>? colorRules = null,
        Color? defaultHighlightColor = null)
    {
        Color defaultColor = defaultHighlightColor ?? Color.Yellow;
        var records = new List<TagRecord>();
        foreach (var row in sheet.Rows)
        {
            if (!row.Values.TryGetValue(tagColumn, out string tag) || string.IsNullOrWhiteSpace(tag))
                continue;

            var noteParts = noteColumns
                .Where(column => row.Values.TryGetValue(column, out string value) && !string.IsNullOrWhiteSpace(value))
                .Select(column => column + ": " + row.Values[column]);

            string note = string.Join(Environment.NewLine, noteParts);
            if (string.IsNullOrWhiteSpace(note))
                note = "Excel row " + row.ExcelRowNumber;

            var watermarkParts = (watermarkColumns ?? Array.Empty<string>())
                .Where(column => !string.Equals(column, tagColumn, StringComparison.OrdinalIgnoreCase))
                .Where(column => row.Values.TryGetValue(column, out string value) && !string.IsNullOrWhiteSpace(value))
                .Select(column => row.Values[column].Trim());

            string watermarkText = string.Join(" / ", watermarkParts);
            string trimmedTag = tag.Trim();
            Color highlightColor = ColorRuleEngine.Apply(trimmedTag, row.Values, colorRules, defaultColor);
            records.Add(new TagRecord(trimmedTag, note, watermarkText, row.ExcelRowNumber, highlightColor));
        }

        return records;
    }

    private static object?[,] ToArray(object? value)
    {
        if (value is object?[,] array)
            return array;

        var singleCell = new object?[2, 2];
        singleCell[1, 1] = value;
        return singleCell;
    }
}
