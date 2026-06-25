using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using PdfSharp.Drawing;
using PdfSharp.Pdf;
using PdfSharp.Pdf.Annotations;
using PdfSharp.Pdf.IO;
using UglyToad.PdfPig;
using UglyToad.PdfPig.Content;

namespace RKPdfAnnotator;

internal sealed class AnnotationResult
{
    public AnnotationResult(
        int totalTags,
        int matchedTags,
        int annotations,
        int watermarks,
        int pdfTags,
        int unmatchedPdfTags,
        int excelTagsNotFound,
        string reportPath,
        TagReport report)
    {
        TotalTags = totalTags;
        MatchedTags = matchedTags;
        Annotations = annotations;
        Watermarks = watermarks;
        PdfTags = pdfTags;
        UnmatchedPdfTags = unmatchedPdfTags;
        ExcelTagsNotFound = excelTagsNotFound;
        ReportPath = reportPath;
        Report = report;
    }

    public int TotalTags { get; }
    public int MatchedTags { get; }
    public int Annotations { get; }
    public int Watermarks { get; }
    public int PdfTags { get; }
    public int UnmatchedPdfTags { get; }
    public int ExcelTagsNotFound { get; }
    public string ReportPath { get; }
    public TagReport Report { get; }
}

internal sealed class TagReport
{
    public TagReport(
        string inputPdfPath,
        string outputPdfPath,
        string csvReportPath,
        DateTime createdAt,
        IReadOnlyList<TagReportSummaryRow> summaryRows,
        IReadOnlyList<FoundTagReportRow> foundTags,
        IReadOnlyList<ExcelNotFoundReportRow> excelTagsNotFound,
        IReadOnlyList<PdfNotInExcelReportRow> pdfTagsNotInExcel,
        IReadOnlyList<DuplicateTagReportRow> duplicateTags,
        IReadOnlyList<PageStatisticReportRow> pageStatistics)
    {
        InputPdfPath = inputPdfPath;
        OutputPdfPath = outputPdfPath;
        CsvReportPath = csvReportPath;
        CreatedAt = createdAt;
        SummaryRows = summaryRows;
        FoundTags = foundTags;
        ExcelTagsNotFound = excelTagsNotFound;
        PdfTagsNotInExcel = pdfTagsNotInExcel;
        DuplicateTags = duplicateTags;
        PageStatistics = pageStatistics;
    }

    public string InputPdfPath { get; }
    public string OutputPdfPath { get; }
    public string CsvReportPath { get; }
    public DateTime CreatedAt { get; }
    public IReadOnlyList<TagReportSummaryRow> SummaryRows { get; }
    public IReadOnlyList<FoundTagReportRow> FoundTags { get; }
    public IReadOnlyList<ExcelNotFoundReportRow> ExcelTagsNotFound { get; }
    public IReadOnlyList<PdfNotInExcelReportRow> PdfTagsNotInExcel { get; }
    public IReadOnlyList<DuplicateTagReportRow> DuplicateTags { get; }
    public IReadOnlyList<PageStatisticReportRow> PageStatistics { get; }
}

internal sealed class TagReportSummaryRow
{
    public TagReportSummaryRow(string metric, string value)
    {
        Metric = metric;
        Value = value;
    }

    public string Metric { get; }
    public string Value { get; }
}

internal sealed class FoundTagReportRow
{
    public FoundTagReportRow(string tag, string pages, int occurrences, int excelRow)
    {
        Tag = tag;
        Pages = pages;
        Occurrences = occurrences;
        ExcelRow = excelRow;
    }

    public string Tag { get; }
    public string Pages { get; }
    public int Occurrences { get; }
    public int ExcelRow { get; }
}

internal sealed class ExcelNotFoundReportRow
{
    public ExcelNotFoundReportRow(string tag, int excelRow, string reason)
    {
        Tag = tag;
        ExcelRow = excelRow;
        Reason = reason;
    }

    public string Tag { get; }
    public int ExcelRow { get; }
    public string Reason { get; }
}

internal sealed class PdfNotInExcelReportRow
{
    public PdfNotInExcelReportRow(string tag, string pages, int occurrences)
    {
        Tag = tag;
        Pages = pages;
        Occurrences = occurrences;
    }

    public string Tag { get; }
    public string Pages { get; }
    public int Occurrences { get; }
}

internal sealed class DuplicateTagReportRow
{
    public DuplicateTagReportRow(string tag, string excelRows, int count)
    {
        Tag = tag;
        ExcelRows = excelRows;
        Count = count;
    }

    public string Tag { get; }
    public string ExcelRows { get; }
    public int Count { get; }
}

internal sealed class PageStatisticReportRow
{
    public PageStatisticReportRow(int page, int totalPdfTagOccurrences, int uniquePdfTags)
    {
        Page = page;
        TotalPdfTagOccurrences = totalPdfTagOccurrences;
        UniquePdfTags = uniquePdfTags;
    }

    public int Page { get; }
    public int TotalPdfTagOccurrences { get; }
    public int UniquePdfTags { get; }
}

internal sealed class AnnotationProgress
{
    public AnnotationProgress(int percent, string message)
    {
        Percent = Math.Max(0, Math.Min(100, percent));
        Message = message;
    }

    public int Percent { get; }
    public string Message { get; }
}

internal sealed class TagMatchingOptions
{
    public TagMatchingOptions(
        int minParts,
        int maxParts,
        string separators,
        int minPartLength = 1,
        int maxPartLength = 5)
    {
        MinParts = Math.Max(1, minParts);
        MaxParts = Math.Max(MinParts, maxParts);
        Separators = string.IsNullOrWhiteSpace(separators) ? "-." : separators;
        MinPartLength = Math.Max(1, minPartLength);
        MaxPartLength = Math.Max(MinPartLength, maxPartLength);
    }

    public int MinParts { get; }
    public int MaxParts { get; }
    public string Separators { get; }
    public int MinPartLength { get; }
    public int MaxPartLength { get; }

    public static TagMatchingOptions Default { get; } = new(3, 5, "-.");

    public IReadOnlyList<char> SeparatorChars
        => Separators
            .Where(c => !char.IsWhiteSpace(c))
            .Distinct()
            .ToList();
}

internal static class PdfAnnotationEngine
{
    public static AnnotationResult Annotate(
        string inputPdfPath,
        string outputPdfPath,
        IReadOnlyList<TagRecord> records,
        WatermarkOptions watermarkOptions,
        TagMatchingOptions? tagMatchingOptions = null,
        IProgress<AnnotationProgress>? progress = null)
    {
        if (!File.Exists(inputPdfPath))
            throw new FileNotFoundException("Input PDF was not found.", inputPdfPath);

        tagMatchingOptions ??= TagMatchingOptions.Default;
        Directory.CreateDirectory(Path.GetDirectoryName(outputPdfPath) ?? Environment.CurrentDirectory);

        var recordsByTag = records
            .SelectMany(record => GetTagVariants(record.Tag, tagMatchingOptions)
                .Select(variant => new { Variant = variant, Record = record }))
            .GroupBy(item => item.Variant, StringComparer.OrdinalIgnoreCase)
            .ToDictionary(group => group.Key, group => group.First().Record, StringComparer.OrdinalIgnoreCase);
        var matchedTags = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var matchedCanonicalTags = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var annotatedLocations = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        int annotationCount = 0;
        int watermarkCount = 0;

        progress?.Report(new AnnotationProgress(0, "Indexing PDF tags..."));
        PdfTagIndex tagIndex = BuildTagIndex(inputPdfPath, tagMatchingOptions, progress);
        progress?.Report(new AnnotationProgress(45, "Annotating matched PDF tags..."));

        using PdfSharp.Pdf.PdfDocument outputDocument = PdfReader.Open(inputPdfPath, PdfDocumentOpenMode.Modify);

        int totalPages = outputDocument.PageCount;
        for (int pageNumber = 1; pageNumber <= totalPages; pageNumber++)
        {
            PdfPage outputPage = outputDocument.Pages[pageNumber - 1];
            var watermarkPlacements = new List<WatermarkPlacement>();
            var pageOccurrences = tagIndex.Occurrences
                .Where(occurrence => occurrence.PageNumber == pageNumber)
                .ToList();

            foreach (PdfTagOccurrence occurrence in pageOccurrences)
            {
                TagRecord? record = null;
                foreach (string variant in GetTagVariants(occurrence.OriginalTag, tagMatchingOptions))
                {
                    if (recordsByTag.TryGetValue(variant, out record))
                        break;
                }

                if (record == null)
                    continue;

                var box = occurrence.Box;
                string locationKey = pageNumber + "|" + record.Tag.ToUpperInvariant() + "|" +
                                     Math.Round(box.Left, 2) + "|" + Math.Round(box.Bottom, 2) + "|" +
                                     Math.Round(box.Right, 2) + "|" + Math.Round(box.Top, 2);
                if (!annotatedLocations.Add(locationKey))
                    continue;

                PdfBox annotationBox = ToPdfAnnotationBox(outputPage, ToVisiblePageBox(outputPage, box));
                XColor highlightColor = XColor.FromArgb(record.HighlightColor.R, record.HighlightColor.G, record.HighlightColor.B);
                if (!AddHighlightPopupAnnotation(outputPage, annotationBox, record.Tag, record.Note, highlightColor))
                    AddNoteAnnotation(outputPage, annotationBox.Right + 2, annotationBox.Bottom, record.Tag, record.Note, highlightColor);

                if (watermarkOptions.Enabled && !string.IsNullOrWhiteSpace(record.WatermarkText))
                    watermarkPlacements.Add(new WatermarkPlacement(record.WatermarkText, ToGraphicsBox(outputPage, box)));

                matchedTags.Add(record.Tag);
                matchedCanonicalTags.Add(CanonicalTagKey(record.Tag, tagMatchingOptions));
                annotationCount++;
            }

            watermarkCount += DrawWatermarks(outputPage, watermarkPlacements, watermarkOptions);
            int percent = 45 + (int)Math.Round(pageNumber / (double)Math.Max(1, totalPages) * 45);
            progress?.Report(new AnnotationProgress(percent, $"Annotating page {pageNumber}/{totalPages}..."));
        }

        progress?.Report(new AnnotationProgress(92, "Saving annotated PDF..."));
        outputDocument.Save(outputPdfPath);
        int totalTags = records
            .Select(record => CanonicalTagKey(record.Tag, tagMatchingOptions))
            .Where(tag => tag.Length > 0)
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .Count();
        var excelTagsNotFound = records
            .Where(record => NormalizeTag(record.Tag).Length > 0)
            .GroupBy(record => CanonicalTagKey(record.Tag, tagMatchingOptions), StringComparer.OrdinalIgnoreCase)
            .Where(group => !matchedCanonicalTags.Contains(group.Key))
            .Select(group => group.First())
            .OrderBy(record => record.Tag, StringComparer.OrdinalIgnoreCase)
            .ToList();
        var unmatchedPdfTags = BuildUnmatchedPdfTags(tagIndex, recordsByTag.Keys, tagMatchingOptions);
        progress?.Report(new AnnotationProgress(96, "Writing tag report..."));
        string reportPath = GetCsvReportPath(outputPdfPath);
        TagReport report = BuildReport(
            inputPdfPath,
            outputPdfPath,
            reportPath,
            tagIndex,
            tagMatchingOptions,
            records,
            matchedTags,
            totalTags,
            annotationCount,
            watermarkCount,
            excelTagsNotFound,
            unmatchedPdfTags);
        WriteReport(reportPath, report);
        progress?.Report(new AnnotationProgress(100, "Annotation complete. Report created."));
        return new AnnotationResult(
            totalTags,
            matchedTags.Count,
            annotationCount,
            watermarkCount,
            tagIndex.UniqueTagCount(tagMatchingOptions),
            unmatchedPdfTags.Count,
            excelTagsNotFound.Count,
            reportPath,
            report);
    }

    private static PdfTagIndex BuildTagIndex(
        string inputPdfPath,
        TagMatchingOptions options,
        IProgress<AnnotationProgress>? progress)
    {
        var tagIndex = new PdfTagIndex();
        Regex tagPattern = BuildTagRegex(options);

        using UglyToad.PdfPig.PdfDocument sourceTextDocument = UglyToad.PdfPig.PdfDocument.Open(inputPdfPath);
        int totalPages = sourceTextDocument.NumberOfPages;
        foreach (Page sourcePage in sourceTextDocument.GetPages())
        {
            foreach (TextCandidate candidate in GetTextCandidates(sourcePage))
            {
                string text = CleanPdfWord(candidate.Text);
                if (text.Length == 0)
                    continue;

                foreach (Match match in tagPattern.Matches(text).Cast<Match>())
                {
                    string tag = NormalizeTag(match.Value);
                    if (tag.Length == 0 || !IsValidTag(tag, options))
                        continue;

                    tagIndex.Add(new PdfTagOccurrence(tag, sourcePage.Number, candidate.Box), options);
                }
            }

            int percent = (int)Math.Round(sourcePage.Number / (double)Math.Max(1, totalPages) * 40);
            progress?.Report(new AnnotationProgress(percent, $"Indexing page {sourcePage.Number}/{totalPages}..."));
        }

        return tagIndex;
    }

    private static Regex BuildTagRegex(TagMatchingOptions options)
    {
        string separators = string.Join(string.Empty, options.SeparatorChars.Select(separator => Regex.Escape(separator.ToString())));
        if (string.IsNullOrEmpty(separators))
            separators = Regex.Escape("-");

        string separatorPattern = "[" + separators + "]";
        string partPattern = $"[A-Za-z0-9]{{{options.MinPartLength},{options.MaxPartLength}}}";
        string required = string.Join(separatorPattern, Enumerable.Repeat(partPattern, options.MinParts));
        string optional = options.MaxParts > options.MinParts
            ? $"(?:{separatorPattern}{partPattern}){{0,{options.MaxParts - options.MinParts}}}"
            : string.Empty;
        return new Regex(@"\b" + required + optional + @"\b", RegexOptions.Compiled | RegexOptions.IgnoreCase);
    }

    private static bool IsValidTag(string tag, TagMatchingOptions options)
    {
        var parts = SplitTagParts(tag, options);
        return parts.Count >= options.MinParts &&
               parts.Count <= options.MaxParts &&
               parts.All(part => part.Length >= options.MinPartLength && part.Length <= options.MaxPartLength);
    }

    private static IReadOnlyList<UnmatchedPdfTag> BuildUnmatchedPdfTags(
        PdfTagIndex tagIndex,
        IEnumerable<string> excelTagVariants,
        TagMatchingOptions options)
    {
        var excelLookup = new HashSet<string>(
            excelTagVariants.Select(variant => CanonicalTagKey(variant, options)).Where(key => key.Length > 0),
            StringComparer.OrdinalIgnoreCase);
        return tagIndex.Occurrences
            .GroupBy(occurrence => CanonicalTagKey(occurrence.NormalizedTag, options), StringComparer.OrdinalIgnoreCase)
            .Where(group => !excelLookup.Contains(group.Key))
            .Select(group => new UnmatchedPdfTag(
                group.Select(occurrence => occurrence.OriginalTag).OrderBy(tag => tag, StringComparer.OrdinalIgnoreCase).First(),
                group.Select(occurrence => occurrence.PageNumber).Distinct().OrderBy(page => page).ToList(),
                group.Count()))
            .OrderBy(tag => tag.Tag, StringComparer.OrdinalIgnoreCase)
            .ToList();
    }

    private static TagReport BuildReport(
        string inputPdfPath,
        string outputPdfPath,
        string reportPath,
        PdfTagIndex tagIndex,
        TagMatchingOptions options,
        IReadOnlyList<TagRecord> records,
        IReadOnlyCollection<string> matchedTags,
        int totalTags,
        int annotationCount,
        int watermarkCount,
        IReadOnlyList<TagRecord> excelTagsNotFound,
        IReadOnlyList<UnmatchedPdfTag> unmatchedPdfTags)
    {
        var foundRows = matchedTags
            .Select(tag => new FoundTagReportRow(
                tag,
                PagesForTag(tagIndex, tag, options),
                OccurrenceCountForTag(tagIndex, tag, options),
                records.FirstOrDefault(record => string.Equals(record.Tag, tag, StringComparison.OrdinalIgnoreCase))?.RowNumber ?? 0))
            .OrderBy(row => row.Tag, StringComparer.OrdinalIgnoreCase)
            .ToList();
        var excelNotFoundRows = excelTagsNotFound
            .Select(record => new ExcelNotFoundReportRow(record.Tag, record.RowNumber, "not_in_pdf"))
            .ToList();
        var pdfNotInExcelRows = unmatchedPdfTags
            .Select(tag => new PdfNotInExcelReportRow(tag.Tag, string.Join(";", tag.Pages), tag.OccurrenceCount))
            .ToList();
        var duplicateRows = records
            .Where(record => NormalizeTag(record.Tag).Length > 0)
            .GroupBy(record => CanonicalTagKey(record.Tag, options), StringComparer.OrdinalIgnoreCase)
            .Where(group => group.Count() > 1)
            .Select(group => new DuplicateTagReportRow(
                group.First().Tag,
                string.Join(";", group.Select(record => record.RowNumber).OrderBy(row => row)),
                group.Count()))
            .OrderBy(row => row.Tag, StringComparer.OrdinalIgnoreCase)
            .ToList();
        var pageStatisticRows = tagIndex.Occurrences
            .GroupBy(occurrence => occurrence.PageNumber)
            .OrderBy(group => group.Key)
            .Select(group => new PageStatisticReportRow(
                group.Key,
                group.Count(),
                group.Select(occurrence => CanonicalTagKey(occurrence.NormalizedTag, options)).Distinct(StringComparer.OrdinalIgnoreCase).Count()))
            .ToList();
        var summaryRows = new List<TagReportSummaryRow>
        {
            new("Input PDF", inputPdfPath),
            new("Output PDF", outputPdfPath),
            new("CSV Report", reportPath),
            new("Excel Tags", totalTags.ToString()),
            new("Matched Excel Tags", foundRows.Count.ToString()),
            new("PDF Tags Indexed", tagIndex.UniqueTagCount(options).ToString()),
            new("Annotations Added", annotationCount.ToString()),
            new("Watermarks Added", watermarkCount.ToString()),
            new("Excel Tags Not Found In PDF", excelNotFoundRows.Count.ToString()),
            new("Duplicate Excel Tags", duplicateRows.Count.ToString()),
            new("Tag Min Parts", options.MinParts.ToString()),
            new("Tag Max Parts", options.MaxParts.ToString()),
            new("Tag Separators", options.Separators)
        };

        return new TagReport(
            inputPdfPath,
            outputPdfPath,
            reportPath,
            DateTime.Now,
            summaryRows,
            foundRows,
            excelNotFoundRows,
            pdfNotInExcelRows,
            duplicateRows,
            pageStatisticRows);
    }

    private static string GetCsvReportPath(string outputPdfPath)
        => Path.Combine(
            Path.GetDirectoryName(outputPdfPath) ?? Environment.CurrentDirectory,
            Path.GetFileNameWithoutExtension(outputPdfPath) + "_tag_report.csv");

    private static void WriteReport(string reportPath, TagReport report)
    {
        using var writer = new StreamWriter(reportPath, false, System.Text.Encoding.UTF8);

        writer.WriteLine("Section,Tag,Pages,Occurrences,Excel Row,Status");
        foreach (DuplicateTagReportRow row in report.DuplicateTags)
            writer.WriteLine(ToCsv("Duplicate Excel Tags", row.Tag, string.Empty, row.Count.ToString(), row.ExcelRows, "Duplicate tag rows in Excel"));

        foreach (FoundTagReportRow row in report.FoundTags)
            writer.WriteLine(ToCsv("Annotated", row.Tag, row.Pages, row.Occurrences.ToString(), row.ExcelRow.ToString(), "PDF tag matched Excel"));

        foreach (ExcelNotFoundReportRow row in report.ExcelTagsNotFound)
            writer.WriteLine(ToCsv("Excel tag not found in PDF", row.Tag, string.Empty, "0", row.ExcelRow.ToString(), "Excel row has no matching PDF tag"));

        writer.WriteLine();
        writer.WriteLine("Page,Total PDF Tag Occurrences,Unique PDF Tags");
        foreach (PageStatisticReportRow row in report.PageStatistics)
            writer.WriteLine(ToCsv(row.Page.ToString(), row.TotalPdfTagOccurrences.ToString(), row.UniquePdfTags.ToString()));
    }

    private static string PagesForTag(PdfTagIndex tagIndex, string tag, TagMatchingOptions options)
    {
        var variants = new HashSet<string>(GetTagVariants(tag, options), StringComparer.OrdinalIgnoreCase);
        return string.Join(";", tagIndex.Occurrences
            .Where(occurrence => variants.Contains(occurrence.NormalizedTag))
            .Select(occurrence => occurrence.PageNumber)
            .Distinct()
            .OrderBy(page => page));
    }

    private static int OccurrenceCountForTag(PdfTagIndex tagIndex, string tag, TagMatchingOptions options)
    {
        var variants = new HashSet<string>(GetTagVariants(tag, options), StringComparer.OrdinalIgnoreCase);
        return tagIndex.Occurrences.Count(occurrence => variants.Contains(occurrence.NormalizedTag));
    }

    private static string ToCsv(params string[] values)
        => string.Join(",", values.Select(value => "\"" + (value ?? string.Empty).Replace("\"", "\"\"") + "\""));

    private static GraphicsBox ToGraphicsBox(PdfPage page, UglyToad.PdfPig.Core.PdfRectangle box)
    {
        int rotation = NormalizeRotation(page.Rotate);
        double mediaWidth = Math.Abs(page.MediaBox.X2 - page.MediaBox.X1);
        double mediaHeight = Math.Abs(page.MediaBox.Y2 - page.MediaBox.Y1);
        double visibleHeight = rotation == 90 || rotation == 270 ? mediaWidth : mediaHeight;
        var visible = new GraphicsBox(
            Math.Min(box.Left, box.Right),
            visibleHeight - Math.Max(box.Top, box.Bottom),
            Math.Max(box.Left, box.Right),
            visibleHeight - Math.Min(box.Top, box.Bottom));

        return rotation switch
        {
            90 => new GraphicsBox(
                visible.Top,
                mediaHeight - visible.Right,
                visible.Bottom,
                mediaHeight - visible.Left),
            180 => new GraphicsBox(
                mediaWidth - visible.Right,
                mediaHeight - visible.Bottom,
                mediaWidth - visible.Left,
                mediaHeight - visible.Top),
            270 => new GraphicsBox(
                mediaWidth - visible.Bottom,
                visible.Left,
                mediaWidth - visible.Top,
                visible.Right),
            _ => visible
        };
    }

    private static int DrawWatermarks(
        PdfPage page,
        IReadOnlyCollection<WatermarkPlacement> placements,
        WatermarkOptions options)
    {
        if (!options.Enabled || placements.Count == 0)
            return 0;

        double mediaWidth = Math.Abs(page.MediaBox.X2 - page.MediaBox.X1);
        double mediaHeight = Math.Abs(page.MediaBox.Y2 - page.MediaBox.Y1);
        double pageWidth = mediaWidth;
        double pageHeight = mediaHeight;
        int pageRotation = page.Rotate;
        int normalizedRotation = NormalizeRotation(pageRotation);
        PdfRectangle mediaBox = page.MediaBox;
        page.Rotate = 0;
        page.Width = XUnit.FromPoint(mediaWidth);
        page.Height = XUnit.FromPoint(mediaHeight);
        try
        {
            using XGraphics graphics = XGraphics.FromPdfPage(page, XGraphicsPdfPageOptions.Append);
            var font = new XFont("Arial", options.FontSize, XFontStyle.Regular);
            var textColor = XColor.FromArgb(options.TextColor.R, options.TextColor.G, options.TextColor.B);
            var textBrush = new XSolidBrush(textColor);

            foreach (WatermarkPlacement placement in placements)
            {
                XSize textSize = graphics.MeasureString(placement.Text, font);
                bool horizontalTag = placement.Box.Width >= placement.Box.Height;
                const double gap = 3;
                const double edge = 2;
                const double padding = 1;

                graphics.Save();
                if (horizontalTag)
                {
                    double x = Clamp(placement.Box.CenterX - (textSize.Width / 2), edge, Math.Max(edge, pageWidth - textSize.Width - edge));
                    double desiredY = normalizedRotation == 180
                        ? placement.Box.Top - gap - textSize.Height
                        : placement.Box.Bottom + gap;
                    double y = Clamp(desiredY, edge, Math.Max(edge, pageHeight - textSize.Height - edge));

                    if (options.BackgroundEnabled)
                        graphics.DrawRectangle(XBrushes.White, x - padding, y - padding, textSize.Width + (padding * 2), textSize.Height + (padding * 2));

                    graphics.DrawString(placement.Text, font, textBrush, new XRect(x, y, textSize.Width, textSize.Height), XStringFormats.TopLeft);
                }
                else
                {
                    // Clockwise text extends left of its anchor. On 270-degree pages,
                    // screen-down points toward decreasing raw x; on 90-degree pages it
                    // points toward increasing raw x.
                    double desiredX = normalizedRotation == 270
                        ? placement.Box.Left - gap
                        : placement.Box.Right + gap + textSize.Height;
                    double x = Clamp(desiredX, edge + textSize.Height, Math.Max(edge + textSize.Height, pageWidth - edge));
                    double y = Clamp(placement.Box.CenterY - (textSize.Width / 2), edge, Math.Max(edge, pageHeight - textSize.Width - edge));
                    graphics.TranslateTransform(x, y, XMatrixOrder.Append);
                    graphics.RotateTransform(90, XMatrixOrder.Append);

                    if (options.BackgroundEnabled)
                        graphics.DrawRectangle(XBrushes.White, -padding, -padding, textSize.Width + (padding * 2), textSize.Height + (padding * 2));

                    graphics.DrawString(placement.Text, font, textBrush, new XRect(0, 0, textSize.Width, textSize.Height), XStringFormats.TopLeft);
                }
                graphics.Restore();
            }
        }
        finally
        {
            page.MediaBox = mediaBox;
            page.Rotate = pageRotation;
        }

        return placements.Count;
    }

    private static double Clamp(double value, double minimum, double maximum)
        => Math.Max(minimum, Math.Min(maximum, value));

    private static VisibleBox ToVisiblePageBox(PdfPage page, UglyToad.PdfPig.Core.PdfRectangle box)
    {
        int rotation = NormalizeRotation(page.Rotate);
        double pageWidth = page.Width.Point;
        double pageHeight = page.Height.Point;

        return rotation switch
        {
            90 => NormalizeVisibleBox(pageWidth - box.Top, pageHeight - box.Right, pageWidth - box.Bottom, pageHeight - box.Left),
            180 => NormalizeVisibleBox(pageWidth - box.Right, box.Bottom, pageWidth - box.Left, box.Top),
            270 => NormalizeVisibleBox(box.Bottom, box.Left, box.Top, box.Right),
            _ => NormalizeVisibleBox(box.Left, pageHeight - box.Top, box.Right, pageHeight - box.Bottom)
        };
    }

    private static PdfBox ToPdfAnnotationBox(PdfPage page, VisibleBox visibleBox)
    {
        double mediaHeight = Math.Abs(page.MediaBox.Y2 - page.MediaBox.Y1);
        return new PdfBox(
            visibleBox.Left,
            mediaHeight - visibleBox.Bottom,
            visibleBox.Right,
            mediaHeight - visibleBox.Top);
    }

    private static int NormalizeRotation(int rotation)
    {
        rotation %= 360;
        return rotation < 0 ? rotation + 360 : rotation;
    }

    private static VisibleBox NormalizeVisibleBox(double left, double top, double right, double bottom)
    {
        return new VisibleBox(
            Math.Min(left, right),
            Math.Min(top, bottom),
            Math.Max(left, right),
            Math.Max(top, bottom));
    }

    private static IReadOnlyCollection<string> GetTagVariants(string tag, TagMatchingOptions options)
    {
        string normalized = NormalizeTag(tag);
        if (normalized.Length == 0)
            return Array.Empty<string>();

        var variants = new HashSet<string>(StringComparer.OrdinalIgnoreCase) { normalized };
        var parts = SplitTagParts(normalized, options);
        if (parts.Count > 1)
        {
            foreach (char separator in options.SeparatorChars)
                variants.Add(string.Join(separator.ToString(), parts));
        }

        return variants;
    }

    private static string CanonicalTagKey(string tag, TagMatchingOptions options)
    {
        string normalized = NormalizeTag(tag);
        if (normalized.Length == 0)
            return string.Empty;

        var parts = SplitTagParts(normalized, options);
        return parts.Count > 1 ? string.Join("\u001F", parts) : normalized;
    }

    private static IReadOnlyList<string> SplitTagParts(string tag, TagMatchingOptions options)
    {
        char[] separators = options.SeparatorChars.ToArray();
        if (separators.Length == 0)
            separators = new[] { '-' };

        return tag.Split(separators, StringSplitOptions.RemoveEmptyEntries)
            .Select(part => part.Trim())
            .Where(part => part.Length > 0)
            .ToList();
    }

    private static IReadOnlyList<TextCandidate> GetTextCandidates(Page page)
    {
        var candidates = page.GetWords()
            .Select(word => new TextCandidate(word.Text, word.BoundingBox))
            .ToList();

        var token = new List<Letter>();
        foreach (Letter letter in page.Letters)
        {
            if (string.IsNullOrWhiteSpace(letter.Value))
            {
                AddLetterToken(candidates, token);
                token.Clear();
            }
            else
            {
                token.Add(letter);
            }
        }
        AddLetterToken(candidates, token);
        AddOrientedLetterRuns(candidates, page.Letters);

        return candidates;
    }

    private static void AddOrientedLetterRuns(ICollection<TextCandidate> candidates, IReadOnlyList<Letter> letters)
    {
        var rotatedLetters = letters
            .Where(letter => !string.IsNullOrWhiteSpace(letter.Value))
            .Where(letter => letter.TextOrientation != TextOrientation.Horizontal)
            .GroupBy(letter => letter.TextOrientation);

        foreach (var orientationGroup in rotatedLetters)
        {
            var groupLetters = orientationGroup.ToList();
            double lineTolerance = Math.Max(2, MedianPointSize(groupLetters) * 0.75);
            foreach (List<Letter> line in ClusterLettersIntoLines(groupLetters, orientationGroup.Key, lineTolerance))
            {
                List<Letter> orderedLine = OrderLettersForReading(line, orientationGroup.Key).ToList();
                AddLineTokens(candidates, orderedLine);
            }
        }
    }

    private static IEnumerable<List<Letter>> ClusterLettersIntoLines(
        IReadOnlyCollection<Letter> letters,
        TextOrientation orientation,
        double tolerance)
    {
        var clusters = new List<List<Letter>>();
        foreach (Letter letter in letters.OrderBy(letter => PerpendicularCenter(letter, orientation)))
        {
            double center = PerpendicularCenter(letter, orientation);
            List<Letter>? current = clusters.LastOrDefault();
            if (current == null || Math.Abs(center - current.Average(item => PerpendicularCenter(item, orientation))) > tolerance)
            {
                clusters.Add(new List<Letter> { letter });
            }
            else
            {
                current.Add(letter);
            }
        }

        return clusters.Where(cluster => cluster.Count > 1);
    }

    private static IEnumerable<Letter> OrderLettersForReading(IEnumerable<Letter> letters, TextOrientation orientation)
        => orientation switch
        {
            TextOrientation.Rotate180 => letters.OrderByDescending(letter => LetterCenterX(letter)),
            TextOrientation.Rotate90 => letters.OrderByDescending(letter => LetterCenterY(letter)),
            TextOrientation.Rotate270 => letters.OrderBy(letter => LetterCenterY(letter)),
            _ => letters.OrderBy(letter => letter.TextSequence)
        };

    private static void AddLineTokens(ICollection<TextCandidate> candidates, IReadOnlyList<Letter> letters)
    {
        if (letters.Count == 0)
            return;

        double maxGap = Math.Max(3, MedianPointSize(letters) * 1.8);
        var token = new List<Letter>();
        Letter? previous = null;
        foreach (Letter letter in letters)
        {
            if (previous != null && InlineGap(previous, letter) > maxGap)
            {
                AddLetterToken(candidates, token);
                token.Clear();
            }

            token.Add(letter);
            previous = letter;
        }

        AddLetterToken(candidates, token);
    }

    private static double InlineGap(Letter previous, Letter current)
        => Math.Abs(
            IsVertical(previous.TextOrientation)
                ? LetterCenterY(current) - LetterCenterY(previous)
                : LetterCenterX(current) - LetterCenterX(previous));

    private static double PerpendicularCenter(Letter letter, TextOrientation orientation)
        => IsVertical(orientation) ? LetterCenterX(letter) : LetterCenterY(letter);

    private static bool IsVertical(TextOrientation orientation)
        => orientation == TextOrientation.Rotate90 || orientation == TextOrientation.Rotate270;

    private static double LetterCenterX(Letter letter)
        => (letter.BoundingBox.Left + letter.BoundingBox.Right) / 2;

    private static double LetterCenterY(Letter letter)
        => (letter.BoundingBox.Bottom + letter.BoundingBox.Top) / 2;

    private static double MedianPointSize(IReadOnlyCollection<Letter> letters)
    {
        if (letters.Count == 0)
            return 6;

        var sizes = letters
            .Select(letter => letter.PointSize > 0 ? letter.PointSize : Math.Max(letter.BoundingBox.Width, letter.BoundingBox.Height))
            .OrderBy(size => size)
            .ToList();
        return sizes[sizes.Count / 2];
    }

    private static void AddLetterToken(ICollection<TextCandidate> candidates, IReadOnlyCollection<Letter> letters)
    {
        if (letters.Count == 0)
            return;

        string text = string.Concat(letters.Select(letter => letter.Value));
        double left = letters.Min(letter => letter.BoundingBox.Left);
        double bottom = letters.Min(letter => letter.BoundingBox.Bottom);
        double right = letters.Max(letter => letter.BoundingBox.Right);
        double top = letters.Max(letter => letter.BoundingBox.Top);
        var box = new UglyToad.PdfPig.Core.PdfRectangle(left, bottom, right, top);
        bool duplicate = candidates.Any(candidate =>
            string.Equals(candidate.Text, text, StringComparison.OrdinalIgnoreCase) && BoxesOverlap(candidate.Box, box));
        if (!duplicate)
            candidates.Add(new TextCandidate(text, box));
    }

    private static bool BoxesOverlap(
        UglyToad.PdfPig.Core.PdfRectangle first,
        UglyToad.PdfPig.Core.PdfRectangle second)
        => first.Left <= second.Right && first.Right >= second.Left &&
           first.Bottom <= second.Top && first.Top >= second.Bottom;

    private static string CleanPdfWord(string text)
    {
        if (string.IsNullOrWhiteSpace(text))
            return string.Empty;

        return text.Trim();
    }

    private static string NormalizeTag(string tag)
    {
        if (string.IsNullOrWhiteSpace(tag))
            return string.Empty;

        string cleaned = tag.Trim().Trim(
            ' ', '\t', '\r', '\n', ':', ';', ',', '.', '(', ')', '[', ']', '{', '}', '<', '>');
        return cleaned.ToUpperInvariant();
    }

    private static bool AddHighlightPopupAnnotation(PdfPage page, PdfBox box, string title, string contents, XColor color)
    {
        try
        {
            const double padding = 1.5;
            var annotation = new PdfTextAnnotation
            {
                Title = title,
                Subject = "PID annotation",
                Contents = contents,
                Color = color,
                Opacity = 0.35,
                Icon = PdfTextAnnotationIcon.Note,
                Rectangle = new PdfRectangle(
                    new XPoint(box.Left - padding, box.Bottom - padding),
                    new XPoint(box.Right + padding, box.Top + padding))
            };

            annotation.Elements.SetName("/Subtype", "/Highlight");
            annotation.Elements.SetName("/Type", "/Annot");
            annotation.Elements.SetValue("/QuadPoints", CreateQuadPoints(box));
            page.Annotations.Add(annotation);
            return true;
        }
        catch
        {
            return false;
        }
    }

    private static PdfArray CreateQuadPoints(PdfBox box)
    {
        var quadPoints = new PdfArray();
        quadPoints.Elements.Add(new PdfReal(box.Left));
        quadPoints.Elements.Add(new PdfReal(box.Top));
        quadPoints.Elements.Add(new PdfReal(box.Right));
        quadPoints.Elements.Add(new PdfReal(box.Top));
        quadPoints.Elements.Add(new PdfReal(box.Left));
        quadPoints.Elements.Add(new PdfReal(box.Bottom));
        quadPoints.Elements.Add(new PdfReal(box.Right));
        quadPoints.Elements.Add(new PdfReal(box.Bottom));
        return quadPoints;
    }

    private static void AddNoteAnnotation(PdfPage page, double x, double y, string title, string contents, XColor color)
    {
        var annotation = new PdfTextAnnotation
        {
            Title = title,
            Contents = contents,
            Icon = PdfTextAnnotationIcon.Note,
            Color = color,
            Rectangle = new PdfRectangle(
                new XPoint(x, y),
                new XPoint(x + 18, y + 18))
        };

        page.Annotations.Add(annotation);
    }

    private readonly struct VisibleBox
    {
        public VisibleBox(double left, double top, double right, double bottom)
        {
            Left = left;
            Top = top;
            Right = right;
            Bottom = bottom;
        }

        public double Left { get; }
        public double Top { get; }
        public double Right { get; }
        public double Bottom { get; }
    }

    private readonly struct PdfBox
    {
        public PdfBox(double left, double bottom, double right, double top)
        {
            Left = left;
            Bottom = bottom;
            Right = right;
            Top = top;
        }

        public double Left { get; }
        public double Bottom { get; }
        public double Right { get; }
        public double Top { get; }
    }

    private readonly struct GraphicsBox
    {
        public GraphicsBox(double left, double top, double right, double bottom)
        {
            Left = left;
            Top = top;
            Right = right;
            Bottom = bottom;
        }

        public double Left { get; }
        public double Top { get; }
        public double Right { get; }
        public double Bottom { get; }
        public double Width => Right - Left;
        public double Height => Bottom - Top;
        public double CenterX => (Left + Right) / 2;
        public double CenterY => (Top + Bottom) / 2;
    }

    private readonly struct WatermarkPlacement
    {
        public WatermarkPlacement(string text, GraphicsBox box)
        {
            Text = text;
            Box = box;
        }

        public string Text { get; }
        public GraphicsBox Box { get; }
    }

    private readonly struct TextCandidate
    {
        public TextCandidate(string text, UglyToad.PdfPig.Core.PdfRectangle box)
        {
            Text = text;
            Box = box;
        }

        public string Text { get; }
        public UglyToad.PdfPig.Core.PdfRectangle Box { get; }
    }

    private sealed class PdfTagIndex
    {
        private readonly HashSet<string> occurrenceKeys = new(StringComparer.OrdinalIgnoreCase);

        public List<PdfTagOccurrence> Occurrences { get; } = new();
        public int UniqueTagCount(TagMatchingOptions options)
            => Occurrences
                .Select(occurrence => CanonicalTagKey(occurrence.NormalizedTag, options))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .Count();

        public void Add(PdfTagOccurrence occurrence, TagMatchingOptions options)
        {
            string key = occurrence.PageNumber + "|" + occurrence.NormalizedTag + "|" +
                         Math.Round(occurrence.Box.Left, 2) + "|" + Math.Round(occurrence.Box.Bottom, 2) + "|" +
                         Math.Round(occurrence.Box.Right, 2) + "|" + Math.Round(occurrence.Box.Top, 2);
            if (occurrenceKeys.Add(key))
                Occurrences.Add(occurrence);
        }
    }

    private sealed class PdfTagOccurrence
    {
        public PdfTagOccurrence(string originalTag, int pageNumber, UglyToad.PdfPig.Core.PdfRectangle box)
        {
            OriginalTag = originalTag;
            NormalizedTag = NormalizeTag(originalTag);
            PageNumber = pageNumber;
            Box = box;
        }

        public string OriginalTag { get; }
        public string NormalizedTag { get; }
        public int PageNumber { get; }
        public UglyToad.PdfPig.Core.PdfRectangle Box { get; }
    }

    private sealed class UnmatchedPdfTag
    {
        public UnmatchedPdfTag(string tag, IReadOnlyList<int> pages, int occurrenceCount)
        {
            Tag = tag;
            Pages = pages;
            OccurrenceCount = occurrenceCount;
        }

        public string Tag { get; }
        public IReadOnlyList<int> Pages { get; }
        public int OccurrenceCount { get; }
    }
}
