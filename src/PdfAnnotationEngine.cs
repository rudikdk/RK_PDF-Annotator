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
    public AnnotationResult(int totalTags, int matchedTags, int annotations, int watermarks)
    {
        TotalTags = totalTags;
        MatchedTags = matchedTags;
        Annotations = annotations;
        Watermarks = watermarks;
    }

    public int TotalTags { get; }
    public int MatchedTags { get; }
    public int Annotations { get; }
    public int Watermarks { get; }
}

internal static class PdfAnnotationEngine
{
    private static readonly Regex TagPattern = new Regex(
        @"\b[A-Za-z0-9]{1,5}[-\.][A-Za-z0-9]{1,5}[-\.][A-Za-z0-9]{1,5}(?:[-\.][A-Za-z0-9]{1,5}){0,2}\b",
        RegexOptions.Compiled | RegexOptions.IgnoreCase);

    public static AnnotationResult Annotate(
        string inputPdfPath,
        string outputPdfPath,
        IReadOnlyList<TagRecord> records,
        WatermarkOptions watermarkOptions)
    {
        if (!File.Exists(inputPdfPath))
            throw new FileNotFoundException("Input PDF was not found.", inputPdfPath);

        Directory.CreateDirectory(Path.GetDirectoryName(outputPdfPath) ?? Environment.CurrentDirectory);

        var recordsByTag = records
            .SelectMany(record => GetTagVariants(record.Tag)
                .Select(variant => new { Variant = variant, Record = record }))
            .GroupBy(item => item.Variant, StringComparer.OrdinalIgnoreCase)
            .ToDictionary(group => group.Key, group => group.First().Record, StringComparer.OrdinalIgnoreCase);
        var matchedTags = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var annotatedLocations = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        int annotationCount = 0;
        int watermarkCount = 0;

        using PdfSharp.Pdf.PdfDocument outputDocument = PdfReader.Open(inputPdfPath, PdfDocumentOpenMode.Modify);
        using UglyToad.PdfPig.PdfDocument sourceTextDocument = UglyToad.PdfPig.PdfDocument.Open(inputPdfPath);

        foreach (Page sourcePage in sourceTextDocument.GetPages())
        {
            PdfPage outputPage = outputDocument.Pages[sourcePage.Number - 1];
            var watermarkPlacements = new List<WatermarkPlacement>();

            foreach (TextCandidate candidate in GetTextCandidates(sourcePage))
            {
                string text = CleanPdfWord(candidate.Text);
                if (text.Length == 0)
                    continue;

                var tagCandidates = TagPattern.Matches(text).Cast<Match>().Select(match => match.Value).ToList();
                if (tagCandidates.Count == 0)
                    tagCandidates.Add(text);

                foreach (string tagCandidate in tagCandidates)
                {
                    string pdfTag = NormalizeTag(tagCandidate);
                    if (pdfTag.Length == 0)
                        continue;

                    TagRecord? record = null;
                    foreach (string variant in GetTagVariants(pdfTag))
                    {
                        if (recordsByTag.TryGetValue(variant, out record))
                            break;
                    }

                    if (record == null)
                        continue;

                    var box = candidate.Box;
                    string locationKey = sourcePage.Number + "|" + record.Tag.ToUpperInvariant() + "|" +
                                         Math.Round(box.Left, 2) + "|" + Math.Round(box.Bottom, 2) + "|" +
                                         Math.Round(box.Right, 2) + "|" + Math.Round(box.Top, 2);
                    if (!annotatedLocations.Add(locationKey))
                        continue;

                    PdfBox annotationBox = ToPdfAnnotationBox(outputPage, ToVisiblePageBox(outputPage, box));
                    XColor highlightColor = XColor.FromArgb(record.HighlightColor.R, record.HighlightColor.G, record.HighlightColor.B);
                    if (!AddHighlightPopupAnnotation(outputPage, annotationBox, record.Tag, record.Note, highlightColor))
                        AddNoteAnnotation(outputPage, annotationBox.Right + 2, annotationBox.Bottom, record.Tag, record.Note, highlightColor);

                    if (watermarkOptions.Enabled && !string.IsNullOrWhiteSpace(record.WatermarkText))
                    {
                        watermarkPlacements.Add(new WatermarkPlacement(record.WatermarkText, ToGraphicsBox(outputPage, box)));
                    }

                    matchedTags.Add(record.Tag);
                    annotationCount++;
                }
            }

            watermarkCount += DrawWatermarks(outputPage, watermarkPlacements, watermarkOptions);
        }

        outputDocument.Save(outputPdfPath);
        int totalTags = records
            .Select(record => NormalizeTag(record.Tag))
            .Where(tag => tag.Length > 0)
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .Count();
        return new AnnotationResult(totalTags, matchedTags.Count, annotationCount, watermarkCount);
    }

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

    private static IReadOnlyCollection<string> GetTagVariants(string tag)
    {
        string normalized = NormalizeTag(tag);
        if (normalized.Length == 0)
            return Array.Empty<string>();

        return new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            normalized,
            normalized.Replace('-', '.'),
            normalized.Replace('.', '-')
        };
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

        return candidates;
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
}
