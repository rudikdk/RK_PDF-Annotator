using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Runtime.InteropServices;

namespace RKPdfAnnotator;

internal static class AddinIcons
{
    public static Bitmap CreateAnnotateBitmap(int size)
    {
        var bitmap = CreateBase(size, out Graphics graphics, out float scale);
        using (graphics)
        using (var pageBrush = new SolidBrush(Color.White))
        using (var foldBrush = new SolidBrush(Color.FromArgb(217, 234, 247)))
        using (var highlightBrush = new SolidBrush(Color.FromArgb(255, 230, 75)))
        using (var pen = new Pen(Color.FromArgb(31, 41, 51), 1.1f * scale) { StartCap = LineCap.Round, EndCap = LineCap.Round })
        {
            graphics.FillRoundedRectangle(pageBrush, ScaleRect(8, 5, 16, 22, scale), 2 * scale);
            graphics.FillPolygon(foldBrush, new[]
            {
                new PointF(19 * scale, 5 * scale),
                new PointF(24 * scale, 10 * scale),
                new PointF(19 * scale, 10 * scale)
            });
            graphics.FillRoundedRectangle(highlightBrush, ScaleRect(11, 14, 10, 4, scale), 1.2f * scale);
            graphics.DrawLine(pen, 11 * scale, 21 * scale, 20 * scale, 21 * scale);
            graphics.DrawLine(pen, 11 * scale, 24 * scale, 17 * scale, 24 * scale);
        }

        return bitmap;
    }

    public static Bitmap CreatePreviewBitmap(int size)
    {
        var bitmap = CreateBase(size, out Graphics graphics, out float scale);
        using (graphics)
        using (var lensBrush = new SolidBrush(Color.FromArgb(217, 234, 247)))
        using (var highlightBrush = new SolidBrush(Color.FromArgb(255, 230, 75)))
        using (var pen = new Pen(Color.White, 2.2f * scale) { StartCap = LineCap.Round, EndCap = LineCap.Round })
        using (var darkPen = new Pen(Color.FromArgb(31, 41, 51), 1.1f * scale))
        {
            graphics.FillEllipse(lensBrush, ScaleRect(8, 7, 13, 13, scale));
            graphics.DrawEllipse(darkPen, ScaleRect(8, 7, 13, 13, scale));
            graphics.FillRoundedRectangle(highlightBrush, ScaleRect(11, 12, 7, 3, scale), 1 * scale);
            graphics.DrawLine(pen, 19 * scale, 19 * scale, 25 * scale, 25 * scale);
        }

        return bitmap;
    }

    public static Bitmap CreateAboutBitmap(int size)
    {
        var bitmap = CreateBase(size, out Graphics graphics, out float scale);
        using (graphics)
        using (var ringBrush = new SolidBrush(Color.FromArgb(217, 234, 247)))
        using (var textBrush = new SolidBrush(Color.FromArgb(31, 41, 51)))
        using (var dotBrush = new SolidBrush(Color.White))
        using (var textFont = new Font("Segoe UI", 17f * scale, FontStyle.Bold, GraphicsUnit.Pixel))
        {
            graphics.FillEllipse(ringBrush, ScaleRect(9, 7, 14, 18, scale));
            graphics.FillEllipse(dotBrush, ScaleRect(14.1f, 10, 3.8f, 3.8f, scale));
            graphics.DrawString("i", textFont, textBrush, new PointF(13.1f * scale, 12.2f * scale));
        }

        return bitmap;
    }

    public static Bitmap CreateGuideBitmap(int size)
    {
        var bitmap = CreateBase(size, out Graphics graphics, out float scale);
        using (graphics)
        using (var pageBrush = new SolidBrush(Color.White))
        using (var foldBrush = new SolidBrush(Color.FromArgb(217, 234, 247)))
        using (var linePen = new Pen(Color.FromArgb(31, 41, 51), 1.2f * scale) { StartCap = LineCap.Round, EndCap = LineCap.Round })
        using (var accentPen = new Pen(Color.FromArgb(15, 118, 110), 1.5f * scale) { StartCap = LineCap.Round, EndCap = LineCap.Round })
        {
            graphics.FillRoundedRectangle(pageBrush, ScaleRect(9, 6, 14, 20, scale), 2 * scale);
            graphics.FillPolygon(foldBrush, new[]
            {
                new PointF(18 * scale, 6 * scale),
                new PointF(23 * scale, 11 * scale),
                new PointF(18 * scale, 11 * scale)
            });
            graphics.DrawLine(accentPen, 12 * scale, 14 * scale, 20 * scale, 14 * scale);
            graphics.DrawLine(linePen, 12 * scale, 18 * scale, 20 * scale, 18 * scale);
            graphics.DrawLine(linePen, 12 * scale, 22 * scale, 17 * scale, 22 * scale);
        }

        return bitmap;
    }

    public static Icon CreateIcon(int size)
    {
        using Bitmap bitmap = CreateAnnotateBitmap(size);
        IntPtr handle = bitmap.GetHicon();
        try
        {
            using Icon icon = Icon.FromHandle(handle);
            return (Icon)icon.Clone();
        }
        finally
        {
            DestroyIcon(handle);
        }
    }

    private static Bitmap CreateBase(int size, out Graphics graphics, out float scale)
    {
        var bitmap = new Bitmap(size, size);
        graphics = Graphics.FromImage(bitmap);
        graphics.SmoothingMode = SmoothingMode.AntiAlias;
        graphics.Clear(Color.Transparent);

        scale = size / 32f;
        using var backBrush = new SolidBrush(Color.FromArgb(15, 118, 110));
        graphics.FillRoundedRectangle(backBrush, ScaleRect(2, 2, 28, 28, scale), 6 * scale);
        return bitmap;
    }

    private static RectangleF ScaleRect(float x, float y, float width, float height, float scale)
        => new(x * scale, y * scale, width * scale, height * scale);

    [DllImport("user32.dll", SetLastError = true)]
    private static extern bool DestroyIcon(IntPtr handle);
}

internal static class GraphicsExtensions
{
    public static void FillRoundedRectangle(this Graphics graphics, Brush brush, RectangleF bounds, float radius)
    {
        using GraphicsPath path = RoundedRectangle(bounds, radius);
        graphics.FillPath(brush, path);
    }

    private static GraphicsPath RoundedRectangle(RectangleF bounds, float radius)
    {
        float diameter = radius * 2;
        var path = new GraphicsPath();
        path.AddArc(bounds.Left, bounds.Top, diameter, diameter, 180, 90);
        path.AddArc(bounds.Right - diameter, bounds.Top, diameter, diameter, 270, 90);
        path.AddArc(bounds.Right - diameter, bounds.Bottom - diameter, diameter, diameter, 0, 90);
        path.AddArc(bounds.Left, bounds.Bottom - diameter, diameter, diameter, 90, 90);
        path.CloseFigure();
        return path;
    }
}
