using System;
using System.IO;
using PdfSharp.Drawing;
using PdfSharp.Pdf;

class Program
{
    static int Main()
    {
        string root = Path.Combine(Directory.GetCurrentDirectory(), "excel-addin", "tmp-smoke");
        Directory.CreateDirectory(root);
        string pdf = Path.Combine(root, "sample.pdf");
        var doc = new PdfDocument();
        var page = doc.AddPage();
        using (var gfx = XGraphics.FromPdfPage(page))
        using (var font = new XFont("Arial", 16))
        {
            gfx.DrawString("PUMP-101-A is installed near the inlet.", font, XBrushes.Black, new XPoint(72, 120));
        }
        doc.Save(pdf);
        Console.WriteLine(pdf);
        return 0;
    }
}
