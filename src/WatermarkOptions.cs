using System.Collections.Generic;
using System.Drawing;

namespace RKPdfAnnotator;

internal sealed class WatermarkOptions
{
    public WatermarkOptions(
        bool enabled,
        IReadOnlyList<string> columns,
        float fontSize,
        Color textColor,
        bool backgroundEnabled)
    {
        Enabled = enabled;
        Columns = columns;
        FontSize = fontSize;
        TextColor = textColor;
        BackgroundEnabled = backgroundEnabled;
    }

    public bool Enabled { get; }
    public IReadOnlyList<string> Columns { get; }
    public float FontSize { get; }
    public Color TextColor { get; }
    public bool BackgroundEnabled { get; }
}
