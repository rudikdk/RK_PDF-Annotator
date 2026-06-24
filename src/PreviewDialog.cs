using System;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;

namespace RKPdfAnnotator;

internal sealed class PreviewDialog : Form
{
    public PreviewDialog(SheetData sheet)
    {
        Text = "Preview Tags";
        StartPosition = FormStartPosition.CenterScreen;
        FormBorderStyle = FormBorderStyle.Sizable;
        MinimizeBox = false;
        ShowInTaskbar = false;
        ClientSize = new Size(760, 480);
        Font = new Font("Segoe UI", 9F);
        BackColor = Color.FromArgb(245, 247, 250);
        Icon = AddinIcons.CreateIcon(32);

        var grid = new DataGridView
        {
            Dock = DockStyle.Fill,
            ReadOnly = true,
            AllowUserToAddRows = false,
            AllowUserToDeleteRows = false,
            AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.DisplayedCells,
            BackgroundColor = Color.White,
            BorderStyle = BorderStyle.None
        };

        grid.Columns.Add("Row", "Row");
        foreach (string header in sheet.Headers.Take(12))
            grid.Columns.Add(header, header);

        foreach (RowData row in sheet.Rows.Take(250))
        {
            object[] values = new object[grid.Columns.Count];
            values[0] = row.ExcelRowNumber;
            for (int i = 0; i < sheet.Headers.Take(12).Count(); i++)
                values[i + 1] = row.Values.TryGetValue(sheet.Headers[i], out string value) ? value : string.Empty;
            grid.Rows.Add(values);
        }

        var footer = new Label
        {
            Dock = DockStyle.Bottom,
            Height = 34,
            TextAlign = ContentAlignment.MiddleLeft,
            Padding = new Padding(12, 0, 0, 0),
            Text = sheet.Rows.Count + " rows found. Showing up to 250 rows and 12 columns.",
            ForeColor = Color.FromArgb(91, 105, 123)
        };

        Controls.Add(grid);
        Controls.Add(footer);
    }
}
