using System;
using System.Diagnostics;
using System.Drawing;
using System.Windows.Forms;

namespace RKPdfAnnotator;

internal sealed class AboutDialog : Form
{
    private const string GitHubUrl = "https://github.com/rudikdk/rk-pdf-annotator";

    public AboutDialog()
    {
        Text = "About RK PDF-Annotator";
        StartPosition = FormStartPosition.CenterScreen;
        FormBorderStyle = FormBorderStyle.FixedDialog;
        MaximizeBox = false;
        MinimizeBox = false;
        ShowInTaskbar = false;
        ClientSize = new Size(540, 350);
        Font = new Font("Segoe UI", 9F);
        BackColor = Color.FromArgb(245, 247, 250);
        Icon = AddinIcons.CreateIcon(32);

        BuildLayout();
    }

    private void BuildLayout()
    {
        var root = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            ColumnCount = 1,
            RowCount = 3,
            Padding = new Padding(20),
            BackColor = BackColor
        };
        root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        root.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
        root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        Controls.Add(root);

        root.Controls.Add(BuildHeader(), 0, 0);
        root.Controls.Add(BuildContent(), 0, 1);
        root.Controls.Add(BuildFooter(), 0, 2);
    }

    private Control BuildHeader()
    {
        var header = new TableLayoutPanel
        {
            Dock = DockStyle.Top,
            AutoSize = true,
            ColumnCount = 2,
            Margin = new Padding(0, 0, 0, 16),
            BackColor = BackColor
        };
        header.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 52));
        header.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));

        header.Controls.Add(new PictureBox
        {
            Image = AddinIcons.CreateAnnotateBitmap(42),
            SizeMode = PictureBoxSizeMode.CenterImage,
            Dock = DockStyle.Fill,
            Margin = new Padding(0, 2, 12, 0)
        }, 0, 0);

        var titleBlock = new TableLayoutPanel { Dock = DockStyle.Top, AutoSize = true, ColumnCount = 1, BackColor = BackColor };
        titleBlock.Controls.Add(new Label
        {
            Text = "RK PDF-Annotator",
            AutoSize = true,
            Font = new Font(Font.FontFamily, 15F, FontStyle.Bold),
            ForeColor = Color.FromArgb(31, 41, 51),
            Margin = new Padding(0, 0, 0, 2)
        });
        titleBlock.Controls.Add(new Label
        {
            Text = "Version 1.0",
            AutoSize = true,
            ForeColor = Color.FromArgb(91, 105, 123),
            Margin = new Padding(0)
        });
        header.Controls.Add(titleBlock, 1, 0);

        return header;
    }

    private Control BuildContent()
    {
        var content = new TableLayoutPanel { Dock = DockStyle.Fill, AutoSize = true, ColumnCount = 1, BackColor = BackColor };

        content.Controls.Add(new Label
        {
            Text = "A Windows Excel-DNA add-in for annotating P&ID PDF drawings from the currently open Excel workbook. It reads tag rows from Excel, finds matching PDF text, highlights the matches, and adds PDF note annotations with selected workbook data.",
            AutoSize = false,
            Height = 78,
            Dock = DockStyle.Top,
            ForeColor = Color.FromArgb(31, 41, 51),
            Margin = new Padding(0, 0, 0, 12)
        });

        content.Controls.Add(new Label
        {
            Text = "Made by Rudi Kaergaard",
            AutoSize = true,
            ForeColor = Color.FromArgb(31, 41, 51),
            Margin = new Padding(0, 0, 0, 4)
        });

        content.Controls.Add(new Label
        {
            Text = "Contact: contact@rkcadtools.com",
            AutoSize = true,
            ForeColor = Color.FromArgb(31, 41, 51),
            Margin = new Padding(0, 0, 0, 12)
        });

        var link = new LinkLabel
        {
            Text = "GitHub: " + GitHubUrl,
            AutoSize = true,
            LinkColor = Color.FromArgb(15, 118, 110),
            ActiveLinkColor = Color.FromArgb(12, 91, 84),
            VisitedLinkColor = Color.FromArgb(15, 118, 110),
            Margin = new Padding(0)
        };
        link.Links.Add("GitHub: ".Length, GitHubUrl.Length, GitHubUrl);
        link.LinkClicked += (_, e) =>
        {
            if (e.Link.LinkData is string url)
                OpenUrl(url);
        };
        content.Controls.Add(link);

        return content;
    }

    private Control BuildFooter()
    {
        var closeButton = new Button
        {
            Text = "Close",
            DialogResult = DialogResult.OK,
            Anchor = AnchorStyles.Right,
            Width = 92,
            Height = 34,
            BackColor = Color.White,
            ForeColor = Color.FromArgb(31, 41, 51),
            FlatStyle = FlatStyle.Flat
        };
        closeButton.FlatAppearance.BorderColor = Color.FromArgb(214, 221, 230);
        AcceptButton = closeButton;
        CancelButton = closeButton;
        return closeButton;
    }

    private static void OpenUrl(string url)
    {
        try
        {
            Process.Start(new ProcessStartInfo(url) { UseShellExecute = true });
        }
        catch (Exception ex)
        {
            MessageBox.Show(
                "Could not open the GitHub repository." + Environment.NewLine + Environment.NewLine + ex.Message,
                "RK PDF-Annotator",
                MessageBoxButtons.OK,
                MessageBoxIcon.Warning);
        }
    }
}
