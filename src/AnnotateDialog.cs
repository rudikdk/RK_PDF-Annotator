using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace RKPdfAnnotator;

internal sealed class AnnotateDialog : Form
{
    private readonly NumericUpDown headerRowInput;
    private readonly ComboBox tagColumnCombo;
    private readonly NumericUpDown minPartsInput;
    private readonly NumericUpDown maxPartsInput;
    private readonly TextBox separatorsBox;
    private readonly CheckedListBox noteColumnsList;
    private readonly CheckBox watermarkEnabledCheck;
    private readonly CheckedListBox watermarkColumnsList;
    private readonly NumericUpDown watermarkFontSizeInput;
    private readonly Button watermarkColorButton;
    private readonly CheckBox watermarkBackgroundCheck;
    private readonly ListBox colorRulesList;
    private readonly Button defaultColorButton;
    private readonly TextBox pdfPathBox;
    private readonly TextBox outputPathBox;
    private readonly ProgressBar progressBar;
    private readonly Label statusLabel;
    private Color watermarkTextColor = Color.Black;
    private Color defaultHighlightColor = Color.Yellow;
    private readonly List<ColorRule> colorRules = new();
    private SheetData sheet;

    public AnnotateDialog()
    {
        Text = "Annotate PDFs";
        StartPosition = FormStartPosition.CenterScreen;
        FormBorderStyle = FormBorderStyle.FixedDialog;
        MaximizeBox = false;
        MinimizeBox = false;
        ShowInTaskbar = false;
        ClientSize = new Size(740, 960);
        Font = new Font("Segoe UI", 9F);
        BackColor = Color.FromArgb(245, 247, 250);
        Icon = AddinIcons.CreateIcon(32);

        sheet = ExcelSheetReader.ReadActiveSheet();

        headerRowInput = new NumericUpDown { Minimum = 1, Maximum = 200, Value = 1, Width = 80 };
        tagColumnCombo = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Width = 260 };
        minPartsInput = new NumericUpDown { Minimum = 1, Maximum = 10, Value = 3, Width = 58 };
        maxPartsInput = new NumericUpDown { Minimum = 1, Maximum = 10, Value = 5, Width = 58 };
        separatorsBox = new TextBox { Text = "-.", Width = 80 };
        noteColumnsList = new CheckedListBox { CheckOnClick = true, Height = 150, Dock = DockStyle.Fill };
        watermarkEnabledCheck = new CheckBox { Text = "Enable watermark", AutoSize = true };
        watermarkColumnsList = new CheckedListBox { CheckOnClick = true, Height = 110, Dock = DockStyle.Fill };
        watermarkFontSizeInput = new NumericUpDown { Minimum = 5, Maximum = 24, Value = 9, Width = 58 };
        watermarkColorButton = new Button { Text = "Text color", Width = 92, Height = 28, BackColor = Color.Black, ForeColor = Color.White };
        watermarkBackgroundCheck = new CheckBox { Text = "White background", AutoSize = true };
        colorRulesList = new ListBox { Height = 110, Dock = DockStyle.Fill, IntegralHeight = false };
        defaultColorButton = new Button { Text = "Default color", Width = 110, Height = 28, BackColor = defaultHighlightColor, ForeColor = Color.Black };
        pdfPathBox = new TextBox { ReadOnly = true, Dock = DockStyle.Fill };
        outputPathBox = new TextBox { ReadOnly = true, Dock = DockStyle.Fill };
        progressBar = new ProgressBar { Minimum = 0, Maximum = 100, Height = 16, Dock = DockStyle.Fill };
        statusLabel = new Label { AutoSize = false, Height = 44, Dock = DockStyle.Fill, ForeColor = Color.FromArgb(91, 105, 123) };

        BuildLayout();
        LoadColumns();
    }

    private void BuildLayout()
    {
        var root = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            ColumnCount = 1,
            RowCount = 9,
            Padding = new Padding(18),
            BackColor = BackColor
        };
        root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        root.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
        root.RowStyles.Add(new RowStyle(SizeType.Absolute, 190));
        root.RowStyles.Add(new RowStyle(SizeType.Absolute, 170));
        root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        Controls.Add(root);

        root.Controls.Add(new Label
        {
            Text = "Workbook: " + sheet.WorkbookName + " / " + sheet.SheetName,
            AutoSize = true,
            Font = new Font(Font.FontFamily, 10F, FontStyle.Bold),
            ForeColor = Color.FromArgb(31, 41, 51),
            Margin = new Padding(0, 0, 0, 12)
        }, 0, 0);

        var settings = new TableLayoutPanel { Dock = DockStyle.Top, AutoSize = true, ColumnCount = 4, BackColor = BackColor };
        settings.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
        settings.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
        settings.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
        settings.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
        settings.Controls.Add(new Label { Text = "Header row", AutoSize = true, Anchor = AnchorStyles.Left, Margin = new Padding(0, 4, 8, 8) }, 0, 0);
        settings.Controls.Add(headerRowInput, 1, 0);
        var reloadButton = new Button { Text = "Reload Columns", Width = 120, Height = 28, Margin = new Padding(12, 0, 0, 8) };
        reloadButton.Click += (_, _) => ReloadColumns();
        settings.Controls.Add(reloadButton, 2, 0);
        settings.Controls.Add(new Label { Text = "Tag column", AutoSize = true, Anchor = AnchorStyles.Left, Margin = new Padding(0, 4, 8, 8) }, 0, 1);
        settings.Controls.Add(tagColumnCombo, 1, 1);
        root.Controls.Add(settings, 0, 1);

        root.Controls.Add(BuildTagFormatPanel(), 0, 2);

        var notesPanel = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 2, BackColor = BackColor };
        notesPanel.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        notesPanel.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
        notesPanel.Controls.Add(new Label { Text = "Note columns", AutoSize = true, Margin = new Padding(0, 0, 0, 6) }, 0, 0);
        notesPanel.Controls.Add(noteColumnsList, 0, 1);
        root.Controls.Add(notesPanel, 0, 3);

        root.Controls.Add(BuildWatermarkPanel(), 0, 4);
        root.Controls.Add(BuildColorRulesPanel(), 0, 5);

        root.Controls.Add(BuildPathPicker("PDF file", pdfPathBox, SelectPdf), 0, 6);
        root.Controls.Add(BuildPathPicker("Output PDF", outputPathBox, SelectOutput), 0, 7);

        var footer = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 3, AutoSize = true, BackColor = BackColor, Margin = new Padding(0, 14, 0, 0) };
        footer.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
        footer.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
        footer.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
        var progressPanel = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 2, AutoSize = true, BackColor = BackColor };
        progressPanel.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        progressPanel.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        progressPanel.Controls.Add(statusLabel, 0, 0);
        progressPanel.Controls.Add(progressBar, 0, 1);
        footer.Controls.Add(progressPanel, 0, 0);
        var cancelButton = new Button { Text = "Close", Width = 92, Height = 34, DialogResult = DialogResult.Cancel, Margin = new Padding(8, 0, 0, 0) };
        var annotateButton = new Button { Text = "Annotate", Width = 104, Height = 34, Margin = new Padding(8, 0, 0, 0) };
        annotateButton.Click += (_, _) => Annotate();
        footer.Controls.Add(annotateButton, 1, 0);
        footer.Controls.Add(cancelButton, 2, 0);
        CancelButton = cancelButton;
        root.Controls.Add(footer, 0, 8);
    }

    private Control BuildTagFormatPanel()
    {
        var panel = new TableLayoutPanel
        {
            Dock = DockStyle.Top,
            AutoSize = true,
            ColumnCount = 7,
            Margin = new Padding(0, 2, 0, 10),
            BackColor = BackColor
        };
        panel.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
        panel.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
        panel.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
        panel.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
        panel.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
        panel.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
        panel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));

        panel.Controls.Add(new Label { Text = "Tag format", AutoSize = true, Anchor = AnchorStyles.Left, Margin = new Padding(0, 4, 12, 0) }, 0, 0);
        panel.Controls.Add(new Label { Text = "Min parts", AutoSize = true, Anchor = AnchorStyles.Left, Margin = new Padding(0, 4, 6, 0) }, 1, 0);
        panel.Controls.Add(minPartsInput, 2, 0);
        panel.Controls.Add(new Label { Text = "Max parts", AutoSize = true, Anchor = AnchorStyles.Left, Margin = new Padding(12, 4, 6, 0) }, 3, 0);
        panel.Controls.Add(maxPartsInput, 4, 0);
        panel.Controls.Add(new Label { Text = "Separators", AutoSize = true, Anchor = AnchorStyles.Left, Margin = new Padding(12, 4, 6, 0) }, 5, 0);
        panel.Controls.Add(separatorsBox, 6, 0);

        minPartsInput.ValueChanged += (_, _) =>
        {
            if (maxPartsInput.Value < minPartsInput.Value)
                maxPartsInput.Value = minPartsInput.Value;
        };
        maxPartsInput.ValueChanged += (_, _) =>
        {
            if (minPartsInput.Value > maxPartsInput.Value)
                minPartsInput.Value = maxPartsInput.Value;
        };

        return panel;
    }

    private Control BuildColorRulesPanel()
    {
        var panel = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            ColumnCount = 2,
            RowCount = 3,
            Margin = new Padding(0, 10, 0, 0),
            BackColor = BackColor
        };
        panel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
        panel.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
        panel.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        panel.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
        panel.RowStyles.Add(new RowStyle(SizeType.AutoSize));

        panel.Controls.Add(new Label { Text = "Color rules (highlight color by tag part or column value)", AutoSize = true }, 0, 0);

        var ruleButtons = new FlowLayoutPanel
        {
            AutoSize = true,
            FlowDirection = FlowDirection.LeftToRight,
            WrapContents = false,
            Anchor = AnchorStyles.Right
        };
        var addButton = new Button { Text = "Add rule", Width = 78, Height = 27 };
        var editButton = new Button { Text = "Edit", Width = 60, Height = 27 };
        var removeButton = new Button { Text = "Remove", Width = 66, Height = 27 };
        var upButton = new Button { Text = "Move up", Width = 78, Height = 27 };
        var downButton = new Button { Text = "Move down", Width = 86, Height = 27 };
        addButton.Click += (_, _) => AddColorRule();
        editButton.Click += (_, _) => EditSelectedColorRule();
        removeButton.Click += (_, _) => RemoveSelectedColorRule();
        upButton.Click += (_, _) => MoveColorRule(-1);
        downButton.Click += (_, _) => MoveColorRule(1);
        ruleButtons.Controls.Add(addButton);
        ruleButtons.Controls.Add(editButton);
        ruleButtons.Controls.Add(removeButton);
        ruleButtons.Controls.Add(upButton);
        ruleButtons.Controls.Add(downButton);
        panel.Controls.Add(ruleButtons, 1, 0);

        colorRulesList.DoubleClick += (_, _) => EditSelectedColorRule();
        panel.Controls.Add(colorRulesList, 0, 1);
        panel.SetColumnSpan(colorRulesList, 2);

        var defaultRow = new FlowLayoutPanel
        {
            AutoSize = true,
            FlowDirection = FlowDirection.LeftToRight,
            WrapContents = false,
            Margin = new Padding(0, 6, 0, 0)
        };
        defaultRow.Controls.Add(new Label { Text = "Default highlight color (used when no rule matches)", AutoSize = true, Margin = new Padding(0, 6, 6, 0) });
        defaultColorButton.Click += (_, _) => ChooseDefaultHighlightColor();
        defaultRow.Controls.Add(defaultColorButton);
        panel.Controls.Add(defaultRow, 0, 2);
        panel.SetColumnSpan(defaultRow, 2);

        return panel;
    }

    private void AddColorRule()
    {
        var rule = new ColorRule { Color = defaultHighlightColor };
        using var dialog = new ColorRuleDialog(rule, sheet.Headers);
        if (dialog.ShowDialog(this) != DialogResult.OK)
            return;

        colorRules.Add(rule);
        RefreshColorRulesList();
        colorRulesList.SelectedIndex = colorRules.Count - 1;
    }

    private void EditSelectedColorRule()
    {
        int index = colorRulesList.SelectedIndex;
        if (index < 0 || index >= colorRules.Count)
            return;

        using var dialog = new ColorRuleDialog(colorRules[index], sheet.Headers);
        if (dialog.ShowDialog(this) != DialogResult.OK)
            return;

        RefreshColorRulesList();
        colorRulesList.SelectedIndex = index;
    }

    private void RemoveSelectedColorRule()
    {
        int index = colorRulesList.SelectedIndex;
        if (index < 0 || index >= colorRules.Count)
            return;

        colorRules.RemoveAt(index);
        RefreshColorRulesList();
    }

    private void MoveColorRule(int offset)
    {
        int index = colorRulesList.SelectedIndex;
        int target = index + offset;
        if (index < 0 || target < 0 || target >= colorRules.Count)
            return;

        ColorRule rule = colorRules[index];
        colorRules.RemoveAt(index);
        colorRules.Insert(target, rule);
        RefreshColorRulesList();
        colorRulesList.SelectedIndex = target;
    }

    private void RefreshColorRulesList()
    {
        colorRulesList.Items.Clear();
        foreach (ColorRule rule in colorRules)
            colorRulesList.Items.Add(rule.Describe() + "  ->  " + ColorHex(rule.Color));
    }

    private static string ColorHex(Color color)
        => "#" + color.R.ToString("X2") + color.G.ToString("X2") + color.B.ToString("X2");

    private void ChooseDefaultHighlightColor()
    {
        using var dialog = new ColorDialog { Color = defaultHighlightColor, FullOpen = true };
        if (dialog.ShowDialog(this) != DialogResult.OK)
            return;

        defaultHighlightColor = dialog.Color;
        defaultColorButton.BackColor = defaultHighlightColor;
        defaultColorButton.ForeColor = defaultHighlightColor.GetBrightness() < 0.5F ? Color.White : Color.Black;
    }

    private Control BuildWatermarkPanel()
    {
        var panel = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            ColumnCount = 2,
            RowCount = 3,
            Margin = new Padding(0, 10, 0, 0),
            BackColor = BackColor
        };
        panel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
        panel.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
        panel.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        panel.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
        panel.RowStyles.Add(new RowStyle(SizeType.AutoSize));

        watermarkEnabledCheck.CheckedChanged += (_, _) => UpdateWatermarkControls();
        panel.Controls.Add(watermarkEnabledCheck, 0, 0);

        var orderButtons = new FlowLayoutPanel
        {
            AutoSize = true,
            FlowDirection = FlowDirection.LeftToRight,
            WrapContents = false,
            Anchor = AnchorStyles.Right
        };
        var upButton = new Button { Text = "Move up", Width = 78, Height = 27 };
        var downButton = new Button { Text = "Move down", Width = 86, Height = 27 };
        upButton.Click += (_, _) => MoveWatermarkColumn(-1);
        downButton.Click += (_, _) => MoveWatermarkColumn(1);
        orderButtons.Controls.Add(upButton);
        orderButtons.Controls.Add(downButton);
        panel.Controls.Add(orderButtons, 1, 0);

        panel.Controls.Add(watermarkColumnsList, 0, 1);
        panel.SetColumnSpan(watermarkColumnsList, 2);

        var options = new FlowLayoutPanel
        {
            AutoSize = true,
            FlowDirection = FlowDirection.LeftToRight,
            WrapContents = false,
            Margin = new Padding(0, 6, 0, 0)
        };
        options.Controls.Add(new Label { Text = "Font size", AutoSize = true, Margin = new Padding(0, 6, 6, 0) });
        options.Controls.Add(watermarkFontSizeInput);
        watermarkColorButton.Click += (_, _) => ChooseWatermarkColor();
        options.Controls.Add(watermarkColorButton);
        watermarkBackgroundCheck.Margin = new Padding(10, 6, 0, 0);
        options.Controls.Add(watermarkBackgroundCheck);
        panel.Controls.Add(options, 0, 2);
        panel.SetColumnSpan(options, 2);

        UpdateWatermarkControls();
        return panel;
    }

    private void ChooseWatermarkColor()
    {
        using var dialog = new ColorDialog { Color = watermarkTextColor, FullOpen = true };
        if (dialog.ShowDialog(this) != DialogResult.OK)
            return;

        watermarkTextColor = dialog.Color;
        watermarkColorButton.BackColor = dialog.Color;
        watermarkColorButton.ForeColor = dialog.Color.GetBrightness() < 0.5F ? Color.White : Color.Black;
    }

    private void UpdateWatermarkControls()
    {
        bool enabled = watermarkEnabledCheck.Checked;
        watermarkColumnsList.Enabled = enabled;
        watermarkFontSizeInput.Enabled = enabled;
        watermarkColorButton.Enabled = enabled;
        watermarkBackgroundCheck.Enabled = enabled;
    }

    private void MoveWatermarkColumn(int offset)
    {
        int index = watermarkColumnsList.SelectedIndex;
        int target = index + offset;
        if (index < 0 || target < 0 || target >= watermarkColumnsList.Items.Count)
            return;

        object item = watermarkColumnsList.Items[index];
        bool wasChecked = watermarkColumnsList.GetItemChecked(index);
        watermarkColumnsList.Items.RemoveAt(index);
        watermarkColumnsList.Items.Insert(target, item);
        watermarkColumnsList.SetItemChecked(target, wasChecked);
        watermarkColumnsList.SelectedIndex = target;
    }

    private static Control BuildPathPicker(string labelText, TextBox textBox, Action browseAction)
    {
        var panel = new TableLayoutPanel { Dock = DockStyle.Top, AutoSize = true, ColumnCount = 3, Margin = new Padding(0, 10, 0, 0) };
        panel.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 88));
        panel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
        panel.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
        panel.Controls.Add(new Label { Text = labelText, AutoSize = true, Anchor = AnchorStyles.Left, Margin = new Padding(0, 5, 8, 0) }, 0, 0);
        panel.Controls.Add(textBox, 1, 0);
        var browse = new Button { Text = "Browse", Width = 86, Height = 28, Margin = new Padding(8, 0, 0, 0) };
        browse.Click += (_, _) => browseAction();
        panel.Controls.Add(browse, 2, 0);
        return panel;
    }

    private void ReloadColumns()
    {
        sheet = ExcelSheetReader.ReadActiveSheet((int)headerRowInput.Value);
        LoadColumns();
    }

    private void LoadColumns()
    {
        tagColumnCombo.Items.Clear();
        noteColumnsList.Items.Clear();
        watermarkColumnsList.Items.Clear();

        foreach (string header in sheet.Headers)
        {
            tagColumnCombo.Items.Add(header);
            noteColumnsList.Items.Add(header);
            watermarkColumnsList.Items.Add(header);
        }

        int tagIndex = FindLikelyTagColumn(sheet.Headers);
        if (tagColumnCombo.Items.Count > 0)
            tagColumnCombo.SelectedIndex = tagIndex;

        for (int i = 0; i < noteColumnsList.Items.Count; i++)
        {
            string header = noteColumnsList.Items[i]?.ToString() ?? string.Empty;
            if (i != tagIndex && IsLikelyNoteColumn(header))
                noteColumnsList.SetItemChecked(i, true);
        }

        statusLabel.Text = sheet.Rows.Count + " data rows found.";
    }

    private void SelectPdf()
    {
        using var dialog = new OpenFileDialog { Filter = "PDF files (*.pdf)|*.pdf", Title = "Choose PDF to annotate" };
        if (dialog.ShowDialog(this) != DialogResult.OK)
            return;

        pdfPathBox.Text = dialog.FileName;
        if (string.IsNullOrWhiteSpace(outputPathBox.Text))
        {
            string folder = Path.GetDirectoryName(dialog.FileName) ?? Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);
            string name = Path.GetFileNameWithoutExtension(dialog.FileName) + "_annotated.pdf";
            outputPathBox.Text = Path.Combine(folder, name);
        }
    }

    private void SelectOutput()
    {
        using var dialog = new SaveFileDialog
        {
            Filter = "PDF files (*.pdf)|*.pdf",
            Title = "Save annotated PDF",
            FileName = string.IsNullOrWhiteSpace(outputPathBox.Text) ? "annotated.pdf" : Path.GetFileName(outputPathBox.Text)
        };
        if (dialog.ShowDialog(this) == DialogResult.OK)
            outputPathBox.Text = dialog.FileName;
    }

    private async void Annotate()
    {
        if (tagColumnCombo.SelectedItem is not string tagColumn)
            throw new InvalidOperationException("Choose a tag column.");
        if (!File.Exists(pdfPathBox.Text))
            throw new InvalidOperationException("Choose an existing PDF file.");
        if (string.IsNullOrWhiteSpace(outputPathBox.Text))
            throw new InvalidOperationException("Choose an output PDF path.");

        var noteColumns = noteColumnsList.CheckedItems.Cast<object>().Select(item => item.ToString() ?? string.Empty).Where(item => item.Length > 0).ToList();
        var watermarkColumns = watermarkColumnsList.CheckedItems.Cast<object>().Select(item => item.ToString() ?? string.Empty).Where(item => item.Length > 0).ToList();
        if (watermarkEnabledCheck.Checked && watermarkColumns.Count == 0)
            throw new InvalidOperationException("Choose at least one watermark column or disable watermarking.");

        var watermarkOptions = new WatermarkOptions(
            watermarkEnabledCheck.Checked,
            watermarkColumns,
            (float)watermarkFontSizeInput.Value,
            watermarkTextColor,
            watermarkBackgroundCheck.Checked);
        var tagMatchingOptions = new TagMatchingOptions(
            (int)minPartsInput.Value,
            (int)maxPartsInput.Value,
            separatorsBox.Text);
        IReadOnlyList<TagRecord> records = ExcelSheetReader.BuildTagRecords(
            sheet, tagColumn, noteColumns, watermarkColumns, colorRules, defaultHighlightColor);
        if (records.Count == 0)
            throw new InvalidOperationException("No tag values were found in the selected tag column.");

        Cursor = Cursors.WaitCursor;
        progressBar.Value = 0;
        statusLabel.Text = "Preparing annotation...";
        SetControlsEnabled(false);
        var progress = new Progress<AnnotationProgress>(UpdateProgress);
        try
        {
            string inputPath = pdfPathBox.Text;
            string outputPath = outputPathBox.Text;
            AnnotationResult result = await Task.Run(() =>
                PdfAnnotationEngine.Annotate(inputPath, outputPath, records, watermarkOptions, tagMatchingOptions, progress));
            UpdateProgress(new AnnotationProgress(99, "Creating Excel report sheet..."));
            string reportSheetName = ExcelReportWriter.WriteToNewSheet(result.Report);
            UpdateProgress(new AnnotationProgress(100, "Annotation complete. Report sheet created."));
            statusLabel.Text = result.MatchedTags + " of " + result.TotalTags + " Excel tags matched; " +
                               result.ExcelTagsNotFound + " Excel tags not found in PDF. Report sheet: " + reportSheetName;
            MessageBox.Show(
                result.MatchedTags + " of " + result.TotalTags + " Excel tags were matched." + Environment.NewLine +
                result.ExcelTagsNotFound + " Excel tags were not found in the PDF." + Environment.NewLine +
                result.Watermarks + " watermarks were added." + Environment.NewLine +
                "PDF: " + outputPathBox.Text + Environment.NewLine +
                "Report sheet: " + reportSheetName + Environment.NewLine +
                "CSV report: " + result.ReportPath,
                "RK PDF-Annotator",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information);
        }
        finally
        {
            SetControlsEnabled(true);
            Cursor = Cursors.Default;
        }
    }

    private void UpdateProgress(AnnotationProgress progress)
    {
        progressBar.Value = Math.Max(progressBar.Minimum, Math.Min(progressBar.Maximum, progress.Percent));
        statusLabel.Text = progress.Percent + "% - " + progress.Message;
    }

    private void SetControlsEnabled(bool enabled)
    {
        foreach (Control control in Controls)
            control.Enabled = enabled;
    }

    private static int FindLikelyTagColumn(IReadOnlyList<string> headers)
    {
        string[] names = { "tag", "tag id", "tag no", "tag number", "component", "component tag" };
        for (int i = 0; i < headers.Count; i++)
        {
            if (names.Any(name => string.Equals(headers[i], name, StringComparison.OrdinalIgnoreCase)))
                return i;
        }

        for (int i = 0; i < headers.Count; i++)
        {
            if (headers[i].IndexOf("tag", StringComparison.OrdinalIgnoreCase) >= 0)
                return i;
        }

        return 0;
    }

    private static bool IsLikelyNoteColumn(string header)
        => header.IndexOf("description", StringComparison.OrdinalIgnoreCase) >= 0 ||
           header.IndexOf("note", StringComparison.OrdinalIgnoreCase) >= 0 ||
           header.IndexOf("comment", StringComparison.OrdinalIgnoreCase) >= 0 ||
           header.IndexOf("location", StringComparison.OrdinalIgnoreCase) >= 0 ||
           header.IndexOf("service", StringComparison.OrdinalIgnoreCase) >= 0;
}
