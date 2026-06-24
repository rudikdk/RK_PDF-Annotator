using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;

namespace RKPdfAnnotator;

internal sealed class ColorRuleDialog : Form
{
    private readonly ComboBox ruleTypeCombo;
    private readonly NumericUpDown partInput;
    private readonly ComboBox columnCombo;
    private readonly ComboBox matchTypeCombo;
    private readonly TextBox valueInput;
    private readonly Button colorButton;
    private readonly Label partLabel;
    private readonly Label columnLabel;
    private readonly Label valueLabel;
    private Color selectedColor;

    public ColorRule Rule { get; }

    public ColorRuleDialog(ColorRule rule, IReadOnlyList<string> headers)
    {
        Rule = rule;
        selectedColor = rule.Color;

        Text = "Color rule";
        StartPosition = FormStartPosition.CenterParent;
        FormBorderStyle = FormBorderStyle.FixedDialog;
        MaximizeBox = false;
        MinimizeBox = false;
        ShowInTaskbar = false;
        ClientSize = new Size(380, 430);
        Font = new Font("Segoe UI", 9F);

        ruleTypeCombo = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Width = 220 };
        ruleTypeCombo.Items.Add("Tag part");
        ruleTypeCombo.Items.Add("Excel column");

        partInput = new NumericUpDown { Minimum = 1, Maximum = 10, Value = Math.Max(1, rule.Part), Width = 220 };

        columnCombo = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Width = 220 };
        foreach (string header in headers)
            columnCombo.Items.Add(header);

        matchTypeCombo = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Width = 220 };
        matchTypeCombo.Items.Add("Exact match");
        matchTypeCombo.Items.Add("Contains");
        matchTypeCombo.Items.Add("Greater than");
        matchTypeCombo.Items.Add("Less than");
        matchTypeCombo.Items.Add("Has any value");

        valueInput = new TextBox { Width = 220, Text = rule.Value };
        colorButton = new Button { Text = "Color", Width = 220, Height = 28, BackColor = selectedColor };
        colorButton.ForeColor = selectedColor.GetBrightness() < 0.5F ? Color.White : Color.Black;
        colorButton.Click += (_, _) => ChooseColor();

        partLabel = new Label { Text = "Part number (1 = first)", AutoSize = true };
        columnLabel = new Label { Text = "Column", AutoSize = true };
        valueLabel = new Label { Text = "Value", AutoSize = true };

        BuildLayout();
        LoadFromRule(rule, headers);

        ruleTypeCombo.SelectedIndexChanged += (_, _) => UpdateVisibility();
        matchTypeCombo.SelectedIndexChanged += (_, _) => UpdateVisibility();
        UpdateVisibility();
    }

    private void BuildLayout()
    {
        var layout = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            ColumnCount = 1,
            Padding = new Padding(16),
            AutoSize = false
        };
        layout.Controls.Add(new Label { Text = "Rule type", AutoSize = true, Margin = new Padding(0, 0, 0, 4) });
        layout.Controls.Add(ruleTypeCombo);
        partLabel.Margin = new Padding(0, 10, 0, 4);
        layout.Controls.Add(partLabel);
        layout.Controls.Add(partInput);
        columnLabel.Margin = new Padding(0, 10, 0, 4);
        layout.Controls.Add(columnLabel);
        layout.Controls.Add(columnCombo);
        layout.Controls.Add(new Label { Text = "Match type", AutoSize = true, Margin = new Padding(0, 10, 0, 4) });
        layout.Controls.Add(matchTypeCombo);
        valueLabel.Margin = new Padding(0, 10, 0, 4);
        layout.Controls.Add(valueLabel);
        layout.Controls.Add(valueInput);
        layout.Controls.Add(colorButton);

        var footer = new FlowLayoutPanel { FlowDirection = FlowDirection.RightToLeft, Dock = DockStyle.Bottom, AutoSize = true, Padding = new Padding(16) };
        var okButton = new Button { Text = "OK", Width = 90, Height = 30 };
        var cancelButton = new Button { Text = "Cancel", Width = 90, Height = 30, DialogResult = DialogResult.Cancel };
        okButton.Click += (_, _) => Accept();
        footer.Controls.Add(cancelButton);
        footer.Controls.Add(okButton);
        CancelButton = cancelButton;

        Controls.Add(layout);
        Controls.Add(footer);
    }

    private void LoadFromRule(ColorRule rule, IReadOnlyList<string> headers)
    {
        ruleTypeCombo.SelectedIndex = rule.RuleType == ColorRuleType.HeaderColumn ? 1 : 0;
        matchTypeCombo.SelectedIndex = rule.MatchType switch
        {
            ColorMatchType.Contains => 1,
            ColorMatchType.GreaterThan => 2,
            ColorMatchType.LessThan => 3,
            ColorMatchType.HasValue => 4,
            _ => 0
        };

        if (!string.IsNullOrEmpty(rule.ColumnName))
        {
            int index = headers.ToList().FindIndex(h => string.Equals(h, rule.ColumnName, StringComparison.OrdinalIgnoreCase));
            columnCombo.SelectedIndex = index >= 0 ? index : (columnCombo.Items.Count > 0 ? 0 : -1);
        }
        else if (columnCombo.Items.Count > 0)
        {
            columnCombo.SelectedIndex = 0;
        }
    }

    private void UpdateVisibility()
    {
        bool isTagPart = ruleTypeCombo.SelectedIndex == 0;
        partLabel.Visible = isTagPart;
        partInput.Visible = isTagPart;
        columnLabel.Visible = !isTagPart;
        columnCombo.Visible = !isTagPart;

        bool hasValue = !isTagPart && matchTypeCombo.SelectedIndex == 4;
        valueLabel.Visible = !hasValue;
        valueInput.Visible = !hasValue;

        if (isTagPart && matchTypeCombo.Items.Count > 2)
        {
            // Tag part rules only support exact/contains in the web app.
            if (matchTypeCombo.SelectedIndex > 1)
                matchTypeCombo.SelectedIndex = 0;
        }
    }

    private void ChooseColor()
    {
        using var dialog = new ColorDialog { Color = selectedColor, FullOpen = true };
        if (dialog.ShowDialog(this) != DialogResult.OK)
            return;

        selectedColor = dialog.Color;
        colorButton.BackColor = selectedColor;
        colorButton.ForeColor = selectedColor.GetBrightness() < 0.5F ? Color.White : Color.Black;
    }

    private void Accept()
    {
        bool isTagPart = ruleTypeCombo.SelectedIndex == 0;
        if (!isTagPart && columnCombo.SelectedItem is not string)
        {
            MessageBox.Show(this, "Choose a column for this rule.", "RK PDF-Annotator", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        ColorMatchType matchType = matchTypeCombo.SelectedIndex switch
        {
            1 => ColorMatchType.Contains,
            2 => ColorMatchType.GreaterThan,
            3 => ColorMatchType.LessThan,
            4 => ColorMatchType.HasValue,
            _ => ColorMatchType.Exact
        };

        if (isTagPart && matchType != ColorMatchType.Contains)
            matchType = ColorMatchType.Exact;

        if (matchType != ColorMatchType.HasValue && string.IsNullOrWhiteSpace(valueInput.Text))
        {
            MessageBox.Show(this, "Enter a value to match.", "RK PDF-Annotator", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        Rule.RuleType = isTagPart ? ColorRuleType.TagPart : ColorRuleType.HeaderColumn;
        Rule.Part = (int)partInput.Value;
        Rule.ColumnName = isTagPart ? string.Empty : (string)(columnCombo.SelectedItem ?? string.Empty);
        Rule.MatchType = matchType;
        Rule.Value = matchType == ColorMatchType.HasValue ? string.Empty : valueInput.Text.Trim();
        Rule.Color = selectedColor;

        DialogResult = DialogResult.OK;
        Close();
    }
}
