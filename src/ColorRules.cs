using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;

namespace RKPdfAnnotator;

internal enum ColorRuleType
{
    TagPart,
    HeaderColumn
}

internal enum ColorMatchType
{
    Exact,
    Contains,
    GreaterThan,
    LessThan,
    HasValue
}

internal sealed class ColorRule
{
    public string Id { get; set; } = Guid.NewGuid().ToString("N");
    public ColorRuleType RuleType { get; set; } = ColorRuleType.TagPart;
    public int Part { get; set; } = 1;
    public string ColumnName { get; set; } = string.Empty;
    public string Value { get; set; } = string.Empty;
    public ColorMatchType MatchType { get; set; } = ColorMatchType.Exact;
    public Color Color { get; set; } = Color.Yellow;

    public string Describe()
    {
        string match = MatchType switch
        {
            ColorMatchType.Exact => "equals",
            ColorMatchType.Contains => "contains",
            ColorMatchType.GreaterThan => "is greater than",
            ColorMatchType.LessThan => "is less than",
            ColorMatchType.HasValue => "has a value",
            _ => MatchType.ToString()
        };

        string subject = RuleType == ColorRuleType.TagPart
            ? "Tag part " + Part
            : "Column \"" + ColumnName + "\"";

        string suffix = MatchType == ColorMatchType.HasValue ? string.Empty : " \"" + Value + "\"";
        return subject + " " + match + suffix;
    }
}

internal static class ColorRuleEngine
{
    public static IReadOnlyList<string> ParseTagParts(string tag)
    {
        if (string.IsNullOrWhiteSpace(tag))
            return Array.Empty<string>();

        foreach (char delimiter in new[] { '-', '.' })
        {
            if (tag.IndexOf(delimiter) >= 0)
            {
                return tag.Split(delimiter)
                    .Select(part => part.Trim())
                    .Where(part => part.Length > 0)
                    .ToList();
            }
        }

        return Array.Empty<string>();
    }

    /// <summary>
    /// Mirrors pid_annotator/tag_engine/color_rules.py apply_color_rules: rules are
    /// evaluated in order and the last matching rule wins.
    /// </summary>
    public static Color Apply(
        string tag,
        IReadOnlyDictionary<string, string>? rowValues,
        IReadOnlyList<ColorRule>? rules,
        Color defaultColor)
    {
        if (rules == null || rules.Count == 0 || string.IsNullOrWhiteSpace(tag))
            return defaultColor;

        IReadOnlyList<string> tagParts = ParseTagParts(tag);
        Color? matched = null;

        foreach (ColorRule rule in rules)
        {
            if (RuleMatches(rule, tagParts, rowValues))
                matched = rule.Color;
        }

        return matched ?? defaultColor;
    }

    private static bool RuleMatches(
        ColorRule rule,
        IReadOnlyList<string> tagParts,
        IReadOnlyDictionary<string, string>? rowValues)
    {
        if (rule.RuleType == ColorRuleType.HeaderColumn)
            return HeaderColumnMatches(rule, rowValues);

        return TagPartMatches(rule, tagParts);
    }

    private static bool TagPartMatches(ColorRule rule, IReadOnlyList<string> tagParts)
    {
        int index = rule.Part - 1;
        if (index < 0 || index >= tagParts.Count)
            return false;

        string tagPart = tagParts[index].ToUpperInvariant();
        string valueUpper = (rule.Value ?? string.Empty).Trim().ToUpperInvariant();

        return rule.MatchType == ColorMatchType.Contains
            ? tagPart.Contains(valueUpper)
            : tagPart == valueUpper;
    }

    private static bool HeaderColumnMatches(ColorRule rule, IReadOnlyDictionary<string, string>? rowValues)
    {
        if (rowValues == null || string.IsNullOrEmpty(rule.ColumnName))
            return false;
        if (!rowValues.TryGetValue(rule.ColumnName, out string? rawValue))
            return false;

        string columnValue = (rawValue ?? string.Empty).Trim();
        bool isEmpty = columnValue.Length == 0 || string.Equals(columnValue, "nan", StringComparison.OrdinalIgnoreCase);

        if (rule.MatchType == ColorMatchType.HasValue)
            return !isEmpty;

        if (isEmpty)
            return false;

        string columnUpper = columnValue.ToUpperInvariant();
        string valueUpper = (rule.Value ?? string.Empty).Trim().ToUpperInvariant();

        switch (rule.MatchType)
        {
            case ColorMatchType.Exact:
                return columnUpper == valueUpper;
            case ColorMatchType.Contains:
                return columnUpper.Contains(valueUpper);
            case ColorMatchType.GreaterThan:
                return double.TryParse(columnValue, out double greaterColumn) && double.TryParse(rule.Value, out double greaterValue)
                    ? greaterColumn > greaterValue
                    : string.CompareOrdinal(columnUpper, valueUpper) > 0;
            case ColorMatchType.LessThan:
                return double.TryParse(columnValue, out double lessColumn) && double.TryParse(rule.Value, out double lessValue)
                    ? lessColumn < lessValue
                    : string.CompareOrdinal(columnUpper, valueUpper) < 0;
            default:
                return false;
        }
    }
}
