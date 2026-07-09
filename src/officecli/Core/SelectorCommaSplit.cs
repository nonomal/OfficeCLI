// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

namespace OfficeCli.Core;

/// <summary>
/// Top-level comma handling for CSS-style selector lists (union). A comma is
/// "top level" only when it sits outside every <c>[]</c> / <c>()</c> group and
/// outside quotes, so a value like <c>row[name="A,B"]</c> or a function arg is
/// never split. Shared by the PowerPoint and Excel query paths so comma-union
/// behaves identically across handlers (CONSISTENCY(comma-union)).
/// </summary>
internal static class SelectorCommaSplit
{
    public static bool ContainsTopLevelComma(string selector)
    {
        int depthBracket = 0, depthParen = 0;
        char? quote = null;
        foreach (var c in selector)
        {
            if (quote.HasValue) { if (c == quote.Value) quote = null; continue; }
            if (c == '"' || c == '\'') { quote = c; continue; }
            if (c == '[') depthBracket++;
            else if (c == ']') depthBracket = System.Math.Max(0, depthBracket - 1);
            else if (c == '(') depthParen++;
            else if (c == ')') depthParen = System.Math.Max(0, depthParen - 1);
            else if (c == ',' && depthBracket == 0 && depthParen == 0) return true;
        }
        return false;
    }

    public static List<string> SplitTopLevelCommas(string selector)
    {
        var parts = new List<string>();
        int depthBracket = 0, depthParen = 0;
        char? quote = null;
        int start = 0;
        for (int i = 0; i < selector.Length; i++)
        {
            var c = selector[i];
            if (quote.HasValue) { if (c == quote.Value) quote = null; continue; }
            if (c == '"' || c == '\'') { quote = c; continue; }
            if (c == '[') depthBracket++;
            else if (c == ']') depthBracket = System.Math.Max(0, depthBracket - 1);
            else if (c == '(') depthParen++;
            else if (c == ')') depthParen = System.Math.Max(0, depthParen - 1);
            else if (c == ',' && depthBracket == 0 && depthParen == 0)
            {
                parts.Add(selector.Substring(start, i - start));
                start = i + 1;
            }
        }
        parts.Add(selector.Substring(start));
        return parts;
    }
}
