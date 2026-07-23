// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

using System;
using System.Collections.Generic;
using System.Text;
using System.Text.RegularExpressions;

namespace OfficeCli.Core.Markdown;

/// <summary>
/// Markdown-subset parser: text → <see cref="MarkdownDocument"/> (semantic IR).
/// The block-level analogue of <c>Core/Diagram/MermaidParser</c>.
///
/// Handles the regular, well-formed markdown that LLMs emit — NOT full
/// CommonMark. Supported:
///   headings      # … ###### (ATX)
///   paragraphs    blank-line separated
///   lists         - / * / + (unordered), 1. 2. (ordered), 2-space nesting
///   code blocks   ```lang fenced
///   blockquote    &gt; line
///   tables        | a | b | GFM pipe (with |---|---| delimiter row)
///   rule          --- *** ___
///   inline        **bold** *italic* `code` [text](url)
///
/// Anything unrecognized degrades to a plain paragraph — the parser NEVER
/// throws (mirrors MermaidParser's "unknown tokens degrade to null" contract).
/// Zero third-party dependencies by design (NativeAOT / WASM clean).
/// </summary>
public static class MarkdownParser
{
    private static readonly Regex HeadingRe = new(@"^(#{1,6})\s+(.*?)\s*#*\s*$");
    private static readonly Regex UnorderedRe = new(@"^(\s*)[-*+]\s+(.*)$");
    private static readonly Regex OrderedRe = new(@"^(\s*)\d+[.)]\s+(.*)$");
    private static readonly Regex FenceRe = new(@"^\s*```+\s*([\w+-]*)\s*$");
    private static readonly Regex RuleRe = new(@"^\s*([-*_])(\s*\1){2,}\s*$");
    private static readonly Regex QuoteRe = new(@"^\s*>\s?(.*)$");
    // The lookahead requires at least one pipe so a single-column delimiter
    // (`|---|`) is accepted while a bare thematic break (`---`) is not — the
    // rule check runs first, but a pipe-less line must never read as a table
    // delimiter regardless of check order.
    private static readonly Regex TableDelimRe = new(@"^(?=.*\|)\s*\|?\s*:?-{1,}:?\s*(\|\s*:?-{1,}:?\s*)*\|?\s*$");

    public static MarkdownDocument Parse(string text)
    {
        var doc = new MarkdownDocument();
        if (string.IsNullOrEmpty(text)) return doc;

        var lines = text.Replace("\r\n", "\n").Replace('\r', '\n').Split('\n');
        int i = 0;
        while (i < lines.Length)
        {
            var line = lines[i];

            // blank line — skip
            if (line.Trim().Length == 0) { i++; continue; }

            // fenced code block
            var fence = FenceRe.Match(line);
            if (fence.Success)
            {
                var lang = fence.Groups[1].Value;
                var sb = new StringBuilder();
                i++;
                while (i < lines.Length && !FenceRe.IsMatch(lines[i]))
                    sb.AppendLine(lines[i++]);
                if (i < lines.Length) i++; // consume closing fence
                doc.Blocks.Add(new MdCodeBlock
                {
                    Language = string.IsNullOrEmpty(lang) ? null : lang,
                    Code = sb.ToString().TrimEnd('\n'),
                });
                continue;
            }

            // thematic break
            if (RuleRe.IsMatch(line)) { doc.Blocks.Add(new MdHorizontalRule()); i++; continue; }

            // heading
            var h = HeadingRe.Match(line);
            if (h.Success)
            {
                doc.Blocks.Add(new MdHeading
                {
                    Level = h.Groups[1].Value.Length,
                    Inlines = ParseInlines(h.Groups[2].Value),
                });
                i++;
                continue;
            }

            // table: a pipe row immediately followed by a delimiter row
            if (LooksLikeTableRow(line) && i + 1 < lines.Length && TableDelimRe.IsMatch(lines[i + 1]))
            {
                var table = ParseTable(lines, ref i);
                doc.Blocks.Add(table);
                continue;
            }

            // blockquote
            if (QuoteRe.IsMatch(line))
            {
                var sb = new StringBuilder();
                while (i < lines.Length && QuoteRe.IsMatch(lines[i]))
                {
                    sb.Append(QuoteRe.Match(lines[i]).Groups[1].Value).Append(' ');
                    i++;
                }
                doc.Blocks.Add(new MdBlockQuote { Inlines = ParseInlines(sb.ToString().Trim()) });
                continue;
            }

            // list (ordered or unordered)
            if (UnorderedRe.IsMatch(line) || OrderedRe.IsMatch(line))
            {
                var list = ParseList(lines, ref i, baseIndent: 0);
                doc.Blocks.Add(list);
                continue;
            }

            // otherwise: paragraph — gather consecutive non-blank, non-structural lines.
            // A table start (row + delimiter on the next line) interrupts the
            // paragraph (GFM) — LLM output routinely omits the blank line
            // before a table; without this check the whole table was swallowed
            // into the paragraph text.
            bool IsTableStart(int at) => LooksLikeTableRow(lines[at])
                && at + 1 < lines.Length && TableDelimRe.IsMatch(lines[at + 1]);
            var para = new StringBuilder();
            while (i < lines.Length && lines[i].Trim().Length > 0
                   && !HeadingRe.IsMatch(lines[i]) && !FenceRe.IsMatch(lines[i])
                   && !RuleRe.IsMatch(lines[i]) && !QuoteRe.IsMatch(lines[i])
                   && !UnorderedRe.IsMatch(lines[i]) && !OrderedRe.IsMatch(lines[i])
                   && !IsTableStart(i))
            {
                if (para.Length > 0) para.Append(' ');
                para.Append(lines[i].Trim());
                i++;
            }
            doc.Blocks.Add(new MdParagraph { Inlines = ParseInlines(para.ToString()) });
        }

        return doc;
    }

    // ─────────────────────────── lists ───────────────────────────

    private static MdList ParseList(string[] lines, ref int i, int baseIndent)
    {
        bool ordered = OrderedRe.IsMatch(lines[i]);
        var list = new MdList { Ordered = ordered };
        MdListItem? current = null;

        while (i < lines.Length)
        {
            var m = UnorderedRe.Match(lines[i]);
            var o = OrderedRe.Match(lines[i]);
            if (!m.Success && !o.Success) break;

            var match = m.Success ? m : o;
            int indent = match.Groups[1].Value.Length;

            if (indent < baseIndent) break;              // belongs to an outer list

            // A marker-type switch at the same level starts a NEW list
            // (CommonMark): `- a` then `1. b` are two lists, not one. Break so
            // the caller opens a fresh list of the other kind.
            if (indent == baseIndent && o.Success != ordered) break;

            if (indent > baseIndent)                     // nested list under current item
            {
                if (current != null)
                    // Append — never assign. One item can own several nested
                    // segments (marker switch mid-nest, partial dedent); a
                    // single-slot assignment overwrote the earlier segment and
                    // its items vanished from the document.
                    current.Children.Add(ParseList(lines, ref i, indent));
                else
                    i++;                                 // orphan indent — skip defensively
                continue;
            }

            current = new MdListItem { Inlines = ParseInlines(match.Groups[2].Value) };
            list.Items.Add(current);
            i++;
        }

        return list;
    }

    // ─────────────────────────── tables ───────────────────────────

    private static bool LooksLikeTableRow(string line) => line.TrimStart().StartsWith("|") || line.Contains(" | ");

    private static MdTable ParseTable(string[] lines, ref int i)
    {
        var table = new MdTable();
        foreach (var cell in SplitRow(lines[i])) table.Header.Add(ParseInlines(cell));
        i += 2; // header + delimiter row

        while (i < lines.Length && lines[i].Contains('|'))
        {
            var row = new List<List<MdSpan>>();
            foreach (var cell in SplitRow(lines[i])) row.Add(ParseInlines(cell));
            table.Rows.Add(row);
            i++;
        }
        return table;
    }

    private static IEnumerable<string> SplitRow(string line)
    {
        var t = line.Trim();
        if (t.StartsWith("|")) t = t[1..];
        if (t.EndsWith("|")) t = t[..^1];
        // NOTE: escaped \| inside cells is not yet handled — MVP scope.
        foreach (var part in t.Split('|')) yield return part.Trim();
    }

    // ─────────────────────────── inline ───────────────────────────

    private static readonly Regex LinkRe = new(@"\[([^\]]*)\]\(([^)]*)\)");

    /// <summary>
    /// Inline scanner for **bold**, *italic*/_italic_, `code`, [text](url).
    /// Emits a flat list of spans, each carrying cumulative flags. Simple
    /// left-to-right scan; nested emphasis collapses onto the same span.
    ///
    /// Delimiter guards (subset of CommonMark's flanking rules, chosen so that
    /// NON-markup characters are never eaten — text loss is worse than a
    /// missed emphasis):
    ///  - an opener needs a matching closer later in the line, else literal
    ///    (`**x` unclosed keeps its asterisks);
    ///  - an opener must be followed by non-space (`2 * 3` stays verbatim);
    ///  - `_`/`__` never toggle intraword (`my_var_name` stays verbatim).
    /// Link text is re-parsed recursively so `[**Bold**](url)` yields a bold
    /// span, not literal asterisks.
    /// </summary>
    public static List<MdSpan> ParseInlines(string text)
    {
        var spans = new List<MdSpan>();
        if (string.IsNullOrEmpty(text)) return spans;

        int pos = 0;
        var buf = new StringBuilder();
        bool bold = false, italic = false;

        void Flush()
        {
            if (buf.Length == 0) return;
            spans.Add(new MdSpan { Text = buf.ToString(), Bold = bold, Italic = italic });
            buf.Clear();
        }

        // `_` must not toggle inside a word (CommonMark: no intraword _ emphasis).
        bool IntrawordUnderscore(char delim, int delimStart, int delimLen)
            => delim == '_'
               && delimStart > 0 && char.IsLetterOrDigit(text[delimStart - 1])
               && delimStart + delimLen < text.Length && char.IsLetterOrDigit(text[delimStart + delimLen]);

        while (pos < text.Length)
        {
            char c = text[pos];

            // inline code `...` (only with a closing backtick; else literal)
            if (c == '`')
            {
                int end = text.IndexOf('`', pos + 1);
                if (end > pos)
                {
                    Flush();
                    spans.Add(new MdSpan { Text = text[(pos + 1)..end], Code = true, Bold = bold, Italic = italic });
                    pos = end + 1;
                    continue;
                }
            }

            // link [text](url) — re-parse the text so inline markers inside
            // the brackets format instead of leaking literally.
            if (c == '[')
            {
                var lm = LinkRe.Match(text, pos);
                if (lm.Success && lm.Index == pos)
                {
                    Flush();
                    foreach (var inner in ParseInlines(lm.Groups[1].Value))
                        spans.Add(new MdSpan
                        {
                            Text = inner.Text,
                            Bold = inner.Bold || bold,
                            Italic = inner.Italic || italic,
                            Code = inner.Code,
                            Href = lm.Groups[2].Value,
                        });
                    pos += lm.Length;
                    continue;
                }
            }

            // strong **...** or __...__
            if ((c == '*' || c == '_') && pos + 1 < text.Length && text[pos + 1] == c)
            {
                var d = new string(c, 2);
                if (bold)
                {
                    Flush(); bold = false; pos += 2; continue;
                }
                bool canOpen = pos + 2 < text.Length
                               && !char.IsWhiteSpace(text[pos + 2])
                               && text.IndexOf(d, pos + 2, StringComparison.Ordinal) >= 0
                               && !IntrawordUnderscore(c, pos, 2);
                if (canOpen)
                {
                    Flush(); bold = true; pos += 2; continue;
                }
                buf.Append(d); pos += 2; continue;
            }

            // emphasis *...* or _..._
            if (c == '*' || c == '_')
            {
                if (italic)
                {
                    // Closer must hug the text (right-flanking): `a * b` stays literal.
                    if (pos > 0 && !char.IsWhiteSpace(text[pos - 1]) && !IntrawordUnderscore(c, pos, 1))
                    {
                        Flush(); italic = false; pos++; continue;
                    }
                    buf.Append(c); pos++; continue;
                }
                bool canOpen = pos + 1 < text.Length
                               && !char.IsWhiteSpace(text[pos + 1])
                               && text.IndexOf(c, pos + 1) >= 0
                               && !IntrawordUnderscore(c, pos, 1);
                if (canOpen)
                {
                    Flush(); italic = true; pos++; continue;
                }
                buf.Append(c); pos++; continue;
            }

            buf.Append(c);
            pos++;
        }
        Flush();
        return spans;
    }
}
