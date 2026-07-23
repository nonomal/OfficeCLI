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
    // Content is optional so a BARE marker line (`-`, `1.`) is still a list item
    // (an empty one) rather than degrading to a literal paragraph — the common
    // "marker on its own line, code fence on the next" shape. `**bold**` etc.
    // stay safe: after the single marker char the remainder must be whitespace+
    // content or end-of-line, so `**x` (no space) never matches.
    private static readonly Regex UnorderedRe = new(@"^(\s*)[-*+](?:\s+(.*))?$");
    private static readonly Regex OrderedRe = new(@"^(\s*)\d+[.)](?:\s+(.*))?$");
    private static readonly Regex FenceRe = new(@"^\s*```+\s*([\w+-]*)\s*$");
    // A fence opener sharing its line with a leading list marker (`1. ``` `,
    // `- ```python`). Handled inside ParseList so the fence becomes nested code
    // content of the list item — the list (and its numbering) stays intact and
    // sibling items survive. Groups: 1=indent, 2=marker, 3=info-string/lang.
    private static readonly Regex ListFenceRe = new(@"^(\s*)(\d+[.)]|[-*+])\s+`{3,}\s*([\w+-]*)\s*$");
    // An indented fence opener on its OWN line (no marker) — the code block of a
    // list item written on the line(s) AFTER the item text (the ubiquitous
    // README "1. Step\n   ```bash\n   …\n   ```" install-step shape). Group 1 =
    // indent, group 2 = info-string. Handled as continuation content of the
    // current item so its body is de-indented exactly like the same-line fence.
    private static readonly Regex ContinuationFenceRe = new(@"^(\s+)`{3,}\s*([\w+-]*)\s*$");
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

            // fenced code block (standalone ``` opener). A fence opener that
            // shares its line with a list marker (`1. ```) is NOT handled here —
            // it flows to the list branch below so ParseList can keep it as
            // nested code content of the list item (see ListFenceRe / ParseList).
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

            // table: a pipe row immediately followed by a delimiter row. A line
            // that is itself a list/blockquote marker YIELDS to those blocks —
            // after LooksLikeTableRow was widened to "any interior pipe", a
            // common `- cost | benefit` list item (or `> a|b` quote) followed by
            // a delimiter-shaped line was otherwise swallowed whole into a table
            // with the marker leaking as literal cell text. A real table row
            // never starts with a `-`/`*`/`+`/`N.`/`>` marker, so this only
            // reclaims the misfire.
            if (LooksLikeTableRow(line) && !IsListOrQuoteMarker(line)
                && i + 1 < lines.Length && TableDelimRe.IsMatch(lines[i + 1]))
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
            bool IsTableStart(int at) => LooksLikeTableRow(lines[at]) && !IsListOrQuoteMarker(lines[at])
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
            var lf = ListFenceRe.Match(lines[i]);
            var m = UnorderedRe.Match(lines[i]);
            var o = OrderedRe.Match(lines[i]);
            if (!m.Success && !o.Success)
            {
                // A next-line indented fence is the current item's code content
                // (the README "1. Step\n   ```bash\n   …\n   ```" step shape).
                // Consume it de-indented — structurally identical to a same-line
                // fence — so ParseList does NOT break here (which used to hand
                // the fence to the top-level scanner as a standalone block and
                // orphan the following sibling markers).
                var cf = ContinuationFenceRe.Match(lines[i]);
                if (cf.Success && current != null && cf.Groups[1].Value.Length > baseIndent)
                {
                    current.CodeBlocks.Add(ConsumeFence(lines, ref i, cf.Groups[2].Value));
                    continue;
                }
                break;
            }

            var match = m.Success ? m : o;
            int indent = match.Groups[1].Value.Length;

            if (indent < baseIndent) break;              // belongs to an outer list

            // A marker-type switch at the same level starts a NEW list
            // (CommonMark): `- a` then `1. b` are two lists, not one. Break so
            // the caller opens a fresh list of the other kind.
            if (indent == baseIndent && o.Success != ordered) break;

            // Nested list under the current item.
            if (indent > baseIndent && current != null)
            {
                // Append — never assign. One item can own several nested
                // segments (marker switch mid-nest, partial dedent); a
                // single-slot assignment overwrote the earlier segment and
                // its items vanished from the document.
                current.Children.Add(ParseList(lines, ref i, indent));
                continue;
            }
            // indent > baseIndent with NO parent item is a genuinely orphaned
            // indented marker (reached straight from the top-level scanner).
            // It must NEVER be silently dropped (the parser's degrade-don't-
            // lose-text contract) — fall through and add it as an item of THIS
            // list at the current level.

            // A fence opener sharing the marker line (`1. ``` `, `- ```py`):
            // the item's content is a fenced code block. Keep it as nested item
            // content (list stays intact, numbering consistent, sibling items
            // survive) rather than hoisting it to a top-level block.
            if (lf.Success)
            {
                current = new MdListItem();
                current.CodeBlocks.Add(ConsumeFence(lines, ref i, lf.Groups[3].Value));
                list.Items.Add(current);
                continue;
            }

            current = new MdListItem { Inlines = ParseInlines(match.Groups[2].Value) };
            list.Items.Add(current);
            i++;
        }

        return list;
    }

    /// <summary>
    /// Consume a fenced code block (opener at the current line, <paramref
    /// name="lang"/> already extracted) that belongs to a list item — either
    /// sharing the marker line or opened on a following continuation line.
    /// Advances <paramref name="i"/> past the body and closing fence. Body lines
    /// are de-indented by the minimum leading-space count across non-blank
    /// lines, capped at the column where the opener's backticks began — so
    /// content aligned under the marker (`   code`) is stripped, while
    /// under-indented content is preserved verbatim (never over-strips
    /// meaningful indentation).
    /// </summary>
    private static MdCodeBlock ConsumeFence(string[] lines, ref int i, string lang)
    {
        int fenceCol = lines[i].IndexOf('`');
        i++; // past the opener line
        var body = new List<string>();
        while (i < lines.Length && !FenceRe.IsMatch(lines[i])) body.Add(lines[i++]);
        if (i < lines.Length) i++; // consume closing fence

        int strip = fenceCol < 0 ? 0 : fenceCol;
        foreach (var bl in body)
            if (bl.Trim().Length > 0) strip = Math.Min(strip, LeadingSpaces(bl));

        var stripped = new List<string>(body.Count);
        foreach (var bl in body) stripped.Add(StripLeading(bl, strip));
        return new MdCodeBlock
        {
            Language = string.IsNullOrEmpty(lang) ? null : lang,
            Code = string.Join("\n", stripped).TrimEnd('\n'),
        };
    }

    private static int LeadingSpaces(string s)
    {
        int n = 0;
        while (n < s.Length && s[n] == ' ') n++;
        return n;
    }

    private static string StripLeading(string s, int n)
    {
        int k = 0;
        while (k < n && k < s.Length && s[k] == ' ') k++;
        return s[k..];
    }

    // ─────────────────────────── tables ───────────────────────────

    // A header/body row candidate: a leading pipe, OR any interior pipe. GFM
    // does NOT require spaces around header pipes ("H1|H2" is a valid header),
    // so a space-flanked " | " check wrongly rejected unspaced headers. Any
    // pipe qualifies; the real gate is the caller's requirement that the NEXT
    // line be a delimiter row (TableDelimRe), which keeps prose that merely
    // contains a '|' from being misread as a table.
    private static bool LooksLikeTableRow(string line) => line.Contains('|');

    // A line that opens a list item or blockquote. Table detection (both the
    // block dispatch AND the row-collection loop) must yield to these: after
    // table-row detection was widened to "any interior pipe", a list item or
    // quote whose text merely contains a '|' would otherwise be swallowed as a
    // table row with its marker leaking as cell text. A genuine table row never
    // starts with a -/*/+/N./> marker.
    private static bool IsListOrQuoteMarker(string line) =>
        UnorderedRe.IsMatch(line) || OrderedRe.IsMatch(line) || QuoteRe.IsMatch(line);

    private static MdTable ParseTable(string[] lines, ref int i)
    {
        var table = new MdTable();
        foreach (var cell in SplitRow(lines[i])) table.Header.Add(ParseInlines(cell));
        i += 2; // header + delimiter row

        // Collect body rows, but YIELD as soon as a line is itself a list/quote
        // marker: an open table otherwise kept swallowing every later "contains
        // a |" line (e.g. `1. item | pipe`) as a table row, ballooning one table
        // over content that was really a mix of list items and prose. Mirrors
        // the block-dispatch yield rule so both entry points agree.
        while (i < lines.Length && lines[i].Contains('|') && !IsListOrQuoteMarker(lines[i]))
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

    /// <summary>
    /// Matches a <c>[label](dest)</c> construct anchored at <paramref name="start"/>
    /// (which must point at the <c>[</c>). Unlike a flat regex this counts bracket
    /// and paren depth so a nested image/link inside the label survives:
    /// <c>[![alt](img)](target)</c> yields label <c>![alt](img)</c> + dest
    /// <c>target</c>, instead of the old <c>LinkRe</c> stopping at the first
    /// inner <c>]</c> and leaking the trailing <c>](target)</c> as literal text.
    /// </summary>
    private static bool TryMatchLinkSyntax(string text, int start, out string label, out string dest, out int length)
    {
        label = dest = "";
        length = 0;
        if (start >= text.Length || text[start] != '[') return false;

        int depth = 0, labelEnd = -1;
        for (int i = start; i < text.Length; i++)
        {
            if (text[i] == '[') depth++;
            else if (text[i] == ']') { if (--depth == 0) { labelEnd = i; break; } }
        }
        if (labelEnd < 0) return false;

        int p = labelEnd + 1;
        if (p >= text.Length || text[p] != '(') return false;

        int pdepth = 0, destEnd = -1;
        for (int i = p; i < text.Length; i++)
        {
            if (text[i] == '(') pdepth++;
            else if (text[i] == ')') { if (--pdepth == 0) { destEnd = i; break; } }
        }
        if (destEnd < 0) return false;

        label = text[(start + 1)..labelEnd];
        dest = text[(p + 1)..destEnd];
        length = destEnd + 1 - start;
        return true;
    }

    /// <summary>
    /// True when index <paramref name="i"/> begins a delimiter run of EXACTLY
    /// two tildes — i.e. <c>~~</c> not flanked by a further <c>~</c> on either
    /// side. GFM strikethrough uses a length-2 run; a 3+ tilde run
    /// (<c>~~~ literal ~~~</c>) is not a strikethrough delimiter, and treating
    /// the <c>~~</c> substring inside it as an open/close pair silently ate two
    /// tildes per side — a text-loss bug the parser's "text loss is worse than a
    /// missed emphasis" guard forbids.
    /// </summary>
    private static bool IsTwoTildeRun(string s, int i)
        => i >= 0 && i + 1 < s.Length && s[i] == '~' && s[i + 1] == '~'
           && (i == 0 || s[i - 1] != '~')
           && (i + 2 >= s.Length || s[i + 2] != '~');

    /// <summary>
    /// Is a clean length-2 tilde run reachable from <paramref name="from"/> to
    /// serve as a strikethrough closer, WITHOUT first crossing a 3+ tilde run?
    ///
    /// A blind "does a clean ~~ exist anywhere later" scan cross-matched
    /// independent regions: in <c>a~~b~~~c and ~~**bold**~~ done</c> it let the
    /// first (invalid) <c>~~</c> opener claim the SECOND region's OPENING <c>~~</c>
    /// as its closer, corrupting both regions. A 3+ tilde run (which is itself
    /// not a valid delimiter) acts as a barrier: an opener may not pair across
    /// it. So here the first opener sees the <c>~~~</c> barrier before any clean
    /// closer and stays literal, while the second region pairs on its own.
    /// </summary>
    private static bool HasReachableTwoTildeCloser(string s, int from)
    {
        int i = from;
        while (i < s.Length)
        {
            // Inline-code spans are opaque to tilde pairing. The main scanner's
            // code branch consumes `...` as literal code BEFORE the tilde branch
            // ever sees it, so a ~~ buried inside a code span is never a real
            // closer — skip the whole span here too, or the opener would be told
            // a closer exists and open a strike that can never close (bleeding
            // strike into all trailing text). Mirror the main scanner: a
            // backtick with a later backtick is a span (skip past it); an
            // unterminated backtick is literal (advance one).
            if (s[i] == '`')
            {
                int close = s.IndexOf('`', i + 1);
                i = close > i ? close + 1 : i + 1;
                continue;
            }
            if (s[i] != '~') { i++; continue; }
            int j = i;
            while (j < s.Length && s[j] == '~') j++;
            int runLen = j - i;
            if (runLen == 2) return true;  // clean 2-run — a usable closer
            if (runLen >= 3) return false; // 3+ run barrier — no pairing across it
            i = j;                         // length-1 run: not a delimiter, skip
        }
        return false;
    }

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
        bool bold = false, italic = false, strike = false;

        void Flush()
        {
            if (buf.Length == 0) return;
            spans.Add(new MdSpan { Text = buf.ToString(), Bold = bold, Italic = italic, Strike = strike });
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
                    spans.Add(new MdSpan { Text = text[(pos + 1)..end], Code = true, Bold = bold, Italic = italic, Strike = strike });
                    pos = end + 1;
                    continue;
                }
            }

            // image ![alt](url) — no picture is embedded (inline markdown is
            // text-only), so it degrades to "!" + alt text, mirroring the flat
            // image path. Recognised BEFORE the link branch so the common
            // "badge links to CI" idiom [![alt](img)](target) doesn't leak raw
            // syntax: here the alt/img are consumed as one token, and the outer
            // link branch (bracket-depth aware) wraps the result.
            if (c == '!' && pos + 1 < text.Length && text[pos + 1] == '['
                && TryMatchLinkSyntax(text, pos + 1, out var alt, out var imgUrl, out var imgLen))
            {
                Flush();
                spans.Add(new MdSpan { Text = "!", Bold = bold, Italic = italic });
                var altSpans = ParseInlines(alt);
                if (altSpans.Count == 0)
                {
                    // Empty alt (`![](url)`): the URL would otherwise vanish
                    // entirely, leaving no trace of the image. Surface the URL
                    // itself as the visible text so nothing is silently lost —
                    // consistent with the "never lose text" degradation.
                    spans.Add(new MdSpan { Text = imgUrl, Bold = bold, Italic = italic, Href = imgUrl });
                }
                foreach (var inner in altSpans)
                    spans.Add(new MdSpan
                    {
                        Text = inner.Text,
                        Bold = inner.Bold || bold,
                        Italic = inner.Italic || italic,
                        Code = inner.Code,
                        Strike = inner.Strike || strike,
                        Href = imgUrl,
                    });
                pos += 1 + imgLen;
                continue;
            }

            // link [text](url) — re-parse the text so inline markers inside
            // the brackets format instead of leaking literally. Bracket-depth
            // aware so a nested image/link in the label survives intact.
            if (c == '[' && TryMatchLinkSyntax(text, pos, out var label, out var url, out var linkLen))
            {
                Flush();
                foreach (var inner in ParseInlines(label))
                    spans.Add(new MdSpan
                    {
                        Text = inner.Text,
                        Bold = inner.Bold || bold,
                        Italic = inner.Italic || italic,
                        Code = inner.Code,
                        Strike = inner.Strike || strike,
                        Href = url,
                    });
                pos += linkLen;
                continue;
            }

            // GFM strikethrough ~~...~~ (single ~ never toggles; unclosed
            // opener stays literal — same text-loss-averse guard as **). Only a
            // CLEAN length-2 tilde run toggles: a 3+ run (`~~~ literal ~~~`) is
            // not a delimiter, and the closer search likewise ignores `~~`
            // buried inside a 3+ run, so no literal tildes are ever swallowed.
            if (c == '~' && pos + 1 < text.Length && text[pos + 1] == '~')
            {
                if (IsTwoTildeRun(text, pos))
                {
                    if (strike)
                    {
                        Flush(); strike = false; pos += 2; continue;
                    }
                    bool canOpen = pos + 2 < text.Length
                                   && !char.IsWhiteSpace(text[pos + 2])
                                   && HasReachableTwoTildeCloser(text, pos + 2);
                    if (canOpen)
                    {
                        Flush(); strike = true; pos += 2; continue;
                    }
                    // Clean 2-run with no reachable closer: keep the two tildes
                    // literal, but still close the current span at this delimiter
                    // boundary (as a real opener would), then emit "~~" verbatim
                    // — no text is lost and nothing is struck.
                    Flush(); buf.Append("~~"); pos += 2; continue;
                }
                // 3+ run (or a ~ that isn't a clean 2-run): emit one tilde
                // literally and re-scan the rest (the left-boundary check in
                // IsTwoTildeRun keeps the run from being mis-read as a delimiter
                // on the next character).
                buf.Append('~'); pos++; continue;
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
