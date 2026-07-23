// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeCli.Core;
using OfficeCli.Core.Markdown;

namespace OfficeCli.Handlers;

public partial class WordHandler
{
    /// <summary>
    /// `add --type markdown` — the docx analogue of `add --type diagram`.
    /// Parses a Markdown subset via the shared, format-neutral
    /// <see cref="MarkdownParser"/> (Core/Markdown, like MermaidParser) and
    /// expands it into ordinary paragraphs / tables at <paramref name="parentPath"/>.
    ///
    /// Like `diagram`, this is ADD-ONLY: there is no persistent "markdown"
    /// node — the expansion produces normal Word elements, so edit those
    /// afterwards (no matching Set/Get/Query/Remove). The block→element and
    /// inline→run mapping reuses the handler's OWN <c>Add</c>/<c>Set</c> entry
    /// points (each returns the real path), so no OOXML is hand-built here and
    /// no Core→handler dependency is introduced.
    ///
    /// Input mirrors `diagram`: canonical `markdown` (+ aliases `text`/`md`)
    /// inline, or `src`/`path` to a .md file.
    /// </summary>
    private string AddMarkdown(OpenXmlElement parent, string parentPath, int? index, Dictionary<string, string> properties)
    {
        var mdText = properties.GetValueOrDefault("markdown")
                     ?? properties.GetValueOrDefault("text")
                     ?? properties.GetValueOrDefault("md");
        if (string.IsNullOrWhiteSpace(mdText)
            && (properties.TryGetValue("src", out var srcFile) || properties.TryGetValue("path", out srcFile))
            && !string.IsNullOrWhiteSpace(srcFile))
        {
            if (!System.IO.File.Exists(srcFile))
                throw new ArgumentException($"markdown source file not found: '{srcFile}'.");
            mdText = System.IO.File.ReadAllText(srcFile);
        }
        if (string.IsNullOrWhiteSpace(mdText))
            throw new ArgumentException("markdown requires inline 'markdown' text (aliases: text, md) or a 'src' .md file path.");

        var doc = MarkdownParser.Parse(mdText);

        // Position: the first block honors the caller's --index/--after/--before;
        // each subsequent block chains AfterElement(previous) so document order
        // is preserved when inserting mid-document. With no position, every
        // block appends (pos = null). Requires the exit-invalidate guard in
        // Add() — anchor navigation used to leave a stale body child-index
        // (see WordAddAnchorCacheTests) which broke the per-cell Set after an
        // anchored `add table`.
        string? firstPath = null;
        string? lastPath = null;

        // Two-phase: phase 1 adds every block (pure appends stay on the
        // append-fast path), phase 2 applies all inline range formatting.
        // Interleaving Add + Set-by-path was O(n²): every body-level Add
        // clears the body child-index/paraId caches, so each Set's path
        // navigation rebuilt them from scratch (8k formatted lines ≈ 147s).
        // Deferred, the first Set rebuilds once and the rest hit the cache.
        var pendingInline = new List<(string Path, List<MdSpan> Inlines)>();

        InsertPosition? PosFor()
        {
            if (lastPath != null) return InsertPosition.AfterElement(lastPath);
            if (index.HasValue) return InsertPosition.AtIndex(index.Value);
            return null;
        }

        void Record(string path)
        {
            firstPath ??= path;
            lastPath = path;
        }

        foreach (var block in doc.Blocks)
        {
            switch (block)
            {
                case MdHeading h:
                    Record(AddHeading(parentPath, h, PosFor(), pendingInline));
                    break;
                case MdParagraph p:
                    Record(AddFlowParagraph(parentPath, p.Inlines, null, pendingInline, pos: PosFor()));
                    break;
                case MdBlockQuote q:
                    // Blockquote: left-indented + italic direct formatting. The
                    // built-in "Quote" style is absent from a blank docx, so we
                    // format directly rather than reference a style that renders
                    // as plain body text.
                    Record(AddFlowParagraph(parentPath, q.Inlines, null, pendingInline,
                        extra: new(StringComparer.OrdinalIgnoreCase) { ["indentLeft"] = "720", ["italic"] = "true" },
                        pos: PosFor()));
                    break;
                case MdList list:
                    AddFlowList(parentPath, list, depth: 0, Record, PosFor, pendingInline);
                    break;
                case MdCodeBlock code:
                    // Monospace via direct font; built-in code styles are absent
                    // from a blank docx.
                    foreach (var line in code.Code.Split('\n'))
                        Record(Add(parentPath, "paragraph", PosFor(),
                            new(StringComparer.OrdinalIgnoreCase) { ["font"] = "Consolas", ["text"] = line }));
                    break;
                case MdHorizontalRule:
                    Record(Add(parentPath, "paragraph", PosFor(),
                        new(StringComparer.OrdinalIgnoreCase) { ["borderBottom"] = "single" }));
                    break;
                case MdTable t:
                    Record(AddFlowTable(parentPath, t, PosFor()));
                    break;
            }
        }

        // Phase 2: all structural adds done — apply inline formatting in one
        // sweep so path navigation builds the body caches once and reuses them.
        // DeferSave for the sweep: each public Set's range branch ends in
        // SaveDoc(), i.e. a full main-part serialization — per formatted line
        // that's the same O(N²) re-serialize batch replay hit (see
        // RunNonResidentBatch). One flush at the end instead; restore the
        // caller's flag so resident semantics are untouched.
        foreach (var (path, inlines) in pendingInline)
            ApplyInlineSpans(path, inlines);

        // Empty markdown still yields at least an anchor so `add` returns a path.
        return firstPath ?? Add(parentPath, "paragraph", PosFor(),
            new(StringComparer.OrdinalIgnoreCase) { ["text"] = "" });
    }

    // Heading font sizes (pt) per level — a compact hierarchy that reads as
    // headings even in a blank docx with no built-in Heading styles defined.
    private static readonly int[] HeadingSizes = { 18, 16, 14, 13, 12, 11 };

    private string AddHeading(string parentPath, MdHeading h, InsertPosition? pos,
        List<(string Path, List<MdSpan> Inlines)> pendingInline)
    {
        int size = HeadingSizes[Math.Clamp(h.Level, 1, 6) - 1];
        // Reference the built-in style id AND apply direct bold+size. The style
        // ref keeps TOC / navigation working if the doc later gains Heading
        // definitions; the direct formatting guarantees the visual hierarchy
        // now (a blank docx ships Heading1..9 as bare scaffolding that renders
        // like body text).
        var extra = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            ["style"] = $"Heading{Math.Clamp(h.Level, 1, 6)}",
            ["bold"] = "true",
            ["size"] = size.ToString(),
        };
        return AddFlowParagraph(parentPath, h.Inlines, null, pendingInline, extra, pos);
    }

    private string AddFlowParagraph(string parentPath, List<MdSpan> inlines, string? style,
        List<(string Path, List<MdSpan> Inlines)> pendingInline,
        Dictionary<string, string>? extra = null, InsertPosition? pos = null)
    {
        var props = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            ["text"] = string.Concat(inlines.Select(s => s.Text)),
        };
        if (style != null) props["style"] = style;
        if (extra != null)
            foreach (var kv in extra) props[kv.Key] = kv.Value;

        var path = Add(parentPath, "paragraph", pos, props);
        if (inlines.Any(s => s.Bold || s.Italic || s.Code))
            pendingInline.Add((path, inlines)); // formatted in phase 2 (perf)
        return path;
    }

    private void AddFlowList(string parentPath, MdList list, int depth, Action<string> record,
        Func<InsertPosition?> posFor,
        List<(string Path, List<MdSpan> Inlines)> pendingInline)
    {
        // Use liststyle (real numbering: bullets / 1. 2.) rather than a
        // ListBullet/ListNumber style id — those style ids are absent from a
        // blank docx and would render as plain paragraphs with no marker.
        // Nesting maps to the numbering ilvl via `level` (0-based, capped 0-8).
        var listStyle = list.Ordered ? "ordered" : "unordered";
        foreach (var item in list.Items)
        {
            var props = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            {
                ["liststyle"] = listStyle,
                ["level"] = Math.Clamp(depth, 0, 8).ToString(),
                ["text"] = string.Concat(item.Inlines.Select(s => s.Text)),
            };

            var path = Add(parentPath, "paragraph", posFor(), props);
            record(path);
            if (item.Inlines.Any(s => s.Bold || s.Italic || s.Code))
                pendingInline.Add((path, item.Inlines)); // formatted in phase 2 (perf)

            // All nested segments, in order — one item can own several
            // (marker switch / partial dedent), see MdListItem.Children.
            foreach (var child in item.Children)
                AddFlowList(parentPath, child, depth + 1, record, posFor, pendingInline);
        }
    }

    private string AddFlowTable(string parentPath, MdTable t, InsertPosition? pos)
    {
        int cols = t.Header.Count;
        int rows = t.Rows.Count + 1;
        var tablePath = Add(parentPath, "table", pos,
            new(StringComparer.OrdinalIgnoreCase) { ["rows"] = rows.ToString(), ["cols"] = cols.ToString() });

        var grid = new List<List<List<MdSpan>>> { t.Header };
        grid.AddRange(t.Rows);
        for (int r = 0; r < grid.Count; r++)
        {
            var row = grid[r];
            for (int c = 0; c < cols && c < row.Count; c++)
            {
                var text = string.Concat(row[c].Select(s => s.Text));
                if (text.Length == 0) continue;
                Set($"{tablePath}/row[{r + 1}]/cell[{c + 1}]",
                    new(StringComparer.OrdinalIgnoreCase) { ["text"] = text });
            }
        }
        return tablePath;
    }

    /// <summary>
    /// Rebuild a freshly-added paragraph's runs so each inline span becomes its
    /// own run with direct formatting (bold/italic/Consolas). Spans tile the
    /// paragraph text in order, so runs are CONSTRUCTED once instead of
    /// range-split after the fact — a paragraph with k formatted spans used to
    /// issue k `set range=` calls, each rescanning every already-split run:
    /// O(k²) per paragraph, ~2.5min for a 2000-span paragraph (blank-line-free
    /// LLM output gathers into exactly that). Direct construction is O(k).
    /// The base run's properties (e.g. a heading's bold+size) are cloned onto
    /// every span run, then span flags overlay them.
    /// </summary>
    private void ApplyInlineSpans(string paraPath, List<MdSpan> inlines)
    {
        if (NavigateToElement(ParsePath(paraPath)) is not Paragraph para) return;

        var baseProps = para.Elements<Run>().FirstOrDefault()?.RunProperties;
        foreach (var r in para.Elements<Run>().ToList()) r.Remove();

        foreach (var span in inlines)
        {
            if (span.Text.Length == 0) continue;
            var run = new Run();
            var rPr = baseProps != null ? (RunProperties)baseProps.CloneNode(true) : new RunProperties();
            if (span.Bold) rPr.Bold ??= new Bold();
            if (span.Italic) rPr.Italic ??= new Italic();
            if (span.Code) rPr.RunFonts = new RunFonts { Ascii = "Consolas", HighAnsi = "Consolas" };
            if (rPr.HasChildren) run.AppendChild(rPr);

            // \t → <w:tab/>, mirroring AddParagraph's text conversion.
            var parts = span.Text.Split('\t');
            for (int p = 0; p < parts.Length; p++)
            {
                if (p > 0) run.AppendChild(new TabChar());
                if (parts[p].Length > 0)
                    run.AppendChild(new Text(parts[p]) { Space = SpaceProcessingModeValues.Preserve });
            }
            para.AppendChild(run);
        }
    }
}
