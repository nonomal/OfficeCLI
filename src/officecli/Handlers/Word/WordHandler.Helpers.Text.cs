// Copyright 2025 OfficeCLI (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeCli.Core;
using Vml = DocumentFormat.OpenXml.Vml;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using M = DocumentFormat.OpenXml.Math;

namespace OfficeCli.Handlers;

public partial class WordHandler
{

    /// <summary>
    /// Get footnote/endnote text, skipping the reference mark run and its trailing space.
    /// </summary>
    private static string GetFootnoteText(OpenXmlElement fnOrEn)
    {
        return string.Join("", fnOrEn.Descendants<Run>()
            .Where(r => r.GetFirstChild<FootnoteReferenceMark>() == null
                     && r.GetFirstChild<EndnoteReferenceMark>() == null)
            .SelectMany(r => r.Elements<Text>())
            .Select(t => t.Text)).TrimStart();
    }

    private static string GetParagraphText(Paragraph para)
    {
        // CONSISTENCY(run-text-tab): use GetRunText so <w:tab/> renders as
        // \t in the paragraph readback (was silently dropped, breaking
        // dump round-trip for tabbed content).
        var sb = new StringBuilder();
        foreach (var child in para.ChildElements)
        {
            if (child is Run run)
                sb.Append(GetRunText(run));
            else if (child is Hyperlink hyperlink)
            {
                foreach (var hChild in hyperlink.ChildElements)
                {
                    if (hChild is Run hRun) sb.Append(GetRunText(hRun));
                    else if (hChild.LocalName == "oMath" || hChild is M.OfficeMath)
                        sb.Append(string.Concat(hChild.Descendants<Text>().Select(t => t.Text))
                            + string.Concat(hChild.Descendants<M.Text>().Select(t => t.Text)));
                }
            }
            else if (child.LocalName == "oMath" || child is M.OfficeMath)
            {
                // BUG-DUMP9-04: inline equations contribute readable text to the
                // paragraph readback so dump round-trip can verify formula
                // survival. Use raw m:t / w:t descendants (not LaTeX) so the
                // glyphs match the source.
                sb.Append(string.Concat(child.Descendants<Text>().Select(t => t.Text))
                    + string.Concat(child.Descendants<M.Text>().Select(t => t.Text)));
            }
        }
        return sb.ToString();
    }

    /// <summary>
    /// Get paragraph text including inline math rendered as readable Unicode.
    /// </summary>
    private static string GetParagraphTextWithMath(Paragraph para)
    {
        var sb = new StringBuilder();
        foreach (var child in para.ChildElements)
        {
            if (child is Run run)
                sb.Append(GetRunText(run));
            else if (child.LocalName == "oMath" || child is M.OfficeMath)
                sb.Append(FormulaParser.ToReadableText(child));
            else if (child is Hyperlink hyperlink)
                sb.Append(string.Concat(hyperlink.Descendants<Text>().Select(t => t.Text)));
        }
        return sb.ToString();
    }

    /// <summary>
    /// Find math elements in a paragraph using both type and localName matching.
    /// </summary>
    private static List<OpenXmlElement> FindMathElements(Paragraph para)
    {
        return para.ChildElements
            .Where(e => e.LocalName == "oMath" || e is M.OfficeMath)
            .ToList();
    }

    /// <summary>
    /// Get all body-level elements, flattening SdtContent containers.
    /// This ensures paragraphs and tables inside w:sdt are not missed.
    /// </summary>
    private static IEnumerable<OpenXmlElement> GetBodyElements(Body body)
    {
        foreach (var element in FlattenWrappers(body.ChildElements))
            yield return element;
    }

    // Descend into SDT (structured document tag) and customXml transparent
    // wrappers so their wrapped paragraphs/tables participate in the body
    // element axis. Without this, docs emitted by e.g. Pages/Google Docs
    // that wrap entire sections in <w:customXml> produce an empty preview.
    private static IEnumerable<OpenXmlElement> FlattenWrappers(IEnumerable<OpenXmlElement> elements)
    {
        foreach (var element in elements)
        {
            if (element is SdtBlock sdt)
            {
                var content = sdt.SdtContentBlock;
                if (content != null)
                    foreach (var child in FlattenWrappers(content.ChildElements))
                        yield return child;
            }
            else if (element.LocalName == "customXml"
                && element.NamespaceUri == "http://schemas.openxmlformats.org/wordprocessingml/2006/main")
            {
                foreach (var child in FlattenWrappers(element.ChildElements))
                    yield return child;
            }
            else
            {
                yield return element;
            }
        }
    }

    /// <summary>
    /// Checks if an element is a structural document element worth displaying
    /// (not inline markers like bookmarkStart, bookmarkEnd, proofErr, etc.)
    /// </summary>
    private static bool IsStructuralElement(OpenXmlElement element)
    {
        var name = element.LocalName;
        return name == "sectPr" || name == "altChunk" || name == "customXml";
    }

    /// <summary>
    /// Get all Run elements in a paragraph, including those nested inside
    /// Hyperlink and SdtContent containers.
    /// </summary>
    private static List<Run> GetAllRuns(Paragraph para)
    {
        return para.Descendants<Run>()
            .Where(r => r.GetFirstChild<CommentReference>() == null)
            // BUG-DUMP4-06: skip runs nested inside an inline SdtRun. Those
            // runs are surfaced separately as a typed `sdt` paragraph child so
            // alias/tag/type metadata round-trips. Without this filter the
            // inner run was emitted twice — once unwrapped (losing metadata)
            // and once via the sdt branch.
            .Where(r => r.Ancestors<SdtRun>().FirstOrDefault() == null)
            // BUG-DUMP6-01: skip runs nested inside <w:fldSimple>. Those
            // runs are surfaced separately as a typed `field` paragraph child
            // carrying the SimpleField.Instruction attribute. Without this
            // filter the inner display run was emitted as a plain run and
            // the field instruction was silently dropped on dump round-trip.
            .Where(r => r.Ancestors<SimpleField>().FirstOrDefault() == null)
            // BUG-DUMP-TXBX: skip runs whose nearest TextBoxContent ancestor
            // sits BELOW the current paragraph (i.e. the run lives inside a
            // textbox that is a descendant of `para`). Those runs are
            // surfaced separately under /<host>/textbox[N]/p[M]/r[K] via the
            // textbox navigation branch and the WordBatchEmitter typed
            // `add textbox` recursion. We must NOT skip runs whose para is
            // itself inside TextBoxContent (the inner paragraphs of a
            // textbox) — for those, no TextBoxContent sits between the run
            // and `para`, so they pass through and emit normally.
            .Where(r =>
            {
                // Drop the run iff its nearest TextBoxContent ancestor is a
                // DESCENDANT of `para` (a textbox lives under this para and
                // this run sits inside it). Keep when no TextBoxContent
                // exists, or when the TextBoxContent ancestor sits at-or-
                // above `para` (meaning `para` itself is the textbox-inner
                // paragraph — emitting its runs is the desired behavior).
                var tbc = r.Ancestors<TextBoxContent>().FirstOrDefault();
                if (tbc == null) return true;
                // tbc is a descendant of `para`? walk tbc's ancestors and
                // check whether `para` is among them.
                foreach (var anc in tbc.Ancestors())
                {
                    if (ReferenceEquals(anc, para)) return false;
                }
                return true;
            })
            .ToList();
    }

    private static string GetRunText(Run run)
    {
        // CONSISTENCY(run-text-tab): walk run children in document order so
        // <w:tab/> renders as \t in the readback. Plain Elements<Text>() drops
        // tabs silently, which broke dump round-trip (the tab IS in the XML
        // because AddText splits on \t and emits TabChar — but Get hid it).
        var sb = new System.Text.StringBuilder();
        foreach (var child in run.Elements())
        {
            switch (child)
            {
                case Text t: sb.Append(t.Text); break;
                case TabChar: sb.Append('\t'); break;
                // CONSISTENCY(text-breaks): mirror AppendTextWithBreaks — \n
                // round-trips through <w:br/> (textWrapping, the OOXML default
                // when w:type is absent). Without this case, Set/Add(text=...)
                // with embedded \n loses the break on dump readback. Skip
                // page/column breaks — they have no \n source representation
                // and a paragraph-level `break` property already captures them.
                case Break br when br.Type == null || br.Type.Value == BreakValues.TextWrapping:
                    sb.Append('\n'); break;
                // BUG-DUMP7-01: <w:sym w:font="Wingdings" w:char="F0E0"/> is a
                // glyph substitution — the run carries no <w:t>. Without a case
                // here, GetRunText returned empty and WordBatchEmitter's run-emit
                // dropped the whole run, silently losing the symbol on dump
                // round-trip. Surface the resolved Unicode code point as Text
                // so the run looks non-empty; the canonical `sym` Format key
                // (set in Navigation.cs) carries the font+char metadata that
                // AddRun consumes to rebuild the SymbolChar element verbatim.
                case SymbolChar symChild:
                {
                    var charHex = symChild.Char?.Value;
                    if (!string.IsNullOrEmpty(charHex)
                        && int.TryParse(charHex, System.Globalization.NumberStyles.HexNumber,
                            System.Globalization.CultureInfo.InvariantCulture, out var symCode))
                        sb.Append(char.ConvertFromUtf32(symCode));
                    break;
                }
                // BUG-DUMP4-01: a Run nested inside a w:del wrapper carries its
                // text in <w:delText> (DeletedText), not <w:t>. Without this
                // case the deleted content was silently dropped from Get
                // readback and dump round-trip — the inner Run was reachable
                // via Descendants<Run>() but appeared empty.
                case DeletedText dt: sb.Append(dt.Text); break;
                // BUG-DUMP5-03: inline character elements that carry no <w:t>
                // child but contribute visible glyphs. Map to their Unicode
                // equivalents so dump→batch round-trip preserves the visible
                // text. Without this, every <w:noBreakHyphen/> / <w:softHyphen/>
                // dropped to an empty run and disappeared on replay.
                case NoBreakHyphen: sb.Append('‑'); break; // non-breaking hyphen
                case SoftHyphen: sb.Append('­'); break;   // soft hyphen
                // BUG-DUMP5-04: date / time placeholder elements (dayLong /
                // monthLong / yearShort / dayShort / monthShort / yearLong)
                // are auto-substituted by Word at render time. They carry no
                // text in OOXML — surface a stable placeholder so dump
                // captures their presence (otherwise the runs vanish on
                // round-trip and Word has nothing to substitute against).
                case DayLong: sb.Append("[dayLong]"); break;
                case DayShort: sb.Append("[dayShort]"); break;
                case MonthLong: sb.Append("[monthLong]"); break;
                case MonthShort: sb.Append("[monthShort]"); break;
                case YearLong: sb.Append("[yearLong]"); break;
                case YearShort: sb.Append("[yearShort]"); break;
            }
        }
        return sb.ToString();
    }

    private static bool HasMixedPunctuation(string text)
    {
        var chinesePunct = "\uff0c\u3002\uff01\uff1f\u3001\uff1b\uff1a\u201c\u201d\u2018\u2019\uff08\uff09\u3010\u3011";
        bool hasChinese = text.Any(c => chinesePunct.Contains(c));
        bool hasEnglish = text.Any(c => ",.!?;:\"'()[]".Contains(c));
        bool hasChineseChars = text.Any(c => c >= 0x4E00 && c <= 0x9FFF);
        return hasChinese && hasEnglish && hasChineseChars;
    }

    private static string GetBookmarkText(BookmarkStart bkStart)
    {
        var bkId = bkStart.Id?.Value;
        if (bkId == null) return "";

        var sb = new System.Text.StringBuilder();
        var sibling = bkStart.NextSibling();
        while (sibling != null)
        {
            if (sibling is BookmarkEnd bkEnd && bkEnd.Id?.Value == bkId)
                break;
            if (sibling is Run run)
                sb.Append(string.Concat(run.Descendants<Text>().Select(t => t.Text)));
            sibling = sibling.NextSibling();
        }
        return sb.ToString();
    }
}
