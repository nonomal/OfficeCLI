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

    // Tolerant BCP-47 shape used to validate run lang.{val,ea,cs} values.
    // RFC 5646 §2.1: language tag is primary (2-3 ALPHA, or 4-8 ALPHA "reserved"
    // for future / "registered"), followed by hyphen-separated subtags each
    // CONSISTENCY(bcp47-validation): shape regex moved to Core/Bcp47LanguageTag.cs
    // so docx and pptx use the same validator. Wrappers kept for readability
    // of the call sites below.
    private const int LangBcp47MaxLength = OfficeCli.Core.Bcp47LanguageTag.MaxLength;
    private static bool LangBcp47IsValid(string value) => OfficeCli.Core.Bcp47LanguageTag.IsValid(value);

    /// <summary>
    /// Resolve the OpenXmlPart that owns a given element. Returns the
    /// HeaderPart/FooterPart/FootnotesPart/EndnotesPart/CommentsPart when the
    /// element lives inside one of those parts, falling back to MainDocumentPart.
    /// Used for part-local relationships like hyperlinks that must be added to
    /// the host part's rels file (e.g. word/_rels/header1.xml.rels) rather than
    /// the document rels.
    /// </summary>
    private OpenXmlPart ResolveHostPart(OpenXmlElement element)
    {
        var main = _doc.MainDocumentPart!;
        // Walk to the part-root element (Header/Footer/Footnotes/Endnotes/Comments/Document)
        var hdr = element.Ancestors<Header>().FirstOrDefault();
        if (hdr != null)
        {
            var hp = main.HeaderParts.FirstOrDefault(p => ReferenceEquals(p.Header, hdr));
            if (hp != null) return hp;
        }
        var ftr = element.Ancestors<Footer>().FirstOrDefault();
        if (ftr != null)
        {
            var fp = main.FooterParts.FirstOrDefault(p => ReferenceEquals(p.Footer, ftr));
            if (fp != null) return fp;
        }
        // Footnote/Endnote: parts live on MainDocumentPart.FootnotesPart / EndnotesPart
        if (element.Ancestors<Footnote>().Any() && main.FootnotesPart != null)
            return main.FootnotesPart;
        if (element.Ancestors<Endnote>().Any() && main.EndnotesPart != null)
            return main.EndnotesPart;
        return main;
    }

    /// <summary>
    /// Resolve a hyperlink relationship by id, searching the element's host
    /// part first, then falling back to MainDocumentPart and other host parts.
    /// </summary>
    private HyperlinkRelationship? ResolveHyperlinkRelationship(OpenXmlElement element, string relId)
    {
        var host = ResolveHostPart(element);
        var rel = host.HyperlinkRelationships.FirstOrDefault(r => r.Id == relId);
        if (rel != null) return rel;
        // Fallback: scan MainDocumentPart and all header/footer parts (handles
        // documents authored with rels in unexpected places).
        var main = _doc.MainDocumentPart!;
        if (!ReferenceEquals(host, main))
        {
            rel = main.HyperlinkRelationships.FirstOrDefault(r => r.Id == relId);
            if (rel != null) return rel;
        }
        foreach (var hp in main.HeaderParts)
        {
            if (ReferenceEquals(hp, host)) continue;
            rel = hp.HyperlinkRelationships.FirstOrDefault(r => r.Id == relId);
            if (rel != null) return rel;
        }
        foreach (var fp in main.FooterParts)
        {
            if (ReferenceEquals(fp, host)) continue;
            rel = fp.HyperlinkRelationships.FirstOrDefault(r => r.Id == relId);
            if (rel != null) return rel;
        }
        return null;
    }

    // ==================== Private Helpers ====================

    /// <summary>
    /// Format twips as a human-readable cm string (e.g., "21cm").
    /// 1 inch = 1440 twips, 1 inch = 2.54 cm.
    /// </summary>
    private static string FormatTwipsToCm(uint twips)
    {
        var cm = twips * 2.54 / 1440.0;
        return $"{cm:0.##}cm";
    }

    private static bool IsTruthy(string? value) =>
        ParseHelpers.IsTruthy(value);

    /// <summary>
    /// BUG-R7-07: a value the user explicitly typed as "false"/"0"/"off" — not
    /// just any non-truthy input (null/empty count as "no override"). Used by
    /// AddParagraph's no-text fallback to decide whether to emit
    /// <c>&lt;w:b w:val="false"/&gt;</c> as an explicit style override vs.
    /// simply removing the element. Set-style call sites continue to use
    /// ApplyRunFormatting's "remove on falsy" semantics so existing tests
    /// (R25/R26 EmptyRpr_NotSurfaced) keep passing.
    /// </summary>
    internal static bool IsExplicitFalseAddOverride(string? value)
    {
        if (string.IsNullOrEmpty(value)) return false;
        var v = value.Trim().ToLowerInvariant();
        return v is "false" or "0" or "off" or "no" or "n";
    }

    /// <summary>
    /// Read a w:val OnOff attribute defensively. Returns null when the
    /// attribute is absent OR when the stored text is not a valid OnOff
    /// token (e.g. <c>&lt;w:bidi w:val="garbage"/&gt;</c>). Default-on
    /// elements (BiDi, Bold, etc.) are conventionally treated as true
    /// when Val is null. R8-fuzz-5: prevents OnOffValue.Parse from
    /// crashing Get/HtmlPreview on a document that loaded fine but
    /// disk-stored a malformed attribute.
    /// </summary>
    internal static bool? TryReadOnOff(DocumentFormat.OpenXml.OnOffValue? val)
    {
        if (val == null) return true; // default-on: <w:bidi/> with no Val
        try { return val.Value; }
        catch (FormatException) { return null; }
    }

    /// <summary>
    /// Find the 1-based run index inside the anchor paragraph where the
    /// CommentRangeStart with <paramref name="commentId"/> sits — i.e. the
    /// number of runs before the range marker plus 1. Returns 0 when the
    /// range marker is not found, or sits before any Run (anchor at paragraph
    /// start).
    /// BUG-DUMP4-03: callers (WordBatchEmitter) need this so dump can preserve
    /// intra-paragraph anchor position; without it replay widens every
    /// comment to the whole paragraph.
    /// </summary>
    public int FindCommentAnchorRunIndex(string commentId)
    {
        var body = _doc.MainDocumentPart?.Document?.Body;
        if (body == null) return 0;
        foreach (var para in body.Descendants<Paragraph>())
        {
            var rs = para.Descendants<CommentRangeStart>()
                .FirstOrDefault(r => r.Id?.Value == commentId);
            if (rs == null) continue;
            // Count Run elements that appear before the CommentRangeStart in
            // document order within the same paragraph.
            int runCount = 0;
            foreach (var el in para.Descendants())
            {
                if (ReferenceEquals(el, rs)) break;
                if (el is Run r && r.GetFirstChild<CommentReference>() == null) runCount++;
            }
            return runCount; // 0 = before any run; N = after run N (1-based)
        }
        return 0;
    }

    /// <summary>
    /// Find the paragraph path where a CommentRangeStart with the given ID is anchored.
    /// Returns "/body/p[N]" or null if not found.
    /// </summary>
    private string? FindCommentAnchorPath(string commentId)
    {
        var body = _doc.MainDocumentPart?.Document?.Body;
        if (body == null) return null;

        var paragraphs = body.Elements<Paragraph>().ToList();
        for (int i = 0; i < paragraphs.Count; i++)
        {
            var hasRange = paragraphs[i].Descendants<CommentRangeStart>()
                .Any(rs => rs.Id?.Value == commentId);
            if (hasRange) return $"/body/{BuildParaPathSegment(paragraphs[i], i + 1)}";
        }
        return null;
    }

    /// <summary>
    /// Get-or-create a sectPr child in CT_SectPr schema order. Replaces the
    /// `?? sectPr.AppendChild(new T())` idiom which violated schema order when
    /// other higher-ranked elements were already present.
    /// </summary>
    private static T EnsureSectPrChild<T>(SectionProperties sectPr) where T : OpenXmlElement, new()
    {
        var existing = sectPr.GetFirstChild<T>();
        if (existing != null) return existing;
        var created = new T();
        InsertSectPrChildInOrder(sectPr, created);
        return created;
    }

    /// <summary>
    /// Get current document protection mode and enforcement status.
    /// </summary>
    private (string mode, bool enforced) GetDocumentProtection()
    {
        var settings = _doc.MainDocumentPart?.DocumentSettingsPart?.Settings;
        var docProtection = settings?.GetFirstChild<DocumentProtection>();
        if (docProtection == null)
            return ("none", false);

        var mode = docProtection.Edit?.InnerText switch
        {
            "readOnly" => "readOnly",
            "comments" => "comments",
            "trackedChanges" => "trackedChanges",
            "forms" => "forms",
            _ => "none"
        };
        var enforced = docProtection.Enforcement?.Value == true
            || (docProtection.Enforcement?.Value == null && docProtection.Edit != null);
        return (mode, enforced);
    }

    /// <summary>
    /// Check if an SDT element is editable based on its lock attribute and the current document protection.
    /// </summary>
    private bool IsSdtEditable(SdtProperties? sdtProps)
    {
        var (mode, enforced) = GetDocumentProtection();

        // No protection or not enforced → all SDTs are editable
        if (!enforced || mode == "none")
            return true;

        // readOnly protection → SDTs are not editable (unless in permRange, P2)
        if (mode == "readOnly")
            return false;

        // forms protection → SDTs are editable unless content-locked
        if (mode == "forms")
        {
            var lockEl = sdtProps?.GetFirstChild<DocumentFormat.OpenXml.Wordprocessing.Lock>();
            var lockVal = lockEl?.Val?.Value;
            return lockVal != LockingValues.ContentLocked && lockVal != LockingValues.SdtContentLocked;
        }

        // comments/trackedChanges → not typically editable
        return false;
    }
}
