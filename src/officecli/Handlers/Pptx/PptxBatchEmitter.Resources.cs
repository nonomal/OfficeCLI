// Copyright 2025 OfficeCLI (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using OfficeCli.Core;

namespace OfficeCli.Handlers;

public static partial class PptxBatchEmitter
{
    // CONSISTENCY(emit-resources-mirror): mirrors WordBatchEmitter.Resources.cs
    // — each whole-part-XML block emits as a single raw-set replace. Theme /
    // master / layout / notesMaster carry rich structured XML (clrScheme,
    // fontScheme, txStyles, fmtScheme, …) that has no typed Set vocabulary; the
    // natural operation is "swap the whole block". Replay's raw-set overwrites
    // whatever the blank deck stamped during BlankDocCreator.

    // CONSISTENCY(raw-xmlns-canonicalize): mirrors
    // WordBatchEmitter.Resources.CanonicalizeRawXml. RawXmlHelper.Execute
    // propagates the root's xmlns declarations onto every direct child so the
    // SDK's InnerXml setter can resolve prefixes (SDK does not inherit root
    // xmlns scope when parsing inner content). After replay, the part's XML
    // carries redundant xmlns:p / xmlns:a attrs on each child of /theme,
    // /slideMaster[N], /slideLayout[N] — observed first-replay growth on a
    // blank-deck round-trip: 16657 → 17923 bytes (≈1.2 KB across 7 raw-set
    // parts), then stable on subsequent rounds. Canonicalise on emit so the
    // first-pass (clean source) and second-pass (post-replay bloated) shapes
    // collapse identically.
    private static string CanonicalizeRawXml(string xml)
    {
        if (string.IsNullOrEmpty(xml) || !xml.StartsWith("<")) return xml;
        try
        {
            var doc = System.Xml.Linq.XDocument.Parse(xml);
            if (doc.Root == null) return xml;
            var rootNsAttrs = doc.Root.Attributes()
                .Where(a => a.IsNamespaceDeclaration)
                .ToDictionary(a => a.Name, a => a.Value);
            foreach (var desc in doc.Root.Descendants())
            {
                var toRemove = desc.Attributes()
                    .Where(a => a.IsNamespaceDeclaration
                                && rootNsAttrs.TryGetValue(a.Name, out var v)
                                && v == a.Value)
                    .ToList();
                foreach (var a in toRemove) a.Remove();
            }
            return doc.Root.ToString(System.Xml.Linq.SaveOptions.DisableFormatting);
        }
        catch
        {
            // Malformed XML — leave as-is rather than corrupting.
            return xml;
        }
    }

    private static void EmitThemeRaw(PowerPointHandler ppt, List<BatchItem> items)
    {
        // Pptx Raw("/theme") returns the presentation-level theme part (first
        // master's theme). Multi-master decks have additional theme parts
        // attached to each master, but the existing Raw/RawSet surface only
        // addresses the primary one — keep parity until per-master theme
        // raw-set lands. Skip silently when the source has none.
        string xml;
        try { xml = ppt.Raw("/theme"); }
        catch { return; }
        if (string.IsNullOrEmpty(xml) || !xml.StartsWith("<") || xml == "(no theme)")
            return;
        xml = CanonicalizeRawXml(xml);

        items.Add(new BatchItem
        {
            Command = "raw-set",
            Part = "/theme",
            Xpath = "/a:theme",
            Action = "replace",
            Xml = xml
        });
    }

    private static void EmitNotesMasterRaw(PowerPointHandler ppt, List<BatchItem> items)
    {
        if (!ppt.HasNotesMaster) return;
        string xml;
        try { xml = ppt.Raw("/notesMaster"); }
        catch { return; }
        if (string.IsNullOrEmpty(xml) || !xml.StartsWith("<")) return;
        xml = CanonicalizeRawXml(xml);

        items.Add(new BatchItem
        {
            Command = "raw-set",
            Part = "/notesMaster",
            Xpath = "/p:notesMaster",
            Action = "replace",
            Xml = xml
        });
    }

    private static void EmitMasterRaw(PowerPointHandler ppt, List<BatchItem> items)
    {
        var n = ppt.SlideMasterCount;
        for (int i = 1; i <= n; i++) EmitMasterRawOne(ppt, i, items);
    }

    private static bool EmitMasterRawOne(PowerPointHandler ppt, int idx, List<BatchItem> items)
    {
        string xml;
        try { xml = ppt.Raw($"/slideMaster[{idx}]"); }
        catch { return false; }
        if (string.IsNullOrEmpty(xml) || !xml.StartsWith("<")) return false;
        xml = CanonicalizeRawXml(xml);

        items.Add(new BatchItem
        {
            Command = "raw-set",
            Part = $"/slideMaster[{idx}]",
            Xpath = "/p:sldMaster",
            Action = "replace",
            Xml = xml
        });
        return true;
    }

    private static void EmitLayoutRaw(PowerPointHandler ppt, List<BatchItem> items)
    {
        var n = ppt.SlideLayoutCount;
        for (int i = 1; i <= n; i++) EmitLayoutRawOne(ppt, i, items);
    }

    private static bool EmitLayoutRawOne(PowerPointHandler ppt, int idx, List<BatchItem> items)
    {
        string xml;
        try { xml = ppt.Raw($"/slideLayout[{idx}]"); }
        catch { return false; }
        if (string.IsNullOrEmpty(xml) || !xml.StartsWith("<")) return false;
        xml = CanonicalizeRawXml(xml);

        items.Add(new BatchItem
        {
            Command = "raw-set",
            Part = $"/slideLayout[{idx}]",
            Xpath = "/p:sldLayout",
            Action = "replace",
            Xml = xml
        });
        return true;
    }

    private static bool EmitNoteSlideRawOne(PowerPointHandler ppt, int idx, List<BatchItem> items)
    {
        string xml;
        try { xml = ppt.Raw($"/noteSlide[{idx}]"); }
        catch { return false; }
        if (string.IsNullOrEmpty(xml) || !xml.StartsWith("<")) return false;
        xml = CanonicalizeRawXml(xml);

        items.Add(new BatchItem
        {
            Command = "raw-set",
            Part = $"/noteSlide[{idx}]",
            Xpath = "/p:notes",
            Action = "replace",
            Xml = xml
        });
        return true;
    }

    // Presentation-level structural children that the typed Add/Set/EmitPresentationProps
    // surface does not round-trip: custShowLst (custom slide shows) and extLst
    // (extension children — sectionLst / modifyVerifier / etc.). Both reference
    // slides by rId; `add slide` on replay mints fresh rIds, so a verbatim
    // raw-set replace would point at stale targets and PowerPoint would refuse
    // to open. Honest path: emit the source XML as a best-effort append AND
    // record an UnsupportedWarning so callers know the references may need
    // manual rewiring. Mirrors the "loud not silent" rule for content we cannot
    // faithfully serialize through the typed vocabulary.
    private static void EmitPresentationExtras(
        PowerPointHandler ppt, List<BatchItem> items, SlideEmitContext ctx)
    {
        string presXml;
        try { presXml = ppt.Raw("/presentation"); }
        catch { return; }
        if (string.IsNullOrEmpty(presXml) || !presXml.StartsWith("<")) return;

        System.Xml.Linq.XDocument doc;
        try { doc = System.Xml.Linq.XDocument.Parse(presXml); }
        catch { return; }
        if (doc.Root == null) return;

        var pNs = System.Xml.Linq.XNamespace.Get(
            "http://schemas.openxmlformats.org/presentationml/2006/main");

        // CT_Presentation child order (ECMA-376 §19.2.1.26) is significant —
        // PowerPoint's strict validator (and replay's OOXML validator) flags
        // any element that appears after a later-schema sibling as an
        // "unexpected child". The relevant tail is:
        //   …, custShowLst, photoAlbum, custDataLst, kinsoku,
        //   defaultTextStyle, modifyVerifier, extLst.
        // Emit in schema order so each `raw-set append` lands on the
        // trailing-most slot at that moment. Previously extLst was appended
        // before kinsoku / defaultTextStyle / photoAlbum, which then chained
        // after extLst in the wrong order — PowerPoint refused the file
        // (0x80070570) on every deck that carried both a section list and
        // a deck-level default text style (gov_bja_template, …).

        // custShowLst — `<p:custShowLst><p:custShow><p:sldLst><p:sld r:id="…"/>`.
        var custShow = doc.Root.Element(pNs + "custShowLst");
        if (custShow != null)
        {
            var xml = CanonicalizeRawXml(custShow.ToString(System.Xml.Linq.SaveOptions.DisableFormatting));
            items.Add(new BatchItem
            {
                Command = "raw-set",
                Part = "/presentation",
                Xpath = "/p:presentation",
                Action = "append",
                Xml = xml,
            });
            ctx.Unsupported.Add(new UnsupportedWarning(
                Element: "presentation.custShowLst",
                SlidePath: "/presentation",
                Reason: "Custom slide shows reference slides by relationship id; replay's `add slide` mints fresh rIds, so the custShow targets may point at stale relationships. Verify in PowerPoint before relying on the round-tripped show."));
        }

        // photoAlbum — flags marking the deck as a photo album
        // (`<p:photoAlbum bw="…" showCaptions="…" layout="…" frame="…"/>`).
        var photo = doc.Root.Element(pNs + "photoAlbum");
        if (photo != null)
        {
            var xml = CanonicalizeRawXml(photo.ToString(System.Xml.Linq.SaveOptions.DisableFormatting));
            items.Add(new BatchItem
            {
                Command = "raw-set",
                Part = "/presentation",
                Xpath = "/p:presentation",
                Action = "append",
                Xml = xml,
            });
            ctx.Unsupported.Add(new UnsupportedWarning(
                Element: "presentation.photoAlbum",
                SlidePath: "/presentation",
                Reason: "photoAlbum (PowerPoint Photo Album metadata: bw / captions / layout / frame attributes) is preserved verbatim via raw-set; no typed Set vocabulary exists for these attributes."));
        }

        // kinsoku — East-Asian line-break rules (`<p:kinsoku invalChars=… hangChars=…/>`).
        var kins = doc.Root.Element(pNs + "kinsoku");
        if (kins != null)
        {
            var xml = CanonicalizeRawXml(kins.ToString(System.Xml.Linq.SaveOptions.DisableFormatting));
            items.Add(new BatchItem
            {
                Command = "raw-set",
                Part = "/presentation",
                Xpath = "/p:presentation",
                Action = "append",
                Xml = xml,
            });
            ctx.Unsupported.Add(new UnsupportedWarning(
                Element: "presentation.kinsoku",
                SlidePath: "/presentation",
                Reason: "kinsoku (East-Asian line-break rules: invalid / hanging chars) is preserved verbatim via raw-set; no typed Set vocabulary exists yet to edit individual rule entries."));
        }

        // defaultTextStyle — body-text level defaults inherited by every
        // slide layout / master that doesn't override them (`<p:defaultTextStyle>
        // <a:defPPr/> <a:lvl1pPr/> …</p:defaultTextStyle>`).
        var dts = doc.Root.Element(pNs + "defaultTextStyle");
        if (dts != null)
        {
            var xml = CanonicalizeRawXml(dts.ToString(System.Xml.Linq.SaveOptions.DisableFormatting));
            items.Add(new BatchItem
            {
                Command = "raw-set",
                Part = "/presentation",
                Xpath = "/p:presentation",
                Action = "append",
                Xml = xml,
            });
            ctx.Unsupported.Add(new UnsupportedWarning(
                Element: "presentation.defaultTextStyle",
                SlidePath: "/presentation",
                Reason: "defaultTextStyle (deck-level paragraph defaults inherited by layouts/masters) is preserved verbatim via raw-set; no typed Set surface for individual level paragraph properties at this level yet."));
        }

        // extLst — MUST be last (CT_Presentation tail). `<p:extLst><p:ext uri="…">`
        // carries sectionLst, modifyVerifier, misc 2010+ extensions.
        var ext = doc.Root.Element(pNs + "extLst");
        if (ext != null)
        {
            var xml = CanonicalizeRawXml(ext.ToString(System.Xml.Linq.SaveOptions.DisableFormatting));
            items.Add(new BatchItem
            {
                Command = "raw-set",
                Part = "/presentation",
                Xpath = "/p:presentation",
                Action = "append",
                Xml = xml,
            });
            ctx.Unsupported.Add(new UnsupportedWarning(
                Element: "presentation.extLst",
                SlidePath: "/presentation",
                Reason: "Presentation extensions (sectionLst / modifyVerifier / …) may reference slides by rId; replay mints fresh rIds, so references can go stale. Section names survive; section → slide membership may need manual rewiring."));
        }
    }
}
