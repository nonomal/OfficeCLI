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

    // ==================== Extended Chart Helpers ====================

    private const string WordChartExUri = "http://schemas.microsoft.com/office/drawing/2014/chartex";
    private const string WordChartUri = "http://schemas.openxmlformats.org/drawingml/2006/chart";

    /// <summary>
    /// Count all charts (both standard ChartPart and ExtendedChartPart) in the document.
    /// </summary>
    private static int CountWordCharts(MainDocumentPart mainPart)
    {
        return mainPart.ChartParts.Count() + mainPart.ExtendedChartParts.Count();
    }

    /// <summary>
    /// Represents a chart part in Word that could be either a standard ChartPart or an ExtendedChartPart.
    /// </summary>
    private class WordChartInfo
    {
        public ChartPart? StandardPart { get; set; }
        public ExtendedChartPart? ExtendedPart { get; set; }
        public DW.DocProperties? DocProperties { get; set; }
        /// <summary>
        /// The <c>wp:inline</c> element that hosts this chart — needed by
        /// chart position Set to mutate the <c>wp:extent</c> child.
        /// </summary>
        public DW.Inline? Inline { get; set; }
        public bool IsExtended => ExtendedPart != null;
    }

    /// <summary>
    /// Get all chart parts (standard + extended) in document order by walking Drawing/Inline elements.
    /// </summary>
    private List<WordChartInfo> GetAllWordCharts()
    {
        var result = new List<WordChartInfo>();
        var mainPart = _doc.MainDocumentPart;
        if (mainPart?.Document?.Body == null) return result;

        // Charts can be inserted in main document body, header parts, or footer parts.
        // Each part owns its own ImagePart/ChartPart relationships (round23 S host-part
        // routing), so look up the chart rel against the part the inline belongs to —
        // not always mainPart. Without this, header/footer charts are dropped from
        // GetAllWordCharts and AddChart's path emission falls back to /chart[0].
        var hostScans = new List<(OpenXmlPart Part, OpenXmlElement? Root)>
        {
            (mainPart, mainPart.Document.Body)
        };
        foreach (var hp in mainPart.HeaderParts)
            hostScans.Add((hp, hp.Header));
        foreach (var fp in mainPart.FooterParts)
            hostScans.Add((fp, fp.Footer));

        foreach (var (hostPart, root) in hostScans)
        {
            if (root == null) continue;
            foreach (var inline in root.Descendants<DW.Inline>())
            {
                var graphicData = inline.Descendants<A.GraphicData>().FirstOrDefault();
                if (graphicData == null) continue;

                var docProps = inline.Descendants<DW.DocProperties>().FirstOrDefault();

                if (graphicData.Uri == WordChartUri)
                {
                    var chartRef = graphicData.Descendants<DocumentFormat.OpenXml.Drawing.Charts.ChartReference>().FirstOrDefault();
                    if (chartRef?.Id?.Value == null) continue;
                    try
                    {
                        var chartPart = (ChartPart)hostPart.GetPartById(chartRef.Id.Value);
                        result.Add(new WordChartInfo { StandardPart = chartPart, DocProperties = docProps, Inline = inline });
                    }
                    catch { /* skip invalid references */ }
                }
                else if (graphicData.Uri == WordChartExUri)
                {
                    var relId = GetWordExtendedChartRelId(inline);
                    if (relId == null) continue;
                    try
                    {
                        var extPart = (ExtendedChartPart)hostPart.GetPartById(relId);
                        result.Add(new WordChartInfo { ExtendedPart = extPart, DocProperties = docProps, Inline = inline });
                    }
                    catch { /* skip invalid references */ }
                }
            }
        }

        return result;
    }

    /// <summary>
    /// Apply <c>width</c> / <c>height</c> to a Word inline chart's
    /// <c>wp:extent</c>. Accepts unit-qualified sizes (`6cm`, `2in`,
    /// `720pt`) or raw EMU integers via EmuConverter.
    ///
    /// CONSISTENCY(chart-position-set): mirrors the PPTX and Excel path.
    /// Word inline charts have no absolute x/y (they flow with text), so
    /// those keys — if provided — are appended to <paramref name="unsupported"/>
    /// rather than silently dropped.
    /// </summary>
    private static void ApplyWordChartPositionSet(
        DW.Inline inline, Dictionary<string, string> properties, List<string> unsupported)
    {
        var extent = inline.Extent;
        if (extent == null) return;

        // x/y are meaningless for inline charts.
        foreach (var k in new[] { "x", "y" })
        {
            var matched = properties.Keys
                .FirstOrDefault(key => key.Equals(k, StringComparison.OrdinalIgnoreCase));
            if (matched == null) continue;
            unsupported.Add(matched);
            Console.Error.WriteLine(
                $"Warning: '{matched}' is ignored on Word inline charts — inline elements have no absolute position. " +
                "For positioned charts, switch to anchor mode (not currently supported).");
        }

        if (properties.TryGetValue("width", out var wStr))
        {
            try { extent.Cx = OfficeCli.Core.EmuConverter.ParseEmu(wStr); }
            catch { unsupported.Add("width"); }
        }

        if (properties.TryGetValue("height", out var hStr))
        {
            try { extent.Cy = OfficeCli.Core.EmuConverter.ParseEmu(hStr); }
            catch { unsupported.Add("height"); }
        }
    }

    /// <summary>
    /// Get the relationship ID from an extended chart inline Drawing element.
    /// </summary>
    private static string? GetWordExtendedChartRelId(DW.Inline inline)
    {
        var gd = inline.Descendants<A.GraphicData>().FirstOrDefault(g => g.Uri == WordChartExUri);
        if (gd == null) return null;
        var typed = gd.Descendants<DocumentFormat.OpenXml.Office2016.Drawing.ChartDrawing.RelId>().FirstOrDefault();
        if (typed?.Id?.Value != null) return typed.Id.Value;
        foreach (var child in gd.ChildElements)
        {
            var rId = child.GetAttributes().FirstOrDefault(a =>
                a.LocalName == "id" && a.NamespaceUri == "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            if (rId.Value != null) return rId.Value;
        }
        return null;
    }
}
