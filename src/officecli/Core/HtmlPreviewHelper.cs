// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

using System.Text;
using DocumentFormat.OpenXml.Packaging;

namespace OfficeCli.Core;

/// <summary>
/// Shared helpers for HTML preview rendering across PowerPoint, Word, and Excel handlers.
/// </summary>
internal static class HtmlPreviewHelper
{
    /// <summary>
    /// HTML-encode text for safe insertion into element content or double-quoted
    /// attribute values: escapes &amp;, &lt;, &gt;, double-quote, and single-quote.
    /// This is the plain entity-encoding shared by the PowerPoint, Excel, and chart
    /// SVG renderers. (Word's preview uses a variant that additionally preserves
    /// consecutive spaces as non-breaking spaces and does not escape the apostrophe —
    /// see WordHandler.HtmlPreview.Css.HtmlEncode, kept separate by design.)
    /// </summary>
    public static string HtmlEncode(string text)
    {
        return text
            .Replace("&", "&amp;")
            .Replace("<", "&lt;")
            .Replace(">", "&gt;")
            .Replace("\"", "&quot;")
            .Replace("'", "&#39;");
    }

    /// <summary>
    /// Load an OpenXML part by its relationship ID and return the content as a base64 data URI.
    /// Returns null if the part cannot be found or read.
    /// </summary>
    public static string? PartToDataUri(OpenXmlPart parentPart, string relId)
    {
        try
        {
            var part = parentPart.GetPartById(relId);
            using var stream = part.GetStream();
            using var ms = new MemoryStream();
            stream.CopyTo(ms);
            var contentType = part.ContentType ?? "image/png";
            if (IsVectorMetafile(contentType))
            {
                // WMF/EMF metafiles cannot be decoded by browsers (no <img> support)
                // and there is no cross-platform .NET rasterizer. Degrade gracefully to
                // a self-contained SVG placeholder so the preview shows a clean framed box
                // instead of a broken-image icon.
                return MetafilePlaceholderDataUri(contentType);
            }
            return $"data:{contentType};base64,{Convert.ToBase64String(ms.ToArray())}";
        }
        catch
        {
            return null;
        }
    }

    /// <summary>
    /// True for WMF/EMF metafile content types that browsers cannot render in an
    /// &lt;img&gt; tag (image/wmf, image/x-wmf, image/emf, image/x-emf).
    /// </summary>
    private static bool IsVectorMetafile(string contentType)
    {
        return contentType.Equals("image/wmf", StringComparison.OrdinalIgnoreCase)
            || contentType.Equals("image/x-wmf", StringComparison.OrdinalIgnoreCase)
            || contentType.Equals("image/emf", StringComparison.OrdinalIgnoreCase)
            || contentType.Equals("image/x-emf", StringComparison.OrdinalIgnoreCase);
    }

    /// <summary>
    /// Build a base64-encoded SVG data URI placeholder for an undecodable metafile.
    /// The SVG uses a viewBox + preserveAspectRatio so it scales to fill the host
    /// &lt;img&gt; width/height, drawing a light-gray bordered rectangle with a centered
    /// label (WMF or EMF). Base64 encoding avoids any data-URI escaping concerns.
    /// </summary>
    private static string MetafilePlaceholderDataUri(string contentType)
    {
        var label = contentType.IndexOf("emf", StringComparison.OrdinalIgnoreCase) >= 0 ? "EMF" : "WMF";
        var svg =
            "<svg xmlns=\"http://www.w3.org/2000/svg\" viewBox=\"0 0 120 120\" " +
            "preserveAspectRatio=\"xMidYMid meet\">" +
            "<rect x=\"2\" y=\"2\" width=\"116\" height=\"116\" rx=\"4\" " +
            "fill=\"#f5f5f5\" stroke=\"#cccccc\" stroke-width=\"2\"/>" +
            "<text x=\"60\" y=\"66\" font-family=\"sans-serif\" font-size=\"22\" " +
            "fill=\"#999999\" text-anchor=\"middle\">" + label + "</text>" +
            "</svg>";
        var b64 = Convert.ToBase64String(Encoding.UTF8.GetBytes(svg));
        return $"data:image/svg+xml;base64,{b64}";
    }
}
