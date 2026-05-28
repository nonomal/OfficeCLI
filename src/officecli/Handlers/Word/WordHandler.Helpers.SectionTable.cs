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
    /// Ensure Columns exists in SectionProperties in correct schema order.
    /// Schema order: ..., PageMargin, ..., Columns, ...
    /// </summary>
    private static Columns EnsureColumns(SectionProperties sectPr)
    {
        var existing = sectPr.GetFirstChild<Columns>();
        if (existing != null) return existing;

        var cols = new Columns();
        var pm = sectPr.GetFirstChild<PageMargin>();
        if (pm != null)
            pm.InsertAfterSelf(cols);
        else
        {
            var pgSz = sectPr.GetFirstChild<PageSize>();
            if (pgSz != null)
                pgSz.InsertAfterSelf(cols);
            else
            {
                // Insert after SectionType, or after last headerReference/footerReference
                var sectionType = sectPr.GetFirstChild<SectionType>();
                if (sectionType != null)
                    sectionType.InsertAfterSelf(cols);
                else
                {
                    OpenXmlElement? lastRef = null;
                    foreach (var child in sectPr.ChildElements)
                    {
                        if (child is HeaderReference || child is FooterReference)
                            lastRef = child;
                    }
                    if (lastRef != null)
                        lastRef.InsertAfterSelf(cols);
                    else
                        sectPr.PrependChild(cols);
                }
            }
        }
        return cols;
    }

    /// <summary>
    /// Ensure PageSize exists in SectionProperties in correct schema order.
    /// Schema order: SectionType, PageSize, PageMargin, ...
    /// </summary>
    private static PageSize EnsureSectPrPageSize(SectionProperties sectPr)
    {
        var existing = sectPr.GetFirstChild<PageSize>();
        if (existing != null) return existing;

        var ps = new PageSize();
        // Insert after SectionType if present, then after FooterReference/HeaderReference,
        // otherwise prepend. OOXML schema order: headerReference*, footerReference*, ..., sectType, pgSz, pgMar
        var sectionType = sectPr.GetFirstChild<SectionType>();
        if (sectionType != null)
        {
            sectionType.InsertAfterSelf(ps);
        }
        else
        {
            // Find the last HeaderReference or FooterReference to insert after
            OpenXmlElement? lastRef = null;
            foreach (var child in sectPr.ChildElements)
            {
                if (child is HeaderReference || child is FooterReference)
                    lastRef = child;
            }
            if (lastRef != null)
                lastRef.InsertAfterSelf(ps);
            else
                sectPr.PrependChild(ps);
        }
        return ps;
    }

    /// <summary>
    /// Ensure PageMargin exists in SectionProperties in correct schema order.
    /// Schema order: SectionType, PageSize, PageMargin, ...
    /// </summary>
    private static PageMargin EnsureSectPrPageMargin(SectionProperties sectPr)
    {
        var existing = sectPr.GetFirstChild<PageMargin>();
        if (existing != null) return existing;

        var pm = new PageMargin();
        // Insert after PageSize if present, after SectionType, after last headerRef/footerRef, or prepend
        var pageSize = sectPr.GetFirstChild<PageSize>();
        if (pageSize != null)
            pageSize.InsertAfterSelf(pm);
        else
        {
            var sectionType = sectPr.GetFirstChild<SectionType>();
            if (sectionType != null)
                sectionType.InsertAfterSelf(pm);
            else
            {
                OpenXmlElement? lastRef = null;
                foreach (var child in sectPr.ChildElements)
                {
                    if (child is HeaderReference || child is FooterReference)
                        lastRef = child;
                }
                if (lastRef != null)
                    lastRef.InsertAfterSelf(pm);
                else
                    sectPr.PrependChild(pm);
            }
        }
        return pm;
    }

    // ==================== sectPr schema-order insertion ====================

    /// <summary>
    /// Canonical CT_SectPr child schema order (subset, in document order):
    ///   headerReference*, footerReference*, footnotePr, endnotePr, type, pgSz,
    ///   pgMar, paperSrc, pgBorders, lnNumType, pgNumType, cols, formProt,
    ///   vAlign, noEndnote, titlePg, textDirection, bidi, rtlGutter, docGrid,
    ///   printerSettings, sectPrChange.
    /// Used to map a child element to its schema-order rank for ordered insertion.
    /// </summary>
    private static int SectPrChildOrder(OpenXmlElement el) => el switch
    {
        HeaderReference => 0,
        FooterReference => 1,
        FootnoteProperties => 2,
        EndnoteProperties => 3,
        SectionType => 4,
        PageSize => 5,
        PageMargin => 6,
        PaperSource => 7,
        PageBorders => 8,
        LineNumberType => 9,
        PageNumberType => 10,
        Columns => 11,
        FormProtection => 12,
        VerticalTextAlignmentOnPage => 13,
        NoEndnote => 14,
        TitlePage => 15,
        TextDirection => 16,
        BiDi => 17,
        GutterOnRight => 18,
        DocGrid => 19,
        PrinterSettingsReference => 20,
        SectionPropertiesChange => 21,
        _ => 99,
    };

    /// <summary>
    /// Insert <paramref name="newChild"/> into <paramref name="sectPr"/> at the
    /// position dictated by CT_SectPr schema order. Required for elements like
    /// &lt;w:bidi/&gt; which Word's schema validator rejects when appended after
    /// &lt;w:docGrid/&gt;. Mirrors the InsertRunPropInSchemaOrder pattern used
    /// for run properties.
    /// </summary>
    private static void InsertSectPrChildInOrder(SectionProperties sectPr, OpenXmlElement newChild)
    {
        var newRank = SectPrChildOrder(newChild);
        OpenXmlElement? successor = null;
        foreach (var child in sectPr.ChildElements)
        {
            if (SectPrChildOrder(child) > newRank)
            {
                successor = child;
                break;
            }
        }
        if (successor != null)
            successor.InsertBeforeSelf(newChild);
        else
            sectPr.AppendChild(newChild);
    }

    /// <summary>
    /// CT_TblPrBase schema order:
    ///   tblStyle, tblpPr, tblOverlap, bidiVisual, tblStyleRowBandSize,
    ///   tblStyleColBandSize, tblW, jc, tblCellSpacing, tblInd, tblBorders,
    ///   shd, tblLayout, tblCellMar, tblLook, tblCaption, tblDescription,
    ///   tblPrChange.
    /// </summary>
    private static int TblPrChildOrder(OpenXmlElement el) => el switch
    {
        TableStyle => 0,
        TablePositionProperties => 1,
        TableOverlap => 2,
        BiDiVisual => 3,
        TableStyleRowBandSize => 4,
        TableStyleColumnBandSize => 5,
        TableWidth => 6,
        TableJustification => 7,
        TableCellSpacing => 8,
        TableIndentation => 9,
        TableBorders => 10,
        Shading => 11,
        TableLayout => 12,
        TableCellMarginDefault => 13,
        TableLook => 14,
        TableCaption => 15,
        TableDescription => 16,
        TablePropertiesChange => 17,
        _ => 99,
    };

    /// <summary>
    /// Insert <paramref name="newChild"/> into <paramref name="tblPr"/> at the
    /// position dictated by CT_TblPrBase schema order. Required for elements
    /// like &lt;w:bidiVisual/&gt; which Word's schema validator rejects when
    /// appended after &lt;w:tblBorders/&gt;.
    /// </summary>
    private static void InsertTblPrChildInOrder(TableProperties tblPr, OpenXmlElement newChild)
    {
        var newRank = TblPrChildOrder(newChild);
        OpenXmlElement? successor = null;
        foreach (var child in tblPr.ChildElements)
        {
            if (TblPrChildOrder(child) > newRank)
            {
                successor = child;
                break;
            }
        }
        if (successor != null)
            successor.InsertBeforeSelf(newChild);
        else
            tblPr.AppendChild(newChild);
    }

    /// <summary>
    /// Get-or-create <w:tblCellMar/> on the given tblPr in CT_TblPrBase schema
    /// order. Prevents the "argv-order produces schema-invalid tblCellMar
    /// position" class of bug — see InsertTblPrChildInOrder docstring.
    /// </summary>
    private static TableCellMarginDefault EnsureTableCellMarginDefault(TableProperties tblPr)
    {
        var cm = tblPr.TableCellMarginDefault;
        if (cm == null)
        {
            cm = new TableCellMarginDefault();
            InsertTblPrChildInOrder(tblPr, cm);
        }
        return cm;
    }
}
