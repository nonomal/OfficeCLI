// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using DocumentFormat.OpenXml;
using Drawing = DocumentFormat.OpenXml.Drawing;
using CX = DocumentFormat.OpenXml.Office2016.Drawing.ChartDrawing;

namespace OfficeCli.Core;

/// <summary>
/// Builder for cx:chart (Office 2016 extended chart types):
/// funnel, treemap, sunburst, boxWhisker, histogram, waterfall (native).
/// </summary>
internal static class ChartExBuilder
{
    internal static readonly HashSet<string> ExtendedChartTypes = new(StringComparer.OrdinalIgnoreCase)
    {
        "funnel", "treemap", "sunburst", "boxwhisker", "histogram"
    };

    internal static bool IsExtendedChartType(string chartType)
    {
        var normalized = chartType.ToLowerInvariant().Replace(" ", "").Replace("_", "").Replace("-", "");
        return ExtendedChartTypes.Contains(normalized);
    }

    /// <summary>
    /// Build a cx:chartSpace for an extended chart type.
    /// </summary>
    internal static CX.ChartSpace BuildExtendedChartSpace(
        string chartType,
        string? title,
        string[]? categories,
        List<(string name, double[] values)> seriesData,
        Dictionary<string, string> properties)
    {
        var normalized = chartType.ToLowerInvariant().Replace(" ", "").Replace("_", "").Replace("-", "");

        var chartSpace = new CX.ChartSpace();

        // 1. Build ChartData
        var chartData = new CX.ChartData();
        for (int si = 0; si < seriesData.Count; si++)
        {
            var data = BuildDataBlock((uint)si, normalized, categories, seriesData[si].values);
            chartData.AppendChild(data);
        }
        chartSpace.AppendChild(chartData);

        // 2. Build Chart
        var chart = new CX.Chart();

        if (!string.IsNullOrEmpty(title))
        {
            chart.AppendChild(BuildChartTitle(title, properties));
        }

        var plotArea = new CX.PlotArea();
        var plotAreaRegion = new CX.PlotAreaRegion();

        var layoutId = normalized switch
        {
            "funnel" => "funnel",
            "treemap" => "treemap",
            "sunburst" => "sunburst",
            "boxwhisker" => "boxWhisker",
            "histogram" => "clusteredColumn",
            _ => "funnel"
        };

        // Parse series fill colors — reuse the `colors=RED,BLUE,GREEN`
        // convention from regular charts, or accept a single `fill=COLOR`
        // for one-series charts like histogram.
        var seriesColors = ChartHelper.ParseSeriesColors(properties);
        if (seriesColors == null && properties.TryGetValue("fill", out var fillStr))
            seriesColors = new[] { fillStr };

        // dataLabels: off by default. Accept "true" / "on" / "1" / "value"
        // (any explicit truthy value enables). "false" / "off" / "0" disables.
        var showDataLabels = IsTruthyProp(properties, "dataLabels", defaultValue: false);

        for (int si = 0; si < seriesData.Count; si++)
        {
            var series = new CX.Series { LayoutId = new EnumValue<CX.SeriesLayout>(
                ParseSeriesLayout(layoutId)) };

            // Schema order for cx:series:
            //   tx → spPr → valueColors → valueColorPositions → dataPoint*
            //   → dataLabels → dataId → layoutPr → axisId* → extLst
            series.AppendChild(new CX.Text(
                new CX.TextData(
                    new CX.Formula(""),
                    new CX.VXsdstring(seriesData[si].name))));

            // Per-series solid fill
            if (seriesColors != null && si < seriesColors.Length && !string.IsNullOrEmpty(seriesColors[si]))
            {
                var (rgb, _) = ParseHelpers.SanitizeColorForOoxml(seriesColors[si]);
                series.AppendChild(new CX.ShapeProperties(
                    new Drawing.SolidFill(
                        new Drawing.RgbColorModelHex { Val = rgb })));
            }

            // Data labels (value count above each bar)
            if (showDataLabels)
            {
                var dl = new CX.DataLabels { Pos = CX.DataLabelPos.OutEnd };
                dl.AppendChild(new CX.DataLabelVisibilities
                {
                    Value = true,
                    SeriesName = false,
                    CategoryName = false,
                });
                series.AppendChild(dl);
            }

            series.AppendChild(new CX.DataId { Val = (uint)si });

            // Chart-type specific layoutPr (histogram binning, treemap label
            // layout, boxWhisker stats, etc.)
            var layoutPr = BuildLayoutProperties(normalized, properties, seriesData[si].values.Length);
            if (layoutPr != null)
                series.AppendChild(layoutPr);

            plotAreaRegion.AppendChild(series);
        }

        plotArea.AppendChild(plotAreaRegion);

        // Axes for chart types that need them (histogram / boxWhisker).
        // Funnel/treemap/sunburst are axis-less.
        if (normalized is "boxwhisker" or "histogram")
        {
            plotArea.AppendChild(BuildCategoryAxis(id: 0, chartType: normalized, properties));
            plotArea.AppendChild(BuildValueAxis(id: 1, properties));
        }

        chart.AppendChild(plotArea);

        // Legend (optional, appears AFTER plotArea per cx:chart schema order)
        if (properties.TryGetValue("legend", out var legendPos) &&
            !string.IsNullOrEmpty(legendPos) &&
            !legendPos.Equals("none", StringComparison.OrdinalIgnoreCase) &&
            !legendPos.Equals("false", StringComparison.OrdinalIgnoreCase) &&
            !legendPos.Equals("off", StringComparison.OrdinalIgnoreCase))
        {
            chart.AppendChild(BuildLegend(legendPos));
        }

        chartSpace.AppendChild(chart);
        return chartSpace;
    }

    private static CX.ChartTitle BuildChartTitle(string title, Dictionary<string, string>? properties = null)
    {
        var rPr = new Drawing.RunProperties { Language = "en-US" };
        // Delegate style parsing to the shared helper so cChart and cxChart
        // stay in vocabulary lockstep. See
        // ChartHelper.ApplyRunStyleProperties.
        if (properties != null)
            ChartHelper.ApplyRunStyleProperties(rPr, properties, keyPrefix: "title");

        var chartTitle = new CX.ChartTitle();
        chartTitle.AppendChild(new CX.Text(
            new CX.RichTextBody(
                new Drawing.BodyProperties(),
                new Drawing.Paragraph(
                    new Drawing.Run(
                        rPr,
                        new Drawing.Text(title))))));
        return chartTitle;
    }

    private static CX.AxisTitle BuildAxisTitle(string title, Dictionary<string, string>? properties = null)
    {
        var rPr = new Drawing.RunProperties { Language = "en-US" };
        if (properties != null)
            ChartHelper.ApplyRunStyleProperties(rPr, properties, keyPrefix: "axisTitle");

        return new CX.AxisTitle(
            new CX.Text(
                new CX.RichTextBody(
                    new Drawing.BodyProperties(),
                    new Drawing.Paragraph(
                        new Drawing.Run(
                            rPr,
                            new Drawing.Text(title))))));
    }

    /// <summary>
    /// Wrap a shared `a:defRPr` (built from a compound `"size:color:fontname"`
    /// spec by <see cref="ChartHelper.BuildDefaultRunPropertiesFromCompoundSpec"/>)
    /// in a <see cref="CX.TxPrTextBody"/>. Only the outer container differs
    /// from the regular-cChart path (<see cref="C.TextProperties"/>).
    /// </summary>
    private static CX.TxPrTextBody? BuildAxisTickLabelStyle(string compoundSpec)
    {
        if (string.IsNullOrEmpty(compoundSpec)) return null;
        var defRp = ChartHelper.BuildDefaultRunPropertiesFromCompoundSpec(compoundSpec);
        return new CX.TxPrTextBody(
            new Drawing.BodyProperties(),
            new Drawing.ListStyle(),
            new Drawing.Paragraph(new Drawing.ParagraphProperties(defRp)));
    }

    /// <summary>
    /// Build a <see cref="CX.ShapeProperties"/> containing a solid-fill outline
    /// for coloring gridlines. Mirrors the regular-chart `gridline.color` knob.
    /// </summary>
    private static CX.ShapeProperties? BuildGridlineShapeProperties(string color)
    {
        if (string.IsNullOrEmpty(color)) return null;
        var (rgb, _) = ParseHelpers.SanitizeColorForOoxml(color);
        var outline = new Drawing.Outline();
        outline.AppendChild(new Drawing.SolidFill(new Drawing.RgbColorModelHex { Val = rgb }));
        return new CX.ShapeProperties(outline);
    }

    private static CX.Legend BuildLegend(string posSpec)
    {
        var legend = new CX.Legend
        {
            Align = CX.PosAlign.Ctr,
            Overlay = false,
        };
        legend.Pos = posSpec.ToLowerInvariant() switch
        {
            "top" or "t"    => CX.SidePos.T,
            "bottom" or "b" => CX.SidePos.B,
            "left" or "l"   => CX.SidePos.L,
            _               => CX.SidePos.R,  // right is the Excel default
        };
        return legend;
    }

    // Build the category axis (X axis for histogram / boxWhisker). Schema
    // order of Axis children: catScaling → title → majorGridlines →
    // tickLabels → ... (only the ones we emit are listed).
    private static CX.Axis BuildCategoryAxis(uint id, string chartType, Dictionary<string, string> properties)
    {
        var axis = new CX.Axis { Id = id };

        // catScaling is required. histogram defaults gapWidth="0" (bars touch)
        // because that's what real Excel emits and it's what users expect.
        var catScaling = new CX.CategoryAxisScaling();
        var gapWidth = properties.GetValueOrDefault("gapWidth");
        if (string.IsNullOrEmpty(gapWidth) && chartType == "histogram")
            gapWidth = "0";
        if (!string.IsNullOrEmpty(gapWidth))
            catScaling.GapWidth = gapWidth;
        axis.AppendChild(catScaling);

        if (properties.TryGetValue("xAxisTitle", out var xTitle) && !string.IsNullOrEmpty(xTitle))
            axis.AppendChild(BuildAxisTitle(xTitle, properties));

        // Category-axis major gridlines are off by default in Excel; opt-in.
        if (IsTruthyProp(properties, "xGridlines", defaultValue: false))
        {
            var gl = new CX.MajorGridlinesGridlines();
            // CONSISTENCY(chart-text-style): category-axis gridline color uses
            // `xGridlineColor` to distinguish from value-axis `gridlineColor`.
            var xglColor = properties.GetValueOrDefault("xGridlineColor")
                        ?? properties.GetValueOrDefault("xGridline.color");
            if (!string.IsNullOrEmpty(xglColor))
                gl.ShapeProperties = BuildGridlineShapeProperties(xglColor);
            axis.AppendChild(gl);
        }

        // Tick labels (bin range labels like "[100, 200]") are ON by default
        // to match real Excel output. Opt out with tickLabels=false. Note
        // that cx:tickLabels itself is an EMPTY element per CT_TickLabels —
        // label text styling lives on the axis's own cx:txPr sibling (below),
        // NOT inside tickLabels. Nesting txPr under tickLabels produces
        // schema-invalid XML that Excel silently "repairs".
        if (IsTruthyProp(properties, "tickLabels", defaultValue: true))
            axis.AppendChild(new CX.TickLabels());

        // CONSISTENCY(chart-text-style): axis-level cx:txPr styles tick
        // labels AND axis title text, matching what ApplyAxisTextProperties
        // does for regular cChart. Compound form `axisfont=size:color:fontname`.
        // Must be AFTER tickLabels per CT_Axis schema sequence
        // (catScaling → title → gridlines → tickLabels → numFmt → spPr → txPr).
        var axisFont = properties.GetValueOrDefault("axisfont")
                    ?? properties.GetValueOrDefault("axis.font");
        if (!string.IsNullOrEmpty(axisFont))
        {
            var tickTxPr = BuildAxisTickLabelStyle(axisFont);
            if (tickTxPr != null) axis.AppendChild(tickTxPr);
        }

        return axis;
    }

    private static CX.Axis BuildValueAxis(uint id, Dictionary<string, string> properties)
    {
        var axis = new CX.Axis { Id = id };
        axis.AppendChild(new CX.ValueAxisScaling());

        if (properties.TryGetValue("yAxisTitle", out var yTitle) && !string.IsNullOrEmpty(yTitle))
            axis.AppendChild(BuildAxisTitle(yTitle, properties));

        // Value-axis gridlines are ON by default — matches Excel's histogram
        // and column charts out of the box.
        if (IsTruthyProp(properties, "gridlines", defaultValue: true))
        {
            var gl = new CX.MajorGridlinesGridlines();
            var glColor = properties.GetValueOrDefault("gridlineColor")
                       ?? properties.GetValueOrDefault("gridline.color");
            if (!string.IsNullOrEmpty(glColor))
                gl.ShapeProperties = BuildGridlineShapeProperties(glColor);
            axis.AppendChild(gl);
        }

        if (IsTruthyProp(properties, "tickLabels", defaultValue: true))
            axis.AppendChild(new CX.TickLabels());

        // cx:txPr must come after tickLabels per CT_Axis schema. See the
        // CONSISTENCY(chart-text-style) note in BuildCategoryAxis above.
        var axisFont = properties.GetValueOrDefault("axisfont")
                    ?? properties.GetValueOrDefault("axis.font");
        if (!string.IsNullOrEmpty(axisFont))
        {
            var tickTxPr = BuildAxisTickLabelStyle(axisFont);
            if (tickTxPr != null) axis.AppendChild(tickTxPr);
        }

        return axis;
    }

    private static bool IsTruthyProp(Dictionary<string, string> properties, string key, bool defaultValue)
    {
        if (!properties.TryGetValue(key, out var v) || string.IsNullOrEmpty(v))
            return defaultValue;
        return !(v.Equals("false", StringComparison.OrdinalIgnoreCase)
              || v.Equals("off", StringComparison.OrdinalIgnoreCase)
              || v == "0"
              || v.Equals("no", StringComparison.OrdinalIgnoreCase));
    }

    private static CX.Data BuildDataBlock(uint id, string chartType, string[]? categories, double[] values)
    {
        var data = new CX.Data { Id = id };

        // String dimension for categories (if provided)
        if (categories != null && chartType is "funnel" or "treemap" or "sunburst" or "boxwhisker")
        {
            var strDim = new CX.StringDimension { Type = CX.StringDimensionType.Cat };
            var strLvl = new CX.StringLevel { PtCount = (uint)categories.Length };
            for (int i = 0; i < categories.Length; i++)
                strLvl.AppendChild(new CX.ChartStringValue(categories[i]) { Index = (uint)i });
            strDim.AppendChild(strLvl);
            data.AppendChild(strDim);
        }

        // Numeric dimension
        var numType = chartType is "treemap" or "sunburst"
            ? CX.NumericDimensionType.Size
            : CX.NumericDimensionType.Val;
        var numDim = new CX.NumericDimension { Type = numType };
        var numLvl = new CX.NumericLevel { PtCount = (uint)values.Length, FormatCode = "General" };
        for (int i = 0; i < values.Length; i++)
            numLvl.AppendChild(new CX.NumericValue(values[i].ToString("G")) { Idx = (uint)i });
        numDim.AppendChild(numLvl);
        data.AppendChild(numDim);

        return data;
    }

    private static CX.SeriesLayoutProperties? BuildLayoutProperties(
        string chartType, Dictionary<string, string> properties, int valueCount)
    {
        switch (chartType)
        {
            case "treemap":
            {
                var lp = new CX.SeriesLayoutProperties();
                var parentLayout = properties.GetValueOrDefault("parentLabelLayout") ?? "overlapping";
                lp.AppendChild(new CX.ParentLabelLayout
                {
                    ParentLabelLayoutVal = parentLayout.ToLowerInvariant() switch
                    {
                        "none" => CX.ParentLabelLayoutVal.None,
                        "banner" => CX.ParentLabelLayoutVal.Banner,
                        _ => CX.ParentLabelLayoutVal.Overlapping
                    }
                });
                return lp;
            }
            case "boxwhisker":
            {
                var lp = new CX.SeriesLayoutProperties();
                lp.AppendChild(new CX.SeriesElementVisibilities
                {
                    MeanLine = false, MeanMarker = true,
                    Nonoutliers = false, Outliers = true
                });
                var method = properties.GetValueOrDefault("quartileMethod") ?? "exclusive";
                lp.AppendChild(new CX.Statistics
                {
                    QuartileMethod = method.ToLowerInvariant() switch
                    {
                        "inclusive" => CX.QuartileMethod.Inclusive,
                        _ => CX.QuartileMethod.Exclusive
                    }
                });
                return lp;
            }
            case "histogram":
            {
                // cx:layoutPr > cx:binning (empty for auto-bin; child cx:binCount
                // OR cx:binSize for explicit bin count/width). `cx:aggregation`
                // is for Pareto charts and causes Excel to render the whole
                // dataset as a single bar.
                //
                // NOTE: the Open XML SDK models cx:binCount as a leaf text
                // element (BinCountXsdunsignedInt → `<cx:binCount>5</cx:binCount>`),
                // but real Excel writes it as an empty element with a `val`
                // attribute (`<cx:binCount val="5"/>`). SDK's form is schema-
                // valid per the generated type metadata but Excel rejects the
                // whole file with "We found a problem with some content"
                // and deletes the drawing. Same applies to cx:binSize. Work
                // around by appending a raw OpenXmlUnknownElement carrying
                // the correct form.
                const string cxNs = "http://schemas.microsoft.com/office/drawing/2014/chartex";
                var lp = new CX.SeriesLayoutProperties();
                var binning = new CX.Binning();

                // intervalClosed: "r" (default, bins are (a,b]) or "l" (bins are [a,b))
                var intervalClosed = properties.GetValueOrDefault("intervalClosed") ?? "r";
                binning.IntervalClosed = intervalClosed.ToLowerInvariant() switch
                {
                    "l" => CX.IntervalClosedSide.L,
                    _   => CX.IntervalClosedSide.R,
                };

                // underflow / overflow: cut-off values for outlier bins
                if (properties.TryGetValue("underflowBin", out var underflow))
                    binning.Underflow = underflow;
                if (properties.TryGetValue("overflowBin", out var overflow))
                    binning.Overflow = overflow;

                // binCount (explicit count) XOR binSize (explicit width). If
                // both are given, binCount wins (it's the more common knob).
                if (properties.TryGetValue("binCount", out var binCountStr) &&
                    uint.TryParse(binCountStr, out var binCount))
                {
                    var binCountEl = new OpenXmlUnknownElement("cx", "binCount", cxNs);
                    binCountEl.SetAttribute(new OpenXmlAttribute("val", "", binCount.ToString()));
                    binning.AppendChild(binCountEl);
                }
                else if (properties.TryGetValue("binSize", out var binSizeStr) &&
                         double.TryParse(binSizeStr, System.Globalization.NumberStyles.Float,
                             System.Globalization.CultureInfo.InvariantCulture, out var binSize))
                {
                    var binSizeEl = new OpenXmlUnknownElement("cx", "binSize", cxNs);
                    binSizeEl.SetAttribute(new OpenXmlAttribute("val", "",
                        binSize.ToString("G", System.Globalization.CultureInfo.InvariantCulture)));
                    binning.AppendChild(binSizeEl);
                }

                lp.AppendChild(binning);
                return lp;
            }
            default:
                return null;
        }
    }

    private static CX.SeriesLayout ParseSeriesLayout(string layoutId)
    {
        return layoutId switch
        {
            "funnel" => CX.SeriesLayout.Funnel,
            "treemap" => CX.SeriesLayout.Treemap,
            "sunburst" => CX.SeriesLayout.Sunburst,
            "boxWhisker" => CX.SeriesLayout.BoxWhisker,
            "clusteredColumn" => CX.SeriesLayout.ClusteredColumn,
            "paretoLine" => CX.SeriesLayout.ParetoLine,
            "regionMap" => CX.SeriesLayout.RegionMap,
            _ => CX.SeriesLayout.Funnel
        };
    }

    /// <summary>
    /// Detect if a cx:chartSpace contains an extended chart type and return the type name.
    /// </summary>
    internal static string? DetectExtendedChartType(CX.ChartSpace chartSpace)
    {
        var series = chartSpace.Descendants<CX.Series>().FirstOrDefault();
        var layoutId = series?.LayoutId?.InnerText;
        if (layoutId == null) return null;
        return layoutId switch
        {
            "funnel" => "funnel",
            "treemap" => "treemap",
            "sunburst" => "sunburst",
            "boxWhisker" => "boxWhisker",
            "clusteredColumn" => "histogram",
            "paretoLine" => "pareto",
            "regionMap" => "regionMap",
            _ => layoutId
        };
    }
}
