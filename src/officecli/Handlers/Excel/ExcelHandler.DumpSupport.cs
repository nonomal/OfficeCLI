// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeCli.Core;

namespace OfficeCli.Handlers;

public partial class ExcelHandler
{
    // ==================== Dump support ====================
    //
    // CONSISTENCY(emit-X-mirror): read-only enumeration surface consumed by
    // ExcelBatchEmitter, mirroring the public helper methods PowerPointHandler
    // grew for PptxBatchEmitter (GetSlideBulletImageParts, GetTimingAudioRels,
    // ...). The emitter lives outside the handler's partial-class family, so
    // everything it needs beyond Get/Query is exposed here.

    /// <summary>Sheet names in workbook (sldIdLst-equivalent) order.</summary>
    public List<string> GetDumpSheetNames() => GetWorksheets().Select(t => t.Name).ToList();

    /// <summary>
    /// Workbook-level settings node (date1904, calc.*, activeTab, protection).
    /// Same Format keys as PopulateWorkbookSettings emits on Get.
    /// </summary>
    public DocumentNode GetDumpWorkbookNode()
    {
        var node = new DocumentNode { Path = "/workbook", Type = "workbook" };
        PopulateWorkbookSettings(node);
        return node;
    }

    /// <summary>
    /// Enumerate every row of a sheet with ALL cells that carry content OR
    /// style. The bulk Get path (GetSheetChildNodes) intentionally omits
    /// styled-empty cells (&lt;c s="1"/&gt;, issue #149 bloat guard); a dump
    /// must include them because their xf holds user-visible formatting
    /// (filled header bands, bordered empty grids). Each cell node is built
    /// by the same CellToNode Get uses, so Format keys match Get exactly.
    /// A dump-only <c>__raw</c> Format key carries the raw stored
    /// &lt;x:v&gt; text so the emitter can reproduce numbers/dates without
    /// going through display formatting.
    /// </summary>
    public List<DocumentNode> GetDumpRowNodes(string sheetName)
    {
        var worksheet = FindWorksheet(sheetName)
            ?? throw new ArgumentException($"Sheet not found: {sheetName}");
        var rows = new List<DocumentNode>();
        var sheetData = GetSheet(worksheet).GetFirstChild<SheetData>();
        if (sheetData == null) return rows;

        // One evaluator per sheet: CellToNode lazily creates a fresh
        // FormulaEvaluator per formula cell when none is passed, which is
        // O(cells × sheet-size) on formula-heavy sheets.
        var eval = new Core.FormulaEvaluator(sheetData, _doc.WorkbookPart);
        var seenRowIndices = new HashSet<uint>();
        foreach (var row in sheetData.Elements<Row>())
        {
            var ridx = row.RowIndex?.Value ?? 0;
            if (ridx != 0 && !seenRowIndices.Add(ridx)) continue;

            var rowNode = new DocumentNode
            {
                Path = $"/{sheetName}/row[{ridx}]",
                Type = "row"
            };
            // CONSISTENCY(unit-qualified-readback): pt-suffix row height,
            // mirroring GetSheetChildNodes.
            if (row.Height?.Value != null)
                rowNode.Format["height"] = $"{row.Height.Value.ToString(System.Globalization.CultureInfo.InvariantCulture)}pt";
            if (row.Hidden?.Value == true)
                rowNode.Format["hidden"] = true;
            if (row.OutlineLevel?.Value is { } rol && rol > 0)
                rowNode.Format["outlineLevel"] = (int)rol;
            if (row.Collapsed?.Value == true)
                rowNode.Format["collapsed"] = true;

            foreach (var cell in row.Elements<Cell>())
            {
                var hasContent = CellHasContent(cell);
                var hasStyle = cell.StyleIndex != null && cell.StyleIndex.Value != 0;
                if (!hasContent && !hasStyle) continue;
                var cellNode = CellToNode(sheetName, cell, worksheet, eval);
                var raw = cell.CellValue?.Text;
                if (!string.IsNullOrEmpty(raw))
                    cellNode.Format["__raw"] = raw;
                // Rich-text marker: inline-string runs or a rich shared-string
                // entry can't ride the CSV baseline; the emitter needs to know.
                if (HasRichTextContent(cell))
                    cellNode.Format["__richtext"] = true;
                rowNode.Children.Add(cellNode);
            }

            if (rowNode.Children.Count == 0 && rowNode.Format.Count == 0) continue;
            rowNode.ChildCount = rowNode.Children.Count;
            rows.Add(rowNode);
        }
        return rows;
    }

    private bool HasRichTextContent(Cell cell)
    {
        var inline = cell.GetFirstChild<InlineString>();
        if (inline != null && inline.Elements<Run>().Any()) return true;
        if (cell.DataType?.Value == CellValues.SharedString
            && int.TryParse(cell.CellValue?.Text, out var ssIdx))
        {
            var ssItems = _doc.WorkbookPart?.SharedStringTablePart?.SharedStringTable;
            if (ssItems != null)
            {
                var item = ssItems.Elements<SharedStringItem>().ElementAtOrDefault(ssIdx);
                if (item != null && item.Elements<Run>().Any()) return true;
            }
        }
        return false;
    }

    /// <summary>
    /// Column definitions of a sheet, one node per column LETTER (Column
    /// elements with min/max ranges are expanded). Format keys mirror the
    /// /Sheet/col[X] Get surface: width, hidden, customWidth, outlineLevel,
    /// collapsed. Range expansion is capped per definition so a stray
    /// A:XFD-wide entry cannot emit 16 000 rows — the overflow is reported
    /// through <paramref name="truncated"/>.
    /// </summary>
    public List<DocumentNode> GetDumpColumnNodes(string sheetName, out bool truncated)
    {
        truncated = false;
        var worksheet = FindWorksheet(sheetName)
            ?? throw new ArgumentException($"Sheet not found: {sheetName}");
        var nodes = new List<DocumentNode>();
        var cols = GetSheet(worksheet).GetFirstChild<Columns>();
        if (cols == null) return nodes;

        const int MaxExpandPerDef = 256;
        foreach (var col in cols.Elements<Column>())
        {
            var min = (int)(col.Min?.Value ?? 0);
            var max = (int)(col.Max?.Value ?? 0);
            if (min < 1 || max < min) continue;
            var span = max - min + 1;
            if (span > MaxExpandPerDef)
            {
                truncated = true;
                max = min + MaxExpandPerDef - 1;
            }
            for (int i = min; i <= max; i++)
            {
                var letter = IndexToColumnName(i);
                var node = new DocumentNode
                {
                    Path = $"/{sheetName}/col[{letter}]",
                    Type = "column",
                    Preview = letter
                };
                if (col.Width?.Value != null && col.CustomWidth?.Value == true)
                    node.Format["width"] = col.Width.Value;
                if (col.Hidden?.Value == true) node.Format["hidden"] = true;
                if (col.OutlineLevel?.Value is { } ol && ol > 0)
                    node.Format["outlineLevel"] = (int)ol;
                if (col.Collapsed?.Value == true) node.Format["collapsed"] = true;
                if (node.Format.Count > 0) nodes.Add(node);
            }
        }
        return nodes;
    }

    /// <summary>
    /// All merged ranges of a sheet, straight from the worksheet's
    /// MergeCells element. Per-cell Format["merge"] readback cannot drive
    /// this — a merge whose cells are all empty and unstyled has no cell
    /// node at all in the dump row enumeration.
    /// </summary>
    public List<string> GetDumpMergeRanges(string sheetName)
    {
        var worksheet = FindWorksheet(sheetName)
            ?? throw new ArgumentException($"Sheet not found: {sheetName}");
        var merges = GetSheet(worksheet).GetFirstChild<MergeCells>();
        if (merges == null) return new List<string>();
        return merges.Elements<MergeCell>()
            .Select(m => m.Reference?.Value)
            .Where(r => !string.IsNullOrEmpty(r))
            .Select(r => r!)
            .ToList();
    }

    /// <summary>
    /// Resolve a dump subtree token ("SheetName" or "sheet[N]") to the
    /// canonical sheet name; null when it does not resolve.
    /// </summary>
    public string? ResolveDumpSheetName(string token)
    {
        try
        {
            var resolved = ResolveSheetName(token);
            return FindWorksheet(resolved) != null ? resolved : null;
        }
        catch { return null; }
    }

    /// <summary>
    /// Cheap package-part scan for per-sheet content the batch emit does not
    /// round-trip yet. Mirrors PptxBatchEmitter's EmitAuxiliaryPartsScan
    /// contract: silent data loss is worse than a noisy warning.
    /// </summary>
    public List<(string Element, string Reason)> GetDumpUnsupportedFeatures(string sheetName)
    {
        var result = new List<(string, string)>();
        var worksheet = FindWorksheet(sheetName);
        if (worksheet == null) return result;
        var ws = GetSheet(worksheet);

        void AddIf(bool present, string element, string reason)
        {
            if (present) result.Add((element, reason));
        }

        AddIf(ws.Elements<ConditionalFormatting>().Any(), "conditionalformatting",
            "conditional formats are not round-tripped by dump yet");
        AddIf(ws.GetFirstChild<DataValidations>()?.Elements<DataValidation>().Any() == true, "validation",
            "data validations are not round-tripped by dump yet");
        AddIf(worksheet.TableDefinitionParts.Any(), "table",
            "tables (listobjects) are not round-tripped by dump yet");
        AddIf(worksheet.PivotTableParts.Any(), "pivottable",
            "pivot tables are not round-tripped by dump yet");
        AddIf(worksheet.DrawingsPart?.ChartParts.Any() == true, "chart",
            "charts are not round-tripped by dump yet");
        AddIf(worksheet.DrawingsPart != null && worksheet.DrawingsPart.ImageParts.Any(), "picture",
            "pictures are not round-tripped by dump yet");
        AddIf(worksheet.DrawingsPart != null
              && !worksheet.DrawingsPart.ChartParts.Any()
              && !worksheet.DrawingsPart.ImageParts.Any(), "drawing",
            "drawing shapes are not round-tripped by dump yet");
        AddIf(worksheet.WorksheetCommentsPart?.Comments != null, "comment",
            "cell comments are not round-tripped by dump yet");
        AddIf(worksheet.EmbeddedObjectParts.Any() || worksheet.EmbeddedPackageParts.Any(), "ole",
            "embedded OLE objects are not round-tripped by dump yet");
        // Sparklines / slicers live in worksheet extLst.
        var extLst = ws.GetFirstChild<WorksheetExtensionList>();
        if (extLst != null)
        {
            var extXml = extLst.OuterXml;
            AddIf(extXml.Contains("sparklineGroups", StringComparison.Ordinal), "sparkline",
                "sparklines are not round-tripped by dump yet");
            AddIf(extXml.Contains("slicerList", StringComparison.OrdinalIgnoreCase), "slicer",
                "slicers are not round-tripped by dump yet");
        }
        return result;
    }
}
