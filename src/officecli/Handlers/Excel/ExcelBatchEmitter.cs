// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

using System.Globalization;
using System.Text;
using OfficeCli.Core;

namespace OfficeCli.Handlers;

// CONSISTENCY(emit-X-mirror): scaffold mirrors WordBatchEmitter.cs /
// PptxBatchEmitter.cs — same public entry shape (full-doc + subtree
// overloads), same Get-driven transcription, same UnsupportedWarning
// contract.
//
// Excel's structural difference from docx/pptx: the document is a sparse
// grid with content addressing (/Sheet1/A1), not an ordered child list.
// That removes the positional-drift machinery (ordinal stubs, deferred
// connectors) but adds a volume problem — a 100k-cell sheet must not emit
// 100k `add cell` rows. The VALUE layer therefore rides `import` items
// (CSV baseline, replayed through ExcelHandler.Import), followed by:
//   1. corrective `set` rows for cells whose stored type import's
//      type-detection would get wrong (string "123", string "TRUE",
//      ISO-date-shaped strings, quote-prefixed text, error cells,
//      array formulas);
//   2. a FORMAT layer of `set` rows (cell styles run-length grouped into
//      ranges per row, merges, row heights, column widths, sheet-level
//      settings).
//
// PR1 scope: sheets + values + formulas + cell styles + merges + links +
// rows/cols + sheet & workbook settings + named ranges. Charts, tables,
// conditional formats, validations, comments, pivots, drawings, sparklines
// and slicers surface as unsupported_element warnings (see
// ExcelHandler.GetDumpUnsupportedFeatures).
public static partial class ExcelBatchEmitter
{
    /// <summary>
    /// Captured at emit time when a sheet carries content we cannot
    /// round-trip through the existing handler vocabulary. Mirrors
    /// PptxBatchEmitter.UnsupportedWarning.
    /// </summary>
    public sealed record UnsupportedWarning(string Element, string Path, string Reason);

    // Style keys CellToNode emits that Set (via ExcelStyleManager) accepts
    // verbatim. Anything else on a cell node is either handled specially
    // (formula/merge/link/...) or skipped via SkippedCellKeys.
    private static readonly HashSet<string> CellStyleKeyPrefixes = new(StringComparer.OrdinalIgnoreCase)
    {
        "font.", "border.", "alignment.", "protection."
    };
    private static readonly HashSet<string> CellStyleKeys = new(StringComparer.OrdinalIgnoreCase)
    {
        "fill", "numberformat", "strike", "underline", "superscript", "subscript",
    };
    // Cell Format keys the emitter consumes through dedicated channels or
    // intentionally drops (derived/readonly state).
    private static readonly HashSet<string> HandledOrDerivedCellKeys = new(StringComparer.OrdinalIgnoreCase)
    {
        "type", "formula", "cachedValue", "computedValue", "evaluated", "empty",
        "merge", "link", "tooltip", "arrayformula", "arrayref", "numFmtId",
        "quotePrefix", "phonetic", "__raw", "__richtext",
    };

    // Sheet-level Format(Get) key → Set key mapping. Only pairs verified on
    // both sides of the handler are listed; unlisted sheet keys are ignored.
    private static readonly (string GetKey, string SetKey)[] SheetSettingMap =
    {
        ("freeze", "freeze"),
        ("zoom", "zoom"),
        ("tabColor", "tabcolor"),
        ("autoFilter", "autofilter"),
        ("orientation", "orientation"),
        ("paperSize", "papersize"),
        ("fitToPage", "fittopage"),
        ("printArea", "printarea"),
        ("header", "header"),
        ("footer", "footer"),
        ("margin.top", "margin.top"),
        ("margin.bottom", "margin.bottom"),
        ("margin.left", "margin.left"),
        ("margin.right", "margin.right"),
        ("margin.header", "margin.header"),
        ("margin.footer", "margin.footer"),
    };

    // Workbook-level keys where Get key == Set case (Set.Workbook.cs).
    // calc.fullPrecision is emitted by Get even at its spec default (true),
    // so it is only carried when false.
    private static readonly string[] WorkbookSettingKeys =
    {
        "workbook.date1904", "workbook.codeName", "workbook.filterPrivacy",
        "workbook.showObjects", "workbook.backupFile", "workbook.dateCompatibility",
        "calc.mode", "calc.iterate", "calc.iterateCount", "calc.iterateDelta",
        "calc.fullCalcOnLoad", "calc.refMode",
        "activeTab", "firstSheet",
        "workbook.lockStructure", "workbook.lockWindows",
    };

    /// <summary>Emit a full Excel workbook as a sequence of BatchItem rows.</summary>
    public static (List<BatchItem> Items, List<UnsupportedWarning> Warnings) EmitExcel(ExcelHandler xl)
    {
        var items = new List<BatchItem>();
        var warnings = new List<UnsupportedWarning>();

        EmitWorkbookSettings(xl, items, warnings);

        var sheetNames = xl.GetDumpSheetNames();
        for (int i = 0; i < sheetNames.Count; i++)
            EmitSheet(xl, sheetNames[i], renameFirstSheet: i == 0, items, warnings);

        EmitNamedRanges(xl, items, warnings);
        EmitDocPropsScan(xl, warnings);

        return (items, warnings);
    }

    /// <summary>
    /// Emit a subtree. Supported paths: `/` (full document), `/SheetName`,
    /// `/sheet[N]`. A single-sheet dump emits `add sheet` (not the
    /// rename-first-sheet form) so it can replay onto a workbook that
    /// already has content; workbook-level settings and named ranges are
    /// NOT included (they live at sibling paths — mirrors the docx/pptx
    /// subtree contract).
    /// </summary>
    public static (List<BatchItem> Items, List<UnsupportedWarning> Warnings) EmitExcel(
        ExcelHandler xl, string path)
    {
        const string SupportedHint = "Supported: /, /SheetName, /sheet[N]";
        if (string.IsNullOrEmpty(path))
            throw new CliException($"dump path cannot be empty. Use '/' for the full document or a sheet path like /Sheet1. {SupportedHint}")
                { Code = "invalid_path" };
        if (path == "/") return EmitExcel(xl);

        var token = path.Trim('/');
        if (token.Length == 0 || token.Contains('/'))
            throw new CliException($"dump path not supported: {path}. {SupportedHint}")
                { Code = "unsupported_path" };

        var sheetName = xl.ResolveDumpSheetName(token)
            ?? throw new CliException($"dump path not found: {path} (no such sheet)")
                { Code = "path_not_found" };

        var items = new List<BatchItem>();
        var warnings = new List<UnsupportedWarning>();
        EmitSheet(xl, sheetName, renameFirstSheet: false, items, warnings);
        return (items, warnings);
    }

    private static void EmitWorkbookSettings(ExcelHandler xl, List<BatchItem> items,
        List<UnsupportedWarning> warnings)
    {
        DocumentNode wb;
        try { wb = xl.GetDumpWorkbookNode(); }
        catch { return; }

        var props = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        foreach (var key in WorkbookSettingKeys)
        {
            if (!wb.Format.TryGetValue(key, out var v) || v == null) continue;
            var s = FormatValue(v);
            if (s.Length > 0) props[key] = s;
        }
        if (wb.Format.TryGetValue("calc.fullPrecision", out var fp) && fp is bool fpB && !fpB)
            props["calc.fullPrecision"] = "false";
        if (wb.Format.ContainsKey("workbook.password"))
            warnings.Add(new UnsupportedWarning("workbook.password", "/workbook",
                "protection password hashes cannot be round-tripped; workbook protection is emitted without a password"));

        if (props.Count == 0) return;
        items.Add(new BatchItem { Command = "set", Path = "/", Props = props });
    }

    private static void EmitNamedRanges(ExcelHandler xl, List<BatchItem> items,
        List<UnsupportedWarning> warnings)
    {
        DocumentNode list;
        try { list = xl.Get("/namedrange"); }
        catch { return; }
        if (list.Children == null) return;
        foreach (var nr in list.Children)
        {
            var name = nr.Format.TryGetValue("name", out var n) ? n as string : null;
            var refVal = nr.Format.TryGetValue("ref", out var r) ? r as string : null;
            if (string.IsNullOrEmpty(name) || string.IsNullOrEmpty(refVal)) continue;
            // Excel's builtin names (_xlnm.Print_Area etc.) are carried by the
            // sheet-level printarea/printtitlerows emits — re-adding them here
            // would duplicate the defined name.
            if (name!.StartsWith("_xlnm.", StringComparison.OrdinalIgnoreCase)) continue;
            var props = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            {
                ["name"] = name!,
                ["ref"] = refVal!,
            };
            var scope = nr.Format.TryGetValue("scope", out var sc) ? sc as string : null;
            // AddNamedRange defaults a sheet-parent to sheet scope; pin the
            // scope explicitly so parent "/" + scope=<sheet> round-trips.
            props["scope"] = string.IsNullOrEmpty(scope) ? "workbook" : scope!;
            if (nr.Format.TryGetValue("comment", out var cm) && cm is string cs && cs.Length > 0)
                props["comment"] = cs;
            try
            {
                items.Add(new BatchItem { Command = "add", Parent = "/", Type = "namedrange", Props = props });
            }
            catch
            {
                warnings.Add(new UnsupportedWarning("namedrange", nr.Path ?? "/namedrange",
                    $"defined name '{name}' could not be emitted"));
            }
        }
    }

    private static void EmitDocPropsScan(ExcelHandler xl, List<UnsupportedWarning> warnings)
    {
        // Core doc properties (title/author/...) are Set-able on "/" but Get
        // does not surface them on the workbook node yet — surface the gap
        // once per dump rather than silently dropping. Cheap probe via Get("/").
        // (Kept as a scan hook for parity with Word/Pptx aux-part scans.)
    }

    // ==================== Per-sheet emit ====================

    private static void EmitSheet(ExcelHandler xl, string sheetName, bool renameFirstSheet,
        List<BatchItem> items, List<UnsupportedWarning> warnings)
    {
        var sheetPath = "/" + sheetName;
        DocumentNode sheetNode;
        try { sheetNode = xl.Get(sheetPath); }
        catch (Exception ex)
        {
            warnings.Add(new UnsupportedWarning("sheet", sheetPath, $"sheet read failed: {ex.Message}"));
            return;
        }

        // 1. Create / claim the sheet slot.
        if (renameFirstSheet)
        {
            // Replay target contract: a blank workbook (officecli create) whose
            // single placeholder sheet is claimed by renaming. `add sheet`
            // would refuse the duplicate name when the source's first sheet is
            // literally "Sheet1", and appending would leave the placeholder
            // dangling — rename sidesteps both.
            if (!string.Equals(sheetName, "Sheet1", StringComparison.Ordinal))
                items.Add(new BatchItem
                {
                    Command = "set",
                    Path = "/sheet[1]",
                    Props = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase) { ["name"] = sheetName },
                });
        }
        else
        {
            items.Add(new BatchItem
            {
                Command = "add",
                Parent = "/",
                Type = "sheet",
                Props = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase) { ["name"] = sheetName },
            });
        }

        // 2. Value baseline: CSV import blocks + corrective typed sets.
        var rows = xl.GetDumpRowNodes(sheetName);
        var corrective = new List<BatchItem>();
        var styleRows = new List<BatchItem>();

        EmitValueBaseline(sheetName, rows, items, corrective, warnings);

        // 3. Cell-level passes off the same row snapshot: corrective typed
        // sets first (they finalize VALUES), then styles (they only touch xf).
        items.AddRange(corrective);

        foreach (var rowNode in rows)
        {
            if (rowNode.Children == null) continue;
            EmitRowCellStyles(sheetName, rowNode, styleRows);
            foreach (var cell in rowNode.Children)
            {
                if (cell.Format.TryGetValue("link", out var lv) && lv is string link && link.Length > 0)
                {
                    var linkProps = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase) { ["link"] = link };
                    if (cell.Format.TryGetValue("tooltip", out var tt) && tt is string tts && tts.Length > 0)
                        linkProps["tooltip"] = tts;
                    styleRows.Add(new BatchItem { Command = "set", Path = cell.Path, Props = linkProps });
                }
                if (cell.Format.TryGetValue("__richtext", out _))
                    warnings.Add(new UnsupportedWarning("richtext", cell.Path ?? sheetPath,
                        "rich-text runs are flattened to plain text on dump"));
                if (cell.Format.TryGetValue("phonetic", out _))
                    warnings.Add(new UnsupportedWarning("phonetic", cell.Path ?? sheetPath,
                        "phonetic (furigana) guides are not round-tripped by dump"));
            }
        }

        items.AddRange(styleRows);

        // 4. Merges — read from the worksheet's MergeCells element (per-cell
        // Format["merge"] misses all-empty merges); one sheet-level set
        // carrying every range.
        var mergeRanges = xl.GetDumpMergeRanges(sheetName);
        if (mergeRanges.Count > 0)
            items.Add(new BatchItem
            {
                Command = "set",
                Path = sheetPath,
                Props = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
                {
                    ["merge"] = string.Join(",", mergeRanges),
                },
            });

        // 5. Row heights / hidden / outline.
        foreach (var rowNode in rows)
        {
            var rp = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            if (rowNode.Format.TryGetValue("height", out var h) && h is string hs) rp["height"] = hs;
            if (rowNode.Format.TryGetValue("hidden", out var hd) && hd is bool hb && hb) rp["hidden"] = "true";
            if (rowNode.Format.TryGetValue("outlineLevel", out var rolv)) rp["outline"] = FormatValue(rolv);
            if (rowNode.Format.TryGetValue("collapsed", out var rc) && rc is bool rcb && rcb) rp["collapsed"] = "true";
            if (rp.Count > 0)
                items.Add(new BatchItem { Command = "set", Path = rowNode.Path, Props = rp });
        }

        // 6. Column widths / hidden / outline.
        var colNodes = xl.GetDumpColumnNodes(sheetName, out var colsTruncated);
        if (colsTruncated)
            warnings.Add(new UnsupportedWarning("column", sheetPath,
                "a column definition spans more than 256 columns; the tail was dropped"));
        foreach (var colNode in colNodes)
        {
            var cp = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            if (colNode.Format.TryGetValue("width", out var w)) cp["width"] = FormatValue(w);
            if (colNode.Format.TryGetValue("hidden", out var chd) && chd is bool chb && chb) cp["hidden"] = "true";
            if (colNode.Format.TryGetValue("outlineLevel", out var colv)) cp["outline"] = FormatValue(colv);
            if (colNode.Format.TryGetValue("collapsed", out var cc) && cc is bool ccb && ccb) cp["collapsed"] = "true";
            if (cp.Count > 0)
                items.Add(new BatchItem { Command = "set", Path = colNode.Path, Props = cp });
        }

        // 7. Sheet-level settings (freeze/zoom/tab color/autofilter/print...).
        EmitSheetSettings(sheetNode, sheetPath, items, warnings);

        // 8. Unsupported-content scan.
        foreach (var (element, reason) in xl.GetDumpUnsupportedFeatures(sheetName))
            warnings.Add(new UnsupportedWarning(element, sheetPath, reason));
    }

    private static void EmitSheetSettings(DocumentNode sheetNode, string sheetPath,
        List<BatchItem> items, List<UnsupportedWarning> warnings)
    {
        var props = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        foreach (var (getKey, setKey) in SheetSettingMap)
        {
            if (!sheetNode.Format.TryGetValue(getKey, out var v) || v == null) continue;
            var s = FormatValue(v);
            if (s.Length > 0) props[setKey] = s;
        }
        // Toggle-off keys: Get emits gridlines/headings only when false.
        if (sheetNode.Format.TryGetValue("gridlines", out var gl) && gl is bool glB && !glB)
            props["showgridlines"] = "false";
        if (sheetNode.Format.TryGetValue("headings", out var hdg) && hdg is bool hdB && !hdB)
            props["showrowcolheaders"] = "false";
        if (sheetNode.Format.TryGetValue("direction", out var dir) && dir is string ds
            && ds.Equals("rtl", StringComparison.OrdinalIgnoreCase))
            props["rtl"] = "true";
        if (props.Count > 0)
            items.Add(new BatchItem { Command = "set", Path = sheetPath, Props = props });

        // Visibility & protection ride separate rows: hiding a sheet mid-way
        // through its own settings row could interact with active-tab logic,
        // and protect must land LAST so earlier replay sets aren't blocked.
        if (sheetNode.Format.TryGetValue("visibility", out var vis) && vis is string visS && visS.Length > 0)
            items.Add(new BatchItem
            {
                Command = "set",
                Path = sheetPath,
                Props = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase) { ["visibility"] = visS },
            });
        if (sheetNode.Format.TryGetValue("protect", out var prot) && prot is bool pb && pb)
        {
            items.Add(new BatchItem
            {
                Command = "set",
                Path = sheetPath,
                Props = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase) { ["protect"] = "true" },
            });
            warnings.Add(new UnsupportedWarning("sheet.password", sheetPath,
                "sheet protection password hashes cannot be round-tripped; protection is emitted without a password"));
        }
    }

    // ==================== Value baseline (CSV import) ====================

    // Split the sheet into row blocks separated by gaps of more than
    // GapRowThreshold contentless rows, so a sparse sheet with data at A1
    // and A100000 doesn't produce a 100k-line CSV of empty rows.
    private const int GapRowThreshold = 50;
    // Rows per import item — bounds the CSV string a single BatchItem carries.
    private const int MaxRowsPerImport = 20_000;

    private sealed record CsvCell(int Col, string Text);

    private static void EmitValueBaseline(string sheetName, List<DocumentNode> rows,
        List<BatchItem> items, List<BatchItem> corrective, List<UnsupportedWarning> warnings)
    {
        // Collect CSV-safe payloads per row index.
        var csvRows = new SortedDictionary<uint, List<CsvCell>>();

        foreach (var rowNode in rows)
        {
            if (rowNode.Children == null) continue;
            foreach (var cell in rowNode.Children)
            {
                var cellRef = LastPathSegment(cell.Path);
                if (!TryParseCellRef(cellRef, out var colIdx, out var rowIdx)) continue;
                var payload = ClassifyCell(cell, corrective, warnings);
                if (payload == null) continue;
                if (!csvRows.TryGetValue(rowIdx, out var list))
                    csvRows[rowIdx] = list = new List<CsvCell>();
                list.Add(new CsvCell(colIdx, payload));
            }
        }

        if (csvRows.Count == 0) return;

        // Walk row indices, grouping into blocks split on large gaps.
        var block = new List<(uint Row, List<CsvCell> Cells)>();
        uint prevRow = 0;
        foreach (var (rowIdx, cells) in csvRows)
        {
            var gapExceeded = block.Count > 0 && rowIdx - prevRow - 1 > GapRowThreshold;
            if (gapExceeded || block.Count >= MaxRowsPerImport)
            {
                FlushImportBlock(sheetName, block, items);
                block.Clear();
            }
            block.Add((rowIdx, cells));
            prevRow = rowIdx;
        }
        FlushImportBlock(sheetName, block, items);
    }

    private static void FlushImportBlock(string sheetName,
        List<(uint Row, List<CsvCell> Cells)> block, List<BatchItem> items)
    {
        if (block.Count == 0) return;
        var minCol = block.Min(b => b.Cells.Min(c => c.Col));
        var startRow = block[0].Row;
        var sb = new StringBuilder();
        uint expectedRow = startRow;
        foreach (var (rowIdx, cells) in block)
        {
            // In-block gaps within the threshold are bridged with blank rows
            // so row alignment holds (import writes rows sequentially from
            // the start cell). The bridge line must be "," (two empty
            // fields), NOT an empty line — ParseCsv drops single-empty-field
            // rows entirely, which would shift every later row up by one.
            while (expectedRow < rowIdx) { sb.Append(",\n"); expectedRow++; }
            var byCol = cells.OrderBy(c => c.Col).ToList();
            int cursor = minCol;
            for (int i = 0; i < byCol.Count; i++)
            {
                while (cursor < byCol[i].Col) { sb.Append(','); cursor++; }
                sb.Append(EscapeCsvField(byCol[i].Text));
                // Field content written; the NEXT cell needs a separator.
                cursor++;
                if (i < byCol.Count - 1) sb.Append(',');
            }
            sb.Append('\n');
            expectedRow++;
        }
        items.Add(new BatchItem
        {
            Command = "import",
            Parent = "/" + sheetName,
            Text = sb.ToString(),
            Props = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            {
                ["start-cell"] = $"{ColumnName(minCol)}{startRow}",
            },
        });
    }

    /// <summary>
    /// Decide how a cell's VALUE replays: return the CSV field text, or null
    /// when the cell has no value payload (style-only) or is handled through
    /// a corrective set row appended to <paramref name="corrective"/>.
    /// </summary>
    private static string? ClassifyCell(DocumentNode cell, List<BatchItem> corrective,
        List<UnsupportedWarning> warnings)
    {
        var raw = cell.Format.TryGetValue("__raw", out var rv) ? rv as string : null;
        var type = cell.Format.TryGetValue("type", out var tv) ? tv as string ?? "" : "";
        var formula = cell.Format.TryGetValue("formula", out var fv) ? fv as string : null;
        var quotePrefix = cell.Format.TryGetValue("quotePrefix", out var qp) && qp is bool qb && qb;

        // Array formulas replay through `set arrayformula=` on the anchor
        // cell only; interior cells of the array range carry the same formula
        // + a Reference on the anchor. Non-anchor cells emit nothing (the
        // spill fills them).
        if (formula != null && cell.Format.TryGetValue("arrayformula", out var af) && af is bool afb && afb)
        {
            var arrayRef = cell.Format.TryGetValue("arrayref", out var ar) ? ar as string : null;
            var cellRef = LastPathSegment(cell.Path);
            var anchor = arrayRef?.Split(':')[0];
            if (arrayRef == null || string.Equals(anchor, cellRef, StringComparison.OrdinalIgnoreCase))
            {
                var props = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
                {
                    ["arrayformula"] = formula,
                };
                if (arrayRef != null) props["ref"] = arrayRef;
                corrective.Add(new BatchItem { Command = "set", Path = cell.Path, Props = props });
            }
            return null;
        }

        if (formula != null)
            return "=" + formula.TrimStart('=');

        // Style-only cell: no stored value at all.
        var text = cell.Text ?? "";
        if (string.IsNullOrEmpty(raw) && text.Length == 0) return null;

        switch (type)
        {
            case "Number":
            case "Date":
                // Raw stored text (serial for dates) — bypasses display
                // formatting AND import's ISO-date detection; the date
                // numberformat rides the style layer.
                return raw ?? text;
            case "Boolean":
                return raw == "1" ? "TRUE" : "FALSE";
            case "Error":
                corrective.Add(new BatchItem
                {
                    Command = "set",
                    Path = cell.Path,
                    Props = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
                    {
                        ["value"] = raw ?? text,
                        ["type"] = "error",
                    },
                });
                return null;
            default: // String / SharedString / InlineString
                if (quotePrefix)
                {
                    // Reproduce the leading-apostrophe idiom: value='<text>
                    // stores the literal + stamps quotePrefix on the xf.
                    corrective.Add(new BatchItem
                    {
                        Command = "set",
                        Path = cell.Path,
                        Props = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
                        {
                            ["value"] = "'" + text,
                        },
                    });
                    return null;
                }
                if (text.StartsWith('='))
                {
                    // `set value==...` unconditionally coerces to formula and
                    // import does the same — no vocabulary reproduces a bare
                    // string "=..." without a quote prefix.
                    warnings.Add(new UnsupportedWarning("cell.literal-equals", cell.Path ?? "",
                        "string cell text starts with '='; replayed with a quote prefix (') to keep it literal"));
                    corrective.Add(new BatchItem
                    {
                        Command = "set",
                        Path = cell.Path,
                        Props = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
                        {
                            ["value"] = "'" + text,
                        },
                    });
                    return null;
                }
                if (NeedsStringCorrection(text))
                {
                    // Import's type detection would store this as number /
                    // bool / date; pin the stored type explicitly instead.
                    corrective.Add(new BatchItem
                    {
                        Command = "set",
                        Path = cell.Path,
                        Props = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
                        {
                            ["value"] = text,
                            ["type"] = "string",
                        },
                    });
                    return null;
                }
                return text;
        }
    }

    /// <summary>
    /// True when import's SetCellValueWithTypeDetection would store the text
    /// as anything other than a plain string (mirrors its detection order:
    /// number → ISO date → boolean; '=' is handled by the caller). Leading
    /// apostrophes must also be pinned — import stores them literally but
    /// `set value='x` strips them, so keep the CSV free of that ambiguity.
    /// </summary>
    private static bool NeedsStringCorrection(string text)
    {
        if (text.Length == 0) return false;
        if (text.StartsWith('\'')) return true;
        if (double.TryParse(text, NumberStyles.Any, CultureInfo.InvariantCulture, out _)) return true;
        if (text.Equals("TRUE", StringComparison.OrdinalIgnoreCase)
            || text.Equals("FALSE", StringComparison.OrdinalIgnoreCase)) return true;
        // ISO-date shapes import converts to serials.
        string[] formats =
        {
            "yyyy-MM-dd", "yyyy-MM-ddTHH:mm:ss", "yyyy-MM-ddTHH:mm:ssZ",
            "yyyy-MM-ddTHH:mm:ss.fff", "yyyy-MM-ddTHH:mm:ss.fffZ", "yyyy-MM-dd HH:mm:ss",
        };
        if (DateTime.TryParseExact(text, formats, CultureInfo.InvariantCulture,
                DateTimeStyles.None, out _)) return true;
        return false;
    }

    // ==================== Cell style layer ====================

    private static void EmitRowCellStyles(string sheetName, DocumentNode rowNode, List<BatchItem> styleRows)
    {
        if (rowNode.Children == null || rowNode.Children.Count == 0) return;

        // Run-length group consecutive cells (same row) with byte-identical
        // style prop dictionaries into A1:D1-style range sets.
        var runs = new List<(int StartCol, int EndCol, uint Row, Dictionary<string, string> Props)>();
        foreach (var cell in rowNode.Children)
        {
            var cellRef = LastPathSegment(cell.Path);
            if (!TryParseCellRef(cellRef, out var colIdx, out var rowIdx)) continue;
            var styleProps = ExtractStyleProps(cell);
            if (styleProps.Count == 0) continue;
            if (runs.Count > 0)
            {
                var (s, e, r, p) = runs[^1];
                if (r == rowIdx && colIdx == e + 1 && SamePropDict(p, styleProps))
                {
                    runs[^1] = (s, colIdx, r, p);
                    continue;
                }
            }
            runs.Add((colIdx, colIdx, rowIdx, styleProps));
        }

        foreach (var (startCol, endCol, row, props) in runs)
        {
            var path = startCol == endCol
                ? $"/{sheetName}/{ColumnName(startCol)}{row}"
                : $"/{sheetName}/{ColumnName(startCol)}{row}:{ColumnName(endCol)}{row}";
            styleRows.Add(new BatchItem { Command = "set", Path = path, Props = props });
        }
    }

    private static Dictionary<string, string> ExtractStyleProps(DocumentNode cell)
    {
        var props = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        foreach (var (key, value) in cell.Format)
        {
            if (value == null) continue;
            if (HandledOrDerivedCellKeys.Contains(key)) continue;
            var isStyle = CellStyleKeys.Contains(key)
                || CellStyleKeyPrefixes.Any(p => key.StartsWith(p, StringComparison.OrdinalIgnoreCase));
            // font.bold / font.italic Get keys map straight through; the
            // remaining top-level toggles (strike/underline/super/subscript)
            // are in CellStyleKeys.
            if (!isStyle) continue;
            var s = FormatValue(value);
            if (s.Length == 0) continue;
            props[key] = s;
        }
        return props;
    }

    private static bool SamePropDict(Dictionary<string, string> a, Dictionary<string, string> b)
    {
        if (a.Count != b.Count) return false;
        foreach (var (k, v) in a)
            if (!b.TryGetValue(k, out var bv) || !string.Equals(v, bv, StringComparison.Ordinal))
                return false;
        return true;
    }

    // ==================== Small helpers ====================

    private static string FormatValue(object v) => v switch
    {
        bool b => b ? "true" : "false",
        double d => d.ToString("R", CultureInfo.InvariantCulture),
        _ => Convert.ToString(v, CultureInfo.InvariantCulture) ?? "",
    };

    private static string LastPathSegment(string? path)
    {
        if (string.IsNullOrEmpty(path)) return "";
        var idx = path!.LastIndexOf('/');
        return idx >= 0 ? path[(idx + 1)..] : path;
    }

    private static bool TryParseCellRef(string cellRef, out int colIdx, out uint rowIdx)
    {
        colIdx = 0; rowIdx = 0;
        int i = 0;
        while (i < cellRef.Length && char.IsAsciiLetter(cellRef[i])) i++;
        if (i == 0 || i == cellRef.Length) return false;
        int col = 0;
        for (int k = 0; k < i; k++)
        {
            var c = char.ToUpperInvariant(cellRef[k]);
            if (c < 'A' || c > 'Z') return false;
            col = col * 26 + (c - 'A' + 1);
        }
        if (!uint.TryParse(cellRef[i..], NumberStyles.None, CultureInfo.InvariantCulture, out rowIdx)
            || rowIdx == 0) return false;
        colIdx = col;
        return true;
    }

    private static string ColumnName(int index)
    {
        var sb = new StringBuilder();
        while (index > 0)
        {
            index--;
            sb.Insert(0, (char)('A' + index % 26));
            index /= 26;
        }
        return sb.ToString();
    }

    private static string EscapeCsvField(string field)
    {
        if (field.Length == 0) return field;
        var needsQuote = field.Contains(',') || field.Contains('"')
            || field.Contains('\n') || field.Contains('\r');
        if (!needsQuote) return field;
        return "\"" + field.Replace("\"", "\"\"") + "\"";
    }
}
