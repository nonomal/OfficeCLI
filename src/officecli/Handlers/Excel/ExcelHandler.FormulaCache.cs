// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeCli.Handlers;

public partial class ExcelHandler
{
    // ===== Formula-cache hygiene (stale <v> refresh at persist) =====
    //
    // officecli evaluates each formula once, at the moment its own cell is
    // written, and has NO cross-cell dependency graph. So a formula written
    // BEFORE the data it references — e.g. `set Summary/B2 =SUMIFS(Data!...)`
    // then `import` the Data rows — keeps the value it computed while the
    // precedents were empty (typically 0); the later import never refreshes it.
    //
    // Real Excel / Google Sheets recompute on open (we already emit
    // fullCalcOnLoad), so a human opening the file sees the right number. But
    // headless readers that TRUST the cache — openpyxl `data_only=True`,
    // pandas.read_excel — read the stale value as a silently-wrong number. That
    // is the one consumer class an automation CLI most needs to serve.
    //
    // At every persist we sweep formula cells and, where our own evaluator
    // CONFIDENTLY disagrees with the cached <v>, act:
    //
    //   L2 (active): drop the <v> so cache-trusting readers get an explicit
    //                "not computed" (openpyxl -> None) instead of a wrong number,
    //                and ensure fullCalcOnLoad so recalc-capable apps refill it.
    //   L1 (seam):   for formulas whose every function is on a differentially-
    //                verified allowlist (FormulaCacheL1Allowlist), write the
    //                FRESH value instead of dropping it. The allowlist is EMPTY
    //                until a function is proven byte-equal to real Excel via the
    //                officeshot golden-diff — so L1 is dormant today. Growing the
    //                allowlist (and, eventually, flipping the default) is the
    //                forward-compatible upgrade path as the evaluator matures;
    //                no architecture change required.
    //
    // Invariant: an on-disk <v> is never a value our evaluator actively
    // contradicts. We only touch a cache we can recompute AND that disagrees;
    // caches we cannot assess (unsupported function, missing-sheet ref) or that
    // already agree are left untouched — so correctly-cached cells never regress.

    /// <summary>
    /// Functions whose evaluator output has been differentially verified against
    /// real Excel and may therefore have a FRESH cached value written on save (L1).
    /// EMPTY by design: a function joins only after it passes the officeshot
    /// golden-diff. Until then every stale cache degrades to the safe L2 path.
    /// </summary>
    internal static readonly HashSet<string> FormulaCacheL1Allowlist =
        new(StringComparer.OrdinalIgnoreCase);

    /// <summary>
    /// True iff <paramref name="formula"/> contains at least one function call and
    /// EVERY function it names is on <paramref name="allowlist"/>. A function token
    /// is an identifier immediately followed by '('. Conservative: any function not
    /// covered (or no function at all) returns false, so an uncertain formula
    /// degrades to the safe L2 path. `_xlfn.`/`_xlfn._xlws.` modern-qualifier
    /// prefixes are stripped before lookup.
    /// </summary>
    internal static bool FormulaFunctionsCovered(string formula, IReadOnlySet<string> allowlist)
    {
        var sawFunction = false;
        foreach (Match m in Regex.Matches(formula, @"([A-Za-z_][A-Za-z0-9_.]*)\s*\("))
        {
            sawFunction = true;
            var fn = m.Groups[1].Value;
            if (fn.StartsWith("_xl", StringComparison.OrdinalIgnoreCase))
            {
                var dot = fn.LastIndexOf('.');
                if (dot >= 0) fn = fn[(dot + 1)..];
            }
            if (!allowlist.Contains(fn)) return false;
        }
        return sawFunction;
    }

    /// <summary>
    /// Sweep every formula cell and reconcile a disagreeing cached value with the
    /// evaluator's current result. See the file-level note for the L1/L2 policy and
    /// the on-disk invariant. Called from <see cref="FlushDirtyParts"/> at persist.
    /// </summary>
    private void RefreshStaleFormulaCaches()
    {
        if (_doc.WorkbookPart == null) return;
        // Run whenever this session changed anything. `import` marks worksheets
        // dirty without flipping Modified, and it is the classic trigger — data
        // landing after a formula that references it — so gate on either signal.
        if (!Modified && _dirtyWorksheets.Count == 0) return;
        var omittedAny = false;
        foreach (var (_, wsPart) in GetWorksheets())
        {
            var sheetData = GetSheet(wsPart).GetFirstChild<SheetData>();
            if (sheetData == null) continue;
            Core.FormulaEvaluator? evaluator = null;
            var sheetChanged = false;
            foreach (var row in sheetData.Elements<Row>())
            {
                foreach (var cell in row.Elements<Cell>())
                {
                    var formula = cell.CellFormula?.Text;
                    if (string.IsNullOrEmpty(formula)) continue;
                    // No cached value -> nothing to invalidate.
                    if (string.IsNullOrEmpty(cell.CellValue?.Text)) continue;
                    // Array / dynamic-array spill cells own a multi-cell region;
                    // their <v> is Excel's to manage — leave it alone.
                    if (cell.CellFormula?.FormulaType?.Value == CellFormulaValues.Array) continue;
                    // Missing-sheet refs are reported separately and the evaluator
                    // silently returns 0 there — that would be a false disagreement.
                    if (FormulaReferencesMissingSheet(formula)) continue;

                    evaluator ??= new Core.FormulaEvaluator(sheetData, _doc.WorkbookPart);
                    var report = evaluator.EvaluateForReport(formula);
                    // Act only when we can confidently recompute the value.
                    if (report.Status != Core.EvalReportStatus.Evaluated) continue;
                    var computed = report.Result!.ToCellValueText();
                    if (CachedComputedAgree(cell.CellValue!.Text, computed)) continue; // fresh -> keep

                    // Stale cache — decideWrite.
                    if (FormulaFunctionsCovered(formula, FormulaCacheL1Allowlist))
                    {
                        WriteFormulaResultToCell(cell, report.Result); // L1: write fresh value
                    }
                    else
                    {
                        cell.CellValue = null;                          // L2: omit (reader -> None)
                        cell.DataType = null;
                        omittedAny = true;
                    }
                    sheetChanged = true;
                }
            }
            if (sheetChanged) _dirtyWorksheets.Add(wsPart);
        }
        if (omittedAny) EnsureFullCalcOnLoad();
    }

    /// <summary>
    /// Compare a cached <x:v> string against a freshly computed value. Numeric
    /// pairs compare with a 1e-9 relative tolerance to absorb IEEE round-trip
    /// jitter Excel itself emits; non-numerics fall back to byte-equal. Shared by
    /// the save-time cache sweep and the `view issues` formula_cache_stale scan.
    /// </summary>
    internal static bool CachedComputedAgree(string cached, string computed)
    {
        if (string.Equals(cached, computed, StringComparison.Ordinal))
            return true;
        if (double.TryParse(cached, System.Globalization.NumberStyles.Float,
                System.Globalization.CultureInfo.InvariantCulture, out var a)
            && double.TryParse(computed, System.Globalization.NumberStyles.Float,
                System.Globalization.CultureInfo.InvariantCulture, out var b)
            && double.IsFinite(a) && double.IsFinite(b))
        {
            var scale = Math.Max(Math.Abs(a), Math.Abs(b));
            return Math.Abs(a - b) <= 1e-9 * Math.Max(scale, 1.0);
        }
        return false;
    }

    /// <summary>
    /// Write a freshly-evaluated formula result into a cell's value/type children,
    /// mirroring how Excel caches a computed result. Unwraps 1x1 Area results,
    /// promotes non-finite numerics to #NUM!, and clears value+type when the result
    /// is null/unevaluable (Excel recomputes on open). Shared by the formula Set
    /// path (cache-on-write) and the L1 branch of <see cref="RefreshStaleFormulaCaches"/>.
    /// </summary>
    private static void WriteFormulaResultToCell(Cell cell, Core.FormulaResult? evalResult)
    {
        // R3 BUG C: ResolveRef now always wraps even single-cell refs in an Area
        // (Round-2 change to preserve BaseRow/BaseCol). When that single cell holds
        // an Error (e.g. INDIRECT to a non-existent sheet), the result reads
        // IsRange:true rather than IsError:true. Unwrap the 1x1 Area-of-Error so the
        // cell still gets t="e" + the error sentinel as its cached value instead of
        // falling through to the "no value" branch.
        if (evalResult is { IsRange: true, RangeValue: { Rows: 1, Cols: 1 } rd1 } &&
            rd1.Cells[0, 0] is { IsError: true } innerErr)
            evalResult = innerErr;
        // BUG R4-C: same Area-of-1x1 unwrap for string / bool / numeric results from
        // OFFSET / INDIRECT. Without this the dispatch below falls through to the
        // "no value" branch — t and <v> are both dropped, producing an on-disk cell
        // that real Excel mis-parses (Get reads correctly only because in-memory
        // eval recomputes).
        if (evalResult is { IsRange: true, RangeValue: { Rows: 1, Cols: 1 } rd2 }
            && rd2.Cells[0, 0] is { } inner
            && (inner.IsString || inner.IsBool || inner.IsNumeric || inner.IsError))
            evalResult = inner;
        if (evalResult is { IsNumeric: true })
        {
            // IEEE-754 ±Infinity / NaN have no OOXML representation; writing
            // "<v>-Infinity</v>" produces a file Excel refuses to open. Promote to
            // #NUM! so the cell switches to t="e".
            var nv = evalResult.NumericValue!.Value;
            if (double.IsNaN(nv) || double.IsInfinity(nv))
            {
                cell.CellValue = new CellValue("#NUM!");
                cell.DataType = new EnumValue<CellValues>(CellValues.Error);
            }
            else
            {
                cell.CellValue = new CellValue(evalResult.ToCellValueText());
                cell.DataType = null;
            }
        }
        else if (evalResult is { IsString: true })
        {
            cell.CellValue = new CellValue(evalResult.StringValue!);
            cell.DataType = new EnumValue<CellValues>(CellValues.String);
        }
        else if (evalResult is { IsBool: true })
        {
            cell.CellValue = new CellValue(evalResult.ToCellValueText());
            cell.DataType = new EnumValue<CellValues>(CellValues.Boolean);
        }
        else if (evalResult is { IsError: true })
        {
            cell.CellValue = new CellValue(evalResult.ErrorValue!);
            cell.DataType = new EnumValue<CellValues>(CellValues.Error);
        }
        else
        {
            // Formula written but not evaluated — will be calculated when opened in Excel.
            cell.CellValue = null;
            cell.DataType = null;
        }
    }

    /// <summary>
    /// Ensure the workbook carries &lt;calcPr fullCalcOnLoad="1"/&gt; so recalc-capable
    /// applications recompute formulas on open. Inserts calcPr in OOXML schema order
    /// when absent. Shared by the formula Set path and the L2 cache-omit branch.
    /// </summary>
    private void EnsureFullCalcOnLoad()
    {
        var workbook = _doc.WorkbookPart!.Workbook!;
        var calcPr = workbook.GetFirstChild<CalculationProperties>();
        if (calcPr == null)
        {
            calcPr = new CalculationProperties();
            // OOXML schema order: ...definedNames, calcPr, oleSize, customWorkbookViews, pivotCaches...
            var insertBefore = (OpenXmlElement?)workbook.GetFirstChild<OleSize>()
                ?? (OpenXmlElement?)workbook.GetFirstChild<CustomWorkbookViews>()
                ?? (OpenXmlElement?)workbook.GetFirstChild<PivotCaches>();
            if (insertBefore != null)
                workbook.InsertBefore(calcPr, insertBefore);
            else
                workbook.AppendChild(calcPr);
        }
        calcPr.FullCalculationOnLoad = true;
    }
}
