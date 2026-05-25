// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0
//
// All Set-side revision logic for Word. Two sections, distinct call shapes:
//
//   SECTION 1 — Creation decorator (was Set.TrackChange.cs)
//     BeginTrackChangeIfRequested + HasTrackChangeKey. Called from Set() on
//     EVERY Run / Paragraph / Table / Row / Cell / SectionProperties target.
//     If the input carries any of revision.type / revision.author / revision.date
//     / revision.id, snapshots the current rPr/pPr/etc and returns a wrap action
//     that appends rPrChange/pPrChange/tblPrChange/trPrChange/tcPrChange/
//     sectPrChange AFTER SetElement mutates state.
//
//   SECTION 2 — Action dispatchers (selector + native-path)
//     IsRevisionSelectorPath / IsRevisionActionRequest plus
//     SetRevisionsBySelector / SetRevisionByNativePath. Routed BEFORE Section 1
//     in Set.cs (revision.action is action, not creation). Handles:
//       set <doc> /revision --prop revision.action=accept                  # all
//       set <doc> '/revision[@author=Alice]' --prop revision.action=accept # by author
//       set <doc> '/revision[@type=ins]' --prop revision.action=reject     # by type
//       set <doc> '/revision[@author=Bob][@type=del]' --prop revision.action=accept
//       set <doc> /revision[@id=42] --prop revision.action=accept          # by w:id (stable)
//       set <doc> /revision[3] --prop revision.action=accept               # positional fallback
//       set <doc> /body/p[N]/r[M] --prop revision.action=accept            # native path
//     All synthetic-selector forms start with `/` to satisfy
//     CONSISTENCY(no-slash-reject) in CommandBuilder.Set.cs.
//
// Key namespace (canonical, no aliases — see d9e812f3):
//   revision.type    — creation kind  (ins/del/format/moveFrom/moveTo)
//   revision.action  — action verb    (accept/reject)
//   revision.author / .date / .id — creation attribution (mixing with
//   revision.action is rejected as ambiguous; see dispatcher guards).
// The bare `revision` key is intentionally absent from both forms.
//
// Path-shift safety (Section 2): action dispatchers process matching elements
// in REVERSE document order. Some accept/reject actions remove sibling content
// (delete-runs, paragraph-merge on ¶ del); walking forward would shift later
// /revision[N] indexes mid-iteration.

using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeCli.Core;

namespace OfficeCli.Handlers;

public partial class WordHandler
{
    // ========================================================================
    // SECTION 1 — Creation decorator (rPrChange / pPrChange / *PrChange capture)
    //
    // Wiring contract (see Set(string, Dictionary<string,string>)):
    //   1. resolve the target element
    //   2. call BeginTrackChangeIfRequested; replace the user's properties dict
    //      with the returned (stripped) dict before dispatching to SetElement
    //   3. after SetElement returns successfully, invoke the returned wrapAction
    //      (no-op when no creation key was present)
    //
    // Scope:
    //   - Run / Paragraph / Table / TableRow / TableCell / SectionProperties
    //   - RTL cascade keys (font.cs / bold.cs / italic.cs / size.cs) are
    //     rejected when combined with revision.* — the cascade writes across
    //     runs and would smear the rPrChange snapshot.
    //   - An element that already carries a pending *PrChange is rejected
    //     (caller must accept/reject the existing one first).
    // ========================================================================

    /// <summary>Sentinel for the no-op case (no creation keys in the input).
    /// Returning the same Dictionary instance lets the caller short-circuit
    /// without allocating a copy.</summary>
    private static readonly Action _trackChangeNoop = static () => { };

    /// <summary>Reject the bare `revision` key with a pointed error.
    ///
    /// The branch `feat/trackchange-redesign` split the overloaded `revision`
    /// key into two disjoint namespaces:
    ///   • `revision.type=ins|del|format|moveFrom|moveTo` — create a new revision
    ///   • `revision.action=accept|reject`                — act on an existing one
    /// A bare `revision=…` literal is no longer valid. Without this guard the
    /// key would fall through to the generic unsupported-property bucket with
    /// a non-actionable error ("unsupported property: revision"); callers
    /// migrating from older scripts/dumps would have no signal pointing them
    /// at the new namespace. Called from Set() entry and from AddRun /
    /// AddParagraph so both surfaces give the same actionable message.</summary>
    internal static void RejectBareRevisionKey(Dictionary<string, string> properties)
    {
        if (properties.TryGetValue("revision", out var v))
            throw new ArgumentException(
                $"bare `revision={v}` is no longer accepted. Use "
                + "`revision.type=ins|del|format|moveFrom|moveTo` to create a "
                + "tracked change, or `revision.action=accept|reject` to act "
                + "on an existing one.");
    }

    /// <summary>Inspect <paramref name="properties"/> for revision.* creation
    /// keys. When present, snapshot the element's current rPr/pPr clone and
    /// return:
    ///   - <c>stripped</c>: copy of <paramref name="properties"/> with creation
    ///     keys removed (so downstream Set helpers don't surface them as
    ///     unsupported);
    ///   - <c>wrapAction</c>: builds and appends the *PrChange (containing the
    ///     snapshot) to the now-mutated parent. The caller must invoke this
    ///     AFTER the Set succeeds.
    /// When no creation key is present, returns the input dict unchanged and a
    /// no-op action (cheap fast path).</summary>
    private (Dictionary<string, string> stripped, Action wrapAction)
        BeginTrackChangeIfRequested(OpenXmlElement element, Dictionary<string, string> properties)
    {
        if (!HasTrackChangeKey(properties))
            return (properties, _trackChangeNoop);

        // ---- guard 1: only Run / Paragraph / Table / TableRow / TableCell /
        // SectionProperties supported ----
        if (element is not Run
            && element is not Paragraph
            && element is not Table
            && element is not TableRow
            && element is not TableCell
            && element is not SectionProperties)
            throw new InvalidOperationException(
                "trackChange capture on set is only supported for run / paragraph / table / "
                + "table-row / table-cell / section elements; other element kinds are not yet implemented.");

        // ---- guard 2: RTL cascade props would smear the snapshot ----
        foreach (var k in properties.Keys)
        {
            var lk = k.ToLowerInvariant();
            if (lk is "font.cs" or "font.complexscript" or "font.complex"
                  or "bold.cs" or "italic.cs" or "size.cs"
                  or "font.bold.cs" or "font.italic.cs" or "font.size.cs"
                  or "boldcs" or "italiccs" or "sizecs")
                throw new InvalidOperationException(
                    "RTL cascade properties are not supported with trackChange yet");
        }

        // ---- extract revision.* sub-keys (case-insensitive) ----
        string? tcAuthor = null, tcDate = null, tcId = null;
        foreach (var (k, v) in properties)
        {
            var lk = k.ToLowerInvariant();
            if (lk == "revision.author") tcAuthor = v;
            else if (lk == "revision.date") tcDate = v;
            else if (lk == "revision.id") tcId = v;
        }
        var author = string.IsNullOrEmpty(tcAuthor) ? "OfficeCLI" : tcAuthor!;
        DateTime date = DateTime.UtcNow;
        if (!string.IsNullOrEmpty(tcDate) && DateTime.TryParse(tcDate, out var parsed))
            date = parsed;
        var idStr = string.IsNullOrEmpty(tcId) ? GenerateRevisionId() : tcId!;

        // ---- strip revision.* creation keys from the dict passed to SetElement ----
        // (revision.action never reaches this decorator — it's routed in Set.cs
        // before TrackChange runs; the strip below is conservatively scoped to
        // creation keys only.)
        var stripped = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        foreach (var (k, v) in properties)
        {
            var lk = k.ToLowerInvariant();
            if (lk is "revision.type" or "revision.author" or "revision.date" or "revision.id")
                continue;
            stripped[k] = v;
        }

        // ---- snapshot + plan the wrap based on element kind ----
        if (element is Run run)
        {
            var existingRPr = run.GetFirstChild<RunProperties>();
            if (existingRPr?.GetFirstChild<RunPropertiesChange>() != null)
                throw new InvalidOperationException(
                    "element already has a pending rPrChange; accept/reject existing first");

            // Snapshot: deep-clone the current rPr's CHILDREN into a
            // <w:rPr> body, which we will host inside
            // <w:rPrChange><w:rPr>...</w:rPr></w:rPrChange>. Schema-wise the
            // inner element is the bare <w:rPr> (ECMA-376 §17.13.5.31) — the
            // SDK exposes it via the PreviousRunProperties strongly-typed
            // subclass; using a plain RunProperties round-trips as an
            // unknown sibling element. Use PreviousRunProperties.
            var snapshotInner = new PreviousRunProperties();
            if (existingRPr != null)
            {
                foreach (var child in existingRPr.ChildElements)
                {
                    if (child is RunPropertiesChange) continue;
                    snapshotInner.AppendChild(child.CloneNode(true));
                }
            }

            Action wrap = () =>
            {
                var rPr = run.GetFirstChild<RunProperties>()
                          ?? run.PrependChild(new RunProperties());
                var rprChange = new RunPropertiesChange
                {
                    Author = author,
                    Date = date,
                    Id = idStr,
                };
                rprChange.AppendChild(snapshotInner);
                // Schema CT_RPr places rPrChange last; AppendChild is correct.
                rPr.AppendChild(rprChange);
            };
            return (stripped, wrap);
        }
        else if (element is Paragraph para)
        {
            var existingPPr = para.ParagraphProperties;
            if (existingPPr?.GetFirstChild<ParagraphPropertiesChange>() != null)
                throw new InvalidOperationException(
                    "element already has a pending pPrChange; accept/reject existing first");

            // Snapshot the pPr. The strongly-typed child class for the <w:pPr>
            // inside <w:pPrChange> is ParagraphPropertiesExtended in
            // DocumentFormat.OpenXml 3.x — NOT PreviousParagraphProperties
            // (despite the parallel naming with PreviousRunProperties used by
            // rPrChange). Confirmed empirically: writing PreviousParagraphProperties
            // round-trips to ParagraphPropertiesExtended after save+reload,
            // breaking strongly-typed reads. Use ParagraphPropertiesExtended on
            // write so write and read see the same SDK type.
            var previous = new ParagraphPropertiesExtended();
            if (existingPPr != null)
            {
                foreach (var child in existingPPr.ChildElements)
                {
                    if (child is ParagraphPropertiesChange) continue;
                    previous.AppendChild(child.CloneNode(true));
                }
            }

            Action wrap = () =>
            {
                var pPr = para.ParagraphProperties ?? para.PrependChild(new ParagraphProperties());
                var pprChange = new ParagraphPropertiesChange
                {
                    Author = author,
                    Date = date,
                    Id = idStr,
                };
                pprChange.AppendChild(previous);
                // Schema CT_PPr places pPrChange last; AppendChild is correct.
                pPr.AppendChild(pprChange);
            };
            return (stripped, wrap);
        }
        else if (element is Table tbl)
        {
            var existingTblPr = tbl.GetFirstChild<TableProperties>();
            if (existingTblPr?.GetFirstChild<TablePropertiesChange>() != null)
                throw new InvalidOperationException(
                    "element already has a pending tblPrChange; accept/reject existing first");

            // For sect/tbl/tc/tr the SDK does NOT have an *Extended quirk —
            // only Previous*Properties classes exist. Use them directly.
            var previous = new PreviousTableProperties();
            if (existingTblPr != null)
            {
                foreach (var child in existingTblPr.ChildElements)
                {
                    if (child is TablePropertiesChange) continue;
                    previous.AppendChild(child.CloneNode(true));
                }
            }

            Action wrap = () =>
            {
                var tblPr = tbl.GetFirstChild<TableProperties>()
                            ?? tbl.PrependChild(new TableProperties());
                var change = new TablePropertiesChange
                {
                    Author = author,
                    Date = date,
                    Id = idStr,
                };
                change.AppendChild(previous);
                tblPr.AppendChild(change);
            };
            return (stripped, wrap);
        }
        else if (element is TableRow tr)
        {
            var existingTrPr = tr.GetFirstChild<TableRowProperties>();
            if (existingTrPr?.GetFirstChild<TableRowPropertiesChange>() != null)
                throw new InvalidOperationException(
                    "element already has a pending trPrChange; accept/reject existing first");

            var previous = new PreviousTableRowProperties();
            if (existingTrPr != null)
            {
                foreach (var child in existingTrPr.ChildElements)
                {
                    if (child is TableRowPropertiesChange) continue;
                    previous.AppendChild(child.CloneNode(true));
                }
            }

            Action wrap = () =>
            {
                var trPr = tr.GetFirstChild<TableRowProperties>()
                           ?? tr.PrependChild(new TableRowProperties());
                var change = new TableRowPropertiesChange
                {
                    Author = author,
                    Date = date,
                    Id = idStr,
                };
                change.AppendChild(previous);
                trPr.AppendChild(change);
            };
            return (stripped, wrap);
        }
        else if (element is TableCell tc)
        {
            var existingTcPr = tc.GetFirstChild<TableCellProperties>();
            if (existingTcPr?.GetFirstChild<TableCellPropertiesChange>() != null)
                throw new InvalidOperationException(
                    "element already has a pending tcPrChange; accept/reject existing first");

            var previous = new PreviousTableCellProperties();
            if (existingTcPr != null)
            {
                foreach (var child in existingTcPr.ChildElements)
                {
                    if (child is TableCellPropertiesChange) continue;
                    previous.AppendChild(child.CloneNode(true));
                }
            }

            Action wrap = () =>
            {
                var tcPr = tc.GetFirstChild<TableCellProperties>()
                           ?? tc.PrependChild(new TableCellProperties());
                var change = new TableCellPropertiesChange
                {
                    Author = author,
                    Date = date,
                    Id = idStr,
                };
                change.AppendChild(previous);
                tcPr.AppendChild(change);
            };
            return (stripped, wrap);
        }
        else
        {
            // SectionProperties — path /body/sectPr resolves to SectionProperties
            // itself, not a parent container. Snapshot SELF's children
            // (excluding existing sectPrChange).
            var sectPr = (SectionProperties)element;
            if (sectPr.GetFirstChild<SectionPropertiesChange>() != null)
                throw new InvalidOperationException(
                    "element already has a pending sectPrChange; accept/reject existing first");

            var previous = new PreviousSectionProperties();
            foreach (var child in sectPr.ChildElements)
            {
                if (child is SectionPropertiesChange) continue;
                previous.AppendChild(child.CloneNode(true));
            }

            Action wrap = () =>
            {
                var change = new SectionPropertiesChange
                {
                    Author = author,
                    Date = date,
                    Id = idStr,
                };
                change.AppendChild(previous);
                sectPr.AppendChild(change);
            };
            return (stripped, wrap);
        }
    }

    private static bool HasTrackChangeKey(Dictionary<string, string> properties)
    {
        // Match creation keys only. revision.action is action — not creation —
        // and must not trigger this decorator.
        foreach (var k in properties.Keys)
        {
            var lk = k.ToLowerInvariant();
            if (lk is "revision.type" or "revision.author" or "revision.date" or "revision.id")
                return true;
        }
        return false;
    }

    // ========================================================================
    // SECTION 2 — Action dispatchers (selector + native-path accept/reject)
    // ========================================================================

    /// <summary>Single revision marker discovered in document order.</summary>
    private sealed record RevisionRef(
        OpenXmlElement Element,
        string Kind,
        string? Author,
        DateTime? Date,
        string? Id);

    /// <summary>True when <paramref name="properties"/> carries
    /// `revision.action=accept` or `revision.action=reject` — i.e. the
    /// caller is asking for an ACTION on existing revisions, not creating
    /// a new one. Used to route native-path mutations
    /// (`/body/p[N]/r[M]`) to <see cref="SetRevisionByNativePath"/>
    /// instead of the creation-side Set.TrackChange decorator.
    ///
    /// Mixing the action with any creation key (revision.type /
    /// revision.author / revision.date / revision.id) is rejected
    /// up-front in <see cref="SetRevisionByNativePath"/> /
    /// <see cref="SetRevisionsBySelector"/> as ambiguous intent.</summary>
    private static bool IsRevisionActionRequest(Dictionary<string, string> properties)
    {
        return properties.ContainsKey("revision.action");
    }

    /// <summary>Accept/reject every revision marker structurally tied to the
    /// element addressed by <paramref name="path"/>. "Tied" means the marker
    /// is the element itself, lives inside the element's subtree (e.g.
    /// pPrChange inside a paragraph's pPr), or wraps the element (a Run
    /// inside an InsertedRun/DeletedRun/MoveFrom/MoveToRun). Markers on
    /// unrelated descendants — e.g. revisions inside other paragraphs that
    /// happen to be inside the same parent body — are NOT touched; scope
    /// follows the structural relationship, not document order.
    ///
    /// 0 matches → throw. 1+ matches → apply action in reverse document
    /// order so removing earlier siblings doesn't invalidate later refs.
    /// Mixing with revision.* attribution keys → throw (ambiguous intent).</summary>
    internal List<string> SetRevisionByNativePath(string path, Dictionary<string, string> properties)
    {
        Modified = true;
        var unsupported = new List<string>();

        if (!properties.TryGetValue("revision.action", out var action))
            throw new ArgumentException("revision action requires --prop revision.action=accept|reject");
        var act = action.Trim().ToLowerInvariant();
        if (act is not ("accept" or "reject"))
            throw new ArgumentException(
                $"revision.action must be `accept` or `reject` (got `{action}`)");
        // Reject mixing action + any creation key — accept/reject takes
        // no type/author/date/id. If the caller wanted attribution, they
        // meant the creation form (revision.type=ins + revision.author=...).
        foreach (var k in properties.Keys)
        {
            if (string.Equals(k, "revision.action", StringComparison.OrdinalIgnoreCase))
                continue;
            if (k.StartsWith("revision.", StringComparison.OrdinalIgnoreCase))
                throw new ArgumentException(
                    $"revision.action={act} cannot be mixed with `{k}` "
                    + "(creation key); use one form at a time");
        }

        var parts = ParsePath(path);
        var target = NavigateToElement(parts)
            ?? throw new ArgumentException($"Path not found: {path}");

        var matching = EnumerateRevisions()
            .Where(r => RevisionTiedToElement(r.Element, target))
            .ToList();
        if (matching.Count == 0)
            throw new ArgumentException(
                $"{path} has no revision markers; use `query revision` to locate one, "
                + "then `set /revision[N]` or `set /revision[@filter]` to accept/reject");

        matching.Reverse();
        foreach (var rev in matching)
        {
            if (act == "accept") AcceptRevision(rev);
            else RejectRevision(rev);
        }
        _doc.MainDocumentPart?.Document?.Save();
        return unsupported;
    }

    /// <summary>True when <paramref name="marker"/> is structurally tied to
    /// <paramref name="target"/> — same element, descendant of target, or
    /// (for run-wrapping marker types) an ancestor of target.</summary>
    private static bool RevisionTiedToElement(OpenXmlElement marker, OpenXmlElement target)
    {
        if (ReferenceEquals(marker, target)) return true;
        // Marker lives inside target's subtree (pPrChange under paragraph,
        // InsertedRun child of paragraph, tcPrChange under cell, …).
        if (marker.Ancestors().Any(a => ReferenceEquals(a, target))) return true;
        // Marker wraps target (target is the inner Run of an InsertedRun /
        // DeletedRun / MoveFromRun / MoveToRun). Restricted to these four
        // wrapper types so `set /body/p[1] --prop revision.action=accept` doesn't
        // also pick up sectPrChange / tblPrChange at parent scope.
        if (marker is InsertedRun || marker is DeletedRun
            || marker is MoveFromRun || marker is MoveToRun)
        {
            if (target.Ancestors().Any(a => ReferenceEquals(a, marker))) return true;
        }
        return false;
    }

    private static bool IsRevisionSelectorPath(string path)
    {
        if (string.IsNullOrEmpty(path)) return false;
        if (path == "/revision") return true;
        if (Regex.IsMatch(path, @"^/revision\[(?:\d+|@\w+=[^\]]+)\](?:\[@\w+=[^\]]+\])*$"))
            return true;
        return false;
    }

    /// <summary>Set entry for revision selectors. Returns the unsupported-key
    /// list (always empty in the happy path — selector dispatch only consumes
    /// `revision`).</summary>
    internal List<string> SetRevisionsBySelector(string path, Dictionary<string, string> properties)
    {
        Modified = true;
        var unsupported = new List<string>();

        if (!properties.TryGetValue("revision.action", out var action) || string.IsNullOrEmpty(action))
        {
            throw new ArgumentException(
                "revision selector requires --prop revision.action=accept|reject");
        }
        var act = action.Trim().ToLowerInvariant();
        if (act is not ("accept" or "reject"))
        {
            throw new ArgumentException(
                $"revision.action must be `accept` or `reject` (got `{action}`)");
        }
        // Reject mixing action + any creation key on the selector path too.
        foreach (var k in properties.Keys)
        {
            if (string.Equals(k, "revision.action", StringComparison.OrdinalIgnoreCase))
                continue;
            if (k.StartsWith("revision.", StringComparison.OrdinalIgnoreCase))
                throw new ArgumentException(
                    $"revision.action={act} cannot be mixed with `{k}` "
                    + "(creation key); use one form at a time");
        }

        var indexMatch = Regex.Match(path, @"^/revision\[(\d+)\]$");
        Predicate<RevisionRef> filter;
        if (indexMatch.Success)
        {
            var idx = int.Parse(indexMatch.Groups[1].Value);
            var all = EnumerateRevisions();
            if (idx < 1 || idx > all.Count)
                throw new ArgumentException(
                    $"/revision[{idx}] out of range (document has {all.Count} revisions)");
            var picked = all[idx - 1];
            filter = r => ReferenceEquals(r.Element, picked.Element);
        }
        else
        {
            // Parse [@attr=value] segments. Multiple are AND-joined.
            var filterDict = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            foreach (Match m in Regex.Matches(path, @"\[@(\w+)=([^\]]+)\]"))
            {
                filterDict[m.Groups[1].Value] = m.Groups[2].Value.Trim('"', '\'');
            }
            filter = r => MatchesFilter(r, filterDict);
        }

        var targets = EnumerateRevisions().Where(r => filter(r)).ToList();
        // Reverse so removing earlier siblings doesn't invalidate later refs.
        targets.Reverse();
        foreach (var rev in targets)
        {
            if (act == "accept") AcceptRevision(rev);
            else RejectRevision(rev);
        }
        _doc.MainDocumentPart?.Document?.Save();
        return unsupported;
    }

    /// <summary>Build the canonical Path for a revision. Prefers
    /// `/revision[@id={w:id}]` when the marker carries a stable OOXML id
    /// (every revision Word/WPS writes does); falls back to positional
    /// `/revision[{index}]` only when id is missing. Mirrors the
    /// paragraph @paraId= / footnote @footnoteId= / sdt @sdtId= convention
    /// — see CONSISTENCY(id-selectors) in Navigation.cs.</summary>
    private static string BuildRevisionPath(RevisionRef rev, int positionalIndex)
        => !string.IsNullOrEmpty(rev.Id)
            ? $"/revision[@id={rev.Id}]"
            : $"/revision[{positionalIndex}]";

    /// <summary>Build a DocumentNode for one revision. Shared between
    /// `query revision` (enumerates all) and `get /revision[...]` (resolves
    /// one) so the two endpoints emit byte-identical shapes.</summary>
    private DocumentNode BuildRevisionNode(RevisionRef rev, int positionalIndex, string text)
    {
        var node = new DocumentNode
        {
            Path = BuildRevisionPath(rev, positionalIndex),
            Type = "revision",
            Text = text,
        };
        node.Format["revision.type"] = rev.Kind;
        if (!string.IsNullOrEmpty(rev.Author))
            node.Format["revision.author"] = rev.Author!;
        if (rev.Date != null)
            node.Format["revision.date"] = rev.Date.Value.ToString("o");
        if (!string.IsNullOrEmpty(rev.Id))
            node.Format["revision.id"] = rev.Id!;
        var nativePath = ComputeRevisionNativePath(rev.Element);
        if (!string.IsNullOrEmpty(nativePath))
            node.Format["revision.nativePath"] = nativePath;
        return node;
    }

    /// <summary>Get-side resolution for `/revision[N]` and `/revision[@id=X]`.
    /// Returns null when the path is not a revision selector or no marker
    /// matches; throws for in-range-but-invalid forms (e.g. unknown filter
    /// attribute) so Get surfaces the same errors as Set.</summary>
    private DocumentNode? TryGetRevisionNode(string path)
    {
        if (string.IsNullOrEmpty(path) || !path.StartsWith("/revision", StringComparison.Ordinal))
            return null;
        // Bare /revision (no selector) is a Query form, not a single-node Get.
        if (path == "/revision") return null;
        if (!IsRevisionSelectorPath(path)) return null;

        var all = EnumerateRevisions();

        // Positional /revision[N] — 1-based, fallback when no @id available.
        var indexMatch = Regex.Match(path, @"^/revision\[(\d+)\]$");
        if (indexMatch.Success)
        {
            var n = int.Parse(indexMatch.Groups[1].Value);
            if (n < 1 || n > all.Count)
                throw new ArgumentException(
                    $"/revision[{n}] out of range (document has {all.Count} revisions)");
            var rev = all[n - 1];
            return BuildRevisionNode(rev, n, ExtractRevisionText(rev));
        }

        // Filter form: /revision[@attr=value] (+ more). Returns the first
        // match in document order — Get is single-node by contract; callers
        // wanting all matches use `query revision` with the same filter.
        var filterDict = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        foreach (Match m in Regex.Matches(path, @"\[@(\w+)=([^\]]+)\]"))
            filterDict[m.Groups[1].Value] = m.Groups[2].Value.Trim('"', '\'');

        for (int i = 0; i < all.Count; i++)
        {
            if (MatchesFilter(all[i], filterDict))
                return BuildRevisionNode(all[i], i + 1, ExtractRevisionText(all[i]));
        }
        throw new ArgumentException($"no revision matches {path}");
    }

    /// <summary>Walk the document body and emit a RevisionRef for every
    /// tracked-change marker. Order is document order (Descendants is
    /// pre-order).</summary>
    private List<RevisionRef> EnumerateRevisions()
    {
        var list = new List<RevisionRef>();
        var body = _doc.MainDocumentPart?.Document?.Body;
        if (body == null) return list;

        foreach (var ins in body.Descendants<InsertedRun>())
            list.Add(new RevisionRef(ins, "insertion", ins.Author?.Value, ins.Date?.Value, ins.Id?.Value?.ToString()));
        foreach (var del in body.Descendants<DeletedRun>())
            list.Add(new RevisionRef(del, "deletion", del.Author?.Value, del.Date?.Value, del.Id?.Value?.ToString()));
        foreach (var rpc in body.Descendants<RunPropertiesChange>())
            list.Add(new RevisionRef(rpc, "formatChange", rpc.Author?.Value, rpc.Date?.Value, rpc.Id?.Value?.ToString()));
        foreach (var ppc in body.Descendants<ParagraphPropertiesChange>())
            list.Add(new RevisionRef(ppc, "paragraphChange", ppc.Author?.Value, ppc.Date?.Value, ppc.Id?.Value?.ToString()));
        foreach (var spc in body.Descendants<SectionPropertiesChange>())
            list.Add(new RevisionRef(spc, "sectionChange", spc.Author?.Value, spc.Date?.Value, spc.Id?.Value?.ToString()));
        foreach (var tpc in body.Descendants<TablePropertiesChange>())
            list.Add(new RevisionRef(tpc, "tableChange", tpc.Author?.Value, tpc.Date?.Value, tpc.Id?.Value?.ToString()));
        foreach (var trpc in body.Descendants<TableRowPropertiesChange>())
            list.Add(new RevisionRef(trpc, "rowChange", trpc.Author?.Value, trpc.Date?.Value, trpc.Id?.Value?.ToString()));
        foreach (var tcpc in body.Descendants<TableCellPropertiesChange>())
            list.Add(new RevisionRef(tcpc, "cellChange", tcpc.Author?.Value, tcpc.Date?.Value, tcpc.Id?.Value?.ToString()));
        // trPr/ins, trPr/del — row-level insertion/deletion markers
        foreach (var trPr in body.Descendants<TableRowProperties>())
        {
            var trIns = trPr.GetFirstChild<Inserted>();
            if (trIns != null)
                list.Add(new RevisionRef(trIns, "rowInsertion", trIns.Author?.Value, trIns.Date?.Value, trIns.Id?.Value?.ToString()));
            var trDel = trPr.GetFirstChild<Deleted>();
            if (trDel != null)
                list.Add(new RevisionRef(trDel, "rowDeletion", trDel.Author?.Value, trDel.Date?.Value, trDel.Id?.Value?.ToString()));
        }
        // cellIns, cellDel — cell-level insertion/deletion
        foreach (var tcPr in body.Descendants<TableCellProperties>())
        {
            var ci = tcPr.GetFirstChild<CellInsertion>();
            if (ci != null)
                list.Add(new RevisionRef(ci, "cellInsertion", ci.Author?.Value, ci.Date?.Value, ci.Id?.Value?.ToString()));
            var cd = tcPr.GetFirstChild<CellDeletion>();
            if (cd != null)
                list.Add(new RevisionRef(cd, "cellDeletion", cd.Author?.Value, cd.Date?.Value, cd.Id?.Value?.ToString()));
        }
        // paraMarkIns / paraMarkDel — paragraph mark insertion/deletion
        foreach (var pMark in body.Descendants<ParagraphMarkRunProperties>())
        {
            var pIns = pMark.GetFirstChild<Inserted>();
            if (pIns != null)
                list.Add(new RevisionRef(pIns, "paragraphMarkInsertion", pIns.Author?.Value, pIns.Date?.Value, pIns.Id?.Value?.ToString()));
            var pDel = pMark.GetFirstChild<Deleted>();
            if (pDel != null)
                list.Add(new RevisionRef(pDel, "paragraphMarkDeletion", pDel.Author?.Value, pDel.Date?.Value, pDel.Id?.Value?.ToString()));
        }
        // moveFrom / moveTo
        foreach (var mf in body.Descendants<MoveFromRun>())
            list.Add(new RevisionRef(mf, "moveFrom", mf.Author?.Value, mf.Date?.Value, mf.Id?.Value?.ToString()));
        foreach (var mt in body.Descendants<MoveToRun>())
            list.Add(new RevisionRef(mt, "moveTo", mt.Author?.Value, mt.Date?.Value, mt.Id?.Value?.ToString()));
        return list;
    }

    /// <summary>Match a RevisionRef against the parsed [@attr=value] filter.
    /// Recognised attrs: author, date, id, type. Unknown attrs are rejected
    /// to avoid silent passes (typo in --filter loses safety).</summary>
    private static bool MatchesFilter(RevisionRef rev, Dictionary<string, string> filter)
    {
        foreach (var (key, want) in filter)
        {
            var k = key.ToLowerInvariant();
            switch (k)
            {
                case "author":
                    if (!string.Equals(rev.Author, want, StringComparison.Ordinal)) return false;
                    break;
                case "id":
                    if (!string.Equals(rev.Id, want, StringComparison.Ordinal)) return false;
                    break;
                case "type":
                    if (!RevisionTypeMatches(rev.Kind, want)) return false;
                    break;
                case "date":
                    // Exact date match — callers wanting ranges can filter
                    // outside Set via repeated calls. Date kept simple.
                    if (rev.Date == null) return false;
                    if (!DateTime.TryParse(want, out var w)) return false;
                    if (rev.Date.Value != w) return false;
                    break;
                default:
                    throw new ArgumentException(
                        $"unknown revision filter attribute `{key}` (valid: author, type, id, date)");
            }
        }
        return true;
    }

    /// <summary>Lenient type alias matching. `ins` matches insertion;
    /// `del` matches deletion; `format` matches formatChange and the
    /// structural *Change types when explicit; etc.</summary>
    private static bool RevisionTypeMatches(string actualKind, string want)
    {
        var w = want.ToLowerInvariant();
        var a = actualKind.ToLowerInvariant();
        if (a == w) return true;
        return (w, a) switch
        {
            ("ins", "insertion") => true,
            ("insertion", "insertion") => true,
            ("del", "deletion") => true,
            ("deletion", "deletion") => true,
            ("format", "formatchange") => true,
            ("format", "paragraphchange") => true,
            ("format", "sectionchange") => true,
            ("format", "tablechange") => true,
            ("format", "rowchange") => true,
            ("format", "cellchange") => true,
            ("paragraph", "paragraphchange") => true,
            ("rowins", "rowinsertion") => true,
            ("rowdel", "rowdeletion") => true,
            ("cellins", "cellinsertion") => true,
            ("celldel", "celldeletion") => true,
            ("paramarkins", "paragraphmarkinsertion") => true,
            ("paramarkdel", "paragraphmarkdeletion") => true,
            ("move", "movefrom") => true,
            ("move", "moveto") => true,
            _ => false,
        };
    }

    /// <summary>Accept one revision marker. Mirrors the per-kind branches in
    /// <see cref="AcceptAllChanges"/> but applied to a single element so the
    /// selector path can iterate matches with reverse-order safety.</summary>
    private void AcceptRevision(RevisionRef rev)
    {
        switch (rev.Kind)
        {
            case "insertion":
                {
                    var ins = (InsertedRun)rev.Element;
                    var parent = ins.Parent;
                    if (parent != null)
                    {
                        foreach (var child in ins.ChildElements.ToList())
                            parent.InsertBefore(child.CloneNode(true), ins);
                    }
                    ins.Remove();
                    break;
                }
            case "deletion":
                rev.Element.Remove();
                break;
            case "formatChange":
            case "paragraphChange":
            case "sectionChange":
            case "tableChange":
            case "rowChange":
            case "cellChange":
                rev.Element.Remove();
                break;
            case "rowInsertion":
            case "rowDeletion":
                // Accept row insertion: keep row, drop marker. Accept row
                // deletion: remove the entire row (the marker said "the row
                // was deleted", accept = commit the deletion).
                if (rev.Kind == "rowInsertion")
                {
                    rev.Element.Remove();
                }
                else
                {
                    var row = rev.Element.Ancestors<TableRow>().FirstOrDefault();
                    rev.Element.Remove();
                    row?.Remove();
                }
                break;
            case "cellInsertion":
                rev.Element.Remove();
                break;
            case "cellDeletion":
                {
                    var cell = rev.Element.Ancestors<TableCell>().FirstOrDefault();
                    rev.Element.Remove();
                    cell?.Remove();
                    break;
                }
            case "paragraphMarkInsertion":
                rev.Element.Remove();
                break;
            case "paragraphMarkDeletion":
                {
                    // Same merge logic as AcceptAllChanges: ¶ del means the
                    // paragraph break was deleted → join with next paragraph.
                    var pMark = (ParagraphMarkRunProperties)rev.Element.Parent!;
                    var thisPara = pMark.Ancestors<Paragraph>().FirstOrDefault();
                    var nextPara = thisPara?.NextSibling<Paragraph>();
                    rev.Element.Remove();
                    if (thisPara != null && nextPara != null)
                    {
                        var movable = thisPara.ChildElements
                            .Where(c => c is not ParagraphProperties)
                            .ToList();
                        var nextPPr = nextPara.GetFirstChild<ParagraphProperties>();
                        OpenXmlElement? insertAfter = nextPPr;
                        foreach (var ch in movable)
                        {
                            ch.Remove();
                            if (insertAfter == null) nextPara.PrependChild(ch);
                            else { insertAfter.InsertAfterSelf(ch); insertAfter = ch; }
                        }
                        thisPara.Remove();
                    }
                    break;
                }
            case "moveFrom":
                rev.Element.Remove();
                break;
            case "moveTo":
                {
                    var mt = (MoveToRun)rev.Element;
                    var parent = mt.Parent;
                    if (parent != null)
                    {
                        foreach (var child in mt.ChildElements.ToList())
                            parent.InsertBefore(child.CloneNode(true), mt);
                    }
                    mt.Remove();
                    break;
                }
        }
    }

    /// <summary>Reject one revision marker. Mirror of <see cref="AcceptRevision"/>
    /// with inverted semantics — discard inserts, restore deletes, restore
    /// prior pPr/rPr/etc. from each *Change's snapshot.</summary>
    private void RejectRevision(RevisionRef rev)
    {
        switch (rev.Kind)
        {
            case "insertion":
                rev.Element.Remove();
                break;
            case "deletion":
                {
                    var del = (DeletedRun)rev.Element;
                    var parent = del.Parent;
                    if (parent != null)
                    {
                        foreach (var child in del.ChildElements.ToList())
                        {
                            var clone = child.CloneNode(true);
                            foreach (var dt in clone.Descendants<DeletedText>().ToList())
                            {
                                var text = new Text(dt.Text);
                                if (dt.Space != null) text.Space = dt.Space;
                                dt.Parent?.ReplaceChild(text, dt);
                            }
                            parent.InsertBefore(clone, del);
                        }
                    }
                    del.Remove();
                    break;
                }
            case "formatChange":
                RestorePropsFromChange<RunProperties, RunPropertiesChange, PreviousRunProperties>(
                    (RunPropertiesChange)rev.Element, () => new RunProperties());
                break;
            case "paragraphChange":
                RestorePropsFromChange<ParagraphProperties, ParagraphPropertiesChange, ParagraphPropertiesExtended>(
                    (ParagraphPropertiesChange)rev.Element, () => new ParagraphProperties());
                break;
            case "sectionChange":
                RestorePropsFromChange<SectionProperties, SectionPropertiesChange, PreviousSectionProperties>(
                    (SectionPropertiesChange)rev.Element, () => new SectionProperties());
                break;
            case "tableChange":
                RestorePropsFromChange<TableProperties, TablePropertiesChange, PreviousTableProperties>(
                    (TablePropertiesChange)rev.Element, () => new TableProperties());
                break;
            case "cellChange":
                RestorePropsFromChange<TableCellProperties, TableCellPropertiesChange, PreviousTableCellProperties>(
                    (TableCellPropertiesChange)rev.Element, () => new TableCellProperties());
                break;
            case "rowChange":
                RestorePropsFromChange<TableRowProperties, TableRowPropertiesChange, PreviousTableRowProperties>(
                    (TableRowPropertiesChange)rev.Element, () => new TableRowProperties());
                break;
            case "rowInsertion":
                {
                    // Reject row insertion: discard the row entirely.
                    var row = rev.Element.Ancestors<TableRow>().FirstOrDefault();
                    rev.Element.Remove();
                    row?.Remove();
                    break;
                }
            case "rowDeletion":
                // Reject row deletion: keep row, drop marker.
                rev.Element.Remove();
                break;
            case "cellInsertion":
                {
                    var cell = rev.Element.Ancestors<TableCell>().FirstOrDefault();
                    rev.Element.Remove();
                    cell?.Remove();
                    break;
                }
            case "cellDeletion":
                rev.Element.Remove();
                break;
            case "paragraphMarkInsertion":
                {
                    // Reject ¶ ins: merge with previous paragraph.
                    var pMark = (ParagraphMarkRunProperties)rev.Element.Parent!;
                    rev.Element.Remove();
                    var thisPara = pMark.Ancestors<Paragraph>().FirstOrDefault();
                    var prevPara = thisPara?.PreviousSibling<Paragraph>();
                    if (thisPara != null && prevPara != null)
                    {
                        foreach (var ch in thisPara.ChildElements.Where(c => c is not ParagraphProperties).ToList())
                        {
                            ch.Remove();
                            prevPara.AppendChild(ch);
                        }
                        thisPara.Remove();
                    }
                    break;
                }
            case "paragraphMarkDeletion":
                rev.Element.Remove();
                break;
            case "moveFrom":
                {
                    var mf = (MoveFromRun)rev.Element;
                    var parent = mf.Parent;
                    if (parent != null)
                    {
                        foreach (var child in mf.ChildElements.ToList())
                            parent.InsertBefore(child.CloneNode(true), mf);
                    }
                    mf.Remove();
                    break;
                }
            case "moveTo":
                rev.Element.Remove();
                break;
        }
    }

    /// <summary>Pull the most representative human-readable text snippet for
    /// a revision marker. Mirrors what `query revision` shipped historically
    /// for ins/del/rPrChange/pPrChange and extends to the rest of the marker
    /// families. Empty when the marker has no associated text (e.g. structural
    /// section/table property changes that don't pin a specific run).</summary>
    private static string ExtractRevisionText(RevisionRef rev)
    {
        switch (rev.Kind)
        {
            case "insertion":
                return string.Join("", rev.Element.Descendants<Text>().Select(t => t.Text));
            case "deletion":
            case "moveFrom":
                return string.Join("", rev.Element.Descendants<DeletedText>().Select(t => t.Text));
            case "moveTo":
                return string.Join("", rev.Element.Descendants<Text>().Select(t => t.Text));
            case "formatChange":
                {
                    var run = rev.Element.Ancestors<Run>().FirstOrDefault();
                    return run != null
                        ? string.Join("", run.Descendants<Text>().Select(t => t.Text))
                        : "";
                }
            case "paragraphChange":
            case "paragraphMarkInsertion":
            case "paragraphMarkDeletion":
                {
                    var para = rev.Element.Ancestors<Paragraph>().FirstOrDefault();
                    return para != null
                        ? string.Join("", para.Descendants<Text>().Select(t => t.Text))
                        : "";
                }
            case "rowInsertion":
            case "rowDeletion":
            case "rowChange":
                {
                    var row = rev.Element.Ancestors<TableRow>().FirstOrDefault();
                    return row != null
                        ? string.Join(" | ",
                            row.Elements<TableCell>().Select(c =>
                                string.Join("", c.Descendants<Text>().Select(t => t.Text))))
                        : "";
                }
            case "cellInsertion":
            case "cellDeletion":
            case "cellChange":
                {
                    var cell = rev.Element.Ancestors<TableCell>().FirstOrDefault();
                    return cell != null
                        ? string.Join("", cell.Descendants<Text>().Select(t => t.Text))
                        : "";
                }
            case "tableChange":
            case "sectionChange":
                return "";
            default:
                return "";
        }
    }

    /// <summary>Compute the OOXML DOM path of the closest navigable ancestor
    /// of a revision marker. Useful for "where in the document is this
    /// revision" — agents can map a `/revision[N]` synthetic path back to a
    /// real `/body/p[@paraId=…]/r[K]` or `/body/tbl[N]/tr[M]/tc[K]` location
    /// for downstream `get` / cross-referencing.
    ///
    /// Path is best-effort and read-only: callers must NOT use it as a Set
    /// target to drive accept/reject (the legitimate path for that is
    /// `set /revision[N]` or a filtered selector — those go through the
    /// same EnumerateRevisions enumerator and are guaranteed to address
    /// the right marker even when revisions overlap on the same anchor).
    /// Empty when no navigable ancestor exists.</summary>
    private static string ComputeRevisionNativePath(OpenXmlElement marker)
    {
        // Walk up to find the closest "anchor" — a Run / Paragraph / TableCell
        // / TableRow / Table. The path is built from /body downward.
        var anchor = (OpenXmlElement?)marker;
        while (anchor != null
               && anchor is not Run
               && anchor is not Paragraph
               && anchor is not TableCell
               && anchor is not TableRow
               && anchor is not Table
               && anchor is not SectionProperties)
        {
            anchor = anchor.Parent;
        }
        if (anchor == null) return "";

        // Build segments root-down by walking back up.
        var segments = new List<string>();
        var current = anchor;
        while (current != null && current is not Body)
        {
            string? seg = current switch
            {
                Run r => $"r[{IndexOfSiblingsByType<Run>(r)}]",
                Paragraph p when !string.IsNullOrEmpty(p.ParagraphId?.Value)
                    => $"p[@paraId={p.ParagraphId!.Value}]",
                Paragraph p => $"p[{IndexOfSiblingsByType<Paragraph>(p)}]",
                TableCell tc => $"tc[{IndexOfSiblingsByType<TableCell>(tc)}]",
                TableRow tr => $"tr[{IndexOfSiblingsByType<TableRow>(tr)}]",
                Table tbl => $"tbl[{IndexOfSiblingsByType<Table>(tbl)}]",
                SectionProperties => "sectPr",
                _ => null,
            };
            if (seg != null) segments.Insert(0, seg);
            current = current.Parent;
        }
        if (segments.Count == 0) return "";
        return "/body/" + string.Join("/", segments);
    }

    private static int IndexOfSiblingsByType<T>(OpenXmlElement element) where T : OpenXmlElement
    {
        var parent = element.Parent;
        if (parent == null) return 1;
        int idx = 0;
        foreach (var sib in parent.Elements<T>())
        {
            idx++;
            if (ReferenceEquals(sib, element)) return idx;
        }
        return idx;
    }

    /// <summary>Common shape for <c>*PrChange</c> rejection: the change
    /// element's parent is the current pPr/rPr/etc.; the change's first child
    /// is the snapshot of the prior state. Replace the parent with a fresh
    /// element rebuilt from the snapshot's children.</summary>
    private static void RestorePropsFromChange<TParent, TChange, TSnapshot>(
        TChange change,
        Func<TParent> newInstance)
        where TParent : OpenXmlElement
        where TChange : OpenXmlElement
        where TSnapshot : OpenXmlElement
    {
        var parent = change.Parent as TParent;
        if (parent == null) { change.Remove(); return; }
        var snapshot = change.GetFirstChild<TSnapshot>();
        if (snapshot == null) { change.Remove(); return; }
        var grand = parent.Parent;
        if (grand == null) { change.Remove(); return; }
        var rebuilt = newInstance();
        foreach (var ch in snapshot.ChildElements.ToList())
            rebuilt.AppendChild(ch.CloneNode(true));
        grand.ReplaceChild(rebuilt, parent);
    }
}
