// Copyright 2025 OfficeCLI (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using OfficeCli.Core;

namespace OfficeCli.Handlers;

public static partial class PptxBatchEmitter
{
    // CONSISTENCY(emit-shape-mirror): mirrors WordBatchEmitter.Paragraph.cs
    // logic shape — get the node, filter props, decide collapsed-single-run
    // vs multi-run, emit the parent then iterate children. PowerPoint
    // shapes can carry many paragraphs (a slide text body is a list of
    // <a:p> elements), so the collapse heuristic is per-paragraph, not
    // per-shape.

    private static void EmitShape(PowerPointHandler ppt, DocumentNode shapeNode, string parentSlidePath,
                                  string replayPath, List<BatchItem> items, SlideEmitContext ctx)
    {
        // depth=3 so paragraph -> run -> any inline runs all materialize. The
        // single-run collapse heuristic needs the run nodes present to read
        // their text / char-prop bag.
        var fullShape = ppt.Get(shapeNode.Path, depth: 3);
        var shapeProps = FilterEmittableProps(fullShape.Format);

        // NodeBuilder emits `geometry=rect` for every shape with the implicit
        // <a:prstGeom prst="rect"/> body — including plain text boxes. Replay
        // routes any shape carrying `geometry=` through AddShape, which (per
        // bbe1a0c8) seeds a default outline when the caller picks a geometry
        // but supplies no explicit fill/line. The result is a round-trip
        // drift: a clean textbox grows a 1pt black border on every dump+replay.
        // Strip the rect default for textbox/title sources; explicit `shape`
        // types keep the geometry so AddShape sees the user's intent.
        if ((shapeNode.Type == "textbox" || shapeNode.Type == "title")
            && shapeProps.TryGetValue("geometry", out var geomVal)
            && geomVal.Equals("rect", StringComparison.OrdinalIgnoreCase))
        {
            shapeProps.Remove("geometry");
        }

        // Emit type matches Add dispatch: "title" / "equation" both reduce to
        // "shape" or "textbox" on Add, and the emitted shape carries its
        // distinguishing prop (isTitle=true / formula=...). For now use
        // "textbox" for plain text shapes (no geometry) and "shape" otherwise.
        // CONSISTENCY(equation-emit-degrade): AddEquation throws when neither
        // `formula` nor `text` is present. NodeBuilder.ShapeToNode emits
        // Format["formula"] (LaTeX from OMath) when available; if it isn't
        // (exotic OMath that ToLatex can't render), degrade to a plain textbox
        // emit rather than crash replay.
        bool isEquation = shapeNode.Type == "equation" && shapeProps.ContainsKey("formula");
        string emitType = shapeNode.Type switch
        {
            "title" => "shape",
            "equation" => isEquation ? "equation" : (shapeProps.ContainsKey("geometry") ? "shape" : "textbox"),
            _ => shapeProps.ContainsKey("geometry") ? "shape" : "textbox",
        };

        items.Add(new BatchItem
        {
            Command = "add",
            Parent = parentSlidePath,
            Type = emitType,
            Props = shapeProps.Count > 0 ? shapeProps : null,
        });

        // Equation shapes' text body is AlternateContent (a14:m + readable
        // fallback run); the math content is fully captured by `formula`.
        // Emitting paragraphs/runs here would inject the fallback string as
        // user text — skip the body walk for equations entirely.
        if (isEquation) return;

        EmitTextBody(ppt, fullShape, replayPath, items);
    }

    private static void EmitPlaceholder(PowerPointHandler ppt, DocumentNode phNode, string parentSlidePath,
                                        string replayPath, List<BatchItem> items, SlideEmitContext ctx)
    {
        var full = ppt.Get(phNode.Path, depth: 3);
        var props = FilterEmittableProps(full.Format);

        items.Add(new BatchItem
        {
            Command = "add",
            Parent = parentSlidePath,
            Type = "placeholder",
            Props = props.Count > 0 ? props : null,
        });

        EmitTextBody(ppt, full, replayPath, items);
    }

    private static void EmitConnector(PowerPointHandler ppt, DocumentNode cxnNode, string parentSlidePath,
                                      List<BatchItem> items, SlideEmitContext ctx)
    {
        var full = ppt.Get(cxnNode.Path);
        var props = FilterEmittableProps(full.Format);

        items.Add(new BatchItem
        {
            Command = "add",
            Parent = parentSlidePath,
            Type = "connector",
            Props = props.Count > 0 ? props : null,
        });
    }

    private static void EmitGroup(PowerPointHandler ppt, DocumentNode grpNode, string parentSlidePath,
                                  string replayPath, List<BatchItem> items, SlideEmitContext ctx)
    {
        var full = ppt.Get(grpNode.Path);
        var props = FilterEmittableProps(full.Format);

        items.Add(new BatchItem
        {
            Command = "add",
            Parent = parentSlidePath,
            Type = "group",
            Props = props.Count > 0 ? props : null,
        });

        if (full.Children == null) return;

        // Group children resolve through the same dispatch as slide-level
        // children. Replay parent for the group's children is the group's
        // positional path; children get fresh ordinals within the group scope.
        var ord = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
        foreach (var child in full.Children)
        {
            switch (child.Type)
            {
                case "textbox":
                case "title":
                case "shape":
                case "equation":
                    ord["shape"] = ord.GetValueOrDefault("shape", 0) + 1;
                    EmitShape(ppt, child, replayPath, $"{replayPath}/shape[{ord["shape"]}]", items, ctx);
                    break;
                case "connector":
                    ord["connector"] = ord.GetValueOrDefault("connector", 0) + 1;
                    EmitConnector(ppt, child, replayPath, items, ctx);
                    break;
                case "group":
                    ord["group"] = ord.GetValueOrDefault("group", 0) + 1;
                    EmitGroup(ppt, child, replayPath, $"{replayPath}/group[{ord["group"]}]", items, ctx);
                    break;
                case "placeholder":
                    ord["placeholder"] = ord.GetValueOrDefault("placeholder", 0) + 1;
                    EmitPlaceholder(ppt, child, replayPath, $"{replayPath}/placeholder[{ord["placeholder"]}]", items, ctx);
                    break;
                default:
                    ctx.Unsupported.Add(new UnsupportedWarning(
                        Element: child.Type ?? "unknown",
                        SlidePath: replayPath,
                        Reason: "group child type deferred to PR2 / unrecognized"));
                    break;
            }
        }
    }

    // Walk an emitted shape's text body. Each paragraph becomes an `add
    // paragraph` entry under the shape; runs become `add run` children of the
    // paragraph (with text carried as the canonical "text" prop). Single-run
    // paragraphs collapse run props onto the paragraph itself, mirroring the
    // docx single-run optimization.
    private static void EmitTextBody(PowerPointHandler ppt, DocumentNode shapeNode, string shapeParent, List<BatchItem> items)
    {
        if (shapeNode.Children == null) return;
        var paragraphs = shapeNode.Children.Where(c => c.Type == "paragraph" || c.Type == "p").ToList();
        if (paragraphs.Count == 0) return;
        // shapeParent is the positional replay path (e.g. /slide[1]/shape[2]),
        // computed by the caller from per-slide ordinal counters. Replaces
        // the previous shapeNode.Path which carried @id= and broke replay.

        int pIdx = 0;
        foreach (var para in paragraphs)
        {
            pIdx++;
            // PPTX-SPECIFIC(shape-auto-empty-paragraph): AddShape / AddTextbox /
            // AddPlaceholder seed the txBody with one empty <a:p>. If we emit
            // every paragraph as `add`, replay produces an off-by-one empty
            // paragraph[1] that accumulates across round-trips. So the first
            // paragraph under a shape rewrites the seeded one via `set`, and
            // subsequent paragraphs append via `add`. docx body has no
            // equivalent auto-empty seed (AddSection initializes an empty body
            // and AddParagraph appends), so WordBatchEmitter uses pure `add`.
            EmitParagraph(ppt, para, shapeParent, pIdx, items, firstParagraph: pIdx == 1);
        }
    }

    private static void EmitParagraph(PowerPointHandler ppt, DocumentNode paraNode, string shapeParent,
                                      int paraIdx, List<BatchItem> items, bool firstParagraph)
    {
        var props = FilterEmittableProps(paraNode.Format);
        var runs = (paraNode.Children ?? new List<DocumentNode>())
            .Where(c => c.Type == "run" || c.Type == "r").ToList();

        // CONSISTENCY(single-run-collapse): mirrors WordBatchEmitter.Paragraph
        // collapseSingleRun — fold a lone run's text + char props onto the
        // paragraph add so simple cases stay one BatchItem.
        bool collapseSingleRun = runs.Count == 1
            && (paraNode.Children?.Count ?? 0) == 1;

        if (collapseSingleRun)
        {
            var runProps = FilterEmittableProps(runs[0].Format);
            foreach (var (k, v) in runProps)
            {
                if (!props.ContainsKey(k)) props[k] = v;
            }
            if (!string.IsNullOrEmpty(runs[0].Text))
                props["text"] = runs[0].Text!;
            if (firstParagraph)
            {
                items.Add(new BatchItem
                {
                    Command = "set",
                    Path = $"{shapeParent}/paragraph[1]",
                    Props = props.Count > 0 ? props : null,
                });
            }
            else
            {
                items.Add(new BatchItem
                {
                    Command = "add",
                    Parent = shapeParent,
                    Type = "paragraph",
                    Props = props.Count > 0 ? props : null,
                });
            }
            return;
        }

        // Multi-run path: emit the paragraph empty (or with paragraph-level
        // props only) then a run per child. First paragraph rewrites the
        // shape's auto-seeded empty <a:p> via `set`; later paragraphs append.
        string paraParent;
        if (firstParagraph)
        {
            items.Add(new BatchItem
            {
                Command = "set",
                Path = $"{shapeParent}/paragraph[1]",
                Props = props.Count > 0 ? props : null,
            });
            paraParent = $"{shapeParent}/paragraph[1]";
        }
        else
        {
            items.Add(new BatchItem
            {
                Command = "add",
                Parent = shapeParent,
                Type = "paragraph",
                Props = props.Count > 0 ? props : null,
            });
            // Target parent path for runs is the just-emitted paragraph at
            // its known positional index (paraIdx). Earlier code used
            // paragraph[last()] but the resolver doesn't walk into a
            // placeholder's txBody to find paragraphs, so explicit index
            // is the portable form.
            // /body/p[last()].
            paraParent = $"{shapeParent}/paragraph[{paraIdx}]";
        }

        foreach (var run in runs)
        {
            EmitRun(run, paraParent, items);
        }
    }

    // Run-level Format keys that AddRun seeds on every new <a:r> regardless
    // of caller input — emitting them adds nothing but noise on round-trip
    // AND triggers drift when the source had MORE than one default-only run.
    // `lang` is the canonical culprit: AddRun hard-codes Language="en-US",
    // so a paragraph carrying N empty <a:r> elements with only lang=en-US
    // produces N+M runs on every dump→replay (M = newly-seeded defaults on
    // the freshly-added paragraph). Treat the empty/lang-only run as a
    // no-op marker and skip it entirely.
    private static readonly HashSet<string> RunDefaultOnlyKeys = new(StringComparer.OrdinalIgnoreCase)
    {
        "lang", "altLang",
    };

    private static void EmitRun(DocumentNode runNode, string paraParent, List<BatchItem> items)
    {
        var props = FilterEmittableProps(runNode.Format);
        bool hasText = !string.IsNullOrEmpty(runNode.Text);

        // Drop runs that carry no text and only default attributes AddRun
        // would seed anyway. Without this, a deck with N lang-only empty
        // runs accumulates N more on each round-trip — the source's N stay
        // (faithfully re-emitted), and AddRun's hard-coded Language="en-US"
        // seeds a fresh lang on every newly-added <a:r>, so the next dump
        // sees N+M runs per paragraph and drifts by M each cycle.
        if (!hasText && props.Count > 0
            && props.Keys.All(k => RunDefaultOnlyKeys.Contains(k)))
        {
            return;
        }
        // Fully empty <a:r> (no text, no props after filtering) — same
        // logic: AddRun would just seed its defaults, no useful content
        // round-trips.
        if (!hasText && props.Count == 0)
            return;

        if (hasText)
            props["text"] = runNode.Text!;

        items.Add(new BatchItem
        {
            Command = "add",
            Parent = paraParent,
            Type = "run",
            Props = props.Count > 0 ? props : null,
        });
    }
}
