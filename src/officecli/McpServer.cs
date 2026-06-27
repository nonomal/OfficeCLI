// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

using System.CommandLine;
using System.Reflection;
using System.Text;
using System.Text.Json;
using System.Text.Json.Nodes;
using OfficeCli.Core;
using OfficeCli.Handlers;

namespace OfficeCli;

/// <summary>
/// Minimal MCP (Model Context Protocol) server over stdio.
/// Implements JSON-RPC 2.0 with initialize, tools/list, and tools/call.
/// All JSON is hand-written via Utf8JsonWriter to avoid reflection (PublishTrimmed).
/// </summary>
public static class McpServer
{
    public static async Task RunAsync()
    {
        using var reader = new StreamReader(Console.OpenStandardInput());
        using var writer = new StreamWriter(Console.OpenStandardOutput()) { AutoFlush = true };

        // MCP is non-resident by design: every command opens, applies, and
        // eager-saves directly (the inline handlers always called Open(), never
        // TryResident). When a command is dispatched through the shared CLI root
        // (RunCli), the CLI handlers DO call TryResident — and with no resident
        // running that auto-spawns a subprocess, which blocks this single
        // long-lived stdio process. Default the opt-out on so shared-grammar
        // dispatch stays non-resident, matching the legacy MCP behaviour. An
        // explicit user value (e.g. to opt INTO residents) is respected.
        if (string.IsNullOrEmpty(Environment.GetEnvironmentVariable("OFFICECLI_NO_AUTO_RESIDENT")))
            Environment.SetEnvironmentVariable("OFFICECLI_NO_AUTO_RESIDENT", "1");

        // MCP server is a long-lived stdio process. The normal
        // per-invocation auto-upgrade path (Program.cs:112) is
        // short-circuited for `officecli mcp` because CheckInBackground
        // is called AFTER the mcp branch in Program.cs — so without
        // this hook, an MCP instance started once and left running for
        // days/weeks would never see a new release.
        //
        // Run the upgrade path in the background: fire once at startup
        // (applies any pending .update from a previous run and kicks a
        // fresh check if >24h stale), then every hour. The hourly wake
        // is cheap because CheckInBackground is debounced by the same
        // 24h timestamp in ~/.officecli/config.json as the normal CLI
        // path, so 23 of 24 wakes no-op. The actual download / verify /
        // File.Move happens in a spawned subprocess whose stdio is
        // redirected (see UpdateChecker.SpawnRefreshProcess), so
        // nothing it does can corrupt our stdout JSON-RPC stream.
        using var upgradeCts = new CancellationTokenSource();
        var upgradeTask = RunPeriodicUpgradeCheckAsync(upgradeCts.Token);

        try
        {
            while (true)
            {
                var line = await reader.ReadLineAsync();
                if (line == null) break;
                if (string.IsNullOrWhiteSpace(line)) continue;

                JsonElement? id = null;
                try
                {
                    using var doc = JsonDocument.Parse(line);
                    var root = doc.RootElement;
                    // The JSON-RPC root must be an Object (single request). Arrays
                    // are valid JSON-RPC 2.0 batch requests that we don't support;
                    // numbers/strings/bools/nulls are malformed entirely. Guard
                    // here before TryGetProperty, which throws on non-Object.
                    if (root.ValueKind != JsonValueKind.Object)
                    {
                        var msg = root.ValueKind == JsonValueKind.Array
                            ? "Invalid Request: batch requests are not supported"
                            : "Invalid Request: request must be a JSON object";
                        await writer.WriteLineAsync(ErrorJson(null, -32600, msg));
                        continue;
                    }
                    // Parse id BEFORE method so a malformed method ('method': 42)
                    // can still echo the original id back per JSON-RPC 2.0 §5.
                    id = root.TryGetProperty("id", out var idEl) ? idEl.Clone() : null;
                    // method must be a string per spec; non-string is an
                    // Invalid Request (-32600), not an internal error.
                    string? method = null;
                    if (root.TryGetProperty("method", out var m))
                    {
                        if (m.ValueKind != JsonValueKind.String)
                        {
                            await writer.WriteLineAsync(ErrorJson(id, -32600, "Invalid Request: 'method' must be a string"));
                            continue;
                        }
                        method = m.GetString();
                    }

                    var response = method switch
                    {
                        "initialize" => HandleInitialize(id),
                        "notifications/initialized" => null,
                        "tools/list" => HandleToolsList(id),
                        "tools/call" => HandleToolsCall(id, root),
                        "ping" => WriteJson(w => { w.WriteStartObject(); Rpc(w, id); w.WriteStartObject("result"); w.WriteEndObject(); w.WriteEndObject(); }),
                        // CONSISTENCY(mcp-error): truncate caller-supplied value to prevent
                        // response amplification (echo arbitrary-length input back unchanged).
                        _ => id.HasValue ? ErrorJson(id, -32601, $"Method not found: {OfficeCli.Help.SchemaHelpLoader.TruncateForError(method ?? "", 64)}") : null,
                    };

                    if (response != null)
                        await writer.WriteLineAsync(response);
                }
                catch (JsonException)
                {
                    await writer.WriteLineAsync(ErrorJson(null, -32700, "Parse error"));
                }
                catch (Exception ex)
                {
                    await writer.WriteLineAsync(ErrorJson(id, -32603, $"Internal error: {ex.Message}"));
                }
            }
        }
        finally
        {
            upgradeCts.Cancel();
            try { await upgradeTask; } catch { }
        }
    }

    private static async Task RunPeriodicUpgradeCheckAsync(CancellationToken token)
    {
        // Fire once at startup — no matter what state the config is in,
        // this applies any pending .update from a previous run and
        // (if stale) spawns a fresh download. Does not block the main
        // loop: this method runs on a background task.
        try { UpdateChecker.CheckInBackground(); } catch { }

        while (!token.IsCancellationRequested)
        {
            try
            {
                await Task.Delay(TimeSpan.FromHours(1), token);
                UpdateChecker.CheckInBackground();
            }
            catch (OperationCanceledException)
            {
                break;
            }
            catch
            {
                // Never crash the MCP server over an update-check failure.
                // UpdateChecker already swallows exceptions internally, so
                // this is belt-and-braces for any future change that might
                // leak one through.
            }
        }
    }

    // ==================== Handlers ====================

    private static string HandleInitialize(JsonElement? id) => WriteJson(w =>
    {
        w.WriteStartObject();
        Rpc(w, id);
        w.WriteStartObject("result");
        w.WriteString("protocolVersion", "2024-11-05");
        w.WriteStartObject("capabilities");
        w.WriteStartObject("tools"); w.WriteBoolean("listChanged", false); w.WriteEndObject();
        w.WriteEndObject();
        var ver = Assembly.GetExecutingAssembly().GetCustomAttribute<AssemblyInformationalVersionAttribute>()?.InformationalVersion ?? "0.0.0";
        w.WriteStartObject("serverInfo"); w.WriteString("name", "officecli"); w.WriteString("version", ver); w.WriteEndObject();
        w.WriteEndObject();
        w.WriteEndObject();
    });

    private static string HandleToolsList(JsonElement? id) => WriteJson(w =>
    {
        w.WriteStartObject();
        Rpc(w, id);
        w.WriteStartObject("result");
        w.WriteStartArray("tools");
        WriteToolDefinitions(w);
        w.WriteEndArray();
        w.WriteEndObject();
        w.WriteEndObject();
    });

    private static string HandleToolsCall(JsonElement? id, JsonElement root)
    {
        if (!root.TryGetProperty("params", out var p))
            return ErrorJson(id, -32602, "Missing params");
        var name = p.TryGetProperty("name", out var n) ? n.GetString() : null;
        var args = p.TryGetProperty("arguments", out var a) ? a : default;
        if (string.IsNullOrEmpty(name))
            return ErrorJson(id, -32602, "Missing tool name");

        try
        {
            // Unified tool: route by "command" arg; legacy: route by tool name.
            // When the unified tool is invoked with no (or empty) "command", the
            // old code fell back to the tool name and the switch reported
            // "Unknown tool: officecli" — which misleads, since the tool is fine
            // and it is the command field that is missing. Detect that here and
            // name the actual gap, listing the valid commands.
            string toolName;
            if (name == "officecli")
            {
                var commandVal = args.ValueKind == JsonValueKind.Object && args.TryGetProperty("command", out var cmd)
                    ? cmd.GetString() : null;
                if (string.IsNullOrWhiteSpace(commandVal))
                {
                    // Do NOT silently route: an inferred verb can be wrong, and
                    // guessing a mutation (batch) from a stray field would hide a
                    // real mistake. Instead, when the caller supplied a field
                    // that belongs to exactly one command, name it as a strong
                    // suggestion and let the caller re-send with an explicit
                    // 'command'. Falls back to the plain list when there is no
                    // unambiguous signal.
                    var hint = SuggestCommandFromExclusiveField(args);
                    throw new ArgumentException(hint is var (field, guess) && field != null
                        ? $"Missing required 'command' field. You provided '{field}', which only '{guess}' uses — "
                          + $"did you mean command=\"{guess}\"? Otherwise set command to one of: {string.Join(", ", CommandNames)}."
                        : "Missing required 'command' field. Set command to one of: " + string.Join(", ", CommandNames) + ".");
                }
                toolName = commandVal;
            }
            else
            {
                toolName = name;
            }
            var contents = ExecuteToolMulti(toolName, args);
            return WriteJson(w =>
            {
                w.WriteStartObject();
                Rpc(w, id);
                w.WriteStartObject("result");
                w.WriteStartArray("content");
                foreach (var c in contents)
                {
                    w.WriteStartObject();
                    w.WriteString("type", c.Type);
                    if (c.Text != null) w.WriteString("text", c.Text);
                    if (c.Data != null) w.WriteString("data", c.Data);
                    if (c.MimeType != null) w.WriteString("mimeType", c.MimeType);
                    w.WriteEndObject();
                }
                w.WriteEndArray();
                w.WriteBoolean("isError", false);
                w.WriteEndObject();
                w.WriteEndObject();
            });
        }
        catch (Exception ex)
        {
            return WriteJson(w =>
            {
                w.WriteStartObject();
                Rpc(w, id);
                w.WriteStartObject("result");
                w.WriteStartArray("content");
                // A CliException already carrying a {success,data} envelope (the
                // batch judgment path) is emitted verbatim so the per-step
                // results survive; everything else gets the "Error: " prefix.
                var errText = ex is OfficeCli.Core.CliException { Code: "batch_step_failed" }
                    ? ex.Message : $"Error: {ex.Message}";
                w.WriteStartObject(); w.WriteString("type", "text"); w.WriteString("text", errText); w.WriteEndObject();
                w.WriteEndArray();
                w.WriteBoolean("isError", true);
                w.WriteEndObject();
                w.WriteEndObject();
            });
        }
    }

    // ==================== Tool Execution ====================

    /// <summary>
    /// MCP content block. Most tool responses are a single text block; screenshot
    /// returns a text caption + an image block (base64 PNG). Fields not relevant
    /// to a given Type are left null and omitted on serialization.
    /// </summary>
    private sealed record McpContent(string Type, string? Text = null, string? Data = null, string? MimeType = null);

    /// <summary>
    /// Multi-modal wrapper around <see cref="ExecuteTool"/>. Special-cases
    /// view+screenshot (returns text caption + base64 PNG); everything else
    /// gets the legacy single-text path. Lets us add image responses without
    /// touching the ~50 string-returning case branches.
    /// </summary>
    private static IReadOnlyList<McpContent> ExecuteToolMulti(string name, JsonElement args)
    {
        if (name == "view" && args.ValueKind == JsonValueKind.Object
            && args.TryGetProperty("mode", out var m) && m.ValueKind == JsonValueKind.String)
        {
            var mode = m.GetString() ?? "";
            if (mode is "screenshot" or "p")
                return RunScreenshot(args);
        }
        return new[] { new McpContent("text", Text: ExecuteTool(name, args)) };
    }

    /// <summary>
    /// Render the document as HTML, headless-screenshot to PNG, return both a
    /// text caption (with the saved tmp PNG path, for agents with fs access)
    /// and the base64 PNG (for MCP-only agents). Mirrors the CLI's
    /// <c>view &lt;file&gt; screenshot</c> path; same backend probing
    /// (playwright → chrome → firefox) via <see cref="HtmlScreenshot"/>.
    /// </summary>
    private static IReadOnlyList<McpContent> RunScreenshot(JsonElement args)
    {
        string Arg(string key) => args.TryGetProperty(key, out var v) ? v.GetString() ?? "" : "";
        int? ArgIntOpt(string key) => args.TryGetProperty(key, out var v) && v.TryGetInt32(out var i) ? i : null;
        int ArgInt(string key, int def) => ArgIntOpt(key) ?? def;

        var file = Arg("file");
        if (string.IsNullOrEmpty(file)) throw new ArgumentException("file= required for screenshot");
        var start = ArgIntOpt("start");
        var end = ArgIntOpt("end");
        var width = ArgInt("screenshot_width", 1600);
        var height = ArgInt("screenshot_height", 1200);
        var grid = ArgInt("grid", 0);
        var renderMode = (Arg("render") is { Length: > 0 } rm ? rm : "auto").ToLowerInvariant();
        if (renderMode is not ("auto" or "native" or "html"))
            throw new ArgumentException($"Invalid render value: {renderMode}. Valid: auto, native, html");

        using var handler = DocumentHandlerFactory.Open(file);
        string? html = null;
        byte[]? directPng = null;
        if (handler is Handlers.PowerPointHandler ppt)
        {
            var pStart = start ?? 1;
            var pEnd = end ?? pStart;

            // Native-first (mirrors CLI --render auto/native/html): export the
            // slide(s) to PNG with the OS-native engine on Windows; grid is
            // HTML-only. Default export size is the slide's 96-DPI native pixels;
            // a custom width overrides it (aspect-matched height).
            var (nativeW, nativeH) = ppt.GetSlideNativePixels();
            int exportW = nativeW, exportH = nativeH;
            if (!(width == 1600 && height == 1200))
            {
                exportW = width;
                exportH = height == 1200 ? Math.Max(1, (int)Math.Round(width * (double)nativeH / nativeW)) : height;
            }
            // -1 = auto: pick columns from slide count + aspect (mirrors CLI).
            int pptGrid = grid < 0
                ? OfficeCli.Core.HtmlScreenshot.AutoGridColumns((end ?? ppt.GetSlideCount()) - pStart + 1, nativeW, nativeH)
                : grid;
            if (renderMode != "html" && OperatingSystem.IsWindows())
            {
                try
                {
                    if (pptGrid > 0)
                    {
                        const int gap = 12, pad = 12;
                        int cellW = Math.Max(1, (int)Math.Round((width - 2 * pad - (pptGrid - 1) * gap) / (double)pptGrid));
                        int cellH = Math.Max(1, (int)Math.Round(cellW * (double)nativeH / nativeW));
                        directPng = OfficeCli.Core.PowerPointPngBackend.RenderGrid(file, pStart, end ?? ppt.GetSlideCount(), cellW, cellH, pptGrid, gap, pad);
                    }
                    else
                    {
                        directPng = OfficeCli.Core.PowerPointPngBackend.Render(file, pStart, pEnd, exportW, exportH);
                    }
                }
                catch { directPng = null; }
            }
            if (renderMode == "native" && directPng == null)
                throw new ArgumentException("render=native requires Windows with Microsoft PowerPoint installed.");
            if (directPng == null)
            {
                html = ppt.ViewAsHtml(pStart, pEnd, pptGrid, width);
                if (pStart == pEnd && grid == 0)
                {
                    if (width == 1600 && height == 1200) { width = nativeW; height = nativeH; }
                    else if (height == 1200) height = Math.Max(1, (int)Math.Round(width * (double)nativeH / nativeW));
                }
            }
        }
        else if (handler is Handlers.ExcelHandler ex) html = ex.ViewAsHtml();
        else if (handler is Handlers.WordHandler whGrid && grid != 0)
        {
            // Contact-sheet grid (native-first on Windows, HTML fallback; incl.
            // -1 = auto) — mirrors CommandBuilder.View.cs's docx grid branch.
            const int gap = 12, pad = 12, maxDim = 1920, scrollbar = 17;
            var (npW, npH) = whGrid.GetPageNativePixels();
            int pageCount = 1;
            var tmpForCount = Path.Combine(Path.GetTempPath(), $"officecli_gridcount_{Path.GetFileNameWithoutExtension(file)}_{Guid.NewGuid():N}.html");
            try
            {
                File.WriteAllText(tmpForCount, whGrid.ViewAsHtml(null));
                pageCount = OfficeCli.Core.HtmlScreenshot.GetPageCountFromDom(tmpForCount) ?? 1;
            }
            catch { /* fall back to 1 row */ }
            finally { try { File.Delete(tmpForCount); } catch { /* ignore */ } }

            int cols = grid < 0 ? OfficeCli.Core.HtmlScreenshot.AutoGridColumns(pageCount, npW, npH) : grid;
            int rows = Math.Max(1, (pageCount + cols - 1) / cols);
            double vpW = width;
            double cellW = Math.Max(1.0, (vpW - scrollbar - pad * 2.0 - (cols - 1) * gap) / cols);
            double cellH = cellW * npH / npW;
            double vpH = pad * 2 + rows * cellH + (rows - 1) * gap;
            double over = Math.Max(vpW, vpH) / maxDim;
            if (over > 1.0) { vpW /= over; cellW /= over; cellH /= over; vpH /= over; }

            // Native-first (the read-only MCP handler coexists with Word's open).
            if (renderMode != "html" && OperatingSystem.IsWindows())
            {
                try { directPng = OfficeCli.Core.WordPdfBackend.RenderGrid(file, $"1-{pageCount}", (int)Math.Round(cellW), (int)Math.Round(cellH), cols, gap, pad); }
                catch { directPng = null; }
            }
            if (renderMode == "native" && directPng == null)
                throw new ArgumentException("render=native requires Windows with Microsoft Word installed.");
            if (directPng == null)
            {
                html = whGrid.ViewAsHtml(null, cols, (int)Math.Round(cellW));
                width = Math.Max(1, (int)Math.Round(vpW));
                height = Math.Max(1, (int)Math.Ceiling(vpH));
            }
        }
        else if (handler is Handlers.WordHandler wh)
        {
            // CONSISTENCY(screenshot-default-first-page): mirror CLI — screenshot
            // mode defaults to page 1 for docx so multi-page docs aren't silently
            // cropped by the viewport. Caller can pass start=N to override.
            var pageFilter = (start ?? 1).ToString();
            if (end is int e && e >= (start ?? 1)) pageFilter = $"{start ?? 1}-{e}";
            if (renderMode != "html" && OperatingSystem.IsWindows())
            {
                try { directPng = OfficeCli.Core.WordPdfBackend.Render(file, pageFilter); } catch { directPng = null; }
            }
            if (renderMode == "native" && directPng == null)
                throw new ArgumentException("render=native requires Windows with Microsoft Word installed.");
            if (directPng == null) html = wh.ViewAsHtml(pageFilter);
        }

        if (html == null && directPng == null)
            throw new ArgumentException("Screenshot mode is only supported for .pptx, .xlsx, and .docx files.");

        var stem = Path.GetFileNameWithoutExtension(file);
        var pngPath = Path.Combine(Path.GetTempPath(), $"officecli_screenshot_{stem}_{Guid.NewGuid():N}.png");
        string backendName;
        if (directPng != null)
        {
            File.WriteAllBytes(pngPath, directPng);
            backendName = handler is Handlers.PowerPointHandler ? "powerpoint" : "word";
        }
        else
        {
            var tmpHtml = Path.Combine(Path.GetTempPath(), $"officecli_preview_{stem}_{Guid.NewGuid():N}.html");
            File.WriteAllText(tmpHtml, html!);
            var r = OfficeCli.Core.HtmlScreenshot.Capture(tmpHtml, pngPath, width, height);
            try { File.Delete(tmpHtml); } catch { /* ignore */ }
            if (!r.Ok)
                throw new InvalidOperationException(
                    "No headless browser available. Install Chrome/Edge/Chromium or Firefox, "
                    + "or `pip install playwright && playwright install chromium`."
                    + (r.Error != null ? $" Last error: {r.Error}" : ""));
            backendName = r.Backend;
        }

        var bytes = File.ReadAllBytes(pngPath);
        var b64 = Convert.ToBase64String(bytes);
        string pagesNote = "";
        if (handler is Handlers.PowerPointHandler pptp)
            pagesNote = $" Slides: {pptp.GetSlideCount()}.";
        var caption = $"Screenshot saved to {pngPath} ({bytes.Length} bytes, backend: {backendName}).{pagesNote}";
        return new[]
        {
            new McpContent("text", Text: caption),
            new McpContent("image", Data: b64, MimeType: "image/png"),
        };
    }

    // ====================================================================
    // Shared-grammar dispatch (Phase 1 of routing MCP through the CLI's one
    // System.CommandLine root). Translating the MCP JSON into the CLI token
    // vector and parsing it with the SAME root the CLI uses means argument
    // validation, business logic, and the {success,data} envelope are shared
    // by construction — not re-marshalled (and re-bugged) by hand here.
    // ====================================================================
    private static RootCommand? _rootCommand;
    private static RootCommand RootCommand => _rootCommand ??= CommandBuilder.BuildRootCommand();

    private readonly record struct CliResult(int Exit, string Stdout, string Stderr);

    /// <summary>
    /// Parse+invoke argv through the shared CLI root, capturing stdout AND
    /// stderr and the exit code. Parse/validation failures (the free win — same
    /// messages the CLI gives) throw before any handler runs. argv is the CLI
    /// token vector, e.g. ["get", file, "/body", "--depth", "1", "--json"].
    /// </summary>
    private static CliResult RunCliRaw(string[] argv)
    {
        var pr = RootCommand.Parse(argv);
        if (pr.Errors.Count > 0)
            throw new ArgumentException(string.Join("; ", pr.Errors.Select(e => e.Message)));
        var prevOut = Console.Out;
        var prevErr = Console.Error;
        var so = new System.IO.StringWriter();
        var se = new System.IO.StringWriter();
        int exit;
        try { Console.SetOut(so); Console.SetError(se); exit = pr.Invoke(); }
        finally { Console.SetOut(prevOut); Console.SetError(prevErr); }
        return new CliResult(exit, so.ToString(), se.ToString());
    }

    /// <summary>
    /// Run a --json CLI command and return its `data` payload, unwrapped from
    /// the {success,data} envelope so the MCP response keeps its existing bare
    /// shape. A success:false envelope throws → MCP isError. (--json commands
    /// emit the error envelope on stdout, so exit code is not consulted here.)
    /// </summary>
    private static string RunCliJsonData(params string[] argv)
    {
        var r = RunCliRaw(argv);
        var outText = r.Stdout.Trim();
        JsonNode? env;
        try { env = JsonNode.Parse(outText); }
        catch
        {
            if (r.Exit != 0) throw new ArgumentException(StripErrPrefix(FirstNonEmpty(r.Stderr, outText)));
            return outText;
        }
        if (env is JsonObject obj)
        {
            if (obj.TryGetPropertyValue("success", out var s) && s is not null && !s.GetValue<bool>())
                throw new ArgumentException(ExtractEnvelopeError(obj));
            if (obj.TryGetPropertyValue("data", out var data) && data is not null)
                return data.ToJsonString(OutputFormatter.PublicJsonOptions);
        }
        return outText;
    }

    /// <summary>
    /// Run a text-output CLI command (raw XML, "Created ...", validate report).
    /// stdout is the result; a non-zero exit throws the stderr message — the
    /// text-mode handlers write "Error: ..." to stderr and signal via exit code
    /// rather than a JSON envelope.
    /// </summary>
    private static string RunCliText(params string[] argv)
    {
        var r = RunCliRaw(argv);
        if (r.Exit != 0)
            throw new ArgumentException(StripErrPrefix(FirstNonEmpty(r.Stderr.Trim(), r.Stdout.Trim())));
        return r.Stdout.TrimEnd('\n', '\r');
    }

    private static string FirstNonEmpty(params string[] xs) =>
        xs.FirstOrDefault(x => !string.IsNullOrWhiteSpace(x)) ?? "Command failed.";

    // The MCP catch block prepends "Error: "; the text-mode handlers already
    // wrote "Error: ..." to stderr. Strip one leading prefix to avoid doubling.
    private static string StripErrPrefix(string s) =>
        s.StartsWith("Error: ", StringComparison.Ordinal) ? s["Error: ".Length..] : s;

    /// <summary>
    /// Pull the human-readable message out of a failure envelope. Two shapes
    /// exist: WrapEnvelopeError → {message:"..."} and WrapErrorEnvelope →
    /// {error:{error:"...", code:"..."}}. Return the inner string, never the
    /// whole object.
    /// </summary>
    private static string ExtractEnvelopeError(JsonObject obj)
    {
        if (obj.TryGetPropertyValue("message", out var m) && m is JsonValue) return m.ToString();
        if (obj.TryGetPropertyValue("error", out var er))
        {
            if (er is JsonObject eo && eo.TryGetPropertyValue("error", out var inner) && inner is not null)
                return inner.ToString();
            if (er is JsonValue) return er.ToString();
        }
        return "Command failed.";
    }

    private static string ExecuteTool(string name, JsonElement args)
    {
        string Arg(string key) => args.ValueKind == JsonValueKind.Object && args.TryGetProperty(key, out var v) ? v.GetString() ?? "" : "";
        int ArgInt(string key, int def) => args.ValueKind == JsonValueKind.Object && args.TryGetProperty(key, out var v) && v.TryGetInt32(out var i) ? i : def;
        int? ArgIntOpt(string key) => args.ValueKind == JsonValueKind.Object && args.TryGetProperty(key, out var v) && v.TryGetInt32(out var i) ? i : null;
        string[] ArgStringArray(string key)
        {
            if (args.ValueKind != JsonValueKind.Object || !args.TryGetProperty(key, out var v) || v.ValueKind != JsonValueKind.Array) return [];
            return v.EnumerateArray().Select(e => e.GetString() ?? "").ToArray();
        }

        switch (name)
        {
            case "create":
            {
                // Routed through the shared root: gains file-extension/required
                // validation and the --type/--force/--locale/--minimal flags the
                // inline path silently ignored.
                var argv = new List<string> { "create", Arg("file") };
                var ctype = Arg("type"); if (!string.IsNullOrEmpty(ctype)) { argv.Add("--type"); argv.Add(ctype); }
                var locale = Arg("locale"); if (!string.IsNullOrEmpty(locale)) { argv.Add("--locale"); argv.Add(locale); }
                if (string.Equals(Arg("force"), "true", StringComparison.OrdinalIgnoreCase)) argv.Add("--force");
                if (string.Equals(Arg("minimal"), "true", StringComparison.OrdinalIgnoreCase)) argv.Add("--minimal");
                return RunCliText(argv.ToArray());
            }
            case "view":
            {
                // screenshot/pdf are intercepted multimodally before this switch
                // (ExecuteToolMulti). Every other mode is routed through the
                // shared CLI root, which gains the windowing/filter options the
                // inline path dropped: issues --type (subtype filter), --limit,
                // --cols, --page, --page-count.
                var mode = Arg("mode");
                var ml = mode.ToLowerInvariant();
                var argv = new List<string> { "view", Arg("file"), mode };
                void Add(string flag, string? val) { if (!string.IsNullOrEmpty(val)) { argv.Add(flag); argv.Add(val); } }
                var start = ArgIntOpt("start"); if (start.HasValue) Add("--start", start.Value.ToString());
                var end = ArgIntOpt("end"); if (end.HasValue) Add("--end", end.Value.ToString());
                var maxLines = ArgIntOpt("max_lines"); if (maxLines.HasValue) Add("--max-lines", maxLines.Value.ToString());
                Add("--type", Arg("type"));   // issues subtype filter
                Add("--cols", Arg("cols"));
                Add("--page", Arg("page"));
                var limit = ArgIntOpt("limit"); if (limit.HasValue) Add("--limit", limit.Value.ToString());
                if (string.Equals(Arg("page_count"), "true", StringComparison.OrdinalIgnoreCase)) argv.Add("--page-count");

                // issues/forms returned bare JSON from the inline path; ask the
                // CLI for --json and unwrap the envelope to keep that shape. All
                // other modes are plain text/HTML/SVG on stdout.
                if (ml is "issues" or "i" or "forms" or "f")
                {
                    argv.Add("--json");
                    return RunCliJsonData(argv.ToArray());
                }
                return RunCliText(argv.ToArray());
            }
            case "get":
            {
                // Phase 1: routed through the shared CLI root. argv mirrors
                // `officecli get <file> <path> --depth N [--save dest] --json`.
                // Validation (required file/path), the resident hop, binary
                // extraction, and the {success,data} envelope all come from the
                // one CLI implementation; we unwrap `data` to keep the MCP
                // response's existing bare {matches,results} shape.
                var path = Arg("path"); if (string.IsNullOrEmpty(path)) path = "/";
                var argv = new List<string> { "get", Arg("file"), path, "--depth", ArgInt("depth", 1).ToString(), "--json" };
                var savePath = Arg("save");
                if (!string.IsNullOrEmpty(savePath)) { argv.Add("--save"); argv.Add(savePath); }
                return RunCliJsonData(argv.ToArray());
            }
            case "query":
            {
                // Routed through the shared root. argv mirrors
                // `officecli query <file> <selector> [--find text] --json`; the
                // Excel cell-alias resolution, r"regex" text filter, and
                // selector warnings all live in the one CLI implementation. We
                // unwrap `data` to keep the bare {matches,results} shape.
                var argv = new List<string> { "query", Arg("file"), Arg("selector"), "--json" };
                var textFilter = Arg("text");
                if (!string.IsNullOrEmpty(textFilter)) { argv.Add("--find"); argv.Add(textFilter); }
                return RunCliJsonData(argv.ToArray());
            }
            case "set":
                return ExecuteMutation(Arg("file"), new BatchItem
                {
                    Command = "set",
                    Path = Arg("path"),
                    Props = ParseProps(ArgStringArray("props")),
                });
            case "add":
            {
                var after = Arg("after");
                var before = Arg("before");
                var fromArg = Arg("from");
                return ExecuteMutation(Arg("file"), new BatchItem
                {
                    Command = "add",
                    Parent = Arg("parent"),
                    Type = Arg("type"),
                    From = string.IsNullOrEmpty(fromArg) ? null : fromArg,
                    Index = ArgIntOpt("index"),
                    After = string.IsNullOrEmpty(after) ? null : after,
                    Before = string.IsNullOrEmpty(before) ? null : before,
                    Props = ParseProps(ArgStringArray("props")),
                });
            }
            case "remove":
                return ExecuteMutation(Arg("file"), new BatchItem
                {
                    Command = "remove",
                    Path = Arg("path"),
                    Props = ParseProps(ArgStringArray("props")),
                });
            case "move":
            {
                var to = Arg("to");
                var mvAfter = Arg("after");
                var mvBefore = Arg("before");
                return ExecuteMutation(Arg("file"), new BatchItem
                {
                    Command = "move",
                    Path = Arg("path"),
                    To = string.IsNullOrEmpty(to) ? null : to,
                    Index = ArgIntOpt("index"),
                    After = string.IsNullOrEmpty(mvAfter) ? null : mvAfter,
                    Before = string.IsNullOrEmpty(mvBefore) ? null : mvBefore,
                    Props = ParseProps(ArgStringArray("props")),
                });
            }
            case "validate":
                // Routed through the shared root (text report on stdout).
                return RunCliText("validate", Arg("file"));
            case "batch":
            {
                var file = Arg("file");
                var commands = Arg("commands");
                var forceStr = Arg("force");
                var stopOnError = !string.Equals(forceStr, "true", StringComparison.OrdinalIgnoreCase);
                // Validate the commands payload before deserializing: an empty or
                // missing 'commands' otherwise surfaces the raw System.Text.Json
                // "input does not contain any JSON tokens" exception, which gives
                // a caller no idea what to fix. The common mistake is putting the
                // array under the wrong key (e.g. 'batch') so 'commands' is empty.
                if (string.IsNullOrWhiteSpace(commands))
                    throw new ArgumentException(
                        "batch requires a 'commands' field: a JSON array (as a string) of command objects, "
                        + "e.g. commands=\"[{\\\"command\\\":\\\"add\\\",\\\"parent\\\":\\\"/body\\\",\\\"type\\\":\\\"paragraph\\\",\\\"props\\\":[\\\"text=Hi\\\"]}]\". "
                        + "Put the array under 'commands' (not under 'batch' or any other key).");
                List<BatchItem>? items;
                try { items = JsonSerializer.Deserialize<List<BatchItem>>(commands, BatchJsonContext.Default.ListBatchItem); }
                catch (JsonException jx)
                {
                    throw new ArgumentException($"'commands' is not a valid JSON array: {jx.Message} "
                        + "It must look like [{\"command\":\"add\", \"parent\":\"/body\", \"type\":\"paragraph\"}].");
                }
                if (items == null || items.Count == 0)
                    throw new ArgumentException("'commands' is an empty array — provide at least one command object.");
                using var handler = DocumentHandlerFactory.Open(file, editable: true);
                // Protection gate against the just-opened in-memory DOM, mirroring
                // the CLI and resident batch paths: a protected .docx rejects
                // batch mutations unless force=true. Surfaced as a thrown
                // CliException → MCP isError, the same way other command errors
                // propagate here.
                var mcpForce = string.Equals(forceStr, "true", StringComparison.OrdinalIgnoreCase);
                if (!mcpForce && file.EndsWith(".docx", StringComparison.OrdinalIgnoreCase))
                {
                    var protBlock = CommandBuilder.GetBatchProtectionBlock(handler, items);
                    if (protBlock != null) throw new CliException(protBlock) { Code = "document_protected" };
                }
                // DeferSave + replay loop, shared with the non-resident CLI batch
                // path (CommandBuilder.RunNonResidentBatch). The using-Dispose
                // performs the single FinalizeDeferredIds + Save flush.
                var results = CommandBuilder.RunNonResidentBatch(handler, items, stopOnError, json: true);
                var sw = new System.IO.StringWriter();
                CommandBuilder.PrintBatchResults(results, json: true, totalCount: items.Count, output: sw);
                // Wrap in the {success,data} envelope like the CLI/resident batch
                // paths (batch is a Judgment command: any failed step → success
                // false). Without this the MCP batch returned a bare
                // {results,summary} with no verdict, so a fully-failed batch was
                // indistinguishable from a clean one. On failure throw the
                // enveloped JSON so isError mirrors the CLI's exit-code-1.
                var batchSuccess = results.Count == 0 || !results.Any(r => !r.Success);
                var envelope = OfficeCli.Core.OutputFormatter.WrapEnvelope(sw.ToString().TrimEnd('\n', '\r'), success: batchSuccess);
                if (!batchSuccess)
                    throw new CliException(envelope) { Code = "batch_step_failed" };
                return envelope;
            }
            case "swap":
                return ExecuteMutation(Arg("file"), new BatchItem
                {
                    Command = "swap",
                    Path = Arg("path"),
                    Path2 = Arg("path2"),
                });
            case "raw":
            {
                // Routed through the shared root: gains the --start/--end/--cols
                // Excel windowing the inline path dropped (it always passed
                // null,null,null), so a huge sheet no longer dumps in full.
                var part = Arg("part"); if (string.IsNullOrEmpty(part)) part = "/document";
                var argv = new List<string> { "raw", Arg("file"), part };
                var start = ArgIntOpt("start"); if (start.HasValue) { argv.Add("--start"); argv.Add(start.Value.ToString()); }
                var end = ArgIntOpt("end"); if (end.HasValue) { argv.Add("--end"); argv.Add(end.Value.ToString()); }
                var cols = Arg("cols"); if (!string.IsNullOrEmpty(cols)) { argv.Add("--cols"); argv.Add(cols); }
                return RunCliText(argv.ToArray());
            }
            case "help":
            {
                // Schema-driven help — single source of truth shared with the CLI's
                // `officecli help` command. The previous implementation was ~150 lines
                // of hardcoded markdown cheat sheets that drifted from schemas/help/*.json
                // (e.g. when chart aliases were added, this block was never updated).
                //
                // Shape (mirrors `officecli help <format> [<element>]`):
                //   {command:"help"}                          → list formats
                //   {command:"help", format:"docx"}           → list elements in that format
                //   {command:"help", format:"docx", type:"paragraph"} → full element schema
                //
                // The Strategy preamble is MCP-specific guidance that schemas don't (and
                // shouldn't) encode — kept inline as McpHelpStrategy.
                var format = Arg("format").ToLowerInvariant();
                var element = Arg("type"); // optional element to drill into

                if (string.IsNullOrEmpty(format))
                    return McpHelpStrategy
                        + "Supported formats: docx, xlsx, pptx.\n"
                        + "Call again with format=<docx|xlsx|pptx> to list elements; "
                        + "add type=<element> for full schema (properties, aliases, examples).";

                if (!OfficeCli.Help.SchemaHelpLoader.IsKnownFormat(format))
                {
                    // CONSISTENCY(mcp-error): truncate user-supplied value in error messages to prevent
                    // response amplification (caller echoes arbitrary-length input back unchanged).
                    var displayFormat = OfficeCli.Help.SchemaHelpLoader.TruncateForError(format, 64);
                    return $"Unknown format '{displayFormat}'. Supported: docx, xlsx, pptx.";
                }

                var canonical = OfficeCli.Help.SchemaHelpLoader.NormalizeFormat(format);
                var sb = new StringBuilder(McpHelpStrategy);

                if (string.IsNullOrEmpty(element))
                {
                    sb.Append("# ").Append(canonical.ToUpperInvariant()).AppendLine(" Elements");
                    sb.AppendLine();
                    foreach (var el in OfficeCli.Help.SchemaHelpLoader.ListElements(canonical))
                        sb.Append("- ").AppendLine(el);
                    sb.AppendLine();
                    var sampleElement = canonical switch { "docx" => "paragraph", "xlsx" => "cell", _ => "shape" };
                    sb.Append("Call again with type=<element> for the full schema. ");
                    sb.Append("Example: {\"command\":\"help\",\"format\":\"").Append(canonical)
                      .Append("\",\"type\":\"").Append(sampleElement).AppendLine("\"}");
                    return sb.ToString();
                }

                try
                {
                    using var doc = OfficeCli.Help.SchemaHelpLoader.LoadSchema(canonical, element);
                    sb.Append(OfficeCli.Help.SchemaHelpRenderer.RenderHuman(doc, null));
                    return sb.ToString();
                }
                catch (Exception ex)
                {
                    return $"{ex.Message}\n\nList available elements via: {{\"command\":\"help\",\"format\":\"{canonical}\"}}";
                }
            }
            case "load_skill":
            {
                // Return the embedded SKILL.md content for the named skill. Pure
                // read — no install side-effect. Identical semantics to the CLI
                // `officecli load_skill <name>` command (both share LoadSkillContent).
                // Agents that want disk-resident skills run `officecli skills install`
                // themselves.
                var skill = Arg("name");
                // path= fetches one bundled reference file (e.g.
                // reference/decision-rules.md) named in the manifest that
                // load_skill (no path) appends. Without it the SKILL.md body's
                // reference pointers are dead links over MCP.
                var skillPath = Arg("path");
                // No name → return the skill catalog so an agent can discover
                // which skill applies (each entry carries its Use-when / Trigger
                // routing guidance). name= is required once a path= is given.
                if (string.IsNullOrEmpty(skill))
                {
                    if (!string.IsNullOrEmpty(skillPath))
                        throw new ArgumentException("name= required when path= is given.");
                    return OfficeCli.Core.SkillInstaller.BuildSkillCatalog();
                }
                try
                {
                    return string.IsNullOrEmpty(skillPath)
                        ? OfficeCli.Core.SkillInstaller.LoadSkillContent(skill)
                        : OfficeCli.Core.SkillInstaller.LoadSkillFile(skill, skillPath);
                }
                catch (ArgumentException ex)
                {
                    // CONSISTENCY(mcp-error): error message already includes the
                    // truncated input via SkillInstaller; re-throw as-is so MCP
                    // returns a structured error to the caller.
                    throw new ArgumentException(ex.Message);
                }
            }
            default:
                // CONSISTENCY(mcp-error): truncate caller-supplied value to prevent
                // response amplification (echo arbitrary-length input back unchanged).
                throw new ArgumentException($"Unknown tool: {OfficeCli.Help.SchemaHelpLoader.TruncateForError(name, 64)}");
        }
    }

    // Single-command mutations (set/add/remove/move/swap) build a BatchItem and
    // run it through the one shared executor (CommandBuilder.ExecuteBatchItem)
    // instead of re-implementing the handler dispatch here. This is what keeps
    // the MCP surface from drifting from batch/CLI on argument shape or verdict
    // semantics — the props-array, swap-path2 and remove-shift bugs were all
    // this surface re-deriving the contract and getting it slightly different.
    // The executor also applies the richer verdicts (unsupported_property,
    // empty-props rejection, NUL-byte guard, Word add-prop warnings) the
    // hand-written cases lacked. json:false → human-readable text, matching the
    // strings the old cases returned.
    private static string ExecuteMutation(string file, BatchItem item)
    {
        using var handler = DocumentHandlerFactory.Open(file, editable: true);
        return CommandBuilder.ExecuteBatchItem(handler, item, json: false);
    }

    private static Dictionary<string, string> ParseProps(string[] propStrs)
    {
        var props = new Dictionary<string, string>();
        foreach (var p in propStrs)
        {
            var eq = p.IndexOf('=');
            if (eq > 0) props[p[..eq]] = p[(eq + 1)..];
        }
        return props;
    }

    // ==================== Tool Definitions ====================

    // Single source of truth for the unified-tool command set: the inputSchema
    // enum AND the missing-command guard both read this, so they cannot drift.
    private static readonly string[] CommandNames =
        { "create", "view", "get", "query", "set", "add", "remove", "move", "swap", "validate", "batch", "raw", "help", "load_skill" };

    // Fields used by exactly one command. When 'command' is omitted, the
    // presence of one of these uniquely identifies the intended verb, so we can
    // route without erroring. Shared (path/parent/props/file/...) fields are
    // deliberately excluded — they are ambiguous across commands.
    private static readonly (string Field, string Command)[] ExclusiveFields =
        { ("commands", "batch"), ("selector", "query"), ("part", "raw"), ("path2", "swap") };

    // Returns the (field, command) pair when the args carry exactly such a
    // command-exclusive field, so the missing-command error can name a precise
    // suggestion. This only *suggests* — it never routes.
    private static (string? Field, string? Command) SuggestCommandFromExclusiveField(JsonElement args)
    {
        if (args.ValueKind != JsonValueKind.Object) return (null, null);
        foreach (var (field, command) in ExclusiveFields)
        {
            if (!args.TryGetProperty(field, out var v)) continue;
            var present = v.ValueKind switch
            {
                JsonValueKind.String => !string.IsNullOrEmpty(v.GetString()),
                JsonValueKind.Array => v.GetArrayLength() > 0,
                JsonValueKind.Null or JsonValueKind.Undefined => false,
                _ => true,
            };
            if (present) return (field, command);
        }
        return (null, null);
    }

    // MCP-specific guidance prepended to every help response. Cannot be derived
    // from schemas/help/*.json — it's about how to use the *tool*, not what the
    // *document model* exposes.
    private const string McpHelpStrategy = @"## Strategy
Use view (outline/stats/issues/annotated) to understand the document first, then get/query to inspect details, then set/add/remove to modify.
View modes: text, annotated, outline, stats, issues, html, svg (pptx only), screenshot, forms (docx only).
Before delivering, pass the delivery gate (see the tool description): validate clean, view issues clean, then a visual audit via view mode=screenshot when layout matters (slide decks most of all). Whether the visual audit is mandatory is format-specific — load_skill name=<pptx|word|excel> carries the authoritative per-format gate.
For 3+ mutations on the same file, use batch (one open/save cycle) instead of separate calls.
Get output keys can be used directly as Set input keys (round-trip safe).
Colors: FF0000, red, rgb(255,0,0), accent1. Sizes: 24pt. Positions: 2cm, 1in, 72pt, or raw EMU.
Paths are 1-based: /slide[1]/shape[2], /body/p[3], /Sheet1/A1.

";

    private const string ToolDescription = @"Create, read, and modify Office documents (.docx, .xlsx, .pptx).

Commands: create (file), view (file, mode: text|annotated|outline|stats|issues|html|svg|screenshot|forms), get (file, path, depth), query (file, selector), set (file, path, props[]), add (file, parent, type, props[], index/after/before), remove (file, path), move (file, path, to, index/after/before), swap (file, path, path2), validate (file), batch (file, commands), raw (file, part), help (format: docx|xlsx|pptx, optional type=<element> for full schema), load_skill (no name: lists all skills with the triggers that say when to use each; name=<pptx|word|excel|word-form|morph-ppt|morph-ppt-3d|pitch-deck|academic-paper|data-dashboard|financial-model>: returns that skill's SKILL.md + a manifest of its reference files; add path=<relpath> to fetch one reference file, e.g. path=reference/decision-rules.md).

Paths are 1-based: /slide[1]/shape[2], /body/p[3], /Sheet1/A1. Props are key=value strings. Call help with format= to list elements, then help with format= and type= to drill into a specific element's schema (properties, aliases, examples).

Delivery gate (before reporting a document finished — any failure = fix and re-check, do NOT deliver; validate passing is NOT delivery, 'looks like a real document' is):
1. Schema: validate -> clean, no errors.
2. Content: view mode=issues -> no overflow/format/structure issues; and scan view mode=text for leftover placeholders (xxxx, lorem/ipsum, <TODO>, {{...}}, $VAR$, empty ()/[]).
3. Visual audit via view mode=screenshot — renders the page/slide and returns it as an image shown to you (--page N per slide/page, or --grid auto for a contact sheet). Judge it adversarially (assume problems exist) for overlap, text overflow, off-slide shapes, dark-on-dark, misalignment; fix positions/sizes (set x/y/width/height) and re-screenshot until right; with no headless browser, say 'not visually verified'. Whether this audit is mandatory is format-specific (slide decks need it most — absolute-positioned shapes overlap invisibly to text modes), so load the format's skill for the authoritative gate: load_skill name=pptx (or word / excel). The per-format SKILL.md, not this blurb, is the source of truth for what 'done' requires.";

    private static void WriteToolDefinitions(Utf8JsonWriter w)
    {
        w.WriteStartObject();
        w.WriteString("name", "officecli");
        // Append a compact always-on skill-trigger summary so the agent is
        // prompted to load the right skill without the full ~1.2k of routing
        // descriptions resident in context. Detail stays lazy behind load_skill.
        w.WriteString("description", ToolDescription + "\n\n" + OfficeCli.Core.SkillInstaller.BuildSkillTriggerSummary());
        w.WriteStartObject("inputSchema");
        w.WriteString("type", "object");
        w.WriteStartObject("properties");
        // command
        w.WriteStartObject("command"); w.WriteString("type", "string");
        w.WriteStartArray("enum");
        foreach (var c in CommandNames)
            w.WriteStringValue(c);
        w.WriteEndArray();
        w.WriteString("description", "Command to execute");
        w.WriteEndObject();
        // file
        w.WriteStartObject("file"); w.WriteString("type", "string"); w.WriteString("description", "Document file path"); w.WriteEndObject();
        // path
        w.WriteStartObject("path"); w.WriteString("type", "string"); w.WriteString("description", "DOM path (e.g. /slide[1]/shape[1], /Sheet1/A1, /body/p[1])"); w.WriteEndObject();
        // parent
        w.WriteStartObject("parent"); w.WriteString("type", "string"); w.WriteString("description", "Parent DOM path for add"); w.WriteEndObject();
        // type
        w.WriteStartObject("type"); w.WriteString("type", "string"); w.WriteString("description", "Element type for add (slide, shape, paragraph, run, table, picture, chart, etc.)"); w.WriteEndObject();
        // from (add: copy an existing element instead of creating a new one)
        w.WriteStartObject("from"); w.WriteString("type", "string"); w.WriteString("description", "Source DOM path for add: copy an existing element to the parent instead of creating a new one (mutually exclusive with type)"); w.WriteEndObject();
        // selector
        w.WriteStartObject("selector"); w.WriteString("type", "string"); w.WriteString("description", "CSS-like selector for query. Valid element types per handler: PPT — shape, textbox, title, picture, table, chart, placeholder, connector, group, zoom, ole, equation (NOT 'slide' — use 'slide[N]>shape' to scope); Excel — cell, sheet, row, column, table, chart, image; Word — paragraph, run, table, image, hyperlink, heading, list. Supports attribute filters ('shape[text=Hello]', 'paragraph[style=Normal] > run[font!=Arial]'), pseudo-selectors (:contains(...), :empty), and Excel cell aliases (bold, size → font.bold, font.size). Path-style selectors starting with '/' are rejected except '/slide[N]/...' scoping in PPT."); w.WriteEndObject();
        // text (query post-filter)
        w.WriteStartObject("text"); w.WriteString("type", "string"); w.WriteString("description", "Filter query results to elements whose text contains this substring (case-insensitive)"); w.WriteEndObject();
        // props
        w.WriteStartObject("props"); w.WriteString("type", "array");
        w.WriteStartObject("items"); w.WriteString("type", "string"); w.WriteEndObject();
        w.WriteString("description", "key=value pairs (e.g. bold=true, color=FF0000, text=Hello)"); w.WriteEndObject();
        // mode
        w.WriteStartObject("mode"); w.WriteString("type", "string"); w.WriteString("description", "View mode: text, annotated, outline, stats, issues, html, svg (pptx), screenshot, forms (docx). screenshot renders the page/slide to a PNG and returns it as an image you can look at — use it to VISUALLY verify layout (overlap, overflow, off-slide shapes, styling) that text modes cannot show. Needs a headless browser (playwright/chrome/firefox); takes seconds."); w.WriteEndObject();
        // screenshot_width / screenshot_height / grid (screenshot mode)
        w.WriteStartObject("screenshot_width"); w.WriteString("type", "number"); w.WriteString("description", "Viewport width for screenshot mode (default 1600)"); w.WriteEndObject();
        w.WriteStartObject("screenshot_height"); w.WriteString("type", "number"); w.WriteString("description", "Viewport height for screenshot mode (default 1200)"); w.WriteEndObject();
        w.WriteStartObject("grid"); w.WriteString("type", "number"); w.WriteString("description", "Tile pages/slides into a thumbnail contact sheet (screenshot mode, pptx + docx). N = column count; -1 = auto (pick columns to keep the sheet roughly square); 0 = off."); w.WriteEndObject();
        // depth
        w.WriteStartObject("depth"); w.WriteString("type", "number"); w.WriteString("description", "Child depth for get (default 1)"); w.WriteEndObject();
        // save (get: extract a node's binary payload to a file)
        w.WriteStartObject("save"); w.WriteString("type", "string"); w.WriteString("description", "For get: destination path to extract the node's binary payload (ole/picture/media/embedded only); response Format gets savedTo/savedBytes/savedContentType"); w.WriteEndObject();
        // index
        w.WriteStartObject("index"); w.WriteString("type", "number"); w.WriteString("description", "Insert position (0-based) for add/move"); w.WriteEndObject();
        // to
        w.WriteStartObject("to"); w.WriteString("type", "string"); w.WriteString("description", "Target parent path for move"); w.WriteEndObject();
        // after, before, path2
        w.WriteStartObject("after"); w.WriteString("type", "string"); w.WriteString("description", "Insert after this sibling path (for add/move)"); w.WriteEndObject();
        w.WriteStartObject("before"); w.WriteString("type", "string"); w.WriteString("description", "Insert before this sibling path (for add/move)"); w.WriteEndObject();
        w.WriteStartObject("path2"); w.WriteString("type", "string"); w.WriteString("description", "Second path for swap"); w.WriteEndObject();
        // start, end, max_lines
        w.WriteStartObject("start"); w.WriteString("type", "number"); w.WriteString("description", "Start line for view"); w.WriteEndObject();
        w.WriteStartObject("end"); w.WriteString("type", "number"); w.WriteString("description", "End line for view"); w.WriteEndObject();
        w.WriteStartObject("max_lines"); w.WriteString("type", "number"); w.WriteString("description", "Max lines for view"); w.WriteEndObject();
        // commands
        w.WriteStartObject("commands"); w.WriteString("type", "string"); w.WriteString("description", "JSON array of batch commands"); w.WriteEndObject();
        // force
        w.WriteStartObject("force"); w.WriteString("type", "string"); w.WriteString("description", "Set to 'true' to continue batch on error (default: stop on first error)"); w.WriteEndObject();
        // part
        w.WriteStartObject("part"); w.WriteString("type", "string"); w.WriteString("description", "Part path for raw (e.g. /document, /styles, /slide[1])"); w.WriteEndObject();
        // cols (raw: Excel column filter)
        w.WriteStartObject("cols"); w.WriteString("type", "string"); w.WriteString("description", "Column filter for raw on Excel sheets, comma-separated (e.g. A,B,C). Pair with start/end to window a large sheet instead of dumping the whole part."); w.WriteEndObject();
        // locale / minimal (create)
        w.WriteStartObject("locale"); w.WriteString("type", "string"); w.WriteString("description", "For create (.docx): locale tag (e.g. zh-CN, ja, ar) setting per-script default fonts and RTL for Arabic/Hebrew. Pass en-US to force a deterministic LTR baseline."); w.WriteEndObject();
        w.WriteStartObject("minimal"); w.WriteString("type", "string"); w.WriteString("description", "For create (.docx): set 'true' to skip Word's Normal baseline and emit a raw OOXML-spec docx (compact / edge-case testing)."); w.WriteEndObject();
        // page / limit / page_count (view)
        w.WriteStartObject("page"); w.WriteString("type", "string"); w.WriteString("description", "Page/slide filter for view html mode (e.g. 1, 2-5, 1,3,5)."); w.WriteEndObject();
        w.WriteStartObject("limit"); w.WriteString("type", "number"); w.WriteString("description", "For view issues mode: cap the number of issues returned."); w.WriteEndObject();
        w.WriteStartObject("page_count"); w.WriteString("type", "string"); w.WriteString("description", "For view stats mode (.docx): set 'true' to also report total page count via Word repagination (Windows + Word required; slow)."); w.WriteEndObject();
        // format
        w.WriteStartObject("format"); w.WriteString("type", "string"); w.WriteString("description", "Document format for help: xlsx, pptx, docx"); w.WriteEndObject();
        // name (for load_skill)
        w.WriteStartObject("name"); w.WriteString("type", "string"); w.WriteString("description", "Skill name for load_skill: pptx, word, excel, word-form, morph-ppt, morph-ppt-3d, pitch-deck, academic-paper, data-dashboard, financial-model. Omit to list all skills with their when-to-use triggers. Pair with path= to fetch a bundled reference file listed in the skill's manifest (e.g. path=reference/decision-rules.md)."); w.WriteEndObject();
        w.WriteEndObject(); // end properties
        w.WriteStartArray("required"); w.WriteStringValue("command"); w.WriteEndArray();
        w.WriteEndObject(); // end inputSchema
        w.WriteEndObject(); // end tool
    }

    // ==================== JSON-RPC Helpers ====================

    private static string WriteJson(Action<Utf8JsonWriter> build)
    {
        using var ms = new MemoryStream();
        using (var w = new Utf8JsonWriter(ms)) build(w);
        return Encoding.UTF8.GetString(ms.ToArray());
    }

    private static void Rpc(Utf8JsonWriter w, JsonElement? id)
    {
        w.WriteString("jsonrpc", "2.0");
        if (id.HasValue) { w.WritePropertyName("id"); id.Value.WriteTo(w); }
        else w.WriteNull("id");
    }

    private static string ErrorJson(JsonElement? id, int code, string message) => WriteJson(w =>
    {
        w.WriteStartObject();
        Rpc(w, id);
        w.WriteStartObject("error");
        w.WriteNumber("code", code);
        w.WriteString("message", message);
        w.WriteEndObject();
        w.WriteEndObject();
    });
}
