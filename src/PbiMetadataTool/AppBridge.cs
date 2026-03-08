using System.Text.Json;
using System.Text.Json.Nodes;
using System.Text.RegularExpressions;
using System.Net.Http.Headers;
using System.Reflection;
using System.Diagnostics;
using System.Text;

namespace PbiMetadataTool;

/// <summary>
/// Handles bidirectional messaging between the WebView2 frontend and C# backend.
/// JS → C#: window.chrome.webview.postMessage(json)
/// C# → JS: webView.CoreWebView2.PostWebMessageAsJson(json)
/// </summary>
internal sealed class AppBridge
{
    private const int MaxHistoryEntries = 500;
    private const string GithubLatestReleaseApiUrl = "https://api.github.com/repos/yczx001/PBIClaw/releases/latest";
    private const string GithubReleasesApiUrl = "https://api.github.com/repos/yczx001/PBIClaw/releases?per_page=20";
    private const string GithubLatestReleasePageUrl = "https://github.com/yczx001/PBIClaw/releases/latest";
    private const string ReleasePageUrl = "https://github.com/yczx001/PBIClaw/releases";
    private const int ChatRequestTimeoutSeconds = 180;
    private static readonly JsonSerializerOptions HistoryJsonOptions = new() { WriteIndented = true };
    private static readonly HttpClient UpdateHttpClient = CreateUpdateHttpClient();

    private readonly MainFormWebView _form;
    private readonly CliOptions _startupOptions;
    private readonly PowerBiInstanceDetector _detector = new();
    private readonly TabularMetadataReader _reader = new();
    private readonly OnDemandMetadataContextBuilder _onDemandContextBuilder = new();
    private readonly ReportMetadataReader _reportReader = new();
    private readonly TabularModelWriter _writer = new();
    private readonly AiChatClient _chatClient = new();
    private readonly AbiSettingsStore _settingsStore = new();
    private readonly List<AiChatMessage> _conversation = [];
    private readonly List<PlanHistoryEntry> _planHistory = [];
    private readonly DateTime _sessionStartedUtc = DateTime.UtcNow;
    private readonly string _sessionToken = $"{DateTime.UtcNow:yyyyMMddHHmmss}-{Guid.NewGuid().ToString("N")[..6]}";
    private readonly string _historyPath =
        Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "PBIClaw", "change-history.json");

    private AbiAssistantSettings _settings = new();
    private ModelMetadata? _model;
    private int? _currentPort;
    private string _lastTabularServer = string.Empty;
    private IReadOnlyList<TabularDatabaseInfo> _tabularDatabases = Array.Empty<TabularDatabaseInfo>();
    private AbiActionPlan? _pendingPlan;
    private AbiActionPlan? _preflightPlan; // plan waiting for user confirmation
    private AbiActionPlan? _preflightSourcePlan;
    private IReadOnlyList<int> _preflightPendingIndices = [];
    private bool _preflightIsRollback;
    private bool _externalToolAutoConnectAttempted;
    private bool _chatCancelRequestedByUser;
    private CancellationTokenSource? _chatCts;
    private string _connectedModelDisplayName = string.Empty;
    private string _currentBackupModelKey = string.Empty;
    private ReportMetadata? _report;
    private ReleaseInfo? _latestRelease;

    public AppBridge(MainFormWebView form, CliOptions startupOptions)
    {
        _form = form;
        _startupOptions = startupOptions;
        LoadPlanHistory();
    }

    // ── Inbound (JS → C#) ────────────────────────────────────────────────────

    public void HandleMessage(string json)
    {
        try
        {
            // WebMessageAsJson wraps string payloads in extra quotes: "\"{ ... }\""
            // Unwrap if needed
            var actual = json;
            if (actual.StartsWith("\"") && actual.EndsWith("\""))
            {
                actual = System.Text.Json.JsonSerializer.Deserialize<string>(actual) ?? actual;
            }

            var node = JsonNode.Parse(actual);
            var type = node?["type"]?.GetValue<string>() ?? string.Empty;
            var payload = node?["payload"] as JsonObject ?? new JsonObject();

            switch (type)
            {
                case "ready":       OnReady(); break;
                case "scan":        OnScan(); break;
                case "connect":     OnConnect(payload); break;
                case "requestTabularDatabases": OnRequestTabularDatabases(payload); break;
                case "connectTabularDatabase": OnConnectTabularDatabase(payload); break;
                case "windowControl": OnWindowControl(payload); break;
                case "chat":        _ = OnChatAsync(payload); break;
                case "cancelChat":  OnCancelChat(); break;
                case "testConnection": _ = OnTestConnectionAsync(payload); break;
                case "discardPlan": OnDiscardPlan(); break;
                case "executePlan": OnExecutePlan(payload); break;
                case "confirmExecute": OnConfirmExecute(); break;
                case "executeRollback": OnExecuteRollback(payload); break;
                case "refreshBackups":  OnRefreshBackups(); break;
                case "openBackupFolder": OnOpenBackupFolder(); break;
                case "saveSettings": OnSaveSettings(payload); break;
                case "markStartupGuideSeen": OnMarkStartupGuideSeen(); break;
                case "checkUpdates": _ = OnCheckUpdatesAsync(); break;
                case "upgradeNow": _ = OnUpgradeNowAsync(payload); break;
            }
        }
        catch (Exception ex)
        {
            Send("status", new { text = $"内部错误: {ex.Message}", level = "error" });
        }
    }

    // ── Handlers ─────────────────────────────────────────────────────────────

    private void OnReady()
    {
        _settings = _settingsStore.Load();
        var detectedInstances = _startupOptions.ExternalToolMode
            ? Array.Empty<PowerBiInstanceInfo>()
            : _detector.DiscoverInstances().ToArray();
        var instances = FilterActiveDesktopInstances(detectedInstances);
        Send("init", new
        {
            settings = SettingsDto(_settings),
            instances = instances.Select(InstanceDto).ToArray(),
            appVersion = CurrentAppVersion()
        });
        OnRefreshBackups();
        SendPlanHistory();
        TryAutoConnectFromExternalTool(instances);
    }

    private void OnScan()
    {
        try
        {
            var list = FilterActiveDesktopInstances(_detector.DiscoverInstances());
            Send("instances", new { instances = list.Select(InstanceDto).ToArray() });
        }
        catch (Exception ex)
        {
            Send("status", new { text = $"扫描失败: {ex.Message}", level = "error" });
        }
    }

    private void OnConnect(JsonObject p)
    {
        var mode = p["mode"]?.GetValue<string>() ?? "pbi";
        var allowWrite = p["allowWrite"]?.GetValue<bool>() ?? false;

        try
        {
            if (mode == "tabular")
            {
                var server = p["server"]?.GetValue<string>() ?? string.Empty;
                if (string.IsNullOrWhiteSpace(server))
                {
                    Send("status", new { text = "请输入服务器地址", level = "warn" });
                    return;
                }

                var normalizedServer = server.Trim();
                try
                {
                    SendTabularDatabaseList(
                        normalizedServer,
                        reason: "connect",
                        currentDatabaseName: _model?.DatabaseName,
                        requestedAllowWrite: allowWrite);
                    Send("status", new { text = "服务器连接成功，请选择数据库。", level = "info" });
                    return;
                }
                catch
                {
                    // Fallback for servers where full database enumeration may fail.
                    ConnectTabularServer(normalizedServer, null);
                    TryApplyAllowWriteSetting(allowWrite);
                    SendConnectedModel();
                    OnRefreshBackups();
                    Send("status", new { text = "未能读取数据库列表，已连接默认数据库。", level = "warn" });
                    return;
                }
            }

            var portStr = p["port"]?.GetValue<string>() ?? string.Empty;
            if (!int.TryParse(portStr, out var port))
            {
                Send("status", new { text = "端口无效", level = "warn" });
                return;
            }
            ConnectPbiPort(port, null, sendReportWarning: true);

            TryApplyAllowWriteSetting(allowWrite);

            SendConnectedModel();
            OnRefreshBackups();
        }
        catch (Exception ex)
        {
            Send("status", new { text = $"连接失败: {ex.Message}", level = "error" });
        }
    }

    private void OnRequestTabularDatabases(JsonObject p)
    {
        try
        {
            var server = p["server"]?.GetValue<string>() ?? _lastTabularServer;
            if (string.IsNullOrWhiteSpace(server))
            {
                Send("status", new { text = "未找到 Tabular 服务器地址，请先连接。", level = "warn" });
                return;
            }

            var reason = p["reason"]?.GetValue<string>() ?? "switch";
            SendTabularDatabaseList(
                server.Trim(),
                reason,
                currentDatabaseName: _model?.DatabaseName,
                requestedAllowWrite: ParseOptionalAllowWrite(p));
        }
        catch (Exception ex)
        {
            Send("status", new { text = $"读取数据库列表失败: {ex.Message}", level = "error" });
        }
    }

    private void OnConnectTabularDatabase(JsonObject p)
    {
        try
        {
            var server = p["server"]?.GetValue<string>() ?? _lastTabularServer;
            var database = p["database"]?.GetValue<string>() ?? string.Empty;
            if (string.IsNullOrWhiteSpace(server))
            {
                Send("connectTabularDatabaseFailed", new { error = "服务器地址为空，请重新选择。" });
                Send("status", new { text = "服务器地址为空，请重新选择。", level = "warn" });
                return;
            }

            if (string.IsNullOrWhiteSpace(database))
            {
                Send("connectTabularDatabaseFailed", new { error = "请选择数据库。" });
                Send("status", new { text = "请选择数据库。", level = "warn" });
                return;
            }

            ConnectTabularServer(server.Trim(), database.Trim());
            var allowWrite = ParseOptionalAllowWrite(p);
            if (allowWrite.HasValue)
            {
                TryApplyAllowWriteSetting(allowWrite.Value);
            }

            SendConnectedModel();
            OnRefreshBackups();
            Send("status", new { text = $"已连接到 Tabular 数据库: {_connectedModelDisplayName}", level = "success" });
        }
        catch (Exception ex)
        {
            Send("connectTabularDatabaseFailed", new { error = ex.Message });
            Send("status", new { text = $"连接数据库失败: {ex.Message}", level = "error" });
        }
    }

    private void OnWindowControl(JsonObject p)
    {
        var cmd = p["cmd"]?.GetValue<string>()?.Trim().ToLowerInvariant() ?? string.Empty;
        try
        {
            switch (cmd)
            {
                case "minimize":
                    _form.MinimizeWindow();
                    break;
                case "maximize":
                    _form.ToggleMaximizeRestore();
                    break;
                case "close":
                    _form.CloseWindow();
                    break;
                case "drag":
                    _form.BeginDragMove();
                    break;
            }
        }
        catch
        {
            // Ignore window command errors.
        }
    }

    private async Task OnChatAsync(JsonObject p)
    {
        var prompt = p["prompt"]?.GetValue<string>() ?? string.Empty;
        if (string.IsNullOrWhiteSpace(prompt)) return;
        if (_model is null)
        {
            Send("chatError", new { error = "请先连接到模型" });
            return;
        }

        _chatCts?.Cancel();
        _chatCts?.Dispose();
        _chatCancelRequestedByUser = false;
        var chatCts = new CancellationTokenSource(TimeSpan.FromSeconds(ChatRequestTimeoutSeconds));
        _chatCts = chatCts;
        var requestId = $"chat-{Guid.NewGuid():N}";
        var startedAt = DateTimeOffset.UtcNow;

        try
        {
            Send("chatStart", new
            {
                requestId,
                startedAt = startedAt.ToString("O")
            });
            Send("chatProgress", new { requestId, phase = "context" });

            var runtimeSettings = BuildSettingsFromPayload(p["runtimeSettings"] as JsonObject, _settings);
            var msgs = new List<AiChatMessage>
            {
                new("system", MetadataPromptBuilder.BuildSystemPrompt(runtimeSettings)),
                new("system", MetadataPromptBuilder.BuildModelContextForPrompt(_model, runtimeSettings.IncludeHiddenObjects, prompt, _report))
            };
            var referencedMeasureContext = MetadataPromptBuilder.BuildReferencedMeasureContext(_model, prompt, runtimeSettings.IncludeHiddenObjects);
            if (!string.IsNullOrWhiteSpace(referencedMeasureContext))
            {
                msgs.Add(new AiChatMessage("system", referencedMeasureContext));
            }
            var referencedCalculatedContext = MetadataPromptBuilder.BuildReferencedCalculatedContext(_model, prompt, runtimeSettings.IncludeHiddenObjects);
            if (!string.IsNullOrWhiteSpace(referencedCalculatedContext))
            {
                msgs.Add(new AiChatMessage("system", referencedCalculatedContext));
            }
            var referencedSourceContext = MetadataPromptBuilder.BuildReferencedTableSourceContext(_model, prompt, runtimeSettings.IncludeHiddenObjects);
            if (!string.IsNullOrWhiteSpace(referencedSourceContext))
            {
                msgs.Add(new AiChatMessage("system", referencedSourceContext));
            }
            if (ShouldBuildOnDemandDeepContext(prompt))
            {
                var onDemandDeepContext = await Task.Run(
                    () => BuildOnDemandDeepContext(prompt, runtimeSettings.IncludeHiddenObjects),
                    chatCts.Token);
                if (!string.IsNullOrWhiteSpace(onDemandDeepContext))
                {
                    msgs.Add(new AiChatMessage("system", onDemandDeepContext));
                }
            }
            msgs.AddRange(BuildConversationContextTail(_conversation, maxMessages: 8, maxChars: 12000));
            msgs.Add(new AiChatMessage("user", prompt));

            Send("chatProgress", new { requestId, phase = "requesting" });

            var streamingStarted = false;
            var streamed = false;
            var streamReply = new StringBuilder();

            var completion = await _chatClient.CompleteStreamAsync(
                new AiChatRequest(runtimeSettings, msgs),
                delta =>
                {
                    if (string.IsNullOrEmpty(delta))
                    {
                        return;
                    }

                    if (!streamingStarted)
                    {
                        streamingStarted = true;
                        Send("chatProgress", new { requestId, phase = "streaming" });
                    }

                    streamed = true;
                    streamReply.Append(delta);
                    Send("chatChunk", new { requestId, delta });
                },
                chatCts.Token);

            var reply = string.IsNullOrWhiteSpace(completion.Text)
                ? streamReply.ToString()
                : completion.Text;

            _conversation.Add(new AiChatMessage("user", prompt));
            _conversation.Add(new AiChatMessage("assistant", reply));
            if (_conversation.Count > 40) _conversation.RemoveRange(0, _conversation.Count - 40);

            var durationMs = Math.Max(1, (int)(DateTimeOffset.UtcNow - startedAt).TotalMilliseconds);
            Send("chatDone", new
            {
                requestId,
                text = reply,
                durationMs,
                streamed = completion.Streamed || streamed
            });

            if (runtimeSettings.AllowModelChanges && AbiActionPlanParser.TryExtract(reply, out var plan, out var preview, out _))
            {
                plan = NormalizeActionPlan(plan);
                preview = AbiActionPlanPreview.BuildText(plan);
                _pendingPlan = plan;
                _preflightSourcePlan = null;
                _preflightPendingIndices = [];
                Send("planDetected", PlanDto(plan, preview));

                AppendPlanHistory(
                    kind: "plan_detected",
                    title: $"检测到变更计划（{plan.Actions.Count} 项）",
                    detail: string.IsNullOrWhiteSpace(plan.Summary) ? preview : plan.Summary,
                    actionCount: plan.Actions.Count);
            }
        }
        catch (OperationCanceledException)
        {
            if (_chatCancelRequestedByUser)
            {
                _chatCancelRequestedByUser = false;
                Send("chatCancelled", new { requestId });
                return;
            }

            Send("chatError", new
            {
                requestId,
                error = $"AI 请求超时（{ChatRequestTimeoutSeconds} 秒），请重试或缩小问题范围。"
            });
        }
        catch (Exception ex)
        {
            Send("chatError", new { requestId, error = ex.Message });
        }
        finally
        {
            if (ReferenceEquals(_chatCts, chatCts))
            {
                _chatCts = null;
            }

            chatCts.Dispose();
        }
    }

    private void OnCancelChat()
    {
        try
        {
            _chatCancelRequestedByUser = true;
            _chatCts?.Cancel();
        }
        catch
        {
            // Ignore cancellation errors.
        }
    }

    private static bool ShouldBuildOnDemandDeepContext(string prompt)
    {
        if (string.IsNullOrWhiteSpace(prompt))
        {
            return false;
        }

        var text = prompt.Trim();
        return ContainsAny(
            text,
            "power query",
            "m代码",
            "m 代码",
            "数据来源",
            "数据源",
            "source",
            "dax",
            "计算表",
            "计算列",
            "calculated table",
            "calculated column",
            "measure",
            "度量",
            "关系",
            "relationship",
            "rls",
            "角色",
            "权限",
            "partition",
            "tmdl",
            "tom",
            "模型",
            "model");
    }

    private static IReadOnlyList<AiChatMessage> BuildConversationContextTail(
        IReadOnlyList<AiChatMessage> conversation,
        int maxMessages,
        int maxChars)
    {
        if (conversation.Count == 0 || maxMessages <= 0 || maxChars <= 0)
        {
            return Array.Empty<AiChatMessage>();
        }

        var selected = new List<AiChatMessage>();
        var usedChars = 0;
        for (var i = conversation.Count - 1; i >= 0; i--)
        {
            var msg = conversation[i];
            if (string.Equals(msg.Role, "system", StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            var content = msg.Content ?? string.Empty;
            var len = content.Length;
            if (selected.Count >= maxMessages)
            {
                break;
            }

            if (usedChars + len > maxChars && selected.Count > 0)
            {
                break;
            }

            selected.Add(msg);
            usedChars += len;
        }

        selected.Reverse();
        return selected;
    }

    private static bool ContainsAny(string text, params string[] keywords)
        => keywords.Any(keyword => text.Contains(keyword, StringComparison.OrdinalIgnoreCase));

    private async Task OnTestConnectionAsync(JsonObject p)
    {
        try
        {
            var runtime = BuildSettingsFromPayload(p, _settings);
            using var cts = new CancellationTokenSource(TimeSpan.FromSeconds(20));
            await _chatClient.TestConnectionAsync(runtime, cts.Token);
            Send("testConnectionResult", new { ok = true, message = "连接测试成功" });
        }
        catch (Exception ex)
        {
            Send("testConnectionResult", new { ok = false, message = $"连接测试失败: {ex.Message}" });
        }
        finally
        {
            Send("testConnectionDone", new { });
        }
    }

    private void OnDiscardPlan()
    {
        if (_pendingPlan is not null || _preflightPlan is not null)
        {
            AppendPlanHistory(
                kind: "plan_discarded",
                title: "已丢弃待执行计划",
                detail: "用户手动丢弃当前计划。");
        }
        _pendingPlan = null;
        _preflightPlan = null;
        _preflightSourcePlan = null;
        _preflightPendingIndices = [];
        _preflightIsRollback = false;
    }

    private void OnExecutePlan(JsonObject p)
    {
        if (!_settings.AllowModelChanges)
        {
            Send("status", new { text = "写回功能已禁用，请在设置中开启", level = "warn" });
            return;
        }
        if (!_currentPort.HasValue || _model is null)
        {
            Send("status", new { text = "未连接到模型", level = "warn" });
            return;
        }
        if (_pendingPlan is null)
        {
            Send("status", new { text = "没有待执行的计划", level = "warn" });
            return;
        }
        _pendingPlan = NormalizeActionPlan(_pendingPlan);

        var indices = p["selectedIndices"]?.AsArray()
            .Select(n => n?.GetValue<int>() ?? -1)
            .Where(i => i >= 0 && i < _pendingPlan.Actions.Count)
            .Distinct()
            .OrderBy(i => i)
            .ToList() ?? [];

        if (indices.Count == 0)
        {
            Send("status", new { text = "请至少选择一个动作", level = "warn" });
            return;
        }

        var autoExecute = p["autoExecute"]?.GetValue<bool>() ?? false;
        var selectedSet = indices.ToHashSet();
        var pendingIndices = p["pendingIndices"]?.AsArray()
            .Select(n => n?.GetValue<int>() ?? -1)
            .Where(i => i >= 0 && i < _pendingPlan.Actions.Count && !selectedSet.Contains(i))
            .Distinct()
            .OrderBy(i => i)
            .ToList() ?? [];
        var selectedActions = indices.Select(i => _pendingPlan.Actions[i]).ToList();
        _preflightPlan = new AbiActionPlan(_pendingPlan.Summary, selectedActions);
        _preflightSourcePlan = _pendingPlan;
        _preflightPendingIndices = pendingIndices;
        _preflightIsRollback = false;

        try
        {
            var analysis = _writer.AnalyzeActions(_currentPort.Value, _model.DatabaseName, _preflightPlan);
            var preflightDetail = $"错误 {analysis.Errors.Count}，警告 {analysis.Warnings.Count}。";
            if (analysis.Errors.Count > 0)
            {
                preflightDetail += $" 首条错误：{analysis.Errors[0]}";
            }
            else if (analysis.Warnings.Count > 0)
            {
                preflightDetail += $" 首条警告：{analysis.Warnings[0]}";
            }

            AppendPlanHistory(
                kind: analysis.HasErrors ? "preflight_failed" : "preflight_passed",
                title: analysis.HasErrors ? "预检失败" : "预检通过",
                detail: preflightDetail,
                actionCount: selectedActions.Count,
                success: !analysis.HasErrors);

            Send("preflight", new
            {
                hasErrors = analysis.HasErrors,
                errors = analysis.Errors.Take(12).ToArray(),
                warnings = analysis.Warnings.ToArray(),
                infos = new[] { $"将执行 {selectedActions.Count} 项变更" },
                autoExecute = autoExecute && !analysis.HasErrors
            });

            if (autoExecute && !analysis.HasErrors)
            {
                Send("status", new { text = "预检通过，已按设置自动执行。", level = "info" });
                OnConfirmExecute();
            }
        }
        catch (Exception ex)
        {
            AppendPlanHistory(
                kind: "preflight_error",
                title: "预检异常",
                detail: ex.Message,
                actionCount: selectedActions.Count,
                success: false);
            Send("status", new { text = $"预检失败: {ex.Message}", level = "error" });
        }
    }

    private void OnConfirmExecute()
    {
        if (_preflightPlan is null || !_currentPort.HasValue || _model is null) return;
        var currentPlan = _preflightPlan;
        var sourcePlan = _preflightSourcePlan ?? _pendingPlan;
        var pendingIndices = _preflightPendingIndices;
        var isRollback = _preflightIsRollback;
        try
        {
            var (snapshot, rollback) = CreateBackup(currentPlan);
            Send("log", new { text = $"执行前已创建备份:\n- {snapshot}\n- {rollback}" });

            var results = _writer.ApplyActions(_currentPort.Value, _model.DatabaseName, currentPlan);
            AbiActionPlan? remainingPlan = null;
            if (!isRollback && sourcePlan is not null && pendingIndices.Count > 0)
            {
                var remaining = pendingIndices
                    .Where(i => i >= 0 && i < sourcePlan.Actions.Count)
                    .Distinct()
                    .OrderBy(i => i)
                    .Select(i => sourcePlan.Actions[i])
                    .ToList();
                if (remaining.Count > 0)
                {
                    remainingPlan = new AbiActionPlan(sourcePlan.Summary, remaining);
                }
            }

            _pendingPlan = remainingPlan;
            _preflightPlan = null;
            _preflightSourcePlan = null;
            _preflightPendingIndices = [];
            _preflightIsRollback = false;

            // Refresh model
            _model = _reader.ReadMetadata(_currentPort.Value, _model.DatabaseName);
            _report = ResolveReportMetadataForPort(_currentPort.Value).Report;
            Send("connected", BuildConnectedPayload(silent: true));

            Send("executeSuccess", new
            {
                results = results.ToArray(),
                remainingPlan = remainingPlan is null
                    ? null
                    : PlanDto(
                        remainingPlan,
                        actionStates: Enumerable.Repeat("pending", remainingPlan.Actions.Count).ToArray())
            });
            AppendPlanHistory(
                kind: isRollback ? "rollback_execute_success" : "execute_success",
                title: isRollback ? $"回滚执行完成（{currentPlan.Actions.Count} 项）" : $"变更执行完成（{currentPlan.Actions.Count} 项）",
                detail: string.Join(Environment.NewLine, results),
                actionCount: currentPlan.Actions.Count,
                success: true);
            OnRefreshBackups();
        }
        catch (Exception ex)
        {
            AppendPlanHistory(
                kind: isRollback ? "rollback_execute_error" : "execute_error",
                title: isRollback ? "回滚执行失败" : "变更执行失败",
                detail: ex.Message,
                actionCount: currentPlan.Actions.Count,
                success: false);
            Send("executeError", new { error = ex.Message });
        }
    }

    private void OnExecuteRollback(JsonObject p)
    {
        var path = p["path"]?.GetValue<string>() ?? string.Empty;
        if (string.IsNullOrWhiteSpace(path) || !_currentPort.HasValue || _model is null)
        {
            Send("status", new { text = "无法执行回滚：未连接或路径无效", level = "warn" });
            return;
        }
        try
        {
            var plan = AbiActionPlanStorage.Load(path);
            if (plan.Actions.Count == 0)
            {
                var warn = "该回滚点不包含可自动执行动作（通常是删除表/列等不可逆变更）。";
                AppendPlanHistory(
                    kind: "rollback_empty",
                    title: "回滚计划为空",
                    detail: $"{Path.GetFileName(path)} | {warn}",
                    success: false);
                Send("status", new { text = warn, level = "warn" });
                return;
            }

            var analysis = _writer.AnalyzeActions(_currentPort.Value, _model.DatabaseName, plan);
            _preflightPlan = plan;
            _preflightSourcePlan = null;
            _preflightPendingIndices = [];
            _preflightIsRollback = true;
            Send("preflight", new
            {
                hasErrors = analysis.HasErrors,
                errors = analysis.Errors.Take(12).ToArray(),
                warnings = analysis.Warnings.ToArray(),
                infos = new[] { $"回滚计划包含 {plan.Actions.Count} 项操作" },
                autoExecute = false
            });
            AppendPlanHistory(
                kind: "rollback_loaded",
                title: $"已加载回滚计划（{plan.Actions.Count} 项）",
                detail: Path.GetFileName(path),
                actionCount: plan.Actions.Count);
        }
        catch (Exception ex)
        {
            var message = ex.Message;
            if (message.Contains("actions 为空", StringComparison.OrdinalIgnoreCase))
            {
                message = "回滚文件不包含可执行动作。此类改动通常无法自动回滚（例如删除表/列）。";
            }

            AppendPlanHistory(
                kind: "rollback_error",
                title: "回滚加载失败",
                detail: message,
                success: false);
            Send("status", new { text = $"回滚失败: {message}", level = "error" });
        }
    }

    private void OnRefreshBackups()
    {
        var root = BackupRoot();
        Directory.CreateDirectory(root);
        var currentPrefix = GetCurrentSessionBackupPrefix();
        if (string.IsNullOrWhiteSpace(currentPrefix))
        {
            Send("backups", new { list = Array.Empty<object>() });
            return;
        }

        var pairs = new Dictionary<string, (FileInfo? Snapshot, FileInfo? Rollback)>(StringComparer.OrdinalIgnoreCase);

        foreach (var path in Directory.EnumerateFiles(root, "*.json", SearchOption.AllDirectories))
        {
            var file = new FileInfo(path);
            if (!file.Name.StartsWith(currentPrefix, StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            if (!TryParseBackupMarker(file.Name, out var key, out var isRollback))
            {
                continue;
            }

            var existing = pairs.TryGetValue(key, out var tuple) ? tuple : (null, null);
            if (isRollback)
            {
                existing.Item2 = file;
            }
            else
            {
                existing.Item1 = file;
            }

            pairs[key] = existing;
        }

        var merged = new List<BackupEntry>();
        foreach (var pair in pairs.Values)
        {
            var primary = pair.Rollback ?? pair.Snapshot;
            if (primary is null)
            {
                continue;
            }

            var totalSize = (pair.Rollback?.Length ?? 0) + (pair.Snapshot?.Length ?? 0);
            var rollbackStatus = "仅快照";
            var rollbackHint = "该条目仅包含模型快照，不包含自动回滚计划。";
            var rollbackSummary = "仅快照，无可执行回滚计划。";
            var canRollback = false;
            if (pair.Rollback is not null)
            {
                (rollbackStatus, rollbackHint, rollbackSummary, canRollback) = EvaluateRollbackEntry(pair.Rollback.FullName);
            }

            merged.Add(new BackupEntry(
                primary.LastWriteTime,
                pair.Rollback is not null ? "回滚" : "快照",
                pair.Rollback is not null && pair.Snapshot is not null
                    ? pair.Rollback.Name.Replace("-rollback-", "-backup-", StringComparison.OrdinalIgnoreCase)
                    : primary.Name,
                $"{Math.Max(1, totalSize / 1024)} KB",
                primary.DirectoryName ?? string.Empty,
                (pair.Rollback ?? pair.Snapshot)!.FullName,
                rollbackStatus,
                rollbackHint,
                rollbackSummary,
                canRollback));
        }

        var files = merged
            .OrderByDescending(x => x.Time)
            .Select(x => new
            {
                time = x.Time.ToString("yyyy-MM-dd HH:mm:ss"),
                type = x.Type,
                name = x.Name,
                size = x.Size,
                dir = x.Dir,
                path = x.Path,
                rollbackStatus = x.RollbackStatus,
                rollbackHint = x.RollbackHint,
                rollbackSummary = x.RollbackSummary,
                canRollback = x.CanRollback
            })
            .ToArray();
        Send("backups", new { list = files });
    }

    private static bool TryParseBackupMarker(string fileName, out string key, out bool isRollback)
    {
        key = string.Empty;
        isRollback = false;

        const string rollbackMarker = "-rollback-";
        const string snapshotMarker = "-before-";

        var marker = string.Empty;
        if (fileName.Contains(rollbackMarker, StringComparison.OrdinalIgnoreCase))
        {
            marker = rollbackMarker;
            isRollback = true;
        }
        else if (fileName.Contains(snapshotMarker, StringComparison.OrdinalIgnoreCase))
        {
            marker = snapshotMarker;
        }

        if (string.IsNullOrWhiteSpace(marker))
        {
            return false;
        }

        var idx = fileName.IndexOf(marker, StringComparison.OrdinalIgnoreCase);
        if (idx <= 0)
        {
            return false;
        }

        var prefix = fileName[..idx];
        var rest = fileName[(idx + marker.Length)..];
        var timestamp = Path.GetFileNameWithoutExtension(rest);
        if (string.IsNullOrWhiteSpace(timestamp))
        {
            return false;
        }

        key = $"{prefix}|{timestamp}";
        return true;
    }

    private (string Status, string Hint, string Summary, bool CanRollback) EvaluateRollbackEntry(string path)
    {
        try
        {
            var plan = AbiActionPlanStorage.Load(path);
            var summary = BuildRollbackOverview(plan);
            if (plan.Actions.Count == 0)
            {
                return ("不可回滚", "该回滚点不包含可自动执行动作。", summary, false);
            }

            if (!_currentPort.HasValue || _model is null)
            {
                return ("待连接验证", $"计划包含 {plan.Actions.Count} 个动作，连接模型后可验证可执行性。", summary, false);
            }

            var analysis = _writer.AnalyzeActions(_currentPort.Value, _model.DatabaseName, plan);
            if (analysis.HasErrors)
            {
                var first = analysis.Errors.FirstOrDefault() ?? "预检失败";
                return ("不可回滚", first, summary, false);
            }

            if (analysis.Warnings.Count > 0)
            {
                return ("可回滚", $"预检通过（{analysis.Warnings.Count} 条警告）", summary, true);
            }

            return ("可回滚", "预检通过。", summary, true);
        }
        catch (Exception ex)
        {
            return ("不可回滚", $"回滚文件解析失败: {ex.Message}", "无法解析变更概述。", false);
        }
    }

    private static string BuildRollbackOverview(AbiActionPlan plan)
    {
        var count = plan.Actions.Count;
        var summary = plan.Summary?.Trim() ?? string.Empty;
        var hasUserSummary = !string.IsNullOrWhiteSpace(summary) &&
                             !summary.StartsWith("回滚计划生成于", StringComparison.OrdinalIgnoreCase);

        if (count == 0)
        {
            return hasUserSummary ? $"{summary}（0 项）" : "0 项动作";
        }

        if (hasUserSummary)
        {
            return $"{summary}（{count} 项）";
        }

        var topKinds = plan.Actions
            .GroupBy(a => NormalizeActionTypeLabel(a.Type), StringComparer.OrdinalIgnoreCase)
            .OrderByDescending(g => g.Count())
            .ThenBy(g => g.Key, StringComparer.OrdinalIgnoreCase)
            .Take(3)
            .Select(g => $"{g.Key}×{g.Count()}")
            .ToArray();

        var target = plan.Actions
            .Select(BuildActionTargetLabel)
            .FirstOrDefault(t => !string.IsNullOrWhiteSpace(t));

        var result = $"共 {count} 项";
        if (topKinds.Length > 0)
        {
            result += $"：{string.Join("，", topKinds)}";
        }
        if (!string.IsNullOrWhiteSpace(target))
        {
            result += $"；例如 {target}";
        }

        return result;
    }

    private static string NormalizeActionTypeLabel(string? type)
    {
        var normalized = (type ?? string.Empty).Trim().ToLowerInvariant();
        return normalized switch
        {
            "create_or_update_measure" => "度量值更新",
            "delete_measure" => "度量值删除",
            "create_relationship" => "关系创建",
            "delete_relationship" => "关系删除",
            "set_relationship_active" => "关系激活状态",
            "set_relationship_cross_filter" => "关系筛选方向",
            "create_calculated_column" => "计算列创建",
            "delete_column" => "列删除",
            "create_calculated_table" => "计算表创建",
            "delete_table" => "表删除",
            "rename_table" => "表重命名",
            "rename_column" => "列重命名",
            "rename_measure" => "度量值重命名",
            "set_table_hidden" => "表显隐",
            "set_column_hidden" => "列显隐",
            "set_measure_hidden" => "度量值显隐",
            "set_format_string" => "格式更新",
            "set_display_folder" => "显示文件夹",
            "create_role" => "角色创建",
            "update_role" => "角色更新",
            "delete_role" => "角色删除",
            "set_role_table_permission" => "角色表权限",
            "remove_role_table_permission" => "角色表权限移除",
            "add_role_member" => "角色成员新增",
            "remove_role_member" => "角色成员移除",
            "update_description" => "描述更新",
            _ => string.IsNullOrWhiteSpace(type) ? "其他" : type.Trim()
        };
    }

    private static string BuildActionTargetLabel(AbiModelAction action)
    {
        if (!string.IsNullOrWhiteSpace(action.Table) && !string.IsNullOrWhiteSpace(action.Name))
        {
            return $"{action.Table}[{action.Name}]";
        }

        if (!string.IsNullOrWhiteSpace(action.Name))
        {
            return action.Name;
        }

        if (!string.IsNullOrWhiteSpace(action.Table))
        {
            return action.Table;
        }

        if (!string.IsNullOrWhiteSpace(action.FromTable) && !string.IsNullOrWhiteSpace(action.ToTable))
        {
            return $"{action.FromTable} -> {action.ToTable}";
        }

        return string.Empty;
    }

    private static AbiActionPlan NormalizeActionPlan(AbiActionPlan plan)
    {
        if (plan.Actions.Count == 0)
        {
            return plan;
        }

        var normalized = plan.Actions.Select(NormalizeAction).ToList();
        return new AbiActionPlan(plan.Summary, normalized);
    }

    private static AbiModelAction NormalizeAction(AbiModelAction action)
    {
        var type = (action.Type ?? string.Empty).Trim().ToLowerInvariant();
        if (!string.Equals(type, "create_calculated_table", StringComparison.Ordinal))
        {
            return action;
        }

        var name = FirstNonEmpty(action.Name, action.Table);
        var expression = action.Expression;
        if (string.IsNullOrWhiteSpace(expression) &&
            IsDateTableIntent(name, action.Reason, action.Table))
        {
            expression = """
ADDCOLUMNS(
    CALENDARAUTO(12),
    "年份", YEAR([Date]),
    "月份序号", MONTH([Date]),
    "月份", FORMAT([Date], "YYYY-MM"),
    "季度", "Q" & FORMAT([Date], "Q"),
    "年季", FORMAT([Date], "YYYY") & "-Q" & FORMAT([Date], "Q"),
    "周数", WEEKNUM([Date], 2),
    "星期", FORMAT([Date], "ddd")
)
""";
        }

        return action with
        {
            Name = string.IsNullOrWhiteSpace(name) ? action.Name : name,
            Expression = string.IsNullOrWhiteSpace(expression) ? action.Expression : expression
        };
    }

    private static bool IsDateTableIntent(params string?[] values)
    {
        var text = string.Join(" ", values.Where(v => !string.IsNullOrWhiteSpace(v))).ToLowerInvariant();
        if (string.IsNullOrWhiteSpace(text))
        {
            return false;
        }

        return text.Contains("日期", StringComparison.Ordinal) ||
               text.Contains("日历", StringComparison.Ordinal) ||
               text.Contains("calendar", StringComparison.Ordinal) ||
               text.Contains("date table", StringComparison.Ordinal) ||
               text.Contains("date_table", StringComparison.Ordinal);
    }

    private static string? FirstNonEmpty(params string?[] values)
    {
        foreach (var value in values)
        {
            if (!string.IsNullOrWhiteSpace(value))
            {
                return value.Trim();
            }
        }

        return null;
    }

    private sealed record BackupEntry(
        DateTime Time,
        string Type,
        string Name,
        string Size,
        string Dir,
        string Path,
        string RollbackStatus,
        string RollbackHint,
        string RollbackSummary,
        bool CanRollback);

    private void OnOpenBackupFolder()
    {
        var root = BackupRoot();
        Directory.CreateDirectory(root);
        System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
        {
            FileName = "explorer.exe",
            Arguments = $"\"{root}\"",
            UseShellExecute = true
        });
    }

    private void OnSaveSettings(JsonObject p)
    {
        try
        {
            _settings = BuildSettingsFromPayload(p, _settings);
            _settingsStore.Save(_settings);
            Send("settingsSaved", SettingsDto(_settings));
        }
        catch (Exception ex)
        {
            Send("settingsSaveFailed", new { error = ex.Message });
        }
    }

    private void OnMarkStartupGuideSeen()
    {
        try
        {
            if (!_settings.ShowStartupGuide)
            {
                return;
            }

            _settings.ShowStartupGuide = false;
            _settingsStore.Save(_settings);
        }
        catch
        {
            // Ignore onboarding persistence failures.
        }
    }

    private async Task OnCheckUpdatesAsync()
    {
        try
        {
            var release = await GetLatestReleaseAsync().ConfigureAwait(false);
            _latestRelease = release;

            var currentVersion = CurrentAppVersion();
            var hasUpdate = CompareVersion(release.Version, currentVersion) > 0;
            Send("updateChecked", new
            {
                hasUpdate,
                currentVersion,
                latestVersion = release.Version,
                publishedAt = release.PublishedAt,
                releaseUrl = release.ReleaseUrl,
                downloadUrl = release.DownloadUrl,
                summary = release.Summary
            });
        }
        catch (Exception ex)
        {
            Send("updateCheckFailed", new { error = ex.Message });
        }
    }

    private async Task OnUpgradeNowAsync(JsonObject payload)
    {
        try
        {
            var releaseUrl = payload["releaseUrl"]?.GetValue<string>()?.Trim();
            if (string.IsNullOrWhiteSpace(releaseUrl))
            {
                releaseUrl = _latestRelease?.ReleaseUrl;
            }

            if (string.IsNullOrWhiteSpace(releaseUrl))
            {
                releaseUrl = ReleasePageUrl;
            }

            Process.Start(new ProcessStartInfo
            {
                FileName = releaseUrl,
                UseShellExecute = true
            });

            Send("upgradeStarted", new { path = releaseUrl });
        }
        catch (Exception ex)
        {
            Send("upgradeFailed", new { error = ex.Message });
        }
    }

    private void AppendPlanHistory(string kind, string title, string detail, int actionCount = 0, bool? success = null)
    {
        _planHistory.Insert(0, new PlanHistoryEntry
        {
            TimeUtc = DateTime.UtcNow,
            Kind = kind,
            Model = _connectedModelDisplayName,
            Title = title,
            Detail = detail,
            ActionCount = actionCount,
            Success = success
        });

        if (_planHistory.Count > MaxHistoryEntries)
        {
            _planHistory.RemoveRange(MaxHistoryEntries, _planHistory.Count - MaxHistoryEntries);
        }

        SavePlanHistory();
        SendPlanHistory();
    }

    private void SendPlanHistory()
    {
        if (string.IsNullOrWhiteSpace(_connectedModelDisplayName))
        {
            Send("planHistory", new { list = Array.Empty<object>() });
            return;
        }

        var currentModel = _connectedModelDisplayName.Trim();
        var list = _planHistory
            .Where(h =>
                h.TimeUtc >= _sessionStartedUtc &&
                string.Equals(h.Model, currentModel, StringComparison.OrdinalIgnoreCase))
            .OrderByDescending(h => h.TimeUtc)
            .Select(h => new
            {
                time = h.TimeUtc.ToLocalTime().ToString("yyyy-MM-dd HH:mm:ss"),
                kind = h.Kind,
                model = h.Model,
                title = h.Title,
                detail = h.Detail,
                actionCount = h.ActionCount,
                success = h.Success
            })
            .ToArray();
        Send("planHistory", new { list });
    }

    private void LoadPlanHistory()
    {
        try
        {
            if (!File.Exists(_historyPath))
            {
                return;
            }

            var json = File.ReadAllText(_historyPath);
            var list = JsonSerializer.Deserialize<List<PlanHistoryEntry>>(json, HistoryJsonOptions);
            if (list is null || list.Count == 0)
            {
                return;
            }

            _planHistory.Clear();
            _planHistory.AddRange(list
                .OrderByDescending(h => h.TimeUtc)
                .Take(MaxHistoryEntries));
        }
        catch
        {
            _planHistory.Clear();
        }
    }

    private void SavePlanHistory()
    {
        try
        {
            var directory = Path.GetDirectoryName(_historyPath);
            if (!string.IsNullOrWhiteSpace(directory) && !Directory.Exists(directory))
            {
                Directory.CreateDirectory(directory);
            }

            var json = JsonSerializer.Serialize(_planHistory, HistoryJsonOptions);
            File.WriteAllText(_historyPath, json);
        }
        catch
        {
            // Ignore history write errors to avoid breaking core workflow.
        }
    }

    // ── Outbound (C# → JS) ───────────────────────────────────────────────────

    public void Send(string type, object payload)
    {
        var json = JsonSerializer.Serialize(new { type, payload });
        if (!string.Equals(type, "chatChunk", StringComparison.OrdinalIgnoreCase) &&
            !string.Equals(type, "chatProgress", StringComparison.OrdinalIgnoreCase))
        {
            var logPath = Path.Combine(Path.GetTempPath(), "abi_scan_debug.txt");
            File.AppendAllText(logPath, $"[{DateTime.Now:HH:mm:ss}] Send({type}) via ExecuteScript\n");
        }
        // Encode as base64 to safely pass arbitrary JSON into JS without escaping issues
        var b64 = Convert.ToBase64String(System.Text.Encoding.UTF8.GetBytes(json));
        _form.ExecuteScript($"try{{App.handle(JSON.parse(atob('{b64}')))}}catch(e){{console.error('Send error',e)}}");
    }

    // ── Helpers ──────────────────────────────────────────────────────────────

    private object BuildConnectedPayload(bool silent)
    {
        if (_model is null)
        {
            return new
            {
                dbName = string.Empty,
                allowChanges = _settings.AllowModelChanges,
                model = (object?)null,
                silent,
                connectionMode = "pbi",
                tabularServer = string.Empty,
                tabularDatabases = Array.Empty<object>()
            };
        }

        var isTabular = !_currentPort.HasValue;
        return new
        {
            dbName = _connectedModelDisplayName,
            allowChanges = _settings.AllowModelChanges,
            model = ModelDto(_model, _connectedModelDisplayName, _report),
            silent,
            connectionMode = isTabular ? "tabular" : "pbi",
            tabularServer = isTabular ? _lastTabularServer : string.Empty,
            tabularDatabases = isTabular
                ? _tabularDatabases.Select(TabularDatabaseDto).ToArray()
                : Array.Empty<object>()
        };
    }

    private void TryApplyAllowWriteSetting(bool allowWrite)
    {
        if (allowWrite == _settings.AllowModelChanges)
        {
            return;
        }

        _settings.AllowModelChanges = allowWrite;
        _settingsStore.Save(_settings);
    }

    private static bool? ParseOptionalAllowWrite(JsonObject payload)
    {
        if (payload["allowWrite"] is not JsonValue value)
        {
            return null;
        }

        try
        {
            return value.GetValue<bool>();
        }
        catch
        {
            return null;
        }
    }

    private static IReadOnlyList<PowerBiInstanceInfo> FilterActiveDesktopInstances(IReadOnlyList<PowerBiInstanceInfo> instances)
    {
        if (instances.Count == 0)
        {
            return instances;
        }

        return instances
            .Where(i => i.DesktopPid > 0)
            .ToArray();
    }

    private void TryAutoConnectFromExternalTool(IReadOnlyList<PowerBiInstanceInfo>? instances = null)
    {
        if (_externalToolAutoConnectAttempted || _model is not null || !_startupOptions.ExternalToolMode)
        {
            return;
        }

        _externalToolAutoConnectAttempted = true;
        var normalizedServer = NormalizeExternalToolValue(_startupOptions.Server);
        var databaseName = NormalizeExternalToolValue(_startupOptions.DatabaseName);
        var preferredPort = _startupOptions.Port ?? ParsePortFromServer(normalizedServer);
        var hasContext = preferredPort.HasValue || !string.IsNullOrWhiteSpace(normalizedServer);
        if (!hasContext)
        {
            return;
        }
        Send("status", new { text = "检测到外部工具上下文，正在自动连接模型...", level = "info" });

        var candidates = new List<int>();
        if (preferredPort.HasValue)
        {
            candidates.Add(preferredPort.Value);
        }

        if (candidates.Count == 0)
        {
            var discovered = instances ?? _detector.DiscoverInstances();
            candidates = BuildPortCandidates(preferredPort, discovered);
        }
        var errors = new List<string>();

        foreach (var port in candidates)
        {
            try
            {
                ConnectPbiPort(port, databaseName, sendReportWarning: false, loadReportMetadata: false);
                SendConnectedModel();
                Send("status", new { text = $"已自动连接当前报表模型（端口 {port}）。", level = "success" });
                _ = LoadReportMetadataForCurrentConnectionAsync(port, showWarning: true);
                OnRefreshBackups();
                return;
            }
            catch (Exception ex)
            {
                errors.Add($"[{port}] {ex.Message}");
            }
        }

        if (!string.IsNullOrWhiteSpace(normalizedServer))
        {
            try
            {
                ConnectTabularServer(normalizedServer, databaseName);
                SendConnectedModel();
                Send("status", new { text = "已根据外部工具参数自动连接模型。", level = "success" });
                OnRefreshBackups();
                return;
            }
            catch (Exception ex)
            {
                errors.Add($"[server={normalizedServer}] {ex.Message}");
            }
        }

        var hint = errors.Count > 0
            ? $"自动连接失败：{errors[^1]}"
            : "自动连接失败：未找到可用模型端口。";
        Send("status", new { text = hint, level = "warn" });
    }

    private void ConnectPbiPort(int port, string? databaseName, bool sendReportWarning, bool loadReportMetadata = true)
    {
        _model = _reader.ReadMetadata(port, databaseName);
        _currentPort = port;
        _tabularDatabases = Array.Empty<TabularDatabaseInfo>();
        _connectedModelDisplayName = ResolvePbiDisplayName(port, _model.DatabaseName);
        _currentBackupModelKey = SafeName(_connectedModelDisplayName);
        if (!loadReportMetadata)
        {
            _report = null;
            return;
        }

        var reportResolve = ResolveReportMetadataForPort(port);
        _report = reportResolve.Report;
        if (sendReportWarning && _report is null && !string.IsNullOrWhiteSpace(reportResolve.Warning))
        {
            Send("status", new { text = reportResolve.Warning, level = "warn" });
        }
    }

    private void ConnectTabularServer(string server, string? databaseName = null)
    {
        _lastTabularServer = server;
        var connStr = $"DataSource={server};";
        _model = _reader.ReadMetadata(connStr, databaseName);
        _currentPort = null;
        _connectedModelDisplayName = _model.DatabaseName;
        _currentBackupModelKey = SafeName(_connectedModelDisplayName);
        _report = null;
    }

    private void SendTabularDatabaseList(
        string server,
        string reason,
        string? currentDatabaseName,
        bool? requestedAllowWrite)
    {
        var connStr = $"DataSource={server};";
        var databases = _reader.ListDatabases(connStr);
        if (databases.Count == 0)
        {
            throw new InvalidOperationException("未找到可用数据库。");
        }

        _lastTabularServer = server;
        _tabularDatabases = databases;

        Send("tabularDatabases", new
        {
            server,
            reason,
            currentDatabase = currentDatabaseName,
            allowWrite = requestedAllowWrite,
            databases = databases.Select(TabularDatabaseDto).ToArray()
        });
    }

    private string BuildOnDemandDeepContext(string prompt, bool includeHiddenObjects)
    {
        if (_model is null)
        {
            return string.Empty;
        }

        var connectionString = _currentPort.HasValue
            ? $"DataSource=localhost:{_currentPort.Value};"
            : (!string.IsNullOrWhiteSpace(_lastTabularServer) ? $"DataSource={_lastTabularServer};" : string.Empty);

        if (string.IsNullOrWhiteSpace(connectionString))
        {
            return string.Empty;
        }

        try
        {
            return _onDemandContextBuilder.Build(
                connectionString,
                _model.DatabaseName,
                _model,
                prompt,
                includeHiddenObjects,
                ResolveCurrentModelSourcePaths());
        }
        catch
        {
            return string.Empty;
        }
    }

    private IReadOnlyList<string> ResolveCurrentModelSourcePaths()
    {
        var candidates = new List<string>();

        // 优先使用报表源路径（PBIP/PBIX），可最大概率命中 TMDL。
        if (_report is not null && !string.IsNullOrWhiteSpace(_report.SourcePath))
        {
            AddPathCandidate(candidates, _report.SourcePath);
        }

        if (_currentPort.HasValue)
        {
            try
            {
                var hit = _detector.DiscoverInstances().FirstOrDefault(i => i.Port == _currentPort.Value);
                if (hit is not null && !string.IsNullOrWhiteSpace(hit.PbixPathHint))
                {
                    AddPathCandidate(candidates, hit.PbixPathHint);
                }

                if (hit is not null && !string.IsNullOrWhiteSpace(hit.WorkspacePath))
                {
                    AddPathCandidate(candidates, hit.WorkspacePath);
                }
            }
            catch
            {
                // Ignore detection errors and continue with collected paths.
            }
        }

        return candidates;
    }

    private static void AddPathCandidate(List<string> candidates, string? rawPath)
    {
        if (string.IsNullOrWhiteSpace(rawPath))
        {
            return;
        }

        try
        {
            var fullPath = Path.GetFullPath(Environment.ExpandEnvironmentVariables(rawPath.Trim()));
            AddDistinctPath(candidates, fullPath);

            if (File.Exists(fullPath))
            {
                var fileDir = Path.GetDirectoryName(fullPath);
                AddDistinctPath(candidates, fileDir);

                if (!string.IsNullOrWhiteSpace(fileDir))
                {
                    var parent = Directory.GetParent(fileDir)?.FullName;
                    AddDistinctPath(candidates, parent);
                }

                return;
            }

            if (Directory.Exists(fullPath))
            {
                var parent = Directory.GetParent(fullPath)?.FullName;
                AddDistinctPath(candidates, parent);
            }
        }
        catch
        {
            // Ignore invalid paths.
        }
    }

    private static void AddDistinctPath(List<string> candidates, string? path)
    {
        if (string.IsNullOrWhiteSpace(path))
        {
            return;
        }

        if (candidates.Any(existing => string.Equals(existing, path, StringComparison.OrdinalIgnoreCase)))
        {
            return;
        }

        candidates.Add(path);
    }

    private void SendConnectedModel()
    {
        if (_model is null)
        {
            return;
        }

        _conversation.Clear();
        Send("connected", BuildConnectedPayload(silent: false));

        if (_report is not null)
        {
            var visualCount = _report.Pages.Sum(page => page.Visuals.Count);
            Send("status", new { text = $"已加载报表信息：{_report.Pages.Count} 页，{visualCount} 个视觉对象。", level = "success" });
        }

        SendPlanHistory();
    }

    private async Task LoadReportMetadataForCurrentConnectionAsync(int port, bool showWarning)
    {
        try
        {
            var resolve = await Task.Run(() => ResolveReportMetadataForPort(port)).ConfigureAwait(false);
            if (!_currentPort.HasValue || _currentPort.Value != port || _model is null)
            {
                return;
            }

            _report = resolve.Report;
            if (_report is not null)
            {
                Send("connected", BuildConnectedPayload(silent: true));
                return;
            }

            if (showWarning && !string.IsNullOrWhiteSpace(resolve.Warning))
            {
                Send("status", new { text = resolve.Warning, level = "warn" });
            }
        }
        catch (Exception ex)
        {
            if (showWarning)
            {
                Send("status", new { text = $"加载报表信息失败: {ex.Message}", level = "warn" });
            }
        }
    }

    private static List<int> BuildPortCandidates(int? preferredPort, IReadOnlyList<PowerBiInstanceInfo> instances)
    {
        var result = new List<int>();
        if (preferredPort.HasValue)
        {
            result.Add(preferredPort.Value);
        }

        foreach (var instance in instances.OrderByDescending(i => i.LastSeenUtc))
        {
            if (!result.Contains(instance.Port))
            {
                result.Add(instance.Port);
            }
        }

        return result;
    }

    private static int? ParsePortFromServer(string? server)
    {
        if (string.IsNullOrWhiteSpace(server))
        {
            return null;
        }

        var trimmed = server.Trim().TrimEnd(';');
        var parts = trimmed.Split(':', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
        if (parts.Length == 2 && int.TryParse(parts[1], out var directPort))
        {
            return directPort;
        }

        if (int.TryParse(trimmed, out var rawPort))
        {
            return rawPort;
        }

        var allPorts = Regex.Matches(trimmed, "\\d{4,6}");
        if (allPorts.Count > 0 && int.TryParse(allPorts[^1].Value, out var matchedPort))
        {
            return matchedPort;
        }

        return null;
    }

    private static string? NormalizeExternalToolValue(string? value)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            return null;
        }

        var trimmed = value.Trim();
        return Regex.IsMatch(trimmed, "^%[^%]+%$") ? null : trimmed;
    }

    private (string Snapshot, string Rollback) CreateBackup(AbiActionPlan plan)
    {
        var before = _reader.ReadMetadata(_currentPort!.Value, _model!.DatabaseName);
        var dir = BackupRoot();
        var ts = DateTime.Now.ToString("yyyyMMdd-HHmmss");
        var modelKey = ResolveCurrentBackupModelKey(before.DatabaseName);
        var snapshot = Path.Combine(dir, $"{modelKey}-sess-{_sessionToken}-before-{ts}.json");
        var rollback = Path.Combine(dir, $"{modelKey}-sess-{_sessionToken}-rollback-{ts}.json");
        Directory.CreateDirectory(dir);
        File.WriteAllText(snapshot, JsonSerializer.Serialize(before, new JsonSerializerOptions { WriteIndented = true }));
        AbiActionPlanStorage.Save(rollback, BuildRollbackPlan(before, plan));
        return (snapshot, rollback);
    }

    private string? GetCurrentSessionBackupPrefix()
    {
        if (_model is null)
        {
            return null;
        }

        var key = ResolveCurrentBackupModelKey(_model.DatabaseName);
        return string.IsNullOrWhiteSpace(key) ? null : $"{key}-sess-{_sessionToken}-";
    }

    private string ResolveCurrentBackupModelKey(string fallback)
    {
        if (!string.IsNullOrWhiteSpace(_currentBackupModelKey))
        {
            return _currentBackupModelKey;
        }

        var source = string.IsNullOrWhiteSpace(_connectedModelDisplayName) ? fallback : _connectedModelDisplayName;
        _currentBackupModelKey = SafeName(source);
        return _currentBackupModelKey;
    }

    private static AbiActionPlan BuildRollbackPlan(ModelMetadata before, AbiActionPlan executed)
    {
        var actions = new List<AbiModelAction>();
        foreach (var a in executed.Actions.Reverse())
        {
            switch (a.Type.ToLowerInvariant())
            {
                case "create_or_update_measure":
                    var oldTable = FindTable(before, a.Table);
                    var old = FindMeasure(before, a.Table, a.Name);
                    actions.Add(old is null
                        ? new AbiModelAction("delete_measure", Table: a.Table, Name: a.Name, Reason: "回滚：删除新增度量值")
                        : new AbiModelAction("create_or_update_measure", Table: a.Table, Name: old.Name, Expression: old.Expression, FormatString: old.FormatString, IsHidden: old.IsHidden, Reason: "回滚：恢复原始度量值"));
                    break;
                case "delete_measure":
                    var del = FindMeasure(before, a.Table, a.Name);
                    if (del is not null)
                        actions.Add(new AbiModelAction("create_or_update_measure", Table: a.Table, Name: del.Name, Expression: del.Expression, FormatString: del.FormatString, IsHidden: del.IsHidden, Reason: "回滚：恢复已删除的度量值"));
                    break;
                case "rename_table":
                    if (!string.IsNullOrWhiteSpace(a.Table) && !string.IsNullOrWhiteSpace(a.NewName))
                    {
                        actions.Add(new AbiModelAction("rename_table", Table: a.NewName, NewName: a.Table, Reason: "回滚：恢复表名"));
                    }
                    break;
                case "rename_column":
                    if (!string.IsNullOrWhiteSpace(a.Table) && !string.IsNullOrWhiteSpace(a.Name) && !string.IsNullOrWhiteSpace(a.NewName))
                    {
                        actions.Add(new AbiModelAction("rename_column", Table: a.Table, Name: a.NewName, NewName: a.Name, Reason: "回滚：恢复列名"));
                    }
                    break;
                case "rename_measure":
                    if (!string.IsNullOrWhiteSpace(a.Table) && !string.IsNullOrWhiteSpace(a.Name) && !string.IsNullOrWhiteSpace(a.NewName))
                    {
                        actions.Add(new AbiModelAction("rename_measure", Table: a.Table, Name: a.NewName, NewName: a.Name, Reason: "回滚：恢复度量值名"));
                    }
                    break;
                case "set_table_hidden":
                    var oldTableVisibility = FindTable(before, a.Table);
                    if (oldTableVisibility is not null)
                    {
                        actions.Add(new AbiModelAction("set_table_hidden", Table: oldTableVisibility.Name, IsHidden: oldTableVisibility.IsHidden, Reason: "回滚：恢复表可见性"));
                    }
                    break;
                case "set_column_hidden":
                    var oldColumn = FindColumn(before, a.Table, a.Name);
                    if (oldColumn is not null)
                    {
                        actions.Add(new AbiModelAction("set_column_hidden", Table: a.Table, Name: oldColumn.Name, IsHidden: oldColumn.IsHidden, Reason: "回滚：恢复列可见性"));
                    }
                    break;
                case "set_measure_hidden":
                    var oldMeasureHidden = FindMeasure(before, a.Table, a.Name);
                    if (oldMeasureHidden is not null)
                    {
                        actions.Add(new AbiModelAction("set_measure_hidden", Table: a.Table, Name: oldMeasureHidden.Name, IsHidden: oldMeasureHidden.IsHidden, Reason: "回滚：恢复度量值可见性"));
                    }
                    break;
                case "set_format_string":
                    var oldMeasureFormat = FindMeasure(before, a.Table, a.Name);
                    if (oldMeasureFormat is not null)
                    {
                        actions.Add(new AbiModelAction("set_format_string", Table: a.Table, Name: oldMeasureFormat.Name, FormatString: oldMeasureFormat.FormatString, Reason: "回滚：恢复格式字符串"));
                    }
                    break;
                case "set_display_folder":
                    var oldMeasureFolder = FindMeasure(before, a.Table, a.Name);
                    if (oldMeasureFolder is not null)
                    {
                        actions.Add(new AbiModelAction("set_display_folder", Table: a.Table, Name: oldMeasureFolder.Name, DisplayFolder: oldMeasureFolder.DisplayFolder, Reason: "回滚：恢复显示文件夹"));
                    }
                    break;
                case "create_calculated_column":
                    if (FindColumn(before, a.Table, a.Name) is null)
                    {
                        actions.Add(new AbiModelAction("delete_column", Table: a.Table, Name: a.Name, Reason: "回滚：删除新增计算列"));
                    }
                    break;
                case "create_calculated_table":
                    if (FindTable(before, a.Name) is null)
                    {
                        actions.Add(new AbiModelAction("delete_table", Table: a.Name, Reason: "回滚：删除新增计算表"));
                    }
                    break;
                case "create_relationship":
                    if (!string.IsNullOrWhiteSpace(a.Name))
                    {
                        actions.Add(new AbiModelAction("delete_relationship", Name: a.Name, Reason: "回滚：删除新增关系"));
                    }
                    else
                    {
                        actions.Add(new AbiModelAction(
                            "delete_relationship",
                            FromTable: a.FromTable,
                            FromColumn: a.FromColumn,
                            ToTable: a.ToTable,
                            ToColumn: a.ToColumn,
                            Reason: "回滚：删除新增关系"));
                    }
                    break;
                case "delete_relationship":
                    var oldRel = FindRelationship(before, a);
                    if (oldRel is not null)
                    {
                        actions.Add(new AbiModelAction(
                            "create_relationship",
                            Name: oldRel.Name,
                            FromTable: oldRel.FromTable,
                            FromColumn: oldRel.FromColumn,
                            ToTable: oldRel.ToTable,
                            ToColumn: oldRel.ToColumn,
                            CrossFilterDirection: oldRel.CrossFilterDirection,
                            IsActive: oldRel.IsActive,
                            Reason: "回滚：恢复已删除关系"));
                    }
                    break;
                case "set_relationship_active":
                    var oldActiveRel = FindRelationship(before, a);
                    if (oldActiveRel is not null)
                    {
                        actions.Add(new AbiModelAction(
                            "set_relationship_active",
                            Name: oldActiveRel.Name,
                            FromTable: oldActiveRel.FromTable,
                            FromColumn: oldActiveRel.FromColumn,
                            ToTable: oldActiveRel.ToTable,
                            ToColumn: oldActiveRel.ToColumn,
                            IsActive: oldActiveRel.IsActive,
                            Reason: "回滚：恢复关系激活状态"));
                    }
                    break;
                case "set_relationship_cross_filter":
                    var oldDirectionRel = FindRelationship(before, a);
                    if (oldDirectionRel is not null)
                    {
                        actions.Add(new AbiModelAction(
                            "set_relationship_cross_filter",
                            Name: oldDirectionRel.Name,
                            FromTable: oldDirectionRel.FromTable,
                            FromColumn: oldDirectionRel.FromColumn,
                            ToTable: oldDirectionRel.ToTable,
                            ToColumn: oldDirectionRel.ToColumn,
                            CrossFilterDirection: oldDirectionRel.CrossFilterDirection,
                            Reason: "回滚：恢复关系筛选方向"));
                    }
                    break;
                case "create_role":
                    var roleBeforeCreate = FindRole(before, a.Name);
                    if (roleBeforeCreate is null)
                    {
                        actions.Add(new AbiModelAction("delete_role", Name: a.Name, Reason: "回滚：删除新增角色"));
                    }
                    else
                    {
                        AddRoleRestoreActions(actions, roleBeforeCreate, a.Name);
                    }
                    break;
                case "update_role":
                    var roleBeforeUpdate = FindRole(before, a.Name);
                    if (roleBeforeUpdate is not null)
                    {
                        var currentRoleName = string.IsNullOrWhiteSpace(a.NewName) ? a.Name : a.NewName;
                        AddRoleRestoreActions(actions, roleBeforeUpdate, currentRoleName);
                    }
                    break;
                case "delete_role":
                    var roleBeforeDelete = FindRole(before, a.Name);
                    if (roleBeforeDelete is not null)
                    {
                        AddRoleRestoreActions(actions, roleBeforeDelete, a.Name);
                    }
                    break;
                case "set_role_table_permission":
                    var oldPermission = FindRoleTablePermission(before, a.Name, a.Table);
                    if (oldPermission is null)
                    {
                        actions.Add(new AbiModelAction("remove_role_table_permission", Name: a.Name, Table: a.Table, Reason: "回滚：移除新增角色表权限"));
                    }
                    else
                    {
                        actions.Add(new AbiModelAction("set_role_table_permission", Name: a.Name, Table: oldPermission.TableName, Expression: oldPermission.FilterExpression, MetadataPermission: oldPermission.MetadataPermission, Reason: "回滚：恢复角色表权限"));
                    }
                    break;
                case "remove_role_table_permission":
                    var permissionBeforeRemove = FindRoleTablePermission(before, a.Name, a.Table);
                    if (permissionBeforeRemove is not null)
                    {
                        actions.Add(new AbiModelAction("set_role_table_permission", Name: a.Name, Table: permissionBeforeRemove.TableName, Expression: permissionBeforeRemove.FilterExpression, MetadataPermission: permissionBeforeRemove.MetadataPermission, Reason: "回滚：恢复角色表权限"));
                    }
                    break;
                case "add_role_member":
                    if (FindRoleMember(before, a.Name, a.MemberName) is null)
                    {
                        actions.Add(new AbiModelAction("remove_role_member", Name: a.Name, MemberName: a.MemberName, Reason: "回滚：移除新增角色成员"));
                    }
                    break;
                case "remove_role_member":
                    var removedMemberBefore = FindRoleMember(before, a.Name, a.MemberName);
                    if (removedMemberBefore is not null)
                    {
                        actions.Add(new AbiModelAction("add_role_member", Name: a.Name, MemberName: removedMemberBefore.Name, IdentityProvider: EmptyToNull(removedMemberBefore.IdentityProvider), MemberType: EmptyToNull(removedMemberBefore.MemberType), Reason: "回滚：恢复角色成员"));
                    }
                    break;
                case "update_description":
                    if (string.Equals(a.ObjectType, "role", StringComparison.OrdinalIgnoreCase))
                    {
                        var roleForDescription = FindRole(before, a.Name);
                        if (roleForDescription is not null)
                        {
                            actions.Add(new AbiModelAction("update_description", ObjectType: "role", Name: roleForDescription.Name, Description: roleForDescription.Description, Reason: "回滚：恢复角色描述"));
                        }
                    }
                    break;
            }
        }
        return new AbiActionPlan($"回滚计划生成于 {DateTime.Now:yyyy-MM-dd HH:mm:ss}", actions);
    }

    private static void AddRoleRestoreActions(ICollection<AbiModelAction> actions, RoleMetadata roleBefore, string? currentName)
    {
        var activeName = string.IsNullOrWhiteSpace(currentName) ? roleBefore.Name : currentName.Trim();
        actions.Add(new AbiModelAction("delete_role", Name: activeName, Reason: "回滚：重建角色"));
        actions.Add(new AbiModelAction("create_role", Name: roleBefore.Name, Description: roleBefore.Description, ModelPermission: roleBefore.ModelPermission, Reason: "回滚：恢复角色"));

        foreach (var permission in roleBefore.TablePermissions)
        {
            actions.Add(new AbiModelAction(
                "set_role_table_permission",
                Name: roleBefore.Name,
                Table: permission.TableName,
                Expression: permission.FilterExpression,
                MetadataPermission: permission.MetadataPermission,
                Reason: "回滚：恢复角色表权限"));
        }

        foreach (var member in roleBefore.Members)
        {
            actions.Add(new AbiModelAction(
                "add_role_member",
                Name: roleBefore.Name,
                MemberName: member.Name,
                IdentityProvider: EmptyToNull(member.IdentityProvider),
                MemberType: EmptyToNull(member.MemberType),
                Reason: "回滚：恢复角色成员"));
        }
    }

    private static string? EmptyToNull(string? value) =>
        string.IsNullOrWhiteSpace(value) ? null : value;

    private static TableMetadata? FindTable(ModelMetadata model, string? tableName) =>
        model.Tables.FirstOrDefault(t => string.Equals(t.Name, tableName, StringComparison.OrdinalIgnoreCase));

    private static MeasureMetadata? FindMeasure(ModelMetadata model, string? tableName, string? measureName)
    {
        if (string.IsNullOrWhiteSpace(measureName))
        {
            return null;
        }

        var inTable = FindTable(model, tableName)?.Measures
            .FirstOrDefault(m => string.Equals(m.Name, measureName, StringComparison.OrdinalIgnoreCase));
        if (inTable is not null)
        {
            return inTable;
        }

        return model.Tables
            .SelectMany(t => t.Measures)
            .FirstOrDefault(m => string.Equals(m.Name, measureName, StringComparison.OrdinalIgnoreCase));
    }

    private static ColumnMetadata? FindColumn(ModelMetadata model, string? tableName, string? columnName)
    {
        if (string.IsNullOrWhiteSpace(columnName))
        {
            return null;
        }

        return FindTable(model, tableName)?.Columns
            .FirstOrDefault(c => string.Equals(c.Name, columnName, StringComparison.OrdinalIgnoreCase));
    }

    private static RelationshipMetadata? FindRelationship(ModelMetadata model, AbiModelAction action)
    {
        if (!string.IsNullOrWhiteSpace(action.Name))
        {
            return model.Relationships.FirstOrDefault(r => string.Equals(r.Name, action.Name, StringComparison.OrdinalIgnoreCase));
        }

        return model.Relationships.FirstOrDefault(r =>
            string.Equals(r.FromTable, action.FromTable, StringComparison.OrdinalIgnoreCase) &&
            string.Equals(r.FromColumn, action.FromColumn, StringComparison.OrdinalIgnoreCase) &&
            string.Equals(r.ToTable, action.ToTable, StringComparison.OrdinalIgnoreCase) &&
            string.Equals(r.ToColumn, action.ToColumn, StringComparison.OrdinalIgnoreCase));
    }

    private static RoleMetadata? FindRole(ModelMetadata model, string? roleName)
    {
        if (string.IsNullOrWhiteSpace(roleName))
        {
            return null;
        }

        return model.Roles.FirstOrDefault(r => string.Equals(r.Name, roleName, StringComparison.OrdinalIgnoreCase));
    }

    private static RoleMemberMetadata? FindRoleMember(ModelMetadata model, string? roleName, string? memberName)
    {
        if (string.IsNullOrWhiteSpace(memberName))
        {
            return null;
        }

        return FindRole(model, roleName)?.Members
            .FirstOrDefault(m => string.Equals(m.Name, memberName, StringComparison.OrdinalIgnoreCase));
    }

    private static RoleTablePermissionMetadata? FindRoleTablePermission(ModelMetadata model, string? roleName, string? tableName)
    {
        if (string.IsNullOrWhiteSpace(tableName))
        {
            return null;
        }

        return FindRole(model, roleName)?.TablePermissions
            .FirstOrDefault(p => string.Equals(p.TableName, tableName, StringComparison.OrdinalIgnoreCase));
    }

    private static string BackupRoot() =>
        Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "PBIClaw", "backups");

    private static string SafeName(string name) =>
        System.Text.RegularExpressions.Regex.Replace(name, @"[^\w\-]", "_");

    private string ResolvePbiDisplayName(int port, string fallbackName)
    {
        try
        {
            var hit = _detector.DiscoverInstances().FirstOrDefault(i => i.Port == port);
            if (hit is not null && !string.IsNullOrWhiteSpace(hit.PbixPathHint))
            {
                var fileName = Path.GetFileNameWithoutExtension(hit.PbixPathHint);
                if (!string.IsNullOrWhiteSpace(fileName))
                {
                    return fileName;
                }
            }
        }
        catch
        {
            // Ignore detection errors and use fallback name.
        }

        return fallbackName;
    }

    // ── DTOs ─────────────────────────────────────────────────────────────────

    private static object InstanceDto(PowerBiInstanceInfo i) => new
    {
        port = i.Port,
        label = string.IsNullOrWhiteSpace(i.PbixPathHint)
            ? (i.DesktopPid > 0
                ? $"Power BI Desktop (PID {i.DesktopPid})"
                : "Power BI Desktop 缓存工作区")
            : Path.GetFileNameWithoutExtension(i.PbixPathHint)
    };

    private static object TabularDatabaseDto(TabularDatabaseInfo db) => new
    {
        name = db.Name,
        id = db.Id,
        compatibilityLevel = db.CompatibilityLevel
    };

    private static object ModelDto(ModelMetadata m, string? databaseDisplayName = null, ReportMetadata? report = null) => new
    {
        databaseName = m.DatabaseName,
        databaseDisplayName = string.IsNullOrWhiteSpace(databaseDisplayName) ? m.DatabaseName : databaseDisplayName,
        tables = m.Tables.Select(t => new
        {
            name = t.Name,
            isHidden = t.IsHidden,
            tableType = t.TableType,
            expression = t.Expression,
            sourceType = t.SourceType,
            sourceExpression = t.SourceExpression,
            dataSourceName = t.DataSourceName,
            sourceSystemType = t.SourceSystemType,
            sourceServer = t.SourceServer,
            sourceDatabase = t.SourceDatabase,
            sourceSchema = t.SourceSchema,
            sourceObjectName = t.SourceObjectName,
            measures = t.Measures.Select(ms => new
            {
                name = ms.Name,
                expression = ms.Expression,
                formatString = ms.FormatString,
                isHidden = ms.IsHidden,
                displayFolder = ms.DisplayFolder
            }).ToArray(),
            columns = t.Columns.Select(c => new
            {
                name = c.Name,
                columnType = c.ColumnType,
                dataType = c.DataType,
                isHidden = c.IsHidden,
                expression = c.Expression
            }).ToArray()
        }).ToArray(),
        relationships = m.Relationships.Select(r => new
        {
            name = r.Name,
            fromTable = r.FromTable,
            fromColumn = r.FromColumn,
            toTable = r.ToTable,
            toColumn = r.ToColumn,
            crossFilterDirection = r.CrossFilterDirection,
            isActive = r.IsActive
        }).ToArray(),
        roles = m.Roles.Select(role => new
        {
            name = role.Name,
            description = role.Description,
            modelPermission = role.ModelPermission,
            members = role.Members.Select(member => new
            {
                name = member.Name,
                identityProvider = member.IdentityProvider,
                memberType = member.MemberType
            }).ToArray(),
            tablePermissions = role.TablePermissions.Select(permission => new
            {
                tableName = permission.TableName,
                filterExpression = permission.FilterExpression,
                metadataPermission = permission.MetadataPermission
            }).ToArray()
        }).ToArray(),
        report = report is null ? null : new
        {
            sourceType = report.SourceType,
            sourcePath = report.SourcePath,
            pages = report.Pages.Select(page => new
            {
                name = page.Name,
                displayName = page.DisplayName,
                visuals = page.Visuals.Select(visual => new
                {
                    name = visual.Name,
                    visualType = visual.VisualType,
                    title = visual.Title
                }).ToArray()
            }).ToArray()
        }
    };

    private (ReportMetadata? Report, string? Warning) ResolveReportMetadataForPort(int port)
    {
        try
        {
            var hit = _detector.DiscoverInstances().FirstOrDefault(i => i.Port == port);
            if (hit is null)
            {
                return (null, null);
            }

            if (string.IsNullOrWhiteSpace(hit.PbixPathHint))
            {
                return (null, "未找到 PBIX/PBIP 路径，无法加载报表信息。");
            }

            var report = _reportReader.TryRead(hit.PbixPathHint);
            if (report is null)
            {
                return (null, $"未能解析报表信息：{Path.GetFileName(hit.PbixPathHint)}");
            }

            return (report, null);
        }
        catch
        {
            return (null, "加载报表信息时发生异常。");
        }
    }

    private static object ActionDto(AbiModelAction a) => new
    {
        type = a.Type,
        table = a.Table,
        name = a.Name,
        newName = a.NewName,
        objectType = a.ObjectType,
        memberName = a.MemberName,
        description = a.Description,
        fromTable = a.FromTable,
        fromColumn = a.FromColumn,
        toTable = a.ToTable,
        toColumn = a.ToColumn,
        isHidden = a.IsHidden,
        formatString = a.FormatString,
        displayFolder = a.DisplayFolder,
        modelPermission = a.ModelPermission,
        metadataPermission = a.MetadataPermission,
        reason = a.Reason
    };

    private static object PlanDto(
        AbiActionPlan plan,
        string? preview = null,
        IReadOnlyList<string>? actionStates = null) => new
    {
        summary = plan.Summary,
        preview = string.IsNullOrWhiteSpace(preview) ? AbiActionPlanPreview.BuildText(plan) : preview,
        actions = plan.Actions.Select(ActionDto).ToArray(),
        actionStates = actionStates?.ToArray()
    };

    private static object SettingsDto(AbiAssistantSettings s) => new
    {
        provider = s.Provider,
        baseUrl = s.BaseUrl,
        model = s.Model,
        apiKey = s.ApiKey,
        temperature = s.Temperature,
        customSystemPrompt = s.CustomSystemPrompt,
        allowModelChanges = s.AllowModelChanges,
        includeHiddenObjects = s.IncludeHiddenObjects,
        quickPrompts = s.QuickPrompts.ToArray(),
        showStartupGuide = s.ShowStartupGuide
    };

    private static AbiAssistantSettings BuildSettingsFromPayload(JsonObject? p, AbiAssistantSettings current)
    {
        var next = new AbiAssistantSettings
        {
            Provider = NormalizeProvider(p?["provider"]?.GetValue<string>() ?? current.Provider),
            BaseUrl = p?["baseUrl"]?.GetValue<string>()?.Trim() ?? current.BaseUrl,
            Model = p?["model"]?.GetValue<string>()?.Trim() ?? current.Model,
            ApiKey = p?["apiKey"]?.GetValue<string>()?.Trim() ?? current.ApiKey,
            Temperature = p?["temperature"]?.GetValue<double>() ?? current.Temperature,
            CustomSystemPrompt = p?["customSystemPrompt"]?.GetValue<string>()?.Trim() ?? current.CustomSystemPrompt,
            AllowModelChanges = p?["allowModelChanges"]?.GetValue<bool>() ?? current.AllowModelChanges,
            IncludeHiddenObjects = p?["includeHiddenObjects"]?.GetValue<bool>() ?? current.IncludeHiddenObjects,
            ShowStartupGuide = p?["showStartupGuide"]?.GetValue<bool>() ?? current.ShowStartupGuide,
            QuickPrompts = p?["quickPrompts"]?.AsArray()
                .Select(n => n?.GetValue<string>() ?? string.Empty)
                .Where(s => !string.IsNullOrWhiteSpace(s))
                .Distinct(StringComparer.Ordinal)
                .Take(12)
                .ToList() ?? [.. current.QuickPrompts]
        };

        if (next.QuickPrompts.Count == 0)
            next.QuickPrompts = ["分析当前模型并给出3个优化建议和可执行步骤。"];

        return next;
    }

    private static string NormalizeProvider(string? provider)
    {
        if (string.Equals(provider, "anthropic", StringComparison.OrdinalIgnoreCase))
        {
            return "anthropic";
        }

        return "openai";
    }

    private static string CurrentAppVersion()
    {
        var productVersion = Application.ProductVersion;
        if (!string.IsNullOrWhiteSpace(productVersion))
        {
            return productVersion.Trim();
        }

        var asmVersion = Assembly.GetExecutingAssembly().GetName().Version;
        if (asmVersion is not null)
        {
            return asmVersion.ToString();
        }

        return "1.0.0";
    }

    private async Task<ReleaseInfo> GetLatestReleaseAsync()
    {
        Exception? listApiError = null;
        try
        {
            return await GetLatestReleaseFromReleasesApiAsync().ConfigureAwait(false);
        }
        catch (Exception ex)
        {
            listApiError = ex;
        }

        Exception? latestApiError = null;
        try
        {
            return await GetLatestReleaseFromApiAsync().ConfigureAwait(false);
        }
        catch (Exception ex)
        {
            latestApiError = ex;
        }

        try
        {
            return await GetLatestReleaseFromLatestPageAsync().ConfigureAwait(false);
        }
        catch (Exception fallbackEx)
        {
            throw new InvalidOperationException(
                "检测更新失败（GitHub API 与回退页面都不可用）。" +
                $"releases接口: {listApiError?.Message}；latest接口: {latestApiError?.Message}；回退页面: {fallbackEx.Message}");
        }
    }

    private async Task<ReleaseInfo> GetLatestReleaseFromReleasesApiAsync()
    {
        using var request = new HttpRequestMessage(HttpMethod.Get, GithubReleasesApiUrl);
        request.Headers.Accept.Clear();
        request.Headers.Accept.Add(new MediaTypeWithQualityHeaderValue("application/vnd.github+json"));
        using var response = await UpdateHttpClient.SendAsync(request).ConfigureAwait(false);
        var content = await response.Content.ReadAsStringAsync().ConfigureAwait(false);
        if (!response.IsSuccessStatusCode)
        {
            throw new InvalidOperationException($"检测更新失败 ({(int)response.StatusCode}): {content}");
        }

        using var doc = JsonDocument.Parse(content);
        if (doc.RootElement.ValueKind != JsonValueKind.Array)
        {
            throw new InvalidOperationException("releases 返回格式异常。");
        }

        ReleaseInfo? best = null;
        Version? bestVersion = null;
        foreach (var release in doc.RootElement.EnumerateArray())
        {
            if (release.TryGetProperty("draft", out var draftNode) &&
                draftNode.ValueKind == JsonValueKind.True)
            {
                continue;
            }

            if (release.TryGetProperty("prerelease", out var prereleaseNode) &&
                prereleaseNode.ValueKind == JsonValueKind.True)
            {
                continue;
            }

            var info = ParseReleaseInfo(release);
            if (string.IsNullOrWhiteSpace(info.Version) ||
                string.Equals(info.Version, "unknown", StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            if (!TryParseVersion(info.Version, out var parsed))
            {
                if (best is null)
                {
                    best = info;
                }
                continue;
            }

            if (best is null || bestVersion is null || parsed > bestVersion)
            {
                best = info;
                bestVersion = parsed;
            }
        }

        if (best is null)
        {
            throw new InvalidOperationException("releases 列表中未找到有效版本。");
        }

        return best;
    }

    private async Task<ReleaseInfo> GetLatestReleaseFromApiAsync()
    {
        using var request = new HttpRequestMessage(HttpMethod.Get, GithubLatestReleaseApiUrl);
        request.Headers.Accept.Clear();
        request.Headers.Accept.Add(new MediaTypeWithQualityHeaderValue("application/vnd.github+json"));
        using var response = await UpdateHttpClient.SendAsync(request).ConfigureAwait(false);
        var content = await response.Content.ReadAsStringAsync().ConfigureAwait(false);
        if (!response.IsSuccessStatusCode)
        {
            throw new InvalidOperationException($"检测更新失败 ({(int)response.StatusCode}): {content}");
        }

        using var doc = JsonDocument.Parse(content);
        return ParseReleaseInfo(doc.RootElement);
    }

    private async Task<ReleaseInfo> GetLatestReleaseFromLatestPageAsync()
    {
        using var request = new HttpRequestMessage(HttpMethod.Get, GithubLatestReleasePageUrl);
        request.Headers.Accept.Clear();
        request.Headers.Accept.Add(new MediaTypeWithQualityHeaderValue("text/html"));
        using var response = await UpdateHttpClient.SendAsync(request).ConfigureAwait(false);
        if (!response.IsSuccessStatusCode)
        {
            var content = await response.Content.ReadAsStringAsync().ConfigureAwait(false);
            throw new InvalidOperationException($"回退检测失败 ({(int)response.StatusCode}): {content}");
        }

        var finalUri = response.RequestMessage?.RequestUri;
        var releaseUrl = finalUri?.ToString() ?? ReleasePageUrl;
        var tagCandidate = string.Empty;

        if (finalUri is not null)
        {
            var absPath = finalUri.AbsolutePath ?? string.Empty;
            var marker = "/releases/tag/";
            var idx = absPath.IndexOf(marker, StringComparison.OrdinalIgnoreCase);
            if (idx >= 0)
            {
                tagCandidate = absPath[(idx + marker.Length)..];
                tagCandidate = Uri.UnescapeDataString(tagCandidate);
            }
        }

        if (string.IsNullOrWhiteSpace(tagCandidate))
        {
            var html = await response.Content.ReadAsStringAsync().ConfigureAwait(false);
            var match = Regex.Match(
                html,
                "/yczx001/PBIClaw/releases/tag/(?<tag>[^\\\"#?]+)",
                RegexOptions.IgnoreCase);
            if (match.Success)
            {
                tagCandidate = Uri.UnescapeDataString(match.Groups["tag"].Value);
                releaseUrl = "https://github.com/yczx001/PBIClaw/releases/tag/" + tagCandidate;
            }
        }

        var version = NormalizeReleaseVersion(tagCandidate, string.Empty);
        if (string.IsNullOrWhiteSpace(version) || string.Equals(version, "unknown", StringComparison.OrdinalIgnoreCase))
        {
            throw new InvalidOperationException("回退检测未能解析最新版本号。");
        }

        return new ReleaseInfo(version, releaseUrl, string.Empty, string.Empty, string.Empty);
    }

    private static HttpClient CreateUpdateHttpClient()
    {
        var client = new HttpClient();
        client.Timeout = TimeSpan.FromSeconds(30);
        client.DefaultRequestHeaders.UserAgent.ParseAdd("PBIClaw-Updater/1.0");
        client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
        return client;
    }

    private static string GetJsonString(JsonElement element, string propertyName)
    {
        if (element.TryGetProperty(propertyName, out var property) && property.ValueKind == JsonValueKind.String)
        {
            return property.GetString() ?? string.Empty;
        }

        return string.Empty;
    }

    private static ReleaseInfo ParseReleaseInfo(JsonElement root)
    {
        var tagName = GetJsonString(root, "tag_name");
        var releaseName = GetJsonString(root, "name");
        var customVersion = GetJsonString(root, "version");
        var version = NormalizeReleaseVersion(tagName, releaseName);
        if (string.IsNullOrWhiteSpace(version) || string.Equals(version, "unknown", StringComparison.OrdinalIgnoreCase))
        {
            version = NormalizeReleaseVersion(customVersion, string.Empty);
        }

        var releaseUrl = GetJsonString(root, "html_url");
        if (string.IsNullOrWhiteSpace(releaseUrl))
        {
            releaseUrl = GetJsonString(root, "releaseUrl");
        }
        if (string.IsNullOrWhiteSpace(releaseUrl))
        {
            releaseUrl = ReleasePageUrl;
        }

        var summary = GetJsonString(root, "body");
        if (string.IsNullOrWhiteSpace(summary))
        {
            summary = GetJsonString(root, "summary");
        }
        if (!string.IsNullOrWhiteSpace(summary) && summary.Length > 500)
        {
            summary = summary[..500] + "...";
        }

        var publishedAt = GetJsonString(root, "published_at");
        if (string.IsNullOrWhiteSpace(publishedAt))
        {
            publishedAt = GetJsonString(root, "publishedAt");
        }

        var downloadUrl = string.Empty;
        if (root.TryGetProperty("assets", out var assets) && assets.ValueKind == JsonValueKind.Array)
        {
            foreach (var asset in assets.EnumerateArray())
            {
                var assetName = GetJsonString(asset, "name");
                if (!assetName.EndsWith(".exe", StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                if (!assetName.Contains("setup", StringComparison.OrdinalIgnoreCase) &&
                    !assetName.Contains("installer", StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                downloadUrl = GetJsonString(asset, "browser_download_url");
                if (!string.IsNullOrWhiteSpace(downloadUrl))
                {
                    break;
                }
            }
        }
        if (string.IsNullOrWhiteSpace(downloadUrl))
        {
            downloadUrl = GetJsonString(root, "downloadUrl");
        }

        return new ReleaseInfo(version, releaseUrl, downloadUrl, publishedAt, summary ?? string.Empty);
    }

    private static string NormalizeReleaseVersion(string tagName, string releaseName)
    {
        static string Clean(string value)
        {
            var v = (value ?? string.Empty).Trim().TrimStart('v', 'V');
            var match = Regex.Match(v, @"\d+(?:\.\d+){1,3}");
            return match.Success ? match.Value : v;
        }

        var byTag = Clean(tagName);
        if (!string.IsNullOrWhiteSpace(byTag))
        {
            return byTag;
        }

        var byName = Clean(releaseName);
        if (!string.IsNullOrWhiteSpace(byName))
        {
            return byName;
        }

        return "unknown";
    }

    private static int CompareVersion(string left, string right)
    {
        if (TryParseVersion(left, out var lv) && TryParseVersion(right, out var rv))
        {
            return lv.CompareTo(rv);
        }

        return string.Compare(left, right, StringComparison.OrdinalIgnoreCase);
    }

    private static bool TryParseVersion(string value, out Version version)
    {
        version = new Version(0, 0);
        var normalized = NormalizeReleaseVersion(value, string.Empty);
        return Version.TryParse(normalized, out version);
    }

    private sealed class PlanHistoryEntry
    {
        public DateTime TimeUtc { get; set; }
        public string Kind { get; set; } = string.Empty;
        public string Model { get; set; } = string.Empty;
        public string Title { get; set; } = string.Empty;
        public string Detail { get; set; } = string.Empty;
        public int ActionCount { get; set; }
        public bool? Success { get; set; }
    }

    private sealed record ReleaseInfo(
        string Version,
        string ReleaseUrl,
        string DownloadUrl,
        string PublishedAt,
        string Summary);
}
