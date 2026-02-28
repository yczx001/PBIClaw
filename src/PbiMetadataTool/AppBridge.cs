using System.Text.Json;
using System.Text.Json.Nodes;

namespace PbiMetadataTool;

/// <summary>
/// Handles bidirectional messaging between the WebView2 frontend and C# backend.
/// JS → C#: window.chrome.webview.postMessage(json)
/// C# → JS: webView.CoreWebView2.PostWebMessageAsJson(json)
/// </summary>
internal sealed class AppBridge
{
    private const int MaxHistoryEntries = 500;
    private static readonly JsonSerializerOptions HistoryJsonOptions = new() { WriteIndented = true };

    private readonly MainFormWebView _form;
    private readonly PowerBiInstanceDetector _detector = new();
    private readonly TabularMetadataReader _reader = new();
    private readonly TabularModelWriter _writer = new();
    private readonly AiChatClient _chatClient = new();
    private readonly AbiSettingsStore _settingsStore = new();
    private readonly List<AiChatMessage> _conversation = [];
    private readonly List<PlanHistoryEntry> _planHistory = [];
    private readonly string _historyPath =
        Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "PBIClaw", "change-history.json");

    private AbiAssistantSettings _settings = new();
    private ModelMetadata? _model;
    private int? _currentPort;
    private string _lastTabularServer = string.Empty;
    private AbiActionPlan? _pendingPlan;
    private AbiActionPlan? _preflightPlan; // plan waiting for user confirmation
    private bool _preflightIsRollback;
    private CancellationTokenSource? _chatCts;
    private string _connectedModelDisplayName = string.Empty;

    public AppBridge(MainFormWebView form)
    {
        _form = form;
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
                case "windowControl": OnWindowControl(payload); break;
                case "chat":        _ = OnChatAsync(payload); break;
                case "testConnection": _ = OnTestConnectionAsync(payload); break;
                case "discardPlan": OnDiscardPlan(); break;
                case "executePlan": OnExecutePlan(payload); break;
                case "confirmExecute": OnConfirmExecute(); break;
                case "executeRollback": OnExecuteRollback(payload); break;
                case "refreshBackups":  OnRefreshBackups(); break;
                case "openBackupFolder": OnOpenBackupFolder(); break;
                case "saveSettings": OnSaveSettings(payload); break;
                case "markStartupGuideSeen": OnMarkStartupGuideSeen(); break;
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
        var instances = _detector.DiscoverInstances();
        Send("init", new
        {
            settings = SettingsDto(_settings),
            instances = instances.Select(InstanceDto).ToArray()
        });
        OnRefreshBackups();
        SendPlanHistory();
    }

    private void OnScan()
    {
        try
        {
            var list = _detector.DiscoverInstances();
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
                _lastTabularServer = server;
                var connStr = $"DataSource={server};";
                _model = _reader.ReadMetadata(connStr, null);
                _currentPort = null;
                _connectedModelDisplayName = _model.DatabaseName;
            }
            else
            {
                var portStr = p["port"]?.GetValue<string>() ?? string.Empty;
                if (!int.TryParse(portStr, out var port))
                {
                    Send("status", new { text = "端口无效", level = "warn" });
                    return;
                }
                _model = _reader.ReadMetadata(port, null);
                _currentPort = port;
                _connectedModelDisplayName = ResolvePbiDisplayName(port, _model.DatabaseName);
            }

            if (allowWrite != _settings.AllowModelChanges)
            {
                _settings.AllowModelChanges = allowWrite;
                _settingsStore.Save(_settings);
            }

            _conversation.Clear();
            Send("connected", new
            {
                dbName = _connectedModelDisplayName,
                allowChanges = _settings.AllowModelChanges,
                model = ModelDto(_model, _connectedModelDisplayName)
            });
            OnRefreshBackups();
        }
        catch (Exception ex)
        {
            Send("status", new { text = $"连接失败: {ex.Message}", level = "error" });
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
        _chatCts = new CancellationTokenSource();

        try
        {
            var runtimeSettings = BuildSettingsFromPayload(p["runtimeSettings"] as JsonObject, _settings);
            var msgs = new List<AiChatMessage>
            {
                new("system", MetadataPromptBuilder.BuildSystemPrompt(runtimeSettings)),
                new("system", MetadataPromptBuilder.BuildModelContext(_model, runtimeSettings.IncludeHiddenObjects))
            };
            var referencedMeasureContext = MetadataPromptBuilder.BuildReferencedMeasureContext(_model, prompt, runtimeSettings.IncludeHiddenObjects);
            if (!string.IsNullOrWhiteSpace(referencedMeasureContext))
            {
                msgs.Add(new AiChatMessage("system", referencedMeasureContext));
            }
            msgs.AddRange(_conversation.TakeLast(10));
            msgs.Add(new AiChatMessage("user", prompt));

            var reply = await _chatClient.CompleteAsync(new AiChatRequest(runtimeSettings, msgs), _chatCts.Token);

            _conversation.Add(new AiChatMessage("user", prompt));
            _conversation.Add(new AiChatMessage("assistant", reply));
            if (_conversation.Count > 40) _conversation.RemoveRange(0, _conversation.Count - 40);

            Send("chatReply", new { text = reply });

            if (runtimeSettings.AllowModelChanges && AbiActionPlanParser.TryExtract(reply, out var plan, out var preview, out _))
            {
                _pendingPlan = plan;
                Send("planDetected", new
                {
                    summary = plan.Summary,
                    preview,
                    actions = plan.Actions.Select(ActionDto).ToArray()
                });

                AppendPlanHistory(
                    kind: "plan_detected",
                    title: $"检测到变更计划（{plan.Actions.Count} 项）",
                    detail: string.IsNullOrWhiteSpace(plan.Summary) ? preview : plan.Summary,
                    actionCount: plan.Actions.Count);
            }
        }
        catch (OperationCanceledException) { /* cancelled */ }
        catch (Exception ex)
        {
            Send("chatError", new { error = ex.Message });
        }
    }

    private async Task OnTestConnectionAsync(JsonObject p)
    {
        try
        {
            var runtime = BuildSettingsFromPayload(p, _settings);
            using var cts = new CancellationTokenSource(TimeSpan.FromSeconds(20));
            await _chatClient.TestConnectionAsync(runtime, cts.Token);
            Send("status", new { text = "连接测试成功", level = "success" });
        }
        catch (Exception ex)
        {
            Send("status", new { text = $"连接测试失败: {ex.Message}", level = "error" });
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

        var indices = p["selectedIndices"]?.AsArray()
            .Select(n => n?.GetValue<int>() ?? -1)
            .Where(i => i >= 0 && i < _pendingPlan.Actions.Count)
            .OrderBy(i => i)
            .ToList() ?? [];

        if (indices.Count == 0)
        {
            Send("status", new { text = "请至少选择一个动作", level = "warn" });
            return;
        }

        var autoExecute = p["autoExecute"]?.GetValue<bool>() ?? false;
        var selectedActions = indices.Select(i => _pendingPlan.Actions[i]).ToList();
        _preflightPlan = new AbiActionPlan(_pendingPlan.Summary, selectedActions);
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
        var isRollback = _preflightIsRollback;
        try
        {
            var (snapshot, rollback) = CreateBackup(currentPlan);
            Send("log", new { text = $"执行前已创建备份:\n- {snapshot}\n- {rollback}" });

            var results = _writer.ApplyActions(_currentPort.Value, _model.DatabaseName, currentPlan);
            _pendingPlan = null;
            _preflightPlan = null;
            _preflightIsRollback = false;

            // Refresh model
            _model = _reader.ReadMetadata(_currentPort.Value, _model.DatabaseName);
            Send("connected", new
            {
                dbName = string.IsNullOrWhiteSpace(_connectedModelDisplayName) ? _model.DatabaseName : _connectedModelDisplayName,
                allowChanges = _settings.AllowModelChanges,
                model = ModelDto(_model, _connectedModelDisplayName)
            });

            Send("executeSuccess", new { results = results.ToArray() });
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
            if (plan.Actions.Count == 0) return;

            var analysis = _writer.AnalyzeActions(_currentPort.Value, _model.DatabaseName, plan);
            _preflightPlan = plan;
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
                message = "回滚文件不包含可执行动作。请使用新版生成的回滚文件，或确认该计划中确实存在可回滚项。";
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
        var files = Directory.EnumerateFiles(root, "*.json", SearchOption.AllDirectories)
            .Select(p => new FileInfo(p))
            .OrderByDescending(f => f.LastWriteTimeUtc)
            .Select(f => new
            {
                time = f.LastWriteTime.ToString("yyyy-MM-dd HH:mm:ss"),
                type = f.Name.Contains("-rollback-", StringComparison.OrdinalIgnoreCase) ? "回滚"
                     : f.Name.Contains("-before-", StringComparison.OrdinalIgnoreCase) ? "快照" : "其他",
                name = f.Name,
                size = $"{Math.Max(1, f.Length / 1024)} KB",
                dir = f.DirectoryName ?? string.Empty,
                path = f.FullName
            })
            .ToArray();
        Send("backups", new { list = files });
    }

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
        var list = _planHistory
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
        var logPath = Path.Combine(Path.GetTempPath(), "abi_scan_debug.txt");
        File.AppendAllText(logPath, $"[{DateTime.Now:HH:mm:ss}] Send({type}) via ExecuteScript\n");
        // Encode as base64 to safely pass arbitrary JSON into JS without escaping issues
        var b64 = Convert.ToBase64String(System.Text.Encoding.UTF8.GetBytes(json));
        _form.ExecuteScript($"try{{App.handle(JSON.parse(atob('{b64}')))}}catch(e){{console.error('Send error',e)}}");
    }

    // ── Helpers ──────────────────────────────────────────────────────────────

    private (string Snapshot, string Rollback) CreateBackup(AbiActionPlan plan)
    {
        var before = _reader.ReadMetadata(_currentPort!.Value, _model!.DatabaseName);
        var dir = BackupRoot();
        var ts = DateTime.Now.ToString("yyyyMMdd-HHmmss");
        var safe = SafeName(before.DatabaseName);
        var snapshot = Path.Combine(dir, $"{safe}-{_currentPort}-before-{ts}.json");
        var rollback = Path.Combine(dir, $"{safe}-{_currentPort}-rollback-{ts}.json");
        Directory.CreateDirectory(dir);
        File.WriteAllText(snapshot, JsonSerializer.Serialize(before, new JsonSerializerOptions { WriteIndented = true }));
        AbiActionPlanStorage.Save(rollback, BuildRollbackPlan(before, plan));
        return (snapshot, rollback);
    }

    private static AbiActionPlan BuildRollbackPlan(ModelMetadata before, AbiActionPlan executed)
    {
        var actions = new List<AbiModelAction>();
        foreach (var a in executed.Actions)
        {
            switch (a.Type.ToLowerInvariant())
            {
                case "create_or_update_measure":
                    var oldTable = before.Tables.FirstOrDefault(t =>
                        string.Equals(t.Name, a.Table, StringComparison.OrdinalIgnoreCase));
                    var old = oldTable?.Measures.FirstOrDefault(m =>
                        string.Equals(m.Name, a.Name, StringComparison.OrdinalIgnoreCase))
                        ?? before.Tables.SelectMany(t => t.Measures).FirstOrDefault(m =>
                            string.Equals(m.Name, a.Name, StringComparison.OrdinalIgnoreCase));
                    actions.Add(old is null
                        ? new AbiModelAction("delete_measure", Table: a.Table, Name: a.Name, Reason: "回滚：删除新增度量值")
                        : new AbiModelAction("create_or_update_measure", Table: a.Table, Name: old.Name, Expression: old.Expression, FormatString: old.FormatString, IsHidden: old.IsHidden, Reason: "回滚：恢复原始度量值"));
                    break;
                case "delete_measure":
                    var delTable = before.Tables.FirstOrDefault(t =>
                        string.Equals(t.Name, a.Table, StringComparison.OrdinalIgnoreCase));
                    var del = delTable?.Measures.FirstOrDefault(m =>
                        string.Equals(m.Name, a.Name, StringComparison.OrdinalIgnoreCase))
                        ?? before.Tables.SelectMany(t => t.Measures).FirstOrDefault(m =>
                            string.Equals(m.Name, a.Name, StringComparison.OrdinalIgnoreCase));
                    if (del is not null)
                        actions.Add(new AbiModelAction("create_or_update_measure", Table: a.Table, Name: del.Name, Expression: del.Expression, FormatString: del.FormatString, IsHidden: del.IsHidden, Reason: "回滚：恢复已删除的度量值"));
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
                    var oldRel = before.Relationships.FirstOrDefault(r =>
                        (!string.IsNullOrWhiteSpace(a.Name) &&
                         string.Equals(r.Name, a.Name, StringComparison.OrdinalIgnoreCase))
                        ||
                        (string.Equals(r.FromTable, a.FromTable, StringComparison.OrdinalIgnoreCase) &&
                         string.Equals(r.FromColumn, a.FromColumn, StringComparison.OrdinalIgnoreCase) &&
                         string.Equals(r.ToTable, a.ToTable, StringComparison.OrdinalIgnoreCase) &&
                         string.Equals(r.ToColumn, a.ToColumn, StringComparison.OrdinalIgnoreCase)));
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
            }
        }
        return new AbiActionPlan($"回滚计划生成于 {DateTime.Now:yyyy-MM-dd HH:mm:ss}", actions);
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

    private static object ModelDto(ModelMetadata m, string? databaseDisplayName = null) => new
    {
        databaseName = m.DatabaseName,
        databaseDisplayName = string.IsNullOrWhiteSpace(databaseDisplayName) ? m.DatabaseName : databaseDisplayName,
        tables = m.Tables.Select(t => new
        {
            name = t.Name,
            isHidden = t.IsHidden,
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
                isHidden = c.IsHidden
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
        }).ToArray()
    };

    private static object ActionDto(AbiModelAction a) => new
    {
        type = a.Type,
        table = a.Table,
        name = a.Name,
        fromTable = a.FromTable,
        fromColumn = a.FromColumn,
        toTable = a.ToTable,
        toColumn = a.ToColumn,
        reason = a.Reason
    };

    private static object SettingsDto(AbiAssistantSettings s) => new
    {
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
}
