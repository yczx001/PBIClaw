using System.Text.Json;
using System.Runtime.InteropServices;
using System.Reflection;
using Microsoft.Web.WebView2.Core;
using Microsoft.Web.WebView2.WinForms;

namespace PBIClawSetup;

internal sealed class SetupWizardForm : Form
{
    private const int WmNclbuttondown = 0xA1;
    private const int HtCaption = 2;

    private enum SetupStep
    {
        Welcome = 1,
        Directory = 2,
        Confirm = 3,
        Installing = 4,
        Complete = 5
    }

    private readonly WebView2 _webView;
    private readonly IReadOnlyList<string> _externalToolDirs;
    private readonly bool _autoInstallOnReady;
    private SetupStep _step = SetupStep.Welcome;
    private string _installDir;
    private InstallResult? _result;
    private bool _autoInstallStarted;
    private bool _installInProgress;
    private bool _lastInstallFailed;

    public int ExitCode { get; private set; } = 1;

    public SetupWizardForm(string? initialInstallDir = null, bool autoInstall = false)
    {
        Text = "PBI Claw 安装向导";
        StartPosition = FormStartPosition.CenterScreen;
        MinimumSize = new Size(920, 620);
        ClientSize = new Size(980, 680);
        FormBorderStyle = FormBorderStyle.None;
        BackColor = Color.FromArgb(15, 17, 23);

        _externalToolDirs = InstallerEngine.GetMachineExternalToolDirs();
        _installDir = InstallerEngine.GetDefaultInstallDir();
        _autoInstallOnReady = autoInstall;
        if (!string.IsNullOrWhiteSpace(initialInstallDir))
        {
            _installDir = initialInstallDir.Trim();
        }

        _webView = new WebView2 { Dock = DockStyle.Fill };
        Controls.Add(_webView);

        _ = InitializeWebAsync();
    }

    protected override void OnFormClosing(FormClosingEventArgs e)
    {
        if (_step == SetupStep.Installing)
        {
            e.Cancel = true;
            return;
        }

        _webView.Dispose();
        base.OnFormClosing(e);
    }

    private async Task InitializeWebAsync()
    {
        try
        {
            var dataFolder = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "PBIClaw",
                "SetupWebView2");

            var env = await CoreWebView2Environment.CreateAsync(null, dataFolder);
            await _webView.EnsureCoreWebView2Async(env);
            _webView.CoreWebView2.Settings.AreDevToolsEnabled = false;
            _webView.CoreWebView2.Settings.AreDefaultContextMenusEnabled = false;
            _webView.CoreWebView2.Settings.IsStatusBarEnabled = false;
            _webView.CoreWebView2.WebMessageReceived += OnWebMessageReceived;
            _webView.CoreWebView2.NavigateToString(BuildHtml());
        }
        catch (Exception ex)
        {
            MessageBox.Show(
                $"安装器界面初始化失败：\n{ex.Message}",
                "PBI Claw 安装程序",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error);
            ExitCode = 1;
            Close();
        }
    }

    private void OnWebMessageReceived(object? sender, CoreWebView2WebMessageReceivedEventArgs e)
    {
        try
        {
            var actual = e.WebMessageAsJson;
            if (actual.StartsWith("\"", StringComparison.Ordinal) &&
                actual.EndsWith("\"", StringComparison.Ordinal))
            {
                actual = JsonSerializer.Deserialize<string>(actual) ?? actual;
            }

            using var doc = JsonDocument.Parse(actual);
            var root = doc.RootElement;
            var type = root.TryGetProperty("type", out var typeNode) ? typeNode.GetString() ?? string.Empty : string.Empty;
            var payload = root.TryGetProperty("payload", out var payloadNode) ? payloadNode : default;

            switch (type)
            {
                case "ready":
                    SendState();
                    if (_autoInstallOnReady && !_autoInstallStarted)
                    {
                        _autoInstallStarted = true;
                        _ = StartInstallAsync();
                    }
                    break;
                case "next":
                    _ = MoveNextAsync();
                    break;
                case "back":
                    MoveBack();
                    break;
                case "browseDir":
                    BrowseDirectory();
                    break;
                case "setInstallDir":
                    if (payload.ValueKind == JsonValueKind.Object && payload.TryGetProperty("value", out var val))
                    {
                        _installDir = val.GetString() ?? _installDir;
                        SendState();
                    }
                    break;
                case "finish":
                    ExitCode = _step == SetupStep.Complete ? 0 : 1;
                    Close();
                    break;
                case "cancel":
                    ExitCode = 1;
                    Close();
                    break;
                case "windowControl":
                    HandleWindowControl(payload);
                    break;
            }
        }
        catch
        {
            // Ignore malformed UI messages.
        }
    }

    private void HandleWindowControl(JsonElement payload)
    {
        if (payload.ValueKind != JsonValueKind.Object)
        {
            return;
        }

        var cmd = payload.TryGetProperty("cmd", out var cmdNode)
            ? (cmdNode.GetString() ?? string.Empty).Trim().ToLowerInvariant()
            : string.Empty;

        switch (cmd)
        {
            case "drag":
                ReleaseCapture();
                SendMessage(Handle, WmNclbuttondown, (IntPtr)HtCaption, IntPtr.Zero);
                break;
            case "minimize":
                WindowState = FormWindowState.Minimized;
                break;
            case "close":
                ExitCode = 1;
                Close();
                break;
        }
    }

    private async Task MoveNextAsync()
    {
        switch (_step)
        {
            case SetupStep.Welcome:
                _lastInstallFailed = false;
                _step = SetupStep.Directory;
                SendState();
                break;
            case SetupStep.Directory:
                if (!TryNormalizeInstallDir(out var error))
                {
                    SendError(error);
                    return;
                }
                _lastInstallFailed = false;
                _step = SetupStep.Confirm;
                SendState();
                break;
            case SetupStep.Confirm:
                await StartInstallAsync();
                break;
            case SetupStep.Complete:
                ExitCode = 0;
                Close();
                break;
        }
    }

    private void MoveBack()
    {
        switch (_step)
        {
            case SetupStep.Directory:
                _lastInstallFailed = false;
                _step = SetupStep.Welcome;
                SendState();
                break;
            case SetupStep.Confirm:
                _lastInstallFailed = false;
                _step = SetupStep.Directory;
                SendState();
                break;
        }
    }

    private void BrowseDirectory()
    {
        using var dialog = new FolderBrowserDialog
        {
            Description = "选择 PBI Claw 安装目录",
            UseDescriptionForTitle = true,
            ShowNewFolderButton = true,
            SelectedPath = _installDir
        };

        if (dialog.ShowDialog(this) == DialogResult.OK)
        {
            _installDir = dialog.SelectedPath;
            SendState();
        }
    }

    private async Task StartInstallAsync()
    {
        if (_installInProgress)
        {
            return;
        }

        if (!TryNormalizeInstallDir(out var error))
        {
            SendError(error);
            return;
        }

        if (!InstallerEngine.IsAdministrator())
        {
            var confirm = MessageBox.Show(
                "安装会写入系统级 External Tools 目录，需要管理员权限。\n是否现在提权继续？",
                "PBI Claw 安装程序",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Question);

            if (confirm != DialogResult.Yes)
            {
                return;
            }

            InstallerEngine.RelaunchAsAdmin(_installDir);
            ExitCode = 0;
            Close();
            return;
        }

        _installInProgress = true;
        _lastInstallFailed = false;
        _result = null;
        _step = SetupStep.Installing;
        SendState();

        try
        {
            _result = await Task.Run(() => InstallerEngine.InstallMachine(_installDir));
            _lastInstallFailed = false;
            _step = SetupStep.Complete;
            SendState();
        }
        catch (Exception ex)
        {
            _result = null;
            _lastInstallFailed = true;
            _step = SetupStep.Confirm;
            SendState();
            SendError($"安装失败：{ex.Message}");
        }
        finally
        {
            _installInProgress = false;
        }
    }

    private bool TryNormalizeInstallDir(out string error)
    {
        error = string.Empty;
        var raw = _installDir.Trim();
        if (string.IsNullOrWhiteSpace(raw))
        {
            error = "请选择安装目录。";
            return false;
        }

        try
        {
            _installDir = Path.GetFullPath(Environment.ExpandEnvironmentVariables(raw));
            return true;
        }
        catch
        {
            error = "安装目录格式无效，请重新选择。";
            return false;
        }
    }

    private void SendState()
    {
        Send("state", new
        {
            step = (int)_step,
            toolName = InstallerEngine.ToolName,
            installDir = _installDir,
            externalToolDirs = _externalToolDirs,
            canBack = _step is SetupStep.Directory or SetupStep.Confirm,
            nextText = _step switch
            {
                SetupStep.Confirm when _lastInstallFailed => "重试安装",
                SetupStep.Confirm => "安装",
                SetupStep.Complete => "完成",
                SetupStep.Installing => "安装中...",
                _ => "下一步"
            },
            canNext = _step != SetupStep.Installing,
            result = _result is null ? null : new
            {
                installDir = _result.InstallDir,
                externalToolDirs = _result.ExternalToolDirs
            }
        });
    }

    private void SendError(string message)
    {
        Send("error", new { message });
    }

    private void Send(string type, object payload)
    {
        if (_webView.CoreWebView2 is null)
        {
            return;
        }

        var json = JsonSerializer.Serialize(new { type, payload });
        _webView.CoreWebView2.PostWebMessageAsJson(json);
    }

    private static string BuildHtml()
    {
        var logoUri = GetInstallerLogoDataUri();

        var html = """
<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>PBI Claw Setup</title>
<style>
:root{
  --bg:#0f1117; --surface:#1a1d27; --surface2:#22263a; --border:#2e3350;
  --accent:#4f8ef7; --text:#e2e8f0; --text2:#8d97ab; --success:#22c55e; --danger:#ef4444;
  --font:'Microsoft YaHei UI','Segoe UI',system-ui,sans-serif;
}
*{box-sizing:border-box}
html,body{margin:0;height:100%;background:var(--bg);color:var(--text);font-family:var(--font);overflow:hidden}
#app{height:100%;display:flex;flex-direction:column}
#header{height:44px;flex-shrink:0;background:var(--surface);border-bottom:1px solid var(--border);padding:0 14px;display:flex;align-items:center}
#logo{width:24px;height:24px;border-radius:7px;display:flex;align-items:center;justify-content:center;margin-right:8px;overflow:hidden;border:1px solid var(--border);background:var(--surface2)}
#logo img{width:100%;height:100%;object-fit:cover;display:block}
#title{font-size:14px;font-weight:700;letter-spacing:.3px}
#topbar-spacer{flex:1}
.window-controls{display:inline-flex;gap:4px}
.win-btn{width:30px;height:26px;border-radius:6px;border:1px solid var(--border);background:var(--surface2);color:var(--text2);font-size:12px;line-height:1;cursor:pointer}
.win-btn:hover{border-color:var(--accent);color:var(--text);background:rgba(79,142,247,.1)}
.win-btn.close:hover{border-color:var(--danger);color:#fff;background:var(--danger)}
#body{flex:1;padding:20px 24px;overflow:auto}
.page{display:none}
.page.active{display:block}
.h1{font-size:18px;font-weight:700;margin:0 0 12px}
.txt{font-size:13px;line-height:1.75;color:var(--text2)}
.card{background:var(--surface);border:1px solid var(--border);border-radius:10px;padding:14px}
.label{font-size:12.5px;font-weight:600;margin-bottom:6px}
.row{display:flex;gap:10px;align-items:center}
.input{flex:1;background:var(--surface2);border:1px solid var(--border);border-radius:8px;padding:10px 12px;color:var(--text);font-size:13px;outline:none}
.input:focus{border-color:var(--accent)}
.btn{height:34px;padding:0 14px;border-radius:8px;border:1px solid var(--border);background:var(--surface2);color:var(--text);font-size:12.5px;cursor:pointer}
.btn:hover{border-color:var(--accent)}
.btn.primary{background:var(--accent);border-color:var(--accent);color:#fff}
.btn.primary:hover{background:#6ba3ff}
.btn:disabled{opacity:.55;cursor:not-allowed}
ul{margin:8px 0 0 18px;padding:0}
li{margin:4px 0;color:var(--text2)}
#dirs{max-height:146px;overflow:auto;background:var(--surface2);border:1px solid var(--border);border-radius:8px;padding:8px 10px}
#dirs div{font-size:12.5px;color:var(--text2);margin:4px 0;word-break:break-all}
#footer{height:62px;flex-shrink:0;border-top:1px solid var(--border);display:flex;justify-content:space-between;align-items:center;padding:0 20px}
#status{font-size:12px;color:var(--text2)}
#actions{display:flex;gap:8px}
#spinner{display:inline-block;width:14px;height:14px;border:2px solid var(--border);border-top-color:var(--accent);border-radius:50%;animation:spin .7s linear infinite;vertical-align:-2px;margin-right:6px}
@keyframes spin{to{transform:rotate(360deg)}}
#toast{position:fixed;top:14px;right:14px;max-width:360px;padding:10px 14px;border-radius:8px;border:1px solid rgba(239,68,68,.4);background:rgba(239,68,68,.12);color:#ffb5b5;display:none}
#result{background:var(--surface2);border:1px solid var(--border);border-radius:8px;padding:10px 12px;max-height:220px;overflow:auto}
#result div{font-size:12.5px;color:var(--text2);margin:5px 0;word-break:break-all}
</style>
</head>
<body>
<div id="toast"></div>
<div id="app">
  <div id="header" title="拖动窗口">
    <div id="logo"><img src="__LOGO_URI__" alt="PBI Claw"></div>
    <div id="title">PBI Claw 安装向导</div>
    <div id="topbar-spacer"></div>
    <div class="window-controls">
      <button class="win-btn" id="btn-win-min">_</button>
      <button class="win-btn close" id="btn-win-close">×</button>
    </div>
  </div>
  <div id="body">
    <section class="page active" data-step="1">
      <h2 class="h1">欢迎使用 PBI Claw</h2>
      <div class="card txt">
        安装向导将完成以下步骤：
        <ul>
          <li>把 <b>PBIClaw.exe</b> 安装到你选择的目录</li>
          <li>在两个系统级 External Tools 目录自动写入 <b>PBIClaw.pbitool.json</b></li>
        </ul>
        <div style="margin-top:10px">安装完成后，你可以直接使用这些核心能力：</div>
        <ul>
          <li>一键连接 PBIX / Tabular 模型，浏览表、字段、度量值、关系和报表页面元数据</li>
          <li>通过 AI 对话进行模型分析、DAX 建议与变更计划生成</li>
          <li>执行前预检、自动备份与回滚，降低模型写回风险</li>
        </ul>
        <div style="margin-top:10px">安装过程需要管理员权限。安装完成后请重启 Power BI Desktop 以加载新工具。</div>
      </div>
    </section>

    <section class="page" data-step="2">
      <h2 class="h1">选择安装目录</h2>
      <div class="card">
        <div class="label">PBIClaw.exe 安装位置</div>
        <div class="row">
          <input id="installDir" class="input" />
          <button class="btn" id="browseBtn">浏览...</button>
        </div>
      </div>
    </section>

    <section class="page" data-step="3">
      <h2 class="h1">确认安装信息</h2>
      <div class="card">
        <div class="label">程序安装目录</div>
        <div id="confirmDir" class="txt"></div>
        <div class="label" style="margin-top:12px">External Tools 写入目录</div>
        <div id="dirs"></div>
      </div>
    </section>

    <section class="page" data-step="4">
      <h2 class="h1">正在安装</h2>
      <div class="card txt"><span id="spinner"></span>正在复制文件并写入 External Tools 配置，请稍候...</div>
    </section>

    <section class="page" data-step="5">
      <h2 class="h1">安装完成</h2>
      <div class="card">
        <div class="txt">安装成功。请重启 Power BI Desktop，然后在 External Tools 中点击 <b>PBI Claw</b>（若未显示请再次重启 Power BI）。</div>
        <div id="result" style="margin-top:10px"></div>
      </div>
    </section>
  </div>
    <div id="footer">
    <div id="status">准备就绪</div>
    <div id="actions">
      <button class="btn" id="cancelBtn">取消</button>
      <button class="btn" id="backBtn" style="display:none">上一步</button>
      <button class="btn primary" id="nextBtn">下一步</button>
    </div>
  </div>
</div>

<script>
const State = { step: 1, installDir: '', externalToolDirs: [], canBack: false, canNext: true, nextText: '下一步', result: null };
let ActionPending = false;

function send(type, payload){ window.chrome.webview.postMessage(JSON.stringify({ type, payload: payload || {} })); }
function $(id){ return document.getElementById(id); }

function toast(msg){
  const el = $('toast');
  el.textContent = msg;
  el.style.display = 'block';
  setTimeout(() => { el.style.display = 'none'; }, 4200);
}

function render(){
  if (ActionPending) {
    ActionPending = false;
  }

  document.querySelectorAll('.page').forEach(p => p.classList.toggle('active', Number(p.dataset.step) === State.step));
  $('installDir').value = State.installDir || '';
  $('confirmDir').textContent = State.installDir || '';
  $('dirs').innerHTML = (State.externalToolDirs || []).map(d => '<div>' + escapeHtml(d) + '</div>').join('');
  $('backBtn').style.display = State.canBack ? '' : 'none';
  $('nextBtn').textContent = State.nextText || '下一步';
  $('nextBtn').disabled = !State.canNext || ActionPending;
  $('cancelBtn').style.display = State.step === 5 ? 'none' : '';
  $('backBtn').disabled = ActionPending;
  $('cancelBtn').disabled = ActionPending;
  $('browseBtn').disabled = ActionPending || State.step !== 2;
  $('installDir').disabled = ActionPending || State.step !== 2;
  $('status').textContent = State.step === 4 ? '安装中...' : (State.step === 5 ? '安装成功' : '准备就绪');

  if (State.step === 5 && State.result){
    const lines = [];
    lines.push('EXE 路径：');
    lines.push(State.result.installDir || '');
    lines.push('');
    lines.push('External Tools 配置：');
    (State.result.externalToolDirs || []).forEach(d => lines.push(d));
    $('result').innerHTML = lines.map(x => '<div>' + escapeHtml(x) + '</div>').join('');
  }
}

function escapeHtml(s){
  return String(s || '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

$('browseBtn').addEventListener('click', () => send('browseDir'));
$('nextBtn').addEventListener('click', () => {
  if (ActionPending) return;
  ActionPending = true;
  render();
  send('next');
});
$('backBtn').addEventListener('click', () => {
  if (ActionPending) return;
  ActionPending = true;
  render();
  send('back');
});
$('cancelBtn').addEventListener('click', () => send('cancel'));
$('btn-win-min').addEventListener('click', () => send('windowControl', { cmd: 'minimize' }));
$('btn-win-close').addEventListener('click', () => send('windowControl', { cmd: 'close' }));
$('header').addEventListener('mousedown', (e) => {
  if (e.button !== 0) return;
  if (e.target.closest('button,input,textarea,select,a,label')) return;
  send('windowControl', { cmd: 'drag' });
});
$('installDir').addEventListener('change', e => send('setInstallDir', { value: e.target.value || '' }));
$('installDir').addEventListener('keyup', e => {
  if (State.step === 2 && e.key === 'Enter') send('next');
});

window.chrome.webview.addEventListener('message', e => {
  try{
    const msg = typeof e.data === 'string' ? JSON.parse(e.data) : e.data;
    if (msg.type === 'state'){
      Object.assign(State, msg.payload || {});
      render();
    } else if (msg.type === 'error'){
      ActionPending = false;
      render();
      toast((msg.payload && msg.payload.message) || '发生未知错误');
    }
  }catch(_){}
});

render();
send('ready');
</script>
</body>
</html>
""";

        return html.Replace("__LOGO_URI__", logoUri);
    }

    private static string GetInstallerLogoDataUri()
    {
        const string defaultLogo = "data:image/svg+xml;utf8,<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 64 64'><defs><linearGradient id='g' x1='0' y1='0' x2='1' y2='1'><stop offset='0' stop-color='%234f8ef7'/><stop offset='1' stop-color='%232d5fb8'/></linearGradient></defs><rect width='64' height='64' rx='12' fill='url(%23g)'/><text x='32' y='39' font-family='Segoe UI' font-size='22' text-anchor='middle' fill='white'>PC</text></svg>";
        try
        {
            using var stream = FindEmbeddedResourceStream("Payload.PBIClaw.pbitool.json");
            if (stream is null)
            {
                return defaultLogo;
            }

            using var reader = new StreamReader(stream);
            var json = reader.ReadToEnd();
            using var doc = JsonDocument.Parse(json);
            if (!doc.RootElement.TryGetProperty("iconData", out var iconNode))
            {
                return defaultLogo;
            }

            var iconData = iconNode.GetString()?.Trim();
            if (string.IsNullOrWhiteSpace(iconData))
            {
                return defaultLogo;
            }

            if (iconData.StartsWith("data:", StringComparison.OrdinalIgnoreCase))
            {
                return iconData;
            }

            if (iconData.Contains(";base64,", StringComparison.OrdinalIgnoreCase))
            {
                return $"data:{iconData}";
            }

            return $"data:image/png;base64,{iconData}";
        }
        catch
        {
            return defaultLogo;
        }
    }

    private static Stream? FindEmbeddedResourceStream(string logicalName)
    {
        var assembly = Assembly.GetExecutingAssembly();
        var exact = assembly.GetManifestResourceStream(logicalName);
        if (exact is not null)
        {
            return exact;
        }

        var fallbackName = assembly
            .GetManifestResourceNames()
            .FirstOrDefault(name => name.EndsWith(logicalName, StringComparison.OrdinalIgnoreCase));

        return fallbackName is null ? null : assembly.GetManifestResourceStream(fallbackName);
    }

    [DllImport("user32.dll")]
    private static extern bool ReleaseCapture();

    [DllImport("user32.dll")]
    private static extern IntPtr SendMessage(IntPtr hWnd, int msg, IntPtr wParam, IntPtr lParam);
}
