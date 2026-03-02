using Microsoft.Web.WebView2.Core;
using Microsoft.Web.WebView2.WinForms;
using System.Reflection;
using System.Runtime.InteropServices;

namespace PbiMetadataTool;

/// <summary>
/// Main application window hosting a WebView2 control.
/// All UI logic lives in ui/index.html; C# handles business logic via AppBridge.
/// </summary>
internal sealed class MainFormWebView : Form
{
    private const int WmNchittest = 0x84;
    private const int HtClient = 1;
    private const int HtCaption = 2;
    private const int HtLeft = 10;
    private const int HtRight = 11;
    private const int HtTop = 12;
    private const int HtTopLeft = 13;
    private const int HtTopRight = 14;
    private const int HtBottom = 15;
    private const int HtBottomLeft = 16;
    private const int HtBottomRight = 17;
    private const int WmNclbuttondown = 0xA1;
    private const int ResizeBorder = 6;
    private const int DwmwaWindowCornerPreference = 33;
    private const int CornerRadius = 14;

    private readonly WebView2 _webView;
    private readonly AppBridge _bridge;
    private readonly CliOptions _startupOptions;
    private bool _useDwmCorners;

    public MainFormWebView(CliOptions startupOptions)
    {
        _startupOptions = startupOptions;
        _bridge = new AppBridge(this, _startupOptions);

        Text = $"PBI Claw v{CurrentVersion()}";
        var workingArea = Screen.PrimaryScreen?.WorkingArea ?? new Rectangle(0, 0, 1600, 900);
        Width = Math.Min(1720, (int)(workingArea.Width * 0.92));
        Height = Math.Min(1040, (int)(workingArea.Height * 0.92));
        MinimumSize = new Size(1100, 700);
        StartPosition = FormStartPosition.CenterScreen;
        FormBorderStyle = FormBorderStyle.None;
        AutoScaleMode = AutoScaleMode.Font;
        Font = new Font("Microsoft YaHei UI", 9F);
        BackColor = Color.FromArgb(15, 17, 23); // matches --bg
        ApplyWindowIcon();

        _webView = new WebView2 { Dock = DockStyle.Fill };
        Controls.Add(_webView);

        Resize += (_, _) =>
        {
            EnsureMaximizedBounds();
            UpdateWindowCornerPreference();
            ApplyRoundedWindowRegion();
        };
        Move += (_, _) => EnsureMaximizedBounds();

        _ = InitWebViewAsync();
    }

    private async Task InitWebViewAsync()
    {
        try
        {
            var dataFolder = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "PBIClaw", "WebView2");

            var env = await CoreWebView2Environment.CreateAsync(null, dataFolder);
            await _webView.EnsureCoreWebView2Async(env);

            _webView.CoreWebView2.Settings.AreDefaultContextMenusEnabled = true;
            _webView.CoreWebView2.Settings.AreDevToolsEnabled = true;
            _webView.CoreWebView2.Settings.IsStatusBarEnabled = false;

            _webView.CoreWebView2.WebMessageReceived += (_, e) =>
                _bridge.HandleMessage(e.WebMessageAsJson);

            _webView.CoreWebView2.NavigateToString(LoadEmbeddedHtml());
        }
        catch (Exception ex)
        {
            MessageBox.Show($"WebView2 初始化失败: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }

    public void PostMessage(string json)
    {
        if (_webView.CoreWebView2 is null) return;
        if (InvokeRequired)
            Invoke(() => _webView.CoreWebView2.PostWebMessageAsJson(json));
        else
            _webView.CoreWebView2.PostWebMessageAsJson(json);
    }

    public void ExecuteScript(string js)
    {
        if (_webView.CoreWebView2 is null) return;
        if (InvokeRequired)
            Invoke(() => _ = _webView.CoreWebView2.ExecuteScriptAsync(js));
        else
            _ = _webView.CoreWebView2.ExecuteScriptAsync(js);
    }

    public void MinimizeWindow()
    {
        if (InvokeRequired)
        {
            BeginInvoke(MinimizeWindow);
            return;
        }

        WindowState = FormWindowState.Minimized;
    }

    public void ToggleMaximizeRestore()
    {
        if (InvokeRequired)
        {
            BeginInvoke(ToggleMaximizeRestore);
            return;
        }

        WindowState = WindowState == FormWindowState.Maximized
            ? FormWindowState.Normal
            : FormWindowState.Maximized;
        UpdateWindowCornerPreference();
        ApplyRoundedWindowRegion();
    }

    public void CloseWindow()
    {
        if (InvokeRequired)
        {
            BeginInvoke(CloseWindow);
            return;
        }

        Close();
    }

    public void BeginDragMove()
    {
        if (InvokeRequired)
        {
            BeginInvoke(BeginDragMove);
            return;
        }

        ReleaseCapture();
        SendMessage(Handle, WmNclbuttondown, (IntPtr)HtCaption, IntPtr.Zero);
    }

    protected override void OnHandleCreated(EventArgs e)
    {
        base.OnHandleCreated(e);
        _useDwmCorners = TrySetDwmCornerPreference(DwmWindowCornerPreference.Round);
        EnsureMaximizedBounds();
        UpdateWindowCornerPreference();
        ApplyRoundedWindowRegion();
    }

    private static string LoadEmbeddedHtml()
    {
        var asm = Assembly.GetExecutingAssembly();
        var resourceName = asm.GetManifestResourceNames()
            .FirstOrDefault(n => n.EndsWith("index.html", StringComparison.OrdinalIgnoreCase));

        if (resourceName is null)
            return "<html><body style='background:#0f1117;color:#e2e8f0;font-family:sans-serif;padding:40px'><h2>UI resource not found</h2></body></html>";

        using var stream = asm.GetManifestResourceStream(resourceName)!;
        using var reader = new StreamReader(stream, System.Text.Encoding.UTF8);
        return reader.ReadToEnd();
    }

    protected override void OnFormClosing(FormClosingEventArgs e)
    {
        _webView.Dispose();
        base.OnFormClosing(e);
    }

    protected override void WndProc(ref Message m)
    {
        if (m.Msg == WmNchittest && FormBorderStyle == FormBorderStyle.None && WindowState == FormWindowState.Normal)
        {
            base.WndProc(ref m);
            if ((int)m.Result == HtClient)
            {
                var cursor = PointToClient(Cursor.Position);
                var left = cursor.X <= ResizeBorder;
                var right = cursor.X >= ClientSize.Width - ResizeBorder;
                var top = cursor.Y <= ResizeBorder;
                var bottom = cursor.Y >= ClientSize.Height - ResizeBorder;

                if (left && top) m.Result = (IntPtr)HtTopLeft;
                else if (right && top) m.Result = (IntPtr)HtTopRight;
                else if (left && bottom) m.Result = (IntPtr)HtBottomLeft;
                else if (right && bottom) m.Result = (IntPtr)HtBottomRight;
                else if (left) m.Result = (IntPtr)HtLeft;
                else if (right) m.Result = (IntPtr)HtRight;
                else if (top) m.Result = (IntPtr)HtTop;
                else if (bottom) m.Result = (IntPtr)HtBottom;
            }

            return;
        }

        base.WndProc(ref m);
    }

    private static string CurrentVersion()
    {
        var v = Assembly.GetExecutingAssembly().GetName().Version;
        return v is null ? "1.0" : $"{v.Major}.{v.Minor}.{v.Build}";
    }

    private void ApplyWindowIcon()
    {
        try
        {
            using var extracted = Icon.ExtractAssociatedIcon(Application.ExecutablePath);
            if (extracted is not null)
            {
                Icon = (Icon)extracted.Clone();
            }
        }
        catch
        {
            // Keep default icon when extraction fails.
        }
    }

    [DllImport("user32.dll")]
    private static extern bool ReleaseCapture();

    [DllImport("user32.dll")]
    private static extern IntPtr SendMessage(IntPtr hWnd, int msg, IntPtr wParam, IntPtr lParam);

    private void EnsureMaximizedBounds()
    {
        if (!IsHandleCreated)
        {
            return;
        }

        var screen = Screen.FromHandle(Handle);
        MaximizedBounds = screen.WorkingArea;
    }

    private void ApplyRoundedWindowRegion()
    {
        if (!IsHandleCreated)
        {
            return;
        }

        if (_useDwmCorners)
        {
            if (Region is not null)
            {
                Region.Dispose();
                Region = null;
            }
            return;
        }

        if (WindowState == FormWindowState.Maximized)
        {
            if (Region is not null)
            {
                Region.Dispose();
                Region = null;
            }
            return;
        }

        var oldRegion = Region;
        var hrgn = CreateRoundRectRgn(0, 0, Width + 1, Height + 1, CornerRadius, CornerRadius);
        try
        {
            Region = Region.FromHrgn(hrgn);
        }
        finally
        {
            DeleteObject(hrgn);
            oldRegion?.Dispose();
        }
    }

    private void UpdateWindowCornerPreference()
    {
        if (!_useDwmCorners || !IsHandleCreated)
        {
            return;
        }

        var pref = WindowState == FormWindowState.Maximized
            ? DwmWindowCornerPreference.DoNotRound
            : DwmWindowCornerPreference.Round;

        TrySetDwmCornerPreference(pref);
    }

    private bool TrySetDwmCornerPreference(DwmWindowCornerPreference preference)
    {
        try
        {
            var pref = (int)preference;
            return DwmSetWindowAttribute(Handle, DwmwaWindowCornerPreference, ref pref, sizeof(int)) == 0;
        }
        catch
        {
            return false;
        }
    }

    [DllImport("gdi32.dll")]
    private static extern IntPtr CreateRoundRectRgn(int nLeftRect, int nTopRect, int nRightRect, int nBottomRect, int nWidthEllipse, int nHeightEllipse);

    [DllImport("gdi32.dll")]
    private static extern bool DeleteObject(IntPtr hObject);

    [DllImport("dwmapi.dll")]
    private static extern int DwmSetWindowAttribute(IntPtr hwnd, int attribute, ref int pvAttribute, int cbAttribute);

    private enum DwmWindowCornerPreference
    {
        Default = 0,
        DoNotRound = 1,
        Round = 2,
        RoundSmall = 3
    }
}
