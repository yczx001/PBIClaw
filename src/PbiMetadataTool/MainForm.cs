using System.Diagnostics;
using System.Drawing.Drawing2D;
using System.Net;
using System.Text.Json;
using System.Text.RegularExpressions;

namespace PbiMetadataTool;

/// <summary>
/// Main application form.
/// Redesigned with a custom-drawn UI to visually mimic DAX Studio.
/// </summary>
internal sealed class MainForm : Form
{
    #region Style & Colors
    // DAX Studio inspired color palette
    private static readonly Color WindowBg = Color.FromArgb(240, 240, 240);
    private static readonly Color RibbonBg = Color.FromArgb(247, 247, 247);
    private static readonly Color RibbonBottomBorder = Color.FromArgb(204, 204, 204);
    private static readonly Color CardBg = Color.White;
    private static readonly Color CardBorder = Color.FromArgb(204, 204, 204);
    private static readonly Color TextPrimary = Color.FromArgb(30, 30, 30);
    private static readonly Color TextSecondary = Color.FromArgb(102, 102, 102);
    private static readonly Color ActiveTabBg = Color.White;
    private static readonly Color InactiveTabBg = Color.FromArgb(230, 230, 230);
    private static readonly Color ButtonHoverBg = Color.FromArgb(229, 243, 255);
    private static readonly Color ButtonHoverBorder = Color.FromArgb(0, 122, 204);
    private static readonly Color StatusBg = Color.FromArgb(0, 122, 204);
    private static readonly Color StatusFg = Color.White;
    #endregion
    private const double LeftPaneWidthRatio = 0.30;
    private const int DesiredRightPaneMinWidth = 260;

    private readonly CliOptions _startupOptions;
    private readonly PowerBiInstanceDetector _detector = new();
    private readonly TabularMetadataReader _reader = new();
    private readonly TabularModelWriter _writer = new();
    private readonly AiChatClient _chatClient = new();
    private readonly AbiSettingsStore _settingsStore = new();
    private readonly List<AiChatMessage> _conversation = [];

    // --- UI Controls ---
    private TableLayoutPanel _ribbon = null!;
    private Panel _mainContent = null!;
    private StatusStrip _statusBar = null!;
    private ToolStripStatusLabel _statusLabel = null!;
    private readonly Dictionary<string, Panel> _pages = new(StringComparer.OrdinalIgnoreCase);
    private readonly Dictionary<string, Button> _mainTabs = new(StringComparer.OrdinalIgnoreCase);
    private readonly TreeView _tree;
    // 连接状态字段（替代原来的控件字段，避免 dispose 问题）
    private bool _useTabularMode = false;
    private string _lastPort = string.Empty;
    private string _lastTabularServer = string.Empty;
    private readonly TextBox _summary;
    private readonly RichTextBox _chatLog;
    private readonly TextBox _chatInput;
    private readonly FlowLayoutPanel _quickPromptsPanel;
    private SplitContainer _dashboardSplitter = null!;
    private readonly CheckedListBox _planActions;
    private readonly TextBox _planPreview;
    private readonly ListView _backups;

    // --- State ---
    private IReadOnlyList<PowerBiInstanceInfo> _instanceData = Array.Empty<PowerBiInstanceInfo>();
    private ModelMetadata? _model;
    private AbiAssistantSettings _settings = new();
    private AbiActionPlan? _pendingPlan;
    private CancellationTokenSource? _chatCts;
    private int? _currentPort;

    public MainForm(CliOptions startupOptions)
    {
        _startupOptions = startupOptions;
        Text = $"PBI Claw v{CurrentVersion()}";
        Width = 1560;
        Height = 980;
        MinimumSize = new Size(1320, 820);
        StartPosition = FormStartPosition.CenterScreen;
        AutoScaleMode = AutoScaleMode.Font;
        Font = new Font("Microsoft YaHei UI", 9F, FontStyle.Regular, GraphicsUnit.Point);
        BackColor = WindowBg;
        ApplyWindowIcon();
        
        _tree = new TreeView { Dock = DockStyle.Fill, BorderStyle = BorderStyle.None, FullRowSelect = true, HideSelection = false, ItemHeight = 28, BackColor = CardBg, DrawMode = TreeViewDrawMode.OwnerDrawText, Font = new Font("Microsoft YaHei UI", 10.5F, FontStyle.Regular) };
        _tree.DrawNode += TreeDrawNode;
        _summary = new TextBox { Dock = DockStyle.Fill, Multiline = true, ReadOnly = true, ScrollBars = ScrollBars.Vertical, Font = new Font("Consolas", 10f), BorderStyle = BorderStyle.None, BackColor = CardBg };
        _chatLog = new RichTextBox { Dock = DockStyle.Fill, BorderStyle = BorderStyle.None, ReadOnly = true, Font = new Font("Microsoft YaHei UI", 10f), BackColor = CardBg };
        _chatInput = new TextBox { Dock = DockStyle.Fill, Multiline = true, PlaceholderText = "输入你的问题，Ctrl+Enter 发送" };
        _quickPromptsPanel = new FlowLayoutPanel { Dock = DockStyle.Fill, FlowDirection = FlowDirection.TopDown, AutoScroll = true, BackColor = CardBg, Padding = new Padding(5) };
        _planActions = new CheckedListBox { Dock = DockStyle.Fill, BorderStyle = BorderStyle.None, CheckOnClick = true, BackColor = CardBg };
        _planPreview = new TextBox { Dock = DockStyle.Fill, Multiline = true, ReadOnly = true, ScrollBars = ScrollBars.Vertical, Font = new Font("Consolas", 10f), BorderStyle = BorderStyle.None, BackColor = CardBg };
        _backups = new ListView { Dock = DockStyle.Fill, View = View.Details, FullRowSelect = true, GridLines = true, BorderStyle = BorderStyle.None, BackColor = CardBg };

        BuildLayout();
        BindEvents();
    }

    private void BuildLayout()
    {
        _ribbon = BuildRibbon();
        _mainContent = BuildMainContent();
        _statusBar = BuildStatusBar();
        
        var mainLayout = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 3 };
        mainLayout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        mainLayout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
        mainLayout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        mainLayout.Controls.Add(_ribbon, 0, 0);
        mainLayout.Controls.Add(_mainContent, 0, 1);
        mainLayout.Controls.Add(_statusBar, 0, 2);

        Controls.Add(mainLayout);
    }

    private MenuStrip BuildMenu()
    {
        var menu = new MenuStrip { Dock = DockStyle.Fill, BackColor = RibbonBg, Renderer = new ToolStripProfessionalRenderer(new DaxStudioColorTable()) };
        var file = new ToolStripMenuItem("文件");
        file.DropDownItems.Add("扫描报告", null, (_, _) => DetectInstances());
        file.DropDownItems.Add("连接模型", null, (_, _) => ConnectAndLoadMetadata());
        file.DropDownItems.Add(new ToolStripSeparator());
        file.DropDownItems.Add("设置", null, (_, _) => OpenSettingsDialog());
        file.DropDownItems.Add(new ToolStripSeparator());
        file.DropDownItems.Add("退出", null, (_, _) => Close());
        var help = new ToolStripMenuItem("帮助");
        help.DropDownItems.Add("关于", null, (_, _) => MessageBox.Show($"PBI Claw\n版本: {CurrentVersion()}\n网址：https://pbihub.cn/", "关于 PBI Claw", MessageBoxButtons.OK, MessageBoxIcon.Information));
        menu.Items.Add(file);
        menu.Items.Add(help);
        return menu;
    }

    private TableLayoutPanel BuildRibbon()
    {
        var ribbonLayout = new TableLayoutPanel { Dock = DockStyle.Top, Height = 110, BackColor = RibbonBg, ColumnCount = 1, RowCount = 2, Padding = new Padding(0) };
        ribbonLayout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        ribbonLayout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
        ribbonLayout.Controls.Add(BuildMenu(), 0, 0);
        var ribbonTabRow = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 4, RowCount = 1, Padding = new Padding(0) };
        ribbonTabRow.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
        ribbonTabRow.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
        ribbonTabRow.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
        ribbonTabRow.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 8));
        var ribbonTabs = new FlowLayoutPanel { Dock = DockStyle.Fill, WrapContents = false, Padding = new Padding(8, 4, 8, 0) };
        _mainTabs["dashboard"] = RibbonTab("仪表盘", "dashboard");
        _mainTabs["writeback"] = RibbonTab("变更执行", "writeback");
        _mainTabs["backups"] = RibbonTab("备份中心", "backups");
        foreach(var tab in _mainTabs.Values) ribbonTabs.Controls.Add(tab);
        var btnConnect = CreateActionButton("连接模型", (_, _) => ShowConnectDialog());
        var btnScan = CreateActionButton("扫描实例", (_, _) => DetectInstances());
        btnConnect.Dock = DockStyle.None;
        btnScan.Dock = DockStyle.None;
        btnConnect.Anchor = AnchorStyles.Right | AnchorStyles.Top;
        btnScan.Anchor = AnchorStyles.Right | AnchorStyles.Top;
        btnConnect.Margin = new Padding(0, 4, 4, 4);
        btnScan.Margin = new Padding(0, 4, 4, 4);
        ribbonTabRow.Controls.Add(ribbonTabs, 0, 0);
        ribbonTabRow.Controls.Add(btnConnect, 1, 0);
        ribbonTabRow.Controls.Add(btnScan, 2, 0);
        ribbonTabRow.Controls.Add(new Panel { Dock = DockStyle.Fill }, 3, 0);
        ribbonLayout.Controls.Add(ribbonTabRow, 0, 1);
        ribbonLayout.Paint += (s, e) => {
            using var p = new Pen(RibbonBottomBorder);
            e.Graphics.DrawLine(p, 0, ribbonLayout.Height - 1, ribbonLayout.Width, ribbonLayout.Height - 1);
        };
        return ribbonLayout;
    }

    private Button RibbonTab(string text, string key)
    {
        var tab = new Button { Text = text, Tag = key, AutoSize = true, Height = 34, FlatStyle = FlatStyle.Flat, Font = new Font("Microsoft YaHei UI", 10F, FontStyle.Bold), Padding = new Padding(12, 4, 12, 4), Margin = new Padding(0, 0, 6, 0), UseVisualStyleBackColor = false, ForeColor = TextPrimary };
        tab.FlatAppearance.BorderSize = 0;
        tab.Click += (s, e) => ShowPage(key);
        return tab;
    }
    
    private Panel BuildMainContent()
    {
        var host = new Panel { Dock = DockStyle.Fill, BackColor = WindowBg, Padding = new Padding(8) };
        _pages["dashboard"] = BuildDashboardPage();
        _pages["writeback"] = BuildWritebackPage();
        _pages["backups"] = BuildBackupsPage();
        foreach(var page in _pages.Values) host.Controls.Add(page);
        return host;
    }

    private StatusStrip BuildStatusBar()
    {
        var status = new StatusStrip { BackColor = StatusBg, ForeColor = StatusFg };
        status.Padding = new Padding(status.Padding.Left, 0, status.Padding.Right, 0);
        _statusLabel = new ToolStripStatusLabel { Spring = true, TextAlign = ContentAlignment.MiddleLeft };
        var versionLabel = new ToolStripStatusLabel { Text = $"v{CurrentVersion()}", BorderSides = ToolStripStatusLabelBorderSides.Left };
        status.Items.Add(_statusLabel);
        status.Items.Add(versionLabel);
        return status;
    }

    #region Page Builders

    private Panel BuildDashboardPage()
    {
        var page = new Panel { Dock = DockStyle.Fill, Visible = false };
        _dashboardSplitter = new SplitContainer
        {
            Dock = DockStyle.Fill,
            FixedPanel = FixedPanel.Panel1,
            BackColor = WindowBg,
            SplitterWidth = 8
        };
        AttachInitialSplitterLayout(_dashboardSplitter);

        // 左侧：模型浏览器
        _dashboardSplitter.Panel1.Controls.Add(Card("模型浏览器", _tree));

        // 右侧：对话历史（上）+ 输入框（下，固定 300px）
        var rightPanel = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 2 };
        rightPanel.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
        rightPanel.RowStyles.Add(new RowStyle(SizeType.Absolute, 300));

        rightPanel.Controls.Add(Card("AI 对话", _chatLog), 0, 0);

        var inputCardContent = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 2 };
        inputCardContent.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
        inputCardContent.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        inputCardContent.Controls.Add(_chatInput, 0, 0);
        var chatActions = new FlowLayoutPanel { Dock = DockStyle.Right, AutoSize = true, Padding = new Padding(0, 4, 0, 0) };
        chatActions.Controls.Add(CreateActionButton("清空会话", (_, _) => { _conversation.Clear(); _chatLog.Clear(); }));
        chatActions.Controls.Add(CreateActionButton("发送", async (_, _) => await SendPromptAsync()));
        inputCardContent.Controls.Add(chatActions, 0, 1);
        rightPanel.Controls.Add(Card("输入", inputCardContent), 0, 1);

        _dashboardSplitter.Panel2.Controls.Add(rightPanel);
        page.Controls.Add(_dashboardSplitter);
        return page;
    }

    private TableLayoutPanel BuildConnectPanel() => new(); // 已废弃，连接逻辑移至 ShowConnectDialog

    private Panel BuildWritebackPage()
    {
        var page = new Panel { Dock = DockStyle.Fill, Visible = false };
        var split = new SplitContainer { Dock = DockStyle.Fill, BackColor = WindowBg, SplitterWidth = 8 };
        var leftActions = new FlowLayoutPanel { Dock = DockStyle.Top, WrapContents = false, AutoSize = true, Padding = new Padding(4, 4, 0, 8) };
        leftActions.Controls.Add(CreateActionButton("全选", (_,_) => SetPlanChecks(true)));
        leftActions.Controls.Add(CreateActionButton("清空选择", (_,_) => SetPlanChecks(false)));
        leftActions.Controls.Add(CreateActionButton("丢弃计划", (_,_) => ClearPendingPlan()));
        leftActions.Controls.Add(CreateActionButton("预检并执行", (_,_) => ExecutePendingPlan(), isPrimary: true));
        var leftCard = Card("待执行动作", _planActions);
        leftCard.Controls.Add(leftActions);
        leftActions.BringToFront();
        split.Panel1.Controls.Add(leftCard);
        split.Panel2.Controls.Add(Card("计划预览", _planPreview));
        page.Controls.Add(split);
        return page;
    }
    
    private Panel BuildBackupsPage()
    {
        var page = new Panel { Dock = DockStyle.Fill, Visible = false };
        _backups.Columns.Clear();
        _backups.Columns.Add("时间", 180);
        _backups.Columns.Add("类型", 90);
        _backups.Columns.Add("文件名", 400);
        _backups.Columns.Add("大小", 90);
        _backups.Columns.Add("路径", 500);
        var backupActions = new FlowLayoutPanel { Dock = DockStyle.Top, WrapContents = false, AutoSize = true, Padding = new Padding(4, 4, 0, 8) };
        backupActions.Controls.Add(CreateActionButton("刷新", (_,_) => RefreshBackupList()));
        backupActions.Controls.Add(CreateActionButton("打开目录", (_,_) => OpenBackupRootFolder()));
        backupActions.Controls.Add(CreateActionButton("执行回滚", (_,_) => RunSelectedBackupRollback(), isPrimary: true));
        var card = Card("备份与回滚历史", _backups);
        card.Controls.Add(backupActions);
        backupActions.BringToFront();
        page.Controls.Add(card);
        return page;
    }
    
    #endregion

    #region UI Helpers
    
    private static Panel Card(string title, Control content) => Card(title, content, autoSizeContent: false);

    private static Panel Card(string title, Control content, bool autoSizeContent)
    {
        var card = new Panel { Dock = DockStyle.Fill, Padding = new Padding(1), BackColor = CardBorder };
        var inner = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 2, BackColor = CardBg };
        inner.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        if (autoSizeContent)
        {
            inner.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            inner.AutoSize = true;
            card.AutoSize = true;
        }
        else
        {
            inner.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
        }
        var header = new Label { Text = title, Dock = DockStyle.Fill, BackColor = InactiveTabBg, ForeColor = TextPrimary, Font = new Font("Microsoft YaHei UI", 10F, FontStyle.Bold), Padding = new Padding(10, 0, 10, 0), Height = 36, TextAlign = ContentAlignment.MiddleLeft };
        content.Padding = new Padding(8);
        inner.Controls.Add(header, 0, 0);
        inner.Controls.Add(content, 0, 1);
        card.Controls.Add(inner);
        return card;
    }

    private static Button CreateActionButton(string text, EventHandler onClick, bool isPrimary = false)
    {
        var b = new Button { Text = text, AutoSize = true, FlatStyle = FlatStyle.Flat, BackColor = isPrimary ? StatusBg : CardBg, ForeColor = isPrimary ? StatusFg : TextPrimary, Font = new Font("Microsoft YaHei UI", 9F, FontStyle.Regular), Margin = new Padding(4, 4, 8, 4), Padding = new Padding(10, 4, 10, 4) };
        b.FlatAppearance.BorderColor = CardBorder;
        b.FlatAppearance.BorderSize = 1;
        b.FlatAppearance.MouseOverBackColor = ButtonHoverBg;
        b.FlatAppearance.MouseDownBackColor = ControlPaint.Dark(ButtonHoverBg);
        b.Click += onClick;
        return b;
    }

    private void ShowPage(string key)
    {
        foreach (var kvp in _pages) kvp.Value.Visible = kvp.Key.Equals(key, StringComparison.OrdinalIgnoreCase);
        foreach (var kvp in _mainTabs)
        {
            kvp.Value.BackColor = kvp.Key.Equals(key, StringComparison.OrdinalIgnoreCase) ? ActiveTabBg : InactiveTabBg;
            kvp.Value.FlatAppearance.BorderSize = kvp.Key.Equals(key, StringComparison.OrdinalIgnoreCase) ? 1 : 0;
            kvp.Value.FlatAppearance.BorderColor = CardBorder;
        }
        ApplyPrimarySplitters();
    }
    
    #endregion
    
    private void BindEvents()
    {
        _chatInput.KeyDown += (s, e) => { if (e.Control && e.KeyCode == Keys.Enter) { e.SuppressKeyPress = true; _ = SendPromptAsync(); } };
        _backups.DoubleClick += (_, _) => RunSelectedBackupRollback();
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
    
    protected override void OnShown(EventArgs e)
    {
        base.OnShown(e);
        ApplyPrimarySplitters();
        Init();

        ApplyPrimarySplitters();
    }

    protected override void OnResize(EventArgs e)
    {
        base.OnResize(e);
        ApplyPrimarySplitters();
    }

    protected override void OnFormClosing(FormClosingEventArgs e)
    {
        _chatCts?.Cancel();
        _chatCts?.Dispose();
        base.OnFormClosing(e);
    }
    
    private void Init()
    {
        _settings = _settingsStore.Load();
        if (_startupOptions.Port.HasValue) _lastPort = _startupOptions.Port.Value.ToString();
        RefreshQuickPrompts();
        RefreshBackupList();
        DetectInstances();
        SetStatus($"欢迎使用 PBI Claw (v{CurrentVersion()})。");
        ShowPage("dashboard");

        if (_startupOptions.ExternalToolMode)
        {
            if (!string.IsNullOrWhiteSpace(NormalizeExternalToolValue(_startupOptions.Server)))
                _lastTabularServer = NormalizeExternalToolValue(_startupOptions.Server)!;
            TryAutoConnectFromExternalTool();
        }
    }
    
    private void RefreshQuickPrompts()
    {
        _quickPromptsPanel.Controls.Clear();
        foreach (var prompt in _settings.QuickPrompts.Take(12))
        {
            var b = new Button { Text = prompt, Font = new Font(this.Font.FontFamily, 9.5f), AutoSize = true, Height = 40, Dock = DockStyle.Top, FlatStyle = FlatStyle.System, TextAlign = ContentAlignment.MiddleLeft, Padding = new Padding(8, 0, 0, 0) };
            b.Click += (_, _) => { _chatInput.Text = prompt; _ = SendPromptAsync(); };
            _quickPromptsPanel.Controls.Add(b);
        }
    }
    
    private void HandlePermissionChanged(bool value)
    {
        if (_settings.AllowModelChanges == value) return;
        _settings.AllowModelChanges = value;
        _settingsStore.Save(_settings);
        if (!_settings.AllowModelChanges) ClearPendingPlan();
        SetStatus(_settings.AllowModelChanges ? "已开启写回模式。" : "已切换为只读模式。");
    }

    private void OpenSettingsDialog()
    {
        if (!AbiSettingsDialog.TryEdit(this, _settings, out var updated)) return;
        _settings = updated;
        _settingsStore.Save(_settings);
        RefreshQuickPrompts();
        SetStatus("设置已更新。");
    }

    private void ShowConnectDialog()
    {
        using var dlg = new Form
        {
            Text = "连接模型",
            Width = 520,
            Height = 360,
            MinimumSize = new Size(420, 320),
            StartPosition = FormStartPosition.CenterParent,
            FormBorderStyle = FormBorderStyle.SizableToolWindow,
            MaximizeBox = false,
            MinimizeBox = false,
            BackColor = CardBg,
            Font = this.Font,
        };

        // 每次创建全新的局部控件，避免 dispose 后复用崩溃
        var radioPbi      = new RadioButton { Text = "Power BI Desktop", AutoSize = true, Checked = !_useTabularMode, Font = new Font("Microsoft YaHei UI", 9F, FontStyle.Bold) };
        var radioTabular  = new RadioButton { Text = "Tabular Server",   AutoSize = true, Checked = _useTabularMode,  Font = new Font("Microsoft YaHei UI", 9F, FontStyle.Bold) };
        var instances     = new ComboBox   { Dock = DockStyle.Fill, DropDownStyle = ComboBoxStyle.DropDownList };
        var port          = new TextBox    { Dock = DockStyle.Fill, PlaceholderText = "端口", Text = _lastPort };
        var tabularServer = new TextBox    { Dock = DockStyle.Fill, PlaceholderText = "例：myserver\\tabular 或 myserver:2383", Text = _lastTabularServer };
        var allowWrite    = new CheckBox   { AutoSize = true, Text = "允许写回（高风险）", Checked = _settings.AllowModelChanges };

        // 填充已知实例
        foreach (var item in _instanceData) instances.Items.Add(new InstanceItem(item));
        if (instances.Items.Count > 0)
        {
            var sel = _instanceData.FirstOrDefault(x => x.Port.ToString() == _lastPort);
            instances.SelectedIndex = sel is not null
                ? _instanceData.ToList().IndexOf(sel)
                : 0;
        }
        instances.SelectedIndexChanged += (_, _) =>
        {
            if (instances.SelectedItem is InstanceItem item)
                port.Text = item.Instance.Port.ToString();
        };

        // PBI 面板
        var pbiPanel = new Panel { Dock = DockStyle.Top, AutoSize = true };
        var pbiGrid  = new TableLayoutPanel { Dock = DockStyle.Top, AutoSize = true, ColumnCount = 2, RowCount = 2, Padding = new Padding(0, 4, 0, 4) };
        pbiGrid.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 70));
        pbiGrid.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 30));
        pbiGrid.Controls.Add(new Label { Text = "报告实例", AutoSize = true, Margin = new Padding(3,3,3,4) }, 0, 0);
        pbiGrid.Controls.Add(new Label { Text = "端口",     AutoSize = true, Margin = new Padding(3,3,3,4) }, 1, 0);
        pbiGrid.Controls.Add(instances, 0, 1);
        pbiGrid.Controls.Add(port,      1, 1);
        var pbiScan = CreateActionButton("扫描实例", (_, _) =>
        {
            var found = _detector.DiscoverInstances();
            instances.Items.Clear();
            foreach (var item in found) instances.Items.Add(new InstanceItem(item));
            if (instances.Items.Count > 0) instances.SelectedIndex = 0;
        });
        // 扫描实例 + 允许写回 同一行
        allowWrite.Margin = new Padding(16, 0, 0, 0);
        var scanAllowRow = new FlowLayoutPanel { Dock = DockStyle.Top, AutoSize = true, WrapContents = false, Padding = new Padding(0, 6, 0, 2) };
        scanAllowRow.Controls.Add(pbiScan);
        scanAllowRow.Controls.Add(allowWrite);
        pbiPanel.Controls.Add(scanAllowRow);
        pbiPanel.Controls.Add(pbiGrid);

        // Tabular 面板
        var tabPanel = new Panel { Dock = DockStyle.Top, AutoSize = true, Visible = _useTabularMode };
        var tabGrid  = new TableLayoutPanel { Dock = DockStyle.Top, AutoSize = true, ColumnCount = 1, RowCount = 2, Padding = new Padding(0, 4, 0, 4) };
        tabGrid.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
        tabGrid.Controls.Add(new Label { Text = "服务器地址", AutoSize = true, Margin = new Padding(3,3,3,4) }, 0, 0);
        tabGrid.Controls.Add(tabularServer, 0, 1);
        // Tabular 模式下的允许写回（独立行，始终可见）
        var tabAllowRow = new FlowLayoutPanel { Dock = DockStyle.Top, AutoSize = true, Padding = new Padding(0, 6, 0, 2) };
        var allowWrite2 = new CheckBox { AutoSize = true, Text = "允许写回（高风险）", Checked = _settings.AllowModelChanges };
        tabAllowRow.Controls.Add(allowWrite2);
        tabPanel.Controls.Add(tabAllowRow);
        tabPanel.Controls.Add(tabGrid);

        // 两个 CheckBox 保持同步
        allowWrite.CheckedChanged  += (_, _) => allowWrite2.Checked = allowWrite.Checked;
        allowWrite2.CheckedChanged += (_, _) => allowWrite.Checked  = allowWrite2.Checked;

        // 模式切换
        pbiPanel.Visible = !_useTabularMode;
        radioPbi.CheckedChanged += (_, _) => { pbiPanel.Visible = radioPbi.Checked; tabPanel.Visible = !radioPbi.Checked; };

        // 模式选择行
        var modeRow = new FlowLayoutPanel { Dock = DockStyle.Top, AutoSize = true, WrapContents = false, Padding = new Padding(0, 4, 0, 8) };
        modeRow.Controls.Add(radioPbi);
        radioTabular.Margin = new Padding(20, 0, 0, 0);
        modeRow.Controls.Add(radioTabular);

        // 外层容器（Dock=Top 控件倒序 Add）
        var content = new Panel { Dock = DockStyle.Top, AutoSize = true, Padding = new Padding(4) };
        content.Controls.Add(tabPanel);
        content.Controls.Add(pbiPanel);
        content.Controls.Add(modeRow);

        // 按钮行
        var btnRow     = new FlowLayoutPanel { Dock = DockStyle.Fill, AutoSize = true, FlowDirection = FlowDirection.RightToLeft, Padding = new Padding(0, 8, 0, 0) };
        var btnConnect = new Button { Text = "连接", DialogResult = DialogResult.OK,     Width = 88, Height = 36, FlatStyle = FlatStyle.Flat, BackColor = StatusBg,       ForeColor = StatusFg };
        var btnCancel  = new Button { Text = "取消", DialogResult = DialogResult.Cancel, Width = 88, Height = 36, FlatStyle = FlatStyle.Flat, BackColor = InactiveTabBg, ForeColor = TextPrimary };
        btnConnect.FlatAppearance.BorderSize = 0;
        btnCancel.FlatAppearance.BorderSize  = 1;
        btnCancel.FlatAppearance.BorderColor = CardBorder;
        btnRow.Controls.Add(btnConnect);
        btnRow.Controls.Add(btnCancel);

        var outer = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 2, Padding = new Padding(16) };
        outer.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        outer.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        outer.Controls.Add(content, 0, 0);
        outer.Controls.Add(btnRow,  0, 1);

        dlg.Controls.Add(outer);
        dlg.AcceptButton = btnConnect;
        dlg.CancelButton = btnCancel;

        if (dlg.ShowDialog(this) == DialogResult.OK)
        {
            // 把对话框里的值同步回字段状态
            _useTabularMode      = radioTabular.Checked;
            _lastPort            = port.Text.Trim();
            _lastTabularServer   = tabularServer.Text.Trim();
            var newAllow         = allowWrite.Checked;

            // 同步 instances 选中到 _instanceData 选中
            if (instances.SelectedItem is InstanceItem sel)
                PopulateInstances(_instanceData, sel.Instance.Port);

            // 更新 allowWrite 设置
            if (newAllow != _settings.AllowModelChanges)
            {
                _settings.AllowModelChanges = newAllow;
                _settingsStore.Save(_settings);
                HandlePermissionChanged(newAllow);
            }

            ConnectAndLoadMetadata();
        }
    }

    private void DetectInstances()
    {
        try
        {
            PopulateInstances(_detector.DiscoverInstances());
            SetStatus(_instanceData.Any() ? $"已找到 {_instanceData.Count} 个 Power BI 实例。" : "未检测到正在运行的 Power BI 实例。请先打开一个 PBIX 文件。");
        }
        catch (Exception ex) { SetStatus($"扫描实例时出错: {ex.Message}"); }
    }

    private void TryAutoConnectFromExternalTool()
    {
        if (_model is not null) return;
        var hasContext = _startupOptions.Port.HasValue || !string.IsNullOrWhiteSpace(NormalizeExternalToolValue(_startupOptions.Server));
        if (!hasContext) return;
        ConnectAndLoadMetadata(showErrorDialog: false);
    }

    private void ApplySelectedInstance()
    {
        // 已废弃：连接状态现在通过 _lastPort / _useTabularMode 字段管理
    }
    
    private void ConnectAndLoadMetadata(bool showErrorDialog = true)
    {
        try
        {
            if (_useTabularMode)
            {
                ConnectTabularServer(showErrorDialog);
                return;
            }

            var database = null as string;
            var preferredPort = int.TryParse(_lastPort, out var textPort) ? textPort : ParsePortFromServer(_lastPort);
            int? selectedPort = null;

            try
            {
                var latestInstances = _detector.DiscoverInstances();
                if (latestInstances.Count > 0 || _instanceData.Count == 0)
                {
                    PopulateInstances(latestInstances, preferredPort ?? selectedPort);
                }
            }
            catch
            {
                // Ignore scan errors here; we still attempt with existing candidates.
            }

            var candidates = BuildConnectionCandidates(preferredPort, selectedPort);
            if (candidates.Count == 0)
            {
                foreach (var fallbackPort in DiscoverWorkspaceFallbackPorts())
                {
                    if (!candidates.Contains(fallbackPort))
                    {
                        candidates.Add(fallbackPort);
                    }
                }
            }

            if (candidates.Count == 0)
            {
                if (showErrorDialog) MessageBox.Show("无法解析模型端口。", Text, MessageBoxButtons.OK, MessageBoxIcon.Warning);
                SetStatus("连接失败：无法解析端口。");
                return;
            }

            var errors = new List<string>();
            foreach (var port in candidates)
            {
                try
                {
                    _model = _reader.ReadMetadata(port, database);
                    _currentPort = port;
                    _lastPort = port.ToString();
                    SelectInstanceByPort(port);
                    RenderMetadata(_model);
                    ClearPendingPlan();
                    ShowPage("dashboard");
                    SetStatus($"已成功连接到模型: {_model.DatabaseName}");
                    Log("system", $"已连接到 {_model.DatabaseName} (port={port})");
                    return;
                }
                catch (Exception ex)
                {
                    errors.Add($"[{port}] {ex.Message}");
                }
            }

            var attempts = string.Join(", ", candidates);
            var lastError = errors.Count > 0 ? errors[^1] : "未知错误";
            throw new InvalidOperationException($"无法建立连接。已尝试端口: {attempts}。最后错误: {lastError}");
        }
        catch (Exception ex)
        {
            if (showErrorDialog) MessageBox.Show($"连接失败:\n{ex.Message}", Text, MessageBoxButtons.OK, MessageBoxIcon.Error);
            SetStatus($"连接失败: {ex.Message}");
        }
    }

    private void ConnectTabularServer(bool showErrorDialog = true)
    {
        try
        {
            var serverAddr = _lastTabularServer;
            if (string.IsNullOrWhiteSpace(serverAddr))
            {
                if (showErrorDialog) MessageBox.Show("请输入服务器地址。", Text, MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            var connStr = $"DataSource={serverAddr};";
            _model = _reader.ReadMetadata(connStr, null);
            _currentPort = null;
            RenderMetadata(_model);
            ClearPendingPlan();
            ShowPage("dashboard");
            SetStatus($"已连接到 Tabular Server: {_model.DatabaseName} ({serverAddr})");
            Log("system", $"已连接到 {_model.DatabaseName} (server={serverAddr})");
        }
        catch (Exception ex)
        {
            if (showErrorDialog) MessageBox.Show($"连接失败:\n{ex.Message}", Text, MessageBoxButtons.OK, MessageBoxIcon.Error);
            SetStatus($"连接失败: {ex.Message}");
        }
    }

    private async Task SendPromptAsync()
    {
        var prompt = _chatInput.Text.Trim();
        if (string.IsNullOrWhiteSpace(prompt)) return;
        if (_model is null) { MessageBox.Show("请先连接到一个模型。", Text, MessageBoxButtons.OK, MessageBoxIcon.Warning); return; }
        _chatInput.Clear();
        Log("user", prompt);
        _chatCts?.Cancel();
        _chatCts?.Dispose();
        _chatCts = new CancellationTokenSource();
        try
        {
            var msgs = new List<AiChatMessage> { new("system", MetadataPromptBuilder.BuildSystemPrompt(_settings)), new("system", MetadataPromptBuilder.BuildModelContext(_model, _settings.IncludeHiddenObjects)) };
            msgs.AddRange(_conversation.TakeLast(10));
            msgs.Add(new AiChatMessage("user", prompt));
            var reply = await _chatClient.CompleteAsync(new AiChatRequest(_settings, msgs), _chatCts.Token);
            Log("assistant", reply);
            _conversation.Add(new AiChatMessage("user", prompt));
            _conversation.Add(new AiChatMessage("assistant", reply));
            if (_conversation.Count > 40) _conversation.RemoveRange(0, _conversation.Count - 40);
            if (_settings.AllowModelChanges && AbiActionPlanParser.TryExtract(reply, out var plan, out var preview, out _))
            {
                SetPendingPlan(plan, preview);
                ShowPage("writeback");
                Log("system", $"检测到 {plan.Actions.Count} 项可执行计划。请在“变更执行”页面中检查。");
            }
            SetStatus("AI 回复已收到。");
        }
        catch (Exception ex) { Log("system", $"AI 请求失败: {ex.Message}"); SetStatus("AI 请求失败。"); }
    }

    private void SetPendingPlan(AbiActionPlan plan, string preview)
    {
        _pendingPlan = plan;
        _planPreview.Text = preview;
        _planActions.Items.Clear();
        for (var i = 0; i < plan.Actions.Count; i++) _planActions.Items.Add(new PendingItem(i, plan.Actions[i], $"{i + 1}. {plan.Actions[i].Type} | {ActionTarget(plan.Actions[i])}"), true);
    }

    private void ClearPendingPlan()
    {
        _pendingPlan = null;
        _planPreview.Clear();
        _planActions.Items.Clear();
    }

    private void SetPlanChecks(bool value)
    {
        for (var i = 0; i < _planActions.Items.Count; i++) _planActions.SetItemChecked(i, value);
    }
    
    private void ExecutePendingPlan()
    {
        if (!_settings.AllowModelChanges) { MessageBox.Show("写回功能已在设置中禁用。", Text, MessageBoxButtons.OK, MessageBoxIcon.Warning); return; }
        if (!_currentPort.HasValue || _model is null) { MessageBox.Show("未连接到模型。", Text, MessageBoxButtons.OK, MessageBoxIcon.Warning); return; }
        var checkedItems = _planActions.CheckedItems.OfType<PendingItem>().ToList();
        if (!checkedItems.Any()) { MessageBox.Show("请至少选择一个要执行的动作。", Text, MessageBoxButtons.OK, MessageBoxIcon.Warning); return; }
        var plan = new AbiActionPlan(_pendingPlan?.Summary ?? "Plan", checkedItems.OrderBy(x => x.Index).Select(x => x.Action).ToList());
        try
        {
            var analysis = _writer.AnalyzeActions(_currentPort.Value, _model.DatabaseName, plan);
            if (analysis.HasErrors)
            {
                MessageBox.Show("预检失败:\n\n" + string.Join("\n", analysis.Errors.Take(12).Select(e => "- " + e)), Text, MessageBoxButtons.OK, MessageBoxIcon.Error);
                Log("system", "预检失败:\n" + string.Join("\n", analysis.Errors.Select(e => "- " + e)));
                return;
            }
            var confirm = MessageBox.Show(PreflightSummary(plan, analysis), "确认执行", MessageBoxButtons.YesNo, analysis.Warnings.Count > 0 ? MessageBoxIcon.Warning : MessageBoxIcon.Question);
            if (confirm != DialogResult.Yes) return;
            var (snapshot, rollback) = CreateBackup(plan);
            Log("system", $"执行前已创建备份:\n- {snapshot}\n- {rollback}");
            var results = _writer.ApplyActions(_currentPort.Value, _model.DatabaseName, plan);
            Log("system", "执行结果:\n" + string.Join("\n", results.Select(r => "- " + r)));
            ClearPendingPlan();
            ConnectAndLoadMetadata();
            RefreshBackupList();
        }
        catch (Exception ex) { Log("system", $"执行失败: {ex.Message}"); }
    }
    
    private (string Snapshot, string Rollback) CreateBackup(AbiActionPlan plan)
    {
        var before = _reader.ReadMetadata(_currentPort!.Value, _model!.DatabaseName);
        var dir = BuildBackupDirectory();
        var ts = DateTime.Now.ToString("yyyyMMdd-HHmmss");
        var safe = SafeName(before.DatabaseName);
        var snapshot = Path.Combine(dir, $"{safe}-{_currentPort}-before-{ts}.json");
        var rollback = Path.Combine(dir, $"{safe}-{_currentPort}-rollback-{ts}.json");
        EnsureDir(snapshot);
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
                    var old = FindMeasure(before, a.Table, a.Name);
                    actions.Add(old is null ? new AbiModelAction("delete_measure", Table: a.Table, Name: a.Name, Reason: "回滚：删除新增度量值") : new AbiModelAction("create_or_update_measure", Table: a.Table, Name: old.Name, Expression: old.Expression, FormatString: old.FormatString, IsHidden: old.IsHidden, Reason: "回滚：恢复原始度量值"));
                    break;
                case "delete_measure":
                    var del = FindMeasure(before, a.Table, a.Name);
                    if (del is not null) actions.Add(new AbiModelAction("create_or_update_measure", Table: a.Table, Name: del.Name, Expression: del.Expression, FormatString: del.FormatString, IsHidden: del.IsHidden, Reason: "回滚：恢复已删除的度量值"));
                    break;
                case "create_relationship":
                    var rel0 = FindRel(before, a);
                    actions.Add(rel0 is null ? new AbiModelAction("delete_relationship", Name: !string.IsNullOrWhiteSpace(a.Name) ? a.Name : RelDefaultName(a), FromTable: a.FromTable, FromColumn: a.FromColumn, ToTable: a.ToTable, ToColumn: a.ToColumn, Reason: "回滚：删除新增关系") : new AbiModelAction("create_relationship", Name: rel0.Name, FromTable: rel0.FromTable, FromColumn: rel0.FromColumn, ToTable: rel0.ToTable, ToColumn: rel0.ToColumn, CrossFilterDirection: rel0.CrossFilterDirection, IsActive: rel0.IsActive, Reason: "回滚：恢复原始关系"));
                    break;
                case "delete_relationship":
                    var rel1 = FindRel(before, a);
                    if (rel1 is not null) actions.Add(new AbiModelAction("create_relationship", Name: rel1.Name, FromTable: rel1.FromTable, FromColumn: rel1.FromColumn, ToTable: rel1.ToTable, ToColumn: rel1.ToColumn, CrossFilterDirection: rel1.CrossFilterDirection, IsActive: rel1.IsActive, Reason: "回滚：恢复已删除的关系"));
                    break;
            }
        }
        return new AbiActionPlan($"回滚计划生成于 {DateTime.Now:yyyy-MM-dd HH:mm:ss}", actions);
    }

    private void RefreshBackupList()
    {
        _backups.BeginUpdate();
        _backups.Items.Clear();
        var root = BackupRoot();
        Directory.CreateDirectory(root);
        var files = Directory.EnumerateFiles(root, "*.json", SearchOption.AllDirectories).Select(p => new FileInfo(p)).OrderByDescending(f => f.LastWriteTimeUtc);
        foreach (var f in files)
        {
            var type = f.Name.Contains("-rollback-", StringComparison.OrdinalIgnoreCase) ? "回滚" : f.Name.Contains("-before-", StringComparison.OrdinalIgnoreCase) ? "快照" : "其他";
            var item = new ListViewItem(f.LastWriteTime.ToString("yyyy-MM-dd HH:mm:ss"));
            item.SubItems.Add(type); item.SubItems.Add(f.Name); item.SubItems.Add($"{Math.Max(1, f.Length / 1024)} KB"); item.SubItems.Add(f.DirectoryName ?? string.Empty); item.Tag = f.FullName;
            _backups.Items.Add(item);
        }
        _backups.EndUpdate();
    }

    private void OpenBackupRootFolder() => Process.Start(new ProcessStartInfo { FileName = "explorer.exe", Arguments = $"\"{BackupRoot()}\"", UseShellExecute = true });
    
    private void RunSelectedBackupRollback()
    {
        if (_backups.SelectedItems.Count == 0) { MessageBox.Show("请选择一个回滚文件。", Text, MessageBoxButtons.OK, MessageBoxIcon.Information); return; }
        if (!string.Equals(_backups.SelectedItems[0].SubItems[1].Text, "回滚", StringComparison.OrdinalIgnoreCase)) { MessageBox.Show("所选文件不是一个回滚文件。", Text, MessageBoxButtons.OK, MessageBoxIcon.Information); return; }
        var path = _backups.SelectedItems[0].Tag?.ToString();
        if (!string.IsNullOrWhiteSpace(path)) ExecuteRollback(path);
    }
    
    private void ExecuteRollback(string rollbackPath)
    {
        if (!_currentPort.HasValue || _model is null) { MessageBox.Show("未连接到模型。", Text, MessageBoxButtons.OK, MessageBoxIcon.Warning); return; }
        try
        {
            var plan = AbiActionPlanStorage.Load(rollbackPath);
            if (plan.Actions.Count == 0) return;
            var analysis = _writer.AnalyzeActions(_currentPort.Value, _model.DatabaseName, plan);
            if (analysis.HasErrors) { MessageBox.Show("回滚预检失败:\n\n" + string.Join("\n", analysis.Errors.Take(12).Select(e => "- " + e)), Text, MessageBoxButtons.OK, MessageBoxIcon.Error); return; }
            var confirm = MessageBox.Show(PreflightSummary(plan, analysis), "确认回滚", MessageBoxButtons.YesNo, analysis.Warnings.Count > 0 ? MessageBoxIcon.Warning : MessageBoxIcon.Question);
            if (confirm != DialogResult.Yes) return;
            var (snapshot, reverse) = CreateBackup(plan);
            Log("system", $"回滚前已创建备份:\n- {snapshot}\n- {reverse}");
            var results = _writer.ApplyActions(_currentPort.Value, _model.DatabaseName, plan);
            Log("system", "回滚结果:\n" + string.Join("\n", results.Select(r => "- " + r)));
            ConnectAndLoadMetadata();
            RefreshBackupList();
        }
        catch (Exception ex) { Log("system", $"回滚失败: {ex.Message}"); }
    }
    
    // 节点类型标记，存在 TreeNode.Tag 里
    private const string TagDatabase = "db";
    private const string TagTableGroup = "tbl-group";
    private const string TagTable = "tbl";
    private const string TagTableHidden = "tbl-hidden";
    private const string TagColGroup = "col-group";
    private const string TagColText = "col-text";
    private const string TagColNum = "col-num";
    private const string TagColDate = "col-date";
    private const string TagColBool = "col-bool";
    private const string TagColCalc = "col-calc";
    private const string TagColHidden = "col-hidden";
    private const string TagMeasureGroup = "msr-group";
    private const string TagMeasure = "msr";
    private const string TagMeasureHidden = "msr-hidden";
    private const string TagRelGroup = "rel-group";
    private const string TagRel = "rel";

    private static void TreeDrawNode(object? sender, DrawTreeNodeEventArgs e)
    {
        var node = e.Node;
        if (node is null) { e.DrawDefault = true; return; }
        var tv = sender as TreeView;
        if (tv is null) { e.DrawDefault = true; return; }
        var g = e.Graphics;
        var bounds = e.Bounds;
        if (bounds.Height <= 0) { e.DrawDefault = true; return; }

        // OwnerDrawText 模式下 bounds.Right 只到文字测量宽度，用控件实际宽度
        var fullRight = tv.ClientSize.Width - 2;
        var fullBounds = new Rectangle(bounds.X, bounds.Y, Math.Max(0, fullRight - bounds.X), bounds.Height);

        var tag = node.Tag as string ?? string.Empty;
        // DAX Studio 配色方案
        var iconColor = tag switch
        {
            TagDatabase                                   => Color.FromArgb(0, 114, 198),
            TagTableGroup or TagColGroup
                or TagMeasureGroup or TagRelGroup         => Color.FromArgb(110, 110, 110),
            TagTable                                      => Color.FromArgb(242, 200, 17),   // DAX Studio 黄色表图标
            TagTableHidden                                => Color.FromArgb(200, 185, 120),
            TagColText                                    => Color.FromArgb(104, 160, 72),   // 绿色 ABC
            TagColNum                                     => Color.FromArgb(37, 130, 196),   // 蓝色 123
            TagColDate                                    => Color.FromArgb(130, 60, 170),   // 紫色日历
            TagColBool                                    => Color.FromArgb(76, 175, 80),    // 绿色勾
            TagColCalc                                    => Color.FromArgb(37, 130, 196),   // 蓝色 fx
            TagColHidden                                  => Color.FromArgb(170, 170, 170),
            TagMeasure                                    => Color.FromArgb(196, 40, 40),    // 红色 Σ
            TagMeasureHidden                              => Color.FromArgb(180, 130, 130),
            TagRel                                        => Color.FromArgb(90, 120, 190),
            _                                             => Color.FromArgb(140, 140, 140),
        };

        // 背景
        var isSelected = (e.State & TreeNodeStates.Selected) != 0;
        var bgColor = isSelected
            ? (tv.Focused ? SystemColors.Highlight : SystemColors.InactiveBorder)
            : Color.White;
        using (var bgBrush = new SolidBrush(bgColor))
            g.FillRectangle(bgBrush, fullBounds);

        var textColor = isSelected && tv.Focused
            ? SystemColors.HighlightText
            : Color.FromArgb(30, 30, 30);

        // 图标区域（14×14，垂直居中）
        const int iconW = 14;
        const int iconGap = 4;
        var iconX = bounds.X + 1;
        var iconY = bounds.Y + (bounds.Height - iconW) / 2;
        var iconRect = new Rectangle(iconX, iconY, iconW, iconW);

        g.SmoothingMode = SmoothingMode.AntiAlias;
        using (var iconBrush = new SolidBrush(iconColor))
        {
            DrawNodeIcon(g, tag, iconRect, iconBrush, iconColor);
        }
        g.SmoothingMode = SmoothingMode.Default;

        // 文字
        var textX = iconX + iconW + iconGap;
        var textW = fullRight - textX;
        if (textW <= 0) return;
        var textRect = new RectangleF(textX, bounds.Y, textW, bounds.Height);
        var nodeFont = node.NodeFont ?? tv.Font ?? SystemFonts.DefaultFont;
        using var textBrush = new SolidBrush(textColor);
        var tsf = new StringFormat(StringFormatFlags.NoWrap)
        {
            LineAlignment = StringAlignment.Center,
            Trimming = StringTrimming.EllipsisCharacter
        };
        g.DrawString(node.Text, nodeFont, textBrush, textRect, tsf);
    }

    private static void DrawNodeIcon(Graphics g, string tag, Rectangle r, SolidBrush brush, Color color)
    {
        var rf = new RectangleF(r.X, r.Y, r.Width, r.Height);
        var cx = r.X + r.Width / 2f;
        var cy = r.Y + r.Height / 2f;
        var sf = new StringFormat { Alignment = StringAlignment.Center, LineAlignment = StringAlignment.Center };

        switch (tag)
        {
            case TagDatabase:
            {
                // 圆柱形数据库（DAX Studio 风格：椭圆顶+矩形体）
                float w = r.Width - 2f, h = r.Height - 2f;
                float ex = r.X + 1f, ey = r.Y + 1f;
                float eh = h * 0.28f;
                using var p = new Pen(color, 1.3f);
                // 圆柱体
                g.DrawRectangle(p, ex, ey + eh / 2f, w, h - eh / 2f);
                // 顶部椭圆
                g.DrawEllipse(p, ex, ey, w, eh);
                // 遮住矩形顶边（用背景色填充椭圆下半）
                using var bg = new SolidBrush(Color.White);
                g.FillRectangle(bg, ex + 1, ey + eh / 2f, w - 2, 2);
                break;
            }

            case TagTable:
            case TagTableHidden:
            {
                // DAX Studio 标志性黄色表格图标：实心矩形 + 白色网格线
                g.FillRectangle(brush, r);
                using var gridPen = new Pen(Color.White, 1f);
                // 横线（分割表头和内容）
                float headerH = r.Height * 0.38f;
                g.DrawLine(gridPen, r.X, r.Y + headerH, r.Right, r.Y + headerH);
                // 竖线（分割列）
                float colX = r.X + r.Width * 0.45f;
                g.DrawLine(gridPen, colX, r.Y + headerH, colX, r.Bottom);
                break;
            }

            case TagTableGroup:
            case TagColGroup:
            case TagMeasureGroup:
            case TagRelGroup:
            {
                // 三条横线，左侧有小圆点
                using var p = new Pen(color, 1.4f);
                float[] ys = { r.Y + r.Height * 0.25f, r.Y + r.Height * 0.5f, r.Y + r.Height * 0.75f };
                foreach (var y in ys)
                {
                    g.FillEllipse(brush, r.X, y - 1.5f, 3f, 3f);
                    g.DrawLine(p, r.X + 5, y, r.Right - 1, y);
                }
                break;
            }

            case TagColText:
            {
                // 绿色 "ABC"（DAX Studio 文本列）
                using var f = new Font("Segoe UI", r.Height * 0.62f, FontStyle.Bold, GraphicsUnit.Pixel);
                g.DrawString("ABC", f, brush, rf, sf);
                break;
            }

            case TagColNum:
            {
                // 蓝色 "123"（DAX Studio 数字列）
                using var f = new Font("Segoe UI", r.Height * 0.62f, FontStyle.Bold, GraphicsUnit.Pixel);
                g.DrawString("123", f, brush, rf, sf);
                break;
            }

            case TagColDate:
            {
                // 紫色日历图标
                float bx = r.X + 0.5f, by = r.Y + r.Height * 0.2f;
                float bw = r.Width - 1f, bh = r.Height * 0.75f;
                using var p = new Pen(color, 1.2f);
                g.DrawRectangle(p, bx, by, bw, bh);
                // 顶部分割线（表头）
                float hdrY = by + bh * 0.32f;
                g.DrawLine(p, bx, hdrY, bx + bw, hdrY);
                // 两个挂钩
                using var hookPen = new Pen(color, 1.5f);
                g.DrawLine(hookPen, cx - 2.5f, r.Y, cx - 2.5f, by + 2);
                g.DrawLine(hookPen, cx + 2.5f, r.Y, cx + 2.5f, by + 2);
                break;
            }

            case TagColBool:
            {
                // 绿色勾选框
                using var p = new Pen(color, 1.3f);
                g.DrawRectangle(p, r.X + 0.5f, r.Y + 0.5f, r.Width - 2f, r.Height - 2f);
                using var cp = new Pen(color, 1.8f);
                cp.StartCap = System.Drawing.Drawing2D.LineCap.Round;
                cp.EndCap = System.Drawing.Drawing2D.LineCap.Round;
                g.DrawLine(cp, r.X + 2.5f, cy, cx - 0.5f, r.Bottom - 2.5f);
                g.DrawLine(cp, cx - 0.5f, r.Bottom - 2.5f, r.Right - 2f, r.Y + 2.5f);
                break;
            }

            case TagColCalc:
            {
                // 蓝色 "fx"（计算列）
                using var f = new Font("Segoe UI", r.Height * 0.62f, FontStyle.Italic, GraphicsUnit.Pixel);
                g.DrawString("fx", f, brush, rf, sf);
                break;
            }

            case TagColHidden:
            {
                // 灰色列（虚线边框 + 斜线）
                using var p = new Pen(color, 1f) { DashStyle = System.Drawing.Drawing2D.DashStyle.Dot };
                g.DrawRectangle(p, r.X + 1, r.Y + 1, r.Width - 3, r.Height - 3);
                break;
            }

            case TagMeasure:
            case TagMeasureHidden:
            {
                // 红色 "Σ"（DAX Studio 度量值）
                using var f = new Font("Segoe UI", r.Height * 0.78f, FontStyle.Bold, GraphicsUnit.Pixel);
                g.DrawString("Σ", f, brush, rf, sf);
                break;
            }

            case TagRel:
            {
                // 双向箭头（关系）
                float mid = cy;
                float arrowSize = 3f;
                using var p = new Pen(color, 1.5f);
                p.StartCap = System.Drawing.Drawing2D.LineCap.Round;
                p.EndCap = System.Drawing.Drawing2D.LineCap.Round;
                g.DrawLine(p, r.X + 1, mid, r.Right - 2, mid);
                // 左箭头
                g.DrawLine(p, r.X + 1, mid, r.X + 1 + arrowSize, mid - arrowSize);
                g.DrawLine(p, r.X + 1, mid, r.X + 1 + arrowSize, mid + arrowSize);
                // 右箭头
                g.DrawLine(p, r.Right - 2, mid, r.Right - 2 - arrowSize, mid - arrowSize);
                g.DrawLine(p, r.Right - 2, mid, r.Right - 2 - arrowSize, mid + arrowSize);
                break;
            }

            default:
                g.FillEllipse(brush, r.X + 3, r.Y + 3, r.Width - 6, r.Height - 6);
                break;
        }
    }

    private void RenderMetadata(ModelMetadata model)
    {
        _tree.BeginUpdate();
        _tree.Nodes.Clear();

        var root = new TreeNode($"{model.DatabaseName}  (兼容级别 {model.CompatibilityLevel})") { Tag = TagDatabase };

        var tablesNode = new TreeNode($"表  ({model.Tables.Count})") { Tag = TagTableGroup };
        foreach (var t in model.Tables.OrderBy(x => x.Name, StringComparer.OrdinalIgnoreCase))
        {
            var tTag = t.IsHidden ? TagTableHidden : TagTable;
            var tNode = new TreeNode(t.IsHidden ? $"{t.Name}  (隐藏)" : t.Name) { Tag = tTag };

            var colsNode = new TreeNode($"列  ({t.Columns.Count})") { Tag = TagColGroup };
            foreach (var c in t.Columns.OrderBy(x => x.Name, StringComparer.OrdinalIgnoreCase))
            {
                var cTag = c.IsHidden ? TagColHidden
                    : c.ColumnType == "Calculated" ? TagColCalc
                    : c.DataType is "Int64" or "Double" or "Decimal" ? TagColNum
                    : c.DataType == "DateTime" ? TagColDate
                    : c.DataType == "Boolean" ? TagColBool
                    : TagColText;
                var cLabel = $"{c.Name}  [{c.DataType}]" + (c.IsHidden ? "  (隐藏)" : "");
                colsNode.Nodes.Add(new TreeNode(cLabel) { Tag = cTag });
            }

            var measuresNode = new TreeNode($"度量值  ({t.Measures.Count})") { Tag = TagMeasureGroup };
            foreach (var m in t.Measures.OrderBy(x => x.Name, StringComparer.OrdinalIgnoreCase))
            {
                var mTag = m.IsHidden ? TagMeasureHidden : TagMeasure;
                measuresNode.Nodes.Add(new TreeNode(m.IsHidden ? $"{m.Name}  (隐藏)" : m.Name) { Tag = mTag });
            }

            tNode.Nodes.Add(colsNode);
            tNode.Nodes.Add(measuresNode);
            tablesNode.Nodes.Add(tNode);
        }

        var relsNode = new TreeNode($"关系  ({model.Relationships.Count})") { Tag = TagRelGroup };
        foreach (var r in model.Relationships.OrderBy(x => x.Name, StringComparer.OrdinalIgnoreCase))
            relsNode.Nodes.Add(new TreeNode($"{r.FromTable}[{r.FromColumn}]  →  {r.ToTable}[{r.ToColumn}]") { Tag = TagRel });

        root.Nodes.Add(tablesNode);
        root.Nodes.Add(relsNode);
        _tree.Nodes.Add(root);
        root.Expand();
        tablesNode.Expand();
        _tree.EndUpdate();
    }
    
    private void Log(string role, string msg)
    {
        if (InvokeRequired) { Invoke(() => Log(role, msg)); return; }
        var roleName = role switch { "assistant" => "AI", "user" => "用户", _ => "系统" };
        var roleColor = role switch { "assistant" => Color.FromArgb(0, 122, 204), "user" => Color.FromArgb(13, 107, 34), _ => Color.DimGray };
        _chatLog.SelectionStart = _chatLog.TextLength;
        _chatLog.SelectionFont = new Font(_chatLog.Font, FontStyle.Bold);
        _chatLog.SelectionColor = roleColor;
        _chatLog.AppendText($"[{DateTime.Now:HH:mm:ss}] {roleName}:{Environment.NewLine}");
        if (role == "assistant")
            AppendMarkdown(msg);
        else
        {
            _chatLog.SelectionStart = _chatLog.TextLength;
            _chatLog.SelectionFont = _chatLog.Font;
            _chatLog.SelectionColor = TextPrimary;
            _chatLog.AppendText(msg);
        }
        _chatLog.AppendText(Environment.NewLine + Environment.NewLine);
        _chatLog.ScrollToCaret();
    }

    private void AppendMarkdown(string text)
    {
        var baseFont = _chatLog.Font;
        var boldFont = new Font(baseFont, FontStyle.Bold);
        var italicFont = new Font(baseFont, FontStyle.Italic);
        var codeFont = new Font("Consolas", baseFont.Size - 0.5f);
        var codeBg = Color.FromArgb(245, 245, 245);
        var codeColor = Color.FromArgb(180, 40, 40);
        var headingColor = Color.FromArgb(0, 90, 160);

        var lines = text.Split('\n');
        bool inCodeBlock = false;

        foreach (var rawLine in lines)
        {
            var line = rawLine.TrimEnd('\r');

            // 代码块 ```
            if (line.TrimStart().StartsWith("```"))
            {
                inCodeBlock = !inCodeBlock;
                _chatLog.AppendText(Environment.NewLine);
                continue;
            }

            if (inCodeBlock)
            {
                _chatLog.SelectionStart = _chatLog.TextLength;
                _chatLog.SelectionFont = codeFont;
                _chatLog.SelectionColor = codeColor;
                _chatLog.SelectionBackColor = codeBg;
                _chatLog.AppendText(line + Environment.NewLine);
                _chatLog.SelectionBackColor = _chatLog.BackColor;
                continue;
            }

            // 标题 ## / #
            if (line.StartsWith("### "))
            {
                AppendRichLine(line[4..], new Font(baseFont.FontFamily, baseFont.Size + 0.5f, FontStyle.Bold), headingColor);
                continue;
            }
            if (line.StartsWith("## "))
            {
                AppendRichLine(line[3..], new Font(baseFont.FontFamily, baseFont.Size + 1.5f, FontStyle.Bold), headingColor);
                continue;
            }
            if (line.StartsWith("# "))
            {
                AppendRichLine(line[2..], new Font(baseFont.FontFamily, baseFont.Size + 2.5f, FontStyle.Bold), headingColor);
                continue;
            }

            // 列表项 - / *
            string lineContent = line;
            string prefix = "";
            if (System.Text.RegularExpressions.Regex.IsMatch(line, @"^\s*[-*]\s"))
            {
                var indent = line.Length - line.TrimStart().Length;
                prefix = new string(' ', indent) + "• ";
                lineContent = System.Text.RegularExpressions.Regex.Replace(line.TrimStart(), @"^[-*]\s", "");
            }
            else if (System.Text.RegularExpressions.Regex.IsMatch(line, @"^\s*\d+\.\s"))
            {
                var m = System.Text.RegularExpressions.Regex.Match(line, @"^(\s*\d+\.\s)(.*)");
                prefix = m.Groups[1].Value;
                lineContent = m.Groups[2].Value;
            }

            if (!string.IsNullOrEmpty(prefix))
            {
                _chatLog.SelectionStart = _chatLog.TextLength;
                _chatLog.SelectionFont = baseFont;
                _chatLog.SelectionColor = TextPrimary;
                _chatLog.AppendText(prefix);
            }

            // 行内解析：**bold**、*italic*、`code`
            AppendInlineMarkdown(lineContent, baseFont, boldFont, italicFont, codeFont, codeColor);
            _chatLog.AppendText(Environment.NewLine);
        }
    }

    private void AppendRichLine(string text, Font font, Color color)
    {
        _chatLog.SelectionStart = _chatLog.TextLength;
        _chatLog.SelectionFont = font;
        _chatLog.SelectionColor = color;
        _chatLog.AppendText(text + Environment.NewLine);
    }

    private void AppendInlineMarkdown(string text, Font baseFont, Font boldFont, Font italicFont, Font codeFont, Color codeColor)
    {
        // 匹配 **bold**、*italic*、`code`
        var pattern = @"(\*\*(.+?)\*\*|\*(.+?)\*|`([^`]+)`)";
        var parts = System.Text.RegularExpressions.Regex.Split(text, pattern);
        // Regex.Split with groups returns interleaved captures; use Matches instead
        int pos = 0;
        var matches = System.Text.RegularExpressions.Regex.Matches(text, pattern);
        foreach (System.Text.RegularExpressions.Match m in matches)
        {
            // 普通文本
            if (m.Index > pos)
            {
                _chatLog.SelectionStart = _chatLog.TextLength;
                _chatLog.SelectionFont = baseFont;
                _chatLog.SelectionColor = TextPrimary;
                _chatLog.AppendText(text[pos..m.Index]);
            }
            // 格式化文本
            if (m.Value.StartsWith("**"))
            {
                _chatLog.SelectionStart = _chatLog.TextLength;
                _chatLog.SelectionFont = boldFont;
                _chatLog.SelectionColor = TextPrimary;
                _chatLog.AppendText(m.Groups[2].Value);
            }
            else if (m.Value.StartsWith("*"))
            {
                _chatLog.SelectionStart = _chatLog.TextLength;
                _chatLog.SelectionFont = italicFont;
                _chatLog.SelectionColor = TextPrimary;
                _chatLog.AppendText(m.Groups[3].Value);
            }
            else if (m.Value.StartsWith("`"))
            {
                _chatLog.SelectionStart = _chatLog.TextLength;
                _chatLog.SelectionFont = codeFont;
                _chatLog.SelectionColor = codeColor;
                _chatLog.AppendText(m.Groups[4].Value);
            }
            pos = m.Index + m.Length;
        }
        // 剩余普通文本
        if (pos < text.Length)
        {
            _chatLog.SelectionStart = _chatLog.TextLength;
            _chatLog.SelectionFont = baseFont;
            _chatLog.SelectionColor = TextPrimary;
            _chatLog.AppendText(text[pos..]);
        }
    }

    #region Helpers
    private void PopulateInstances(IReadOnlyList<PowerBiInstanceInfo> instances, int? preferredPort = null)
    {
        _instanceData = instances;

        // 如果有偏好端口，更新 _lastPort
        if (preferredPort.HasValue)
        {
            _lastPort = preferredPort.Value.ToString();
        }
        else if (_instanceData.Count > 0)
        {
            // 优先选有 DesktopPid 的实例
            var best = _instanceData.FirstOrDefault(x => x.DesktopPid > 0) ?? _instanceData[0];
            if (string.IsNullOrEmpty(_lastPort))
                _lastPort = best.Port.ToString();
        }
    }

    private List<int> BuildConnectionCandidates(int? preferredPort, int? selectedPort)
    {
        var result = new List<int>();
        var dedup = new HashSet<int>();

        void AddPort(int? value)
        {
            if (!value.HasValue || value.Value <= 0)
            {
                return;
            }

            if (dedup.Add(value.Value))
            {
                result.Add(value.Value);
            }
        }

        AddPort(preferredPort);
        AddPort(selectedPort);

        foreach (var instance in _instanceData
                     .OrderByDescending(x => x.DesktopPid > 0)
                     .ThenByDescending(x => x.LastSeenUtc))
        {
            AddPort(instance.Port);
        }

        return result;
    }

    private static List<int> DiscoverWorkspaceFallbackPorts()
    {
        var result = new List<(int Port, DateTime LastSeenUtc)>();
        foreach (var root in WorkspaceRoots())
        {
            if (!Directory.Exists(root))
            {
                continue;
            }

            IEnumerable<string> workspaces;
            try
            {
                workspaces = Directory.EnumerateDirectories(root, "*", SearchOption.TopDirectoryOnly);
            }
            catch
            {
                continue;
            }

            foreach (var workspace in workspaces)
            {
                var candidates = new[]
                {
                    Path.Combine(workspace, "msmdsrv.port.txt"),
                    Path.Combine(workspace, "Data", "msmdsrv.port.txt")
                };

                foreach (var candidate in candidates)
                {
                    if (!File.Exists(candidate))
                    {
                        continue;
                    }

                    try
                    {
                        var text = File.ReadAllText(candidate).Replace("\0", string.Empty).Trim();
                        if (int.TryParse(text, out var port) && port > 0 && port <= 65535)
                        {
                            var lastSeen = Directory.GetLastWriteTimeUtc(workspace);
                            result.Add((port, lastSeen));
                            break;
                        }

                        var matches = Regex.Matches(text, "\\d{4,6}");
                        if (matches.Count > 0 && int.TryParse(matches[^1].Value, out var parsed) && parsed > 0 && parsed <= 65535)
                        {
                            var lastSeen = Directory.GetLastWriteTimeUtc(workspace);
                            result.Add((parsed, lastSeen));
                            break;
                        }
                    }
                    catch
                    {
                        // Ignore bad workspace entries.
                    }
                }
            }
        }

        return result
            .OrderByDescending(x => x.LastSeenUtc)
            .Select(x => x.Port)
            .Distinct()
            .ToList();
    }

    private static IReadOnlyList<string> WorkspaceRoots()
    {
        var local = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
        var userProfile = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
        var roots = new List<string>
        {
            Path.Combine(local, "Microsoft", "Power BI Desktop", "AnalysisServicesWorkspaces"),
            Path.Combine(local, "Microsoft", "Power BI Desktop Store App", "AnalysisServicesWorkspaces"),
            Path.Combine(local, "Microsoft", "Power BI Desktop SSRS", "AnalysisServicesWorkspaces"),
            Path.Combine(userProfile, "Microsoft", "Power BI Desktop", "AnalysisServicesWorkspaces"),
            Path.Combine(userProfile, "Microsoft", "Power BI Desktop Store App", "AnalysisServicesWorkspaces"),
            Path.Combine(userProfile, "Microsoft", "Power BI Desktop SSRS", "AnalysisServicesWorkspaces")
        };

        var packagesRoot = Path.Combine(local, "Packages");
        if (Directory.Exists(packagesRoot))
        {
            foreach (var pkg in Directory.EnumerateDirectories(packagesRoot, "Microsoft.MicrosoftPowerBIDesktop_*", SearchOption.TopDirectoryOnly))
            {
                roots.Add(Path.Combine(pkg, "LocalCache", "Local", "Microsoft", "Power BI Desktop", "AnalysisServicesWorkspaces"));
                roots.Add(Path.Combine(pkg, "LocalState", "AnalysisServicesWorkspaces"));
            }
        }

        return roots
            .Where(path => !string.IsNullOrWhiteSpace(path))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();
    }

    private void SelectInstanceByPort(int port)
    {
        _lastPort = port.ToString();
    }

    private static int? ParsePortFromServer(string server)
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

    private void AttachInitialSplitterLayout(SplitContainer split)
    {
        void OnLayout(object? s, LayoutEventArgs e)
        {
            var preferred = ComputePreferredLeftWidth(split);
            if (preferred <= 0) return;
            split.Layout -= OnLayout;
            TrySetSplitterDistance(split, preferred);
        }
        split.Layout += OnLayout;
    }

    private void ApplyPrimarySplitters()
    {
        TrySetSplitterDistance(_dashboardSplitter, ComputePreferredLeftWidth(_dashboardSplitter));
    }

    private static int ComputePreferredLeftWidth(SplitContainer split)
    {
        if (split is null)
        {
            return 0;
        }

        var available = split.ClientSize.Width - split.SplitterWidth;
        if (available <= 0)
        {
            return 0;
        }

        return (int)Math.Round(available * LeftPaneWidthRatio, MidpointRounding.AwayFromZero);
    }

    private static void TrySetSplitterDistance(SplitContainer split, int preferred)
    {
        if (split is null || preferred <= 0)
        {
            return;
        }

        var available = split.ClientSize.Width - split.SplitterWidth;
        if (available <= 0)
        {
            return;
        }

        var minLeft = Math.Max(split.Panel1MinSize, 200);
        var minRight = Math.Max(split.Panel2MinSize, DesiredRightPaneMinWidth);
        var maxLeft = Math.Max(minLeft, available - minRight);
        var safeDistance = Math.Max(minLeft, Math.Min(preferred, maxLeft));
        if (safeDistance > minLeft - 1 && safeDistance < available - minRight + 1)
        {
            try
            {
                split.SplitterDistance = safeDistance;
            }
            catch
            {
                // Ignore intermittent layout timing issues.
            }
        }
    }

    private static string? NormalizeExternalToolValue(string? value) => string.IsNullOrWhiteSpace(value) || value.StartsWith('%') ? null : value.Trim();
    private static string BuildDefaultOutputPath() => Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "PBIClaw", $"abi-model-{DateTime.Now:yyyyMMdd-HHmmss}.json");
    private static string BackupRoot() => Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "PBIClaw", "backups");
    private static string BuildBackupDirectory() => Path.Combine(BackupRoot(), DateTime.Now.ToString("yyyyMMdd"));
    private static void EnsureDir(string outputPath) => Directory.CreateDirectory(Path.GetDirectoryName(outputPath)!);
    private static string SafeName(string text) => string.IsNullOrWhiteSpace(text) ? "model" : new string(text.Select(ch => Path.GetInvalidFileNameChars().Contains(ch) ? '_' : ch).ToArray());
    private static string ActionTarget(AbiModelAction a) => a.Type switch { "create_or_update_measure" or "delete_measure" => $"{a.Table}.{a.Name}", "create_relationship" => $"{a.FromTable}.{a.FromColumn}->{a.ToTable}.{a.ToColumn}", "delete_relationship" when !string.IsNullOrWhiteSpace(a.Name) => a.Name!, "delete_relationship" => $"{a.FromTable}.{a.FromColumn}->{a.ToTable}.{a.ToColumn}", _ => a.Name ?? a.Table ?? "N/A" };

    private static string PreflightSummary(AbiActionPlan plan, TabularActionAnalysis analysis)
    {
        var lines = new List<string> { $"待执行动作: {plan.Actions.Count}", $"信息: {analysis.Infos.Count}", $"警告: {analysis.Warnings.Count}" };
        if (analysis.Warnings.Any()) { lines.Add(string.Empty); lines.Add("主要警告:"); lines.AddRange(analysis.Warnings.Take(8).Select(w => $"- {w}")); }
        lines.Add(string.Empty); lines.Add("是否继续执行？");
        return string.Join(Environment.NewLine, lines);
    }
    
    private static MeasureMetadata? FindMeasure(ModelMetadata model, string? table, string? measure) =>
        model.Tables.FirstOrDefault(t => t.Name.Equals(table, StringComparison.OrdinalIgnoreCase))
            ?.Measures.FirstOrDefault(m => m.Name.Equals(measure, StringComparison.OrdinalIgnoreCase));

    private static RelationshipMetadata? FindRel(ModelMetadata model, AbiModelAction action)
    {
        if (!string.IsNullOrWhiteSpace(action.Name)) return model.Relationships.FirstOrDefault(r => r.Name.Equals(action.Name, StringComparison.OrdinalIgnoreCase));
        return model.Relationships.FirstOrDefault(r => r.FromTable.Equals(action.FromTable, StringComparison.OrdinalIgnoreCase) && r.FromColumn.Equals(action.FromColumn, StringComparison.OrdinalIgnoreCase) && r.ToTable.Equals(action.ToTable, StringComparison.OrdinalIgnoreCase) && r.ToColumn.Equals(action.ToColumn, StringComparison.OrdinalIgnoreCase));
    }
    
    private static string RelDefaultName(AbiModelAction a) => $"{a.FromTable ?? "T1"}_{a.FromColumn ?? "C1"}_to_{a.ToTable ?? "T2"}_{a.ToColumn ?? "C2"}";

    private void SetStatus(string text)
    {
        if (InvokeRequired) { Invoke(() => _statusLabel.Text = text); }
        else { _statusLabel.Text = text; }
    }

    private static string CurrentVersion() => Application.ProductVersion ?? "dev";

    private sealed record PendingItem(int Index, AbiModelAction Action, string Label)
    {
        public override string ToString() => Label;
    }

    private sealed class InstanceItem(PowerBiInstanceInfo instance)
    {
        public PowerBiInstanceInfo Instance { get; } = instance;
        public override string ToString() => $"Port {Instance.Port} - {(string.IsNullOrWhiteSpace(Instance.PbixPathHint) ? "(未知 PBIX)" : Path.GetFileName(Instance.PbixPathHint))}";
    }

    private class DaxStudioColorTable : ProfessionalColorTable
    {
        public override Color MenuStripGradientBegin => RibbonBg;
        public override Color MenuStripGradientEnd => RibbonBg;
        public override Color ToolStripDropDownBackground => CardBg;
        public override Color MenuItemSelected => ButtonHoverBg;
        public override Color MenuItemBorder => ButtonHoverBorder;
        public override Color MenuItemPressedGradientBegin => CardBg;
        public override Color MenuItemPressedGradientEnd => CardBg;
        public override Color SeparatorLight => CardBorder;
        public override Color SeparatorDark => CardBorder;
    }
}
#endregion
