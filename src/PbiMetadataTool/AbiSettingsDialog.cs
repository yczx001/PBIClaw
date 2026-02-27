namespace PbiMetadataTool;

internal sealed class AbiSettingsDialog : Form
{
    private readonly TabControl _tabControl;
    private readonly TextBox _baseUrlBox;
    private readonly TextBox _modelBox;
    private readonly TextBox _apiKeyBox;
    private readonly NumericUpDown _temperatureInput;
    private readonly CheckBox _allowChangesCheck;
    private readonly CheckBox _includeHiddenCheck;
    private readonly TextBox _customPromptBox;
    private readonly TextBox _quickPromptsBox;

    public AbiAssistantSettings Result { get; private set; }

    private AbiSettingsDialog(AbiAssistantSettings source)
    {
        Result = Clone(source);

        Text = "PBI Claw 设置";
        StartPosition = FormStartPosition.CenterParent;
        Width = 780;
        Height = 620;
        MinimizeBox = false;
        MaximizeBox = false;
        FormBorderStyle = FormBorderStyle.FixedDialog;
        Font = new Font("Microsoft YaHei UI", 9.5F, FontStyle.Regular, GraphicsUnit.Point);
        BackColor = Color.FromArgb(240, 240, 240);

        // Initialize controls
        _tabControl = new TabControl { Dock = DockStyle.Fill, Padding = new Point(10, 6) };
        _baseUrlBox = new TextBox { Dock = DockStyle.Fill };
        _modelBox = new TextBox { Dock = DockStyle.Fill };
        _apiKeyBox = new TextBox { Dock = DockStyle.Fill, UseSystemPasswordChar = true };
        _temperatureInput = new NumericUpDown { DecimalPlaces = 2, Minimum = 0, Maximum = 2, Increment = 0.05M, Width = 120 };
        _allowChangesCheck = new CheckBox { AutoSize = true, Text = "允许 PBI Claw 修改数据模型（高风险）" };
        _includeHiddenCheck = new CheckBox { AutoSize = true, Text = "在模型上下文中包含隐藏对象" };
        _customPromptBox = new TextBox { Multiline = true, ScrollBars = ScrollBars.Vertical, Dock = DockStyle.Fill };
        _quickPromptsBox = new TextBox { Multiline = true, ScrollBars = ScrollBars.Vertical, Dock = DockStyle.Fill };

        // Set initial values
        _baseUrlBox.Text = Result.BaseUrl;
        _modelBox.Text = Result.Model;
        _apiKeyBox.Text = Result.ApiKey;
        _temperatureInput.Value = (decimal)Math.Clamp(Result.Temperature, 0, 2);
        _allowChangesCheck.Checked = Result.AllowModelChanges;
        _includeHiddenCheck.Checked = Result.IncludeHiddenObjects;
        _customPromptBox.Text = Result.CustomSystemPrompt;
        _quickPromptsBox.Text = string.Join(Environment.NewLine, Result.QuickPrompts);

        BuildLayout();
    }

    public static bool TryEdit(IWin32Window owner, AbiAssistantSettings source, out AbiAssistantSettings updated)
    {
        using var dialog = new AbiSettingsDialog(source);
        var ok = dialog.ShowDialog(owner) == DialogResult.OK;
        updated = ok ? dialog.Result : source;
        return ok;
    }

    private void BuildLayout()
    {
        var mainLayout = new TableLayoutPanel { Dock = DockStyle.Fill, RowCount = 2, ColumnCount = 1, Padding = new Padding(10) };
        mainLayout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
        mainLayout.RowStyles.Add(new RowStyle(SizeType.AutoSize));

        // Tabs
        _tabControl.TabPages.Add(new TabPage("AI 引擎") { Controls = { BuildAiEnginePage() } });
        _tabControl.TabPages.Add(new TabPage("提示词") { Controls = { BuildPromptsPage() } });
        _tabControl.TabPages.Add(new TabPage("通用") { Controls = { BuildGeneralPage() } });
        mainLayout.Controls.Add(_tabControl, 0, 0);

        // Buttons
        var buttonPanel = new FlowLayoutPanel { Dock = DockStyle.Fill, FlowDirection = FlowDirection.RightToLeft, AutoSize = true, WrapContents = false, Padding = new Padding(0, 10, 0, 0) };
        var okButton = new Button { Text = "保存", DialogResult = DialogResult.OK, AutoSize = true, Padding = new Padding(10, 4, 10, 4) };
        var cancelButton = new Button { Text = "取消", DialogResult = DialogResult.Cancel, AutoSize = true, Padding = new Padding(10, 4, 10, 4) };
        okButton.Click += (_, _) => SaveSettings();
        buttonPanel.Controls.Add(okButton);
        buttonPanel.Controls.Add(cancelButton);
        mainLayout.Controls.Add(buttonPanel, 0, 1);

        Controls.Add(mainLayout);
        AcceptButton = okButton;
        CancelButton = cancelButton;
    }

    private Control BuildAiEnginePage()
    {
        var layout = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 2, RowCount = 5, Padding = new Padding(20) };
        layout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 130));
        layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
        for (int i = 0; i < 4; i++) layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100)); // Spacer

        layout.Controls.Add(new Label { Text = "Base URL:", Anchor = AnchorStyles.Left, AutoSize = true, TextAlign = ContentAlignment.MiddleLeft }, 0, 0);
        layout.Controls.Add(_baseUrlBox, 1, 0);
        layout.Controls.Add(new Label { Text = "模型:", Anchor = AnchorStyles.Left, AutoSize = true, TextAlign = ContentAlignment.MiddleLeft }, 0, 1);
        layout.Controls.Add(_modelBox, 1, 1);
        layout.Controls.Add(new Label { Text = "API Key:", Anchor = AnchorStyles.Left, AutoSize = true, TextAlign = ContentAlignment.MiddleLeft }, 0, 2);
        layout.Controls.Add(_apiKeyBox, 1, 2);
        layout.Controls.Add(new Label { Text = "Temperature:", Anchor = AnchorStyles.Left, AutoSize = true, TextAlign = ContentAlignment.MiddleLeft }, 0, 3);
        layout.Controls.Add(_temperatureInput, 1, 3);
        
        // Set row padding
        _baseUrlBox.Margin = new Padding(5, 8, 5, 8);
        _modelBox.Margin = new Padding(5, 8, 5, 8);
        _apiKeyBox.Margin = new Padding(5, 8, 5, 8);
        _temperatureInput.Margin = new Padding(5, 8, 5, 8);

        return layout;
    }
    
    private Control BuildPromptsPage()
    {
        var layout = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 4, Padding = new Padding(20) };
        layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        layout.RowStyles.Add(new RowStyle(SizeType.Percent, 50));
        layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        layout.RowStyles.Add(new RowStyle(SizeType.Percent, 50));

        layout.Controls.Add(new Label { Text = "自定义系统提示词:", AutoSize = true }, 0, 0);
        layout.Controls.Add(_customPromptBox, 0, 1);
        layout.Controls.Add(new Label { Text = "快捷提示词 (每行一个):", AutoSize = true, Margin = new Padding(0, 10, 0, 0)}, 0, 2);
        layout.Controls.Add(_quickPromptsBox, 0, 3);

        return layout;
    }
    
    private Control BuildGeneralPage()
    {
        var layout = new FlowLayoutPanel { Dock = DockStyle.Fill, FlowDirection = FlowDirection.TopDown, Padding = new Padding(20) };
        layout.Controls.Add(_allowChangesCheck);
        layout.Controls.Add(_includeHiddenCheck);
        _includeHiddenCheck.Margin = new Padding(3, 10, 3, 3);
        return layout;
    }

    private void SaveSettings()
    {
        if (string.IsNullOrWhiteSpace(_baseUrlBox.Text))
        {
            _tabControl.SelectedTab = _tabControl.TabPages["AI 引擎"];
            MessageBox.Show("Base URL 不能为空。", Text, MessageBoxButtons.OK, MessageBoxIcon.Warning);
            DialogResult = DialogResult.None; // Prevent closing
            return;
        }

        if (string.IsNullOrWhiteSpace(_modelBox.Text))
        {
            _tabControl.SelectedTab = _tabControl.TabPages["AI 引擎"];
            MessageBox.Show("模型名称不能为空。", Text, MessageBoxButtons.OK, MessageBoxIcon.Warning);
            DialogResult = DialogResult.None; // Prevent closing
            return;
        }

        Result = new AbiAssistantSettings
        {
            BaseUrl = _baseUrlBox.Text.Trim(),
            Model = _modelBox.Text.Trim(),
            ApiKey = _apiKeyBox.Text.Trim(),
            Temperature = (double)_temperatureInput.Value,
            AllowModelChanges = _allowChangesCheck.Checked,
            IncludeHiddenObjects = _includeHiddenCheck.Checked,
            CustomSystemPrompt = _customPromptBox.Text.Trim(),
            QuickPrompts = _quickPromptsBox.Text
                .Split(["\r", "\n"], StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
                .Distinct(StringComparer.Ordinal)
                .Take(20)
                .ToList()
        };

        if (Result.QuickPrompts.Count == 0)
        {
            Result.QuickPrompts = [ "分析当前模型并给出3个优化建议和可执行步骤。" ];
        }
    }

    private static AbiAssistantSettings Clone(AbiAssistantSettings source)
    {
        return new AbiAssistantSettings
        {
            BaseUrl = source.BaseUrl,
            Model = source.Model,
            ApiKey = source.ApiKey,
            Temperature = source.Temperature,
            AllowModelChanges = source.AllowModelChanges,
            IncludeHiddenObjects = source.IncludeHiddenObjects,
            CustomSystemPrompt = source.CustomSystemPrompt,
            QuickPrompts = [..source.QuickPrompts]
        };
    }
}
