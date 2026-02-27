using System.Diagnostics;

namespace PbiMetadataTool;

internal sealed class BackupCenterDialog : Form
{
    private readonly string _rootDirectory;
    private readonly ListView _list;
    private readonly Button _useRollbackButton;

    public string? SelectedRollbackPath { get; private set; }

    public BackupCenterDialog(string rootDirectory)
    {
        _rootDirectory = rootDirectory;

        Text = "备份中心";
        Width = 1024;
        Height = 640;
        StartPosition = FormStartPosition.CenterParent;
        MinimizeBox = false;
        MaximizeBox = false;
        Font = new Font("Microsoft YaHei UI", 9.5F, FontStyle.Regular, GraphicsUnit.Point);
        BackColor = Color.FromArgb(240, 240, 240);

        // Initialize controls
        _list = new ListView
        {
            Dock = DockStyle.Fill,
            View = View.Details,
            FullRowSelect = true,
            MultiSelect = false,
            GridLines = true
        };
        _useRollbackButton = new Button { Text = "使用此回滚", Enabled = false, DialogResult = DialogResult.OK, AutoSize = true, Padding = new Padding(10, 4, 10, 4) };

        BuildLayout();
        BindEvents();
        RefreshList();
    }

    private void BuildLayout()
    {
        _list.Columns.Add("时间", 180);
        _list.Columns.Add("类型", 100);
        _list.Columns.Add("文件名", 380);
        _list.Columns.Add("大小", 100);
        _list.Columns.Add("目录", 250);

        var mainLayout = new TableLayoutPanel { Dock = DockStyle.Fill, RowCount = 2, ColumnCount = 1, Padding = new Padding(10) };
        mainLayout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
        mainLayout.RowStyles.Add(new RowStyle(SizeType.AutoSize));

        mainLayout.Controls.Add(_list, 0, 0);

        var buttonPanel = new FlowLayoutPanel { Dock = DockStyle.Fill, FlowDirection = FlowDirection.RightToLeft, AutoSize = true, WrapContents = false, Padding = new Padding(0, 10, 0, 0) };
        var refreshButton = new Button { Text = "刷新", AutoSize = true, Padding = new Padding(10, 4, 10, 4) };
        var openFolderButton = new Button { Text = "打开目录", AutoSize = true, Padding = new Padding(10, 4, 10, 4) };
        var closeButton = new Button { Text = "关闭", DialogResult = DialogResult.Cancel, AutoSize = true, Padding = new Padding(10, 4, 10, 4) };
        
        buttonPanel.Controls.Add(_useRollbackButton);
        buttonPanel.Controls.Add(closeButton);
        buttonPanel.Controls.Add(openFolderButton);
        buttonPanel.Controls.Add(refreshButton);
        
        mainLayout.Controls.Add(buttonPanel, 0, 1);

        // Assign event handlers that are tied to the buttons
        refreshButton.Click += (_, _) => RefreshList();
        openFolderButton.Click += (_, _) => OpenRootFolder();
        _useRollbackButton.Click += (_, _) => UseSelectedRollback();

        Controls.Add(mainLayout);
        AcceptButton = _useRollbackButton;
        CancelButton = closeButton;
    }

    private void BindEvents()
    {
        _list.SelectedIndexChanged += (_, _) => SyncSelectionState();
        _list.DoubleClick += (_, _) =>
        {
            if (IsSelectedRollback())
            {
                UseSelectedRollback();
            }
        };
    }

    private void RefreshList()
    {
        _list.BeginUpdate();
        _list.Items.Clear();

        EnsureRootDirectory();
        var files = Directory.EnumerateFiles(_rootDirectory, "*.json", SearchOption.AllDirectories)
            .Select(path => new FileInfo(path))
            .OrderByDescending(x => x.LastWriteTimeUtc)
            .ToList();

        foreach (var file in files)
        {
            var type = ResolveType(file.Name);
            var item = new ListViewItem(file.LastWriteTime.ToString("yyyy-MM-dd HH:mm:ss"))
            {
                Tag = file.FullName
            };
            item.SubItems.Add(type);
            item.SubItems.Add(file.Name);
            item.SubItems.Add($"{Math.Max(1, file.Length / 1024)} KB");
            item.SubItems.Add(file.DirectoryName ?? string.Empty);
            _list.Items.Add(item);
        }

        _list.EndUpdate();
        SyncSelectionState();
    }

    private void SyncSelectionState()
    {
        _useRollbackButton.Enabled = IsSelectedRollback();
    }

    private bool IsSelectedRollback()
    {
        if (_list.SelectedItems.Count == 0) return false;
        var type = _list.SelectedItems[0].SubItems[1].Text;
        return string.Equals(type, "回滚", StringComparison.OrdinalIgnoreCase);
    }

    private void UseSelectedRollback()
    {
        if (!IsSelectedRollback()) return;
        SelectedRollbackPath = _list.SelectedItems[0].Tag?.ToString();
        if (string.IsNullOrWhiteSpace(SelectedRollbackPath)) return;
        this.Close();
    }

    private void OpenRootFolder()
    {
        EnsureRootDirectory();
        Process.Start(new ProcessStartInfo
        {
            FileName = "explorer.exe",
            Arguments = $"\"{_rootDirectory}\"",
            UseShellExecute = true
        });
    }

    private void EnsureRootDirectory()
    {
        if (!Directory.Exists(_rootDirectory))
        {
            Directory.CreateDirectory(_rootDirectory);
        }
    }

    private static string ResolveType(string fileName)
    {
        if (fileName.Contains("-rollback-", StringComparison.OrdinalIgnoreCase)) return "回滚";
        if (fileName.Contains("-before-", StringComparison.OrdinalIgnoreCase)) return "快照";
        return "其他";
    }
}
