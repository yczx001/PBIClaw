using System.Text;
using System.Windows.Forms;

namespace PBIClawSetup;

internal sealed class SetupWizardForm : Form
{
    private enum SetupStep
    {
        Welcome,
        Directory,
        Confirm,
        Installing,
        Complete
    }

    private readonly IReadOnlyList<string> _externalToolDirs;

    private readonly Panel _welcomePage;
    private readonly Panel _directoryPage;
    private readonly Panel _confirmPage;
    private readonly Panel _installingPage;
    private readonly Panel _completePage;

    private readonly TextBox _installDirText;
    private readonly TextBox _confirmSummaryText;
    private readonly TextBox _completeSummaryText;
    private readonly Label _installStatusLabel;
    private readonly ProgressBar _installProgressBar;

    private readonly Button _backButton;
    private readonly Button _nextButton;
    private readonly Button _cancelButton;

    private SetupStep _step = SetupStep.Welcome;
    private InstallResult? _result;

    public int ExitCode { get; private set; } = 1;

    public SetupWizardForm()
    {
        Text = "PBI Claw 安装向导";
        StartPosition = FormStartPosition.CenterScreen;
        FormBorderStyle = FormBorderStyle.FixedDialog;
        MaximizeBox = false;
        MinimizeBox = false;
        ClientSize = new Size(760, 470);

        _externalToolDirs = InstallerEngine.GetMachineExternalToolDirs();

        var root = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            ColumnCount = 1,
            RowCount = 3
        };
        root.RowStyles.Add(new RowStyle(SizeType.Absolute, 88));
        root.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
        root.RowStyles.Add(new RowStyle(SizeType.Absolute, 62));
        Controls.Add(root);

        var header = BuildHeader();
        root.Controls.Add(header, 0, 0);

        var pageHost = new Panel
        {
            Dock = DockStyle.Fill,
            Padding = new Padding(24, 18, 24, 10)
        };
        root.Controls.Add(pageHost, 0, 1);

        _welcomePage = BuildWelcomePage();
        _directoryPage = BuildDirectoryPage(out _installDirText);
        _confirmPage = BuildConfirmPage(out _confirmSummaryText);
        _installingPage = BuildInstallingPage(out _installStatusLabel, out _installProgressBar);
        _completePage = BuildCompletePage(out _completeSummaryText);

        pageHost.Controls.Add(_welcomePage);
        pageHost.Controls.Add(_directoryPage);
        pageHost.Controls.Add(_confirmPage);
        pageHost.Controls.Add(_installingPage);
        pageHost.Controls.Add(_completePage);

        var footer = new Panel
        {
            Dock = DockStyle.Fill,
            Padding = new Padding(16, 8, 16, 10),
            BackColor = SystemColors.Control
        };
        root.Controls.Add(footer, 0, 2);

        var buttons = new FlowLayoutPanel
        {
            Dock = DockStyle.Right,
            AutoSize = true,
            AutoSizeMode = AutoSizeMode.GrowAndShrink,
            FlowDirection = FlowDirection.RightToLeft,
            WrapContents = false
        };
        footer.Controls.Add(buttons);

        _nextButton = new Button
        {
            Width = 92,
            Height = 30
        };
        _nextButton.Click += OnNextClicked;
        buttons.Controls.Add(_nextButton);

        _backButton = new Button
        {
            Text = "< 上一步",
            Width = 92,
            Height = 30
        };
        _backButton.Click += OnBackClicked;
        buttons.Controls.Add(_backButton);

        _cancelButton = new Button
        {
            Text = "取消",
            Width = 92,
            Height = 30
        };
        _cancelButton.Click += OnCancelClicked;
        buttons.Controls.Add(_cancelButton);

        _installDirText.Text = InstallerEngine.GetDefaultInstallDir();

        SetStep(SetupStep.Welcome);
    }

    protected override void OnFormClosing(FormClosingEventArgs e)
    {
        if (_step == SetupStep.Installing)
        {
            e.Cancel = true;
            return;
        }

        base.OnFormClosing(e);
    }

    private static Panel BuildHeader()
    {
        var panel = new Panel
        {
            Dock = DockStyle.Fill,
            BackColor = Color.FromArgb(28, 82, 138)
        };

        var title = new Label
        {
            AutoSize = true,
            ForeColor = Color.White,
            Font = new Font("Segoe UI", 16, FontStyle.Bold),
            Text = "PBI Claw 安装向导",
            Location = new Point(20, 16)
        };
        panel.Controls.Add(title);

        var subtitle = new Label
        {
            AutoSize = true,
            ForeColor = Color.FromArgb(235, 245, 255),
            Font = new Font("Segoe UI", 9.5f),
            Text = "按“下一步”完成安装，自动注册 Power BI External Tools。",
            Location = new Point(22, 50)
        };
        panel.Controls.Add(subtitle);

        return panel;
    }

    private static Panel BuildWelcomePage()
    {
        var panel = CreatePagePanel();

        var title = new Label
        {
            AutoSize = true,
            Font = new Font("Segoe UI", 12, FontStyle.Bold),
            Text = "欢迎使用 PBI Claw 安装程序",
            Location = new Point(0, 0)
        };
        panel.Controls.Add(title);

        var detail = new Label
        {
            AutoSize = false,
            Width = 670,
            Height = 180,
            Font = new Font("Segoe UI", 10),
            Text = "此向导会执行以下操作：\r\n\r\n" +
                   "1. 将 PBIClaw.exe 安装到你选择的目录\r\n" +
                   "2. 在两个系统 External Tools 目录写入 PBIClaw.pbitool.json\r\n" +
                   "3. 菜单名称固定为“PBI Claw”\r\n\r\n" +
                   "安装时会请求管理员权限。",
            Location = new Point(0, 42)
        };
        panel.Controls.Add(detail);

        return panel;
    }

    private static Panel BuildDirectoryPage(out TextBox installDirText)
    {
        var panel = CreatePagePanel();

        var title = new Label
        {
            AutoSize = true,
            Font = new Font("Segoe UI", 12, FontStyle.Bold),
            Text = "选择安装目录",
            Location = new Point(0, 0)
        };
        panel.Controls.Add(title);

        var tip = new Label
        {
            AutoSize = true,
            Font = new Font("Segoe UI", 9.5f),
            Text = "请选择 PBIClaw.exe 的安装位置：",
            Location = new Point(0, 44)
        };
        panel.Controls.Add(tip);

        var dirText = new TextBox
        {
            Width = 540,
            Location = new Point(0, 76)
        };
        panel.Controls.Add(dirText);

        var browseButton = new Button
        {
            Text = "浏览...",
            Width = 95,
            Height = 27,
            Location = new Point(552, 74)
        };
        browseButton.Click += (_, _) =>
        {
            using var dialog = new FolderBrowserDialog
            {
                Description = "选择 PBI Claw 安装目录",
                UseDescriptionForTitle = true,
                ShowNewFolderButton = true,
                SelectedPath = dirText.Text
            };

            if (dialog.ShowDialog() == DialogResult.OK)
            {
                dirText.Text = dialog.SelectedPath;
            }
        };
        panel.Controls.Add(browseButton);
        installDirText = dirText;

        return panel;
    }

    private static Panel BuildConfirmPage(out TextBox confirmSummaryText)
    {
        var panel = CreatePagePanel();

        var title = new Label
        {
            AutoSize = true,
            Font = new Font("Segoe UI", 12, FontStyle.Bold),
            Text = "确认安装信息",
            Location = new Point(0, 0)
        };
        panel.Controls.Add(title);

        var tip = new Label
        {
            AutoSize = true,
            Font = new Font("Segoe UI", 9.5f),
            Text = "点击“安装”开始执行：",
            Location = new Point(0, 44)
        };
        panel.Controls.Add(tip);

        confirmSummaryText = new TextBox
        {
            Multiline = true,
            ReadOnly = true,
            ScrollBars = ScrollBars.Vertical,
            Width = 650,
            Height = 220,
            Location = new Point(0, 72),
            BackColor = Color.White
        };
        panel.Controls.Add(confirmSummaryText);

        return panel;
    }

    private static Panel BuildInstallingPage(out Label statusLabel, out ProgressBar progressBar)
    {
        var panel = CreatePagePanel();

        var title = new Label
        {
            AutoSize = true,
            Font = new Font("Segoe UI", 12, FontStyle.Bold),
            Text = "正在安装",
            Location = new Point(0, 0)
        };
        panel.Controls.Add(title);

        statusLabel = new Label
        {
            AutoSize = true,
            Font = new Font("Segoe UI", 9.5f),
            Text = "正在复制文件并写入外部工具配置...",
            Location = new Point(0, 50)
        };
        panel.Controls.Add(statusLabel);

        progressBar = new ProgressBar
        {
            Width = 650,
            Height = 20,
            Style = ProgressBarStyle.Marquee,
            MarqueeAnimationSpeed = 25,
            Location = new Point(0, 84)
        };
        panel.Controls.Add(progressBar);

        return panel;
    }

    private static Panel BuildCompletePage(out TextBox completeSummaryText)
    {
        var panel = CreatePagePanel();

        var title = new Label
        {
            AutoSize = true,
            Font = new Font("Segoe UI", 12, FontStyle.Bold),
            Text = "安装完成",
            Location = new Point(0, 0)
        };
        panel.Controls.Add(title);

        var tip = new Label
        {
            AutoSize = true,
            Font = new Font("Segoe UI", 9.5f),
            Text = "请重启 Power BI Desktop，然后在 External Tools 中点击 PBI Claw。",
            Location = new Point(0, 44)
        };
        panel.Controls.Add(tip);

        completeSummaryText = new TextBox
        {
            Multiline = true,
            ReadOnly = true,
            ScrollBars = ScrollBars.Vertical,
            Width = 650,
            Height = 220,
            Location = new Point(0, 72),
            BackColor = Color.White
        };
        panel.Controls.Add(completeSummaryText);

        return panel;
    }

    private static Panel CreatePagePanel()
    {
        return new Panel
        {
            Dock = DockStyle.Fill,
            Visible = false
        };
    }

    private async void OnNextClicked(object? sender, EventArgs e)
    {
        switch (_step)
        {
            case SetupStep.Welcome:
                SetStep(SetupStep.Directory);
                return;
            case SetupStep.Directory:
                if (!TryNormalizeInstallDir(out _))
                {
                    return;
                }
                SetStep(SetupStep.Confirm);
                return;
            case SetupStep.Confirm:
                await InstallAsync();
                return;
            case SetupStep.Complete:
                ExitCode = 0;
                Close();
                return;
        }
    }

    private void OnBackClicked(object? sender, EventArgs e)
    {
        switch (_step)
        {
            case SetupStep.Directory:
                SetStep(SetupStep.Welcome);
                return;
            case SetupStep.Confirm:
                SetStep(SetupStep.Directory);
                return;
        }
    }

    private void OnCancelClicked(object? sender, EventArgs e)
    {
        ExitCode = 1;
        Close();
    }

    private void SetStep(SetupStep step)
    {
        _step = step;

        _welcomePage.Visible = step == SetupStep.Welcome;
        _directoryPage.Visible = step == SetupStep.Directory;
        _confirmPage.Visible = step == SetupStep.Confirm;
        _installingPage.Visible = step == SetupStep.Installing;
        _completePage.Visible = step == SetupStep.Complete;

        _backButton.Visible = step is SetupStep.Directory or SetupStep.Confirm;
        _cancelButton.Visible = step is not SetupStep.Installing and not SetupStep.Complete;

        _nextButton.Enabled = step != SetupStep.Installing;
        _nextButton.Text = step switch
        {
            SetupStep.Confirm => "安装",
            SetupStep.Complete => "完成",
            SetupStep.Installing => "安装中...",
            _ => "下一步 >"
        };

        if (step == SetupStep.Confirm)
        {
            UpdateConfirmSummary();
        }

        if (step == SetupStep.Complete)
        {
            UpdateCompleteSummary();
        }
    }

    private void UpdateConfirmSummary()
    {
        _ = TryNormalizeInstallDir(out var normalizedInstallDir);
        var installDir = normalizedInstallDir ?? _installDirText.Text.Trim();

        var summary = new StringBuilder();
        summary.AppendLine("程序文件安装路径：");
        summary.AppendLine(installDir);
        summary.AppendLine();
        summary.AppendLine("将写入 External Tools 的目录：");
        foreach (var dir in _externalToolDirs)
        {
            summary.AppendLine(dir);
        }
        summary.AppendLine();
        summary.Append("外部工具名称：");
        summary.Append(InstallerEngine.ToolName);

        _confirmSummaryText.Text = summary.ToString();
    }

    private void UpdateCompleteSummary()
    {
        if (_result is null)
        {
            _completeSummaryText.Text = string.Empty;
            return;
        }

        var summary = new StringBuilder();
        summary.AppendLine("安装成功。");
        summary.AppendLine();
        summary.AppendLine("EXE 路径：");
        summary.AppendLine(_result.InstallDir);
        summary.AppendLine();
        summary.AppendLine("外部工具配置写入：");
        foreach (var dir in _result.ExternalToolDirs)
        {
            summary.AppendLine(dir);
        }

        _completeSummaryText.Text = summary.ToString();
    }

    private async Task InstallAsync()
    {
        if (!TryNormalizeInstallDir(out var installDir))
        {
            return;
        }
        var finalInstallDir = installDir!;

        if (!InstallerEngine.IsAdministrator())
        {
            var relaunch = MessageBox.Show(
                "安装需要管理员权限。点击“是”将重新以管理员方式启动安装器。",
                "PBI Claw 安装程序",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Question);

            if (relaunch != DialogResult.Yes)
            {
                return;
            }

            InstallerEngine.RelaunchAsAdmin(finalInstallDir);
            ExitCode = 0;
            Close();
            return;
        }

        SetStep(SetupStep.Installing);
        _installStatusLabel.Text = "正在复制文件并写入外部工具配置...";

        try
        {
            _result = await Task.Run(() => InstallerEngine.InstallMachine(finalInstallDir));
            _installStatusLabel.Text = "安装完成。";
            SetStep(SetupStep.Complete);
        }
        catch (Exception ex)
        {
            MessageBox.Show(
                $"安装失败：\n{ex.Message}",
                "PBI Claw 安装程序",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error);
            SetStep(SetupStep.Confirm);
        }
    }

    private bool TryNormalizeInstallDir(out string? installDir)
    {
        installDir = null;
        var raw = _installDirText.Text.Trim();
        if (string.IsNullOrWhiteSpace(raw))
        {
            MessageBox.Show("请先选择安装目录。", "PBI Claw 安装程序", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return false;
        }

        try
        {
            installDir = Path.GetFullPath(Environment.ExpandEnvironmentVariables(raw));
            _installDirText.Text = installDir;
            return true;
        }
        catch
        {
            MessageBox.Show("安装目录格式无效，请重新选择。", "PBI Claw 安装程序", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return false;
        }
    }
}
