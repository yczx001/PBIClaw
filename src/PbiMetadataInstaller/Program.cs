using System.Windows.Forms;

namespace PBIClawSetup;

internal static class Program
{
    [STAThread]
    private static int Main(string[] args)
    {
        try
        {
            var options = ParseArgs(args);
            if (options.ShowHelp)
            {
                ShowUsage();
                return 0;
            }

            if (!string.IsNullOrWhiteSpace(options.InstallDir))
            {
                return InstallFromCommandLine(options.InstallDir);
            }

            ApplicationConfiguration.Initialize();
            using var wizard = new SetupWizardForm();
            Application.Run(wizard);
            return wizard.ExitCode;
        }
        catch (Exception ex)
        {
            MessageBox.Show(
                $"Install failed:\n{ex.Message}",
                "PBI Claw 安装程序",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error);
            return 1;
        }
    }

    private static int InstallFromCommandLine(string installDirArg)
    {
        var installDir = Path.GetFullPath(Environment.ExpandEnvironmentVariables(installDirArg));

        if (!InstallerEngine.IsAdministrator())
        {
            var relaunch = MessageBox.Show(
                "安装会把外部工具配置写入系统目录，需要管理员权限。\n\n点击“是”继续提权安装，点击“否”取消。",
                "PBI Claw 安装程序",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Question);

            if (relaunch != DialogResult.Yes)
            {
                return 1;
            }

            InstallerEngine.RelaunchAsAdmin(installDir);
            return 0;
        }

        var result = InstallerEngine.InstallMachine(installDir);
        ShowSuccess(result);
        return 0;
    }

    private static void ShowSuccess(InstallResult result)
    {
        var directories = string.Join(Environment.NewLine, result.ExternalToolDirs);
        MessageBox.Show(
            "安装完成。\n\n" +
            $"EXE 路径:\n{result.InstallDir}\n\n" +
            $"外部工具配置已写入:\n{directories}\n\n" +
            "请重启 Power BI Desktop，然后在 External Tools 中点击 PBI Claw。",
            "PBI Claw 安装程序",
            MessageBoxButtons.OK,
            MessageBoxIcon.Information);
    }

    private static InstallOptions ParseArgs(string[] args)
    {
        string? installDir = null;
        var showHelp = false;

        for (var i = 0; i < args.Length; i++)
        {
            var arg = args[i];
            if (string.Equals(arg, "--machine", StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            if (string.Equals(arg, "--help", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(arg, "-h", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(arg, "/?", StringComparison.OrdinalIgnoreCase))
            {
                showHelp = true;
                continue;
            }

            if (arg.StartsWith("--install-dir=", StringComparison.OrdinalIgnoreCase))
            {
                installDir = arg.Substring("--install-dir=".Length).Trim().Trim('"');
                continue;
            }

            if (string.Equals(arg, "--install-dir", StringComparison.OrdinalIgnoreCase))
            {
                if (i + 1 >= args.Length)
                {
                    throw new ArgumentException("Missing value for --install-dir");
                }

                installDir = args[++i].Trim().Trim('"');
                continue;
            }
        }

        return new InstallOptions(installDir, showHelp);
    }

    private static void ShowUsage()
    {
        MessageBox.Show(
            "Usage:\n" +
            "PBIClawSetup.exe [--install-dir \"C:\\\\Path\\\\To\\\\Folder\"]\n\n" +
            "安装时会自动把 PBIClaw.pbitool.json 写入两个系统 External Tools 目录。",
            "PBI Claw 安装程序",
            MessageBoxButtons.OK,
            MessageBoxIcon.Information);
    }

    private sealed record InstallOptions(string? InstallDir, bool ShowHelp);
}
