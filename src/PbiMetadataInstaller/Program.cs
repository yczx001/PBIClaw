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

            if (options.IsUninstall)
            {
                var installDir = options.InstallDir ?? InstallerEngine.GetExistingInstallDir();
                if (string.IsNullOrWhiteSpace(installDir) || !Directory.Exists(installDir))
                {
                    MessageBox.Show(
                        "未找到 PBI Claw 的安装目录，请手动指定安装路径。",
                        "PBI Claw 卸载程序",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning);
                    return 1;
                }

                if (!InstallerEngine.IsAdministrator())
                {
                    // 重新以管理员权限启动卸载
                    InstallerEngine.RelaunchAsAdminForUninstall(installDir);
                    return 0;
                }

                var confirm = MessageBox.Show(
                    $"确定要卸载 PBI Claw 吗？\n\n安装目录：{installDir}",
                    "确认卸载",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Question);

                if (confirm != DialogResult.Yes)
                {
                    return 0;
                }

                InstallerEngine.Uninstall(installDir);

                MessageBox.Show(
                    "PBI Claw 已成功卸载！",
                    "卸载完成",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
                return 0;
            }

            ApplicationConfiguration.Initialize();
            using var wizard = new SetupWizardForm(options.InstallDir, options.AutoInstall);
            Application.Run(wizard);
            return wizard.ExitCode;
        }
        catch (Exception ex)
        {
            var title = args.Contains("--uninstall", StringComparer.OrdinalIgnoreCase) ? "卸载失败" : "安装失败";
            MessageBox.Show(
                $"{title}:\n{ex.Message}",
                $"PBI Claw {title}",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error);
            return 1;
        }
    }

    private static InstallOptions ParseArgs(string[] args)
    {
        string? installDir = null;
        var showHelp = false;
        var autoInstall = false;
        var isUninstall = false;

        for (var i = 0; i < args.Length; i++)
        {
            var arg = args[i];
            if (string.Equals(arg, "--machine", StringComparison.OrdinalIgnoreCase))
            {
                autoInstall = true;
                continue;
            }

            if (string.Equals(arg, "--uninstall", StringComparison.OrdinalIgnoreCase))
            {
                isUninstall = true;
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

        return new InstallOptions(installDir, showHelp, autoInstall, isUninstall);
    }

    private static void ShowUsage()
    {
        MessageBox.Show(
            "Usage:\n" +
            "安装：PBIClawSetup.exe [--install-dir \"C:\\\\Path\\\\To\\\\Folder\"]\n" +
            "卸载：PBIClawSetup.exe --uninstall [--install-dir \"C:\\\\Path\\\\To\\\\Folder\"]\n\n" +
            "安装时会自动把 PBIClaw.pbitool.json 写入两个系统 External Tools 目录。\n" +
            "卸载后会自动清理所有相关文件和配置。",
            "PBI Claw 安装程序",
            MessageBoxButtons.OK,
            MessageBoxIcon.Information);
    }

    private sealed record InstallOptions(string? InstallDir, bool ShowHelp, bool AutoInstall, bool IsUninstall);
}
