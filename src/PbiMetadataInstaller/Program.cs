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

            ApplicationConfiguration.Initialize();
            using var wizard = new SetupWizardForm(options.InstallDir, options.AutoInstall);
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

    private static InstallOptions ParseArgs(string[] args)
    {
        string? installDir = null;
        var showHelp = false;
        var autoInstall = false;

        for (var i = 0; i < args.Length; i++)
        {
            var arg = args[i];
            if (string.Equals(arg, "--machine", StringComparison.OrdinalIgnoreCase))
            {
                autoInstall = true;
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

        return new InstallOptions(installDir, showHelp, autoInstall);
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

    private sealed record InstallOptions(string? InstallDir, bool ShowHelp, bool AutoInstall);
}
