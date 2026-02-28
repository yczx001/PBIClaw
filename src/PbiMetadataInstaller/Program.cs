using System.Reflection;
using System.Security.Principal;
using System.Text.Json;
using System.Windows.Forms;
using System.Diagnostics;

namespace PBIClawSetup;

internal static class Program
{
    private const string ToolExeName = "PBIClaw.exe";
    private const string ToolJsonName = "PBIClaw.pbitool.json";
    private const string ToolDisplayName = "PBI Claw";
    private const string ToolArguments = "--server \"%server%\" --database \"%database%\" --external-tool";

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

            var installDir = ResolveInstallDirectory(options.InstallDir);
            if (string.IsNullOrWhiteSpace(installDir))
            {
                return 0;
            }

            if (!IsAdministrator())
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

                RelaunchAsAdmin($"--machine --install-dir \"{installDir}\"");
                return 0;
            }

            InstallMachine(installDir);
            return 0;
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

    private static int InstallMachine(string installDir)
    {
        if (!IsAdministrator())
        {
            MessageBox.Show(
                "Machine install requires Administrator privileges.\nRun setup as administrator.",
                "PBI Claw 安装程序",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error);
            return 1;
        }

        Directory.CreateDirectory(installDir);
        var exePath = Path.Combine(installDir, ToolExeName);
        ExtractEmbeddedResource($"Payload.{ToolExeName}", exePath);

        var toolModel = LoadPbiToolModel(exePath);

        var externalToolDirs = GetMachineExternalToolDirs();
        foreach (var externalToolDir in externalToolDirs)
        {
            Directory.CreateDirectory(externalToolDir);
            var jsonPath = Path.Combine(externalToolDir, ToolJsonName);
            WritePbiToolJson(jsonPath, toolModel);
        }

        ShowSuccess(installDir, externalToolDirs);
        return 0;
    }

    private static string? ResolveInstallDirectory(string? fromArgs)
    {
        if (!string.IsNullOrWhiteSpace(fromArgs))
        {
            return Path.GetFullPath(Environment.ExpandEnvironmentVariables(fromArgs));
        }

        var defaultDir = GetDefaultInstallDir();
        var startPath = defaultDir;
        if (!Directory.Exists(startPath))
        {
            var programFiles = Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles);
            if (!string.IsNullOrWhiteSpace(programFiles))
            {
                startPath = programFiles;
            }
        }

        using var dialog = new FolderBrowserDialog
        {
            Description = "选择 PBI Claw 可执行文件安装目录",
            UseDescriptionForTitle = true,
            ShowNewFolderButton = true,
            SelectedPath = defaultDir,
            InitialDirectory = startPath
        };

        if (dialog.ShowDialog() != DialogResult.OK)
        {
            return null;
        }

        return Path.GetFullPath(dialog.SelectedPath);
    }

    private static void ShowSuccess(string installDir, IReadOnlyList<string> externalToolDirs)
    {
        var directories = string.Join(Environment.NewLine, externalToolDirs);
        MessageBox.Show(
            "安装完成。\n\n" +
            $"EXE 路径:\n{installDir}\n\n" +
            $"外部工具配置已写入:\n{directories}\n\n" +
            "请重启 Power BI Desktop，然后在 External Tools 中点击 PBI Claw。",
            "PBI Claw 安装程序",
            MessageBoxButtons.OK,
            MessageBoxIcon.Information);
    }

    private static string GetDefaultInstallDir()
    {
        var programFiles = Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles);
        if (string.IsNullOrWhiteSpace(programFiles))
        {
            return Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), ToolDisplayName);
        }

        return Path.Combine(programFiles, ToolDisplayName);
    }

    private static IReadOnlyList<string> GetMachineExternalToolDirs()
    {
        var dirs = new List<string>();

        var commonProgramFiles = Environment.GetFolderPath(Environment.SpecialFolder.CommonProgramFiles);
        if (!string.IsNullOrWhiteSpace(commonProgramFiles))
        {
            dirs.Add(Path.Combine(commonProgramFiles, "Microsoft Shared", "Power BI Desktop", "External Tools"));
        }

        var commonProgramFilesX86 = Environment.GetEnvironmentVariable("CommonProgramFiles(x86)");
        if (!string.IsNullOrWhiteSpace(commonProgramFilesX86))
        {
            dirs.Add(Path.Combine(commonProgramFilesX86, "Microsoft Shared", "Power BI Desktop", "External Tools"));
        }

        return dirs
            .Where(d => !string.IsNullOrWhiteSpace(d))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();
    }

    private static Dictionary<string, object?> LoadPbiToolModel(string exePath)
    {
        using var stream = FindResourceStream($"Payload.{ToolJsonName}");
        using var reader = new StreamReader(stream);
        var json = reader.ReadToEnd();
        using var doc = JsonDocument.Parse(json);

        var root = doc.RootElement;
        var model = new Dictionary<string, object?>(StringComparer.OrdinalIgnoreCase);
        foreach (var prop in root.EnumerateObject())
        {
            model[prop.Name] = ReadValue(prop.Value);
        }

        model["name"] = ToolDisplayName;
        model["path"] = exePath;
        model["arguments"] = ToolArguments;
        return model;
    }

    private static void WritePbiToolJson(string jsonPath, Dictionary<string, object?> model)
    {
        var normalized = JsonSerializer.Serialize(model, new JsonSerializerOptions
        {
            WriteIndented = true
        });
        File.WriteAllText(jsonPath, normalized);
    }

    private static bool IsAdministrator()
    {
        using var identity = WindowsIdentity.GetCurrent();
        var principal = new WindowsPrincipal(identity);
        return principal.IsInRole(WindowsBuiltInRole.Administrator);
    }

    private static void ExtractEmbeddedResource(string logicalName, string outputPath)
    {
        using var stream = FindResourceStream(logicalName);
        using var file = File.Create(outputPath);
        stream.CopyTo(file);
    }

    private static Stream FindResourceStream(string logicalName)
    {
        var assembly = Assembly.GetExecutingAssembly();
        var exact = assembly.GetManifestResourceStream(logicalName);
        if (exact is not null)
        {
            return exact;
        }

        // Resource logical names may include namespace prefixes in SDK-style projects.
        var fallbackName = assembly
            .GetManifestResourceNames()
            .FirstOrDefault(name => name.EndsWith(logicalName, StringComparison.OrdinalIgnoreCase));

        if (fallbackName is null)
        {
            throw new InvalidOperationException($"Missing embedded resource: {logicalName}");
        }

        var fallback = assembly.GetManifestResourceStream(fallbackName);
        if (fallback is null)
        {
            throw new InvalidOperationException($"Cannot open embedded resource: {fallbackName}");
        }

        return fallback;
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

    private static object? ReadValue(JsonElement element)
    {
        return element.ValueKind switch
        {
            JsonValueKind.String => element.GetString(),
            JsonValueKind.Number => element.TryGetInt64(out var i) ? i : element.GetDouble(),
            JsonValueKind.True => true,
            JsonValueKind.False => false,
            JsonValueKind.Null => null,
            _ => element.GetRawText()
        };
    }

    private static void RelaunchAsAdmin(string args)
    {
        var exe = Environment.ProcessPath;
        if (string.IsNullOrWhiteSpace(exe))
        {
            throw new InvalidOperationException("Cannot determine current executable path.");
        }

        var info = new ProcessStartInfo
        {
            FileName = exe,
            Arguments = args,
            UseShellExecute = true,
            Verb = "runas"
        };

        Process.Start(info);
    }
}
