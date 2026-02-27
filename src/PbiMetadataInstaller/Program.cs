using System.Reflection;
using System.Security.Principal;
using System.Text.Json;
using System.Windows.Forms;
using System.Diagnostics;

namespace PBIClawSetup;

internal static class Program
{
    [STAThread]
    private static int Main(string[] args)
    {
        try
        {
            var forceMachine = args.Any(a => string.Equals(a, "--machine", StringComparison.OrdinalIgnoreCase));
            var forceUser = args.Any(a => string.Equals(a, "--user", StringComparison.OrdinalIgnoreCase));

            if (forceMachine && forceUser)
            {
                MessageBox.Show(
                    "Cannot use --machine and --user together.",
                    "PBI Claw 安装程序",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
                return 1;
            }

            if (forceMachine)
            {
                return InstallMachine();
            }

            if (forceUser)
            {
                InstallToDirectory(GetUserTargetDir());
                ShowSuccess(GetUserTargetDir(), "User");
                return 0;
            }

            // Default behavior: prefer machine-wide install so Power BI can discover the tool reliably.
            if (IsAdministrator())
            {
                return InstallMachine();
            }

            var relaunch = MessageBox.Show(
                "Administrator permission is recommended for machine-wide install.\n\n" +
                "Click Yes to relaunch setup as Administrator.\n" +
                "Click No to install for current user only.",
                "PBI Claw 安装程序",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Question);

            if (relaunch == DialogResult.Yes)
            {
                RelaunchAsAdmin("--machine");
                return 0;
            }

            var userDir = GetUserTargetDir();
            InstallToDirectory(userDir);
            ShowSuccess(userDir, "User");
            MessageBox.Show(
                "If the tool is not visible in Power BI, rerun setup as Administrator and choose machine-wide install.",
                "PBI Claw 安装程序",
                MessageBoxButtons.OK,
                MessageBoxIcon.Warning);

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

    private static int InstallMachine()
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

        var targets = GetMachineTargetDirs();
        foreach (var dir in targets)
        {
            InstallToDirectory(dir);
        }

        ShowSuccess(string.Join(Environment.NewLine, targets), "Machine");
        return 0;
    }

    private static void InstallToDirectory(string targetDir)
    {
        Directory.CreateDirectory(targetDir);

        var exePath = Path.Combine(targetDir, "PBIClaw.exe");
        var jsonPath = Path.Combine(targetDir, "PBIClaw.pbitool.json");

        ExtractEmbeddedResource("Payload.PBIClaw.exe", exePath);
        ExtractEmbeddedResource("Payload.PBIClaw.pbitool.json", jsonPath);
        NormalizePbiToolJson(jsonPath, exePath);
    }

    private static void ShowSuccess(string target, string scope)
    {
        MessageBox.Show(
            $"Install completed.\n\nScope: {scope}\nTarget:\n{target}\n\nRestart Power BI Desktop, then open External Tools and click PBI Claw.",
            "PBI Claw 安装程序",
            MessageBoxButtons.OK,
            MessageBoxIcon.Information);
    }

    private static string GetUserTargetDir()
    {
        return Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
            "Power BI Desktop",
            "External Tools");
    }

    private static IReadOnlyList<string> GetMachineTargetDirs()
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

    private static void NormalizePbiToolJson(string jsonPath, string exePath)
    {
        var json = File.ReadAllText(jsonPath);
        using var doc = JsonDocument.Parse(json);

        var root = doc.RootElement;
        var model = new Dictionary<string, object?>(StringComparer.OrdinalIgnoreCase);
        foreach (var prop in root.EnumerateObject())
        {
            model[prop.Name] = ReadValue(prop.Value);
        }

        model["path"] = exePath;
        if (!model.ContainsKey("name"))
        {
            model["name"] = "PBI Claw";
        }

        var normalized = JsonSerializer.Serialize(model, new JsonSerializerOptions
        {
            WriteIndented = true
        });

        File.WriteAllText(jsonPath, normalized);
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
}
