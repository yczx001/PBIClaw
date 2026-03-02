using System.Diagnostics;
using System.Reflection;
using System.Security.Principal;
using System.Text.Json;

namespace PBIClawSetup;

internal sealed record InstallResult(string InstallDir, IReadOnlyList<string> ExternalToolDirs);

internal static class InstallerEngine
{
    private const string ToolExeName = "PBIClaw.exe";
    private const string ToolJsonName = "PBIClaw.pbitool.json";
    private const string ToolDisplayName = "PBI Claw";
    private const string ToolArguments = "--server \"%server%\" --database \"%database%\" --external-tool";

    public static string ToolName => ToolDisplayName;

    public static string GetDefaultInstallDir()
    {
        var programFiles = Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles);
        if (string.IsNullOrWhiteSpace(programFiles))
        {
            return Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), ToolDisplayName);
        }

        return Path.Combine(programFiles, ToolDisplayName);
    }

    public static IReadOnlyList<string> GetMachineExternalToolDirs()
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

    public static bool IsAdministrator()
    {
        using var identity = WindowsIdentity.GetCurrent();
        var principal = new WindowsPrincipal(identity);
        return principal.IsInRole(WindowsBuiltInRole.Administrator);
    }

    public static InstallResult InstallMachine(string installDir)
    {
        if (!IsAdministrator())
        {
            throw new InvalidOperationException("Machine install requires Administrator privileges.");
        }

        var normalizedInstallDir = Path.GetFullPath(Environment.ExpandEnvironmentVariables(installDir));
        Directory.CreateDirectory(normalizedInstallDir);

        var exePath = Path.Combine(normalizedInstallDir, ToolExeName);
        if (IsExecutableRunning(exePath))
        {
            throw new InvalidOperationException("检测到正在运行的 PBIClaw.exe，请先关闭后再重试安装。");
        }

        ExtractEmbeddedResource($"Payload.{ToolExeName}", exePath);

        var toolModel = LoadPbiToolModel(exePath);
        var externalToolDirs = GetMachineExternalToolDirs();
        foreach (var externalToolDir in externalToolDirs)
        {
            Directory.CreateDirectory(externalToolDir);
            var jsonPath = Path.Combine(externalToolDir, ToolJsonName);
            WritePbiToolJson(jsonPath, toolModel);
        }

        return new InstallResult(normalizedInstallDir, externalToolDirs);
    }

    public static void RelaunchAsAdmin(string installDir)
    {
        var exe = Environment.ProcessPath;
        if (string.IsNullOrWhiteSpace(exe))
        {
            throw new InvalidOperationException("Cannot determine current executable path.");
        }

        var args = $"--machine --install-dir \"{installDir}\"";
        var info = new ProcessStartInfo
        {
            FileName = exe,
            Arguments = args,
            UseShellExecute = true,
            Verb = "runas"
        };

        Process.Start(info);
    }

    private static Dictionary<string, object?> LoadPbiToolModel(string exePath)
    {
        using var stream = FindResourceStream($"Payload.{ToolJsonName}");
        using var reader = new StreamReader(stream);
        var json = reader.ReadToEnd();
        using var doc = JsonDocument.Parse(json);

        var model = new Dictionary<string, object?>(StringComparer.OrdinalIgnoreCase);
        foreach (var prop in doc.RootElement.EnumerateObject())
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

    private static void ExtractEmbeddedResource(string logicalName, string outputPath)
    {
        using var stream = FindResourceStream(logicalName);
        using var file = File.Create(outputPath);
        stream.CopyTo(file);
    }

    private static bool IsExecutableRunning(string targetExePath)
    {
        var normalizedTarget = Path.GetFullPath(targetExePath);
        var currentPid = Environment.ProcessId;

        foreach (var process in Process.GetProcessesByName(Path.GetFileNameWithoutExtension(ToolExeName)))
        {
            try
            {
                if (process.Id == currentPid)
                {
                    continue;
                }

                var processPath = process.MainModule?.FileName;
                if (!string.IsNullOrWhiteSpace(processPath) &&
                    string.Equals(Path.GetFullPath(processPath), normalizedTarget, StringComparison.OrdinalIgnoreCase))
                {
                    return true;
                }
            }
            catch
            {
                // Ignore inaccessible processes.
            }
            finally
            {
                process.Dispose();
            }
        }

        return false;
    }

    private static Stream FindResourceStream(string logicalName)
    {
        var assembly = Assembly.GetExecutingAssembly();
        var exact = assembly.GetManifestResourceStream(logicalName);
        if (exact is not null)
        {
            return exact;
        }

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
