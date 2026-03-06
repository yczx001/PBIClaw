using System.Diagnostics;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Security.Principal;
using System.Text.Json;

namespace PBIClawSetup;

internal sealed record InstallResult(string InstallDir, IReadOnlyList<string> ExternalToolDirs);
internal sealed record ShortcutResult(string? DesktopShortcutPath, string? StartMenuShortcutPath);

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
        if (IsAnyToolProcessRunning())
        {
            throw new InvalidOperationException("检测到 PBIClaw 仍在运行，请先关闭所有 PBIClaw 进程后再重试安装。");
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

    public static ShortcutResult ApplyShortcuts(string installDir, bool createDesktopShortcut, bool createStartMenuShortcut)
    {
        var normalizedInstallDir = Path.GetFullPath(Environment.ExpandEnvironmentVariables(installDir));
        var exePath = Path.Combine(normalizedInstallDir, ToolExeName);
        if (!File.Exists(exePath))
        {
            throw new InvalidOperationException($"未找到可执行文件：{exePath}");
        }

        var desktopPath = ApplyShortcutSetting(
            BuildShortcutCandidates(Environment.SpecialFolder.CommonDesktopDirectory, Environment.SpecialFolder.DesktopDirectory),
            createDesktopShortcut,
            exePath,
            normalizedInstallDir);

        var startMenuPath = ApplyShortcutSetting(
            BuildShortcutCandidates(Environment.SpecialFolder.CommonPrograms, Environment.SpecialFolder.Programs),
            createStartMenuShortcut,
            exePath,
            normalizedInstallDir);

        return new ShortcutResult(desktopPath, startMenuPath);
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

    private static IReadOnlyList<string> BuildShortcutCandidates(Environment.SpecialFolder preferred, Environment.SpecialFolder fallback)
    {
        var result = new List<string>();

        var preferredPath = Environment.GetFolderPath(preferred);
        if (!string.IsNullOrWhiteSpace(preferredPath))
        {
            result.Add(Path.Combine(preferredPath, $"{ToolDisplayName}.lnk"));
        }

        var fallbackPath = Environment.GetFolderPath(fallback);
        if (!string.IsNullOrWhiteSpace(fallbackPath))
        {
            result.Add(Path.Combine(fallbackPath, $"{ToolDisplayName}.lnk"));
        }

        return result
            .Where(path => !string.IsNullOrWhiteSpace(path))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToArray();
    }

    private static string? ApplyShortcutSetting(
        IReadOnlyList<string> candidates,
        bool enabled,
        string exePath,
        string workingDirectory)
    {
        if (candidates.Count == 0)
        {
            return null;
        }

        var primary = candidates[0];
        if (!enabled)
        {
            foreach (var path in candidates)
            {
                TryDeleteFile(path);
            }

            return null;
        }

        var dir = Path.GetDirectoryName(primary);
        if (!string.IsNullOrWhiteSpace(dir))
        {
            Directory.CreateDirectory(dir);
        }

        CreateShortcut(primary, exePath, workingDirectory);

        foreach (var other in candidates.Skip(1))
        {
            TryDeleteFile(other);
        }

        return primary;
    }

    private static void TryDeleteFile(string path)
    {
        try
        {
            if (File.Exists(path))
            {
                File.Delete(path);
            }
        }
        catch
        {
            // Ignore cleanup failures to avoid blocking installation flow.
        }
    }

    private static void CreateShortcut(string shortcutPath, string targetPath, string workingDirectory)
    {
        var shellType = Type.GetTypeFromProgID("WScript.Shell")
            ?? throw new InvalidOperationException("无法创建快捷方式：系统缺少 WScript.Shell。");

        object? shell = null;
        object? shortcut = null;
        try
        {
            shell = Activator.CreateInstance(shellType);
            if (shell is null)
            {
                throw new InvalidOperationException("无法创建快捷方式：COM 初始化失败。");
            }

            shortcut = shellType.InvokeMember(
                "CreateShortcut",
                BindingFlags.InvokeMethod,
                null,
                shell,
                [shortcutPath]);

            if (shortcut is null)
            {
                throw new InvalidOperationException("无法创建快捷方式对象。");
            }

            var shortcutType = shortcut.GetType();
            shortcutType.InvokeMember("TargetPath", BindingFlags.SetProperty, null, shortcut, [targetPath]);
            shortcutType.InvokeMember("WorkingDirectory", BindingFlags.SetProperty, null, shortcut, [workingDirectory]);
            shortcutType.InvokeMember("Description", BindingFlags.SetProperty, null, shortcut, [ToolDisplayName]);
            shortcutType.InvokeMember("IconLocation", BindingFlags.SetProperty, null, shortcut, [$"{targetPath},0"]);
            shortcutType.InvokeMember("Save", BindingFlags.InvokeMethod, null, shortcut, null);
        }
        finally
        {
            if (shortcut is not null && Marshal.IsComObject(shortcut))
            {
                Marshal.FinalReleaseComObject(shortcut);
            }

            if (shell is not null && Marshal.IsComObject(shell))
            {
                Marshal.FinalReleaseComObject(shell);
            }
        }
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

    private static bool IsAnyToolProcessRunning()
    {
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
                if (string.IsNullOrWhiteSpace(processPath))
                {
                    return true;
                }

                var fileName = Path.GetFileName(processPath);
                if (string.Equals(fileName, ToolExeName, StringComparison.OrdinalIgnoreCase))
                {
                    return true;
                }
            }
            catch
            {
                // Cannot inspect process details safely; block install to avoid false-success overwrite.
                return true;
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
