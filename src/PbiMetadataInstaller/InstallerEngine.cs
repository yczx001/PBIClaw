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
    private const string ShortcutIconName = "PBIClaw.round.ico";
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

    public static string? GetExistingInstallDir()
    {
        // 从注册表查找已安装路径
        using var key = Microsoft.Win32.Registry.LocalMachine.OpenSubKey(@"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\PBI Claw");
        if (key != null)
        {
            var installLocation = key.GetValue("InstallLocation") as string;
            if (!string.IsNullOrWhiteSpace(installLocation) && Directory.Exists(installLocation))
            {
                return installLocation;
            }
        }

        // 检查默认安装路径
        var defaultDir = GetDefaultInstallDir();
        if (Directory.Exists(defaultDir) && File.Exists(Path.Combine(defaultDir, ToolExeName)))
        {
            return defaultDir;
        }

        return null;
    }

    public static void Uninstall(string installDir)
    {
        if (!IsAdministrator())
        {
            throw new InvalidOperationException("卸载需要管理员权限。");
        }

        var normalizedInstallDir = Path.GetFullPath(Environment.ExpandEnvironmentVariables(installDir));
        if (!Directory.Exists(normalizedInstallDir))
        {
            throw new InvalidOperationException($"安装目录不存在：{normalizedInstallDir}");
        }

        if (IsAnyToolProcessRunning())
        {
            throw new InvalidOperationException("检测到 PBIClaw 仍在运行，请先关闭所有 PBIClaw 进程后再重试卸载。");
        }

        // 1. 删除Power BI外部工具配置
        var externalToolDirs = GetMachineExternalToolDirs();
        foreach (var externalToolDir in externalToolDirs)
        {
            var jsonPath = Path.Combine(externalToolDir, ToolJsonName);
            TryDeleteFile(jsonPath);
        }

        // 2. 删除快捷方式
        var shortcutCandidates = new List<string>();
        // 桌面快捷方式
        shortcutCandidates.AddRange(BuildShortcutCandidates(Environment.SpecialFolder.CommonDesktopDirectory, Environment.SpecialFolder.DesktopDirectory));
        // 开始菜单快捷方式
        shortcutCandidates.AddRange(BuildShortcutCandidates(Environment.SpecialFolder.CommonPrograms, Environment.SpecialFolder.Programs));

        foreach (var shortcutPath in shortcutCandidates)
        {
            TryDeleteFile(shortcutPath);
        }

        // 3. 删除注册表卸载信息
        try
        {
            Microsoft.Win32.Registry.LocalMachine.DeleteSubKeyTree(@"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\PBI Claw", false);
        }
        catch
        {
            // 忽略注册表删除错误
        }

        // 4. 删除安装目录文件
        try
        {
            Directory.Delete(normalizedInstallDir, recursive: true);
        }
        catch
        {
            // 如果删除失败，安排下次重启时删除
            try
            {
                var moveDir = Path.Combine(Path.GetTempPath(), $"PBIClaw_Uninstall_{Guid.NewGuid():N}");
                Directory.Move(normalizedInstallDir, moveDir);

                // 使用MoveFileEx安排重启删除
                MoveFileEx(moveDir, null, MoveFileFlags.MOVEFILE_DELAY_UNTIL_REBOOT);
            }
            catch
            {
                // 忽略最终删除错误，至少已经移除了所有入口
            }
        }
    }

    public static void RegisterUninstallInformation(string installDir, string version)
    {
        if (!IsAdministrator())
        {
            return; // 非管理员安装不写入注册表
        }

        try
        {
            using var key = Microsoft.Win32.Registry.LocalMachine.CreateSubKey(@"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\PBI Claw");
            if (key == null) return;

            var exePath = Path.Combine(installDir, ToolExeName);
            var setupExePath = Environment.ProcessPath ?? exePath;

            key.SetValue("DisplayName", ToolDisplayName);
            key.SetValue("DisplayVersion", version);
            key.SetValue("Publisher", "PBI Hub");
            key.SetValue("InstallLocation", installDir);
            key.SetValue("InstallDate", DateTime.Now.ToString("yyyyMMdd"));
            key.SetValue("UninstallString", $"\"{setupExePath}\" --uninstall");
            key.SetValue("ModifyPath", $"\"{setupExePath}\" --install-dir \"{installDir}\"");
            key.SetValue("DisplayIcon", $"{exePath},0");
            key.SetValue("EstimatedSize", GetDirectorySize(installDir) / 1024); // KB
            key.SetValue("URLInfoAbout", "https://pbihub.cn");
            key.SetValue("NoRepair", 1);
            key.SetValue("NoModify", 0);
        }
        catch
        {
            // 忽略注册表写入错误，不影响安装
        }
    }

    private static long GetDirectorySize(string path)
    {
        long size = 0;
        try
        {
            var dirInfo = new DirectoryInfo(path);
            foreach (var file in dirInfo.GetFiles("*", SearchOption.AllDirectories))
            {
                size += file.Length;
            }
        }
        catch
        {
            // 忽略大小计算错误
        }
        return size;
    }

    [DllImport("kernel32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
    private static extern bool MoveFileEx(string lpExistingFileName, string? lpNewFileName, MoveFileFlags dwFlags);

    [Flags]
    private enum MoveFileFlags : uint
    {
        MOVEFILE_DELAY_UNTIL_REBOOT = 0x00000004
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
        var shortcutIconPath = Path.Combine(normalizedInstallDir, ShortcutIconName);
        TryExtractEmbeddedResource($"Payload.{ShortcutIconName}", shortcutIconPath);

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
        var shortcutIconPath = Path.Combine(normalizedInstallDir, ShortcutIconName);
        var effectiveIconPath = File.Exists(shortcutIconPath) ? shortcutIconPath : exePath;

        var desktopPath = ApplyShortcutSetting(
            BuildShortcutCandidates(Environment.SpecialFolder.CommonDesktopDirectory, Environment.SpecialFolder.DesktopDirectory),
            createDesktopShortcut,
            exePath,
            normalizedInstallDir,
            effectiveIconPath);

        var startMenuPath = ApplyShortcutSetting(
            BuildShortcutCandidates(Environment.SpecialFolder.CommonPrograms, Environment.SpecialFolder.Programs),
            createStartMenuShortcut,
            exePath,
            normalizedInstallDir,
            effectiveIconPath);

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

    public static void RelaunchAsAdminForUninstall(string installDir)
    {
        var exe = Environment.ProcessPath;
        if (string.IsNullOrWhiteSpace(exe))
        {
            throw new InvalidOperationException("Cannot determine current executable path.");
        }

        var args = $"--uninstall --install-dir \"{installDir}\"";
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
        string workingDirectory,
        string iconPath)
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

        CreateShortcut(primary, exePath, workingDirectory, iconPath);

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

    private static void CreateShortcut(string shortcutPath, string targetPath, string workingDirectory, string iconPath)
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
            shortcutType.InvokeMember("IconLocation", BindingFlags.SetProperty, null, shortcut, [$"{iconPath},0"]);
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

    private static bool TryExtractEmbeddedResource(string logicalName, string outputPath)
    {
        try
        {
            ExtractEmbeddedResource(logicalName, outputPath);
            return true;
        }
        catch (InvalidOperationException)
        {
            return false;
        }
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
