using System.Globalization;
using System.Management;
using System.Text.RegularExpressions;

namespace PbiMetadataTool;

internal sealed class PowerBiInstanceDetector
{
    private static readonly Regex StartPathRegex = new(
        "-s\\s+(?:\"(?<start>[^\"]+)\"|(?<start>\\S+))",
        RegexOptions.IgnoreCase | RegexOptions.Compiled);

    private static readonly Regex PbixPathRegex = new(
        "(?:\"(?<pbix>[^\"]+\\.(?:pbix|pbit|pbip))\"|(?<pbix>\\S+\\.(?:pbix|pbit|pbip)))",
        RegexOptions.IgnoreCase | RegexOptions.Compiled);

    private static readonly Regex IniPortRegex = new(
        "<Port>(?<port>\\d+)</Port>",
        RegexOptions.IgnoreCase | RegexOptions.Compiled);

    public IReadOnlyList<PowerBiInstanceInfo> DiscoverInstances()
    {
        var discovered = new List<PowerBiInstanceInfo>();
        var desktopProcesses = QueryDesktopProcesses().ToDictionary(x => x.ProcessId);
        var msmdsrvProcesses = QueryProcesses("msmdsrv.exe");

        foreach (var msmd in msmdsrvProcesses)
        {
            if (!desktopProcesses.ContainsKey(msmd.ParentProcessId))
            {
                continue;
            }

            var workspacePath = ExtractWorkspacePath(msmd.CommandLine);
            if (string.IsNullOrWhiteSpace(workspacePath))
            {
                continue;
            }

            var normalizedWorkspace = NormalizePath(workspacePath);
            if (normalizedWorkspace is null)
            {
                continue;
            }

            var port = TryReadPort(normalizedWorkspace);
            if (!port.HasValue)
            {
                continue;
            }

            var desktopProcess = desktopProcesses[msmd.ParentProcessId];
            var pbixPath = ExtractPbixPath(desktopProcess.CommandLine);
            var lastSeenUtc = GetWorkspaceTimestampUtc(workspacePath);

            discovered.Add(new PowerBiInstanceInfo(
                desktopProcess.ProcessId,
                msmd.ProcessId,
                port.Value,
                normalizedWorkspace,
                pbixPath,
                lastSeenUtc));
        }

        // Only use workspace fallback when process-based discovery found nothing.
        // This avoids showing stale "PID 0" cache entries alongside real instances.
        if (discovered.Count == 0)
        {
            AddWorkspaceFallback(discovered);
        }

        return discovered
            .GroupBy(x => x.Port)
            .Select(g => g.OrderByDescending(x => x.LastSeenUtc).First())
            .OrderByDescending(x => x.LastSeenUtc)
            .ThenBy(x => x.Port)
            .ToList();
    }

    private static IEnumerable<ProcessCommandInfo> QueryDesktopProcesses()
    {
        var names = new[]
        {
            "PBIDesktop.exe",
            "PBIDesktopStore.exe",
            "PBIDesktopRS.exe",
            "PBIRSDesktop.exe"
        };

        return names
            .SelectMany(QueryProcesses)
            .GroupBy(p => p.ProcessId)
            .Select(g => g.First());
    }

    private static void AddWorkspaceFallback(ICollection<PowerBiInstanceInfo> discovered)
    {
        // 已通过进程匹配发现的端口，fallback 不重复添加
        var knownPorts = new HashSet<int>(discovered.Select(x => x.Port));

        var roots = WorkspaceRoots();
        foreach (var root in roots.Where(Directory.Exists))
        {
            IEnumerable<string> workspacePaths;
            try
            {
                workspacePaths = Directory.EnumerateDirectories(root);
            }
            catch
            {
                continue;
            }

            foreach (var workspacePath in workspacePaths)
            {
                var port = TryReadPort(workspacePath);
                if (!port.HasValue)
                {
                    continue;
                }

                // 跳过已经由进程匹配发现的端口
                if (knownPorts.Contains(port.Value))
                {
                    continue;
                }

                var time = Directory.GetLastWriteTimeUtc(workspacePath);
                discovered.Add(new PowerBiInstanceInfo(
                    DesktopPid: 0,
                    MsmdsrvPid: 0,
                    Port: port.Value,
                    WorkspacePath: workspacePath,
                    PbixPathHint: null,
                    LastSeenUtc: time));
                knownPorts.Add(port.Value);
            }
        }
    }

    private static IReadOnlyList<string> WorkspaceRoots()
    {
        var local = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
        var userProfile = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
        var roots = new List<string>
        {
            Path.Combine(local, "Microsoft", "Power BI Desktop", "AnalysisServicesWorkspaces"),
            Path.Combine(local, "Microsoft", "Power BI Desktop Store App", "AnalysisServicesWorkspaces"),
            Path.Combine(local, "Microsoft", "Power BI Desktop SSRS", "AnalysisServicesWorkspaces"),
            Path.Combine(userProfile, "Microsoft", "Power BI Desktop", "AnalysisServicesWorkspaces"),
            Path.Combine(userProfile, "Microsoft", "Power BI Desktop Store App", "AnalysisServicesWorkspaces"),
            Path.Combine(userProfile, "Microsoft", "Power BI Desktop SSRS", "AnalysisServicesWorkspaces")
        };

        var packagesRoot = Path.Combine(local, "Packages");
        if (Directory.Exists(packagesRoot))
        {
            foreach (var pkg in Directory.EnumerateDirectories(packagesRoot, "Microsoft.MicrosoftPowerBIDesktop_*", SearchOption.TopDirectoryOnly))
            {
                roots.Add(Path.Combine(pkg, "LocalCache", "Local", "Microsoft", "Power BI Desktop", "AnalysisServicesWorkspaces"));
                roots.Add(Path.Combine(pkg, "LocalState", "AnalysisServicesWorkspaces"));
            }
        }

        return roots
            .Where(path => !string.IsNullOrWhiteSpace(path))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();
    }

    private static IEnumerable<ProcessCommandInfo> QueryProcesses(string processName)
    {
        var results = new List<ProcessCommandInfo>();
        var query = $"SELECT ProcessId, ParentProcessId, CommandLine FROM Win32_Process WHERE Name='{processName}'";

        try
        {
            using var searcher = new ManagementObjectSearcher(query);
            using var collection = searcher.Get();
            foreach (var item in collection.Cast<ManagementObject>())
            {
                var processId = Convert.ToInt32(item["ProcessId"], CultureInfo.InvariantCulture);
                var parentProcessId = Convert.ToInt32(item["ParentProcessId"], CultureInfo.InvariantCulture);
                var commandLine = item["CommandLine"]?.ToString() ?? string.Empty;
                results.Add(new ProcessCommandInfo(processId, parentProcessId, commandLine));
            }
        }
        catch
        {
            return Array.Empty<ProcessCommandInfo>();
        }

        return results;
    }

    private static string? ExtractWorkspacePath(string? commandLine)
    {
        if (string.IsNullOrWhiteSpace(commandLine))
        {
            return null;
        }

        var match = StartPathRegex.Match(commandLine);
        if (!match.Success)
        {
            return null;
        }

        var path = match.Groups["start"].Value;
        if (string.IsNullOrWhiteSpace(path))
        {
            return null;
        }

        var normalized = NormalizePath(path);
        if (normalized is null)
        {
            return null;
        }

        var canonical = normalized
            .Replace(Path.AltDirectorySeparatorChar, Path.DirectorySeparatorChar)
            .TrimEnd(Path.DirectorySeparatorChar);

        if (canonical.EndsWith($"{Path.DirectorySeparatorChar}msmdsrv.ini", StringComparison.OrdinalIgnoreCase))
        {
            var dataDir = Path.GetDirectoryName(canonical);
            return dataDir is null ? null : Directory.GetParent(dataDir)?.FullName;
        }

        if (canonical.EndsWith($"{Path.DirectorySeparatorChar}Data", StringComparison.OrdinalIgnoreCase))
        {
            return Directory.GetParent(canonical)?.FullName ?? canonical;
        }

        var workspaceMatch = Regex.Match(
            canonical,
            @"(?<ws>.+[\\/](AnalysisServicesWorkspace_[^\\/]+))([\\/].*)?$",
            RegexOptions.IgnoreCase);

        if (workspaceMatch.Success)
        {
            return workspaceMatch.Groups["ws"].Value;
        }

        return canonical;
    }

    private static string? ExtractPbixPath(string? commandLine)
    {
        if (string.IsNullOrWhiteSpace(commandLine))
        {
            return null;
        }

        var match = PbixPathRegex.Match(commandLine);
        if (!match.Success)
        {
            return null;
        }

        var path = match.Groups["pbix"].Value;
        return NormalizePath(path);
    }

    private static string? NormalizePath(string path)
    {
        try
        {
            return Path.GetFullPath(Environment.ExpandEnvironmentVariables(path));
        }
        catch
        {
            return null;
        }
    }

    private static int? TryReadPort(string workspacePath)
    {
        var iniPath = Path.Combine(workspacePath, "Data", "msmdsrv.ini");
        var candidates = new[]
        {
            Path.Combine(workspacePath, "msmdsrv.port.txt"),
            Path.Combine(workspacePath, "Data", "msmdsrv.port.txt"),
            Path.Combine(Path.GetDirectoryName(iniPath) ?? string.Empty, "msmdsrv.port.txt")
        };

        foreach (var candidate in candidates.Where(path => !string.IsNullOrWhiteSpace(path)))
        {
            if (!File.Exists(candidate))
            {
                continue;
            }

            var text = File.ReadAllText(candidate);
            if (TryParsePort(text, out var parsed))
            {
                return parsed;
            }
        }

        if (File.Exists(iniPath))
        {
            var iniContent = File.ReadAllText(iniPath);
            var match = IniPortRegex.Match(iniContent);
            if (match.Success &&
                int.TryParse(match.Groups["port"].Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out var iniPort) &&
                IsValidPort(iniPort))
            {
                return iniPort;
            }
        }

        return null;
    }

    private static bool TryParsePort(string text, out int port)
    {
        port = 0;
        if (string.IsNullOrWhiteSpace(text))
        {
            return false;
        }

        // Strip null bytes and whitespace
        var normalized = text.Replace("\0", string.Empty).Trim();

        // Direct parse (e.g. "4721")
        if (int.TryParse(normalized, NumberStyles.Integer, CultureInfo.InvariantCulture, out var direct) &&
            IsValidPort(direct))
        {
            port = direct;
            return true;
        }

        // Store App writes port as space-separated digits: "4 7 2 1" → "4721"
        var noSpaces = normalized.Replace(" ", string.Empty);
        if (int.TryParse(noSpaces, NumberStyles.Integer, CultureInfo.InvariantCulture, out var joined) &&
            IsValidPort(joined))
        {
            port = joined;
            return true;
        }

        // Fallback: find last 4-6 digit sequence
        var allPorts = Regex.Matches(normalized, "\\d{4,6}");
        if (allPorts.Count > 0 &&
            int.TryParse(allPorts[^1].Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out var matched) &&
            IsValidPort(matched))
        {
            port = matched;
            return true;
        }

        return false;
    }

    private static bool IsValidPort(int port) => port is > 0 and <= 65535;

    private static DateTime GetWorkspaceTimestampUtc(string workspacePath)
    {
        try
        {
            return Directory.GetLastWriteTimeUtc(workspacePath);
        }
        catch
        {
            return DateTime.UtcNow;
        }
    }

    private sealed record ProcessCommandInfo(int ProcessId, int ParentProcessId, string CommandLine);
}
