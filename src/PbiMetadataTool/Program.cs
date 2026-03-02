using System.Text.Json;
using System.Windows.Forms;
using System.Text.RegularExpressions;

namespace PbiMetadataTool;

internal static class Program
{
    [STAThread]
    private static int Main(string[] args)
    {
        try
        {
            InitializeWinFormsRendering();
            var options = CliOptions.Parse(args);
            if (options.ShowHelp)
            {
                PrintHelp();
                return 0;
            }

            if (options.ExternalToolMode)
            {
                Application.Run(new MainFormWebView(options));
                return 0;
            }

            var detector = new PowerBiInstanceDetector();
            var instances = detector.DiscoverInstances();

            if (options.ListInstances)
            {
                PrintInstances(instances);
                return 0;
            }

            var targetPort = ResolvePort(options, instances);
            if (!targetPort.HasValue)
            {
                var error = "未找到可用的 Power BI Desktop 模型端口。请确认 PBIDesktop 已打开并加载了模型。";
                Console.Error.WriteLine(error);
                MessageBox.Show(
                    $"Run failed:\n{error}",
                    "PBI Claw",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
                return 2;
            }

            var metadataReader = new TabularMetadataReader();
            var metadata = metadataReader.ReadMetadata(targetPort.Value, options.DatabaseName);
            var outputPath = options.OutputPath ?? BuildDefaultOutputPath(targetPort.Value);

            var serializerOptions = new JsonSerializerOptions
            {
                WriteIndented = true
            };

            var json = JsonSerializer.Serialize(metadata, serializerOptions);
            EnsureOutputDirectory(outputPath);
            File.WriteAllText(outputPath, json);

            Console.WriteLine($"连接端口: {targetPort.Value}");
            Console.WriteLine($"数据库: {metadata.DatabaseName}");
            Console.WriteLine($"表数量: {metadata.Tables.Count}");
            Console.WriteLine($"关系数量: {metadata.Relationships.Count}");
            Console.WriteLine($"角色数量: {metadata.Roles.Count}");
            Console.WriteLine($"元数据输出: {Path.GetFullPath(outputPath)}");

            return 0;
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"执行失败: {ex.Message}");
            MessageBox.Show(
                $"Run failed:\n{ex.Message}",
                "PBI Claw",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error);
            return 1;
        }
    }

    private static void InitializeWinFormsRendering()
    {
        // Improve text and control sharpness on modern/high-DPI displays.
        Application.SetHighDpiMode(HighDpiMode.PerMonitorV2);
        Application.EnableVisualStyles();
        Application.SetCompatibleTextRenderingDefault(false);
    }

    private static int? ResolvePort(CliOptions options, IReadOnlyList<PowerBiInstanceInfo> instances)
    {
        var server = NormalizeExternalToolValue(options.Server);
        if (!string.IsNullOrWhiteSpace(server))
        {
            var parsedFromServer = TryParsePortFromServer(server);
            if (parsedFromServer.HasValue)
            {
                return parsedFromServer.Value;
            }
        }

        if (options.Port.HasValue)
        {
            return options.Port.Value;
        }

        if (instances.Count == 0)
        {
            return null;
        }

        if (options.InstanceIndex.HasValue)
        {
            var index = options.InstanceIndex.Value;
            if (index < 0 || index >= instances.Count)
            {
                throw new ArgumentOutOfRangeException(nameof(options.InstanceIndex), $"实例下标超出范围: {index}");
            }

            return instances[index].Port;
        }

        return instances[0].Port;
    }

    private static int? TryParsePortFromServer(string server)
    {
        if (string.IsNullOrWhiteSpace(server))
        {
            return null;
        }

        // Power BI external tools may pass values like localhost:xxxxx or localhost:xxxxx;
        var parts = server.Split(':', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
        if (parts.Length == 2 && int.TryParse(parts[1], out var directPort))
        {
            return directPort;
        }

        if (int.TryParse(server, out var rawPort))
        {
            return rawPort;
        }

        var allPorts = Regex.Matches(server, "\\d{4,6}");
        if (allPorts.Count > 0 && int.TryParse(allPorts[^1].Value, out var matchedPort))
        {
            return matchedPort;
        }

        return null;
    }

    private static string? NormalizeExternalToolValue(string? value)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            return null;
        }

        var trimmed = value.Trim();
        return Regex.IsMatch(trimmed, "^%[^%]+%$") ? null : trimmed;
    }

    private static string BuildDefaultOutputPath(int port)
    {
        var timestamp = DateTime.Now.ToString("yyyyMMdd-HHmmss");
        var baseDir = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
            "PBIClaw");
        return Path.Combine(baseDir, $"abi-model-{port}-{timestamp}.json");
    }

    private static void EnsureOutputDirectory(string outputPath)
    {
        var directory = Path.GetDirectoryName(outputPath);
        if (!string.IsNullOrWhiteSpace(directory) && !Directory.Exists(directory))
        {
            Directory.CreateDirectory(directory);
        }
    }

    private static void PrintInstances(IReadOnlyList<PowerBiInstanceInfo> instances)
    {
        if (instances.Count == 0)
        {
            Console.WriteLine("未发现可用实例。");
            return;
        }

        for (var i = 0; i < instances.Count; i++)
        {
            var instance = instances[i];
            Console.WriteLine($"[{i}] port={instance.Port} desktopPid={instance.DesktopPid} msmdsrvPid={instance.MsmdsrvPid}");
            Console.WriteLine($"    workspace={instance.WorkspacePath}");
            if (!string.IsNullOrWhiteSpace(instance.PbixPathHint))
            {
                Console.WriteLine($"    pbix={instance.PbixPathHint}");
            }
        }
    }

    private static void PrintHelp()
    {
        Console.WriteLine("PBI Claw");
        Console.WriteLine("自动识别 Power BI Desktop 端口，连接语义模型并支持 AI 对话分析。");
        Console.WriteLine();
        Console.WriteLine("参数:");
        Console.WriteLine("  --list-instances          列出检测到的实例");
        Console.WriteLine("  --instance-index <index>  指定实例下标");
        Console.WriteLine("  --port <port>             手动指定端口");
        Console.WriteLine("  --server <host:port>      直接指定服务器（支持 localhost:xxxxx）");
        Console.WriteLine("  --database <name>         指定数据库名(默认首个数据库)");
        Console.WriteLine("  --out <path>              指定输出 JSON 路径");
        Console.WriteLine("  --external-tool           以外部工具模式运行（PBI Claw 图形界面）");
        Console.WriteLine("  --help                    查看帮助");
    }
}
