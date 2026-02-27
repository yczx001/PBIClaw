namespace PbiMetadataTool;

internal sealed record CliOptions(
    bool ShowHelp,
    bool ListInstances,
    int? InstanceIndex,
    int? Port,
    string? Server,
    string? DatabaseName,
    string? OutputPath,
    bool ExternalToolMode)
{
    public static CliOptions Parse(string[] args)
    {
        var showHelp = false;
        var listInstances = false;
        int? instanceIndex = null;
        int? port = null;
        string? server = null;
        string? databaseName = null;
        string? outputPath = null;
        var externalToolMode = false;

        for (var i = 0; i < args.Length; i++)
        {
            var arg = args[i];

            switch (arg)
            {
                case "--help":
                case "-h":
                case "/?":
                    showHelp = true;
                    break;
                case "--list-instances":
                    listInstances = true;
                    break;
                case "--instance-index":
                    instanceIndex = ParseRequiredInt(args, ref i, "--instance-index");
                    break;
                case "--port":
                    port = ParseRequiredInt(args, ref i, "--port");
                    break;
                case "--server":
                    server = ParseRequiredString(args, ref i, "--server");
                    break;
                case "--database":
                    databaseName = ParseRequiredString(args, ref i, "--database");
                    break;
                case "--out":
                    outputPath = ParseRequiredString(args, ref i, "--out");
                    break;
                case "--external-tool":
                    externalToolMode = true;
                    break;
                default:
                    throw new ArgumentException($"未知参数: {arg}");
            }
        }

        if (args.Length == 0)
        {
            externalToolMode = true;
        }

        return new CliOptions(showHelp, listInstances, instanceIndex, port, server, databaseName, outputPath, externalToolMode);
    }

    private static int ParseRequiredInt(string[] args, ref int index, string optionName)
    {
        var value = ParseRequiredString(args, ref index, optionName);
        if (int.TryParse(value, out var parsed))
        {
            return parsed;
        }

        throw new ArgumentException($"参数 {optionName} 需要整数值，当前为: {value}");
    }

    private static string ParseRequiredString(string[] args, ref int index, string optionName)
    {
        if (index + 1 >= args.Length)
        {
            throw new ArgumentException($"参数 {optionName} 缺少值");
        }

        index++;
        return args[index];
    }
}
