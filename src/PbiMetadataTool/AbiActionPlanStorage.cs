using System.Text.Json;

namespace PbiMetadataTool;

internal static class AbiActionPlanStorage
{
    public static void Save(string path, AbiActionPlan plan)
    {
        var directory = Path.GetDirectoryName(path);
        if (!string.IsNullOrWhiteSpace(directory) && !Directory.Exists(directory))
        {
            Directory.CreateDirectory(directory);
        }

        var payload = new
        {
            abi_action_plan = new
            {
                summary = plan.Summary,
                actions = plan.Actions
            }
        };

        var json = JsonSerializer.Serialize(payload, new JsonSerializerOptions
        {
            WriteIndented = true,
            PropertyNamingPolicy = JsonNamingPolicy.CamelCase
        });
        File.WriteAllText(path, json);
    }

    public static AbiActionPlan Load(string path)
    {
        var text = File.ReadAllText(path);
        if (!AbiActionPlanParser.TryParseJsonText(text, out var plan, out var error, allowEmptyActions: true))
        {
            throw new InvalidOperationException(error);
        }

        return plan;
    }
}
