using System.Text.Json;
using System.Text.RegularExpressions;

namespace PbiMetadataTool;

internal static class AbiActionPlanParser
{
    private static readonly Regex JsonCodeBlockRegex = new(
        "```(?:json)?\\s*(?<json>\\{.*?\\})\\s*```",
        RegexOptions.IgnoreCase | RegexOptions.Singleline | RegexOptions.Compiled);

    public static bool TryExtract(string assistantReply, out AbiActionPlan plan, out string preview, out string error)
    {
        foreach (Match match in JsonCodeBlockRegex.Matches(assistantReply))
        {
            var raw = match.Groups["json"].Value;
            if (TryParseFromJson(raw, out plan, out preview, out error))
            {
                return true;
            }
        }

        plan = new AbiActionPlan(string.Empty, []);
        preview = string.Empty;
        error = "回复中未找到可执行的 abi_action_plan。";
        return false;
    }

    public static bool TryParseJsonText(string jsonText, out AbiActionPlan plan, out string error, bool allowEmptyActions = false)
    {
        if (TryParseFromJson(jsonText, out plan, out _, out error, allowEmptyActions))
        {
            return true;
        }

        return false;
    }

    private static bool TryParseFromJson(string json, out AbiActionPlan plan, out string preview, out string error, bool allowEmptyActions = false)
    {
        try
        {
            using var doc = JsonDocument.Parse(json);
            var root = doc.RootElement;
            var payload = root;

            if (root.ValueKind == JsonValueKind.Object &&
                TryGetPropertyIgnoreCase(root, "abi_action_plan", out var nestedPlan) &&
                nestedPlan.ValueKind == JsonValueKind.Object)
            {
                payload = nestedPlan;
            }

            if (payload.ValueKind != JsonValueKind.Object ||
                !TryGetPropertyIgnoreCase(payload, "actions", out var actionsElement) ||
                actionsElement.ValueKind != JsonValueKind.Array)
            {
                plan = new AbiActionPlan(string.Empty, []);
                preview = string.Empty;
                error = "JSON 未包含 actions 数组。";
                return false;
            }

            var summary = TryGetPropertyIgnoreCase(payload, "summary", out var summaryElement)
                ? summaryElement.GetString() ?? string.Empty
                : string.Empty;

            var actions = new List<AbiModelAction>();
            foreach (var item in actionsElement.EnumerateArray())
            {
                if (item.ValueKind != JsonValueKind.Object)
                {
                    continue;
                }

                var type = GetString(item, "type");
                if (string.IsNullOrWhiteSpace(type))
                {
                    continue;
                }

                actions.Add(new AbiModelAction(
                    Type: type.Trim(),
                    Reason: GetString(item, "reason"),
                    Table: GetString(item, "table"),
                    Name: GetString(item, "name"),
                    NewName: GetString(item, "newName"),
                    Expression: GetString(item, "expression"),
                    FormatString: GetString(item, "formatString"),
                    DisplayFolder: GetString(item, "displayFolder"),
                    Description: GetString(item, "description"),
                    ObjectType: GetString(item, "objectType"),
                    DataType: GetString(item, "dataType"),
                    ModelPermission: GetString(item, "modelPermission"),
                    MetadataPermission: GetString(item, "metadataPermission"),
                    MemberName: GetString(item, "memberName"),
                    IdentityProvider: GetString(item, "identityProvider"),
                    MemberType: GetString(item, "memberType"),
                    Operation: GetString(item, "operation"),
                    IsHidden: GetBool(item, "isHidden"),
                    FromTable: GetString(item, "fromTable"),
                    FromColumn: GetString(item, "fromColumn"),
                    ToTable: GetString(item, "toTable"),
                    ToColumn: GetString(item, "toColumn"),
                    CrossFilterDirection: GetString(item, "crossFilterDirection"),
                    IsActive: GetBool(item, "isActive")));
            }

            if (actions.Count == 0)
            {
                if (allowEmptyActions)
                {
                    plan = new AbiActionPlan(summary, []);
                    preview = AbiActionPlanPreview.BuildText(plan);
                    error = string.Empty;
                    return true;
                }

                plan = new AbiActionPlan(string.Empty, []);
                preview = string.Empty;
                error = "actions 为空，无法执行。";
                return false;
            }

            plan = new AbiActionPlan(summary, actions);
            preview = AbiActionPlanPreview.BuildText(plan);
            error = string.Empty;
            return true;
        }
        catch (Exception ex)
        {
            plan = new AbiActionPlan(string.Empty, []);
            preview = string.Empty;
            error = $"JSON 解析失败: {ex.Message}";
            return false;
        }
    }

    private static string? GetString(JsonElement element, string name)
    {
        if (!TryGetPropertyIgnoreCase(element, name, out var value))
        {
            return null;
        }

        return value.ValueKind switch
        {
            JsonValueKind.String => value.GetString(),
            JsonValueKind.Number => value.GetRawText(),
            JsonValueKind.True => "true",
            JsonValueKind.False => "false",
            _ => value.GetRawText()
        };
    }

    private static bool? GetBool(JsonElement element, string name)
    {
        if (!TryGetPropertyIgnoreCase(element, name, out var value))
        {
            return null;
        }

        if (value.ValueKind == JsonValueKind.True)
        {
            return true;
        }

        if (value.ValueKind == JsonValueKind.False)
        {
            return false;
        }

        if (value.ValueKind == JsonValueKind.String &&
            bool.TryParse(value.GetString(), out var parsed))
        {
            return parsed;
        }

        return null;
    }

    private static bool TryGetPropertyIgnoreCase(JsonElement element, string propertyName, out JsonElement value)
    {
        if (element.ValueKind != JsonValueKind.Object)
        {
            value = default;
            return false;
        }

        if (element.TryGetProperty(propertyName, out value))
        {
            return true;
        }

        foreach (var property in element.EnumerateObject())
        {
            if (string.Equals(property.Name, propertyName, StringComparison.OrdinalIgnoreCase))
            {
                value = property.Value;
                return true;
            }
        }

        value = default;
        return false;
    }
}
