using System.Text;

namespace PbiMetadataTool;

internal sealed record AbiActionPlan(string Summary, IReadOnlyList<AbiModelAction> Actions);

internal sealed record AbiModelAction(
    string Type,
    string? Reason = null,
    string? Table = null,
    string? Name = null,
    string? Expression = null,
    string? FormatString = null,
    bool? IsHidden = null,
    string? FromTable = null,
    string? FromColumn = null,
    string? ToTable = null,
    string? ToColumn = null,
    string? CrossFilterDirection = null,
    bool? IsActive = null);

internal static class AbiActionPlanPreview
{
    public static string BuildText(AbiActionPlan plan)
    {
        var sb = new StringBuilder();
        sb.AppendLine(string.IsNullOrWhiteSpace(plan.Summary) ? "待执行变更计划" : plan.Summary.Trim());
        sb.AppendLine($"动作数量: {plan.Actions.Count}");
        sb.AppendLine();

        for (var i = 0; i < plan.Actions.Count; i++)
        {
            var action = plan.Actions[i];
            sb.AppendLine($"{i + 1}. {action.Type}");

            var target = BuildTarget(action);
            if (!string.IsNullOrWhiteSpace(target))
            {
                sb.AppendLine($"   目标: {target}");
            }

            if (!string.IsNullOrWhiteSpace(action.Expression))
            {
                sb.AppendLine("   表达式:");
                sb.AppendLine(Indent(action.Expression.Trim(), "      "));
            }

            if (!string.IsNullOrWhiteSpace(action.Reason))
            {
                sb.AppendLine($"   原因: {action.Reason.Trim()}");
            }

            sb.AppendLine();
        }

        return sb.ToString().TrimEnd();
    }

    private static string BuildTarget(AbiModelAction action)
    {
        return action.Type switch
        {
            "create_or_update_measure" => $"{action.Table}.{action.Name}",
            "delete_measure" => $"{action.Table}.{action.Name}",
            "create_relationship" => $"{action.FromTable}.{action.FromColumn} -> {action.ToTable}.{action.ToColumn}",
            "delete_relationship" when !string.IsNullOrWhiteSpace(action.Name) => action.Name!,
            "delete_relationship" => $"{action.FromTable}.{action.FromColumn} -> {action.ToTable}.{action.ToColumn}",
            _ => action.Name ?? action.Table ?? string.Empty
        };
    }

    private static string Indent(string text, string prefix)
    {
        return string.Join(Environment.NewLine, text.Split('\n').Select(line => prefix + line.TrimEnd('\r')));
    }
}
