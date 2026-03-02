using System.Text;

namespace PbiMetadataTool;

internal sealed record AbiActionPlan(string Summary, IReadOnlyList<AbiModelAction> Actions);

internal sealed record AbiModelAction(
    string Type,
    string? Reason = null,
    string? Table = null,
    string? Name = null,
    string? NewName = null,
    string? Expression = null,
    string? FormatString = null,
    string? DisplayFolder = null,
    string? Description = null,
    string? ObjectType = null,
    string? DataType = null,
    string? ModelPermission = null,
    string? MetadataPermission = null,
    string? MemberName = null,
    string? IdentityProvider = null,
    string? MemberType = null,
    string? Operation = null,
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
        return action.Type.ToLowerInvariant() switch
        {
            "create_or_update_measure" => $"{action.Table}.{action.Name}",
            "delete_measure" => $"{action.Table}.{action.Name}",
            "delete_table" => action.Table ?? string.Empty,
            "rename_table" => $"{action.Table} -> {action.NewName}",
            "rename_column" => $"{action.Table}.{action.Name} -> {action.NewName}",
            "rename_measure" => $"{action.Table}.{action.Name} -> {action.NewName}",
            "set_table_hidden" => $"{action.Table} | hidden={action.IsHidden}",
            "set_column_hidden" => $"{action.Table}.{action.Name} | hidden={action.IsHidden}",
            "set_measure_hidden" => $"{action.Table}.{action.Name} | hidden={action.IsHidden}",
            "set_format_string" => $"{action.Table}.{action.Name} | {action.FormatString}",
            "set_display_folder" => $"{action.Table}.{action.Name} | {action.DisplayFolder}",
            "create_calculated_column" => $"{action.Table}.{action.Name}",
            "delete_column" => $"{action.Table}.{action.Name}",
            "create_calculated_table" => action.Name ?? string.Empty,
            "set_relationship_active" when !string.IsNullOrWhiteSpace(action.Name) => $"{action.Name} | active={action.IsActive}",
            "set_relationship_active" => $"{action.FromTable}.{action.FromColumn} -> {action.ToTable}.{action.ToColumn} | active={action.IsActive}",
            "set_relationship_cross_filter" when !string.IsNullOrWhiteSpace(action.Name) => $"{action.Name} | {action.CrossFilterDirection}",
            "set_relationship_cross_filter" => $"{action.FromTable}.{action.FromColumn} -> {action.ToTable}.{action.ToColumn} | {action.CrossFilterDirection}",
            "update_description" => $"{action.ObjectType}:{action.Table}.{action.Name}",
            "create_role" => action.Name ?? string.Empty,
            "update_role" => $"{action.Name} -> {action.NewName}",
            "delete_role" => action.Name ?? string.Empty,
            "set_role_table_permission" => $"role={action.Name}, table={action.Table}",
            "remove_role_table_permission" => $"role={action.Name}, table={action.Table}",
            "add_role_member" => $"role={action.Name}, member={action.MemberName}",
            "remove_role_member" => $"role={action.Name}, member={action.MemberName}",
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
