using System.Text;

namespace PbiMetadataTool;

internal static class MetadataPromptBuilder
{
    public static string BuildSystemPrompt(AbiAssistantSettings settings)
    {
        var writeMode = settings.AllowModelChanges ? "允许模型变更" : "只读建议";
        var prompt = $"""
你是 PBI Claw 智能助手，服务于 Power BI 建模与分析。
当前模式：{writeMode}。
输出要求：
1) 先给结论，再给可执行步骤。
2) 涉及 DAX 时，提供可直接粘贴的完整表达式。
3) 如用户需求不明确，先提出最多 3 个关键澄清问题。
4) 在只读建议模式下，不要宣称已修改模型，只能给步骤和代码。
""";

        if (settings.AllowModelChanges)
        {
            prompt += """

5) 【重要】当用户要求创建、新增、修改、删除任何模型对象（度量值、关系等）时，你必须直接输出可执行的 JSON 代码块，而不是给出操作步骤或说明文字让用户自己去操作。系统会自动解析并执行该代码块。
代码块格式必须严格如下（字段名一致，放在回复末尾）：
```json
{
  "abi_action_plan": {
    "summary": "一句话概述本次变更",
    "actions": [
      {
        "type": "create_or_update_measure",
        "table": "表名",
        "name": "度量值名",
        "expression": "DAX表达式",
        "formatString": "#,0.00",
        "isHidden": false,
        "reason": "变更原因"
      }
    ]
  }
}
```
支持的 type：
- create_or_update_measure / delete_measure
- create_relationship / delete_relationship
- delete_table
- rename_table / rename_column / rename_measure
- set_table_hidden / set_column_hidden / set_measure_hidden
- set_format_string / set_display_folder
- create_calculated_column / delete_column
- create_calculated_table
- set_relationship_active / set_relationship_cross_filter
- update_description（objectType 支持 table/column/measure/role）
- create_role / update_role / delete_role
- set_role_table_permission / remove_role_table_permission
- add_role_member / remove_role_member

字段约定：
- rename_* 使用 `newName`
- set_*_hidden 使用 `isHidden`
- set_format_string 使用 `formatString`
- set_display_folder 使用 `displayFolder`
- create_calculated_* 使用 `expression`
- set_relationship_active 使用 `isActive`
- set_relationship_cross_filter 使用 `crossFilterDirection`
- 角色相关：`name` 为角色名，成员使用 `memberName`，表权限可用 `expression`(RLS筛选) 与 `metadataPermission`

关系定位可用 name，或用 fromTable/fromColumn/toTable/toColumn。
如果用户没有明确要求变更，不要输出 abi_action_plan。
如果用户要求变更但信息不完整（如未指定表名），先询问缺失信息，不要猜测输出。
""";
        }
        else
        {
            prompt += """

5) 当前是只读建议模式，禁止输出 abi_action_plan，禁止声称“已执行写回”。
""";
        }

        if (!string.IsNullOrWhiteSpace(settings.CustomSystemPrompt))
        {
            prompt += Environment.NewLine + Environment.NewLine + "附加要求：" + Environment.NewLine + settings.CustomSystemPrompt.Trim();
        }

        return prompt;
    }

    public static string BuildModelContext(ModelMetadata metadata, bool includeHiddenObjects, ReportMetadata? report = null)
    {
        var sb = new StringBuilder();
        sb.AppendLine("以下是当前 Power BI 模型上下文：");
        sb.AppendLine($"数据库: {metadata.DatabaseName}");
        sb.AppendLine($"兼容级别: {metadata.CompatibilityLevel}");
        sb.AppendLine($"表数量: {metadata.Tables.Count}");
        sb.AppendLine($"关系数量: {metadata.Relationships.Count}");
        sb.AppendLine($"角色数量: {metadata.Roles.Count}");
        sb.AppendLine();

        foreach (var table in metadata.Tables.OrderBy(t => t.Name, StringComparer.OrdinalIgnoreCase))
        {
            if (!includeHiddenObjects && table.IsHidden)
            {
                continue;
            }

            sb.AppendLine($"[表] {table.Name}" + (table.IsHidden ? " (Hidden)" : string.Empty));

            var columns = table.Columns
                .Where(c => includeHiddenObjects || !c.IsHidden)
                .OrderBy(c => c.Name, StringComparer.OrdinalIgnoreCase)
                .Select(c => $"  - 列: {c.Name} | 类型: {c.DataType} | 列类型: {c.ColumnType}" + (c.IsHidden ? " | Hidden" : string.Empty));

            foreach (var column in columns)
            {
                sb.AppendLine(column);
            }

            var measures = table.Measures
                .Where(m => includeHiddenObjects || !m.IsHidden)
                .OrderBy(m => m.Name, StringComparer.OrdinalIgnoreCase)
                .Select(m => $"  - 度量值: {m.Name}" + (m.IsHidden ? " | Hidden" : string.Empty));

            foreach (var measure in measures)
            {
                sb.AppendLine(measure);
            }

            sb.AppendLine();
        }

        sb.AppendLine("关系:");
        foreach (var relationship in metadata.Relationships.OrderBy(r => r.Name, StringComparer.OrdinalIgnoreCase))
        {
            sb.AppendLine($"- {relationship.FromTable}.{relationship.FromColumn} -> {relationship.ToTable}.{relationship.ToColumn} | {relationship.CrossFilterDirection} | {(relationship.IsActive ? "Active" : "Inactive")}");
        }

        if (metadata.Roles.Count > 0)
        {
            sb.AppendLine();
            sb.AppendLine("角色:");
            foreach (var role in metadata.Roles.OrderBy(r => r.Name, StringComparer.OrdinalIgnoreCase))
            {
                sb.AppendLine($"- 角色: {role.Name} | 模型权限: {role.ModelPermission}");
                if (!string.IsNullOrWhiteSpace(role.Description))
                {
                    sb.AppendLine($"  说明: {role.Description}");
                }

                foreach (var member in role.Members)
                {
                    var memberType = string.IsNullOrWhiteSpace(member.MemberType) ? string.Empty : $" | 类型: {member.MemberType}";
                    var provider = string.IsNullOrWhiteSpace(member.IdentityProvider) ? string.Empty : $" | 身份源: {member.IdentityProvider}";
                    sb.AppendLine($"  - 成员: {member.Name}{memberType}{provider}");
                }

                foreach (var permission in role.TablePermissions)
                {
                    var filter = string.IsNullOrWhiteSpace(permission.FilterExpression) ? "(空)" : permission.FilterExpression;
                    var metadataPermission = string.IsNullOrWhiteSpace(permission.MetadataPermission) ? string.Empty : $" | 元数据权限: {permission.MetadataPermission}";
                    sb.AppendLine($"  - 表权限: {permission.TableName}{metadataPermission} | 过滤: {filter}");
                }
            }
        }

        if (report is not null)
        {
            sb.AppendLine();
            sb.AppendLine("报表前端信息:");
            sb.AppendLine($"- 来源类型: {report.SourceType}");
            sb.AppendLine($"- 报表页数量: {report.Pages.Count}");
            foreach (var page in report.Pages.Take(30))
            {
                var pageName = string.IsNullOrWhiteSpace(page.DisplayName) ? page.Name : page.DisplayName;
                sb.AppendLine($"[报表页] {pageName} | 视觉对象: {page.Visuals.Count}");
                foreach (var visual in page.Visuals.Take(40))
                {
                    var title = string.IsNullOrWhiteSpace(visual.Title) ? "(无标题)" : visual.Title;
                    sb.AppendLine($"  - 视觉对象: {visual.VisualType} | 标题: {title} | 名称: {visual.Name}");
                }
            }
        }

        return sb.ToString();
    }

    public static string BuildReferencedMeasureContext(ModelMetadata metadata, string userPrompt, bool includeHiddenObjects)
    {
        if (string.IsNullOrWhiteSpace(userPrompt))
        {
            return string.Empty;
        }

        var prompt = userPrompt.Trim();
        var hits = new List<(string Table, MeasureMetadata Measure)>();

        foreach (var table in metadata.Tables)
        {
            if (!includeHiddenObjects && table.IsHidden)
            {
                continue;
            }

            foreach (var measure in table.Measures)
            {
                if (!includeHiddenObjects && measure.IsHidden)
                {
                    continue;
                }

                if (!IsPromptMentioningMeasure(prompt, table.Name, measure.Name))
                {
                    continue;
                }

                hits.Add((table.Name, measure));
                if (hits.Count >= 8)
                {
                    break;
                }
            }

            if (hits.Count >= 8)
            {
                break;
            }
        }

        if (hits.Count == 0)
        {
            return string.Empty;
        }

        var sb = new StringBuilder();
        sb.AppendLine("用户问题中疑似提到了以下度量值，请优先基于这些定义回答：");
        foreach (var hit in hits)
        {
            sb.AppendLine($"- {hit.Table}[{hit.Measure.Name}]");
            if (!string.IsNullOrWhiteSpace(hit.Measure.DisplayFolder))
            {
                sb.AppendLine($"  文件夹: {hit.Measure.DisplayFolder}");
            }

            if (!string.IsNullOrWhiteSpace(hit.Measure.FormatString))
            {
                sb.AppendLine($"  格式: {hit.Measure.FormatString}");
            }

            var expression = (hit.Measure.Expression ?? string.Empty).Trim();
            if (expression.Length > 1200)
            {
                expression = expression[..1200] + Environment.NewLine + "...(已截断)";
            }

            sb.AppendLine("  DAX:");
            sb.AppendLine(IndentBlock(expression, "    "));
        }

        return sb.ToString().TrimEnd();
    }

    private static bool IsPromptMentioningMeasure(string prompt, string tableName, string measureName)
    {
        if (string.IsNullOrWhiteSpace(measureName))
        {
            return false;
        }

        if (prompt.Contains($"{tableName}[{measureName}]", StringComparison.OrdinalIgnoreCase))
        {
            return true;
        }

        if (prompt.Contains($"[{measureName}]", StringComparison.OrdinalIgnoreCase))
        {
            return true;
        }

        return measureName.Length >= 2 &&
               prompt.Contains(measureName, StringComparison.OrdinalIgnoreCase);
    }

    private static string IndentBlock(string text, string indent)
    {
        if (string.IsNullOrWhiteSpace(text))
        {
            return indent + "(空)";
        }

        return string.Join(Environment.NewLine, text
            .Replace("\r\n", "\n")
            .Split('\n')
            .Select(line => indent + line));
    }
}
