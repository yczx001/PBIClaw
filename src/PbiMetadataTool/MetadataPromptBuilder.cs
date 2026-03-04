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
- delete_table（可删除模型中的任意表，包括 Import/PowerQuery/Calculated）
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
- `create_calculated_table` 必须同时提供 `name` 和完整 `expression`（DAX），禁止留空或只写占位文字。
- 当用户要求“创建日期表/日历表”时，你必须自行生成可执行的日期表 DAX（例如基于 CALENDARAUTO），不要反问用户提供表达式。

关系定位可用 name，或用 fromTable/fromColumn/toTable/toColumn。
删除表时，不要因为来源类型是 PowerQuery 就拒绝；只要该表在模型中存在，即可使用 delete_table。
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

    public static string BuildModelContextForPrompt(
        ModelMetadata metadata,
        bool includeHiddenObjects,
        string userPrompt,
        ReportMetadata? report = null)
    {
        return ShouldUseDetailedModelContext(userPrompt)
            ? BuildModelContext(metadata, includeHiddenObjects, report)
            : BuildModelIndexContext(metadata, includeHiddenObjects, report);
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

            var tableType = string.IsNullOrWhiteSpace(table.TableType) ? "Unknown" : table.TableType;
            sb.AppendLine($"[表] {table.Name} | 表类型: {tableType}" + (table.IsHidden ? " (Hidden)" : string.Empty));
            if (!string.IsNullOrWhiteSpace(table.SourceType))
            {
                sb.AppendLine($"  来源类型: {table.SourceType}");
            }
            if (!string.IsNullOrWhiteSpace(table.DataSourceName))
            {
                sb.AppendLine($"  数据源: {table.DataSourceName}");
            }
            if (HasUsableValue(table.SourceSystemType))
            {
                sb.AppendLine($"  来源系统: {table.SourceSystemType}");
            }
            if (!string.IsNullOrWhiteSpace(table.SourceServer))
            {
                sb.AppendLine($"  来源服务器: {table.SourceServer}");
            }
            if (!string.IsNullOrWhiteSpace(table.SourceDatabase))
            {
                sb.AppendLine($"  来源数据库: {table.SourceDatabase}");
            }
            if (!string.IsNullOrWhiteSpace(table.SourceSchema))
            {
                sb.AppendLine($"  来源架构: {table.SourceSchema}");
            }
            if (!string.IsNullOrWhiteSpace(table.SourceObjectName))
            {
                sb.AppendLine($"  来源对象: {table.SourceObjectName}");
            }

            if (string.Equals(tableType, "Calculated", StringComparison.OrdinalIgnoreCase) &&
                !string.IsNullOrWhiteSpace(table.Expression))
            {
                sb.AppendLine("  计算表 DAX:（按需查询时提供完整表达式）");
            }

            var columns = table.Columns
                .Where(c => includeHiddenObjects || !c.IsHidden)
                .OrderBy(c => c.Name, StringComparer.OrdinalIgnoreCase)
                .ToList();

            foreach (var column in columns)
            {
                sb.AppendLine($"  - 列: {column.Name} | 类型: {column.DataType} | 列类型: {column.ColumnType}" + (column.IsHidden ? " | Hidden" : string.Empty));

                if (string.Equals(column.ColumnType, "Calculated", StringComparison.OrdinalIgnoreCase) &&
                    !string.IsNullOrWhiteSpace(column.Expression))
                {
                    sb.AppendLine("    计算列 DAX:（按需查询时提供完整表达式）");
                }
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

    public static string BuildModelIndexContext(ModelMetadata metadata, bool includeHiddenObjects, ReportMetadata? report = null)
    {
        var sb = new StringBuilder();
        sb.AppendLine("以下是当前 Power BI 模型索引（精简上下文）：");
        sb.AppendLine($"数据库: {metadata.DatabaseName}");
        sb.AppendLine($"兼容级别: {metadata.CompatibilityLevel}");
        sb.AppendLine($"表数量: {metadata.Tables.Count}");
        sb.AppendLine($"关系数量: {metadata.Relationships.Count}");
        sb.AppendLine($"角色数量: {metadata.Roles.Count}");
        sb.AppendLine();

        var tables = metadata.Tables
            .Where(t => includeHiddenObjects || !t.IsHidden)
            .OrderBy(t => t.Name, StringComparer.OrdinalIgnoreCase)
            .ToList();
        sb.AppendLine($"表索引 ({tables.Count}):");
        foreach (var table in tables.Take(220))
        {
            var columnCount = table.Columns.Count(c => includeHiddenObjects || !c.IsHidden);
            var measureCount = table.Measures.Count(m => includeHiddenObjects || !m.IsHidden);
            var type = string.IsNullOrWhiteSpace(table.TableType) ? "-" : table.TableType;
            sb.AppendLine($"- {table.Name} | 类型: {type} | 列: {columnCount} | 度量值: {measureCount}");
        }
        if (tables.Count > 220)
        {
            sb.AppendLine($"- ...(其余 {tables.Count - 220} 张表已省略)");
        }

        var relationships = metadata.Relationships
            .OrderBy(r => r.Name, StringComparer.OrdinalIgnoreCase)
            .ToList();
        if (relationships.Count > 0)
        {
            sb.AppendLine();
            sb.AppendLine($"关系索引 ({relationships.Count}):");
            foreach (var relationship in relationships.Take(120))
            {
                sb.AppendLine($"- {relationship.FromTable}.{relationship.FromColumn} -> {relationship.ToTable}.{relationship.ToColumn}");
            }
            if (relationships.Count > 120)
            {
                sb.AppendLine($"- ...(其余 {relationships.Count - 120} 条关系已省略)");
            }
        }

        if (metadata.Roles.Count > 0)
        {
            sb.AppendLine();
            sb.AppendLine("角色索引:");
            foreach (var role in metadata.Roles.OrderBy(r => r.Name, StringComparer.OrdinalIgnoreCase).Take(60))
            {
                sb.AppendLine($"- {role.Name}");
            }
            if (metadata.Roles.Count > 60)
            {
                sb.AppendLine($"- ...(其余 {metadata.Roles.Count - 60} 个角色已省略)");
            }
        }

        if (report is not null)
        {
            sb.AppendLine();
            sb.AppendLine($"报表页数量: {report.Pages.Count}");
            foreach (var page in report.Pages.Take(20))
            {
                var pageName = string.IsNullOrWhiteSpace(page.DisplayName) ? page.Name : page.DisplayName;
                sb.AppendLine($"- 页面: {pageName} | 视觉对象: {page.Visuals.Count}");
            }
            if (report.Pages.Count > 20)
            {
                sb.AppendLine($"- ...(其余 {report.Pages.Count - 20} 页已省略)");
            }
        }

        sb.AppendLine();
        sb.AppendLine("若需要完整字段/度量值/DAX/M 代码，请按对象名继续追问。");
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

            sb.AppendLine("  DAX:");
            sb.AppendLine(IndentBlock(hit.Measure.Expression ?? string.Empty, "    "));
        }

        return sb.ToString().TrimEnd();
    }

    public static string BuildReferencedCalculatedContext(ModelMetadata metadata, string userPrompt, bool includeHiddenObjects)
    {
        if (string.IsNullOrWhiteSpace(userPrompt))
        {
            return string.Empty;
        }

        var prompt = userPrompt.Trim();
        var rows = new List<(string Kind, string Name, string Expression)>();

        foreach (var table in metadata.Tables)
        {
            if (!includeHiddenObjects && table.IsHidden)
            {
                continue;
            }

            if (!string.IsNullOrWhiteSpace(table.Expression) &&
                IsPromptMentioningObject(prompt, table.Name))
            {
                rows.Add(("计算表", table.Name, table.Expression));
            }

            foreach (var column in table.Columns)
            {
                if (!includeHiddenObjects && column.IsHidden)
                {
                    continue;
                }

                if (!string.Equals(column.ColumnType, "Calculated", StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                if (!IsPromptMentioningObject(prompt, table.Name, column.Name))
                {
                    continue;
                }

                if (!string.IsNullOrWhiteSpace(column.Expression))
                {
                    rows.Add(("计算列", $"{table.Name}[{column.Name}]", column.Expression));
                }
            }
        }

        if (rows.Count == 0)
        {
            return string.Empty;
        }

        var sb = new StringBuilder();
        sb.AppendLine("用户问题中疑似提到了以下计算对象，回答时请基于它们的真实 DAX：");
        foreach (var row in rows.Take(8))
        {
            sb.AppendLine($"- {row.Kind}: {row.Name}");
            sb.AppendLine("  DAX:");
            sb.AppendLine(IndentBlock(row.Expression, "    "));
        }

        return sb.ToString().TrimEnd();
    }

    public static string BuildReferencedTableSourceContext(ModelMetadata metadata, string userPrompt, bool includeHiddenObjects)
    {
        if (string.IsNullOrWhiteSpace(userPrompt))
        {
            return string.Empty;
        }

        var prompt = userPrompt.Trim();
        var rows = new List<(
            string Table,
            string SourceType,
            string DataSourceName,
            string SourceSystemType,
            string SourceServer,
            string SourceDatabase,
            string SourceSchema,
            string SourceObjectName,
            string SourceExpression)>();

        foreach (var table in metadata.Tables)
        {
            if (!includeHiddenObjects && table.IsHidden)
            {
                continue;
            }

            if (string.IsNullOrWhiteSpace(table.SourceExpression))
            {
                continue;
            }

            if (!IsPromptMentioningObject(prompt, table.Name))
            {
                continue;
            }

            rows.Add((
                table.Name,
                table.SourceType,
                table.DataSourceName,
                table.SourceSystemType,
                table.SourceServer,
                table.SourceDatabase,
                table.SourceSchema,
                table.SourceObjectName,
                table.SourceExpression));
        }

        if (rows.Count == 0)
        {
            return string.Empty;
        }

        var sb = new StringBuilder();
        sb.AppendLine("用户问题中疑似提到了以下表的数据来源，请基于真实查询脚本回答：");
        foreach (var row in rows.Take(8))
        {
            sb.AppendLine($"- 表: {row.Table}");
            if (!string.IsNullOrWhiteSpace(row.SourceType))
            {
                sb.AppendLine($"  来源类型: {row.SourceType}");
            }
            if (!string.IsNullOrWhiteSpace(row.DataSourceName))
            {
                sb.AppendLine($"  数据源: {row.DataSourceName}");
            }
            if (HasUsableValue(row.SourceSystemType))
            {
                sb.AppendLine($"  来源系统: {row.SourceSystemType}");
            }
            if (!string.IsNullOrWhiteSpace(row.SourceServer))
            {
                sb.AppendLine($"  来源服务器: {row.SourceServer}");
            }
            if (!string.IsNullOrWhiteSpace(row.SourceDatabase))
            {
                sb.AppendLine($"  来源数据库: {row.SourceDatabase}");
            }
            if (!string.IsNullOrWhiteSpace(row.SourceSchema))
            {
                sb.AppendLine($"  来源架构: {row.SourceSchema}");
            }
            if (!string.IsNullOrWhiteSpace(row.SourceObjectName))
            {
                sb.AppendLine($"  来源对象: {row.SourceObjectName}");
            }
            sb.AppendLine("  查询脚本:");
            sb.AppendLine(IndentBlock(row.SourceExpression, "    "));
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

    private static bool IsPromptMentioningObject(string prompt, string tableName, string? columnOrMeasureName = null)
    {
        if (string.IsNullOrWhiteSpace(tableName))
        {
            return false;
        }

        if (!string.IsNullOrWhiteSpace(columnOrMeasureName))
        {
            if (prompt.Contains($"{tableName}[{columnOrMeasureName}]", StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }

            if (prompt.Contains($"[{columnOrMeasureName}]", StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }

            if (columnOrMeasureName.Length >= 2 &&
                prompt.Contains(columnOrMeasureName, StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }
        }

        return tableName.Length >= 2 &&
               prompt.Contains(tableName, StringComparison.OrdinalIgnoreCase);
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

    private static bool HasUsableValue(string? value)
        => !string.IsNullOrWhiteSpace(value) &&
           !string.Equals(value, "Unknown", StringComparison.OrdinalIgnoreCase);

    private static bool ShouldUseDetailedModelContext(string prompt)
    {
        if (string.IsNullOrWhiteSpace(prompt))
        {
            return false;
        }

        return ContainsAnyCaseInsensitive(
            prompt,
            "全部",
            "所有",
            "完整",
            "全量",
            "每张",
            "每个",
            "全模型",
            "模型全貌",
            "all metadata",
            "full metadata",
            "entire model",
            "all tables",
            "all columns",
            "all measures",
            "all relationships");
    }

    private static bool ContainsAnyCaseInsensitive(string text, params string[] keywords)
        => keywords.Any(keyword => text.Contains(keyword, StringComparison.OrdinalIgnoreCase));
}
