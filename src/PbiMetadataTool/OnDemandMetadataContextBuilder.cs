using System.Text;
using Microsoft.AnalysisServices.Tabular;

namespace PbiMetadataTool;

internal sealed class OnDemandMetadataContextBuilder
{
    private readonly PowerQueryMetadataReader _powerQueryReader = new();

    public string Build(string connectionString, string? databaseName, ModelMetadata modelSnapshot, string userPrompt, bool includeHiddenObjects, string? modelSourcePath)
    {
        if (string.IsNullOrWhiteSpace(connectionString) || string.IsNullOrWhiteSpace(userPrompt))
        {
            return string.Empty;
        }

        var intent = AnalyzeIntent(modelSnapshot, userPrompt);
        if (!intent.ShouldQuery)
        {
            return string.Empty;
        }

        var server = new Server();
        server.Connect(connectionString);

        try
        {
            var database = ResolveDatabase(server, databaseName);
            if (database is null)
            {
                return string.Empty;
            }

            var model = database.Model;
            var sb = new StringBuilder();
            sb.AppendLine("按需深度元数据（实时读取，未截断）：");

            if (intent.IncludeModel)
            {
                AppendModelSection(database, model, sb);
            }

            if (intent.IncludeDataSources)
            {
                AppendDataSourcesSection(model, sb);
                AppendPowerQueryQueriesSection(modelSnapshot, modelSourcePath, sb);
            }

            if (intent.IncludeExpressions)
            {
                AppendModelExpressionsSection(model, sb);
            }

            var targetTables = ResolveTargetTables(model, intent, includeHiddenObjects).ToList();
            if (targetTables.Count > 0)
            {
                sb.AppendLine();
                sb.AppendLine("表级深度信息:");
                foreach (var table in targetTables)
                {
                    AppendTableSection(table, includeHiddenObjects, sb);
                }
            }

            if (intent.IncludeRelationships)
            {
                AppendRelationshipsSection(model, intent.MentionedTables, sb);
            }

            if (intent.IncludeRoles)
            {
                AppendRolesSection(model, sb);
            }

            if (intent.IncludePerspectives)
            {
                AppendPerspectivesSection(model, sb);
            }

            if (intent.IncludeCultures)
            {
                AppendCulturesSection(model, sb);
            }

            if (intent.IncludeCalculationGroups && targetTables.Count == 0)
            {
                AppendCalculationGroupsSection(model, sb);
            }

            var text = sb.ToString().Trim();
            return string.Equals(text, "按需深度元数据（实时读取，未截断）：", StringComparison.Ordinal)
                ? string.Empty
                : text;
        }
        finally
        {
            if (server.Connected)
            {
                server.Disconnect();
            }
        }
    }

    private static QueryIntent AnalyzeIntent(ModelMetadata modelSnapshot, string prompt)
    {
        var text = prompt.Trim();
        var mentionedTables = FindMentionedTables(modelSnapshot, text);

        var includeAllTables = ContainsAny(
            text,
            "全部",
            "所有",
            "完整",
            "全量",
            "每张",
            "每个表",
            "all metadata",
            "full metadata",
            "each table",
            "per table",
            "tmdl");
        var includeModel = includeAllTables || ContainsAny(text, "模型", "model", "兼容级别", "compatibility", "annotation", "注解");
        var includeDataSources = includeAllTables || ContainsAny(text, "power query", "m代码", "m code", "数据来源", "数据源", "query", "source");
        var includeRelationships = includeAllTables || ContainsAny(text, "关系", "relationship", "cardinality", "cross filter", "security filtering");
        var includeRoles = includeAllTables || ContainsAny(text, "角色", "权限", "rls", "ols", "security");
        var includePerspectives = includeAllTables || ContainsAny(text, "perspective", "透视图");
        var includeCultures = includeAllTables || ContainsAny(text, "culture", "translation", "翻译", "本地化");
        var includeExpressions = includeAllTables || ContainsAny(text, "named expression", "model expression", "表达式");
        var includeCalculationGroups = includeAllTables || ContainsAny(text, "计算组", "calculation group", "calculation item");

        return new QueryIntent(
            IncludeModel: includeModel,
            IncludeDataSources: includeDataSources,
            IncludeRelationships: includeRelationships,
            IncludeRoles: includeRoles,
            IncludePerspectives: includePerspectives,
            IncludeCultures: includeCultures,
            IncludeExpressions: includeExpressions,
            IncludeCalculationGroups: includeCalculationGroups,
            IncludeAllTables: includeAllTables,
            MentionedTables: mentionedTables);
    }

    private static HashSet<string> FindMentionedTables(ModelMetadata modelSnapshot, string prompt)
    {
        var result = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var table in modelSnapshot.Tables)
        {
            if (table.Name.Length >= 2 && ContainsIgnoreCase(prompt, table.Name))
            {
                result.Add(table.Name);
                continue;
            }

            if (table.Measures.Any(measure => IsPromptMentioningObject(prompt, table.Name, measure.Name)))
            {
                result.Add(table.Name);
                continue;
            }

            if (table.Columns.Any(column => IsPromptMentioningObject(prompt, table.Name, column.Name)))
            {
                result.Add(table.Name);
            }
        }

        return result;
    }

    private static IEnumerable<Table> ResolveTargetTables(Model model, QueryIntent intent, bool includeHiddenObjects)
    {
        var tables = model.Tables.AsEnumerable();
        if (!includeHiddenObjects && !intent.IncludeAllTables)
        {
            tables = tables.Where(t => !t.IsHidden || intent.MentionedTables.Contains(t.Name));
        }

        if (intent.IncludeAllTables)
        {
            return tables.OrderBy(t => t.Name, StringComparer.OrdinalIgnoreCase);
        }

        if (intent.MentionedTables.Count > 0)
        {
            return tables
                .Where(t => intent.MentionedTables.Contains(t.Name))
                .OrderBy(t => t.Name, StringComparer.OrdinalIgnoreCase);
        }

        return Enumerable.Empty<Table>();
    }

    private static void AppendModelSection(Database database, Model model, StringBuilder sb)
    {
        sb.AppendLine($"数据库: {database.Name}");
        sb.AppendLine($"数据库ID: {database.ID}");
        sb.AppendLine($"兼容级别: {database.CompatibilityLevel}");
        AppendIfNotEmpty(sb, string.Empty, "默认存储模式", GetStringProperty(model, "DefaultMode"));
        AppendIfNotEmpty(sb, string.Empty, "禁用隐式度量值", GetStringProperty(model, "DiscourageImplicitMeasures"));
        AppendAnnotations(model, sb, string.Empty);
    }

    private static void AppendDataSourcesSection(Model model, StringBuilder sb)
    {
        var list = model.DataSources.ToList();
        sb.AppendLine();
        sb.AppendLine($"数据源 ({list.Count}):");
        foreach (var dataSource in list.OrderBy(d => d.Name, StringComparer.OrdinalIgnoreCase))
        {
            sb.AppendLine($"- 名称: {dataSource.Name}");
            AppendIfNotEmpty(sb, "  ", "类型", dataSource.Type.ToString());
            AppendIfNotEmpty(sb, "  ", "连接串", GetStringProperty(dataSource, "ConnectionString"));
            AppendIfNotEmpty(sb, "  ", "服务器", GetStringProperty(dataSource, "Server"));
            AppendIfNotEmpty(sb, "  ", "数据库", GetStringProperty(dataSource, "Database"));
            AppendIfNotEmpty(sb, "  ", "账号", GetStringProperty(dataSource, "Account"));
            AppendAnnotations(dataSource, sb, "  ");
        }
    }

    private void AppendPowerQueryQueriesSection(ModelMetadata modelSnapshot, string? modelSourcePath, StringBuilder sb)
    {
        var loadedTableNames = modelSnapshot.Tables
            .Where(table => string.Equals(table.SourceType, "PowerQuery", StringComparison.OrdinalIgnoreCase))
            .Select(table => table.Name)
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();

        var queries = _powerQueryReader.TryReadQueries(modelSourcePath, loadedTableNames);
        if (queries.Count == 0)
        {
            return;
        }

        var unloadedCount = queries.Count(query => !query.IsLoadedToModel);
        sb.AppendLine();
        sb.AppendLine($"Power Query 查询 ({queries.Count})：未加载 {unloadedCount}");
        foreach (var query in queries.OrderBy(q => q.Name, StringComparer.OrdinalIgnoreCase).Take(80))
        {
            sb.AppendLine($"- 名称: {query.Name}" + (query.IsLoadedToModel ? " | 已加载到模型" : " | 未加载到模型"));
            if (query.IsParameter)
            {
                sb.AppendLine("  类型: 参数查询");
            }
            else if (query.IsFunction)
            {
                sb.AppendLine("  类型: 函数查询");
            }

            AppendIfNotEmpty(sb, "  ", "来源系统", NormalizeUnknown(query.SourceSystemType));
            AppendIfNotEmpty(sb, "  ", "来源服务器", query.SourceServer);
            AppendIfNotEmpty(sb, "  ", "来源数据库", query.SourceDatabase);
            AppendIfNotEmpty(sb, "  ", "来源架构", query.SourceSchema);
            AppendIfNotEmpty(sb, "  ", "来源对象", query.SourceObjectName);
            AppendIfNotEmpty(sb, "  ", "M 代码", query.Expression);
        }
    }

    private static void AppendModelExpressionsSection(Model model, StringBuilder sb)
    {
        var expressions = GetEnumerableProperty(model, "Expressions").ToList();
        if (expressions.Count == 0)
        {
            return;
        }

        sb.AppendLine();
        sb.AppendLine($"模型表达式 ({expressions.Count}):");
        foreach (var expression in expressions.OrderBy(e => GetStringProperty(e, "Name"), StringComparer.OrdinalIgnoreCase))
        {
            var name = FirstNonEmpty(GetStringProperty(expression, "Name"), "(未命名)");
            sb.AppendLine($"- 名称: {name}");
            AppendIfNotEmpty(sb, "  ", "类型", GetStringProperty(expression, "Kind"));
            AppendIfNotEmpty(sb, "  ", "表达式", GetStringProperty(expression, "Expression"));
            AppendAnnotations(expression, sb, "  ");
        }
    }

    private static void AppendTableSection(Table table, bool includeHiddenObjects, StringBuilder sb)
    {
        var tableSourceInfo = ResolveTableSourceInfo(table);
        sb.AppendLine();
        sb.AppendLine($"[表] {table.Name}");
        AppendIfNotEmpty(sb, "  ", "隐藏", table.IsHidden ? "是" : "否");
        AppendIfNotEmpty(sb, "  ", "表类型", ResolvePartitionSourceType(table.Partitions.FirstOrDefault()?.Source));
        AppendIfNotEmpty(sb, "  ", "来源类型", tableSourceInfo.SourceType);
        AppendIfNotEmpty(sb, "  ", "数据源", tableSourceInfo.DataSourceName);
        AppendIfNotEmpty(sb, "  ", "来源系统", NormalizeUnknown(tableSourceInfo.Lineage.SystemType));
        AppendIfNotEmpty(sb, "  ", "来源服务器", tableSourceInfo.Lineage.Server);
        AppendIfNotEmpty(sb, "  ", "来源数据库", tableSourceInfo.Lineage.Database);
        AppendIfNotEmpty(sb, "  ", "来源架构", tableSourceInfo.Lineage.Schema);
        AppendIfNotEmpty(sb, "  ", "来源对象", tableSourceInfo.Lineage.ObjectName);
        AppendIfNotEmpty(sb, "  ", "描述", table.Description);
        AppendAnnotations(table, sb, "  ");

        AppendPartitions(table, sb);
        AppendColumns(table, includeHiddenObjects, sb);
        AppendMeasures(table, includeHiddenObjects, sb);
        AppendHierarchies(table, includeHiddenObjects, sb);
        AppendCalculationGroup(table, sb);
    }

    private static void AppendPartitions(Table table, StringBuilder sb)
    {
        var partitions = table.Partitions.ToList();
        sb.AppendLine($"  分区 ({partitions.Count}):");
        foreach (var partition in partitions.OrderBy(p => p.Name, StringComparer.OrdinalIgnoreCase))
        {
            var sourceType = ResolvePartitionSourceType(partition.Source);
            var dataSourceName = ResolvePartitionDataSourceName(partition.Source);
            var sourceExpression = ResolvePartitionSourceExpression(partition.Source);
            var lineage = QuerySourceParser.Parse(sourceType, sourceExpression, dataSourceName);

            sb.AppendLine($"  - 名称: {partition.Name}");
            AppendIfNotEmpty(sb, "    ", "模式", partition.Mode.ToString());
            AppendIfNotEmpty(sb, "    ", "来源类型", sourceType);
            AppendIfNotEmpty(sb, "    ", "数据源", dataSourceName);
            AppendIfNotEmpty(sb, "    ", "来源系统", NormalizeUnknown(lineage.SystemType));
            AppendIfNotEmpty(sb, "    ", "来源服务器", lineage.Server);
            AppendIfNotEmpty(sb, "    ", "来源数据库", lineage.Database);
            AppendIfNotEmpty(sb, "    ", "来源架构", lineage.Schema);
            AppendIfNotEmpty(sb, "    ", "来源对象", lineage.ObjectName);
            AppendIfNotEmpty(sb, "    ", "查询脚本", sourceExpression);
            AppendAnnotations(partition, sb, "    ");
        }
    }

    private static void AppendColumns(Table table, bool includeHiddenObjects, StringBuilder sb)
    {
        var columns = table.Columns
            .Where(c => includeHiddenObjects || !c.IsHidden)
            .OrderBy(c => c.Name, StringComparer.OrdinalIgnoreCase)
            .ToList();

        sb.AppendLine($"  列 ({columns.Count}):");
        foreach (var column in columns)
        {
            sb.AppendLine($"  - 名称: {column.Name}");
            AppendIfNotEmpty(sb, "    ", "列类型", column.Type.ToString());
            AppendIfNotEmpty(sb, "    ", "数据类型", column.DataType.ToString());
            AppendIfNotEmpty(sb, "    ", "隐藏", column.IsHidden ? "是" : "否");
            AppendIfNotEmpty(sb, "    ", "描述", column.Description);
            AppendIfNotEmpty(sb, "    ", "显示文件夹", GetStringProperty(column, "DisplayFolder"));
            AppendIfNotEmpty(sb, "    ", "格式", GetStringProperty(column, "FormatString"));
            AppendIfNotEmpty(sb, "    ", "汇总方式", GetStringProperty(column, "SummarizeBy"));
            AppendIfNotEmpty(sb, "    ", "数据分类", GetStringProperty(column, "DataCategory"));
            AppendIfNotEmpty(sb, "    ", "来源列", GetStringProperty(column, "SourceColumn"));
            AppendIfNotEmpty(sb, "    ", "是否键列", GetStringProperty(column, "IsKey"));
            AppendIfNotEmpty(sb, "    ", "是否可空", GetStringProperty(column, "IsNullable"));
            var sortByColumn = GetObjectProperty(column, "SortByColumn");
            AppendIfNotEmpty(sb, "    ", "排序列", GetStringProperty(sortByColumn, "Name"));

            var expression = column is CalculatedColumn calculatedColumn
                ? calculatedColumn.Expression
                : GetStringProperty(column, "Expression");
            AppendIfNotEmpty(sb, "    ", "表达式", expression);
            AppendAnnotations(column, sb, "    ");
        }
    }

    private static void AppendMeasures(Table table, bool includeHiddenObjects, StringBuilder sb)
    {
        var measures = table.Measures
            .Where(m => includeHiddenObjects || !m.IsHidden)
            .OrderBy(m => m.Name, StringComparer.OrdinalIgnoreCase)
            .ToList();

        sb.AppendLine($"  度量值 ({measures.Count}):");
        foreach (var measure in measures)
        {
            sb.AppendLine($"  - 名称: {measure.Name}");
            AppendIfNotEmpty(sb, "    ", "隐藏", measure.IsHidden ? "是" : "否");
            AppendIfNotEmpty(sb, "    ", "描述", measure.Description);
            AppendIfNotEmpty(sb, "    ", "显示文件夹", measure.DisplayFolder);
            AppendIfNotEmpty(sb, "    ", "格式", measure.FormatString);
            AppendIfNotEmpty(sb, "    ", "FormatStringExpression", GetStringProperty(measure, "FormatStringExpression"));
            AppendIfNotEmpty(sb, "    ", "DetailRowsExpression", GetStringProperty(measure, "DetailRowsExpression"));
            AppendIfNotEmpty(sb, "    ", "表达式", measure.Expression);
            AppendAnnotations(measure, sb, "    ");
        }
    }

    private static void AppendHierarchies(Table table, bool includeHiddenObjects, StringBuilder sb)
    {
        var hierarchies = table.Hierarchies
            .Where(h => includeHiddenObjects || !h.IsHidden)
            .OrderBy(h => h.Name, StringComparer.OrdinalIgnoreCase)
            .ToList();

        if (hierarchies.Count == 0)
        {
            return;
        }

        sb.AppendLine($"  层级 ({hierarchies.Count}):");
        foreach (var hierarchy in hierarchies)
        {
            sb.AppendLine($"  - 名称: {hierarchy.Name}");
            AppendIfNotEmpty(sb, "    ", "隐藏", hierarchy.IsHidden ? "是" : "否");
            AppendIfNotEmpty(sb, "    ", "描述", hierarchy.Description);

            var levels = hierarchy.Levels
                .OrderBy(l => l.Ordinal)
                .Select(l => $"{l.Ordinal}:{l.Name}({l.Column?.Name})")
                .ToList();
            if (levels.Count > 0)
            {
                sb.AppendLine($"    级别: {string.Join(" | ", levels)}");
            }
            AppendAnnotations(hierarchy, sb, "    ");
        }
    }

    private static void AppendCalculationGroup(Table table, StringBuilder sb)
    {
        var calcGroup = GetObjectProperty(table, "CalculationGroup");
        if (calcGroup is null)
        {
            return;
        }

        sb.AppendLine("  计算组:");
        AppendIfNotEmpty(sb, "    ", "Precedence", GetStringProperty(calcGroup, "Precedence"));
        AppendAnnotations(calcGroup, sb, "    ");

        var items = GetEnumerableProperty(calcGroup, "CalculationItems").ToList();
        sb.AppendLine($"    CalculationItems ({items.Count}):");
        foreach (var item in items.OrderBy(i => GetStringProperty(i, "Name"), StringComparer.OrdinalIgnoreCase))
        {
            sb.AppendLine($"    - 名称: {FirstNonEmpty(GetStringProperty(item, "Name"), "(未命名)")}");
            AppendIfNotEmpty(sb, "      ", "Ordinal", GetStringProperty(item, "Ordinal"));
            AppendIfNotEmpty(sb, "      ", "表达式", GetStringProperty(item, "Expression"));
            AppendIfNotEmpty(sb, "      ", "FormatStringExpression", FirstNonEmpty(
                GetStringProperty(item, "FormatStringExpression"),
                GetStringProperty(GetObjectProperty(item, "FormatStringDefinition"), "Expression")));
            AppendIfNotEmpty(sb, "      ", "NoSelectionExpression", GetStringProperty(item, "NoSelectionExpression"));
            AppendAnnotations(item, sb, "      ");
        }
    }

    private static void AppendRelationshipsSection(Model model, HashSet<string> mentionedTables, StringBuilder sb)
    {
        var relationships = model.Relationships.ToList();
        if (mentionedTables.Count > 0)
        {
            relationships = relationships
                .Where(r => RelationshipInvolvesAnyTable(r, mentionedTables))
                .ToList();
        }

        sb.AppendLine();
        sb.AppendLine($"关系 ({relationships.Count}):");
        foreach (var relationship in relationships.OrderBy(r => r.Name, StringComparer.OrdinalIgnoreCase))
        {
            sb.AppendLine($"- 名称: {relationship.Name}");
            AppendIfNotEmpty(sb, "  ", "类型", relationship.Type.ToString());
            AppendIfNotEmpty(sb, "  ", "激活", relationship.IsActive ? "是" : "否");
            AppendIfNotEmpty(sb, "  ", "交叉筛选", GetStringProperty(relationship, "CrossFilteringBehavior"));
            AppendIfNotEmpty(sb, "  ", "安全筛选", GetStringProperty(relationship, "SecurityFilteringBehavior"));
            AppendIfNotEmpty(sb, "  ", "FromCardinality", GetStringProperty(relationship, "FromCardinality"));
            AppendIfNotEmpty(sb, "  ", "ToCardinality", GetStringProperty(relationship, "ToCardinality"));
            AppendIfNotEmpty(sb, "  ", "JoinOnDateBehavior", GetStringProperty(relationship, "JoinOnDateBehavior"));
            AppendIfNotEmpty(sb, "  ", "RelyOnReferentialIntegrity", GetStringProperty(relationship, "RelyOnReferentialIntegrity"));

            if (relationship is SingleColumnRelationship single)
            {
                var from = $"{single.FromColumn?.Table?.Name}.{single.FromColumn?.Name}";
                var to = $"{single.ToColumn?.Table?.Name}.{single.ToColumn?.Name}";
                AppendIfNotEmpty(sb, "  ", "起点", from);
                AppendIfNotEmpty(sb, "  ", "终点", to);
            }
            AppendAnnotations(relationship, sb, "  ");
        }
    }

    private static void AppendRolesSection(Model model, StringBuilder sb)
    {
        var roles = model.Roles.ToList();
        sb.AppendLine();
        sb.AppendLine($"角色 ({roles.Count}):");
        foreach (var role in roles.OrderBy(r => r.Name, StringComparer.OrdinalIgnoreCase))
        {
            sb.AppendLine($"- 角色: {role.Name}");
            AppendIfNotEmpty(sb, "  ", "模型权限", role.ModelPermission.ToString());
            AppendIfNotEmpty(sb, "  ", "描述", role.Description);

            var members = GetEnumerableProperty(role, "Members").ToList();
            if (members.Count > 0)
            {
                sb.AppendLine($"  成员 ({members.Count}):");
                foreach (var member in members)
                {
                    var name = FirstNonEmpty(
                        GetStringProperty(member, "MemberName"),
                        GetStringProperty(member, "Name"),
                        GetStringProperty(member, "MemberID"));
                    sb.AppendLine($"  - {name}");
                    AppendIfNotEmpty(sb, "    ", "类型", GetStringProperty(member, "MemberType"));
                    AppendIfNotEmpty(sb, "    ", "身份源", GetStringProperty(member, "IdentityProvider"));
                }
            }

            var tablePermissions = GetEnumerableProperty(role, "TablePermissions").ToList();
            if (tablePermissions.Count > 0)
            {
                sb.AppendLine($"  表权限 ({tablePermissions.Count}):");
                foreach (var permission in tablePermissions)
                {
                    var tableObject = GetObjectProperty(permission, "Table");
                    var tableName = FirstNonEmpty(
                        GetStringProperty(tableObject, "Name"),
                        GetStringProperty(permission, "Name"),
                        "(未知表)");
                    sb.AppendLine($"  - 表: {tableName}");
                    AppendIfNotEmpty(sb, "    ", "MetadataPermission", GetStringProperty(permission, "MetadataPermission"));
                    AppendIfNotEmpty(sb, "    ", "FilterExpression", GetStringProperty(permission, "FilterExpression"));

                    var columnPermissions = GetEnumerableProperty(permission, "ColumnPermissions").ToList();
                    if (columnPermissions.Count > 0)
                    {
                        sb.AppendLine($"    列权限 ({columnPermissions.Count}):");
                        foreach (var cp in columnPermissions)
                        {
                            sb.AppendLine($"    - 列: {FirstNonEmpty(GetStringProperty(GetObjectProperty(cp, "Column"), "Name"), GetStringProperty(cp, "Name"))}");
                            AppendIfNotEmpty(sb, "      ", "Permission", GetStringProperty(cp, "Permission"));
                        }
                    }
                }
            }

            AppendAnnotations(role, sb, "  ");
        }
    }

    private static void AppendPerspectivesSection(Model model, StringBuilder sb)
    {
        var perspectives = GetEnumerableProperty(model, "Perspectives").ToList();
        if (perspectives.Count == 0)
        {
            return;
        }

        sb.AppendLine();
        sb.AppendLine($"透视图 ({perspectives.Count}):");
        foreach (var perspective in perspectives.OrderBy(p => GetStringProperty(p, "Name"), StringComparer.OrdinalIgnoreCase))
        {
            var name = FirstNonEmpty(GetStringProperty(perspective, "Name"), "(未命名)");
            sb.AppendLine($"- 名称: {name}");
            var perspectiveTables = GetEnumerableProperty(perspective, "PerspectiveTables")
                .Select(pt => FirstNonEmpty(GetStringProperty(GetObjectProperty(pt, "Table"), "Name"), GetStringProperty(pt, "Name")))
                .Where(n => !string.IsNullOrWhiteSpace(n))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .OrderBy(n => n, StringComparer.OrdinalIgnoreCase)
                .ToList();
            if (perspectiveTables.Count > 0)
            {
                sb.AppendLine($"  表: {string.Join(", ", perspectiveTables)}");
            }
            AppendAnnotations(perspective, sb, "  ");
        }
    }

    private static void AppendCulturesSection(Model model, StringBuilder sb)
    {
        var cultures = GetEnumerableProperty(model, "Cultures").ToList();
        if (cultures.Count == 0)
        {
            return;
        }

        sb.AppendLine();
        sb.AppendLine($"Culture/翻译 ({cultures.Count}):");
        foreach (var culture in cultures.OrderBy(c => GetStringProperty(c, "Name"), StringComparer.OrdinalIgnoreCase))
        {
            var name = FirstNonEmpty(GetStringProperty(culture, "Name"), "(未命名)");
            sb.AppendLine($"- 名称: {name}");
            AppendIfNotEmpty(sb, "  ", "Description", GetStringProperty(culture, "Description"));
            var translations = GetEnumerableProperty(culture, "ObjectTranslations").ToList();
            if (translations.Count > 0)
            {
                sb.AppendLine($"  ObjectTranslations: {translations.Count}");
            }
            AppendAnnotations(culture, sb, "  ");
        }
    }

    private static void AppendCalculationGroupsSection(Model model, StringBuilder sb)
    {
        var calcGroupTables = model.Tables
            .Where(t => GetObjectProperty(t, "CalculationGroup") is not null)
            .OrderBy(t => t.Name, StringComparer.OrdinalIgnoreCase)
            .ToList();

        if (calcGroupTables.Count == 0)
        {
            return;
        }

        sb.AppendLine();
        sb.AppendLine($"计算组表 ({calcGroupTables.Count}):");
        foreach (var table in calcGroupTables)
        {
            AppendTableSection(table, includeHiddenObjects: true, sb);
        }
    }

    private static bool RelationshipInvolvesAnyTable(Relationship relationship, HashSet<string> tableNames)
    {
        if (tableNames.Count == 0)
        {
            return true;
        }

        if (relationship is SingleColumnRelationship single)
        {
            var fromTable = single.FromColumn?.Table?.Name ?? string.Empty;
            var toTable = single.ToColumn?.Table?.Name ?? string.Empty;
            return tableNames.Contains(fromTable) || tableNames.Contains(toTable);
        }

        var reflectionFrom = FirstNonEmpty(
            GetStringProperty(GetObjectProperty(relationship, "FromTable"), "Name"),
            GetStringProperty(relationship, "FromTable"));
        var reflectionTo = FirstNonEmpty(
            GetStringProperty(GetObjectProperty(relationship, "ToTable"), "Name"),
            GetStringProperty(relationship, "ToTable"));
        return tableNames.Contains(reflectionFrom) || tableNames.Contains(reflectionTo);
    }

    private static Database? ResolveDatabase(Server server, string? databaseName)
    {
        if (!string.IsNullOrWhiteSpace(databaseName))
        {
            var db = server.Databases
                .OfType<Database>()
                .FirstOrDefault(d => string.Equals(d.Name, databaseName, StringComparison.OrdinalIgnoreCase));
            if (db is not null)
            {
                return db;
            }
        }

        return server.Databases
            .OfType<Database>()
            .FirstOrDefault();
    }

    private static void AppendIfNotEmpty(StringBuilder sb, string indent, string label, string? value)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            return;
        }

        sb.AppendLine($"{indent}{label}: {value}");
    }

    private static void AppendAnnotations(object target, StringBuilder sb, string indent)
    {
        var annotations = GetEnumerableProperty(target, "Annotations").ToList();
        if (annotations.Count == 0)
        {
            return;
        }

        sb.AppendLine($"{indent}注解 ({annotations.Count}):");
        foreach (var annotation in annotations.OrderBy(a => GetStringProperty(a, "Name"), StringComparer.OrdinalIgnoreCase))
        {
            var name = FirstNonEmpty(GetStringProperty(annotation, "Name"), "(未命名)");
            var value = FirstNonEmpty(GetStringProperty(annotation, "Value"), GetStringProperty(annotation, "Expression"));
            sb.AppendLine($"{indent}- {name}: {value}");
        }
    }

    private static IEnumerable<object> GetEnumerableProperty(object target, string propertyName)
    {
        var prop = target.GetType().GetProperty(propertyName);
        if (prop is null)
        {
            return Enumerable.Empty<object>();
        }

        var value = prop.GetValue(target);
        if (value is not System.Collections.IEnumerable enumerable)
        {
            return Enumerable.Empty<object>();
        }

        var list = new List<object>();
        foreach (var item in enumerable)
        {
            if (item is not null)
            {
                list.Add(item);
            }
        }

        return list;
    }

    private static object? GetObjectProperty(object? target, string propertyName)
    {
        if (target is null)
        {
            return null;
        }

        var prop = target.GetType().GetProperty(propertyName);
        return prop?.GetValue(target);
    }

    private static string GetStringProperty(object? target, string propertyName)
    {
        if (target is null)
        {
            return string.Empty;
        }

        var prop = target.GetType().GetProperty(propertyName);
        if (prop is null)
        {
            return string.Empty;
        }

        var value = prop.GetValue(target);
        return value?.ToString() ?? string.Empty;
    }

    private static string FirstNonEmpty(params string[] values)
    {
        foreach (var value in values)
        {
            if (!string.IsNullOrWhiteSpace(value))
            {
                return value;
            }
        }

        return string.Empty;
    }

    private static string NormalizeUnknown(string? value)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            return string.Empty;
        }

        return string.Equals(value, "Unknown", StringComparison.OrdinalIgnoreCase)
            ? string.Empty
            : value;
    }

    private static bool ContainsAny(string input, params string[] keywords)
        => keywords.Any(keyword => ContainsIgnoreCase(input, keyword));

    private static bool ContainsIgnoreCase(string text, string value)
        => text.Contains(value, StringComparison.OrdinalIgnoreCase);

    private static bool IsPromptMentioningObject(string prompt, string tableName, string objectName)
    {
        if (string.IsNullOrWhiteSpace(objectName))
        {
            return false;
        }

        if (ContainsIgnoreCase(prompt, $"{tableName}[{objectName}]"))
        {
            return true;
        }

        if (ContainsIgnoreCase(prompt, $"[{objectName}]"))
        {
            return true;
        }

        return objectName.Length >= 2 && ContainsIgnoreCase(prompt, objectName);
    }

    private static (string SourceType, string DataSourceName, string SourceExpression, SourceLineageInfo Lineage) ResolveTableSourceInfo(Table table)
    {
        if (table.Partitions.Count == 0)
        {
            var emptyLineage = QuerySourceParser.Parse(string.Empty, string.Empty, string.Empty);
            return (string.Empty, string.Empty, string.Empty, emptyLineage);
        }

        var sourceType = string.Empty;
        var dataSourceName = string.Empty;
        var sourceExpression = string.Empty;

        foreach (var partition in table.Partitions)
        {
            if (string.IsNullOrWhiteSpace(sourceType))
            {
                sourceType = ResolvePartitionSourceType(partition.Source);
            }

            if (string.IsNullOrWhiteSpace(dataSourceName))
            {
                dataSourceName = ResolvePartitionDataSourceName(partition.Source);
            }

            if (string.IsNullOrWhiteSpace(sourceExpression))
            {
                sourceExpression = ResolvePartitionSourceExpression(partition.Source);
            }

            if (!string.IsNullOrWhiteSpace(dataSourceName) && !string.IsNullOrWhiteSpace(sourceExpression))
            {
                break;
            }
        }

        return (sourceType, dataSourceName, sourceExpression, QuerySourceParser.Parse(sourceType, sourceExpression, dataSourceName));
    }

    private static string ResolvePartitionSourceType(PartitionSource? source)
    {
        return source switch
        {
            CalculatedPartitionSource => "Calculated",
            MPartitionSource => "PowerQuery",
            QueryPartitionSource => "Query",
            EntityPartitionSource => "Entity",
            null => string.Empty,
            _ => source.GetType().Name
        };
    }

    private static string ResolvePartitionSourceExpression(PartitionSource? source)
    {
        return source switch
        {
            CalculatedPartitionSource calculatedSource => calculatedSource.Expression ?? string.Empty,
            MPartitionSource mSource => mSource.Expression ?? string.Empty,
            QueryPartitionSource qSource => qSource.Query ?? string.Empty,
            null => string.Empty,
            _ => FirstNonEmpty(GetStringProperty(source, "Expression"), GetStringProperty(source, "Query"))
        };
    }

    private static string ResolvePartitionDataSourceName(PartitionSource? source)
    {
        if (source is null)
        {
            return string.Empty;
        }

        object? dataSource = source switch
        {
            QueryPartitionSource qSource => qSource.DataSource,
            _ => GetObjectProperty(source, "DataSource")
        };

        return FirstNonEmpty(
            GetStringProperty(dataSource, "Name"),
            GetStringProperty(dataSource, "ID"));
    }

    private sealed record QueryIntent(
        bool IncludeModel,
        bool IncludeDataSources,
        bool IncludeRelationships,
        bool IncludeRoles,
        bool IncludePerspectives,
        bool IncludeCultures,
        bool IncludeExpressions,
        bool IncludeCalculationGroups,
        bool IncludeAllTables,
        HashSet<string> MentionedTables)
    {
        public bool ShouldQuery =>
            IncludeModel ||
            IncludeDataSources ||
            IncludeRelationships ||
            IncludeRoles ||
            IncludePerspectives ||
            IncludeCultures ||
            IncludeExpressions ||
            IncludeCalculationGroups ||
            IncludeAllTables ||
            MentionedTables.Count > 0;
    }
}
