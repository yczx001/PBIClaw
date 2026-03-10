using Microsoft.AnalysisServices.Tabular;

namespace PbiMetadataTool;

internal sealed record TabularActionAnalysis(
    IReadOnlyList<string> Infos,
    IReadOnlyList<string> Warnings,
    IReadOnlyList<string> Errors)
{
    public bool HasErrors => Errors.Count > 0;
}

internal sealed class TabularModelWriter
{
    public TabularActionAnalysis AnalyzeActions(int port, string? databaseName, AbiActionPlan plan)
        => AnalyzeActions($"DataSource=localhost:{port};", databaseName, plan);

    public TabularActionAnalysis AnalyzeActions(string connectionString, string? databaseName, AbiActionPlan plan)
    {
        if (plan.Actions.Count == 0)
        {
            return new TabularActionAnalysis(
                ["未检测到需要执行的动作。"],
                [],
                []);
        }

        var infos = new List<string>();
        var warnings = new List<string>();
        var errors = new List<string>();

        var server = new Server();
        server.Connect(connectionString);

        try
        {
            var database = ResolveDatabase(server, databaseName);
            if (database is null)
            {
                return new TabularActionAnalysis([], [], ["未找到可写入的 Tabular 数据库。"]);
            }

            foreach (var action in plan.Actions)
            {
                AnalyzeAction(database, action, infos, warnings, errors);
            }

            if (errors.Count == 0)
            {
                infos.Add($"预检完成：{plan.Actions.Count} 项动作可进入执行确认。");
            }

            return new TabularActionAnalysis(infos, warnings, errors);
        }
        finally
        {
            if (server.Connected)
            {
                server.Disconnect();
            }
        }
    }

    public IReadOnlyList<string> ApplyActions(int port, string? databaseName, AbiActionPlan plan)
        => ApplyActions($"DataSource=localhost:{port};", databaseName, plan);

    public IReadOnlyList<string> ApplyActions(string connectionString, string? databaseName, AbiActionPlan plan)
    {
        if (plan.Actions.Count == 0)
        {
            return ["未检测到需要执行的动作。"];
        }

        var results = new List<string>();
        var server = new Server();
        server.Connect(connectionString);

        try
        {
            var database = ResolveDatabase(server, databaseName);
            if (database is null)
            {
                throw new InvalidOperationException("未找到可写入的 Tabular 数据库。");
            }

            var changed = false;
            foreach (var action in plan.Actions)
            {
                changed |= ApplyAction(database, action, results);
            }

            if (changed)
            {
                database.Model.SaveChanges();
                results.Add("已提交模型变更。");
            }
            else
            {
                results.Add("没有需要提交的实际变更。");
            }

            return results;
        }
        finally
        {
            if (server.Connected)
            {
                server.Disconnect();
            }
        }
    }

    private static void AnalyzeAction(Database database, AbiModelAction action, ICollection<string> infos, ICollection<string> warnings, ICollection<string> errors)
    {
        try
        {
            var type = action.Type.Trim().ToLowerInvariant();
            switch (type)
            {
                case "create_or_update_measure":
                    AnalyzeCreateOrUpdateMeasure(database, action, infos, warnings, errors);
                    break;
                case "delete_measure":
                    AnalyzeDeleteMeasure(database, action, infos, warnings, errors);
                    break;
                case "create_relationship":
                    AnalyzeCreateRelationship(database, action, infos, warnings, errors);
                    break;
                case "delete_relationship":
                    AnalyzeDeleteRelationship(database, action, infos, warnings, errors);
                    break;
                default:
                    if (!TabularExtendedActionHandler.TryAnalyze(database, action, infos, warnings, errors))
                    {
                        errors.Add($"不支持的动作类型: {action.Type}");
                    }
                    break;
            }
        }
        catch (Exception ex)
        {
            errors.Add($"动作预检失败 [{action.Type}]：{ex.Message}");
        }
    }

    private static void AnalyzeCreateOrUpdateMeasure(Database database, AbiModelAction action, ICollection<string> infos, ICollection<string> warnings, ICollection<string> errors)
    {
        var table = TryFindTable(database, action.Table);
        if (table is null)
        {
            errors.Add($"表不存在: {action.Table}");
            return;
        }

        var name = action.Name?.Trim();
        if (string.IsNullOrWhiteSpace(name))
        {
            errors.Add("create_or_update_measure 缺少 name");
            return;
        }

        if (string.IsNullOrWhiteSpace(action.Expression))
        {
            errors.Add($"度量值 {table.Name}.{name} 缺少 expression");
            return;
        }

        var existing = table.Measures.FirstOrDefault(m => string.Equals(m.Name, name, StringComparison.OrdinalIgnoreCase));
        if (existing is null)
        {
            infos.Add($"将新增度量值: {table.Name}.{name}");
            return;
        }

        warnings.Add($"将覆盖已有度量值: {table.Name}.{name}");

        if (string.Equals(existing.Expression?.Trim(), action.Expression?.Trim(), StringComparison.Ordinal))
        {
            infos.Add($"度量值表达式与现有一致: {table.Name}.{name}");
        }

        if (!string.IsNullOrWhiteSpace(action.FormatString) &&
            !string.Equals(existing.FormatString, action.FormatString, StringComparison.Ordinal))
        {
            warnings.Add($"将更新格式字符串: {table.Name}.{name}");
        }
    }

    private static void AnalyzeDeleteMeasure(Database database, AbiModelAction action, ICollection<string> infos, ICollection<string> warnings, ICollection<string> errors)
    {
        var table = TryFindTable(database, action.Table);
        if (table is null)
        {
            errors.Add($"表不存在: {action.Table}");
            return;
        }

        var name = action.Name?.Trim();
        if (string.IsNullOrWhiteSpace(name))
        {
            errors.Add("delete_measure 缺少 name");
            return;
        }

        var existing = table.Measures.FirstOrDefault(m => string.Equals(m.Name, name, StringComparison.OrdinalIgnoreCase));
        if (existing is null)
        {
            warnings.Add($"待删除度量值不存在，将跳过: {table.Name}.{name}");
            return;
        }

        warnings.Add($"将删除度量值: {table.Name}.{name}");
        infos.Add($"删除动作已通过预检: {table.Name}.{name}");
    }

    private static void AnalyzeCreateRelationship(Database database, AbiModelAction action, ICollection<string> infos, ICollection<string> warnings, ICollection<string> errors)
    {
        var fromTable = TryFindTable(database, action.FromTable);
        var toTable = TryFindTable(database, action.ToTable);
        if (fromTable is null)
        {
            errors.Add($"FromTable 不存在: {action.FromTable}");
            return;
        }

        if (toTable is null)
        {
            errors.Add($"ToTable 不存在: {action.ToTable}");
            return;
        }

        var fromColumnName = action.FromColumn?.Trim();
        var toColumnName = action.ToColumn?.Trim();
        if (string.IsNullOrWhiteSpace(fromColumnName) || string.IsNullOrWhiteSpace(toColumnName))
        {
            errors.Add("create_relationship 缺少 fromColumn 或 toColumn");
            return;
        }

        var fromColumn = fromTable.Columns.FirstOrDefault(c => string.Equals(c.Name, fromColumnName, StringComparison.OrdinalIgnoreCase));
        var toColumn = toTable.Columns.FirstOrDefault(c => string.Equals(c.Name, toColumnName, StringComparison.OrdinalIgnoreCase));
        if (fromColumn is null)
        {
            errors.Add($"列不存在: {fromTable.Name}.{fromColumnName}");
            return;
        }

        if (toColumn is null)
        {
            errors.Add($"列不存在: {toTable.Name}.{toColumnName}");
            return;
        }

        var existingByName = TryFindRelationshipByName(database, action.Name);
        var existingByEndpoints = TryFindRelationshipByEndpoints(database, fromTable.Name, fromColumnName, toTable.Name, toColumnName);

        if (existingByName is not null)
        {
            if (!IsSameEndpoints(existingByName, fromTable.Name, fromColumnName, toTable.Name, toColumnName))
            {
                warnings.Add($"将重定向已有关系: {existingByName.Name}");
            }

            infos.Add($"将更新关系: {existingByName.Name}");
            return;
        }

        if (existingByEndpoints is not null)
        {
            warnings.Add($"相同端点关系已存在: {existingByEndpoints.Name}，将更新现有关系而非新增。");
            return;
        }

        infos.Add($"将新增关系: {fromTable.Name}.{fromColumnName} -> {toTable.Name}.{toColumnName}");
    }

    private static void AnalyzeDeleteRelationship(Database database, AbiModelAction action, ICollection<string> infos, ICollection<string> warnings, ICollection<string> errors)
    {
        var existing = FindRelationship(database, action);
        if (existing is null)
        {
            warnings.Add("待删除关系不存在，将跳过。") ;
            return;
        }

        warnings.Add($"将删除关系: {existing.Name}");
        infos.Add($"删除动作已通过预检: {existing.Name}");
    }

    private static bool ApplyAction(Database database, AbiModelAction action, ICollection<string> results)
    {
        var type = action.Type.Trim().ToLowerInvariant();
        return type switch
        {
            "create_or_update_measure" => ApplyCreateOrUpdateMeasure(database, action, results),
            "delete_measure" => ApplyDeleteMeasure(database, action, results),
            "create_relationship" => ApplyCreateRelationship(database, action, results),
            "delete_relationship" => ApplyDeleteRelationship(database, action, results),
            _ => TabularExtendedActionHandler.TryApply(database, action, results)
                ? true
                : HandleUnsupportedAction(type, results)
        };
    }

    private static bool ApplyCreateOrUpdateMeasure(Database database, AbiModelAction action, ICollection<string> results)
    {
        var table = FindTable(database, action.Table);
        var measureName = RequireValue(action.Name, "name");
        var expression = RequireValue(action.Expression, "expression");

        var existing = table.Measures.FirstOrDefault(m => string.Equals(m.Name, measureName, StringComparison.OrdinalIgnoreCase));
        if (existing is null)
        {
            var created = new Measure
            {
                Name = measureName,
                Expression = expression
            };
            if (!string.IsNullOrWhiteSpace(action.FormatString))
            {
                created.FormatString = action.FormatString;
            }

            if (action.IsHidden.HasValue)
            {
                created.IsHidden = action.IsHidden.Value;
            }

            table.Measures.Add(created);
            results.Add($"在表 [{table.Name}] 下创建度量值 [{measureName}]");
            return true;
        }

        existing.Expression = expression;
        if (!string.IsNullOrWhiteSpace(action.FormatString))
        {
            existing.FormatString = action.FormatString;
        }

        if (action.IsHidden.HasValue)
        {
            existing.IsHidden = action.IsHidden.Value;
        }

        results.Add($"在表 [{table.Name}] 下更新度量值 [{measureName}]");
        return true;
    }

    private static bool ApplyDeleteMeasure(Database database, AbiModelAction action, ICollection<string> results)
    {
        var table = FindTable(database, action.Table);
        var measureName = RequireValue(action.Name, "name");
        var existing = table.Measures.FirstOrDefault(m => string.Equals(m.Name, measureName, StringComparison.OrdinalIgnoreCase));
        if (existing is null)
        {
            results.Add($"未找到度量值，跳过删除: {table.Name}.{measureName}");
            return false;
        }

        table.Measures.Remove(existing);
        results.Add($"在表 [{table.Name}] 下删除度量值 [{measureName}]");
        return true;
    }

    private static bool ApplyCreateRelationship(Database database, AbiModelAction action, ICollection<string> results)
    {
        var fromTable = FindTable(database, action.FromTable);
        var toTable = FindTable(database, action.ToTable);
        var fromColumnName = RequireValue(action.FromColumn, "fromColumn");
        var toColumnName = RequireValue(action.ToColumn, "toColumn");

        var fromColumn = fromTable.Columns.FirstOrDefault(c => string.Equals(c.Name, fromColumnName, StringComparison.OrdinalIgnoreCase))
                         ?? throw new InvalidOperationException($"未找到列: {fromTable.Name}.{fromColumnName}");
        var toColumn = toTable.Columns.FirstOrDefault(c => string.Equals(c.Name, toColumnName, StringComparison.OrdinalIgnoreCase))
                       ?? throw new InvalidOperationException($"未找到列: {toTable.Name}.{toColumnName}");

        var existing = FindRelationship(database, action);
        if (existing is null)
        {
            // Safety: if same endpoints already exist, update that relationship to avoid duplicate edge.
            existing = TryFindRelationshipByEndpoints(database, fromTable.Name, fromColumnName, toTable.Name, toColumnName);
        }

        if (existing is not null)
        {
            existing.FromColumn = fromColumn;
            existing.ToColumn = toColumn;
            existing.CrossFilteringBehavior = ParseCrossFilter(action.CrossFilterDirection);
            if (action.IsActive.HasValue)
            {
                existing.IsActive = action.IsActive.Value;
            }

            results.Add($"更新关系: {existing.Name}");
            return true;
        }

        var relationName = string.IsNullOrWhiteSpace(action.Name)
            ? $"{fromTable.Name}_{fromColumnName}_to_{toTable.Name}_{toColumnName}"
            : action.Name.Trim();

        var relationship = new SingleColumnRelationship
        {
            Name = relationName,
            FromColumn = fromColumn,
            ToColumn = toColumn,
            CrossFilteringBehavior = ParseCrossFilter(action.CrossFilterDirection),
            IsActive = action.IsActive ?? true
        };
        database.Model.Relationships.Add(relationship);
        results.Add($"新增关系: {relationName}");
        return true;
    }

    private static bool ApplyDeleteRelationship(Database database, AbiModelAction action, ICollection<string> results)
    {
        var existing = FindRelationship(database, action);
        if (existing is null)
        {
            results.Add("未找到关系，跳过删除。");
            return false;
        }

        database.Model.Relationships.Remove(existing);
        results.Add($"删除关系: {existing.Name}");
        return true;
    }

    private static bool HandleUnsupportedAction(string type, ICollection<string> results)
    {
        results.Add($"不支持的动作类型，已跳过: {type}");
        return false;
    }

    private static Table FindTable(Database database, string? tableName)
    {
        var name = RequireValue(tableName, "table/fromTable/toTable");
        var normalized = NormalizeTableReference(name);
        return database.Model.Tables.FirstOrDefault(t => NameEquals(t.Name, normalized))
               ?? throw new InvalidOperationException($"未找到表: {name}");
    }

    private static Table? TryFindTable(Database database, string? tableName)
    {
        if (string.IsNullOrWhiteSpace(tableName))
        {
            return null;
        }

        var normalized = NormalizeTableReference(tableName);
        return database.Model.Tables.FirstOrDefault(t => NameEquals(t.Name, normalized));
    }

    private static SingleColumnRelationship? FindRelationship(Database database, AbiModelAction action)
    {
        var all = database.Model.Relationships.OfType<SingleColumnRelationship>().ToList();
        if (!string.IsNullOrWhiteSpace(action.Name))
        {
            return all.FirstOrDefault(r => string.Equals(r.Name, action.Name.Trim(), StringComparison.OrdinalIgnoreCase));
        }

        if (string.IsNullOrWhiteSpace(action.FromTable) ||
            string.IsNullOrWhiteSpace(action.FromColumn) ||
            string.IsNullOrWhiteSpace(action.ToTable) ||
            string.IsNullOrWhiteSpace(action.ToColumn))
        {
            return null;
        }

        return all.FirstOrDefault(r =>
            string.Equals(r.FromColumn?.Table?.Name, action.FromTable, StringComparison.OrdinalIgnoreCase) &&
            string.Equals(r.FromColumn?.Name, action.FromColumn, StringComparison.OrdinalIgnoreCase) &&
            string.Equals(r.ToColumn?.Table?.Name, action.ToTable, StringComparison.OrdinalIgnoreCase) &&
            string.Equals(r.ToColumn?.Name, action.ToColumn, StringComparison.OrdinalIgnoreCase));
    }

    private static SingleColumnRelationship? TryFindRelationshipByName(Database database, string? name)
    {
        if (string.IsNullOrWhiteSpace(name))
        {
            return null;
        }

        return database.Model.Relationships
            .OfType<SingleColumnRelationship>()
            .FirstOrDefault(r => string.Equals(r.Name, name.Trim(), StringComparison.OrdinalIgnoreCase));
    }

    private static SingleColumnRelationship? TryFindRelationshipByEndpoints(Database database, string fromTable, string fromColumn, string toTable, string toColumn)
    {
        return database.Model.Relationships
            .OfType<SingleColumnRelationship>()
            .FirstOrDefault(r => IsSameEndpoints(r, fromTable, fromColumn, toTable, toColumn));
    }

    private static bool IsSameEndpoints(SingleColumnRelationship relationship, string fromTable, string fromColumn, string toTable, string toColumn)
    {
        return NameEquals(relationship.FromColumn?.Table?.Name ?? string.Empty, fromTable) &&
               string.Equals(relationship.FromColumn?.Name, fromColumn, StringComparison.OrdinalIgnoreCase) &&
               NameEquals(relationship.ToColumn?.Table?.Name ?? string.Empty, toTable) &&
               string.Equals(relationship.ToColumn?.Name, toColumn, StringComparison.OrdinalIgnoreCase);
    }

    private static bool NameEquals(string actual, string request)
        => string.Equals(NormalizeObjectName(actual), NormalizeObjectName(request), StringComparison.OrdinalIgnoreCase);

    private static string NormalizeTableReference(string? value)
    {
        var normalized = NormalizeObjectName(value);
        var bracketIndex = normalized.IndexOf('[');
        if (bracketIndex > 0 && normalized.EndsWith("]", StringComparison.Ordinal))
        {
            normalized = normalized[..bracketIndex].Trim();
            normalized = NormalizeObjectName(normalized);
        }
        return normalized;
    }

    private static string NormalizeObjectName(string? value)
    {
        var name = (value ?? string.Empty).Trim();
        while (name.Length >= 2)
        {
            var wrappedByQuotes =
                (name.StartsWith("'", StringComparison.Ordinal) && name.EndsWith("'", StringComparison.Ordinal)) ||
                (name.StartsWith("\"", StringComparison.Ordinal) && name.EndsWith("\"", StringComparison.Ordinal)) ||
                (name.StartsWith("`", StringComparison.Ordinal) && name.EndsWith("`", StringComparison.Ordinal));
            var wrappedByBrackets =
                name.StartsWith("[", StringComparison.Ordinal) && name.EndsWith("]", StringComparison.Ordinal);

            if (!wrappedByQuotes && !wrappedByBrackets)
            {
                break;
            }

            name = name[1..^1].Trim();
        }

        return name;
    }

    private static CrossFilteringBehavior ParseCrossFilter(string? value)
    {
        if (!string.IsNullOrWhiteSpace(value) &&
            Enum.TryParse<CrossFilteringBehavior>(value.Trim(), true, out var parsed))
        {
            return parsed;
        }

        if (string.Equals(value, "Both", StringComparison.OrdinalIgnoreCase))
        {
            return CrossFilteringBehavior.BothDirections;
        }

        if (string.Equals(value, "OneWay", StringComparison.OrdinalIgnoreCase))
        {
            return CrossFilteringBehavior.OneDirection;
        }

        return CrossFilteringBehavior.OneDirection;
    }

    private static string RequireValue(string? value, string fieldName)
    {
        return string.IsNullOrWhiteSpace(value)
            ? throw new InvalidOperationException($"动作缺少字段: {fieldName}")
            : value.Trim();
    }

    private static Database? ResolveDatabase(Server server, string? databaseName)
    {
        if (!string.IsNullOrWhiteSpace(databaseName))
        {
            return server.Databases
                .OfType<Database>()
                .FirstOrDefault(db => string.Equals(db.Name, databaseName, StringComparison.OrdinalIgnoreCase));
        }

        return server.Databases
            .OfType<Database>()
            .FirstOrDefault();
    }
}
