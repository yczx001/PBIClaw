using Microsoft.AnalysisServices.Tabular;

namespace PbiMetadataTool;

internal static class TabularExtendedActionHandler
{
    public static bool TryAnalyze(Database database, AbiModelAction action, ICollection<string> infos, ICollection<string> warnings, ICollection<string> errors)
    {
        var type = NormalizeType(action.Type);
        switch (type)
        {
            case "delete_table":
                AnalyzeDeleteTable(database, action, infos, warnings, errors);
                return true;
            case "rename_table":
                AnalyzeRenameTable(database, action, infos, warnings, errors);
                return true;
            case "rename_column":
                AnalyzeRenameColumn(database, action, infos, warnings, errors);
                return true;
            case "rename_measure":
                AnalyzeRenameMeasure(database, action, infos, warnings, errors);
                return true;
            case "set_table_hidden":
                AnalyzeSetTableHidden(database, action, infos, errors);
                return true;
            case "set_column_hidden":
                AnalyzeSetColumnHidden(database, action, infos, errors);
                return true;
            case "set_measure_hidden":
                AnalyzeSetMeasureHidden(database, action, infos, errors);
                return true;
            case "set_format_string":
                AnalyzeSetFormatString(database, action, infos, errors);
                return true;
            case "set_display_folder":
                AnalyzeSetDisplayFolder(database, action, infos, errors);
                return true;
            case "create_calculated_column":
                AnalyzeCreateCalculatedColumn(database, action, infos, warnings, errors);
                return true;
            case "delete_column":
                AnalyzeDeleteColumn(database, action, infos, warnings, errors);
                return true;
            case "create_calculated_table":
                AnalyzeCreateCalculatedTable(database, action, infos, warnings, errors);
                return true;
            case "set_relationship_active":
                AnalyzeSetRelationshipActive(database, action, infos, errors);
                return true;
            case "set_relationship_cross_filter":
                AnalyzeSetRelationshipCrossFilter(database, action, infos, errors);
                return true;
            case "update_description":
                AnalyzeUpdateDescription(database, action, infos, errors);
                return true;
            case "create_role":
                AnalyzeCreateRole(database, action, infos, warnings, errors);
                return true;
            case "update_role":
                AnalyzeUpdateRole(database, action, infos, warnings, errors);
                return true;
            case "delete_role":
                AnalyzeDeleteRole(database, action, infos, warnings, errors);
                return true;
            case "set_role_table_permission":
                AnalyzeSetRoleTablePermission(database, action, infos, warnings, errors);
                return true;
            case "remove_role_table_permission":
                AnalyzeRemoveRoleTablePermission(database, action, infos, warnings, errors);
                return true;
            case "add_role_member":
                AnalyzeAddRoleMember(database, action, infos, warnings, errors);
                return true;
            case "remove_role_member":
                AnalyzeRemoveRoleMember(database, action, infos, warnings, errors);
                return true;
            default:
                return false;
        }
    }

    public static bool TryApply(Database database, AbiModelAction action, ICollection<string> results)
    {
        var type = NormalizeType(action.Type);
        return type switch
        {
            "delete_table" => ApplyDeleteTable(database, action, results),
            "rename_table" => ApplyRenameTable(database, action, results),
            "rename_column" => ApplyRenameColumn(database, action, results),
            "rename_measure" => ApplyRenameMeasure(database, action, results),
            "set_table_hidden" => ApplySetTableHidden(database, action, results),
            "set_column_hidden" => ApplySetColumnHidden(database, action, results),
            "set_measure_hidden" => ApplySetMeasureHidden(database, action, results),
            "set_format_string" => ApplySetFormatString(database, action, results),
            "set_display_folder" => ApplySetDisplayFolder(database, action, results),
            "create_calculated_column" => ApplyCreateCalculatedColumn(database, action, results),
            "delete_column" => ApplyDeleteColumn(database, action, results),
            "create_calculated_table" => ApplyCreateCalculatedTable(database, action, results),
            "set_relationship_active" => ApplySetRelationshipActive(database, action, results),
            "set_relationship_cross_filter" => ApplySetRelationshipCrossFilter(database, action, results),
            "update_description" => ApplyUpdateDescription(database, action, results),
            "create_role" => ApplyCreateRole(database, action, results),
            "update_role" => ApplyUpdateRole(database, action, results),
            "delete_role" => ApplyDeleteRole(database, action, results),
            "set_role_table_permission" => ApplySetRoleTablePermission(database, action, results),
            "remove_role_table_permission" => ApplyRemoveRoleTablePermission(database, action, results),
            "add_role_member" => ApplyAddRoleMember(database, action, results),
            "remove_role_member" => ApplyRemoveRoleMember(database, action, results),
            _ => false
        };
    }

    private static string NormalizeType(string? type)
    {
        var t = (type ?? string.Empty).Trim().ToLowerInvariant();
        return t switch
        {
            "hide_or_show_table" => "set_table_hidden",
            "hide_or_show_column" => "set_column_hidden",
            "hide_or_show_measure" => "set_measure_hidden",
            "set_relationship_state" => "set_relationship_active",
            "create_or_update_role" => "update_role",
            _ => t
        };
    }

    private static void AnalyzeDeleteTable(Database database, AbiModelAction action, ICollection<string> infos, ICollection<string> warnings, ICollection<string> errors)
    {
        var table = TryFindTable(database, action.Table);
        if (table is null)
        {
            warnings.Add($"表不存在，将跳过删除: {action.Table}");
            return;
        }

        var relatedRelationships = GetRelationshipsTouchingTable(database, table.Name);
        if (relatedRelationships.Count > 0)
        {
            warnings.Add($"删除表 {table.Name} 时将同时删除 {relatedRelationships.Count} 条关联关系。");
        }

        var rolePermissionImpacts = new List<string>();
        foreach (var role in database.Model.Roles.OfType<ModelRole>())
        {
            var hitCount = role.TablePermissions
                .OfType<TablePermission>()
                .Count(p => IsRolePermissionForTable(p, table.Name));
            if (hitCount > 0)
            {
                rolePermissionImpacts.Add($"{role.Name}({hitCount})");
            }
        }

        if (rolePermissionImpacts.Count > 0)
        {
            warnings.Add($"删除表 {table.Name} 时将移除角色表权限: {string.Join(", ", rolePermissionImpacts)}");
        }

        warnings.Add($"将删除表: {table.Name}");
        infos.Add($"删除动作已通过预检: {table.Name}");
    }

    private static void AnalyzeRenameTable(Database database, AbiModelAction action, ICollection<string> infos, ICollection<string> warnings, ICollection<string> errors)
    {
        var table = TryFindTable(database, action.Table);
        var newName = action.NewName?.Trim();
        if (table is null)
        {
            errors.Add($"表不存在: {action.Table}");
            return;
        }
        if (string.IsNullOrWhiteSpace(newName))
        {
            errors.Add("rename_table 缺少 newName");
            return;
        }
        if (TryFindTable(database, newName) is not null && !string.Equals(table.Name, newName, StringComparison.OrdinalIgnoreCase))
        {
            errors.Add($"目标表名已存在: {newName}");
            return;
        }
        warnings.Add($"将重命名表: {table.Name} -> {newName}");
        infos.Add($"重命名动作已通过预检: {table.Name}");
    }

    private static void AnalyzeRenameColumn(Database database, AbiModelAction action, ICollection<string> infos, ICollection<string> warnings, ICollection<string> errors)
    {
        var table = TryFindTable(database, action.Table);
        var name = action.Name?.Trim();
        var newName = action.NewName?.Trim();
        if (table is null || string.IsNullOrWhiteSpace(name) || string.IsNullOrWhiteSpace(newName))
        {
            errors.Add("rename_column 缺少 table/name/newName 或对象不存在");
            return;
        }
        if (TryFindColumn(table, name) is null)
        {
            errors.Add($"列不存在: {table.Name}.{name}");
            return;
        }
        if (TryFindColumn(table, newName) is not null && !string.Equals(name, newName, StringComparison.OrdinalIgnoreCase))
        {
            errors.Add($"目标列名已存在: {table.Name}.{newName}");
            return;
        }
        warnings.Add($"将重命名列: {table.Name}.{name} -> {newName}");
    }

    private static void AnalyzeRenameMeasure(Database database, AbiModelAction action, ICollection<string> infos, ICollection<string> warnings, ICollection<string> errors)
    {
        var table = TryFindTable(database, action.Table);
        var name = action.Name?.Trim();
        var newName = action.NewName?.Trim();
        if (table is null || string.IsNullOrWhiteSpace(name) || string.IsNullOrWhiteSpace(newName))
        {
            errors.Add("rename_measure 缺少 table/name/newName 或对象不存在");
            return;
        }
        if (TryFindMeasure(table, name) is null)
        {
            errors.Add($"度量值不存在: {table.Name}.{name}");
            return;
        }
        if (TryFindMeasure(table, newName) is not null && !string.Equals(name, newName, StringComparison.OrdinalIgnoreCase))
        {
            errors.Add($"目标度量值名已存在: {table.Name}.{newName}");
            return;
        }
        warnings.Add($"将重命名度量值: {table.Name}.{name} -> {newName}");
    }

    private static void AnalyzeSetTableHidden(Database database, AbiModelAction action, ICollection<string> infos, ICollection<string> errors)
    {
        var table = TryFindTable(database, action.Table);
        if (table is null || !action.IsHidden.HasValue)
        {
            errors.Add("set_table_hidden 缺少 table/isHidden 或对象不存在");
            return;
        }
        infos.Add($"将设置表显隐: {table.Name} -> {(action.IsHidden.Value ? "Hidden" : "Visible")}");
    }

    private static void AnalyzeSetColumnHidden(Database database, AbiModelAction action, ICollection<string> infos, ICollection<string> errors)
    {
        var table = TryFindTable(database, action.Table);
        if (table is null || string.IsNullOrWhiteSpace(action.Name) || !action.IsHidden.HasValue || TryFindColumn(table, action.Name) is null)
        {
            errors.Add("set_column_hidden 缺少 table/name/isHidden 或对象不存在");
            return;
        }
        infos.Add($"将设置列显隐: {table.Name}.{action.Name}");
    }

    private static void AnalyzeSetMeasureHidden(Database database, AbiModelAction action, ICollection<string> infos, ICollection<string> errors)
    {
        var table = TryFindTable(database, action.Table);
        if (table is null || string.IsNullOrWhiteSpace(action.Name) || !action.IsHidden.HasValue || TryFindMeasure(table, action.Name) is null)
        {
            errors.Add("set_measure_hidden 缺少 table/name/isHidden 或对象不存在");
            return;
        }
        infos.Add($"将设置度量值显隐: {table.Name}.{action.Name}");
    }

    private static void AnalyzeSetFormatString(Database database, AbiModelAction action, ICollection<string> infos, ICollection<string> errors)
    {
        var table = TryFindTable(database, action.Table);
        if (table is null || string.IsNullOrWhiteSpace(action.Name) || action.FormatString is null || TryFindMeasure(table, action.Name) is null)
        {
            errors.Add("set_format_string 缺少 table/name/formatString 或对象不存在");
            return;
        }
        infos.Add($"将更新格式字符串: {table.Name}.{action.Name}");
    }

    private static void AnalyzeSetDisplayFolder(Database database, AbiModelAction action, ICollection<string> infos, ICollection<string> errors)
    {
        var table = TryFindTable(database, action.Table);
        if (table is null || string.IsNullOrWhiteSpace(action.Name) || TryFindMeasure(table, action.Name) is null)
        {
            errors.Add("set_display_folder 缺少 table/name 或对象不存在");
            return;
        }
        infos.Add($"将更新显示文件夹: {table.Name}.{action.Name}");
    }

    private static void AnalyzeCreateCalculatedColumn(Database database, AbiModelAction action, ICollection<string> infos, ICollection<string> warnings, ICollection<string> errors)
    {
        var table = TryFindTable(database, action.Table);
        if (table is null || string.IsNullOrWhiteSpace(action.Name) || string.IsNullOrWhiteSpace(action.Expression))
        {
            errors.Add("create_calculated_column 缺少 table/name/expression 或对象不存在");
            return;
        }
        var existing = TryFindColumn(table, action.Name);
        if (existing is null)
        {
            infos.Add($"将新增计算列: {table.Name}.{action.Name}");
            return;
        }
        if (existing is not CalculatedColumn)
        {
            errors.Add($"同名列已存在且不是计算列: {table.Name}.{action.Name}");
            return;
        }
        warnings.Add($"将更新计算列: {table.Name}.{action.Name}");
    }

    private static void AnalyzeDeleteColumn(Database database, AbiModelAction action, ICollection<string> infos, ICollection<string> warnings, ICollection<string> errors)
    {
        var table = TryFindTable(database, action.Table);
        if (table is null || string.IsNullOrWhiteSpace(action.Name))
        {
            errors.Add("delete_column 缺少 table/name 或对象不存在");
            return;
        }
        if (TryFindColumn(table, action.Name) is null)
        {
            warnings.Add($"列不存在，将跳过删除: {table.Name}.{action.Name}");
            return;
        }
        warnings.Add($"将删除列: {table.Name}.{action.Name}");
        infos.Add($"删除动作已通过预检: {table.Name}.{action.Name}");
    }

    private static void AnalyzeCreateCalculatedTable(Database database, AbiModelAction action, ICollection<string> infos, ICollection<string> warnings, ICollection<string> errors)
    {
        if (string.IsNullOrWhiteSpace(action.Name) || string.IsNullOrWhiteSpace(action.Expression))
        {
            errors.Add("create_calculated_table 缺少 name/expression");
            return;
        }
        if (TryFindTable(database, action.Name) is null)
        {
            infos.Add($"将新增计算表: {action.Name}");
            return;
        }
        warnings.Add($"同名表已存在，将更新其计算表达式: {action.Name}");
    }

    private static void AnalyzeSetRelationshipActive(Database database, AbiModelAction action, ICollection<string> infos, ICollection<string> errors)
    {
        if (!action.IsActive.HasValue || FindRelationship(database, action) is null)
        {
            errors.Add("set_relationship_active 缺少 isActive 或关系不存在");
            return;
        }
        infos.Add("将更新关系激活状态");
    }

    private static void AnalyzeSetRelationshipCrossFilter(Database database, AbiModelAction action, ICollection<string> infos, ICollection<string> errors)
    {
        if (string.IsNullOrWhiteSpace(action.CrossFilterDirection) || FindRelationship(database, action) is null)
        {
            errors.Add("set_relationship_cross_filter 缺少 crossFilterDirection 或关系不存在");
            return;
        }
        infos.Add("将更新关系筛选方向");
    }

    private static void AnalyzeUpdateDescription(Database database, AbiModelAction action, ICollection<string> infos, ICollection<string> errors)
    {
        var objectType = NormalizeObjectType(action.ObjectType);
        if (string.IsNullOrWhiteSpace(objectType))
        {
            errors.Add("update_description 缺少 objectType（table/column/measure/role）");
            return;
        }
        infos.Add($"将更新描述: {objectType}");
    }

    private static void AnalyzeCreateRole(Database database, AbiModelAction action, ICollection<string> infos, ICollection<string> warnings, ICollection<string> errors)
    {
        if (string.IsNullOrWhiteSpace(action.Name))
        {
            errors.Add("create_role 缺少 name");
            return;
        }
        if (TryFindRole(database, action.Name) is null) infos.Add($"将新增角色: {action.Name}");
        else warnings.Add($"角色已存在，将按更新处理: {action.Name}");
    }

    private static void AnalyzeUpdateRole(Database database, AbiModelAction action, ICollection<string> infos, ICollection<string> warnings, ICollection<string> errors)
    {
        if (string.IsNullOrWhiteSpace(action.Name) || TryFindRole(database, action.Name) is null)
        {
            errors.Add($"角色不存在: {action.Name}");
            return;
        }
        infos.Add($"将更新角色: {action.Name}");
    }

    private static void AnalyzeDeleteRole(Database database, AbiModelAction action, ICollection<string> infos, ICollection<string> warnings, ICollection<string> errors)
    {
        if (string.IsNullOrWhiteSpace(action.Name))
        {
            errors.Add("delete_role 缺少 name");
            return;
        }
        if (TryFindRole(database, action.Name) is null) warnings.Add($"角色不存在，将跳过删除: {action.Name}");
        else warnings.Add($"将删除角色: {action.Name}");
    }

    private static void AnalyzeSetRoleTablePermission(Database database, AbiModelAction action, ICollection<string> infos, ICollection<string> warnings, ICollection<string> errors)
    {
        var role = TryFindRole(database, action.Name);
        var table = TryFindTable(database, action.Table);
        if (role is null || table is null)
        {
            errors.Add("set_role_table_permission 缺少 role/table 或对象不存在");
            return;
        }
        infos.Add($"将设置角色表权限: {role.Name}.{table.Name}");
    }

    private static void AnalyzeRemoveRoleTablePermission(Database database, AbiModelAction action, ICollection<string> infos, ICollection<string> warnings, ICollection<string> errors)
    {
        var role = TryFindRole(database, action.Name);
        if (role is null || string.IsNullOrWhiteSpace(action.Table))
        {
            errors.Add("remove_role_table_permission 缺少 role/table 或对象不存在");
            return;
        }
        infos.Add($"将删除角色表权限: {role.Name}.{action.Table}");
    }

    private static void AnalyzeAddRoleMember(Database database, AbiModelAction action, ICollection<string> infos, ICollection<string> warnings, ICollection<string> errors)
    {
        var role = TryFindRole(database, action.Name);
        if (role is null || string.IsNullOrWhiteSpace(action.MemberName))
        {
            errors.Add("add_role_member 缺少 role/memberName 或对象不存在");
            return;
        }
        infos.Add($"将新增角色成员: {role.Name}.{action.MemberName}");
    }

    private static void AnalyzeRemoveRoleMember(Database database, AbiModelAction action, ICollection<string> infos, ICollection<string> warnings, ICollection<string> errors)
    {
        var role = TryFindRole(database, action.Name);
        if (role is null || string.IsNullOrWhiteSpace(action.MemberName))
        {
            errors.Add("remove_role_member 缺少 role/memberName 或对象不存在");
            return;
        }
        infos.Add($"将删除角色成员: {role.Name}.{action.MemberName}");
    }

    private static bool ApplyDeleteTable(Database database, AbiModelAction action, ICollection<string> results)
    {
        var table = TryFindTable(database, action.Table);
        if (table is null)
        {
            results.Add($"未找到表，跳过删除: {action.Table}");
            return false;
        }

        var relationships = GetRelationshipsTouchingTable(database, table.Name);
        foreach (var relationship in relationships)
        {
            database.Model.Relationships.Remove(relationship);
            results.Add($"删除关联关系: {relationship.Name}");
        }

        foreach (var role in database.Model.Roles.OfType<ModelRole>())
        {
            var permissions = role.TablePermissions
                .OfType<TablePermission>()
                .Where(p => IsRolePermissionForTable(p, table.Name))
                .ToList();
            foreach (var permission in permissions)
            {
                role.TablePermissions.Remove(permission);
                results.Add($"删除角色表权限: {role.Name}.{table.Name}");
            }
        }

        database.Model.Tables.Remove(table);
        results.Add($"删除表: {table.Name}");
        return true;
    }

    private static bool ApplyRenameTable(Database database, AbiModelAction action, ICollection<string> results)
    {
        var table = FindTable(database, action.Table);
        var newName = RequireValue(action.NewName, "newName");
        if (TryFindTable(database, newName) is not null &&
            !string.Equals(table.Name, newName, StringComparison.OrdinalIgnoreCase))
        {
            throw new InvalidOperationException($"目标表名已存在: {newName}");
        }

        if (string.Equals(table.Name, newName, StringComparison.OrdinalIgnoreCase))
        {
            results.Add($"表名无变化，跳过: {table.Name}");
            return false;
        }

        var oldName = table.Name;
        table.Name = newName;
        results.Add($"重命名表: {oldName} -> {newName}");
        return true;
    }

    private static bool ApplyRenameColumn(Database database, AbiModelAction action, ICollection<string> results)
    {
        var table = FindTable(database, action.Table);
        var oldName = RequireValue(action.Name, "name");
        var newName = RequireValue(action.NewName, "newName");
        var column = FindColumn(table, oldName);
        if (TryFindColumn(table, newName) is not null &&
            !string.Equals(oldName, newName, StringComparison.OrdinalIgnoreCase))
        {
            throw new InvalidOperationException($"目标列名已存在: {table.Name}.{newName}");
        }

        if (string.Equals(column.Name, newName, StringComparison.OrdinalIgnoreCase))
        {
            results.Add($"列名无变化，跳过: {table.Name}.{column.Name}");
            return false;
        }

        column.Name = newName;
        results.Add($"重命名列: {table.Name}.{oldName} -> {newName}");
        return true;
    }

    private static bool ApplyRenameMeasure(Database database, AbiModelAction action, ICollection<string> results)
    {
        var table = FindTable(database, action.Table);
        var oldName = RequireValue(action.Name, "name");
        var newName = RequireValue(action.NewName, "newName");
        var measure = FindMeasure(table, oldName);
        if (TryFindMeasure(table, newName) is not null &&
            !string.Equals(oldName, newName, StringComparison.OrdinalIgnoreCase))
        {
            throw new InvalidOperationException($"目标度量值名已存在: {table.Name}.{newName}");
        }

        if (string.Equals(measure.Name, newName, StringComparison.OrdinalIgnoreCase))
        {
            results.Add($"度量值名无变化，跳过: {table.Name}.{measure.Name}");
            return false;
        }

        measure.Name = newName;
        results.Add($"重命名度量值: {table.Name}.{oldName} -> {newName}");
        return true;
    }

    private static bool ApplySetTableHidden(Database database, AbiModelAction action, ICollection<string> results)
    {
        var table = FindTable(database, action.Table);
        var hidden = action.IsHidden ?? throw new InvalidOperationException("动作缺少字段: isHidden");
        if (table.IsHidden == hidden)
        {
            results.Add($"表显隐无变化，跳过: {table.Name}");
            return false;
        }
        table.IsHidden = hidden;
        results.Add($"设置表显隐: {table.Name} -> {(hidden ? "Hidden" : "Visible")}");
        return true;
    }

    private static bool ApplySetColumnHidden(Database database, AbiModelAction action, ICollection<string> results)
    {
        var table = FindTable(database, action.Table);
        var column = FindColumn(table, action.Name);
        var hidden = action.IsHidden ?? throw new InvalidOperationException("动作缺少字段: isHidden");
        if (column.IsHidden == hidden)
        {
            results.Add($"列显隐无变化，跳过: {table.Name}.{column.Name}");
            return false;
        }
        column.IsHidden = hidden;
        results.Add($"设置列显隐: {table.Name}.{column.Name} -> {(hidden ? "Hidden" : "Visible")}");
        return true;
    }

    private static bool ApplySetMeasureHidden(Database database, AbiModelAction action, ICollection<string> results)
    {
        var table = FindTable(database, action.Table);
        var measure = FindMeasure(table, action.Name);
        var hidden = action.IsHidden ?? throw new InvalidOperationException("动作缺少字段: isHidden");
        if (measure.IsHidden == hidden)
        {
            results.Add($"度量值显隐无变化，跳过: {table.Name}.{measure.Name}");
            return false;
        }
        measure.IsHidden = hidden;
        results.Add($"设置度量值显隐: {table.Name}.{measure.Name} -> {(hidden ? "Hidden" : "Visible")}");
        return true;
    }

    private static bool ApplySetFormatString(Database database, AbiModelAction action, ICollection<string> results)
    {
        var table = FindTable(database, action.Table);
        var measure = FindMeasure(table, action.Name);
        var formatString = action.FormatString ?? throw new InvalidOperationException("动作缺少字段: formatString");
        if (string.Equals(measure.FormatString ?? string.Empty, formatString, StringComparison.Ordinal))
        {
            results.Add($"格式字符串无变化，跳过: {table.Name}.{measure.Name}");
            return false;
        }
        measure.FormatString = formatString;
        results.Add($"更新格式字符串: {table.Name}.{measure.Name}");
        return true;
    }

    private static bool ApplySetDisplayFolder(Database database, AbiModelAction action, ICollection<string> results)
    {
        var table = FindTable(database, action.Table);
        var measure = FindMeasure(table, action.Name);
        var displayFolder = action.DisplayFolder ?? string.Empty;
        if (string.Equals(measure.DisplayFolder ?? string.Empty, displayFolder, StringComparison.Ordinal))
        {
            results.Add($"显示文件夹无变化，跳过: {table.Name}.{measure.Name}");
            return false;
        }
        measure.DisplayFolder = displayFolder;
        results.Add($"更新显示文件夹: {table.Name}.{measure.Name}");
        return true;
    }

    private static bool ApplyCreateCalculatedColumn(Database database, AbiModelAction action, ICollection<string> results)
    {
        var table = FindTable(database, action.Table);
        var name = RequireValue(action.Name, "name");
        var expression = RequireValue(action.Expression, "expression");
        var existing = TryFindColumn(table, name);
        CalculatedColumn target;
        if (existing is null)
        {
            target = new CalculatedColumn
            {
                Name = name,
                Expression = expression
            };
            table.Columns.Add(target);
            results.Add($"创建计算列: {table.Name}.{name}");
        }
        else if (existing is CalculatedColumn calculated)
        {
            target = calculated;
            target.Expression = expression;
            results.Add($"更新计算列: {table.Name}.{name}");
        }
        else
        {
            throw new InvalidOperationException($"同名列已存在且不是计算列: {table.Name}.{name}");
        }

        if (!string.IsNullOrWhiteSpace(action.DataType) && TryParseDataType(action.DataType, out var dt))
        {
            target.DataType = dt;
        }
        if (!string.IsNullOrWhiteSpace(action.FormatString))
        {
            target.FormatString = action.FormatString;
        }
        if (action.IsHidden.HasValue)
        {
            target.IsHidden = action.IsHidden.Value;
        }
        if (action.Description is not null)
        {
            target.Description = action.Description;
        }
        return true;
    }

    private static bool ApplyDeleteColumn(Database database, AbiModelAction action, ICollection<string> results)
    {
        var table = FindTable(database, action.Table);
        var name = RequireValue(action.Name, "name");
        var column = TryFindColumn(table, name);
        if (column is null)
        {
            results.Add($"未找到列，跳过删除: {table.Name}.{name}");
            return false;
        }
        table.Columns.Remove(column);
        results.Add($"删除列: {table.Name}.{name}");
        return true;
    }

    private static bool ApplyCreateCalculatedTable(Database database, AbiModelAction action, ICollection<string> results)
    {
        var tableName = RequireValue(action.Name, "name");
        var expression = RequireValue(action.Expression, "expression");
        var existing = TryFindTable(database, tableName);
        if (existing is null)
        {
            var table = new Table { Name = tableName };
            if (action.IsHidden.HasValue) table.IsHidden = action.IsHidden.Value;
            if (action.Description is not null) table.Description = action.Description;
            var partition = new Partition
            {
                Name = tableName,
                Source = new CalculatedPartitionSource
                {
                    Expression = expression
                }
            };
            table.Partitions.Add(partition);
            database.Model.Tables.Add(table);
            results.Add($"创建计算表: {tableName}");
            return true;
        }

        var partition0 = existing.Partitions.FirstOrDefault();
        if (partition0 is null)
        {
            partition0 = new Partition { Name = tableName };
            existing.Partitions.Add(partition0);
        }
        if (partition0.Source is not CalculatedPartitionSource cps)
        {
            cps = new CalculatedPartitionSource();
            partition0.Source = cps;
        }
        cps.Expression = expression;
        if (action.IsHidden.HasValue) existing.IsHidden = action.IsHidden.Value;
        if (action.Description is not null) existing.Description = action.Description;
        results.Add($"更新计算表表达式: {tableName}");
        return true;
    }

    private static bool ApplySetRelationshipActive(Database database, AbiModelAction action, ICollection<string> results)
    {
        var relationship = FindRelationship(database, action)
                           ?? throw new InvalidOperationException("未找到关系。");
        var isActive = action.IsActive ?? throw new InvalidOperationException("动作缺少字段: isActive");
        if (relationship.IsActive == isActive)
        {
            results.Add($"关系激活状态无变化，跳过: {relationship.Name}");
            return false;
        }
        relationship.IsActive = isActive;
        results.Add($"设置关系激活状态: {relationship.Name} -> {isActive}");
        return true;
    }

    private static bool ApplySetRelationshipCrossFilter(Database database, AbiModelAction action, ICollection<string> results)
    {
        var relationship = FindRelationship(database, action)
                           ?? throw new InvalidOperationException("未找到关系。");
        var behavior = ParseCrossFilter(action.CrossFilterDirection);
        if (relationship.CrossFilteringBehavior == behavior)
        {
            results.Add($"关系筛选方向无变化，跳过: {relationship.Name}");
            return false;
        }
        relationship.CrossFilteringBehavior = behavior;
        results.Add($"设置关系筛选方向: {relationship.Name} -> {behavior}");
        return true;
    }

    private static bool ApplyUpdateDescription(Database database, AbiModelAction action, ICollection<string> results)
    {
        var objectType = NormalizeObjectType(action.ObjectType);
        var description = action.Description ?? string.Empty;
        switch (objectType)
        {
            case "table":
                {
                    var table = FindTable(database, action.Table);
                    table.Description = description;
                    results.Add($"更新表描述: {table.Name}");
                    return true;
                }
            case "column":
                {
                    var table = FindTable(database, action.Table);
                    var column = FindColumn(table, action.Name);
                    column.Description = description;
                    results.Add($"更新列描述: {table.Name}.{column.Name}");
                    return true;
                }
            case "measure":
                {
                    var table = FindTable(database, action.Table);
                    var measure = FindMeasure(table, action.Name);
                    measure.Description = description;
                    results.Add($"更新度量值描述: {table.Name}.{measure.Name}");
                    return true;
                }
            case "role":
                {
                    var role = FindRole(database, action.Name);
                    role.Description = description;
                    results.Add($"更新角色描述: {role.Name}");
                    return true;
                }
            default:
                throw new InvalidOperationException($"update_description 暂不支持 objectType: {action.ObjectType}");
        }
    }

    private static bool ApplyCreateRole(Database database, AbiModelAction action, ICollection<string> results)
    {
        var roleName = RequireValue(action.Name, "name");
        var role = TryFindRole(database, roleName);
        if (role is null)
        {
            role = new ModelRole { Name = roleName };
            database.Model.Roles.Add(role);
            results.Add($"创建角色: {roleName}");
        }
        else
        {
            results.Add($"角色已存在，按更新处理: {roleName}");
        }
        if (action.Description is not null) role.Description = action.Description;
        if (!string.IsNullOrWhiteSpace(action.ModelPermission)) role.ModelPermission = ParseModelPermission(action.ModelPermission);
        return true;
    }

    private static bool ApplyUpdateRole(Database database, AbiModelAction action, ICollection<string> results)
    {
        var role = FindRole(database, action.Name);
        if (!string.IsNullOrWhiteSpace(action.NewName))
        {
            var newName = action.NewName.Trim();
            var conflict = TryFindRole(database, newName);
            if (conflict is not null && !string.Equals(conflict.Name, role.Name, StringComparison.OrdinalIgnoreCase))
            {
                throw new InvalidOperationException($"目标角色名已存在: {newName}");
            }
            role.Name = newName;
        }
        if (action.Description is not null) role.Description = action.Description;
        if (!string.IsNullOrWhiteSpace(action.ModelPermission)) role.ModelPermission = ParseModelPermission(action.ModelPermission);
        results.Add($"更新角色: {role.Name}");
        return true;
    }

    private static bool ApplyDeleteRole(Database database, AbiModelAction action, ICollection<string> results)
    {
        var role = TryFindRole(database, action.Name);
        if (role is null)
        {
            results.Add($"未找到角色，跳过删除: {action.Name}");
            return false;
        }
        database.Model.Roles.Remove(role);
        results.Add($"删除角色: {role.Name}");
        return true;
    }

    private static bool ApplySetRoleTablePermission(Database database, AbiModelAction action, ICollection<string> results)
    {
        var role = FindRole(database, action.Name);
        var table = FindTable(database, action.Table);
        var permission = TryFindRoleTablePermission(role, table.Name);
        var filter = action.Expression ?? string.Empty;
        var metaPermission = ParseMetadataPermission(action.MetadataPermission);
        if (permission is null)
        {
            permission = new TablePermission
            {
                Table = table,
                FilterExpression = filter,
                MetadataPermission = metaPermission
            };
            role.TablePermissions.Add(permission);
            results.Add($"新增角色表权限: {role.Name}.{table.Name}");
            return true;
        }
        permission.FilterExpression = filter;
        permission.MetadataPermission = metaPermission;
        results.Add($"更新角色表权限: {role.Name}.{table.Name}");
        return true;
    }

    private static bool ApplyRemoveRoleTablePermission(Database database, AbiModelAction action, ICollection<string> results)
    {
        var role = FindRole(database, action.Name);
        var tableName = RequireValue(action.Table, "table");
        var permission = TryFindRoleTablePermission(role, tableName);
        if (permission is null)
        {
            results.Add($"未找到角色表权限，跳过: {role.Name}.{tableName}");
            return false;
        }
        role.TablePermissions.Remove(permission);
        results.Add($"删除角色表权限: {role.Name}.{tableName}");
        return true;
    }

    private static bool ApplyAddRoleMember(Database database, AbiModelAction action, ICollection<string> results)
    {
        var role = FindRole(database, action.Name);
        var memberName = RequireValue(action.MemberName, "memberName");
        if (TryFindRoleMember(role, memberName) is not null)
        {
            results.Add($"角色成员已存在，跳过新增: {role.Name}.{memberName}");
            return false;
        }
        var member = CreateRoleMember(memberName, action);
        role.Members.Add(member);
        results.Add($"新增角色成员: {role.Name}.{memberName}");
        return true;
    }

    private static bool ApplyRemoveRoleMember(Database database, AbiModelAction action, ICollection<string> results)
    {
        var role = FindRole(database, action.Name);
        var memberName = RequireValue(action.MemberName, "memberName");
        var member = TryFindRoleMember(role, memberName);
        if (member is null)
        {
            results.Add($"未找到角色成员，跳过删除: {role.Name}.{memberName}");
            return false;
        }
        role.Members.Remove(member);
        results.Add($"删除角色成员: {role.Name}.{memberName}");
        return true;
    }

    private static Table FindTable(Database database, string? tableName)
    {
        var name = RequireValue(tableName, "table");
        return TryFindTable(database, name)
               ?? throw new InvalidOperationException($"未找到表: {name}");
    }

    private static Table? TryFindTable(Database database, string? tableName)
    {
        if (string.IsNullOrWhiteSpace(tableName))
        {
            return null;
        }
        var requestName = NormalizeTableReference(tableName);
        return database.Model.Tables.FirstOrDefault(t => NameEquals(t.Name, requestName));
    }

    private static Column FindColumn(Table table, string? name)
    {
        var columnName = RequireValue(name, "name");
        return TryFindColumn(table, columnName)
               ?? throw new InvalidOperationException($"未找到列: {table.Name}.{columnName}");
    }

    private static Column? TryFindColumn(Table table, string? name)
    {
        if (string.IsNullOrWhiteSpace(name))
        {
            return null;
        }
        var requestName = NormalizeObjectName(name);
        return table.Columns.FirstOrDefault(c => NameEquals(c.Name, requestName));
    }

    private static Measure FindMeasure(Table table, string? name)
    {
        var measureName = RequireValue(name, "name");
        return TryFindMeasure(table, measureName)
               ?? throw new InvalidOperationException($"未找到度量值: {table.Name}.{measureName}");
    }

    private static Measure? TryFindMeasure(Table table, string? name)
    {
        if (string.IsNullOrWhiteSpace(name))
        {
            return null;
        }
        var requestName = NormalizeObjectName(name);
        return table.Measures.FirstOrDefault(m => NameEquals(m.Name, requestName));
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

    private static ModelRole? TryFindRole(Database database, string? roleName)
    {
        if (string.IsNullOrWhiteSpace(roleName))
        {
            return null;
        }
        return database.Model.Roles.OfType<ModelRole>()
            .FirstOrDefault(r => string.Equals(r.Name, roleName.Trim(), StringComparison.OrdinalIgnoreCase));
    }

    private static ModelRole FindRole(Database database, string? roleName)
    {
        var name = RequireValue(roleName, "name(role)");
        return TryFindRole(database, name)
               ?? throw new InvalidOperationException($"未找到角色: {name}");
    }

    private static TablePermission? TryFindRoleTablePermission(ModelRole role, string tableName)
    {
        return role.TablePermissions
            .OfType<TablePermission>()
            .FirstOrDefault(p => string.Equals(p.Table?.Name, tableName, StringComparison.OrdinalIgnoreCase));
    }

    private static List<SingleColumnRelationship> GetRelationshipsTouchingTable(Database database, string tableName)
    {
        return database.Model.Relationships
            .OfType<SingleColumnRelationship>()
            .Where(r =>
                string.Equals(r.FromColumn?.Table?.Name, tableName, StringComparison.OrdinalIgnoreCase) ||
                string.Equals(r.ToColumn?.Table?.Name, tableName, StringComparison.OrdinalIgnoreCase))
            .ToList();
    }

    private static bool IsRolePermissionForTable(TablePermission permission, string tableName)
    {
        return string.Equals(permission.Table?.Name, tableName, StringComparison.OrdinalIgnoreCase);
    }

    private static ModelRoleMember? TryFindRoleMember(ModelRole role, string memberName)
    {
        return role.Members
            .OfType<ModelRoleMember>()
            .FirstOrDefault(m =>
                string.Equals(m.MemberName, memberName, StringComparison.OrdinalIgnoreCase) ||
                string.Equals(m.Name, memberName, StringComparison.OrdinalIgnoreCase) ||
                string.Equals(m.MemberID, memberName, StringComparison.OrdinalIgnoreCase));
    }

    private static ModelRoleMember CreateRoleMember(string memberName, AbiModelAction action)
    {
        if (!string.IsNullOrWhiteSpace(action.IdentityProvider) || !string.IsNullOrWhiteSpace(action.MemberType))
        {
            return new ExternalModelRoleMember
            {
                MemberName = memberName,
                IdentityProvider = action.IdentityProvider ?? string.Empty,
                MemberType = ParseRoleMemberType(action.MemberType)
            };
        }
        return new WindowsModelRoleMember
        {
            MemberName = memberName
        };
    }

    private static bool TryParseDataType(string? value, out DataType dataType)
    {
        dataType = DataType.String;
        if (string.IsNullOrWhiteSpace(value))
        {
            return false;
        }
        return Enum.TryParse(value.Trim(), true, out dataType);
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

    private static ModelPermission ParseModelPermission(string? value)
    {
        if (!string.IsNullOrWhiteSpace(value) &&
            Enum.TryParse<ModelPermission>(value.Trim(), true, out var parsed))
        {
            return parsed;
        }
        return ModelPermission.Read;
    }

    private static MetadataPermission ParseMetadataPermission(string? value)
    {
        if (!string.IsNullOrWhiteSpace(value) &&
            Enum.TryParse<MetadataPermission>(value.Trim(), true, out var parsed))
        {
            return parsed;
        }
        return MetadataPermission.Default;
    }

    private static RoleMemberType ParseRoleMemberType(string? value)
    {
        if (!string.IsNullOrWhiteSpace(value) &&
            Enum.TryParse<RoleMemberType>(value.Trim(), true, out var parsed))
        {
            return parsed;
        }
        return RoleMemberType.Auto;
    }

    private static string NormalizeObjectType(string? value)
    {
        return (value ?? string.Empty).Trim().ToLowerInvariant();
    }

    private static string RequireValue(string? value, string fieldName)
    {
        return string.IsNullOrWhiteSpace(value)
            ? throw new InvalidOperationException($"动作缺少字段: {fieldName}")
            : value.Trim();
    }
}
