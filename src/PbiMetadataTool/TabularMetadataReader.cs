using Microsoft.AnalysisServices.Tabular;

namespace PbiMetadataTool;

internal sealed class TabularMetadataReader
{
    public ModelMetadata ReadMetadata(int port, string? databaseName)
        => ReadMetadata($"DataSource=localhost:{port};", databaseName);

    public ModelMetadata ReadMetadata(string connectionString, string? databaseName)
    {
        var server = new Server();
        server.Connect(connectionString);

        try
        {
            var database = ResolveDatabase(server, databaseName);
            if (database is null)
            {
                throw new InvalidOperationException("未找到可用的 Tabular 数据库。");
            }

            var tables = database.Model.Tables
                .Select(table => new TableMetadata(
                    Name: table.Name,
                    IsHidden: table.IsHidden,
                    Columns: table.Columns
                        .Select(column => new ColumnMetadata(
                            Name: column.Name,
                            ColumnType: column.Type.ToString(),
                            DataType: column.DataType.ToString(),
                            IsHidden: column.IsHidden,
                            Expression: column is CalculatedColumn calculatedColumn
                                ? (calculatedColumn.Expression ?? string.Empty)
                                : string.Empty))
                        .ToList(),
                    Measures: table.Measures
                        .Select(measure => new MeasureMetadata(
                            Name: measure.Name,
                            Expression: measure.Expression ?? string.Empty,
                            FormatString: measure.FormatString ?? string.Empty,
                            IsHidden: measure.IsHidden,
                            DisplayFolder: measure.DisplayFolder ?? string.Empty))
                        .ToList(),
                    TableType: ResolveTableType(table),
                    Expression: ResolveCalculatedTableExpression(table)))
                .ToList();

            var relationships = database.Model.Relationships
                .OfType<SingleColumnRelationship>()
                .Select(relationship => new RelationshipMetadata(
                    Name: relationship.Name,
                    FromTable: relationship.FromColumn?.Table?.Name ?? string.Empty,
                    FromColumn: relationship.FromColumn?.Name ?? string.Empty,
                    ToTable: relationship.ToColumn?.Table?.Name ?? string.Empty,
                    ToColumn: relationship.ToColumn?.Name ?? string.Empty,
                    CrossFilterDirection: relationship.CrossFilteringBehavior.ToString(),
                    IsActive: relationship.IsActive))
                .ToList();

            var roles = database.Model.Roles
                .Cast<object>()
                .Select(ReadRoleMetadata)
                .OrderBy(role => role.Name, StringComparer.OrdinalIgnoreCase)
                .ToList();

            return new ModelMetadata(
                DatabaseName: database.Name,
                DatabaseId: database.ID,
                CompatibilityLevel: database.CompatibilityLevel,
                Tables: tables,
                Relationships: relationships,
                Roles: roles);
        }
        finally
        {
            if (server.Connected)
            {
                server.Disconnect();
            }
        }
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

    private static RoleMetadata ReadRoleMetadata(object roleObject)
    {
        var members = GetEnumerableProperty(roleObject, "Members")
            .Select(ReadRoleMemberMetadata)
            .ToList();

        var tablePermissions = GetEnumerableProperty(roleObject, "TablePermissions")
            .Select(ReadRoleTablePermissionMetadata)
            .ToList();

        return new RoleMetadata(
            Name: GetStringProperty(roleObject, "Name"),
            Description: GetStringProperty(roleObject, "Description"),
            ModelPermission: GetStringProperty(roleObject, "ModelPermission"),
            Members: members,
            TablePermissions: tablePermissions);
    }

    private static RoleMemberMetadata ReadRoleMemberMetadata(object memberObject)
    {
        return new RoleMemberMetadata(
            Name: FirstNonEmpty(
                GetStringProperty(memberObject, "MemberName"),
                GetStringProperty(memberObject, "Name"),
                GetStringProperty(memberObject, "MemberId")),
            IdentityProvider: GetStringProperty(memberObject, "IdentityProvider"),
            MemberType: FirstNonEmpty(
                GetStringProperty(memberObject, "MemberType"),
                memberObject.GetType().Name));
    }

    private static RoleTablePermissionMetadata ReadRoleTablePermissionMetadata(object permissionObject)
    {
        var tableObject = GetObjectProperty(permissionObject, "Table");
        return new RoleTablePermissionMetadata(
            TableName: FirstNonEmpty(
                GetStringProperty(tableObject, "Name"),
                GetStringProperty(permissionObject, "Name"),
                "(未知表)"),
            FilterExpression: GetStringProperty(permissionObject, "FilterExpression"),
            MetadataPermission: GetStringProperty(permissionObject, "MetadataPermission"));
    }

    private static IEnumerable<object> GetEnumerableProperty(object target, string propertyName)
    {
        if (target is null)
        {
            return Enumerable.Empty<object>();
        }

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

    private static string ResolveTableType(Table table)
    {
        if (table.Partitions.Count == 0)
        {
            return "Unknown";
        }

        var source = table.Partitions[0].Source;
        return source switch
        {
            CalculatedPartitionSource => "Calculated",
            QueryPartitionSource => "Import",
            MPartitionSource => "PowerQuery",
            EntityPartitionSource => "Entity",
            _ => source?.GetType().Name ?? "Unknown"
        };
    }

    private static string ResolveCalculatedTableExpression(Table table)
    {
        foreach (var partition in table.Partitions)
        {
            if (partition.Source is CalculatedPartitionSource calculatedSource)
            {
                return calculatedSource.Expression ?? string.Empty;
            }
        }

        return string.Empty;
    }
}
