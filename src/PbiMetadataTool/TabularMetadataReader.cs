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
                            IsHidden: column.IsHidden))
                        .ToList(),
                    Measures: table.Measures
                        .Select(measure => new MeasureMetadata(
                            Name: measure.Name,
                            Expression: measure.Expression ?? string.Empty,
                            FormatString: measure.FormatString ?? string.Empty,
                            IsHidden: measure.IsHidden,
                            DisplayFolder: measure.DisplayFolder ?? string.Empty))
                        .ToList()))
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

            return new ModelMetadata(
                DatabaseName: database.Name,
                DatabaseId: database.ID,
                CompatibilityLevel: database.CompatibilityLevel,
                Tables: tables,
                Relationships: relationships);
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
}
