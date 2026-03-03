using System.Text.RegularExpressions;

namespace PbiMetadataTool;

internal sealed record SourceLineageInfo(
    string SystemType,
    string Server,
    string Database,
    string Schema,
    string ObjectName);

internal static class QuerySourceParser
{
    public static SourceLineageInfo Parse(string sourceType, string sourceExpression, string dataSourceName)
    {
        var expression = sourceExpression ?? string.Empty;
        var systemType = ResolveSystemType(sourceType, expression, dataSourceName);

        var server = FirstMatch(expression, [
            @"Sql\.Database\s*\(\s*""(?<v>[^""]+)""\s*,\s*""[^""]+""",
            @"Sql\.Databases\s*\(\s*""(?<v>[^""]+)""",
            @"PostgreSQL\.Database\s*\(\s*""(?<v>[^""]+)""\s*,\s*""[^""]+""",
            @"MySQL\.Database\s*\(\s*""(?<v>[^""]+)""\s*,\s*""[^""]+""",
            @"Oracle\.Database\s*\(\s*""(?<v>[^""]+)""",
            @"SapHana\.Database\s*\(\s*""(?<v>[^""]+)""",
            @"Snowflake\.Databases\s*\(\s*""(?<v>[^""]+)""",
            @"Odbc\.DataSource\s*\(\s*""(?<v>[^""]+)""",
            @"OleDb\.DataSource\s*\(\s*""(?<v>[^""]+)"""
        ]);

        var database = FirstMatch(expression, [
            @"Sql\.Database\s*\(\s*""[^""]+""\s*,\s*""(?<v>[^""]+)""",
            @"PostgreSQL\.Database\s*\(\s*""[^""]+""\s*,\s*""(?<v>[^""]+)""",
            @"MySQL\.Database\s*\(\s*""[^""]+""\s*,\s*""(?<v>[^""]+)""",
            @"\[\s*Name\s*=\s*""(?<v>[^""]+)""\s*,\s*Kind\s*=\s*""Database""\s*\]",
            @"\[\s*Kind\s*=\s*""Database""\s*,\s*Name\s*=\s*""(?<v>[^""]+)""\s*\]"
        ]);

        var schema = FirstMatch(expression, [
            @"\[\s*Schema\s*=\s*""(?<v>[^""]+)""\s*,\s*Item\s*=\s*""[^""]+""\s*\]",
            @"\[\s*Item\s*=\s*""[^""]+""\s*,\s*Schema\s*=\s*""(?<v>[^""]+)""\s*\]",
            @"\[\s*Name\s*=\s*""(?<v>[^""]+)""\s*,\s*Kind\s*=\s*""Schema""\s*\]"
        ]);

        var objectName = FirstMatch(expression, [
            @"\[\s*Schema\s*=\s*""[^""]+""\s*,\s*Item\s*=\s*""(?<v>[^""]+)""\s*\]",
            @"\[\s*Item\s*=\s*""(?<v>[^""]+)""\s*,\s*Schema\s*=\s*""[^""]+""\s*\]",
            @"\[\s*Name\s*=\s*""(?<v>[^""]+)""\s*,\s*Kind\s*=\s*""Table""\s*\]",
            @"\[\s*Kind\s*=\s*""Table""\s*,\s*Name\s*=\s*""(?<v>[^""]+)""\s*\]",
            @"\[\s*Name\s*=\s*""(?<v>[^""]+)""\s*,\s*Kind\s*=\s*""View""\s*\]",
            @"\[\s*Kind\s*=\s*""View""\s*,\s*Name\s*=\s*""(?<v>[^""]+)""\s*\]"
        ]);

        if (string.IsNullOrWhiteSpace(objectName))
        {
            var sqlSchema = FirstMatch(expression, [
                @"\bFROM\s+\[(?<v>[^\]]+)\]\.\[[^\]]+\]",
                @"\bFROM\s+(?<v>[A-Za-z_][\w$]*)\.[A-Za-z_][\w$]*"
            ]);
            if (!string.IsNullOrWhiteSpace(sqlSchema))
            {
                schema = FirstNonEmpty(schema, sqlSchema);
            }

            objectName = FirstMatch(expression, [
                @"\bFROM\s+\[[^\]]+\]\.\[(?<v>[^\]]+)\]",
                @"\bFROM\s+\[(?<v>[^\]]+)\]",
                @"\bFROM\s+[A-Za-z_][\w$]*\.(?<v>[A-Za-z_][\w$]*)",
                @"\bFROM\s+(?<v>[A-Za-z_][\w$]*)"
            ]);
        }

        return new SourceLineageInfo(
            SystemType: systemType,
            Server: server,
            Database: database,
            Schema: schema,
            ObjectName: objectName);
    }

    private static string ResolveSystemType(string sourceType, string expression, string dataSourceName)
    {
        if (ContainsCall(expression, "PostgreSQL.Database"))
            return "PostgreSQL";
        if (ContainsCall(expression, "MySQL.Database"))
            return "MySQL";
        if (ContainsCall(expression, "Sql.Database") || ContainsCall(expression, "Sql.Databases"))
            return "SQL Server";
        if (ContainsCall(expression, "Oracle.Database"))
            return "Oracle";
        if (ContainsCall(expression, "Snowflake.Databases"))
            return "Snowflake";
        if (ContainsCall(expression, "SapHana.Database"))
            return "SAP HANA";
        if (ContainsCall(expression, "Odbc.DataSource"))
            return "ODBC";
        if (ContainsCall(expression, "OleDb.DataSource"))
            return "OLE DB";
        if (ContainsCall(expression, "AnalysisServices.Database"))
            return "Analysis Services";
        if (ContainsCall(expression, "Excel.Workbook"))
            return "Excel";
        if (ContainsCall(expression, "Csv.Document"))
            return "CSV";
        if (ContainsCall(expression, "SharePoint.Contents") || ContainsCall(expression, "SharePoint.Files"))
            return "SharePoint";
        if (ContainsCall(expression, "Web.Contents"))
            return "Web API";

        if (!string.IsNullOrWhiteSpace(dataSourceName))
        {
            if (Contains(dataSourceName, "postgres")) return "PostgreSQL";
            if (Contains(dataSourceName, "mysql")) return "MySQL";
            if (Contains(dataSourceName, "sql")) return "SQL Server";
            if (Contains(dataSourceName, "oracle")) return "Oracle";
            if (Contains(dataSourceName, "snowflake")) return "Snowflake";
            if (Contains(dataSourceName, "odbc")) return "ODBC";
        }

        if (!string.IsNullOrWhiteSpace(sourceType))
        {
            return sourceType;
        }

        return "Unknown";
    }

    private static string FirstMatch(string text, IEnumerable<string> patterns)
    {
        foreach (var pattern in patterns)
        {
            var match = Regex.Match(text, pattern, RegexOptions.IgnoreCase | RegexOptions.Singleline);
            if (!match.Success)
            {
                continue;
            }

            var value = match.Groups["v"].Value.Trim();
            if (!string.IsNullOrWhiteSpace(value))
            {
                return value;
            }
        }

        return string.Empty;
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

    private static bool Contains(string text, string keyword)
        => text.Contains(keyword, StringComparison.OrdinalIgnoreCase);

    private static bool ContainsCall(string text, string functionName)
    {
        if (string.IsNullOrWhiteSpace(text) || string.IsNullOrWhiteSpace(functionName))
        {
            return false;
        }

        var escaped = Regex.Escape(functionName);
        return Regex.IsMatch(
            text,
            $@"(?<![A-Za-z0-9_]){escaped}\s*\(",
            RegexOptions.IgnoreCase | RegexOptions.Singleline);
    }
}
