using System.IO.Compression;
using System.Text;
using System.Text.RegularExpressions;

namespace PbiMetadataTool;

internal sealed record PowerQueryQueryMetadata(
    string Name,
    string Expression,
    bool IsLoadedToModel,
    bool IsParameter,
    bool IsFunction,
    string SourceSystemType,
    string SourceServer,
    string SourceDatabase,
    string SourceSchema,
    string SourceObjectName);

internal sealed class PowerQueryMetadataReader
{
    public IReadOnlyList<PowerQueryQueryMetadata> TryReadQueries(string? sourcePath, IReadOnlyCollection<string> loadedTableNames)
    {
        if (string.IsNullOrWhiteSpace(sourcePath))
        {
            return Array.Empty<PowerQueryQueryMetadata>();
        }

        try
        {
            var fullPath = Path.GetFullPath(Environment.ExpandEnvironmentVariables(sourcePath.Trim()));
            if (Directory.Exists(fullPath))
            {
                return MergeQueryRows(
                    ParseFromPbipDirectory(fullPath, loadedTableNames),
                    ParseFromWorkspaceDirectory(fullPath, loadedTableNames),
                    ParseFromTmdlDirectory(fullPath, loadedTableNames));
            }

            if (!File.Exists(fullPath))
            {
                return Array.Empty<PowerQueryQueryMetadata>();
            }

            var ext = Path.GetExtension(fullPath).ToLowerInvariant();
            if (ext is ".pbix" or ".pbit")
            {
                return MergeQueryRows(
                    ParseFromPbixArchive(fullPath, loadedTableNames),
                    ParseFromTmdlDirectory(Path.GetDirectoryName(fullPath) ?? string.Empty, loadedTableNames));
            }

            if (ext == ".pbip")
            {
                var projectDir = Path.GetDirectoryName(fullPath);
                if (!string.IsNullOrWhiteSpace(projectDir))
                {
                    return MergeQueryRows(
                        ParseFromPbipDirectory(projectDir, loadedTableNames),
                        ParseFromTmdlDirectory(projectDir, loadedTableNames));
                }
            }
        }
        catch
        {
            // Ignore parse failures and return empty list.
        }

        return Array.Empty<PowerQueryQueryMetadata>();
    }

    private static IReadOnlyList<PowerQueryQueryMetadata> ParseFromPbixArchive(string archivePath, IReadOnlyCollection<string> loadedTableNames)
    {
        using var archive = ZipFile.OpenRead(archivePath);
        var dataMashupEntry = archive.Entries
            .FirstOrDefault(entry => string.Equals(entry.FullName.Replace('\\', '/'), "DataMashup", StringComparison.OrdinalIgnoreCase));
        if (dataMashupEntry is null)
        {
            return Array.Empty<PowerQueryQueryMetadata>();
        }

        using var stream = dataMashupEntry.Open();
        using var ms = new MemoryStream();
        stream.CopyTo(ms);
        var bytes = ms.ToArray();
        return ParseFromDataMashupBytes(bytes, loadedTableNames);
    }

    private static IReadOnlyList<PowerQueryQueryMetadata> ParseFromPbipDirectory(string projectDir, IReadOnlyCollection<string> loadedTableNames)
    {
        var semanticModelDir = Directory.EnumerateDirectories(projectDir, "*.SemanticModel", SearchOption.TopDirectoryOnly)
            .FirstOrDefault();
        if (string.IsNullOrWhiteSpace(semanticModelDir))
        {
            return Array.Empty<PowerQueryQueryMetadata>();
        }

        var sectionCandidates = Directory.EnumerateFiles(semanticModelDir, "*.m", SearchOption.AllDirectories)
            .Where(path =>
                path.EndsWith("Section1.m", StringComparison.OrdinalIgnoreCase) ||
                path.EndsWith("Section.m", StringComparison.OrdinalIgnoreCase) ||
                path.EndsWith("Mashup.m", StringComparison.OrdinalIgnoreCase))
            .OrderBy(path => path.Length)
            .ToList();

        if (sectionCandidates.Count == 0)
        {
            sectionCandidates = Directory.EnumerateFiles(semanticModelDir, "*.m", SearchOption.AllDirectories)
                .OrderBy(path => path.Length)
                .ToList();
        }

        foreach (var path in sectionCandidates)
        {
            try
            {
                var text = File.ReadAllText(path, Encoding.UTF8);
                var rows = ParseQueriesFromSectionM(text, loadedTableNames);
                if (rows.Count > 0)
                {
                    return rows;
                }
            }
            catch
            {
                // Try next candidate.
            }
        }

        return Array.Empty<PowerQueryQueryMetadata>();
    }

    private static IReadOnlyList<PowerQueryQueryMetadata> ParseFromWorkspaceDirectory(string workspaceDir, IReadOnlyCollection<string> loadedTableNames)
    {
        var candidates = new List<string>();
        try
        {
            var dataDir = Path.Combine(workspaceDir, "Data");
            if (Directory.Exists(dataDir))
            {
                candidates.AddRange(Directory.EnumerateFiles(dataDir, "*", SearchOption.TopDirectoryOnly)
                    .Where(path =>
                        string.Equals(Path.GetFileName(path), "DataMashup", StringComparison.OrdinalIgnoreCase) ||
                        path.EndsWith(".mashup", StringComparison.OrdinalIgnoreCase)));
            }

            if (candidates.Count == 0)
            {
                candidates.AddRange(Directory.EnumerateFiles(workspaceDir, "*", SearchOption.AllDirectories)
                    .Where(path =>
                        string.Equals(Path.GetFileName(path), "DataMashup", StringComparison.OrdinalIgnoreCase) ||
                        path.EndsWith(".mashup", StringComparison.OrdinalIgnoreCase))
                    .Take(20));
            }
        }
        catch
        {
            return Array.Empty<PowerQueryQueryMetadata>();
        }

        foreach (var path in candidates.Distinct(StringComparer.OrdinalIgnoreCase))
        {
            try
            {
                var bytes = File.ReadAllBytes(path);
                var rows = ParseFromDataMashupBytes(bytes, loadedTableNames);
                if (rows.Count > 0)
                {
                    return rows;
                }
            }
            catch
            {
                // Try next candidate.
            }
        }

        return Array.Empty<PowerQueryQueryMetadata>();
    }

    private static IReadOnlyList<PowerQueryQueryMetadata> ParseFromTmdlDirectory(string rootDir, IReadOnlyCollection<string> loadedTableNames)
    {
        if (string.IsNullOrWhiteSpace(rootDir) || !Directory.Exists(rootDir))
        {
            return Array.Empty<PowerQueryQueryMetadata>();
        }

        var files = new List<string>();
        try
        {
            var semanticModelDirs = Directory.EnumerateDirectories(rootDir, "*.SemanticModel", SearchOption.TopDirectoryOnly).ToList();
            if (semanticModelDirs.Count > 0)
            {
                foreach (var semanticModelDir in semanticModelDirs)
                {
                    files.AddRange(Directory.EnumerateFiles(semanticModelDir, "*.tmdl", SearchOption.AllDirectories));
                }
            }
            else
            {
                files.AddRange(Directory.EnumerateFiles(rootDir, "*.tmdl", SearchOption.AllDirectories).Take(1200));
            }
        }
        catch
        {
            return Array.Empty<PowerQueryQueryMetadata>();
        }

        if (files.Count == 0)
        {
            return Array.Empty<PowerQueryQueryMetadata>();
        }

        var loadedNames = new HashSet<string>(loadedTableNames ?? Array.Empty<string>(), StringComparer.OrdinalIgnoreCase);
        var rows = new List<PowerQueryQueryMetadata>();
        foreach (var path in files.Distinct(StringComparer.OrdinalIgnoreCase))
        {
            try
            {
                var text = File.ReadAllText(path, Encoding.UTF8);
                rows.AddRange(ParseQueriesFromTmdlText(text, loadedNames));
            }
            catch
            {
                // Skip invalid files.
            }
        }

        return rows
            .GroupBy(q => q.Name, StringComparer.OrdinalIgnoreCase)
            .Select(group => group.OrderByDescending(ScoreQuery).First())
            .OrderBy(q => q.Name, StringComparer.OrdinalIgnoreCase)
            .ToList();
    }

    private static IReadOnlyList<PowerQueryQueryMetadata> ParseFromDataMashupBytes(byte[] bytes, IReadOnlyCollection<string> loadedTableNames)
    {
        var sections = TryExtractSectionMFromMashupBytes(bytes);
        if (sections.Count > 0)
        {
            var parsed = sections
                .Select(section => ParseQueriesFromSectionM(section, loadedTableNames))
                .Where(rows => rows.Count > 0)
                .ToArray();
            if (parsed.Length > 0)
            {
                return MergeQueryRows(parsed);
            }
        }

        var sectionM = TryExtractSectionMFromBinaryText(bytes);
        if (string.IsNullOrWhiteSpace(sectionM))
        {
            return Array.Empty<PowerQueryQueryMetadata>();
        }

        return ParseQueriesFromSectionM(sectionM, loadedTableNames);
    }

    private static IReadOnlyList<string> TryExtractSectionMFromMashupBytes(byte[] mashupBytes)
    {
        var zipStart = FindZipHeaderOffset(mashupBytes);
        if (zipStart < 0)
        {
            return Array.Empty<string>();
        }

        try
        {
            using var zipStream = new MemoryStream(mashupBytes, writable: false);
            using var archive = new ZipArchive(zipStream, ZipArchiveMode.Read, leaveOpen: false);
            var sections = TryReadSectionsFromZipArchive(archive);
            if (sections.Count > 0)
            {
                return sections;
            }
        }
        catch
        {
            // Fallback to offset stream below.
        }

        try
        {
            using var offsetZipStream = new MemoryStream(mashupBytes, zipStart, mashupBytes.Length - zipStart, writable: false);
            using var offsetArchive = new ZipArchive(offsetZipStream, ZipArchiveMode.Read, leaveOpen: false);
            return TryReadSectionsFromZipArchive(offsetArchive);
        }
        catch
        {
            return Array.Empty<string>();
        }
    }

    private static IReadOnlyList<string> TryReadSectionsFromZipArchive(ZipArchive archive)
    {
        var formulaEntries = archive.Entries
            .Where(entry =>
            {
                if (entry.Length <= 0)
                {
                    return false;
                }

                var normalized = entry.FullName.Replace('\\', '/');
                return normalized.EndsWith(".m", StringComparison.OrdinalIgnoreCase) &&
                       (normalized.StartsWith("Formulas/", StringComparison.OrdinalIgnoreCase) ||
                        normalized.Contains("/Formulas/", StringComparison.OrdinalIgnoreCase));
            })
            .OrderBy(entry => GetMashupSectionOrder(entry.FullName))
            .ThenBy(entry => entry.FullName, StringComparer.OrdinalIgnoreCase)
            .ToList();

        if (formulaEntries.Count == 0)
        {
            formulaEntries = archive.Entries
                .Where(entry =>
                {
                    if (entry.Length <= 0)
                    {
                        return false;
                    }

                    var normalized = entry.FullName.Replace('\\', '/');
                    return normalized.EndsWith(".m", StringComparison.OrdinalIgnoreCase);
                })
                .OrderBy(entry => GetMashupSectionOrder(entry.FullName))
                .ThenBy(entry => entry.FullName, StringComparer.OrdinalIgnoreCase)
                .ToList();
        }

        var sections = new List<string>();
        foreach (var entry in formulaEntries)
        {
            try
            {
                using var reader = new StreamReader(entry.Open(), Encoding.UTF8, detectEncodingFromByteOrderMarks: true);
                var text = reader.ReadToEnd();
                if (string.IsNullOrWhiteSpace(text))
                {
                    continue;
                }

                sections.Add(text);
            }
            catch
            {
                // Continue parsing other sections.
            }
        }

        return sections;
    }

    private static int GetMashupSectionOrder(string path)
    {
        var fileName = Path.GetFileName(path);
        if (string.Equals(fileName, "Section1.m", StringComparison.OrdinalIgnoreCase))
        {
            return 0;
        }

        if (string.Equals(fileName, "Section.m", StringComparison.OrdinalIgnoreCase))
        {
            return 1;
        }

        if (fileName.StartsWith("Section", StringComparison.OrdinalIgnoreCase))
        {
            return 2;
        }

        return 10;
    }

    private static string TryExtractSectionMFromBinaryText(byte[] bytes)
    {
        var utf8 = Encoding.UTF8.GetString(bytes);
        var utf16 = Encoding.Unicode.GetString(bytes);

        var candidate = ExtractSectionBlock(utf8);
        if (!string.IsNullOrWhiteSpace(candidate))
        {
            return candidate;
        }

        return ExtractSectionBlock(utf16);
    }

    private static string ExtractSectionBlock(string text)
    {
        if (string.IsNullOrWhiteSpace(text))
        {
            return string.Empty;
        }

        var start = text.IndexOf("section ", StringComparison.OrdinalIgnoreCase);
        if (start < 0)
        {
            return string.Empty;
        }

        var slice = text[start..];
        var match = Regex.Match(
            slice,
            @"(?is)^section\s+[^\r\n;]+;\s*(?:shared\s+.+)$");
        return match.Success ? match.Value : slice;
    }

    private static int FindZipHeaderOffset(byte[] bytes)
    {
        for (var i = 0; i <= bytes.Length - 4; i++)
        {
            if (bytes[i] == 0x50 && bytes[i + 1] == 0x4B && bytes[i + 2] == 0x03 && bytes[i + 3] == 0x04)
            {
                return i;
            }
        }

        return -1;
    }

    private static IReadOnlyList<PowerQueryQueryMetadata> ParseQueriesFromSectionM(string sectionM, IReadOnlyCollection<string> loadedTableNames)
    {
        var loadedNames = new HashSet<string>(loadedTableNames ?? Array.Empty<string>(), StringComparer.OrdinalIgnoreCase);
        var rows = new List<PowerQueryQueryMetadata>();
        if (string.IsNullOrWhiteSpace(sectionM))
        {
            return rows;
        }

        var matches = Regex.Matches(
            sectionM,
            @"(?ms)^\s*shared\s+(?<name>#""(?:[^""]|"""")*""|[^\s=]+)\s*=\s*(?<expr>.*?);\s*(?=^\s*shared\s+|\z)");

        foreach (Match match in matches)
        {
            if (!match.Success)
            {
                continue;
            }

            var rawName = match.Groups["name"].Value.Trim();
            var name = NormalizeQueryName(rawName);
            if (string.IsNullOrWhiteSpace(name))
            {
                continue;
            }

            var expression = match.Groups["expr"].Value.Trim();
            if (string.IsNullOrWhiteSpace(expression))
            {
                continue;
            }

            var isParameter = Regex.IsMatch(expression, @"IsParameterQuery\s*=\s*true", RegexOptions.IgnoreCase);
            var isFunction = expression.StartsWith("(", StringComparison.Ordinal) ||
                             Regex.IsMatch(expression, @"=>");

            var lineage = QuerySourceParser.Parse("PowerQuery", expression, string.Empty);
            rows.Add(new PowerQueryQueryMetadata(
                Name: name,
                Expression: expression,
                IsLoadedToModel: loadedNames.Contains(name),
                IsParameter: isParameter,
                IsFunction: isFunction,
                SourceSystemType: lineage.SystemType,
                SourceServer: lineage.Server,
                SourceDatabase: lineage.Database,
                SourceSchema: lineage.Schema,
                SourceObjectName: lineage.ObjectName));
        }

        return rows
            .GroupBy(q => q.Name, StringComparer.OrdinalIgnoreCase)
            .Select(g => g.First())
            .OrderBy(q => q.Name, StringComparer.OrdinalIgnoreCase)
            .ToList();
    }

    private static IReadOnlyList<PowerQueryQueryMetadata> ParseQueriesFromTmdlText(string text, HashSet<string> loadedNames)
    {
        if (string.IsNullOrWhiteSpace(text))
        {
            return Array.Empty<PowerQueryQueryMetadata>();
        }

        var rows = new List<PowerQueryQueryMetadata>();
        rows.AddRange(ParseQueriesFromTmdlTableBlocks(text, loadedNames));

        var eqMatches = Regex.Matches(
            text,
            @"(?ms)^\s*(?:expression|shared)\s+(?<name>#""(?:[^""]|"""")*""|'[^']+'|""[^""]+""|[^\s=]+)\s*=\s*(?<expr>.*?)(?=^\s*(?:expression|shared|table|role|relationship|culture|perspective|dataSource|model|database|partition)\b|\z)");

        foreach (Match match in eqMatches)
        {
            if (!match.Success)
            {
                continue;
            }

            var name = NormalizeQueryName(match.Groups["name"].Value);
            var expr = match.Groups["expr"].Value.Trim();
            if (string.IsNullOrWhiteSpace(name) || string.IsNullOrWhiteSpace(expr))
            {
                continue;
            }

            if (!LooksLikePowerQueryExpression(expr))
            {
                continue;
            }

            rows.Add(CreateQueryRow(name, expr, loadedNames));
        }

        return rows
            .GroupBy(q => q.Name, StringComparer.OrdinalIgnoreCase)
            .Select(group => group.OrderByDescending(ScoreQuery).First())
            .ToList();
    }

    private static IReadOnlyList<PowerQueryQueryMetadata> ParseQueriesFromTmdlTableBlocks(string text, HashSet<string> loadedNames)
    {
        var lines = text.Replace("\r\n", "\n", StringComparison.Ordinal).Split('\n');
        if (lines.Length == 0)
        {
            return Array.Empty<PowerQueryQueryMetadata>();
        }

        var rows = new List<PowerQueryQueryMetadata>();
        string? currentTableName = null;
        var currentTableIndent = -1;

        for (var i = 0; i < lines.Length; i++)
        {
            var line = lines[i];
            var trimmed = line.Trim();
            if (trimmed.Length == 0)
            {
                continue;
            }

            var indent = GetLeadingWhitespaceWidth(line);
            if (TryMatchTmdlTableHeader(trimmed, out var matchedTableName))
            {
                currentTableName = matchedTableName;
                currentTableIndent = indent;
                continue;
            }

            if (currentTableName is null)
            {
                continue;
            }

            if (indent <= currentTableIndent && IsTmdlTopLevelHeader(trimmed))
            {
                currentTableName = null;
                currentTableIndent = -1;
                continue;
            }

            if (!trimmed.StartsWith("source", StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            if (!TryMatchTmdlSourceLine(trimmed, out var inlineSource))
            {
                continue;
            }

            var sourceIndent = indent;
            var sourceLines = new List<string>();
            if (!string.IsNullOrWhiteSpace(inlineSource))
            {
                sourceLines.Add(inlineSource);
            }

            var cursor = i + 1;
            for (; cursor < lines.Length; cursor++)
            {
                var candidateLine = lines[cursor];
                if (candidateLine.Length == 0)
                {
                    sourceLines.Add(string.Empty);
                    continue;
                }

                var candidateTrimmed = candidateLine.Trim();
                if (candidateTrimmed.Length == 0)
                {
                    sourceLines.Add(string.Empty);
                    continue;
                }

                var candidateIndent = GetLeadingWhitespaceWidth(candidateLine);
                if (candidateIndent <= sourceIndent)
                {
                    break;
                }

                sourceLines.Add(candidateLine);
            }

            i = Math.Max(i, cursor - 1);
            if (sourceLines.Count == 0)
            {
                continue;
            }

            var expression = NormalizeTmdlSourceBlock(sourceLines).Trim();
            if (string.IsNullOrWhiteSpace(expression))
            {
                continue;
            }

            if (!LooksLikePowerQueryExpression(expression))
            {
                continue;
            }

            rows.Add(CreateQueryRow(currentTableName, expression, loadedNames));
        }

        return rows;
    }

    private static bool TryMatchTmdlTableHeader(string trimmedLine, out string tableName)
    {
        var match = Regex.Match(trimmedLine, @"^table\s+(?<name>#""(?:[^""]|"""")*""|'[^']+'|""[^""]+""|[^\r\n=]+?)(?:\s*=.*)?$", RegexOptions.IgnoreCase);
        if (match.Success)
        {
            tableName = NormalizeQueryName(match.Groups["name"].Value);
            return !string.IsNullOrWhiteSpace(tableName);
        }

        tableName = string.Empty;
        return false;
    }

    private static bool TryMatchTmdlSourceLine(string trimmedLine, out string inlineSource)
    {
        var match = Regex.Match(trimmedLine, @"^source\s*(?:=|:)\s*(?<expr>.*)$", RegexOptions.IgnoreCase);
        if (!match.Success)
        {
            inlineSource = string.Empty;
            return false;
        }

        inlineSource = match.Groups["expr"].Value.Trim();
        return true;
    }

    private static bool IsTmdlTopLevelHeader(string trimmedLine)
    {
        return Regex.IsMatch(trimmedLine, @"^(table|role|relationship|culture|perspective|dataSource|model|database|expression|shared)\b", RegexOptions.IgnoreCase);
    }

    private static string NormalizeTmdlSourceBlock(List<string> lines)
    {
        if (lines.Count == 0)
        {
            return string.Empty;
        }

        var start = 0;
        while (start < lines.Count && string.IsNullOrWhiteSpace(lines[start]))
        {
            start++;
        }

        if (start >= lines.Count)
        {
            return string.Empty;
        }

        var normalized = lines.Skip(start).ToList();
        var minIndent = int.MaxValue;
        for (var i = 0; i < normalized.Count; i++)
        {
            var line = normalized[i];
            if (string.IsNullOrWhiteSpace(line))
            {
                continue;
            }

            minIndent = Math.Min(minIndent, GetLeadingWhitespaceWidth(line));
        }

        if (minIndent == int.MaxValue)
        {
            minIndent = 0;
        }

        for (var i = 0; i < normalized.Count; i++)
        {
            var line = normalized[i];
            if (string.IsNullOrWhiteSpace(line))
            {
                normalized[i] = string.Empty;
                continue;
            }

            var remove = Math.Min(minIndent, line.Length);
            normalized[i] = line[remove..].TrimEnd();
        }

        return string.Join("\n", normalized).TrimEnd();
    }

    private static int GetLeadingWhitespaceWidth(string text)
    {
        var count = 0;
        while (count < text.Length && char.IsWhiteSpace(text[count]))
        {
            count++;
        }

        return count;
    }

    private static PowerQueryQueryMetadata CreateQueryRow(string name, string expression, HashSet<string> loadedNames)
    {
        var isParameter = Regex.IsMatch(expression, @"IsParameterQuery\s*=\s*true", RegexOptions.IgnoreCase);
        var isFunction = expression.StartsWith("(", StringComparison.Ordinal) || Regex.IsMatch(expression, @"=>");
        var lineage = QuerySourceParser.Parse("PowerQuery", expression, string.Empty);
        return new PowerQueryQueryMetadata(
            Name: name,
            Expression: expression,
            IsLoadedToModel: loadedNames.Contains(name),
            IsParameter: isParameter,
            IsFunction: isFunction,
            SourceSystemType: lineage.SystemType,
            SourceServer: lineage.Server,
            SourceDatabase: lineage.Database,
            SourceSchema: lineage.Schema,
            SourceObjectName: lineage.ObjectName);
    }

    private static bool LooksLikePowerQueryExpression(string expression)
    {
        if (string.IsNullOrWhiteSpace(expression))
        {
            return false;
        }

        return Regex.IsMatch(expression, @"(?is)\blet\b.+\bin\b") ||
               expression.Contains("Sql.Database", StringComparison.OrdinalIgnoreCase) ||
               expression.Contains("PostgreSQL.Database", StringComparison.OrdinalIgnoreCase) ||
               expression.Contains("MySQL.Database", StringComparison.OrdinalIgnoreCase) ||
               expression.Contains("Odbc.DataSource", StringComparison.OrdinalIgnoreCase) ||
               expression.Contains("Web.Contents", StringComparison.OrdinalIgnoreCase) ||
               expression.Contains("Table.", StringComparison.OrdinalIgnoreCase) ||
               expression.Contains("#table", StringComparison.OrdinalIgnoreCase);
    }

    private static IReadOnlyList<PowerQueryQueryMetadata> MergeQueryRows(params IReadOnlyList<PowerQueryQueryMetadata>[] sources)
    {
        var merged = new Dictionary<string, PowerQueryQueryMetadata>(StringComparer.OrdinalIgnoreCase);
        foreach (var source in sources)
        {
            if (source is null || source.Count == 0)
            {
                continue;
            }

            foreach (var item in source)
            {
                if (!merged.TryGetValue(item.Name, out var current))
                {
                    merged[item.Name] = item;
                    continue;
                }

                if (ScoreQuery(item) > ScoreQuery(current))
                {
                    merged[item.Name] = item;
                }
            }
        }

        return merged.Values
            .OrderBy(q => q.Name, StringComparer.OrdinalIgnoreCase)
            .ToList();
    }

    private static int ScoreQuery(PowerQueryQueryMetadata query)
    {
        var score = 0;
        if (!string.IsNullOrWhiteSpace(query.Expression))
        {
            score += Math.Min(4000, query.Expression.Length);
        }
        if (query.IsLoadedToModel)
        {
            score += 200;
        }
        if (!string.IsNullOrWhiteSpace(query.SourceSystemType) && !string.Equals(query.SourceSystemType, "Unknown", StringComparison.OrdinalIgnoreCase))
        {
            score += 100;
        }
        return score;
    }

    private static string NormalizeQueryName(string rawName)
    {
        if (string.IsNullOrWhiteSpace(rawName))
        {
            return string.Empty;
        }

        var name = rawName.Trim();
        if (name.StartsWith("#\"", StringComparison.Ordinal) && name.EndsWith("\"", StringComparison.Ordinal) && name.Length >= 3)
        {
            name = name[2..^1].Replace("\"\"", "\"", StringComparison.Ordinal);
        }
        else if ((name.StartsWith("\"", StringComparison.Ordinal) && name.EndsWith("\"", StringComparison.Ordinal)) ||
                 (name.StartsWith("'", StringComparison.Ordinal) && name.EndsWith("'", StringComparison.Ordinal)))
        {
            name = name[1..^1];
        }

        return name.Trim();
    }
}
