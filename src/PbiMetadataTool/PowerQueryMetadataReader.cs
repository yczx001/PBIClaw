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
                return ParseFromPbipDirectory(fullPath, loadedTableNames);
            }

            if (!File.Exists(fullPath))
            {
                return Array.Empty<PowerQueryQueryMetadata>();
            }

            var ext = Path.GetExtension(fullPath).ToLowerInvariant();
            if (ext is ".pbix" or ".pbit")
            {
                return ParseFromPbixArchive(fullPath, loadedTableNames);
            }

            if (ext == ".pbip")
            {
                var projectDir = Path.GetDirectoryName(fullPath);
                if (!string.IsNullOrWhiteSpace(projectDir))
                {
                    return ParseFromPbipDirectory(projectDir, loadedTableNames);
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

        var sectionM = TryExtractSectionMFromMashupBytes(bytes);
        if (string.IsNullOrWhiteSpace(sectionM))
        {
            sectionM = TryExtractSectionMFromBinaryText(bytes);
        }

        if (string.IsNullOrWhiteSpace(sectionM))
        {
            return Array.Empty<PowerQueryQueryMetadata>();
        }

        return ParseQueriesFromSectionM(sectionM, loadedTableNames);
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

    private static string TryExtractSectionMFromMashupBytes(byte[] mashupBytes)
    {
        var zipStart = FindZipHeaderOffset(mashupBytes);
        if (zipStart < 0)
        {
            return string.Empty;
        }

        try
        {
            using var zipStream = new MemoryStream(mashupBytes, zipStart, mashupBytes.Length - zipStart, writable: false);
            using var archive = new ZipArchive(zipStream, ZipArchiveMode.Read, leaveOpen: false);
            var sectionEntry = archive.Entries
                .FirstOrDefault(entry =>
                    string.Equals(entry.FullName.Replace('\\', '/'), "Formulas/Section1.m", StringComparison.OrdinalIgnoreCase) ||
                    entry.FullName.EndsWith("/Section1.m", StringComparison.OrdinalIgnoreCase));

            if (sectionEntry is null)
            {
                return string.Empty;
            }

            using var reader = new StreamReader(sectionEntry.Open(), Encoding.UTF8, detectEncodingFromByteOrderMarks: true);
            return reader.ReadToEnd();
        }
        catch
        {
            return string.Empty;
        }
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

        return name.Trim();
    }
}
