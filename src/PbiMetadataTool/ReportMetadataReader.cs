using System.IO.Compression;
using System.Text;
using System.Text.Json;

namespace PbiMetadataTool;

internal sealed class ReportMetadataReader
{
    public ReportMetadata? TryRead(string? sourcePath)
    {
        if (string.IsNullOrWhiteSpace(sourcePath))
        {
            return null;
        }

        try
        {
            var fullPath = Path.GetFullPath(Environment.ExpandEnvironmentVariables(sourcePath.Trim()));
            if (Directory.Exists(fullPath))
            {
                return TryReadFromPbipDirectory(fullPath, fullPath);
            }

            if (!File.Exists(fullPath))
            {
                return null;
            }

            var ext = Path.GetExtension(fullPath).ToLowerInvariant();
            return ext switch
            {
                ".pbix" or ".pbit" => TryReadFromPbixArchive(fullPath),
                ".pbip" => TryReadFromPbipFile(fullPath),
                _ => null
            };
        }
        catch
        {
            return null;
        }
    }

    private static ReportMetadata? TryReadFromPbixArchive(string archivePath)
    {
        using var archive = ZipFile.OpenRead(archivePath);
        if (TryReadLegacyLayoutFromArchive(archive, out var legacyPages))
        {
            return new ReportMetadata("PBIX", archivePath, legacyPages);
        }

        if (TryReadDefinitionFromArchive(archive, out var definitionPages))
        {
            return new ReportMetadata("PBIX", archivePath, definitionPages);
        }

        return null;
    }

    private static ReportMetadata? TryReadFromPbipFile(string pbipPath)
    {
        var projectDir = Path.GetDirectoryName(pbipPath);
        if (string.IsNullOrWhiteSpace(projectDir) || !Directory.Exists(projectDir))
        {
            return null;
        }

        return TryReadFromPbipDirectory(projectDir, pbipPath);
    }

    private static ReportMetadata? TryReadFromPbipDirectory(string projectDir, string sourcePath)
    {
        var reportDir = Directory.EnumerateDirectories(projectDir, "*.Report", SearchOption.TopDirectoryOnly)
            .FirstOrDefault();
        if (string.IsNullOrWhiteSpace(reportDir))
        {
            return null;
        }

        if (TryReadDefinitionFromFolder(reportDir, out var definitionPages))
        {
            return new ReportMetadata("PBIP", sourcePath, definitionPages);
        }

        var reportJsonPath = Directory.EnumerateFiles(reportDir, "report.json", SearchOption.AllDirectories)
            .OrderBy(path => path.Length)
            .FirstOrDefault();

        if (string.IsNullOrWhiteSpace(reportJsonPath))
        {
            return null;
        }

        var json = File.ReadAllText(reportJsonPath, Encoding.UTF8);
        var pages = ParsePagesFromLayoutJson(json);
        return pages.Count == 0 ? null : new ReportMetadata("PBIP", sourcePath, pages);
    }

    private static IReadOnlyList<ReportPageMetadata> ParsePagesFromLayoutJson(string json)
    {
        using var doc = JsonDocument.Parse(json);
        if (!doc.RootElement.TryGetProperty("sections", out var sections) || sections.ValueKind != JsonValueKind.Array)
        {
            return Array.Empty<ReportPageMetadata>();
        }

        var pages = new List<ReportPageMetadata>();
        foreach (var section in sections.EnumerateArray())
        {
            var pageName = GetString(section, "name");
            var displayName = GetString(section, "displayName");
            var visualContainers = section.TryGetProperty("visualContainers", out var visualsNode) && visualsNode.ValueKind == JsonValueKind.Array
                ? visualsNode
                : default;

            var visuals = new List<ReportVisualMetadata>();
            if (visualContainers.ValueKind == JsonValueKind.Array)
            {
                foreach (var container in visualContainers.EnumerateArray())
                {
                    var visualName = GetString(container, "name");
                    var configJson = GetString(container, "config");
                    var (visualType, title) = ParseVisualInfo(configJson);
                    visuals.Add(new ReportVisualMetadata(
                        Name: visualName,
                        VisualType: visualType,
                        Title: title));
                }
            }

            pages.Add(new ReportPageMetadata(
                Name: pageName,
                DisplayName: displayName,
                Visuals: visuals));
        }

        return pages;
    }

    private static bool TryReadLegacyLayoutFromArchive(ZipArchive archive, out IReadOnlyList<ReportPageMetadata> pages)
    {
        pages = Array.Empty<ReportPageMetadata>();
        var entry = FindArchiveEntry(archive, "Report/Layout");
        if (entry is null)
        {
            return false;
        }

        using var stream = entry.Open();
        var json = ReadLayoutText(stream);
        pages = ParsePagesFromLayoutJson(json);
        return pages.Count > 0;
    }

    private static bool TryReadDefinitionFromArchive(ZipArchive archive, out IReadOnlyList<ReportPageMetadata> pages)
    {
        pages = Array.Empty<ReportPageMetadata>();
        var pagesIndex = FindArchiveEntry(archive, "Report/definition/pages/pages.json");
        if (pagesIndex is null)
        {
            return false;
        }

        var pageIds = ParsePageOrderFromPagesMetadata(ReadZipEntryText(pagesIndex));
        if (pageIds.Count == 0)
        {
            var prefix = "report/definition/pages/";
            pageIds = archive.Entries
                .Select(e => e.FullName.Replace('\\', '/'))
                .Where(path =>
                    path.StartsWith(prefix, StringComparison.OrdinalIgnoreCase) &&
                    path.EndsWith("/page.json", StringComparison.OrdinalIgnoreCase))
                .Select(path => path[prefix.Length..(path.Length - "/page.json".Length)])
                .Where(id => !string.IsNullOrWhiteSpace(id))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .OrderBy(id => id, StringComparer.OrdinalIgnoreCase)
                .ToList();
        }

        var parsedPages = new List<ReportPageMetadata>();
        foreach (var pageId in pageIds)
        {
            var pageEntry = FindArchiveEntry(archive, $"Report/definition/pages/{pageId}/page.json");
            if (pageEntry is null)
            {
                continue;
            }

            var pageJson = ReadZipEntryText(pageEntry);
            var page = ParseDefinitionPage(pageJson, pageId);
            if (page is null)
            {
                continue;
            }

            var visuals = ParseDefinitionVisualsFromArchive(archive, pageId);
            parsedPages.Add(page with { Visuals = visuals });
        }

        pages = parsedPages;
        return parsedPages.Count > 0;
    }

    private static bool TryReadDefinitionFromFolder(string reportDir, out IReadOnlyList<ReportPageMetadata> pages)
    {
        pages = Array.Empty<ReportPageMetadata>();
        var pagesMetaPath = Path.Combine(reportDir, "definition", "pages", "pages.json");
        if (!File.Exists(pagesMetaPath))
        {
            return false;
        }

        var pageIds = ParsePageOrderFromPagesMetadata(File.ReadAllText(pagesMetaPath, Encoding.UTF8));
        var pagesRoot = Path.Combine(reportDir, "definition", "pages");
        if (pageIds.Count == 0 && Directory.Exists(pagesRoot))
        {
            pageIds = Directory.EnumerateDirectories(pagesRoot, "*", SearchOption.TopDirectoryOnly)
                .Select(Path.GetFileName)
                .Where(name => !string.IsNullOrWhiteSpace(name))
                .OrderBy(name => name, StringComparer.OrdinalIgnoreCase)
                .ToList()!;
        }

        var parsedPages = new List<ReportPageMetadata>();
        foreach (var pageId in pageIds)
        {
            if (string.IsNullOrWhiteSpace(pageId))
            {
                continue;
            }

            var pagePath = Path.Combine(pagesRoot, pageId, "page.json");
            if (!File.Exists(pagePath))
            {
                continue;
            }

            var page = ParseDefinitionPage(File.ReadAllText(pagePath, Encoding.UTF8), pageId);
            if (page is null)
            {
                continue;
            }

            var visuals = ParseDefinitionVisualsFromFolder(pagesRoot, pageId);
            parsedPages.Add(page with { Visuals = visuals });
        }

        pages = parsedPages;
        return parsedPages.Count > 0;
    }

    private static List<string> ParsePageOrderFromPagesMetadata(string json)
    {
        try
        {
            using var doc = JsonDocument.Parse(json);
            if (!doc.RootElement.TryGetProperty("pageOrder", out var pageOrder) || pageOrder.ValueKind != JsonValueKind.Array)
            {
                return [];
            }

            return pageOrder
                .EnumerateArray()
                .Where(node => node.ValueKind == JsonValueKind.String)
                .Select(node => node.GetString())
                .Where(name => !string.IsNullOrWhiteSpace(name))
                .Select(name => name!)
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();
        }
        catch
        {
            return [];
        }
    }

    private static ReportPageMetadata? ParseDefinitionPage(string pageJson, string fallbackPageName)
    {
        try
        {
            using var doc = JsonDocument.Parse(pageJson);
            var root = doc.RootElement;
            var name = GetString(root, "name");
            if (string.IsNullOrWhiteSpace(name))
            {
                name = fallbackPageName;
            }

            var displayName = GetString(root, "displayName");
            if (string.IsNullOrWhiteSpace(displayName))
            {
                displayName = name;
            }

            return new ReportPageMetadata(
                Name: name,
                DisplayName: displayName,
                Visuals: []);
        }
        catch
        {
            return null;
        }
    }

    private static IReadOnlyList<ReportVisualMetadata> ParseDefinitionVisualsFromArchive(ZipArchive archive, string pageId)
    {
        var prefix = $"report/definition/pages/{pageId}/visuals/";
        var entries = archive.Entries
            .Where(e =>
            {
                var normalized = e.FullName.Replace('\\', '/');
                return normalized.StartsWith(prefix, StringComparison.OrdinalIgnoreCase)
                    && normalized.EndsWith("/visual.json", StringComparison.OrdinalIgnoreCase);
            })
            .OrderBy(e => e.FullName, StringComparer.OrdinalIgnoreCase)
            .ToList();

        var visuals = new List<ReportVisualMetadata>();
        foreach (var entry in entries)
        {
            var visual = ParseDefinitionVisual(ReadZipEntryText(entry));
            if (visual is not null)
            {
                visuals.Add(visual);
            }
        }

        return visuals;
    }

    private static IReadOnlyList<ReportVisualMetadata> ParseDefinitionVisualsFromFolder(string pagesRoot, string pageId)
    {
        var visualsRoot = Path.Combine(pagesRoot, pageId, "visuals");
        if (!Directory.Exists(visualsRoot))
        {
            return [];
        }

        var visuals = new List<ReportVisualMetadata>();
        foreach (var path in Directory.EnumerateFiles(visualsRoot, "visual.json", SearchOption.AllDirectories)
                     .OrderBy(p => p, StringComparer.OrdinalIgnoreCase))
        {
            var visual = ParseDefinitionVisual(File.ReadAllText(path, Encoding.UTF8));
            if (visual is not null)
            {
                visuals.Add(visual);
            }
        }

        return visuals;
    }

    private static ReportVisualMetadata? ParseDefinitionVisual(string visualJson)
    {
        try
        {
            using var doc = JsonDocument.Parse(visualJson);
            var root = doc.RootElement;
            var visualName = GetString(root, "name");
            var visualNode = root.TryGetProperty("visual", out var vNode) && vNode.ValueKind == JsonValueKind.Object
                ? vNode
                : default;
            var visualType = GetString(visualNode, "visualType");
            if (string.IsNullOrWhiteSpace(visualType))
            {
                visualType = "unknown";
            }

            var title = ExtractTitleFromDefinitionVisual(visualNode);
            return new ReportVisualMetadata(
                Name: visualName,
                VisualType: visualType,
                Title: title);
        }
        catch
        {
            return null;
        }
    }

    private static string ExtractTitleFromDefinitionVisual(JsonElement visualNode)
    {
        if (visualNode.ValueKind != JsonValueKind.Object)
        {
            return string.Empty;
        }

        if (!visualNode.TryGetProperty("objects", out var objectsNode) || objectsNode.ValueKind != JsonValueKind.Object)
        {
            return string.Empty;
        }

        var title = ExtractTitleFromObjectsArray(objectsNode, "title");
        if (!string.IsNullOrWhiteSpace(title))
        {
            return title;
        }

        return ExtractTitleFromObjectsArray(objectsNode, "text");
    }

    private static string ExtractTitleFromObjectsArray(JsonElement objectsNode, string key)
    {
        if (!objectsNode.TryGetProperty(key, out var arrayNode) || arrayNode.ValueKind != JsonValueKind.Array)
        {
            return string.Empty;
        }

        foreach (var item in arrayNode.EnumerateArray())
        {
            if (!item.TryGetProperty("properties", out var props) || props.ValueKind != JsonValueKind.Object)
            {
                continue;
            }

            if (!props.TryGetProperty("text", out var textNode) || textNode.ValueKind != JsonValueKind.Object)
            {
                continue;
            }

            var literal = TryExtractLiteralText(textNode);
            if (!string.IsNullOrWhiteSpace(literal))
            {
                return literal;
            }
        }

        return string.Empty;
    }

    private static string TryExtractLiteralText(JsonElement textNode)
    {
        if (!textNode.TryGetProperty("expr", out var exprNode) || exprNode.ValueKind != JsonValueKind.Object)
        {
            return string.Empty;
        }

        if (!exprNode.TryGetProperty("Literal", out var literal) || literal.ValueKind != JsonValueKind.Object)
        {
            return string.Empty;
        }

        var raw = GetString(literal, "Value");
        return string.IsNullOrWhiteSpace(raw) ? string.Empty : raw.Trim().Trim('\'');
    }

    private static ZipArchiveEntry? FindArchiveEntry(ZipArchive archive, string path)
    {
        var normalizedPath = path.Replace('\\', '/');
        return archive.GetEntry(normalizedPath)
            ?? archive.Entries.FirstOrDefault(e =>
                string.Equals(
                    e.FullName.Replace('\\', '/'),
                    normalizedPath,
                    StringComparison.OrdinalIgnoreCase));
    }

    private static string ReadZipEntryText(ZipArchiveEntry entry)
    {
        using var stream = entry.Open();
        using var reader = new StreamReader(stream, Encoding.UTF8);
        return reader.ReadToEnd();
    }

    private static string ReadLayoutText(Stream stream)
    {
        using var ms = new MemoryStream();
        stream.CopyTo(ms);
        var bytes = ms.ToArray();

        if (TryReadJson(bytes, Encoding.UTF8, out var utf8))
        {
            return utf8;
        }

        if (TryReadJson(bytes, Encoding.Unicode, out var utf16Le))
        {
            return utf16Le;
        }

        if (TryReadJson(bytes, Encoding.BigEndianUnicode, out var utf16Be))
        {
            return utf16Be;
        }

        return Encoding.UTF8.GetString(bytes);
    }

    private static bool TryReadJson(byte[] bytes, Encoding encoding, out string text)
    {
        try
        {
            text = encoding.GetString(bytes);
            using var _ = JsonDocument.Parse(text);
            return true;
        }
        catch
        {
            text = string.Empty;
            return false;
        }
    }

    private static (string VisualType, string Title) ParseVisualInfo(string configJson)
    {
        if (string.IsNullOrWhiteSpace(configJson))
        {
            return ("unknown", string.Empty);
        }

        try
        {
            using var doc = JsonDocument.Parse(configJson);
            var root = doc.RootElement;
            var singleVisual = root.TryGetProperty("singleVisual", out var sv) ? sv : default;
            var visualType = singleVisual.ValueKind == JsonValueKind.Object
                ? GetString(singleVisual, "visualType")
                : string.Empty;

            var title = ExtractTitleFromConfig(singleVisual);
            return (string.IsNullOrWhiteSpace(visualType) ? "unknown" : visualType, title);
        }
        catch
        {
            return ("unknown", string.Empty);
        }
    }

    private static string ExtractTitleFromConfig(JsonElement singleVisual)
    {
        if (singleVisual.ValueKind != JsonValueKind.Object)
        {
            return string.Empty;
        }

        if (!singleVisual.TryGetProperty("objects", out var objectsNode) || objectsNode.ValueKind != JsonValueKind.Object)
        {
            return string.Empty;
        }

        if (!objectsNode.TryGetProperty("title", out var titleNode) || titleNode.ValueKind != JsonValueKind.Array)
        {
            return string.Empty;
        }

        foreach (var titleObj in titleNode.EnumerateArray())
        {
            if (!titleObj.TryGetProperty("properties", out var props) || props.ValueKind != JsonValueKind.Object)
            {
                continue;
            }

            if (!props.TryGetProperty("text", out var textNode) || textNode.ValueKind != JsonValueKind.Object)
            {
                continue;
            }

            if (!textNode.TryGetProperty("expr", out var exprNode) || exprNode.ValueKind != JsonValueKind.Object)
            {
                continue;
            }

            var raw = TryExtractLiteralText(textNode);
            if (!string.IsNullOrWhiteSpace(raw))
            {
                return raw;
            }
        }

        return string.Empty;
    }

    private static string GetString(JsonElement element, string propertyName)
    {
        if (element.ValueKind != JsonValueKind.Object || !element.TryGetProperty(propertyName, out var node))
        {
            return string.Empty;
        }

        return node.ValueKind switch
        {
            JsonValueKind.String => node.GetString() ?? string.Empty,
            JsonValueKind.Number => node.GetRawText(),
            JsonValueKind.True => "true",
            JsonValueKind.False => "false",
            _ => string.Empty
        };
    }
}
