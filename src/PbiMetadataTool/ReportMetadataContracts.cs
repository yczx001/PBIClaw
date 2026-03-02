namespace PbiMetadataTool;

public sealed record ReportMetadata(
    string SourceType,
    string SourcePath,
    IReadOnlyList<ReportPageMetadata> Pages);

public sealed record ReportPageMetadata(
    string Name,
    string DisplayName,
    IReadOnlyList<ReportVisualMetadata> Visuals);

public sealed record ReportVisualMetadata(
    string Name,
    string VisualType,
    string Title);
