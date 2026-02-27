namespace PbiMetadataTool;

public sealed record ModelMetadata(
    string DatabaseName,
    string DatabaseId,
    int CompatibilityLevel,
    IReadOnlyList<TableMetadata> Tables,
    IReadOnlyList<RelationshipMetadata> Relationships);

public sealed record TableMetadata(
    string Name,
    bool IsHidden,
    IReadOnlyList<ColumnMetadata> Columns,
    IReadOnlyList<MeasureMetadata> Measures);

public sealed record ColumnMetadata(
    string Name,
    string ColumnType,
    string DataType,
    bool IsHidden);

public sealed record MeasureMetadata(
    string Name,
    string Expression,
    string FormatString,
    bool IsHidden,
    string DisplayFolder = "");

public sealed record RelationshipMetadata(
    string Name,
    string FromTable,
    string FromColumn,
    string ToTable,
    string ToColumn,
    string CrossFilterDirection,
    bool IsActive);

public sealed record PowerBiInstanceInfo(
    int DesktopPid,
    int MsmdsrvPid,
    int Port,
    string WorkspacePath,
    string? PbixPathHint,
    DateTime LastSeenUtc);
