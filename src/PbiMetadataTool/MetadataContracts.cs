namespace PbiMetadataTool;

public sealed record ModelMetadata(
    string DatabaseName,
    string DatabaseId,
    int CompatibilityLevel,
    IReadOnlyList<TableMetadata> Tables,
    IReadOnlyList<RelationshipMetadata> Relationships,
    IReadOnlyList<RoleMetadata> Roles);

public sealed record TableMetadata(
    string Name,
    bool IsHidden,
    IReadOnlyList<ColumnMetadata> Columns,
    IReadOnlyList<MeasureMetadata> Measures,
    string TableType = "",
    string Expression = "");

public sealed record ColumnMetadata(
    string Name,
    string ColumnType,
    string DataType,
    bool IsHidden,
    string Expression = "");

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

public sealed record RoleMetadata(
    string Name,
    string Description,
    string ModelPermission,
    IReadOnlyList<RoleMemberMetadata> Members,
    IReadOnlyList<RoleTablePermissionMetadata> TablePermissions);

public sealed record RoleMemberMetadata(
    string Name,
    string IdentityProvider,
    string MemberType);

public sealed record RoleTablePermissionMetadata(
    string TableName,
    string FilterExpression,
    string MetadataPermission);

public sealed record PowerBiInstanceInfo(
    int DesktopPid,
    int MsmdsrvPid,
    int Port,
    string WorkspacePath,
    string? PbixPathHint,
    DateTime LastSeenUtc);
