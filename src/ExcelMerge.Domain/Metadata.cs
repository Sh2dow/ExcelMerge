namespace ExcelMerge.Domain;

public enum WorkbookFormat
{
    Unknown,
    Xlsx,
    Csv,
    Tsv,
}

public enum SheetVisibility
{
    Visible,
    Hidden,
    VeryHidden,
}

public sealed record WorkbookMetadata(
    string DisplayName,
    WorkbookFormat Format,
    IReadOnlyList<SheetMetadata> Sheets,
    string? SourcePath = null,
    long? ContentLength = null,
    DateTimeOffset? LastModifiedUtc = null,
    bool Uses1904DateSystem = false);

/// <param name="Id">An adapter-defined stable identifier; it is not required to be the sheet name.</param>
/// <param name="Position">The zero-based position in workbook order.</param>
/// <param name="FirstRowIndex">The zero-based first used row, when known.</param>
/// <param name="LastRowIndex">The zero-based last used row, when known.</param>
/// <param name="FirstColumnIndex">The zero-based first used column, when known.</param>
/// <param name="LastColumnIndex">The zero-based last used column, when known.</param>
public sealed record SheetMetadata(
    string Id,
    string Name,
    int Position,
    SheetVisibility Visibility = SheetVisibility.Visible,
    int? FirstRowIndex = null,
    int? LastRowIndex = null,
    int? FirstColumnIndex = null,
    int? LastColumnIndex = null,
    long? NonEmptyCellCount = null);
