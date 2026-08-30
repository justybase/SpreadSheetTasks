namespace SpreadSheetTasks;

/// <summary>
/// Selects the style fallback used when a replacement cell has no suitable
/// style in the original worksheet.
/// </summary>
public enum ReplaceSheetDataStyleFallback
{
    /// <summary>Reuse the dominant style found in the original column.</summary>
    Inherit,

    /// <summary>Use the workbook's General style (style index 0).</summary>
    General
}

/// <summary>
/// Options used by <see cref="XlsxUpdater.ReplaceSheetData"/> and
/// <see cref="XlsbUpdater.ReplaceSheetData"/>.
/// </summary>
public sealed class ReplaceSheetDataOptions
{
    /// <summary>
    /// Optional header values. When supplied they are written as row 1 and the
    /// supplied input rows start at row 2. When omitted, input rows start at row 1.
    /// </summary>
    public IReadOnlyList<string>? Headers { get; init; }

    /// <summary>Style strategy for replacement cells. Defaults to inheritance.</summary>
    public ReplaceSheetDataStyleFallback StyleFallback { get; init; } = ReplaceSheetDataStyleFallback.Inherit;
}
