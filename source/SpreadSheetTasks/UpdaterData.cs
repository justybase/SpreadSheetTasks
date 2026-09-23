using System.Collections;
using System.Data;

namespace SpreadSheetTasks;

internal sealed class UpdaterRows
{
    private UpdaterRows(object?[][] rows, IReadOnlyList<string>? headers)
    {
        Rows = rows;
        Headers = headers;
    }

    internal object?[][] Rows { get; }

    internal IReadOnlyList<string>? Headers { get; }

    internal int DataRowCount => Rows.Length;

    internal int TotalRowCount => Rows.Length + (Headers is null ? 0 : 1);

    internal int LastColumnIndex
    {
        get
        {
            int max = Headers?.Count ?? 0;
            foreach (var row in Rows)
            {
                if (row is not null && row.Length > max)
                    max = row.Length;
            }
            return max - 1;
        }
    }

    internal static UpdaterRows FromReader(IDataReader reader, ReplaceSheetDataOptions? options)
    {
        ArgumentNullException.ThrowIfNull(reader);
        var rows = new List<object?[]>();
        int fieldCount = reader.FieldCount;

        while (reader.Read())
        {
            var row = new object?[fieldCount];
            for (int column = 0; column < fieldCount; column++)
            {
                object value = reader.GetValue(column);
                row[column] = value == DBNull.Value ? null : value;
            }
            rows.Add(row);
        }

        return Create(rows, options?.Headers);
    }

    internal static UpdaterRows FromRows(List<object?[]> rows, ReplaceSheetDataOptions? options)
    {
        ArgumentNullException.ThrowIfNull(rows);
        return Create(rows, options?.Headers);
    }

    internal static List<object?[]> ReadReaderPrefix(IDataReader reader, int maximumRows, out bool hasMore)
    {
        ArgumentNullException.ThrowIfNull(reader);
        ArgumentOutOfRangeException.ThrowIfNegative(maximumRows);

        int fieldCount = reader.FieldCount;
        var rows = new List<object?[]>();
        while (rows.Count <= maximumRows && reader.Read())
        {
            var row = new object?[fieldCount];
            for (int column = 0; column < fieldCount; column++)
            {
                object value = reader.GetValue(column);
                row[column] = value == DBNull.Value ? null : value;
            }
            rows.Add(row);
        }

        hasMore = rows.Count > maximumRows;
        return rows;
    }

    internal static UpdaterRows FromDataTable(DataTable table, ReplaceSheetDataOptions? options)
    {
        ArgumentNullException.ThrowIfNull(table);
        using var reader = table.CreateDataReader();
        return FromReader(reader, options);
    }

    internal static UpdaterRows FromArray(object?[][] rows, ReplaceSheetDataOptions? options)
    {
        ArgumentNullException.ThrowIfNull(rows);
        return Create(rows, options?.Headers);
    }

    internal static UpdaterRows FromList(IReadOnlyList<object?[]> rows, ReplaceSheetDataOptions? options)
    {
        ArgumentNullException.ThrowIfNull(rows);
        var copy = new object?[rows.Count][];
        for (int i = 0; i < rows.Count; i++)
            copy[i] = rows[i] ?? Array.Empty<object?>();
        return Create(copy, options?.Headers);
    }

    private static UpdaterRows Create(IEnumerable<object?[]> rows, IReadOnlyList<string>? headers)
    {
        var materialized = rows
            .Select(row => row ?? Array.Empty<object?>())
            .ToList();

        int end = materialized.Count;
        while (end > 0 && IsEmptyRow(materialized[end - 1]))
            end--;

        if (end != materialized.Count)
            materialized.RemoveRange(end, materialized.Count - end);

        if (headers is not null)
        {
            for (int i = 0; i < headers.Count; i++)
                ArgumentNullException.ThrowIfNull(headers[i]);
        }

        return new UpdaterRows(materialized.ToArray(), headers);
    }

    internal static object? Unwrap(object? value)
    {
        if (value is FormattedCell formatted)
            return formatted.Value;
        return value;
    }

    internal static bool IsEmpty(object? value)
    {
        value = Unwrap(value);
        return value is null || value == DBNull.Value;
    }

    internal static bool IsEmptyRow(IReadOnlyList<object?> row)
    {
        for (int i = 0; i < row.Count; i++)
        {
            if (!IsEmpty(row[i]))
                return false;
        }
        return true;
    }

    internal static void ValidateHeaders(IReadOnlyList<string>? headers)
    {
        if (headers is null)
            return;

        for (int i = 0; i < headers.Count; i++)
            ArgumentNullException.ThrowIfNull(headers[i]);
    }
}
