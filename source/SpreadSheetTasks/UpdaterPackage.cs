using System.IO.Compression;

namespace SpreadSheetTasks;

/// <summary>
/// In-memory representation of an Office ZIP package. Keeping parts separate
/// lets an updater replace only the parts it actually changes.
/// </summary>
internal sealed class UpdaterPackage : IDisposable
{
    private sealed class Part
    {
        internal Part(string name, byte[] data)
        {
            Name = name;
            Data = data;
        }

        internal string Name { get; }
        internal byte[] Data { get; set; }
    }

    private readonly string _sourcePath;
    private readonly List<Part> _parts = [];
    private readonly Dictionary<string, Part> _lookup = new(StringComparer.OrdinalIgnoreCase);
    private bool _disposed;

    private UpdaterPackage(string sourcePath)
    {
        _sourcePath = sourcePath;
    }

    internal static UpdaterPackage Load(string path)
    {
        ArgumentNullException.ThrowIfNull(path);
        ArgumentException.ThrowIfNullOrWhiteSpace(path);
        if (!File.Exists(path))
            throw new FileNotFoundException("Excel package was not found.", path);

        var package = new UpdaterPackage(Path.GetFullPath(path));
        try
        {
            using var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read);
            using var archive = new ZipArchive(stream, ZipArchiveMode.Read, leaveOpen: false);
            foreach (var entry in archive.Entries)
            {
                using var input = entry.Open();
                using var output = new MemoryStream(entry.Length > 0 && entry.Length <= int.MaxValue
                    ? (int)entry.Length
                    : 0);
                input.CopyTo(output);
                package.AddPart(entry.FullName, output.ToArray());
            }
            return package;
        }
        catch
        {
            package.Dispose();
            throw;
        }
    }

    internal static void EnsureExtension(string path, string expectedExtension, string apiName)
    {
        string actualExtension = Path.GetExtension(path);
        if (!string.Equals(actualExtension, expectedExtension, StringComparison.OrdinalIgnoreCase))
        {
            throw new NotSupportedException(
                $"{apiName} supports only '{expectedExtension}' files; received '{actualExtension}'.");
        }
    }

    internal IReadOnlyList<string> PartNames
    {
        get
        {
            ThrowIfDisposed();
            return _parts.Select(part => part.Name).ToArray();
        }
    }

    internal bool Contains(string name)
    {
        ThrowIfDisposed();
        return _lookup.ContainsKey(name);
    }

    internal byte[]? TryGetPart(string name)
    {
        ThrowIfDisposed();
        return _lookup.TryGetValue(name, out var part) ? part.Data.ToArray() : null;
    }

    internal byte[] GetPart(string name)
    {
        return TryGetPart(name) ?? throw new InvalidDataException($"Required ZIP part '{name}' is missing.");
    }

    internal void SetPart(string name, byte[] data)
    {
        ThrowIfDisposed();
        ArgumentNullException.ThrowIfNull(name);
        ArgumentNullException.ThrowIfNull(data);

        if (_lookup.TryGetValue(name, out var existing))
        {
            existing.Data = data.ToArray();
            return;
        }

        AddPart(name, data);
    }

    internal byte[] ToArray()
    {
        ThrowIfDisposed();
        using var output = new MemoryStream();
        using (var archive = new ZipArchive(output, ZipArchiveMode.Create, leaveOpen: true))
        {
            foreach (var part in _parts)
            {
                var entry = archive.CreateEntry(part.Name, CompressionLevel.Optimal);
                using var destination = entry.Open();
                destination.Write(part.Data, 0, part.Data.Length);
            }
        }
        return output.ToArray();
    }

    internal void Save(string? outputPath = null)
    {
        ThrowIfDisposed();
        string target = Path.GetFullPath(outputPath ?? _sourcePath);
        string? directory = Path.GetDirectoryName(target);
        if (string.IsNullOrEmpty(directory))
            throw new InvalidOperationException("The output path has no directory.");

        Directory.CreateDirectory(directory);
        string temporary = Path.Combine(directory, $".{Path.GetFileName(target)}.{Guid.NewGuid():N}.tmp");
        try
        {
            File.WriteAllBytes(temporary, ToArray());
            File.Move(temporary, target, overwrite: true);
        }
        finally
        {
            if (File.Exists(temporary))
                File.Delete(temporary);
        }
    }

    public void Dispose()
    {
        if (_disposed)
            return;
        _disposed = true;
        _parts.Clear();
        _lookup.Clear();
    }

    private void AddPart(string name, byte[] data)
    {
        var part = new Part(name, data.ToArray());
        _parts.Add(part);
        _lookup[name] = part;
    }

    private void ThrowIfDisposed()
    {
        ObjectDisposedException.ThrowIf(_disposed, this);
    }
}
