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

        internal Part(string name, string filePath)
        {
            Name = name;
            FilePath = filePath;
            Data = [];
        }

        internal string Name { get; }
        internal byte[] Data { get; private set; }
        internal string? FilePath { get; private set; }
        internal long Length => FilePath is null ? Data.LongLength : new FileInfo(FilePath).Length;

        internal byte[] ReadAll()
        {
            return FilePath is null ? Data.ToArray() : File.ReadAllBytes(FilePath);
        }

        internal void CopyTo(Stream destination)
        {
            if (FilePath is null)
            {
                destination.Write(Data, 0, Data.Length);
                return;
            }

            using var input = new FileStream(FilePath, FileMode.Open, FileAccess.Read, FileShare.Read,
                bufferSize: 1024 * 1024, FileOptions.SequentialScan);
            input.CopyTo(destination, 1024 * 1024);
        }

        internal void SetBytes(byte[] data)
        {
            DeleteStagedFile();
            Data = data.ToArray();
        }

        internal void SetFile(string filePath)
        {
            if (!string.Equals(FilePath, filePath, StringComparison.Ordinal))
                DeleteStagedFile();
            FilePath = filePath;
            Data = [];
        }

        internal void Dispose()
        {
            DeleteStagedFile();
        }

        private void DeleteStagedFile()
        {
            if (FilePath is null)
                return;

            try
            {
                File.Delete(FilePath);
            }
            catch (FileNotFoundException)
            {
                // The caller may already have removed a failed staging file.
            }
            FilePath = null;
        }
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
        return _lookup.TryGetValue(name, out var part) ? part.ReadAll() : null;
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
            existing.SetBytes(data);
            return;
        }

        AddPart(name, data);
    }

    /// <summary>
    /// Replaces a part with a staged file. Ownership of the file transfers to
    /// this package; it remains available for repeated saves and is removed on
    /// replacement or disposal.
    /// </summary>
    internal void SetPartFromFile(string name, string filePath)
    {
        ThrowIfDisposed();
        ArgumentNullException.ThrowIfNull(name);
        ArgumentException.ThrowIfNullOrWhiteSpace(filePath);
        if (!File.Exists(filePath))
            throw new FileNotFoundException("Staged ZIP part was not found.", filePath);

        if (_lookup.TryGetValue(name, out var existing))
        {
            existing.SetFile(filePath);
            return;
        }

        var part = new Part(name, filePath);
        _parts.Add(part);
        _lookup[name] = part;
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
                part.CopyTo(destination);
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
            long uncompressedLength = 0;
            foreach (Part part in _parts)
            {
                uncompressedLength = Math.Min(1024L * 1024, uncompressedLength + part.Length);
                if (uncompressedLength >= 1024L * 1024)
                    break;
            }
            int bufferSize = uncompressedLength >= 1024L * 1024 ? 1024 * 1024 : 64 * 1024;
            using (var output = new FileStream(temporary, FileMode.CreateNew, FileAccess.ReadWrite, FileShare.None,
                bufferSize, FileOptions.SequentialScan))
            using (var archive = new ZipArchive(output, ZipArchiveMode.Create, leaveOpen: false))
            {
                foreach (var part in _parts)
                {
                    var entry = archive.CreateEntry(part.Name, CompressionLevel.Optimal);
                    using var destination = entry.Open();
                    part.CopyTo(destination);
                }
            }
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
        foreach (var part in _parts)
            part.Dispose();
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
