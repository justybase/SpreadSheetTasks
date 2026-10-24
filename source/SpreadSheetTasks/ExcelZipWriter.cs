using System;
using System.Buffers.Binary;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Text;

namespace SpreadSheetTasks
{
    /// <summary>
    /// Maps the public <see cref="CompressionLevel"/> values (and the library default)
    /// to numeric zlib levels. <c>null</c> means the balanced default (level 3).
    /// </summary>
    internal static class ExcelZipCompression
    {
        internal const int DefaultLevel = 3;

        internal static int ToLevel(CompressionLevel? level) => level switch
        {
            null => DefaultLevel,
            CompressionLevel.NoCompression => 0,
            CompressionLevel.Fastest => 1,
            CompressionLevel.Optimal => 6,
            CompressionLevel.SmallestSize => 9,
            _ => DefaultLevel,
        };

        /// <summary>Fallback mapping used where only the enum is available (net8.0, non-seekable streams).</summary>
        internal static CompressionLevel ToEnum(int level) => level switch
        {
            <= 0 => CompressionLevel.NoCompression,
            <= 2 => CompressionLevel.Fastest,
            <= 6 => CompressionLevel.Optimal,
            _ => CompressionLevel.SmallestSize,
        };
    }

    /// <summary>
    /// A single entry opened for writing. Mirrors the <c>ZipArchiveEntry.Open()</c> shape used by
    /// the writers so the ZIP implementation can be swapped without touching every call site.
    /// </summary>
    internal sealed class ExcelZipEntry
    {
        private readonly Stream _stream;

        internal ExcelZipEntry(Stream stream) => _stream = stream;

        public Stream Open() => _stream;
    }

    /// <summary>
    /// ZIP writer used by the workbook writers.
    /// <para>
    /// Seekable outputs use the built-in writer so an exact zlib level (3 by default) can be
    /// selected; non-seekable outputs fall back to <see cref="ZipArchive"/>, which cannot expose
    /// numeric levels but supports streaming without seeking.
    /// </para>
    /// </summary>
    internal sealed class ExcelZipWriter : IDisposable
    {
        private readonly IBackend _backend;

        internal ExcelZipWriter(Stream output, CompressionLevel? compressionLevel, bool leaveOpen)
        {
            // The numeric-level writer is used for the new balanced default (zlib 3). When the
            // caller explicitly selects an enum level, keep the BCL writer: it is faster than
            // the custom path at those levels and preserves the previous behaviour exactly.
            _backend = compressionLevel is null && output.CanSeek
                ? new SeekableBackend(output, ExcelZipCompression.DefaultLevel, leaveOpen)
                : new ZipArchiveBackend(output, ExcelZipCompression.ToLevel(compressionLevel), leaveOpen);
        }

        public ExcelZipEntry CreateEntry(string name) => _backend.CreateEntry(name);

        public void CreateEntryFromFile(string sourceFileName, string entryName)
            => _backend.CreateEntryFromFile(sourceFileName, entryName);

        public void Dispose() => _backend.Dispose();

        private interface IBackend : IDisposable
        {
            ExcelZipEntry CreateEntry(string name);
            void CreateEntryFromFile(string sourceFileName, string entryName);
        }

        /// <summary>Non-seekable fallback: the BCL writer, restricted to the enum levels.</summary>
        private sealed class ZipArchiveBackend : IBackend
        {
            private readonly ZipArchive _archive;
            private readonly CompressionLevel _level;

            internal ZipArchiveBackend(Stream output, int compressionLevel, bool leaveOpen)
            {
                _level = ExcelZipCompression.ToEnum(compressionLevel);
                _archive = new ZipArchive(output, ZipArchiveMode.Create, leaveOpen);
            }

            public ExcelZipEntry CreateEntry(string name)
                => new(_archive.CreateEntry(name, _level).Open());

            public void CreateEntryFromFile(string sourceFileName, string entryName)
                => _archive.CreateEntryFromFile(sourceFileName, entryName, _level);

            public void Dispose() => _archive.Dispose();
        }

        /// <summary>
        /// ZIP writer for seekable streams. It always emits ZIP64 records, so entries and the
        /// archive may exceed 4 GiB; sizes in the local header are patched after the data is
        /// written, exactly like the BCL writer does for seekable output.
        /// </summary>
        private sealed class SeekableBackend : IBackend
        {
            private const uint LocalFileHeaderSignature = 0x04034b50;
            private const uint CentralDirectoryHeaderSignature = 0x02014b50;
            private const uint EndOfCentralDirectorySignature = 0x06054b50;
            private const uint Zip64EndOfCentralDirectorySignature = 0x06064b50;
            private const uint Zip64EndOfCentralDirectoryLocatorSignature = 0x07064b50;
            private const ushort Zip64Version = 45;
            private const ushort Utf8Flag = 0x0800;
            private const ushort StoredMethod = 0;
            private const ushort DeflateMethod = 8;

            private readonly Stream _output;
            private readonly bool _leaveOpen;
            private readonly int _level;
            private readonly long _origin;
            private readonly List<EntryInfo> _entries = new();
            private EntryStream? _current;
            private bool _disposed;

            internal SeekableBackend(Stream output, int compressionLevel, bool leaveOpen)
            {
                _output = output;
                _leaveOpen = leaveOpen;
                _level = Math.Clamp(compressionLevel, 0, 9);
                _origin = output.Position;
            }

            public ExcelZipEntry CreateEntry(string name)
            {
                ThrowIfDisposed();
                FinishCurrent();

                byte[] nameBytes = Encoding.UTF8.GetBytes(name);
                var info = new EntryInfo(nameBytes)
                {
                    Offset = _output.Position - _origin,
                    Method = _level == 0 ? StoredMethod : DeflateMethod,
                    Time = DosTime(),
                    Date = DosDate(),
                };
                info.NameLength = nameBytes.Length;
                info.ExtraOffset = info.Offset + 30 + nameBytes.Length + 4;

                WriteLocalHeader(info);
                Stream sink = _level == 0 ? _output : CreateDeflate();
                var entry = new EntryStream(this, info, sink, ownsSink: _level != 0);
                _current = entry;
                return new ExcelZipEntry(entry);
            }

            public void CreateEntryFromFile(string sourceFileName, string entryName)
            {
                using var input = new FileStream(sourceFileName, FileMode.Open, FileAccess.Read, FileShare.Read,
                    bufferSize: 64 * 1024, FileOptions.SequentialScan);
                using var entry = CreateEntry(entryName).Open();
                input.CopyTo(entry, 64 * 1024);
            }

            private Stream CreateDeflate()
            {
#if NET9_0_OR_GREATER
                return new DeflateStream(_output,
                    new ZLibCompressionOptions { CompressionLevel = _level },
                    leaveOpen: true);
#else
                return new DeflateStream(_output, ExcelZipCompression.ToEnum(_level), leaveOpen: true);
#endif
            }

            private void WriteLocalHeader(EntryInfo info)
            {
                Span<byte> header = stackalloc byte[30];
                BinaryPrimitives.WriteUInt32LittleEndian(header, LocalFileHeaderSignature);
                BinaryPrimitives.WriteUInt16LittleEndian(header[4..], Zip64Version);
                BinaryPrimitives.WriteUInt16LittleEndian(header[6..], Utf8Flag);
                BinaryPrimitives.WriteUInt16LittleEndian(header[8..], info.Method);
                BinaryPrimitives.WriteUInt16LittleEndian(header[10..], info.Time);
                BinaryPrimitives.WriteUInt16LittleEndian(header[12..], info.Date);
                BinaryPrimitives.WriteUInt32LittleEndian(header[14..], 0); // CRC, patched
                BinaryPrimitives.WriteUInt32LittleEndian(header[18..], uint.MaxValue); // compressed, ZIP64 extra
                BinaryPrimitives.WriteUInt32LittleEndian(header[22..], uint.MaxValue); // uncompressed, ZIP64 extra
                BinaryPrimitives.WriteUInt16LittleEndian(header[26..], (ushort)info.NameLength);
                BinaryPrimitives.WriteUInt16LittleEndian(header[28..], 20); // ZIP64 extra: id + size + 16 bytes
                _output.Write(header);
                _output.Write(info.NameBytes);

                Span<byte> extra = stackalloc byte[20];
                BinaryPrimitives.WriteUInt16LittleEndian(extra, 0x0001);
                BinaryPrimitives.WriteUInt16LittleEndian(extra[2..], 16);
                BinaryPrimitives.WriteUInt64LittleEndian(extra[4..], 0);
                BinaryPrimitives.WriteUInt64LittleEndian(extra[12..], 0);
                _output.Write(extra);
            }

            /// <summary>Called by <see cref="EntryStream"/> once its data was flushed.</summary>
            private void FinishEntry(EntryStream entry)
            {
                if (entry.IsFinished)
                    return;

                entry.MarkFinished();
                entry.DisposeSink();

                long end = _output.Position;
                EntryInfo info = entry.Info;
                info.CompressedSize = end - entry.DataStart;
                info.UncompressedSize = entry.UncompressedLength;
                info.Crc = entry.Crc;

                PatchLocalHeader(info);
                _entries.Add(info);
                if (ReferenceEquals(_current, entry))
                    _current = null;
            }

            private void PatchLocalHeader(EntryInfo info)
            {
                long position = _output.Position;
                _output.Position = _origin + info.Offset + 14;
                Span<byte> crc = stackalloc byte[4];
                BinaryPrimitives.WriteUInt32LittleEndian(crc, info.Crc);
                _output.Write(crc);

                _output.Position = _origin + info.ExtraOffset;
                Span<byte> sizes = stackalloc byte[16];
                BinaryPrimitives.WriteUInt64LittleEndian(sizes, (ulong)info.UncompressedSize);
                BinaryPrimitives.WriteUInt64LittleEndian(sizes[8..], (ulong)info.CompressedSize);
                _output.Write(sizes);

                _output.Position = position;
            }

            private void FinishCurrent()
            {
                _current?.FinishFromWriter();
            }

            private void WriteCentralDirectory()
            {
                long centralStart = _output.Position - _origin;
                foreach (EntryInfo info in _entries)
                {
                    WriteCentralDirectoryHeader(info);
                }
                long centralSize = _output.Position - _origin - centralStart;

                bool needsZip64End = _entries.Count > ushort.MaxValue
                    || centralStart > uint.MaxValue
                    || centralSize > uint.MaxValue;

                if (needsZip64End)
                {
                    WriteZip64EndOfCentralDirectory(centralStart, centralSize);
                }

                WriteEndOfCentralDirectory(centralStart, centralSize, needsZip64End);
                _output.Flush();
            }

            private void WriteCentralDirectoryHeader(EntryInfo info)
            {
                Span<byte> header = stackalloc byte[46];
                BinaryPrimitives.WriteUInt32LittleEndian(header, CentralDirectoryHeaderSignature);
                BinaryPrimitives.WriteUInt16LittleEndian(header[4..], 0x0014); // version made by (MS-DOS, 2.0)
                BinaryPrimitives.WriteUInt16LittleEndian(header[6..], Zip64Version);
                BinaryPrimitives.WriteUInt16LittleEndian(header[8..], Utf8Flag);
                BinaryPrimitives.WriteUInt16LittleEndian(header[10..], info.Method);
                BinaryPrimitives.WriteUInt16LittleEndian(header[12..], info.Time);
                BinaryPrimitives.WriteUInt16LittleEndian(header[14..], info.Date);
                BinaryPrimitives.WriteUInt32LittleEndian(header[16..], info.Crc);
                BinaryPrimitives.WriteUInt32LittleEndian(header[20..], uint.MaxValue);
                BinaryPrimitives.WriteUInt32LittleEndian(header[24..], uint.MaxValue);
                BinaryPrimitives.WriteUInt16LittleEndian(header[28..], (ushort)info.NameLength);
                BinaryPrimitives.WriteUInt16LittleEndian(header[30..], 28); // ZIP64 extra
                BinaryPrimitives.WriteUInt16LittleEndian(header[32..], 0);  // comment
                BinaryPrimitives.WriteUInt16LittleEndian(header[34..], 0);  // disk
                BinaryPrimitives.WriteUInt16LittleEndian(header[36..], 0);  // internal attrs
                BinaryPrimitives.WriteUInt32LittleEndian(header[38..], 0);  // external attrs
                BinaryPrimitives.WriteUInt32LittleEndian(header[42..], uint.MaxValue);
                _output.Write(header);
                _output.Write(info.NameBytes);

                Span<byte> extra = stackalloc byte[28];
                BinaryPrimitives.WriteUInt16LittleEndian(extra, 0x0001);
                BinaryPrimitives.WriteUInt16LittleEndian(extra[2..], 24);
                BinaryPrimitives.WriteUInt64LittleEndian(extra[4..], (ulong)info.UncompressedSize);
                BinaryPrimitives.WriteUInt64LittleEndian(extra[12..], (ulong)info.CompressedSize);
                BinaryPrimitives.WriteUInt64LittleEndian(extra[20..], (ulong)info.Offset);
                _output.Write(extra);
            }

            private void WriteZip64EndOfCentralDirectory(long centralStart, long centralSize)
            {
                long zip64RecordOffset = _output.Position - _origin;

                Span<byte> record = stackalloc byte[56];
                BinaryPrimitives.WriteUInt32LittleEndian(record, Zip64EndOfCentralDirectorySignature);
                BinaryPrimitives.WriteUInt64LittleEndian(record[4..], 44);
                BinaryPrimitives.WriteUInt16LittleEndian(record[12..], 0x0014);
                BinaryPrimitives.WriteUInt16LittleEndian(record[14..], Zip64Version);
                BinaryPrimitives.WriteUInt32LittleEndian(record[16..], 0);
                BinaryPrimitives.WriteUInt32LittleEndian(record[20..], 0);
                BinaryPrimitives.WriteUInt64LittleEndian(record[24..], (ulong)_entries.Count);
                BinaryPrimitives.WriteUInt64LittleEndian(record[32..], (ulong)_entries.Count);
                BinaryPrimitives.WriteUInt64LittleEndian(record[40..], (ulong)centralSize);
                BinaryPrimitives.WriteUInt64LittleEndian(record[48..], (ulong)centralStart);
                _output.Write(record);

                Span<byte> locator = stackalloc byte[20];
                BinaryPrimitives.WriteUInt32LittleEndian(locator, Zip64EndOfCentralDirectoryLocatorSignature);
                BinaryPrimitives.WriteUInt32LittleEndian(locator[4..], 0);
                BinaryPrimitives.WriteUInt64LittleEndian(locator[8..], (ulong)zip64RecordOffset);
                BinaryPrimitives.WriteUInt32LittleEndian(locator[16..], 1);
                _output.Write(locator);
            }

            private void WriteEndOfCentralDirectory(long centralStart, long centralSize, bool zip64)
            {
                Span<byte> end = stackalloc byte[22];
                BinaryPrimitives.WriteUInt32LittleEndian(end, EndOfCentralDirectorySignature);
                BinaryPrimitives.WriteUInt16LittleEndian(end[4..], 0);
                BinaryPrimitives.WriteUInt16LittleEndian(end[6..], 0);
                BinaryPrimitives.WriteUInt16LittleEndian(end[8..],
                    zip64 ? ushort.MaxValue : (ushort)_entries.Count);
                BinaryPrimitives.WriteUInt16LittleEndian(end[10..],
                    zip64 ? ushort.MaxValue : (ushort)_entries.Count);
                BinaryPrimitives.WriteUInt32LittleEndian(end[12..],
                    zip64 ? uint.MaxValue : (uint)centralSize);
                BinaryPrimitives.WriteUInt32LittleEndian(end[16..],
                    zip64 ? uint.MaxValue : (uint)centralStart);
                BinaryPrimitives.WriteUInt16LittleEndian(end[20..], 0);
                _output.Write(end);
            }

            private static ushort DosTime()
            {
                DateTime now = DateTime.Now;
                return (ushort)((now.Hour << 11) | (now.Minute << 5) | (now.Second / 2));
            }

            private static ushort DosDate()
            {
                DateTime now = DateTime.Now;
                int year = Math.Clamp(now.Year, 1980, 2107);
                return (ushort)(((year - 1980) << 9) | (now.Month << 5) | now.Day);
            }

            private void ThrowIfDisposed()
            {
                if (_disposed)
                    throw new ObjectDisposedException(nameof(ExcelZipWriter));
            }

            public void Dispose()
            {
                if (_disposed)
                    return;

                try
                {
                    FinishCurrent();
                    WriteCentralDirectory();
                }
                finally
                {
                    _disposed = true;
                    if (!_leaveOpen)
                        _output.Dispose();
                }
            }

            private sealed class EntryInfo
            {
                internal EntryInfo(byte[] nameBytes)
                {
                    NameBytes = nameBytes;
                }

                internal byte[] NameBytes { get; }
                internal int NameLength;
                internal ushort Method;
                internal ushort Time;
                internal ushort Date;
                internal uint Crc;
                internal long CompressedSize;
                internal long UncompressedSize;
                internal long Offset;
                internal long ExtraOffset;
            }

            /// <summary>Entry data stream: tracks CRC and uncompressed length, then finishes the entry.</summary>
            private sealed class EntryStream : Stream
            {
                private readonly SeekableBackend _owner;
                private readonly Stream _sink;
                private readonly bool _ownsSink;
                private uint _crc;
                private long _uncompressed;
                private bool _finished;

                internal EntryStream(SeekableBackend owner, EntryInfo info, Stream sink, bool ownsSink)
                {
                    _owner = owner;
                    Info = info;
                    _sink = sink;
                    _ownsSink = ownsSink;
                    DataStart = owner._output.Position;
                }

                internal EntryInfo Info { get; }
                internal long DataStart { get; }
                internal uint Crc => _crc;
                internal long UncompressedLength => _uncompressed;
                internal bool IsFinished => _finished;

                internal void MarkFinished() => _finished = true;

                internal void DisposeSink()
                {
                    if (_ownsSink)
                        _sink.Dispose();
                    else
                        _sink.Flush();
                }

                internal void FinishFromWriter()
                {
                    if (!_finished)
                        _owner.FinishEntry(this);
                }

                public override bool CanRead => false;
                public override bool CanSeek => false;
                public override bool CanWrite => true;
                public override long Length => throw new NotSupportedException();

                public override long Position
                {
                    get => throw new NotSupportedException();
                    set => throw new NotSupportedException();
                }

                public override void Flush() => _sink.Flush();

                public override int Read(byte[] buffer, int offset, int count)
                    => throw new NotSupportedException();

                public override long Seek(long offset, SeekOrigin origin)
                    => throw new NotSupportedException();

                public override void SetLength(long value)
                    => throw new NotSupportedException();

                public override void Write(byte[] buffer, int offset, int count)
                    => Write(buffer.AsSpan(offset, count));

                public override void Write(ReadOnlySpan<byte> buffer)
                {
                    if (_finished)
                        throw new ObjectDisposedException(nameof(ExcelZipEntry));

                    if (buffer.Length == 0)
                        return;

                    _crc = Crc32.Append(_crc, buffer);
                    _uncompressed += buffer.Length;
                    _sink.Write(buffer);
                }

                public override void WriteByte(byte value)
                {
                    Span<byte> buffer = stackalloc byte[1];
                    buffer[0] = value;
                    Write(buffer);
                }

                protected override void Dispose(bool disposing)
                {
                    if (disposing)
                        FinishFromWriter();
                    base.Dispose(disposing);
                }
            }
        }
    }

    /// <summary>Standard CRC-32 (IEEE 802.3) used by the ZIP format.</summary>
    internal static class Crc32
    {
        // Slicing-by-8: the simple byte-at-a-time loop costs tens of milliseconds on large
        // worksheets, while this version processes eight bytes per iteration.
        private static readonly uint[][] Tables = CreateTables();

        internal static uint Append(uint crc, ReadOnlySpan<byte> data)
        {
            uint value = ~crc;

            while (data.Length >= 8)
            {
                uint first = BinaryPrimitives.ReadUInt32LittleEndian(data);
                uint second = BinaryPrimitives.ReadUInt32LittleEndian(data[4..]);
                value =
                    Tables[7][(value ^ first) & 0xFF] ^
                    Tables[6][((value >> 8) ^ (first >> 8)) & 0xFF] ^
                    Tables[5][((value >> 16) ^ (first >> 16)) & 0xFF] ^
                    Tables[4][((value >> 24) ^ (first >> 24)) & 0xFF] ^
                    Tables[3][second & 0xFF] ^
                    Tables[2][(second >> 8) & 0xFF] ^
                    Tables[1][(second >> 16) & 0xFF] ^
                    Tables[0][(second >> 24) & 0xFF];
                data = data[8..];
            }

            foreach (byte b in data)
            {
                value = Tables[0][(value ^ b) & 0xFF] ^ (value >> 8);
            }

            return ~value;
        }

        private static uint[][] CreateTables()
        {
            var tables = new uint[8][];
            tables[0] = new uint[256];
            for (uint i = 0; i < 256; i++)
            {
                uint value = i;
                for (int bit = 0; bit < 8; bit++)
                {
                    value = (value & 1) != 0 ? 0xEDB88320u ^ (value >> 1) : value >> 1;
                }
                tables[0][i] = value;
            }

            for (int table = 1; table < tables.Length; table++)
            {
                tables[table] = new uint[256];
                for (int i = 0; i < 256; i++)
                {
                    uint previous = tables[table - 1][i];
                    tables[table][i] = (previous >> 8) ^ tables[0][previous & 0xFF];
                }
            }
            return tables;
        }
    }
}
