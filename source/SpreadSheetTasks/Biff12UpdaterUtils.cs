using System.Buffers.Binary;
using System.Text;

namespace SpreadSheetTasks;

internal readonly struct Biff12Record
{
    internal Biff12Record(int headerStart, int dataStart, int dataEnd, uint id, int length)
    {
        HeaderStart = headerStart;
        DataStart = dataStart;
        DataEnd = dataEnd;
        Id = id;
        Length = length;
    }

    internal int HeaderStart { get; }
    internal int DataStart { get; }
    internal int DataEnd { get; }
    internal uint Id { get; }
    internal int Length { get; }
}

internal static class Biff12UpdaterUtils
{
    internal static bool TryReadRecord(ReadOnlySpan<byte> buffer, int position, out Biff12Record record)
    {
        record = default;
        if (position < 0 || position >= buffer.Length)
            return false;

        if (!TryReadVlq(buffer, position, out uint id, out int afterId))
            return false;
        if (!TryReadVlq(buffer, afterId, out uint length, out int afterLength))
            return false;
        if (length > int.MaxValue || afterLength > buffer.Length - (int)length)
            return false;

        record = new Biff12Record(position, afterLength, afterLength + (int)length, id, (int)length);
        return true;
    }

    internal static bool TryReadVlq(ReadOnlySpan<byte> buffer, int position, out uint value, out int nextPosition)
    {
        value = 0;
        nextPosition = position;
        if (position < 0 || position >= buffer.Length)
            return false;

        int shift = 0;
        for (int i = 0; i < 5; i++)
        {
            if (nextPosition >= buffer.Length)
                return false;
            byte current = buffer[nextPosition++];
            if (i == 4 && (current & 0xF0) != 0)
                return false;

            value |= (uint)(current & 0x7F) << shift;
            if ((current & 0x80) == 0)
                return true;
            shift += 7;
        }
        return false;
    }

    internal static byte[] BuildRecord(uint id, ReadOnlySpan<byte> payload)
    {
        Span<byte> idBuffer = stackalloc byte[5];
        Span<byte> lengthBuffer = stackalloc byte[5];
        int idLength = WriteVlq(idBuffer, id);
        int payloadLength = WriteVlq(lengthBuffer, checked((uint)payload.Length));

        var result = new byte[idLength + payloadLength + payload.Length];
        idBuffer[..idLength].CopyTo(result);
        lengthBuffer[..payloadLength].CopyTo(result.AsSpan(idLength));
        payload.CopyTo(result.AsSpan(idLength + payloadLength));
        return result;
    }

    internal static int WriteVlq(Span<byte> destination, uint value)
    {
        int position = 0;
        while (value >= 0x80)
        {
            destination[position++] = (byte)((value & 0x7F) | 0x80);
            value >>= 7;
        }
        destination[position++] = (byte)value;
        return position;
    }

    internal static string ReadUtf16(ReadOnlySpan<byte> buffer, int start, int characterCount)
    {
        if (characterCount < 0 || start < 0 || characterCount > (buffer.Length - start) / 2)
            throw new InvalidDataException("BIFF12 UTF-16 string exceeds the record boundary.");
        return Encoding.Unicode.GetString(buffer.Slice(start, characterCount * 2));
    }

    internal static int ReadInt32(ReadOnlySpan<byte> buffer, int offset)
    {
        if (offset < 0 || offset > buffer.Length - sizeof(int))
            throw new InvalidDataException("BIFF12 integer exceeds the record boundary.");
        return BinaryPrimitives.ReadInt32LittleEndian(buffer[offset..]);
    }

    internal static uint ReadUInt32(ReadOnlySpan<byte> buffer, int offset)
    {
        if (offset < 0 || offset > buffer.Length - sizeof(uint))
            throw new InvalidDataException("BIFF12 unsigned integer exceeds the record boundary.");
        return BinaryPrimitives.ReadUInt32LittleEndian(buffer[offset..]);
    }

    internal static void WriteInt32(Span<byte> buffer, int offset, int value)
    {
        if (offset < 0 || offset > buffer.Length - sizeof(int))
            throw new InvalidDataException("BIFF12 integer exceeds the record boundary.");
        BinaryPrimitives.WriteInt32LittleEndian(buffer[offset..], value);
    }

    internal static void WriteUInt32(Span<byte> buffer, int offset, uint value)
    {
        if (offset < 0 || offset > buffer.Length - sizeof(uint))
            throw new InvalidDataException("BIFF12 unsigned integer exceeds the record boundary.");
        BinaryPrimitives.WriteUInt32LittleEndian(buffer[offset..], value);
    }

    internal static void WriteDouble(Span<byte> buffer, int offset, double value)
    {
        if (offset < 0 || offset > buffer.Length - sizeof(double))
            throw new InvalidDataException("BIFF12 double exceeds the record boundary.");
        BinaryPrimitives.WriteDoubleLittleEndian(buffer[offset..], value);
    }

    internal sealed class SharedStringsState
    {
        internal byte[] Buffer { get; set; } = [];
        internal List<string> Values { get; } = [];
        internal Dictionary<string, int> Index { get; } = new(StringComparer.Ordinal);
        internal uint TotalCount { get; set; }
        internal uint UniqueCount { get; set; }
        internal int EndSstOffset { get; set; }
        internal int CommittedValueCount { get; set; }
        internal bool Dirty { get; set; }
    }

    internal static SharedStringsState ParseSharedStrings(byte[] buffer)
    {
        var state = new SharedStringsState
        {
            Buffer = buffer.ToArray(),
            EndSstOffset = buffer.Length
        };

        int position = 0;
        while (TryReadRecord(buffer, position, out var record))
        {
            if (record.Id == 0x009F && record.Length >= 8)
            {
                state.TotalCount = ReadUInt32(buffer, record.DataStart);
                state.UniqueCount = ReadUInt32(buffer, record.DataStart + 4);
            }
            else if (record.Id == 0x0013 && record.Length >= 5)
            {
                uint count = ReadUInt32(buffer, record.DataStart + 1);
                if (count <= int.MaxValue && count <= (uint)((record.DataEnd - (record.DataStart + 5)) / 2))
                {
                    string value = ReadUtf16(buffer, record.DataStart + 5, (int)count);
                    state.Index.TryAdd(value, state.Values.Count);
                    state.Values.Add(value);
                }
            }
            else if (record.Id == 0x00A0)
            {
                state.EndSstOffset = record.HeaderStart;
            }
            position = record.DataEnd;
        }

        state.CommittedValueCount = state.Values.Count;
        return state;
    }
}
