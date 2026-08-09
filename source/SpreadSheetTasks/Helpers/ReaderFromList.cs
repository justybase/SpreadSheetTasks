using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Data.Common;
using System.Diagnostics.CodeAnalysis;
using System.Linq;

namespace SpreadSheetTasks.Helpers;

internal sealed class ReaderFromList : DbDataReader
{
    private readonly List<object?[]> _rows;
    private readonly List<string> _headers;
    private readonly int _fieldCount;
    private readonly int _rowsCnt;
    private readonly string[] _typeNames;
    private readonly TypeCode[] _typeCodes;
    public ReaderFromList(List<object?[]> rows, List<string> headers, List<TypeCode> typeCodes)
    {
        _rows = rows;
        _headers = headers;

        _fieldCount = _headers.Count;
        _rowsCnt = _rows.Count();
        if (_rows.Count() > 0)
        {
            _currentRow = _rows[0];
        }
        _typeNames = new string[_fieldCount];
        _typeCodes = new TypeCode[_fieldCount];

        for (int i = 0; i < _fieldCount; i++)
        {
            _typeNames[i] = typeCodes[i].ToString();
            _typeCodes[i] = typeCodes[i];
        }
    }

    [return: DynamicallyAccessedMembers(DynamicallyAccessedMemberTypes.PublicProperties | DynamicallyAccessedMemberTypes.PublicFields)]
    private static Type GetTypeFor(TypeCode typeCode) => typeCode switch
    {
        TypeCode.Boolean => typeof(bool),
        TypeCode.Char => typeof(char),
        TypeCode.SByte => typeof(sbyte),
        TypeCode.Byte => typeof(byte),
        TypeCode.Int16 => typeof(short),
        TypeCode.UInt16 => typeof(ushort),
        TypeCode.Int32 => typeof(int),
        TypeCode.UInt32 => typeof(uint),
        TypeCode.Int64 => typeof(long),
        TypeCode.UInt64 => typeof(ulong),
        TypeCode.Single => typeof(float),
        TypeCode.Double => typeof(double),
        TypeCode.Decimal => typeof(decimal),
        TypeCode.DateTime => typeof(DateTime),
        TypeCode.String => typeof(string),
        TypeCode.DBNull => typeof(DBNull),
        TypeCode.Empty or TypeCode.Object => typeof(object),
        _ => typeof(object)
    };

    private int _currentRowNum = -1;
    private object?[] _currentRow;

    public override DataTable? GetSchemaTable()
    {
        return null;
    }

    public override object this[int ordinal] => _rows[_currentRowNum][ordinal];

    public override object this[string name] => throw new NotImplementedException();

    public override int Depth => throw new NotImplementedException();

    public override int FieldCount => _fieldCount;

    public override bool HasRows => _rowsCnt > 0;

    public override bool IsClosed => _currentRowNum > _rowsCnt;

    public override int RecordsAffected => throw new NotImplementedException();

    public override bool GetBoolean(int ordinal)
    {
        return (bool)_currentRow[ordinal];
    }

    public override byte GetByte(int ordinal)
    {
        return (byte)_currentRow[ordinal];
    }

    public override long GetBytes(int ordinal, long dataOffset, byte[] buffer, int bufferOffset, int length)
    {
        throw new NotImplementedException();
    }

    public override char GetChar(int ordinal)
    {
        return (char)_currentRow[ordinal];
    }

    public override long GetChars(int ordinal, long dataOffset, char[] buffer, int bufferOffset, int length)
    {
        throw new NotImplementedException();
    }

    public override string GetDataTypeName(int ordinal)
    {
        return _typeNames[ordinal];
    }

    public override DateTime GetDateTime(int ordinal)
    {
        return (DateTime)_currentRow[ordinal];
    }

    public override decimal GetDecimal(int ordinal)
    {
        return (Decimal)_currentRow[ordinal];
    }

    public override double GetDouble(int ordinal)
    {
        return (Double)_currentRow[ordinal];
    }

    [return: DynamicallyAccessedMembers(DynamicallyAccessedMemberTypes.PublicProperties | DynamicallyAccessedMemberTypes.PublicFields)]
    public override Type GetFieldType(int ordinal)
    {
        return GetTypeFor(_typeCodes[ordinal]);
    }

    public override float GetFloat(int ordinal)
    {
        return (float)_currentRow[ordinal];
    }

    public override Guid GetGuid(int ordinal)
    {
        return (Guid)_currentRow[ordinal];
    }

    public override short GetInt16(int ordinal)
    {
        return (short)_currentRow[ordinal];
    }

    public override int GetInt32(int ordinal)
    {
        return (int)_currentRow[ordinal];
    }

    public override long GetInt64(int ordinal)
    {
        return (long)_currentRow[ordinal];
    }

    public override string GetName(int ordinal)
    {
        return _headers[ordinal];
    }

    public override int GetOrdinal(string name)
    {
        throw new NotSupportedException();
    }

    public override string GetString(int ordinal)
    {
        return (string)_currentRow[ordinal];
    }

    public override object GetValue(int ordinal)
    {
        return _currentRow[ordinal];
    }

    public override int GetValues(object[] values)
    {
        for (int i = 0; i < _fieldCount; i++)
        {
            values[i] = _currentRow[i] ?? DBNull.Value;
        }
        return _fieldCount;
    }

    public override bool IsDBNull(int ordinal)
    {
        var val = _currentRow[ordinal];
        return val == null || val == DBNull.Value;
    }

    public override bool NextResult() => false;

    public override bool Read()
    {
        ++_currentRowNum;
        bool res = _currentRowNum < _rowsCnt;
        if (res)
        {
            _currentRow = _rows[_currentRowNum];
        }

        return res;
    }

    public override IEnumerator GetEnumerator()
    {
        throw new NotImplementedException();
    }

    public override bool Equals(object? obj)
    {
        return obj is ReaderFromList list &&
               EqualityComparer<object?[]>.Default.Equals(_currentRow, list._currentRow);
    }
}
