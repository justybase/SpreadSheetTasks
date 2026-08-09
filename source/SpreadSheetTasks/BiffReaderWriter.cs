using System;
using System.Buffers;
using System.Buffers.Binary;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Text;

namespace SpreadSheetTasks
{
    internal sealed class BiffReaderWriter : IDisposable
    {
        //private const int WorkbookPr = 0x99;
        private const int _sheet = 0x9C; // 156

        private const int _xf = 0x2f;

        private const int _cellXfStart = 0x269;
        private const int _cellXfEnd = 0x26a;

        private const int _cellStyleXfStart = 0x272;
        private const int _cellStyleXfEnd = 0x273;

        private const int _numberFormatStart = 0x267;
        private const int _numberFormat = 0x2c;
        private const int _numberFormatEnd = 0x268;

        private const int _sharedStringStart = 159;
        private const int _stringItem = 0x13; //19

        private const uint _row = 0x00;
        private const uint _blank = 0x01;
        private const uint _number = 0x02; // BrtCellRk
        private const uint _boolError = 0x03;
        private const uint _bool = 0x04;
        private const uint _float = 0x05;
        private const uint _string = 0x06;
        private const uint _sharedString = 0x07;
        private const uint _formulaString = 0x08;
        private const uint _formulaNumber = 0x09;
        private const uint _formulaBool = 0x0a;
        private const uint _formulaError = 0x0b;

        private readonly byte[] _buffer = new byte[128];
        private readonly Stream? Stream;

        // fast path: cale dane w pamieci (byte[]), zero kopii przy odczycie rekordow
        private readonly byte[]? _data;
        private int _pos;

        public BiffReaderWriter(Stream stream)
        {
            Stream = stream ?? throw new ArgumentNullException(nameof(stream));
        }

        public BiffReaderWriter(byte[] data)
        {
            _data = data ?? throw new ArgumentNullException(nameof(data));
        }

        private enum SheetVisibility : byte
        {
            Visible = 0x0,
            Hidden = 0x1,
            VeryHidden = 0x2
        }

        internal uint _workbookId;
        internal string _recId;
        internal string _workbookName;
        internal bool _isSheet;

        internal bool ReadWorkbook()
        {
            if (_data != null)
            {
                if (!BeginRecord(out var recordId, out var data))
                {
                    return false;
                }
                ParseWorkbookRecord(recordId, data);
                return true;
            }

            if (!TryReadVariableValue(out var recordId2) ||
                !TryReadVariableValue(out var recordLength))
                return false;
            byte[]? rented = null;
            byte[] buffer;
            try
            {
                if (recordLength <= _buffer.Length)
                    buffer = _buffer;
                else
                    buffer = rented = ArrayPool<byte>.Shared.Rent((int)recordLength);
                if (Stream!.Read(buffer, 0, (int)recordLength) != recordLength)
                    return false;

                ParseWorkbookRecord(recordId2, buffer.AsSpan(0, (int)recordLength));
                return true;
            }
            finally
            {
                if (rented != null)
                    ArrayPool<byte>.Shared.Return(rented);
            }
        }

        private void ParseWorkbookRecord(uint recordId, ReadOnlySpan<byte> data)
        {
            _isSheet = false;
            if (recordId == _sheet)
            {
                _workbookId = GetDWord(data, 4);

                int offset = 8;
                _recId = GetNullableString(data, ref offset);

                // Must be between 1 and 31 characters
                uint nameLength = GetDWord(data, offset);
                _workbookName = GetString(data, offset + 4, nameLength);
                _isSheet = true;
            }
        }

        internal bool _inCellXf;
        internal bool _inCellStyleXf;
        internal bool _inNumberFormat;

        internal ushort _parentCellStyleXf;
        internal ushort _numberFormatIndex;
        //public ushort FontIndex;

        internal int _format;
        internal string _formatString;

        public bool ReadStyles()
        {
            if (_data != null)
            {
                if (!BeginRecord(out var recordId, out var data))
                {
                    return false;
                }
                ParseStylesRecord(recordId, data);
                return true;
            }

            if (!TryReadVariableValue(out var recordId2) ||
                !TryReadVariableValue(out var recordLength))
                return false;

            byte[]? rented = null;
            byte[] buffer;
            try
            {
                if (recordLength <= _buffer.Length)
                    buffer = _buffer;
                else
                    buffer = rented = ArrayPool<byte>.Shared.Rent((int)recordLength);
                if (Stream!.Read(buffer, 0, (int)recordLength) != recordLength)
                    return false;

                ParseStylesRecord(recordId2, buffer.AsSpan(0, (int)recordLength));
                return true;
            }
            finally
            {
                if (rented != null)
                    ArrayPool<byte>.Shared.Return(rented);
            }
        }

        private void ParseStylesRecord(uint recordId, ReadOnlySpan<byte> data)
        {
            switch (recordId)
            {
                case _cellXfStart:
                    _inCellXf = true;
                    break;
                case _cellXfEnd:
                    _inCellXf = false;
                    break;
                case _cellStyleXfStart:
                    _inCellStyleXf = true;
                    break;
                case _cellStyleXfEnd:
                    _inCellStyleXf = false;
                    break;
                case _numberFormatStart:
                    _inNumberFormat = true;
                    break;
                case _numberFormatEnd:
                    _inNumberFormat = false;
                    break;

                case _xf when _inCellStyleXf:
                    break;
                case _xf when _inCellXf:
                    {
                        _parentCellStyleXf = GetWord(data, 0);
                        _numberFormatIndex = GetWord(data, 2);
                        //var FontIndex = GetWord(buffer, 4);
                        break;
                    }

                case _numberFormat when _inNumberFormat:
                    {
                        // Must be between 1 and 255 characters
                        _format = GetWord(data, 0);
                        uint length = GetDWord(data, 2);
                        _formatString = GetString(data, 2 + 4, length);

                        break;
                    }
            }
        }

        internal string? _sharedStringValue;
        internal uint _sharedStringUniqueCount = 0;
        public bool ReadSharedStrings()
        {
            if (_data != null)
            {
                if (!BeginRecord(out var recordId, out var data))
                {
                    return false;
                }
                ParseSharedStringsRecord(recordId, data);
                return true;
            }

            if (!TryReadVariableValue(out var recordId2) ||
                !TryReadVariableValue(out var recordLength))
                return false;

            byte[]? rented = null;
            byte[] buffer;
            try
            {
                if (recordLength <= _buffer.Length)
                    buffer = _buffer;
                else
                    buffer = rented = ArrayPool<byte>.Shared.Rent((int)recordLength);

                uint readed = 0;
                do
                {
                    readed += (uint)Stream!.Read(buffer, (int)readed, (int)(recordLength - readed));
                    if (readed == 0)
                    {
                        return false;
                    }
                } while (readed < recordLength);

                ParseSharedStringsRecord(recordId2, buffer.AsSpan(0, (int)recordLength));
                return true;
            }
            finally
            {
                if (rented != null)
                    ArrayPool<byte>.Shared.Return(rented);
            }
        }

        private void ParseSharedStringsRecord(uint recordId, ReadOnlySpan<byte> data)
        {
            if (recordId == _stringItem)
            {
                uint length = GetDWord(data, 1);
                _sharedStringValue = GetString(data, 1 + 4, length);
            }
            else if (recordId == _sharedStringStart)
            {
                _sharedStringUniqueCount = GetDWord(data, 4);
                _sharedStringValue = null;
            }
            else
            {
                _sharedStringValue = null;
            }
        }

        //public object cellValue;
        internal CellType _cellType;
        internal int _intValue;
        internal double _doubleVal;
        internal bool _boolValue;
        internal string _stringValue;

        internal int _columnNum = -1;
        internal uint _xfIndex;
        //public bool isSharedStringVal = false;
        internal bool _readCell = false;
        internal int _rowIndex = -1;

        internal bool ReadWorksheet()
        {
            if (_data != null)
            {
                if (!BeginRecord(out var recordId, out var data))
                {
                    return false;
                }
                ParseWorksheetRecord(recordId, data);
                return true;
            }

            if (!TryReadVariableValue(out var recordId2) ||
                !TryReadVariableValue(out var recordLength))
                return false;

            byte[]? rented = null;
            byte[] buffer;
            try
            {
                if (recordLength <= _buffer.Length)
                    buffer = _buffer;
                else
                    buffer = rented = ArrayPool<byte>.Shared.Rent((int)recordLength);
                if (Stream!.Read(buffer, 0, (int)recordLength) != recordLength)
                    return false;

                ParseWorksheetRecord(recordId2, buffer.AsSpan(0, (int)recordLength));
                return true;
            }
            finally
            {
                if (rented != null)
                    ArrayPool<byte>.Shared.Return(rented);
            }
        }

        private void ParseWorksheetRecord(uint recordId, ReadOnlySpan<byte> data)
        {
            _readCell = false;
            _columnNum = -1;
            //isSharedStringVal = false;

            switch (recordId)
            {
                case _row: // BrtRowHdr 0 = 0x0000
                    {
                        _rowIndex = GetInt32(data, 0);
                        //    byte flags = buffer[11];
                        //    bool hidden = (flags & 0b10000) != 0;
                        //    bool unsynced = (flags & 0b100000) != 0;

                        //    double? height = null;
                        //    if (unsynced)
                        //        height = GetWord(buffer, 8) / 20.0; // Where does 20 come from?

                        //    // TODO: Default format ?
                        break;
                    }
                case _blank: //BrtCellBlank (1 = 0x0001)
                case _boolError:
                case _formulaError: // BrtFmlaError (11 = 0x000B)
                    _readCell = true;
                    _cellType = CellType.nullValue;
                    break;
                case _number:
                    _doubleVal = GetRkNumber(data, 8);
                    _readCell = true;
                    _cellType = CellType.doubleVal;
                    break;
                case _bool:
                case _formulaBool:
                    _boolValue = (data[8] == 1);
                    _readCell = true;
                    _cellType = CellType.boolVal;
                    break;
                case _formulaNumber:
                case _float:
                    _doubleVal = GetDouble(data, 8);
                    _readCell = true;
                    _cellType = CellType.doubleVal;
                    break;
                case _string:
                case _formulaString:
                    {
                        // Must be less than 32768 characters
                        var length = GetDWord(data, 8);
                        _stringValue = GetString(data, 8 + 4, length);
                        _readCell = true;
                        _cellType = CellType.stringVal;
                        break;
                    }
                case _sharedString:
                    _intValue = (int)GetDWord(data, 8);
                    _readCell = true;
                    _cellType = CellType.sharedString;
                    break;
            }

            if (_readCell)
            {
                _columnNum = (int)GetDWord(data, 0);
                _xfIndex = GetDWord(data, 4) & 0xffffff;
            }
        }

        //https://github.com/ExcelDataReader/ExcelDataReader
        static uint GetDWord(ReadOnlySpan<byte> buffer, int offset)
        {
            return BinaryPrimitives.ReadUInt32LittleEndian(buffer.Slice(offset, 4));
        }

        //https://github.com/ExcelDataReader/ExcelDataReader
        static int GetInt32(ReadOnlySpan<byte> buffer, int offset)
        {
            return BinaryPrimitives.ReadInt32LittleEndian(buffer.Slice(offset, 4));
        }

        //https://github.com/ExcelDataReader/ExcelDataReader
        static ushort GetWord(ReadOnlySpan<byte> buffer, int offset)
        {
            return BinaryPrimitives.ReadUInt16LittleEndian(buffer.Slice(offset, 2));
        }

        private static string GetString(ReadOnlySpan<byte> buffer, int offset, uint length)
        {
            return Encoding.Unicode.GetString(buffer.Slice(offset, (int)length * 2));
        }

        //https://github.com/ExcelDataReader/ExcelDataReader
        static string GetNullableString(ReadOnlySpan<byte> buffer, ref int offset)
        {
            var length = GetDWord(buffer, offset);
            offset += 4;
            if (length == uint.MaxValue)
                return null;

            string result = Encoding.Unicode.GetString(buffer.Slice(offset, (int)length * 2));
            offset += (int)length * 2;
            return result;
        }

        //https://github.com/ExcelDataReader/ExcelDataReader
        //2.5.122 RkNumber
        static double GetRkNumber(ReadOnlySpan<byte> buffer, int offset)
        {
            double result;

            byte flags = buffer[offset];

            if ((flags & 0x02) != 0)
            {
                result = GetInt32(buffer, offset) >> 2;
            }
            else
            {
                result = BitConverter.Int64BitsToDouble((long)((ulong)(GetDWord(buffer, offset) & -4) << 32));
            }

            if ((flags & 0x01) != 0)
            {
                result /= 100;
            }

            return result;
        }

        //https://github.com/ExcelDataReader/ExcelDataReader
        static double GetDouble(ReadOnlySpan<byte> buffer, int offset)
        {
            return BitConverter.Int64BitsToDouble(BinaryPrimitives.ReadInt64LittleEndian(buffer.Slice(offset, 8)));
        }

        //https://github.com/ExcelDataReader/ExcelDataReader

        private bool TryReadVariableValue(out uint value)
        {
            value = 0;

            if (Stream!.Read(_buffer, 0, 1) == 0)
                return false;

            byte b1 = _buffer[0];
            value = (uint)(b1 & 0x7F);

            if ((b1 & 0x80) == 0)
                return true;

            if (Stream.Read(_buffer, 0, 1) == 0)
                return false;
            byte b2 = _buffer[0];
            value = ((uint)(b2 & 0x7F) << 7) | value;

            if ((b2 & 0x80) == 0)
                return true;

            if (Stream.Read(_buffer, 0, 1) == 0)
                return false;
            byte b3 = _buffer[0];
            value = ((uint)(b3 & 0x7F) << 14) | value;

            if ((b3 & 0x80) == 0)
                return true;

            if (Stream.Read(_buffer, 0, 1) == 0)
                return false;
            byte b4 = _buffer[0];
            value = ((uint)(b4 & 0x7F) << 21) | value;

            return true;
        }

        // fast path: naglowek rekordu (varint id + varint dlugosc) z bufora w pamieci
        private bool BeginRecord(out uint recordId, out ReadOnlySpan<byte> data)
        {
            recordId = 0;
            data = default;
            if (!ReadVarint(out recordId) || !ReadVarint(out uint length))
            {
                return false;
            }
            byte[] bytes = _data!;
            if (length > (uint)(bytes.Length - _pos))
            {
                return false;
            }
            data = bytes.AsSpan(_pos, (int)length);
            _pos += (int)length;
            return true;
        }

        private bool ReadVarint(out uint value)
        {
            value = 0;
            byte[] bytes = _data!;
            int pos = _pos;

            if (pos >= bytes.Length)
            {
                return false;
            }
            byte b1 = bytes[pos++];
            value = (uint)(b1 & 0x7F);
            if ((b1 & 0x80) == 0)
            {
                _pos = pos;
                return true;
            }

            if (pos >= bytes.Length)
            {
                return false;
            }
            byte b2 = bytes[pos++];
            value = ((uint)(b2 & 0x7F) << 7) | value;
            if ((b2 & 0x80) == 0)
            {
                _pos = pos;
                return true;
            }

            if (pos >= bytes.Length)
            {
                return false;
            }
            byte b3 = bytes[pos++];
            value = ((uint)(b3 & 0x7F) << 14) | value;
            if ((b3 & 0x80) == 0)
            {
                _pos = pos;
                return true;
            }

            if (pos >= bytes.Length)
            {
                return false;
            }
            byte b4 = bytes[pos++];
            value = ((uint)(b4 & 0x7F) << 21) | value;
            _pos = pos;
            return true;
        }

        public void Dispose()
        {
            Stream?.Dispose();
        }

        public override bool Equals(object? obj)
        {
            return obj is BiffReaderWriter writer &&
                   _workbookId == writer._workbookId;
        }
        //void Dispose(bool disposing)
        //{
        //    if (disposing)
        //        Stream.Dispose();
        //}
    }

    internal class DataColReader
    {
        internal readonly IDataReader _dataReader;
        internal DataTable _dataTable;
        private readonly object[,] _tabelarData;
        private readonly bool _isDataReader;
        private readonly bool _isDataTable;
        internal int _dataTableRowsCount;

        private readonly bool _headers;
        private int _rowNum = 0;

        internal string[] _databaseTypes;

        public DataColReader(IDataReader reader, Boolean headers = false, int maxRows = -1)
        {
            this._dataReader = reader;
            this._headers = headers;
            this._isDataReader = true;
            this._overLimit = maxRows;

            _databaseTypes = new string[_dataReader.FieldCount];
            for (int i = 0; i < _databaseTypes.Length; i++)
            {
                _databaseTypes[i] = _dataReader.GetDataTypeName(i);
            }
        }

        public DataColReader(DataTable dataTable, Boolean headers = false, int maxRows = -1)
        {
            this._dataTable = dataTable;
            this._headers = headers;
            this._isDataTable = true;
            this._overLimit = maxRows;
            this._dataTableRowsCount = dataTable.Rows.Count;

            _databaseTypes = new string[_dataTable.Columns.Count];

            // WORK TO DO !!
            for (int i = 0; i < _databaseTypes.Length; i++)
            {
                _databaseTypes[i] = _dataTable.Columns[i].DataType.ToString();
            }
        }

        public DataColReader(string[,] tabelarData)
        {
            this._tabelarData = tabelarData;
            _isDataReader = false;
            _databaseTypes = new string[tabelarData.Length];
            for (int i = 0; i < _databaseTypes.Length; i++)
            {
                _databaseTypes[i] = "-1";
            }
        }

        private readonly int _overLimit = -1;
        public int FieldCount    // the Name property
        {
            get
            {
                if (_isDataReader && _overLimit > 0)
                {
                    return _overLimit;
                }
                else if (_isDataReader)
                {
                    return _dataReader.FieldCount;
                }
                else if (_isDataTable)
                {
                    return _dataTable.Columns.Count;
                }
                else
                {
                    return _tabelarData.GetUpperBound(1) + 1;
                }
            }
        }
        public bool Read()
        {
            ++_rowNum;

            if (_isDataReader)
            {
                if (_isDataReader && _rowNum <= 1 && _headers)
                {
                    return true;
                }
                else if (top100 != null && _topNum <= top100.Count)
                {
                    _topNum++;
                    if (_topNum == top100.Count + 1)
                    {
                        top100 = null;
                        return AreNextRows;
                    }
                    return true;
                }
                else
                {
                    return _dataReader.Read();
                }
            }
            else if (_isDataTable)
            {
                int dataStartRow = _headers ? 2 : 1;
                int maxDataRow = _dataTableRowsCount + dataStartRow;
                if (_rowNum >= dataStartRow && _rowNum < maxDataRow)
                {
                    _dataTableRow = _dataTable.Rows[_rowNum - dataStartRow].ItemArray;
                    return true;
                }
                else if (_headers && _rowNum == 1)
                {
                    return true;
                }
                return false;
            }
            else
            {
                return (_rowNum < _tabelarData.GetUpperBound(0) + 2);
            }
        }

        private object[]? _dataTableRow;
        public object GetValue(int j)
        {
            if (_isDataReader)
            {
                if ((_rowNum > 1 || !_headers) && top100 == null)
                {
                    return _dataReader.GetValue(j);
                }
                else if (_headers && _rowNum == 1)
                {
                    return _dataReader.GetName(j);
                }
                else
                {
                    return top100[_topNum - 1][j];
                }
            }
            else if (_isDataTable)
            {
                if (_rowNum > 1 || !_headers)
                {
                    return _dataTableRow[j];
                }
                else
                {
                    return _dataTable.Columns[j].ColumnName;
                }
            }
            else
            {
                return _tabelarData[_rowNum - 1, j];
            }
        }

        
        public bool GetBoolean(int j)
        {
            if (_isDataReader)
            {
                if ((_rowNum > 1 || !_headers) && top100 == null)
                {
                    return _dataReader.GetBoolean(j);
                }
                else if (_headers && _rowNum == 1)
                {
                    throw new Exception("bool for header ?");
                }
                else
                {
                    return (bool)top100[_topNum - 1][j];
                }
            }
            else if (_isDataTable)
            {
                if (_rowNum > 1 || !_headers)
                {
                    //return DataTable.Rows[_rowNum-2][j];
                    return (bool)_dataTableRow[j];
                }
                else
                {
                    throw new Exception("bool for header ?");
                }
            }
            else
            {
                return (bool)_tabelarData[_rowNum - 1, j];
            }
        }

        public char GetChar(int j)
        {
            if (_isDataReader)
            {
                if ((_rowNum > 1 || !_headers) && top100 == null)
                {
                    return _dataReader.GetChar(j);
                }
                else if (_headers && _rowNum == 1)
                {
                    throw new Exception("char for header ?");
                }
                else
                {
                    return (char)top100[_topNum - 1][j];
                }
            }
            else if (_isDataTable)
            {
                if (_rowNum > 1 || !_headers)
                {
                    //return DataTable.Rows[_rowNum-2][j];
                    return (char)_dataTableRow[j];
                }
                else
                {
                    throw new Exception("char for header ?");
                }
            }
            else
            {
                return (char)_tabelarData[_rowNum - 1, j];
            }
        }

        public byte GetByte(int j)
        {
            if (_isDataReader)
            {
                if ((_rowNum > 1 || !_headers) && top100 == null)
                {
                    return _dataReader.GetByte(j);
                }
                else if (_headers && _rowNum == 1)
                {
                    throw new Exception("byte for header ?");
                }
                else
                {
                    return (byte)top100[_topNum - 1][j];
                }
            }
            else if (_isDataTable)
            {
                if (_rowNum > 1 || !_headers)
                {
                    //return DataTable.Rows[_rowNum-2][j];
                    return (byte)_dataTableRow[j];
                }
                else
                {
                    throw new Exception("byte for header ?");
                }
            }
            else
            {
                return (byte)_tabelarData[_rowNum - 1, j];
            }
        }

        public sbyte GetSByte(int j)
        {
            if (_isDataReader)
            {
                if ((_rowNum > 1 || !_headers) && top100 == null)
                {
                    return (sbyte)_dataReader.GetValue(j);
                }
                else if (_headers && _rowNum == 1)
                {
                    throw new Exception("byte for header ?");
                }
                else
                {
                    return (sbyte)top100[_topNum - 1][j];
                }
            }
            else if (_isDataTable)
            {
                if (_rowNum > 1 || !_headers)
                {
                    //return DataTable.Rows[_rowNum-2][j];
                    return (sbyte)_dataTableRow[j];
                }
                else
                {
                    throw new Exception("sbyte for header ?");
                }
            }
            else
            {
                return (sbyte)_tabelarData[_rowNum - 1, j];
            }
        }

        public Int16 GetInt16(int j)
        {
            if (_isDataReader)
            {
                if ((_rowNum > 1 || !_headers) && top100 == null)
                {
                    return _dataReader.GetInt16(j);
                }
                else if (_headers && _rowNum == 1)
                {
                    throw new Exception("Int16 for header ?");
                }
                else
                {
                    return (Int16)top100[_topNum - 1][j];
                }
            }
            else if (_isDataTable)
            {
                if (_rowNum > 1 || !_headers)
                {
                    //return DataTable.Rows[_rowNum-2][j];
                    return (Int16)_dataTableRow[j];
                }
                else
                {
                    throw new Exception("Int16 for header ?");
                }
            }
            else
            {
                return (Int16)_tabelarData[_rowNum - 1, j];
            }
        }

        public Int32 GetInt32(int j)
        {
            if (_isDataReader)
            {
                if ((_rowNum > 1 || !_headers) && top100 == null)
                {
                    return _dataReader.GetInt32(j);
                }
                else if (_headers && _rowNum == 1)
                {
                    throw new Exception("Int32 for header ?");
                }
                else
                {
                    return (Int32)top100[_topNum - 1][j];
                }
            }
            else if (_isDataTable)
            {
                if (_rowNum > 1 || !_headers)
                {
                    //return DataTable.Rows[_rowNum-2][j];
                    return (Int32)_dataTableRow[j];
                }
                else
                {
                    throw new Exception("Int32 for header ?");
                }
            }
            else
            {
                return (Int32)_tabelarData[_rowNum - 1, j];
            }
        }

        public Int64 GetInt64(int j)
        {
            if (_isDataReader)
            {
                if ((_rowNum > 1 || !_headers) && top100 == null)
                {
                    return _dataReader.GetInt64(j);
                }
                else if (_headers && _rowNum == 1)
                {
                    throw new Exception("Int64 for header ?");
                }
                else
                {
                    return (Int64)top100[_topNum - 1][j];
                }
            }
            else if (_isDataTable)
            {
                if (_rowNum > 1 || !_headers)
                {
                    //return DataTable.Rows[_rowNum-2][j];
                    return (Int64)_dataTableRow[j];
                }
                else
                {
                    throw new Exception("Int64 for header ?");
                }
            }
            else
            {
                return (Int32)_tabelarData[_rowNum - 1, j];
            }
        }

        public float GetFloat(int j)
        {
            if (_isDataReader)
            {
                if ((_rowNum > 1 || !_headers) && top100 == null)
                {
                    return _dataReader.GetFloat(j);
                }
                else if (_headers && _rowNum == 1)
                {
                    throw new Exception("float for header ?");
                }
                else
                {
                    return (float)top100[_topNum - 1][j];
                }
            }
            else if (_isDataTable)
            {
                if (_rowNum > 1 || !_headers)
                {
                    //return DataTable.Rows[_rowNum-2][j];
                    return (float)_dataTableRow[j];
                }
                else
                {
                    throw new Exception("float for header ?");
                }
            }
            else
            {
                return (float)_tabelarData[_rowNum - 1, j];
            }
        }
        public double GetDouble(int j)
        {
            if (_isDataReader)
            {
                if ((_rowNum > 1 || !_headers) && top100 == null)
                {
                    return _dataReader.GetDouble(j);
                }
                else if (_headers && _rowNum == 1)
                {
                    throw new Exception("double for header ?");
                }
                else
                {
                    return (double)top100[_topNum - 1][j];
                }
            }
            else if (_isDataTable)
            {
                if (_rowNum > 1 || !_headers)
                {
                    //return DataTable.Rows[_rowNum-2][j];
                    return (double)_dataTableRow[j];
                }
                else
                {
                    throw new Exception("double for header ?");
                }
            }
            else
            {
                return (double)_tabelarData[_rowNum - 1, j];
            }
        }
        public decimal GetDecimal(int j)
        {
            if (_isDataReader)
            {
                if ((_rowNum > 1 || !_headers) && top100 == null)
                {
                    return _dataReader.GetDecimal(j);
                }
                else if (_headers && _rowNum == 1)
                {
                    throw new Exception("decimal for header ?");
                }
                else
                {
                    return (decimal)top100[_topNum - 1][j];
                }
            }
            else if (_isDataTable)
            {
                if (_rowNum > 1 || !_headers)
                {
                    //return DataTable.Rows[_rowNum-2][j];
                    return (decimal)_dataTableRow[j];
                }
                else
                {
                    throw new Exception("decimal for header ?");
                }
            }
            else
            {
                return (decimal)_tabelarData[_rowNum - 1, j];
            }
        }

        public DateTime GetDateTime(int j)
        {
            if (_isDataReader)
            {
                if ((_rowNum > 1 || !_headers) && top100 == null)
                {
                    return _dataReader.GetDateTime(j);
                }
                else if (_headers && _rowNum == 1)
                {
                    throw new Exception("DateTime for header ?");
                }
                else
                {
                    return (DateTime)top100[_topNum - 1][j];
                }
            }
            else if (_isDataTable)
            {
                if (_rowNum > 1 || !_headers)
                {
                    return (DateTime)_dataTableRow[j];
                }
                else
                {
                    throw new Exception("decimal for header ?");
                }
            }
            else
            {
                return (DateTime)_tabelarData[_rowNum - 1, j];
            }
        }

        public string GetString(int j)
        {
            if (_isDataReader)
            {
                if ((_rowNum > 1 || !_headers) && top100 == null)
                {
                    return _dataReader.GetString(j);
                }
                else if (_headers && _rowNum == 1)
                {
                    return _dataReader.GetName(j);
                }
                else
                {
                    return top100[_topNum - 1][j].ToString();
                }
            }
            else if (_isDataTable)
            {
                if (_rowNum > 1 || !_headers)
                {
                    //return DataTable.Rows[_rowNum-2][j];
                    return _dataTableRow[j].ToString();
                }
                else
                {
                    return _dataTable.Columns[j].ColumnName;
                }
            }
            else
            {
                return _tabelarData[_rowNum - 1, j].ToString();
            }
        }

        public bool IsDBNull(int j)
        {
            if (_isDataReader)
            {
                if ((_rowNum > 1 || !_headers) && top100 == null)
                {
                    return _dataReader.IsDBNull(j);
                }
                else if (_headers && _rowNum == 1)
                {
                    return false;
                }
                else
                {
                    return top100[_topNum - 1][j] == null || top100[_topNum - 1][j] == DBNull.Value;
                }
            }
            else if (_isDataTable)
            {
                if (_rowNum > 1 || !_headers)
                {
                    //return DataTable.Rows[_rowNum-2][j];
                    return _dataTableRow[j] == null || _dataTableRow[j] == DBNull.Value;
                }
                else
                {
                    return _dataTable.Columns[j].ColumnName == null;
                }
            }
            else
            {
                return _tabelarData[_rowNum - 1, j] == null || _tabelarData[_rowNum - 1, j] == DBNull.Value;
            }
        }

        public void GetWidthFromDataTable(Span<double> width, double maxWidth, bool doAutofilter)
        {
            int n = _dataTableRowsCount > 100 ? 100 : _dataTableRowsCount;
            int m = FieldCount;

            for (int j = 0; j < m; j++)
            {
                double valTemp = 1.25 * _dataTable.Columns[j].ToString().Length + 2;
                if (doAutofilter)
                {
                    valTemp += 2;
                }

                if (valTemp > maxWidth)
                {
                    valTemp = maxWidth;
                }

                if (width[j] < valTemp)
                {
                    width[j] = valTemp;
                }
            }

            for (int i = 0; i < n; i++)
            {
                for (int j = 0; j < m; j++)
                {
                    double valTemp = 1.25 * _dataTable.Rows[i][j].ToString().Length + 2;
                    if (valTemp > maxWidth)
                    {
                        valTemp = maxWidth;
                    }

                    if (width[j] < valTemp)
                    {
                        width[j] = valTemp;
                    }
                }
            }
        }
        public bool AreNextRows { get; set; }
        private int _topNum = 0;
        public List<object[]> top100;
    }

    //https://github.com/ExcelDataReader/ExcelDataReader
    //https://docs.microsoft.com/en-us/openspecs/office_file_formats/ms-xlsb/aa9f2bac-991a-42a8-8cfa-507de84017b6


    internal enum CellType
    {
        doubleVal,
        boolVal,
        stringVal,
        sharedString,
        nullValue
    }
}
